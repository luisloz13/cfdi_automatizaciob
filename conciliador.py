"""Conciliacion bancaria contra CFDI descargados, con clasificacion contable opcional.

Pasos:
  1. Lee los XML (CFDI 3.3 y 4.0) que dejo cfdi_sat.py para el cliente/mes.
  2. Lee el estado de cuenta (CSV/Excel) y lo concilia contra los CFDI.
  3. (Opcional, con --provider) clasifica con IA los CFDI recibidos y los cargos sin CFDI.
  4. Genera un Excel con: Resumen, Conciliados, Mov sin CFDI, CFDI sin mov y CFDI (detalle).

Uso:
    python conciliador.py --rfc XAXX010101000 --anio 2026 --mes 9 --estado-cuenta banco.xlsx
    python conciliador.py --rfc XAXX010101000 --anio 2026 --mes 9 --estado-cuenta banco.xlsx --provider azure
    python conciliador.py --rfc XAXX010101000 --anio 2026 --mes 9          # solo leer/clasificar CFDI

Estado de cuenta: columnas `fecha`, `concepto` (o descripcion) y `monto` (negativo = cargo),
o bien `cargo` y `abono` por separado. Ver README para el detalle.
"""
import argparse
import itertools
import logging
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import pandas as pd

import auto_clasificador as ia

log = logging.getLogger("conciliador")

DESTINO_DEFAULT = r"G:\Mi unidad\CFDI"
RE_RFC = re.compile(r"^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$")
PALABRAS_RUIDO = {"SA", "DE", "CV", "SAPI", "SC", "RL", "LA", "EL", "LOS", "LAS", "Y", "SOCIEDAD",
                  "ANONIMA", "CAPITAL", "VARIABLE", "MEXICO", "MEXICANA", "SERVICIOS", "GRUPO"}
ALIAS = {
    "fecha": {"fecha", "fecha operacion", "fecha de operacion", "fecha movimiento", "date"},
    "concepto": {"concepto", "descripcion", "descripción", "detalle", "movimiento", "referencia"},
    "monto": {"monto", "importe"},
    "cargo": {"cargo", "cargos", "retiro", "retiros"},
    "abono": {"abono", "abonos", "deposito", "depositos", "depósito", "depósitos"},
}
FORMATOS_FECHA = ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%Y/%m/%d", "%d.%m.%Y")

# Signo esperado del movimiento bancario (+1 abono, -1 cargo) segun (direccion, tipo de CFDI).
SENTIDO = {
    ("Emitido", "I"): 1, ("Emitido", "P"): 1, ("Emitido", "E"): -1, ("Emitido", "N"): -1,
    ("Recibido", "I"): -1, ("Recibido", "P"): -1, ("Recibido", "E"): 1, ("Recibido", "N"): 1,
}


# --------------------------------------------------------------------------- #
# Modelos
# --------------------------------------------------------------------------- #
@dataclass
class Cfdi:
    uuid: str
    fecha: date
    tipo: str                 # I ingreso, E egreso, P pago, N nomina, T traslado
    serie_folio: str
    direccion: str            # Emitido | Recibido (respecto al cliente)
    emisor_rfc: str
    emisor_nombre: str
    receptor_rfc: str
    receptor_nombre: str
    uso_cfdi: str
    metodo_pago: str
    forma_pago: str
    moneda: str
    subtotal: float
    iva: float
    total: float              # total del comprobante en MXN
    monto: float              # importe a conciliar en MXN (en tipo P: suma de los pagos)
    conceptos: str
    clave_prod: str
    archivo: str
    nota: str = ""
    clasificacion: dict = field(default_factory=dict)

    @property
    def sentido(self) -> int:
        return SENTIDO[(self.direccion, self.tipo)]

    @property
    def conciliable(self) -> bool:
        """Las facturas PPD se pagan con su complemento (tipo P): ese es el que cruza con el banco."""
        return self.monto > 0 and not (self.tipo == "I" and self.metodo_pago == "PPD")

    @property
    def contraparte(self) -> tuple[str, str]:
        return ((self.emisor_rfc, self.emisor_nombre) if self.direccion == "Recibido"
                else (self.receptor_rfc, self.receptor_nombre))


@dataclass
class Mov:
    idx: int
    fila: int
    fecha: date
    concepto: str
    monto: float              # con signo: negativo = cargo
    clasificacion: dict = field(default_factory=dict)


@dataclass
class Par:
    mov: Mov
    cfdis: list
    score: int
    confianza: str
    criterio: str


# --------------------------------------------------------------------------- #
# Lectura de CFDI
# --------------------------------------------------------------------------- #
def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _num(valor, defecto: float = 0.0) -> float:
    try:
        return float(valor)
    except (TypeError, ValueError):
        return defecto


def parsear_cfdi(path: Path, rfc_cliente: str) -> Cfdi:
    """Extrae los datos relevantes de un XML CFDI 3.3/4.0. Levanta ValueError si no es utilizable."""
    try:
        root = ET.parse(path).getroot()
    except ET.ParseError as exc:
        raise ValueError(f"XML mal formado: {exc}") from exc
    if _local(root.tag) != "Comprobante":
        raise ValueError("no es un Comprobante CFDI")

    a = root.attrib
    hijos = {_local(e.tag): e for e in root}
    emisor, receptor = hijos.get("Emisor"), hijos.get("Receptor")
    timbre = next((e for e in root.iter() if _local(e.tag) == "TimbreFiscalDigital"), None)
    uuid = (timbre.get("UUID", "") if timbre is not None else "").strip().upper()
    if emisor is None or receptor is None or not uuid:
        raise ValueError("falta Emisor, Receptor o timbre (UUID)")

    e_rfc, r_rfc = emisor.get("Rfc", "").upper(), receptor.get("Rfc", "").upper()
    if rfc_cliente == e_rfc:
        direccion = "Emitido"
    elif rfc_cliente == r_rfc:
        direccion = "Recibido"
    else:
        raise ValueError(f"no pertenece al cliente {rfc_cliente} (emisor {e_rfc}, receptor {r_rfc})")

    tipo = a.get("TipoDeComprobante", "").upper()
    try:
        fecha = datetime.fromisoformat(a["Fecha"]).date()
    except (KeyError, ValueError) as exc:
        raise ValueError(f"fecha invalida: {a.get('Fecha')!r}") from exc

    moneda = a.get("Moneda", "MXN").upper()
    nota = ""
    tc = 1.0
    if moneda not in ("MXN", "XXX"):
        tc = _num(a.get("TipoCambio"), 1.0)
        nota = f"Moneda {moneda} (TC {tc:g}); convertido a MXN"
    total = _num(a.get("Total")) * tc

    monto = total
    if tipo == "P":
        pagos = [e for e in root.iter() if _local(e.tag) == "Pago"]
        monto = sum(_num(p.get("Monto")) * (_num(p.get("TipoCambioP"), 1.0) if p.get("MonedaP", "MXN") != "MXN" else 1.0)
                    for p in pagos)
    elif tipo == "T":
        monto = 0.0

    conceptos = [c for c in root.iter() if _local(c.tag) == "Concepto"]
    descripciones = "; ".join(dict.fromkeys(c.get("Descripcion", "").strip() for c in conceptos[:4]))[:200]
    claves = ",".join(dict.fromkeys(c.get("ClaveProdServ", "") for c in conceptos[:4] if c.get("ClaveProdServ")))
    impuestos = hijos.get("Impuestos")
    iva = _num(impuestos.get("TotalImpuestosTrasladados")) * tc if impuestos is not None else 0.0

    cfdi = Cfdi(
        uuid=uuid, fecha=fecha, tipo=tipo, direccion=direccion,
        serie_folio=f"{a.get('Serie', '')}{a.get('Folio', '')}",
        emisor_rfc=e_rfc, emisor_nombre=emisor.get("Nombre", ""),
        receptor_rfc=r_rfc, receptor_nombre=receptor.get("Nombre", ""),
        uso_cfdi=receptor.get("UsoCFDI", ""), metodo_pago=a.get("MetodoPago", ""),
        forma_pago=a.get("FormaPago", ""), moneda=moneda,
        subtotal=_num(a.get("SubTotal")) * tc, iva=iva, total=total, monto=round(monto, 2),
        conceptos=descripciones, clave_prod=claves, archivo=path.name, nota=nota,
    )
    if (direccion, tipo) not in SENTIDO and tipo != "T":
        raise ValueError(f"tipo de comprobante no soportado: {tipo!r}")
    if cfdi.tipo == "I" and cfdi.metodo_pago == "PPD":
        cfdi.nota = (cfdi.nota + "; " if cfdi.nota else "") + "PPD: se concilia con su complemento de pago"
    return cfdi


def periodos(anio: int, mes: int, atras: int) -> list[tuple[int, int]]:
    """El mes pedido y `atras` meses anteriores (un pago puede cubrir facturas del mes pasado)."""
    return [((anio * 12 + mes - 1 - k) // 12, (anio * 12 + mes - 1 - k) % 12 + 1) for k in range(atras + 1)]


def cargar_cfdis(destino: Path, rfc: str, anio: int, mes: int, atras: int) -> list[Cfdi]:
    base = next(iter(sorted(destino.glob(f"{rfc} - *"))), destino / rfc)
    if not base.is_dir():
        raise SystemExit(f"No existe la carpeta del cliente en {destino} (se esperaba '{rfc} - <Nombre>').")
    cfdis: dict[str, Cfdi] = {}
    errores = 0
    for a, m in periodos(anio, mes, atras):
        for xml in sorted((base / f"{a}" / f"{m:02d}").rglob("*.xml")):
            try:
                c = parsear_cfdi(xml, rfc)
            except ValueError as exc:
                log.warning("%s: %s", xml.name, exc)
                errores += 1
                continue
            cfdis.setdefault(c.uuid, c)
    log.info("CFDI leidos: %s (omitidos por error: %s)", len(cfdis), errores)
    return [c for c in cfdis.values() if c.tipo != "T"]


# --------------------------------------------------------------------------- #
# Lectura del estado de cuenta
# --------------------------------------------------------------------------- #
def parsear_fecha(valor) -> date | None:
    txt = str(valor).strip().split(" ")[0].split("T")[0]
    for fmt in FORMATOS_FECHA:
        try:
            return datetime.strptime(txt, fmt).date()
        except ValueError:
            continue
    return None


def _columna(df: pd.DataFrame, clave: str) -> str | None:
    return next((c for c in df.columns if c in ALIAS[clave]), None)


def leer_movimientos(ruta: Path, cargos_positivos: bool = False) -> list[Mov]:
    """Lee el estado de cuenta; devuelve movimientos con monto firmado (negativo = cargo)."""
    df = ia.leer_tabla(ruta, requeridas=())
    col = {k: _columna(df, k) for k in ALIAS}
    if not (col["fecha"] and col["concepto"] and (col["monto"] or (col["cargo"] and col["abono"]))):
        raise ia.ClasificadorError("El estado de cuenta necesita: fecha, concepto y monto (o cargo y abono).")

    def valor(fila, clave):
        v = fila[col[clave]]
        return 0.0 if pd.isna(v) or not str(v).strip() else ia.parsear_monto(v)

    movs: list[Mov] = []
    for i, fila in df.iterrows():
        try:
            fecha = parsear_fecha(fila[col["fecha"]])
            if fecha is None:
                raise ia.ClasificadorError(f"fecha invalida {fila[col['fecha']]!r}")
            if col["monto"]:
                monto = valor(fila, "monto")
                monto = -monto if cargos_positivos else monto
            else:
                monto = valor(fila, "abono") - abs(valor(fila, "cargo"))
        except ia.ClasificadorError as exc:
            log.warning("Estado de cuenta fila %s omitida: %s", i + 2, exc)
            continue
        if monto == 0:
            continue
        movs.append(Mov(len(movs), i + 2, fecha, ia.sanear_concepto(fila[col["concepto"]]), round(monto, 2)))
    if not movs:
        raise ia.ClasificadorError("El estado de cuenta no tiene movimientos utilizables.")
    return movs


# --------------------------------------------------------------------------- #
# Conciliacion
# --------------------------------------------------------------------------- #
def _normalizar(txt: str) -> str:
    txt = unicodedata.normalize("NFKD", txt or "").encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Z0-9 ]", " ", txt.upper())


def contraparte_en_concepto(mov: Mov, cfdi: Cfdi) -> bool:
    """True si el RFC o una palabra distintiva del nombre de la contraparte aparece en el concepto."""
    rfc, nombre = cfdi.contraparte
    concepto = _normalizar(mov.concepto)
    if rfc and rfc in concepto.replace(" ", ""):
        return True
    return any(len(p) >= 4 and p not in PALABRAS_RUIDO and p in concepto for p in _normalizar(nombre).split())


def puntuar(mov: Mov, cfdi: Cfdi, dias: int, ventana: int) -> int:
    """Monto exacto = 50; cercania de fecha hasta 30; contraparte 20; folio 10 (tope 100)."""
    score = 50 + round(30 * (1 - max(dias, 0) / ventana))
    if contraparte_en_concepto(mov, cfdi):
        score += 20
    if len(cfdi.serie_folio) >= 3 and cfdi.serie_folio.upper() in _normalizar(mov.concepto).replace(" ", ""):
        score += 10
    return min(score, 100)


def _confianza(score: int) -> str:
    return "Alto" if score >= 85 else "Medio" if score >= 65 else "Bajo"


def _bajar(conf: str) -> str:
    return {"Alto": "Medio", "Medio": "Bajo"}.get(conf, "Bajo")


def conciliar(movs: list[Mov], cfdis: list[Cfdi], tolerancia: float = 0.01, ventana: int = 45) -> list[Par]:
    """
    Fase 1: un movimiento <-> un CFDI (monto, signo y fecha compatibles; asignacion voraz por puntaje).
    Fase 2: un movimiento <-> 2 o 3 CFDI del mismo proveedor/cliente cuya suma coincide.
    """
    candidatos = [c for c in cfdis if c.conciliable]
    por_uuid = {c.uuid: c for c in candidatos}
    pares_posibles = []
    for m in movs:
        signo = 1 if m.monto > 0 else -1
        for c in candidatos:
            if c.sentido != signo or abs(abs(m.monto) - c.monto) > tolerancia:
                continue
            dias = (m.fecha - c.fecha).days
            if -3 <= dias <= ventana:
                pares_posibles.append((puntuar(m, c, dias, ventana), -abs(dias), m.idx, c.uuid))
    pares_posibles.sort(reverse=True)

    por_mov: dict[int, list[tuple[int, str]]] = {}
    por_cfdi: dict[str, list[tuple[int, int]]] = {}
    for score, _, mi, cu in pares_posibles:
        por_mov.setdefault(mi, []).append((score, cu))
        por_cfdi.setdefault(cu, []).append((score, mi))

    resultado: list[Par] = []
    usados_m: set[int] = set()
    usados_c: set[str] = set()
    for score, _, mi, cu in pares_posibles:
        if mi in usados_m or cu in usados_c:
            continue
        usados_m.add(mi)
        usados_c.add(cu)
        # Ambiguo: otro candidato muy parecido compitio por el mismo movimiento o el mismo CFDI.
        rivales = [s for s, u in por_mov[mi] if u != cu] + [s for s, i in por_cfdi[cu] if i != mi]
        ambiguo = any(abs(score - s) <= 5 for s in rivales)
        conf = _bajar(_confianza(score)) if ambiguo else _confianza(score)
        criterio = "monto exacto" + (", varios candidatos parecidos" if ambiguo else "")
        resultado.append(Par(movs[mi], [por_uuid[cu]], score, conf, criterio))

    resultado += _conciliar_sumas(movs, candidatos, usados_m, usados_c, tolerancia, ventana)
    return sorted(resultado, key=lambda p: p.mov.idx)


def _conciliar_sumas(movs, candidatos, usados_m, usados_c, tolerancia, ventana) -> list[Par]:
    """Pagos que cubren varias facturas de la misma contraparte en un solo movimiento."""
    grupos: dict[tuple, list[Cfdi]] = {}
    for c in candidatos:
        if c.uuid not in usados_c:
            grupos.setdefault((c.contraparte[0], c.sentido), []).append(c)

    resultado = []
    for m in movs:
        if m.idx in usados_m:
            continue
        signo = 1 if m.monto > 0 else -1
        objetivo = abs(m.monto)
        encontrados = []
        for (_, sentido), lista in grupos.items():
            if sentido != signo:
                continue
            cerca = sorted((c for c in lista if c.uuid not in usados_c and -3 <= (m.fecha - c.fecha).days <= ventana),
                           key=lambda c: abs((m.fecha - c.fecha).days))[:15]
            for n in (2, 3):
                for combo in itertools.combinations(cerca, n):
                    if abs(sum(c.monto for c in combo) - objetivo) <= tolerancia * n:
                        encontrados.append(combo)
        if not encontrados:
            continue
        combo = encontrados[0]
        con_nombre = contraparte_en_concepto(m, combo[0])
        conf = "Medio" if con_nombre and len(encontrados) == 1 else "Bajo"
        criterio = f"suma de {len(combo)} CFDI de la misma contraparte" + (
            f" ({len(encontrados)} combinaciones posibles)" if len(encontrados) > 1 else "")
        usados_m.add(m.idx)
        usados_c.update(c.uuid for c in combo)
        resultado.append(Par(m, list(combo), 0, conf, criterio))
    return resultado


# --------------------------------------------------------------------------- #
# Clasificacion con IA
# --------------------------------------------------------------------------- #
class Clasificador:
    """Envuelve auto_clasificador: cache, corte ante fallas repetidas y degradacion sin romper el reporte."""

    def __init__(self, proveedor) -> None:
        self.proveedor = proveedor
        self.cache: dict = {}
        self.fallas = 0
        self.activo = proveedor is not None

    def __call__(self, clave, concepto: str, monto: float) -> dict:
        vacio = {c: "" for c in ia.CAMPOS_SALIDA}
        if not self.activo:
            return vacio
        if clave in self.cache:
            return self.cache[clave]
        try:
            res = ia.clasificar_gasto(self.proveedor, concepto, monto)
            self.fallas = 0
        except ia.ErrorFatal as exc:
            log.error("Clasificacion detenida: %s", exc)
            self.activo = False
            return vacio
        except ia.ClasificadorError as exc:
            log.error("No se clasifico %r: %s", concepto[:40], exc)
            self.fallas += 1
            if self.fallas >= ia.MAX_FALLAS_CONSECUTIVAS:
                log.error("Demasiadas fallas seguidas; se desactiva la clasificacion.")
                self.activo = False
            res = {**vacio, "justificacion_fiscal": f"ERROR: {exc}"}
        self.cache[clave] = res
        return res


def clasificar_cfdi(c: Cfdi, clasificar: Clasificador) -> None:
    """Solo gastos: CFDI recibidos de tipo I (facturas) o E (notas de credito)."""
    if c.direccion != "Recibido" or c.tipo not in ("I", "E"):
        return
    texto = (f"CFDI recibido de {c.emisor_nombre} (RFC {c.emisor_rfc}). Conceptos: {c.conceptos}. "
             f"Uso CFDI: {c.uso_cfdi}. Producto SAT: {c.clave_prod}")
    c.clasificacion = clasificar((c.emisor_rfc, c.conceptos[:80], c.uso_cfdi), texto, c.total)


# --------------------------------------------------------------------------- #
# Reporte
# --------------------------------------------------------------------------- #
def _cls(d: dict) -> dict:
    return {
        "Cuenta sugerida": d.get("cuenta_sugerida", ""), "Codigo": d.get("codigo_cuenta", ""),
        "Confianza IA": d.get("nivel_confianza", ""), "Justificacion fiscal": d.get("justificacion_fiscal", ""),
    }


def construir_hojas(movs: list[Mov], cfdis: list[Cfdi], pares: list[Par], rfc: str, periodo: str) -> dict:
    conciliados = []
    for p in pares:
        for c in p.cfdis:
            conciliados.append({
                "Fecha mov": p.mov.fecha, "Concepto banco": p.mov.concepto, "Monto banco": p.mov.monto,
                "UUID": c.uuid, "Tipo": c.tipo, "Contraparte": c.contraparte[1], "RFC contraparte": c.contraparte[0],
                "Serie/Folio": c.serie_folio, "Fecha CFDI": c.fecha, "Monto CFDI": c.monto,
                "Dias": (p.mov.fecha - c.fecha).days, "Confianza": p.confianza, "Criterio": p.criterio,
                "Fila banco": p.mov.fila,
            })
    uuids_ok = {c.uuid for p in pares for c in p.cfdis}
    idx_ok = {p.mov.idx for p in pares}

    sin_cfdi = [{"Fila banco": m.fila, "Fecha": m.fecha, "Concepto": m.concepto, "Monto": m.monto, **_cls(m.clasificacion)}
                for m in movs if m.idx not in idx_ok]
    sin_mov = [{"Fecha": c.fecha, "UUID": c.uuid, "Tipo": c.tipo, "Direccion": c.direccion,
                "Contraparte": c.contraparte[1], "Monto": c.monto, "Metodo pago": c.metodo_pago, "Nota": c.nota}
               for c in cfdis if c.uuid not in uuids_ok]
    detalle = [{"Fecha": c.fecha, "UUID": c.uuid, "Tipo": c.tipo, "Direccion": c.direccion, "Serie/Folio": c.serie_folio,
                "RFC emisor": c.emisor_rfc, "Emisor": c.emisor_nombre, "RFC receptor": c.receptor_rfc,
                "Receptor": c.receptor_nombre, "Uso CFDI": c.uso_cfdi, "Metodo pago": c.metodo_pago,
                "Forma pago": c.forma_pago, "Moneda": c.moneda, "Subtotal": c.subtotal, "IVA": c.iva, "Total": c.total,
                "Conceptos": c.conceptos, "Conciliado": "Si" if c.uuid in uuids_ok else "No", "Nota": c.nota,
                **_cls(c.clasificacion)} for c in cfdis]

    conciliables = [c for c in cfdis if c.conciliable]
    cargos_sin = sum(-m["Monto"] for m in sin_cfdi if m["Monto"] < 0)
    resumen = [
        ("Cliente (RFC)", rfc), ("Periodo", periodo), ("CFDI leidos", len(cfdis)),
        ("Movimientos bancarios", len(movs)),
        ("Movimientos conciliados", f"{len(idx_ok)} ({len(idx_ok) / max(len(movs), 1):.0%})" if movs else 0),
        ("Movimientos sin CFDI", len(sin_cfdi)), ("Cargos sin CFDI (importe)", round(cargos_sin, 2)),
        ("CFDI conciliables", len(conciliables)),
        ("CFDI conciliables sin movimiento", sum(1 for c in conciliables if c.uuid not in uuids_ok)),
        ("Conciliaciones de confianza Baja (revisar)", sum(1 for p in pares if p.confianza == "Bajo")),
    ]
    if not movs:  # modo solo-CFDI: no hay nada que conciliar
        resumen = resumen[:3] + [("Modo", "Solo CFDI (sin estado de cuenta)"),
                                 ("CFDI clasificados con IA", sum(1 for c in cfdis if c.clasificacion))]
    return {
        "Resumen": pd.DataFrame(resumen, columns=["Concepto", "Valor"]),
        "Conciliados": pd.DataFrame(conciliados), "Mov sin CFDI": pd.DataFrame(sin_cfdi),
        "CFDI sin mov": pd.DataFrame(sin_mov), "CFDI": pd.DataFrame(detalle),
    }


def escribir_excel(hojas: dict, salida: Path) -> None:
    from openpyxl.styles import Font
    with pd.ExcelWriter(salida, engine="openpyxl") as xl:
        for nombre, df in hojas.items():
            df.to_excel(xl, sheet_name=nombre, index=False)
            ws = xl.sheets[nombre]
            ws.freeze_panes = "A2"
            for celda in ws[1]:
                celda.font = Font(bold=True)
            for col in ws.columns:
                ancho = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(ancho + 2, 60)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Conciliacion bancaria contra CFDI + clasificacion contable.")
    ap.add_argument("--rfc", required=True, help="RFC del cliente")
    ap.add_argument("--anio", type=int, required=True)
    ap.add_argument("--mes", type=int, required=True, choices=range(1, 13), metavar="{1-12}")
    ap.add_argument("--estado-cuenta", help="CSV/Excel del banco (sin el solo se leen/clasifican los CFDI)")
    ap.add_argument("--destino", default=DESTINO_DEFAULT, help="carpeta raiz donde cfdi_sat.py guardo los XML")
    ap.add_argument("--provider", choices=["azure", "gcp"], help="activa la clasificacion con IA")
    ap.add_argument("--salida", help="ruta del Excel (default: conciliacion_<RFC>_<AAAA>-<MM>.xlsx)")
    ap.add_argument("--tolerancia", type=float, default=0.01, help="diferencia maxima de monto en MXN")
    ap.add_argument("--ventana-dias", type=int, default=45, help="dias maximos entre CFDI y pago")
    ap.add_argument("--meses-atras", type=int, default=1, help="meses previos de CFDI a considerar")
    ap.add_argument("--cargos-positivos", action="store_true", help="en la columna monto los cargos vienen positivos")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    rfc = args.rfc.strip().upper()
    if not RE_RFC.match(rfc):
        ap.error(f"RFC invalido: {rfc}")

    try:
        cfdis = cargar_cfdis(Path(args.destino), rfc, args.anio, args.mes, args.meses_atras)
        if not cfdis:
            raise ia.ClasificadorError("No se encontraron CFDI para ese cliente/periodo. ¿Ya corriste cfdi_sat.py?")

        clasificar = Clasificador(ia.crear_proveedor(args.provider) if args.provider else None)
        movs, pares = [], []
        if args.estado_cuenta:
            todos = leer_movimientos(Path(args.estado_cuenta), args.cargos_positivos)
            movs = [m for m in todos if (m.fecha.year, m.fecha.month) == (args.anio, args.mes)]
            if len(movs) < len(todos):
                log.warning("%s movimientos fuera de %s-%02d se ignoraron", len(todos) - len(movs), args.anio, args.mes)
            pares = conciliar(movs, cfdis, args.tolerancia, args.ventana_dias)
            ok = {p.mov.idx for p in pares}
            for m in movs:
                if m.idx not in ok and m.monto < 0:
                    m.clasificacion = clasificar(("mov", m.concepto.lower()), m.concepto, abs(m.monto))

        for c in cfdis:
            clasificar_cfdi(c, clasificar)

        hojas = construir_hojas(movs, cfdis, pares, rfc, f"{args.anio}-{args.mes:02d}")
        salida = Path(args.salida or f"conciliacion_{rfc}_{args.anio}-{args.mes:02d}.xlsx")
        escribir_excel(hojas, salida)
    except ia.ClasificadorError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    except PermissionError:
        print("ERROR: no se pudo escribir el Excel (¿esta abierto?). Cierralo e intenta de nuevo.", file=sys.stderr)
        return 1

    print(hojas["Resumen"].to_string(index=False, header=False))
    print(f"\nReporte: {salida}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
