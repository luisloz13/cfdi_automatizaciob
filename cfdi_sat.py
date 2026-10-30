"""Descarga de CFDI (emitidos y recibidos) del portal del SAT con CIEC.

Guarda los XML en  <DESTINO>/<RFC - Nombre>/<AAAA>/<MM>/<Emitidos|Recibidos>/<UUID>.xml
Si DESTINO es la carpeta de Drive for Desktop, la sincronizacion a Drive es automatica.

Uso:
    python cfdi_sat.py --anio 2026 --mes 9                  # todos los clientes
    python cfdi_sat.py --anio 2026 --mes 9 --rfc XAXX010101000
    python cfdi_sat.py --anio 2026 --mes 9 --tipo recibidos
    python cfdi_sat.py --anio 2026 --mes 9 --destino D:\\CFDI

El destino tambien puede fijarse con la variable de entorno CFDI_DESTINO.
Codigo de salida: 0 = todo ok, 1 = hubo clientes o descargas con error.
"""
import argparse
import calendar
import csv
import logging
import os
import re
import sys
from dataclasses import dataclass
from logging.handlers import RotatingFileHandler
from pathlib import Path
from xml.etree import ElementTree

import keyring
from playwright.sync_api import Page, TimeoutError as PWTimeout, sync_playwright

BASE = Path(__file__).parent
DESTINO_DEFAULT = r"G:\Mi unidad\CFDI"  # carpeta de Drive for Desktop
CLIENTES_CSV = BASE / "clientes.csv"
ERRORES_DIR = BASE / "errores"
SERVICIO_KEYRING = "cfdi_sat"
LIMITE_SAT = 500  # el portal corta los resultados a 500 por busqueda
INTENTOS_DESCARGA = 3

URL_LOGIN = "https://cfdiau.sat.gob.mx/nidp/app/login?id=SATUPCFDiCon&sid=0&option=credential&sid=0"
URL_RECIBIDOS = "https://portalcfdi.facturaelectronica.sat.gob.mx/ConsultaReceptor.aspx"
URL_EMITIDOS = "https://portalcfdi.facturaelectronica.sat.gob.mx/ConsultaEmisor.aspx"

RE_RFC = re.compile(r"^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$")
RE_UUID = re.compile(r"^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}$")

# Selectores del portal. Si el SAT cambia la pagina, solo se ajusta aqui.
SEL = {
    "rfc": "#rfc",
    "password": "#password",
    "captcha_img": "#divCaptcha img",
    "captcha_input": "#userCaptcha",
    "submit": "#submit",
    "radio_fechas": "#ctl00_MainContent_RdoFechas",
    "anio": "#DdlAnio",
    "mes": "#ctl00_MainContent_CldFecha_DdlMes",
    "dia": "#ctl00_MainContent_CldFecha_DdlDia",
    "emi_ini": "#ctl00_MainContent_CldFechaInicial2_Calendario_text",
    "emi_fin": "#ctl00_MainContent_CldFechaFinal2_Calendario_text",
    "buscar": "#ctl00_MainContent_BtnBusqueda",
    "filas": "#ctl00_MainContent_tblResult tr:has(span#BtnDescarga), #ctl00_MainContent_tblResult tr:has(#BtnDescarga)",
    "descarga": "#BtnDescarga",
    "uuid": "td:nth-child(2) span",
}

log = logging.getLogger("cfdi")


@dataclass
class Resumen:
    """Contadores de una corrida; se pueden sumar entre periodos/tipos/clientes."""
    encontrados: int = 0
    nuevos: int = 0
    errores: int = 0

    def __iadd__(self, otro: "Resumen") -> "Resumen":
        self.encontrados += otro.encontrados
        self.nuevos += otro.nuevos
        self.errores += otro.errores
        return self


# --------------------------------------------------------------------------- #
# Utilidades puras (sin navegador)
# --------------------------------------------------------------------------- #
def limpiar_nombre(nombre: str) -> str:
    """Quita caracteres que Windows no permite en nombres de carpeta."""
    limpio = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", nombre).strip(" .")
    return limpio or "SIN_NOMBRE"


def es_xml_valido(ruta: Path) -> bool:
    """El SAT a veces devuelve una pagina de error con status 200; descarta lo que no sea XML."""
    try:
        ElementTree.parse(ruta)
        return True
    except (ElementTree.ParseError, OSError):
        return False


def cargar_clientes(ruta: Path, solo_rfc: str | None) -> list[dict]:
    if not ruta.is_file():
        raise SystemExit(f"No existe {ruta}. Debe tener las columnas: rfc,nombre")
    with open(ruta, newline="", encoding="utf-8-sig") as f:
        filas = list(csv.DictReader(f))
    if filas and not {"rfc", "nombre"} <= set(filas[0]):
        raise SystemExit(f"{ruta.name} debe tener las columnas: rfc,nombre")
    clientes, vistos = [], set()
    for n, fila in enumerate(filas, start=2):
        rfc = (fila.get("rfc") or "").strip().upper()
        if not RE_RFC.match(rfc):
            log.error("%s fila %s: RFC invalido %r, se omite", ruta.name, n, rfc)
            continue
        if rfc in vistos or (solo_rfc and rfc != solo_rfc):
            continue
        vistos.add(rfc)
        clientes.append({"rfc": rfc, "nombre": (fila.get("nombre") or "").strip()})
    return clientes


# --------------------------------------------------------------------------- #
# Portal SAT
# --------------------------------------------------------------------------- #
def login(page: Page, rfc: str, password: str) -> None:
    captcha = BASE / "captcha.png"
    try:
        for intento in range(3):
            page.goto(URL_LOGIN)
            page.fill(SEL["rfc"], rfc)
            page.fill(SEL["password"], password)
            page.locator(SEL["captcha_img"]).screenshot(path=str(captcha))
            print(f"\n[{rfc}] Captcha guardado en {captcha}")
            texto = input("Escribe el captcha (o Enter para reintentar): ").strip()
            if not texto:
                continue
            page.fill(SEL["captcha_input"], texto)
            page.click(SEL["submit"])
            try:
                page.wait_for_url("**/portalcfdi.facturaelectronica.sat.gob.mx/**", timeout=20000)
                return
            except PWTimeout:
                log.warning("%s: login fallido (intento %s), revisa captcha/CIEC", rfc, intento + 1)
        raise RuntimeError(f"No se pudo iniciar sesion para {rfc}")
    finally:
        captcha.unlink(missing_ok=True)


def buscar(page: Page, tipo: str, anio: int, mes: int, dia: int | None = None) -> None:
    """Busca el mes completo, o un solo dia si se indica `dia`."""
    if tipo == "recibidos":
        page.goto(URL_RECIBIDOS)
        page.check(SEL["radio_fechas"])
        page.select_option(SEL["anio"], str(anio))
        page.select_option(SEL["mes"], str(mes))
        page.select_option(SEL["dia"], str(dia or 0))  # 0 = todo el mes
    else:
        page.goto(URL_EMITIDOS)
        page.check(SEL["radio_fechas"])
        ultimo = calendar.monthrange(anio, mes)[1]
        page.fill(SEL["emi_ini"], f"{dia or 1:02d}/{mes:02d}/{anio}")
        page.fill(SEL["emi_fin"], f"{dia or ultimo:02d}/{mes:02d}/{anio}")
    page.click(SEL["buscar"])
    page.wait_for_load_state("networkidle")


def descargar(page: Page, carpeta: Path) -> Resumen:
    """Descarga los XML de la pagina de resultados actual. Un fallo no aborta el resto."""
    carpeta.mkdir(parents=True, exist_ok=True)
    filas = page.locator(SEL["filas"])
    res = Resumen(encontrados=filas.count())
    for i in range(res.encontrados):
        fila = filas.nth(i)
        uuid = fila.locator(SEL["uuid"]).first.inner_text().strip().upper()
        if not RE_UUID.match(uuid):
            log.error("Fila %s: UUID con formato inesperado %r, se omite", i + 1, uuid)
            res.errores += 1
            continue
        destino = carpeta / f"{uuid}.xml"
        if destino.exists():
            continue
        if _descargar_uno(page, fila, destino):
            res.nuevos += 1
        else:
            res.errores += 1
    return res


def _descargar_uno(page: Page, fila, destino: Path) -> bool:
    """Descarga a un .part, valida que sea XML y solo entonces lo renombra (nunca queda un XML a medias)."""
    parcial = destino.with_suffix(".part")
    for intento in range(1, INTENTOS_DESCARGA + 1):
        try:
            with page.expect_download(timeout=30000) as dl:
                fila.locator(SEL["descarga"]).first.click()
            dl.value.save_as(parcial)
            if not es_xml_valido(parcial):
                raise ValueError("el archivo descargado no es un XML valido")
            os.replace(parcial, destino)
            return True
        except Exception as exc:
            log.warning("%s: intento %s/%s fallido (%s)", destino.name, intento, INTENTOS_DESCARGA, exc)
            parcial.unlink(missing_ok=True)
    log.error("%s: no se pudo descargar", destino.name)
    return False


def descargar_periodo(page: Page, tipo: str, anio: int, mes: int, carpeta: Path) -> Resumen:
    """Descarga el mes; si el SAT lo trunca (500 resultados) repite la busqueda dia por dia."""
    buscar(page, tipo, anio, mes)
    if page.locator(SEL["filas"]).count() < LIMITE_SAT:
        return descargar(page, carpeta)
    log.warning("%s: %s-%02d llego al limite de %s; se divide por dia", tipo, anio, mes, LIMITE_SAT)
    total = Resumen()
    for dia in range(1, calendar.monthrange(anio, mes)[1] + 1):
        buscar(page, tipo, anio, mes, dia)
        parcial = descargar(page, carpeta)
        if parcial.encontrados >= LIMITE_SAT:
            log.error("%s: el dia %02d/%02d/%s sigue al limite; pueden faltar CFDI", tipo, dia, mes, anio)
            parcial.errores += 1
        total += parcial
    return total


def procesar_cliente(browser, rfc: str, nombre: str, anio: int, mes: int, tipos: list[str],
                     destino: Path) -> Resumen | None:
    """Devuelve el resumen, o None si el cliente fallo por completo (sin CIEC, login, etc.)."""
    password = keyring.get_password(SERVICIO_KEYRING, rfc)
    if not password:
        log.error("%s: sin CIEC guardada (usa guardar_credencial.py)", rfc)
        return None
    ctx = browser.new_context(accept_downloads=True)
    page = ctx.new_page()
    total = Resumen()
    try:
        login(page, rfc, password)
        for tipo in tipos:
            carpeta = destino / f"{rfc} - {limpiar_nombre(nombre)}" / f"{anio}" / f"{mes:02d}" / tipo.capitalize()
            res = descargar_periodo(page, tipo, anio, mes, carpeta)
            total += res
            log.info("%s %s %s-%02d: %s encontrados, %s nuevos, %s con error",
                     rfc, tipo, anio, mes, res.encontrados, res.nuevos, res.errores)
        return total
    except Exception:
        log.exception("%s: error procesando cliente", rfc)
        _captura_error(page, rfc)
        return None
    finally:
        ctx.close()


def _captura_error(page: Page, rfc: str) -> None:
    """Deja una captura de pantalla para diagnosticar cambios en el portal."""
    try:
        ERRORES_DIR.mkdir(exist_ok=True)
        ruta = ERRORES_DIR / f"{rfc}.png"
        page.screenshot(path=str(ruta))
        log.info("Captura del error: %s", ruta)
    except Exception:
        pass  # la captura es solo ayuda; no debe tapar el error original


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main() -> int:
    ap = argparse.ArgumentParser(description="Descarga masiva de CFDI del portal del SAT.")
    ap.add_argument("--anio", type=int, required=True)
    ap.add_argument("--mes", type=int, required=True, choices=range(1, 13), metavar="{1-12}")
    ap.add_argument("--rfc", help="procesar solo este RFC")
    ap.add_argument("--tipo", choices=["emitidos", "recibidos", "ambos"], default="ambos")
    ap.add_argument("--destino", default=os.environ.get("CFDI_DESTINO", DESTINO_DEFAULT),
                    help="carpeta raiz de destino (default: $CFDI_DESTINO o %(default)s)")
    ap.add_argument("--clientes", default=str(CLIENTES_CSV), help="CSV con columnas rfc,nombre")
    args = ap.parse_args()
    if not 2014 <= args.anio <= 2100:
        ap.error("--anio fuera de rango")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            RotatingFileHandler(BASE / "cfdi.log", maxBytes=2_000_000, backupCount=3, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    tipos = ["emitidos", "recibidos"] if args.tipo == "ambos" else [args.tipo]
    destino = Path(args.destino)
    if not destino.parent.exists():
        log.error("La ruta de destino %s no es accesible (¿Drive for Desktop esta abierto?)", destino)
        return 1

    solo_rfc = args.rfc.strip().upper() if args.rfc else None
    clientes = cargar_clientes(Path(args.clientes), solo_rfc)
    if not clientes:
        log.error("No hay clientes para procesar.")
        return 1

    total, fallidos = Resumen(), []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # visible: necesitas ver el captcha
        try:
            for c in clientes:
                res = procesar_cliente(browser, c["rfc"], c["nombre"], args.anio, args.mes, tipos, destino)
                if res is None:
                    fallidos.append(c["rfc"])
                else:
                    total += res
        except KeyboardInterrupt:
            log.warning("Interrumpido por el usuario.")
            fallidos.append("(interrumpido)")
        finally:
            browser.close()

    log.info("RESUMEN %s-%02d: %s clientes, %s CFDI encontrados, %s nuevos, %s descargas con error, "
             "clientes fallidos: %s", args.anio, args.mes, len(clientes), total.encontrados,
             total.nuevos, total.errores, ", ".join(fallidos) or "ninguno")
    return 1 if (fallidos or total.errores) else 0


if __name__ == "__main__":
    sys.exit(main())
