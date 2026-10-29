"""Descarga de CFDI (emitidos y recibidos) del portal del SAT con CIEC.

Guarda los XML en  <DESTINO>/<RFC - Nombre>/<AAAA>/<MM>/<Emitidos|Recibidos>/<UUID>.xml
Si DESTINO es la carpeta de Drive for Desktop, la sincronizacion a Drive es automatica.

Uso:
    python cfdi_sat.py --anio 2026 --mes 9              # todos los clientes
    python cfdi_sat.py --anio 2026 --mes 9 --rfc XAXX010101000
    python cfdi_sat.py --anio 2026 --mes 9 --tipo recibidos
"""
import argparse
import calendar
import csv
import logging
import sys
from pathlib import Path

import keyring
from playwright.sync_api import Page, TimeoutError as PWTimeout, sync_playwright

BASE = Path(__file__).parent
DESTINO = Path(r"G:\Mi unidad\CFDI")  # <-- cambia a tu carpeta de Drive for Desktop
CLIENTES_CSV = BASE / "clientes.csv"
SERVICIO_KEYRING = "cfdi_sat"
LIMITE_SAT = 500  # el portal corta los resultados a 500 por busqueda

URL_LOGIN = "https://cfdiau.sat.gob.mx/nidp/app/login?id=SATUPCFDiCon&sid=0&option=credential&sid=0"
URL_RECIBIDOS = "https://portalcfdi.facturaelectronica.sat.gob.mx/ConsultaReceptor.aspx"
URL_EMITIDOS = "https://portalcfdi.facturaelectronica.sat.gob.mx/ConsultaEmisor.aspx"

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


def login(page: Page, rfc: str, password: str) -> None:
    for intento in range(3):
        page.goto(URL_LOGIN)
        page.fill(SEL["rfc"], rfc)
        page.fill(SEL["password"], password)
        img = BASE / "captcha.png"
        page.locator(SEL["captcha_img"]).screenshot(path=str(img))
        print(f"\n[{rfc}] Captcha guardado en {img}")
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


def buscar(page: Page, tipo: str, anio: int, mes: int) -> None:
    if tipo == "recibidos":
        page.goto(URL_RECIBIDOS)
        page.check(SEL["radio_fechas"])
        page.select_option(SEL["anio"], str(anio))
        page.select_option(SEL["mes"], str(mes))
        page.select_option(SEL["dia"], "0")  # 0 = todo el mes
    else:
        page.goto(URL_EMITIDOS)
        page.check(SEL["radio_fechas"])
        ultimo = calendar.monthrange(anio, mes)[1]
        page.fill(SEL["emi_ini"], f"01/{mes:02d}/{anio}")
        page.fill(SEL["emi_fin"], f"{ultimo:02d}/{mes:02d}/{anio}")
    page.click(SEL["buscar"])
    page.wait_for_load_state("networkidle")


def descargar(page: Page, tipo: str, carpeta: Path) -> tuple[int, int]:
    carpeta.mkdir(parents=True, exist_ok=True)
    filas = page.locator(SEL["filas"])
    total = filas.count()
    if total >= LIMITE_SAT:
        log.warning("%s: %s resultados, el SAT trunca a %s. Dividir la busqueda por dia.", tipo, total, LIMITE_SAT)
    nuevos = 0
    for i in range(total):
        fila = filas.nth(i)
        uuid = fila.locator(SEL["uuid"]).first.inner_text().strip().upper()
        destino = carpeta / f"{uuid}.xml"
        if destino.exists():
            continue
        with page.expect_download(timeout=30000) as dl:
            fila.locator(SEL["descarga"]).first.click()
        dl.value.save_as(destino)
        nuevos += 1
    return total, nuevos


def procesar_cliente(browser, rfc: str, nombre: str, anio: int, mes: int, tipos: list[str]) -> None:
    password = keyring.get_password(SERVICIO_KEYRING, rfc)
    if not password:
        log.error("%s: sin CIEC guardada (usa guardar_credencial.py)", rfc)
        return
    ctx = browser.new_context(accept_downloads=True)
    page = ctx.new_page()
    try:
        login(page, rfc, password)
        for tipo in tipos:
            buscar(page, tipo, anio, mes)
            carpeta = DESTINO / f"{rfc} - {nombre}" / f"{anio}" / f"{mes:02d}" / tipo.capitalize()
            total, nuevos = descargar(page, tipo, carpeta)
            log.info("%s %s %s-%02d: %s encontrados, %s nuevos", rfc, tipo, anio, mes, total, nuevos)
    except Exception:
        log.exception("%s: error procesando cliente", rfc)
    finally:
        ctx.close()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--anio", type=int, required=True)
    ap.add_argument("--mes", type=int, required=True)
    ap.add_argument("--rfc", help="procesar solo este RFC")
    ap.add_argument("--tipo", choices=["emitidos", "recibidos", "ambos"], default="ambos")
    args = ap.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(BASE / "cfdi.log", encoding="utf-8"), logging.StreamHandler(sys.stdout)],
    )
    tipos = ["emitidos", "recibidos"] if args.tipo == "ambos" else [args.tipo]

    with open(CLIENTES_CSV, newline="", encoding="utf-8") as f:
        clientes = [c for c in csv.DictReader(f) if not args.rfc or c["rfc"] == args.rfc]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # visible: necesitas ver el captcha
        for c in clientes:
            procesar_cliente(browser, c["rfc"], c["nombre"], args.anio, args.mes, tipos)
        browser.close()


if __name__ == "__main__":
    main()
