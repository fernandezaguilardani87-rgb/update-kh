#!/usr/bin/env python3
"""
Descargador de listados de precios — Portal MxM (Más por Menos)
https://intranet.inmobiliariamaspormenos.com/

Pasos:
  1. Login con usuario + contraseña
  2. En /inmuebles/, recoger todas las promociones (foto → URL detalle)
  3. Por cada promoción: entrar, ir a pestaña Documents, descargar listado de precios
  4. Guardar como Price_List_{fecha}_{nombre}.pdf

Uso:
    python descargar_mxm.py             — descarga todo
    python descargar_mxm.py --explorar  — solo lista las promociones, sin descargar
"""

import sys, os, re, time, argparse
from pathlib import Path
from datetime import datetime

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

URL_BASE  = "https://intranet.inmobiliariamaspormenos.com"
URL_LOGIN = URL_BASE + "/"
URL_LIST  = URL_BASE + "/inmuebles/"

USUARIO  = "Admin"
PASSWORD = "MxMON"

CARPETA_DESTINO = Path(r"C:\Users\User\Docs\Update\Update KH")

# Mapeo nombre portal → nombre Excel KH (para que los PDFs encajen con property_manager)
MAPEO_NOMBRES = {
    "spinto blu":                          "Spinto Blu",
    "skye casares golf":                   "Skye",
    "idyllic":                             "Ydillic",
    "camarate hills ii":                   "Camarate Hills",
    "eleven41":                            "Eleven 41",
    "pure sun ii":                         "PureSun II",
    "natura estepona":                     "Natura",
    # lagumare41 / local comercial / sanctuary → confirmar con Dan
}

# Palabras clave para reconocer un listado de precios
KEYWORDS_PRECIO = [
    'price list', 'pricelist', 'price-list', 'price_list',
    'lista de precio', 'lista precios', 'listado precio',
    'tarifa', 'precios',
]

def es_precio(texto: str) -> bool:
    t = texto.lower().replace('_', ' ').replace('-', ' ')
    return any(k.replace('_', ' ').replace('-', ' ') in t for k in KEYWORDS_PRECIO)

# ─── PLAYWRIGHT ───────────────────────────────────────────────────────────────

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    os.system(f'"{sys.executable}" -m pip install playwright --quiet')
    os.system(f'"{sys.executable}" -m playwright install chromium')
    from playwright.sync_api import sync_playwright

def ts():
    return datetime.now().strftime("%H:%M:%S")

def log(msg):
    print(f"[{ts()}] {msg}")

def limpiar(txt):
    return "".join(c if c.isalnum() or c in " -_." else "_" for c in txt).strip()

# ─── PASO 1: LOGIN ────────────────────────────────────────────────────────────

def login(page):
    log(f"Abriendo {URL_LOGIN}")
    page.goto(URL_LOGIN, wait_until="networkidle", timeout=30_000)
    time.sleep(1)

    page.locator('input[type="text"]').first.fill(USUARIO)
    page.locator('input[type="password"]').first.fill(PASSWORD)
    page.locator('input[type="submit"]').first.click()
    page.wait_for_load_state("networkidle", timeout=20_000)
    time.sleep(1.5)

    if "/inmuebles" in page.url:
        log(f"  Login OK → {page.url}")
        return True
    log(f"  Login fallido — URL: {page.url}")
    return False

# ─── PASO 2: RECOGER PROMOCIONES ─────────────────────────────────────────────

def recoger_promociones(page):
    """
    Devuelve lista de dicts {nombre, url} leyendo los div.row de /inmuebles/.
    Cada fila tiene: imagen (enlace a detalle) + span.dblock (nombre).
    """
    if "/inmuebles" not in page.url:
        page.goto(URL_LIST, wait_until="networkidle", timeout=20_000)
        time.sleep(1.5)

    promos = page.evaluate("""
        () => {
            const result = [];
            for (const row of document.querySelectorAll('div.row')) {
                const nameEl = row.querySelector('span.dblock');
                const linkEl = row.querySelector('a[href^="/p-"]');
                if (nameEl && linkEl) {
                    result.push({
                        nombre: nameEl.textContent.trim(),
                        url:    linkEl.href,
                    });
                }
            }
            return result;
        }
    """)

    # Paginación
    pagina = 1
    while True:
        sig = page.evaluate("""
            () => {
                const a = document.querySelector('a[rel="next"], a.next, a[aria-label="Next"]');
                return a ? a.href : null;
            }
        """)
        if not sig:
            break
        pagina += 1
        log(f"  Paginando → página {pagina}")
        page.goto(sig, wait_until="networkidle", timeout=15_000)
        time.sleep(1)
        mas = page.evaluate("""
            () => {
                const result = [];
                for (const row of document.querySelectorAll('div.row')) {
                    const nameEl = row.querySelector('span.dblock');
                    const linkEl = row.querySelector('a[href^="/p-"]');
                    if (nameEl && linkEl) {
                        result.push({ nombre: nameEl.textContent.trim(), url: linkEl.href });
                    }
                }
                return result;
            }
        """)
        if not mas:
            break
        promos.extend(mas)

    # Deduplicar por URL
    seen = set()
    unicos = []
    for p in promos:
        if p['url'] not in seen:
            seen.add(p['url'])
            unicos.append(p)

    return unicos

# ─── PASO 3: PROCESAR UNA PROMOCIÓN ──────────────────────────────────────────

def procesar(page, context, promo, fecha, descargados, errores, sin_doc):
    nombre_portal = promo['nombre']
    url_detalle   = promo['url']
    nombre_kh     = MAPEO_NOMBRES.get(nombre_portal.lower().strip(), nombre_portal)

    log(f"\n{'─'*55}")
    log(f"  {nombre_portal}  →  {url_detalle}")

    try:
        # ── Entrar a la página de detalle (clic en la foto = misma URL) ───────
        page.goto(url_detalle, wait_until="networkidle", timeout=20_000)
        time.sleep(1.5)
        log(f"  Detalle cargado: {page.url}")

        # ── Buscar y pulsar la pestaña Documents ──────────────────────────────
        tab_ok = False
        for texto in ["Documents", "Documentos", "Docs", "Files", "Archivos"]:
            loc = page.get_by_role("link", name=texto).or_(
                  page.get_by_role("tab",  name=texto)).or_(
                  page.get_by_role("button", name=texto))
            if loc.count() > 0:
                loc.first.click()
                time.sleep(1.5)
                tab_ok = True
                log(f"  Pestaña '{texto}' abierta ✓")
                break

        # Si no hay pestaña con esos textos, buscar por texto parcial en cualquier enlace
        if not tab_ok:
            for texto in ["ocument", "rchivo", "ile"]:
                loc = page.locator(f'a:has-text("{texto}"), button:has-text("{texto}")')
                if loc.count() > 0:
                    loc.first.click()
                    time.sleep(1.5)
                    tab_ok = True
                    log(f"  Pestaña (parcial '{texto}') abierta ✓")
                    break

        if not tab_ok:
            # Puede que los documentos ya estén visibles sin pestaña
            log("  Sin pestaña Documents — buscando documentos en la página actual")

        # ── Mostrar todas las pestañas/tabs disponibles para diagnóstico ──────
        tabs_visibles = page.evaluate("""
            () => Array.from(document.querySelectorAll(
                'a[href], [role="tab"], [role="link"], nav a, .tab, .nav-link'
            )).map(el => ({tag: el.tagName, text: el.textContent.trim().slice(0,60), href: el.href||''}))
             .filter(x => x.text.length > 0)
        """)
        log(f"  Tabs/links en detalle: {[t['text'] for t in tabs_visibles[:15]]}")

        # ── Buscar documentos descargables ────────────────────────────────────
        # Estrategia 1: enlace cuyo texto contiene palabras clave de precio
        docs = page.evaluate("""
            () => {
                const result = [];
                for (const a of document.querySelectorAll('a[href]')) {
                    const txt  = a.textContent.trim();
                    const href = a.href;
                    result.push({ texto: txt, href: href });
                }
                return result;
            }
        """)

        url_descarga  = None
        texto_descarga = None

        # Primero: coincidencia por texto de precio
        for d in docs:
            if es_precio(d['texto'] + ' ' + d['href']):
                url_descarga   = d['href']
                texto_descarga = d['texto']
                break

        # Segundo: cualquier PDF en la página
        if not url_descarga:
            for d in docs:
                if d['href'].lower().endswith('.pdf'):
                    url_descarga   = d['href']
                    texto_descarga = d['texto'] or d['href'].split('/')[-1]
                    break

        # Tercer: enlace de descarga genérico
        if not url_descarga:
            for d in docs:
                if 'download' in d['href'].lower() or 'descargar' in d['href'].lower():
                    url_descarga   = d['href']
                    texto_descarga = d['texto'] or 'descarga'
                    break

        if not url_descarga:
            log(f"  ⬜ Sin listado de precios disponible")
            log(f"  Links en página: {[d['texto'][:30] for d in docs if d['texto']][:10]}")
            sin_doc.append(nombre_portal)
            return

        log(f"  Descargando: «{texto_descarga}» → {url_descarga[:80]}")

        # ── Descargar ─────────────────────────────────────────────────────────
        nombre_archivo = limpiar(nombre_kh)
        destino = CARPETA_DESTINO / f"Price_List_{fecha}_{nombre_archivo}.pdf"

        ok = False

        # Intento 1: petición HTTP directa
        try:
            resp = context.request.get(url_descarga, timeout=30_000)
            if resp.ok and len(resp.body()) > 1000:
                destino.write_bytes(resp.body())
                log(f"  ✓ Guardado: {destino.name}  ({len(resp.body())//1024} KB)")
                descargados.append(str(destino))
                ok = True
            else:
                raise Exception(f"HTTP {resp.status} o respuesta vacía")
        except Exception as e1:
            log(f"  HTTP directo falló ({e1}) — usando navegador...")

        # Intento 2: clic con expect_download
        if not ok:
            try:
                with page.expect_download(timeout=30_000) as dl:
                    page.locator(f'a[href="{url_descarga}"]').first.click()
                dl.value.save_as(str(destino))
                log(f"  ✓ Guardado (navegador): {destino.name}")
                descargados.append(str(destino))
                ok = True
            except Exception as e2:
                log(f"  ERROR: {e2}")
                errores.append(f"{nombre_portal}: {e2}")

    except Exception as e:
        log(f"  ERROR procesando {nombre_portal}: {e}")
        errores.append(f"{nombre_portal}: {e}")

# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main(explorar=False):
    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)
    fecha       = datetime.now().strftime("%Y-%m-%d")
    descargados = []
    errores     = []
    sin_doc     = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, slow_mo=200)
        context = browser.new_context(accept_downloads=True)
        page    = context.new_page()

        if not login(page):
            log("Login fallido — revisa usuario/contraseña en el script")
            input("Pulsa Enter para cerrar...")
            browser.close()
            return

        log("\nRecogiendo promociones...")
        promos = recoger_promociones(page)
        log(f"  {len(promos)} promociones encontradas")

        if explorar:
            print()
            print("─" * 60)
            print("  PROMOCIONES EN EL PORTAL MxM")
            print("─" * 60)
            for i, p in enumerate(promos, 1):
                print(f"  {i:2d}.  {p['nombre']:<40}  {p['url']}")
            print("─" * 60)
            browser.close()
            return

        for promo in promos:
            procesar(page, context, promo, fecha, descargados, errores, sin_doc)

        browser.close()

    # ── Resumen ───────────────────────────────────────────────────────────────
    print()
    print("=" * 60)
    print("  RESULTADO")
    print("=" * 60)

    if descargados:
        print(f"\n✓ Descargados ({len(descargados)}):")
        for f in descargados:
            print(f"    {Path(f).name}")
    else:
        print("\n  No se descargó ningún listado.")

    if sin_doc:
        print(f"\n⬜ Sin listado disponible ({len(sin_doc)}):")
        for d in sin_doc:
            print(f"    {d}")

    if errores:
        print(f"\n!! Errores ({len(errores)}):")
        for e in errores:
            print(f"    {e}")

    print()
    input("Pulsa Enter para cerrar...")

# ─── ENTRY POINT ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--explorar", action="store_true",
                        help="Lista las promociones sin descargar")
    args = parser.parse_args()

    print()
    print("=" * 60)
    print("  Descargador MxM — Más por Menos")
    if args.explorar:
        print("  MODO EXPLORACIÓN")
    print("=" * 60)
    print()

    main(explorar=args.explorar)
