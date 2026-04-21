#!/usr/bin/env python3
"""
Descargador automático — ON3 Portal Brokers
Descarga los PDFs de cada promoción en https://on3.es/brokers/

Uso: python descargar_on3.py
"""

import sys, os, time
from pathlib import Path
from datetime import datetime

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

URL_LOGIN   = "https://on3.es/acceso-broker/"
URL_BROKERS = "https://on3.es/brokers/"
USUARIO     = "kalantari Homes"
PASSWORD    = "G0CVq(vxJT%)H3bRP3EYho!@"

CARPETA_DESTINO = Path(r"C:\Users\User\Docs\Update\Update KH\Listados de Precios")

# ─── PALABRAS CLAVE PARA FILTRAR LISTADOS DE PRECIOS ─────────────────────────

KEYWORDS_PRECIO = [
    'lp', ' lp ', '_lp_', '-lp-', 'lp_', '_lp',
    'price list', 'pricelist', 'price-list', 'price_list',
    'lista de precio', 'lista precios', 'listado de precio',
    'listado precios', 'tarifa', 'precios', 'lista_precios',
]

# ─── INSTALACIÓN AUTOMÁTICA DE PLAYWRIGHT ────────────────────────────────────

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Instalando playwright (solo la primera vez)...")
    os.system(f'"{sys.executable}" -m pip install playwright --quiet')
    os.system(f'"{sys.executable}" -m playwright install chromium')
    from playwright.sync_api import sync_playwright

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def ts():
    return datetime.now().strftime("%H:%M:%S")

def log(msg):
    print(f"[{ts()}] {msg}")

def limpiar(texto):
    return "".join(c if c.isalnum() or c in " -_." else "_" for c in texto).strip()

def es_listado_precios(texto):
    n = texto.lower().replace('_', ' ').replace('-', ' ')
    # "lp" solo como token separado para evitar falsos positivos
    import re
    if re.search(r'\blp\b', n):
        return True
    return any(kw.replace('_', ' ').replace('-', ' ') in n
               for kw in KEYWORDS_PRECIO if kw not in ('lp', ' lp ', '_lp_', '-lp-', 'lp_', '_lp'))

# ─── LÓGICA PRINCIPAL ─────────────────────────────────────────────────────────

def main():
    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)
    fecha       = datetime.now().strftime("%Y-%m-%d")
    descargados = []
    errores     = []
    sin_pdf     = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True, slow_mo=0)
        context = browser.new_context(accept_downloads=True)
        page    = context.new_page()

        # ── 1. LOGIN ──────────────────────────────────────────────────────────
        log(f"Cargando {URL_LOGIN}...")
        page.goto(URL_LOGIN, wait_until="domcontentloaded", timeout=30_000)

        # Cookie modal (DOM principal, clase cmplz-accept)
        try:
            btn_cookie = page.locator("button.cmplz-accept")
            btn_cookie.wait_for(state="visible", timeout=5_000)
            btn_cookie.click()
            log("  ✓ Cookie modal aceptado")
        except Exception:
            pass  # Modal puede no aparecer si ya se aceptó antes

        # Credenciales
        try:
            page.locator('input[name="log"]').first.fill(USUARIO)
            page.locator('input[name="pwd"]').first.fill(PASSWORD)
            log("  ✓ Credenciales rellenadas")
        except Exception as e:
            log(f"  ERROR rellenando credenciales: {e}")
            browser.close()
            return [], [f"Login fallido: {e}"], []

        # Pulsar Acceder
        page.locator('button:has-text("Acceder")').first.click()
        log("  Acceder pulsado")
        try:
            page.wait_for_url("**/brokers/**", timeout=15_000)
        except Exception:
            pass
        log(f"  URL tras login: {page.url}")

        if "on3.es" not in page.url:
            log("  !! Login fallido — revisa credenciales")
            browser.close()
            return [], ["Login fallido: URL inesperada"], []

        # ── 2. CARGAR PANEL DE BROKERS ────────────────────────────────────────
        if "brokers" not in page.url:
            page.goto(URL_BROKERS, wait_until="domcontentloaded", timeout=20_000)

        # Scroll rápido para forzar lazy-load de todas las tarjetas
        log("  Cargando todas las promociones (scroll)...")
        altura_prev = 0
        for _ in range(20):
            page.evaluate("window.scrollBy(0, 2000)")
            time.sleep(0.2)
            altura = page.evaluate("document.body.scrollHeight")
            if altura == altura_prev:
                break
            altura_prev = altura

        # ── 3. EXTRAER TODAS LAS PROMOCIONES ──────────────────────────────────
        promociones_raw = page.evaluate("""
            () => {
                const seen  = new Set();
                const named = {};   // href → nombre con texto
                const all   = [];   // todos los hrefs

                for (const a of document.querySelectorAll('a[href*="/promocion/"]')) {
                    const href   = a.href.split('?')[0].split('#')[0].replace(/\\/+$/, '') + '/';
                    const nombre = a.textContent.trim();
                    if (!seen.has(href)) {
                        seen.add(href);
                        all.push(href);
                    }
                    if (nombre && !named[href])
                        named[href] = nombre;
                }
                return all.map(href => ({
                    href,
                    nombre: named[href] ||
                            href.split('/').filter(Boolean).pop()
                                .replace(/-broker$/, '').replace(/-/g, ' ')
                                .replace(/\b\w/g, c => c.toUpperCase())
                }));
            }
        """)

        if not promociones_raw:
            log("  !! No se encontraron promociones. ¿Login correcto?")
            browser.close()
            return [], ["No se encontraron promociones"], []

        log(f"\n  {len(promociones_raw)} promociones encontradas:")
        for p in promociones_raw:
            log(f"    · {p['nombre']}  →  {p['href']}")

        # ── 4. PROCESAR CADA PROMOCIÓN ────────────────────────────────────────
        for idx, promo in enumerate(promociones_raw, 1):
            nombre_promo = promo['nombre']
            url_promo    = promo['href']

            log(f"\n{'─'*55}")
            log(f"[{idx}/{len(promociones_raw)}] {nombre_promo}")

            try:
                page.goto(url_promo, wait_until="domcontentloaded", timeout=20_000)

                # Recoger todos los enlaces a PDF
                pdfs = page.evaluate("""
                    () => {
                        const result = [];
                        const seen   = new Set();
                        for (const a of document.querySelectorAll('a[href]')) {
                            const href = a.href;
                            const txt  = a.textContent.trim();
                            if (seen.has(href)) continue;
                            if (href.toLowerCase().includes('.pdf') ||
                                txt.toLowerCase() === 'descargar' ||
                                a.getAttribute('download') !== null) {
                                seen.add(href);
                                result.push({href, txt});
                            }
                        }
                        return result;
                    }
                """)

                if not pdfs:
                    log(f"  ⚠  Sin PDFs — se omite")
                    sin_pdf.append(nombre_promo)
                    continue

                log(f"  {len(pdfs)} PDF(s) disponibles")

                # Solo descargar si hay coincidencia con keywords de precio
                listados = [p for p in pdfs
                            if es_listado_precios(p['href']) or es_listado_precios(p['txt'])]
                if not listados:
                    log(f"  ⚠  Sin listado de precios identificado — se omite")
                    sin_pdf.append(nombre_promo)
                    continue

                log(f"  {len(listados)} listado(s) de precios encontrado(s)")

                for pdf in listados:
                    href          = pdf['href']
                    nombre_doc    = href.split('/')[-1].split('?')[0] or 'doc.pdf'
                    nombre_limpio = limpiar(nombre_promo)[:45]
                    destino       = CARPETA_DESTINO / \
                        f"Price_List_ON3_{fecha}_{nombre_limpio}_{nombre_doc}"

                    log(f"  Descargando: {nombre_doc}...")
                    try:
                        resp = context.request.get(href, timeout=30_000)
                        if resp.ok:
                            destino.write_bytes(resp.body())
                            log(f"  ✓ Guardado: {destino.name}")
                            descargados.append(str(destino))
                        else:
                            raise Exception(f"HTTP {resp.status}")
                    except Exception as e:
                        log(f"  ERROR: {e}")
                        errores.append(f"{nombre_promo} [{nombre_doc}]: {e}")

            except Exception as e:
                log(f"  ERROR procesando {nombre_promo}: {e}")
                errores.append(f"{nombre_promo}: {e}")

        browser.close()

    return descargados, errores, sin_pdf


# ─── PUNTO DE ENTRADA ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    print()
    print("=" * 58)
    print("  Descargador de PDFs — ON3 Portal Brokers")
    print("=" * 58)
    print()

    descargados, errores, sin_pdf = main()

    print()
    print("=" * 58)
    print("  RESULTADO")
    print("=" * 58)

    if descargados:
        print(f"\nPDFs descargados ({len(descargados)}):")
        for f in descargados:
            print(f"  OK  {Path(f).name}")
    else:
        print("\nNo se descargó ningún PDF.")

    if sin_pdf:
        print(f"\n{'─'*58}")
        print(f"  ⚠  Promociones sin PDFs ({len(sin_pdf)}):")
        for p in sin_pdf:
            print(f"  !!  {p}")

    if errores:
        print(f"\nErrores técnicos ({len(errores)}):")
        for e in errores:
            print(f"  !!  {e}")

    print()
    input("Pulsa Enter para cerrar...")
