#!/usr/bin/env python3
"""
Descargador de listados de precios — Prime Invest
https://www.primeinvest.es/es/mi-cuenta

Flujo:
  1. Cerrar modal de cookies
  2. Login con email + contraseña → Acceder
  3. Panel izquierdo → Promociones
  4. Poner 100 entradas por página
  5. Raspar todas las páginas de la tabla (REF, PROMOCIÓN, HAB, M², PRECIO)
  6. Agrupar por promoción y generar un PDF de tabla por cada una
  7. Guardar como Price_List_{fecha}_{nombre_promocion}.pdf

Uso:
    python descargar_primeinvest.py             — descarga todo
    python descargar_primeinvest.py --explorar  — lista promociones sin descargar
"""

import sys, os, time, argparse
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

URL_LOGIN = "https://www.primeinvest.es/es/mi-cuenta"
USUARIO   = "info@kalantarihomes.com"
PASSWORD  = "kalantarihomes2023#"

CARPETA_DESTINO = Path(r"C:\Users\User\Docs\Update\Update KH\Listados de Precios")
FECHA = datetime.now().strftime("%Y-%m-%d")

# ─── DEPENDENCIAS ─────────────────────────────────────────────────────────────

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    os.system(f'"{sys.executable}" -m pip install playwright --quiet')
    os.system(f'"{sys.executable}" -m playwright install chromium')
    from playwright.sync_api import sync_playwright

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
except ImportError:
    os.system(f'"{sys.executable}" -m pip install reportlab --quiet')
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet

# ─── HELPERS ──────────────────────────────────────────────────────────────────

def ts():
    return datetime.now().strftime("%H:%M:%S")

def log(msg):
    print(f"[{ts()}] {msg}")

def limpiar(txt):
    return "".join(c if c.isalnum() or c in " -_." else "_" for c in txt).strip()

def scroll_continuo(page, segundos=8, pausa=0.35):
    """Scroll arriba y abajo sin parar durante N segundos."""
    log(f"  Scroll continuo durante {segundos}s...")
    fin = time.time() + segundos
    bajando = True
    while time.time() < fin:
        if bajando:
            page.evaluate("window.scrollBy(0, 300)")
        else:
            page.evaluate("window.scrollBy(0, -300)")
        bajando = not bajando
        time.sleep(pausa)

def scroll_vaiven(page, ciclos=4, pausa=0.4):
    """Hace scroll abajo y arriba varias veces para activar la página."""
    for i in range(ciclos):
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(pausa)
        page.evaluate("window.scrollTo(0, 0)")
        time.sleep(pausa)
    log(f"  Scroll vaivén: {ciclos} ciclos completados")

def cerrar_modal_cookies(page, timeout=15):
    """
    Intenta cerrar el banner de cookies durante 'timeout' segundos.
    Combina scroll continuo + clics repetidos + JS de emergencia.
    """
    SELECTORES = [
        'button:has-text("Aceptar")',
        'button:has-text("Aceptar todo")',
        'button:has-text("Accept all")',
        'button:has-text("Accept")',
        'button:has-text("OK")',
        'a:has-text("Aceptar")',
        '#cookie-accept', '.cookie-accept',
        '[data-action="accept"]',
        '.cc-btn.cc-allow', '.cc-accept',
        '[aria-label*="cookie" i] button',
        '.cookie-notice-accept',
        '#cn-accept-cookie',
    ]
    fin = time.time() + timeout
    paso = 0
    while time.time() < fin:
        paso += 1
        # scroll alternado
        if paso % 2 == 0:
            page.evaluate("window.scrollBy(0, 250)")
        else:
            page.evaluate("window.scrollBy(0, -250)")

        for sel in SELECTORES:
            try:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    loc.first.click(timeout=2_000)
                    log(f"  ✓ Modal cookies cerrado ({sel})")
                    time.sleep(0.8)
                    return True
            except Exception:
                pass

        # Fallback JS: ocultar cualquier banner de cookies conocido
        eliminado = page.evaluate("""
            () => {
                const sels = [
                    '.cookie-notice', '#cookie-notice', '.cookie-banner',
                    '#cookie-banner', '.cookie-law-info-bar', '#cookie-law-info-bar',
                    '.cc-window', '#cookieChoiceInfo', '.cookie-consent',
                    '[class*="cookie"]', '[id*="cookie"]',
                ];
                for (const s of sels) {
                    const el = document.querySelector(s);
                    if (el && el.offsetParent !== null) {
                        el.remove();
                        return s;
                    }
                }
                return null;
            }
        """)
        if eliminado:
            log(f"  ✓ Banner cookies eliminado por JS ({eliminado})")
            time.sleep(0.5)
            return True

        time.sleep(0.3)

    log("  ⚠  No se detectó modal de cookies (puede que ya estuviera cerrado)")
    return False

# ─── GENERACIÓN DE PDF ────────────────────────────────────────────────────────

def generar_pdf(nombre_promo: str, unidades: list, destino: Path) -> bool:
    try:
        doc = SimpleDocTemplate(
            str(destino),
            pagesize=landscape(A4),
            leftMargin=1.5*cm, rightMargin=1.5*cm,
            topMargin=1.5*cm,  bottomMargin=1.5*cm,
        )
        styles   = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph(
            f"<b>PRICE LIST — {nombre_promo.upper()}</b>", styles['Title']))
        elements.append(Paragraph(
            f"Fecha: {datetime.now().strftime('%d/%m/%Y')}   ·   "
            f"Unidades disponibles: {len(unidades)}", styles['Normal']))
        elements.append(Spacer(1, 0.5*cm))

        header = ["REF", "HAB", "M²", "PRECIO"]
        rows   = [header] + [
            [u.get('ref',''), u.get('hab',''), u.get('m2',''), u.get('precio','')]
            for u in sorted(unidades, key=lambda x: x.get('ref', ''))
        ]

        t = Table(rows, colWidths=[7*cm, 2.5*cm, 4*cm, 5*cm], repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND',    (0,0), (-1,0),  colors.HexColor('#1E293B')),
            ('TEXTCOLOR',     (0,0), (-1,0),  colors.white),
            ('FONTNAME',      (0,0), (-1,0),  'Helvetica-Bold'),
            ('FONTSIZE',      (0,0), (-1,0),  10),
            ('ALIGN',         (0,0), (-1,0),  'CENTER'),
            ('FONTNAME',      (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE',      (0,1), (-1,-1), 9),
            ('ALIGN',         (0,1), (0,-1),  'LEFT'),
            ('ALIGN',         (1,1), (-1,-1), 'CENTER'),
            ('GRID',          (0,0), (-1,-1), 0.4, colors.HexColor('#CBD5E1')),
            ('TOPPADDING',    (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            *[('BACKGROUND',  (0,i), (-1,i),  colors.HexColor('#F8FAFC'))
              for i in range(2, len(rows), 2)],
        ]))
        elements.append(t)
        doc.build(elements)
        return True
    except Exception as e:
        log(f"  ERROR generando PDF: {e}")
        return False

# ─── MAIN ─────────────────────────────────────────────────────────────────────

def _siguiente(page):
    """Pulsa 'Siguiente/Next' si está activo. Devuelve True si pulsó."""
    return page.evaluate("""
        () => {
            const b = [...document.querySelectorAll('a,button')].find(el => {
                const t = el.textContent.trim().toLowerCase();
                return (t === 'siguiente' || t === 'next' || t === '>') &&
                       !el.classList.contains('disabled') &&
                       !el.hasAttribute('disabled');
            });
            if (b) { b.click(); return true; }
            return false;
        }
    """)


def _poner_100(page):
    """Intenta poner 100 o 'Todos' en el selector de entradas por página."""
    for sel in ['.dataTables_length select', 'select[name*="length"]',
                'select[name*="entries"]', 'select[name*="per"]']:
        loc = page.locator(sel)
        if loc.count() > 0 and loc.first.is_visible():
            for val in ['100', '-1']:
                try:
                    loc.first.select_option(val)
                    time.sleep(0.8)
                    return val
                except Exception:
                    continue
    return None


JS_RASPAR = """
(nombrePromo) => {
    const resultado = [];
    for (const tabla of document.querySelectorAll('table')) {
        const rows = tabla.querySelectorAll('tbody tr');
        if (rows.length === 0) continue;
        const ths = Array.from(tabla.querySelectorAll('thead th, thead td'))
                         .map(th => th.textContent.trim().toLowerCase());
        const iP = ths.findIndex(h => h.includes('precio') || h.includes('price'));
        if (iP === -1) continue;
        const iR  = ths.findIndex(h => h.includes('ref'));
        const iH  = ths.findIndex(h => h.includes('hab') || h.includes('dorm'));
        const iM  = ths.findIndex(h => h.includes('m²') || h.includes('m2') || h.includes('metro'));
        const iE  = ths.findIndex(h => h.includes('estado') || h.includes('status') || h.includes('disponib'));
        const iPl = ths.findIndex(h => h.includes('planta') || h.includes('floor') || h.includes('nivel'));
        for (const row of rows) {
            const c = row.querySelectorAll('td');
            if (c.length < 2) continue;
            const g = i => (i >= 0 && c[i]) ? c[i].textContent.trim() : '';
            resultado.push({promo: nombrePromo, ref: g(iR), hab: g(iH),
                            m2: g(iM), precio: g(iP), estado: g(iE), planta: g(iPl)});
        }
        break;
    }
    return resultado;
}
"""

JS_LINKS = """
() => {
    const result = [];
    for (const tabla of document.querySelectorAll('table')) {
        for (const row of tabla.querySelectorAll('tbody tr')) {
            const link   = row.querySelector('a[href]');
            const nombre = row.querySelector('td');
            if (link && link.href)
                result.push({nombre: nombre ? nombre.textContent.trim() : '', href: link.href});
        }
    }
    return result;
}
"""


def main(explorar=False):
    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)
    descargados, errores, sin_listado = [], [], []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True, slow_mo=0)
        page    = browser.new_page()

        # ── 1. Login ───────────────────────────────────────────────────────────
        log(f"Abriendo {URL_LOGIN}")
        page.goto(URL_LOGIN, wait_until="domcontentloaded", timeout=30_000)
        time.sleep(1.5)   # el modal de cookies carga asíncrono

        # ── 2. Cerrar modal de cookies (iframe de terceros) ────────────────────
        log("Cerrando modal de cookies...")
        modal_cerrado = False
        for frame in page.frames:
            for sel in ['[id*="accept"]', 'button:has-text("Aceptar")',
                        'button:has-text("Accept")', '[class*="accept"]']:
                try:
                    loc = frame.locator(sel)
                    if loc.count() > 0:
                        loc.first.click(timeout=2_000)
                        log(f"  ✓ Modal cerrado ({sel})")
                        modal_cerrado = True
                        break
                except Exception:
                    pass
            if modal_cerrado:
                break
        if not modal_cerrado:
            page.keyboard.press("Escape")

        # ── 3. Credenciales ────────────────────────────────────────────────────
        page.locator('#emailInput').fill(USUARIO)
        page.locator('#passwordInput').fill(PASSWORD)
        page.locator('button:has-text("Acceder")').click()
        log("  Credenciales enviadas")

        # Esperar a que aparezca el enlace de Promociones (indica login OK)
        try:
            page.wait_for_selector('a[href*="promotion"]', timeout=15_000)
        except Exception:
            page.wait_for_load_state("networkidle", timeout=15_000)
        log(f"  Login OK → {page.url}")

        # ── 4. Ir a Promociones ────────────────────────────────────────────────
        promo_link = page.locator('a[href*="promotion"]').first
        promo_link.click()
        page.wait_for_selector('table', timeout=10_000)
        log(f"  Promociones → {page.url}")
        url_lista = page.url

        # ── 5. Extraer links de todas las promociones ──────────────────────────
        _poner_100(page)

        seen, promos_unicas = set(), []
        pag = 0
        while True:
            pag += 1
            for p in page.evaluate(JS_LINKS):
                if p['href'] not in seen and p['href'] != url_lista:
                    seen.add(p['href'])
                    promos_unicas.append(p)
            log(f"  Pág lista {pag}: {len(promos_unicas)} total")
            if not _siguiente(page):
                break
            page.wait_for_load_state("domcontentloaded", timeout=8_000)
            time.sleep(0.5)

        log(f"  {len(promos_unicas)} promociones encontradas")

        # ── 6. Entrar en cada promoción y raspar unidades ──────────────────────
        por_promo: dict = defaultdict(list)

        for idx, promo in enumerate(promos_unicas, 1):
            nombre = promo['nombre'] or f"Promo_{idx}"
            href   = promo['href']
            log(f"  [{idx}/{len(promos_unicas)}] {nombre}")
            try:
                page.goto(href, wait_until="domcontentloaded", timeout=15_000)
                page.wait_for_selector('table', timeout=8_000)
                _poner_100(page)

                filas_todas = []
                pg = 0
                while True:
                    pg += 1
                    filas = page.evaluate(JS_RASPAR, nombre)
                    filas_todas.extend(filas)
                    if not _siguiente(page):
                        break
                    page.wait_for_load_state("domcontentloaded", timeout=8_000)
                    time.sleep(0.3)

                if filas_todas:
                    por_promo[nombre].extend(filas_todas)
                    log(f"    ✓ {len(filas_todas)} unidades")
                else:
                    log(f"    ⚠  Sin tabla")
                    sin_listado.append(nombre)
            except Exception as e:
                log(f"    ERROR: {e}")
                errores.append(f"{nombre}: {e}")

        total = sum(len(v) for v in por_promo.values())
        log(f"\n  Total unidades: {total}  |  Promociones: {len(por_promo)}")
        browser.close()

    if explorar:
        print()
        print("─" * 58)
        print("  PRIME INVEST — PROMOCIONES DISPONIBLES")
        print("─" * 58)
        for i, (nombre, units) in enumerate(sorted(por_promo.items()), 1):
            print(f"  {i:2d}.  {nombre:<40}  ({len(units)} unidades)")
        print("─" * 58)
        return

    # ── 9. Generar un PDF por promoción ───────────────────────────────────────
    print()
    for nombre_promo, unidades in sorted(por_promo.items()):
        if not unidades:
            sin_listado.append(nombre_promo)
            continue
        destino = CARPETA_DESTINO / f"Price_List_Prime_Invest_{FECHA}_{limpiar(nombre_promo)}.pdf"
        log(f"  Generando PDF: {destino.name}  ({len(unidades)} unidades)")
        if generar_pdf(nombre_promo, unidades, destino):
            log(f"  ✓ Guardado")
            descargados.append(str(destino))
        else:
            errores.append(f"{nombre_promo}: error generando PDF")

    # ── Resumen ────────────────────────────────────────────────────────────────
    print()
    print("=" * 58)
    print("  RESULTADO")
    print("=" * 58)

    if descargados:
        print(f"\n✓ PDFs generados ({len(descargados)}):")
        for f in descargados:
            print(f"    {Path(f).name}")
    else:
        print("\n  No se generó ningún PDF.")

    if sin_listado:
        print(f"\n⬜ Sin unidades ({len(sin_listado)}): {sin_listado}")

    if errores:
        print(f"\n!! Errores: {errores}")

    print()
    input("Pulsa Enter para cerrar...")


# ─── ENTRY POINT ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Descargador Prime Invest")
    parser.add_argument("--explorar", action="store_true",
                        help="Lista promociones sin descargar")
    args = parser.parse_args()

    print()
    print("=" * 58)
    print("  Descargador — Prime Invest")
    if args.explorar:
        print("  MODO EXPLORACIÓN")
    print("=" * 58)
    print()

    main(explorar=args.explorar)
