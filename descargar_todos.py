#!/usr/bin/env python3
"""
Descargador unificado de listados de precios — KH Inmobiliaria
Portales: Bromley · Dream · Magnum & Partners · MxM · Prime Invest · ON3

Uso:
    python descargar_todos.py                         — ejecuta los 6 portales
    python descargar_todos.py --portal bromley         — solo Bromley
    python descargar_todos.py --portal magnum mxm primeinvest on3
    python descargar_todos.py --explorar              — lista promociones
"""

import sys, os, re, time, argparse
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN GLOBAL
# ══════════════════════════════════════════════════════════════════════════════

CARPETA_DESTINO = Path(r"C:\Users\User\Docs\Update\Update KH\Listados de Precios")
FECHA = datetime.now().strftime("%Y-%m-%d")

# ── Bromley ───────────────────────────────────────────────────────────────────
BROMLEY_URL_LOGIN = "https://ap.bromleyestatesmarbella.com/welcome/"
BROMLEY_USUARIO   = "info@kalantarihomes.com"
BROMLEY_PASSWORD  = "kalantarihomes2023#"
BROMLEY_DESARROLLOS = [
    {"nombre": "Infinity Homes",       "url": "https://ap.bromleyestatesmarbella.com/development-doc/las-mesas-infinity-homes/"},
    {"nombre": "Las Mesas Collection", "url": "https://ap.bromleyestatesmarbella.com/development-doc/las-mesas-collection/"},
    {"nombre": "Las Mesas Sea Suites", "url": "https://ap.bromleyestatesmarbella.com/development-doc/las-mesas-sea-suites/"},
]

# ── Dream ─────────────────────────────────────────────────────────────────────
DREAM_URL_LOGIN  = "https://agentsportal.dreamexclusives.com/index.php"
DREAM_USUARIO    = "info@kalantarihomes.com"
DREAM_PASSWORD   = "kalantarihomes2023#"
DREAM_DESARROLLOS = [
    {"nombre": "Capri",                 "url_downloads": "https://agentsportal.dreamexclusives.com/index.php?opt=Promociones&buscar=24&tab=downloads"},
    {"nombre": "Adel San Roque",        "url_downloads": "https://agentsportal.dreamexclusives.com/index.php?opt=Promociones&buscar=26&tab=downloads"},
    {"nombre": "The Palms At Estepona", "url_downloads": "https://agentsportal.dreamexclusives.com/index.php?opt=Promociones&buscar=39&tab=downloads"},
]

# ── Magnum & Partners ─────────────────────────────────────────────────────────
MAGNUM_URL_LOGIN       = "https://www.magnum-partners.com/brokers/"
MAGNUM_URL_PROMOCIONES = "https://www.magnum-partners.com/brokers/promociones/"
MAGNUM_N_PAGINAS       = 6
MAGNUM_USUARIO         = "info@kalantarihomes.com"
MAGNUM_PASSWORD        = "kalantarihomes2023#"
MAGNUM_DESARROLLOS = [
    "Aby", "Adagio", "Aire", "Altara Alcaidesa", "Be Aloha",
    "Oceana Gardens", "Symphony Suites", "Zenith", "Zinnia",
]
MAGNUM_MAPEO_NOMBRES = {
    "aby middle": "Aby Estepona",
    "aby upper":  "Aby Estepona",
}

# ── MxM (Más por Menos) ───────────────────────────────────────────────────────
MXM_URL_BASE  = "https://intranet.inmobiliariamaspormenos.com"
MXM_URL_LOGIN = MXM_URL_BASE + "/"
MXM_URL_LIST  = MXM_URL_BASE + "/inmuebles/"
MXM_USUARIO   = "Admin"
MXM_PASSWORD  = "MxMON"
MXM_MAPEO_NOMBRES = {
    "spinto blu":          "Spinto Blu",
    "skye casares golf":   "Skye",
    "idyllic":             "Ydillic",
    "camarate hills ii":   "Camarate Hills",
    "eleven41":            "Eleven 41",
    "pure sun ii":         "PureSun II",
    "natura estepona":     "Natura",
}

# ── Prime Invest ──────────────────────────────────────────────────────────────
PRIMEINVEST_URL_LOGIN = "https://www.primeinvest.es/es/mi-cuenta"
PRIMEINVEST_USUARIO   = "info@kalantarihomes.com"
PRIMEINVEST_PASSWORD  = "kalantarihomes2023#"

# ── ON3 ───────────────────────────────────────────────────────────────────────
ON3_URL_LOGIN   = "https://on3.es/acceso-broker/"
ON3_URL_BROKERS = "https://on3.es/brokers/"
ON3_USUARIO     = "kalantari Homes"
ON3_PASSWORD    = "G0CVq(vxJT%)H3bRP3EYho!@"

# ── Palabras clave para identificar listados de precios ───────────────────────
KEYWORDS_PRECIO = [
    'price list', 'pricelist', 'price-list', 'price_list',
    'lista de precio', 'lista precios', 'listado precio',
    'tarifa', 'precios',
]

# ══════════════════════════════════════════════════════════════════════════════
#  DEPENDENCIAS
# ══════════════════════════════════════════════════════════════════════════════

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Instalando playwright (solo la primera vez)...")
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

# ══════════════════════════════════════════════════════════════════════════════
#  UTILIDADES COMUNES
# ══════════════════════════════════════════════════════════════════════════════

def ts():
    return datetime.now().strftime("%H:%M:%S")

def log(msg):
    print(f"[{ts()}] {msg}")

def limpiar(txt):
    return "".join(c if c.isalnum() or c in " -_." else "_" for c in txt).strip()

def es_precio(texto):
    t = texto.lower().replace('_', ' ').replace('-', ' ')
    if re.search(r'\blp\b', t):
        return True
    return any(k.replace('_', ' ').replace('-', ' ') in t for k in KEYWORDS_PRECIO)

def nombre_coincide(nombre_portal, filtros):
    if not filtros:
        return True
    np = nombre_portal.lower().strip()
    return any(f.lower().strip() in np or np in f.lower().strip() for f in filtros)

def guardar_pdf(context, url_descarga, destino, page=None):
    destino = Path(destino)
    if destino.exists() and destino.stat().st_size > 1000:
        log(f"  ↩ Ya existe: {destino.name}")
        return True
    try:
        resp = context.request.get(url_descarga, timeout=30_000)
        body = resp.body()
        if resp.ok and len(body) > 1000:
            destino.write_bytes(body)
            log(f"  ✓ {destino.name}  ({len(body)//1024} KB)")
            return True
        raise Exception(f"HTTP {resp.status}")
    except Exception:
        pass
    if page:
        try:
            with page.expect_download(timeout=30_000) as dl:
                page.locator(f'a[href="{url_descarga}"]').first.click()
            dl.value.save_as(str(destino))
            log(f"  ✓ {destino.name} (navegador)")
            return True
        except Exception as e:
            log(f"  ERROR descarga: {e}")
    return False

def _siguiente(page):
    return page.evaluate("""
        () => {
            const b = [...document.querySelectorAll('a,button')].find(el => {
                const t = el.textContent.trim().toLowerCase();
                return (t === 'siguiente' || t === 'next' || t === '>') &&
                       !el.classList.contains('disabled') && !el.hasAttribute('disabled');
            });
            if (b) { b.click(); return true; }
            return false;
        }
    """)

def _poner_100(page):
    for sel in ['.dataTables_length select', 'select[name*="length"]',
                'select[name*="entries"]', 'select[name*="per"]']:
        loc = page.locator(sel)
        if loc.count() > 0 and loc.first.is_visible():
            for val in ['100', '-1']:
                try:
                    loc.first.select_option(val)
                    time.sleep(0.3)
                    return val
                except Exception:
                    continue
    return None

def _login_generico(page, usuario, password, selectores_user, selectores_pass, selectores_btn):
    for sel in selectores_user:
        loc = page.locator(sel)
        if loc.count() > 0 and loc.first.is_visible():
            loc.first.fill(usuario); break
    for sel in selectores_pass:
        loc = page.locator(sel)
        if loc.count() > 0 and loc.first.is_visible():
            loc.first.fill(password); break
    for sel in selectores_btn:
        loc = page.locator(sel)
        if loc.count() > 0 and loc.first.is_visible():
            loc.first.click(); return
    page.locator('input[type="password"]').first.press("Enter")

def _nuevo_browser(pw, headless=True):
    browser = pw.chromium.launch(headless=headless, slow_mo=0)
    context = browser.new_context(accept_downloads=True)
    return browser, context, context.new_page()

# ══════════════════════════════════════════════════════════════════════════════
#  PORTAL 1 — BROMLEY ESTATES
# ══════════════════════════════════════════════════════════════════════════════

def run_bromley(explorar=False):
    log("\n" + "═"*58)
    log("  BROMLEY ESTATES")
    log("═"*58)
    descargados, errores, sin_listado = [], [], []

    with sync_playwright() as pw:
        browser, context, page = _nuevo_browser(pw)

        page.goto(BROMLEY_URL_LOGIN, wait_until="domcontentloaded", timeout=30_000)
        _login_generico(page, BROMLEY_USUARIO, BROMLEY_PASSWORD,
            ['input[name="log"]', 'input[name="user_login"]', 'input[type="email"]', 'input[type="text"]'],
            ['input[name="pwd"]', 'input[name="user_pass"]', 'input[type="password"]'],
            ['input[value="LOGIN"]', 'input[value="Login"]', 'input[name="wp-submit"]',
             'button:has-text("LOGIN")', 'button:has-text("Login")', 'button[type="submit"]'])
        page.wait_for_load_state("domcontentloaded", timeout=20_000)
        log(f"  Login OK → {page.url}")

        if explorar:
            for dev in BROMLEY_DESARROLLOS:
                log(f"    · {dev['nombre']}  →  {dev['url']}")
            browser.close(); return [], [], []

        for dev in BROMLEY_DESARROLLOS:
            nombre, url = dev["nombre"], dev["url"]
            log(f"  ── {nombre}")
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=20_000)
                try:
                    page.wait_for_selector('text="Availability"', timeout=6_000)
                    page.get_by_text("Availability", exact=True).first.click()
                    page.wait_for_selector('table tbody tr', timeout=5_000)
                except Exception:
                    pass

                archivos = page.evaluate("""
                    () => {
                        let panel = document.querySelector('.su-tabs-pane-open');
                        if (!panel) {
                            for (const p of document.querySelectorAll('.su-tabs-pane')) {
                                if (window.getComputedStyle(p).display !== 'none') { panel = p; break; }
                            }
                        }
                        if (!panel) return [];
                        const result = [];
                        for (const row of panel.querySelectorAll('table tbody tr')) {
                            const cells = row.querySelectorAll('td');
                            if (cells.length < 4) continue;
                            const link = cells[3].querySelector('a');
                            if (link?.href) result.push({nombre: cells[0].textContent.trim(), href: link.href});
                        }
                        return result;
                    }
                """)

                listados = [a for a in archivos if es_precio(a['nombre'])]
                if not listados:
                    log(f"    ⬜ Sin listado"); sin_listado.append(nombre); continue

                for arch in listados:
                    destino = CARPETA_DESTINO / f"Price_List_Bromley_Estates_{FECHA}_{limpiar(arch['nombre'])}.pdf"
                    if guardar_pdf(context, arch['href'], destino, page):
                        descargados.append(str(destino))
                    else:
                        errores.append(f"Bromley/{nombre}: descarga fallida")
            except Exception as e:
                log(f"  ERROR: {e}"); errores.append(f"Bromley/{nombre}: {e}")

        browser.close()
    return descargados, errores, sin_listado


# ══════════════════════════════════════════════════════════════════════════════
#  PORTAL 2 — DREAM EXCLUSIVES
# ══════════════════════════════════════════════════════════════════════════════

def run_dream(explorar=False):
    log("\n" + "═"*58)
    log("  DREAM EXCLUSIVES")
    log("═"*58)
    descargados, errores, sin_listado = [], [], []

    with sync_playwright() as pw:
        browser, context, page = _nuevo_browser(pw)

        page.goto(DREAM_URL_LOGIN, wait_until="domcontentloaded", timeout=30_000)
        _login_generico(page, DREAM_USUARIO, DREAM_PASSWORD,
            ['input[name="email"]', 'input[name="username"]', 'input[type="email"]', 'input[type="text"]'],
            ['input[name="password"]', 'input[name="pwd"]', 'input[type="password"]'],
            ['button[type="submit"]', 'input[type="submit"]', 'button:has-text("Sign in")', 'button:has-text("Login")'])
        page.wait_for_load_state("domcontentloaded", timeout=20_000)
        log(f"  Login OK → {page.url}")

        if explorar:
            for dev in DREAM_DESARROLLOS:
                log(f"    · {dev['nombre']}  →  {dev['url_downloads']}")
            browser.close(); return [], [], []

        for dev in DREAM_DESARROLLOS:
            nombre, url = dev["nombre"], dev["url_downloads"]
            log(f"  ── {nombre}")
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=20_000)

                archivos = page.evaluate("""
                    () => {
                        const result = [];
                        for (const a of document.querySelectorAll('a[href]')) {
                            const href   = a.href || '';
                            const nombre = a.textContent.trim() || href.split('/').pop();
                            if (href.toLowerCase().includes('.pdf') || href.includes('download') ||
                                a.getAttribute('download') !== null)
                                result.push({nombre, href});
                        }
                        return result;
                    }
                """)

                listados = [a for a in archivos if es_precio(a['nombre'])]
                if not listados:
                    log(f"    ⬜ Sin listado"); sin_listado.append(nombre); continue

                for arch in listados:
                    destino = CARPETA_DESTINO / f"Price_List_Dream_Exclusives_{FECHA}_{limpiar(arch['nombre'])}.pdf"
                    if guardar_pdf(context, arch['href'], destino, page):
                        descargados.append(str(destino))
                    else:
                        errores.append(f"Dream/{nombre}: descarga fallida")
            except Exception as e:
                log(f"  ERROR: {e}"); errores.append(f"Dream/{nombre}: {e}")

        browser.close()
    return descargados, errores, sin_listado


# ══════════════════════════════════════════════════════════════════════════════
#  PORTAL 3 — MAGNUM & PARTNERS
# ══════════════════════════════════════════════════════════════════════════════

def run_magnum(explorar=False):
    log("\n" + "═"*58)
    log("  MAGNUM & PARTNERS")
    log("═"*58)
    descargados, errores, sin_listado = [], [], []

    with sync_playwright() as pw:
        browser, context, page = _nuevo_browser(pw)

        page.goto(MAGNUM_URL_LOGIN, wait_until="domcontentloaded", timeout=30_000)
        _login_generico(page, MAGNUM_USUARIO, MAGNUM_PASSWORD,
            ['input[name="email"]', 'input[name="username"]', 'input[type="email"]', 'input[type="text"]'],
            ['input[type="password"]', 'input[name="password"]', 'input[name="pwd"]'],
            ['input[value="Acceder"]', 'input[value="ACCEDER"]', 'input[name="wp-submit"]',
             'button:has-text("Acceder")', 'button[type="submit"]', 'input[type="submit"]'])
        page.wait_for_load_state("domcontentloaded", timeout=25_000)
        log(f"  Login OK → {page.url}")

        # Recoger todas las promociones
        todas = []
        for num_pag in range(1, MAGNUM_N_PAGINAS + 1):
            url_pag = MAGNUM_URL_PROMOCIONES if num_pag == 1 \
                      else f"{MAGNUM_URL_PROMOCIONES.rstrip('/')}/page/{num_pag}/"
            try:
                page.goto(url_pag, wait_until="domcontentloaded", timeout=15_000)
                promos_pag = page.evaluate("""
                    () => {
                        const cards = [];
                        for (const a of document.querySelectorAll('a')) {
                            const txt = a.textContent.trim().toLowerCase();
                            if (!txt.includes('saber') && !txt.includes('más')) continue;
                            let nombre = '', el = a;
                            for (let i = 0; i < 8 && !nombre; i++) {
                                el = el.parentElement;
                                if (!el) break;
                                for (const h of el.querySelectorAll('h1,h2,h3,h4,h5,.title,.name')) {
                                    const t = h.textContent.trim();
                                    if (t) { nombre = t; break; }
                                }
                            }
                            if (a.href && nombre) cards.push({nombre, url_detalle: a.href});
                        }
                        return cards;
                    }
                """)
                if not promos_pag:
                    break
                todas.extend(promos_pag)
            except Exception:
                break

        EXCLUIR = {'hola', 'perfil', 'cerrar sesión', 'logout', 'inicio', 'home', 'contacto'}
        seen, unicas = set(), []
        for p in todas:
            nn = p['nombre'].strip().lower().rstrip(',')
            if nn in EXCLUIR or len(nn) < 3: continue
            if p['url_detalle'] not in seen:
                seen.add(p['url_detalle']); unicas.append(p)

        log(f"  {len(unicas)} promociones en el portal")

        if explorar:
            for i, p in enumerate(unicas, 1):
                print(f"  {i:2d}. {p['nombre']:<40}  {p['url_detalle']}")
            browser.close(); return [], [], []

        promos = [p for p in unicas if nombre_coincide(p['nombre'], MAGNUM_DESARROLLOS)]
        no_enc = [d for d in MAGNUM_DESARROLLOS
                  if not any(nombre_coincide(p['nombre'], [d]) for p in unicas)]
        if no_enc:
            log(f"  ⚠  No encontrados: {', '.join(no_enc)}")
            errores.extend([f"Magnum/{d}: no encontrado" for d in no_enc])

        for promo in promos:
            nombre = promo['nombre']
            log(f"  ── {nombre}")
            try:
                page.goto(promo['url_detalle'], wait_until="domcontentloaded", timeout=15_000)

                planos_links = page.evaluate(
                    "() => Array.from(document.querySelectorAll('a[href]'))"
                    ".filter(a => a.href.includes('/planos/')).map(a => a.href)")
                if not planos_links:
                    log(f"    ⬜ Sin planos"); sin_listado.append(nombre); continue

                page.goto(planos_links[0], wait_until="domcontentloaded", timeout=15_000)

                url_descarga = page.evaluate("""
                    () => {
                        for (const a of document.querySelectorAll('a[href]')) {
                            const txt  = a.textContent.trim().toLowerCase();
                            const href = a.href.toLowerCase();
                            if ((txt.includes('descargar') && (txt.includes('precio') || txt.includes('lista'))) ||
                                txt.includes('price list') || (href.includes('price') && href.endsWith('.pdf')))
                                return a.href;
                        }
                        for (const a of document.querySelectorAll('a[href]'))
                            if (a.href.toLowerCase().endsWith('.pdf')) return a.href;
                        return null;
                    }
                """)

                if not url_descarga:
                    log(f"    ⬜ Sin listado"); sin_listado.append(nombre); continue

                nombre_kh = MAGNUM_MAPEO_NOMBRES.get(nombre.lower().strip(), nombre)
                destino   = CARPETA_DESTINO / f"Price_List_Magnum_{FECHA}_{limpiar(nombre_kh)}.pdf"
                if guardar_pdf(context, url_descarga, destino, page):
                    descargados.append(str(destino))
                else:
                    errores.append(f"Magnum/{nombre}: descarga fallida")

            except Exception as e:
                log(f"  ERROR: {e}"); errores.append(f"Magnum/{nombre}: {e}")

        browser.close()
    return descargados, errores, sin_listado


# ══════════════════════════════════════════════════════════════════════════════
#  PORTAL 4 — MXM (MÁS POR MENOS)
# ══════════════════════════════════════════════════════════════════════════════

def run_mxm(explorar=False):
    log("\n" + "═"*58)
    log("  MÁS POR MENOS (MxM)")
    log("═"*58)
    descargados, errores, sin_listado = [], [], []

    with sync_playwright() as pw:
        browser, context, page = _nuevo_browser(pw)

        page.goto(MXM_URL_LOGIN, wait_until="domcontentloaded", timeout=30_000)
        page.locator('input[type="text"]').first.fill(MXM_USUARIO)
        page.locator('input[type="password"]').first.fill(MXM_PASSWORD)
        page.locator('input[type="submit"]').first.click()
        page.wait_for_load_state("domcontentloaded", timeout=20_000)
        log(f"  Login OK → {page.url}")

        if "/inmuebles" not in page.url:
            page.goto(MXM_URL_LIST, wait_until="domcontentloaded", timeout=20_000)

        promos = page.evaluate("""
            () => {
                const result = [];
                for (const row of document.querySelectorAll('div.row')) {
                    const nameEl = row.querySelector('span.dblock');
                    const linkEl = row.querySelector('a[href^="/p-"]');
                    if (nameEl && linkEl)
                        result.push({nombre: nameEl.textContent.trim(), url: linkEl.href});
                }
                return result;
            }
        """)
        seen, unicos = set(), []
        for p in promos:
            if p['url'] not in seen:
                seen.add(p['url']); unicos.append(p)

        log(f"  {len(unicos)} promociones en el portal")

        if explorar:
            for i, p in enumerate(unicos, 1):
                print(f"  {i:2d}. {p['nombre']:<40}  {p['url']}")
            browser.close(); return [], [], []

        for promo in unicos:
            nombre_portal = promo['nombre']
            nombre_kh     = MXM_MAPEO_NOMBRES.get(nombre_portal.lower().strip(), nombre_portal)
            log(f"  ── {nombre_portal}")
            try:
                page.goto(promo['url'], wait_until="domcontentloaded", timeout=15_000)

                for texto in ["Documents", "Documentos", "Docs", "Files", "Archivos"]:
                    loc = page.get_by_role("link", name=texto).or_(page.get_by_role("tab", name=texto))
                    if loc.count() > 0:
                        loc.first.click()
                        try:
                            page.wait_for_selector('a[href$=".pdf"]', timeout=4_000)
                        except Exception:
                            pass
                        break

                docs = page.evaluate(
                    "() => Array.from(document.querySelectorAll('a[href]'))"
                    ".map(a => ({texto: a.textContent.trim(), href: a.href}))")

                url_descarga = None
                for d in docs:
                    if es_precio(d['texto'] + ' ' + d['href']):
                        url_descarga = d['href']; break
                if not url_descarga:
                    for d in docs:
                        if d['href'].lower().endswith('.pdf'):
                            url_descarga = d['href']; break

                if not url_descarga:
                    log(f"    ⬜ Sin listado"); sin_listado.append(nombre_portal); continue

                destino = CARPETA_DESTINO / f"Price_List_MxM_{FECHA}_{limpiar(nombre_kh)}.pdf"
                if guardar_pdf(context, url_descarga, destino, page):
                    descargados.append(str(destino))
                else:
                    errores.append(f"MxM/{nombre_portal}: descarga fallida")

            except Exception as e:
                log(f"  ERROR: {e}"); errores.append(f"MxM/{nombre_portal}: {e}")

        browser.close()
    return descargados, errores, sin_listado


# ══════════════════════════════════════════════════════════════════════════════
#  PORTAL 5 — PRIME INVEST
# ══════════════════════════════════════════════════════════════════════════════

# JavaScript reutilizable para raspar unidades de una promoción
_JS_RASPAR_PI = """
(nombrePromo) => {
    for (const tabla of document.querySelectorAll('table')) {
        const rows = tabla.querySelectorAll('tbody tr');
        if (rows.length === 0) continue;
        const ths = Array.from(tabla.querySelectorAll('thead th, thead td'))
                         .map(th => th.textContent.trim().toLowerCase());
        const iP  = ths.findIndex(h => h.includes('precio') || h.includes('price'));
        if (iP === -1) continue;
        const iR  = ths.findIndex(h => h.includes('ref'));
        const iH  = ths.findIndex(h => h.includes('hab') || h.includes('dorm'));
        const iM  = ths.findIndex(h => h.includes('m²') || h.includes('m2') || h.includes('metro'));
        const iE  = ths.findIndex(h => h.includes('estado') || h.includes('status') || h.includes('disponib'));
        const iPl = ths.findIndex(h => h.includes('planta') || h.includes('floor') || h.includes('nivel'));
        const resultado = [];
        for (const row of rows) {
            const c = row.querySelectorAll('td');
            if (c.length < 2) continue;
            const g = i => (i >= 0 && c[i]) ? c[i].textContent.trim() : '';
            resultado.push({promo: nombrePromo, ref: g(iR), hab: g(iH),
                            m2: g(iM), precio: g(iP), estado: g(iE), planta: g(iPl)});
        }
        return resultado;
    }
    return [];
}
"""

_JS_LINKS_PI = """
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


def _floor_from_ref(ref: str) -> str:
    """
    Deriva la etiqueta de planta desde el código de referencia de la unidad.
    Se usa como fallback cuando el portal no proporciona columna PLANTA.

    Patrones soportados:
      AE-1.02.A   → "2"      (bloque.planta.unidad con puntos como separador)
      AE-2.00.D   → "0"      (planta baja numérica)
      AE-1.S01.A  → "Sótano" (prefijo S = sótano)
      AE-1.AT.A   → "Ático"  (prefijo AT/A = ático)
      360-10A     → "10"     (prefijo-2dígitos+letra)
      AS-11B      → "11"
      AN-24A      → "2"      (solo primer dígito cuando 2dígitos+letra larga)
      AN-2        → "2"      (solo número tras guión)
    """
    ref = str(ref).strip()
    if not ref:
        return ''

    # Patrón 1: PREFIX-block.FLOOR.unit  (ej. AE-1.02.A, AE-2.00.D, AE-1.S01.A)
    m = re.search(r'\d+\.([^.]+)\.[^.]+$', ref)
    if m:
        seg = m.group(1)
        prefix = re.sub(r'\d', '', seg).upper()
        digits = re.sub(r'\D', '', seg).lstrip('0') or '0'
        if prefix in ('S', 'SS', 'SOT'):
            return 'Sótano'
        if prefix in ('A', 'AT', 'PH'):
            return 'Ático'
        return digits

    # Patrón 2: PREFIX-{2+digits}{letter(s)}  (ej. 360-10A, AS-11B, AS-40A)
    m = re.search(r'-(\d{2,})[A-Za-z]+$', ref)
    if m:
        return m.group(1).lstrip('0') or '0'

    # Patrón 3: PREFIX-{digit}{digit}{letter+}  (ej. AN-24A → piso 2)
    # Solo aplica si hay exactamente 2 dígitos seguidos de al menos 1 letra
    m = re.search(r'-(\d)(\d[A-Za-z]+)$', ref)
    if m:
        return m.group(1)

    # Patrón 4: PREFIX-{digits}  (ej. AN-2, AN-11 sin letra final)
    m = re.search(r'-(\d+)$', ref)
    if m:
        return m.group(1)

    # Patrón 5: Último recurso — primer dígito(s) justo tras cualquier guión
    # (ej. 360-4-1A → "4", AL4-2113 → "2113")
    m = re.search(r'-(\d+)', ref)
    if m:
        return m.group(1)

    return ''


def _generar_pdf_pi(nombre_promo: str, unidades: list, destino: Path) -> bool:
    try:
        doc = SimpleDocTemplate(str(destino), pagesize=landscape(A4),
                                leftMargin=1.5*cm, rightMargin=1.5*cm,
                                topMargin=1.5*cm, bottomMargin=1.5*cm)
        styles = getSampleStyleSheet()

        # Determinar valor de PLANTA para cada unidad:
        # 1º usa lo que el portal proporcionó, 2º deriva del código REF
        def _planta(u):
            p = str(u.get('planta', '') or '').strip()
            return p if p else _floor_from_ref(u.get('ref', ''))

        sorted_units = sorted(unidades, key=lambda x: x.get('ref', ''))
        rows = [["REF", "HAB", "M²", "PLANTA", "PRECIO"]] + [
            [u.get('ref',''), str(u.get('hab','')), u.get('m2',''),
             _planta(u), u.get('precio','')]
            for u in sorted_units
        ]
        t = Table(rows, colWidths=[6.5*cm, 2*cm, 3.5*cm, 3.5*cm, 5*cm], repeatRows=1)
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
        doc.build([
            Paragraph(f"<b>PRICE LIST — {nombre_promo.upper()}</b>", styles['Title']),
            Paragraph(f"Fecha: {datetime.now().strftime('%d/%m/%Y')}  ·  Unidades: {len(unidades)}", styles['Normal']),
            Spacer(1, 0.4*cm),
            t,
        ])
        return True
    except Exception as e:
        log(f"  ERROR PDF {nombre_promo}: {e}"); return False


def run_primeinvest(explorar=False):
    log("\n" + "═"*58)
    log("  PRIME INVEST")
    log("═"*58)
    descargados, errores, sin_listado = [], [], []

    with sync_playwright() as pw:
        browser, context, page = _nuevo_browser(pw)

        # ── Login ──────────────────────────────────────────────────────────────
        page.goto(PRIMEINVEST_URL_LOGIN, wait_until="domcontentloaded", timeout=30_000)
        time.sleep(0.8)   # modal de cookies carga asíncrono en iframe

        # Cerrar modal de cookies (está en iframe de terceros)
        modal_ok = False
        for frame in page.frames:
            for sel in ['[id*="accept"]', 'button:has-text("Aceptar")',
                        'button:has-text("Accept")', '[class*="accept"]']:
                try:
                    loc = frame.locator(sel)
                    if loc.count() > 0:
                        loc.first.click(timeout=2_000)
                        log(f"  ✓ Modal cookies cerrado")
                        modal_ok = True; break
                except Exception:
                    pass
            if modal_ok: break
        if not modal_ok:
            page.keyboard.press("Escape")

        # Rellenar credenciales con IDs confirmados
        page.locator('#emailInput').fill(PRIMEINVEST_USUARIO)
        page.locator('#passwordInput').fill(PRIMEINVEST_PASSWORD)
        page.locator('button:has-text("Acceder")').click()

        try:
            page.wait_for_selector('a[href*="promotion"]', timeout=15_000)
        except Exception:
            page.wait_for_load_state("networkidle", timeout=15_000)
        log(f"  Login OK → {page.url}")

        # ── Ir a Promociones ───────────────────────────────────────────────────
        page.locator('a[href*="promotion"]').first.click()
        page.wait_for_selector('table', timeout=10_000)
        url_lista = page.url
        log(f"  Promociones → {url_lista}")

        # ── Extraer links de todas las promociones (paginando) ─────────────────
        _poner_100(page)

        seen, promos_unicas = set(), []
        while True:
            for p in page.evaluate(_JS_LINKS_PI):
                if p['href'] not in seen and p['href'] != url_lista:
                    seen.add(p['href']); promos_unicas.append(p)
            if not _siguiente(page): break
            page.wait_for_load_state("domcontentloaded", timeout=8_000)

        log(f"  {len(promos_unicas)} promociones")

        if explorar:
            for i, p in enumerate(promos_unicas, 1):
                print(f"  {i:2d}. {p['nombre']}")
            browser.close(); return [], [], []

        # ── Entrar en cada promoción y raspar unidades ─────────────────────────
        por_promo: dict = defaultdict(list)
        for idx, promo in enumerate(promos_unicas, 1):
            nombre = promo['nombre'] or f"Promo_{idx}"
            log(f"  [{idx}/{len(promos_unicas)}] {nombre}")
            try:
                page.goto(promo['href'], wait_until="domcontentloaded", timeout=15_000)
                page.wait_for_selector('table', timeout=8_000)
                _poner_100(page)

                filas = []
                while True:
                    filas.extend(page.evaluate(_JS_RASPAR_PI, nombre))
                    if not _siguiente(page): break
                    page.wait_for_load_state("domcontentloaded", timeout=8_000)

                if filas:
                    por_promo[nombre].extend(filas)
                    log(f"    ✓ {len(filas)} unidades")
                else:
                    log(f"    ⬜ Sin tabla"); sin_listado.append(nombre)
            except Exception as e:
                log(f"    ERROR: {e}"); errores.append(f"PrimeInvest/{nombre}: {e}")

        # ── Generar PDFs ───────────────────────────────────────────────────────
        for nombre_promo, unidades in sorted(por_promo.items()):
            destino = CARPETA_DESTINO / f"Price_List_Prime_Invest_{FECHA}_{limpiar(nombre_promo)}.pdf"
            if destino.exists() and destino.stat().st_size > 1000:
                log(f"  ↩ Ya existe: {destino.name}")
                descargados.append(str(destino))
            elif _generar_pdf_pi(nombre_promo, unidades, destino):
                log(f"  ✓ PDF: {destino.name}")
                descargados.append(str(destino))
            else:
                errores.append(f"PrimeInvest/{nombre_promo}: error PDF")

        browser.close()
    return descargados, errores, sin_listado


# ══════════════════════════════════════════════════════════════════════════════
#  PORTAL 6 — ON3
# ══════════════════════════════════════════════════════════════════════════════

def run_on3(explorar=False):
    log("\n" + "═"*58)
    log("  ON3 PORTAL BROKERS")
    log("═"*58)
    descargados, errores, sin_listado = [], [], []

    with sync_playwright() as pw:
        browser, context, page = _nuevo_browser(pw)

        # ── Login ──────────────────────────────────────────────────────────────
        page.goto(ON3_URL_LOGIN, wait_until="domcontentloaded", timeout=30_000)

        # Cookie modal (DOM principal, clase cmplz-accept)
        try:
            btn_cookie = page.locator("button.cmplz-accept")
            btn_cookie.wait_for(state="visible", timeout=5_000)
            btn_cookie.click()
            log("  ✓ Cookie modal aceptado")
        except Exception:
            pass

        page.locator('input[name="log"]').first.fill(ON3_USUARIO)
        page.locator('input[name="pwd"]').first.fill(ON3_PASSWORD)
        page.locator('button:has-text("Acceder")').first.click()
        try:
            page.wait_for_url("**/brokers/**", timeout=15_000)
        except Exception:
            pass
        log(f"  Login OK → {page.url}")

        if "on3.es" not in page.url:
            log("  !! Login fallido"); browser.close()
            return [], ["ON3: login fallido"], []

        # ── Cargar panel de brokers ────────────────────────────────────────────
        if "brokers" not in page.url:
            page.goto(ON3_URL_BROKERS, wait_until="domcontentloaded", timeout=20_000)

        # Scroll para lazy-load de todas las tarjetas
        altura_prev = 0
        for _ in range(20):
            page.evaluate("window.scrollBy(0, 2000)")
            time.sleep(0.2)
            altura = page.evaluate("document.body.scrollHeight")
            if altura == altura_prev:
                break
            altura_prev = altura

        # ── Extraer todas las promociones ──────────────────────────────────────
        promociones = page.evaluate("""
            () => {
                const seen  = new Set();
                const named = {};
                const all   = [];
                for (const a of document.querySelectorAll('a[href*="/promocion/"]')) {
                    const href   = a.href.split('?')[0].split('#')[0].replace(/\\/+$/, '') + '/';
                    const nombre = a.textContent.trim();
                    if (!seen.has(href)) { seen.add(href); all.push(href); }
                    if (nombre && !named[href]) named[href] = nombre;
                }
                return all.map(href => ({
                    href,
                    nombre: named[href] ||
                            href.split('/').filter(Boolean).pop()
                                .replace(/-broker$/, '').replace(/-/g, ' ')
                                .replace(/\\b\\w/g, c => c.toUpperCase())
                }));
            }
        """)

        if not promociones:
            log("  !! No se encontraron promociones")
            browser.close()
            return [], ["ON3: no se encontraron promociones"], []

        log(f"  {len(promociones)} promociones encontradas")

        if explorar:
            for i, p in enumerate(promociones, 1):
                print(f"  {i:2d}. {p['nombre']:<45}  {p['href']}")
            browser.close(); return [], [], []

        # ── Procesar cada promoción ────────────────────────────────────────────
        for idx, promo in enumerate(promociones, 1):
            nombre_promo = promo['nombre']
            url_promo    = promo['href']
            log(f"  ── [{idx}/{len(promociones)}] {nombre_promo}")

            try:
                page.goto(url_promo, wait_until="domcontentloaded", timeout=20_000)

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

                listados = [p for p in pdfs
                            if es_precio(p['href']) or es_precio(p['txt'])]

                if not listados:
                    log(f"    ⬜ Sin listado de precios")
                    sin_listado.append(nombre_promo)
                    continue

                for pdf in listados:
                    href          = pdf['href']
                    nombre_doc    = href.split('/')[-1].split('?')[0] or 'doc.pdf'
                    nombre_limpio = limpiar(nombre_promo)[:45]
                    destino       = CARPETA_DESTINO / \
                        f"Price_List_ON3_{FECHA}_{nombre_limpio}_{nombre_doc}"

                    if guardar_pdf(context, href, destino):
                        descargados.append(str(destino))
                    else:
                        errores.append(f"ON3/{nombre_promo}: descarga fallida")

            except Exception as e:
                log(f"    ERROR: {e}"); errores.append(f"ON3/{nombre_promo}: {e}")

        browser.close()
    return descargados, errores, sin_listado


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

PORTALES = {
    "bromley":     run_bromley,
    "dream":       run_dream,
    "magnum":      run_magnum,
    "mxm":         run_mxm,
    "primeinvest": run_primeinvest,
    "on3":         run_on3,
}

def main():
    parser = argparse.ArgumentParser(description="Descargador unificado KH Inmobiliaria")
    parser.add_argument("--portal", nargs="+", choices=PORTALES.keys(), metavar="PORTAL",
                        help="bromley dream magnum mxm primeinvest (defecto: todos)")
    parser.add_argument("--explorar", action="store_true",
                        help="Lista promociones sin descargar")
    args = parser.parse_args()

    portales_a_ejecutar = args.portal or list(PORTALES.keys())

    print()
    print("═"*58)
    print("  Descargador unificado — KH Inmobiliaria")
    if args.explorar:
        print("  MODO EXPLORACIÓN")
    else:
        print(f"  Portales: {', '.join(p.upper() for p in portales_a_ejecutar)}")
    print("═"*58)

    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)

    total_desc, total_err, total_sin = [], [], []

    if args.explorar or len(portales_a_ejecutar) == 1:
        # Exploración o portal único: ejecución secuencial
        for nombre_portal in portales_a_ejecutar:
            desc, err, sin = PORTALES[nombre_portal](explorar=args.explorar)
            total_desc.extend(desc)
            total_err.extend(err)
            total_sin.extend(sin)
    else:
        # Varios portales: ejecución en paralelo
        log(f"\n  Ejecutando {len(portales_a_ejecutar)} portales en paralelo...")
        with ThreadPoolExecutor(max_workers=len(portales_a_ejecutar)) as ex:
            futuros = {ex.submit(PORTALES[n], False): n for n in portales_a_ejecutar}
            for futuro in as_completed(futuros):
                try:
                    desc, err, sin = futuro.result()
                    total_desc.extend(desc)
                    total_err.extend(err)
                    total_sin.extend(sin)
                except Exception as e:
                    total_err.append(f"{futuros[futuro]}: error inesperado — {e}")

    if args.explorar:
        print(); input("Pulsa Enter para cerrar..."); return

    print()
    print("═"*58)
    print("  RESUMEN FINAL")
    print("═"*58)
    if total_desc:
        print(f"\n✓ Descargados ({len(total_desc)}):")
        for f in total_desc: print(f"    {Path(f).name}")
    else:
        print("\n  No se descargó ningún listado.")
    if total_sin:
        print(f"\n⬜ Sin listado ({len(total_sin)}): {', '.join(total_sin)}")
    if total_err:
        print(f"\n!! Errores ({len(total_err)}):")
        for e in total_err: print(f"    {e}")
    print()
    input("Pulsa Enter para cerrar...")


if __name__ == "__main__":
    main()
