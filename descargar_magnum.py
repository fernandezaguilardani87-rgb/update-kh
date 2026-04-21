#!/usr/bin/env python3
"""
Descargador automático — Magnum & Partners Portal Brokers
https://www.magnum-partners.com/brokers/

Flujo:
  1. Login con email + contraseña → botón ACCEDER
  2. Recorrer páginas de promociones (páginas 1–6)
  3. Para cada promoción que esté en DESARROLLOS (o todas si DESARROLLOS vacío):
       a. Clic en "Saber más" → página de detalle
       b. Buscar enlace "Planos y Precios" → página /brokers/planos/?promoid=XXX
       c. Si hay "Descargar lista de precios" → descargar PDF
       d. Si no → anotar como sin listado disponible

Uso:
    python descargar_magnum.py             — descarga los desarrollos de DESARROLLOS
                                             (si está vacío, intenta TODOS)
    python descargar_magnum.py --explorar  — lista todas las promociones del portal
                                             con sus URLs sin descargar nada
"""

import sys
import os
import re
import time
import argparse
from pathlib import Path
from datetime import datetime

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

URL_LOGIN        = "https://www.magnum-partners.com/brokers/"
URL_PROMOCIONES  = "https://www.magnum-partners.com/brokers/promociones/"
N_PAGINAS        = 6          # páginas 1–6 de promociones

USUARIO  = "info@kalantarihomes.com"
PASSWORD = "kalantarihomes2023#"

# Nombres de los desarrollos que quieres descargar.
# Se busca por coincidencia parcial (sin mayúsculas/minúsculas).
# Deja vacío para intentar TODOS los desarrollos del portal.
DESARROLLOS = [
    "Aby",              # → Aby Middle y/o Aby Upper en el portal
    "Adagio",
    "Aire",
    "Altara Alcaidesa",
    "Be Aloha",
    "Oceana Gardens",   # → Oceana Gardens I y Oceana Gardens II
    "Symphony Suites",  # → Symphony Suites Fase I y Fase II
    "Zenith",           # → Zenith Estepona
    "Zinnia",
]

# Mapeo nombre del portal → nombre en el Excel KH.
# Necesario cuando el nombre del portal difiere del nombre en el Excel maestro.
# Las variantes de "Aby" son el caso crítico: "Aby" solo tiene 3 letras y el
# matcher de property_manager ignora palabras cortas, así que necesitamos
# "Estepona" en el nombre del archivo para que haga la asociación correctamente.
MAPEO_NOMBRES = {
    "aby middle": "Aby Estepona",
    "aby upper":  "Aby Estepona",
}

# Carpeta donde se guardan los PDFs (misma que Bromley y Dream)
CARPETA_DESTINO = Path(r"C:\Users\User\Docs\Update\Update KH")

# ─── PALABRAS CLAVE PARA IDENTIFICAR LISTADOS DE PRECIOS ─────────────────────

KEYWORDS_PRECIO = [
    'price list', 'pricelist', 'price-list', 'price_list',
    'lista de precio', 'lista precios', 'listado de precio',
    'lista_de_precio', 'lista_precios',
    'tarifa', 'precios', 'descargar lista',
]

def es_listado_precios(texto: str) -> bool:
    n = texto.lower().replace('_', ' ').replace('-', ' ')
    return any(kw.replace('_', ' ').replace('-', ' ') in n for kw in KEYWORDS_PRECIO)

def nombre_coincide(nombre_portal: str, filtros: list) -> bool:
    """True si el nombre del portal coincide con alguno de los filtros (parcial, sin mayúsculas)."""
    if not filtros:
        return True   # sin filtro → aceptar todos
    np = nombre_portal.lower().strip()
    return any(f.lower().strip() in np or np in f.lower().strip() for f in filtros)

# ─── INSTALACIÓN AUTOMÁTICA DE PLAYWRIGHT ────────────────────────────────────

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    print("Instalando playwright (solo la primera vez)...")
    os.system(f'"{sys.executable}" -m pip install playwright --quiet')
    os.system(f'"{sys.executable}" -m playwright install chromium')
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def ts() -> str:
    return datetime.now().strftime("%H:%M:%S")

def log(msg: str) -> None:
    print(f"[{ts()}] {msg}")

def limpiar_nombre(texto: str) -> str:
    return "".join(c if c.isalnum() or c in " -_." else "_" for c in texto).strip()

# ─── 1. LOGIN ─────────────────────────────────────────────────────────────────

def hacer_login(page) -> bool:
    """Login en el portal. Devuelve True si tiene éxito."""
    log(f"Abriendo portal: {URL_LOGIN}")
    page.goto(URL_LOGIN, wait_until="networkidle", timeout=30_000)

    try:
        # Campo email/usuario
        for sel in ['input[name="email"]', 'input[name="username"]',
                    'input[name="user_login"]', 'input[name="log"]',
                    'input[type="email"]', 'input[type="text"]']:
            loc = page.locator(sel)
            if loc.count() > 0 and loc.first.is_visible():
                loc.first.fill(USUARIO)
                log(f"  Email rellenado ({sel})")
                break

        # Campo contraseña
        for sel in ['input[type="password"]', 'input[name="password"]',
                    'input[name="pwd"]', 'input[name="user_pass"]']:
            loc = page.locator(sel)
            if loc.count() > 0 and loc.first.is_visible():
                loc.first.fill(PASSWORD)
                log(f"  Contraseña rellenada ({sel})")
                break

        # Botón ACCEDER
        boton_ok = False
        for sel in [
            'input[value="Acceder"]', 'input[value="ACCEDER"]',
            'button:has-text("Acceder")', 'button:has-text("ACCEDER")',
            'button[type="submit"]', 'input[type="submit"]',
            'input[name="wp-submit"]',
            'button:has-text("Login")', 'button:has-text("Entrar")',
        ]:
            loc = page.locator(sel)
            if loc.count() > 0:
                for i in range(loc.count()):
                    if loc.nth(i).is_visible():
                        loc.nth(i).click(timeout=6_000)
                        boton_ok = True
                        log(f"  Botón pulsado ({sel})")
                        break
            if boton_ok:
                break

        if not boton_ok:
            page.locator('input[type="password"]').first.press("Enter")
            log("  Enter enviado (botón no encontrado)")

        page.wait_for_load_state("networkidle", timeout=25_000)
        time.sleep(1.5)   # margen extra para portales que no redirigen
        log(f"  URL tras login: {page.url}")

        # Verificar login: buscar señal de sesión activa.
        # Este portal puede quedarse en la misma URL con el formulario en el DOM
        # pero oculto por CSS, así que NO usamos :visible en el campo contraseña.
        # En su lugar buscamos elementos que solo aparecen con sesión abierta.
        sesion_ok = False

        # a) Enlace de logout / "Cerrar sesión" / "Mi cuenta"
        for sel in ['a[href*="logout"]', 'a[href*="log-out"]', 'a[href*="salir"]',
                    'a:has-text("Cerrar sesión")', 'a:has-text("Logout")',
                    'a:has-text("Mi cuenta")', 'a:has-text("Mi perfil")',
                    'a[href*="promociones"]', 'a[href*="brokers/"]']:
            if page.locator(sel).count() > 0:
                sesion_ok = True
                log(f"  Sesión detectada ({sel}) ✓")
                break

        # b) Mensaje de error explícito de WordPress ("contraseña incorrecta", etc.)
        error_wp = page.locator('.login-error, #login_error, .error, p.error').count()
        if error_wp > 0:
            msg = page.locator('.login-error, #login_error, .error, p.error').first.inner_text()
            log(f"  ERROR de login: {msg.strip()[:120]}")
            return False

        # c) Si no detectamos sesión pero tampoco error explícito → asumir OK
        #    (algunos portales no añaden clases reconocibles)
        if not sesion_ok:
            log("  Sesión no confirmada automáticamente — continuando de todas formas")

        log("  Login completado ✓")
        return True

    except Exception as e:
        log(f"  ERROR en login: {e}")
        return False

# ─── 2. RECOGER TODAS LAS PROMOCIONES ────────────────────────────────────────

def recoger_promociones(page) -> list[dict]:
    """
    Recorre las páginas 1–N_PAGINAS de promociones y devuelve lista de dicts:
        { nombre, url_detalle }
    """
    todas = []
    for num_pag in range(1, N_PAGINAS + 1):
        if num_pag == 1:
            url_pag = URL_PROMOCIONES
        else:
            url_pag = f"{URL_PROMOCIONES.rstrip('/')}/page/{num_pag}/"

        log(f"  Página {num_pag}: {url_pag}")
        try:
            page.goto(url_pag, wait_until="networkidle", timeout=20_000)
            time.sleep(1.0)

            # Comprobar si la página existe (si no hay promociones, es la última)
            promos_pag = page.evaluate("""
                () => {
                    // Las tarjetas de promoción tienen el nombre en un h2/h3/h4
                    // y el botón "Saber más" en un enlace
                    const cards = [];

                    // Buscar todos los enlaces que contengan "saber" en el texto
                    const links = document.querySelectorAll('a');
                    for (const a of links) {
                        const txt = a.textContent.trim().toLowerCase();
                        if (txt.includes('saber') || txt.includes('más') || txt.includes('ver')) {
                            // El nombre está en un elemento cercano (hermano o padre)
                            let nombre = '';
                            // Subir en el DOM buscando un título
                            let el = a;
                            for (let i = 0; i < 8 && !nombre; i++) {
                                el = el.parentElement;
                                if (!el) break;
                                const headings = el.querySelectorAll('h1,h2,h3,h4,h5,.title,.name');
                                if (headings.length > 0) {
                                    nombre = headings[0].textContent.trim();
                                }
                            }
                            if (a.href && nombre) {
                                cards.push({ nombre, url_detalle: a.href });
                            }
                        }
                    }
                    return cards;
                }
            """)

            if not promos_pag:
                log(f"    Sin promociones — fin de páginas")
                break

            log(f"    {len(promos_pag)} promociones encontradas")
            for p in promos_pag:
                log(f"    · {p['nombre']}")
            todas.extend(promos_pag)

        except Exception as e:
            log(f"    Error en página {num_pag}: {e}")

    # Deduplicar por URL y filtrar falsos positivos (nav items, saludos, etc.)
    EXCLUIR_NOMBRES = {'hola', 'hola,', 'perfil', 'mi perfil', 'cerrar sesión',
                       'logout', 'inicio', 'home', 'contacto'}
    EXCLUIR_URLS    = {'/perfil/', '/login/', '/logout/', '/mi-cuenta/', '/contacto/'}

    seen = set()
    unicas = []
    for p in todas:
        nombre_norm = p['nombre'].strip().lower().rstrip(',')
        url_lower   = p['url_detalle'].lower()

        # Ignorar si nombre es un saludo/nav o URL es de perfil/login
        if nombre_norm in EXCLUIR_NOMBRES:
            continue
        if any(excl in url_lower for excl in EXCLUIR_URLS):
            continue
        # Ignorar nombres muy cortos (< 3 caracteres) o que no sean desarrollos
        if len(nombre_norm) < 3:
            continue

        if p['url_detalle'] not in seen:
            seen.add(p['url_detalle'])
            unicas.append(p)
    return unicas

# ─── 3. PROCESAR UNA PROMOCIÓN ────────────────────────────────────────────────

def procesar_promocion(page, context, promo: dict, fecha: str,
                       descargados: list, errores: list, sin_listado: list) -> None:
    """
    Dada una promoción (nombre + url_detalle):
      a. Navegar a la página de detalle
      b. Encontrar el enlace "Planos y Precios" → /brokers/planos/?promoid=XXX
      c. En esa página buscar "Descargar lista de precios" y descargar el PDF
    """
    nombre       = promo['nombre']
    url_detalle  = promo['url_detalle']

    log(f"\n{'─'*55}")
    log(f"Promoción: {nombre}")
    log(f"URL: {url_detalle}")

    try:
        # ── a. Página de detalle ───────────────────────────────────────────────
        page.goto(url_detalle, wait_until="networkidle", timeout=20_000)
        time.sleep(1.0)

        # ── b. Buscar enlace "Planos y Precios" → /brokers/planos/?promoid=XXX ──
        url_planos = None

        # Prioridad 1: enlace cuya href contiene /planos/ (ruta específica del portal)
        planos_links = page.evaluate("""
            () => {
                const links = document.querySelectorAll('a[href]');
                const result = [];
                for (const a of links) {
                    if (a.href.includes('/planos/')) {
                        result.push(a.href);
                    }
                }
                return result;
            }
        """)
        if planos_links:
            url_planos = planos_links[0]
            log(f"  Enlace 'Planos y Precios' encontrado: {url_planos}")

        # Prioridad 2: texto del enlace contiene "planos" o "precios"
        if not url_planos:
            for sel in [
                'a:has-text("Planos y Precios")',
                'a:has-text("planos y precios")',
                'a:has-text("Planos")',
                'a:has-text("Precios")',
            ]:
                loc = page.locator(sel)
                if loc.count() > 0:
                    url_planos = loc.first.get_attribute('href')
                    log(f"  Enlace por texto ({sel}): {url_planos}")
                    break

        if not url_planos:
            log(f"  ⚠  No se encontró enlace a Planos y Precios")
            sin_listado.append(f"{nombre} (sin sección planos)")
            return

        # ── c. Navegar a la página de planos ─────────────────────────────────
        page.goto(url_planos, wait_until="networkidle", timeout=20_000)
        time.sleep(1.5)
        log(f"  Página planos: {page.url}")

        # ── d. Buscar botón "Descargar lista de precios" ──────────────────────
        # El portal Magnum muestra este botón específico cuando hay lista disponible.
        # Si no está → el desarrollo no tiene lista publicada.
        url_descarga = None

        # Prioridad 1: texto exacto del botón/enlace del portal
        descarga_links = page.evaluate("""
            () => {
                const links = document.querySelectorAll('a[href]');
                const result = [];
                for (const a of links) {
                    const txt = a.textContent.trim().toLowerCase();
                    const href = a.href.toLowerCase();
                    // "Descargar lista de precios" o variantes
                    if (
                        (txt.includes('descargar') && (txt.includes('precio') || txt.includes('lista'))) ||
                        (txt.includes('download') && (txt.includes('price') || txt.includes('list'))) ||
                        (txt.includes('price list')) ||
                        (txt.includes('lista de precios'))
                    ) {
                        result.push({ texto: a.textContent.trim(), href: a.href });
                    }
                }
                return result;
            }
        """)

        if descarga_links:
            url_descarga = descarga_links[0]['href']
            log(f"  Botón descarga: «{descarga_links[0]['texto']}» → {url_descarga[:80]}")
        else:
            # Prioridad 2: cualquier PDF en la página de planos
            # (algunos portales enlazan directamente al PDF sin texto estándar)
            pdf_links = page.evaluate("""
                () => {
                    const links = document.querySelectorAll('a[href]');
                    const result = [];
                    for (const a of links) {
                        if (a.href.toLowerCase().includes('.pdf')) {
                            result.push({ texto: a.textContent.trim(), href: a.href });
                        }
                    }
                    return result;
                }
            """)
            if pdf_links:
                mejor = next(
                    (l for l in pdf_links if es_listado_precios(l['texto'] + ' ' + l['href'])),
                    None   # No fallback a cualquier PDF — solo los que parezcan listas de precios
                )
                if mejor:
                    url_descarga = mejor['href']
                    log(f"  PDF precio encontrado: «{mejor['texto']}» → {url_descarga[:80]}")
                else:
                    log(f"  PDFs en planos: {[l['texto'] for l in pdf_links[:5]]} (ninguno parece lista de precios)")

        if not url_descarga:
            log(f"  ⬜ Sin lista de precios disponible (todo vendido o no publicado)")
            sin_listado.append(nombre)
            return

        # ── Descargar PDF ──────────────────────────────────────────────────────
        # Aplicar mapeo de nombre (portal → Excel KH) si existe
        nombre_kh      = MAPEO_NOMBRES.get(nombre.lower().strip(), nombre)
        nombre_archivo = limpiar_nombre(nombre_kh)
        # Formato: Price_List_{fecha}_{nombre_promocion}.pdf
        # → "Price_List" hace que property_manager lo reconozca como listado de precios
        # → nombre al final para identificarlo fácilmente
        archivo_dest   = CARPETA_DESTINO / f"Price_List_{fecha}_{nombre_archivo}.pdf"

        log(f"  Descargando → {archivo_dest.name}")
        try:
            response = context.request.get(url_descarga, timeout=30_000)
            if response.ok:
                archivo_dest.write_bytes(response.body())
                log(f"  ✓ Guardado: {archivo_dest.name}")
                descargados.append(str(archivo_dest))
            else:
                raise Exception(f"HTTP {response.status}")
        except Exception as e_dl:
            # Fallback: usar Playwright download si la petición directa falla
            log(f"  Descarga directa fallida ({e_dl}), intentando con navegador...")
            try:
                with page.expect_download(timeout=30_000) as dl_info:
                    page.locator(f'a[href="{url_descarga}"]').first.click()
                download = dl_info.value
                download.save_as(str(archivo_dest))
                log(f"  ✓ Guardado (navegador): {archivo_dest.name}")
                descargados.append(str(archivo_dest))
            except Exception as e_dl2:
                log(f"  ERROR: {e_dl2}")
                errores.append(f"{nombre}: {e_dl2}")

    except Exception as e:
        log(f"  ERROR procesando {nombre}: {e}")
        errores.append(f"{nombre}: {e}")

# ─── LÓGICA PRINCIPAL ─────────────────────────────────────────────────────────

def main(modo_explorar: bool = False):
    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)
    fecha       = datetime.now().strftime("%Y-%m-%d")
    descargados = []
    errores     = []
    sin_listado = []

    with sync_playwright() as p:
        log("Abriendo navegador Chromium...")
        browser = p.chromium.launch(headless=False, slow_mo=250)
        context = browser.new_context(accept_downloads=True)
        page    = context.new_page()

        # ── Login ──────────────────────────────────────────────────────────────
        if not hacer_login(page):
            browser.close()
            return [], ["Login fallido — revisa credenciales"], []

        # ── Recoger todas las promociones ──────────────────────────────────────
        log(f"\nRecorriendo páginas de promociones (1–{N_PAGINAS})...")
        todas_las_promos = recoger_promociones(page)
        log(f"\nTotal promociones encontradas en el portal: {len(todas_las_promos)}")

        if modo_explorar:
            print()
            print("─" * 58)
            print("  LISTADO COMPLETO DE PROMOCIONES EN MAGNUM & PARTNERS")
            print("─" * 58)
            for idx, p_item in enumerate(todas_las_promos, 1):
                print(f"  {idx:2d}. {p_item['nombre']:<40}  {p_item['url_detalle']}")
            print("─" * 58)
            print()
            print("Copia los nombres que quieras en la lista DESARROLLOS del script.")
            browser.close()
            return [], [], []

        # ── Filtrar según DESARROLLOS ──────────────────────────────────────────
        if DESARROLLOS:
            promos_a_procesar = [
                p_item for p_item in todas_las_promos
                if nombre_coincide(p_item['nombre'], DESARROLLOS)
            ]
            no_encontrados = [
                d for d in DESARROLLOS
                if not any(nombre_coincide(p_item['nombre'], [d]) for p_item in todas_las_promos)
            ]
            if no_encontrados:
                log(f"\n⚠  No encontrados en el portal: {', '.join(no_encontrados)}")
                errores.extend([f"{d}: no encontrado en el portal" for d in no_encontrados])
        else:
            promos_a_procesar = todas_las_promos
            log("  (Sin filtro — procesando TODAS las promociones)")

        log(f"  Promociones a procesar: {len(promos_a_procesar)}")

        # ── Procesar cada promoción ────────────────────────────────────────────
        for promo in promos_a_procesar:
            procesar_promocion(page, context, promo, fecha,
                               descargados, errores, sin_listado)

        browser.close()

    return descargados, errores, sin_listado


# ─── PUNTO DE ENTRADA ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Descargador Magnum & Partners")
    parser.add_argument("--explorar", action="store_true",
                        help="Lista todas las promociones del portal sin descargar")
    args = parser.parse_args()

    print()
    print("=" * 58)
    print("  Descargador de precios — Magnum & Partners")
    if args.explorar:
        print("  MODO EXPLORACIÓN")
    print("=" * 58)
    print()

    descargados, errores, sin_listado = main(modo_explorar=args.explorar)

    if args.explorar:
        input("Pulsa Enter para cerrar...")
        sys.exit(0)

    print()
    print("=" * 58)
    print("  RESULTADO")
    print("=" * 58)

    if descargados:
        print(f"\nListados de precios descargados ({len(descargados)}):")
        for f in descargados:
            print(f"  OK  {Path(f).name}")
    else:
        print("\nNo se descargó ningún listado de precios.")

    if sin_listado:
        print(f"\n{'─'*58}")
        print(f"  ⬜ Sin listado disponible ({len(sin_listado)}) — contactar promotora:")
        print(f"{'─'*58}")
        for d in sin_listado:
            print(f"  !!  {d}")
        print()
        print("  Puede estar todo vendido o el listado no publicado aún.")
        print("  Marca estas filas en rojo claro en el Excel maestro.")

    if errores:
        print(f"\nErrores ({len(errores)}):")
        for e in errores:
            print(f"  !! {e}")

    print()
    input("Pulsa Enter para cerrar...")
