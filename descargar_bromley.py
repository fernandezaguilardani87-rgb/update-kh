#!/usr/bin/env python3
"""
Descargador automático — Bromley Estates Agency Portal
Descarga los PDFs de la pestaña "Availability" de cada desarrollo.

Uso: python descargar_bromley.py
"""

import sys
import os
import time
from pathlib import Path
from datetime import datetime

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

URL_LOGIN = "https://ap.bromleyestatesmarbella.com/welcome/"
USUARIO   = "info@kalantarihomes.com"
PASSWORD  = "kalantarihomes2023#"

# URL directa de cada desarrollo + nombre limpio para el archivo descargado
# Patrón: https://ap.bromleyestatesmarbella.com/development-doc/[slug]/
DESARROLLOS = [
    {
        "nombre": "Infinity Homes",
        "url": "https://ap.bromleyestatesmarbella.com/development-doc/las-mesas-infinity-homes/",
    },
    {
        "nombre": "Las Mesas Collection",
        "url": "https://ap.bromleyestatesmarbella.com/development-doc/las-mesas-collection/",
    },
    {
        "nombre": "Las Mesas Sea Suites",
        "url": "https://ap.bromleyestatesmarbella.com/development-doc/las-mesas-sea-suites/",
    },
]

# Carpeta donde se guardan los PDFs
CARPETA_DESTINO = Path(r"C:\Users\User\Docs\Update\Update KH\Listados de Precios")

# ─── INSTALACIÓN AUTOMÁTICA DE PLAYWRIGHT ────────────────────────────────────

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    print("Instalando playwright (solo la primera vez)...")
    os.system(f'"{sys.executable}" -m pip install playwright --quiet')
    os.system(f'"{sys.executable}" -m playwright install chromium')
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ─── PALABRAS CLAVE PARA IDENTIFICAR LISTADOS DE PRECIOS ─────────────────────
# Solo se descargan archivos cuyo nombre contenga alguna de estas palabras clave
# (sin distinguir mayúsculas/minúsculas ni acentos)

KEYWORDS_PRECIO = [
    'price list', 'pricelist', 'price-list', 'price_list',
    'lista de precio', 'lista precios', 'listado de precio',
    'lista_de_precio', 'lista_precios',
    'tarifa', 'precios',
]

def es_listado_precios(nombre_archivo):
    """True si el nombre del archivo parece una lista de precios."""
    n = nombre_archivo.lower().replace('_', ' ').replace('-', ' ')
    return any(kw.replace('_', ' ').replace('-', ' ') in n for kw in KEYWORDS_PRECIO)

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def ts():
    return datetime.now().strftime("%H:%M:%S")

def log(msg):
    print(f"[{ts()}] {msg}")

def limpiar_nombre(texto):
    """Convierte texto en nombre de archivo seguro."""
    return "".join(c if c.isalnum() or c in " -_." else "_" for c in texto).strip()

# ─── LÓGICA PRINCIPAL ─────────────────────────────────────────────────────────

def main():
    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)
    fecha = datetime.now().strftime("%Y-%m-%d")
    descargados   = []
    errores       = []
    sin_listado   = []   # desarrollos sin ningún PDF de precios (contactar promotora)

    with sync_playwright() as p:

        # ── Abrir navegador ───────────────────────────────────────────────────
        log("Abriendo navegador Chromium...")
        browser = p.chromium.launch(headless=False, slow_mo=300)
        context = browser.new_context(accept_downloads=True)
        page    = context.new_page()

        # ── 1. LOGIN ──────────────────────────────────────────────────────────
        log(f"Abriendo portal: {URL_LOGIN}")
        page.goto(URL_LOGIN, wait_until="networkidle", timeout=30_000)

        try:
            # Campo usuario — probar varios selectores + fallback a cualquier input texto visible
            usuario_ok = False
            for sel in ['input[name="log"]', 'input[name="user_login"]',
                        'input[name="email"]', 'input[name="username"]',
                        'input[type="email"]', 'input[type="text"]']:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    loc.first.fill(USUARIO)
                    log(f"  Usuario rellenado ({sel})")
                    usuario_ok = True
                    break
            if not usuario_ok:
                raise Exception("No se encontró el campo de usuario/email en la página")

            # Campo contraseña
            for sel in ['input[name="pwd"]', 'input[name="user_pass"]',
                        'input[name="password"]', 'input[type="password"]']:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    loc.first.fill(PASSWORD)
                    log(f"  Contraseña rellenada ({sel})")
                    break

            # Botón LOGIN — buscar el visible con valor LOGIN
            # (hay 2 input[type=submit]: uno de búsqueda oculto y uno de login)
            boton_ok = False
            for sel in ['input[value="LOGIN"]', 'input[value="Login"]',
                        'input[value="Log In"]', 'input[name="wp-submit"]',
                        '#wp-submit',
                        'button:has-text("LOGIN")', 'button:has-text("Login")']:
                loc = page.locator(sel)
                if loc.count() > 0:
                    for i in range(loc.count()):
                        item = loc.nth(i)
                        if item.is_visible():
                            item.click(timeout=6_000)
                            boton_ok = True
                            log(f"  Botón LOGIN pulsado ({sel})")
                            break
                if boton_ok:
                    break

            if not boton_ok:
                # Último recurso: Enter en el campo de contraseña
                page.locator('input[type="password"]').first.press("Enter")
                log("  Formulario enviado con Enter")

            page.wait_for_load_state("networkidle", timeout=20_000)

            # Verificar login: si la URL cambió o aparece el menú de usuario, es correcto
            # No usamos presencia de input[password] porque la página puede mantenerlo
            log(f"  URL tras login: {page.url}")
            log("Login completado")

        except Exception as e:
            log(f"ERROR en login: {e}")
            browser.close()
            return [], [f"Login fallido: {e}"], []

        # ── 2. PROCESAR CADA DESARROLLO ───────────────────────────────────────
        for dev in DESARROLLOS:
            nombre = dev["nombre"]
            url    = dev["url"]
            log(f"\n{'─'*55}")
            log(f"Desarrollo: {nombre}")
            log(f"URL: {url}")

            try:
                page.goto(url, wait_until="networkidle", timeout=20_000)

                # ── Clic en pestaña "Availability" ─────────────────────────────
                # Esperar a que el tab sea visible y pulsarlo
                try:
                    page.wait_for_selector('text="Availability"', timeout=8_000)
                    page.get_by_text("Availability", exact=True).first.click()
                    time.sleep(1.2)  # esperar que el contenido del tab cargue
                    log("  Pestaña Availability abierta")
                except Exception as e_tab:
                    log(f"  AVISO: no se pudo abrir pestaña Availability: {e_tab}")

                # ── Recoger archivos del panel ACTIVO de la pestaña ──────────────
                # El portal usa el plugin "Shortcodes Ultimate" de WordPress:
                # el panel abierto tiene la clase extra "su-tabs-pane-open"
                time.sleep(1.5)

                archivos = page.evaluate("""
                    () => {
                        // Panel activo = .su-tabs-pane-open (Shortcodes Ultimate)
                        // Fallback: cualquier .su-tabs-pane con display != none
                        let panel = document.querySelector('.su-tabs-pane-open');
                        if (!panel) {
                            for (const p of document.querySelectorAll('.su-tabs-pane')) {
                                if (window.getComputedStyle(p).display !== 'none') {
                                    panel = p; break;
                                }
                            }
                        }
                        if (!panel) return [];

                        const rows = panel.querySelectorAll('table tbody tr');
                        const result = [];
                        for (const row of rows) {
                            const cells = row.querySelectorAll('td');
                            if (cells.length < 4) continue;
                            const link = cells[3].querySelector('a');
                            if (link && link.href) {
                                result.push({nombre: cells[0].textContent.trim(), href: link.href});
                            }
                        }
                        return result;
                    }
                """)

                log(f"  {len(archivos)} archivos en Availability")

                # ── Filtrar: solo listados de precios ──────────────────────────
                listados = [a for a in archivos if es_listado_precios(a['nombre'])]
                descartados = [a['nombre'] for a in archivos if not es_listado_precios(a['nombre'])]

                if descartados:
                    log(f"  Descartados (no son listas de precios): {len(descartados)}")
                    for d in descartados:
                        log(f"    · {d[:70]}")

                if not listados:
                    # No hay listado de precios → puede estar todo vendido o sin disponibilidad
                    log(f"  ⚠  Sin listado de precios en Availability — contactar promotora")
                    sin_listado.append(nombre)
                    continue

                # ── Descargar solo los listados de precios ─────────────────────
                for arch in listados:
                    nombre_doc    = limpiar_nombre(arch['nombre'])
                    href          = arch['href']
                    archivo_dest  = CARPETA_DESTINO / f"Price_List_Bromley_Estates_{fecha}_{nombre_doc}.pdf"

                    log(f"  Descargando: {arch['nombre'][:65]}...")

                    try:
                        response = context.request.get(href, timeout=30_000)
                        if response.ok:
                            archivo_dest.write_bytes(response.body())
                            log(f"  Guardado: {archivo_dest.name}")
                            descargados.append(str(archivo_dest))
                        else:
                            raise Exception(f"HTTP {response.status} — {href}")
                    except Exception as e:
                        log(f"  ERROR: {e}")
                        errores.append(f"{nombre} [{nombre_doc}]: {e}")

            except Exception as e:
                log(f"  ERROR procesando {nombre}: {e}")
                errores.append(f"{nombre}: {e}")

        browser.close()

    return descargados, errores, sin_listado


# ─── PUNTO DE ENTRADA ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    print()
    print("=" * 58)
    print("  Descargador de precios — Bromley Estates")
    print("=" * 58)
    print()

    descargados, errores, sin_listado = main()

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
        print(f"  ⚠  CONTACTAR PROMOTORA — sin listado disponible ({len(sin_listado)}):")
        print(f"{'─'*58}")
        for d in sin_listado:
            print(f"  !!  {d}")
        print()
        print("  Puede estar todo vendido o el listado no publicado aún.")
        print("  Marca estas filas en rojo claro en el Excel maestro.")

    if errores:
        print(f"\nErrores técnicos ({len(errores)}):")
        for e in errores:
            print(f"  !! {e}")
        print()
        print("Revisa las capturas _debug_*.png en la carpeta.")

    print()
    input("Pulsa Enter para cerrar...")
