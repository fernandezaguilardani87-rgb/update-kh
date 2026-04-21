#!/usr/bin/env python3
"""
Descargador automático — Dream Exclusives Agents Portal
Descarga los PDFs de "Downloads" de cada desarrollo (solo listados de precios).

Uso: python descargar_dream.py
"""

import sys
import os
import time
from pathlib import Path
from datetime import datetime

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

URL_LOGIN      = "https://agentsportal.dreamexclusives.com/index.php"
URL_DESARROLLOS = "https://agentsportal.dreamexclusives.com/index.php?opt=Promociones"
USUARIO        = "info@kalantarihomes.com"
PASSWORD       = "kalantarihomes2023#"

# URL de Downloads de cada desarrollo
# Patrón: ?opt=Promociones&buscar={ID}&tab=downloads
# El ID numérico se obtiene abriendo cada desarrollo en el portal y mirando la URL.
# Capri confirmado = buscar=24. El resto hay que verificarlos.
DESARROLLOS = [
    {
        "nombre": "Capri",
        "url_downloads": "https://agentsportal.dreamexclusives.com/index.php?opt=Promociones&buscar=24&tab=downloads",
    },
    {
        "nombre": "Adel San Roque",
        "url_downloads": "https://agentsportal.dreamexclusives.com/index.php?opt=Promociones&buscar=26&tab=downloads",
    },
    {
        "nombre": "The Palms At Estepona",
        "url_downloads": "https://agentsportal.dreamexclusives.com/index.php?opt=Promociones&buscar=39&tab=downloads",
    },
]

# Carpeta donde se guardan los PDFs
CARPETA_DESTINO = Path(r"C:\Users\User\Docs\Update\Update KH\Listados de Precios")

# ─── PALABRAS CLAVE PARA IDENTIFICAR LISTADOS DE PRECIOS ─────────────────────

KEYWORDS_PRECIO = [
    'price list', 'pricelist', 'price-list', 'price_list',
    'lista de precio', 'lista precios', 'listado de precio',
    'lista_de_precio', 'lista_precios',
    'tarifa', 'precios',
]

def es_listado_precios(nombre_archivo):
    # Normalizar: guiones y guiones bajos → espacio, para detectar
    # tanto "Price_List" como "Price-List" como "Price List"
    n = nombre_archivo.lower().replace('_', ' ').replace('-', ' ')
    return any(kw.replace('_', ' ').replace('-', ' ') in n for kw in KEYWORDS_PRECIO)

# ─── INSTALACIÓN AUTOMÁTICA DE PLAYWRIGHT ────────────────────────────────────

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    print("Instalando playwright (solo la primera vez)...")
    os.system(f'"{sys.executable}" -m pip install playwright --quiet')
    os.system(f'"{sys.executable}" -m playwright install chromium')
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def ts():
    return datetime.now().strftime("%H:%M:%S")

def log(msg):
    print(f"[{ts()}] {msg}")

def limpiar_nombre(texto):
    return "".join(c if c.isalnum() or c in " -_." else "_" for c in texto).strip()

# ─── LÓGICA PRINCIPAL ─────────────────────────────────────────────────────────

def main():
    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)
    fecha       = datetime.now().strftime("%Y-%m-%d")
    descargados = []
    errores     = []
    sin_listado = []

    with sync_playwright() as p:

        log("Abriendo navegador Chromium...")
        browser = p.chromium.launch(headless=False, slow_mo=300)
        context = browser.new_context(accept_downloads=True)
        page    = context.new_page()

        # ── 1. LOGIN ──────────────────────────────────────────────────────────
        log(f"Abriendo portal: {URL_LOGIN}")
        page.goto(URL_LOGIN, wait_until="networkidle", timeout=30_000)

        try:
            # Campo usuario
            for sel in ['input[name="email"]', 'input[name="username"]',
                        'input[name="user"]', 'input[type="email"]',
                        'input[name="log"]', 'input[type="text"]']:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    loc.first.fill(USUARIO)
                    log(f"  Usuario rellenado ({sel})")
                    break

            # Campo contraseña
            for sel in ['input[name="password"]', 'input[name="pwd"]',
                        'input[name="pass"]', 'input[type="password"]']:
                loc = page.locator(sel)
                if loc.count() > 0 and loc.first.is_visible():
                    loc.first.fill(PASSWORD)
                    log(f"  Contraseña rellenada ({sel})")
                    break

            # Botón de login
            boton_ok = False
            for sel in ['button[type="submit"]', 'input[type="submit"]',
                        'button:has-text("Sign in")', 'button:has-text("Login")',
                        'button:has-text("Entrar")', 'input[value="Sign in"]',
                        'input[value="Login"]']:
                loc = page.locator(sel)
                if loc.count() > 0:
                    for i in range(loc.count()):
                        if loc.nth(i).is_visible():
                            loc.nth(i).click(timeout=6_000)
                            boton_ok = True
                            log(f"  Botón login pulsado ({sel})")
                            break
                if boton_ok:
                    break

            if not boton_ok:
                page.locator('input[type="password"]').first.press("Enter")
                log("  Formulario enviado con Enter")

            page.wait_for_load_state("networkidle", timeout=20_000)
            log(f"  URL tras login: {page.url}")
            log("Login completado")

        except Exception as e:
            log(f"ERROR en login: {e}")
            browser.close()
            return [], [f"Login fallido: {e}"], []

        # ── 2. PROCESAR CADA DESARROLLO ───────────────────────────────────────
        for dev in DESARROLLOS:
            nombre_dev   = dev["nombre"]
            url_downloads = dev["url_downloads"]

            log(f"\n{'─'*55}")
            log(f"Desarrollo: {nombre_dev}")
            log(f"URL: {url_downloads}")

            try:
                page.goto(url_downloads, wait_until="networkidle", timeout=20_000)
                time.sleep(1.0)

                # ── Recoger todos los archivos disponibles ──────────────────
                archivos = page.evaluate("""
                    () => {
                        const links = document.querySelectorAll('a[href]');
                        const result = [];
                        for (const a of links) {
                            const href = a.href || '';
                            const nombre = a.textContent.trim() ||
                                           a.getAttribute('download') ||
                                           href.split('/').pop();
                            // Solo enlaces que parezcan archivos descargables
                            if (href && (
                                href.includes('.pdf') || href.includes('.PDF') ||
                                href.includes('download') || href.includes('Download') ||
                                a.getAttribute('download') !== null
                            )) {
                                result.push({nombre, href});
                            }
                        }
                        return result;
                    }
                """)

                log(f"  {len(archivos)} archivos disponibles en Downloads")

                # ── Filtrar: solo listados de precios ───────────────────────
                listados = [a for a in archivos if es_listado_precios(a['nombre'])]
                descartados = [a['nombre'] for a in archivos if not es_listado_precios(a['nombre'])]

                if descartados:
                    log(f"  Descartados (no son listas de precios): {len(descartados)}")
                    for d in descartados:
                        log(f"    · {d[:70]}")

                if not listados:
                    log(f"  ⚠  Sin listado de precios — contactar promotora")
                    sin_listado.append(nombre_dev)
                    continue

                # ── Descargar solo listados de precios ──────────────────────
                for arch in listados:
                    nombre_doc   = limpiar_nombre(arch['nombre'])
                    href         = arch['href']
                    archivo_dest = CARPETA_DESTINO / f"Price_List_Dream_Exclusives_{fecha}_{nombre_doc}.pdf"

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
                        errores.append(f"{nombre_dev} [{nombre_doc}]: {e}")

            except Exception as e:
                log(f"  ERROR procesando {nombre_dev}: {e}")
                errores.append(f"{nombre_dev}: {e}")

        browser.close()

    return descargados, errores, sin_listado


# ─── PUNTO DE ENTRADA ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    print()
    print("=" * 58)
    print("  Descargador de precios — Dream Exclusives")
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

    if errores:
        print(f"\nErrores técnicos ({len(errores)}):")
        for e in errores:
            print(f"  !! {e}")

    print()
    input("Pulsa Enter para cerrar...")
