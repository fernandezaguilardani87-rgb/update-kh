#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# Property Manager — KH Inmobiliaria Costa del Sol
# Lanzador para macOS (doble-clic en el Finder)
# ─────────────────────────────────────────────────────────────────────────────

# Ir al directorio donde está este script
cd "$(dirname "$0")"

# Verificar Python 3
if ! command -v python3 &> /dev/null; then
    osascript -e 'display dialog "Python 3 no está instalado.\n\nDescárgalo desde https://www.python.org" buttons {"OK"} with icon stop'
    exit 1
fi

# Verificar e instalar dependencias si hacen falta
python3 -c "import pdfplumber, openpyxl, pandas" 2>/dev/null || {
    osascript -e 'display dialog "Instalando dependencias (una sola vez)…\nEsto puede tardar unos segundos." buttons {"OK"} default button 1'
    pip3 install pdfplumber openpyxl pandas --break-system-packages --quiet 2>/dev/null || \
    pip3 install pdfplumber openpyxl pandas --quiet
}

# Lanzar la aplicación
python3 property_manager.py
