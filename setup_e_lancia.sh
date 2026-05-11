#!/bin/bash
# Setup e lancio BDE Assignment Tool (Mac/Linux)
# Esegui con: bash setup_e_lancia.sh

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_DIR="$SCRIPT_DIR/.venv"

echo "================================"
echo " BDE Assignment Tool - Setup"
echo "================================"

# Crea virtual environment se non esiste
if [ ! -d "$VENV_DIR" ]; then
    echo "Creo virtual environment..."
    python3 -m venv "$VENV_DIR"
    echo "Installo dipendenze..."
    "$VENV_DIR/bin/pip" install --quiet pandas openpyxl
    echo "Setup completato!"
else
    echo "Virtual environment trovato."
fi

echo ""
echo "Lancio BDE Assignment Tool..."
echo "Si apre nel browser: http://127.0.0.1:8787"
echo "Per chiudere: Ctrl+C"
echo ""

"$VENV_DIR/bin/python3" "$SCRIPT_DIR/auto_bde_gui.py"
