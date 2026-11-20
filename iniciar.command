#!/bin/bash
# Abre la interfaz en macOS. Doble clic o: ./iniciar.command
cd "$(dirname "$0")" || exit 1
if [ -f ".venv/bin/activate" ]; then
  # shellcheck disable=SC1091
  source .venv/bin/activate
fi
python3 app.py
