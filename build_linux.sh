#!/bin/bash
# Build a self-contained Linux executable. Run inside the project folder.
# Needs: python3, python3-tk, pip. Output: dist/AdviseeDocFiller
set -e
pip install -r requirements.txt pyinstaller --break-system-packages 2>/dev/null \
  || pip install -r requirements.txt pyinstaller
pyinstaller --onedir --windowed --name AdviseeDocFiller \
  --add-data "templates_docx:templates_docx" \
  --add-data "assets:assets" \
  --hidden-import docxcompose \
  --collect-data docxcompose \
  gui_app.py
echo "Done. App folder at dist/AdviseeDocFiller (run dist/AdviseeDocFiller/AdviseeDocFiller)"
