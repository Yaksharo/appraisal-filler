@echo off
REM Build a self-contained Windows executable. Run inside the project folder.
REM Needs: Python 3 from python.org with "py launcher". Output: dist\AdviseeDocFiller.exe
py -m pip install -r requirements.txt pyinstaller
py -m PyInstaller --onedir --windowed --name AdviseeDocFiller ^
  --add-data "templates_docx;templates_docx" ^
  --add-data "assets;assets" ^
  --icon assets\logo.ico ^
  --hidden-import docxcompose ^
  --collect-data docxcompose ^
  gui_app.py
echo Done. App folder at dist\AdviseeDocFiller (run AdviseeDocFiller.exe inside it)
pause
