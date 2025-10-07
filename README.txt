PhanTichDAO - PyQt5 Windows App (implements VBA logic)
------------------------------------------------------
What this contains:
- Main PyQt5 app: main.py
- requirements.txt (pyqt5, python-docx, openpyxl)
- build.bat to build exe locally using PyInstaller
- .github/workflows/windows-build.yml to build exe on GitHub Actions (windows-latest)

How to run (dev):
1) Create and activate a venv (recommended)
   python -m venv venv
   venv\Scripts\activate
2) pip install -r requirements.txt
3) python main.py
4) To build standalone exe locally (Windows): run build.bat (installs pyinstaller and builds)

How to build on GitHub Actions:
- Create a GitHub repo, push this project.
- GitHub Actions will run the workflow in .github/workflows/windows-build.yml on push.
- The workflow builds the exe and uploads artifact 'phantichdao-windows-exe' (dist\PhanTichDAO.exe).

Notes:
- The app lets user pick any folder containing .docx files.
- It reads paragraphs, extracts digits, keeps last 3 digits, normalizes by sorting digits ascending,
  groups counts for mocs [5,10,15,20], shows Top3 for each moc, and a combined comparison table.
- Export to CSV and Excel (.xlsx) available.
