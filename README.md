# CRISCO — Replacement Management App (Python + Access)

Desktop application for managing academic schedule replacements ("заміни") with a simple UI and MS Access database as a data source.

## Features

- **Create replacement documents** in Word format (`.docx`)
- **Data loading** from `dataBase.mdb` (teachers, groups, subjects, rooms, departments)
- **Academic year / month structure** for organizing generated documents
- **Search/autocomplete-like workflow** for faster input

## Tech stack

- **Python** (Tkinter UI)
- **MS Access (.mdb)** via **ODBC** (`pyodbc`)
- **Word document generation** (`python-docx`)
- Optional calendar widget (`tkcalendar`)

## Requirements

- Windows 7/8/10/11
- Python 3.8+
- Microsoft Access Database Engine (ODBC driver)
  - https://www.microsoft.com/en-us/download/details.aspx?id=54920

## Run (dev)

```bash
pip install -r requirements.txt
python main.py
```

> Important: `dataBase.mdb` must be in the same folder as `main.py` (or the application expects it next to the executable in distribution mode).

## Build EXE (PyInstaller)

```bash
pip install pyinstaller
pyinstaller Crisco_Optimized.spec
```

The executable will be produced in `dist/`.

## Repository notes

- `build/`, `dist/`, `*.exe` are intentionally ignored.
- Folders with generated replacement documents (`Zaminy/`, `Заміни/`) are ignored.

## Troubleshooting

- **Cannot connect to the database**
  - Ensure `dataBase.mdb` is present next to the app
  - Install Microsoft Access Database Engine

## Screenshots

Add screenshots here (recommended for CV):

- `docs/screenshots/main.png`
- `docs/screenshots/form.png`

---

## Українською (коротко)

CRISCO — програма для ведення замін з UI на Tkinter та базою даних Access (`dataBase.mdb`). Генерує документи замін у форматі Word (`.docx`).
