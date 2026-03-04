# Triple Duty Bond Generator

Automatically extract data from Bill of Entry PDFs and generate a formatted Triple Duty Bond Word Document. The application conforms to the Nagarkot Brand Standard.

## Tech Stack
- Python 3.10+
- `tkinter` (GUI)
- `pdfplumber` (Data extraction)
- `python-docx` (Document template rendering)
- `Pillow` (Image/logo rendering)

---

## Installation

### Clone
```bash
git clone https://github.com/username/Triple-Duty-Bond-Generator.git
cd Triple-Duty-Bond-Generator
```

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. Create virtual environment
```cmd
python -m venv venv
```

2. Activate (REQUIRED)

Windows:
```cmd
venv\Scripts\activate
```

Mac/Linux:
```bash
source venv/bin/activate
```

3. Install dependencies
```cmd
pip install -r requirements.txt
```

4. Run application
```cmd
python generate_bond.py
```

---

### Build Executable (For Desktop Apps)

1. Ensure the virtual environment is activated and dependencies installed (including PyInstaller):
```cmd
pip install pyinstaller
```

2. Build using the included Spec file (Ensure you do not run main.py directly):
```cmd
pyinstaller Triple_Duty_Bond_Generator.spec
```

3. Locate Executable:
The standalone application will be generated in the `dist/` folder. The templates and optional logo are bundled inside the `.exe`.

---

## Environment Variables

```cmd
cp .env.example .env
```
(Modify if required for future APIs)

---

## Notes
- **ALWAYS use virtual environment for Python.**
- Do not commit `venv` or `__pycache__`.
- Run and test before pushing.
