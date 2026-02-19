Readme here...

## User guide — Windows executable

This repository includes a packaged Windows executable option for the CTP processor. The executable presents a small GUI where users can select the input cadet CSV/Excel file and choose the output file name and destination.

- **Run the EXE:** Double-click the produced `ctp_processor.exe` (or run from Command Prompt). If you built with `--windowed` it runs without a console window.
- **Select input file:** Browse to your cadet data file (CSV, .xls, .xlsx). The default input is the `downloads/Cadet Qualifications Report.csv` file if present.
- **Select output file:** Choose a destination and filename for the generated tracker (default `CTP_Tracker.xlsx`).
- **Master module list:** If your master module file is named `Module_list.csv` (or `module_list.csv`), include it alongside the exe or in the same working folder so the program can find it. If you used a one-file bundle, include it using PyInstaller's `--add-data` when building.

### Building an executable (recommended commands)

From the project root using the project's virtualenv, install requirements and run PyInstaller:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m PyInstaller --onefile --windowed ctp_processor.py
```

If you prefer a distributable folder (easier to bundle data files):

```powershell
.venv\Scripts\python.exe -m PyInstaller --onedir --windowed --add-data "Module_list.csv;." --add-data "downloads;downloads" ctp_processor.py
```

Notes:
- `--onefile` creates a single EXE (extracts at runtime). Use `--onedir` to create a folder with the EXE and additional files.
- Ensure `Module_list.csv` and any required data files are included via `--add-data` or placed next to the exe.
- Test the generated exe from the `dist\` folder before sharing.

## Downloading raw data

Describe how to obtain the raw cadet data (CSV/Excel) here — e.g. download from the cadet database or export from the reporting system. Include steps, required permissions, and any filters to apply. (Fill in your organisation's specific instructions.)
