# OFD Exceptions Report Generator

A small Windows desktop application that converts an incomplete-report Excel export into a polished exceptions report for distribution.

## Report types

- Draft Reports
- In Review Reports

The selected report type is shown in the exported Excel report title.

## Output columns

- CAD Incident Number
- Incident Date / Time
- Shift
- Station
- Units
- Address

The application does **not infer or fill missing incident information**. Blank source values remain blank.

The source export can provide the address under either `Address` or `Location`. HTML `<br/>` tags in the Units field are converted to readable line breaks for presentation only.

## Downloading the Windows app

GitHub Actions builds a standalone Windows executable.

1. Open the **Actions** tab in this repository.
2. Select **Build Windows EXE**.
3. Open the most recent successful run.
4. Under **Artifacts**, download **OFD-Exceptions-Report-Generator-Windows**.
5. Extract the ZIP.
6. Double-click **OFD Exceptions Report Generator.exe**.

No Python installation is required on the computer running the finished executable.

## Building manually

The GitHub workflow uses Python 3.12, openpyxl, and PyInstaller.
