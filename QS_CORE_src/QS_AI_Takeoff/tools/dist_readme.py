"""README_INSTALL.txt text for the colleague distribution."""


def write_readme(path: str, version: str, has_python: bool) -> None:
    py_line = ("Python 3.12 is installed automatically if missing (bundled installer, current user only)."
               if has_python else "Python 3.12 (64-bit) must be installed first: https://www.python.org")
    text = f"""DWG QS Tools {version}
QS CORE - the all-in-one QS platform (main window + Walls & Rooms Takeoff)
=====================================================================

INSTALL
  1. Extract the whole ZIP to any folder (do NOT run Setup from inside the ZIP).
  2. Double-click  Setup.bat
     - Installs to  %LOCALAPPDATA%\\DWG_QS_Tools   (no admin rights needed)
     - {py_line}
     - Packages are installed offline from the "wheels" folder (internet used only as fallback).
     - Creates Desktop + Start-menu shortcuts.
  3. Start  "QS CORE"  from the Desktop.
     Walls & Rooms opens from its tab (2D Takeoff > Plaster · Paint · Floor · Ceiling) or its own shortcut.

  Custom folder:  Setup.bat "D:\\Tools\\DWG_QS"
  Update:         run the new Setup.bat again - code is replaced, your config\\*.json settings are kept.
  Uninstall:      Start menu > DWG QS Tools > Uninstall   (or Uninstall.bat in the install folder)

AUTOCAD
  - DXF files work without AutoCAD.
  - DWG files: open the drawing in AutoCAD (COM), or APPLOAD the LISP files in
    %LOCALAPPDATA%\\DWG_QS_Tools\\QS_AI_Takeoff\\autocad   and run QS_SCAN.
  - Tested with AutoCAD 2026 on the author's PC only.

BEFORE USING QUANTITIES
  - Check Project Setup > Drawing units and Settings > wall layers for YOUR drawing.
  - Exported quantities are PROVISIONAL until walls are approved in Geometry Review.
  - Reconcile a few walls against a manual takeoff before relying on totals.

PROBLEMS
  Setup log:      %TEMP%\\DWG_QS_Tools_setup.log
  Program log:    <install folder>\\QS_AI_Takeoff\\logs\\launcher_stderr.txt
  Send these files to the tool owner.
"""
    with open(path, "w", encoding="utf-8", newline="\r\n") as fh:
        fh.write(text)
