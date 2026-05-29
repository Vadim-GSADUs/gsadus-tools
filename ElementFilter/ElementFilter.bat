@echo off
rem Launch the Revit Element Poison Finder (no console window).
rem Tries the windowed Python launchers first, falls back to console python.
where pyw >nul 2>&1 && (start "" pyw "%~dp0element_filter.py" & exit /b)
where pythonw >nul 2>&1 && (start "" pythonw "%~dp0element_filter.py" & exit /b)
where py >nul 2>&1 && (start "" py "%~dp0element_filter.py" & exit /b)
start "" python "%~dp0element_filter.py"
