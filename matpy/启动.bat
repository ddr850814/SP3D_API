@echo off
rem Launcher: prefer packaged exe in dist\, fallback to Python source.
if exist "%~dp0dist" (
    for %%f in ("%~dp0dist\*.exe") do (
        start "" "%%f"
        exit /b
    )
)
start "" pythonw "%~dp0main.py"
