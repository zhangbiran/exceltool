@echo off
setlocal

set "EXCELTOOL_SOURCE=%USERPROFILE%\.local\src\exceltool"
set "LIBREOFFICE_PROGRAM=%ProgramFiles%\LibreOffice\program"

if not exist "%LIBREOFFICE_PROGRAM%\python.exe" (
  echo LibreOffice Python was not found: "%LIBREOFFICE_PROGRAM%\python.exe" 1>&2
  exit /b 4
)

if not exist "%EXCELTOOL_SOURCE%\exceltool" (
  echo ExcelTool source was not found: "%EXCELTOOL_SOURCE%\exceltool" 1>&2
  exit /b 4
)

set "SAL_DISABLE_SYNCHRONOUS_PRINTER_DETECTION=1"
set "SAL_DISABLE_PRINTERLIST=1"
set "SAL_DISABLE_DEFAULTPRINTER=1"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PATH=%LIBREOFFICE_PROGRAM%;%PATH%"

"%LIBREOFFICE_PROGRAM%\python.exe" "%EXCELTOOL_SOURCE%\exceltool" %*
exit /b %ERRORLEVEL%
