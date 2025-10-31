@echo off
REM  Downloads a file from GitHub.  Saves locally, not wherever this .bat file is in path.

echo.
echo  =========================== REMINDER: To Github URLs, README.md != readme.md =============================
echo  USAGE: utilget.bat (to prompt), or utilget.bat [subfolder\filename] where subfolder to create is optional.
echo  ==========================================================================================================

REM First enabling delayedexpansion only here, only after != in echo above, so display of ! character not misprocessed.
setlocal EnableExtensions EnableDelayedExpansion

REM Get file name
set "FILENAME=%~1"

if not "%FILENAME%"=="" goto :havefile

echo  [94m[INPUT][0m [93mEnter the file name to download from utilgets:[0m
set /p FILENAME=

:havefile
REM Covers error if merely enter was pressed when prompted for a file name
if "%FILENAME%"=="" (
  echo  [31m[ERROR][0m No file name provided. Nothing to do.
  goto :eof
)

set "BASE=https://raw.githubusercontent.com/Sikoat/utilgets/refs/heads/main"

REM Compute local save path (create subfolder if implied)
set "LOCALPATH=%FILENAME:/=\%"
for %%I in ("%LOCALPATH%") do (
  set "DIR=%%~dpI"
  set "LEAF=%%~nxI"
)
if not "%DIR%"=="" if not exist "%DIR%" (
  echo  [94m[INFO ][0m Creating directory: [32m%DIR%[0m
  mkdir "%DIR%" >nul 2>nul
)

REM Build URL using only the leaf filename (no added subfolder in the URL)
set "URL=%BASE%/%LEAF%"

echo  [94m[INFO ][0m Source URL: [36m%URL%[0m
for %%I in ("%LOCALPATH%") do echo  [94m[INFO ][0m Save as:    [93m%%~fI[0m

echo  [31m[WORK ][0m Starting download with curl ...

REM Use -f to fail on HTTP errors, -L to follow redirects
curl -f -L --retry 2 --connect-timeout 20 -o "%LOCALPATH%" "%URL%"

REM Get byte count (0 if file missing)
if exist "%LOCALPATH%" (
  for %%Z in ("%LOCALPATH%") do set "BYTES=%%~zZ"
) else (
  set "BYTES=0"
)

REM Show success only when nonzero bytes.  Avoids easy error in batch script if ( ) else ( ) style code.
if not "%BYTES%"=="0" echo  [92m[SUCCESS][0m Downloaded [93m%LOCALPATH%[0m  [92m(%BYTES% bytes)[0m
if not "%BYTES%"=="0" goto :success

REM Line immediately below runs if no success, as then not jumped over:
echo [91m[FAIL][0m Download failed or file empty: %LOCALPATH%

:success

endlocal
