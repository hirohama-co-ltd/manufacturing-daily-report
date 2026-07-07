@echo off
setlocal
REM clasp は UNC を cwd にできないため、一旦 TEMP から pushd で NAS をドライブにマップする
cd /d %TEMP%
pushd "%~dp0"
if errorlevel 1 (
  echo pushd failed: %~dp0
  exit /b 1
)
echo PWD: %CD%
echo.
echo === clasp push ===
call clasp.cmd push --force
if errorlevel 1 (
  popd
  exit /b 1
)
echo.
echo === clasp deploy ===
call clasp.cmd deploy -i AKfycbxdfVghWyrDFN5QTqpKRm9OrV4bsnlr0SwHLgsU54Mxw_UUVuvmsPCoSgv1nqjxosFy -d "SwitchoverC9Master"
set ERR=%ERRORLEVEL%
popd
exit /b %ERR%
