@echo off
setlocal
cd /d "%~dp0"

echo ====================================
echo ECharts editor build tool
echo ====================================
echo.

echo [1/4] Installing build/runtime dependencies...
python -m pip install --disable-pip-version-check -r requirements.txt pyinstaller
if errorlevel 1 goto failed
echo Dependencies are ready.
echo.

echo [2/4] Cleaning old build files...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo Clean complete.
echo.

echo [3/4] Building executable...
python -m PyInstaller echarts_editor_build.spec --clean
if errorlevel 1 goto failed
echo.

if exist "dist\*.exe" (
    echo [4/4] Build succeeded.
    for %%F in ("dist\*.exe") do echo Executable: %%~fF
    echo.
    pause
    explorer dist
    exit /b 0
)

echo Build finished but no executable was found in dist.

:failed
echo.
echo [4/4] Build failed.
echo Please check the error messages above.
pause
exit /b 1
