@echo off
echo === Kappa Windows Build ===
echo.

where uv >nul 2>&1 || (
    echo ERROR: uv not found.
    echo Install it from: https://docs.astral.sh/uv/getting-started/installation/
    pause
    exit /b 1
)

echo Installing / updating dependencies...
uv sync --group dev
if errorlevel 1 ( echo Dependency install failed. & pause & exit /b 1 )

echo.
echo Building executable...
uv run pyinstaller kappa.spec --clean
if errorlevel 1 ( echo Build FAILED. See output above. & pause & exit /b 1 )

echo.
echo ============================================================
echo  Build complete!
echo  Output folder : dist\Kappa\
echo  Run with      : dist\Kappa\Kappa.exe
echo  Distribute by zipping the entire dist\Kappa\ folder.
echo ============================================================
pause
