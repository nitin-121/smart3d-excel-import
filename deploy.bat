@echo off
echo ============================================
echo  Smart3D Excel Import - Build ^& Deploy
echo ============================================
echo.

:: === CONFIGURATION ===
:: Try multiple possible Smart3D V14 install paths
set SMART3D_HOME=C:\Program Files\Intergraph\SmartPlant 3D V14
if not exist "%SMART3D_HOME%" set SMART3D_HOME=C:\Program Files\Hexagon\SmartPlant 3D V14
if not exist "%SMART3D_HOME%" set SMART3D_HOME=C:\Program Files\Intergraph\Smart 3D V14
if not exist "%SMART3D_HOME%" set SMART3D_HOME=C:\Program Files\Intergraph\SP3D V14

set BUILD_CONFIG=Release
set SRC_DIR=%~dp0src
set OUTPUT_DIR=%SRC_DIR%\bin\%BUILD_CONFIG%
set DLL_NAME=Smart3D.ExcelImport.dll

echo Smart3D Home:  %SMART3D_HOME%
echo Build Config:  %BUILD_CONFIG%
echo Output Dir:    %OUTPUT_DIR%
echo.

:: === STEP 1: RESTORE ===
echo [1/5] Restoring NuGet packages...
cd /d %SRC_DIR%
dotnet restore
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] NuGet restore failed!
    pause
    exit /b 1
)
echo [OK] Packages restored.
echo.

:: === STEP 2: BUILD ===
echo [2/5] Building project...
dotnet build -c %BUILD_CONFIG% --no-restore
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Build failed!
    pause
    exit /b 1
)
echo [OK] Build succeeded.
echo.

:: === STEP 3: VERIFY OUTPUT ===
if not exist "%OUTPUT_DIR%\%DLL_NAME%" (
    echo [ERROR] Output DLL not found: %OUTPUT_DIR%\%DLL_NAME%
    echo Available files in output dir:
    dir "%OUTPUT_DIR%\*.dll" 2>nul
    pause
    exit /b 1
)

:: === STEP 4: DEPLOY TO SMART3D ===
echo [3/5] Checking Smart3D installation...
if not exist "%SMART3D_HOME%\bin" (
    echo [ERROR] Smart3D not found at: %SMART3D_HOME%
    echo.
    echo Please edit this script and set the correct SMART3D_HOME path.
    echo Common locations:
    echo   C:\Program Files\Intergraph\SmartPlant 3D V14
    echo   C:\Program Files\Hexagon\SmartPlant 3D V14
    pause
    exit /b 1
)

echo [4/5] Copying DLLs to Smart3D...
copy /Y "%OUTPUT_DIR%\%DLL_NAME%" "%SMART3D_HOME%\" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [WARN] Could not copy to Smart3D home. Trying bin\...
    copy /Y "%OUTPUT_DIR%\%DLL_NAME%" "%SMART3D_HOME%\" >nul 2>&1
)
copy /Y "%OUTPUT_DIR%\EPPlus.dll" "%SMART3D_HOME%\" >nul 2>&1
copy /Y "%OUTPUT_DIR%\Serilog.dll" "%SMART3D_HOME%\" >nul 2>&1
copy /Y "%OUTPUT_DIR%\Serilog.Sinks.File.dll" "%SMART3D_HOME%\" >nul 2>&1
echo [OK] DLLs copied.
echo.

:: === STEP 5: REGISTER COMMAND FILE ===
echo [5/5] Copying command registration...
if exist "%~dp0CommandRegistration.xml" (
    if not exist "%SMART3D_HOME%\CommandFiles" mkdir "%SMART3D_HOME%\CommandFiles" 2>nul
    copy /Y "%~dp0CommandRegistration.xml" "%SMART3D_HOME%\CommandFiles\" >nul 2>&1
    echo [OK] Command registration copied.
) else (
    echo [WARN] CommandRegistration.xml not found in script directory.
)

echo.
echo ============================================
echo  Deployment Complete!
echo ============================================
echo.
echo Output DLL: %OUTPUT_DIR%\%DLL_NAME%
echo.
echo Next steps:
echo  1. Restart Smart3D V14
echo  2. Go to Tools ^> Custom Commands
echo  3. Click "Bulk Property Import from Excel"
echo  4. Select your Excel file and import!
echo.
echo Logs: %%APPDATA%%\Smart3DExcelImporter\
echo.
pause
