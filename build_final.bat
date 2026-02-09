@echo off
setlocal enabledelayedexpansion

echo Cleaning previous builds...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del CMDA_Scraper.spec 2>nul

echo Installing dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ❌ Failed to install dependencies
    pause
    exit /b 1
)

echo Installing Playwright browsers...
python -m playwright install
if errorlevel 1 (
    echo ❌ Failed to install Playwright browsers
    pause
    exit /b 1
)

echo Building executable...
python -m PyInstaller --onefile ^
  --add-data ".env;." ^
  --add-data "zoho_tokens.json;." ^
  --add-data "ExistData.xlsx;." ^
  --add-data "loader.gif;." ^
  --add-data "client_logo.png;." ^
  --add-data "%USERPROFILE%\AppData\Local\ms-playwright;ms-playwright" ^
  --hidden-import="playwright._impl._generated" ^
  --hidden-import="playwright._impl" ^
  --hidden-import="playwright.sync_api" ^
  --hidden-import="helper" ^
  --hidden-import="extractor" ^
  --hidden-import="pdf_report" ^
  --hidden-import="approved_letter" ^
  --hidden-import="Integration" ^
  --hidden-import="ZohoCRMAutomatedAuth" ^
  --hidden-import="requests" ^
  --hidden-import="urllib3" ^
  --hidden-import="urllib.parse" ^
  --hidden-import="io" ^
  --hidden-import="pathlib" ^
  --name "CMDA_Scraper" ^
  main.py

if errorlevel 1 (
    echo ❌ PyInstaller build failed
    echo Checking for error logs...
    if exist build\CMDA_Scraper\warn-CMDA_Scraper.txt (
        echo === BUILD WARNINGS ===
        type build\CMDA_Scraper\warn-CMDA_Scraper.txt
    )
    pause
    exit /b 1
)

echo Checking dist folder...
dir dist

if exist "dist\CMDA_Scraper.exe" (
    echo ✅ Build successful! Executable created.
    echo File size: 
    for %%F in ("dist\CMDA_Scraper.exe") do echo %%~zF bytes
    
    echo Copying data files to dist folder...
    copy "ExistData.xlsx" "dist\" >nul
    copy ".env" "dist\" >nul
    copy "zoho_tokens.json" "dist\" >nul
    copy "loader.gif" "dist\" >nul
    copy "client_logo.png" "dist\" >nul
    echo ✅ Data files copied to dist folder!
) else (
    echo ❌ Build failed - No executable found in dist folder
    echo Checking build directory...
    dir build
    pause
    exit /b 1
)

echo Copying browsers to dist folder...
if exist "%USERPROFILE%\AppData\Local\ms-playwright" (
    xcopy "%USERPROFILE%\AppData\Local\ms-playwright" "dist\ms-playwright" /E /I /H /Y
    echo ✅ Browsers copied successfully!
) else (
    echo ❌ Browser source path not found!
)

echo.
echo ✅ Build complete! Check 'dist' folder for CMDA_Scraper.exe
echo.
echo Important: The entire 'dist' folder must be distributed together.
pause