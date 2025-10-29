@echo off
echo Building Tatli Siparis application...

:: Check if Python is available
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Python is not found! Please install Python and add it to PATH
    pause
    exit /b 1
)

:: Install required packages if not already installed
echo Checking/installing required packages...
python -m pip install --upgrade pip
python -m pip install pyinstaller
python -m pip install -r requirements.txt

:: Clean previous builds
echo Cleaning previous builds...
if exist "build" rd /s /q "build"
if exist "dist" rd /s /q "dist"

:: Create the executable with PyInstaller
echo Creating executable...
pyinstaller --clean ^
    --name "Tatlı Sipariş" ^
    --icon "appicon.ico" ^
    --add-data "appicon.ico;." ^
    --noconsole ^
    --onedir ^
    --win-private-assemblies ^
    tatli_siparis.py

:: Check if build was successful
if %ERRORLEVEL% neq 0 (
    echo Build failed!
    pause
    exit /b 1
)

:: Copy additional required files to dist folder
echo Copying additional files...
copy requirements.txt "dist\Tatlı Sipariş\"
copy parse_gptfix.py "dist\Tatlı Sipariş\"
copy shipment_oop.py "dist\Tatlı Sipariş\"
if exist "appicon.ico" copy "appicon.ico" "dist\Tatlı Sipariş\"

echo Build completed successfully!
echo The executable is in the dist folder.
pause