@echo off
:: =============================================================================
:: Build-Setup.bat  —  MPP Pipeline Setup Builder
:: =============================================================================
setlocal enabledelayedexpansion

set "SCRIPT_DIR=%~dp0"
if "!SCRIPT_DIR:~-1!"=="\" set "SCRIPT_DIR=!SCRIPT_DIR:~0,-1!"

set "CS_FILE=!SCRIPT_DIR!\MPP-Pipeline-Setup.cs"
set "OUT_EXE=!SCRIPT_DIR!\MPP-Pipeline-Setup.exe"
set "ICON_PATH="

cls
echo.
echo =====================================================
echo   MPP Pipeline  -  Setup Builder
echo =====================================================
echo.
echo   Working folder: !SCRIPT_DIR!
echo.

echo [1/4] Checking for MPP-Pipeline-Setup.cs...
if not exist "!CS_FILE!" (
    echo.
    echo ERROR: Cannot find MPP-Pipeline-Setup.cs
    echo        Expected at: !CS_FILE!
    echo.
    pause
    exit /b 1
)
echo       Found: !CS_FILE!
echo.

echo [2/4] Locating C# compiler (csc.exe)...
set "CSC="

for %%P in (
    "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
    "C:\Windows\Microsoft.NET\Framework64\v3.5\csc.exe"
    "C:\Windows\Microsoft.NET\Framework\v3.5\csc.exe"
) do (
    echo       Checking: %%~P
    if exist %%P (
        if "!CSC!"=="" set "CSC=%%~P"
    )
)

if "!CSC!"=="" (
    echo.
    echo ERROR: csc.exe not found.
    echo Please install .NET Framework 4.x from:
    echo   https://dotnet.microsoft.com/download/dotnet-framework
    echo.
    pause
    exit /b 1
)
echo       Found: !CSC!
echo.

echo [3/4] Compiling MPP-Pipeline-Setup.exe...
echo       (this usually takes 5-15 seconds)
echo.

set "ICON_ARG="
if not "!ICON_PATH!"=="" (
    if exist "!ICON_PATH!" (
        set "ICON_ARG=/win32icon:"!ICON_PATH!""
    ) else (
        echo       WARNING: Icon not found at !ICON_PATH! -- skipping.
    )
)

"!CSC!" /nologo /target:winexe /optimize+ ^
    /r:System.dll ^
    /r:System.Drawing.dll ^
    /r:System.Windows.Forms.dll ^
    /out:"!OUT_EXE!" ^
    !ICON_ARG! ^
    "!CS_FILE!"

set "COMPILE_RESULT=!ERRORLEVEL!"
echo.
echo       csc.exe exit code: !COMPILE_RESULT!

if !COMPILE_RESULT! neq 0 (
    echo.
    echo ERROR: Compilation failed. See output above.
    echo.
    pause
    exit /b 1
)

echo.
echo [4/4] Verifying output...

if not exist "!OUT_EXE!" (
    echo.
    echo ERROR: Output file not created.
    echo.
    pause
    exit /b 1
)

for %%A in ("!OUT_EXE!") do set "FSIZE=%%~zA"
set /a FSIZE_KB=!FSIZE! / 1024
echo       Created: !OUT_EXE!
echo       Size   : ~!FSIZE_KB! KB

echo.
echo =====================================================
echo   BUILD SUCCESSFUL
echo =====================================================
echo.
echo   Distribute MPP-Pipeline-Setup.exe to your users.
echo   They double-click it -- it will prompt for admin
echo   rights automatically, then install silently.
echo.
if "!ICON_PATH!"=="" (
    echo   REMINDER: When you have a .ico file, set ICON_PATH
    echo   at the top of this script and rebuild.
    echo.
)
pause
endlocal
