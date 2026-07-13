@echo off
REM ═══════════════════════════════════════════════════════════════════════
REM  Build the StudentSIS Lao Windows installer end-to-end.
REM
REM  Prerequisites (one-time setup on the build machine):
REM    • .NET 8 SDK     — https://dotnet.microsoft.com/download
REM    • Inno Setup 6.x — https://jrsoftware.org/isinfo.php
REM
REM  Usage:
REM    Just double-click this file, or run from a cmd prompt:
REM        installer\build.cmd
REM
REM  Output:
REM    publish\                                     ← intermediate dotnet publish output
REM    installer-output\StudentSIS_Lao_Setup_v*.exe ← the installer (give this to users)
REM ═══════════════════════════════════════════════════════════════════════

setlocal enabledelayedexpansion

REM Walk up to the project root regardless of where this script was invoked from.
pushd "%~dp0\.."
set "PROJECT_ROOT=%CD%"
echo.
echo === Project root: %PROJECT_ROOT%
echo.

REM ── 1. dotnet publish ─────────────────────────────────────────────────
echo === Step 1/3: dotnet publish (Release, self-contained x64) ===
echo.
if exist "%PROJECT_ROOT%\publish" rmdir /s /q "%PROJECT_ROOT%\publish"

dotnet publish "%PROJECT_ROOT%\StudentSIS.csproj" ^
    -c Release ^
    -r win-x64 ^
    --self-contained true ^
    -p:PublishSingleFile=false ^
    -p:IncludeNativeLibrariesForSelfExtract=true ^
    -p:DebugType=none ^
    -p:DebugSymbols=false ^
    -o "%PROJECT_ROOT%\publish"

if errorlevel 1 (
    echo.
    echo *** dotnet publish FAILED ***
    popd
    exit /b 1
)

echo.
echo Publish output:
dir /b "%PROJECT_ROOT%\publish" | findstr /R "StudentSIS_Lao.exe"
if errorlevel 1 (
    echo *** publish folder is missing StudentSIS_Lao.exe ***
    popd
    exit /b 1
)
echo.

REM ── 2. Locate Inno Setup compiler ─────────────────────────────────────
echo === Step 2/3: Locating Inno Setup compiler (ISCC.exe) ===
set "ISCC="
where ISCC.exe >nul 2>&1
if not errorlevel 1 (
    for /f "delims=" %%I in ('where ISCC.exe') do set "ISCC=%%I"
    goto :iscc_found
)
REM Common install paths.
for %%P in (
    "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
    "%ProgramFiles%\Inno Setup 6\ISCC.exe"
    "%ProgramFiles(x86)%\Inno Setup 5\ISCC.exe"
    "%ProgramFiles%\Inno Setup 5\ISCC.exe"
) do (
    if exist %%P set "ISCC=%%~P"
)
if not defined ISCC (
    echo.
    echo *** Inno Setup compiler ISCC.exe not found. ***
    echo Install Inno Setup 6 from https://jrsoftware.org/isinfo.php
    echo and re-run this script.
    popd
    exit /b 2
)
:iscc_found
echo Using: %ISCC%
echo.

REM ── 3. Compile installer ──────────────────────────────────────────────
echo === Step 3/3: Compiling installer ===
if exist "%PROJECT_ROOT%\installer-output" rmdir /s /q "%PROJECT_ROOT%\installer-output"

REM Pull <Version> from StudentSIS.csproj so the installer's MyAppVersion +
REM output filename ALWAYS match the app version. Single source of truth =
REM csproj — no need to also bump the number in the .iss file.
for /f "delims=" %%V in ('powershell -NoProfile -Command "([xml](Get-Content '%PROJECT_ROOT%\StudentSIS.csproj')).Project.PropertyGroup.Version"') do set "APP_VERSION=%%V"
if "%APP_VERSION%"=="" (
    echo *** Could not read ^<Version^> from StudentSIS.csproj — falling back to .iss default ***
) else (
    echo Compiling installer for v%APP_VERSION%
)

"%ISCC%" /DMyAppVersion=%APP_VERSION% "%PROJECT_ROOT%\installer\StudentSIS_Lao.iss"
if errorlevel 1 (
    echo.
    echo *** Inno Setup compile FAILED ***
    popd
    exit /b 3
)

echo.
echo ═══════════════════════════════════════════════════════════
echo  SUCCESS — installer ready:
for %%F in ("%PROJECT_ROOT%\installer-output\*.exe") do echo    %%~F  (%%~zF bytes)
echo ═══════════════════════════════════════════════════════════
echo.

popd
endlocal
exit /b 0
