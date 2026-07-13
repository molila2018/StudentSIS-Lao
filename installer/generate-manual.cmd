@echo off
REM ═══════════════════════════════════════════════════════════════════════
REM  Generate the StudentSIS Lao user manual (PDF, in Lao).
REM
REM  Prerequisites (one-time):
REM    • .NET 8 SDK
REM    • Microsoft Word  — OR — LibreOffice (for .docx → .pdf conversion)
REM
REM  Output:
REM    installer\User-Manual.pdf
REM ═══════════════════════════════════════════════════════════════════════

setlocal
pushd "%~dp0\.."
set "PROJECT_ROOT=%CD%"

echo === Building integration-tests project (contains the manual generator) ===
dotnet build "%PROJECT_ROOT%\tests\IntegrationTests\IntegrationTests.csproj" -c Debug -nologo -v q
if errorlevel 1 (
    echo *** Build failed ***
    popd
    exit /b 1
)

echo.
echo === Generating user manual ===
dotnet run --project "%PROJECT_ROOT%\tests\IntegrationTests\IntegrationTests.csproj" --no-build -- manual "%PROJECT_ROOT%\installer\User-Manual.pdf"
if errorlevel 1 (
    echo *** Manual generation failed ***
    popd
    exit /b 2
)

echo.
echo ═══════════════════════════════════════════════════════════
echo  Done. Open the PDF to verify:
echo    %PROJECT_ROOT%\installer\User-Manual.pdf
echo ═══════════════════════════════════════════════════════════
popd
endlocal
