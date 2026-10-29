@echo off
rem Builds Decksterity: the add-in (bin\Release) and its installer (dist\DecksteritySetup-<version>.msi).
rem Needs Visual Studio 2022 or the VS 2022 Build Tools with the "Office/SharePoint build tools" workload,
rem and internet access on the first run (the WiX installer SDK comes from NuGet).
rem Everything is logged to build.log.
setlocal
cd /d "%~dp0"
set "VERSION=1.1.0"

set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
set "MSBUILD="
for /f "usebackq tokens=*" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe`) do set "MSBUILD=%%i"
if not defined MSBUILD (
  echo MSBuild was not found. Install Visual Studio 2022 or its Build Tools. > build.log
  goto :failed
)
echo Using %MSBUILD% > build.log

rem The add-in's manifests must be signed; this reuses or creates a local self-signed certificate
set "THUMB="
for /f "usebackq tokens=*" %%t in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\make-signing-cert.ps1"`) do set "THUMB=%%t"
if not defined THUMB (
  echo Could not find or create the signing certificate. >> build.log
  goto :failed
)
echo Signing certificate %THUMB% >> build.log

echo. >> build.log
echo === Building the add-in >> build.log
"%MSBUILD%" decksterity.csproj /t:Rebuild /p:Configuration=Release /p:ManifestKeyFile= /p:ManifestCertificateThumbprint=%THUMB% /v:minimal /nologo >> build.log 2>&1
if errorlevel 1 goto :failed

echo. >> build.log
echo === Building the installer >> build.log
"%MSBUILD%" Setup\DecksteritySetup.wixproj /restore /t:Rebuild /p:Configuration=Release /p:DecksterityVersion=%VERSION% /v:minimal /nologo >> build.log 2>&1
if errorlevel 1 goto :failed
if not exist "dist\DecksteritySetup-%VERSION%.msi" (
  echo The installer was not produced. >> build.log
  goto :failed
)

echo. >> build.log
echo BUILD SUCCEEDED: dist\DecksteritySetup-%VERSION%.msi >> build.log
exit /b 0

:failed
echo. >> build.log
echo BUILD FAILED - see the errors above >> build.log
exit /b 1
