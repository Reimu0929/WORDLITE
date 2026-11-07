@echo off
setlocal EnableExtensions

rem ---------------------------------------------------------------------------
rem  Build WordPadLite.WinUI on Windows from sources stored on WSL.
rem
rem  Building directly from \\wsl.localhost\... fails because the Windows SDK
rem  manifest tool (mt.exe) cannot handle UNC paths (error C1010070).
rem  This script mirrors the sources to a local folder first, then builds.
rem
rem  Usage (from Windows cmd/PowerShell):
rem      build-windows.cmd [Debug|Release]
rem
rem  Optional environment variables:
rem      WORDLITE_BUILD_ROOT  destination folder
rem                           default: %USERPROFILE%\source\repos\WORDLITE-build
rem ---------------------------------------------------------------------------

set "SRC=%~dp0"
if "%SRC:~-1%"=="\" set "SRC=%SRC:~0,-1%"

if not defined WORDLITE_BUILD_ROOT set "WORDLITE_BUILD_ROOT=%USERPROFILE%\source\repos\WORDLITE-build"
set "BUILD_ROOT=%WORDLITE_BUILD_ROOT%"

set "CONFIG=%~1"
if not defined CONFIG set "CONFIG=Debug"

where dotnet >nul 2>&1
if errorlevel 1 (
    echo [ERROR] dotnet was not found in PATH.
    exit /b 1
)

echo Source : %SRC%
echo Destination : %BUILD_ROOT%
echo Configuration : %CONFIG%
echo.

echo [1/3] Mirroring sources...
if not exist "%BUILD_ROOT%" mkdir "%BUILD_ROOT%" >nul 2>&1
robocopy "%SRC%" "%BUILD_ROOT%" /MIR /XD bin obj .git .vs .idea /XF *.user *.suo /NFL /NDL /NP
set "RC=%ERRORLEVEL%"
if %RC% GEQ 8 (
    echo.
    echo [ERROR] Failed to mirror the sources ^(robocopy exit code %RC%^).
    exit /b 1
)

echo [2/3] Stopping a running WordPadLite.exe...
taskkill /IM WordPadLite.exe /F >nul 2>&1

echo [3/3] Building...
pushd "%BUILD_ROOT%\WordPadLite.WinUI"
dotnet build -c %CONFIG% -p:Platform=x64 -v:m
set "RC=%ERRORLEVEL%"
popd

if not "%RC%"=="0" (
    echo.
    echo [ERROR] Build failed.
    exit /b %RC%
)

echo.
echo Build succeeded.
echo   %BUILD_ROOT%\WordPadLite.WinUI\bin\x64\%CONFIG%\net10.0-windows10.0.19041.0\win-x64\WordPadLite.exe
exit /b 0
