@echo off
REM Outlook Desktop MCP - Launcher
REM Usage:
REM   outlook-desktop-mcp.cmd mcp    Start the MCP server (stdio)
REM   outlook-desktop-mcp.cmd test   Run COM validation tests

setlocal

REM Run from the install directory and keep the caller's working directory
REM off Python's module search path (-P), so a planted module in whatever
REM directory the MCP client starts us from can never be imported.
cd /d "%~dp0"

set PYTHON=%~dp0.venv\Scripts\python.exe

if not exist "%PYTHON%" (
    echo ERROR: Virtual environment not found. Run setup first. 1>&2
    exit /b 1
)

if "%1"=="mcp" (
    "%PYTHON%" -P -m outlook_desktop_mcp.server
) else if "%1"=="test" (
    "%PYTHON%" -P "%~dp0tests\phase1_com_test.py"
) else (
    echo Usage: outlook-desktop-mcp.cmd [mcp^|test] 1>&2
    exit /b 1
)
