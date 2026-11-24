@echo off
rem Cierra opencode y los servidores de este MCP (no cierra Excel).
cd /d "%~dp0"
set "PATH=%APPDATA%\npm;%PATH%"
node scripts\cerrar-mcp.mjs
pause
