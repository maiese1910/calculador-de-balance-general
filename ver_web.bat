@echo off
echo ===================================================
echo   Iniciando Calculador de Balance (Modo Web)
echo ===================================================
echo.
echo   Se abrira tu navegador automaticamente...
echo   Presiona Ctrl+C para detener el servidor.
echo.
call npx http-server ./web -p 8080 -o
pause
