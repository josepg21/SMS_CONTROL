@echo off
REM ============================================================
REM  Iniciar panel de seguimiento SMS en red local
REM  Las demas maquinas abren:  http://IP_DE_ESTA_PC:8501
REM ============================================================
echo ============================================================
echo  Seguimiento SMS - Servidor local
echo.
echo  Para que OTRA maquina vea el panel, abre en esa maquina:
echo  http://%%COMPUTERNAME%%:8501   (o la IP de este PC)
echo.
echo  Para detener: cierra esta ventana o presiona Ctrl+C
echo ============================================================
echo.
REM Mostrar IP del equipo
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do echo  Tu IP local es:%%a
echo.
REM Iniciar streamlit accesible desde la red
python -m streamlit run "%~dp0app.py" --server.address 0.0.0.0 --server.port 8501
pause
