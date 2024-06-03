@echo off
echo activating environment

set mypath=%cd%
echo %mypath%

call C:\temp\reporte-horas\.venv\Scripts\python.exe  C:\temp\reporte-horas\main.py
exit