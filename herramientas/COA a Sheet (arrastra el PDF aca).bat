@echo off
REM Arrastra un PDF de COA sobre este archivo para generar la cadena de la columna "coa".
REM Tambien podes hacer doble click y elegir el PDF en el selector.
py "%~dp0coa_a_sheet.py" %*
