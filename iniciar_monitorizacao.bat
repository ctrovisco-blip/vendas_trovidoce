@echo off
title Monitorizar Mapas de Vendas - Trovidoce
echo A instalar dependencias...
pip install pandas openpyxl watchdog -q
echo.
echo A iniciar monitorizacao...
python monitorizar.py
pause
