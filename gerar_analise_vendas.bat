@echo off
title Analise de Vendas (relatorio unico) - Trovidoce
echo A instalar dependencias...
pip install pandas openpyxl reportlab pymupdf -q
echo.
echo A gerar relatorio unico (Analise_Vendas.pdf)...
python "%~dp0gerar_analise_vendas.py"
echo.
pause
