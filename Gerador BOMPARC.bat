@echo off
cd /d "%~dp0"
echo Iniciando o Gerador de Relatorios BOMPARC...
echo Por favor, aguarde enquanto a janela do navegador e aberta.
python -m streamlit run form_app.py
pause
