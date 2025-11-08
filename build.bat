@echo off
setlocal
echo ---------------------------------------------------
echo  BUILDER DO APLICATIVO DE REDIRECIONAMENTO TRAY
echo ---------------------------------------------------

:: Verifica se o Python esta instalado
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: Python nao encontrado no PATH.
    echo Por favor, instale o Python 3 (python.org) e marque "Add to PATH".
    pause
    goto :eof
)

echo [1/4] Criando ambiente virtual (venv)...
python -m venv venv
if %errorlevel% neq 0 (
    echo ERRO ao criar venv.
    pause
    goto :eof
)

echo [2/4] Ativando venv e instalando pacotes (isso pode demorar)...
call venv\Scripts\activate.bat
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERRO ao instalar pacotes do requirements.txt.
    pause
    goto :eof
)

echo [3/4] Construindo o executavel (.exe) com PyInstaller...
::
:: COMANDO CORRIGIDO ABAIXO:
:: 1. Juntei tudo em uma linha para evitar erros de 'continuacao' (^) no PowerShell.
:: 2. Corrigi o separador do '--add-data' de ':' para ';' (o correto para Windows).
:: 3. Adicionei '--hidden-import' para garantir que a 'automation_logic' seja incluída.
::
pyinstaller --onefile --windowed --noconsole --name="Automacao_Redirect_Tray" --add-data="de_para_redirecionamento.xlsx;." --hidden-import="automation_logic" automation_gui.py

if %errorlevel% neq 0 (
    echo ERRO ao construir o executavel com PyInstaller.
    pause
    goto :eof
)

echo [4/4] Limpando arquivos temporarios...
rmdir /s /q build
del /q Automacao_Redirect_Tray.spec

echo.
echo ---------------------------------------------------
echo  CONSTRUIDO COM SUCESSO!
echo ---------------------------------------------------
echo O seu executavel (Automacao_Redirect_Tray.exe) esta na pasta 'dist'.
echo.
pause