@echo off
chcp 65001 >nul
rem Preenche a planilha DRES a partir dos prints do Relatorio de coletas.
rem Deixe este arquivo na mesma pasta do preencher_dres.py e da planilha .xlsx,
rem com os prints dentro da subpasta "prints". Depois e so dar dois cliques.

cd /d "%~dp0"

if not exist "preencher_dres.py" (
  echo Nao achei o preencher_dres.py nesta pasta: %~dp0
  pause & exit /b
)
if not exist "prints\" (
  mkdir prints
  echo Criei a pasta "prints". Coloque os prints dos cartoes nela e rode de novo.
  pause & exit /b
)

rem Chave da API: pede uma vez e guarda no Windows (setx)
if "%ANTHROPIC_API_KEY%"=="" (
  set /p ANTHROPIC_API_KEY=Cole a sua chave da API da Anthropic ^(sk-ant-...^) e aperte Enter:
)
if "%ANTHROPIC_API_KEY%"=="" (
  echo Sem chave, nao da para continuar.
  pause & exit /b
)
setx ANTHROPIC_API_KEY "%ANTHROPIC_API_KEY%" >nul

rem Usa a primeira planilha .xlsx da pasta que nao seja um resultado anterior
set "PLANILHA="
for %%F in (*.xlsx) do (
  echo %%~nF | findstr /i /c:"_preenchida" >nul || if not defined PLANILHA set "PLANILHA=%%F"
)
if not defined PLANILHA (
  echo Nao achei a planilha .xlsx nesta pasta.
  pause & exit /b
)

echo Planilha: %PLANILHA%
echo.
python preencher_dres.py --planilha "%PLANILHA%" --fotos "prints" %*
echo.
pause
