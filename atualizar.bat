@echo off
title Atualizar Televendas
echo ========================================
echo   Atualizador Televendas
echo ========================================
echo.

set "BASEDIR=C:\Users\mar\OneDrive - SPADER DISTRIBUIDORA DE ALIMENTOS L\Area de Trabalho"
set "CAMDIR=%BASEDIR%\dashboards\televendas"
set "SCRIPTDIR=%~dp0"

echo [1/3] Extraindo dados da base...
node "%SCRIPTDIR%extrair_televendas.js"
if %ERRORLEVEL% neq 0 (
    echo ERRO ao extrair dados!
    pause
    exit /b 1
)
echo OK!
echo.

echo [2/3] Enviando para GitHub...
git add dados_julho.js dados_mensal.js dados_julho.json
git commit -m "feat: atualizacao televendas %date% %time%" 2>nul
git push
echo.

echo [3/3] Resumo:
node -e "var d=require('%CAMDIR%\\dados_julho.json');console.log('  Mes: '+d.mes);console.log('  Pedidos: '+d.rows.length);var f=d.rows.reduce((s,r)=>s+r.fat,0);console.log('  Faturamento: R\$ '+f.toFixed(2));"
echo.
echo Dashboard atualizado!
pause
