@echo off
setlocal enabledelayedexpansion

:: Autoelevação
net session >nul 2>&1
if %errorlevel% NEQ 0 (
    echo [INFO] Reabrindo como administrador...
    powershell -Command "Start-Process cmd -ArgumentList '/c %~dpnx0' -Verb RunAs"
    exit /b
)

set "INTERFACE="
set "IP_ATUAL="
set "ACHOU_IP=0"

for /f "tokens=*" %%i in ('netsh interface ip show config') do (
    set "linha=%%i"

    rem Remove espaços no início
    for /f "tokens=* delims= " %%j in ("!linha!") do set "linha=%%j"

    rem Detecta interface
    echo !linha! | findstr /C:"Configuração da interface" >nul && (
        for /f "tokens=2 delims=\" %%a in ("!linha!") do (
            set "INTERFACE=%%a"
            set "INTERFACE=!INTERFACE:"=!"
            echo [DEBUG] Interface encontrada: !INTERFACE!
        )
    )

    rem Detecta IP e extrai
    echo !linha! | findstr /C:"Endereço IP" >nul && (
        for /f "tokens=2 delims=:" %%a in ("!linha!") do (
            set "IP_TEST=%%a"
            for %%x in (!IP_TEST!) do (
                set "IP_ATUAL=%%x"
                echo [DEBUG] IP detectado: !IP_ATUAL!
                call :check_prefix !IP_ATUAL!
            )
        )
    )
)

goto :menu

:check_prefix
set "CHECK=%1"
if not defined CHECK goto :eof
set "PREFIX=!CHECK:~0,7!"
if "!PREFIX!"=="10.14.5" (
    set "ACHOU_IP=1"
    echo [DEBUG] IP válido identificado por prefixo: !CHECK!
)
goto :eof

:menu
cls
echo ========================================================
echo      ALTERAR CONFIGURACOES DE REDE
echo ========================================================

if "!ACHOU_IP!"=="1" (
    echo Interface ativa detectada: !INTERFACE!
    echo IP atual: !IP_ATUAL!
) else (
    echo [ERRO] Nao foi possivel detectar uma interface com IP 10.14.*
    echo.
    echo [DEBUG] Exibindo saída bruta:
    netsh interface ip show config
    pause
    exit /b
)

echo.
echo 1) PRODEMGE (Gateway: 10.14.56.1 / DNS: 200.198.34.80)
echo 2) INET     (Gateway: 10.14.56.2 / DNS: 8.8.8.8)
echo.
set /p ESCOLHA="Escolha a rede que deseja configurar (1 ou 2): "

if "!ESCOLHA!"=="1" goto PRODEMGE
if "!ESCOLHA!"=="2" goto INET

echo [ERRO] Opcao invalida.
pause
exit /b

:PRODEMGE
echo.
echo [SITUACAO] Aplicando rede PRODEMGE...
netsh interface ip set address name="!INTERFACE!" static 10.14.56.62 255.255.255.0 10.14.56.1 1 || echo [ERRO] Falha ao definir IP
netsh interface ip set dns name="!INTERFACE!" static 200.198.34.80 || echo [ERRO] Falha ao definir DNS
echo [OK] Configuração PRODEMGE aplicada.
goto FIM

:INET
echo.
echo [SITUACAO] Aplicando rede INET...
netsh interface ip set address name="!INTERFACE!" static 10.14.56.62 255.255.255.0 10.14.56.2 1 || echo [ERRO] Falha ao definir IP
netsh interface ip set dns name="!INTERFACE!" static 8.8.8.8 || echo [ERRO] Falha ao definir DNS
echo [OK] Configuração INET aplicada.
goto FIM

:FIM
echo.
echo Pressione qualquer tecla para sair...
pause >nul
exit
