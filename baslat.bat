@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0"
if errorlevel 1 goto directory_error
if not exist ".venv\Scripts\python.exe" goto missing_venv

".venv\Scripts\python.exe" "baslat.py" %*
set "APP_EXIT_CODE=%ERRORLEVEL%"
if not "%APP_EXIT_CODE%"=="0" (
    echo.
    echo Baslatma tamamlanamadi. Yukaridaki mesaji kontrol edin.
    pause
)
exit /b %APP_EXIT_CODE%

:missing_venv
echo Sanal ortam bulunamadi: .venv\Scripts\python.exe
echo Projenin README.md dosyasindaki Windows kurulumu adimlarini uygulayin.
echo Bu baslatici otomatik paket kurulumu yapmaz.
pause
exit /b 1

:directory_error
echo Proje klasorune gecilemedi. Baslatma dosyasinin konumunu kontrol edin.
pause
exit /b 1
