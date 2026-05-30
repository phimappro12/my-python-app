@echo off
title TexFlow - Share Mode
echo ============================================
echo    TexFlow - Chia se qua Ngrok
echo ============================================
echo.

:: Ngrok duong dan Windows Store
set NGROK_PATH=C:\Program Files\WindowsApps\ngrok.ngrok_3.39.1.0_x64__1g87z0zv29zzc\ngrok.exe

:: Kiem tra ngrok
if not exist "%NGROK_PATH%" (
    :: Thu dang luoi PATH truoc
    where ngrok >nul 2>&1
    if errorlevel 1 (
        echo [LOI] Khong tim thay ngrok!
        echo Ngrok da cai qua Windows Store - thu chay lai voi quyen Admin.
        pause
        exit /b
    )
    set NGROK_PATH=ngrok
)

:: Kiem tra da dang ky Ngrok account chua
echo [Kiem tra] Neu lan dau dung, can dang nhap Ngrok tai:
echo   https://dashboard.ngrok.com/get-started/your-authtoken
echo   Sau do chay lenh: ngrok config add-authtoken ^<token cua ban^>
echo.

:: Chuyen vao thu muc app
cd /d "%~dp0"

:: Chay app readonly tren port 8502
echo [1/2] Khoi dong TexFlow Read-Only tren port 8502...
start "TexFlow ReadOnly" cmd /k "cd /d %~dp0 && streamlit run app_readonly.py --server.port 8502 --server.headless true"

:: Cho app khoi dong
echo Dang cho app khoi dong (5 giay)...
timeout /t 5 /nobreak > nul

:: Chay ngrok
echo [2/2] Tao tunnel Ngrok...
echo.
echo ============================================
echo  Link se hien o dong "Forwarding" ben duoi
echo  Vi du: https://abc123.ngrok-free.app
echo  Copy link do gui cho nguoi dung!
echo ============================================
echo.
"%NGROK_PATH%" http 8502