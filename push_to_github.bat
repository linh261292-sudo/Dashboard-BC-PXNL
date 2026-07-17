@echo off
setlocal
cd /d "%~dp0"

set LOGFILE=%~dp0push_to_github_log.txt

echo [%date% %time%] Bat dau kiem tra... >> "%LOGFILE%"

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
    echo [%date% %time%] LOI: Thu muc nay chua duoc git init / chua ket noi GitHub. Xem huong dan thiet lap. >> "%LOGFILE%"
    exit /b 1
)

git add index.html >nul 2>&1
git diff --cached --quiet
if %errorlevel%==0 (
    echo [%date% %time%] Khong co thay doi moi trong index.html - bo qua. >> "%LOGFILE%"
) else (
    git commit -m "Auto update dashboard %date% %time%" >nul 2>&1
    git push origin main >> "%LOGFILE%" 2>&1
    if errorlevel 1 (
        echo [%date% %time%] LOI: Push len GitHub that bai. Kiem tra ket noi mang / dang nhap Git. >> "%LOGFILE%"
        exit /b 1
    ) else (
        echo [%date% %time%] Da push index.html len GitHub thanh cong. >> "%LOGFILE%"
    )
)

endlocal
