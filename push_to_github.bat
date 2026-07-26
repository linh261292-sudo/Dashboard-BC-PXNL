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

REM ---- Buoc 1: tu dong cap nhat index.html tu file Excel (neu co Python) ----
set PYEXE=
where python >nul 2>&1
if not errorlevel 1 (
    set PYEXE=python
) else (
    where py >nul 2>&1
    if not errorlevel 1 set PYEXE=py
)

if "%PYEXE%"=="" (
    echo [%date% %time%] CANH BAO: Khong tim thay Python tren may - bo qua buoc tu dong cap nhat tu Excel, chi day index.html hien tai len GitHub. Cai Python tai python.org va chay "pip install openpyxl" de bat tinh nang nay. >> "%LOGFILE%"
) else (
    %PYEXE% "%~dp0generate_dashboard.py" >> "%LOGFILE%" 2>&1
    if errorlevel 1 (
        echo [%date% %time%] LOI: generate_dashboard.py that bai - GIU NGUYEN index.html cu, khong day len GitHub lan nay. Xem chi tiet loi o tren. >> "%LOGFILE%"
        exit /b 1
    ) else (
        echo [%date% %time%] Da chay generate_dashboard.py xong. >> "%LOGFILE%"
    )
)

REM ---- Buoc 2: day len GitHub neu index.html co thay doi ----
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
