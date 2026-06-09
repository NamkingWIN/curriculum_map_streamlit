@echo off
chcp 65001 >nul
setlocal

REM =====================================================
REM  RUN APP CURRICULUM MAPPING
REM  - Tự tạo môi trường ảo .venv nếu chưa có
REM  - Tự cài thư viện trong requirements_curriculum_app.txt
REM  - Chạy ứng dụng Streamlit
REM =====================================================

cd /d "%~dp0"

echo.
echo ============================================
echo   Curriculum Mapping App
echo ============================================
echo Thu muc hien tai: %cd%
echo.

REM Kiem tra file app
if not exist "curriculum_mapping_app.py" (
    echo [LOI] Khong tim thay file curriculum_mapping_app.py trong thu muc nay.
    echo Hay dat run.bat cung thu muc voi curriculum_mapping_app.py.
    echo.
    pause
    exit /b 1
)

REM Kiem tra file requirements; neu chua co thi tao nhanh
if not exist "requirements_curriculum_app.txt" (
    echo [CANH BAO] Khong tim thay requirements_curriculum_app.txt. Dang tao file mac dinh...
    > "requirements_curriculum_app.txt" echo streamlit^>=1.32
    >> "requirements_curriculum_app.txt" echo pandas^>=2.0
    >> "requirements_curriculum_app.txt" echo openpyxl^>=3.1
)

REM Tim Python
set "PYTHON_CMD=python"
%PYTHON_CMD% --version >nul 2>&1
if errorlevel 1 (
    set "PYTHON_CMD=py -3"
    %PYTHON_CMD% --version >nul 2>&1
    if errorlevel 1 (
        echo [LOI] Khong tim thay Python.
        echo Hay cai Python va nho tick "Add python.exe to PATH" khi cai dat.
        echo.
        pause
        exit /b 1
    )
)

echo Dang su dung Python:
%PYTHON_CMD% --version
echo.

REM Tao moi truong ao neu chua co
if not exist ".venv\Scripts\python.exe" (
    echo Dang tao moi truong ao .venv ...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo [LOI] Tao moi truong ao khong thanh cong.
        echo Hay kiem tra lai cai dat Python.
        echo.
        pause
        exit /b 1
    )
) else (
    echo Da co moi truong ao .venv.
)

echo.
echo Dang nang cap pip va cai dat thu vien can thiet...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 (
    echo [CANH BAO] Khong nang cap duoc pip. Tiep tuc cai thu vien...
)

".venv\Scripts\python.exe" -m pip install -r requirements_curriculum_app.txt
if errorlevel 1 (
    echo.
    echo [LOI] Cai dat thu vien khong thanh cong.
    echo Hay kiem tra ket noi internet hoac noi dung file requirements_curriculum_app.txt.
    echo.
    pause
    exit /b 1
)

echo.
echo Dang mo ung dung Streamlit...
echo Neu trinh duyet khong tu mo, hay copy dia chi http://localhost:8501 vao Chrome/Edge.
echo De tat app: quay lai cua so nay va bam Ctrl + C.
echo.

".venv\Scripts\python.exe" -m streamlit run curriculum_mapping_app.py

echo.
echo Ung dung da dung.
pause
