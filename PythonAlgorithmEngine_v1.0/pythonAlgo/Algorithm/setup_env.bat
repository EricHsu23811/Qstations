@echo off
REM ==============================================
REM  Python Virtual Environment Setup Script
REM  在 .bat 所在資料夾建立/啟動虛擬環境並安裝套件
REM ==============================================

echo [1] Creating virtual environment...
python -m venv venv

IF NOT EXIST venv (
    echo Error: Failed to create virtual environment
    pause
    exit /b
)

echo [2] Activating virtual environment...
call venv\Scripts\activate

echo [3] Upgrading pip...
python -m pip install --upgrade pip

echo [4] Installing packages...
pip install numpy matplotlib

echo ----------------------------------------------
echo Virtual environment setup complete!
echo Location: %CD%\venv
echo Installed packages: numpy, matplotlib
echo ----------------------------------------------

pause
