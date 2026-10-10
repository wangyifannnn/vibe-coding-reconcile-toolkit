@echo off
chcp 65001 >nul
cd /d "%~dp0"

where py >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
) else (
    set "PYTHON_CMD=python"
)

%PYTHON_CMD% -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo 首次运行，正在安装所需组件...
    %PYTHON_CMD% -m pip install -r requirements.txt
    if errorlevel 1 goto error
)

%PYTHON_CMD% main.py
if errorlevel 1 goto error

echo.
echo 结果已生成在 output 文件夹中。
pause
exit /b 0

:error
echo.
echo 运行失败，请检查上方错误信息。
pause
exit /b 1
