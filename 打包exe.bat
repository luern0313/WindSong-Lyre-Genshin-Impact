@echo off
chcp 65001 >nul
echo ========================================
echo     疯物之诗琴 打包脚本
echo ========================================
echo.

REM 检查是否安装了 PyInstaller
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo [信息] 正在安装 PyInstaller...
    pip install pyinstaller
)

echo [信息] 开始打包...
echo.

REM 使用 spec 文件打包
pyinstaller --clean build_exe.spec

echo.
if %errorlevel% equ 0 (
    echo ========================================
    echo [成功] 打包完成！
    echo 输出路径: dist\疯物之诗琴.exe
    echo ========================================
    echo.
    echo [提示] 运行前请确保 midi 文件夹与 exe 在同一目录下
) else (
    echo [错误] 打包失败，请检查错误信息
)

echo.
pause
