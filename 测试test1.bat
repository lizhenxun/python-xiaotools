@echo off
:: 检查是否已以管理员身份运行
openfiles >nul 2>nul
if %errorlevel% NEQ 0 (
    echo 需要管理员权限。正在以管理员权限重新运行...
    powershell -Command "Start-Process cmd -ArgumentList '/c %~s0' -Verb RunAs"
    exit
)

:: 删除当前文件夹下所有 .xlsx 文件
echo 正在删除 .xlsx 文件...
del /q *.xlsx
echo .xlsx 文件已删除。

:: 解压所有 .zip 压缩包到当前文件夹
echo 正在解压 .zip 压缩包...
for %%i in (*.zip) do (
    echo 解压 %%i...
    powershell -Command "Expand-Archive -Path '%%i' -DestinationPath '.'"
)
echo 解压完成。

:: 结束并自动退出
exit
