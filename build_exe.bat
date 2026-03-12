@echo off
echo 正在准备打包环境...
pip install -r requirements.txt
pip install pyinstaller

echo 正在开始打包 (这可能需要 1-2 分钟)...
pyinstaller --onefile --name "文件匹配统计工具" file_processor.py

echo.
echo ========================================================
echo 打包成功！
echo 请在 dist 文件夹下找到 "文件匹配统计工具.exe"
echo ========================================================
pause
