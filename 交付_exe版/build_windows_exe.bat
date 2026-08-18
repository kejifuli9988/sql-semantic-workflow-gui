@echo off
setlocal

cd /d %~dp0

if not exist dist mkdir dist

pyinstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --icon sql_compare_tool.ico ^
  --name SQL语义比对工具 ^
  sql_semantic_workflow_gui.py

echo.
echo 打包完成，exe 路径：
echo %~dp0dist\SQL语义比对工具.exe
pause
