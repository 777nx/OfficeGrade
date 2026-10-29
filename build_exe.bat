@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/3] 安装打包依赖...
python -m pip install -q pyinstaller pypinyin python-docx openpyxl

echo [2/3] 开始打包 OfficeGrade.exe ...
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name OfficeGrade ^
  --icon "assets\app.ico" ^
  --version-file version_info.txt ^
  --add-data "答案\打字测试.docx;答案" ^
  --add-data "答案\简历表.docx;答案" ^
  --add-data "评分标准.md;." ^
  --add-data "assets\app.ico;assets" ^
  --add-data "assets\app_icon.png;assets" ^
  --add-data "assets\app_icon_64.png;assets" ^
  --collect-all pypinyin ^
  --hidden-import docx ^
  --hidden-import openpyxl ^
  --hidden-import grade_exam ^
  --hidden-import grade_gui ^
  grade_gui.py
rem 以后改代码再打包，直接再跑本 bat 即可
rem 需要命令行时：OfficeGrade.exe --cli --submissions "..."

if errorlevel 1 (
  echo 打包失败
  exit /b 1
)

echo [3/3] 整理发布目录 release\ ...
if not exist "release" mkdir release
copy /Y "dist\OfficeGrade.exe" "release\OfficeGrade.exe" >nul
copy /Y "使用说明.txt" "release\使用说明.txt" >nul
if not exist "release\答案" mkdir "release\答案"
copy /Y "答案\打字测试.docx" "release\答案\" >nul
copy /Y "答案\简历表.docx" "release\答案\" >nul
copy /Y "评分标准.md" "release\评分标准.md" >nul

rem 从 version_info.txt 读出版本号，另存带版本号的 exe，便于 GitHub Release 上传
set "VER="
for /f "usebackq tokens=2 delims='" %%A in (`findstr /C:"ProductVersion" "version_info.txt"`) do (
  if not defined VER set "VER=%%A"
)
if defined VER (
  copy /Y "dist\OfficeGrade.exe" "release\OfficeGrade_v%VER%.exe" >nul
  echo 发版附件：release\OfficeGrade_v%VER%.exe
) else (
  echo 警告：未能从 version_info.txt 读取 ProductVersion
)

echo.
echo 完成！把整个 release 文件夹拷到任意 Windows 电脑即可用。
echo 主程序：release\OfficeGrade.exe
echo.
pause
