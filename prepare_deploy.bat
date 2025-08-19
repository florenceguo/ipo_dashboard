@echo off
echo 正在准备部署文件...

REM 创建部署文件夹
mkdir "%USERPROFILE%\Desktop\ipo-dashboard" 2>nul

REM 复制网站文件
copy "index.html" "%USERPROFILE%\Desktop\ipo-dashboard\"
copy "app.js" "%USERPROFILE%\Desktop\ipo-dashboard\"
copy "excel_data_summary.json" "%USERPROFILE%\Desktop\ipo-dashboard\"
copy "vercel.json" "%USERPROFILE%\Desktop\ipo-dashboard\"
copy "README.md" "%USERPROFILE%\Desktop\ipo-dashboard\"
copy "package.json" "%USERPROFILE%\Desktop\ipo-dashboard\"

echo.
echo 部署文件已准备完成！
echo 文件位置: %USERPROFILE%\Desktop\ipo-dashboard
echo.
echo 请将这个文件夹压缩成ZIP文件，然后上传到Vercel
pause

