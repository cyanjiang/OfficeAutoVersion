@echo off
rem 打包安装程序：先在 Visual Studio 里生成 OfficeAutoVersion.sln，再双击本文件
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-setup.ps1" %*
if errorlevel 1 pause
