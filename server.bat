@echo off
echo 正在测试服务器环境
gacutil /i Microsoft.Matrix.WebHost.dll
ping -n 1 127.0.0.1 >nul
echo 环境配置完成
echo 正在启动web虚拟服务器
WebServer.EXE /port:84 /path:"E:\早期作品\nbmailsystem" /vpath:"/"
pause




