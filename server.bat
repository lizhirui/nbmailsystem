@echo off
echo ���ڲ��Է���������
gacutil /i Microsoft.Matrix.WebHost.dll
ping -n 1 127.0.0.1 >nul
echo �����������
echo ��������web���������
WebServer.EXE /port:84 /path:"E:\������Ʒ\nbmailsystem" /vpath:"/"
pause




