@echo off

cd /d %~sdp0
%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe FixtureBookAddin.dll %*

echo.
pause
exit
