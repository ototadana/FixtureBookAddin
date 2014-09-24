@echo off

cd /d %~sdp0
start wscript runas.js "regasm.bat /codebase /tlb"
