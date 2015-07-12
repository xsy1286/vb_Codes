@FOR %%A IN (*.REG) DO (REGEDIT /S %%A)

ping -n 127.0.01 2>nul
@echo off
cd /d %PROGRAMFILES%\client\
start client.exe