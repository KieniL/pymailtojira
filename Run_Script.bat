@echo off

REM iterateTimeInSeconds
set VAR1=600
REM addJIRAKeytoMailName
set VAR2=True
REM mailCounter
set VAR3=50
REM desiredFolder
set VAR4="Posteingang"
REM Path to PythonDirectory
set PythonPath=C:\Users\U11643\Desktop\pymailtojira
set VAR6=\Scripts\python.exe
set VAR7=%PythonPath%%VAR6%


REM Use Python Interpreter run python program and pass all variables to it
"%VAR7%" main.py %1 %VAR1% %VAR2% %VAR3% %VAR4%
pause