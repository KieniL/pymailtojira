@echo off

REM iterateTimeInSeconds
set VAR1=60
REM addJIRAKeytoMailName
set VAR2=True
REM mailCounter
set VAR3=50
REM desiredFolder
set VAR4="Posteingang"

REM pass them all
C:\Users\U11643\Desktop\pymailtojira\Scripts\python.exe main.py %1 %VAR1% %VAR2% %VAR3% %VAR4%
pause