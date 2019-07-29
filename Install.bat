@echo on
REM Path to PythonDirectory
set PythonPath=C:\Users\U11643\Desktop\pymailtojira
set VAR6=\Scripts\pip.exe
set VAR7=%PythonPath%%VAR6%


REM Use Python Interpreter to install python packages
"%VAR7%" install pymsgbox pandas xlrd pywin32 jira