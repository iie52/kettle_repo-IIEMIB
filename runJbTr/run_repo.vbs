'*******************************************************
' метод wscript shell run 
'*******************************************************
Option Explicit

dim WshShell
Set WshShell = WScript.CreateObject("Wscript.Shell") ' создаём объект типа Shell
WshShell.Run "C:\pdi-ce-5.4.0.1-130\data-integration\Kitchen.bat /level:Basic /rep:iiemib /job:/runJbTr/Job_1 >> "&WshShell.CurrentDirectory&"\run_repo_vbs.log", 0, false  'исполнение команды
Set WshShell = Nothing 'очищаем переменную