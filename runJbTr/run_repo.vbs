'*******************************************************
' ����� wscript shell run 
'*******************************************************
Option Explicit

dim WshShell
Set WshShell = WScript.CreateObject("Wscript.Shell") ' ������ ������ ���� Shell
WshShell.Run "C:\pdi-ce-5.4.0.1-130\data-integration\Kitchen.bat /level:Basic /rep:iiemib /job:/runJbTr/Job_1 >> "&WshShell.CurrentDirectory&"\run_repo_vbs.log", 0, false  '���������� �������
Set WshShell = Nothing '������� ����������