'Script for adding programs to startup using registry
'Author: Chechkenev Andrew (DarkCat09/CodePicker13)

Option Explicit

Dim Wsh, pathToProgram, isForAllUsers, root, pathToProgramArr

Set Wsh = CreateObject("WScript.Shell")
pathToProgram = InputBox("Enter full path to program", "AddToStartup")
isForAllUsers = MsgBox("Add this program to startup for all users?", vbYesNo, "AddToStartup")
'MsgBox pathToProgram & " " & isForAllUsers

If isForAllUsers = vbYes then
root = "HKEY_LOCAL_MACHINE"
Else
root = "HKEY_CURRENT_USER"
End If

pathToProgramArr = Split(pathToProgram, "\")
'MsgBox root & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & Split(pathToProgramArr(UBound(pathToProgramArr)), ".")(0)
Wsh.RegWrite root & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & Split(pathToProgramArr(UBound(pathToProgramArr)), ".")(0), pathToProgram, "REG_SZ"
