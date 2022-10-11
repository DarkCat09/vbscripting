'Automatically rename files, adding entered prefix
'by DarkCat09

Option Explicit

Dim prefix, directory
Dim Fso, Dir, File
Set Fso = CreateObject("Scripting.FileSystemObject")

prefix = InputBox("Enter file prefix:", "Autorename")

If Wsh.Arguments.Count > 0 Then
	directory = Wsh.Arguments(0)
Else
	directory = Fso.GetParentFolderName(WScript.ScriptFullName)
End If

Set Dir = Fso.GetFolder(directory)

For Each File in Dir.Files
	Fso.MoveFile directory & "\" & File.Name, directory & "\" & prefix & File.Name
Next

MsgBox "OK!"
