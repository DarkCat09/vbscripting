Dim Fso
Set Fso = CreateObject("Scripting.FileSystemObject")
MsgBox "0 = " & Fso.GetSpecialFolder(0) & vbCrLf & "1 = " & Fso.GetSpecialFolder(1) & vbCrLf & "2 = " & Fso.GetSpecialFolder(2)
MsgBox "System drive = " & Fso.GetSpecialFolder(0).Drive
