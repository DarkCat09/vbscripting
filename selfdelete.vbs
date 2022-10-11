Rem Немножко хулиганим с FSO
Rem Самоудаляемый скрипт

Rem Запустите этот скрипт и он удалится!

Option Explicit

Dim FSO, backuping, scriptName
Set FSO = CreateObject("Scripting.FileSystemObject")
backuping = True					'если хотите, чтобы скрипт сам себя забэкапил, установите
							'backuping = True
scriptName = WScript.ScriptName

Sub BackupScript()
FSO.CreateFolder "backup"
FSO.CopyFile scriptName, "backup\" & scriptName
End Sub

If backuping = True then Call BackupScript() End if
FSO.DeleteFile scriptName
