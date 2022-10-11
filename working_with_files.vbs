Rem Работа с файлами и папками
Rem Секретные заметки

Option Explicit

Dim FSO, myText
Set FSO = CreateObject("Scripting.FileSystemObject")	'для работы с файлами нужен объект ФС

Rem Создание каталога
If Not FSO.FolderExists("Notes") Then FSO.CreateFolder "Notes" End if

Rem Создание файла и запись данных
Set myText = FSO.CreateTextFile("Notes\1.txt",True,True)	'Создаём текстовый файл 1.txt, если есть - перезаписываем (True), с кодировкой Unicode (True).
myText.Close							'Закрываем TextStream, нам он больше не нужен.

Set myText = FSO.OpenTextFile("Notes\1.txt",2,False,-1)				'Открываем текстовый файл 1.txt в режиме запись (2), с кодировкой Unicode (-1)
myText.WriteLine(InputBox("Введите текст новой секретной заметки", "Заметки"))	'записываем текст
MsgBox "Записано! Читай!"							'оповещаем об этом пользователя
Rem MsgBox myText.ReadAll()							'чтение не работает, т.к. мы открыли в режиме (2).
myText.Close									'закрываем поток

FSO.MoveFile "Notes\1.txt", "2.txt"						'перемещаем файл
WScript.Sleep 5000								'ждём 5000мс=5сек
FSO.DeleteFile "2.txt"								'удаляем эту конфиденциальную заметку
FSO.DeleteFolder "Notes", True							'удаляем всю папку, даже если там файлы с read-only
