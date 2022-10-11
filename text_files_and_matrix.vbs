Rem Текстовые файлы и Матрицы
Rem Имитация Светодиодной Матрицы в Блокноте.

Rem НЕ работает!

Option Explicit

Dim i, j
Dim FSO, myFile
Dim matrix(7,3)					'инициализируем двумерный массив или матрицу 8x4

Set FSO = CreateObject("Scripting.FileSystemObject")
Set myFile = FSO.CreateTextFile("matrix.txt",True,True)
myFile.Close

Sub writeMatrix()
Set myFile = FSO.OpenTextFile("matrix.txt",2,True,0)	'открываем matrix.txt в режиме записи (2) с кодировкой ASCII (0)

For j=0 To 3
For i=0 To 7
If matrix(i,j)=True then myFile.Write(".") else myFile.Write(" ") End if
next
myFile.Write(vbCrLf)
next

myFile.Close		'закрываем поток
End Sub

Rem Заполнение
For i=0 To 7
For j=0 To 3
matrix(i, j)=True	'включено
next
next

writeMatrix()

Rem Каёмочка
For i=0 To 7
matrix(i, 0)=True
next
For j=0 To 3
matrix(0, j)=True
next
For i=0 To 7
matrix(i, 7)=True
next
For j=0 To 3
matrix(7, j)=True
next

writeMatrix()

Rem ...
