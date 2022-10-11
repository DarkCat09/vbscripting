Rem Циклы
Rem Простой брутфорс

Option Explicit

Dim num, i
Dim res1, res2, res3, res4 'Ненужные переменные для msgbox

num = CInt(InputBox("Введи целое число", "Брутфорс"))

Rem Do-While-Loop
i = 0
Do While (i < num)
i = i + 1
loop

res1 = MsgBox("Ваше число: " & i & ".", vbOK, "Do-While-Loop")

Rem Do-Until-Loop
i = 0
Do Until (i = num)
i = i + 1
loop

res2 = MsgBox("Ваше число: " & i & ".", vbOK, "Do-Until-Loop")

Rem While-Wend
i = 0
While i < num
i = i + 1
Wend

res3 = MsgBox("Ваше число: " & i & ".", vbOK, "While-Wend")

Rem For, с защитой от перегрузки
i = 0
For i=0 To 3000
If i = num then Exit For End if
next

If Not i = num then
MsgBox "Перегрузка! (Итераций более 3000)."
else
res4 = MsgBox("Ваше число: " & i & ".", vbOK, "For с защитой")
End if
