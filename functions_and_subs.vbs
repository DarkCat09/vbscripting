Rem Функции и Процедуры
Rem Калькулятор

Rem Option Explicit

Dim numA, numB, strS

Rem Функция прибавить - add
Function mathAdd(a, b)
c = a+b
mathAdd = c
End Function

Rem Функция отнять - subtract
Function mathSbt(a, b)
c = a-b
mathSbt = c
End Function

Rem Функция умножить - multiple
Function mathMlt(a, b)
c = a*b
mathMlt = c
End Function

Rem Функция разделить - divide
Function mathDvd(a, b)
c = a/b
mathDvd = c
End Function

Sub calc(a, b, sign)
If sign = "+" then
MsgBox a & "+" & b & "=" & mathAdd(a, b)
elseif sign = "-" then
MsgBox a & "-" & b & "=" & mathSbt(a, b)
elseif sign = "*" then
MsgBox a & "*" & b & "=" & mathMlt(a, b)
elseif sign = "/" then
MsgBox a & "/" & b & "=" & mathDvd(a, b)
else
MsgBox "Ошибка!"
End if
End Sub

numA = CInt(InputBox("Введите первое число", "Калькулятор"))
numB = CInt(InputBox("Введите второе число","Калькулятор"))
strS = InputBox("Введите знак арифметического действия","Калькулятор")
Call calc(numA,numB,strS)
