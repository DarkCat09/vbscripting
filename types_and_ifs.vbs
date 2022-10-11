Rem Типы данных и Циклы
Rem +немного о работе с датой

Option Explicit

Dim a, b
Dim num
Dim name, result
Dim brthdate
Dim instrChar

a = 10
b = 20

MsgBox "Привет!"

Rem Бессмысленная проверка переменных
if a = 10 then MsgBox "a = 10!" End if
if b = 20 then MsgBox "b = 20!" else MsgBox "b не= 20!" End if

Rem Проверка введённого числа
Rem CInt - Convert to Integer
num = CInt(InputBox("Введи любое число", "Инпут"))

Rem if, elseif, else, End if - почти как в Bash или Python
If num < a then
MsgBox "Твоё число меньше " & a & "."
elseif num > b then
MsgBox "Твоё число больше " & b & "."
else
Rem vbCrLf - это символ новой строки
MsgBox "Твоё число в пределах " & a & " и " & b & "," & vbCrLf & "т.е. больше, чем " & a & ", но меньше, чем " & b & "."
End if

Rem Проверка кнопки
name = InputBox("Введи имя", "Инпут")
result = MsgBox("Привет, " & name & "!", vbOKCancel, "Сообщение")
If result = vbOK then MsgBox "Пока!" End if

Rem Проверка имени и даты рождения
Rem CDate - Convert to Date
brthdate = CDate(InputBox("Введи дату рождения в формате ДД.ММ.ГГГГ или ДД/ММ/ГГГГ", "Инпут"))

instrChar = InStr(name, "Андрей")	'ищем моё имя - Андрей - в указанном имени

If (Not instrChar = 0) And Day(brthdate) = 13 And Month(brthdate) = 7 And Year(brthdate) = 2009 then
MsgBox "Возможно, ты разработчик этого скрипта!"
else
MsgBox "OK"
End if
