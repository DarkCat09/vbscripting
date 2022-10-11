Rem Работа с массивами и строками

Rem Option Explicit

Dim res1		'ненужная переменная для MsgBox
Dim values, strChar
Dim oldStrChar

Dim dyn_arr()		'динамический массив
Dim fix_arr(10)		'массив фиксированного размера - 11
Dim matrix_arr(3,2) 	'двумерный массив или матрица 4x3

Rem Отсчёт индекса начинается с нуля, как во всех ЯП!
Rem Т.е. 2=0,1,2=(по-человечески)3.

Rem Заполнение массива вручную
fix_arr(0)  = "Привет, "
fix_arr(1)  = "Мир! "
fix_arr(2)  = "Как дела? "
fix_arr(3)  = 5			'это означает отлично
fix_arr(4)  = ". Кошка "	'а это мой стих
fix_arr(5)  = "села на "
fix_arr(6)  = "порог, "
fix_arr(7)  = "вот и "
fix_arr(8)  = "съела весь "
fix_arr(9)  = "мой пирог. "
fix_arr(10) = "Пока!"
Rem В массиве VBS можно хранить всё что угодно!
Rem И неважно, что там было первое - строка или число.

Rem Заполнение массива через цикл
Rem Например, мы работаем со светодиодной матрицей.
Rem Нужно всю её залить белым цветом - вкл. все светодиоды.
Rem Вручную это прописывать слишком долго.
For i=0 To 3
For j=0 To 2
matrix_arr(i,j) = True 'вкл.
next
next
Rem А применение матрице рассмотрим в работе с файлами.


Rem Получение массива из строки, парсинг.
Rem Запрашиваем значения у пользователя.
values = InputBox("Введите значения массива через запятую", "Динамический массив")

Rem Сложный, но железобетонный и гибкий вариант парсинга из строк.
Rem Я всё закомментировал, так как данный код не работает!
Rem strChar = InStr(values, ",")		'находим в строке разделитель
Rem MsgBox strChar & " " & Len(values)
Rem dyn_arr(0) = Mid(values, 1, strChar)	'вырезаем нужную часть строки

Rem k = 1
Rem Do Until (strChar = 0)
Rem oldStrChar = strChar
Rem strChar = InStr(oldStrChar, values, ",")		'находим разделитель начиная с предыдущего
Rem dyn_arr(k) = Mid(values, oldStrChar, strChar)	'вырезаем нужную часть строки
Rem loop

Rem Выводим пользователю полученный массив,
Rem соеденяя его в строку с помощью Join.
Rem Разделитель по умолчанию - пробел.
Rem res1 = MsgBox(dyn_arr.Join, vbOK, "Сложный парсинг с помощью InStr")


Rem Очень лёгкий способ.
Rem Эта функция существует в .NET (в т.ч. VB) и Java.
Rem И называется...
Rem Split

Rem Однако, из-за странностей Visual Basic
Rem придётся делать вот так:
temp_arr = Split(values, ",")
For l=0 To UBound(temp_arr)
Redim Preserve dyn_arr(l)	'перезаписываем размер массива с сохранением данных - ReDim Preserve
dyn_arr(l) = temp_arr(l)
next

res1 = MsgBox(Join(dyn_arr), vbOK, "Простой парсинг с помощью Split")

Rem Да, всё настолько просто. И это работает!
Rem Хотя повозиться пришлось.
Rem Но String.Split есть не во всех ЯП и оно слишком автоматическое.


Rem Выводим массив фиксированного размера всё тем же Join.
MsgBox Join(fix_arr, "")	'с пустым разделителем
