Rem Большой проект
Rem +словари, Select ... case и Err
Rem Локальная БД Магазина

On Error Resume Next		'Обрабатываем все ошибки - теперь они нас не побеспокоят
Rem Option Explicit

Dim res1			'Ненужная переменная для MsgBox
Dim exitFromMenu

Dim prodFileName, optionsDelim
prodFileName = "products.csv"	'Имя файла с продуктами
optionsDelim = ","		'Разделитель опций (продукт,цена)
				'Если в названии товара может присутсвовать запятая,
				'Лучше ставить optionsDelim = ";"
exitFromMenu = False

Rem Словарь, как в Java и .NET:
Set Products = CreateObject("Scripting.Dictionary")

Dim FSO, ProdFile
Set FSO = CreateObject("Scripting.FileSystemObject")

Sub addProducts()
Set ProdFile = FSO.OpenTextFile(prodFileName, 8, True, -1)
prodOptions = "наименование; цена"

Do Until (prodOptions = "")

prodOptions = InputBox("Введите опции товара (наименование и цена), разделяя точкой с запятой." &_
vbCrLf & "Пример: Вода Волжанка 0.5 л; 25", "Локальная База Данных Магазина")

splitedOpts = Split(prodOptions, "; ")
If UBound(splitedOpts) > 0 then Products.Add splitedOpts(0), splitedOpts(1) End if	'проверка - опций не 0? и запись в словарь
loop

For Each prodkey In Products.keys					'для каждого ключа в словаре...
ProdFile.Write(prodkey & optionsDelim &_
Products(prodkey) & vbCrLf)						'записываем в файл (ключ - наим.товара + разделитель + значение(ключа) - цену).
next

ProdFile.Close
End Sub

Sub readProducts()
Set ProdFile = FSO.OpenTextFile(prodFileName, 1, False, -1)
MsgBox Err.Number
MsgBox ProdFile.ReadAll()
ProdFile.Close
End Sub

Sub removeAllProducts()
FSO.DeleteFile prodFileName, True
End Sub

Function computeDiscount(cost, pct)
Rem cost - обычная стоимость.
Rem pct  - скидка, в процентах.
itog = cost - (cost/100*pct)
computeDiscount = itog
End Function

Rem Делаем своеобразное меню на Инпутах

Do While (exitFromMenu = False)	'пока пользователь не вышел из меню, показываем его

selected_function = InputBox("Выберите функцию:" & vbCrLf & "1 - добавить продукты в файл" & vbCrLf &_
"2 - просмотреть список продуктов" & vbCrLf & "3 - очистить список продуктов" & vbCrLf &_
"4 - рассчитать цену товара со скидкой" & vbCrLf & "5 - выйти", "Локальная База Данных Магазина")

Rem select...case - то же самое,
Rem что и знакомый всем C-программистам switch...case.
Rem Только синтаксис немножко отличается.
Select case selected_function
case "1"
Call addProducts()
case "2"
Call readProducts()
Rem If Err.Number = 53 then MsgBox "Нету файлика с продуктами!" End if
case "3"
Call removeAllProducts()
case "4"
header = "Рассчитать цену со скидкой"
cost = CInt(InputBox("Обычная стоимость:", header))
prct = CInt(InputBox("Скидка, в %:", header))
res1 = MsgBox(computeDiscount(cost, prct), vbOK, header)
case "5"
exitFromMenu = True
case else
MsgBox "Ошибка!" & vbCrLf & "Укажите правильное значение." & vbCrLf & "Чтобы выйти, выберите 5."
End Select
loop
