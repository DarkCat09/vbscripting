Rem ������� ������
Rem +�������, Select ... case � Err
Rem ��������� �� ��������

On Error Resume Next		'������������ ��� ������ - ������ ��� ��� �� �����������
Rem Option Explicit

Dim res1			'�������� ���������� ��� MsgBox
Dim exitFromMenu

Dim prodFileName, optionsDelim
prodFileName = "products.csv"	'��� ����� � ����������
optionsDelim = ","		'����������� ����� (�������,����)
				'���� � �������� ������ ����� ������������� �������,
				'����� ������� optionsDelim = ";"
exitFromMenu = False

Rem �������, ��� � Java � .NET:
Set Products = CreateObject("Scripting.Dictionary")

Dim FSO, ProdFile
Set FSO = CreateObject("Scripting.FileSystemObject")

Sub addProducts()
Set ProdFile = FSO.OpenTextFile(prodFileName, 8, True, -1)
prodOptions = "������������; ����"

Do Until (prodOptions = "")

prodOptions = InputBox("������� ����� ������ (������������ � ����), �������� ������ � �������." &_
vbCrLf & "������: ���� �������� 0.5 �; 25", "��������� ���� ������ ��������")

splitedOpts = Split(prodOptions, "; ")
If UBound(splitedOpts) > 0 then Products.Add splitedOpts(0), splitedOpts(1) End if	'�������� - ����� �� 0? � ������ � �������
loop

For Each prodkey In Products.keys					'��� ������� ����� � �������...
ProdFile.Write(prodkey & optionsDelim &_
Products(prodkey) & vbCrLf)						'���������� � ���� (���� - ����.������ + ����������� + ��������(�����) - ����).
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
Rem cost - ������� ���������.
Rem pct  - ������, � ���������.
itog = cost - (cost/100*pct)
computeDiscount = itog
End Function

Rem ������ ������������ ���� �� �������

Do While (exitFromMenu = False)	'���� ������������ �� ����� �� ����, ���������� ���

selected_function = InputBox("�������� �������:" & vbCrLf & "1 - �������� �������� � ����" & vbCrLf &_
"2 - ����������� ������ ���������" & vbCrLf & "3 - �������� ������ ���������" & vbCrLf &_
"4 - ���������� ���� ������ �� �������" & vbCrLf & "5 - �����", "��������� ���� ������ ��������")

Rem select...case - �� �� �����,
Rem ��� � �������� ���� C-������������� switch...case.
Rem ������ ��������� �������� ����������.
Select case selected_function
case "1"
Call addProducts()
case "2"
Call readProducts()
Rem If Err.Number = 53 then MsgBox "���� ������� � ����������!" End if
case "3"
Call removeAllProducts()
case "4"
header = "���������� ���� �� �������"
cost = CInt(InputBox("������� ���������:", header))
prct = CInt(InputBox("������, � %:", header))
res1 = MsgBox(computeDiscount(cost, prct), vbOK, header)
case "5"
exitFromMenu = True
case else
MsgBox "������!" & vbCrLf & "������� ���������� ��������." & vbCrLf & "����� �����, �������� 5."
End Select
loop
