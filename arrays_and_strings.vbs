Rem ������ � ��������� � ��������

Rem Option Explicit

Dim res1		'�������� ���������� ��� MsgBox
Dim values, strChar
Dim oldStrChar

Dim dyn_arr()		'������������ ������
Dim fix_arr(10)		'������ �������������� ������� - 11
Dim matrix_arr(3,2) 	'��������� ������ ��� ������� 4x3

Rem ������ ������� ���������� � ����, ��� �� ���� ��!
Rem �.�. 2=0,1,2=(��-�����������)3.

Rem ���������� ������� �������
fix_arr(0)  = "������, "
fix_arr(1)  = "���! "
fix_arr(2)  = "��� ����? "
fix_arr(3)  = 5			'��� �������� �������
fix_arr(4)  = ". ����� "	'� ��� ��� ����
fix_arr(5)  = "���� �� "
fix_arr(6)  = "�����, "
fix_arr(7)  = "��� � "
fix_arr(8)  = "����� ���� "
fix_arr(9)  = "��� �����. "
fix_arr(10) = "����!"
Rem � ������� VBS ����� ������� �� ��� ������!
Rem � �������, ��� ��� ���� ������ - ������ ��� �����.

Rem ���������� ������� ����� ����
Rem ��������, �� �������� �� ������������ ��������.
Rem ����� ��� � ������ ����� ������ - ���. ��� ����������.
Rem ������� ��� ����������� ������� �����.
For i=0 To 3
For j=0 To 2
matrix_arr(i,j) = True '���.
next
next
Rem � ���������� ������� ���������� � ������ � �������.


Rem ��������� ������� �� ������, �������.
Rem ����������� �������� � ������������.
values = InputBox("������� �������� ������� ����� �������", "������������ ������")

Rem �������, �� �������������� � ������ ������� �������� �� �����.
Rem � �� ���������������, ��� ��� ������ ��� �� ��������!
Rem strChar = InStr(values, ",")		'������� � ������ �����������
Rem MsgBox strChar & " " & Len(values)
Rem dyn_arr(0) = Mid(values, 1, strChar)	'�������� ������ ����� ������

Rem k = 1
Rem Do Until (strChar = 0)
Rem oldStrChar = strChar
Rem strChar = InStr(oldStrChar, values, ",")		'������� ����������� ������� � �����������
Rem dyn_arr(k) = Mid(values, oldStrChar, strChar)	'�������� ������ ����� ������
Rem loop

Rem ������� ������������ ���������� ������,
Rem �������� ��� � ������ � ������� Join.
Rem ����������� �� ��������� - ������.
Rem res1 = MsgBox(dyn_arr.Join, vbOK, "������� ������� � ������� InStr")


Rem ����� ����� ������.
Rem ��� ������� ���������� � .NET (� �.�. VB) � Java.
Rem � ����������...
Rem Split

Rem ������, ��-�� ����������� Visual Basic
Rem ������� ������ ��� ���:
temp_arr = Split(values, ",")
For l=0 To UBound(temp_arr)
Redim Preserve dyn_arr(l)	'�������������� ������ ������� � ����������� ������ - ReDim Preserve
dyn_arr(l) = temp_arr(l)
next

res1 = MsgBox(Join(dyn_arr), vbOK, "������� ������� � ������� Split")

Rem ��, �� ��������� ������. � ��� ��������!
Rem ���� ���������� ��������.
Rem �� String.Split ���� �� �� ���� �� � ��� ������� ��������������.


Rem ������� ������ �������������� ������� �� ��� �� Join.
MsgBox Join(fix_arr, "")	'� ������ ������������
