Rem ���� ������ � �����
Rem +������� � ������ � �����

Option Explicit

Dim a, b
Dim num
Dim name, result
Dim brthdate
Dim instrChar

a = 10
b = 20

MsgBox "������!"

Rem ������������� �������� ����������
if a = 10 then MsgBox "a = 10!" End if
if b = 20 then MsgBox "b = 20!" else MsgBox "b ��= 20!" End if

Rem �������� ��������� �����
Rem CInt - Convert to Integer
num = CInt(InputBox("����� ����� �����", "�����"))

Rem if, elseif, else, End if - ����� ��� � Bash ��� Python
If num < a then
MsgBox "��� ����� ������ " & a & "."
elseif num > b then
MsgBox "��� ����� ������ " & b & "."
else
Rem vbCrLf - ��� ������ ����� ������
MsgBox "��� ����� � �������� " & a & " � " & b & "," & vbCrLf & "�.�. ������, ��� " & a & ", �� ������, ��� " & b & "."
End if

Rem �������� ������
name = InputBox("����� ���", "�����")
result = MsgBox("������, " & name & "!", vbOKCancel, "���������")
If result = vbOK then MsgBox "����!" End if

Rem �������� ����� � ���� ��������
Rem CDate - Convert to Date
brthdate = CDate(InputBox("����� ���� �������� � ������� ��.��.���� ��� ��/��/����", "�����"))

instrChar = InStr(name, "������")	'���� �� ��� - ������ - � ��������� �����

If (Not instrChar = 0) And Day(brthdate) = 13 And Month(brthdate) = 7 And Year(brthdate) = 2009 then
MsgBox "��������, �� ����������� ����� �������!"
else
MsgBox "OK"
End if
