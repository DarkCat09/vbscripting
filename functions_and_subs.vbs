Rem ������� � ���������
Rem �����������

Rem Option Explicit

Dim numA, numB, strS

Rem ������� ��������� - add
Function mathAdd(a, b)
c = a+b
mathAdd = c
End Function

Rem ������� ������ - subtract
Function mathSbt(a, b)
c = a-b
mathSbt = c
End Function

Rem ������� �������� - multiple
Function mathMlt(a, b)
c = a*b
mathMlt = c
End Function

Rem ������� ��������� - divide
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
MsgBox "������!"
End if
End Sub

numA = CInt(InputBox("������� ������ �����", "�����������"))
numB = CInt(InputBox("������� ������ �����","�����������"))
strS = InputBox("������� ���� ��������������� ��������","�����������")
Call calc(numA,numB,strS)
