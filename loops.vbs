Rem �����
Rem ������� ��������

Option Explicit

Dim num, i
Dim res1, res2, res3, res4 '�������� ���������� ��� msgbox

num = CInt(InputBox("����� ����� �����", "��������"))

Rem Do-While-Loop
i = 0
Do While (i < num)
i = i + 1
loop

res1 = MsgBox("���� �����: " & i & ".", vbOK, "Do-While-Loop")

Rem Do-Until-Loop
i = 0
Do Until (i = num)
i = i + 1
loop

res2 = MsgBox("���� �����: " & i & ".", vbOK, "Do-Until-Loop")

Rem While-Wend
i = 0
While i < num
i = i + 1
Wend

res3 = MsgBox("���� �����: " & i & ".", vbOK, "While-Wend")

Rem For, � ������� �� ����������
i = 0
For i=0 To 3000
If i = num then Exit For End if
next

If Not i = num then
MsgBox "����������! (�������� ����� 3000)."
else
res4 = MsgBox("���� �����: " & i & ".", vbOK, "For � �������")
End if
