Rem �������������� �������
Rem ��� ��������� �����������,
Rem ����� ��� ������ ��� �������� VBS.
Rem (��� � � ����� - ������ ���������).

'���� ��� ��� ���� �� ������� � �����
'http://vbhack.ru/uroki-vbscript/urok-vbscript-n8-matematicheskie-funkcii-funk.html
'� �������� ������ ����.

Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, o

a = "������ ����� 10: " & Abs(10)
b = "������ ����� -10: " & Abs(-10)
c = "���������� ����� 10: " & Atn(10)
d = "������� ���� 10 ������: " & Cos(10)
e = "����� ���� 10 ������: " & Sin(10)
f = "������� ���� 10 ������: " & Tan(10)
g = "���������� ���������� � 2: " & exp(2)
h = "�������� �� 15: " & log(15)
i = "���������� �������� �� 15: " & log(15)/log(10)
j = "���������� ������ 15: " & Sqr(15)
k = "���������� ������ 0: " & Sqr(0)

min = 3
max = 18

Randomize
l = "������������ 1: " & (Int(1 + (Rnd(3) * 10))) '(Int(min + (Rnd() * max)))
m = "������������ 2: " & Int((10 - 1 + 1) * Rnd + 1) ' Int((max - min + 1) * Rnd + min)
n = "������������ 3: " & Int((10 * Rnd) + 1) ' Int((max * Rnd) + min)
o = "������� ������������: " & Rnd(10)

p = "����� ����� (Int) �� 10.3: " & Int(10.3)
q = "����� ����� (Fix) �� 10.3: " & Fix(10.3)
r = "����� ����� (Int) �� -10.3: " & Int(-10.3)
s = "����� ����� (Fix) �� -10.3: " & Fix(-10.3)

Rem ������� ����� Int � Fix ����������� ������ ��� ������������� ���������.
Rem Fix ���������� ��������� ������� �� ������ �����,
Rem � Int ��������� ������� �� ������ �����.
Rem �������� ��� ��:
Rem http://vbhack.ru/uroki-vbscript/urok-vbscript-n8-matematicheskie-funkcii-funk.html

MsgBox a & vbCrLf & b & vbCrLf & c & vbCrLf & d & vbCrLf & e & vbCrLf & f & vbCrLf &_
g & vbCrLf & h & vbCrLf & i & vbCrLf & j & vbCrLf & k & vbCrLf & l & vbCrLf & m &_
vbCrLf & n & vbCrLf & o & vbCrLf & p & vbCrLf & q & vbCrLf & r & vbCrLf & s
