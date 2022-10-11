Rem Математические функции
Rem Эта программа бесполезная,
Rem разве что только для изучения VBS.
Rem (что я и делаю - изучаю скриптинг).

'Весь код был взят из примера с сайта
'http://vbhack.ru/uroki-vbscript/urok-vbscript-n8-matematicheskie-funkcii-funk.html
'И немножко изменён мною.

Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, o

a = "Модуль числа 10: " & Abs(10)
b = "Модуль числа -10: " & Abs(-10)
c = "Арктангенс числа 10: " & Atn(10)
d = "Косинус угла 10 радиан: " & Cos(10)
e = "Синус угла 10 радиан: " & Sin(10)
f = "Тангенс угла 10 радиан: " & Tan(10)
g = "Экспонента возведённая в 2: " & exp(2)
h = "Логарифм из 15: " & log(15)
i = "Десятичный логарифм из 15: " & log(15)/log(10)
j = "Квадратный корень 15: " & Sqr(15)
k = "Квадратный корень 0: " & Sqr(0)

min = 3
max = 18

Randomize
l = "Рандомизация 1: " & (Int(1 + (Rnd(3) * 10))) '(Int(min + (Rnd() * max)))
m = "Рандомизация 2: " & Int((10 - 1 + 1) * Rnd + 1) ' Int((max - min + 1) * Rnd + min)
n = "Рандомизация 3: " & Int((10 * Rnd) + 1) ' Int((max * Rnd) + min)
o = "Простая рандомизация: " & Rnd(10)

p = "Целое число (Int) из 10.3: " & Int(10.3)
q = "Целое число (Fix) из 10.3: " & Fix(10.3)
r = "Целое число (Int) из -10.3: " & Int(-10.3)
s = "Целое число (Fix) из -10.3: " & Fix(-10.3)

Rem Разница между Int и Fix проявляется только при отрицательных значениях.
Rem Fix возвращает ближайшее меньшее по модулю число,
Rem а Int ближайшее большее по модулю число.
Rem Источник тот же:
Rem http://vbhack.ru/uroki-vbscript/urok-vbscript-n8-matematicheskie-funkcii-funk.html

MsgBox a & vbCrLf & b & vbCrLf & c & vbCrLf & d & vbCrLf & e & vbCrLf & f & vbCrLf &_
g & vbCrLf & h & vbCrLf & i & vbCrLf & j & vbCrLf & k & vbCrLf & l & vbCrLf & m &_
vbCrLf & n & vbCrLf & o & vbCrLf & p & vbCrLf & q & vbCrLf & r & vbCrLf & s
