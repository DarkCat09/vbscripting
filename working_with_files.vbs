Rem ������ � ������� � �������
Rem ��������� �������

Option Explicit

Dim FSO, myText
Set FSO = CreateObject("Scripting.FileSystemObject")	'��� ������ � ������� ����� ������ ��

Rem �������� ��������
If Not FSO.FolderExists("Notes") Then FSO.CreateFolder "Notes" End if

Rem �������� ����� � ������ ������
Set myText = FSO.CreateTextFile("Notes\1.txt",True,True)	'������ ��������� ���� 1.txt, ���� ���� - �������������� (True), � ���������� Unicode (True).
myText.Close							'��������� TextStream, ��� �� ������ �� �����.

Set myText = FSO.OpenTextFile("Notes\1.txt",2,False,-1)				'��������� ��������� ���� 1.txt � ������ ������ (2), � ���������� Unicode (-1)
myText.WriteLine(InputBox("������� ����� ����� ��������� �������", "�������"))	'���������� �����
MsgBox "��������! �����!"							'��������� �� ���� ������������
Rem MsgBox myText.ReadAll()							'������ �� ��������, �.�. �� ������� � ������ (2).
myText.Close									'��������� �����

FSO.MoveFile "Notes\1.txt", "2.txt"						'���������� ����
WScript.Sleep 5000								'��� 5000��=5���
FSO.DeleteFile "2.txt"								'������� ��� ���������������� �������
FSO.DeleteFolder "Notes", True							'������� ��� �����, ���� ���� ��� ����� � read-only
