Attribute VB_Name = "Numerator"
Dim UIDs As Object      '������� ����������
Dim Prefixes As Object  '������� ���������
Dim Codes As Object     '������� �����

'������������� �������
Sub Init()
    
    '�������� ������� ����������
    NUM.Cells(1, 1) = "��������! ����� ��������� ��������� ����������. ������ �������������� �� ��������������."
    NUM.Cells(3, 1) = "�������"
    NUM.Cells(3, 2) = "�����"
    Range(NUM.Cells(1, 1), NUM.Cells(3, 100)).Interior.Color = RGB(214, 214, 214)
    Set UIDs = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = 4
    Do While NUM.Cells(i, 1) <> ""
        pref = NUM.Cells(i, 1)
        UIDs.Add pref, NUM.Cells(i, 2)
        i = i + 1
    Loop
    
    '�������� ��������� � �����
    i = firstDic
    'while
End Sub

'���������� ������� �� ��������
Sub Save()
    Dim i As Long
    i = 4
    For Each Key In UIDs.keys
        NUM.Cells(i, 1) = Key
        NUM.Cells(i, 2) = UIDs(Key)
        i = i + 1
    Next
End Sub

'������� �������
Sub Clear()
    On Error GoTo er
    Sheets(tabNum).Cells.Clear
er:
End Sub

'��������� ����������� ������
Function Generate(DAT As Date, buyer As String) As String
    pref = GetPrefix(buyer) + Right(CStr(Year(DAT)), 2) + CStr(Month(DAT)) + CStr(Day(DAT))
    If Not UIDs.exists(pref) Then UIDs.Add pref, 0
    UIDs(pref) = UIDs(pref) + 1
    Generate = pref + Right(CStr(UIDs(pref) + 1000), 3)
End Function

'����� � ������� ��� ��������� ��������
Function GetPrefix(buyer As String) As String
    GetPrefix = UCase(Left(buyer, 1))
End Function