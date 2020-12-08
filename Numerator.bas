Attribute VB_Name = "Numerator"
Dim UIDs As Object      '������� ����������
Dim Prefixes As Object  '������� ���������
Dim Codes As Object     '������� �����
Dim dTab As Variant

'������������� �������
Sub Init()
    
    '�������� ������� ����������
    Call NewTab(tabNum, False)
    Set dTab = Sheets(tabNum)
    dTab.Cells(1, 1) = "��������! ����� ��������� ��������� ����������. ������ �������������� �� ��������������."
    dTab.Cells(3, 1) = "�������"
    dTab.Cells(3, 2) = "�����"
    Range(dTab.Cells(1, 1), dTab.Cells(3, 100)).Interior.Color = RGB(214, 214, 214)
    Set UIDs = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = 4
    Do While dTab.Cells(i, 1) <> ""
        pref = dTab.Cells(i, 1)
        UIDs.Add pref, dTab.Cells(i, 2)
        i = i + 1
    Loop
    
    '�������� ��������� � �����
    Set bTab = Sheets(tabDic)
    i=
End Sub

'���������� ������� �� ��������
Sub Save()
    Dim i As Long
    i = 4
    For Each Key In UIDs.keys
        dTab.Cells(i, 1) = Key
        dTab.Cells(i, 2) = UIDs(Key)
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
Function Generate(dat As Date, buyer As String) As String
    pref = GetPrefix(buyer) + Right(CStr(Year(dat)), 2) + CStr(Month(dat)) + CStr(Day(dat))
    If Not UIDs.exists(pref) Then UIDs.Add pref, 0
    UIDs(pref) = UIDs(pref) + 1
    Generate = pref + Right(CStr(UIDs(pref) + 1000), 3)
End Function

'����� � ������� ��� ��������� ��������
Function GetPrefix(buyer As String) As String
    GetPrefix = UCase(Left(buyer, 1))
End Function