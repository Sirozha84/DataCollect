Attribute VB_Name = "Numerator"
Const NumPage = "������� ����������"
Dim Nums As Object
Dim dTab As Variant

'������������� �������
Sub Init()
    Message "������������� ������� ����������"
    Call NewTab(NumPage, False)
    Set dTab = Sheets(NumPage)
    Set Nums = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = 1
    Do While dTab.Cells(i, 1) <> ""
        Nums.Add dTab.Cells(i, 1), dTab.Cells(i, 2)
        i = i + 1
    Loop
End Sub

'���������� ������� �� ��������
Sub Save()
    Dim i As Long
    i = 1
    For Each Key In Nums.keys
        dTab.Cells(i, 1) = Key
        dTab.Cells(i, 2) = Nums(Key)
        i = i + 1
    Next
End Sub

Sub Clear()
    On Error GoTo er
    Sheets(NumPage).Cells.Clear
er:
End Sub

Function Generate(dat As Date, Buyer As String) As String
    '����������� ��� ���������� �� ����� ���� �� ����������, ��� ��� �������� ������ ���� ��� �� ���������� ������
    pref = UCase(Left(Buyer, 1)) + Right(CStr(Year(dat)), 2) + CStr(Month(dat)) + CStr(Day(dat))
    
    If Not Nums.exists(pref) Then Nums.Add pref, 0
    Nums(pref) = Nums(pref) + 1
    Generate = pref + Right(CStr(Nums(pref) + 1000), 3)
End Function