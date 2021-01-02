Attribute VB_Name = "Verify"
Dim Comment As String   '������ � �������������
Dim errors As Boolean   '���� ������� ������
Dim groups As Variant   '������� �����
Dim dateS As Variant    '������� ��� �����������
Dim limitPrs As Variant '������� ������� �� ��������
Dim limitOne As Variant  '����� ����� �� �������� ������ ����������
Dim limitAll As Variant '����� ����� �� ��������
Dim summPrs As Variant    '�������� ���� �� ��������
Dim summOne As Variant   '�������� ���� ������� ������ ����������
Dim summAll As Variant  '�������� ���� ������� ����

'������������� �������� �������
Sub Init()
    
    Set dateS = CreateObject("Scripting.Dictionary")
    Set limitPrs = CreateObject("Scripting.Dictionary")
    Set summPrs = CreateObject("Scripting.Dictionary")
    Set summOne = CreateObject("Scripting.Dictionary")
    Set summAll = CreateObject("Scripting.Dictionary")
    Set groups = CreateObject("Scripting.Dictionary")
    Dim i As Long
    
    '������ ������� ��� ����������� ��������
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        cmp = DIC.Cells(i, 1).text
        dtt = DIC.Cells(i, 2)
        dateS(cmp) = dtt
        i = i + 1
    Loop
    
    '������ ����� �������
    limitOne = DIC.Cells(1, 4)
    limitAll = DIC.Cells(2, 4)

    '������ ������� ������� ��������
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        cmp = DIC.Cells(i, 1).text
        lim = DIC.Cells(i, 4)
        limitPrs(cmp) = lim
        i = i + 1
    Loop
    
    '������ ������� �����
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        cmp = DIC.Cells(i, 1).text
        grp = DIC.Cells(i, 3).text
        groups(cmp) = grp
        i = i + 1
    Loop
   
End Sub

'���������� ������� �������� ��������
Sub SaveValues()

End Sub

'�������� ������������ ������, ���������� true ���� ���� ������
'dat - ������� � �������
'src - ������� � �����������
'iC - ������ � ������
'iI - ������ � ����������
'changed - true ���� ������ ��� ���� ���������������� � ������ ����������� �� ���������
Function Verify(ByRef DAT As Variant, ByRef SRC As Variant, ByVal iC As Long, ByVal iI As Long, _
    changed As Boolean) As Boolean
    
    Comment = ""
    errors = False
    Verify = True
    
    '2 - ����
    DAT.Cells(iC, 2).NumberFormat = "dd.MM.yyyy"
    If Not IsDate(DAT.Cells(iC, 2)) Then
        DAT.Cells(iC, 2).Interior.Color = colRed
        SRC.Cells(iI, 2).Interior.Color = colRed
        AddCom "���� ������� �� ���������"
    Else
        Call DateTest(DAT, iC)
    End If
    
    '3 - ���
    If Not isINNKPP(DAT.Cells(iC, 3).text) Then
        DAT.Cells(iC, 3).Interior.Color = colRed
        SRC.Cells(iI, 3).Interior.Color = colRed
        AddCom "���/��� ������� �� ���������"
    End If
    
    '5 - ���
    If Not isINNKPP(DAT.Cells(iC, 5).text) Then
        DAT.Cells(iC, 5).Interior.Color = colRed
        SRC.Cells(iI, 5).Interior.Color = colRed
        AddCom "��� ����� �� ���������"
    End If
    
    '7 - ���������
    DAT.Cells(iC, 7).NumberFormat = "### ### ##0.00"
    If Not isPrice(DAT.Cells(iC, 7)) Then
        DAT.Cells(iC, 7).Interior.Color = colRed
        SRC.Cells(iI, 7).Interior.Color = colRed
        AddCom "��������� ������� �� ���������"
    End If
    
    '8 - ������ ���
    If Not isNDS(DAT.Cells(iC, 8).text) Then
        DAT.Cells(iC, 8).Interior.Color = colRed
        SRC.Cells(iI, 8).Interior.Color = colRed
        AddCom "��� ����� �� ���������"
    End If
    
    '9-11 - ��������� ������ ���������� �������
    For i = 9 To 11
        DAT.Cells(iC, i).NumberFormat = "### ### ##0.00"
        If Not isPriceNDS(DAT.Cells(iC, i)) Then
            DAT.Cells(iC, i).Interior.Color = colRed
            SRC.Cells(iI, i).Interior.Color = colRed
            AddCom "��������� ������ ���������� ������� ������� �� ���������"
        End If
    Next
    
    '12-14 - ����� ���
    e = False
    For i = 12 To 14
        DAT.Cells(iC, i).NumberFormat = "### ### ##0.00"
        If Not isPriceNDS(DAT.Cells(iC, i)) Then e = True
    Next
    If e Then
        DAT.Cells(iC, i).Interior.Color = colRed
        SRC.Cells(iI, i).Interior.Color = colRed
        AddCom "����� ��� ������� �� ���������"
    Else
        Call LimitsTest(DAT, iC)
    End If
    
    '����� ����������� � ������������� ���
    col = colRed
    If Not errors Then col = colGreen: Comment = "�������"
    DAT.Cells(iC, cCom) = Comment
    DAT.Cells(iC, cCom).Interior.Color = col
    SRC.Cells(iI, cCom) = Comment
    SRC.Cells(iI, cCom).Interior.Color = col
    
    Verify = errors
    
End Function

'�������� ������������ ����
Sub DateTest(ByRef DAT As Variant, ByVal i As Long)
    sel = DAT.Cells(i, 6)
    dtt = DAT.Cells(i, 2)
    If dtt < dateS(sel) Then AddCom "���� �������� �� ����� ���� ����� ����������� ��������"
End Sub

'��������� �� ���� ���+�������
Function Kvartal(DAT As Date) As String
    Kvartal = CStr(Year(DAT)) + CStr((Month(DAT) - 1) \ 3 + 1)
End Function

'�������� �������
Sub LimitsTest(ByRef DAT As Variant, ByVal i As Long)
    sel = DAT.Cells(i, 6)
    selCur = sel + "!" + Kvartal(DAT.Cells(i, 2)) + "!"
    buy = DAT.Cells(i, 4)
    grp = groups(sel)
    Sum = 0
    For j = 12 To 14
        If IsNumeric(DAT.Cells(i, j)) Then Sum = Sum + DAT.Cells(i, j)
    Next
    summPrs(selCur) = summPrs(selCur) + Sum
    summOne(selCur + buy) = summAll(selCur + buy) + Sum
    summAll(selCur) = summAll(selCur) + Sum
    If summPrs(selCur) > limitPrs(sel) Then AddCom "�������� ����� ��������" '������������
    If summOne(selCur + buy) > limitOne Then AddCom "�������� ����� ����� ������ ������ ����������"
    If summAll(selCur) > limitAll Then AddCom "�������� ����� ����� ������"
End Sub

'���������� ����������� � ������
Sub AddCom(str As String)
    If Comment <> "" Then Comment = Comment + ", "
    Comment = Comment + str
    errors = True
End Sub

'�������� �� ������������ ����
Function isPrice(ByVal var As Variant)
    isPrice = False
    If IsNumeric(var) Then
        If var >= 0 And var <> "" Then isPrice = True
    End If
End Function

'�������� �� ������������ ���� ��� ��� � ��� �����
Function isPriceNDS(ByVal var As Variant)
    isPriceNDS = False
    If IsNumeric(var) Then
        If var >= 0 Then isPriceNDS = True
    Else
        If Not IsError(var) Then
            If var = "" Then isPriceNDS = True
        End If
    End If
End Function

'�������� �� ������������ ���/���
Function isINNKPP(ByVal str As String) As Boolean
    If str = "" Then isINNKPP = False: Exit Function
    Dim s() As String
    s = Split(str, "/")
    isINNKPP = True
    If Not IsNumeric(s(0)) Then isINNKPP = False
    If Len(s(0)) <> 10 And Len(s(0)) <> 12 Then isINNKPP = False
    If UBound(s) > 0 Then
        If Not IsNumeric(s(1)) Then isINNKPP = False
        If Len(s(1)) <> 9 Then isINNKPP = False
    End If
End Function

'�������� �� ������������ ���
Function isNDS(ByVal str As String) As Boolean
    isNDS = False
    If str = "10" Then isNDS = True
    If str = "18" Then isNDS = True
    If str = "20" Then isNDS = True
End Function