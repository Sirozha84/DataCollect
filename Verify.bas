Attribute VB_Name = "Verify"
Dim Comment As String   '������ � �������������
Dim errors As Boolean   '���� ������� ������
Dim groups As Variant   '������� �����
Dim dateS As Variant    '������� ��� �����������
Dim limitPrs As Variant '������� ������� �� ��������
Dim limitOne As Variant '����� ����� �� �������� ������ ����������
Dim limitAll As Variant '����� ����� �� ��������
Dim summOne As Variant  '�������� ���� ������� ������ ����������
Dim summAll As Variant  '�������� ���� ������� ����
Dim buyers As Variant   '������� ����������� "� ���� ��������"

'������������� �������� �������
Sub Init()
    
    Set dateS = CreateObject("Scripting.Dictionary")
    Set limitPrs = CreateObject("Scripting.Dictionary")
    Set summOne = CreateObject("Scripting.Dictionary")
    Set summAll = CreateObject("Scripting.Dictionary")
    Set groups = CreateObject("Scripting.Dictionary")
    Set buyers = CreateObject("Scripting.Dictionary")
    
    '������ ����� �������
    limitOne = DIC.Cells(1, cLimits)
    limitAll = DIC.Cells(2, cLimits)
    
    '������ �������� ��� �����������, ������� �������� � �����
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        cmp = DIC.Cells(i, cINN).text
        dtt = DIC.Cells(i, cSDate)
        dateS(cmp) = dtt
        lim = DIC.Cells(i, cLimits)
        limitPrs(cmp) = lim
        grp = DIC.Cells(i, cGroup).text
        groups(cmp) = grp
        i = i + 1
    Loop
   
End Sub

'���������� ������� �������� ��������
Sub SaveValues()
    Dim i As Long
    i = 1
    VAL.Cells.Clear
    VAL.Columns(1).ColumnWidth = 7
    VAL.Columns(2).ColumnWidth = 20
    VAL.Columns(3).ColumnWidth = 20
    VAL.Columns(4).ColumnWidth = 10
    DrawTable summAll, "������ ����� �������� ��������", i
    DrawTable summOne, "����� �������� �� �����������", i
End Sub

'����� �
Sub DrawTable(tabl As Variant, name As String, i As Long)
    VAL.Cells(i, 1) = name
    i = i + 1
    VAL.Cells(i, 1) = "�������"
    VAL.Cells(i, 2) = "��������"
    VAL.Cells(i, 3) = "����������"
    VAL.Cells(i, 4) = "�����"
    Range(VAL.Cells(i, 1), VAL.Cells(i, 100)).Interior.Color = colGray
    i = i + 1
    For Each sel In tabl
        s = Split(sel, "!")
        VAL.Cells(i, 1) = s(1)
        VAL.Cells(i, 2) = s(0)
        VAL.Cells(i, 3) = s(2)
        VAL.Cells(i, 4) = tabl(sel)
        i = i + 1
    Next
    i = i + 1
End Sub

'�������� ������������ ������, ���������� true ���� ���� ������
'dat - ������� � �������
'src - ������� � �����������
'iC - ������ � ������
'iI - ������ � ����������
Function Verify(ByRef DAT As Variant, ByRef SRC As Variant, ByVal iC As Long, _
                                                            ByVal iI As Long) As Boolean
    
    Comment = ""
    errors = False
    Verify = True
    
    '2 - ����
    If Not isDateMy(DAT.Cells(iC, 2).text) Then
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
        If Not isPriceNDS(DAT.Cells(iC, i)) Then
            DAT.Cells(iC, i).Interior.Color = colRed
            SRC.Cells(iI, i).Interior.Color = colRed
            AddCom "��������� ������ ���������� ������� ������� �� ���������"
        End If
    Next
    
    '12-14 - ����� ���
    e = False
    For i = 12 To 14
        If Not isPriceNDS(DAT.Cells(iC, i)) Then e = True
    Next
    If e Then
        DAT.Cells(iC, i).Interior.Color = colRed
        SRC.Cells(iI, i).Interior.Color = colRed
        AddCom "����� ��� ������� �� ���������"
    Else
        LimitsTest DAT, iC
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
Function Kvartal(DAT As Variant) As String
    On Error GoTo er
    Kvartal = CStr(Year(DAT)) + CStr((Month(DAT) - 1) \ 3 + 1)
    Exit Function
er:
    Kvartal = ""
End Function

'�������� �������
Sub LimitsTest(ByRef DAT As Variant, ByVal i As Long)
    kv = "!" + Kvartal(DAT.Cells(i, 2)) + "!"
    sel = DAT.Cells(i, cSellINN).text
    selCur = sel + kv
    buy = DAT.Cells(i, cBuyINN).text
    buyCur = buy + kv
    grp = groups(sel)
    Sum = 0
    For j = 12 To 14
        If IsNumeric(DAT.Cells(i, j)) Then Sum = Sum + DAT.Cells(i, j)
    Next
    summOne(selCur + buy) = summOne(selCur + buy) + Sum
    summAll(selCur) = summAll(selCur) + Sum
    If summOne(selCur + buy) > limitOne Then AddCom "�������� ����� ����� ������ ������ ����������"
    If summAll(selCur) > limitPrs(sel) Then AddCom "�������� ����� ��������" '������������
    If summAll(selCur) > limitAll Then AddCom "�������� ����� ����� ������"
    If buyers(buyCur + grp) = "" Then
        buyers(buyCur + grp) = sel
    Else
        If buyers(buyCur + grp) <> sel Then AddCom "������� � ������� �������� ������"
    End If
End Sub

'���������� ����������� � ������
Sub AddCom(str As String)
    If Comment <> "" Then Comment = Comment + ", "
    Comment = Comment + str
    errors = True
End Sub

'�������� �� ������������ ����
Function isDateMy(ByVal var As String)
    isDateMy = True
    On Error GoTo er
    s = Split(var, ".")
    If UBound(s) <> 2 Then GoTo er
    If CInt(s(0)) < 1 Or CInt(s(0)) > 31 Then GoTo er
    If CInt(s(1)) < 1 Or CInt(s(1)) > 12 Then GoTo er
    If CInt(s(2)) < 1900 Or CInt(s(2)) > 2100 Then GoTo er
    If Not IsDate(var) Then GoTo er
    Exit Function
er:
    isDateMy = False
End Function

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