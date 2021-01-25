Attribute VB_Name = "Verify"
Dim Comment As String       '������ � �������������
Dim errors As Boolean       '���� ������� ������
Dim groups As Variant       '������� �����
Dim dateS As Variant        '������� ��� �����������
Dim limitOne As Variant     '����� ����� �� �������� ������ ����������
Dim limitAll As Variant     '����� ����� �� ��������
Dim buyers As Variant       '������� ����������� "� ���� ��������"
Dim qrtIndexes As Variant   '������� ������� ��������

'������������� �������� �������
Sub Init()
    
    Set dateS = CreateObject("Scripting.Dictionary")
    Set limitPrs = CreateObject("Scripting.Dictionary")
    Set summS = CreateObject("Scripting.Dictionary")
    Set summOne = CreateObject("Scripting.Dictionary")
    Set summAll = CreateObject("Scripting.Dictionary")
    Set groups = CreateObject("Scripting.Dictionary")
    Set buyers = CreateObject("Scripting.Dictionary")
    Set selIndexes = CreateObject("Scripting.Dictionary")
    Set qrtIndexes = CreateObject("Scripting.Dictionary")
    
    '������ ����� �������
    limitOne = DIC.Cells(1, 5)
    limitAll = DIC.Cells(2, 5)
    
    '������ �������� ��� �����������, ������� �������� � �����
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        cmp = DIC.Cells(i, cINN).text
        dtt = DIC.Cells(i, cSDate)
        dateS(cmp) = dtt
        grp = DIC.Cells(i, cGroup).text
        groups(cmp) = grp
        selIndexes(cmp) = i
        Range(DIC.Cells(i, cPFact), DIC.Cells(i, cPFact + quartCount - 1)).NumberFormat = "### ### ##0.00"
        i = i + 1
    Loop
    
    '�������������� ���������
    For i = 0 To quartCount - 1
        qrtIndexes(IndexToQuartal(i)) = i
    Next

End Sub

'�������� ������������ ������, ���������� true ���� ��� ������
'iC - ������ � ������
'iI - ������ � ����������
Function Verify(ByVal iC As Long, ByVal iI As Long, ByVal oldINN, ByVal oldSum) As Boolean
    
    Comment = ""
    errors = False
    Verify = True
    
    '2 - ����
    If Not isDateMy(DAT.Cells(iC, 2).text) Then
        DAT.Cells(iC, 2).Interior.Color = colRed
        SRC.Cells(iI, 2).Interior.Color = colRed
        AddCom "���� ������� �� ���������"
    Else
        DateTest iC
    End If
    
    '3 - ���/���
    If Not isINNKPP(DAT.Cells(iC, 3).text) Then
        DAT.Cells(iC, 3).Interior.Color = colRed
        SRC.Cells(iI, 3).Interior.Color = colRed
        AddCom "�������� ���/��� ����������"
    End If
    
    '5 - ���
    If Not isINN(DAT.Cells(iC, 5).text) Then
        DAT.Cells(iC, 5).Interior.Color = colRed
        SRC.Cells(iI, 5).Interior.Color = colRed
        AddCom "�������� ��� ��������"
    Else
        If selIndexes(DAT.Cells(iC, 5).text) = Empty Then _
                AddCom "��� " + sel + " �� ������ � �����������"
    End If
    
    
    '7 - ���������
    If Not isPrice(DAT.Cells(iC, 7)) Then
        DAT.Cells(iC, 7).Interior.Color = colRed
        SRC.Cells(iI, 7).Interior.Color = colRed
        AddCom "����� � ��� ������� �� ���������"
    End If
    
    '8 - ������ ���
    If Not isNDS(DAT.Cells(iC, 8).text) Then
        DAT.Cells(iC, 8).Interior.Color = colRed
        SRC.Cells(iI, 8).Interior.Color = colRed
        AddCom "�������� ������ ���"
    End If
    
    '9-11 - ��������� ������ ���������� �������
    For i = 9 To 11
        If Not isPriceNDS(DAT.Cells(iC, i)) Then
            DAT.Cells(iC, i).Interior.Color = colRed
            SRC.Cells(iI, i).Interior.Color = colRed
            errors = True
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
        errors = True
    End If
    
    '���� ��� ������ � ������������ �����, ��������� �������� �� ������
    If Not errors Then LimitsTest iC, iI, oldINN, oldSum
    
    '����� ����������� � ������������� ���
    col = colRed
    If Not errors Then col = colGreen: Comment = "�������"
    DAT.Cells(iC, cCom) = Comment
    DAT.Cells(iC, cCom).Interior.Color = col
    SRC.Cells(iI, cCom) = Comment
    SRC.Cells(iI, cCom).Interior.Color = col
    
    Verify = Not errors
    
End Function

'�������� ������������ ����
Sub DateTest(ByVal i As Long)
    sel = DAT.Cells(i, 6)
    dtt = DAT.Cells(i, 2)
    If dtt < dateS(sel) Then AddCom "���� �� �� ����� ���� ����� ����������� ��������"
End Sub

'��������� �� ���� ���+�������
Function Kvartal(sdata As Variant) As String
    On Error GoTo er
    Kvartal = CStr(Year(sdata)) + CStr((Month(sdata) - 1) \ 3 + 1)
    Exit Function
er:
    Kvartal = ""
End Function

'�������� �������
'i, si - ������ ����� ������ � �������
'oldINN, oldSum - ������� ��� �������� � ������� ����� (���� ��� ������������)
Sub LimitsTest(ByVal i As Long, ByVal si As Long, ByVal oldINN, ByVal oldSum)
    cod = DAT.Cells(i, cCode).text + "!"
    kv = Kvartal(DAT.Cells(i, 2))
    kvin = qrtIndexes(kv)
    sel = DAT.Cells(i, cSellINN).text
    selCur = sel + "!" + kv + "!"
    buy = DAT.Cells(i, cBuyINN).text
    buyCur = buy + "!" + kv + "!"
    grp = groups(sel)
    Sum = 0
    For j = 12 To 14
        If IsNumeric(DAT.Cells(i, j)) Then Sum = Sum + DAT.Cells(i, j)
    Next
    summS(cod + selCur + buy) = summS(cod + selCur + buy) + Sum
    summOne(selCur + buy) = summOne(selCur + buy) + Sum
    summAll(selCur) = summAll(selCur) + Sum
    e = False
    
    '�������� �� ����� ������ ����������
    If summOne(selCur + buy) > limitOne Then _
            AddCom "�������� ����� ������ ������� �������� ������� ����������": e = True
    
    '�������� �� �������
    'RestoreBalance DAT.Cells(i, 2), oldINN, oldSum
    ind = selIndexes(sel)
    over = False
    For j = 0 To kvin
        If Sum > DIC.Cells(ind, cLimits + j) Then over = True
    Next
    If Not over Then
        DIC.Cells(ind, cPFact + kvin) = DIC.Cells(ind, cPFact + kvin) + Sum
    Else
        AddCom "����� ��������� ��������� ������� � ������� ��������": e = True
    End If
    
    '�������� �� ����� ����� ��������
    If summAll(selCur) > limitAll Then AddCom "�������� ����� ����� ������ ������� ��������": e = True
    
    '������� ����� ���������, ���� ���� ���� ���� ������ � ��������
    If e Then
        DAT.Cells(i, cPrice).Interior.Color = colRed
        SRC.Cells(si, cPrice).Interior.Color = colRed
    End If
    
    '�������� �� ��������� ��������� ��� ������ ����������
    If buyers(buyCur + grp) = "" Then
        buyers(buyCur + grp) = sel
    Else
        If buyers(buyCur + grp) <> sel Then AddCom "������� ��������� �������� ��� ������� ����������"
    End If
End Sub

'�������������� �������
Sub RestoreBalance(dt, oldINN, oldSum)
    kvin = qrtIndexes(Kvartal(dt))
    If oldSum > 0 Then
        ind = selIndexes(oldINN)
        If ind <> Empty Then _
                DIC.Cells(ind, cPFact + kvin) = DIC.Cells(ind, cPFact + kvin) - oldSum
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
        '������ ����� ��� � ������, ��� ���� ���������
        If Not IsError(var) Then
            If var = "" Then isPriceNDS = True
        End If
    End If
End Function

'�������� �� ������������ ���
Function isINN(ByVal str As String) As Boolean
    isINN = False
    If str = "" Then Exit Function
    If IsNumeric(str) And Len(str) = 10 Then isINN = True
End Function

'�������� �� ������������ ���/���
Function isINNKPP(ByVal str As String) As Boolean
    isINNKPP = False
    If str = "" Then isINNKPP = False: Exit Function
    s = Split(str, "/")
    '����������� ����
    If IsNumeric(s(0)) And Len(s(0)) = 10 And UBound(s) > 0 Then
        If IsNumeric(s(1)) And Len(s(1)) = 9 Then isINNKPP = True
    End If
    '��
    If IsNumeric(s(0)) And Len(s(0)) = 12 And UBound(s) = 0 Then isINNKPP = True
End Function

'�������� �� ������������ ���
Function isNDS(ByVal str As String) As Boolean
    isNDS = False
    If str = "10" Then isNDS = True
    If str = "18" Then isNDS = True
    If str = "20" Then isNDS = True
End Function