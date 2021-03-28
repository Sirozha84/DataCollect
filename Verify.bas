Attribute VB_Name = "Verify"
'��������� ������: 28.03.2021 18:42

Dim Comment As String       '������ � �������������
Dim errors As Boolean       '���� ������� ������
Dim groups As Variant       '������� �����
Dim dateS As Variant        '������� ��� �����������
Dim limitOne As Variant     '����� ����� �� �������� ������ ����������
Dim limitAll As Variant     '����� ����� �� ��������
Dim buyers As Variant       '������� ����������� "� ���� ��������"
Dim qrtIndexes As Variant   '������� ������� ��������
Dim summOne As Variant      '�������� ���� ������� ������ ����������
Dim summAll As Variant      '�������� ���� ������� ����

'������������� �������� �������
Sub Init()
    
    Set dateS = CreateObject("Scripting.Dictionary")
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

'�������� ������������ ������ ��������, ���������� true ���� ��� ������
'Di - ������ � ������
'Si - ������ � ����������
Function Verify(ByVal Di As Long, ByVal Si As Long, ByVal oldINN, ByVal oldSum) As Boolean
    
    Comment = ""
    errors = False
    Verify = True
    
    '2 - ����
    If Not isDateMy(DAT.Cells(Di, 2).text) Then
        DAT.Cells(Di, 2).Interior.Color = colRed
        SRC.Cells(Si, 2).Interior.Color = colRed
        AddCom "���� ������� �� ���������"
    Else
        DateTest Di
    End If
    
    '3 - ���/���
    If Not isINNKPP(DAT.Cells(Di, 3).text) Then
        DAT.Cells(Di, 3).Interior.Color = colRed
        SRC.Cells(Si, 3).Interior.Color = colRed
        AddCom "�������� ���/��� ����������"
    End If
    
    '5 - ���
    If Not isINN(DAT.Cells(Di, 5).text) Then
        DAT.Cells(Di, 5).Interior.Color = colRed
        SRC.Cells(Si, 5).Interior.Color = colRed
        AddCom "�������� ��� ��������"
    Else
        If selIndexes(DAT.Cells(Di, 5).text) = Empty Then AddCom "��� �� ������ � �����������"
    End If
    
    '7 - ���������
    If Not isPrice(DAT.Cells(Di, 7)) Then
        DAT.Cells(Di, 7).Interior.Color = colRed
        SRC.Cells(Si, 7).Interior.Color = colRed
        AddCom "����� � ��� ������� �� ���������"
    End If
    
    '8 - ������ ���
    If Not isNDS(DAT.Cells(Di, 8).text) Then
        DAT.Cells(Di, 8).Interior.Color = colRed
        SRC.Cells(Si, 8).Interior.Color = colRed
        AddCom "�������� ������ ���"
    End If
    
    '9-11 - ��������� ������ ���������� �������
    For i = 9 To 11
        If Not isPriceNDS(DAT.Cells(Di, i)) Then
            DAT.Cells(Di, i).Interior.Color = colRed
            SRC.Cells(Si, i).Interior.Color = colRed
            errors = True
        End If
    Next
    
    '12-14 - ����� ���
    e = False
    For i = 12 To 14
        If Not isPriceNDS(DAT.Cells(Di, i)) Then e = True
    Next
    If e Then
        DAT.Cells(Di, i).Interior.Color = colRed
        SRC.Cells(Si, i).Interior.Color = colRed
        errors = True
    End If
    
    '���� ��� ������ � ������������ �����, ��������� �������� �� ������
    If Not errors Then LimitsTest Di, Si, oldINN, oldSum
    
    '����� ����������� � ������������� ���
    col = colRed
    If Not errors Then col = colGreen: Comment = "�������"
    DAT.Cells(Di, cCom) = Comment
    DAT.Cells(Di, cCom).Interior.Color = col
    SRC.Cells(Si, cCom) = Comment
    SRC.Cells(Si, cCom).Interior.Color = col
    
    Verify = Not errors
    
End Function

'�������� ������������ ������ �����������, ���������� true ���� ��� ������
'i - ����� ������
Function VerifyLoad(ByVal i As Long) As Boolean
    
    Comment = ""
    errors = False
    VerifyLoad = True

    '����
    If Not isDateMy(DTL.Cells(i, clDate).text) Then
        DTL.Cells(i, clDate).Interior.Color = colRed
        AddCom "���� ������� �� ���������"
    End If

    '��� ����������
    If Not isINNKPP(DTL.Cells(i, clProvINN).text) Then
        DTL.Cells(i, clProvINN).Interior.Color = colRed
        AddCom "�������� ��� ����������"
    End If

    '��� ��������
    If Not isINN(DTL.Cells(i, clSaleINN).text) Then
        DTL.Cells(i, clSaleINN).Interior.Color = colRed
        AddCom "�������� ��� ��������"
    Else
        If Not CorrectSaler(DTL.Cells(i, clSaleINN).text, DTL.Cells(i, clSaleName).text) Then
            AddCom "������������ � ��� �������� �� ��������� � ������� � �����������"
        End If
    End If

    '���������
    If Not isPrice(DTL.Cells(i, clPrice)) Then
        DTL.Cells(i, clPrice).Interior.Color = colRed
        AddCom "����� � ��� ������� �� ���������"
    End If
    
    '9-11 - ��������� ������ ���������� �������
    For j = 9 To 11
        If Not isPriceNDS(DTL.Cells(i, j)) Then
            DTL.Cells(i, j).Interior.Color = colRed
            errors = True
        End If
    Next
    
    '12-14 - ����� ���
    e = False
    For j = 12 To 14
        If Not isPriceNDS(DTL.Cells(i, j)) Then e = True
    Next
    If e Then
        DTL.Cells(i, j).Interior.Color = colRed
        errors = True
    End If
    
    '����� ����������� � ������������� ���
    col = colRed
    If Not errors Then col = colGreen: Comment = "�������"
    DTL.Cells(i, clCom) = Comment
    DTL.Cells(i, clCom).Interior.Color = col
    
    VerifyLoad = Not errors
    
End Function

'�������� ������������ ����
Sub DateTest(ByVal i As Long)
    SEL = DAT.Cells(i, 6)
    dtt = DAT.Cells(i, 2)
    If dtt < dateS(SEL) Then AddCom "���� �� �� ����� ���� ����� ����������� ��������"
End Sub

'�������� �������
'Di - ������ � ������
'Si - ������ � ����������
'oldINN, oldSum - ������� ��� �������� � ������� ����� (���� ��� ������������)
Sub LimitsTest(ByVal Di As Long, ByVal Si As Long, ByVal oldINN, ByVal oldSum)
    
    dt = DAT.Cells(Di, 2)
    cod = DAT.Cells(Di, cCode).text + "!"
    kv = Kvartal(dt)
    kvin = qrtIndexes(kv)
    SEL = DAT.Cells(Di, cSellINN).text
    selCur = SEL + "!" + kv + "!"
    BUY = DAT.Cells(Di, cBuyINN).text
    buyCur = BUY + "!" + kv + "!"
    grp = groups(SEL)
    Sum = 0
    For j = 12 To 14
        If IsNumeric(DAT.Cells(Di, j)) Then Sum = Sum + DAT.Cells(Di, j)
    Next
    summOne(selCur + BUY) = summOne(selCur + BUY) + Sum
    summAll(selCur) = summAll(selCur) + Sum
    e = False
    ind = selIndexes(SEL)
    
    '�������� �� ������ �������� � ���� �������
    If DIC.Cells(ind, cSaleProtect + DateToQIndex(dt)) = "��" Then
        AddCom "������ �������� �� ����� ��������� � ������ �������"
        Exit Sub '���������� �������� ������������
    End If
    
    '�������� �� ����� ������ ����������
    If summOne(selCur + BUY) > limitOne Then _
            AddCom "�������� ����� ������ ������� �������� ������� ����������": e = True
    
    '�������� �� �������
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
        DAT.Cells(Di, cPrice).Interior.Color = colRed
        SRC.Cells(Si, cPrice).Interior.Color = colRed
    End If
    
    '�������� �� ��������� ��������� ��� ������ ����������
    If buyers(buyCur + grp) = "" Then
        buyers(buyCur + grp) = SEL
    Else
        If buyers(buyCur + grp) <> SEL Then AddCom "������� ��������� �������� ��� ������� ����������"
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
        If var >= 1 And var <> "" Then isPrice = True
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

'******************** End of File ********************