Attribute VB_Name = "Values"
'������������ ������ "������ ������"
Sub CreateReport()

    Message "������������ ������ �� ������� ������..."
    Dim i As Long
    
    '�������� ������� ������������ � ��� ��������
    Set summS = CreateObject("Scripting.Dictionary")
    Set buyList = CreateObject("Scripting.Dictionary")
    Set sellList = CreateObject("Scripting.Dictionary")
    i = firstDat
    Do While DAT.Cells(i, cAccept) <> ""
        If DAT.Cells(i, cAccept) = "OK" Then
            cod = DAT.Cells(i, cCode).text
            kv = Kvartal(DAT.Cells(i, cDates))
            buy = DAT.Cells(i, cBuyINN).text
            sell = DAT.Cells(i, cSellINN).text
            Sum = 0
            For j = 12 To 14
                If DAT.Cells(i, j) <> "" Then Sum = Sum + DAT.Cells(i, j)
            Next
            ID = cod + "!" + kv + "!" + sell + "!" + buy
            summS(ID) = summS(ID) + Sum
            buyList(DAT.Cells(i, cBuyINN).text) = DAT.Cells(i, cBuyer).text
            'sellList(DAT.Cells(i, cSellINN).text) = DAT.Cells(i, cSeller).text
        End If
        i = i + 1
    Loop
    
    '�������� ������� ��� � ������� ��������
    Dim statList As Variant
    Set statList = CreateObject("Scripting.Dictionary")
    i = firstDic
    Do While DIC.Cells(i, cINN) <> ""
        sellList(DIC.Cells(i, cINN).text) = DIC.Cells(i, cSellerName).text
        statList(DIC.Cells(i, cINN).text) = DIC.Cells(i, cPStat).text
        i = i + 1
    Loop
    
    '�������� ������� ������� � �����
    Set clients = CreateObject("Scripting.Dictionary")
    Set templates = CreateObject("Scripting.Dictionary")
    i = firstTempl
    Do While TMP.Cells(i, 3).text <> ""
        clients(TMP.Cells(i, 3).text) = TMP.Cells(i, 1).text
        templates(TMP.Cells(i, 3).text) = TMP.Cells(i, 2).text
        i = i + 1
    Loop
    
    '���������� �����
    Range(VAL.Cells(4, 1), VAL.Cells(maxRow, 7)).Clear
    VAL.Cells(4, 1) = "������"
    VAL.Cells(4, 2) = "�����"
    VAL.Cells(4, 3) = "��������"
    VAL.Cells(4, 4) = "�����"
    VAL.Cells(4, 3) = "�������"
    VAL.Cells(4, 4) = "��������"
    VAL.Cells(4, 5) = "������"
    VAL.Cells(4, 6) = "����������"
    VAL.Cells(4, 7) = "�����"
    Range(VAL.Cells(4, 1), VAL.Cells(4, 7)).Interior.Color = colGray
    
    '������������ ������
    i = 5
    Dim s As Variant
    Dim sel As Variant
    For Each sel In summS
        s = Split(sel, "!")
        VAL.Cells(i, 1) = clients(s(0))
        VAL.Cells(i, 2) = templates(s(0))
        VAL.Cells(i, 3) = s(1)
        VAL.Cells(i, 4) = sellList(s(2))
        VAL.Cells(i, 5) = statList(s(2))
        VAL.Cells(i, 6) = buyList(s(3)) + " (" + s(3) + ")"
        VAL.Cells(i, 7).NumberFormat = "### ### ##0.00"
        VAL.Cells(i, 7) = summS(sel)
        i = i + 1
    Next
    Range(VAL.Cells(4, 1), VAL.Cells(i - 1, 7)).Rows.AutoFilter
    Message "������!"
    
End Sub