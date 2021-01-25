Attribute VB_Name = "Values"
Public summS As Variant  '�������� ���� ������� ������ ���������� � ������� ����
Public summOne As Variant  '�������� ���� ������� ������ ����������
Public summAll As Variant  '�������� ���� ������� ����

'���������� ������� �������� ��������
Sub CreateReport()
    
    Message "������������ ������ �� ������� ������"
    Dim i As Long
    
    '�������� ������� ������������ � ��� ��������
    Set buyList = CreateObject("Scripting.Dictionary")
    Set sellList = CreateObject("Scripting.Dictionary")
    i = firstDat
    Do While DAT.Cells(i, cAccept) <> ""
        buyList(Cells(i, cBuyINN).text) = Cells(i, cBuyer).text
        sellList(Cells(i, cSellINN).text) = Cells(i, cSeller).text
        i = i + 1
    Loop
    
    '�������� ������� ��� � ������� ��������
    Dim statList As Variant
    Set statList = CreateObject("Scripting.Dictionary")
    i = firstDic
    Do While DIC.Cells(i, cINN) <> ""
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
    
    '��������� �����
    i = 1
    VAL.Cells.Clear
    VAL.Columns(1).ColumnWidth = 20
    VAL.Columns(2).ColumnWidth = 20
    VAL.Columns(3).ColumnWidth = 10
    VAL.Columns(4).ColumnWidth = 20
    VAL.Columns(5).ColumnWidth = 20
    VAL.Columns(6).ColumnWidth = 30
    VAL.Columns(7).ColumnWidth = 15
    VAL.Cells(1, 1) = "������"
    VAL.Cells(1, 2) = "�����"
    VAL.Cells(1, 3) = "��������"
    VAL.Cells(1, 4) = "�����"
    VAL.Cells(1, 3) = "�������"
    VAL.Cells(1, 4) = "��������"
    VAL.Cells(1, 5) = "������"
    VAL.Cells(1, 6) = "����������"
    VAL.Cells(1, 7) = "�����"
    Range(VAL.Cells(1, 1), VAL.Cells(1, 100)).Interior.Color = colGray
    i = i + 1
    Dim s As Variant
    Dim sel As Variant
    For Each sel In summS
        s = Split(sel, "!")
        cl = clients(s(0))
        If cl <> Empty Then VAL.Cells(i, 1) = cl Else VAL.Cells(i, 1) = "���: " + s(0)
        frm = templates(s(0))
        If frm <> Empty Then VAL.Cells(i, 2) = frm
        VAL.Cells(i, 3) = s(2)
        VAL.Cells(i, 4) = sellList(s(1))
        VAL.Cells(i, 5) = statList(s(1))
        VAL.Cells(i, 6) = buyList(s(3)) + " (" + s(3) + ")"
        VAL.Cells(i, 7).NumberFormat = "### ### ##0.00"
        VAL.Cells(i, 7) = summS(sel)
        i = i + 1
    Next
    Range(VAL.Cells(1, 1), VAL.Cells(1, 7)).Rows.AutoFilter
    
End Sub