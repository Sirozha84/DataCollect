Attribute VB_Name = "Values"
Option Explicit

Public summOne As Variant  '�������� ���� ������� ������ ����������

'���������� ������� �������� ��������
Sub CreateReport()
    
    Message "������������ ������ �� ������� ������"
    Dim i As Long
    
    '�������� ������� ������������ � ��� ��������
    Dim buyList As Variant
    Set buyList = CreateObject("Scripting.Dictionary")
    Dim sellList As Variant
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
    
    '��������� �����
    i = 1
    VAL.Cells.Clear
    VAL.Columns(1).ColumnWidth = 7
    VAL.Columns(2).ColumnWidth = 20
    VAL.Columns(3).ColumnWidth = 20
    VAL.Columns(4).ColumnWidth = 20
    VAL.Columns(5).ColumnWidth = 20
    VAL.Cells(1, 1) = "�������"
    VAL.Cells(1, 2) = "��������"
    VAL.Cells(1, 3) = "������"
    VAL.Cells(1, 4) = "����������"
    VAL.Cells(1, 5) = "�����"
    Range(VAL.Cells(1, 1), VAL.Cells(1, 100)).Interior.Color = colGray
    i = i + 1
    Dim s As Variant
    Dim sel As Variant
    For Each sel In summOne
        s = Split(sel, "!")
        VAL.Cells(i, 1) = s(1)
        VAL.Cells(i, 2) = sellList(s(0))
        VAL.Cells(i, 3) = statList(s(0))
        VAL.Cells(i, 4) = buyList(s(2))
        VAL.Cells(i, 5).NumberFormat = "### ### ##0.00"
        VAL.Cells(i, 5) = summOne(sel)
        i = i + 1
    Next
    
    Range(VAL.Cells(1, 1), VAL.Cells(1, 5)).Rows.AutoFilter
End Sub