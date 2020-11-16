Attribute VB_Name = "Template"
Const MaxRecords = 100  '������������ ���������� �������
Const FirstClient = 6   '������ ������ ������ ��������
Const Secret = "123"
Const maxBuyers = 100   '������������ ���������� �����������
Const maxSellers = 100  '������������ ���������� ���������

Dim temp As Variant
Dim dat As Variant

Public Sub Generate()
    '������� ������������ ��� (���� �� ����)
    Dim i As Long
    i = FirstClient
    max = 0
    Do While Cells(i, 1) <> ""
        If max < Cells(i, 2) Then max = Cells(i, 2)
        i = i + 1
    Loop
    
    '����������� ���� ��������� ��������, � ������� ��� ���
    i = FirstClient
    Do While Cells(i, 1) <> ""
        If Cells(i, 2) = "" Then max = max + 1: Cells(i, 2) = max
        i = i + 1
    Loop
        
    '���������� �������
    Dim total As Long
    total = i - 1
    Set dat = Application.ActiveSheet
    For i = FirstClient To total
        Message "�������� ������� " + CStr(i - FirstClient + 1) + " �� " + _
        CStr(total - FirstClient + 1)
        Call NewTemplate(Cells(i, 1), Cells(i, 2))
    Next
    
    Message "������!"
End Sub

'�������� ������ �����
Sub NewTemplate(name As String, cod As Long)
    '������ ���� � ������� ���������
    Filename = Cells(1, 3) + "\" + name + ".xlsx"
    Workbooks.Add
    On Error GoTo er2
    Application.DisplayAlerts = False
    Sheets.Add
    Sheets.Add
    Sheets(1).name = name
    Sheets(2).name = "����������"
    Sheets(3).name = "��������"
    Sheets(4).Delete
    Sheets(4).Delete
er2:
    
    On Error GoTo er
    Set temp = Application.ActiveSheet
    Set listb = Sheets(2)
    Set lists = Sheets(3)
    Cells(1, 1) = cod
    
    '������� �� �������������
    listb.Columns(1).ColumnWidth = 20
    listb.Columns(2).ColumnWidth = 10
    listb.Cells(1, 1) = "����������"
    listb.Cells(1, 2) = "���/���"
    lists.Columns(1).ColumnWidth = 20
    lists.Columns(2).ColumnWidth = 10
    lists.Cells(1, 1) = "��������"
    lists.Cells(1, 2) = "���"
    
    '�������� �������. ������ ����� �����
    Columns(1).ColumnWidth = 20
    Columns(2).ColumnWidth = 15
    Columns(3).ColumnWidth = 30
    Columns(4).ColumnWidth = 15
    Columns(5).ColumnWidth = 10
    Columns(6).ColumnWidth = 15
    Columns(7).ColumnWidth = 15
    Columns(8).ColumnWidth = 10
    Columns(9).ColumnWidth = 12
    Columns(10).ColumnWidth = 12
    Columns(11).ColumnWidth = 12
    Columns(12).ColumnWidth = 12
    Columns(13).ColumnWidth = 12
    Columns(14).ColumnWidth = 12
    Rows(3).RowHeight = 30
    Rows(4).RowHeight = 30
    Cells(3, 1) = "��"
    Range(Cells(3, 1), Cells(3, 2)).Merge
    Cells(3, 3) = "�������� � ����������"
    Range(Cells(3, 3), Cells(3, 4)).Merge
    Cells(3, 5) = "�������� � ��������"
    Range(Cells(3, 5), Cells(3, 6)).Merge
    Cells(3, 7) = "���������" + Chr(10) + "������ � ���"
    Cells(3, 8) = "������" + Chr(10) + "���, %"
    Range(Cells(3, 8), Cells(4, 8)).Merge
    Cells(3, 9) = "��������� ������ ���������� �������" + Chr(10) + "(� ���.) ��� ���"
    Range(Cells(3, 9), Cells(3, 11)).Merge
    Cells(3, 12) = "����� ���"
    Range(Cells(3, 12), Cells(3, 14)).Merge
    Cells(4, 1) = "�" + Chr(10) + "(���. 020)"
    Cells(4, 2) = "����" + Chr(10) + "(���. 030)"
    Cells(4, 3) = "���/���"
    Cells(4, 4) = "������������"
    Cells(4, 5) = "���"
    Cells(4, 6) = "������������"
    Cells(4, 7) = "� ���. � ���."
    Cells(4, 9) = "20%" + Chr(10) + "(���. 170)"
    Cells(4, 10) = "18%" + Chr(10) + "(���. 200)"
    Cells(4, 11) = "10%" + Chr(10) + "(���. 205)"
    Cells(4, 12) = "20%" + Chr(10) + "(���. 200)"
    Cells(4, 13) = "18%" + Chr(10) + "(���. 205)"
    Cells(4, 14) = "10%" + Chr(10) + "(���. 210)"
    Set hat = Range(Cells(3, 1), Cells(4, 14))
    hat.HorizontalAlignment = xlCenter
    hat.VerticalAlignment = xlCenter
    hat.Interior.Color = RGB(224, 224, 224)
    hat.Borders.Weight = 3
    
    '���� 2 - ����
    Call setFormat(2, "date")
    Call setValidation(2, "date")
    Call allowEdit(2, "����")
    '���� 3 - ��� ����������, ��������� � ������� ���
    For i = 5 To 4 + MaxRecords
        Cells(i, 3).FormulaLocal = "=���(D" + CStr(i) + ";����������!A2:B" + _
        CStr(maxBuyers) + ";2;0)"
    Next
    setFormatConditions (3)
    '���� 4 - ����������, �������� �� ������
    Call setValidation(4, "b")
    Call allowEdit(4, "����������")
    '���� 5 - ��� ��������, ��������� � ������� ���
    For i = 5 To 4 + MaxRecords
        Cells(i, 5).FormulaLocal = "=���(F" + CStr(i) + ";��������!A2:B" + _
        CStr(maxSellers) + ";2;0)"
    Next
    setFormatConditions (5)
    '���� 6 - ��������, �������� �� ������
    Call setValidation(6, "s")
    Call allowEdit(6, "��������")
    '���� 7 - ���������
    Call setFormat(7, "money")
    Cells(1, 7).Borders.Weight = 3
    Cells(1, 7).FormulaLocal = "=����(G5:G" + CStr(4 + MaxRecords) + ")"
    Call allowEdit(7, "���������")
    '���� 8 - ������ ���
    Call setValidation(8, "nds")
    Call allowEdit(8, "������ ���")
    '����� 9-14
    For i = 9 To 14
        Call setFormat(i, "money")
        Cells(1, i).Borders.Weight = 3
    Next
    '���� 9-11 - ����� � ��� 20,18,10%      ������� G/(100+H)*100
    For i = 5 To 4 + MaxRecords
        Cells(i, 9).FormulaLocal = "=����(�(G" + CStr(i) + "<>"""";H" + CStr(i) + "=20);" + _
        "������(G" + CStr(i) + "/(100+H" + CStr(i) + ")*100;2);"""")"
        Cells(i, 10).FormulaLocal = "=����(�(G" + CStr(i) + "<>"""";H" + CStr(i) + "=18);" + _
        "������(G" + CStr(i) + "/(100+H" + CStr(i) + ")*100;2);"""")"
        Cells(i, 11).FormulaLocal = "=����(�(G" + CStr(i) + "<>"""";H" + CStr(i) + "=10);" + _
        "������(G" + CStr(i) + "/(100+H" + CStr(i) + ")*100;2);"""")"
    Next
    Cells(1, 9).FormulaLocal = "=����(I5:I" + CStr(4 + MaxRecords) + ")"
    Cells(1, 10).FormulaLocal = "=����(J5:J" + CStr(4 + MaxRecords) + ")"
    Cells(1, 11).FormulaLocal = "=����(K5:K" + CStr(4 + MaxRecords) + ")"
    '���� 12-14 - ����� ��� ��� 20,18,10%   ������� G/(100+H)*H
    For i = 5 To 4 + MaxRecords
        Cells(i, 12).FormulaLocal = "=����(�(G" + CStr(i) + "<>"""";H" + CStr(i) + "=20);" + _
        "������(G" + CStr(i) + "/(100+H" + CStr(i) + ")*H" + CStr(i) + ";2);"""")"
        Cells(i, 13).FormulaLocal = "=����(�(G" + CStr(i) + "<>"""";H" + CStr(i) + "=18);" + _
        "������(G" + CStr(i) + "/(100+H" + CStr(i) + ")*H" + CStr(i) + ";2);"""")"
        Cells(i, 14).FormulaLocal = "=����(�(G" + CStr(i) + "<>"""";H" + CStr(i) + "=10);" + _
        "������(G" + CStr(i) + "/(100+H" + CStr(i) + ")*H" + CStr(i) + ";2);"""")"
    Next
    Cells(1, 12).FormulaLocal = "=����(L5:L" + CStr(4 + MaxRecords) + ")"
    Cells(1, 13).FormulaLocal = "=����(M5:M" + CStr(4 + MaxRecords) + ")"
    Cells(1, 14).FormulaLocal = "=����(N5:N" + CStr(4 + MaxRecords) + ")"
    
    
    '������ �����
    temp.Protect Secret, UserInterfaceOnly:=True
    'lists.Protect Secret, UserInterfaceOnly:=True
    
    '���� ��������� ����� ������� ��������� - ���������������� ���������� ������
    ActiveWorkbook.SaveAs Filename:=Filename
    ActiveWorkbook.Close
er:
End Sub

'��������� ������� ��� �������
Sub setFormat(ByVal c As Integer, format As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    If format = "date" Then rang.NumberFormat = "dd.MM.yyyy"
    If format = "money" Then rang.NumberFormat = "### ### ##0.00"
End Sub

'��������� ��������� �������������� ��� �������
Sub setFormatConditions(c As Integer)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    With rang.FormatConditions
        .Add Type:=16
        .Item(.count).Font.Color = vbWhite
    End With
End Sub

'��������� �������� ��������
Sub setValidation(c As Integer, typ As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    If typ = "b" Then formul = "=����������!$A$2:$A$" + CStr(maxBuyers)
    If typ = "s" Then formul = "=��������!$A$2:$A$" + CStr(maxSellers)
    If typ = "nds" Then formul = "10,18,20"
    If formul <> "" Then
        With rang.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formul
            .ErrorMessage = "������ �� ������, ����������!"
        End With
    End If
    If typ = "date" Then
        With rang.Validation
            .Delete
            '30000 - �����-�� ���� 82-�� ����, ��� � �� ����� ��� �������� ������������ ����
            .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="30000"
            .ErrorMessage = "��� ������ ���� ����!"
        End With
    End If
End Sub

'��������� ���������� �������������� ��� �������
Sub allowEdit(c As Integer, name As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    temp.Protection.AllowEditRanges.Add Title:=name, Range:=rang, Password:=""
    rang.Interior.Color = RGB(255, 255, 192)
End Sub