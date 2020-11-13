Attribute VB_Name = "Template"
Const MaxRecords = 100  '������������ ���������� �������
Const FirstClient = 6   '������ ������ ������ ��������
Const Secret = "123"

Dim temp As Variant
Dim dat As Variant
Dim maxBuyers As Integer
Dim maxSellers As Integer

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
    Filename = Cells(1, 3) + "\" + name + ".xlsx"
    Workbooks.Add
    'On Error GoTo er2
    Application.DisplayAlerts = False
    Sheets.Add
    Sheets(1).name = name
    Sheets(2).name = "�����������"
    Sheets(3).Delete
    Sheets(3).Delete
er2:
    
    'On Error GoTo er
    Set temp = Application.ActiveSheet
    Set lists = Sheets(2)
    Cells(1, 1) = cod
    
    '�������� �����������
    For i = 1 To 4: lists.Columns(i).ColumnWidth = 20: Next
    i = 5
    j = 0
    Do While dat.Cells(i, 3) <> ""
        j = j + 1
        lists.Cells(j, 1) = dat.Cells(i, 3)
        lists.Cells(j, 2) = dat.Cells(i, 4)
        i = i + 1
    Loop
    maxBuyers = j
    i = 5
    j = 0
    Do While dat.Cells(i, 5) <> ""
        j = j + 1
        lists.Cells(j, 3) = dat.Cells(i, 5)
        lists.Cells(j, 4) = dat.Cells(i, 6)
        i = i + 1
    Loop
    maxSellers = j
    
    '������ ����� �����
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
    
    '����2 - ����
    Call setFormat(2, "date")
    Call allowEdit(2, "����")
    '����3 - ���, ��������� ����� ���
    For i = 5 To 4 + MaxRecords
        Cells(i, 3).FormulaLocal = "=���(D" + CStr(i) + ";�����������!A2:B" + _
        CStr(maxBuyers) + ";2;0)"
    Next
    setFormatConditions (3)
    '����3 - ����������, �������� �� ������
    Call setValidation(4, "b")
    Call allowEdit(4, "����������")
    
    '������ �����
    temp.Protect Secret, UserInterfaceOnly:=True
    lists.Protect Secret, UserInterfaceOnly:=True
    
    '���� ��������� ����� ������� ��������� - ���������������� ���������� ������
    ActiveWorkbook.SaveAs Filename:=Filename
    ActiveWorkbook.Close
er:
End Sub

'��������� ������� ��� �������
Sub setFormat(c As Integer, format As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    If format = "date" Then rang.NumberFormat = "dd.MM.yyyy"
End Sub

'��������� ��������� �������������� ��� �������
Sub setFormatConditions(c As Integer)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    With rang.FormatConditions
        .Add Type:=16
        .Item(.count).Font.Color = vbWhite
    End With
End Sub

'��������� ������ ��������
Sub setValidation(c As Integer, list As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    If list = "b" Then
        formul = "=�����������!$A$2:$A$" + CStr(maxBuyers)
    Else
        formul = "=�����������!$C$2:$C$" + CStr(maxSellers)
    End If
    With rang.Validation
        .Delete
        .Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:=formul
        .ErrorMessage = "������ �� ������, ����������!"
    End With
End Sub

'��������� ���������� �������������� ��� �������
Sub allowEdit(c As Integer, name As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    temp.Protection.AllowEditRanges.Add _
        Title:=name, _
        Range:=rang, _
        Password:=""
    rang.Interior.Color = RGB(255, 255, 192)
End Sub