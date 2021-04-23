Attribute VB_Name = "Template"
'Last change: 23.04.2021 18:20

Const LastRec = 10000   '��������� ������ ������� (������ ������ 5, ����� ��������)
Const maxComps = 100    '������������ ���������� �������� (��������� ��� �����������)

Sub Generate()
    
    Main.Init
    If IsNumeric(PRP.Cells(7, 2)) Then last = PRP.Cells(7, 2)
    Dim max As Long
    i = firstTempl
    Do While Cells(i, cTClient) <> "" Or Cells(i, cTForm) <> ""
        i = i + 1
    Loop
    '���������� �������
    Set namelist = CreateObject("Scripting.Dictionary")
    max = i - 1
    fold = DirImportSale
    For i = firstTempl To max
        Message "�������� ������� " + CStr(i - firstTempl + 1) + " �� " + CStr(max - firstTempl + 1)
        cln = cutBadSymbols(Cells(i, cTClient).text)
        brk = cutBadSymbols(Cells(i, cTBroker).text)
        tem = cutBadSymbols(Cells(i, cTForm).text)
        '��������, ���������� �� �����
        uname = cln + "!" + tem
        If namelist(uname) = "" Then
            namelist(uname) = 0
            If Not IsCode(Cells(i, cTCode)) Then
                cod = last + 1
                last = cod
                Cells(i, cTCode) = cod
            End If
            If Cells(i, cTStat).text <> "OK" Then
                '������ ����� � ����
                If brk <> "" Then brk = "\" + brk
                MakeDir fold + "\" + cln
                MakeDir fold + "\" + cln + brk
                MakeDir fold + "\" + cln + brk + "\" + tem
                name = fold + "\" + cln + brk + "\" + tem + "\" + tem + ".xlsx"
                res = NewTemplate(cln, tem, name, Cells(i, cTCode).text)
                If res = 0 Then
                    Cells(i, cTFile) = "��������� ������ ��� �������� �����"
                    Cells(i, cTResult) = "������"
                End If
                If res = 1 Then
                    Cells(i, cTFile) = name
                    Cells(i, cTResult) = "�������!"
                    Cells(i, cTStat) = "OK"
                End If
                If res = 2 Then
                    Cells(i, cTFile) = name
                    Cells(i, cTResult) = "���� ��� ����������, ���������"
                    Cells(i, cTStat) = "OK"
                End If
            Else
                Cells(i, cTResult) = "������ ��� ������ �����"
            End If
        Else
            Cells(i, cTResult) = "��� ������� ��� ������� �� ���������."
        End If
    Next
    PRP.Cells(7, 2) = last
    
    ActiveWorkbook.Save
    Message "������! ���� �������."
    
End Sub

'��������, ������ �� ������ �� ���
Function IsCode(n As Variant)
    IsCode = False
    If IsNumeric(n) Then
        If n > 0 Then IsCode = True
    End If
End Function

'�������� ������ �����
'���������� 0 - ���� �� ������, 1 - ���� ������, 2 - ���� ��� ����, ���������
Function NewTemplate(ByVal cln As String, ByVal tem As String, _
    ByVal fileName As String, ByVal cod As String) As Byte
    
    '���� ���� ���������� - ���������
    If Dir$(fileName) <> "" Then NewTemplate = 2: Exit Function
    
    '������ ���� � ������� ���������
    Workbooks.Add
    Sheets.Add
    Sheets.Add
    Sheets(1).name = cln
    Sheets(2).name = "����������"
    Sheets(3).name = "��������"
    On Error Resume Next '�� ������ ���� ���������� ���� 1 �������, (� 2010 ��������� �� ��������� 3)
    Application.DisplayAlerts = False
    Sheets(4).Delete
    Sheets(4).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set temp = Application.ActiveSheet
    Cells(1, 1) = cod
    Cells(2, 1) = tmpVersion
    Range(Cells(1, 1), Cells(2, 1)).Font.Color = vbWhite
    Cells(1, 2) = "������: " + cln
    Cells(2, 2) = "������: " + tem
    
    '�������� �������. ������ ����� �����
    Columns(1).ColumnWidth = 20
    Columns(2).ColumnWidth = 15
    Columns(3).ColumnWidth = 22
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
    Range(Cells(3, 1), Cells(3, 2)).merge
    Cells(3, 3) = "�������� � ����������"
    Range(Cells(3, 3), Cells(3, 4)).merge
    Cells(3, 5) = "�������� � ��������"
    Range(Cells(3, 5), Cells(3, 6)).merge
    Cells(3, 7) = "���������" + Chr(10) + "������ � ���"
    Cells(3, 8) = "������" + Chr(10) + "���, %"
    Range(Cells(3, 8), Cells(4, 8)).merge
    Cells(3, 9) = "��������� ������ ���������� �������" + Chr(10) + "(� ���.) ��� ���"
    Range(Cells(3, 9), Cells(3, 11)).merge
    Cells(3, 12) = "����� ���"
    Range(Cells(3, 12), Cells(3, 14)).merge
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
    hat.Interior.Color = colGray
    hat.Borders.Weight = 3
    
    '���� 2 - ����
    SetFormat 2, "date"
    SetValidation 2, "date"
    AllowEdit 2, "����"
    
    '���� 3 - ��� ����������, ��������� � ������� ���
    SetRange(3).FormulaLocal = "=���(D5;����������!A$2:B$" + CStr(maxComps) + ";2;0)"
    SetFormatConditions 3
    
    '���� 4 - ����������, �������� �� ������
    SetValidation 4, "buy"
    AllowEdit 4, "����������"
    
    '���� 5 - ��� ��������, ��������� � ������� ���
    SetRange(5).FormulaLocal = "=���(F5;��������!A$2:B$" + CStr(maxComps) + ";2;0)"
    SetFormatConditions 5
    
    '���� 6 - ��������, �������� �� ������
    SetValidation 6, "sale"
    AllowEdit 6, "��������"
    
    '���� 7 - ���������
    SetValidation 7, "num"
    SetFormat 7, "money"
    AllowEdit 7, "���������"
    Cells(1, 7).Borders.Weight = 3
    Cells(1, 7).FormulaLocal = "=����(G5:G" + CStr(LastRec) + ")"
    
    '���� 8 - ������ ���
    SetValidation 8, "nds"
    AllowEdit 8, "������ ���"
    
    '����� 9-14
    For i = 9 To 14
        SetFormat i, "money"
        Cells(1, i).Borders.Weight = 3
    Next
    
    '���� 9-11 - ����� � ��� 20,18,10%      ������� G/(100+H)*100
    SetRange(9).FormulaLocal = "=����(�(G5<>"""";H5=20);������(G5-L5;2);"""")"
    SetRange(10).FormulaLocal = "=����(�(G5<>"""";H5=18);������(G5-M5;2);"""")"
    SetRange(11).FormulaLocal = "=����(�(G5<>"""";H5=10);������(G5-N5;2);"""")"
    Cells(1, 9).FormulaLocal = "=����(I5:I" + CStr(LastRec) + ")"
    Cells(1, 10).FormulaLocal = "=����(J5:J" + CStr(LastRec) + ")"
    Cells(1, 11).FormulaLocal = "=����(K5:K" + CStr(LastRec) + ")"
    
    '���� 12-14 - ����� ��� ��� 20,18,10%   ������� G/(100+H)*H
    SetRange(12).FormulaLocal = "=����(�(G5<>"""";H5=20);������(G5/(100+H5)*H5;2);"""")"
    SetRange(13).FormulaLocal = "=����(�(G5<>"""";H5=18);������(G5/(100+H5)*H5;2);"""")"
    SetRange(14).FormulaLocal = "=����(�(G5<>"""";H5=10);������(G5/(100+H5)*H5;2);"""")"
    Cells(1, 12).FormulaLocal = "=����(L5:L" + CStr(LastRec) + ")"
    Cells(1, 13).FormulaLocal = "=����(M5:M" + CStr(LastRec) + ")"
    Cells(1, 14).FormulaLocal = "=����(N5:N" + CStr(LastRec) + ")"
        
    '����������
    Range(Cells(4, 1), Cells(4, 14)).Rows.AutoFilter
    
    '����������� �������
    Range("A5").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    
    '������� �� �������������
    For i = 2 To 3
        Sheets(i).Activate
        Columns(1).ColumnWidth = 30
        Columns(2).ColumnWidth = 20
        Cells(1, 1) = "������������"
        Cells(1, 2) = "���/���"
        Range(Cells(2, 2), Cells(maxComps, 2)).NumberFormat = "@"
        Rows(2).Hidden = True
        With Range(Cells(3, 2), Cells(maxComps, 2)).Validation
            .Delete
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:= _
                "=OR(LEN(B3)=12,LEN(B3)=20)"
            .ErrorMessage = "�� ���������� ����� ������. ������ ���� 12 ��� 20 ��������."
        End With
        Range("A3").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
    Next
    
    '������ � ���������� �����
    Sheets(1).Activate
    SetProtect ActiveSheet
    On Error GoTo er
    ActiveWorkbook.SaveAs fileName:=fileName    '��� ������ ��� ������ ������������ � �������
    ActiveWorkbook.Close                        '��������� ����� (������ ������ ������ �� ������ �����)
    NewTemplate = 1
    Exit Function
er:
    ActiveWorkbook.Close
    NewTemplate = 0
End Function

Function SetRange(ByVal c As Integer) As Range
    Set SetRange = Range(Cells(5, c), Cells(LastRec, c))
End Function

'��������� ������� ��� �������
Sub SetFormat(ByVal c As Integer, format As String)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    If format = "date" Then rang.NumberFormat = "dd.MM.yyyy"
    If format = "money" Then rang.NumberFormat = "### ### ##0.00"
End Sub

'��������� ��������� �������������� ��� �������
Sub SetFormatConditions(c As Integer)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    With rang.FormatConditions
        .Add Type:=xlErrorsCondition
        .Item(.Count).Font.Color = vbWhite
    End With
End Sub

'��������� �������� ��������
Sub SetValidation(c As Integer, typ As String)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    If typ = "buy" Then
        formul = "=����������!$A$2:$A$" + CStr(maxComps)
        With rang.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formul
            .ErrorMessage = "������ �� ������, ����������!"
        End With
    End If
    If typ = "sale" Then
        formul = "=��������!$A$2:$A$" + CStr(maxComps)
        With rang.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formul
            .ErrorMessage = "������ �� ������, ����������!"
        End With
    End If
    If typ = "date" Then
        formul = "=OR(AND(H5=10),AND(H5=18,B5<43466),AND(H5=20,B5>=43466))"
        With rang.Validation
            .Delete
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:=formul
            .ErrorMessage = "�� 01.01.2019 ��� ��� 18%, ����� - 20%, ��� 10% � ����� �����"
        End With
    End If
    If typ = "num" Then
        With rang.Validation
            .Delete
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="0"
            .ErrorMessage = "����� ������ ���� ������ 0"
        End With
    End If
    If typ = "nds" Then
        formul = "=OR(AND(H5=10),AND(H5=18,B5<43466),AND(H5=20,B5>=43466))"
        With rang.Validation
            .Delete
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:=formul
            .ErrorMessage = "�� 01.01.2019 ��� ��� 18%, ����� - 20%, ��� 10% � ����� �����"
        End With
    End If
End Sub

'��������� ���������� �������������� ��� �������
Sub AllowEdit(c As Integer, name As String)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    ActiveSheet.Protection.AllowEditRanges.Add Title:=name, Range:=rang, Password:=""
    rang.Interior.Color = RGB(255, 255, 192)
End Sub

'******************** End of File ********************