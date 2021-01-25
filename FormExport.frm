VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormExport 
   Caption         =   "�������� ������"
   ClientHeight    =   3241
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4564
   OleObjectBlob   =   "FormExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim Sellers As Variant

'������������� ���������� �������
Private Sub UserForm_Initialize()
    
    Message "���������� � ��������..."
    Set Sellers = CreateObject("Scripting.Dictionary")
    Set Quartals = CreateObject("Scripting.Dictionary")
    Set Months = CreateObject("Scripting.Dictionary")
    
    Verify.Init
    i = firstDat
    Do While Cells(i, cAccept) <> ""
        If Cells(i, cAccept).text = "OK" Then
            Sellers(SellFileName(Cells(i, cSellINN).text)) = 1
            Months(YearAndMonth(Cells(i, cDates).text)) = 1
            Quartals(YearAndQuartal(Cells(i, cDates).text)) = 1
        End If
        i = i + 1
    Loop
    
    If Months.Count = 0 Or Quartals.Count = 0 Then
        MsgBox "��� ������ ��� ��������"
        End
    End If
    
    ComboBoxBuyers.AddItem "���"
    For Each seller In Sellers
        ComboBoxBuyers.AddItem seller
    Next
    ComboBoxBuyers.ListIndex = 0
    
    For Each m In Months
        ComboBoxMonths.AddItem m
    Next
    ComboBoxMonths.ListIndex = 0
    
    For Each q In Quartals
        ComboBoxQuartals.AddItem q
    Next
    ComboBoxQuartals.ListIndex = 0
    
    Message "������!"
    
End Sub

'��� ����� �� ��� ��������
Function SellFileName(inn) As String
    ind = selIndexes(inn)
    If ind <> Empty Then SellFileName = inn + "-" + DIC.Cells(ind, 1)
End Function

Function YearAndMonth(ByVal d As String) As String
    On Error GoTo er:
    YearAndMonth = CStr(Year(d)) + " - "
    dy = Month(d)
    If dy = 1 Then YearAndMonth = YearAndMonth + "������"
    If dy = 2 Then YearAndMonth = YearAndMonth + "�������"
    If dy = 3 Then YearAndMonth = YearAndMonth + "����"
    If dy = 4 Then YearAndMonth = YearAndMonth + "������"
    If dy = 5 Then YearAndMonth = YearAndMonth + "���"
    If dy = 6 Then YearAndMonth = YearAndMonth + "����"
    If dy = 7 Then YearAndMonth = YearAndMonth + "����"
    If dy = 8 Then YearAndMonth = YearAndMonth + "������"
    If dy = 9 Then YearAndMonth = YearAndMonth + "��������"
    If dy = 10 Then YearAndMonth = YearAndMonth + "�������"
    If dy = 11 Then YearAndMonth = YearAndMonth + "������"
    If dy = 12 Then YearAndMonth = YearAndMonth + "�������"
    Exit Function
er:
    YearAndMonth = ""
End Function

Function YearAndQuartal(ByVal d As String) As String
    On Error GoTo er
    YearAndQuartal = CStr(Year(d)) + " - " + CStr((Month(d) - 1) \ 3 + 1) + " �������"
    Exit Function
er:
    YearAndQuartal = ""
End Function

Private Sub OptionAll_Click()
    ComboBoxMonths.Enabled = False
    ComboBoxQuartals.Enabled = False
End Sub

Private Sub OptionMonth_Click()
    ComboBoxMonths.Enabled = True
    ComboBoxQuartals.Enabled = False
End Sub

Private Sub OptionQuartal_Click()
    ComboBoxMonths.Enabled = False
    ComboBoxQuartals.Enabled = True
End Sub

Private Sub CommandExit_Click()
    End
End Sub

Private Sub CommandExport_Click()
    If ComboBoxBuyers.ListIndex = 0 Then
        n = 1
        a = Sellers.Count
        For Each seller In Sellers
            ExportFile seller, CStr(n) + " �� " + CStr(a) + ": "
            n = n + 1
        Next
    Else
        ExportFile ComboBoxBuyers.Value, ""
    End If
    Message "������!"
    End
End Sub

'������� �����
Private Sub ExportFile(ByVal seller As String, NUM As String)
    
    Message "������� ����� " + NUM + seller
    
    '������������ � ���� � ������ �����
    Patch = DAT.Cells(2, 3)
    fol = ""
    mnC = OptionMonth.Value
    mn = ComboBoxMonths.Value
    qrC = OptionQuartal.Value
    qr = ComboBoxQuartals.Value
    If mnC Then fol = "\" + mn
    If qrC Then fol = "\" + qr
    If fol <> "" Then folder (Patch + fol)
    fileName = Patch + fol + "\" + cutBadSymbols(seller) + ".xlsx"
    
    '������ �����
    Workbooks.Add
    i = 1
    Cells(i, 1) = "��� ����" + Chr(10) + "��������"
    Cells(i, 2) = "� ����" + Chr(10) + "�������"
    Cells(i, 3) = "���� ����" + Chr(10) + "�������"
    Cells(i, 4) = "���"
    Cells(i, 5) = "���"
    Cells(i, 6) = "������������"
    Cells(i, 7) = "����� � ���." + Chr(10) + "� ���."
    Cells(i, 8) = "�����" + Chr(10) + "��� ��� 20%"
    Cells(i, 9) = "�����" + Chr(10) + "��� ��� 18%"
    Cells(i, 10) = "�����" + Chr(10) + "��� ��� 10%"
    Cells(i, 11) = "��� 20%"
    Cells(i, 12) = "��� 18%"
    Cells(i, 13) = "��� 10%"
    Cells(i, 14) = "������ ��"
    Columns(1).ColumnWidth = 10
    Columns(2).ColumnWidth = 13
    Columns(3).ColumnWidth = 10
    Columns(4).ColumnWidth = 11
    Columns(5).ColumnWidth = 10
    Columns(6).ColumnWidth = 15
    Columns(7).ColumnWidth = 12
    Columns(8).ColumnWidth = 12
    Columns(9).ColumnWidth = 12
    Columns(10).ColumnWidth = 12
    Columns(11).ColumnWidth = 10
    Columns(12).ColumnWidth = 10
    Columns(13).ColumnWidth = 10
    Columns(14).ColumnWidth = 10
    Rows(1).RowHeight = 30
    Set hat = Range(Cells(1, 1), Cells(1, 14))
    hat.HorizontalAlignment = xlCenter
    hat.VerticalAlignment = xlCenter
    hat.Interior.Color = colGray
    hat.Borders.Weight = 3
    firstEx = i + 1
    
    
    '��������� �����
    i = firstDat
    j = firstEx
    Do While DAT.Cells(i, cAccept) <> ""
        If DAT.Cells(i, cAccept) = "OK" Then
            cp = True
            If SellFileName(DAT.Cells(i, cSellINN).text) <> seller Then cp = False
            d = DAT.Cells(i, cDates)
            If mnC Then If YearAndMonth(d) <> mn Then cp = False
            If qrC Then If YearAndQuartal(d) <> qr Then cp = False
            If cp Then
                Cells(j, 1).NumberFormat = "@"
                Cells(j, 1) = "01"
                Cells(j, 2) = DAT.Cells(i, 1)
                Cells(j, 3).NumberFormat = "dd.MM.yyyy"
                Cells(j, 3) = DAT.Cells(i, 2)
                innkpp = Split(DAT.Cells(i, 3), "/")
                Cells(j, 4).NumberFormat = "@"
                Cells(j, 4) = innkpp(0)
                Cells(j, 5).NumberFormat = "@"
                If UBound(innkpp) > 0 Then Cells(j, 5) = innkpp(1)
                Cells(j, 6) = DAT.Cells(i, 4)
                Cells(j, 7).NumberFormat = "### ### ##0.00"
                Cells(j, 7) = DAT.Cells(i, 7)
                For c = 0 To 5
                    Cells(j, 8 + c).NumberFormat = "### ### ##0.00"
                    Cells(j, 8 + c) = DAT.Cells(i, 9 + c)
                Next
                j = j + 1
            End If
        End If
        i = i + 1
    Loop
    
    '���������� � �������� ���������
    On Error GoTo er
    Application.DisplayAlerts = False
    If j > firstEx Then ActiveWorkbook.SaveAs fileName:=fileName
    ActiveWorkbook.Close
    Exit Sub
er:
    ActiveWorkbook.Close
    MsgBox "��������� ������ ��� ���������� ����� " + fileName
End Sub