VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormExport 
   Caption         =   "�������� ������"
   ClientHeight    =   2520
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4557
   OleObjectBlob   =   "FormExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstDate As Date
Dim LastDate As Date

'������������� ���������� �������
Private Sub UserForm_Initialize()
    
    Verify.Init
    
    '���������� ������ ���������
    ComboBoxBuyers.AddItem "���"
    For Each seller In selIndexes
        ComboBoxBuyers.AddItem SellFileName(seller)
    Next
    ComboBoxBuyers.ListIndex = 0
        
    '������ �����
    TextBoxFirstCollect = PRP.Cells(8, 2)
    TextBoxLastCollect = PRP.Cells(9, 2)
    
End Sub

'��� ����� �� ��� ��������
Function SellFileName(INN) As String
    ind = selIndexes(INN)
    If ind <> Empty Then SellFileName = INN + "-" + DIC.Cells(ind, 1)
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

Private Sub CommandExit_Click()
    End
End Sub

'������ "�������"
Private Sub CommandExport_Click()
    
    On Error GoTo er
    FirstDate = CDate(TextBoxFirstCollect)
    LastDate = CDate(TextBoxLastCollect)
    On Error GoTo 0
    
    If ComboBoxBuyers.ListIndex = 0 Then
        n = 1
        a = selIndexes.Count
        For Each seller In selIndexes
            ExportFile seller, CStr(n) + " �� " + CStr(a) + ": "
            n = n + 1
        Next
    Else
        ExportFile Left(ComboBoxBuyers.Value, 10), ""
    End If
    
    '���������� ��� ������� �����
    PRP.Cells(8, 2) = TextBoxFirstCollect
    PRP.Cells(9, 2) = TextBoxLastCollect
    
    Message "������!"
    End

er:
    MsgBox "���� �� ������� ��� ������� �� ���������"

End Sub

'������� �����
Private Sub ExportFile(ByVal INN As String, NUM As String)

    seller = SellFileName(INN)
    Message "������� ����� " + NUM + seller
    
    '�������� ������������ ������ � �����������
    si = selIndexes(INN)
    limit = DIC.Cells(si, cLimND)
    ermsg = "� �������� " + DIC.Cells(si, 1) + " � ��� " + INN + " "
    If limit = Empty Then
        MsgBox ermsg + "�� ������ �����!"
        End
    End If
    oND = StupidQToQIndex(DIC.Cells(si, cOPND))
    If oND < 0 Then
        MsgBox ermsg + "�� ������ ��� ������ �� ��������� �������� ������ ��!"
        End
    End If
    
    '������������ � ���� � ������ �����
    Patch = DirExport + "\��������"
    MakeDir Patch
    fol = ""
    If fol <> "" Then MakeDir (Patch + fol)
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
            dc = DAT.Cells(i, cDateCol)
            If dc >= FirstDate And dc < LastDate + 1 Then
                cp = True
                If DAT.Cells(i, cSellINN).text <> INN Then cp = False
                d = DAT.Cells(i, cDates)
                If cp Then
                    '����������� ������ �� �����
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
                    '��������� ������� - ������ �������� � ����� ���
                    'Cells(j, 15) = YearAndQuartal(DAT.Cells(i, 2))
                    Cells(j, 15) = DateToQIndex(DAT.Cells(i, 2))
                    Sum = 0
                    For j2 = 11 To 13
                        If IsNumeric(Cells(j, j2)) Then Sum = Sum + Cells(j, j2)
                    Next
                    Cells(j, 16) = Sum
                    j = j + 1
                End If
            End If
        End If
        i = i + 1
    Loop
    
    '���������� �� ��������
    Cells(1, 15) = "�������"
    Cells(1, 16) = "���"
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("O2") '������ ������� ����������
        .SortFields.Add Key:=Range("F2") '������ ������� ����������
        .setRange Range("A2:P" + CStr(j - 1)) '�������� ����������� �������
        .Apply
    End With
    
    PeriodND limit, oND
    
    '�������� ��������� ��������
    Columns(15).Delete
    Columns(15).Delete
    
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

'������ �������� ��������� ����������
'oND - �������� ������ ��
Sub PeriodND(ByVal limit As Double, ByVal oND)
    
    
    
    End
    
End Sub