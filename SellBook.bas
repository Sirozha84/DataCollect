Attribute VB_Name = "SellBook"
'Last change: 20.07.2021 19:25

Dim Patch As String
Dim BuyersList As Variant
Dim SellersList As Variant
Dim Where As Collection
Dim Quartals As Object
Dim BUY As Object
Dim SEL As Object

Sub Run()
    Verify.Init
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Patch = diag.SelectedItems(1)
    Set files = getFiles(Patch, False)
    Range(SBK.Cells(7, 1), SBK.Cells(maxRow, 2)).Clear
    i = 7
    ClearOldBooks Path
    For Each file In files
        SBK.Cells(i, 1).Hyperlinks.Add Anchor:=SBK.Cells(i, 1), _
            Address:="file:" + file, TextToDisplay:=file
        er = ExportBook(file)
        If er = 0 Then SBK.Cells(i, 2) = "������ ��� ������ � ������"
        If er = 1 Then
            If BookCount > 0 Then
                SBK.Cells(i, 2) = "������� ����� ������ (" + CStr(BookCount) + ")"
            Else
                SBK.Cells(i, 2) = "������ ������"
            End If
        End If
        If er = 2 Then SBK.Cells(i, 2) = "������ ����� ������������ ������"
        i = i + 1
    Next
    Message "������!"
    MsgBox "������������ ���� ������ ���������!"
End Sub

'������������ ����� ������
'���������� 0 - ��� ������ ��������, 1 - ����� �� �������, 2 - ������� ��������� ������
Function ExportBook(ByVal file As String) As Byte
    
    Message "������ ����� " + file
    
    '������������� �������
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Patch = FSO.getparentfoldername(file) + "\"
    On Error GoTo er
    Set templ = Workbooks.Open(file, False, False)
    Set TMP = templ.Worksheets(1)
    Set BUY = templ.Worksheets("����������")
    Set SEL = templ.Worksheets("��������")
    
    '�������� �� "������������" ������� �������
    cod = TMP.Cells(1, 1).text
    If cod = "" Or TMP.Cells(2, 1).text <> tmpVersion Then
        ExportBook = 0
        templ.Close
        Exit Function
    End If
    GetLists
    templ.Close
    
    '���������� � �������� ������
    If Not Prepare(cod) Then
        ExportBook = 2
        Exit Function
    End If
    
    '��������� ����
    BookCount = 0
    For Each q In Quartals
        For Each b In BuyersList
            For Each s In SellersList
                MakeBook q, b, s
            Next
        Next
    Next
    
    ExportBook = 1
    Exit Function
er:
    ExportBook = 0
End Function

'����������: �������� ������� �� ���������� ���������, ��������������, ���������� ������� ���������
'���������� True, ���� ������ ���
Function Prepare(ByVal cod As String) As Boolean
    Set Where = New Collection
    Set Quartals = CreateObject("Scripting.Dictionary")
    Prepare = True
    i = firstDat
    Do While DAT.Cells(i, cAccept) <> ""
        If DAT.Cells(i, cCode).text = cod Then
            If DAT.Cells(i, cAccept) = "OK" Then
                Where.Add i
                Quartals(GetQuartal(DAT.Cells(i, cDates))) = 1
            Else
                Prepare = False
                Exit Function
            End If
        End If
        i = i + 1
    Loop
End Function

'������ ������������ ����������� � ��������� �� �������
Sub GetLists()
    
    On Error Resume Next
    
    Set BuyersList = CreateObject("Scripting.Dictionary")
    Set SellersList = CreateObject("Scripting.Dictionary")
    
    i = 2
    For i = 2 To 1000
        If BUY.Cells(i, 1).text <> "" Then _
            BuyersList(BUY.Cells(i, 2).text) = BUY.Cells(i, 1).text
    Next
    
    i = 2
    For i = 2 To 1000
        If SEL.Cells(i, 1).text <> "" Then
            Si = Left(SEL.Cells(i, 2).text, 10)
            ind = selIndexes(Si)
            If Not ind = Empty Then SellersList(Si) = DIC.Cells(selIndexes(Si), 1).text
        End If
    Next

End Sub

'���������� ������ �������� � ������� "1-20"
Function GetQuartal(d As Date) As String
    GetQuartal = CStr((Month(d) - 1) \ 3 + 1) + "-" + Right(CStr(Year(d)), 2)
End Function

'���������� ������� �� ������ ��������
Function Period(q As String)
    y = ".20" + Right(q, 2)
    If Left(q, 1) = "1" Then Period = "� 01.01" + y + " �� 31.03" + y
    If Left(q, 1) = "2" Then Period = "� 01.04" + y + " �� 30.06" + y
    If Left(q, 1) = "3" Then Period = "� 01.07" + y + " �� 30.09" + y
    If Left(q, 1) = "4" Then Period = "� 01.10" + y + " �� 31.12" + y
End Function

'������ ���������� �� ���������� ����
Sub ClearOldBooks(ByVal pat As String)
    On Error GoTo er
    If InStr(1, pat, ".sync") > 0 Then Exit Sub
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set curfold = FSO.GetFolder(pat)
    For Each file In curfold.files
        If file.name Like "������*.xls*" Then Kill file
    Next
    For Each subfolder In curfold.subFolders
         ClearOldBooks subfolder
    Next subfolder
    Exit Sub
er:
    MsgBox "��������� ������ ��� �������� ������ ���� ������. ������������ ���� ��������."
    End
End Sub

'������������ �����
Sub MakeBook(ByVal q As String, ByVal b As String, ByVal s As String)
    
    '����� ������ ��� ������� ���������� �������+����������+��������
    Dim Finded As Collection
    Set Finded = New Collection
    For Each j In Where
        If q = GetQuartal(DAT.Cells(j, 2)) And b = DAT.Cells(j, cBuyINN).text And _
            s = DAT.Cells(j, cSellINN).text Then Finded.Add j
    Next
    If Finded.Count = 0 Then Exit Sub
    
    '�����-�� ������ ���� �����, ������ �����
    name = cutBadSymbols("������ " + SellersList(s) + " (" + s + ") - " + _
                                     BuyersList(b) + " (" + b + ") " + q)
    fileName = Patch + name + ".xlsx"
    Message "������������ ����� " + name
    Workbooks.Add
    Range(Cells(1, 1), Cells(1048576, 24)).Font.name = "Arial"
    Range(Cells(1, 1), Cells(1048576, 24)).Font.Size = 9
    
    '���������
    e = Chr(10)
    Rows(1).RowHeight = 18.8
    bigCell "����� ������", 1, 1, 1, 24
    Cells(1, 1).Font.Size = 14
    Rows(2).RowHeight = 10.9
    Rows(3).RowHeight = 12
    Cells(3, 1) = "�������� " + SellersList(s)
    Rows(4).RowHeight = 12
    Cells(4, 1) = "����������������� ����� � ��� ������� ���������� �� ��e� �����������������-�������� " + _
        DAT.Cells(Finded(1), cSellINN).text
    Rows(5).RowHeight = 12
    Cells(5, 1) = "������� �� ������ " + Period(q)
    Rows(6).RowHeight = 12.8
    Cells(6, 1) = "�����: ���������� = " + DAT.Cells(Finded(1), cBuyer)
    Cells(6, 1).Font.Bold = True
    
    '����� �������
    Rows(7).RowHeight = 90.8
    Rows(8).RowHeight = 40.9
    Rows(9).RowHeight = 10.9
    Range(Cells(7, 1), Cells(9, 24)).Font.Size = 8
    Range(Cells(7, 1), Cells(9, 24)).Font.Bold = True
    Range(Cells(7, 1), Cells(9, 24)).HorizontalAlignment = xlCenter
    Range(Cells(7, 1), Cells(9, 24)).VerticalAlignment = xlCenter
    '1
    Columns(1).ColumnWidth = 6.08
    bigCell "�" + e + "�/�", 7, 1, 2, 1
    Cells(9, 1) = "1"
    '2
    Columns(2).ColumnWidth = 6.75
    bigCell "���" + e + "����" + e + "�����-" + e + "���", 7, 2, 2, 1
    Cells(9, 2) = "2"
    '3
    Columns(3).ColumnWidth = 14.58
    bigCell "����� � ����" + e + "�����-�������" + e + "��������", 7, 3, 2, 1
    Cells(9, 3) = "3"
    '3�
    Columns(4).ColumnWidth = 14.58
    bigCell "����������-" + e + "����� �����" + e + "����������" + e + "����������", 7, 4, 2, 1
    Cells(9, 4) = "3�"
    '3�
    Columns(5).ColumnWidth = 12.25
    bigCell "��� ����" + e + "������", 7, 5, 2, 1
    Cells(9, 5) = "3�"
    '4
    Columns(6).ColumnWidth = 14.58
    bigCell "����� � ����" + e + "�����������" + e + "�����-�������" + e + "��������", 7, 6, 2, 1
    Cells(9, 6) = "4"
    '5
    Columns(7).ColumnWidth = 14.16
    bigCell "����� � ����" + e + "����������-" + e + "�������" + e + "�����-�������" + e + _
        "��������", 7, 7, 2, 1
    Cells(9, 7) = "5"
    '6
    Columns(8).ColumnWidth = 16.92
    bigCell "����� � ����" + e + "�����������" + e + "����������-" + e + "������� �����-" + e + _
        "������� ��������", 7, 8, 2, 1
    Cells(9, 8) = "6"
    '7
    Columns(9).ColumnWidth = 16.5
    bigCell "������������" + e + "����������", 7, 9, 2, 1
    Cells(9, 9) = "7"
    '8
    Columns(10).ColumnWidth = 12.25
    bigCell "���/���" + e + "����������", 7, 10, 2, 1
    Cells(9, 10) = "8"
    '9-10
    Columns(11).ColumnWidth = 15.75
    Columns(12).ColumnWidth = 15.75
    bigCell "�������� � ����������" + e + "(������������, ������)", 7, 11, 1, 2
    bigCell "������������" + e + "����������", 8, 11, 1, 1
    bigCell "���/���" + e + "����������", 8, 12, 1, 1
    Cells(9, 11) = "9"
    Cells(9, 12) = "10"
    '11
    Columns(13).ColumnWidth = 13.08
    bigCell "����� � ����" + e + "���������," + e + "����������-" + e + "�����" + e + "������", 7, 13, 2, 1
    Cells(9, 13) = "11"
    '12
    Columns(14).ColumnWidth = 9.92
    bigCell "�����-" + e + "�������" + e + "� ���" + e + "������", 7, 14, 2, 1
    Cells(9, 14) = "12"
    '13�-�
    Columns(15).ColumnWidth = 15.75
    Columns(16).ColumnWidth = 15.75
    bigCell "��������� ������ �� �����-" + e + "�������, ������� ��������� ��" + e + _
        "����������������� �����-" + e + "������� (������� ���) � ������" + e + "�����-�������", 7, 15, 1, 2
    bigCell "� ������" + e + "�����-�������", 8, 15, 1, 1
    bigCell "� ������ �" + e + "��������", 8, 16, 1, 1
    Cells(9, 15) = "13�"
    Cells(9, 16) = "13�"
    '14-16
    Columns(17).ColumnWidth = 15.75
    Columns(18).ColumnWidth = 15.75
    Columns(19).ColumnWidth = 15.75
    Columns(20).ColumnWidth = 15.75
    bigCell "��������� ������, ���������� �������, �� �����-�������, " + e + _
        "������� ��������� �� ����������������� �����-������� " + e + _
        "(��� ���) � ������ � ��������, �� ������", 7, 17, 1, 4
    bigCell "20 ���������", 8, 17, 1, 1
    bigCell "18 ���������", 8, 18, 1, 1
    bigCell "10 ���������", 8, 19, 1, 1
    bigCell "0 ���������", 8, 20, 1, 1
    Cells(9, 17) = "14"
    Cells(9, 18) = "14�"
    Cells(9, 19) = "15"
    Cells(9, 20) = "16"
    '17-18
    Columns(21).ColumnWidth = 15.75
    Columns(22).ColumnWidth = 15.75
    Columns(23).ColumnWidth = 15.75
    bigCell "����� ��� �� �����-�������," + e + "������� ����� ������ �� �����������������" + e + _
        "�����-������� � ������ � ��������, �� ������", 7, 21, 1, 3
    bigCell "20 ���������", 8, 21, 1, 1
    bigCell "18 ���������", 8, 22, 1, 1
    bigCell "10 ���������", 8, 23, 1, 1
    Cells(9, 21) = "17"
    Cells(9, 22) = "17�"
    Cells(9, 23) = "18"
    '19
    Columns(24).ColumnWidth = 15.75
    bigCell "���������" + e + "������," + e + "�������������" + e + "�� ������, ��" + e + _
        "�����-�������," + e + "�������" + e + "���������" + e + "�� ����������-" + e + _
        "�������" + e + "�����-�������" + e + "� ������ �" + e + "��������", 7, 24, 2, 1
    Cells(9, 24) = "19"
    
    '������
    i = 10
    n = 1
    s1 = 0: s2 = 0: s3 = 0: s4 = 0: s5 = 0: s6 = 0
    For Each j In Finded
        Rows(i).RowHeight = 24
        Rows(i).VerticalAlignment = xlTop
        Cells(i, 1) = n
        Cells(i, 2).NumberFormat = "@"
        Cells(i, 2) = "01"
        Cells(i, 3) = DAT.Cells(j, 1).text + " ��" + e + DAT.Cells(j, cDates).text
        Cells(i, 9) = DAT.Cells(j, cBuyer)
        Cells(i, 9).WrapText = True
        Cells(i, 10) = DAT.Cells(j, cBuyINN)
        Cells(i, 10).WrapText = True
        Cells(i, 16) = DAT.Cells(j, cPrice)
        Cells(i, 17) = DAT.Cells(j, 9): If Cells(i, 17) <> "" Then s1 = s1 + Cells(i, 17)
        Cells(i, 18) = DAT.Cells(j, 10): If Cells(i, 18) <> "" Then s2 = s2 + Cells(i, 18)
        Cells(i, 19) = DAT.Cells(j, 11): If Cells(i, 19) <> "" Then s3 = s3 + Cells(i, 19)
        Cells(i, 21) = DAT.Cells(j, 12): If Cells(i, 21) <> "" Then s4 = s4 + Cells(i, 21)
        Cells(i, 22) = DAT.Cells(j, 13): If Cells(i, 22) <> "" Then s5 = s5 + Cells(i, 22)
        Cells(i, 23) = DAT.Cells(j, 14): If Cells(i, 23) <> "" Then s6 = s6 + Cells(i, 23)
        Range(Cells(i, 9), Cells(i, 10)).WrapText = True
        Range(Cells(i, 15), Cells(i, 23)).NumberFormat = numFormat
        i = i + 1
        n = n + 1
    Next
    
    '������
    Rows(i).RowHeight = 12.8
    Cells(i, 1) = "�����"
    Range(Cells(i, 1), Cells(i, 16)).merge
    Cells(i, 1).HorizontalAlignment = xlRight
    Range(Cells(i, 1), Cells(i, 24)).Font.Bold = True
    If s1 > 0 Then Cells(i, 17) = s1
    If s2 > 0 Then Cells(i, 18) = s2
    If s3 > 0 Then Cells(i, 19) = s3
    If s4 > 0 Then Cells(i, 21) = s4
    If s5 > 0 Then Cells(i, 22) = s5
    If s6 > 0 Then Cells(i, 23) = s6
    Range(Cells(i, 15), Cells(i, 23)).NumberFormat = numFormat
    Range(Cells(7, 1), Cells(i, 24)).Borders.Weight = 2
    
    '���������� � �������� ���������
    On Error GoTo er
    Application.DisplayAlerts = False
    With Sheets(1).PageSetup
        .Orientation = xlLandscape
        .LeftMargin = 0.64
        .TopMargin = 0.64
        .RightMargin = 0.64
        .BottomMargin = 0.64
        .FitToPagesWide = 1
        .Zoom = False
    End With
    ActiveWorkbook.SaveAs fileName:=fileName
    ActiveWorkbook.Close
    BookCount = BookCount + 1
    Exit Sub
er:
    ActiveWorkbook.Close
    MsgBox "��������� ������ ��� ���������� ����� " + fileName
End Sub

Sub bigCell(txt As String, r As Variant, c As Variant, height As Variant, width As Variant)
    Cells(r, c) = txt
    Range(Cells(r, c), Cells(r + height - 1, c + width - 1)).merge
    Cells(r, c).HorizontalAlignment = xlCenter
    Cells(r, c).VerticalAlignment = xlCenter
    Cells(r, c).Font.Bold = True
End Sub

'******************** End of File ********************