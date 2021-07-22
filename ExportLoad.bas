Attribute VB_Name = "ExportLoad"
'Last change: 22.07.2021 14:41

Dim datesD As Collection    '��������� ���, ��������� � ������������
Dim datesI As Collection    '��������� �������� ���� ���

Sub Run()
    
    Message "����������..."
    Dictionary.Init
    
    '���������� �������� ��� ��������
    Patch = DirExport + "\�����������"
    MakeDir Patch
    Set files = Source.getFiles(DirExport + "\�����������", False)
    For Each file In files
        Kill file
    Next
    
    '��������� ������ ��� ��������� ��� ��������
    Set files = Source.getFiles(DirExport + "\��������", False)
    Set INNs = FilesToINNs(files)
    
    n = 1
    a = INNs.Count
    For Each INN In INNs
        Message "������������� �����������... (" + CStr(n) + " �� " + CStr(a) + ")"
        LoadAllocation INN
        n = n + 1
    Next
    
    n = 1
    For Each INN In INNs
        Message "������� ������... (" + CStr(n) + " �� " + CStr(a) + ")"
        CreateExportFile INN
        n = n + 1
    Next
    
    Message "������!"
    
End Sub

'������������ ����� ��������
Sub CreateExportFile(ByVal INN As String)
    
    '������������ � ���� � ������ �����
    Patch = DirExport + "\�����������"
    MakeDir Patch
    fileName = Patch + "\" + cutBadSymbols(SellFileName(INN)) + ".xlsx"
    
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
    Cells(i, 8) = "����� ���"
    Cells(i, 9) = "������ ��"
    Cells(i, 10) = "���� ��������" + Chr(10) + "� ����� ��"
    Columns(1).ColumnWidth = 10
    Columns(2).ColumnWidth = 13
    Columns(3).ColumnWidth = 10
    Columns(4).ColumnWidth = 11
    Columns(5).ColumnWidth = 10
    Columns(6).ColumnWidth = 15
    Columns(7).ColumnWidth = 12
    Columns(8).ColumnWidth = 12
    Columns(9).ColumnWidth = 12
    Columns(10).ColumnWidth = 15
    Rows(1).RowHeight = 30
    Set hat = Range(Cells(1, 1), Cells(1, 10))
    hat.HorizontalAlignment = xlCenter
    hat.VerticalAlignment = xlCenter
    hat.Interior.Color = colGray
    hat.Borders.Weight = 3
    firstEx = i + 1
    
    '��������� �����
    i = firstDat
    j = firstEx
    Do While DTL.Cells(i, clAccept) <> ""
        If DTL.Cells(i, clAccept) = "OK" Then
            If Left(DTL.Cells(i, clSaleINN).text, 10) = INN Then
                '����������� ������ �� �����
                Cells(j, 1).NumberFormat = "@"
                Cells(j, 1) = "01"
                Cells(j, 2) = DTL.Cells(i, 1)
                Cells(j, 3).NumberFormat = "dd.MM.yyyy"
                Cells(j, 3) = DTL.Cells(i, clDate)
                innkpp = Split(dat.Cells(i, 3), "/")
                Cells(j, 4).NumberFormat = "@"
                Cells(j, 4) = DTL.Cells(i, clSaleINN)
                Cells(j, 6) = DTL.Cells(i, clSaleName)
                Cells(j, 7).NumberFormat = numFormat
                Cells(j, 7) = DTL.Cells(i, clPrice)
                Cells(j, 8).NumberFormat = numFormat
                Cells(j, 8) = DTL.Cells(i, clNDS)
                Cells(j, 9) = DTL.Cells(i, clPND).text
                Cells(j, 10) = LastDateOfQuartal(DTL.Cells(i, clPND).text)
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

'������������� �����������
Sub LoadAllocation(ByVal INN As String)
    
    oND = StupidQToQIndex(DIC.Cells(selIndexes(INN), cOPND))
    For cPer = oND To quartCount - 1
        
        '������ ����� ���� �������� �� ������� ������
        Sum = GetSaleSumm(INN, cPer)
        If Sum > minSale Then
        
            '��������� ������ �����������
            GetDatesList cPer, INN
            For d = 1 To datesD.Count
                i = datesI(d)
                Post = DTL.Cells(i, clNDS)  '�����������
                If Sum >= Post Then
                    If DTL.Cells(i, clRasp) = "" Then
                        Sum = Sum - Post
                        DTL.Cells(i, clRasp) = Post
                        DTL.Cells(i, clPND) = IndexToQYYYY(cPer)
                    Else
                        Sum = Sum - OneCellSum(DTL.Cells(i, clRasp))
                    End If
                    If Sum < maxDif Then Exit For
                End If
            Next
            
        End If
        
    Next
    
End Sub

'������� ������ ������ � ������ ���
Function FilesToINNs(files As Object) As Object
    Set FilesToINNs = New Collection
    For Each file In files
        FilesToINNs.Add Left(Source.FSO.GetFileName(file), 10)
    Next
End Function

'������ ����� �������� �������� � INN �� ������� Q
Function GetSaleSumm(ByVal INN As String, ByVal q As Integer) As Double
    i = firstDat
    Sum = 0
    Do While dat.Cells(i, cAccept) <> ""
        If dat.Cells(i, cAccept) = "OK" And Left(dat.Cells(i, cSellINN).text, 10) = INN Then
            If StupidQToQIndex(dat.Cells(i, cPND)) = q Then
                Sum = Sum + WorksheetFunction.Sum(Range(dat.Cells(i, cNDS), dat.Cells(i, cNDS + 2)))
            End If
        End If
        i = i + 1
    Loop
    GetSaleSumm = Sum
End Function

'������������ �������������� ������ ��� �������� � 12 ��������� ������� �� cPer
Sub GetDatesList(ByVal cPer As Integer, ByVal INN As String)
            
    Dim datesDt As Collection
    Dim datesIt As Collection
    Set datesDt = New Collection
    Set datesIt = New Collection
    Set datesD = New Collection
    Set datesI = New Collection
        
    '�������� �������, ��������������� ������� ��������
    Dim d As Date
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        'Debug.Assert i <> 13822
        If DTL.Cells(i, clAccept) = "OK" And DTL.Cells(i, clSaleINN) = INN Then
            d = DTL.Cells(i, clDate)
            q = DateToQIndex(d)
            If q >= cPer And q <= cPer + 11 Then
                datesDt.Add d
                datesIt.Add i
            End If
        End If
        i = i + 1
    Loop
    
    '��������� ��������� ����
    Do While datesDt.Count > 0
        max = datesDt(1)
        im = 1
        For i = 1 To datesDt.Count
            If max < datesDt(i) Then
                max = datesDt(i)
                im = i
            End If
        Next
        datesD.Add datesDt(im)
        datesI.Add datesIt(im)
        datesDt.Remove (im)
        datesIt.Remove (im)
    Loop
    
End Sub

'���������� ��������� ���� ��������
Function LastDateOfQuartal(ByVal q) As String
    n = Left(q, 1)
    y = Right(q, 4)
    If n = 1 Then LastDateOfQuartal = "31.03." + y
    If n = 2 Then LastDateOfQuartal = "30.06." + y
    If n = 3 Then LastDateOfQuartal = "30.09." + y
    If n = 4 Then LastDateOfQuartal = "31.12." + y
End Function

'******************** End of File ********************