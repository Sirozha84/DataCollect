Attribute VB_Name = "ExportLoad"
'Last change: 04.04.2021 18:47

Sub Run()
    
    Message "����������..."
    Dictionary.Init
    LoadAllocation
    
    '���������� �������� ��� ��������
    Patch = DirExport + "\�����������"
    MakeDir Patch
    Set files = Source.getFiles(DirExport + "\�����������", False)
    For Each file In files
        Kill file
    Next
    
    '��������� ������ ��� ��������� ��� ��������
    Dim SalersINN As Collection
    Set SalersINN = New Collection
    Set files = Source.getFiles(DirExport + "\��������", False)
    
    n = 1
    a = selIndexes.Count
    For Each file In files
        CreateExportFile Left(Source.FSO.GetFileName(file), 10), CStr(n) + " �� " + CStr(a) + ": "
    Next
    
    Message "������!"
    
End Sub

'������������ ����� ��������
Sub CreateExportFile(ByVal INN As String, ByVal NUM As String)
    
    saler = SellFileName(INN)
    Message "������� ����� " + NUM + saler
    
    '������������ � ���� � ������ �����
    Patch = DirExport + "\�����������"
    MakeDir Patch
    fileName = Patch + "\" + cutBadSymbols(saler) + ".xlsx"
    
    '������ �����
    Workbooks.Add
    
    
    
    '��� ����� ������������ ����� ������
    
    
    
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
Sub LoadAllocation()
    
    For Each INN In selIndexes
        
        Si = selIndexes(INN)
        oND = StupidQToQIndex(DIC.Cells(Si, cOPND))
        
        For cPer = oND To quartCount - 1
            
            '������ ����� ���� �������� �� ������� ������
            Sum = GetSaleSumm(INN, cPer)
            If Sum > minSale Then
            
                Set dateS = GetDatesList(cPer)
                
                '��������� ������ �����������
                Dim ndlist As Collection
                Set ndlist = New Collection
                For Each postdate In dateS
                    i = dateS(postdate)
                    Post = DTL.Cells(i, clNDS)
                    If Sum - Post >= 0 Then
                        Sum = Sum - Post
                        If Sum >= 0 Then ndlist.Add i
                        If Sum < maxDif Then Exit For
                    End If
                Next
                
                '�� ���������� ������ ����������� ����������� ������� ������
                For Each i In ndlist
                    DTL.Cells(i, clPND) = IndexToQYYYY(cPer)
                Next
                
            End If
            
        Next
        
    Next
    
    End
    
End Sub

'������ ����� �������� �������� � INN �� ������� Q
Function GetSaleSumm(ByVal INN As String, ByVal Q As Integer) As Double
    i = firstDat
    Sum = 0
    Do While DAT.Cells(i, cAccept) <> ""
        If DAT.Cells(i, cAccept) = "OK" And DAT.Cells(i, cSellINN) = INN Then
            If StupidQToQIndex(DAT.Cells(i, cPND)) = Q Then
                Sum = Sum + WorksheetFunction.Sum(Range(DAT.Cells(i, clNDS), DAT.Cells(i, clNDS + 2)))
            End If
        End If
        i = i + 1
    Loop
    GetSaleSumm = Sum
End Function

'������������ �������������� ������ ��� �������� � 12 ��������� ������� �� cPer
Function GetDatesList(ByVal cPer As Integer) As Object
            
    '�������� ���������� ����
    Set dateS = CreateObject("Scripting.Dictionary")
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        Dim d As Date
        d = DTL.Cells(i, clDate)
        Q = DateToQIndex(d)
        If Q >= cPer And Q <= cPer + 11 And DTL.Cells(i, clPND) = "" Then dateS(d) = i
        i = i + 1
    Loop
    
    '��������� ��������� ����
    Set datesSorted = CreateObject("Scripting.Dictionary")
    Do While dateS.Count > 0
        Dim max As Date
        max = 0
        For Each dt In dateS
            If max < dt Then max = dt
        Next
        datesSorted(max) = dateS(max)
        dateS.Remove (max)
    Loop
    
    Set GetDatesList = datesSorted
    
End Function

'******************** End of File ********************