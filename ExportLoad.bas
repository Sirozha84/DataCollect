Attribute VB_Name = "ExportLoad"
'��������� ������: 03.04.2021 21:36

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
Sub CreateExportFile(ByVal inn As String, ByVal NUM As String)
    
    saler = SellFileName(inn)
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
    
    For Each inn In selIndexes
        
        Si = selIndexes(inn)
        oND = StupidQToQIndex(DIC.Cells(Si, cOPND))
        
        For cPer = oND To Quartals
            
            '������ ����� ���� �������� �� ������� ������
            Sum = 0
            i = firstDat
            Do While DAT.Cells(i, cAccept) <> ""
                If DAT.Cells(i, cAccept) = "OK" And DAT.Cells(i, cSellINN) = inn Then
                    If StupidQToQIndex(DAT.Cells(i, cPND)) = cPer Then
                        Sum = Sum + WorksheetFunction.Sum(Range(DAT.Cells(i, clNDS), DAT.Cells(i, clNDS + 2)))
                    End If
                End If
                i = i + 1
            Loop
            If Sum > minSale Then
            
                Set dateS = GetDatesList(cPer)

                '����� ������ ������ ������� � ������ ������ � dataS

            
            End If
            
        Next
        
    Next
    
    End
    
End Sub

'������������ �������������� ������ ��� �������� � 12 ��������� ������� �� ���������
Function GetDatesList(ByVal cPer As Integer) As Object
            
    '�������� ���������� ����
    Set dateS = CreateObject("Scripting.Dictionary")
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        Dim d As Date
        d = DTL.Cells(i, clDate)
        q = DateToQIndex(d)
        If q >= cPer And q <= cPer + 11 Then dateS(d) = i
        i = i + 1
    Loop
    
    '��������� ��������� ����
    Set datesSorted = CreateObject("Scripting.Dictionary")
    Do While dateS.Count > 0
        '������� ������������ ����
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