Attribute VB_Name = "CollectLoad"
Dim LastRec As Long
Dim curFile As String
Dim curMark As String
Dim curProv As String
Dim curProvINN As String

'������ �������� ����� ������
Sub Run()
    
    Log.Init
    Range(DTL.Cells(firstDtL, 1), DTL.Cells(maxRow, 18)).Clear
    Range(DTL.Cells(firstDtL, 17), DTL.Cells(maxRow, 18)).Interior.Color = colGray
    Range(DTL.Cells(firstDtL, 17), DTL.Cells(maxRow, 18)).Font.Color = RGB(166, 166, 166)
    LastRec = firstDtL
    
    '�������� ��������� ������
    Set files = Source.getFiles(DirImportLoad, False)

    n = 1
    s = 0
    e = 0
    For Each file In files
        curf = file
        If Len(curf) > 40 Then curf = "..." + Right(curf, 40)
        Message ("��������� ����� " + CStr(n) + " �� " + CStr(files.Count) + " (" + curf) + ")"
        er = AddFile(file)
        If er > 0 Then
            Log.Rec file, er
            e = e + 1
        Else
            s = s + 1
        End If
        n = n + 1
    Next

    'ActiveWorkbook.Save
    Message "������! ���� �������."
    Application.DisplayAlerts = True
    
    If isRelease Then MsgBox ("��������� ���������!" + Chr(13) + "������ ����������� �������: " + _
                                                CStr(s) + Chr(13) + "����� � ��������: " + CStr(e))
    
End Sub

'���������� ������ �� �����. ����������:
'0 - �� ������
'1 - ������ ��������
'2 - ������ � ������ (errors=true)
'3 - ��� �������, ��� �� �� ������
Function AddFile(ByVal file As String) As Byte
    
    '����������
    Application.DisplayAlerts = False
    If Not TrySave(file) Then AddFile = 6: Exit Function
    errors = False
    Application.ScreenUpdating = False
    If isRelease Then On Error GoTo er
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    
    If Not impBook Is Nothing Then
        
        Set SRC = impBook.Worksheets(1) '���� ���� ������ � ������� �����
        curMark = UCase(SRC.Cells(2, 2).text)
        If curMark <> "�" And curMark <> "�" Then
            AddFile = 3
            impBook.Close False
            Exit Function
        End If
        
        curProv = Mid(SRC.Cells(3, 1).text, 10, Len(SRC.Cells(3, 1).text) - 9)
        curProvINN = Right(SRC.Cells(4, 1).text, 10)
        
        i = 10
        Do While SRC.Cells(i, 2).text = "01"
            If Not copyRecord(i) Then errors = True
            DTL.Cells(LastRec, clFile) = file
            LastRec = LastRec + 1
            i = i + 1
        Loop
        
        impBook.Close False
        
    End If
    
    Application.ScreenUpdating = True
    DoEvents
    If errors Then AddFile = 2
    Exit Function

er:
    AddFile = 1
    
End Function

'����������� ������. ���������� True, ���� ������ ��������� ��� ������
'si - ������ � ����������
Function copyRecord(ByVal si As Long) As Boolean
    
    DTL.Cells(LastRec, 1) = curMark
    DTL.Cells(LastRec, 3).NumberFormat = "@"
    DTL.Cells(LastRec, 3) = curProvINN
    DTL.Cells(LastRec, 4) = curProv
    DTL.Cells(LastRec, 5).NumberFormat = "@"
    DTL.Cells(LastRec, 5) = SRC.Cells(si, 10)
    DTL.Cells(LastRec, 6) = SRC.Cells(si, 11)
    
    DTL.Cells(LastRec, 7) = SRC.Cells(si, 16)
    
    DTL.Cells(LastRec, 8) = SRC.Cells(si, 17)
    DTL.Cells(LastRec, 9) = SRC.Cells(si, 18)
    DTL.Cells(LastRec, 10) = SRC.Cells(si, 19)
    
    DTL.Cells(LastRec, 11) = SRC.Cells(si, 21)
    DTL.Cells(LastRec, 12) = SRC.Cells(si, 22)
    DTL.Cells(LastRec, 13) = SRC.Cells(si, 23)
    
    copyRecord = True
    
End Function