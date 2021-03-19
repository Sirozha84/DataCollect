Attribute VB_Name = "CollectLoad"
Dim LastRec As Long
Dim curFile As String
Dim curMark As String
Dim curProv As String
Dim curProvINN As String

'������ �������� ����� ������
Sub Run()
    
    Log.Init
    Range(DTL.Cells(firstDtL, 1), DTL.Cells(maxRow, clAccept)).Clear
    Range(DTL.Cells(firstDtL, clFile), DTL.Cells(maxRow, clAccept)).Interior.Color = colGray
    Range(DTL.Cells(firstDtL, clFile), DTL.Cells(maxRow, clAccept)).Font.Color = RGB(166, 166, 166)
    LastRec = firstDtL
    
    '�������� ��������� ������ � ������ ����
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

    '��������� ������ � �����������
    Message "������ ����������� �������"
    Range(DIC.Cells(firstDic, cPBalance), DIC.Cells(maxRow, cPBalance + quartCount * 2 - 1)).Clear
    Set salers = CreateObject("Scripting.Dictionary")
    i = firstDic
    Do While DIC.Cells(i, cINN) <> ""
        salers(DIC.Cells(i, cINN).text) = i
        i = i + 1
    Loop
    lastdic = i
    
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        If DTL.Cells(i, clAccept) = "OK" Then
            INN = DTL.Cells(i, clInINN).text
            '���������� ������ �������� � ����������
            If salers(INN) = "" Then
                salers(INN) = lastdic
                DIC.Cells(lastdic, cSellerName) = DTL.Cells(i, clInName)
                DIC.Cells(lastdic, cINN).NumberFormat = "@"
                DIC.Cells(lastdic, cINN) = INN
                For j = 0 To quartCount - 1
                    DIC.Cells(lastdic, cLimits + j).NumberFormat = "### ### ##0.00"
                    DIC.Cells(lastdic, cLimits + j).FormulaR1C1 = _
                            "=SUM(RC[" + CStr(24 + j) + "]:RC[" + CStr(47 - j) + "])-" + _
                            "SUM(RC[12]:RC[" + CStr(23 - j) + "])"
                Next
                lastdic = lastdic + 1
            End If
            '���������� �����������
            qi = DateToQIndex(DTL.Cells(i, 3))
            If qi >= 0 Then
                Sum = 0
                For j = 12 To 14
                    If IsNumeric(DTL.Cells(i, j)) Then Sum = Sum + DTL.Cells(i, j)
                Next
                s = salers(INN) '������
                qi = qi * 2 + cPBalance
                If DTL.Cells(i, 1).text = "�" Then qi = qi + 1
                DIC.Cells(s, qi) = DIC.Cells(s, qi) + Sum
            End If
        End If
        i = i + 1
    Loop

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
            If Not copyRecord(i) Then
                errors = True
                DTL.Cells(LastRec, clAccept) = "fail"
            Else
                DTL.Cells(LastRec, clDateCol) = DateTime.Now
                DTL.Cells(LastRec, clAccept) = "OK"
            End If
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
    
    DTL.Cells(LastRec, clMark) = curMark
    DTL.Cells(LastRec, clNum) = SRC.Cells(si, 1)
    DTL.Cells(LastRec, clDate).NumberFormat = "dd.MM.yyyy"
    DTL.Cells(LastRec, clDate) = SRC.Cells(si, 3)
    DTL.Cells(LastRec, clOutINN).NumberFormat = "@"
    DTL.Cells(LastRec, clOutINN) = curProvINN
    DTL.Cells(LastRec, clOutName) = curProv
    DTL.Cells(LastRec, clInINN).NumberFormat = "@"
    DTL.Cells(LastRec, clInINN) = SRC.Cells(si, 10)
    DTL.Cells(LastRec, clInName) = SRC.Cells(si, 9)
    DTL.Cells(LastRec, clPrice) = SRC.Cells(si, 16)
    DTL.Cells(LastRec, clPrice + 1) = SRC.Cells(si, 17)
    DTL.Cells(LastRec, clPrice + 2) = SRC.Cells(si, 18)
    DTL.Cells(LastRec, clPrice + 3) = SRC.Cells(si, 19)
    DTL.Cells(LastRec, clPrice + 4) = SRC.Cells(si, 21)
    DTL.Cells(LastRec, clPrice + 5) = SRC.Cells(si, 22)
    DTL.Cells(LastRec, clPrice + 6) = SRC.Cells(si, 23)
    
    copyRecord = VerifyLoad(LastRec)
    
End Function