Attribute VB_Name = "CollectSale"
Dim LastRec As Long
Dim curFile As String
Dim curCode As String

'������ �������� ����� ������
Sub Run()
    
    Message "����������..."
    Numerator.Init
    Log.Init
    Verify.Init
    
    '�������� ��������� ������
    Set files = Source.getFiles(DirImportSale, True)
    
    '������������ ������ ������
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
    
    Values.CreateReport
    ActiveWorkbook.Save
    Message "������! ���� �������."
    Application.DisplayAlerts = True
    
    If isRelease Then MsgBox ("��������� ���������!" + Chr(13) + "������ ����������� �������: " + _
                                                CStr(s) + Chr(13) + "����� � ��������: " + CStr(e))
    
End Sub

'���������� ������ �� �����. ����������:
'0 - �� ������
'1 - ������ ��������
'2 - ������ � ������ (errors=true)
'3 - ��� ����
'4 - ������ ����� �� ��������������
'5 �� ������������, ��� ��� ��������� (���������� � Source)
'6 - ���� ��� ������
'7 - �������� � �������
'��������� �� ������� �� ���� ����� ������� � Log
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
        SetProtect SRC
        ver = SRC.Cells(2, 1).text
        If ver <> tmpVersion Then
            AddFile = 4
            impBook.Close False
            Exit Function
        End If
        curFile = file
        curCode = SRC.Cells(1, 1)
        If curCode <> "" Then
            
            '������� ���������� ������ ��� �������
            i = firstDat
            Do While DAT.Cells(i, cAccept) <> ""
                If DAT.Cells(i, cUIN) = "" And DAT.Cells(i, cCode) = curCode Then
                    DAT.Rows(i).Delete
                Else
                    i = i + 1
                End If
            Loop
        
            '����������� ������������ ������
            Set Indexes = CreateObject("Scripting.Dictionary")
            i = firstDat
            Do While DAT.Cells(i, cAccept) <> ""
                UID = DAT.Cells(i, cUIN)
                If UID <> "" Then Indexes.Add UID, i
                i = i + 1
            Loop
            LastRec = i
        
            '������������ ������ ���������
            Set resUIDs = CreateObject("Scripting.Dictionary")
            i = firstSrc
            Do While NotEmpty(i)
                UID = SRC.Cells(i, 1)
                '������ ��� ���� (��������)
                If UID <> "" Then
                    
                    ind = Indexes(UID)
                    If ind <> Empty Then
                        
                        '� ������ ������������� ����, ��������� ������
                        If Not copyRecord(ind, i, True) Then errors = True
                        
                        '������ �� ���������
                        stat = DAT.Cells(ind, cStatus).text
                        If stat = "0" Then
                            DAT.Cells(ind, cCom) = "������ ������������!"
                            DAT.Cells(ind, cCom).Interior.Color = colRed
                            SRC.Cells(i, cCom) = "������ ������������!"
                            SRC.Cells(i, cCom).Interior.Color = colRed
                        End If
                        If stat = "2" Then
                            DAT.Cells(ind, cCom) = "������ �������������!"
                            DAT.Cells(ind, cCom).Interior.Color = colGreen
                            SRC.Cells(i, cCom) = "������ �������������!"
                            SRC.Cells(i, cCom).Interior.Color = colGreen
                        End If
                        
                    Else
                        '� ��� � ���, ����� ������ ���, ����� ���������� UID, �������� � ��� ���
                        UID = ""
                    End If
                End If
                '����� ������
                If UID = "" Then If Not copyRecord(LastRec, i, False) Then errors = True
                rUID = SRC.Cells(i, 1).text
                
                '���������� ������� resUIDs - ��� ������, ������� ���� � �������
                '����� � ����� ���� ��� ������ �� ���� �� ����� �������, ������� ����������� � ����
                '�������, �� ������� ���������.
                '���� � ������� ����� ��� ���������� ������, �� ��� ����� ������!
                On Error Resume Next
                If rUID <> "" Then resUIDs.Add rUID, 1
                
                i = i + 1
            Loop
            
            '��������� �������� �� �������� ������ (�� ������� ��������� �����)
            i = firstDat
            Do While DAT.Cells(i, cAccept) <> ""
                UID = DAT.Cells(i, cUIN).text
                If UID <> "" And DAT.Cells(i, cCode) = curCode Then
                    If resUIDs(UID) = Empty Then
                        DAT.Cells(i, cCom) = "������ ������� ���������� (������ � ���)"
                        DAT.Cells(i, cCom).Interior.Color = colYellow
                        DAT.Cells(i, cAccept) = "lost"
                        AddFile = 2
                    End If
                End If
                i = i + 1
            Loop
            
        Else
            AddFile = 3
        End If
        impBook.Close saveSource
        
    End If
    
    Numerator.Save
    Application.ScreenUpdating = True
    DoEvents
    If errors Then AddFile = 2
    Exit Function

er:
    AddFile = 1

End Function

'�������� �� ������ ������
'���������� True, ���� ������ � ��������� �� ������
Function NotEmpty(ByVal i As Long) As Boolean
    NotEmpty = False
    For j = 1 To 15
        txt = SRC.Cells(i, j).text
        If txt <> "" And txt <> "#�/�" Then NotEmpty = True: Exit For
    Next
End Function

'����������� ������. ���������� True, ���� ������ ��������� ��� ������
'di - ������ � ������
'si - ������ � ����������
'refresh - true, ���� ���������� ������ (��������� ��� ����������)
Function copyRecord(ByVal di As Long, ByVal si As Long, refresh As Boolean) As Boolean
    
    stat = DAT.Cells(di, cStatus).text
    If stat = "0" Then
        copyRecord = False
        Exit Function
    End If
    
    SetFormates di
    SRC.Cells(si, 1).ClearFormats
    
    '������ �������������, ���������� ��������� �� ��������� ������ � ������
    If stat = "2" Then
        For j = 2 To 14
            CheckChanges di, si, j
            SRC.Cells(si, j) = DAT.Cells(di, j)
        Next
        copyRecord = True
        Exit Function
    End If
    
    If refresh And DAT.Cells(di, cAccept) = "OK" Then
        oldSum = 0
        For i = 12 To 14
            If DAT.Cells(di, i) <> "" Then oldSum = oldSum + DAT.Cells(di, i)
        Next
        RestoreBalance DAT.Cells(di, cDates), DAT.Cells(di, cSellINN).text, oldSum
    End If
    
    '����������� ������� � ��������� �� ���������
    For j = 2 To 14
        If j <> 6 Then
            CheckChanges di, si, j
            If Not IsError(SRC.Cells(si, j)) Then DAT.Cells(di, j) = SRC.Cells(si, j)
        Else
            s = selIndexes(DAT.Cells(di, 5).text)
            If s <> Empty Then
                DAT.Cells(di, 6) = DIC.Cells(s, 1)
            Else
                AddCom ("��� �� ������ � �����������")
            End If
        End If
    Next
    DAT.Cells(di, cFile) = curFile
    DAT.Cells(di, cCode) = curCode
    DAT.Cells(di, cAccept) = "fail" '�� ��������� ����� ������� ������ �� ������
    
    '�������� �� �������� ������ (���� ��� ���������� � ������ � ����� ������)
    If refresh And SRC.Cells(si, cDates).text = "" Then
        SRC.Cells(si, 1).Font.Color = colWhite
        SRC.Cells(si, cCom) = "������ ������� ����������"
        SRC.Cells(si, cCom).Interior.Color = colYellow
        DAT.Cells(di, cCom) = "������ ������� ����������"
        DAT.Cells(di, cCom).Interior.Color = colYellow
        DAT.Cells(di, cAccept) = "lost"
        copyRecord = True
        Exit Function
        '���������� �������� � ���� ������ �� ���������, �������...
    End If
    
    copyRecord = Verify.Verify(di, si, oldINN, oldSum)
    
    '���� �����, ����������� ������ ����� �����
    If copyRecord Then
        Dim needNum As Boolean
        If refresh Then
            needNum = Not Numerator.CheckPrefix(DAT.Cells(di, 1).text, _
                DAT.Cells(di, 2), DAT.Cells(di, cSellINN).text)
        Else
            needNum = True
        End If
        If needNum Then
            n = Numerator.Generate(DAT.Cells(di, 2), DAT.Cells(di, cSellINN).text)
            DAT.Cells(di, cUIN).NumberFormat = "@"
            DAT.Cells(di, cUIN) = n
            DAT.Cells(di, cDateCol) = DateTime.Now
            SRC.Cells(si, 1).NumberFormat = "@"
            SRC.Cells(si, 1) = n
        End If
        DAT.Cells(di, cAccept) = "OK"
    End If
    
    If Not refresh Then LastRec = LastRec + 1
    If DAT.Cells(di, cStatus).text = "" Then DAT.Cells(di, cStatus) = 1
    
End Function

Sub SetFormates(ByVal i As Long)
    DAT.Cells(i, 2).NumberFormat = "dd.MM.yyyy"
    DAT.Cells(i, 7).NumberFormat = "### ### ##0.00"
    For j = 9 To 11
        DAT.Cells(i, j).NumberFormat = "### ### ##0.00"
    Next
    For j = 12 To 14
        DAT.Cells(i, j).NumberFormat = "### ### ##0.00"
    Next
End Sub

'������������ ��������� � ������� �� ������
Sub CheckChanges(ByVal di As Long, ByVal si As Long, ByVal j As Long)
    
    '����� �������
    DAT.Cells(di, j).Interior.Color = colWhite
    If j = 2 Or j = 4 Or j = 6 Or j = 7 Or j = 8 Then
        SRC.Cells(si, j).Interior.Color = colYellow
    Else
        SRC.Cells(si, j).Interior.Color = colWhite
    End If
    
    '���������, ���� ���� �������
    If DAT.Cells(di, j).text <> SRC.Cells(si, j).text Then
        DAT.Cells(di, j).Interior.Color = colBlue
        SRC.Cells(si, j).Interior.Color = colBlue
    End If

End Sub

'******************** End of File ********************