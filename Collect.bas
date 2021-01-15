Attribute VB_Name = "Collect"
Dim LastRec As Long
Dim curFile As String
Dim curCode As String

'������ �������� ����� ������
Sub Run()
    
    Numerator.Init
    Log.Init
    Verify.Init
    
    '�������� ��������� ������
    Set files = Source.getFiles(DAT.Cells(1, 3))
    
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
    
    Verify.SaveValues
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
                If DAT.Cells(i, 1) = "" And DAT.Cells(i, cCode) = curCode Then
                    DAT.Rows(i).Delete
                Else
                    i = i + 1
                End If
            Loop
        
            '����������� ������������ ������
            Set Indexes = CreateObject("Scripting.Dictionary")
            i = firstDat
            Do While DAT.Cells(i, cAccept) <> ""
                UID = DAT.Cells(i, 1)
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
                If rUID <> "" Then resUIDs.Add rUID, 1
                '���������� ������� resUIDs - ��� ������, ������� ���� � �������
                '����� � ����� ���� ��� ������ �� ���� �� ����� �������, ������� ����������� � ����
                '�������, �� ������� ���������.
                '���� � ������� ����� ��� ���������� ������, �� ��� ����� ������!
                i = i + 1
            Loop
            
            '��������� �������� �� �������� ������
            i = firstDat
            Do While DAT.Cells(i, cAccept) <> ""
                UID = DAT.Cells(i, 1).text
                If UID <> "" And DAT.Cells(i, cCode) = curCode Then
                    If resUIDs(UID) = Empty Then
                        DAT.Cells(i, cCom) = "������ �������!"
                        DAT.Cells(i, cCom).Interior.Color = colRed
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
    For j = 1 To 14
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
        Exit Function
    End If
    
    SetFormates di
    
    '������ �������������, ���������� ��������� �� ��������� ������ � ������
    If stat = "2" Then
        For j = 2 To 14
            CheckChanges di, si, j
            SRC.Cells(si, j) = DAT.Cells(di, j)
        Next
        Exit Function
    End If
    
    '����������� ������� � ��������� �� ���������
    For j = 2 To 14
        CheckChanges di, si, j
        DAT.Cells(di, j) = SRC.Cells(si, j)
    Next
    DAT.Cells(di, cFile) = curFile
    DAT.Cells(di, cCode) = curCode
    errors = Verify.Verify(di, si)
    
    '���� �����, ����������� ������ ����� �����
    If Not errors Then
        Dim needNum As Boolean
        If refresh Then
            needNum = Not Numerator.CheckPrefix(DAT.Cells(di, 1).text, _
                DAT.Cells(di, 2), DAT.Cells(di, cSeller).text)
        Else
            needNum = True
        End If
        If needNum Then
            n = Numerator.Generate(DAT.Cells(di, 2), DAT.Cells(di, cSeller).text)
            DAT.Cells(di, 1) = n
            SRC.Cells(si, 1) = n
        End If
        DAT.Cells(di, cAccept) = "OK"
        copyRecord = True
    Else
        DAT.Cells(di, cAccept) = "fail"
        copyRecord = False
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