Attribute VB_Name = "CollectLoad"
'��������� ������: 28.03.2021 16:22

Dim LastRec As Long
Dim curFile As String
Dim curMark As String
Dim curProv As String
Dim curProvINN As String
Dim UINs As Object

'������ �������� ����� ������
Sub Run()
    
    Message "����������..."
    Dictionary.Init
    Log.Init
    
    '������� ���� �� ������ ���������� �������
    Set UINs = CreateObject("Scripting.Dictionary")
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        If DTL.Cells(i, clAccept) = "OK" Then
            UINs(DTL.Cells(i, clUIN).text) = i
        Else
            DTL.Rows(i).Delete
            i = i - 1
        End If
        i = i + 1
    Loop
    LastRec = i
    
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
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        If DTL.Cells(i, clAccept) = "OK" Then
            inn = DTL.Cells(i, clSaleINN).text
            '���������� �����������
            Qi = DateToQIndex(DTL.Cells(i, 3))
            If Qi >= 0 Then
                Sum = 0
                For j = 12 To 14
                    If IsNumeric(DTL.Cells(i, j)) Then Sum = Sum + DTL.Cells(i, j)
                Next
                Si = selIndexes(inn) '������
                Qi = Qi * 2 + cPBalance
                If DTL.Cells(i, 1).text = "�" Then Qi = Qi + 1
                DIC.Cells(Si, Qi) = DIC.Cells(Si, Qi) + Sum
            End If
        End If
        i = i + 1
    Loop

    ActiveWorkbook.Save
    Message "������! ���� �������."
    Application.DisplayAlerts = True
    
    MsgBox ("��������� ���������!" + Chr(13) + "������ ����������� �������: " + _
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
    On Error GoTo er
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    On Error GoTo 0
    
    If Not impBook Is Nothing Then
        
        Set SRC = impBook.Worksheets(1)
        
        If Left(SRC.Cells(1, 1), 6) = "������" Then
            curMark = UCase(SRC.Cells(6, 4).text)
            If curMark <> "�" And curMark <> "�" Then
                AddFile = 3
                impBook.Close False
                Exit Function
            End If
            
            curProv = Split(SRC.Cells(3, 2).text, ": ")(1)
            curProvINN = Right(SRC.Cells(4, 2).text, 20)
            
            i = 12
            Do While SRC.Cells(i, 2).text <> ""
                If UINs(SRC.Cells(i, 21).text) = "" Then
                    If Not copyRecord(i) Then
                        errors = True
                        DTL.Cells(LastRec, clAccept) = "fail"
                    Else
                        DTL.Cells(LastRec, clDateCol) = DateTime.Now
                        uin = GenerateLoad
                        DTL.Cells(LastRec, clUIN) = uin
                        SRC.Cells(i, 21) = uin
                        DTL.Cells(LastRec, clAccept) = "OK"
                    End If
                    DTL.Cells(LastRec, clFile) = file
                    LastRec = LastRec + 1
                End If
                i = i + 1
            Loop
        End If
        
        On Error GoTo er
        impBook.Close True
        
    End If
    
    Application.ScreenUpdating = True
    DoEvents
    If errors Then AddFile = 2
    Exit Function

er:
    AddFile = 1
    
End Function

'����������� ������. ���������� True, ���� ������ ��������� ��� ������
'Si - ������ � ����������
Function copyRecord(ByVal Si As Long) As Boolean
    
    DTL.Cells(LastRec, clMark) = curMark
    DTL.Cells(LastRec, clNum) = SRC.Cells(Si, 1)
    DTL.Cells(LastRec, clDate).NumberFormat = "dd.MM.yyyy"
    DTL.Cells(LastRec, clDate) = Right(SRC.Cells(Si, 5), 10)
    DTL.Cells(LastRec, clProvINN).NumberFormat = "@"
    DTL.Cells(LastRec, clProvINN) = curProvINN
    DTL.Cells(LastRec, clProvName) = curProv
    DTL.Cells(LastRec, clSaleINN).NumberFormat = "@"
    DTL.Cells(LastRec, clSaleINN) = Left(SRC.Cells(Si, 10), 10)
    DTL.Cells(LastRec, clSaleName) = SRC.Cells(Si, 9)
    DTL.Cells(LastRec, clPrice) = SRC.Cells(Si, 15)
    DTL.Cells(LastRec, clNDS) = SRC.Cells(Si, 16)
    
    copyRecord = VerifyLoad(LastRec)
    
End Function

Function CompNameSeparate(ByVal s As String) As String
    CompNameSeparate = Trim(Split(s, ":"))
End Function

'******************** End of File ********************