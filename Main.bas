Attribute VB_Name = "Main"
Public Const isRelease = True   'True - ����������� ������, False - ����� ������� (��� ��������, ��� ������ � �����)

Public Const firstDat = 8       '������ ������ � ��������� ������
Public Const firstSrc = 5       '������ ������ � �������� ������
Public Const firstDic = 5       '������ ������ � �����������
Public Const cCom = 15          '������� ��� �����������
Public Const cStatus = 16       '������� �������
Public Const cFile = 17         '������� � ������ �����
Public Const cCode = 18         '������� � ����� �����

Public Const tabDic = "����������"
Public Const tabErr = "������"
Public Const tabNum = "������� ����������"

Public colWhite As Long '�����
Public colRed As Long
Public colGreen As Long
Public colYellow As Long

Dim dat As Variant      '������� � �������
Dim src As Variant      '������� � �����������
Dim Indexes As Object   '������� ��������
Dim max As Long         '��������� ������ � ������
Dim i As Long
Dim file As Variant
Dim cod As String

'����� ���������� � �������
Sub DirSelect()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Cells(1, 3) = diag.SelectedItems(1)
End Sub

'�������� ���� ������ (�������� �����)
Sub Clear()
    On Error GoTo er
    If isRelease Then If MsgBox("��������! " + Chr(10) + Chr(10) + _
        "������ ��������� ������� ��� ��������� ������ ������ ������ � ����������. " + _
        "��� ������������������ ������ ��� ��������� ����������� ����� ��������� ������ ���." + _
        Chr(10) + Chr(10) + "����������?", vbYesNo) = vbNo Then Exit Sub
    Range(Cells(firstDat, 1), Cells(1048576, 50)).Clear
    Sheets(errName).Cells.Clear
er:
    Numerator.Clear
End Sub

'���� ������
Sub DataCollect()
    
    Set dat = ActiveSheet
    noEmpty = (dat.Cells(firstDat, 2) <> "")
    If isRelease And noEmpty Then If MsgBox("���������� ���� ������. ����������?", vbYesNo) = vbNo Then Exit Sub
    
    '�������������
    Message "����������"
    colWhite = RGB(255, 255, 255)
    colRed = RGB(255, 192, 192)
    colGreen = RGB(192, 255, 192)
    colYellow = RGB(255, 255, 192)
    Numerator.Init
    Log.Init
    Verify.Init
    n = 1
    s = 0
    e = 0
    
    '�������� ��������� ������
    Set files = Source.getFiles(dat.Cells(1, 3))
    
    '������������ ������ ������
    For Each file In files
        curf = file
        If Len(curf) > 40 Then curf = "..." + Right(curf, 40)
        Message ("��������� ����� " + CStr(n) + " �� " + CStr(files.Count) + " (" + curf) + ")"
        er = AddFile(file)
        If er > 0 Then
            Call Log.Rec(file, er)
            e = e + 1
        Else
            s = s + 1
        End If
        n = n + 1
    Next
    Message ("������!")
    If isRelease Then MsgBox ("��������� ���������!" + Chr(13) + "������ ����������� �������: " + CStr(s) + Chr(13) + "����� � ��������: " + CStr(e))
    
End Sub

'���������� ������ �� �����. ����������:
'0 - �� ������
'1 - ������ ��������
'2 - ������ � ������
'3 - ��� ����
'4 - ������ ������������
Function AddFile(ByVal file As String) As Byte
    errors = False
    If isRelease Then On Error GoTo er
    Application.ScreenUpdating = False
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    If Not impBook Is Nothing Then
        Set src = impBook.Worksheets(1) '���� ���� ������ � ������� �����
        src.Unprotect Template.Secret
        cod = src.Cells(1, 1)
        If cod <> "" Then
            
            '������� ���������� ������ � ��������
            i = firstDat
            Do While dat.Cells(i, 2) <> ""
                If dat.Cells(i, 1) = "" And dat.Cells(i, cCode) = cod Then
                    dat.Rows(i).Delete
                    max = max - 1
                Else
                    i = i + 1
                End If
            Loop
        
            '����������� ������������ ������
            Set Indexes = CreateObject("Scripting.Dictionary")
            i = firstDat
            Do While dat.Cells(i, 2) <> ""
                uid = dat.Cells(i, 1)
                If uid <> "" Then Indexes.Add uid, i
                i = i + 1
            Loop
            max = i
        
            '������������ ������ ���������
            Set resuids = CreateObject("Scripting.Dictionary")
            i = firstSrc
            Do While NotEmpty(i)
                uid = src.Cells(i, 1)
                '������ ��� ���� (��������)
                If uid <> "" Then
                    
                    ind = Indexes(uid)
                    If ind <> Empty Then
                        
                        '� ������ ������������� ����, ��������� ������
                        If copyRecord(ind, i, True) Then errors = True
                        
                        '������ �� ���������
                        stat = dat.Cells(ind, cStatus).text
                        If stat = "0" Then
                            dat.Cells(ind, cCom) = "������ ������������!"
                            dat.Cells(ind, cCom).Interior.Color = colRed
                            src.Cells(i, cCom) = "������ ������������!"
                            src.Cells(i, cCom).Interior.Color = colRed
                        End If
                        If stat = "2" Then
                            dat.Cells(ind, cCom) = "������ �������������!"
                            dat.Cells(ind, cCom).Interior.Color = colGreen
                            src.Cells(i, cCom) = "������ �������������!"
                            src.Cells(i, cCom).Interior.Color = colGreen
                        End If
                        
                    Else
                        '� ��� � ���, ����� ������ ���, ����� ���������� UID, �������� � ��� ���
                        uid = ""
                    End If
                End If
                '����� ������
                If uid = "" Then If copyRecord(max, i, False) Then errors = True
                resuids.Add src.Cells(i, 1).text, 1
                i = i + 1
            Loop
            
            '��������� �������� �� �������� ������
            i = firstDat
            Do While dat.Cells(i, 2) <> ""
                uid = dat.Cells(i, 1)
                If uid <> "" And dat.Cells(i, cCode) = cod Then
                    If resuids(uid) = Empty Then
                        dat.Cells(i, cCom) = "������ �������!"
                        dat.Cells(i, cCom).Interior.Color = colRed
                        AddFile = 2
                    End If
                End If
                i = i + 1
            Loop
            
        Else
            AddFile = 3
        End If
        src.Protect Template.Secret
        impBook.Close isRelease
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
Function NotEmpty(i As Long) As Boolean
    NotEmpty = False
    For j = 1 To 14
        txt = src.Cells(i, j).text
        If txt <> "" And txt <> "#�/�" Then NotEmpty = True: Exit For
    Next
End Function

'����������� ������. ���������� True, ���� � ������ ���� ������
'di - ������ � ������
'si - ������ � ����������
'refresh - true, ���� ���������� ������ (��������� ��� ����������)
Function copyRecord(ByVal di As Long, ByVal si As Long, refresh As Boolean) As Boolean
    
    stat = dat.Cells(di, cStatus).text
    If stat = "0" Or stat = "2" Then Exit Function
    
    Dim changed As Boolean
    For j = 2 To 14
        ravno = dat.Cells(di, j).text = src.Cells(si, j).text
        dat.Cells(di, j) = src.Cells(si, j)
        dat.Cells(di, j).ClearFormats
        If j = 2 Or j = 4 Or j = 6 Or j = 7 Or j = 8 Then
            src.Cells(si, j).Interior.Color = colYellow
        Else
            src.Cells(si, j).Interior.Color = colWhite
        End If
        If refresh And Not ravno Then
            dat.Cells(di, j).Interior.Color = colYellow
            src.Cells(si, j).Interior.Color = colYellow
            changed = True
        End If
    Next
    dat.Cells(di, cFile) = file
    dat.Cells(di, cCode) = cod
    Range(dat.Cells(di, cFile), dat.Cells(di, cCode)).Font.Color = RGB(192, 192, 192)
    errors = Verify.Verify(dat, src, di, si, changed)
    If errors Then
        copyRecord = True
    Else
        '���� ��� ������, � ��� �� ����������, ����������� �����
        If Not refresh Then
            num = Numerator.Generate(dat.Cells(di, 2), dat.Cells(di, 4))
            dat.Cells(di, 1) = num
            src.Cells(si, 1) = num
        End If
    End If
    If Not refresh Then max = max + 1
    If dat.Cells(di, cStatus).text = "" Then dat.Cells(di, cStatus) = 1
End Function