Attribute VB_Name = "Main"
Const isRelease = True 'True - ����������� ������, False - ����� ������� (��� ��������, ��� ������ � �����)

Const FirstD = 8        '������ ������ � ��������� ������
Const FirstS = 5        '������ ������ � �������� ������
Const cFile = 16        '������� � ������ �����
Const cCode = 17        '������� � ����� �����

Const errName = "������"

Dim dat As Variant      '������� � �������
Dim src As Variant      '������� � �����������
Dim err As Variant      '������� � ��������
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
    If isRelease Then If MsgBox("������ ��������� ������� ��������� ������ ������ ������ � ����������. ����������?", vbYesNo) = vbNo Then Exit Sub
    Range(Cells(FirstD, 1), Cells(1048576, 50)).Clear
    Sheets(errName).Cells.Clear
er:
    Numerator.Clear
End Sub

'���� ������
Sub DataCollect()
    
    Set dat = ActiveSheet
    noEmpty = (dat.Cells(FirstD, 2) <> "")
    If isRelease And noEmpty Then If MsgBox("���������� ���� ������. ����������?", vbYesNo) = vbNo Then Exit Sub
    Message "����������"
    
    '�������� ��������� ������
    Set files = Source.getFiles(dat.Cells(1, 3))
        
    '������ ������� (���� � ���) ��� ������ ������
    Call NewTab(errName, True)
    Set err = Sheets(errName)
    err.Columns(1).ColumnWidth = 100
    err.Columns(2).ColumnWidth = 20
    err.Cells(1, 1) = "����"
    err.Cells(1, 2) = "���������"
            
    '������� ���������� ������ � ��������
    Application.ScreenUpdating = False
    i = FirstD
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
    i = FirstD
    Do While dat.Cells(i, 2) <> ""
        uid = dat.Cells(i, 1)
        If uid <> "" Then Indexes.Add uid, i
        i = i + 1
    Loop
    max = i
    
    '�������������� ������� � ���������� ��������
    Numerator.Init
    Verify.Init
    n = 1
    s = 0
    e = 0
    
    '������������ ������ ������
    For Each file In files
        curf = file
        If Len(curf) > 40 Then curf = "..." + Right(curf, 40)
        Message ("��������� ����� " + CStr(n) + " �� " + CStr(files.count) + " (" + curf) + ")"
        er = AddFile(file)
        If er > 0 Then
            e = e + 1
            err.Cells(e + 1, 1) = file
            If er = 1 Then err.Cells(e + 1, 2) = "������ �������� �����"
            If er = 2 Then err.Cells(e + 1, 2) = "������ � ������"
            If er = 3 Then err.Cells(e + 1, 2) = "����������� ���"
        Else
            s = s + 1
        End If
        n = n + 1
    Next
    Message ("������!")
    If isRelease Then MsgBox ("��������� ���������!" + Chr(13) + "������ ����������� �������: " + CStr(s) + Chr(13) + "����� � ��������: " + CStr(e))
    
End Sub

'���������� ������ �� ����� (���������� 0 - �� ������, 1 - ������ ��������, 2 - ������ � ������, 3 - ��� ����)
Function AddFile(ByVal file As String) As Byte
    errors = False
    On Error GoTo er
    Application.ScreenUpdating = False
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    If Not impBook Is Nothing Then
        Set src = impBook.Worksheets(1) '���� ���� ������ � ������� �����
        src.Unprotect Template.Secret
        cod = src.Cells(1, 1)
        If cod <> "" Then
        
            '������������ ������ ���������
            i = FirstS
            Do While NotEmpty(i)
                uid = src.Cells(i, 1)
                '������ ��� ���� (��������)
                If uid <> "" Then
                    ind = Indexes(uid)
                    If ind <> Empty Then
                        '� ������ ������������� ����, ��������� ������
                        If copyRecord(file, ind, i, True) Then errors = True
                    Else
                        '� ��� � ���, ����� ������ ���, ����� ���������� UID, �������� � ��� ���
                        uid = ""
                    End If
                End If
                '����� ������
                If uid = "" Then If copyRecord(file, max, i, False) Then errors = True
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

'���������� True ���� ������ si � ��������� �� ������
Function NotEmpty(si As Long) As Boolean
    NotEmpty = False
    For j = 1 To 14
        txt = src.Cells(si, j).text
        If txt <> "" And txt <> "#�/�" Then NotEmpty = True
    Next
End Function

'����������� ������. refresh - ���������� ������ (��������� ��� ����������)
'���������� True - ���� � ������ ���� ������
Function copyRecord(file As String, ByVal di As Long, ByVal si As Long, refresh As Boolean) As Boolean
    Dim changed As Boolean
    wht = RGB(255, 255, 255)
    yel = RGB(256, 256, 192)
    For j = 2 To 14
        ravno = dat.Cells(di, j).text = src.Cells(si, j).text
        dat.Cells(di, j) = src.Cells(si, j)
        dat.Cells(di, j).ClearFormats
        If j = 2 Or j = 4 Or j = 6 Or j = 7 Or j = 8 Then
            src.Cells(si, j).Interior.Color = yel
        Else
            src.Cells(si, j).Interior.Color = wht
        End If
        If refresh And Not ravno Then
            dat.Cells(di, j).Interior.Color = yel
            src.Cells(si, j).Interior.Color = yel
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
End Function