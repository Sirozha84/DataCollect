Attribute VB_Name = "Main"
'Const isRelease = True  'True - ����������� ������, False - ����� ������� (��� ��������, ��� ������ � �����)
Const isRelease = False 'True - ����������� ������, False - ����� ������� (��� ��������, ��� ������ � �����)

Const FirstD = 6        '������ ������ � ��������� ������
Const FirstS = 5        '������ ������ � �������� ������
Const cFile = 17        '������� � ������ �����
Const cCode = 18        '������� � ����� �����
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
    Set Files = Source.GetList(dat.Cells(1, 3))
        
    '������ ������� (���� � ���) ��� ������ ������
    Call NewTab(errName, True)
    Set err = Sheets(errName)
    err.Columns(1).ColumnWidth = 100
    err.Columns(2).ColumnWidth = 20
    err.Cells(1, 1) = "����"
    err.Cells(1, 2) = "���������"
    
    '����������� ������������ ������
    Set Indexes = CreateObject("Scripting.Dictionary")
    i = FirstD
    Do While dat.Cells(i, 2) <> ""
        uid = dat.Cells(i, 1)
        If uid <> "" Then Indexes.Add uid, i
        i = i + 1
    Loop
    max = i
    
    '�������������� ������� ����������
    Numerator.Init
    
    n = 1
    s = 0
    e = 0
    For Each file In Files
        Message ("��������� ����� " + CStr(n) + " �� " + CStr(Files.Count) + " (" + Source.FSO.getfilename(file)) + ")"
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
    'On Error GoTo er
    Application.ScreenUpdating = False
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    If Not impBook Is Nothing Then
        Set src = impBook.Worksheets(1) '���� ���� ������ � ������� �����
        cod = src.Cells(1, 1)
        If cod <> "" Then
        
            '������� ���������� ������ � ��������
            i = FirstD
            Do While dat.Cells(i, 2) <> ""
                If dat.Cells(i, 1) = "" And dat.Cells(i, cCode) = cod Then
                    dat.Rows(i).Delete
                    max = max - 1
                Else
                    i = i + 1
                End If
            Loop
            
            '������������ ������ ���������
            i = FirstS
            Do While src.Cells(i, 2) <> ""
                
                uid = src.Cells(i, 1)
                If uid = "" Then
                    '������ ���
                    AddFile = copyRecord(file, max, i, False)
                    max = max + 1
                Else
                    '������ ����
                    ind = Indexes(uid)
                    AddFile = copyRecord(file, ind, i, True)
                End If
                i = i + 1
            Loop
            
        Else
            AddFile = 3
        End If
        impBook.Close isRelease
    End If
    Numerator.Save
    Application.ScreenUpdating = True
    DoEvents
    Exit Function
er:
    AddFile = 1
End Function

'����������� ������. refresh - ���������� ������ (��������� ��� ����������)
Function copyRecord(file As String, ByVal di As Long, ByVal si As Long, refresh As Byte) As Byte
    For j = 2 To 14
        If refresh Then
            If dat.Cells(di, j) = src.Cells(si, j) Then
                dat.Cells(di, j).ClearFormats
            Else
                dat.Cells(di, j).Interior.Color = RGB(256, 256, 192)
            End If
        End If
        dat.Cells(di, j) = src.Cells(si, j)
    Next
    dat.Cells(di, cFile) = file
    dat.Cells(di, cCode) = cod
    Range(dat.Cells(di, cFile), dat.Cells(di, cCode)).Font.Color = RGB(192, 192, 192)
    errors = Verify.Verify(dat, src, di, si)
    If errors Then
        copyRecord = 2
    Else
        '���� ��� ������, ����������� �����, ���� ��� ���
        If dat.Cells(di, 1) = "" Then
            num = Numerator.Generate(dat.Cells(di, 2), dat.Cells(di, 4))
            dat.Cells(di, 1) = num
            src.Cells(si, 1) = num
        End If
    End If
End Function