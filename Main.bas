Attribute VB_Name = "Main"
'Const isRelease = True  'True - ����������� ������, False - ����� ������� (��� ��������, ��� ������ � �����)
Const isRelease = False 'True - ����������� ������, False - ����� ������� (��� ��������, ��� ������ � �����)
Const FirstC = 6        '������ ������ � ��������� ������
Const FirstS = 5        '������ ������ � �������� ������
Const cFile = 17        '������� � ������ �����
Const cCode = 18        '������� � ����� �����
Const errName = "������"
Dim dat As Variant      '������� � �������
Dim err As Variant      '������� � ��������
Dim Indexes As Object   '������� ��������
Dim max As Long         '��������� ������ � ������

'����� ���������� � �������
Sub DirSelect()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Cells(1, 3) = diag.SelectedItems(1)
End Sub

'�������� ���� ������ (�������� �����)
Sub Clear()
    On Error GoTo er
    'If MsgBox("������ ��������� ������� ��������� ������ ������ ������ � ����������. ����������?", vbYesNo) = vbNo Then Exit Sub
    Range(Cells(FirstC, 1), Cells(1048576, 50)).Clear
    Sheets(errName).Cells.Clear
er:
    Numerator.Clear
End Sub

'���� ������
Sub DataCollect()
    
    Set dat = ActiveSheet
    noEmpty = (dat.Cells(FirstC, 2) <> "")
    If isRelease And noEmpty Then If MsgBox("���������� ���� ������. ����������?", vbYesNo) = vbNo Then Exit Sub
    Message "����������"
    
    '�������� ��������� ������
    Set Files = Source.GetList("C:\Users\SG\OneDrive\������\�������� ������\���� ������\������")
        
    '������ ������� (���� � ���) ��� ������ ������
    Call NewTab(errName, True)
    Set err = Sheets(errName)
    err.Columns(1).ColumnWidth = 100
    err.Columns(2).ColumnWidth = 20
    err.Cells(1, 1) = "����"
    err.Cells(1, 2) = "���������"
    
    '����������� ��������� ������
    
    
    '�������������� ������� ����������
    Numerator.Init
    
    n = 1
    s = 0
    e = 0
    'max = FindMax(dat, FirstC, 2) - noEmpty '���� �� ����� +1, ������ ��� � ���� ������ ����� ������������ ������ ����� �����
    
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
    
    On Error GoTo er
    
    Application.ScreenUpdating = False
    
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    If Not impBook Is Nothing Then
        Set src = impBook.Worksheets(1) '���� ���� ������ � ������� �����
        cod = src.Cells(1, 1)
        If cod <> "" Then
        
            '������� ���������� ������ � ��������
            Dim i As Long
            i = FirstC
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
                
                If src.Cells(i, 1) = "" Then
                    '������ ���
                    For j = 2 To 14
                        dat.Cells(max, j) = src.Cells(i, j)
                    Next
                    dat.Cells(max, cFile) = file
                    dat.Cells(max, cCode) = cod
                    Range(dat.Cells(max, cFile), dat.Cells(max, cCode)).Font.Color = RGB(192, 192, 192)
                    
                    errors = Verify.Verify(dat, src, max, i)
                    If errors Then
                        AddFile = 2
                    Else
                        '���� ��� ������, ����������� �����
                        num = Numerator.Generate(dat.Cells(max, 2), dat.Cells(max, 4))
                        dat.Cells(max, 1) = num
                        src.Cells(i, 1) = num
                    End If
                    max = max + 1
                Else
                    '������ ����
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