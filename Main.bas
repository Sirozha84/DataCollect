Attribute VB_Name = "Main"
Const isRelease = False 'True - ����������� ������, False - ����� ������� (��� ��������, ��� ������ � �����)
Const FirstC = 6        '������ ������ � ��������� ������
Const FirstS = 5        '������ ������ � �������� ������
Const cFile = 17        '������� � ������ �����
Const cCode = 18        '������� � ����� �����
Const erTabName = "������"

Dim erTab As Variant

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
    Range(Cells(StartString, 1), Cells(1048576, 50)).Clear
    Sheets(erTabName).Cells.Clear
er:
    Numerator.Clear
End Sub

'���� ������
Sub DataCollect()
    If isRelease Then If MsgBox("���������� ���� ������. ����������?", vbYesNo) = vbNo Then Exit Sub
    
    '�������� ��������� ������
    Set Files = Source.GetList("C:\Users\SG\OneDrive\������\�������� ������\���� ������\������")
        
    '������ ������� (���� � ���) ��� ������ ������
    Call NewTab(erTabName, True)
    Set erTab = Sheets("������")
    erTab.Columns(1).ColumnWidth = 100
    erTab.Columns(2).ColumnWidth = 20
    erTab.Cells(1, 1) = "����"
    erTab.Cells(1, 2) = "���������"
    
    '�������������� ������� ����������
    Numerator.Init
    
    Dim str As Long
    str = FirstC
    n = 1
    s = 0
    e = 0
    Max = Files.Count
    For Each file In Files
        Message ("��������� ����� " + CStr(n) + " �� " + CStr(Files.Count) + " (" + Source.FSO.getfilename(file)) + ")"
        er = AddFile(file, str)
        If er > 0 Then
            e = e + 1
            erTab.Cells(e + 1, 1) = file
            If er = 1 Then erTab.Cells(e + 1, 2) = "������ �������� �����"
            If er = 2 Then erTab.Cells(e + 1, 2) = "������ � ������"
            If er = 3 Then erTab.Cells(e + 1, 2) = "����������� ���"
        Else
            s = s + 1
        End If
        n = n + 1
    Next
    Message ("������!")
    If isRelease Then MsgBox ("��������� ���������!" + Chr(13) + "������ ����������� �������: " + CStr(s) + Chr(13) + "����� � ��������: " + CStr(e))
    
End Sub

'���������� ������ �� ����� (���������� 0 - �� ������, 1 - ������ ��������, 2 - ������ � ������, 3 - ��� ����)
Function AddFile(ByVal file As String, ByRef str As Long) As Byte
    
    On Error GoTo er
    
    Application.ScreenUpdating = False
    Set cur = ActiveSheet
    Set imBook = Nothing
    Set imBook = Workbooks.Open(file, False, False)
    If Not imBook Is Nothing Then
        Set imSh = imBook.Worksheets(1) '���� ���� ������ � ������� �����
        cod = imSh.Cells(1, 1)
        If cod <> "" Then
        
            '������� ���������� ������ � ��������
            
            
            '������������ ������ ���������
            i = FirstString
            Do While imSh.Cells(i, 2) <> ""
                
                '�������� �������
                For j = 2 To 14
                    cur.Cells(str, j) = imSh.Cells(i, j)
                Next
                cur.Cells(str, cFile) = file
                cur.Cells(str, cCode) = cod
                Range(cur.Cells(str, cFile), cur.Cells(str, cSheet)).Font.Color = RGB(192, 192, 192)
                
                errors = Verify.Verify(cur, imSh, str, i)
                If errors Then
                    AddFile = 2
                Else
                    '���� ��� ������, ����������� �����
                    cur.Cells(str, 1) = Numerator.Generate(cur.Cells(str, 2), cur.Cells(str, 4))
                End If
                str = str + 1
                i = i + 1
            Loop
        Else
            AddFile = 3
        End If
        imBook.Close isRelease
    End If
    Numerator.Save
    Application.ScreenUpdating = True
    DoEvents
    Exit Function
er:
    AddFile = 1
End Function