Attribute VB_Name = "Main"
Public Const isRelease = False  'True - ����������� ������, False - ����� ������� (��� ��������, ��� ������ � �����)
Public Const saveSource = True  'True - ���������� ������ � ������, False - ������ �� ������������ (�������)
Public Const maxRow = 1048576   '��������� ������ ����� (��� �������)
Public Const maxCol = 50        '��������� ������� ����� (��� �������)

'�������
Public Const cBuyer = 6         '��������
Public Const cCom = 15          '�����������
Public Const cStatus = 16       '������
Public Const cFile = 17         '��� �����
Public Const cCode = 18         '��� �����

'������ ������
Public Const firstDat = 8       '������ ������ � ��������� ������
Public Const firstSrc = 5       '������ ������ � �������� ������
Public Const firstTempl = 7     '������ ������ � ������ ��������
Public Const firstDic = 5       '������ ������ � �����������
Public Const firstErr = 2       '������ ������ � ������ ������
Public Const firstNum = 4       '������ ������ � ������� ����������

'�����
Public colWhite As Long
Public colRed As Long
Public colGreen As Long
Public colYellow As Long

'������ �� �������
Public DAT As Variant   '������
Public SRC As Variant   '�������� ������
Public DIC As Variant   '�����������
Public ERR As Variant   '������ ������
Public NUM As Variant   '������� ����������

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

    Message "�������� ������"
    Init
    'On Error GoTo er
    If isRelease Then If MsgBox("��������! " + Chr(10) + Chr(10) + _
        "������ ��������� ������� ��� ��������� ������ ������ ������ � ����������. " + _
        "��� ������������������ ������ ��� ��������� ����������� ����� ��������� ������ ���." + _
        Chr(10) + Chr(10) + "����������?", vbYesNo) = vbNo Then Exit Sub
    Range(DAT.Cells(firstDat, 1), DAT.Cells(maxRow, maxCol)).Clear
    Range(ERR.Cells(firstErr, 1), ERR.Cells(maxRow, maxCol)).Clear
    Range(NUM.Cells(firstNum, 1), NUM.Cells(maxRow, maxCol)).Clear
    Exit Sub
    
    Message "������!"
    
er:
    MsgBox ("������ ����������� ���������!")
End Sub

'���� ������
Sub DataCollect()
    
    If isRelease Then If MsgBox("���������� ���� ������. ����������?", vbYesNo) = vbNo Then Exit Sub
    
    Message "����������..."
    
    Init
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

'������������� ������, ������
Sub Init()
    
    'On Error GoTo er
    
    Set DAT = ActiveSheet
    Set DIC = Sheets("����������")
    Set ERR = Sheets("������")
    Set NUM = Sheets("������� ����������")
    
    colWhite = RGB(255, 255, 255)
    colRed = RGB(255, 192, 192)
    colGreen = RGB(192, 255, 192)
    colYellow = RGB(255, 255, 192)
    
    Exit Sub
er:
    MsgBox ("������ ����������� ���������!")
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
        Set SRC = impBook.Worksheets(1) '���� ���� ������ � ������� �����
        SRC.Unprotect Template.Secret
        cod = SRC.Cells(1, 1)
        If cod <> "" Then
            
            '������� ���������� ������ � ��������
            i = firstDat
            Do While DAT.Cells(i, 2) <> ""
                If DAT.Cells(i, 1) = "" And DAT.Cells(i, cCode) = cod Then
                    DAT.Rows(i).Delete
                    max = max - 1
                Else
                    i = i + 1
                End If
            Loop
        
            '����������� ������������ ������
            Set Indexes = CreateObject("Scripting.Dictionary")
            i = firstDat
            Do While DAT.Cells(i, 2) <> ""
                UID = DAT.Cells(i, 1)
                If UID <> "" Then Indexes.Add UID, i
                i = i + 1
            Loop
            max = i
        
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
                        If copyRecord(ind, i, True) Then errors = True
                        
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
                If UID = "" Then If copyRecord(max, i, False) Then errors = True
                rUID = SRC.Cells(i, 1).text
                If rUID <> "" Then resUIDs.Add rUID, 1
                i = i + 1
            Loop
            
            '��������� �������� �� �������� ������
            i = firstDat
            Do While DAT.Cells(i, 2) <> ""
                UID = DAT.Cells(i, 1)
                If UID <> "" And DAT.Cells(i, cCode) = cod Then
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
        SRC.Protect Template.Secret
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
Function NotEmpty(i As Long) As Boolean
    NotEmpty = False
    For j = 1 To 14
        txt = SRC.Cells(i, j).text
        If txt <> "" And txt <> "#�/�" Then NotEmpty = True: Exit For
    Next
End Function

'����������� ������. ���������� True, ���� � ������ ���� ������
'di - ������ � ������
'si - ������ � ����������
'refresh - true, ���� ���������� ������ (��������� ��� ����������)
Function copyRecord(ByVal di As Long, ByVal si As Long, refresh As Boolean) As Boolean
    
    stat = DAT.Cells(di, cStatus).text
    If stat = "0" Or stat = "2" Then Exit Function
    
    '����������� ������� � ��������� �� ���������
    Dim changed As Boolean
    For j = 2 To 14
        ravno = DAT.Cells(di, j).text = SRC.Cells(si, j).text
        DAT.Cells(di, j) = SRC.Cells(si, j)
        DAT.Cells(di, j).ClearFormats
        If j = 2 Or j = 4 Or j = 6 Or j = 7 Or j = 8 Then
            SRC.Cells(si, j).Interior.Color = colYellow
        Else
            SRC.Cells(si, j).Interior.Color = colWhite
        End If
        If refresh And Not ravno Then
            DAT.Cells(di, j).Interior.Color = colYellow
            SRC.Cells(si, j).Interior.Color = colYellow
            changed = True
        End If
    Next
    DAT.Cells(di, cFile) = file
    DAT.Cells(di, cCode) = cod
    Range(DAT.Cells(di, cFile), DAT.Cells(di, cCode)).Font.Color = RGB(192, 192, 192)
    errors = Verify.Verify(DAT, SRC, di, si, changed)
    
    '���� �����, ����������� ������ ����� �����
    If Not errors Then
        Dim needNum As Boolean
        If refresh Then
            needNum = Not Numerator.CheckPrefix(DAT.Cells(di, 1).text, _
                DAT.Cells(di, 2), DAT.Cells(di, cBuyer).text)
        Else
            needNum = True
        End If
        If needNum Then
            n = Numerator.Generate(DAT.Cells(di, 2), DAT.Cells(di, cBuyer).text)
            DAT.Cells(di, 1) = n
            SRC.Cells(si, 1) = n
        End If
    Else
        copyRecord = True
    End If
    
    If Not refresh Then max = max + 1
    If DAT.Cells(di, cStatus).text = "" Then DAT.Cells(di, cStatus) = 1
    
End Function