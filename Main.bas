Attribute VB_Name = "Main"
Public Const isRelease = True   'True - ����������� ������, False - ����� ������� (��� ��������, ��� ������ � �����)
Public Const saveSource = True  'True - ���������� ������ � ������, False - ������ �� ������������ (�������)

Public Const Secret = "123"     '������ ��� ������

Public Const maxRow = 1048576   '��������� ������ ����� (��� �������)
Public Const tmpVersion = "20210108"    '������ �������

'������� "������"
Public Const cDates = 2         '����
Public Const cBuyINN = 3        '��� ����������
Public Const cBuyer = 4
Public Const cSellINN = 5       '��� ��������
Public Const cSeller = 6        '��������
Public Const cPrice = 7         '��������� � ���
Public Const cCom = 15          '�����������
Public Const cStatus = 16       '������
Public Const cFile = 17         '��� �����
Public Const cCode = 18         '��� �����
Public Const cAccept = 19       '�������/�� �������

'������� "����������"
Public Const cSellerName = 1    '������������ ��������
Public Const cINN = 2           '���
Public Const cSDate = 3         '���� �����������
Public Const cGroup = 4         '������
Public Const cLimits = 5        '������
Public Const cPLiter = 6        '������� - �����
Public Const cPCode = 7         '������� - ���

'������ ������
Public Const firstDat = 8       '������ ������ � ��������� ������
Public Const firstSrc = 5       '������ ������ � �������� ������
Public Const firstTempl = 7     '������ ������ � ������ �������
Public Const firstDic = 4       '������ ������ � �����������
Public Const firstErr = 2       '������ ������ � ������ ������
Public Const firstNum = 4       '������ ������ � ������� ����������

'�����
Public colWhite As Long
Public colRed As Long
Public colGreen As Long
Public colYellow As Long
Public colGray As Long
Public colBlue As Long

'������ �� �������
Public DAT As Variant   '������
Public SRC As Variant   '�������� ������
Public DIC As Variant   '�����������
Public ERR As Variant   '������ ������
Public NUM As Variant   '������� ����������
Public VAL As Variant   '�������� �������

'����� ���������� � �������
Sub ButtonDirSelect()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Cells(1, 3) = diag.SelectedItems(1)
End Sub

'����� ���������� ��� ��������
Sub ButtonDirSelectExport()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Cells(2, 3) = diag.SelectedItems(1)
End Sub

'������ "����� ������"
Sub ButtonSellBook()
    file = Application.GetOpenFilename("����� Excel(*.xls*),*.xls*", 1, _
        "�������� ���� �������", , False)
    If VarType(file) = vbBoolean Then Exit Sub
    ExportBook ByVal CStr(file)
End Sub

'������ "������� � 1�"
Sub ButtonExport()
    Init
    FormExport.Show
End Sub

'������ "�������� ������"
Sub ButtonClear()
    Init
    If isRelease Then
        e = Chr(10)
        If InputBox("��������! " + e + e + _
            "������ ��������� ������� ��� ��������� ������. " + _
            "��� ������������������ ������ ��� ��������� ����������� ����� ��������� ������ ���. " + _
            "���������� � ������� ���������� ��������� �� �����." + e + e + _
            "��� ����������� ������� ������.", "�������� ������") <> Secret Then Exit Sub
    End If
    SetProtect DAT
    Range(Cells(firstDat, 1), Cells(maxRow, cAccept)).Clear
    Range(Cells(firstDat, cStatus), Cells(maxRow, cStatus)).Interior.Color = colYellow
    Range(Cells(firstDat, cFile), Cells(maxRow, cAccept)).Interior.Color = colGray
    Range(Cells(firstDat, cFile), Cells(maxRow, cAccept)).Font.Color = RGB(166, 166, 166)
    Message "������!"
End Sub

'������ "���� ������"
Sub ButtonDataCollect()
    Init
    If isRelease Then If MsgBox("���������� ���� ������. ����������?", vbYesNo) = vbNo Then Exit Sub
    Message "����������..."
    SetProtect DAT
    Collect.Run
End Sub

'������ "������������ �������"
Sub ButtonCreateTemplates()
    Init
    Template.Generate
End Sub

'������������� ������, ������
Sub Init()
    colWhite = RGB(255, 255, 255)
    colRed = RGB(255, 192, 192)
    colGreen = RGB(192, 255, 192)
    colYellow = RGB(255, 255, 192)
    colGray = RGB(217, 217, 217)
    colBlue = RGB(192, 217, 255)
    
    If isRelease Then On Error GoTo er
    Set DAT = Sheets("������")
    Set DIC = Sheets("����������")
    Set ERR = Sheets("������")
    Set NUM = Sheets("������� ����������")
    Set VAL = Sheets("������")
    
    Exit Sub
er:
    MsgBox ("������ ����������� ���������!")
    End
End Sub

'��������� ������
Sub SetProtect(table As Variant)
    table.Protect Secret, AllowFormattingColumns:=True, UserInterfaceOnly:=True, AllowFiltering:=True
End Sub