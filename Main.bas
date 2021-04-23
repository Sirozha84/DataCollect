Attribute VB_Name = "Main"
'Last change: 23.04.2021 14:09

'���������
Public Const maxRow = 1048576   '��������� ������ ����� (��� �������)
Public Const tmpVersion = "20210108"    '����������� ������ �������

'���������
Public Const Secret = "123"     '������ ��� ������
Public Const quartCount = 12    '���������� ��������� � �������� �������
Public Const lastYear = 2020    '������ ��������� ���
Public Const lastQuartal = 4    '������ ��������� �������
Public Const limitOND = 9000000 '����� � �������� ������ �� (9�)
Public Const minLim = 5000000   '����������� ����� ������, ���� ������, ������ ������������ (5�)
Public Const minSale = 20000    '����������� ����� ������, ����������� ��� ������������� ������� (20�)
Public Const maxDif = 15000     '����������� ������� ����� ��������� � ������������ (15�)

'������� "��������"
Public Const cUIN = 1           '���
Public Const cDates = 2         '����
Public Const cBuyINN = 3        '���������� ���
Public Const cBuyer = 4         '���������� ������������
Public Const cSellINN = 5       '�������� ���
Public Const cSeller = 6        '�������� ������������
Public Const cPrice = 7         '��������� � ���
Public Const cCom = 15          '�����������
Public Const cStatus = 16       '������
Public Const cDateCol = 17      '���� �����
Public Const cFile = 18         '��� �����
Public Const cCode = 19         '��� �����
Public Const cAccept = 20       '�������/�� �������
Public Const cPND = 21          '������ ��

'������� "�����������"
Public Const clMark = 1         '������
Public Const clKVO = 2          '��� ���� �������� (���)
Public Const clNum = 3          '�����
Public Const clDate = 4         '����
Public Const clProvINN = 5      '��������� ���
Public Const clProvName = 6     '��������� ������������
Public Const clSaleINN = 7      '�������� ���
Public Const clSaleName = 8     '�������� ������������
Public Const clPrice = 9        '��������� � ���
Public Const clNDS = 13         '����� ���
Public Const clCom = 14         '�����������
Public Const clStatus = 15      '������
Public Const clRasp = 16        '������������
Public Const clPND = 17         '������ ��
Public Const clOst = 18         '������� ���
Public Const clDateCol = 19     '���� �����
Public Const clUIN = 20         '���
Public Const clFile = 21        '��� �����
Public Const clAccept = 22      '�������/�� �������

'������� "����������"
Public Const cSellerName = 1    '������������ ��������
Public Const cINN = 2           '���
Public Const cSDate = 3         '���� �����������
Public Const cGroup = 4         '������
Public Const cLimND = 5         '����� �� �� �������
Public Const cPLiter = 6        '������� - �����
Public Const cPCode = 7         '������� - ���
Public Const cOPND = 8          '�������� ������ ��
Public Const cPStat = 9         '������
Public Const cLimits = 10       '������ ������� � ���������
Public Const cPFact = 22        '������ ������� � ������������ ��������
Public Const cPBalance = 34     '������ ������� � ��������� (*2)
Public Const cCorrect = 58      '������ ������� � ��������������� �������
Public Const cPRev = 70         '������ ������� � ������������ ���������� (��� ������� ��������)
Public Const cSaleProtect = 82  '������ ������� � ��������� ��������

'������� "�������"
Public Const cTClient = 1       '������
Public Const cTBroker = 2       '���������
Public Const cTForm = 3         '�����
Public Const cTCode = 4         '���
Public Const cTFile = 5         '����
Public Const cTResult = 6       '���������
Public Const cTStat = 7         '������

'������ ������
Public Const firstDat = 6       '��������
Public Const firstDtL = 6       '�����������
Public Const firstSrc = 5       '�������
Public Const firstTempl = 6     '������ ��������
Public Const firstDic = 4       '����������
Public Const firstErr = 2       '������
Public Const firstNum = 4       '������� ����������
Public Const firstValues = 6    '����� "������"

'������ � �����������
Public Const pImportSale = 4    '������ ��������
Public Const pImportLoad = 5    '������ �����������
Public Const pExport = 6        '�������

'�����
Public colWhite As Long         '��� ����������� �����
Public colRed As Long           '������
Public colGreen As Long         '��������
Public colYellow As Long        '���������� ��� ��������������
Public colGray As Long          '��������� ����
Public colBlue As Long          '�������� ���������

'������ �� �������
Public DAT As Variant           '������ � ��������
Public DTL As Variant           '������ � ������������
Public SRC As Variant           '�������� ������
Public DIC As Variant           '�����������
Public ERR As Variant           '������ ������
Public NUM As Variant           '������� ����������
Public VAL As Variant           '�������� �������
Public VLS As Variant           '������� �������
Public TMP As Variant           '�������
Public SBK As Variant           '����� ������
Public PRP As Variant           '���������

'���������
Public DirImportSale As String  '������� ������� ��������
Public DirImportLoad As String  '������� ������� �����������
Public DirExport As String      '������� ��������

'����� ����������
Public selIndexes As Variant    '������� �������� ��������� (������ ����� � ����������� �� ���)
Public BookCount As Long        '������� ��������������� ����

'������������� ������, ������
Sub Init()
    colWhite = RGB(255, 255, 255)
    colRed = RGB(255, 192, 192)
    colGreen = RGB(192, 255, 192)
    colYellow = RGB(255, 255, 192)
    colGray = RGB(217, 217, 217)
    colBlue = RGB(192, 217, 255)
    
    On Error GoTo er
    Set DAT = Sheets("��������")
    Set DTL = Sheets("�����������")
    Set DIC = Sheets("����������")
    Set VAL = Sheets("������")
    Set VLS = Sheets("������� �������")
    Set TMP = Sheets("�������")
    Set SBK = Sheets("����� ������")
    Set ERR = Sheets("������")
    Set NUM = Sheets("���������")
    Set PRP = Sheets("���������")
    
    DirImportSale = PRP.Cells(pImportSale, 2).text
    DirImportLoad = PRP.Cells(pImportLoad, 2).text
    DirExport = PRP.Cells(pExport, 2).text
    
    Exit Sub
er:
    MsgBox ("������ ����������� ���������! ����������� ������� ���� ������� ��� �������������.")
    End
End Sub

Sub ButtonProperties()
    Init
    FormProperties.Show
End Sub

'******************** ������� "��������" ********************

'������ "���� ������"
Sub ButtonDataCollect()
    Init
    If MsgBox("���������� ���� ������ �� ���������. ����������?", vbYesNo) = vbNo Then Exit Sub
    SetProtect DAT
    CollectSale.Run
    DAT.Activate
End Sub

'������ "������� � 1�"
Sub ButtonExport()
    Init
    FormExport.Show
End Sub

'������ "�������� ������"
Sub ButtonClear()
    Init
    e = Chr(10)
    If InputBox("��������! " + e + e + _
        "������ ��������� ������� ��� ��������� ������. " + _
        "��� ������������������ ������ ��� ��������� ����������� ����� ��������� ������ ���. " + _
        "���������� � ������� ���������� ��������� �� �����." + e + e + _
        "��� ����������� ������� ������.", "�������� ������") <> Secret Then Exit Sub
    SetProtect DAT
    Range(DAT.Cells(firstDat, 1), DAT.Cells(maxRow, cAccept)).Clear
    Range(DAT.Cells(firstDat, cStatus), DAT.Cells(maxRow, cDateCol)).Interior.Color = colYellow
    Range(DAT.Cells(firstDat, cFile), DAT.Cells(maxRow, cAccept)).Interior.Color = colGray
    Range(DAT.Cells(firstDat, cFile), DAT.Cells(maxRow, cAccept)).Font.Color = RGB(166, 166, 166)
    Range(DTL.Cells(firstDtL, 1), DTL.Cells(maxRow, clAccept)).Clear
    Range(DTL.Cells(firstDtL, clFile), DTL.Cells(maxRow, clAccept)).Interior.Color = colGray
    Range(DTL.Cells(firstDtL, clFile), DTL.Cells(maxRow, clAccept)).Font.Color = RGB(166, 166, 166)
    Range(DIC.Cells(firstDic, cPFact), DIC.Cells(maxRow, cPFact + quartCount * 6 - 1)).Clear
    Range(DIC.Cells(firstDic, cSaleProtect), DIC.Cells(maxRow, cSaleProtect + quartCount - 1)). _
            Interior.Color = colGray
    
    Message "������! ���� �� ��� �������. " + _
            "���� ���������� - �������� ���� �� ���������� � �������� �����."
End Sub

'******************** ������� "�����������" ********************

'������ "���� �����������"
Sub ButtonCollectLoad()
    Init
    If MsgBox("���������� ���� ������ �� ������������. " + _
            "����������?", vbYesNo) = vbNo Then Exit Sub
    CollectLoad.Run
End Sub

'������ "������� ����������� � 1�"
Sub ButtonExportLoad()
    Init
    If MsgBox("���������� ������� ������ � ������������. ����������?", vbYesNo) = vbNo Then Exit Sub
    ExportLoad.Run
End Sub

'******************** ������� "������" ********************

'������ ������� ��������
Sub ButtonRevisionVolumes()
    Init
    DIC.Activate
    DIC.Cells(firstDic, cPRev).Activate
    Revision.Run
End Sub

'������ "������������ �����" �� �������
Sub ButtonReportVolumes()
    Init
    Values.CreateReport
    VAL.Activate
End Sub

'******************** ������� "�������" ********************

'������ "������������ �������"
Sub ButtonCreateTemplates()
    Init
    TMP.Activate
    Template.Generate
End Sub

'******************** ������� "����� ������" ********************

'������ "������������"
Public Sub ButtonSellBook()
    Init
    SellBook.Run
End Sub

'******************** End of File ********************