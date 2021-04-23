VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProperties 
   Caption         =   "���������"
   ClientHeight    =   5761
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   7560
   OleObjectBlob   =   "FormProperties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Last change: 23.04.2021 16:02

Private Sub UserForm_Activate()
    TextBoxImportSale = PRP.Cells(pImportSale, 2).text
    TextBoxImportLoad = PRP.Cells(pImportLoad, 2).text
    TextBoxExport = PRP.Cells(pExport, 2).text
    RefreshPeriod
End Sub

Sub RefreshPeriod()
    LabelLD.Caption = CStr(lastQuartal) + CStr(lastYear)
    lq = lastQuartal
    ly = lastYear
    For i = 2 To quartCount
        lq = lq - 1
        If lq < 1 Then lq = 4: ly = ly - 1
    Next
    LabelFD.Caption = CStr(lq) + CStr(ly)
    PRP.Cells(pLastYear, 2) = lastYear
    PRP.Cells(pLastQuartal, 2) = lastQuartal
End Sub

'******************** ����� "����" ********************

'������ ������ ���� ������� ��������
Private Sub ButtonExploreImportSale_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxImportSale = diag.SelectedItems(1)
End Sub

'������ ������ ���� ������� �����������
Private Sub CommandImportLoad_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxImportLoad = diag.SelectedItems(1)
End Sub

'������ ������ ���� �������� � 1�
Private Sub ButtonExport_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxExport = diag.SelectedItems(1)
End Sub

'******************** ����� "������" ********************

'������ "������� � ���������� ��������"
Private Sub CommandButtonNext_Click()
    
    '�������� "�������" ���������� ��������
    cl = True
    i = firstDic
    Do While DIC.Cells(i, 2) <> ""
        If DIC.Cells(i, cPFact + quartCount - 1) <> "" Then cl = False: Exit Do
        If DIC.Cells(i, cPBalance + quartCount * 2 - 1) <> "" Then cl = False: Exit Do
        If DIC.Cells(i, cPBalance + quartCount * 2 - 2) <> "" Then cl = False: Exit Do
        If DIC.Cells(i, cCorrect + quartCount - 1) <> "" Then cl = False: Exit Do
        i = i + 1
    Loop
    If Not cl Then
        If MsgBox("���������� ������ �� ������� " + LabelFD.Caption + ", ������� ����� �������. ����������?", _
            vbYesNo) = vbNo Then Exit Sub
    End If
    
    '����� ������� ������
    
    
    
    lastQuartal = lastQuartal + 1
    If lastQuartal > 4 Then lastQuartal = 1: lastYear = lastYear + 1
    RefreshPeriod
End Sub

'������ "������� � ����������� ��������"
Private Sub CommandButton1_Click()
    
    '�������� "�������" ���������� ��������
    cl = True
    i = firstDic
    Do While DIC.Cells(i, 2) <> ""
        If DIC.Cells(i, cPFact) <> "" Then cl = False: Exit Do
        If DIC.Cells(i, cPBalance) <> "" Then cl = False: Exit Do
        If DIC.Cells(i, cPBalance + 1) <> "" Then cl = False: Exit Do
        If DIC.Cells(i, cCorrect) <> "" Then cl = False: Exit Do
        i = i + 1
    Loop
    If Not cl Then
        If MsgBox("���������� ������ �� ������� " + LabelLD.Caption + ", ������� ����� �������. ����������?", _
            vbYesNo) = vbNo Then Exit Sub
    End If
    
    '����� ������� �����
    
    
    
    lastQuartal = lastQuartal - 1
    If lastQuartal < 1 Then lastQuartal = 4: lastYear = lastYear - 1
    RefreshPeriod
End Sub


'������ "OK"
Private Sub CommandOK_Click()
    PRP.Cells(pImportSale, 2) = TextBoxImportSale
    PRP.Cells(pImportLoad, 2) = TextBoxImportLoad
    PRP.Cells(pExport, 2) = TextBoxExport
    End
End Sub

'������ "������"
Private Sub CommandCancel_Click()
    End
End Sub

'******************** End of File ********************