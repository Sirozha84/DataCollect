VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProperties 
   Caption         =   "Настройки"
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
'Last change: 23.04.2021 18:06

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

'******************** Фрейм "Пути" ********************

'Кнопка обзора пути импорта отгрузок
Private Sub ButtonExploreImportSale_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxImportSale = diag.SelectedItems(1)
End Sub

'Кнопка обзора пути импорта поступлений
Private Sub CommandImportLoad_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxImportLoad = diag.SelectedItems(1)
End Sub

'Кнопка обзора пути экспорта в 1С
Private Sub ButtonExport_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxExport = diag.SelectedItems(1)
End Sub

'******************** Фрейм "Период" ********************

'Кнопка "Переход к следующему кварталу"
Private Sub CommandButtonNext_Click()
    
    Application.ScreenUpdating = False
    
    'Проверка "чистоты" удаляемого квартала
    cl = True
    i = firstDic
    Do While DIC.Cells(i, 2) <> ""
        If cl Then
            If DIC.Cells(i, cPFact + quartCount - 1) <> "" Then cl = False
            If DIC.Cells(i, cPBalance + quartCount * 2 - 1) <> "" Then cl = False
            If DIC.Cells(i, cPBalance + quartCount * 2 - 2) <> "" Then cl = False
            If DIC.Cells(i, cCorrect + quartCount - 1) <> "" Then cl = False
        End If
        i = i + 1
    Loop
    If Not cl Then
        If MsgBox("Обнаружены данные за квартал " + LabelFD.Caption + ", которые будут удалены. Продолжить?", _
            vbYesNo) = vbNo Then Exit Sub
    End If
    
    'Сдвиг колонок вправо
    i = i - 1
    MoveColumnsRight firstDic, i, cPFact, 1
    MoveColumnsRight firstDic, i, cPBalance, 2
    MoveColumnsRight firstDic, i, cCorrect, 1
    MoveColumnsRight firstDic, i, cPRev, 1
    MoveColumnsRight firstDic, i, cSaleProtect, 1, True
    
    'Сдвиг периода в настройках
    lastQuartal = lastQuartal + 1
    If lastQuartal > 4 Then lastQuartal = 1: lastYear = lastYear + 1
    QHeadCreate
    RefreshPeriod
    
    Application.ScreenUpdating = True
    
End Sub

'Кнопка "Возврат к предыдущему кварталу"
Private Sub CommandButton1_Click()
    
    Application.ScreenUpdating = False
    
    'Проверка "чистоты" удаляемого квартала
    cl = True
    i = firstDic
    Do While DIC.Cells(i, 2) <> ""
        If cl Then
            If DIC.Cells(i, cPFact) <> "" Then cl = False
            If DIC.Cells(i, cPBalance) <> "" Then cl = False
            If DIC.Cells(i, cPBalance + 1) <> "" Then cl = False
            If DIC.Cells(i, cCorrect) <> "" Then cl = False
        End If
        i = i + 1
    Loop
    If Not cl Then
        If MsgBox("Обнаружены данные за квартал " + LabelLD.Caption + ", которые будут удалены. Продолжить?", _
            vbYesNo) = vbNo Then Exit Sub
    End If
    
    'Сдвиг колонок влево
    i = i - 1
    MoveColumnsLeft firstDic, i, cPFact, 1
    MoveColumnsLeft firstDic, i, cPBalance, 2
    MoveColumnsLeft firstDic, i, cCorrect, 1
    MoveColumnsLeft firstDic, i, cPRev, 1
    MoveColumnsLeft firstDic, i, cSaleProtect, 1, True
    
    'Сдвиг периода в настройках
    lastQuartal = lastQuartal - 1
    If lastQuartal < 1 Then lastQuartal = 4: lastYear = lastYear - 1
    QHeadCreate
    RefreshPeriod
    
    Application.ScreenUpdating = True
    
End Sub

'Смещение колонок вправо
Sub MoveColumnsRight(ByVal i1 As Long, ByVal i2 As Long, ByVal c, ByVal m, Optional gray As Boolean)
    Range(DIC.Cells(i1, c), DIC.Cells(i2, c + (quartCount - 1) * m - 1)).Select
    Selection.Copy
    DIC.Cells(i1, c + m).Select
    ActiveSheet.Paste
    Range(DIC.Cells(i1, c), DIC.Cells(i2, c + m - 1)).Clear
    If gray Then Range(DIC.Cells(i1, c), DIC.Cells(i2, c + m - 1)).Interior.Color = colGray
End Sub

'Смещение колонок влево
Sub MoveColumnsLeft(ByVal i1 As Long, ByVal i2 As Long, ByVal c, ByVal m, Optional gray As Boolean)
    Range(DIC.Cells(i1, c + m), DIC.Cells(i2, c + quartCount * m - 1)).Select
    Selection.Copy
    DIC.Cells(i1, c).Select
    ActiveSheet.Paste
    Range(DIC.Cells(i1, c + quartCount * m - m), DIC.Cells(i2, c + quartCount * m - 1)).Clear
    If gray Then Range(DIC.Cells(i1, c + quartCount * m - m), _
            DIC.Cells(i2, c + quartCount * m - 1)).Interior.Color = colGray
End Sub

'Обновление шапки колонок с кварталами
Sub QHeadCreate()
    For i = 0 To quartCount - 1
        DIC.Cells(3, cLimits + i) = IndexToQYYYY(i)
        DIC.Cells(3, cPFact + i) = IndexToQYYYY(i)
        DIC.Cells(3, cPBalance + i * 2) = IndexToQYYYY(i)
        DIC.Cells(3, cPBalance + i * 2 + 1) = IndexToQYYYY(i)
        DIC.Cells(3, cCorrect + i) = IndexToQYYYY(i)
        DIC.Cells(3, cPRev + i) = IndexToQYYYY(i)
        DIC.Cells(3, cSaleProtect + i) = IndexToQYYYY(i)
    Next
End Sub

'Кнопка "OK"
Private Sub CommandOK_Click()
    PRP.Cells(pImportSale, 2) = TextBoxImportSale
    PRP.Cells(pImportLoad, 2) = TextBoxImportLoad
    PRP.Cells(pExport, 2) = TextBoxExport
    End
End Sub

'Кнопка "Отмена"
Private Sub CommandCancel_Click()
    End
End Sub

'******************** End of File ********************