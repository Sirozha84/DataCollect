VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormExport 
   Caption         =   "Выгрузка данных"
   ClientHeight    =   4914
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   9114.001
   OleObjectBlob   =   "FormExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Последняя правка: 04.07.2021 20:13

'Инициализация диалогового окна
Private Sub UserForm_Initialize()
    
    Dictionary.Init
    
    'Cписок продавцов
    For Each seller In selIndexes
        ListSalersAll.AddItem SellFileName(seller)
    Next
    ReSort
        
    'Период сбора
    TextBoxFirstCollect = PRP.Cells(8, 2)
    TextBoxLastCollect = PRP.Cells(9, 2)
    
End Sub

'Добавление продавцов

Private Sub ListSalersAll_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommandAdd_Click
End Sub

Private Sub CommandAdd_Click()
    If ListSalersAll.ListIndex < 0 Then Exit Sub
    ListSalersSelected.AddItem ListSalersAll.List(ListSalersAll.ListIndex)
    ListSalersAll.RemoveItem (ListSalersAll.ListIndex)
    ReSort
End Sub

Private Sub CommandAddAll_Click()
    For i = 1 To ListSalersAll.ListCount
        ListSalersSelected.AddItem ListSalersAll.List(0)
        ListSalersAll.RemoveItem (0)
    Next
    ReSort
End Sub

'Удаление продавцов

Private Sub ListSalersSelected_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommandRemove_Click
End Sub

Private Sub CommandRemove_Click()
    If ListSalersSelected.ListIndex < 0 Then Exit Sub
    ListSalersAll.AddItem ListSalersSelected.List(ListSalersSelected.ListIndex)
    ListSalersSelected.RemoveItem (ListSalersSelected.ListIndex)
    ReSort
End Sub

Private Sub CommandRemoveAll_Click()
    For i = 1 To ListSalersSelected.ListCount
        ListSalersAll.AddItem ListSalersSelected.List(0)
        ListSalersSelected.RemoveItem (0)
    Next
    ReSort
End Sub

'Сортировка листбокса
Sub ReSort()
    For i = 0 To ListSalersAll.ListCount - 2
        For i2 = i To ListSalersAll.ListCount - 1
            If StrComp(ListSalersAll.List(i), ListSalersAll.List(i2)) > 0 Then
                temp = ListSalersAll.List(i)
                ListSalersAll.List(i) = ListSalersAll.List(i2)
                ListSalersAll.List(i2) = temp
            End If
        Next
    Next
    For i = 0 To ListSalersSelected.ListCount - 2
        For i2 = i To ListSalersSelected.ListCount - 1
            If StrComp(ListSalersSelected.List(i), ListSalersSelected.List(i2)) > 0 Then
                temp = ListSalersSelected.List(i)
                ListSalersSelected.List(i) = ListSalersSelected.List(i2)
                ListSalersSelected.List(i2) = temp
            End If
        Next
    Next
End Sub

'Кнопка "Экспорт"
Private Sub CommandExport_Click()
    
    On Error GoTo er
    FirstDate = CDate(TextBoxFirstCollect)
    LastDate = CDate(TextBoxLastCollect)
    On Error GoTo 0
    
    If ComboBoxBuyers.ListIndex = 0 Then
        n = 1
        a = selIndexes.Count
        For Each seller In selIndexes
            ExportSale.Run seller, CStr(n) + " из " + CStr(a) + ": ", FirstDate, LastDate
            n = n + 1
        Next
    Else
        ExportSale.Run Left(ComboBoxBuyers.Value, 10), "", FirstDate, LastDate
    End If
    
    'Сохранение дат периода сбора
    PRP.Cells(8, 2) = TextBoxFirstCollect
    PRP.Cells(9, 2) = TextBoxLastCollect
    
    Message "Готово!"
    End

er:
    MsgBox "Даты не введены или введены не корректно"

End Sub

'Кнопка выхода
Private Sub CommandExit_Click()
    End
End Sub

'******************** End of File ********************