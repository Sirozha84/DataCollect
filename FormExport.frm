VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormExport 
   Caption         =   "Выгрузка данных"
   ClientHeight    =   2520
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4557
   OleObjectBlob   =   "FormExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Инициализация диалогового окна
Private Sub UserForm_Initialize()
    
    Dictionary.Init
    
    'Выпадающий список продавцов
    ComboBoxBuyers.AddItem "Все"
    For Each seller In selIndexes
        ComboBoxBuyers.AddItem SellFileName(seller)
    Next
    ComboBoxBuyers.ListIndex = 0
        
    'Период сбора
    TextBoxFirstCollect = PRP.Cells(8, 2)
    TextBoxLastCollect = PRP.Cells(9, 2)
    
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