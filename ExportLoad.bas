Attribute VB_Name = "ExportLoad"
Sub Run()
    
    Message "Подготовка..."
    Dictionary.Init
    
    'Подготовка каталога для экспорта
    Patch = DirExport + "\Поступления"
    MakeDir Patch
    Set files = Source.getFiles(DirExport + "\Поступления", False)
    For Each file In files
        Kill file
    Next
    
    'Формируем список инн продавцов для выгрузки
    Dim SalersINN As Collection
    Set SalersINN = New Collection
    Set files = Source.getFiles(DirExport + "\Отгрузки", False)
    
    n = 1
    a = selIndexes.Count
    For Each file In files
        CreateExportFile Left(Source.FSO.GetFileName(file), 10), CStr(n) + " из " + CStr(a) + ": "
    Next
    
    Message "Готово!"
    
End Sub

'Формирование файла выгрузки
Sub CreateExportFile(ByVal inn As String, ByVal NUM As String)
    
    saler = SellFileName(inn)
    Message "Экспорт файла " + NUM + saler
    
    'Определяемся с путём и именем файла
    Patch = DirExport + "\Поступления"
    MakeDir Patch
    fileName = Patch + "\" + cutBadSymbols(saler) + ".xlsx"
    
    'Создаём книгу
    Workbooks.Add
    
    
    
    
    End
    
    'Сохранение и закрытие документа
    On Error GoTo er
    Application.DisplayAlerts = False
    If j > firstEx Then ActiveWorkbook.SaveAs fileName:=fileName
    ActiveWorkbook.Close
    Exit Sub

er:
    ActiveWorkbook.Close
    MsgBox "Произошла ошибка при сохранении файла " + fileName

End Sub

'******************** End of File ********************