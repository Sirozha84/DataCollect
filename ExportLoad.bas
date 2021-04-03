Attribute VB_Name = "ExportLoad"
'Последняя правка: 03.04.2021 21:36

Sub Run()
    
    Message "Подготовка..."
    Dictionary.Init
    LoadAllocation
    
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
    
    
    
    'Тут будет формирование строк данных
    
    
    
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

'Распределение поступлений
Sub LoadAllocation()
    
    For Each inn In selIndexes
        
        Si = selIndexes(inn)
        oND = StupidQToQIndex(DIC.Cells(Si, cOPND))
        
        For cPer = oND To Quartals
            
            'Расчёт суммы всех отгрузок за текущий период
            Sum = 0
            i = firstDat
            Do While DAT.Cells(i, cAccept) <> ""
                If DAT.Cells(i, cAccept) = "OK" And DAT.Cells(i, cSellINN) = inn Then
                    If StupidQToQIndex(DAT.Cells(i, cPND)) = cPer Then
                        Sum = Sum + WorksheetFunction.Sum(Range(DAT.Cells(i, clNDS), DAT.Cells(i, clNDS + 2)))
                    End If
                End If
                i = i + 1
            Loop
            If Sum > minSale Then
            
                Set dateS = GetDatesList(cPer)

                'Далее делаем подбор начиная с первой строки в dataS

            
            End If
            
        Next
        
    Next
    
    End
    
End Sub

'Формирование сортированного списка дат входящих в 12 кварталов начиная от заданного
Function GetDatesList(ByVal cPer As Integer) As Object
            
    'Собираем подходящие даты
    Set dateS = CreateObject("Scripting.Dictionary")
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        Dim d As Date
        d = DTL.Cells(i, clDate)
        q = DateToQIndex(d)
        If q >= cPer And q <= cPer + 11 Then dateS(d) = i
        i = i + 1
    Loop
    
    'Сортируем собранные даты
    Set datesSorted = CreateObject("Scripting.Dictionary")
    Do While dateS.Count > 0
        'Находим максимальную дату
        Dim max As Date
        max = 0
        For Each dt In dateS
            If max < dt Then max = dt
        Next
        datesSorted(max) = dateS(max)
        dateS.Remove (max)
    Loop
    Set GetDatesList = datesSorted
End Function

'******************** End of File ********************