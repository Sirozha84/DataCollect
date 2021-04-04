Attribute VB_Name = "ExportLoad"
'Last change: 04.04.2021 18:47

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
Sub CreateExportFile(ByVal INN As String, ByVal NUM As String)
    
    saler = SellFileName(INN)
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
    
    For Each INN In selIndexes
        
        Si = selIndexes(INN)
        oND = StupidQToQIndex(DIC.Cells(Si, cOPND))
        
        For cPer = oND To quartCount - 1
            
            'Расчёт суммы всех отгрузок за текущий период
            Sum = GetSaleSumm(INN, cPer)
            If Sum > minSale Then
            
                Set dateS = GetDatesList(cPer)
                
                'Подбираем список поступлений
                Dim ndlist As Collection
                Set ndlist = New Collection
                For Each postdate In dateS
                    i = dateS(postdate)
                    Post = DTL.Cells(i, clNDS)
                    If Sum - Post >= 0 Then
                        Sum = Sum - Post
                        If Sum >= 0 Then ndlist.Add i
                        If Sum < maxDif Then Exit For
                    End If
                Next
                
                'По собранному списку поступлений проставляем текущий период
                For Each i In ndlist
                    DTL.Cells(i, clPND) = IndexToQYYYY(cPer)
                Next
                
            End If
            
        Next
        
    Next
    
    End
    
End Sub

'Расчёт суммы отгрузок продавца с INN за квартал Q
Function GetSaleSumm(ByVal INN As String, ByVal Q As Integer) As Double
    i = firstDat
    Sum = 0
    Do While DAT.Cells(i, cAccept) <> ""
        If DAT.Cells(i, cAccept) = "OK" And DAT.Cells(i, cSellINN) = INN Then
            If StupidQToQIndex(DAT.Cells(i, cPND)) = Q Then
                Sum = Sum + WorksheetFunction.Sum(Range(DAT.Cells(i, clNDS), DAT.Cells(i, clNDS + 2)))
            End If
        End If
        i = i + 1
    Loop
    GetSaleSumm = Sum
End Function

'Формирование сортированного списка дат входящих в 12 кварталов начиная от cPer
Function GetDatesList(ByVal cPer As Integer) As Object
            
    'Собираем подходящие даты
    Set dateS = CreateObject("Scripting.Dictionary")
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        Dim d As Date
        d = DTL.Cells(i, clDate)
        Q = DateToQIndex(d)
        If Q >= cPer And Q <= cPer + 11 And DTL.Cells(i, clPND) = "" Then dateS(d) = i
        i = i + 1
    Loop
    
    'Сортируем собранные даты
    Set datesSorted = CreateObject("Scripting.Dictionary")
    Do While dateS.Count > 0
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