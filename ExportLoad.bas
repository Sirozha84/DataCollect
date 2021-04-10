Attribute VB_Name = "ExportLoad"
'Last change: 10.04.2021 20:35

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
    i = 1
    Cells(i, 1) = "Код вида" + Chr(10) + "операции"
    Cells(i, 2) = "№ счет" + Chr(10) + "фактуры"
    Cells(i, 3) = "Дата счет" + Chr(10) + "фактуры"
    Cells(i, 4) = "ИНН"
    Cells(i, 5) = "КПП"
    Cells(i, 6) = "Наименование"
    Cells(i, 7) = "Сумма в руб." + Chr(10) + "и коп."
    Cells(i, 8) = "Сумма НДС" '+ Chr(10) + "без НДС 20%"
    'Cells(i, 9) = "Сумма" + Chr(10) + "без НДС 18%"
    'Cells(i, 10) = "Сумма" + Chr(10) + "без НДС 10%"
    'Cells(i, 11) = "НДС 20%"
    'Cells(i, 12) = "НДС 18%"
    'Cells(i, 13) = "НДС 10%"
    Cells(i, 14) = "Период НД"
    Columns(1).ColumnWidth = 10
    Columns(2).ColumnWidth = 13
    Columns(3).ColumnWidth = 10
    Columns(4).ColumnWidth = 11
    Columns(5).ColumnWidth = 10
    Columns(6).ColumnWidth = 15
    Columns(7).ColumnWidth = 12
    Columns(8).ColumnWidth = 12
    Columns(9).ColumnWidth = 12
    Columns(10).ColumnWidth = 12
    Columns(11).ColumnWidth = 10
    Columns(12).ColumnWidth = 10
    Columns(13).ColumnWidth = 10
    Columns(14).ColumnWidth = 10
    Rows(1).RowHeight = 30
    Set hat = Range(Cells(1, 1), Cells(1, 14))
    hat.HorizontalAlignment = xlCenter
    hat.VerticalAlignment = xlCenter
    hat.Interior.Color = colGray
    hat.Borders.Weight = 3
    firstEx = i + 1
    
    'Заполняем книгу
    i = firstDat
    j = firstEx
    Do While DTL.Cells(i, clAccept) <> ""
        If DTL.Cells(i, clAccept) = "OK" Then
            If DTL.Cells(i, clSaleINN).text = INN Then
                'Копирование данных из сбора
                Cells(j, 1).NumberFormat = "@"
                Cells(j, 1) = "01"
                Cells(j, 2) = DTL.Cells(i, 1)
                Cells(j, 3).NumberFormat = "dd.MM.yyyy"
                Cells(j, 3) = DTL.Cells(i, clDate)
                innkpp = Split(DAT.Cells(i, 3), "/")
                Cells(j, 4).NumberFormat = "@"
                Cells(j, 4) = DTL.Cells(i, clSaleINN)
                Cells(j, 6) = DTL.Cells(i, clSaleName)
                Cells(j, 7).NumberFormat = "### ### ##0.00"
                Cells(j, 7) = DTL.Cells(i, clPrice)
                Cells(j, 8).NumberFormat = "### ### ##0.00"
                Cells(j, 8) = DTL.Cells(i, clNDS)
                Cells(j, 14) = LastDateOfQuartal(DTL.Cells(i, clPND).text)
                j = j + 1
            End If
        End If
        i = i + 1
    Loop
    
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
                
                'По собранному списку поступлений проставляем текущий период и распределённую сумму
                For Each i In ndlist
                    DTL.Cells(i, clPND) = IndexToQYYYY(cPer)
                    If DTL.Cells(i, clRasp).text = "" Then _
                            DTL.Cells(i, clRasp) = DTL.Cells(i, clNDS)
                Next
                
            End If
            
        Next
        
    Next
    
End Sub

'Расчёт суммы отгрузок продавца с INN за квартал Q
Function GetSaleSumm(ByVal INN As String, ByVal q As Integer) As Double
    i = firstDat
    Sum = 0
    Do While DAT.Cells(i, cAccept) <> ""
        If DAT.Cells(i, cAccept) = "OK" And DAT.Cells(i, cSellINN) = INN Then
            If StupidQToQIndex(DAT.Cells(i, cPND)) = q Then
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
        q = DateToQIndex(d)
        If q >= cPer And q <= cPer + 11 And DTL.Cells(i, clPND) = "" Then dateS(d) = i
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

'Вычисление последней даты квартала
Function LastDateOfQuartal(ByVal q) As String
    n = Left(q, 1)
    y = Right(q, 4)
    If n = 1 Then LastDateOfQuartal = "31.03." + y
    If n = 2 Then LastDateOfQuartal = "30.06." + y
    If n = 3 Then LastDateOfQuartal = "30.09." + y
    If n = 4 Then LastDateOfQuartal = "31.12." + y
End Function

'******************** End of File ********************