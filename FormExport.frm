VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormExport 
   Caption         =   "Выгрузка данных"
   ClientHeight    =   3241
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4564
   OleObjectBlob   =   "FormExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Инициализация выпадающих списков
Private Sub UserForm_Initialize()
    
    Verify.Init
    
    ComboBoxBuyers.AddItem "Все"
    For Each seller In selIndexes
        ComboBoxBuyers.AddItem SellFileName(seller)
    Next
    ComboBoxBuyers.ListIndex = 0
    
    For y = lastYear To lastYear - Int(quartCount / 4) + 1 Step -1
        For m = 12 To 1 Step -1
            ComboBoxMonths.AddItem _
                YearAndMonth("01." + Right(CStr(m + 100), 2) + "." + CStr(y))
        Next
    Next
    ComboBoxMonths.ListIndex = 0
    
    For y = lastYear To lastYear - Int(quartCount / 4) + 1 Step -1
        For m = 10 To 1 Step -3
            ComboBoxQuartals.AddItem _
                YearAndQuartal("01." + Right(CStr(m + 100), 2) + "." + CStr(y))
        Next
    Next
    ComboBoxQuartals.ListIndex = 0
    
End Sub

'Имя файла по ИНН продавца
Function SellFileName(INN) As String
    ind = selIndexes(INN)
    If ind <> Empty Then SellFileName = INN + "-" + DIC.Cells(ind, 1)
End Function

Function YearAndMonth(ByVal d As String) As String
    On Error GoTo er:
    YearAndMonth = CStr(Year(d)) + " - "
    dy = Month(d)
    If dy = 1 Then YearAndMonth = YearAndMonth + "Январь"
    If dy = 2 Then YearAndMonth = YearAndMonth + "Февраль"
    If dy = 3 Then YearAndMonth = YearAndMonth + "Март"
    If dy = 4 Then YearAndMonth = YearAndMonth + "Апрель"
    If dy = 5 Then YearAndMonth = YearAndMonth + "Май"
    If dy = 6 Then YearAndMonth = YearAndMonth + "Июнь"
    If dy = 7 Then YearAndMonth = YearAndMonth + "Июль"
    If dy = 8 Then YearAndMonth = YearAndMonth + "Август"
    If dy = 9 Then YearAndMonth = YearAndMonth + "Сентябрь"
    If dy = 10 Then YearAndMonth = YearAndMonth + "Октябрь"
    If dy = 11 Then YearAndMonth = YearAndMonth + "Ноябрь"
    If dy = 12 Then YearAndMonth = YearAndMonth + "Декабрь"
    Exit Function
er:
    YearAndMonth = ""
End Function

Function YearAndQuartal(ByVal d As String) As String
    On Error GoTo er
    YearAndQuartal = CStr(Year(d)) + " - " + CStr((Month(d) - 1) \ 3 + 1) + " квартал"
    Exit Function
er:
    YearAndQuartal = ""
End Function

Private Sub OptionAll_Click()
    ComboBoxMonths.Enabled = False
    ComboBoxQuartals.Enabled = False
End Sub

Private Sub OptionMonth_Click()
    ComboBoxMonths.Enabled = True
    ComboBoxQuartals.Enabled = False
End Sub

Private Sub OptionQuartal_Click()
    ComboBoxMonths.Enabled = False
    ComboBoxQuartals.Enabled = True
End Sub

Private Sub CommandExit_Click()
    End
End Sub

'Кнопка "Экспорт"
Private Sub CommandExport_Click()
    If ComboBoxBuyers.ListIndex = 0 Then
        n = 1
        a = selIndexes.Count
        For Each seller In selIndexes
            ExportFile seller, CStr(n) + " из " + CStr(a) + ": "
            n = n + 1
        Next
    Else
        ExportFile Left(ComboBoxBuyers.Value, 10), "" ' ComboBoxBuyers.Value, ""
    End If
    Message "Готово!"
    End
End Sub

'Экспорт файла
Private Sub ExportFile(ByVal INN As String, NUM As String)
    
    seller = SellFileName(INN)
    Message "Экспорт файла " + NUM + seller
    
    'Определяемся с путём и именем файла
    Patch = DAT.Cells(2, 3)
    fol = ""
    mnC = OptionMonth.Value
    mn = ComboBoxMonths.Value
    qrC = OptionQuartal.Value
    qr = ComboBoxQuartals.Value
    If mnC Then fol = "\" + mn
    If qrC Then fol = "\" + qr
    If fol <> "" Then folder (Patch + fol)
    fileName = Patch + fol + "\" + cutBadSymbols(seller) + ".xlsx"
    
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
    Cells(i, 8) = "Сумма" + Chr(10) + "без НДС 20%"
    Cells(i, 9) = "Сумма" + Chr(10) + "без НДС 18%"
    Cells(i, 10) = "Сумма" + Chr(10) + "без НДС 10%"
    Cells(i, 11) = "НДС 20%"
    Cells(i, 12) = "НДС 18%"
    Cells(i, 13) = "НДС 10%"
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
    Do While DAT.Cells(i, cAccept) <> ""
        If DAT.Cells(i, cAccept) = "OK" Then
            cp = True
            If DAT.Cells(i, cSellINN).text <> INN Then cp = False
            d = DAT.Cells(i, cDates)
            If mnC Then If YearAndMonth(d) <> mn Then cp = False
            If qrC Then If YearAndQuartal(d) <> qr Then cp = False
            If cp Then
                Cells(j, 1).NumberFormat = "@"
                Cells(j, 1) = "01"
                Cells(j, 2) = DAT.Cells(i, 1)
                Cells(j, 3).NumberFormat = "dd.MM.yyyy"
                Cells(j, 3) = DAT.Cells(i, 2)
                innkpp = Split(DAT.Cells(i, 3), "/")
                Cells(j, 4).NumberFormat = "@"
                Cells(j, 4) = innkpp(0)
                Cells(j, 5).NumberFormat = "@"
                If UBound(innkpp) > 0 Then Cells(j, 5) = innkpp(1)
                Cells(j, 6) = DAT.Cells(i, 4)
                Cells(j, 7).NumberFormat = "### ### ##0.00"
                Cells(j, 7) = DAT.Cells(i, 7)
                For c = 0 To 5
                    Cells(j, 8 + c).NumberFormat = "### ### ##0.00"
                    Cells(j, 8 + c) = DAT.Cells(i, 9 + c)
                Next
                Cells(j, 15) = YearAndQuartal(DAT.Cells(i, 2))
                j = j + 1
            End If
        End If
        i = i + 1
    Loop
    
    'Сортировка и удаление временного столбца
    Cells(1, 15) = "Квартал"
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("O2")
        .SortFields.Add Key:=Range("F2")
        .setRange Range("A2:O" + CStr(j - 1))
        .Apply
    End With
    Columns(15).Delete
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