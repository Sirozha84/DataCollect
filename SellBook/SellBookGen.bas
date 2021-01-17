Attribute VB_Name = "SellBookGen"
Const firstDat = 5  'Первая строка данных

'Колонки реестра
Const cDates = 2    'Дата
Const cBuyINN = 3   'Покупатель ИНН
Const cBuyer = 4    'Покупатель Наименование
Const cSellINN = 5  'Продавец ИНН
Const cSeller = 6   'Продавец Наименование
Const cPrice = 7    'Стоимость
Const cCom = 15     'Комментарий

'Положительные комментарии
Const OK1 = "Принято"
Const OK2 = "Данные зафиксированы!"

Dim DAT As Object
Dim BUY As Object
Dim SEL As Object
Dim BuyersList As Variant
Dim SellersList As Variant
Dim Quartals As Variant
Dim Where As Collection
Dim Patch As String
'Здесь коллекция Where не нужна, так как, теоретически, любая строка должна соответствовать требованию,
'но оставлю её для совместимости, чтоб код формирования как можно меньше отличался от того что в сборе.

Sub ButtonGenerate()
    file = Application.GetOpenFilename("Файлы Excel(*.xls*),*.xls*", 1, _
        "Выберите файл реестра", , False)
    If VarType(file) = vbBoolean Then Exit Sub
    ExportBook ByVal CStr(file)
    'ExportBook "c:\DataCollect\BookTest\Crn001 Дубликат! Код формы 9.xlsx"
End Sub

Public Sub ExportBook(ByVal file As String)
    Message "Подготовка..."
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Patch = FSO.GetParentFolderName(file) + "\"
    Application.ScreenUpdating = False
    On Error GoTo er
    Set templ = Workbooks.Open(file, False, False)
    Set DAT = templ.Worksheets(1)
    Set BUY = templ.Worksheets("Покупатели")
    Set SEL = templ.Worksheets("Продавцы")
    
    GetLists
    GetQuartalsAndIndexes
    ClearOldBooks

    For Each q In Quartals
        For Each b In BuyersList
            For Each s In SellersList
                MakeBook q, b, s
            Next
        Next
    Next
    
    templ.Close
    Message "Готово!"
    Exit Sub
er:
    MsgBox "Произошла ошибка при открытии файла реестра"
End Sub

'Чтение справочников покупателей и продавцов из реестра
Sub GetLists()
    
    Set BuyersList = New Collection
    Set SellersList = New Collection
    
    i = 2
    Do While BUY.Cells(i, 1) <> ""
        BuyersList.Add BUY.Cells(i, 1).text, BUY.Cells(i, 1).text
        i = i + 1
    Loop
    
    i = 2
    Do While SEL.Cells(i, 1) <> ""
        SellersList.Add SEL.Cells(i, 1).text, SEL.Cells(i, 1).text
        i = i + 1
    Loop

End Sub

'Чтение базы, подготовка списка кварталов и индексов строк, в которых есть совпадения
Sub GetQuartalsAndIndexes()
    Set Where = New Collection
    Set Quartals = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    i = firstDat
    Do While DAT.Cells(i, cCom) <> ""
        If DAT.Cells(i, cCom) = OK1 Or DAT.Cells(i.cCom) = OK2 Then
            b = ""
            b = BuyersList(DAT.Cells(i, cBuyer))
            s = ""
            s = SellersList(DAT.Cells(i, cSeller))
            If b <> "" And s <> "" Then
                Where.Add i
                Quartals(GetQuartal(DAT.Cells(i, cDates))) = 1
            End If
        End If
        i = i + 1
    Loop
End Sub

'Вычисление номера квартала в формате "1-20"
Function GetQuartal(d As Date) As String
    GetQuartal = CStr((Month(d) - 1) \ 3 + 1) + "-" + Right(CStr(Year(d)), 2)
End Function

'Вычисление периода по номеру квартала
Function Period(q As String)
    y = ".20" + Right(q, 2)
    If Left(q, 1) = "1" Then Period = "с 01.01" + y + " по 31.03" + y
    If Left(q, 1) = "2" Then Period = "с 01.04" + y + " по 30.06" + y
    If Left(q, 1) = "3" Then Period = "с 01.07" + y + " по 30.09" + y
    If Left(q, 1) = "4" Then Period = "с 01.10" + y + " по 31.12" + y
End Function

'Чистка директории от предыдущих книг
Sub ClearOldBooks()
    On Error GoTo er
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set curfold = FSO.GetFolder(Patch)
    For Each file In curfold.Files
        If file.name Like "КнПрод*.xls*" Then
            Kill Patch + file.name
        End If
    Next
    
    Exit Sub
er:
    MsgBox "Произошла ошибка при удалении старых книг продаж. Формирование книг отменено."
    End
End Sub

'Формирование книги
Sub MakeBook(ByVal q As String, ByVal b As String, ByVal s As String)
    
    'Поиск данных для текущей комбинации квартал+покупатель+продавец
    Dim Finded As Collection
    Set Finded = New Collection
    For Each j In Where
        If q = GetQuartal(DAT.Cells(j, 2)) And b = DAT.Cells(j, cBuyer).text And _
            s = DAT.Cells(j, cSeller).text Then Finded.Add j
    Next
    If Finded.Count = 0 Then Exit Sub
    
    'Какие-то данные всёже нашли, создаём книгу
    name = cutBadSymbols("КнПрод " + s + " - " + b + " " + q)
    Filename = Patch + name + ".xlsx"
    Message "Формирование книги " + name
    Workbooks.Add
    Range(Cells(1, 1), Cells(1048576, 24)).Font.name = "Arial"
    Range(Cells(1, 1), Cells(1048576, 24)).Font.Size = 9
    
    'Заголовок
    e = Chr(10)
    Rows(1).RowHeight = 18.8
    bigCell "Книга продаж", 1, 1, 1, 24
    Cells(1, 1).Font.Size = 14
    Rows(2).RowHeight = 10.9
    Rows(3).RowHeight = 12
    Cells(3, 1) = "Продавец " + s
    Rows(4).RowHeight = 12
    Cells(4, 1) = "Идентификационный номер и код причины постановки на учeт налогоплательщика-продавца " + _
        DAT.Cells(Finded(1), cSellINN).text
    Rows(5).RowHeight = 12
    Cells(5, 1) = "Продажа за период " + Period(q)
    Rows(6).RowHeight = 12.8
    Cells(6, 1) = "Отбор: Контрагент = " + DAT.Cells(Finded(1), cBuyer)
    Cells(6, 1).Font.Bold = True
    
    'Шапка таблицы
    Rows(7).RowHeight = 90.8
    Rows(8).RowHeight = 40.9
    Rows(9).RowHeight = 10.9
    Range(Cells(7, 1), Cells(9, 24)).Font.Size = 8
    Range(Cells(7, 1), Cells(9, 24)).Font.Bold = True
    Range(Cells(7, 1), Cells(9, 24)).HorizontalAlignment = xlCenter
    Range(Cells(7, 1), Cells(9, 24)).VerticalAlignment = xlCenter
    '1
    Columns(1).ColumnWidth = 6.08
    bigCell "№" + e + "п/п", 7, 1, 2, 1
    Cells(9, 1) = "1"
    '2
    Columns(2).ColumnWidth = 6.75
    bigCell "Код" + e + "вида" + e + "опера-" + e + "ции", 7, 2, 2, 1
    Cells(9, 2) = "2"
    '3
    Columns(3).ColumnWidth = 14.58
    bigCell "Номер и дата" + e + "счета-фактуры" + e + "продавца", 7, 3, 2, 1
    Cells(9, 3) = "3"
    '3а
    Columns(4).ColumnWidth = 14.58
    bigCell "Регистраци-" + e + "онный номер" + e + "таможенной" + e + "декларации", 7, 4, 2, 1
    Cells(9, 4) = "3а"
    '3б
    Columns(5).ColumnWidth = 12.25
    bigCell "Код вида" + e + "товара", 7, 5, 2, 1
    Cells(9, 5) = "3б"
    '4
    Columns(6).ColumnWidth = 14.58
    bigCell "Номер и дата" + e + "исправления" + e + "счета-фактуры" + e + "продавца", 7, 6, 2, 1
    Cells(9, 6) = "4"
    '5
    Columns(7).ColumnWidth = 14.16
    bigCell "Номер и дата" + e + "корректиро-" + e + "вочного" + e + "счета-фактуры" + e + _
        "продавца", 7, 7, 2, 1
    Cells(9, 7) = "5"
    '6
    Columns(8).ColumnWidth = 16.92
    bigCell "Номер и дата" + e + "исправления" + e + "корректиро-" + e + "вочного счета-" + e + _
        "фактуры продавца", 7, 8, 2, 1
    Cells(9, 8) = "6"
    '7
    Columns(9).ColumnWidth = 16.5
    bigCell "Наименование" + e + "покупателя", 7, 9, 2, 1
    Cells(9, 9) = "7"
    '8
    Columns(10).ColumnWidth = 12.25
    bigCell "ИНН/КПП" + e + "покупателя", 7, 10, 2, 1
    Cells(9, 10) = "8"
    '9-10
    Columns(11).ColumnWidth = 15.75
    Columns(12).ColumnWidth = 15.75
    bigCell "Сведения о посреднике" + e + "(комиссионере, агенте)", 7, 11, 1, 2
    bigCell "Наименование" + e + "посредника", 8, 11, 1, 1
    bigCell "ИНН/КПП" + e + "посредника", 8, 12, 1, 1
    Cells(9, 11) = "9"
    Cells(9, 12) = "10"
    '11
    Columns(13).ColumnWidth = 13.08
    bigCell "Номер и дата" + e + "документа," + e + "подтвержда-" + e + "ющего" + e + "оплату", 7, 13, 2, 1
    Cells(9, 13) = "11"
    '12
    Columns(14).ColumnWidth = 9.92
    bigCell "Наиме-" + e + "нование" + e + "и код" + e + "валюты", 7, 14, 2, 1
    Cells(9, 14) = "12"
    '13а-б
    Columns(15).ColumnWidth = 15.75
    Columns(16).ColumnWidth = 15.75
    bigCell "Стоимость продаж по счету-" + e + "фактуре, разница стоимости по" + e + _
        "корректировочному счету-" + e + "фактуре (включая НДС) в валюте" + e + "счета-фактуры", 7, 15, 1, 2
    bigCell "в валюте" + e + "счета-фактуры", 8, 15, 1, 1
    bigCell "в рублях и" + e + "копейках", 8, 16, 1, 1
    Cells(9, 15) = "13а"
    Cells(9, 16) = "13б"
    '14-16
    Columns(17).ColumnWidth = 15.75
    Columns(18).ColumnWidth = 15.75
    Columns(19).ColumnWidth = 15.75
    Columns(20).ColumnWidth = 15.75
    bigCell "Стоимость продаж, облагаемых налогом, по счету-фактуре, " + e + _
        "разница стоимости по корректировочному счету-фактуре " + e + _
        "(без НДС) в рублях и копейках, по ставке", 7, 17, 1, 4
    bigCell "20 процентов", 8, 17, 1, 1
    bigCell "18 процентов", 8, 18, 1, 1
    bigCell "10 процентов", 8, 19, 1, 1
    bigCell "0 процентов", 8, 20, 1, 1
    Cells(9, 17) = "14"
    Cells(9, 18) = "14а"
    Cells(9, 19) = "15"
    Cells(9, 20) = "16"
    '17-18
    Columns(21).ColumnWidth = 15.75
    Columns(22).ColumnWidth = 15.75
    Columns(23).ColumnWidth = 15.75
    bigCell "Сумма НДС по счету-фактуре," + e + "разница суммы налога по корректировочному" + e + _
        "счету-фактуре в рублях и копейках, по ставке", 7, 21, 1, 3
    bigCell "20 процентов", 8, 21, 1, 1
    bigCell "18 процентов", 8, 22, 1, 1
    bigCell "10 процентов", 8, 23, 1, 1
    Cells(9, 21) = "17"
    Cells(9, 22) = "17а"
    Cells(9, 23) = "18"
    '19
    Columns(24).ColumnWidth = 15.75
    bigCell "Стоимость" + e + "продаж," + e + "освобождаемых" + e + "от налога, по" + e + _
        "счету-фактуре," + e + "разница" + e + "стоимости" + e + "по корректиро-" + e + _
        "вочному" + e + "счету-фактуре" + e + "в рублях и" + e + "копейках", 7, 24, 2, 1
    Cells(9, 24) = "19"
    
    'Строки
    i = 10
    n = 1
    s1 = 0: s2 = 0: s3 = 0: s4 = 0: s5 = 0: s6 = 0
    For Each j In Finded
        Rows(i).RowHeight = 24
        Rows(i).VerticalAlignment = xlTop
        Cells(i, 1) = n
        Cells(i, 2).NumberFormat = "@"
        Cells(i, 2) = "01"
        Cells(i, 3) = DAT.Cells(j, 1).text + " от" + e + DAT.Cells(j, cDates).text
        Cells(i, 9) = DAT.Cells(j, cBuyer)
        Cells(i, 9).WrapText = True
        Cells(i, 10) = DAT.Cells(j, cBuyINN)
        Cells(i, 10).WrapText = True
        Cells(i, 16) = DAT.Cells(j, cPrice)
        Cells(i, 17) = DAT.Cells(j, 9): If Cells(i, 17) <> "" Then s1 = s1 + Cells(i, 17)
        Cells(i, 18) = DAT.Cells(j, 10): If Cells(i, 18) <> "" Then s2 = s2 + Cells(i, 18)
        Cells(i, 19) = DAT.Cells(j, 11): If Cells(i, 19) <> "" Then s3 = s3 + Cells(i, 19)
        Cells(i, 21) = DAT.Cells(j, 12): If Cells(i, 21) <> "" Then s4 = s4 + Cells(i, 21)
        Cells(i, 22) = DAT.Cells(j, 13): If Cells(i, 22) <> "" Then s5 = s5 + Cells(i, 22)
        Cells(i, 23) = DAT.Cells(j, 14): If Cells(i, 23) <> "" Then s6 = s6 + Cells(i, 23)
        Range(Cells(i, 9), Cells(i, 10)).WrapText = True
        Range(Cells(i, 15), Cells(i, 23)).NumberFormat = "### ### ##0.00"
        i = i + 1
        n = n + 1
    Next
    
    'Подвал
    Rows(i).RowHeight = 12.8
    Cells(i, 1) = "Итого"
    Range(Cells(i, 1), Cells(i, 16)).Merge
    Cells(i, 1).HorizontalAlignment = xlRight
    Range(Cells(i, 1), Cells(i, 24)).Font.Bold = True
    If s1 > 0 Then Cells(i, 17) = s1
    If s2 > 0 Then Cells(i, 18) = s2
    If s3 > 0 Then Cells(i, 19) = s3
    If s4 > 0 Then Cells(i, 21) = s4
    If s5 > 0 Then Cells(i, 22) = s5
    If s6 > 0 Then Cells(i, 23) = s6
    Range(Cells(i, 15), Cells(i, 23)).NumberFormat = "### ### ##0.00"
    Range(Cells(7, 1), Cells(i, 24)).Borders.Weight = 2
    
    'End
    
    'Сохранение и закрытие документа
    On Error GoTo er
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=Filename
    ActiveWorkbook.Close
    Exit Sub
er:
    ActiveWorkbook.Close
    MsgBox "Произошла ошибка при сохранении файла " + Filename
End Sub

Sub bigCell(txt As String, r As Variant, c As Variant, height As Variant, width As Variant)
    Cells(r, c) = txt
    Range(Cells(r, c), Cells(r + height - 1, c + width - 1)).Merge
    Cells(r, c).HorizontalAlignment = xlCenter
    Cells(r, c).VerticalAlignment = xlCenter
    Cells(r, c).Font.Bold = True
End Sub