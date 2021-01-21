Attribute VB_Name = "Template"
Const LastRec = 10000           'Последняя строка записей (Первая всегда 5, вбита гвоздями)
Const maxBuyers = 100           'Максимальное количество покупателей
Const maxSellers = 100          'Максимальное количество продавцов

Public Sub Generate()
    
    Main.Init
    If IsNumeric(NUM.Cells(2, 1)) Then last = NUM.Cells(2, 1)
    Dim max As Long
    i = firstTempl
    Do While Cells(i, 1) <> "" Or Cells(i, 2) <> ""
        i = i + 1
    Loop
    'Генерируем шаблоны
    Set namelist = CreateObject("Scripting.Dictionary")
    max = i - 1
    fold = DAT.Cells(1, 3).text
    For i = firstTempl To max
        Message "Создение шаблона " + CStr(i - firstTempl + 1) + " из " + CStr(max - FirstClient + 1)
        cln = cutBadSymbols(Cells(i, 1).text)
        tem = cutBadSymbols(Cells(i, 2).text)
        'Проверим, уникальные ли имена
        uname = cln + "!" + tem
        If namelist(uname) = "" Then
            namelist(uname) = 0
            If Not isCode(Cells(i, 3)) Then
                cod = last + 1
                last = cod
                Cells(i, 3) = cod
            End If
            If Cells(i, cTStat).text <> "OK" Then
                'Создаём папку и файл
                folder fold + "\" + cln
                folder fold + "\" + cln + "\" + tem
                name = fold + "\" + cln + "\" + tem + "\" + tem + ".xlsx"
                res = NewTemplate(cln, tem, name, Cells(i, 3).text)
                If res = 0 Then
                    Cells(i, 4) = "Произошла ошибка при создании файла"
                    Cells(i, 5) = "Ошибка"
                End If
                If res = 1 Then
                    Cells(i, 4) = name
                    Cells(i, 5) = "Успешно!"
                    Cells(i, cTStat) = "OK"
                End If
                If res = 2 Then
                    Cells(i, 4) = name
                    Cells(i, 5) = "Файл уже существует, пропущено"
                End If
            Else
                Cells(i, 5) = "Шаблон уже был создан ранее"
            End If
        Else
            Cells(i, 5) = "Имя клиента или шаблона не уникально."
        End If
    Next
    NUM.Cells(2, 1) = last
    
    ActiveWorkbook.Save
    Message "Готово! Файл сохранён."
    
End Sub

'Проверка, похоже ли ячейка на код
Function isCode(n As Variant)
    isCode = False
    If IsNumeric(n) Then
        If n > 0 Then isCode = True
    End If
End Function

'Создание нового файла
'Возвращает 0 - файл не создан, 1 - файл создан, 2 - файл уже есть, промущено
Function NewTemplate(ByVal cln As String, ByVal tem As String, _
    ByVal fileName As String, ByVal cod As String) As Byte
    
    'Если файл существует - пропустим
    If Dir$(fileName) <> "" Then NewTemplate = 2: Exit Function
    
    'Создаём файл с нужными вкладками
    Workbooks.Add
    If isRelease Then On Error GoTo er2
    Application.DisplayAlerts = False
    Sheets.Add
    Sheets.Add
    Sheets(1).name = cln
    Sheets(2).name = "Покупатели"
    Sheets(3).name = "Продавцы"
    Sheets(4).Delete
    Sheets(4).Delete
er2:
    If isRelease Then On Error GoTo er
    Set temp = Application.ActiveSheet
    Set listb = Sheets(2)
    Set lists = Sheets(3)
    Cells(1, 1) = cod
    Cells(2, 1) = tmpVersion
    Range(Cells(1, 1), Cells(2, 1)).Font.Color = vbWhite
    Cells(1, 2) = "Клиент: " + cln
    Cells(2, 2) = "Реестр: " + tem
    
    'Вкладки со справочниками
    listb.Columns(1).ColumnWidth = 30
    listb.Columns(2).ColumnWidth = 20
    listb.Cells(1, 1) = "Наименование"
    listb.Cells(1, 2) = "ИНН/КПП"
    Range(listb.Cells(2, 2), listb.Cells(maxBuyers, 2)).NumberFormat = "@"
    lists.Columns(1).ColumnWidth = 30
    lists.Columns(2).ColumnWidth = 20
    lists.Cells(1, 1) = "Наименование"
    lists.Cells(1, 2) = "ИНН"
    Range(lists.Cells(2, 2), lists.Cells(maxSellers, 2)).NumberFormat = "@"
    
    'Основная вкладка. Рисуем шапку формы
    Columns(1).ColumnWidth = 20
    Columns(2).ColumnWidth = 15
    Columns(3).ColumnWidth = 22
    Columns(4).ColumnWidth = 15
    Columns(5).ColumnWidth = 10
    Columns(6).ColumnWidth = 15
    Columns(7).ColumnWidth = 15
    Columns(8).ColumnWidth = 10
    Columns(9).ColumnWidth = 12
    Columns(10).ColumnWidth = 12
    Columns(11).ColumnWidth = 12
    Columns(12).ColumnWidth = 12
    Columns(13).ColumnWidth = 12
    Columns(14).ColumnWidth = 12
    Rows(3).RowHeight = 30
    Rows(4).RowHeight = 30
    Cells(3, 1) = "СФ"
    Range(Cells(3, 1), Cells(3, 2)).merge
    Cells(3, 3) = "Сведения о покупателе"
    Range(Cells(3, 3), Cells(3, 4)).merge
    Cells(3, 5) = "Сведения о продавце"
    Range(Cells(3, 5), Cells(3, 6)).merge
    Cells(3, 7) = "Стоимость" + Chr(10) + "продаж с НДС"
    Cells(3, 8) = "Ставка" + Chr(10) + "НДС, %"
    Range(Cells(3, 8), Cells(4, 8)).merge
    Cells(3, 9) = "Стоимость продаж облагаемых налогом" + Chr(10) + "(в руб.) без НДС"
    Range(Cells(3, 9), Cells(3, 11)).merge
    Cells(3, 12) = "Сумма НДС"
    Range(Cells(3, 12), Cells(3, 14)).merge
    Cells(4, 1) = "№" + Chr(10) + "(стр. 020)"
    Cells(4, 2) = "Дата" + Chr(10) + "(стр. 030)"
    Cells(4, 3) = "ИНН/КПП"
    Cells(4, 4) = "Наименование"
    Cells(4, 5) = "ИНН"
    Cells(4, 6) = "Наименование"
    Cells(4, 7) = "в руб. и коп."
    Cells(4, 9) = "20%" + Chr(10) + "(стр. 170)"
    Cells(4, 10) = "18%" + Chr(10) + "(стр. 200)"
    Cells(4, 11) = "10%" + Chr(10) + "(стр. 205)"
    Cells(4, 12) = "20%" + Chr(10) + "(стр. 200)"
    Cells(4, 13) = "18%" + Chr(10) + "(стр. 205)"
    Cells(4, 14) = "10%" + Chr(10) + "(стр. 210)"
    Set hat = Range(Cells(3, 1), Cells(4, 14))
    hat.HorizontalAlignment = xlCenter
    hat.VerticalAlignment = xlCenter
    hat.Interior.Color = colGray
    hat.Borders.Weight = 3
    
    'Поле 2 - Дата
    setFormat 2, "date"
    setValidation 2, "date"
    allowEdit temp, 2, "Дата"
    
    'Поле 3 - ИНН покупателя, находится с помощью ВПР
    setRange(3).FormulaLocal = "=ВПР(D5;Покупатели!A$2:B$" + CStr(maxBuyers) + ";2;0)"
    setFormatConditions 3
    
    'Поле 4 - Покупатель, выбираем из списка
    setValidation 4, "buy"
    allowEdit temp, 4, "Покупатель"
    
    'Поле 5 - ИНН продавца, находится с помлщью ВПР
    setRange(5).FormulaLocal = "=ВПР(F5;Продавцы!A$2:B$" + CStr(maxSellers) + ";2;0)"
    setFormatConditions 5
    
    'Поле 6 - Продавец, выбираем из списка
    setValidation 6, "sell"
    allowEdit temp, 6, "Продавец"
    
    'Поле 7 - Стоимость
    setValidation 7, "num"
    setFormat 7, "money"
    allowEdit temp, 7, "Стоимость"
    Cells(1, 7).Borders.Weight = 3
    Cells(1, 7).FormulaLocal = "=СУММ(G5:G" + CStr(LastRec) + ")"
    
    'Поле 8 - Ставка НДС
    setValidation 8, "nds"
    allowEdit temp, 8, "Ставка НДС"
    
    'Общее 9-14
    For i = 9 To 14
        setFormat i, "money"
        Cells(1, i).Borders.Weight = 3
    Next
    
    'Поле 9-11 - Сумма с НДС 20,18,10%      Формула G/(100+H)*100
    setRange(9).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=20);ОКРУГЛ(G5/(100+H5)*100;2);"""")"
    setRange(10).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=18);ОКРУГЛ(G5/(100+H5)*100;2);"""")"
    setRange(11).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=10);ОКРУГЛ(G5/(100+H5)*100;2);"""")"
    Cells(1, 9).FormulaLocal = "=СУММ(I5:I" + CStr(LastRec) + ")"
    Cells(1, 10).FormulaLocal = "=СУММ(J5:J" + CStr(LastRec) + ")"
    Cells(1, 11).FormulaLocal = "=СУММ(K5:K" + CStr(LastRec) + ")"
    
    'Поле 12-14 - Сумма без НДС 20,18,10%   Формула G/(100+H)*H
    setRange(12).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=20);ОКРУГЛ(G5/(100+H5)*H5;2);"""")"
    setRange(13).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=18);ОКРУГЛ(G5/(100+H5)*H5;2);"""")"
    setRange(14).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=10);ОКРУГЛ(G5/(100+H5)*H5;2);"""")"
    Cells(1, 12).FormulaLocal = "=СУММ(L5:L" + CStr(LastRec) + ")"
    Cells(1, 13).FormulaLocal = "=СУММ(M5:M" + CStr(LastRec) + ")"
    Cells(1, 14).FormulaLocal = "=СУММ(N5:N" + CStr(LastRec) + ")"
    
    Range(Cells(4, 1), Cells(4, 14)).Rows.AutoFilter
    Range("A5").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    
    'Защита и сохранение книги
    SetProtect temp
    ActiveWorkbook.SaveAs fileName:=fileName    'Для тестов эти строки комментируем и смотрим
    ActiveWorkbook.Close                        'результат сразу (список только делаем из одного файла)
    NewTemplate = 1
    Exit Function
er:
    ActiveWorkbook.Close
    NewTemplate = 0
End Function

Function setRange(ByVal c As Integer) As Range
    Set setRange = Range(Cells(5, c), Cells(LastRec, c))
End Function

'Установка формата для колонки
Sub setFormat(ByVal c As Integer, format As String)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    If format = "date" Then rang.NumberFormat = "dd.MM.yyyy"
    If format = "money" Then rang.NumberFormat = "### ### ##0.00"
End Sub

'Установка условного форматирования для колонки
Sub setFormatConditions(c As Integer)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    With rang.FormatConditions
        .Add Type:=xlErrorsCondition
        .Item(.Count).Font.Color = vbWhite
    End With
End Sub

'Установка проверки значений
Sub setValidation(c As Integer, typ As String)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    If typ = "buy" Then
        formul = "=Покупатели!$A$2:$A$" + CStr(maxBuyers)
        With rang.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formul
            .ErrorMessage = "Только из списка, пожалуйста!"
        End With
    End If
    If typ = "sell" Then
        formul = "=Продавцы!$A$2:$A$" + CStr(maxSellers)
        With rang.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formul
            .ErrorMessage = "Только из списка, пожалуйста!"
        End With
    End If
    If typ = "date" Then
        formul = "=OR(AND(H5=10),AND(H5=18,B5<43466),AND(H5=20,B5>=43466))"
        With rang.Validation
            .Delete
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:=formul
            .ErrorMessage = "До 01.01.2019 ндс был 18%, после - 20%, или 10% в любое время"
        End With
    End If
    If typ = "num" Then
        With rang.Validation
            .Delete
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="0"
            .ErrorMessage = "Число должно быть больше 0"
        End With
    End If
    If typ = "nds" Then
        formul = "=OR(AND(H5=10),AND(H5=18,B5<43466),AND(H5=20,B5>=43466))"
        With rang.Validation
            .Delete
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:=formul
            .ErrorMessage = "До 01.01.2019 ндс был 18%, после - 20%, или 10% в любое время"
        End With
    End If
End Sub

'Установка разрешения редактирования для колонки
Sub allowEdit(sh As Variant, c As Integer, name As String)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    sh.Protection.AllowEditRanges.Add Title:=name, Range:=rang, Password:=""
    rang.Interior.Color = RGB(255, 255, 192)
End Sub