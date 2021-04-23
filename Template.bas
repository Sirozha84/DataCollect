Attribute VB_Name = "Template"
'Last change: 23.04.2021 18:20

Const LastRec = 10000   'Последняя строка записей (Первая всегда 5, вбита гвоздями)
Const maxComps = 100    'Максимальное количество компаний (продавцов или покупателей)

Sub Generate()
    
    Main.Init
    If IsNumeric(PRP.Cells(7, 2)) Then last = PRP.Cells(7, 2)
    Dim max As Long
    i = firstTempl
    Do While Cells(i, cTClient) <> "" Or Cells(i, cTForm) <> ""
        i = i + 1
    Loop
    'Генерируем шаблоны
    Set namelist = CreateObject("Scripting.Dictionary")
    max = i - 1
    fold = DirImportSale
    For i = firstTempl To max
        Message "Создение шаблона " + CStr(i - firstTempl + 1) + " из " + CStr(max - firstTempl + 1)
        cln = cutBadSymbols(Cells(i, cTClient).text)
        brk = cutBadSymbols(Cells(i, cTBroker).text)
        tem = cutBadSymbols(Cells(i, cTForm).text)
        'Проверим, уникальные ли имена
        uname = cln + "!" + tem
        If namelist(uname) = "" Then
            namelist(uname) = 0
            If Not IsCode(Cells(i, cTCode)) Then
                cod = last + 1
                last = cod
                Cells(i, cTCode) = cod
            End If
            If Cells(i, cTStat).text <> "OK" Then
                'Создаём папку и файл
                If brk <> "" Then brk = "\" + brk
                MakeDir fold + "\" + cln
                MakeDir fold + "\" + cln + brk
                MakeDir fold + "\" + cln + brk + "\" + tem
                name = fold + "\" + cln + brk + "\" + tem + "\" + tem + ".xlsx"
                res = NewTemplate(cln, tem, name, Cells(i, cTCode).text)
                If res = 0 Then
                    Cells(i, cTFile) = "Произошла ошибка при создании файла"
                    Cells(i, cTResult) = "Ошибка"
                End If
                If res = 1 Then
                    Cells(i, cTFile) = name
                    Cells(i, cTResult) = "Успешно!"
                    Cells(i, cTStat) = "OK"
                End If
                If res = 2 Then
                    Cells(i, cTFile) = name
                    Cells(i, cTResult) = "Файл уже существует, пропущено"
                    Cells(i, cTStat) = "OK"
                End If
            Else
                Cells(i, cTResult) = "Шаблон был создан ранее"
            End If
        Else
            Cells(i, cTResult) = "Имя клиента или шаблона не уникально."
        End If
    Next
    PRP.Cells(7, 2) = last
    
    ActiveWorkbook.Save
    Message "Готово! Файл сохранён."
    
End Sub

'Проверка, похоже ли ячейка на код
Function IsCode(n As Variant)
    IsCode = False
    If IsNumeric(n) Then
        If n > 0 Then IsCode = True
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
    Sheets.Add
    Sheets.Add
    Sheets(1).name = cln
    Sheets(2).name = "Покупатели"
    Sheets(3).name = "Продавцы"
    On Error Resume Next 'На случай если изначально было 1 вкладка, (в 2010 создаются по умолчанию 3)
    Application.DisplayAlerts = False
    Sheets(4).Delete
    Sheets(4).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set temp = Application.ActiveSheet
    Cells(1, 1) = cod
    Cells(2, 1) = tmpVersion
    Range(Cells(1, 1), Cells(2, 1)).Font.Color = vbWhite
    Cells(1, 2) = "Клиент: " + cln
    Cells(2, 2) = "Реестр: " + tem
    
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
    SetFormat 2, "date"
    SetValidation 2, "date"
    AllowEdit 2, "Дата"
    
    'Поле 3 - ИНН покупателя, находится с помощью ВПР
    SetRange(3).FormulaLocal = "=ВПР(D5;Покупатели!A$2:B$" + CStr(maxComps) + ";2;0)"
    SetFormatConditions 3
    
    'Поле 4 - Покупатель, выбираем из списка
    SetValidation 4, "buy"
    AllowEdit 4, "Покупатель"
    
    'Поле 5 - ИНН продавца, находится с помлщью ВПР
    SetRange(5).FormulaLocal = "=ВПР(F5;Продавцы!A$2:B$" + CStr(maxComps) + ";2;0)"
    SetFormatConditions 5
    
    'Поле 6 - Продавец, выбираем из списка
    SetValidation 6, "sale"
    AllowEdit 6, "Продавец"
    
    'Поле 7 - Стоимость
    SetValidation 7, "num"
    SetFormat 7, "money"
    AllowEdit 7, "Стоимость"
    Cells(1, 7).Borders.Weight = 3
    Cells(1, 7).FormulaLocal = "=СУММ(G5:G" + CStr(LastRec) + ")"
    
    'Поле 8 - Ставка НДС
    SetValidation 8, "nds"
    AllowEdit 8, "Ставка НДС"
    
    'Общее 9-14
    For i = 9 To 14
        SetFormat i, "money"
        Cells(1, i).Borders.Weight = 3
    Next
    
    'Поле 9-11 - Сумма с НДС 20,18,10%      Формула G/(100+H)*100
    SetRange(9).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=20);ОКРУГЛ(G5-L5;2);"""")"
    SetRange(10).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=18);ОКРУГЛ(G5-M5;2);"""")"
    SetRange(11).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=10);ОКРУГЛ(G5-N5;2);"""")"
    Cells(1, 9).FormulaLocal = "=СУММ(I5:I" + CStr(LastRec) + ")"
    Cells(1, 10).FormulaLocal = "=СУММ(J5:J" + CStr(LastRec) + ")"
    Cells(1, 11).FormulaLocal = "=СУММ(K5:K" + CStr(LastRec) + ")"
    
    'Поле 12-14 - Сумма без НДС 20,18,10%   Формула G/(100+H)*H
    SetRange(12).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=20);ОКРУГЛ(G5/(100+H5)*H5;2);"""")"
    SetRange(13).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=18);ОКРУГЛ(G5/(100+H5)*H5;2);"""")"
    SetRange(14).FormulaLocal = "=ЕСЛИ(И(G5<>"""";H5=10);ОКРУГЛ(G5/(100+H5)*H5;2);"""")"
    Cells(1, 12).FormulaLocal = "=СУММ(L5:L" + CStr(LastRec) + ")"
    Cells(1, 13).FormulaLocal = "=СУММ(M5:M" + CStr(LastRec) + ")"
    Cells(1, 14).FormulaLocal = "=СУММ(N5:N" + CStr(LastRec) + ")"
        
    'Автофильтр
    Range(Cells(4, 1), Cells(4, 14)).Rows.AutoFilter
    
    'Закрепление области
    Range("A5").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    
    'Вкладки со справочниками
    For i = 2 To 3
        Sheets(i).Activate
        Columns(1).ColumnWidth = 30
        Columns(2).ColumnWidth = 20
        Cells(1, 1) = "Наименование"
        Cells(1, 2) = "ИНН/КПП"
        Range(Cells(2, 2), Cells(maxComps, 2)).NumberFormat = "@"
        Rows(2).Hidden = True
        With Range(Cells(3, 2), Cells(maxComps, 2)).Validation
            .Delete
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:= _
                "=OR(LEN(B3)=12,LEN(B3)=20)"
            .ErrorMessage = "Не корректная длина строки. Должно быть 12 или 20 символов."
        End With
        Range("A3").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
    Next
    
    'Защита и сохранение книги
    Sheets(1).Activate
    SetProtect ActiveSheet
    On Error GoTo er
    ActiveWorkbook.SaveAs fileName:=fileName    'Для тестов эти строки комментируем и смотрим
    ActiveWorkbook.Close                        'результат сразу (список только делаем из одного файла)
    NewTemplate = 1
    Exit Function
er:
    ActiveWorkbook.Close
    NewTemplate = 0
End Function

Function SetRange(ByVal c As Integer) As Range
    Set SetRange = Range(Cells(5, c), Cells(LastRec, c))
End Function

'Установка формата для колонки
Sub SetFormat(ByVal c As Integer, format As String)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    If format = "date" Then rang.NumberFormat = "dd.MM.yyyy"
    If format = "money" Then rang.NumberFormat = "### ### ##0.00"
End Sub

'Установка условного форматирования для колонки
Sub SetFormatConditions(c As Integer)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    With rang.FormatConditions
        .Add Type:=xlErrorsCondition
        .Item(.Count).Font.Color = vbWhite
    End With
End Sub

'Установка проверки значений
Sub SetValidation(c As Integer, typ As String)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    If typ = "buy" Then
        formul = "=Покупатели!$A$2:$A$" + CStr(maxComps)
        With rang.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formul
            .ErrorMessage = "Только из списка, пожалуйста!"
        End With
    End If
    If typ = "sale" Then
        formul = "=Продавцы!$A$2:$A$" + CStr(maxComps)
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
Sub AllowEdit(c As Integer, name As String)
    Set rang = Range(Cells(5, c), Cells(LastRec, c))
    ActiveSheet.Protection.AllowEditRanges.Add Title:=name, Range:=rang, Password:=""
    rang.Interior.Color = RGB(255, 255, 192)
End Sub

'******************** End of File ********************