Attribute VB_Name = "Template"
Public Const Secret = "123"     'Пароль для защиты

Const FirstClient = 7   'Первая строка списка клиентов
Const MaxRecords = 100  'Максимальное количество записей
Const maxBuyers = 100   'Максимальное количество покупателей
Const maxSellers = 100  'Максимальное количество продавцов

Public Sub Generate()
        
    'Ищем максимальный код, из существующих и проверяем на дубликаты
    Dim max As Long
    max = 0
    Dim i As Long
    i = FirstClient
    Do While Cells(i, 1) <> "" Or Cells(i, 2) <> ""
        If isCode(Cells(i, 3)) Then
            If max < Cells(i, 3) Then max = Cells(i, 3)
        Else
            Cells(i, 3) = ""
        End If
        Cells(i, 4) = ""
        i = i + 1
    Loop
           
    'Генерируем шаблоны
    Set namelist = CreateObject("Scripting.Dictionary")
    Set codes = CreateObject("Scripting.Dictionary")
    Dim total As Long
    total = i - 1
    fold = Cells(1, 3).text
    For i = FirstClient To total
        Message "Создение шаблона " + CStr(i - FirstClient + 1) + " из " + _
            CStr(total - FirstClient + 1)
        cln = Cells(i, 1).text
        tem = Cells(i, 2).text
        If validName(cln) And validName(tem) Then
            'Проверим, уникальные ли имена
            uname = cln + "!" + tem
            If namelist(uname) = "" Then
                namelist(uname) = 0
                need = False
                If isCode(Cells(i, 3)) Then
                    cod = Cells(i, 3)
                    If codes(cod) = "" Then
                        codes(cod) = 0
                    Else
                        need = True
                    End If
                Else
                    need = True
                End If
                If need Then
                    max = max + 1
                    cod = max
                    Cells(i, 3) = cod
                End If
                name = fold + "\" + cln + "\" + tem + ".xlsx"
                'Создаём папку и файл
                folder (fold + "\" + cln)
                If NewTemplate(cln, tem, name, cod) Then
                    Cells(i, 4) = name
                Else
                    Cells(i, 4) = "Произошла ошибка при создании файла"
                End If
            Else
                Cells(i, 4) = "Имя клиента или шаблона не уникально."
            End If
        Else
            Cells(i, 4) = "Имя клиента или шаблона не указано или указано некорректно."
        End If
    Next
    
    Message "Готово!"
    
End Sub

'Проверка, похоже ли ячейка на код
Function isCode(n As Variant)
    isCode = False
    If IsNumeric(n) Then
        If n > 0 Then isCode = True
    End If
End Function

'Проверка на правильность имени файда/папки
Function validName(ByVal name As String) As Boolean
    validName = True
    If name = "" Then validName = False
    If InStr(name, """") Then validName = False
    If InStr(name, "*") Then validName = False
    If InStr(name, "\") Then validName = False
    If InStr(name, "|") Then validName = False
    If InStr(name, "/") Then validName = False
    If InStr(name, "?") Then validName = False
    If InStr(name, ":") Then validName = False
    If InStr(name, "<") Then validName = False
    If InStr(name, ">") Then validName = False
End Function

'Создание папки
Sub folder(name As String)
    On Error GoTo er
    MkDir (name)
er:
End Sub

'Создание нового файла
Function NewTemplate(ByVal cln As String, ByVal tem As String, _
    ByVal fileName As String, ByVal cod As String) As Boolean
    
    'Если файл существует - пропустим
    If Dir$(fileName) <> "" Then NewTemplate = True: Exit Function
    
    'Создаём файл с нужными вкладками
    Workbooks.Add
    On Error GoTo er2
    Application.DisplayAlerts = False
    Sheets.Add
    Sheets.Add
    Sheets(1).name = cln
    Sheets(2).name = "Покупатели"
    Sheets(3).name = "Продавцы"
    Sheets(4).Delete
    Sheets(4).Delete
er2:
    On Error GoTo er
    Set temp = Application.ActiveSheet
    Set listb = Sheets(2)
    Set lists = Sheets(3)
    Cells(1, 1) = cod
    Cells(1, 1).Font.Color = vbWhite
    Cells(1, 2) = "Клиент: " + cln
    Cells(2, 2) = "Шаблон: " + tem
    
    'Вкладки со справочниками
    listb.Columns(1).ColumnWidth = 30
    listb.Columns(2).ColumnWidth = 20
    listb.Cells(1, 1) = "Наименование"
    listb.Cells(1, 2) = "ИНН/КПП"
    lists.Columns(1).ColumnWidth = 30
    lists.Columns(2).ColumnWidth = 20
    lists.Cells(1, 1) = "Наименование"
    lists.Cells(1, 2) = "ИНН"
    
    'Основная вкладка. Рисуем шапку формы
    Columns(1).ColumnWidth = 20
    Columns(2).ColumnWidth = 15
    Columns(3).ColumnWidth = 30
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
    Range(Cells(3, 1), Cells(3, 2)).Merge
    Cells(3, 3) = "Сведения о покупателе"
    Range(Cells(3, 3), Cells(3, 4)).Merge
    Cells(3, 5) = "Сведения о продавце"
    Range(Cells(3, 5), Cells(3, 6)).Merge
    Cells(3, 7) = "Стоимость" + Chr(10) + "продаж с НДС"
    Cells(3, 8) = "Ставка" + Chr(10) + "НДС, %"
    Range(Cells(3, 8), Cells(4, 8)).Merge
    Cells(3, 9) = "Стоимость продаж облагаемых налогом" + Chr(10) + "(в руб.) без НДС"
    Range(Cells(3, 9), Cells(3, 11)).Merge
    Cells(3, 12) = "Сумма НДС"
    Range(Cells(3, 12), Cells(3, 14)).Merge
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
    hat.Interior.Color = RGB(224, 224, 224)
    hat.Borders.Weight = 3
    
    'Поле 2 - Дата
    Call setFormat(2, "date")
    Call setValidation(2, "date")
    Call allowEdit(temp, 2, "Дата")
    'Поле 3 - ИНН покупателя, находится с помощью ВПР
    For i = 5 To 4 + MaxRecords
        Cells(i, 3).FormulaLocal = "=ВПР(D" + CStr(i) + ";Покупатели!A2:B" + _
        CStr(maxBuyers) + ";2;0)"
    Next
    setFormatConditions (3)
    'Поле 4 - Покупатель, выбираем из списка
    Call setValidation(4, "b")
    Call allowEdit(temp, 4, "Покупатель")
    'Поле 5 - ИНН продавца, находится с помлщью ВПР
    For i = 5 To 4 + MaxRecords
        Cells(i, 5).FormulaLocal = "=ВПР(F" + CStr(i) + ";Продавцы!A2:B" + _
        CStr(maxSellers) + ";2;0)"
    Next
    setFormatConditions (5)
    'Поле 6 - Продавец, выбираем из списка
    Call setValidation(6, "s")
    Call allowEdit(temp, 6, "Продавец")
    'Поле 7 - Стоимость
    Call setFormat(7, "money")
    Cells(1, 7).Borders.Weight = 3
    Cells(1, 7).FormulaLocal = "=СУММ(G5:G" + CStr(4 + MaxRecords) + ")"
    Call allowEdit(temp, 7, "Стоимость")
    'Поле 8 - Ставка НДС
    Call setValidation(8, "nds")
    Call allowEdit(temp, 8, "Ставка НДС")
    'Общее 9-14
    For i = 9 To 14
        Call setFormat(i, "money")
        Cells(1, i).Borders.Weight = 3
    Next
    'Поле 9-11 - Сумма с НДС 20,18,10%      Формула G/(100+H)*100
    For i = 5 To 4 + MaxRecords
        Cells(i, 9).FormulaLocal = "=ЕСЛИ(И(G" + CStr(i) + "<>"""";H" + CStr(i) + "=20);" + _
        "ОКРУГЛ(G" + CStr(i) + "/(100+H" + CStr(i) + ")*100;2);"""")"
        Cells(i, 10).FormulaLocal = "=ЕСЛИ(И(G" + CStr(i) + "<>"""";H" + CStr(i) + "=18);" + _
        "ОКРУГЛ(G" + CStr(i) + "/(100+H" + CStr(i) + ")*100;2);"""")"
        Cells(i, 11).FormulaLocal = "=ЕСЛИ(И(G" + CStr(i) + "<>"""";H" + CStr(i) + "=10);" + _
        "ОКРУГЛ(G" + CStr(i) + "/(100+H" + CStr(i) + ")*100;2);"""")"
    Next
    Cells(1, 9).FormulaLocal = "=СУММ(I5:I" + CStr(4 + MaxRecords) + ")"
    Cells(1, 10).FormulaLocal = "=СУММ(J5:J" + CStr(4 + MaxRecords) + ")"
    Cells(1, 11).FormulaLocal = "=СУММ(K5:K" + CStr(4 + MaxRecords) + ")"
    'Поле 12-14 - Сумма без НДС 20,18,10%   Формула G/(100+H)*H
    For i = 5 To 4 + MaxRecords
        Cells(i, 12).FormulaLocal = "=ЕСЛИ(И(G" + CStr(i) + "<>"""";H" + CStr(i) + "=20);" + _
        "ОКРУГЛ(G" + CStr(i) + "/(100+H" + CStr(i) + ")*H" + CStr(i) + ";2);"""")"
        Cells(i, 13).FormulaLocal = "=ЕСЛИ(И(G" + CStr(i) + "<>"""";H" + CStr(i) + "=18);" + _
        "ОКРУГЛ(G" + CStr(i) + "/(100+H" + CStr(i) + ")*H" + CStr(i) + ";2);"""")"
        Cells(i, 14).FormulaLocal = "=ЕСЛИ(И(G" + CStr(i) + "<>"""";H" + CStr(i) + "=10);" + _
        "ОКРУГЛ(G" + CStr(i) + "/(100+H" + CStr(i) + ")*H" + CStr(i) + ";2);"""")"
    Next
    Cells(1, 12).FormulaLocal = "=СУММ(L5:L" + CStr(4 + MaxRecords) + ")"
    Cells(1, 13).FormulaLocal = "=СУММ(M5:M" + CStr(4 + MaxRecords) + ")"
    Cells(1, 14).FormulaLocal = "=СУММ(N5:N" + CStr(4 + MaxRecords) + ")"
    
    'Защита и сохранение книги
    temp.Protect Secret, UserInterfaceOnly:=True
    ActiveWorkbook.SaveAs fileName:=fileName    'Для тестов эти строки комментируем и смотрим
    ActiveWorkbook.Close                        'результат сразу (список только делаем из одного файла)
    NewTemplate = True
    Exit Function
er:
    NewTemplate = False
End Function

'Установка формата для колонки
Sub setFormat(ByVal c As Integer, format As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    If format = "date" Then rang.NumberFormat = "dd.MM.yyyy"
    If format = "money" Then rang.NumberFormat = "### ### ##0.00"
End Sub

'Установка условного форматирования для колонки
Sub setFormatConditions(c As Integer)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    With rang.FormatConditions
        .Add Type:=16
        .Item(.count).Font.Color = vbWhite
    End With
End Sub

'Установка проверки значений
Sub setValidation(c As Integer, typ As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    If typ = "b" Then formul = "=Покупатели!$A$2:$A$" + CStr(maxBuyers)
    If typ = "s" Then formul = "=Продавцы!$A$2:$A$" + CStr(maxSellers)
    If typ = "nds" Then formul = "10,18,20"
    If formul <> "" Then
        With rang.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formul
            .ErrorMessage = "Только из списка, пожалуйста!"
        End With
    End If
    If typ = "date" Then
        With rang.Validation
            .Delete
            '30000 - какая-то дата 82-го года, так и не понял как записать человеческую дату
            .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="30000"
            .ErrorMessage = "Тут должна быть дата!"
        End With
    End If
End Sub

'Установка разрешения редактирования для колонки
Sub allowEdit(sh As Variant, c As Integer, name As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    sh.Protection.AllowEditRanges.Add Title:=name, Range:=rang, Password:=""
    rang.Interior.Color = RGB(255, 255, 192)
End Sub