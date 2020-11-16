Attribute VB_Name = "Template"
Const MaxRecords = 100  'Максимальное количество записей
Const FirstClient = 6   'Первая строка списка клиентов
Const Secret = "123"
Const maxBuyers = 100   'Максимальное количество покупателей
Const maxSellers = 100  'Максимальное количество продавцов

Dim temp As Variant
Dim dat As Variant

Public Sub Generate()
    'Находим максимальный код (если он есть)
    Dim i As Long
    i = FirstClient
    max = 0
    Do While Cells(i, 1) <> ""
        If max < Cells(i, 2) Then max = Cells(i, 2)
        i = i + 1
    Loop
    
    'Проставляем коды остальным клиентам, у которых его нет
    i = FirstClient
    Do While Cells(i, 1) <> ""
        If Cells(i, 2) = "" Then max = max + 1: Cells(i, 2) = max
        i = i + 1
    Loop
        
    'Генерируем шаблоны
    Dim total As Long
    total = i - 1
    Set dat = Application.ActiveSheet
    For i = FirstClient To total
        Message "Создение шаблона " + CStr(i - FirstClient + 1) + " из " + _
        CStr(total - FirstClient + 1)
        Call NewTemplate(Cells(i, 1), Cells(i, 2))
    Next
    
    Message "Готово!"
End Sub

'Создание нового файла
Sub NewTemplate(name As String, cod As Long)
    'Создаём файл с нужными вкладками
    Filename = Cells(1, 3) + "\" + name + ".xlsx"
    Workbooks.Add
    On Error GoTo er2
    Application.DisplayAlerts = False
    Sheets.Add
    Sheets.Add
    Sheets(1).name = name
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
    
    'Вкладки со справочниками
    listb.Columns(1).ColumnWidth = 20
    listb.Columns(2).ColumnWidth = 10
    listb.Cells(1, 1) = "Покупатели"
    listb.Cells(1, 2) = "ИНН/КПП"
    lists.Columns(1).ColumnWidth = 20
    lists.Columns(2).ColumnWidth = 10
    lists.Cells(1, 1) = "Продавцы"
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
    Call allowEdit(2, "Дата")
    'Поле 3 - ИНН покупателя, находится с помощью ВПР
    For i = 5 To 4 + MaxRecords
        Cells(i, 3).FormulaLocal = "=ВПР(D" + CStr(i) + ";Покупатели!A2:B" + _
        CStr(maxBuyers) + ";2;0)"
    Next
    setFormatConditions (3)
    'Поле 4 - Покупатель, выбираем из списка
    Call setValidation(4, "b")
    Call allowEdit(4, "Покупатель")
    'Поле 5 - ИНН продавца, находится с помлщью ВПР
    For i = 5 To 4 + MaxRecords
        Cells(i, 5).FormulaLocal = "=ВПР(F" + CStr(i) + ";Продавцы!A2:B" + _
        CStr(maxSellers) + ";2;0)"
    Next
    setFormatConditions (5)
    'Поле 6 - Продавец, выбираем из списка
    Call setValidation(6, "s")
    Call allowEdit(6, "Продавец")
    'Поле 7 - Стоимость
    Call setFormat(7, "money")
    Cells(1, 7).Borders.Weight = 3
    Cells(1, 7).FormulaLocal = "=СУММ(G5:G" + CStr(4 + MaxRecords) + ")"
    Call allowEdit(7, "Стоимость")
    'Поле 8 - Ставка НДС
    Call setValidation(8, "nds")
    Call allowEdit(8, "Ставка НДС")
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
    
    
    'Защита книги
    temp.Protect Secret, UserInterfaceOnly:=True
    'lists.Protect Secret, UserInterfaceOnly:=True
    
    'Если требуется сразу открыть результат - закомментировать оставшиеся строки
    ActiveWorkbook.SaveAs Filename:=Filename
    ActiveWorkbook.Close
er:
End Sub

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
Sub allowEdit(c As Integer, name As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    temp.Protection.AllowEditRanges.Add Title:=name, Range:=rang, Password:=""
    rang.Interior.Color = RGB(255, 255, 192)
End Sub