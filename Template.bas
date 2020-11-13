Attribute VB_Name = "Template"
Const MaxRecords = 100  'Максимальное количество записей
Const FirstClient = 6   'Первая строка списка клиентов
Const Secret = "123"

Dim temp As Variant
Dim dat As Variant
Dim maxBuyers As Integer
Dim maxSellers As Integer

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
    Filename = Cells(1, 3) + "\" + name + ".xlsx"
    Workbooks.Add
    'On Error GoTo er2
    Application.DisplayAlerts = False
    Sheets.Add
    Sheets(1).name = name
    Sheets(2).name = "Справочники"
    Sheets(3).Delete
    Sheets(3).Delete
er2:
    
    'On Error GoTo er
    Set temp = Application.ActiveSheet
    Set lists = Sheets(2)
    Cells(1, 1) = cod
    
    'Копируем справочники
    For i = 1 To 4: lists.Columns(i).ColumnWidth = 20: Next
    i = 5
    j = 0
    Do While dat.Cells(i, 3) <> ""
        j = j + 1
        lists.Cells(j, 1) = dat.Cells(i, 3)
        lists.Cells(j, 2) = dat.Cells(i, 4)
        i = i + 1
    Loop
    maxBuyers = j
    i = 5
    j = 0
    Do While dat.Cells(i, 5) <> ""
        j = j + 1
        lists.Cells(j, 3) = dat.Cells(i, 5)
        lists.Cells(j, 4) = dat.Cells(i, 6)
        i = i + 1
    Loop
    maxSellers = j
    
    'Рисуем шапку формы
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
    
    'Поле2 - Дата
    Call setFormat(2, "date")
    Call allowEdit(2, "Дата")
    'Поле3 - ИНН, находится через впр
    For i = 5 To 4 + MaxRecords
        Cells(i, 3).FormulaLocal = "=ВПР(D" + CStr(i) + ";Справочники!A2:B" + _
        CStr(maxBuyers) + ";2;0)"
    Next
    setFormatConditions (3)
    'Поле3 - Покупатель, выбираем из списка
    Call setValidation(4, "b")
    Call allowEdit(4, "Покупатель")
    
    'Защита книги
    temp.Protect Secret, UserInterfaceOnly:=True
    lists.Protect Secret, UserInterfaceOnly:=True
    
    'Если требуется сразу открыть результат - закомментировать оставшиеся строки
    ActiveWorkbook.SaveAs Filename:=Filename
    ActiveWorkbook.Close
er:
End Sub

'Установка формата для колонки
Sub setFormat(c As Integer, format As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    If format = "date" Then rang.NumberFormat = "dd.MM.yyyy"
End Sub

'Установка условного форматирования для колонки
Sub setFormatConditions(c As Integer)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    With rang.FormatConditions
        .Add Type:=16
        .Item(.count).Font.Color = vbWhite
    End With
End Sub

'Установка списка проверки
Sub setValidation(c As Integer, list As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    If list = "b" Then
        formul = "=Справочники!$A$2:$A$" + CStr(maxBuyers)
    Else
        formul = "=Справочники!$C$2:$C$" + CStr(maxSellers)
    End If
    With rang.Validation
        .Delete
        .Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:=formul
        .ErrorMessage = "Только из списка, пожалуйста!"
    End With
End Sub

'Установка разрешения редактирования для колонки
Sub allowEdit(c As Integer, name As String)
    Set rang = Range(Cells(5, c), Cells(4 + MaxRecords, c))
    temp.Protection.AllowEditRanges.Add _
        Title:=name, _
        Range:=rang, _
        Password:=""
    rang.Interior.Color = RGB(255, 255, 192)
End Sub