Attribute VB_Name = "Verify"
Dim Comment As String   'Строка с комментариями
Dim errors As Boolean   'Флаг наличия ошибок
Dim groups As Variant   'Словарь групп
Dim dateS As Variant    'Словарь дат регистраций
Dim limitPrs As Variant 'Словарь лимитов на отгрузку
Dim limitOne As Variant 'Общий лимит на отгрузку одному покупателю
Dim limitAll As Variant 'Общий лимит на отгрузку
Dim summOne As Variant  'Счётчики сумм продажи одному покупателю
Dim summAll As Variant  'Счётчики сумм продажи всем
Dim buyers As Variant   'Словарь покупателей "у кого покупаем"

'Инициализация словарей лимитов
Sub Init()
    
    Set dateS = CreateObject("Scripting.Dictionary")
    Set limitPrs = CreateObject("Scripting.Dictionary")
    Set summOne = CreateObject("Scripting.Dictionary")
    Set summAll = CreateObject("Scripting.Dictionary")
    Set groups = CreateObject("Scripting.Dictionary")
    Set buyers = CreateObject("Scripting.Dictionary")
    Dim i As Long
    
    'Чтение общих лимитов
    limitOne = DIC.Cells(1, cLimits)
    limitAll = DIC.Cells(2, cLimits)
    
    'Чтение словарей дат регистрации, лимитов отгрузок и групп
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        cmp = DIC.Cells(i, cINN).text
        dtt = DIC.Cells(i, cSDate)
        dateS(cmp) = dtt
        lim = DIC.Cells(i, cLimits)
        limitPrs(cmp) = lim
        grp = DIC.Cells(i, cGroup).text
        groups(cmp) = grp
        i = i + 1
    Loop
   
End Sub

'Сохранение текущих значений отгрузок
Sub SaveValues()
    Dim i As Long
    i = 1
    VAL.Cells.Clear
    VAL.Columns(1).ColumnWidth = 7
    VAL.Columns(2).ColumnWidth = 20
    VAL.Columns(3).ColumnWidth = 20
    VAL.Columns(4).ColumnWidth = 10
    DrawTable summAll, "Полный объём отгрузки продавца", i
    DrawTable summOne, "Объём отгрузки по покупателям", i
End Sub

'Шапка в
Sub DrawTable(tabl As Variant, name As String, i As Long)
    VAL.Cells(i, 1) = name
    i = i + 1
    VAL.Cells(i, 1) = "Квартал"
    VAL.Cells(i, 2) = "Продавец"
    VAL.Cells(i, 3) = "Покупатель"
    VAL.Cells(i, 4) = "Объём"
    Range(VAL.Cells(i, 1), VAL.Cells(i, 100)).Interior.Color = colGray
    i = i + 1
    For Each sel In tabl
        s = Split(sel, "!")
        VAL.Cells(i, 1) = s(1)
        VAL.Cells(i, 2) = s(0)
        VAL.Cells(i, 3) = s(2)
        VAL.Cells(i, 4) = tabl(sel)
        i = i + 1
    Next
    i = i + 1
End Sub

'Проверка корректности данных, возвращает true если есть ошибки
'dat - таблица с данными
'src - таблица с исходниками
'iC - строка в данных
'iI - строка в исходниках
'changed - true если данные уже были зарегистрированы и запись проверяется на изменения
Function Verify(ByRef DAT As Variant, ByRef SRC As Variant, ByVal iC As Long, ByVal iI As Long, _
    changed As Boolean) As Boolean
    
    Comment = ""
    errors = False
    Verify = True
    
    '2 - Дата
    DAT.Cells(iC, 2).NumberFormat = "dd.MM.yyyy"
    If Not IsDate(DAT.Cells(iC, 2)) Then
        DAT.Cells(iC, 2).Interior.Color = colRed
        SRC.Cells(iI, 2).Interior.Color = colRed
        AddCom "Дата введена не корректно"
    Else
        Call DateTest(DAT, iC)
    End If
    
    '3 - ИНН
    If Not isINNKPP(DAT.Cells(iC, 3).text) Then
        DAT.Cells(iC, 3).Interior.Color = colRed
        SRC.Cells(iI, 3).Interior.Color = colRed
        AddCom "ИНН/КПП введены не корректно"
    End If
    
    '5 - ИНН
    If Not isINNKPP(DAT.Cells(iC, 5).text) Then
        DAT.Cells(iC, 5).Interior.Color = colRed
        SRC.Cells(iI, 5).Interior.Color = colRed
        AddCom "ИНН введён не корректно"
    End If
    
    '7 - Стоимость
    DAT.Cells(iC, 7).NumberFormat = "### ### ##0.00"
    If Not isPrice(DAT.Cells(iC, 7)) Then
        DAT.Cells(iC, 7).Interior.Color = colRed
        SRC.Cells(iI, 7).Interior.Color = colRed
        AddCom "Стоимость введена не корректно"
    End If
    
    '8 - Ставка НДС
    If Not isNDS(DAT.Cells(iC, 8).text) Then
        DAT.Cells(iC, 8).Interior.Color = colRed
        SRC.Cells(iI, 8).Interior.Color = colRed
        AddCom "НДС введён не корректно"
    End If
    
    '9-11 - Стоимость продаж облагаемых налогом
    For i = 9 To 11
        DAT.Cells(iC, i).NumberFormat = "### ### ##0.00"
        If Not isPriceNDS(DAT.Cells(iC, i)) Then
            DAT.Cells(iC, i).Interior.Color = colRed
            SRC.Cells(iI, i).Interior.Color = colRed
            AddCom "Стоимость продаж облагаемых налогом введена не корректно"
        End If
    Next
    
    '12-14 - Сумма НДС
    e = False
    For i = 12 To 14
        DAT.Cells(iC, i).NumberFormat = "### ### ##0.00"
        If Not isPriceNDS(DAT.Cells(iC, i)) Then e = True
    Next
    If e Then
        DAT.Cells(iC, i).Interior.Color = colRed
        SRC.Cells(iI, i).Interior.Color = colRed
        AddCom "Сумма НДС введена не корректно"
    Else
        LimitsTest DAT, iC
    End If
    
    'Пишем комментарий и расскрашиваем его
    col = colRed
    If Not errors Then col = colGreen: Comment = "Принято"
    DAT.Cells(iC, cCom) = Comment
    DAT.Cells(iC, cCom).Interior.Color = col
    SRC.Cells(iI, cCom) = Comment
    SRC.Cells(iI, cCom).Interior.Color = col
    
    Verify = errors
    
End Function

'Проверка правильности даты
Sub DateTest(ByRef DAT As Variant, ByVal i As Long)
    sel = DAT.Cells(i, 6)
    dtt = DAT.Cells(i, 2)
    If dtt < dateS(sel) Then AddCom "Дата операции не может быть ранее регистрации компании"
End Sub

'Вычисляет из даты Год+Квартал
Function Kvartal(DAT As Variant) As String
    On Error GoTo er
    Kvartal = CStr(Year(DAT)) + CStr((Month(DAT) - 1) \ 3 + 1)
    Exit Function
er:
    Kvartal = ""
End Function

'Проверка лимитов
Sub LimitsTest(ByRef DAT As Variant, ByVal i As Long)
    kv = "!" + Kvartal(DAT.Cells(i, 2)) + "!"
    sel = DAT.Cells(i, cSellINN).text
    selCur = sel + kv
    buy = DAT.Cells(i, cBuyINN).text
    buyCur = buy + kv
    grp = groups(sel)
    Sum = 0
    For j = 12 To 14
        If IsNumeric(DAT.Cells(i, j)) Then Sum = Sum + DAT.Cells(i, j)
    Next
    summOne(selCur + buy) = summOne(selCur + buy) + Sum
    summAll(selCur) = summAll(selCur) + Sum
    If summOne(selCur + buy) > limitOne Then AddCom "Превышен общий лимит продаж одному покупателю"
    If summAll(selCur) > limitPrs(sel) Then AddCom "Превышен лимит отгрузок" 'Персональный
    If summAll(selCur) > limitAll Then AddCom "Превышен общий лимит продаж"
    If buyers(buyCur + grp) = "" Then
        buyers(buyCur + grp) = sel
    Else
        If buyers(buyCur + grp) <> sel Then AddCom "Покупка у другого продавца группы"
    End If

End Sub

'Добавление комментария к строке
Sub AddCom(str As String)
    If Comment <> "" Then Comment = Comment + ", "
    Comment = Comment + str
    errors = True
End Sub

'Проверка на корректность цены
Function isPrice(ByVal var As Variant)
    isPrice = False
    If IsNumeric(var) Then
        If var >= 0 And var <> "" Then isPrice = True
    End If
End Function

'Проверка на корректность цены без НДС и его суммы
Function isPriceNDS(ByVal var As Variant)
    isPriceNDS = False
    If IsNumeric(var) Then
        If var >= 0 Then isPriceNDS = True
    Else
        If Not IsError(var) Then
            If var = "" Then isPriceNDS = True
        End If
    End If
End Function

'Проверка на корректность ИНН/КПП
Function isINNKPP(ByVal str As String) As Boolean
    If str = "" Then isINNKPP = False: Exit Function
    Dim s() As String
    s = Split(str, "/")
    isINNKPP = True
    If Not IsNumeric(s(0)) Then isINNKPP = False
    If Len(s(0)) <> 10 And Len(s(0)) <> 12 Then isINNKPP = False
    If UBound(s) > 0 Then
        If Not IsNumeric(s(1)) Then isINNKPP = False
        If Len(s(1)) <> 9 Then isINNKPP = False
    End If
End Function

'Проверка на корректность НДС
Function isNDS(ByVal str As String) As Boolean
    isNDS = False
    If str = "10" Then isNDS = True
    If str = "18" Then isNDS = True
    If str = "20" Then isNDS = True
End Function