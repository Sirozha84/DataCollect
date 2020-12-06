Attribute VB_Name = "Verify"
Const cComment = 15     'Колонка для комментария
Const startLimits = 4   'Первые строки в таблицах лимитов

Dim Comment As String   'Строка с комментариями
Dim errors As Boolean   'Флаг наличия ошибок
Dim groups As Variant   'Словарь групп

Dim dates As Variant    'Словарь дат регистраций
Dim limitO As Variant   'Словарь лимитов на отгрузку
Dim limitP As Variant   'Словарь лимитов на покупку
Dim limitP1 As Variant  'Общий лимит на отгрузку одному покупателю
Dim limitPA As Variant  'Общий лимит на отгрузку
Dim summO As Variant    'Счётчики сумм на отгрузку
Dim summP As Variant    'Счётчики сумм на покупку
Dim summP1 As Variant   'Счётчики сумм продажи одному покупателю
Dim summPA As Variant   'Счётчики сумм продажи всем

'Инициализация словарей лимитов
Sub Init()
    
    Set dic = Sheets("Справочник")
    Set dates = CreateObject("Scripting.Dictionary")
    Set limitO = CreateObject("Scripting.Dictionary")
    Set limitP = CreateObject("Scripting.Dictionary")
    limitP1 = dic.Cells(1, 4)
    limitPA = dic.Cells(2, 4)
    Set summO = CreateObject("Scripting.Dictionary")
    Set summP = CreateObject("Scripting.Dictionary")
    Set summP1 = CreateObject("Scripting.Dictionary")
    Set summPA = CreateObject("Scripting.Dictionary")
    Set groups = CreateObject("Scripting.Dictionary")
    Dim i As Long

    'Чтение словаря дат регистраций компаний
    i = startLimits
    Do While dic.Cells(i, 1) <> ""
        cmp = dic.Cells(i, 1).text
        dtt = dic.Cells(i, 2)
        dates(cmp) = dtt
        i = i + 1
    Loop

    'Чтение словаря лимитов отгрузок
    i = startLimits
    Do While dic.Cells(i, 1) <> ""
        cmp = dic.Cells(i, 1).text
        lim = dic.Cells(i, 4)
        limitO(cmp) = lim
        i = i + 1
    Loop
    
    'Чтение словаря групп
    i = startLimits
    Do While dic.Cells(i, 1) <> ""
        cmp = dic.Cells(i, 1).text
        grp = dic.Cells(i, 3).text
        groups(cmp) = grp
        i = i + 1
    Loop
    
    'Чтение словаря лимитов продаж
    i = startLimits
    Do While dic.Cells(i, 3) <> ""
        grp = dic.Cells(i, 3).text
        If dic.Cells(i, 5).text <> "" And IsNumeric(dic.Cells(i, 5)) Then
            lim = dic.Cells(i, 5)
            limitP(grp) = lim
        End If
        i = i + 1
    Loop
    
End Sub

'Проверка корректности данных, возвращает true если есть ошибки
'dat - таблица с данными
'src - таблица с исходниками
'iC - строка в данных
'iI - строка в исходниках
'changed - true если данные уже были зарегистрированы и запись проверяется на изменения
Function Verify(ByRef dat As Variant, ByRef src As Variant, ByVal iC As Long, ByVal iI As Long, _
    changed As Boolean) As Boolean
    
    red = RGB(255, 192, 192)
    grn = RGB(192, 255, 192)
    yel = RGB(255, 255, 192)
    
    Comment = ""
    errors = False
    Verify = True
    
    '2 - Дата
    dat.Cells(iC, 2).NumberFormat = "dd.MM.yyyy"
    If Not IsDate(dat.Cells(iC, 2)) Then
        dat.Cells(iC, 2).Interior.Color = red
        src.Cells(iI, 2).Interior.Color = red
        AddCom "Дата введена не корректно"
    Else
        Call DateTest(dat, iC)
    End If
    
    '3 - ИНН
    If Not isINNKPP(dat.Cells(iC, 3).text) Then
        dat.Cells(iC, 3).Interior.Color = red
        src.Cells(iI, 3).Interior.Color = red
        AddCom "ИНН/КПП введены не корректно"
    End If
    
    '5 - ИНН
    If Not isINNKPP(dat.Cells(iC, 5).text) Then
        dat.Cells(iC, 5).Interior.Color = red
        src.Cells(iI, 5).Interior.Color = red
        AddCom "ИНН введён не корректно"
    End If
    
    '7 - Стоимость
    dat.Cells(iC, 7).NumberFormat = "### ### ##0.00"
    If Not isPrice(dat.Cells(iC, 7)) Then
        dat.Cells(iC, 7).Interior.Color = red
        src.Cells(iI, 7).Interior.Color = red
        AddCom "Стоимость введена не корректно"
    End If
    
    '8 - Ставка НДС
    If Not isNDS(dat.Cells(iC, 8).text) Then
        dat.Cells(iC, 8).Interior.Color = red
        src.Cells(iI, 8).Interior.Color = red
        AddCom "НДС введён не корректно"
    End If
    
    '9-11 - Стоимость продаж облагаемых налогом
    For i = 9 To 11
        dat.Cells(iC, i).NumberFormat = "### ### ##0.00"
        If Not isPriceNDS(dat.Cells(iC, i)) Then
            dat.Cells(iC, i).Interior.Color = red
            src.Cells(iI, i).Interior.Color = red
            AddCom "Стоимость продаж облагаемых налогом введена не корректно"
        End If
    Next
    
    '12-14 - Сумма НДС
    e = False
    For i = 12 To 14
        dat.Cells(iC, i).NumberFormat = "### ### ##0.00"
        If Not isPriceNDS(dat.Cells(iC, i)) Then e = True
    Next
    If e Then
        dat.Cells(iC, i).Interior.Color = red
        src.Cells(iI, i).Interior.Color = red
        AddCom "Сумма НДС введена не корректно"
    Else
        Call LimitsTest(dat, iC)
    End If
    
    'Пишем комментарий и расскрашиваем его
    col = red
    If Not errors Then col = grn: Comment = "Принято"
    dat.Cells(iC, cComment) = Comment
    dat.Cells(iC, cComment).Interior.Color = col
    src.Cells(iI, cComment) = Comment
    src.Cells(iI, cComment).Interior.Color = col
    
    Verify = errors
    
End Function

'Проверка правильности даты
Sub DateTest(ByRef dat As Variant, ByVal i As Long)
    sel = dat.Cells(i, 6)
    dtt = dat.Cells(i, 2)
    If dtt < dates(sel) Then AddCom "Дата операции не может быть ранее регистрации компании"
End Sub

'Вычисляет из даты Год+Квартал
Function Kvartal(dat As Date) As String
    Kvartal = CStr(Year(dat)) + CStr((Month(dat) - 1) \ 3 + 1)
End Function

'Проверка лимитов
Sub LimitsTest(ByRef dat As Variant, ByVal i As Long)
    kvr = Kvartal(dat.Cells(i, 2))
    sel = dat.Cells(i, 6)
    selK = sel + "!" + kvr
    grp = groups(sel)
    buy = dat.Cells(i, 4) + "!" + grp
    Sum = 0
    For j = 12 To 14
        If IsNumeric(dat.Cells(i, j)) Then Sum = Sum + dat.Cells(i, j)
    Next
    summO(selK) = summO(selK) + Sum
    summP(buy) = summP(buy) + Sum
    summP1(sel + buy) = summPA(sel + buy) + Sum
    summPA(selK) = summPA(selK) + Sum
    If summO(selK) > limitO(sel) Then AddCom "Превышен лимит отгрузок" 'Персональный
    If summP(buy) > limitP(grp) Then AddCom "Превышен лимит покупок"  'Персональный
    If summP1(sel + buy) > limitP1 Then AddCom "Превышен общий лимит продаж одному покупателю"
    If summPA(selK) > limitPA Then AddCom "Превышен общий лимит продаж"
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