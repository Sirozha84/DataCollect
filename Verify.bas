Attribute VB_Name = "Verify"
Dim Comment As String       'Строка с комментариями
Dim errors As Boolean       'Флаг наличия ошибок
Dim groups As Variant       'Словарь групп
Dim dateS As Variant        'Словарь дат регистраций
Dim limitOne As Variant     'Общий лимит на отгрузку одному покупателю
Dim limitAll As Variant     'Общий лимит на отгрузку
Dim buyers As Variant       'Словарь покупателей "у кого покупаем"
Dim qrtIndexes As Variant   'Индексы колонок квартала
Dim summOne As Variant      'Счётчики сумм продажи одному покупателю
Dim summAll As Variant      'Счётчики сумм продажи всем

'Инициализация словарей лимитов
Sub Init()
    
    Set dateS = CreateObject("Scripting.Dictionary")
    Set summOne = CreateObject("Scripting.Dictionary")
    Set summAll = CreateObject("Scripting.Dictionary")
    Set groups = CreateObject("Scripting.Dictionary")
    Set buyers = CreateObject("Scripting.Dictionary")
    Set selIndexes = CreateObject("Scripting.Dictionary")
    Set qrtIndexes = CreateObject("Scripting.Dictionary")
    
    'Чтение общих лимитов
    limitOne = DIC.Cells(1, 5)
    limitAll = DIC.Cells(2, 5)
    
    'Чтение словарей дат регистрации, лимитов отгрузок и групп
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        cmp = DIC.Cells(i, cINN).text
        dtt = DIC.Cells(i, cSDate)
        dateS(cmp) = dtt
        grp = DIC.Cells(i, cGroup).text
        groups(cmp) = grp
        selIndexes(cmp) = i
        Range(DIC.Cells(i, cPFact), DIC.Cells(i, cPFact + quartCount - 1)).NumberFormat = "### ### ##0.00"
        i = i + 1
    Loop
    
    'Индексирование кварталов
    For i = 0 To quartCount - 1
        qrtIndexes(IndexToQuartal(i)) = i
    Next

End Sub

'Проверка корректности данных отгрузок, возвращает true если нет ошибок
'iC - строка в данных
'iI - строка в исходниках
Function Verify(ByVal iC As Long, ByVal iI As Long, ByVal oldINN, ByVal oldSum) As Boolean
    
    Comment = ""
    errors = False
    Verify = True
    
    '2 - Дата
    If Not isDateMy(DAT.Cells(iC, 2).text) Then
        DAT.Cells(iC, 2).Interior.Color = colRed
        SRC.Cells(iI, 2).Interior.Color = colRed
        AddCom "Дата введена не корректно"
    Else
        DateTest iC
    End If
    
    '3 - ИНН/КПП
    If Not isINNKPP(DAT.Cells(iC, 3).text) Then
        DAT.Cells(iC, 3).Interior.Color = colRed
        SRC.Cells(iI, 3).Interior.Color = colRed
        AddCom "Неверные ИНН/КПП покупателя"
    End If
    
    '5 - ИНН
    If Not isINN(DAT.Cells(iC, 5).text) Then
        DAT.Cells(iC, 5).Interior.Color = colRed
        SRC.Cells(iI, 5).Interior.Color = colRed
        AddCom "Неверный ИНН продавца"
    Else
        If selIndexes(DAT.Cells(iC, 5).text) = Empty Then AddCom "ИНН не найден в справочнике"
    End If
    
    '7 - Стоимость
    If Not isPrice(DAT.Cells(iC, 7)) Then
        DAT.Cells(iC, 7).Interior.Color = colRed
        SRC.Cells(iI, 7).Interior.Color = colRed
        AddCom "Сумма с НДС введена не корректно"
    End If
    
    '8 - Ставка НДС
    If Not isNDS(DAT.Cells(iC, 8).text) Then
        DAT.Cells(iC, 8).Interior.Color = colRed
        SRC.Cells(iI, 8).Interior.Color = colRed
        AddCom "Неверная ставка НДС"
    End If
    
    '9-11 - Стоимость продаж облагаемых налогом
    For i = 9 To 11
        If Not isPriceNDS(DAT.Cells(iC, i)) Then
            DAT.Cells(iC, i).Interior.Color = colRed
            SRC.Cells(iI, i).Interior.Color = colRed
            errors = True
        End If
    Next
    
    '12-14 - Сумма НДС
    e = False
    For i = 12 To 14
        If Not isPriceNDS(DAT.Cells(iC, i)) Then e = True
    Next
    If e Then
        DAT.Cells(iC, i).Interior.Color = colRed
        SRC.Cells(iI, i).Interior.Color = colRed
        errors = True
    End If
    
    'Если нет ошибок в корректности ввода, запускаем проверку на лимиты
    If Not errors Then LimitsTest iC, iI, oldINN, oldSum
    
    'Пишем комментарий и расскрашиваем его
    col = colRed
    If Not errors Then col = colGreen: Comment = "Принято"
    DAT.Cells(iC, cCom) = Comment
    DAT.Cells(iC, cCom).Interior.Color = col
    SRC.Cells(iI, cCom) = Comment
    SRC.Cells(iI, cCom).Interior.Color = col
    
    Verify = Not errors
    
End Function

'Проверка корректности данных поступлений, возвращает true если нет ошибок
'i - номер строки
Function VerifyLoad(ByVal i As Long) As Boolean
    
    Comment = ""
    errors = False
    VerifyLoad = True

    'Дата
    If Not isDateMy(DTL.Cells(i, clDate).text) Then
        DTL.Cells(i, clDate).Interior.Color = colRed
        AddCom "Дата введена не корректно"
    End If

    'ИНН на выходе
    If Not isINN(DTL.Cells(i, clOutINN).text) Then
        DTL.Cells(i, clOutINN).Interior.Color = colRed
        AddCom "Неверный ИНН на выходе"
    End If

    'ИНН на входе
    If Not isINN(DTL.Cells(i, clInINN).text) Then
        DTL.Cells(i, clInINN).Interior.Color = colRed
        AddCom "Неверный ИНН на входе"
    End If

    'Стоимость
    If Not isPrice(DTL.Cells(i, clPrice)) Then
        DTL.Cells(i, clPrice).Interior.Color = colRed
        AddCom "Сумма с НДС введена не корректно"
    End If
    
    '9-11 - Стоимость продаж облагаемых налогом
    For j = 9 To 11
        If Not isPriceNDS(DTL.Cells(i, j)) Then
            DTL.Cells(i, j).Interior.Color = colRed
            errors = True
        End If
    Next
    
    '12-14 - Сумма НДС
    e = False
    For j = 12 To 14
        If Not isPriceNDS(DTL.Cells(i, j)) Then e = True
    Next
    If e Then
        DTL.Cells(i, j).Interior.Color = colRed
        errors = True
    End If
    
    'Пишем комментарий и расскрашиваем его
    col = colRed
    If Not errors Then col = colGreen: Comment = "Принято"
    DTL.Cells(i, clCom) = Comment
    DTL.Cells(i, clCom).Interior.Color = col
    
    VerifyLoad = Not errors
    
End Function

'Проверка правильности даты
Sub DateTest(ByVal i As Long)
    SEL = DAT.Cells(i, 6)
    dtt = DAT.Cells(i, 2)
    If dtt < dateS(SEL) Then AddCom "Дата СФ не может быть ранее регистрации продавца"
End Sub

'Проверка лимитов
'i, si - номера строк данных и реестра
'oldINN, oldSum - прежние инн продавца и прежняя сумма (если это перепроверка)
Sub LimitsTest(ByVal i As Long, ByVal si As Long, ByVal oldINN, ByVal oldSum)
    cod = DAT.Cells(i, cCode).text + "!"
    kv = Kvartal(DAT.Cells(i, 2))
    kvin = qrtIndexes(kv)
    SEL = DAT.Cells(i, cSellINN).text
    selCur = SEL + "!" + kv + "!"
    BUY = DAT.Cells(i, cBuyINN).text
    buyCur = BUY + "!" + kv + "!"
    grp = groups(SEL)
    Sum = 0
    For j = 12 To 14
        If IsNumeric(DAT.Cells(i, j)) Then Sum = Sum + DAT.Cells(i, j)
    Next
    summOne(selCur + BUY) = summOne(selCur + BUY) + Sum
    summAll(selCur) = summAll(selCur) + Sum
    e = False
    
    'Проверка на лимит одному покупателю
    If summOne(selCur + BUY) > limitOne Then _
            AddCom "Превышен лимит продаж данного продавца данному покупателю": e = True
    
    'Проверка на остатки
    'RestoreBalance DAT.Cells(i, 2), oldINN, oldSum
    ind = selIndexes(SEL)
    over = False
    For j = 0 To kvin
        If Sum > DIC.Cells(ind, cLimits + j) Then over = True
    Next
    If Not over Then
        DIC.Cells(ind, cPFact + kvin) = DIC.Cells(ind, cPFact + kvin) + Sum
    Else
        AddCom "Сумма превышает свободный остаток у данного продавца": e = True
    End If
    
    'Проверка на общий лимит продавца
    If summAll(selCur) > limitAll Then AddCom "Превышен общий лимит продаж данного продавца": e = True
    
    'Пометка суммы ошибочной, если есть хоть одна ошибка с лимитами
    If e Then
        DAT.Cells(i, cPrice).Interior.Color = colRed
        SRC.Cells(si, cPrice).Interior.Color = colRed
    End If
    
    'Проверка на связанных продавцов для одного покупателя
    If buyers(buyCur + grp) = "" Then
        buyers(buyCur + grp) = SEL
    Else
        If buyers(buyCur + grp) <> SEL Then AddCom "Указаны связанные продавцы для данного покупателя"
    End If
End Sub

'Восстановление остатка
Sub RestoreBalance(dt, oldINN, oldSum)
    kvin = qrtIndexes(Kvartal(dt))
    If oldSum > 0 Then
        ind = selIndexes(oldINN)
        If ind <> Empty Then _
                DIC.Cells(ind, cPFact + kvin) = DIC.Cells(ind, cPFact + kvin) - oldSum
    End If
End Sub

'Добавление комментария к строке
Sub AddCom(str As String)
    If Comment <> "" Then Comment = Comment + ", "
    Comment = Comment + str
    errors = True
End Sub

'Проверка на корректность даты
Function isDateMy(ByVal var As String)
    isDateMy = True
    On Error GoTo er
    s = Split(var, ".")
    If UBound(s) <> 2 Then GoTo er
    If CInt(s(0)) < 1 Or CInt(s(0)) > 31 Then GoTo er
    If CInt(s(1)) < 1 Or CInt(s(1)) > 12 Then GoTo er
    If CInt(s(2)) < 1900 Or CInt(s(2)) > 2100 Then GoTo er
    If Not IsDate(var) Then GoTo er
    Exit Function
er:
    isDateMy = False
End Function

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
        'Строка может быт и пустой, это тоже нормально
        If Not IsError(var) Then
            If var = "" Then isPriceNDS = True
        End If
    End If
End Function

'Проверка на корректность ИНН
Function isINN(ByVal str As String) As Boolean
    isINN = False
    If str = "" Then Exit Function
    If IsNumeric(str) And Len(str) = 10 Then isINN = True
End Function

'Проверка на корректность ИНН/КПП
Function isINNKPP(ByVal str As String) As Boolean
    isINNKPP = False
    If str = "" Then isINNKPP = False: Exit Function
    s = Split(str, "/")
    'Юридическое лицо
    If IsNumeric(s(0)) And Len(s(0)) = 10 And UBound(s) > 0 Then
        If IsNumeric(s(1)) And Len(s(1)) = 9 Then isINNKPP = True
    End If
    'ИП
    If IsNumeric(s(0)) And Len(s(0)) = 12 And UBound(s) = 0 Then isINNKPP = True
End Function

'Проверка на корректность НДС
Function isNDS(ByVal str As String) As Boolean
    isNDS = False
    If str = "10" Then isNDS = True
    If str = "18" Then isNDS = True
    If str = "20" Then isNDS = True
End Function