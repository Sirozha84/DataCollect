Attribute VB_Name = "Verify"
Const cComment = 15

Dim Comment As String
Dim errors As Boolean
Dim limitO As Variant
Dim summO As Variant

'Инициализация словарей лимитов
Sub Init()
    Set limitO = CreateObject("Scripting.Dictionary")
    Set summO = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = 2
    Set dic = Sheets("Лимиты отгрузок")
    Do While dic.Cells(i, 1) <> ""
        cmp = dic.Cells(i, 1).text
        lim = dic.Cells(i, 2)
        limitO(cmp) = lim
        i = i + 1
    Loop
End Sub

'Проверка корректности данных
Function Verify(ByRef dat As Variant, ByRef src As Variant, ByVal iC As Long, ByVal iI As Long, _
    changed As Boolean) As Boolean
    
    Comment = ""
    errors = False
    red = RGB(255, 192, 192)
    grn = RGB(192, 255, 192)
    yel = RGB(255, 255, 192)
    Verify = True
    
    '2 - Дата
    dat.Cells(iC, 2).NumberFormat = "dd.MM.yyyy"
    If Not IsDate(dat.Cells(iC, 2)) Then
        dat.Cells(iC, 2).Interior.Color = red
        src.Cells(iI, 2).Interior.Color = red
        AddCom "Дата введена не корректно"
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
    Else
        'Проверка на выход из лимита
        c = dat.Cells(iC, 6)
        s = dat.Cells(iC, 7)
        summO(c) = summO(c) + s
        If summO(c) > limitO(c) Then AddCom "Общая сумма превышает лимит отгрузок"
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
    For i = 12 To 14
        dat.Cells(iC, i).NumberFormat = "### ### ##0.00"
        If Not isPriceNDS(dat.Cells(iC, i)) Then
            dat.Cells(iC, i).Interior.Color = red
            src.Cells(iI, i).Interior.Color = red
            AddCom "Сумма НДС введена не корректно"
        End If
    Next
    
    'Пишем комментарий и расскрашиваем его
    col = red
    If Not errors Then col = grn: Comment = "Принято"
    dat.Cells(iC, cComment) = Comment
    dat.Cells(iC, cComment).Interior.Color = col
    src.Cells(iI, cComment) = Comment
    src.Cells(iI, cComment).Interior.Color = col
    
    Verify = errors
    
End Function

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