Attribute VB_Name = "Misc"
'Последняя правка: 05.05.2021 18:47

'Сообщение в строке статуса
Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub

'Создание папки
Public Sub MakeDir(ByVal name As String)
    On Error Resume Next
    MkDir (name)
End Sub

'Удаление неугодных символов для имени файла
Function cutBadSymbols(ByVal name As String) As String
    name = Replace(name, """", "")
    name = Replace(name, "*", "")
    name = Replace(name, "\", "")
    name = Replace(name, "|", "")
    name = Replace(name, "/", "")
    name = Replace(name, "?", "")
    name = Replace(name, ":", "")
    name = Replace(name, "<", "")
    name = Replace(name, ">", "")
    name = Replace(name, Chr(10), "")
    name = Replace(name, Chr(13), "")
    cutBadSymbols = Trim(name)
End Function

'Проверка на возможность созранениф
Function TrySave(file As Variant)
    On Error GoTo er
    newname = file + "_"
    Name file As newname
    Name newname As file
    TrySave = True
    Exit Function
er:
    TrySave = False
End Function

'Расчёт квартала по номеру индекса
Function IndexToQuartal(ByVal i As Integer) As String
    IndexToQuartal = CStr(lastYear - Int((lastQuartal + i) / 4) + 1) + CStr(4 - i Mod 4)
End Function

'Вычисляет из индекса КГГГГ
Function IndexToQYYYY(ByVal i As Integer) As String
    q = lastQuartal
    y = lastYear
    For j = 1 To i
        q = q - 1
        If q < 1 Then q = 4: y = y - 1
    Next
    IndexToQYYYY = CStr(q) + CStr(y)
    Exit Function
End Function

'Вычисляет индекс квартала из даты
'Если ошибка или дата в не диапазона - возвращает -1
Function DateToQIndex(ByVal d As Date) As Integer
    On Error GoTo er
    i = (lastYear - Year(d)) * 4
    i = i + lastQuartal - ((Month(d) - 1) \ 3) - 1
    If i < 0 Or i >= quartCount Then i = -1
    DateToQIndex = i
    Exit Function
er:
    DateToQIndex = -1
End Function

'Вычисляет индекс квартала из строки типа КГГГГ
'Если ошибка или дата в не диапазона - возвращает -1
Function StupidQToQIndex(ByVal d As String) As Integer
    On Error GoTo er
    i = (lastYear - CInt(Right(d, 4))) * 4
    i = i + lastQuartal - (CInt(Left(d, 1))) '- 1
    If i < 0 Or i >= quartCount Then i = -1
    StupidQToQIndex = i
    Exit Function
er:
    StupidQToQIndex = -1
End Function

'Установка защиты
'Возвращает True, если удалось, если возвращается False - значит пароль не подошёл
Function SetProtect(table As Variant) As Boolean
    On Error GoTo er
    table.Protect Secret, AllowFormattingColumns:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    SetProtect = True
    Exit Function
er:
    SetProtect = False
End Function

'Сумма значений, разделённых ";"
Function OneCellSum(cell As String)
    sts = Split(cell, ";")
    Sum = 0
    For Each st In sts
        Sum = Sum + CStr(st)
    Next
    OneCellSum = Sum
End Function

'******************** End of File ********************