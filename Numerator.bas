Attribute VB_Name = "Numerator"
Dim Prefixes As Object  'Словарь префиксов
Dim Liters As Object    'Словарь литеров
Dim Codes As Object     'Словарь кодов

'Инициализация словаря
Sub Init()
    
    'Загрузка словаря нумератора
    Range(NUM.Cells(1, 1), NUM.Cells(3, 100)).Interior.Color = RGB(214, 214, 214)
    Set Prefixes = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = firstNum
    Do While NUM.Cells(i, 1) <> ""
        pref = NUM.Cells(i, 1)
        Prefixes.Add pref, NUM.Cells(i, 2)
        i = i + 1
    Loop
    
    'Загрузка префиксов и кодов
    Set Liters = CreateObject("Scripting.Dictionary")
    Set Codes = CreateObject("Scripting.Dictionary")
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        buyer = DIC.Cells(i, 1)
        Liters.Add buyer, GetLiter(buyer, DIC.Cells(i, 8).text)
        Codes.Add buyer, DIC.Cells(i, 9).text
        i = i + 1
    Loop

End Sub

'Сохранение словаря на страницу
Sub Save()
    Dim i As Long
    i = firstNum
    For Each Key In Prefixes.keys
        NUM.Cells(i, 1) = Key
        NUM.Cells(i, 2) = Prefixes(Key)
        i = i + 1
    Next
End Sub

'Генерация уникального номера
Public Function Generate(recDate As Date, buyer As String) As String
    
    lit = Liters(buyer)
    Dim d As Byte
    cod = GetCode(buyer, recDate, d)
    
    pref = lit + cod
    If Not Prefixes.exists(pref) Then Prefixes.Add pref, 0
    Prefixes(pref) = Prefixes(pref) + 1

    Generate = pref + Right(CStr(Prefixes(pref) + 10 ^ d), d)

End Function

'Поиск в словаре или генерация литера
Function GetLiter(ByVal buyer As String, lit As String) As String
    If lit <> "" Then
        GetLiter = UCase(lit)
    Else
        GetLiter = UCase(Left(buyer, 1))
    End If
End Function

'Поиск с вловаре или генерация кода
Function GetCode(ByVal buyer As String, dateR As Date, ByRef d As Byte)
    GetCode = Codes(buyer)
    d = 5
    If GetCode = "" Then
        GetCode = Right(CStr(Year(recDate)), 2) + _
            Right(CStr(Month(recDate) + 100), 2) + _
            Right(CStr(Day(recDate) + 100), 2)
        d = 3
    End If
End Function

'Проверка правильности префикса (сравнивается существующий уид и наименование продавца)
Function CheckPrefix(UID As String, ByVal dateS As Date, buyer As String) As Boolean
    Dim d As Byte
    pref = Liters(buyer) + GetCode(buyer, dateS, d)
    CheckPrefix = pref = Left(UID, Len(pref))
End Function