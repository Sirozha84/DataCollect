Attribute VB_Name = "Numerator"
Dim Prefixes As Object  'Словарь префиксов
Dim Liters As Object    'Словарь литеров
Dim Codes As Object     'Словарь кодов

'Инициализация словаря
Sub Init()
    
    'Загрузка словаря нумератора
    Range(NUM.Cells(1, 1), NUM.Cells(3, 100)).Interior.Color = RGB(214, 214, 214)
    Set Prefixes = CreateObject("Scripting.Dictionary")
    i = firstNum
    Do While NUM.Cells(i, 1).text <> ""
        pref = NUM.Cells(i, 1).text
        Prefixes.Add pref, NUM.Cells(i, 2)
        i = i + 1
    Loop
    
    'Загрузка префиксов и кодов
    Set Liters = CreateObject("Scripting.Dictionary")
    Set Codes = CreateObject("Scripting.Dictionary")
    i = firstDic
    Do While DIC.Cells(i, cINN) <> ""
        seller = DIC.Cells(i, cINN).text
        Liters.Add seller, GetLiter(seller, DIC.Cells(i, cPLiter).text)
        Codes.Add seller, DIC.Cells(i, cPCode).text
        i = i + 1
    Loop

End Sub

'Сохранение словаря на страницу
Sub Save()
    i = firstNum
    For Each Key In Prefixes.keys
        NUM.Cells(i, 1) = "'" + Key
        NUM.Cells(i, 2) = Prefixes(Key)
        i = i + 1
    Next
    'На всякий случай делаем следующую строку пустой
    NUM.Cells(i, 1) = ""
    NUM.Cells(i, 2) = ""
End Sub

'Генерация уникального номера
Public Function Generate(recDate As Date, seller As String) As String
    
    lit = Liters(seller)
    Dim d As Byte
    cod = GetCode(seller, recDate, d)
    
    pref = lit + cod
    If Not Prefixes.exists(pref) Then Prefixes.Add pref, 0
    Prefixes(pref) = Prefixes(pref) + 1

    Generate = pref + Right(CStr(Prefixes(pref) + 10 ^ d), d)

End Function

'Поиск в словаре или генерация литера
Function GetLiter(ByVal seller As String, lit As String) As String
    If lit <> "" Then
        GetLiter = UCase(lit)
    Else
        GetLiter = UCase(Left(seller, 1))
    End If
End Function

'Поиск в словаре или генерация кода
'd - количество цифр номера
Function GetCode(ByVal seller As String, dateR As Date, ByRef d As Byte)
    GetCode = Codes(seller)
    d = 5
    If GetCode = "" Then
        GetCode = Right(CStr(Year(dateR)), 2) + _
            Right(CStr(Month(dateR) + 100), 2) + _
            Right(CStr(Day(dateR) + 100), 2)
        d = 3
    End If
End Function

'Проверка правильности префикса (сравнивается существующий уид и наименование продавца)
Function CheckPrefix(UID As String, ByVal dateS As Date, seller As String) As Boolean
    Dim d As Byte
    pref = Liters(seller) + GetCode(seller, dateS, d)
    CheckPrefix = pref = Left(UID, Len(pref))
End Function