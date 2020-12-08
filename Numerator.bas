Attribute VB_Name = "Numerator"
Dim Nums As Object
Dim dTab As Variant

'Инициализация словаря
Sub Init()
    Call NewTab(tabNum, False)
    Set dTab = Sheets(tabNum)
    dTab.Cells(1, 1) = "Внимание! Здесь находится служебная информация. Ручное редактирование не рекоммендуется."
    dTab.Cells(3, 1) = "Префикс"
    dTab.Cells(3, 2) = "Номер"
    Range(dTab.Cells(1, 1), dTab.Cells(3, 100)).Interior.Color = RGB(214, 214, 214)
    Set Nums = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = 4
    Do While dTab.Cells(i, 1) <> ""
        pref = dTab.Cells(i, 1)
        Nums.Add pref, dTab.Cells(i, 2)
        i = i + 1
    Loop
End Sub

'Сохранение словаря на страницу
Sub Save()
    Dim i As Long
    i = 4
    For Each Key In Nums.keys
        dTab.Cells(i, 1) = Key
        dTab.Cells(i, 2) = Nums(Key)
        i = i + 1
    Next
End Sub

'Очистка словаря
Sub Clear()
    On Error GoTo er
    Sheets(tabNum).Cells.Clear
er:
End Sub

'Генерация уникального номера
Function Generate(dat As Date, Buyer As String) As String
    pref = UCase(Left(Buyer, 1)) + Right(CStr(Year(dat)), 2) + CStr(Month(dat)) + CStr(Day(dat))
    If Not Nums.exists(pref) Then Nums.Add pref, 0
    Nums(pref) = Nums(pref) + 1
    Generate = pref + Right(CStr(Nums(pref) + 1000), 3)
End Function