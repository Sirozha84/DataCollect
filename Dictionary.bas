Attribute VB_Name = "Dictionary"
'Последняя правка: 10.04.2021 16:58

'Чтение словаря продавцов
Public Sub Init()
    Set selIndexes = CreateObject("Scripting.Dictionary")
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        INN = DIC.Cells(i, cINN).text
        selIndexes(INN) = i
        i = i + 1
    Loop
End Sub

'Имя файла по ИНН продавца
Function SellFileName(INN) As String
    ind = selIndexes(INN)
    If ind <> Empty Then SellFileName = INN + "-" + DIC.Cells(ind, 1)
End Function

'Проверка на совпадение имени и ИНН продавца, если ИНН не существует - добавление нового
'Возвращает false, если имя не совпадает с тем что в
Function CorrectSaler(ByVal INN As String, ByVal name As String) As Boolean
    CorrectSaler = True
    If selIndexes(INN) = 0 Then
        i = firstDic
        Do While DIC.Cells(i, cINN) <> ""
            i = i + 1
        Loop
        lastdic = i
        selIndexes(INN) = lastdic
        DIC.Cells(lastdic, cSellerName) = name
        DIC.Cells(lastdic, cINN).NumberFormat = "@"
        DIC.Cells(lastdic, cINN) = INN
        For j = 0 To quartCount - 1
            DIC.Cells(lastdic, cLimits + j).NumberFormat = "### ### ##0.00"
            DIC.Cells(lastdic, cLimits + j).FormulaR1C1 = _
                    "=SUM(RC[" + CStr(24 + j) + "]:RC[" + CStr(47 - j) + "])-" + _
                    "SUM(RC[12]:RC[" + CStr(23 - j) + "])"
        Next
        lastdic = lastdic + 1
    Else
        If DIC.Cells(selIndexes(INN), 1) <> name Then CorrectSaler = False
    End If
End Function

'******************** End of File ********************