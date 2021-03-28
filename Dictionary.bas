Attribute VB_Name = "Dictionary"
'Последняя правка: 28.03.2021 13:15

'Чтение словаря продавцов
Public Sub Init()
    Set selIndexes = CreateObject("Scripting.Dictionary")
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        inn = DIC.Cells(i, cINN).text
        selIndexes(inn) = i
        i = i + 1
    Loop
End Sub

'Имя файла по ИНН продавца
Function SellFileName(inn) As String
    ind = selIndexes(inn)
    If ind <> Empty Then SellFileName = inn + "-" + DIC.Cells(ind, 1)
End Function

'Проверка на совпадение имени и ИНН продавца, если ИНН не существует - добавление нового
'Возвращает false, если имя не совпадает с тем что в
Function CorrectSaler(ByVal inn As String, ByVal name As String) As Boolean
    CorrectSaler = True
    If selIndexes(inn) = 0 Then
        i = firstDic
        Do While DIC.Cells(i, cINN) <> ""
            i = i + 1
        Loop
        lastdic = i
        selIndexes(inn) = lastdic
        DIC.Cells(lastdic, cSellerName) = DTL.Cells(i, clSaleName)
        DIC.Cells(lastdic, cINN).NumberFormat = "@"
        DIC.Cells(lastdic, cINN) = inn
        For j = 0 To quartCount - 1
            DIC.Cells(lastdic, cLimits + j).NumberFormat = "### ### ##0.00"
            DIC.Cells(lastdic, cLimits + j).FormulaR1C1 = _
                    "=SUM(RC[" + CStr(24 + j) + "]:RC[" + CStr(47 - j) + "])-" + _
                    "SUM(RC[12]:RC[" + CStr(23 - j) + "])"
        Next
        lastdic = lastdic + 1
    Else
        If DIC.Cells(selIndexes(inn), 1) <> name Then CorrectSaler = False
    End If
End Function

'******************** End of File ********************