Attribute VB_Name = "Dictionary"
'Последняя правка: 27.06.2021 15:20

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
'Возвращает новое имя для нового продавца, или имя из справочника, если он там есть
Function CorrectSaler(ByVal inn As String, ByVal name As String) As String
    If Len(inn) > 10 Then inn = Left(inn, 10)
    CorrectSaler = True
    If selIndexes(inn) = 0 Then
        i = firstDic
        Do While DIC.Cells(i, cINN) <> ""
            i = i + 1
        Loop
        lastdic = i
        selIndexes(inn) = lastdic
        DIC.Cells(lastdic, cSellerName) = name
        DIC.Cells(lastdic, cINN).NumberFormat = "@"
        DIC.Cells(lastdic, cINN) = inn
        
        For j = 0 To quartCount - 1
            DIC.Cells(lastdic, cLimits + j).NumberFormat = numFormat
            DIC.Cells(lastdic, cLimits + j).FormulaR1C1 = _
                    "=SUM(RC[" + CStr(quartCount * 2 + j) + "]:RC[" + CStr(quartCount * 4 - 1 - j) + "])-" + _
                    "SUM(RC[" + CStr(quartCount) + "]:RC[" + CStr(quartCount * 2 - 1 - j) + "])"
        Next
        lastdic = lastdic + 1
    End If
    CorrectSaler = DIC.Cells(selIndexes(inn), 1)
End Function

'Возвращает индекс в справочнике по ИНН[/КПП]
Function IndexByINN(ByVal inn As String) As Integer
    If Len(inn) > 10 Then inn = Left(inn, 10)
    IndexByINN = selIndexes(inn)
End Function

'******************** End of File ********************