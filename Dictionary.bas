Attribute VB_Name = "Dictionary"
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

'******************** End of File ********************