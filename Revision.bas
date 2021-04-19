Attribute VB_Name = "Revision"
'Последняя правка: 19.04.2021 20:51

'Ревизия остатков
Public Sub Run()
    
    Message "Ревизия остатков..."
    
    'Формирование
    Set selIndexes = CreateObject("Scripting.Dictionary")
    Set qrtIndexes = CreateObject("Scripting.Dictionary")
    Range(DIC.Cells(firstDic, cPRev), DIC.Cells(maxRow, cPRev + quartCount - 1)).Clear
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        cmp = DIC.Cells(i, cINN).text
        selIndexes(cmp) = i
        Range(DIC.Cells(i, cPFact), DIC.Cells(i, cPFact + quartCount - 1)).NumberFormat = "### ### ##0.00"
        i = i + 1
    Loop
    For i = 0 To quartCount - 1
        qrtIndexes(IndexToQuartal(i)) = i
    Next
    
    i = firstDat
    Do While DAT.Cells(i, cAccept) <> ""
        If DAT.Cells(i, cAccept) = "OK" Then
            s = WorksheetFunction.Sum(Range(DAT.Cells(i, 12), DAT.Cells(i, 14)))
            sl = selIndexes(DAT.Cells(i, cSellINN).text)
            If sl = Empty Then
                MsgBox "Произошла неожиданная ошибка:" + Chr(10) + "Продавец " + DAT.Cells(i, cSeller) + _
                        " c ИНН " + DAT.Cells(i, cSellINN).text + " отсутствует в справочнике!"
                End
            End If
            q = DateToQIndex(DAT.Cells(i, cDates))
            If q >= 0 Then DIC.Cells(sl, cPRev + q) = DIC.Cells(sl, cPRev + q) + s
        End If
        i = i + 1
    Loop
    
    'Проверка с текущими значениями
    i = firstDic
    Do While DIC.Cells(i, 1) <> ""
        For j = 0 To quartCount - 1
            If DIC.Cells(i, cPFact + j) = DIC.Cells(i, cPRev + j) Then
                DIC.Cells(i, cPRev + j).Interior.Color = colGreen
            Else
                DIC.Cells(i, cPRev + j).Interior.Color = colRed
            End If
        Next
        i = i + 1
    Loop
    
    Message "Готово"
    
End Sub

'******************** End of File ********************