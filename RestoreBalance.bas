Attribute VB_Name = "RestoreBalance"
'Восстановление остатков
Public Sub Run()
    Main.Init
    Set selIndexes = CreateObject("Scripting.Dictionary")
    Set qrtIndexes = CreateObject("Scripting.Dictionary")
    Range(DIC.Cells(firstDic, cPFact), DIC.Cells(maxRow, cPFact + quartCount - 1)).Clear
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
            s = 0
            For j = 12 To 14
                If DAT.Cells(i, j).text <> "" Then s = s + DAT.Cells(i, j)
            Next
            sl = selIndexes(DAT.Cells(i, 5).text)
            kv = Kvartal(DAT.Cells(i, 2))
            kvin = cPFact + qrtIndexes(kv)
            DIC.Cells(sl, kvin) = DIC.Cells(sl, kvin) + s
        End If
        i = i + 1
    Loop
End Sub