Attribute VB_Name = "Values"
'Последняя правка: 30.06.2021 20:57

'Формирование отчёта "Объёмы продаж"
Public Sub CreateReport()

    Message "Формирование отчёта по объёмам продаж..."
    Dim i As Long
    
    'Собираем словари наименований и ИНН компаний
    Set summPrice = CreateObject("Scripting.Dictionary")
    Set summNDS = CreateObject("Scripting.Dictionary")
    Set buyList = CreateObject("Scripting.Dictionary")
    Set sellList = CreateObject("Scripting.Dictionary")
    i = firstDat
    Do While DAT.Cells(i, cAccept) <> ""
        If DAT.Cells(i, cAccept) = "OK" Then
            cod = DAT.Cells(i, cCode).text
            q = DateToQIndex(DAT.Cells(i, cDates))
            BUY = DAT.Cells(i, cBuyINN).text
            sell = DAT.Cells(i, cSellINN).text
            nds = WorksheetFunction.Sum(Range(DAT.Cells(i, 12), DAT.Cells(i, 14)))
            ID = cod + "!" + CStr(q) + "!" + sell + "!" + BUY
            summPrice(ID) = summPrice(ID) + DAT.Cells(i, cPrice)
            summNDS(ID) = summNDS(ID) + nds
            buyList(DAT.Cells(i, cBuyINN).text) = DAT.Cells(i, cBuyer).text
        End If
        i = i + 1
    Loop
    
    'Собираем словарь ИНН и статуса продавца
    Dim statList As Variant
    Set statList = CreateObject("Scripting.Dictionary")
    i = firstDic
    Do While DIC.Cells(i, cINN) <> ""
        sellList(DIC.Cells(i, cINN).text) = DIC.Cells(i, cSellerName).text
        statList(DIC.Cells(i, cINN).text) = DIC.Cells(i, cPStat).text
        i = i + 1
    Loop
    
    'Собираем словари Клиента и Формы
    Set clients = CreateObject("Scripting.Dictionary")
    Set brokers = CreateObject("Scripting.Dictionary")
    Set templates = CreateObject("Scripting.Dictionary")
    i = firstTempl
    Do While TMP.Cells(i, cTCode).text <> ""
        clients(TMP.Cells(i, cTCode).text) = TMP.Cells(i, cTClient).text
        brokers(TMP.Cells(i, cTCode).text) = TMP.Cells(i, cTBroker).text
        templates(TMP.Cells(i, cTCode).text) = TMP.Cells(i, cTForm).text
        i = i + 1
    Loop
    
    'Подготовка листа
    cols = 9
    hat = firstValues - 1
    Range(VAL.Cells(firstValues, 1), VAL.Cells(maxRow, cols)).Clear
    VAL.Cells(hat, 1) = "Клиент"
    VAL.Cells(hat, 2) = "Посредник"
    VAL.Cells(hat, 3) = "Форма"
    VAL.Cells(hat, 4) = "Квартал"
    VAL.Cells(hat, 5) = "Продавец"
    VAL.Cells(hat, 6) = "Статус"
    VAL.Cells(hat, 7) = "Покупатель"
    VAL.Cells(hat, 8) = "Стоимость с НДС"
    VAL.Cells(hat, 9) = "НДС"
    Range(VAL.Cells(hat, 1), VAL.Cells(hat, cols)).Interior.Color = colGray
    Range(VAL.Cells(hat, 1), VAL.Cells(hat, cols)).Borders.Weight = 2
    
    'Формирование отчёта
    i = firstValues
    Dim s As Variant
    Dim SEL As Variant
    For Each SEL In summPrice
        s = Split(SEL, "!")
        VAL.Cells(i, 1) = clients(s(0))
        VAL.Cells(i, 2) = brokers(s(0))
        VAL.Cells(i, 3) = templates(s(0))
        VAL.Cells(i, 4) = s(1)
        VAL.Cells(i, 5) = sellList(s(2)) + " (" + s(2) + ")"
        VAL.Cells(i, 6) = statList(s(2))
        VAL.Cells(i, 7) = buyList(s(3)) + " (" + s(3) + ")"
        VAL.Cells(i, 8).NumberFormat = "### ### ##0.00"
        VAL.Cells(i, 8) = summPrice(SEL)
        VAL.Cells(i, 9).NumberFormat = "### ### ##0.00"
        VAL.Cells(i, 9) = summNDS(SEL)
        i = i + 1
    Next
    On Error Resume Next
    Range(VAL.Cells(hat, 1), VAL.Cells(hat, cols)).Rows.AutoFilter
    
    'Сводная таблица
    VLS.Cells.Clear
    VLS.PivotTableWizard SourceType:=xlDatabase, _
        SourceData:=Range(VAL.Cells(hat, 1), VAL.Cells(i - 1, cols)), _
        TableDestination:=VLS.Cells(1, 1)
    
    Message "Готово!"
    
End Sub

'******************** End of File ********************