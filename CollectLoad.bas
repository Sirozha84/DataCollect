Attribute VB_Name = "CollectLoad"
'Последняя правка: 18.04.2021 13:55

Dim LastRec As Long
Dim curFile As String
Dim curMark As String
Dim curProv As String
Dim curProvINN As String
Dim UINs As Object

'Запуск процесса сбора данных
Sub Run()
    
    Message "Подготовка..."
    Dictionary.Init
    Numerator.InitLoad
    Log.Init
    
    'Очищаем сбор от старых непринятых записей
    Set UINs = CreateObject("Scripting.Dictionary")
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        If DTL.Cells(i, clAccept) = "OK" Then
            UINs(DTL.Cells(i, clUIN).text) = i
        Else
            DTL.Rows(i).Delete
            i = i - 1
        End If
        i = i + 1
    Loop
    LastRec = i
    
    'Получаем коллекцию файлов и делаем сбор из них
    Set files = Source.getFiles(DirImportLoad, False)
    n = 1
    s = 0
    e = 0
    For Each file In files
        curf = file
        If Len(curf) > 40 Then curf = "..." + Right(curf, 40)
        Message ("Обработка файла " + CStr(n) + " из " + CStr(files.Count) + " (" + curf) + ")"
        er = AddFile(file)
        If er > 0 Then
            Log.Rec file, er
            e = e + 1
        Else
            s = s + 1
        End If
        n = n + 1
    Next

    'Обновляем данные в справочнике
    Message "Расчёт квартальных лимитов"
    Range(DIC.Cells(firstDic, cPBalance), DIC.Cells(maxRow, cPBalance + quartCount * 2 - 1)).Clear
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        If DTL.Cells(i, clAccept) = "OK" Then
            Qi = DateToQIndex(DTL.Cells(i, clDate))
            If Qi >= 0 Then
                Si = selIndexes(DTL.Cells(i, clSaleINN).text)
                Qi = Qi * 2 + cPBalance
                Sum = DTL.Cells(i, clNDS)
                If DTL.Cells(i, 1).text = "З" Then Qi = Qi + 1
                DIC.Cells(Si, Qi) = DIC.Cells(Si, Qi) + Sum
            End If
            
        End If
        i = i + 1
    Loop
    
    FindDuplicates

    'Завершение
    ActiveWorkbook.Save
    Message "Готово! Файл сохранён."
    Application.DisplayAlerts = True
    
    MsgBox ("Обработка завершена!" + Chr(13) + "Файлов загруженные успешно: " + _
            CStr(s) + Chr(13) + "Файлы с ошибками: " + CStr(e))
    
End Sub

'Добавление данных из файла. Возвращает:
'0 - всё хорошо
'1 - ошибка загрузки
'2 - ошибка в данных (errors=true)
'3 - нет маркера, или он не верный
Function AddFile(ByVal file As String) As Byte
    
    'Подготовки
    Application.DisplayAlerts = False
    If Not TrySave(file) Then AddFile = 6: Exit Function
    errors = False
    Application.ScreenUpdating = False
    On Error GoTo er
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    Set SRC = impBook.Worksheets(1)
    On Error GoTo 0
        
    'Чтение и проверка маркера
    curMark = UCase(SRC.Cells(1, 1).text)
    If curMark <> "К" And curMark <> "З" Then
        AddFile = 3
        impBook.Close False
        Exit Function
    End If
    
    'Определяемся с типом файла
    If LCase(Left(SRC.Cells(2, 1), 5)) = "книга" Then ftyp = "b"
    If LCase(Left(SRC.Cells(2, 22), 5)) = "книга" Then ftyp = "b"
    If LCase(Left(SRC.Cells(2, 1), 6)) = "журнал" Then ftyp = "j"
    If LCase(Left(SRC.Cells(2, 27), 6)) = "журнал" Then ftyp = "j"
    
    'Читаем данные
    c = 50  'Колонка с обратной связью
    If ftyp = "j" Then
        curProv = Split(SRC.Cells(4, 2).text, ": ")(1)
        curProvINN = Right(SRC.Cells(5, 2).text, 20)
        i = 13  'Первая строка данных
        'Чтение данных о закупках
        Do While SRC.Cells(i, 2).text <> ""
            If UINs(SRC.Cells(i, c).text) = "" Then
                If Not copyRecordZH(i) Then
                    errors = True
                    DTL.Cells(LastRec, clAccept) = "fail"
                Else
                    DTL.Cells(LastRec, clDateCol) = DateTime.Now
                    uin = GenerateLoad
                    DTL.Cells(LastRec, clUIN) = uin
                    SRC.Cells(i, c) = uin
                    DTL.Cells(LastRec, clAccept) = "OK"
                End If
                DTL.Cells(LastRec, clFile) = file
                LastRec = LastRec + 1
            End If
            i = i + 1
        Loop
    End If
    If ftyp = "b" Then
        curProv = Replace(SRC.Cells(4, 1).text, "Продавец  ", "")
        curProvINN = Right(SRC.Cells(5, 1).text, 20)
        i = 14  'Первая строка данных
        fNDS = 20
        If Left(SRC.Cells(11, 34).text, 2) = "18" And Left(SRC.Cells(11, 40).text, 2) = "18" Then fNDS = 18
        'Чтение данных о закупках
        Do While SRC.Cells(i, 2).text <> ""
            If UINs(SRC.Cells(i, c).text) = "" Then
                If Not copyRecordSB(i, fNDS) Then
                    errors = True
                    DTL.Cells(LastRec, clAccept) = "fail"
                Else
                    DTL.Cells(LastRec, clDateCol) = DateTime.Now
                    uin = GenerateLoad
                    DTL.Cells(LastRec, clUIN) = uin
                    SRC.Cells(i, c) = uin
                    DTL.Cells(LastRec, clAccept) = "OK"
                End If
                DTL.Cells(LastRec, clFile) = file
                LastRec = LastRec + 1
            End If
            i = i + 1
        Loop
    End If
        

    'Завершение
    On Error GoTo er
    impBook.Close True
    Application.ScreenUpdating = True
    DoEvents    'Не помню для чего это, вроде как без этого всё зависало, а потом открывалось много окон
    If errors Then AddFile = 2
    Exit Function

er:
    AddFile = 1
    
End Function

'Копирование записи из журнала. Возвращает True, если запись принялась без ошибок
'Si - строка в исходниках
Function copyRecordZH(ByVal Si As Long) As Boolean
    
    On Error GoTo er
    DTL.Cells(LastRec, clMark) = curMark
    DTL.Cells(LastRec, clKVO).NumberFormat = "@"
    DTL.Cells(LastRec, clKVO) = SRC.Cells(Si, 4)                'КВО
    nd = SRC.Cells(Si, 6).text                                  'Номер и дата
    DTL.Cells(LastRec, clNum) = Split(nd, " от ")(0)
    DTL.Cells(LastRec, clDate).NumberFormat = "dd.MM.yyyy"
    DTL.Cells(LastRec, clDate) = Right(nd, 10)
    DTL.Cells(LastRec, clProvINN).NumberFormat = "@"
    DTL.Cells(LastRec, clProvINN) = curProvINN
    DTL.Cells(LastRec, clProvName) = curProv
    DTL.Cells(LastRec, clSaleINN).NumberFormat = "@"
    DTL.Cells(LastRec, clSaleINN) = Left(SRC.Cells(Si, 15), 10) 'ИНН/КПП
    DTL.Cells(LastRec, clSaleName) = SRC.Cells(Si, 13)          'Продавец
    DTL.Cells(LastRec, clPrice) = SRC.Cells(Si, 27)             'Стоимость
    DTL.Cells(LastRec, clNDS) = SRC.Cells(Si, 29)               'НДС
    copyRecordZH = VerifyLoad(LastRec)
    On Error GoTo 0
    AddFormuls
    Exit Function
    
er:
    copyRecordZH = False
    
End Function

'Копирование записи из книги продаж. Возвращает True, если запись принялась без ошибок
'Si - строка в исходниках
'fNDS - первый НДС в колонке (18 или 20)
Function copyRecordSB(ByVal Si As Long, ByVal fNDS As Integer) As Boolean
    
    On Error GoTo er
    DTL.Cells(LastRec, clMark) = curMark
    DTL.Cells(LastRec, clKVO).NumberFormat = "@"
    kvo = SRC.Cells(Si, 2)
    DTL.Cells(LastRec, clKVO) = kvo
    DTL.Cells(LastRec, clSaleINN).NumberFormat = "@"
    DTL.Cells(LastRec, clSaleINN) = Left(SRC.Cells(Si, 10), 20) 'ИННКПП
    DTL.Cells(LastRec, clSaleName) = SRC.Cells(Si, 17)          'Продавец
    If kvo = "02" Then
        DTL.Cells(LastRec, clKVO) = "22"
        DTL.Cells(LastRec, clSaleINN).NumberFormat = "@"
        DTL.Cells(LastRec, clSaleINN) = curProvINN
        DTL.Cells(LastRec, clSaleName) = curProv
        kvochange = True
    End If
    nd = SRC.Cells(Si, 3).text
    DTL.Cells(LastRec, clNum).NumberFormat = "@"
    DTL.Cells(LastRec, clNum) = Split(nd, " от")(0)
    DTL.Cells(LastRec, clDate).NumberFormat = "dd.MM.yyyy"
    DTL.Cells(LastRec, clDate) = Right(nd, 10)
    DTL.Cells(LastRec, clProvINN).NumberFormat = "@"
    DTL.Cells(LastRec, clProvINN) = curProvINN
    DTL.Cells(LastRec, clProvName) = curProv
    DTL.Cells(LastRec, clPrice) = SRC.Cells(Si, 32)             'Стоимость
    If fNDS = 18 Then
        DTL.Cells(LastRec, clPrice + 1) = SRC.Cells(Si, 34)     'Стоимость без НДС 18
        DTL.Cells(LastRec, clPrice + 2) = SRC.Cells(Si, 36)     'Стоимость без НДС 10
        DTL.Cells(LastRec, clNDS) = WorksheetFunction.Sum _
                (Range(Cells(Si, 40), Cells(Si, 43)))           'сумма НДС
    End If
    If fNDS = 20 Then
        DTL.Cells(LastRec, clPrice + 1) = SRC.Cells(Si, 34)     'Стоимость без НДС 20
        DTL.Cells(LastRec, clPrice + 2) = SRC.Cells(Si, 36)     'Стоимость без НДС 18
        DTL.Cells(LastRec, clPrice + 3) = SRC.Cells(Si, 38)     'Стоимость без НДС 10
        DTL.Cells(LastRec, clNDS) = WorksheetFunction.Sum _
                (Range(Cells(Si, 43), Cells(Si, 48)))           'сумма НДС
    End If
    copyRecordSB = VerifyLoad(LastRec)
    'КВО менялся с 02 на 22, делаем связанные с этим событием действия
    If kvochange Then
        i = selIndexes(DTL.Cells(LastRec, clSaleINN).text)
        j = DateToQIndex(DTL.Cells(LastRec, clDate))
        DIC.Cells(i, cSaleProtect + j) = "Да"
    End If
    On Error GoTo 0
    AddFormuls
    Exit Function
    
er:
    copyRecordSB = False
    
End Function

'Добавление формул и проверки данных
Sub AddFormuls()
    s = CStr(LastRec)
    DTL.Cells(LastRec, clOst).Formula = "=M" + s + "-OneCellSum(P" + s + ")"
    formul = "=R" + s + ">=0"
    With DTL.Cells(LastRec, clRasp).Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:=formul
        .ErrorMessage = "Распределённая сумма превысила сумму НДС"
    End With
End Sub

'Проверка собранных записей на отсутствие повторяющихся номеров
Sub FindDuplicates()
    Set numbers = CreateObject("Scripting.Dictionary")
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        NUM = DTL.Cells(i, clNum).text
        If numbers(NUM) = Empty Then
            numbers(NUM) = i
        Else
            io = numbers(NUM)
            DTL.Cells(io, clCom) = "Номер СФ повторяется"
            DTL.Cells(i, clCom) = "Номер СФ повторяется"
            DTL.Cells(io, clCom).Interior.Color = colRed
            DTL.Cells(i, clCom).Interior.Color = colRed
            DTL.Cells(io, clAccept) = "fail"
            DTL.Cells(i, clAccept) = "fail"
        End If
        i = i + 1
    Loop
End Sub

'******************** End of File ********************