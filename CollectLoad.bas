Attribute VB_Name = "CollectLoad"
'Last change: 12.07.2021 15:44

Dim LastRec As Long
Dim curFile As String
Dim curMark As String
Dim curProv As String
Dim curProvINN As String
Dim UINs As Object

'Колонки данных в обрабатываемых файлах
Dim scKVO As Integer
Dim scND As Integer
Dim scSeller As Integer
Dim scSellerINN As Integer
Dim scPrice As Integer
Dim scPWN20 As Integer
Dim scPWN18 As Integer
Dim scPWN10 As Integer
Dim scNDS20 As Integer
Dim scNDS18 As Integer
Dim scNDS10 As Integer

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
                Si = IndexByINN(DTL.Cells(i, clSaleINN).text)
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
'7 - нет маркера, или он не верный
'8 - поля не распознаны
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
        AddFile = 7
        impBook.Close False
        Exit Function
    End If
    
    'Определяемся с типом файла
    If LCase(Left(SRC.Cells(2, 1), 5)) = "книга" Then ftyp = "b"
    If LCase(Left(SRC.Cells(2, 22), 5)) = "книга" Then ftyp = "b"
    If LCase(Left(SRC.Cells(2, 1), 6)) = "журнал" Then ftyp = "j"
    If LCase(Left(SRC.Cells(2, 27), 6)) = "журнал" Then ftyp = "j"
    
    'Читаем данные
    c = 60  'Колонка с обратной связью
    If ftyp = "j" Then
        curProv = Split(SRC.Cells(4, 2).text, ": ")(1)
        curProvINN = Right(SRC.Cells(5, 2).text, 20)
        i = 13
        'Чтение данных о закупках
        Do While SRC.Cells(i, 2).text <> ""
            If UINs(SRC.Cells(i, c).text) = "" And SRC.Cells(i, 2).text <> "005" Then
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
        If Not SBFieldRecognition Then AddFile = 8: GoTo ex
        curProv = Replace(SRC.Cells(4, 1).text, "Продавец  ", "")
        curProvINN = Right(SRC.Cells(5, 1).text, 20)
        i = 13
        'Чтение данных о закупках
        Do While SRC.Cells(i, 2).text <> ""
            If UINs(SRC.Cells(i, c).text) = "" And SRC.Cells(i, 1).text <> "005" Then
                If Not copyRecordSB(i) Then
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
    If ftyp = "" Then AddFile = 8
    If errors Then AddFile = 2

ex:
    'Завершение
    On Error GoTo er
    impBook.Close True
    Application.ScreenUpdating = True
    DoEvents    'Не помню для чего это, вроде как без этого всё зависало, а потом открывалось много окон
    Exit Function

er:
    AddFile = 1
    
End Function

'Копирование записи из журнала. Возвращает True, если запись принялась без ошибок
'Si - строка в исходниках
Function copyRecordZH(ByVal Si As Long) As Boolean
    
    SetFormates LastRec
    On Error GoTo er
    DTL.Cells(LastRec, clMark) = curMark
    DTL.Cells(LastRec, clKVO) = SRC.Cells(Si, 4)
    nd = SRC.Cells(Si, 6).text
    DTL.Cells(LastRec, clNum) = NumFromND(nd)
    DTL.Cells(LastRec, clDate).NumberFormat = "dd.MM.yyyy"
    DTL.Cells(LastRec, clDate) = Right(nd, 10)
    DTL.Cells(LastRec, clProvINN) = curProvINN
    DTL.Cells(LastRec, clProvName) = curProv
    DTL.Cells(LastRec, clSaleINN) = Left(SRC.Cells(Si, 15), 10)
    DTL.Cells(LastRec, clSaleName) = SRC.Cells(Si, 13)
    DTL.Cells(LastRec, clPrice) = SRC.Cells(Si, 27)
    DTL.Cells(LastRec, clNDS) = SRC.Cells(Si, 29)
    
    copyRecordZH = VerifyLoad(LastRec)
    AddFormuls
    Exit Function
    
er:
    copyRecordZH = False
    
End Function

'Копирование записи из книги продаж. Возвращает True, если запись принялась без ошибок
'Si - строка в исходниках
Function copyRecordSB(ByVal Si As Long) As Boolean
    
    SetFormates LastRec
    On Error GoTo er
    DTL.Cells(LastRec, clMark) = curMark
    kvo = SRC.Cells(Si, scKVO)
    DTL.Cells(LastRec, clKVO) = kvo
    DTL.Cells(LastRec, clSaleINN) = Left(SRC.Cells(Si, scSellerINN), 10)
    DTL.Cells(LastRec, clSaleName) = SRC.Cells(Si, scSeller)
    If kvo = "02" Then
        DTL.Cells(LastRec, clKVO) = "22"
        DTL.Cells(LastRec, clSaleINN) = curProvINN
        DTL.Cells(LastRec, clSaleName) = curProv
        kvochange = True
    End If
    nd = SRC.Cells(Si, 3).text
    DTL.Cells(LastRec, clNum) = NumFromND(nd)
    DTL.Cells(LastRec, clDate).NumberFormat = "dd.MM.yyyy"
    DTL.Cells(LastRec, clDate) = Right(nd, 10)
    DTL.Cells(LastRec, clProvINN) = curProvINN
    DTL.Cells(LastRec, clProvName) = curProv
    DTL.Cells(LastRec, clPrice) = SRC.Cells(Si, scPrice)
    If scPWN20 <> 0 Then
        DTL.Cells(LastRec, clPrice + 1) = SRC.Cells(Si, scPWN20)
        DTL.Cells(LastRec, clPrice + 2) = SRC.Cells(Si, scPWN18)
        DTL.Cells(LastRec, clPrice + 3) = SRC.Cells(Si, scPWN10)
        DTL.Cells(LastRec, clNDS) = _
            WorksheetFunction.Sum(Range(SRC.Cells(Si, scNDS20), SRC.Cells(Si, scNDS10)))
    Else
        DTL.Cells(LastRec, clPrice + 2) = SRC.Cells(Si, scPWN18)
        DTL.Cells(LastRec, clPrice + 3) = SRC.Cells(Si, scPWN10)
        DTL.Cells(LastRec, clNDS) = _
            WorksheetFunction.Sum(Range(SRC.Cells(Si, scNDS18), SRC.Cells(Si, scNDS10)))
    End If
    copyRecordSB = VerifyLoad(LastRec)
    
    'КВО менялся с 02 на 22, делаем связанные с этим событием действия
    If kvochange Then
        i = IndexByINN(DTL.Cells(LastRec, clSaleINN).text)
        j = DateToQIndex(DTL.Cells(LastRec, clDate))
        If j >= 0 Then DIC.Cells(i, cSaleProtect + j) = "Да"
    End If
    
    AddFormuls
    Exit Function
    
er:
    copyRecordSB = False
    
End Function

'Установка форматов полей
Sub SetFormates(ByVal i As Integer)
    DTL.Cells(LastRec, clKVO).NumberFormat = "@"
    DTL.Cells(LastRec, clNum).NumberFormat = "@"
    DTL.Cells(LastRec, clProvINN).NumberFormat = "@"
    DTL.Cells(LastRec, clSaleINN).NumberFormat = "@"
    For i = 0 To 4
        DTL.Cells(LastRec, clPrice + i).NumberFormat = numFormat
    Next
End Sub

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

'Отделение номера от "Номер и дата"
Function NumFromND(ByVal nd As String) As String
    ss = Split(nd, " от")
    If UBound(ss) > 0 Then NumFromND = ss(0)
    ss = Split(nd, ";")
    If UBound(ss) > 0 Then NumFromND = ss(0)
End Function

'Распознавание колонок в книге продаж
Function SBFieldRecognition() As Boolean
    
    scKVO = 0
    scND = 0
    scSeller = 0
    scSellerINN = 0
    scPrice = 0
    scPWN20 = 0
    scPWN18 = 0
    scPWN10 = 0
    scNDS20 = 0
    scNDS18 = 0
    scNDS10 = 0

    For i = 1 To 50
        If SRC.Cells(8, i) = "Код вида опера-ции" Then scKVO = i
        If Left(SRC.Cells(8, i), 12) = "Номер и дата" Then scND = i
        If Left(SRC.Cells(8, i), 12) = "Наименование" Then scSeller = i
        If Left(SRC.Cells(8, i), 7) = "ИНН/КПП" Then scSellerINN = i
        If Left(SRC.Cells(11, i), 10) = "в рублях и" Then scPrice = i
        If SRC.Cells(11, i) = "20 процентов" Then If scPWN20 = 0 Then scPWN20 = i Else scNDS20 = i
        If SRC.Cells(11, i) = "18 процентов" Then If scPWN18 = 0 Then scPWN18 = i Else scNDS18 = i
        If SRC.Cells(11, i) = "10 процентов" Then If scPWN10 = 0 Then scPWN10 = i Else scNDS10 = i
    Next
    
    SBFieldRecognition = _
        scKVO <> 0 And _
        scND <> 0 And _
        scSeller <> 0 And _
        scSellerINN <> 0 And _
        scPrice <> 0 And _
        scPWN18 <> 0 And _
        scPWN10 <> 0 And _
        scNDS18 <> 0 And _
        scNDS10 <> 0

End Function

'******************** End of File ********************