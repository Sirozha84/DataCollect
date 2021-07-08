Attribute VB_Name = "CollectSale"
'Last change: 08.07.2021 21:27

Dim LastRec As Long
Dim curFile As String
Dim curCode As String

'Запуск процесса сбора данных
Sub Run()
    
    Message "Подготовка..."
    Numerator.Init
    Log.Init
    Verify.Init
    
    'Получаем коллекцию файлов
    Set files = Source.getFiles(DirImportSale, True)
    
    'Обрабатываем список файлов
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
    
    Values.CreateReport
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
'3 - нет кода
'4 - версия файла не поддерживается
'5 не использовать, это код дубликата (провеяется в Source)
'6 - файл уже открыт
'7 - проблема с записью
'Сообщения об ошибках по этим кодам пишется в Log
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
    SetProtect SRC
    On Error GoTo 0
    
    'Проверка версии шаблона реестра
    ver = SRC.Cells(2, 1).text
    If ver <> tmpVersion Then
        AddFile = 4
        impBook.Close False
        Exit Function
    End If
    curFile = file
    curCode = SRC.Cells(1, 1)
    If curCode <> "" Then
        
        'Очищаем предыдущие строки без номеров
        i = firstDat
        Do While DAT.Cells(i, cAccept) <> ""
            If DAT.Cells(i, cUIN) = "" And DAT.Cells(i, cCode) = curCode Then
                DAT.Rows(i).Delete
            Else
                i = i + 1
            End If
        Loop
    
        'Индексируем существующие записи
        Set Indexes = CreateObject("Scripting.Dictionary")
        i = firstDat
        Do While DAT.Cells(i, cAccept) <> ""
            UID = DAT.Cells(i, cUIN)
            If UID <> "" Then Indexes.Add UID, i
            i = i + 1
        Loop
        LastRec = i
    
        'Обрабатываем строки исходника
        Set resUIDs = CreateObject("Scripting.Dictionary")
        i = firstSrc
        Do While NotEmpty(i)
            UID = SRC.Cells(i, 1)
            'Строка уже есть (наверное)
            If UID <> "" Then
                
                ind = Indexes(UID)
                If ind <> Empty Then
                    
                    'И строка действительно есть, обновляем данные
                    If Not copyRecord(ind, i, True) Then errors = True
                    
                    'Данные не обновлены
                    stat = DAT.Cells(ind, cStatus).text
                    If stat = "0" Then
                        DAT.Cells(ind, cCom) = "Данные аннулированы!"
                        DAT.Cells(ind, cCom).Interior.Color = colRed
                        SRC.Cells(i, cCom) = "Данные аннулированы!"
                        SRC.Cells(i, cCom).Interior.Color = colRed
                    End If
                    If stat = "2" Then
                        DAT.Cells(ind, cCom) = "Данные зафиксированы!"
                        DAT.Cells(ind, cCom).Interior.Color = colGreen
                        SRC.Cells(i, cCom) = "Данные зафиксированы!"
                        SRC.Cells(i, cCom).Interior.Color = colGreen
                    End If
                    
                Else
                    'А вот и нет, такой строки нет, стоит непонятный UID, которого у нас нет
                    UID = ""
                End If
            End If
            'Новая строка
            If UID = "" Then If Not copyRecord(LastRec, i, False) Then errors = True
            rUID = SRC.Cells(i, 1).text
            
            'Составляем словарь resUIDs - все номера, которые есть в реестре
            'Далее в сборе ищем все записи по коду из этого реестра, которые отсутствуют в этом
            'словаре, их считаем удалёнными.
            'Если в реестре будет два одинаковых номера, то тут будет ошибка!
            On Error Resume Next
            If rUID <> "" Then resUIDs.Add rUID, 1
            
            i = i + 1
        Loop
        
        'Проверяем исходник на удалённые записи (на предмет пропавших УИНов)
        i = firstDat
        Do While DAT.Cells(i, cAccept) <> ""
            UID = DAT.Cells(i, cUIN).text
            If UID <> "" And DAT.Cells(i, cCode) = curCode Then
                If resUIDs(UID) = Empty Then
                    DAT.Cells(i, cCom) = "Данные удалены заказчиком (вместе с УИН)"
                    DAT.Cells(i, cCom).Interior.Color = colYellow
                    DAT.Cells(i, cAccept) = "lost"
                    AddFile = 2
                End If
            End If
            i = i + 1
        Loop
        
    Else
        AddFile = 3
    End If
    
    'Завершение
    On Error GoTo er
    impBook.Close True
    Numerator.Save
    Application.ScreenUpdating = True
    DoEvents    'Не помню для чего это, вроде как без этого всё зависало, а потом открывалось много окон
    If errors Then AddFile = 2
    Exit Function

er:
    AddFile = 1

End Function

'Проверка на пустую строку
'Возвращает True, если строка в исходнике не пустая
Function NotEmpty(ByVal i As Long) As Boolean
    NotEmpty = False
    For j = 1 To 15
        txt = SRC.Cells(i, j).text
        If txt <> "" And txt <> "#Н/Д" Then NotEmpty = True: Exit For
    Next
End Function

'Копирование записи. Возвращает True, если запись принялась без ошибок
'Di - строка в данных
'Si - строка в исходниках
'refresh - true, если обновление данных (проверять что поменялось)
Function copyRecord(ByVal Di As Long, ByVal Si As Long, refresh As Boolean) As Boolean
    
    stat = DAT.Cells(Di, cStatus).text
    If stat = "0" Then
        copyRecord = False
        Exit Function
    End If
    
    SetFormates Di
    SRC.Cells(Si, 1).ClearFormats
    
    'Запись зафиксирована, возвращаем изменение из собранных данных в реестр
    If stat = "2" Then
        For j = 2 To 14
            CheckChanges Di, Si, j
            SRC.Cells(Si, j) = DAT.Cells(Di, j)
        Next
        copyRecord = True
        Exit Function
    End If
    
    If refresh And DAT.Cells(Di, cAccept) = "OK" Then
        oldSum = 0
        For i = 12 To 14
            If DAT.Cells(Di, i) <> "" Then oldSum = oldSum + DAT.Cells(Di, i)
        Next
        RestoreBalance DAT.Cells(Di, cDates), DAT.Cells(Di, cSellINN).text, oldSum
    End If
    
    'Копирование записей с проверкой на изменение
    For j = 2 To 14
        If j <> cSeller Then
            CheckChanges Di, Si, j
            If Not IsError(SRC.Cells(Si, j)) Then DAT.Cells(Di, j) = SRC.Cells(Si, j)
            If j = cSellINN Then DAT.Cells(Di, cSellINN) = Left(DAT.Cells(Di, cSellINN).text, 10)
        Else
            s = selIndexes(DAT.Cells(Di, cSellINN).text)
            If s <> Empty Then
                DAT.Cells(Di, 6) = DIC.Cells(s, 1)
            Else
                AddCom ("ИНН не найден в справочнике")
            End If
        End If
    Next
    DAT.Cells(Di, cFile) = curFile
    DAT.Cells(Di, cCode) = curCode
    DAT.Cells(Di, cAccept) = "fail" 'По умолчанию будем считать строку не верной
    
    'Проверка на удалённую запись (если это обновление и строка с датой пустая)
    If refresh And SRC.Cells(Si, cDates).text = "" Then
        SRC.Cells(Si, 1).Font.Color = colWhite
        SRC.Cells(Si, cCom) = "Данные удалены заказчиком"
        SRC.Cells(Si, cCom).Interior.Color = colYellow
        DAT.Cells(Di, cCom) = "Данные удалены заказчиком"
        DAT.Cells(Di, cCom).Interior.Color = colYellow
        DAT.Cells(Di, cAccept) = "lost"
        copyRecord = True
        Exit Function
        'Дальнейшие действия в этом случае не требуются, выходим...
    End If
    
    copyRecord = Verify.Verify(Di, Si, oldINN, oldSum)
    
    'Если нужно, присваиваем записи новый номер
    If copyRecord Then
        Dim needNum As Boolean
        If refresh Then
            needNum = Not Numerator.CheckPrefix(DAT.Cells(Di, 1).text, _
                DAT.Cells(Di, 2), DAT.Cells(Di, cSellINN).text)
        Else
            needNum = True
        End If
        If needNum Then
            n = Numerator.Generate(DAT.Cells(Di, 2), DAT.Cells(Di, cSellINN).text)
            DAT.Cells(Di, cUIN).NumberFormat = "@"
            DAT.Cells(Di, cUIN) = n
            DAT.Cells(Di, cDateCol) = DateTime.Now
            SRC.Cells(Si, 1).NumberFormat = "@"
            SRC.Cells(Si, 1) = n
        End If
        DAT.Cells(Di, cAccept) = "OK"
    End If
    
    If Not refresh Then LastRec = LastRec + 1
    If DAT.Cells(Di, cStatus).text = "" Then DAT.Cells(Di, cStatus) = 1
    
End Function

Sub SetFormates(ByVal i As Long)
    DAT.Cells(i, 2).NumberFormat = "dd.MM.yyyy"
    DAT.Cells(i, 7).NumberFormat = "### ### ##0.00"
    For j = 9 To 11
        DAT.Cells(i, j).NumberFormat = "### ### ##0.00"
    Next
    For j = 12 To 14
        DAT.Cells(i, j).NumberFormat = "### ### ##0.00"
    Next
End Sub

'Отслеживание изменений и пометка их цветом
'Di - строка в данных
'Si - строка в исходниках
Sub CheckChanges(ByVal Di As Long, ByVal Si As Long, ByVal j As Long)
    
    'Сброс формата
    DAT.Cells(Di, j).Interior.Color = colWhite
    If j = 2 Or j = 4 Or j = 6 Or j = 7 Or j = 8 Then
        SRC.Cells(Si, j).Interior.Color = colYellow
    Else
        SRC.Cells(Si, j).Interior.Color = colWhite
    End If
    
    'Подсветка, если есть разница
    If DAT.Cells(Di, j).text <> SRC.Cells(Si, j).text Then
        DAT.Cells(Di, j).Interior.Color = colBlue
        SRC.Cells(Si, j).Interior.Color = colBlue
    End If

End Sub

'******************** End of File ********************