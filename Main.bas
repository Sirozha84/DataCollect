Attribute VB_Name = "Main"
Public Const isRelease = True   'True - полноценная работа, False - режим отладки (нет вопросов, нет записи в файлы)

Public Const firstDat = 8       'Первая строка в коллекции данных
Public Const firstSrc = 5       'Первая строка в исходных файлах
Public Const firstDic = 5       'Первая строка в справочнике
Public Const cCom = 15          'Колонка для комментария
Public Const cStatus = 16       'Колонка статуса
Public Const cFile = 17         'Колонка с именем файла
Public Const cCode = 18         'Колонка с кодом файла

Public Const tabDic = "Справочник"
Public Const tabErr = "Ошибки"
Public Const tabNum = "Словарь нумератора"

Public colWhite As Long 'Цвета
Public colRed As Long
Public colGreen As Long
Public colYellow As Long

Dim dat As Variant      'Таблица с данными
Dim src As Variant      'Таблица с исходниками
Dim Indexes As Object   'Словарь индексов
Dim max As Long         'Последняя строка в данных
Dim i As Long
Dim file As Variant
Dim cod As String

'Выбор директории с данными
Sub DirSelect()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Cells(1, 3) = diag.SelectedItems(1)
End Sub

'Удаление всех данных (оставляя шапку)
Sub Clear()
    On Error GoTo er
    If isRelease Then If MsgBox("Внимание! " + Chr(10) + Chr(10) + _
        "Данная процедура очистит все собранные данные список ошибок и нумераторы. " + _
        "Уже зарегистрированные данные при повторной регистрации могут присвоить другой код." + _
        Chr(10) + Chr(10) + "Продолжить?", vbYesNo) = vbNo Then Exit Sub
    Range(Cells(firstDat, 1), Cells(1048576, 50)).Clear
    Sheets(errName).Cells.Clear
er:
    Numerator.Clear
End Sub

'Сбор данных
Sub DataCollect()
    
    Set dat = ActiveSheet
    noEmpty = (dat.Cells(firstDat, 2) <> "")
    If isRelease And noEmpty Then If MsgBox("Начинается сбор данных. Продолжить?", vbYesNo) = vbNo Then Exit Sub
    
    'Инициализация
    Message "Подготовка"
    colWhite = RGB(255, 255, 255)
    colRed = RGB(255, 192, 192)
    colGreen = RGB(192, 255, 192)
    colYellow = RGB(255, 255, 192)
    Numerator.Init
    Log.Init
    Verify.Init
    n = 1
    s = 0
    e = 0
    
    'Получаем коллекцию файлов
    Set files = Source.getFiles(dat.Cells(1, 3))
    
    'Обрабатываем список файлов
    For Each file In files
        curf = file
        If Len(curf) > 40 Then curf = "..." + Right(curf, 40)
        Message ("Обработка файла " + CStr(n) + " из " + CStr(files.Count) + " (" + curf) + ")"
        er = AddFile(file)
        If er > 0 Then
            Call Log.Rec(file, er)
            e = e + 1
        Else
            s = s + 1
        End If
        n = n + 1
    Next
    Message ("Готово!")
    If isRelease Then MsgBox ("Обработка завершена!" + Chr(13) + "Файлов загруженные успешно: " + CStr(s) + Chr(13) + "Файлы с ошибками: " + CStr(e))
    
End Sub

'Добавление данных из файла. Возвращает:
'0 - всё хорошо
'1 - ошибка загрузки
'2 - ошибка в данных
'3 - нет кода
'4 - запись аннулирована
Function AddFile(ByVal file As String) As Byte
    errors = False
    If isRelease Then On Error GoTo er
    Application.ScreenUpdating = False
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    If Not impBook Is Nothing Then
        Set src = impBook.Worksheets(1) 'Пока берём данные с первого листа
        src.Unprotect Template.Secret
        cod = src.Cells(1, 1)
        If cod <> "" Then
            
            'Очищаем предыдущие строки с ошибками
            i = firstDat
            Do While dat.Cells(i, 2) <> ""
                If dat.Cells(i, 1) = "" And dat.Cells(i, cCode) = cod Then
                    dat.Rows(i).Delete
                    max = max - 1
                Else
                    i = i + 1
                End If
            Loop
        
            'Индексируем существующие записи
            Set Indexes = CreateObject("Scripting.Dictionary")
            i = firstDat
            Do While dat.Cells(i, 2) <> ""
                uid = dat.Cells(i, 1)
                If uid <> "" Then Indexes.Add uid, i
                i = i + 1
            Loop
            max = i
        
            'Обрабатываем строки исходника
            Set resuids = CreateObject("Scripting.Dictionary")
            i = firstSrc
            Do While NotEmpty(i)
                uid = src.Cells(i, 1)
                'Строка уже есть (наверное)
                If uid <> "" Then
                    
                    ind = Indexes(uid)
                    If ind <> Empty Then
                        
                        'И строка действительно есть, обновляем данные
                        If copyRecord(ind, i, True) Then errors = True
                        
                        'Данные не обновлены
                        stat = dat.Cells(ind, cStatus).text
                        If stat = "0" Then
                            dat.Cells(ind, cCom) = "Данные аннулированы!"
                            dat.Cells(ind, cCom).Interior.Color = colRed
                            src.Cells(i, cCom) = "Данные аннулированы!"
                            src.Cells(i, cCom).Interior.Color = colRed
                        End If
                        If stat = "2" Then
                            dat.Cells(ind, cCom) = "Данные зафиксированы!"
                            dat.Cells(ind, cCom).Interior.Color = colGreen
                            src.Cells(i, cCom) = "Данные зафиксированы!"
                            src.Cells(i, cCom).Interior.Color = colGreen
                        End If
                        
                    Else
                        'А вот и нет, такой строки нет, стоит непонятный UID, которого у нас нет
                        uid = ""
                    End If
                End If
                'Новая строка
                If uid = "" Then If copyRecord(max, i, False) Then errors = True
                resuids.Add src.Cells(i, 1).text, 1
                i = i + 1
            Loop
            
            'Проверяем исходник на удалённые записи
            i = firstDat
            Do While dat.Cells(i, 2) <> ""
                uid = dat.Cells(i, 1)
                If uid <> "" And dat.Cells(i, cCode) = cod Then
                    If resuids(uid) = Empty Then
                        dat.Cells(i, cCom) = "Данные удалены!"
                        dat.Cells(i, cCom).Interior.Color = colRed
                        AddFile = 2
                    End If
                End If
                i = i + 1
            Loop
            
        Else
            AddFile = 3
        End If
        src.Protect Template.Secret
        impBook.Close isRelease
    End If
    Numerator.Save
    Application.ScreenUpdating = True
    DoEvents
    If errors Then AddFile = 2
    Exit Function
er:
    AddFile = 1
End Function

'Проверка на пустую строку
'Возвращает True, если строка в исходнике не пустая
Function NotEmpty(i As Long) As Boolean
    NotEmpty = False
    For j = 1 To 14
        txt = src.Cells(i, j).text
        If txt <> "" And txt <> "#Н/Д" Then NotEmpty = True: Exit For
    Next
End Function

'Копирование записи. Возвращает True, если в данных есть ошибка
'di - строка в данных
'si - строка в исходниках
'refresh - true, если обновление данных (проверять что поменялось)
Function copyRecord(ByVal di As Long, ByVal si As Long, refresh As Boolean) As Boolean
    
    stat = dat.Cells(di, cStatus).text
    If stat = "0" Or stat = "2" Then Exit Function
    
    Dim changed As Boolean
    For j = 2 To 14
        ravno = dat.Cells(di, j).text = src.Cells(si, j).text
        dat.Cells(di, j) = src.Cells(si, j)
        dat.Cells(di, j).ClearFormats
        If j = 2 Or j = 4 Or j = 6 Or j = 7 Or j = 8 Then
            src.Cells(si, j).Interior.Color = colYellow
        Else
            src.Cells(si, j).Interior.Color = colWhite
        End If
        If refresh And Not ravno Then
            dat.Cells(di, j).Interior.Color = colYellow
            src.Cells(si, j).Interior.Color = colYellow
            changed = True
        End If
    Next
    dat.Cells(di, cFile) = file
    dat.Cells(di, cCode) = cod
    Range(dat.Cells(di, cFile), dat.Cells(di, cCode)).Font.Color = RGB(192, 192, 192)
    errors = Verify.Verify(dat, src, di, si, changed)
    If errors Then
        copyRecord = True
    Else
        'Если нет ошибок, и это не обновление, присваиваем номер
        If Not refresh Then
            num = Numerator.Generate(dat.Cells(di, 2), dat.Cells(di, 4))
            dat.Cells(di, 1) = num
            src.Cells(si, 1) = num
        End If
    End If
    If Not refresh Then max = max + 1
    If dat.Cells(di, cStatus).text = "" Then dat.Cells(di, cStatus) = 1
End Function