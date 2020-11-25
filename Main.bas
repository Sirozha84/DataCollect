Attribute VB_Name = "Main"
Const isRelease = True 'True - полноценная работа, False - режим отладки (нет вопросов, нет записи в файлы)

Const FirstD = 8        'Первая строка в коллекции данных
Const FirstS = 5        'Первая строка в исходных файлах
Const cFile = 16        'Колонка с именем файла
Const cCode = 17        'Колонка с кодом файла

Const errName = "Ошибки"

Dim dat As Variant      'Таблица с данными
Dim src As Variant      'Таблица с исходниками
Dim err As Variant      'Таблица с ошибками
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
    If isRelease Then If MsgBox("Данная процедура очистит собранные данные список ошибок и нумераторы. Продолжить?", vbYesNo) = vbNo Then Exit Sub
    Range(Cells(FirstD, 1), Cells(1048576, 50)).Clear
    Sheets(errName).Cells.Clear
er:
    Numerator.Clear
End Sub

'Сбор данных
Sub DataCollect()
    
    Set dat = ActiveSheet
    noEmpty = (dat.Cells(FirstD, 2) <> "")
    If isRelease And noEmpty Then If MsgBox("Начинается сбор данных. Продолжить?", vbYesNo) = vbNo Then Exit Sub
    Message "Подготовка"
    
    'Получаем коллекцию файлов
    Set files = Source.getFiles(dat.Cells(1, 3))
        
    'Создаём вкладку (если её нет) для списка ошибок
    Call NewTab(errName, True)
    Set err = Sheets(errName)
    err.Columns(1).ColumnWidth = 100
    err.Columns(2).ColumnWidth = 20
    err.Cells(1, 1) = "Файл"
    err.Cells(1, 2) = "Результат"
            
    'Очищаем предыдущие строки с ошибками
    Application.ScreenUpdating = False
    i = FirstD
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
    i = FirstD
    Do While dat.Cells(i, 2) <> ""
        uid = dat.Cells(i, 1)
        If uid <> "" Then Indexes.Add uid, i
        i = i + 1
    Loop
    max = i
    
    'Инициализируем словари и сбрасываем счётчики
    Numerator.Init
    Verify.Init
    n = 1
    s = 0
    e = 0
    
    'Обрабатываем список файлов
    For Each file In files
        curf = file
        If Len(curf) > 40 Then curf = "..." + Right(curf, 40)
        Message ("Обработка файла " + CStr(n) + " из " + CStr(files.count) + " (" + curf) + ")"
        er = AddFile(file)
        If er > 0 Then
            e = e + 1
            err.Cells(e + 1, 1) = file
            If er = 1 Then err.Cells(e + 1, 2) = "Ошибка загрузки файла"
            If er = 2 Then err.Cells(e + 1, 2) = "Ошибка в данных"
            If er = 3 Then err.Cells(e + 1, 2) = "Отсутствует код"
        Else
            s = s + 1
        End If
        n = n + 1
    Next
    Message ("Готово!")
    If isRelease Then MsgBox ("Обработка завершена!" + Chr(13) + "Файлов загруженные успешно: " + CStr(s) + Chr(13) + "Файлы с ошибками: " + CStr(e))
    
End Sub

'Добавление данных из файла (возвращает 0 - всё хорошо, 1 - ошибка загрузки, 2 - ошибка в данных, 3 - нет кода)
Function AddFile(ByVal file As String) As Byte
    errors = False
    On Error GoTo er
    Application.ScreenUpdating = False
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    If Not impBook Is Nothing Then
        Set src = impBook.Worksheets(1) 'Пока берём данные с первого листа
        src.Unprotect Template.Secret
        cod = src.Cells(1, 1)
        If cod <> "" Then
        
            'Обрабатываем строки исходника
            i = FirstS
            Do While NotEmpty(i)
                uid = src.Cells(i, 1)
                'Строка уже есть (наверное)
                If uid <> "" Then
                    ind = Indexes(uid)
                    If ind <> Empty Then
                        'И строка действительно есть, обновляем данные
                        If copyRecord(file, ind, i, True) Then errors = True
                    Else
                        'А вот и нет, такой строки нет, стоит непонятный UID, которого у нас нет
                        uid = ""
                    End If
                End If
                'Новая строка
                If uid = "" Then If copyRecord(file, max, i, False) Then errors = True
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

'Возвращает True если строка si в исходнике не пустая
Function NotEmpty(si As Long) As Boolean
    NotEmpty = False
    For j = 1 To 14
        txt = src.Cells(si, j).text
        If txt <> "" And txt <> "#Н/Д" Then NotEmpty = True
    Next
End Function

'Копирование записи. refresh - обновление данных (проверять что поменялось)
'Возвращает True - если в данных есть ошибка
Function copyRecord(file As String, ByVal di As Long, ByVal si As Long, refresh As Boolean) As Boolean
    Dim changed As Boolean
    wht = RGB(255, 255, 255)
    yel = RGB(256, 256, 192)
    For j = 2 To 14
        ravno = dat.Cells(di, j).text = src.Cells(si, j).text
        dat.Cells(di, j) = src.Cells(si, j)
        dat.Cells(di, j).ClearFormats
        If j = 2 Or j = 4 Or j = 6 Or j = 7 Or j = 8 Then
            src.Cells(si, j).Interior.Color = yel
        Else
            src.Cells(si, j).Interior.Color = wht
        End If
        If refresh And Not ravno Then
            dat.Cells(di, j).Interior.Color = yel
            src.Cells(si, j).Interior.Color = yel
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
End Function