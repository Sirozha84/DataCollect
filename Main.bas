Attribute VB_Name = "Main"
'Const isRelease = True  'True - полноценная работа, False - режим отладки (нет вопросов, нет записи в файлы)
Const isRelease = False 'True - полноценная работа, False - режим отладки (нет вопросов, нет записи в файлы)
Const FirstC = 6        'Первая строка в коллекции данных
Const FirstS = 5        'Первая строка в исходных файлах
Const cFile = 17        'Колонка с именем файла
Const cCode = 18        'Колонка с кодом файла
Const errName = "Ошибки"
Dim dat As Variant      'Таблица с данными
Dim err As Variant      'Таблица с ошибками
Dim Indexes As Object   'Словарь индексов
Dim max As Long         'Последняя строка в данных

'Выбор директории с данными
Sub DirSelect()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Cells(1, 3) = diag.SelectedItems(1)
End Sub

'Удаление всех данных (оставляя шапку)
Sub Clear()
    On Error GoTo er
    'If MsgBox("Данная процедура очистит собранные данные список ошибок и нумераторы. Продолжить?", vbYesNo) = vbNo Then Exit Sub
    Range(Cells(FirstC, 1), Cells(1048576, 50)).Clear
    Sheets(errName).Cells.Clear
er:
    Numerator.Clear
End Sub

'Сбор данных
Sub DataCollect()
    
    Set dat = ActiveSheet
    noEmpty = (dat.Cells(FirstC, 2) <> "")
    If isRelease And noEmpty Then If MsgBox("Начинается сбор данных. Продолжить?", vbYesNo) = vbNo Then Exit Sub
    Message "Подготовка"
    
    'Получаем коллекцию файлов
    Set Files = Source.GetList("C:\Users\SG\OneDrive\Работа\Цифровая сибирь\Сбор данных\Данные")
        
    'Создаём вкладку (если её нет) для списка ошибок
    Call NewTab(errName, True)
    Set err = Sheets(errName)
    err.Columns(1).ColumnWidth = 100
    err.Columns(2).ColumnWidth = 20
    err.Cells(1, 1) = "Файл"
    err.Cells(1, 2) = "Результат"
    
    'Индексируем собранные данные
    
    
    'Инициализируем словарь нумератора
    Numerator.Init
    
    n = 1
    s = 0
    e = 0
    'max = FindMax(dat, FirstC, 2) - noEmpty 'Если не пусто +1, потому что с этой строки будет продолжаться запись новых строк
    
    For Each file In Files
        Message ("Обработка файла " + CStr(n) + " из " + CStr(Files.Count) + " (" + Source.FSO.getfilename(file)) + ")"
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
    
    On Error GoTo er
    
    Application.ScreenUpdating = False
    
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    If Not impBook Is Nothing Then
        Set src = impBook.Worksheets(1) 'Пока берём данные с первого листа
        cod = src.Cells(1, 1)
        If cod <> "" Then
        
            'Очищаем предыдущие строки с ошибками
            Dim i As Long
            i = FirstC
            Do While dat.Cells(i, 2) <> ""
                If dat.Cells(i, 1) = "" And dat.Cells(i, cCode) = cod Then
                    dat.Rows(i).Delete
                    max = max - 1
                Else
                    i = i + 1
                End If
            Loop
            
            'Обрабатываем строки исходника
            i = FirstS
            Do While src.Cells(i, 2) <> ""
                
                If src.Cells(i, 1) = "" Then
                    'Строки нет
                    For j = 2 To 14
                        dat.Cells(max, j) = src.Cells(i, j)
                    Next
                    dat.Cells(max, cFile) = file
                    dat.Cells(max, cCode) = cod
                    Range(dat.Cells(max, cFile), dat.Cells(max, cCode)).Font.Color = RGB(192, 192, 192)
                    
                    errors = Verify.Verify(dat, src, max, i)
                    If errors Then
                        AddFile = 2
                    Else
                        'Если нет ошибок, присваиваем номер
                        num = Numerator.Generate(dat.Cells(max, 2), dat.Cells(max, 4))
                        dat.Cells(max, 1) = num
                        src.Cells(i, 1) = num
                    End If
                    max = max + 1
                Else
                    'Строка есть
                End If
                
                i = i + 1
            Loop
        Else
            AddFile = 3
        End If
        impBook.Close isRelease
    End If
    Numerator.Save
    Application.ScreenUpdating = True
    DoEvents
    Exit Function
er:
    AddFile = 1
End Function