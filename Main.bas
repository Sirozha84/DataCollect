Attribute VB_Name = "Main"
Const isRelease = False 'True - полноценная работа, False - режим отладки (нет вопросов, нет записи в файлы)
Const FirstC = 6        'Первая строка в коллекции данных
Const FirstS = 5        'Первая строка в исходных файлах
Const cFile = 17        'Колонка с именем файла
Const cCode = 18        'Колонка с кодом файла
Const erTabName = "Ошибки"

Dim erTab As Variant

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
    Range(Cells(StartString, 1), Cells(1048576, 50)).Clear
    Sheets(erTabName).Cells.Clear
er:
    Numerator.Clear
End Sub

'Сбор данных
Sub DataCollect()
    If isRelease Then If MsgBox("Начинается сбор данных. Продолжить?", vbYesNo) = vbNo Then Exit Sub
    
    'Получаем коллекцию файлов
    Set Files = Source.GetList("C:\Users\SG\OneDrive\Работа\Цифровая сибирь\Сбор данных\Данные")
        
    'Создаём вкладку (если её нет) для списка ошибок
    Call NewTab(erTabName, True)
    Set erTab = Sheets("Ошибки")
    erTab.Columns(1).ColumnWidth = 100
    erTab.Columns(2).ColumnWidth = 20
    erTab.Cells(1, 1) = "Файл"
    erTab.Cells(1, 2) = "Результат"
    
    'Инициализируем словарь нумератора
    Numerator.Init
    
    Dim str As Long
    str = FirstC
    n = 1
    s = 0
    e = 0
    Max = Files.Count
    For Each file In Files
        Message ("Обработка файла " + CStr(n) + " из " + CStr(Files.Count) + " (" + Source.FSO.getfilename(file)) + ")"
        er = AddFile(file, str)
        If er > 0 Then
            e = e + 1
            erTab.Cells(e + 1, 1) = file
            If er = 1 Then erTab.Cells(e + 1, 2) = "Ошибка загрузки файла"
            If er = 2 Then erTab.Cells(e + 1, 2) = "Ошибка в данных"
            If er = 3 Then erTab.Cells(e + 1, 2) = "Отсутствует код"
        Else
            s = s + 1
        End If
        n = n + 1
    Next
    Message ("Готово!")
    If isRelease Then MsgBox ("Обработка завершена!" + Chr(13) + "Файлов загруженные успешно: " + CStr(s) + Chr(13) + "Файлы с ошибками: " + CStr(e))
    
End Sub

'Добавление данных из файла (возвращает 0 - всё хорошо, 1 - ошибка загрузки, 2 - ошибка в данных, 3 - нет кода)
Function AddFile(ByVal file As String, ByRef str As Long) As Byte
    
    On Error GoTo er
    
    Application.ScreenUpdating = False
    Set cur = ActiveSheet
    Set imBook = Nothing
    Set imBook = Workbooks.Open(file, False, False)
    If Not imBook Is Nothing Then
        Set imSh = imBook.Worksheets(1) 'Пока берём данные с первого листа
        cod = imSh.Cells(1, 1)
        If cod <> "" Then
        
            'Очищаем предыдущие строки с ошибками
            
            
            'Обрабатываем строки исходника
            i = FirstString
            Do While imSh.Cells(i, 2) <> ""
                
                'Копируем строчку
                For j = 2 To 14
                    cur.Cells(str, j) = imSh.Cells(i, j)
                Next
                cur.Cells(str, cFile) = file
                cur.Cells(str, cCode) = cod
                Range(cur.Cells(str, cFile), cur.Cells(str, cSheet)).Font.Color = RGB(192, 192, 192)
                
                errors = Verify.Verify(cur, imSh, str, i)
                If errors Then
                    AddFile = 2
                Else
                    'Если нет ошибок, присваиваем номер
                    cur.Cells(str, 1) = Numerator.Generate(cur.Cells(str, 2), cur.Cells(str, 4))
                End If
                str = str + 1
                i = i + 1
            Loop
        Else
            AddFile = 3
        End If
        imBook.Close isRelease
    End If
    Numerator.Save
    Application.ScreenUpdating = True
    DoEvents
    Exit Function
er:
    AddFile = 1
End Function