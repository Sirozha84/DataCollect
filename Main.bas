Attribute VB_Name = "Main"
Public Const isRelease = False  'True - полноценная работа, False - режим отладки (нет вопросов, нет записи в файлы)
Public Const saveSource = True  'True - сохранение данных в формах, False - данные не записываются (отладка)
Public Const maxRow = 1048576   'Последняя строка везде (для очистки)
Public Const maxCol = 50        'Последняя колонка везде (для очистки)

'Колонки
Public Const cBuyer = 6         'Продавец
Public Const cCom = 15          'Комментарий
Public Const cStatus = 16       'Статус
Public Const cFile = 17         'Имя файла
Public Const cCode = 18         'Код формы

'Первые строки
Public Const firstDat = 8       'Первая строка в коллекции данных
Public Const firstSrc = 5       'Первая строка в исходных файлах
Public Const firstTempl = 7     'Первая строка в списке шаблонов
Public Const firstDic = 5       'Первая строка в справочнике
Public Const firstErr = 2       'Первая строка в списке ошибок
Public Const firstNum = 4       'Первая строка в словаре нумератора

'Цвета
Public colWhite As Long
Public colRed As Long
Public colGreen As Long
Public colYellow As Long

'Ссылки на таблицы
Public DAT As Variant   'Данные
Public SRC As Variant   'Исходные данные
Public DIC As Variant   'Справочники
Public ERR As Variant   'Список ошибок
Public NUM As Variant   'Словарь нумератора

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

    Message "Удаление данных"
    Init
    'On Error GoTo er
    If isRelease Then If MsgBox("Внимание! " + Chr(10) + Chr(10) + _
        "Данная процедура очистит все собранные данные список ошибок и нумераторы. " + _
        "Уже зарегистрированные данные при повторной регистрации могут присвоить другой код." + _
        Chr(10) + Chr(10) + "Продолжить?", vbYesNo) = vbNo Then Exit Sub
    Range(DAT.Cells(firstDat, 1), DAT.Cells(maxRow, maxCol)).Clear
    Range(ERR.Cells(firstErr, 1), ERR.Cells(maxRow, maxCol)).Clear
    Range(NUM.Cells(firstNum, 1), NUM.Cells(maxRow, maxCol)).Clear
    Exit Sub
    
    Message "Готово!"
    
er:
    MsgBox ("Ошибка целостности документа!")
End Sub

'Сбор данных
Sub DataCollect()
    
    If isRelease Then If MsgBox("Начинается сбор данных. Продолжить?", vbYesNo) = vbNo Then Exit Sub
    
    Message "Подготовка..."
    
    Init
    Numerator.Init
    Log.Init
    Verify.Init
    
    'Получаем коллекцию файлов
    Set files = Source.getFiles(DAT.Cells(1, 3))
    
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

'Инициализация таблиц, цветов
Sub Init()
    
    'On Error GoTo er
    
    Set DAT = ActiveSheet
    Set DIC = Sheets("Справочник")
    Set ERR = Sheets("Ошибки")
    Set NUM = Sheets("Словарь нумератора")
    
    colWhite = RGB(255, 255, 255)
    colRed = RGB(255, 192, 192)
    colGreen = RGB(192, 255, 192)
    colYellow = RGB(255, 255, 192)
    
    Exit Sub
er:
    MsgBox ("Ошибка целостности документа!")
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
        Set SRC = impBook.Worksheets(1) 'Пока берём данные с первого листа
        SRC.Unprotect Template.Secret
        cod = SRC.Cells(1, 1)
        If cod <> "" Then
            
            'Очищаем предыдущие строки с ошибками
            i = firstDat
            Do While DAT.Cells(i, 2) <> ""
                If DAT.Cells(i, 1) = "" And DAT.Cells(i, cCode) = cod Then
                    DAT.Rows(i).Delete
                    max = max - 1
                Else
                    i = i + 1
                End If
            Loop
        
            'Индексируем существующие записи
            Set Indexes = CreateObject("Scripting.Dictionary")
            i = firstDat
            Do While DAT.Cells(i, 2) <> ""
                UID = DAT.Cells(i, 1)
                If UID <> "" Then Indexes.Add UID, i
                i = i + 1
            Loop
            max = i
        
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
                        If copyRecord(ind, i, True) Then errors = True
                        
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
                If UID = "" Then If copyRecord(max, i, False) Then errors = True
                rUID = SRC.Cells(i, 1).text
                If rUID <> "" Then resUIDs.Add rUID, 1
                i = i + 1
            Loop
            
            'Проверяем исходник на удалённые записи
            i = firstDat
            Do While DAT.Cells(i, 2) <> ""
                UID = DAT.Cells(i, 1)
                If UID <> "" And DAT.Cells(i, cCode) = cod Then
                    If resUIDs(UID) = Empty Then
                        DAT.Cells(i, cCom) = "Данные удалены!"
                        DAT.Cells(i, cCom).Interior.Color = colRed
                        AddFile = 2
                    End If
                End If
                i = i + 1
            Loop
            
        Else
            AddFile = 3
        End If
        SRC.Protect Template.Secret
        impBook.Close saveSource
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
        txt = SRC.Cells(i, j).text
        If txt <> "" And txt <> "#Н/Д" Then NotEmpty = True: Exit For
    Next
End Function

'Копирование записи. Возвращает True, если в данных есть ошибка
'di - строка в данных
'si - строка в исходниках
'refresh - true, если обновление данных (проверять что поменялось)
Function copyRecord(ByVal di As Long, ByVal si As Long, refresh As Boolean) As Boolean
    
    stat = DAT.Cells(di, cStatus).text
    If stat = "0" Or stat = "2" Then Exit Function
    
    'Копирование записей с проверкой на изменение
    Dim changed As Boolean
    For j = 2 To 14
        ravno = DAT.Cells(di, j).text = SRC.Cells(si, j).text
        DAT.Cells(di, j) = SRC.Cells(si, j)
        DAT.Cells(di, j).ClearFormats
        If j = 2 Or j = 4 Or j = 6 Or j = 7 Or j = 8 Then
            SRC.Cells(si, j).Interior.Color = colYellow
        Else
            SRC.Cells(si, j).Interior.Color = colWhite
        End If
        If refresh And Not ravno Then
            DAT.Cells(di, j).Interior.Color = colYellow
            SRC.Cells(si, j).Interior.Color = colYellow
            changed = True
        End If
    Next
    DAT.Cells(di, cFile) = file
    DAT.Cells(di, cCode) = cod
    Range(DAT.Cells(di, cFile), DAT.Cells(di, cCode)).Font.Color = RGB(192, 192, 192)
    errors = Verify.Verify(DAT, SRC, di, si, changed)
    
    'Если нужно, присваиваем записи новый номер
    If Not errors Then
        Dim needNum As Boolean
        If refresh Then
            needNum = Not Numerator.CheckPrefix(DAT.Cells(di, 1).text, _
                DAT.Cells(di, 2), DAT.Cells(di, cBuyer).text)
        Else
            needNum = True
        End If
        If needNum Then
            n = Numerator.Generate(DAT.Cells(di, 2), DAT.Cells(di, cBuyer).text)
            DAT.Cells(di, 1) = n
            SRC.Cells(si, 1) = n
        End If
    Else
        copyRecord = True
    End If
    
    If Not refresh Then max = max + 1
    If DAT.Cells(di, cStatus).text = "" Then DAT.Cells(di, cStatus) = 1
    
End Function