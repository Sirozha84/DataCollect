Attribute VB_Name = "Main"
Public Const isRelease = True   'True - полноценная работа, False - режим отладки (нет вопросов, нет записи в файлы)
Public Const saveSource = True  'True - сохранение данных в формах, False - данные не записываются (отладка)

Public Const Secret = "123"     'Пароль для защиты

Public Const maxRow = 1048576   'Последняя строка везде (для очистки)
Public Const tmpVersion = "20210108"    'Версия реестра

'Колонки "Данные"
Public Const cUIN = 1           'УИН
Public Const cDates = 2         'Дата
Public Const cBuyINN = 3        'ИНН покупателя
Public Const cBuyer = 4         'Наименование покупателя
Public Const cSellINN = 5       'ИНН продавца
Public Const cSeller = 6        'Наименование продавец
Public Const cPrice = 7         'Стоимость с НДС
Public Const cCom = 15          'Комментарий
Public Const cStatus = 16       'Статус
Public Const cDateCol = 17      'Дата сбора
Public Const cFile = 18         'Имя файла
Public Const cCode = 19         'Код формы
Public Const cAccept = 20       'Принято/не принято

'Колонки "Справочник"
Public Const cSellerName = 1    'Наименование продавца
Public Const cINN = 2           'ИНН
Public Const cSDate = 3         'Дата регистрации
Public Const cGroup = 4         'Группы
Public Const cPLiter = 6        'Префикс - литер
Public Const cPCode = 7         'Префикс - код
Public Const cPStat = 8         'Статус
Public Const cLimits = 9        'Первая колонка с остатками
Public Const cPFact = 21        'Первая колонка с фактическими объёмами
Public Const cPBalance = 33     'Первая колонка с остатками
Public Const cPRev = 45
Public Const quartCount = 12    'Количество кварталов в расчётах лимитов
Public Const lastYear = 2020    'Первый расчётный год (потом это будет переменной, но пока статика)
Public Const lastQuartal = 4    'Первыё расчётный квартал (аналогично)

'Колонки "Шаблоны"
Public Const cTClient = 1       'Клиент
Public Const cTBroker = 2       'Посредник
Public Const cTForm = 3         'Форма
Public Const cTCode = 4         'Код
Public Const cTFile = 5         'Файл
Public Const cTResult = 6       'Результат
Public Const cTStat = 7         'Статус

'Первые строки
Public Const firstDat = 8       'Первая строка в коллекции данных
Public Const firstSrc = 5       'Первая строка в исходных файлах
Public Const firstTempl = 6     'Первая строка в списке реестра
Public Const firstDic = 4       'Первая строка в справочнике
Public Const firstErr = 2       'Первая строка в списке ошибок
Public Const firstNum = 4       'Первая строка в словаре нумератора
Public Const firstValues = 6    'Первая строка в отчёте "Объёмы"

'Цвета
Public colWhite As Long
Public colRed As Long
Public colGreen As Long
Public colYellow As Long
Public colGray As Long
Public colBlue As Long

'Ссылки на таблицы
Public DAT As Variant           'Данные
Public SRC As Variant           'Исходные данные
Public DIC As Variant           'Справочники
Public ERR As Variant           'Список ошибок
Public NUM As Variant           'Словарь нумератора
Public VAL As Variant           'Значения объёмов
Public VLS As Variant           'Сводная таблица
Public TMP As Variant           'Шаблоны
Public SBK As Variant           'Книги продаж

'Общие переменные
Public selIndexes As Variant    'Словарь индексов продавцов (номера строк в справочнике по ИНН)
Public BookCount As Long        'Счётчик сгенерированных книг

'Инициализация таблиц, цветов
Sub Init()
    colWhite = RGB(255, 255, 255)
    colRed = RGB(255, 192, 192)
    colGreen = RGB(192, 255, 192)
    colYellow = RGB(255, 255, 192)
    colGray = RGB(217, 217, 217)
    colBlue = RGB(192, 217, 255)
    
    If isRelease Then On Error GoTo er
    Set DAT = Sheets("Данные")
    Set DIC = Sheets("Справочник")
    Set ERR = Sheets("Ошибки")
    Set NUM = Sheets("Словарь нумератора")
    Set VAL = Sheets("Объёмы")
    Set VLS = Sheets("Сводная таблица")
    Set TMP = Sheets("Шаблоны")
    Set SBK = Sheets("Книги продаж")
    
    Exit Sub
er:
    MsgBox ("Ошибка целостности документа! Необходимые вкладки были удалены или переименовены.")
    End
End Sub

'******************** Вкладка "Данные" ********************

'Выбор директории с данными
Sub ButtonDirSelectImport()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Cells(1, 3) = diag.SelectedItems(1)
End Sub

'Выбор директории для экспорта
Sub ButtonDirSelectExport()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Cells(2, 3) = diag.SelectedItems(1)
End Sub

'Кнопка "Сбор данных"
Sub ButtonDataCollect()
    Init
    If isRelease Then If MsgBox("Начинается сбор данных. Продолжить?", vbYesNo) = vbNo Then Exit Sub
    Message "Подготовка..."
    SetProtect DAT
    Collect.Run
    DAT.Activate
End Sub

'Кнопка "Экспорт в 1С"
Sub ButtonExport()
    Init
    FormExport.Show
End Sub

'Кнопка "Удаление данных"
Sub ButtonClear()
    Init
    If isRelease Then
        e = Chr(10)
        If InputBox("Внимание! " + e + e + _
            "Данная процедура очистит все собранные данные. " + _
            "Уже зарегистрированные данные при повторной регистрации могут присвоить другой код. " + _
            "Справочник и словари нумератора удаляться не будут." + e + e + _
            "Для продолжения введите пароль.", "Удаление данных") <> Secret Then Exit Sub
    End If
    SetProtect DAT
    Range(Cells(firstDat, 1), Cells(maxRow, cAccept)).Clear
    Range(Cells(firstDat, cStatus), Cells(maxRow, cDateCol)).Interior.Color = colYellow
    Range(Cells(firstDat, cFile), Cells(maxRow, cAccept)).Interior.Color = colGray
    Range(Cells(firstDat, cFile), Cells(maxRow, cAccept)).Font.Color = RGB(166, 166, 166)
    Range(DIC.Cells(firstDic, cPFact), DIC.Cells(maxRow, cPFact + quartCount - 1)).Clear
    Message "Готово!"
End Sub

'******************** Вкладка "Объёмы" ********************

'Кнопка ревизии остатков
Sub ButtonRevisionVolumes()
    Init
    Revision.Run
End Sub

'Кнопка "Сформировать отчёт" по объёмам
Sub ButtonReportVolumes()
    Init
    Values.CreateReport
    VAL.Activate
End Sub

'******************** Вкладка "Шаблоны" ********************

'Кнопка "Генерировать шаблоны"
Sub ButtonCreateTemplates()
    Init
    Template.Generate
End Sub

'******************** Вкладка "Книги продаж" ********************

'Кнопка "Сформировать"
Public Sub ButtonSellBook()
    Init
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Patch = diag.SelectedItems(1)
    Set files = getFiles(Patch, False)
    Range(SBK.Cells(7, 1), SBK.Cells(maxRow, 2)).Clear
    i = 7
    For Each file In files
        SBK.Cells(i, 1) = file
        er = ExportBook(file)
        If er = 0 Then SBK.Cells(i, 2) = "Ошибка при работе с файлом"
        If er = 1 Then
            If BookCount > 0 Then
                SBK.Cells(i, 2) = "Созданы книги продаж (" + CStr(BookCount) + ")"
            Else
                SBK.Cells(i, 2) = "Реестр пустой"
            End If
        End If
        If er = 2 Then SBK.Cells(i, 2) = "Реестр имеет некорректные записи"
        i = i + 1
    Next
    VAL.Activate
    Message "Готово!"
    MsgBox "Формирование книг продаж завершено!"
End Sub