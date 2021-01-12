Attribute VB_Name = "Main"
Public Const isRelease = True   'True - полноценная работа, False - режим отладки (нет вопросов, нет записи в файлы)
Public Const saveSource = True  'True - сохранение данных в формах, False - данные не записываются (отладка)

Public Const Secret = "123"     'Пароль для защиты

Public Const maxRow = 1048576   'Последняя строка везде (для очистки)
Public Const tmpVersion = "20210108"    'Версия реестра

'Колонки "Данные"
Public Const cDates = 2         'Дата
Public Const cBuyINN = 3        'ИНН покупателя
Public Const cBuyer = 4
Public Const cSellINN = 5       'ИНН продавца
Public Const cSeller = 6        'Продавец
Public Const cPrice = 7         'Стоимость с НДС
Public Const cCom = 15          'Комментарий
Public Const cStatus = 16       'Статус
Public Const cFile = 17         'Имя файла
Public Const cCode = 18         'Код формы
Public Const cAccept = 19       'Принято/не принято

'Колонки "Справочник"
Public Const cSellerName = 1    'Наименование продавца
Public Const cINN = 2           'ИНН
Public Const cSDate = 3         'Дата регистрации
Public Const cGroup = 4         'Группы
Public Const cLimits = 5        'Лимиты
Public Const cPLiter = 6        'Префикс - литер
Public Const cPCode = 7         'Префикс - код

'Первые строки
Public Const firstDat = 8       'Первая строка в коллекции данных
Public Const firstSrc = 5       'Первая строка в исходных файлах
Public Const firstTempl = 7     'Первая строка в списке реестра
Public Const firstDic = 4       'Первая строка в справочнике
Public Const firstErr = 2       'Первая строка в списке ошибок
Public Const firstNum = 4       'Первая строка в словаре нумератора

'Цвета
Public colWhite As Long
Public colRed As Long
Public colGreen As Long
Public colYellow As Long
Public colGray As Long
Public colBlue As Long

'Ссылки на таблицы
Public DAT As Variant   'Данные
Public SRC As Variant   'Исходные данные
Public DIC As Variant   'Справочники
Public ERR As Variant   'Список ошибок
Public NUM As Variant   'Словарь нумератора
Public VAL As Variant   'Значения объёмов

'Выбор директории с данными
Sub ButtonDirSelect()
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

'Кнопка "Книги продаж"
Sub ButtonSellBook()
    file = Application.GetOpenFilename("Файлы Excel(*.xls*),*.xls*", 1, _
        "Выберите файл реестра", , False)
    If VarType(file) = vbBoolean Then Exit Sub
    ExportBook ByVal CStr(file)
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
    Range(Cells(firstDat, cStatus), Cells(maxRow, cStatus)).Interior.Color = colYellow
    Range(Cells(firstDat, cFile), Cells(maxRow, cAccept)).Interior.Color = colGray
    Range(Cells(firstDat, cFile), Cells(maxRow, cAccept)).Font.Color = RGB(166, 166, 166)
    Message "Готово!"
End Sub

'Кнопка "Сбор данных"
Sub ButtonDataCollect()
    Init
    If isRelease Then If MsgBox("Начинается сбор данных. Продолжить?", vbYesNo) = vbNo Then Exit Sub
    Message "Подготовка..."
    SetProtect DAT
    Collect.Run
End Sub

'Кнопка "Генерировать шаблоны"
Sub ButtonCreateTemplates()
    Init
    Template.Generate
End Sub

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
    
    Exit Sub
er:
    MsgBox ("Ошибка целостности документа!")
    End
End Sub

'Установка защиты
Sub SetProtect(table As Variant)
    table.Protect Secret, AllowFormattingColumns:=True, UserInterfaceOnly:=True, AllowFiltering:=True
End Sub