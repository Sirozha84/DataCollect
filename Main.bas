Attribute VB_Name = "Main"
'Last change: 23.04.2021 14:09

'Константы
Public Const maxRow = 1048576   'Последняя строка везде (для очистки)
Public Const tmpVersion = "20210108"    'Необходимая версия реестра

'Настройки
Public Const Secret = "123"     'Пароль для защиты
Public Const quartCount = 12    'Количество кварталов в расчётах лимитов
Public Const lastYear = 2020    'Первый расчётный год
Public Const lastQuartal = 4    'Первый расчётный квартал
Public Const limitOND = 9000000 'Лимит в основной период НД (9м)
Public Const minLim = 5000000   'Минимальная сумма продаж, если меньше, период пропускается (5м)
Public Const minSale = 20000    'Минимальная сумма продаж, достаточная для распределения закупок (20т)
Public Const maxDif = 15000     'Минимальная разница между отгрузкой и поступлением (15т)

'Колонки "Отгрузки"
Public Const cUIN = 1           'УИН
Public Const cDates = 2         'Дата
Public Const cBuyINN = 3        'Покупатель ИНН
Public Const cBuyer = 4         'Покупатель Наименование
Public Const cSellINN = 5       'Продавец ИНН
Public Const cSeller = 6        'Продавец Наименование
Public Const cPrice = 7         'Стоимость с НДС
Public Const cCom = 15          'Комментарий
Public Const cStatus = 16       'Статус
Public Const cDateCol = 17      'Дата сбора
Public Const cFile = 18         'Имя файла
Public Const cCode = 19         'Код формы
Public Const cAccept = 20       'Принято/не принято
Public Const cPND = 21          'Период НД

'Колонки "Поступления"
Public Const clMark = 1         'Маркер
Public Const clKVO = 2          'Код вида операции (КВО)
Public Const clNum = 3          'Номер
Public Const clDate = 4         'Дата
Public Const clProvINN = 5      'Поставщик ИНН
Public Const clProvName = 6     'Поставшик Наименование
Public Const clSaleINN = 7      'Продавец ИНН
Public Const clSaleName = 8     'Продавец Наименование
Public Const clPrice = 9        'Стоимость с НДС
Public Const clNDS = 13         'Сумма НДС
Public Const clCom = 14         'Комментарий
Public Const clStatus = 15      'Статус
Public Const clRasp = 16        'Распределено
Public Const clPND = 17         'Период НД
Public Const clOst = 18         'Остаток НДС
Public Const clDateCol = 19     'Дата сбора
Public Const clUIN = 20         'УИН
Public Const clFile = 21        'Имя файла
Public Const clAccept = 22      'Принято/не принято

'Колонки "Справочник"
Public Const cSellerName = 1    'Наименование продавца
Public Const cINN = 2           'ИНН
Public Const cSDate = 3         'Дата регистрации
Public Const cGroup = 4         'Группы
Public Const cLimND = 5         'Лимит НД за квартал
Public Const cPLiter = 6        'Префикс - литер
Public Const cPCode = 7         'Префикс - код
Public Const cOPND = 8          'Основной период НД
Public Const cPStat = 9         'Статус
Public Const cLimits = 10       'Первая колонка с остатками
Public Const cPFact = 22        'Первая колонка с фактическими объёмами
Public Const cPBalance = 34     'Первая колонка с остатками (*2)
Public Const cCorrect = 58      'Первая колонка с корректировками лимитов
Public Const cPRev = 70         'Первая колонка с фактическими отгрузками (для ревизии остатков)
Public Const cSaleProtect = 82  'Первая колонка с запретами отгрузок

'Колонки "Шаблоны"
Public Const cTClient = 1       'Клиент
Public Const cTBroker = 2       'Посредник
Public Const cTForm = 3         'Форма
Public Const cTCode = 4         'Код
Public Const cTFile = 5         'Файл
Public Const cTResult = 6       'Результат
Public Const cTStat = 7         'Статус

'Первые строки
Public Const firstDat = 6       'Отгрузки
Public Const firstDtL = 6       'Поступления
Public Const firstSrc = 5       'Реестры
Public Const firstTempl = 6     'Список шаблонов
Public Const firstDic = 4       'Справочник
Public Const firstErr = 2       'Ошибки
Public Const firstNum = 4       'Словарь нумератора
Public Const firstValues = 6    'Отчёт "Объёмы"

'Строки с параметрами
Public Const pImportSale = 4    'Импорт отгрузок
Public Const pImportLoad = 5    'Импорт поступлений
Public Const pExport = 6        'Экспорт

'Цвета
Public colWhite As Long         'Для невидимости строк
Public colRed As Long           'Ошибки
Public colGreen As Long         'Принятые
Public colYellow As Long        'Разрешёные для редактирования
Public colGray As Long          'Служебные поля
Public colBlue As Long          'Замечены изменения

'Ссылки на таблицы
Public DAT As Variant           'Данные о продажах
Public DTL As Variant           'Данные о поступлениях
Public SRC As Variant           'Исходные данные
Public DIC As Variant           'Справочники
Public ERR As Variant           'Список ошибок
Public NUM As Variant           'Словарь нумератора
Public VAL As Variant           'Значения объёмов
Public VLS As Variant           'Сводная таблица
Public TMP As Variant           'Шаблоны
Public SBK As Variant           'Книги продаж
Public PRP As Variant           'Настройки

'Настройки
Public DirImportSale As String  'Каталог импорта отгрузок
Public DirImportLoad As String  'Каталог импорта поступлений
Public DirExport As String      'Каталог экспорта

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
    
    On Error GoTo er
    Set DAT = Sheets("Отгрузки")
    Set DTL = Sheets("Поступления")
    Set DIC = Sheets("Справочник")
    Set VAL = Sheets("Объёмы")
    Set VLS = Sheets("Сводная таблица")
    Set TMP = Sheets("Шаблоны")
    Set SBK = Sheets("Книги продаж")
    Set ERR = Sheets("Ошибки")
    Set NUM = Sheets("Нумератор")
    Set PRP = Sheets("Настройки")
    
    DirImportSale = PRP.Cells(pImportSale, 2).text
    DirImportLoad = PRP.Cells(pImportLoad, 2).text
    DirExport = PRP.Cells(pExport, 2).text
    
    Exit Sub
er:
    MsgBox ("Ошибка целостности документа! Необходимые вкладки были удалены или переименовены.")
    End
End Sub

Sub ButtonProperties()
    Init
    FormProperties.Show
End Sub

'******************** Вкладка "Отгрузки" ********************

'Кнопка "Сбор данных"
Sub ButtonDataCollect()
    Init
    If MsgBox("Начинается сбор данных по отгрузкам. Продолжить?", vbYesNo) = vbNo Then Exit Sub
    SetProtect DAT
    CollectSale.Run
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
    e = Chr(10)
    If InputBox("Внимание! " + e + e + _
        "Данная процедура очистит все собранные данные. " + _
        "Уже зарегистрированные данные при повторной регистрации могут присвоить другой код. " + _
        "Справочник и словари нумератора удаляться не будут." + e + e + _
        "Для продолжения введите пароль.", "Удаление данных") <> Secret Then Exit Sub
    SetProtect DAT
    Range(DAT.Cells(firstDat, 1), DAT.Cells(maxRow, cAccept)).Clear
    Range(DAT.Cells(firstDat, cStatus), DAT.Cells(maxRow, cDateCol)).Interior.Color = colYellow
    Range(DAT.Cells(firstDat, cFile), DAT.Cells(maxRow, cAccept)).Interior.Color = colGray
    Range(DAT.Cells(firstDat, cFile), DAT.Cells(maxRow, cAccept)).Font.Color = RGB(166, 166, 166)
    Range(DTL.Cells(firstDtL, 1), DTL.Cells(maxRow, clAccept)).Clear
    Range(DTL.Cells(firstDtL, clFile), DTL.Cells(maxRow, clAccept)).Interior.Color = colGray
    Range(DTL.Cells(firstDtL, clFile), DTL.Cells(maxRow, clAccept)).Font.Color = RGB(166, 166, 166)
    Range(DIC.Cells(firstDic, cPFact), DIC.Cells(maxRow, cPFact + quartCount * 6 - 1)).Clear
    Range(DIC.Cells(firstDic, cSaleProtect), DIC.Cells(maxRow, cSaleProtect + quartCount - 1)). _
            Interior.Color = colGray
    
    Message "Готово! Файл не был сохранён. " + _
            "Если передумали - закройте файл не сохраняясь и откройте снова."
End Sub

'******************** Вкладка "Поступления" ********************

'Кнопка "Сбор поступлений"
Sub ButtonCollectLoad()
    Init
    If MsgBox("Начинается сбор данных по поступлениям. " + _
            "Продолжить?", vbYesNo) = vbNo Then Exit Sub
    CollectLoad.Run
End Sub

'Кнопка "Экспорт поступлений в 1С"
Sub ButtonExportLoad()
    Init
    If MsgBox("Начинается экспорт данных о поступлениях. Продолжить?", vbYesNo) = vbNo Then Exit Sub
    ExportLoad.Run
End Sub

'******************** Вкладка "Объёмы" ********************

'Кнопка ревизии остатков
Sub ButtonRevisionVolumes()
    Init
    DIC.Activate
    DIC.Cells(firstDic, cPRev).Activate
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
    TMP.Activate
    Template.Generate
End Sub

'******************** Вкладка "Книги продаж" ********************

'Кнопка "Сформировать"
Public Sub ButtonSellBook()
    Init
    SellBook.Run
End Sub

'******************** End of File ********************