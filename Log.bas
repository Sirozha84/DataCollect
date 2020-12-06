Attribute VB_Name = "Log"
Dim err As Variant  'Таблица с ошибками
Dim recN As Long    'Текущий номер строки

Sub Init()
    'Создаём вкладку (если её нет) для списка ошибок
    Call NewTab(tabErr, True)
    Set err = Sheets(tabErr)
    err.Columns(1).ColumnWidth = 100
    err.Columns(2).ColumnWidth = 20
    err.Cells(1, 1) = "Файл"
    err.Cells(1, 2) = "Результат"
    Range(err.Cells(1, 1), err.Cells(1, 100)).Interior.Color = RGB(214, 214, 214)
    recN = 2
End Sub

Sub Rec(ByVal file As String, ByVal code As Integer)
    msg = "Неопознанная ошибка"
    If code = 1 Then msg = "Ошибка загрузки файла"
    If code = 2 Then msg = "Ошибка в данных"
    If code = 3 Then msg = "Отсутствует код"
    If code = 4 Then msg = "Дубликат! Обработка пропущена"
    err.Cells(recN, 1) = file
    err.Cells(recN, 2) = msg
    recN = recN + 1
End Sub