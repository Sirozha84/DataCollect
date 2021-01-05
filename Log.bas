Attribute VB_Name = "Log"
Dim recN As Long    'Текущий номер строки

'Инициализация
Sub Init()
    ERR.Cells.Clear
    ERR.Columns(1).ColumnWidth = 100
    ERR.Columns(2).ColumnWidth = 30
    ERR.Cells(1, 1) = "Файл"
    ERR.Cells(1, 2) = "Результат"
    Range(ERR.Cells(1, 1), ERR.Cells(1, 100)).Interior.Color = colGray
    recN = firstErr
End Sub

'Запись ошибки
Sub Rec(ByVal file As String, ByVal code As Integer)
    msg = "Неопознанная ошибка"
    If code = 1 Then msg = "Ошибка загрузки файла"
    If code = 2 Then msg = "Ошибка в данных"
    If code = 3 Then msg = "Отсутствует код"
    If code = 4 Then msg = "Версия формы не поддерживается"
    If code = 5 Then msg = "Дубликат! Обработка пропущена"
    ERR.Cells(recN, 1) = file
    ERR.Cells(recN, 2) = msg
    recN = recN + 1
End Sub