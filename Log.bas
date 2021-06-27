Attribute VB_Name = "Log"
'Последняя правка: 27.06.2021 08:36

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
    If code = 6 Then msg = "Файл заблокирован"
    If code = 7 Then msg = "Отсутствует маркер, либо он не верный"
    If code = 8 Then msg = "Поля не распознаны"
    ERR.Cells(recN, 1) = file
    ERR.Cells(recN, 2) = msg
    recN = recN + 1
End Sub

'******************** End of File ********************