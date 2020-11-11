Attribute VB_Name = "Misc"
Private SearchMethod As Byte

'Сообщение в строке статуса
Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub

'Создание таблицы
Sub NewTab(name As String, cl As Boolean)
    Application.ScreenUpdating = False
    Set cur = ActiveSheet
    name = Left(name, 31)
    If Not SheetExist(name) Then
        Sheets.Add(, Sheets(Sheets.Count)).name = name
    End If
    If cl Then Sheets(name).Cells.Clear
    cur.Activate
    Application.ScreenUpdating = True
End Sub

'Проверка на существование листа
Private Function SheetExist(name As String) As Boolean
    Dim objSheet As Object
    On Error GoTo HandleError
    ThisWorkbook.Worksheets(name).Activate
    SheetExist = True
    Exit Function
HandleError:
    SheetExist = False
End Function