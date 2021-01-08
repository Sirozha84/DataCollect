Attribute VB_Name = "Misc"
'Сообщение в строке статуса
Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub

'Создание папки
Sub folder(name As String)
    On Error GoTo er
    MkDir (name)
er:
End Sub

'Удаление неугодных символов для имени файла
Function cutBadSymbols(ByVal name As String) As String
    name = Replace(name, """", "")
    name = Replace(name, "*", "")
    name = Replace(name, "\", "")
    name = Replace(name, "|", "")
    name = Replace(name, "/", "")
    name = Replace(name, "?", "")
    name = Replace(name, ":", "")
    name = Replace(name, "<", "")
    name = Replace(name, ">", "")
    cutBadSymbols = name
End Function

'Проверка файла на то что он уже открыт
Function IsBookOpen(fileName As String) As Boolean
    Dim FSO As Object, strFileName$, strFilePath$
    Set FSO = CreateObject("Scripting.FileSystemObject")
    name = FSO.getfilename(fileName)
    Dim wbBook As Workbook
    On Error Resume Next
    Set wbBook = Workbooks(name)
    IsBookOpen = Not wbBook Is Nothing
End Function