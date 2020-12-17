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