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

'Проверка на возможность созранениф
Function TrySave(file As Variant)
    On Error GoTo er
    newname = file + "_"
    Name file As newname
    Name newname As file
    TrySave = True
    Exit Function
er:
    TrySave = False
End Function

'Расчёт квартала по номеру индекса
Function IndexToQuartal(ByVal i As Integer) As String
    IndexToQuartal = CStr(lastYear - Int((lastQuartal + i) / 4) + 1) + CStr(4 - i Mod 4)
End Function