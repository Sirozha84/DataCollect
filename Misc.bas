Attribute VB_Name = "Misc"
'��������� � ������ �������
Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub

'�������� �����
Sub folder(name As String)
    On Error GoTo er
    MkDir (name)
er:
End Sub

'�������� ��������� �������� ��� ����� �����
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