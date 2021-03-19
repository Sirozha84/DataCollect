Attribute VB_Name = "Misc"
'��������� � ������ �������
Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub

'�������� �����
Public Sub MakeDir(ByVal name As String)
    On Error Resume Next
    MkDir (name)
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

'�������� �� ����������� ����������
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

'������ �������� �� ������ �������
Function IndexToQuartal(ByVal i As Integer) As String
    IndexToQuartal = CStr(lastYear - Int((lastQuartal + i) / 4) + 1) + CStr(4 - i Mod 4)
End Function

'��������� �� ���� ���+�������
Function Kvartal(sdata As Variant) As String
    On Error GoTo er
    Kvartal = CStr(Year(sdata)) + CStr((Month(sdata) - 1) \ 3 + 1)
    Exit Function
er:
    Kvartal = ""
End Function

'��������� ������ �������� �� ����
'���� ������ ��� ���� � �� ��������� - ���������� -1
Function DateToQIndex(ByVal d As Date) As Integer
    On Error GoTo er
    i = (lastYear - Year(d)) * 4
    i = i + lastQuartal - ((Month(d) - 1) \ 3) - 1
    If i < 0 Or i >= quartCount Then i = -1
    DateToQIndex = i
    Exit Function
er:
    DateToQIndex = -1
End Function

'��������� ������ �������� �� ������ ���� �����
'���� ������ ��� ���� � �� ��������� - ���������� -1
Function StupidQToQIndex(ByVal d As String) As Integer
    On Error GoTo er
    i = (lastYear - CInt(Right(d, 4))) * 4
    i = i + lastQuartal - (CInt(Left(d, 1))) '- 1
    If i < 0 Or i >= quartCount Then i = -1
    StupidQToQIndex = i
    Exit Function
er:
    StupidQToQIndex = -1
End Function

'��������� ������
'���������� True, ���� �������, ���� ������������ False - ������ ������ �� �������
Function SetProtect(table As Variant) As Boolean
    On Error GoTo er
    table.Protect Secret, AllowFormattingColumns:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    SetProtect = True
    Exit Function
er:
    SetProtect = False
End Function