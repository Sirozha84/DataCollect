Attribute VB_Name = "Source"
Public files As Collection
Public Path As String
Public FSO As Object
Private curfold As Variant

'��������� ������ ������ � �������� ����������
Function getFiles(ByVal pat As String) As Collection
    Path = pat
    Set files = New Collection
    Set FSO = CreateObject("Scripting.FileSystemObject")
    readDir pat
    'DuplicateFinder
    Set getFiles = files
End Function

'������ ������ ������ � �����
Private Sub readDir(ByVal pat As String)
    On Error GoTo er:
    If InStr(1, pat, ".sync") > 0 Then Exit Sub
    Set curfold = FSO.GetFolder(pat)
    For Each file In curfold.files
        If file.name Like "*.xls*" And _
                InStr(1, file.name, "~$") = 0 And _
                InStr(1, file.name, "������ ") = 0 _
            Then files.Add file.Path
    Next
    For Each subfolder In curfold.subFolders
         readDir subfolder
    Next subfolder
er:
End Sub

'���������� ���� �� �����
Function GetCode(ByVal file As String) As String
    If Not TrySave(file) Then Exit Function
    On Error GoTo er
    Application.ScreenUpdating = False
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    If Not impBook Is Nothing Then
        Set SRC = impBook.Worksheets(1)
        SetProtect SRC
        GetCode = SRC.Cells(1, 1)
        impBook.Close False
    End If
er:
End Function

Sub RenameFile(ByVal file As String, ByVal c As String)
    ins = " ��������! ��� ����� "
    On Error GoTo er
    If InStr(1, file, ins) = 0 Then
        newname = Replace(file, ".xls", ins + c + ".xls")
        Name file As newname
    End If
er:
End Sub