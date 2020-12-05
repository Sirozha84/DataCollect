Attribute VB_Name = "Source"
Public files As Collection
Public Path As String
Public FSO As Object
Private curfold As Variant

'Получение списка файлов в заданной директории
Function getFiles(ByVal pat As String) As Collection
    Path = pat
    Set files = New Collection
    Set FSO = CreateObject("Scripting.FileSystemObject")
    readDir pat
    Set getFiles = files
End Function

'Чтение списка файлов и папок
Private Sub readDir(ByVal pat As String)
    On Error GoTo er:
    If InStr(1, pat, ".sync") > 0 Then Exit Sub
    Set curfold = FSO.getfolder(pat)
    For Each file In curfold.files
        If file.name Like "*.xls*" And InStr(1, file.name, "~$") = 0 Then files.Add file.Path
    Next
    For Each subfolder In curfold.subFolders
         readDir subfolder
    Next subfolder
er:
End Sub