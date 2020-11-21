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
    Set curfold = FSO.getfolder(pat)
    For Each file In curfold.files
        If file.name Like "*.xls*" Then files.Add file.Path
    Next
    For Each subfolder In curfold.subFolders
         readDir subfolder
    Next subfolder
End Sub