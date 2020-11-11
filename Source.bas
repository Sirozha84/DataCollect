Attribute VB_Name = "Source"
Public Path As String
Public FSO As Object

'Получение списка файлов в заданной директории
Function GetList(ByVal pat As String) As Collection
    
    Path = pat
    
    Set Files = New Collection
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set curfold = FSO.GetFolder(Cells(1, 3))
    
    For Each file In curfold.Files
        If file.name Like "*.xls*" Then Files.Add file.Path
    Next
    
    Set GetList = Files
End Function