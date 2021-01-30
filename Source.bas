Attribute VB_Name = "Source"
Public files As Collection
Public Path As String
Public FSO As Object
Private curfold As Variant

'Получение списка файлов в заданной директории
Function getFiles(ByVal pat As String, FindDuplicate As Boolean) As Collection
    Path = pat
    Set files = New Collection
    Set FSO = CreateObject("Scripting.FileSystemObject")
    readDir pat
    If FindDuplicate Then DuplicateFinder
    Set getFiles = files
End Function

'Чтение списка файлов и папок
Private Sub readDir(ByVal pat As String)
    On Error GoTo er:
    If InStr(1, pat, ".sync") > 0 Then Exit Sub
    Set curfold = FSO.GetFolder(pat)
    For Each file In curfold.files
        If file.name Like "*.xls*" And _
                InStr(1, file.name, "~$") = 0 And _
                InStr(1, file.name, "КнПрод ") = 0 _
            Then files.Add file.Path
    Next
    For Each subfolder In curfold.subFolders
         readDir subfolder
    Next subfolder
er:
End Sub

'Проверка файлов на наличие дубликатов по коду
Sub DuplicateFinder()
    Set Codes = CreateObject("Scripting.Dictionary")
    Set dupl = New Collection
    n = 1
    a = files.Count
    For Each file1 In files
        Message ("Проверка на дубликаты: " + CStr(n) + " из " + CStr(a))
        c = GetCode(file1)
        file2 = Codes(c)
        If Codes(c) = Empty Then
            Codes(c) = file1
        Else
            dupl.Add file1 + "*" + c
            dupl.Add file2 + "*" + c
        End If
        n = n + 1
    Next
    'Удаляем из списка файлов файлы, найденные как дубликаты и переименовываем их
    For Each file In dupl
        s = Split(file, "*")
        f = s(0)
        c = s(1)
        Log.Rec f, 5
        RenameFile f, c
        For i = files.Count To 1 Step -1
            If files(i) = f Then files.Remove (i)
        Next
    Next
End Sub

'Извлечение кода из файла
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
    ins = " Дубликат! Код формы "
    On Error GoTo er
    If InStr(1, file, ins) = 0 Then
        newname = Replace(file, ".xls", ins + c + ".xls")
        Name file As newname
    End If
er:
End Sub