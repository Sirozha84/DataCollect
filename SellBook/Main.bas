Attribute VB_Name = "Main"
Public Patch As String

Public Sub ButtonGenerate()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Patch = diag.SelectedItems(1)
    
    Set files = getFiles(Patch)
    Range(Cells(7, 1), Cells(1000000, 2)).Clear
    i = 7
    For Each file In files
        Cells(i, 1) = file
        er = ExportBook(file)
        If er = 0 Then Cells(i, 2) = "Ошибка при работе с файлом"
        If er = 1 Then Cells(i, 2) = "Созданы книги продаж"
        If er = 2 Then Cells(i, 2) = "Реестр имеет некорректные записи"
        i = i + 1
    Next
    MsgBox "Формирование книг продаж завершено!"
End Sub