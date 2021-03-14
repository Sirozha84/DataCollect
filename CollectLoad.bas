Attribute VB_Name = "CollectLoad"
Sub Run()
    
    'Получаем коллекцию файлов
    Set files = Source.getFiles(DirImportLoad, False)

    For Each file In files
        curf = file
        If Len(curf) > 40 Then curf = "..." + Right(curf, 40)
        Message ("Обработка файла " + CStr(n) + " из " + CStr(files.Count) + " (" + curf) + ")"
        er = AddFile(file)
        If er > 0 Then
            Log.Rec file, er
            e = e + 1
        Else
            s = s + 1
        End If
        n = n + 1
    Next

    Message "Готово!"
    
End Sub