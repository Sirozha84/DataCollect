Attribute VB_Name = "Misc"
Private SearchMethod As Byte

'Сообщение в строку статуса
Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub

'Создание таблицы
Sub NewTab(name As String, cl As Boolean)
    Application.ScreenUpdating = False
    Set cur = ActiveSheet
    
    name = Left(name, 31)
    If Not SheetExist(name) Then
        Sheets.Add(, Sheets(Sheets.Count)).name = name
    End If
    If cl Then Sheets(name).Cells.Clear
    cur.Activate
    Application.ScreenUpdating = True
End Sub

'Проверка на существование листа
Private Function SheetExist(name As String) As Boolean
    Dim objSheet As Object
    On Error GoTo HandleError
    ThisWorkbook.Worksheets(name).Activate
    SheetExist = True
    Exit Function
HandleError:
    SheetExist = False
End Function

'Поиск максимального количества строк (последняя строка где первая колонка непрерывно заполнена)
Function FindMax(ByVal name As String) As Long
    i = 0
    Do While Sheets(name).Cells(i + 1000, 1) <> ""
        i = i + 1000
    Loop
    Do While Sheets(name).Cells(i + 1, 1) <> ""
        i = i + 1
    Loop
    FindMax = i
End Function

'Расчёт прогресса (сколько из скольки + процент)
Function Progress(ByVal cur As Long, ByVal all As Long)
    Progress = text + ":" + str(cur) + " из" + str(all) + " (" + str(Int(cur / all * 100)) + "% )"
End Function

'Выбор метода поиска (проверка на сортировку, если сортировано - бинарный, если нет - перебор
Sub MethodSelect(ByVal name As String, ByVal first As Long, ByVal last As Long)
    For i = first To last - 1
        If StrComp(Sheets(name).Cells(i, 1), Sheets(name).Cells(i + 1, 1), vbTextCompare) > 0 Then
            SearchMethod = 1
            Exit For
        End If
    Next
End Sub

'Поиск значения в таблице
Function Search(ByVal name As String, ByVal str As String, ByVal first As Long, ByVal last As Long) As Long
    
    If SearchMethod = 0 Then
        
        'Бинарный поиск
        Find = 0
        Do
            middle = first + Int((last - first) / 2)
            If StrComp(str, Sheets(name).Cells(first, 1), vbTextCompare) = 0 Then Find = first
            If StrComp(str, Sheets(name).Cells(last, 1), vbTextCompare) = 0 Then Find = last
            If StrComp(str, Sheets(name).Cells(middle, 1), vbTextCompare) = 0 Then Find = middle
            If StrComp(str, Sheets(name).Cells(middle, 1), vbTextCompare) < 0 Then last = middle
            If StrComp(str, Sheets(name).Cells(middle, 1), vbTextCompare) > 0 Then first = middle
        Loop Until Find > 0 Or last - first < 2
    
    Else
        
        'Поиск перебором
        For i = first To last
            If str = Sheets(name).Cells(i, 1) Then
                Find = i
                Exit For
            End If
        Next
    
    End If
    Search = Find

End Function