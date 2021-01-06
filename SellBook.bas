Attribute VB_Name = "SellBook"
Dim Patch As String
Dim Buyers As Collection
Dim Sellers As Collection
Dim Where As Collection
Dim Quartals As Object
Dim i As Long

Public Sub ExportBook(ByVal file As String)

    Main.Init
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Patch = FSO.GetParentFolderName(file) + "\"
    
    GetLists file
    GetQuartalsAndIndexes

    For Each q In Quartals
        For Each b In Buyers
            For Each s In Sellers
                MakeBook q, b, s
            Next
        Next
    Next
    
    Message "Готово!"
End Sub

'Чтение справочников покупателей и продавцов из шаблона
Sub GetLists(ByVal file As String)
    Message "Чтение данных из шаблона"
    Set Buyers = New Collection
    Set Sellers = New Collection
    On Error GoTo er
    Application.ScreenUpdating = False
    Set templ = Workbooks.Open(file, False, False)
    If Not templ Is Nothing Then
        Set SRC = templ.Worksheets("Покупатели")
        i = 2
        Do While SRC.Cells(i, 1) <> ""
            Buyers.Add SRC.Cells(i, 1).text, SRC.Cells(i, 1).text
            i = i + 1
        Loop
        Set SRC = templ.Worksheets("Продавцы")
        i = 2
        Do While SRC.Cells(i, 1) <> ""
            Sellers.Add SRC.Cells(i, 1).text, SRC.Cells(i, 1).text
            i = i + 1
        Loop
        templ.Close False
    End If
    Exit Sub
er:
    MsgBox "Произошла ошибка при открытии файла шаблона"
    End
End Sub

'Чтение базы, подготовка списка кварталов и индексов строк, в которых есть совпадения
Sub GetQuartalsAndIndexes()
    Message "Подготовка списка кварталов"
    Set Where = New Collection
    Set Quartals = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    i = firstDat
    Do While Cells(i, 4) <> "" Or Cells(i, cSeller) <> ""
        If Cells(i, 1) <> "" And Cells(i, cDates) <> "" And _
                Cells(i, cBuyer) <> "" And Cells(i, cSeller) <> "" Then
            b = ""
            b = Buyers(Cells(i, cBuyer))
            s = ""
            s = Sellers(Cells(i, cSeller))
            If b <> "" And s <> "" Then
                Where.Add i
                Quartals(GetQuartal(Cells(i, cDates))) = 1
            End If
        End If
        i = i + 1
    Loop
End Sub

'Вычисление номера квартала в формате "1-20"
Function GetQuartal(d As Date) As String
    GetQuartal = CStr((Month(d) - 1) \ 3 + 1) + "-" + Right(CStr(Year(d)), 2)
End Function

'Вычисление периода по номеру квартала
Function Period(q As String)
    y = ".20" + Right(q, 2)
    If Left(q, 1) = "1" Then Period = "с 01.01" + y + " по 31.03" + y
    If Left(q, 1) = "2" Then Period = "с 01.04" + y + " по 30.06" + y
    If Left(q, 1) = "3" Then Period = "с 01.07" + y + " по 30.09" + y
    If Left(q, 1) = "4" Then Period = "с 01.10" + y + " по 31.12" + y
End Function
'Формирование книги
Sub MakeBook(ByVal q As String, ByVal b As String, ByVal s As String)
    
    'Поиск данных для текущей комбинации квартал+покупатель+продавец
    Dim Finded As Collection
    Set Finded = New Collection
    For Each j In Where
        If q = GetQuartal(DAT.Cells(j, 2)) And b = DAT.Cells(j, cBuyer).text And _
            s = DAT.Cells(j, cSeller).text Then Finded.Add j
    Next
    If Finded.Count = 0 Then Exit Sub
    
    'Какие-то данные всёже нашли, создаём книгу
    name = "КнПрод " + s + " - " + b + " " + q
    filename = Patch + name + ".xlsx"
    Message "Формирование книги " + name
    Workbooks.Add
    Application.DisplayAlerts = False
    
    'Шапка
    bigCell "Книга продаж", 1, 1, 1, 24
    Cells(1, 1).Font.Size = 14
    Cells(1, 1).Font.Bold = True
    Cells(3, 1) = "Продавец " + s
    Cells(4, 1) = "Идентификационный номер и код причины постановки на учeт налогоплательщика-продавца " + _
        DAT.Cells(Finded(1), cSellINN).text
    Cells(5, 1) = "Продажа за период " + Period(q)
    Cells(6, 1) = "Отбор: Контрагент = " + DAT.Cells(Finded(1), cBuyer)
    Cells(6, 1).Font.Bold = True
    bigCell "№" + Chr(10) + "п/п", 7, 1, 2, 1
    Cells(9, 1) = "1"
    Cells(9, 2) = "2"
    Cells(9, 3) = "3"
    Cells(9, 4) = "3а"
    Cells(9, 5) = "3б"
    Cells(9, 6) = "4"
    Cells(9, 7) = "5"
    Cells(9, 8) = "6"
    Cells(9, 9) = "7"
    Cells(9, 10) = "8"
    Cells(9, 11) = "9"
    Cells(9, 12) = "10"
    Cells(9, 13) = "11"
    Cells(9, 14) = "12"
    Cells(9, 15) = "13а"
    Cells(9, 16) = "13б"
    Cells(9, 17) = "14"
    Cells(9, 18) = "14а"
    Cells(9, 19) = "15"
    Cells(9, 20) = "16"
    Cells(9, 21) = "17"
    Cells(9, 22) = "17а"
    Cells(9, 23) = "18"
    Cells(9, 24) = "19"
    Range(Cells(7, 1), Cells(9, 24)).HorizontalAlignment = xlCenter
    Range(Cells(7, 1), Cells(9, 24)).VerticalAlignment = xlCenter
    'Строки
    i = 10
    n = 1
    s1 = 0: s2 = 0: s3 = 0: s4 = 0: s5 = 0: s6 = 0
    For Each j In Finded
        Cells(i, 1) = n
        Cells(i, 2) = 1
        Cells(i, 3) = DAT.Cells(j, 1).text + " от " + DAT.Cells(j, cDates).text
        Cells(i, 9) = DAT.Cells(j, cBuyer)
        Cells(i, 10) = DAT.Cells(j, cBuyINN)
        Cells(i, 16) = DAT.Cells(j, cPrice)
        Cells(i, 17) = DAT.Cells(j, 9): If Cells(i, 17) <> "" Then s1 = s1 + Cells(i, 17)
        Cells(i, 18) = DAT.Cells(j, 10): If Cells(i, 18) <> "" Then s2 = s2 + Cells(i, 18)
        Cells(i, 19) = DAT.Cells(j, 11): If Cells(i, 19) <> "" Then s3 = s3 + Cells(i, 19)
        Cells(i, 21) = DAT.Cells(j, 12): If Cells(i, 21) <> "" Then s4 = s4 + Cells(i, 21)
        Cells(i, 22) = DAT.Cells(j, 13): If Cells(i, 22) <> "" Then s5 = s5 + Cells(i, 22)
        Cells(i, 23) = DAT.Cells(j, 14): If Cells(i, 23) <> "" Then s6 = s6 + Cells(i, 23)
        i = i + 1
        n = n + 1
    Next
    
    'Подвал
    If s1 > 0 Then Cells(i, 17) = s1
    If s2 > 0 Then Cells(i, 18) = s2
    If s3 > 0 Then Cells(i, 19) = s3
    If s4 > 0 Then Cells(i, 21) = s4
    If s5 > 0 Then Cells(i, 22) = s5
    If s6 > 0 Then Cells(i, 23) = s6
    Range(Cells(7, 1), Cells(i, 24)).Borders.Weight = 2
    
    End
    
    On Error Resume Next
    ActiveWorkbook.SaveAs filename:=filename
    ActiveWorkbook.Close

End Sub

Sub bigCell(txt As String, r As Variant, c As Variant, height As Variant, width As Variant)
    Cells(r, c) = txt
    Range(Cells(r, c), Cells(r + height - 1, c + width - 1)).merge
    Cells(r, c).HorizontalAlignment = xlCenter
    Cells(r, c).VerticalAlignment = xlCenter
End Sub