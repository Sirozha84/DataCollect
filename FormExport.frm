VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormExport 
   Caption         =   "Выгрузка данных"
   ClientHeight    =   2520
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4557
   OleObjectBlob   =   "FormExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstDate As Date
Dim LastDate As Date

'Инициализация выпадающих списков
Private Sub UserForm_Initialize()
    
    Verify.Init
    
    'Выпадающий список продавцов
    ComboBoxBuyers.AddItem "Все"
    For Each seller In selIndexes
        ComboBoxBuyers.AddItem SellFileName(seller)
    Next
    ComboBoxBuyers.ListIndex = 0
        
    'Период сбора
    TextBoxFirstCollect = PRP.Cells(8, 2)
    TextBoxLastCollect = PRP.Cells(9, 2)
    
End Sub

'Имя файла по ИНН продавца
Function SellFileName(INN) As String
    ind = selIndexes(INN)
    If ind <> Empty Then SellFileName = INN + "-" + DIC.Cells(ind, 1)
End Function

Private Sub CommandExit_Click()
    End
End Sub

'Кнопка "Экспорт"
Private Sub CommandExport_Click()
    
    On Error GoTo er
    FirstDate = CDate(TextBoxFirstCollect)
    LastDate = CDate(TextBoxLastCollect)
    On Error GoTo 0
    
    If ComboBoxBuyers.ListIndex = 0 Then
        n = 1
        a = selIndexes.Count
        For Each seller In selIndexes
            ExportFile seller, CStr(n) + " из " + CStr(a) + ": "
            n = n + 1
        Next
    Else
        ExportFile Left(ComboBoxBuyers.Value, 10), ""
    End If
    
    'Сохранение дат периода сбора
    PRP.Cells(8, 2) = TextBoxFirstCollect
    PRP.Cells(9, 2) = TextBoxLastCollect
    
    Message "Готово!"
    End

er:
    MsgBox "Даты не введены или введены не корректно"

End Sub

'Экспорт файла
Private Sub ExportFile(ByVal INN As String, NUM As String)

    seller = SellFileName(INN)
    Message "Экспорт файла " + NUM + seller
    
    'Проверка обязательных данных в справочнике
    si = selIndexes(INN)
    Limit = DIC.Cells(si, cLimND)
    ermsg = "У продавца " + DIC.Cells(si, 1) + " с ИНН " + INN + " "
    If Limit = Empty Then
        MsgBox ermsg + "не указан лимит!"
        End
    End If
    If StupidQToQIndex(DIC.Cells(si, cOPND)) < 0 Then
        MsgBox ermsg + "не указан или указан не корректно основной период НД!"
        End
    End If
    
    'Определяемся с путём и именем файла
    Patch = DirExport + "\Отгрузки"
    MakeDir Patch
    fol = ""
    If fol <> "" Then MakeDir (Patch + fol)
    fileName = Patch + fol + "\" + cutBadSymbols(seller) + ".xlsx"
    
    'Создаём книгу
    Workbooks.Add
    i = 1
    Cells(i, 1) = "Код вида" + Chr(10) + "операции"
    Cells(i, 2) = "№ счет" + Chr(10) + "фактуры"
    Cells(i, 3) = "Дата счет" + Chr(10) + "фактуры"
    Cells(i, 4) = "ИНН"
    Cells(i, 5) = "КПП"
    Cells(i, 6) = "Наименование"
    Cells(i, 7) = "Сумма в руб." + Chr(10) + "и коп."
    Cells(i, 8) = "Сумма" + Chr(10) + "без НДС 20%"
    Cells(i, 9) = "Сумма" + Chr(10) + "без НДС 18%"
    Cells(i, 10) = "Сумма" + Chr(10) + "без НДС 10%"
    Cells(i, 11) = "НДС 20%"
    Cells(i, 12) = "НДС 18%"
    Cells(i, 13) = "НДС 10%"
    Cells(i, 14) = "Период НД"
    Columns(1).ColumnWidth = 10
    Columns(2).ColumnWidth = 13
    Columns(3).ColumnWidth = 10
    Columns(4).ColumnWidth = 11
    Columns(5).ColumnWidth = 10
    Columns(6).ColumnWidth = 15
    Columns(7).ColumnWidth = 12
    Columns(8).ColumnWidth = 12
    Columns(9).ColumnWidth = 12
    Columns(10).ColumnWidth = 12
    Columns(11).ColumnWidth = 10
    Columns(12).ColumnWidth = 10
    Columns(13).ColumnWidth = 10
    Columns(14).ColumnWidth = 10
    Rows(1).RowHeight = 30
    Set hat = Range(Cells(1, 1), Cells(1, 14))
    hat.HorizontalAlignment = xlCenter
    hat.VerticalAlignment = xlCenter
    hat.Interior.Color = colGray
    hat.Borders.Weight = 3
    firstEx = i + 1
    
    'Заполняем книгу
    i = firstDat
    j = firstEx
    Do While DAT.Cells(i, cAccept) <> ""
        If DAT.Cells(i, cAccept) = "OK" Then
            dc = DAT.Cells(i, cDateCol)
            If dc >= FirstDate And dc < LastDate + 1 Then
                cp = True
                If DAT.Cells(i, cSellINN).text <> INN Then cp = False
                d = DAT.Cells(i, cDates)
                If cp Then
                    'Копирование данных из сбора
                    Cells(j, 1).NumberFormat = "@"
                    Cells(j, 1) = "01"
                    Cells(j, 2) = DAT.Cells(i, 1)
                    Cells(j, 3).NumberFormat = "dd.MM.yyyy"
                    Cells(j, 3) = DAT.Cells(i, 2)
                    innkpp = Split(DAT.Cells(i, 3), "/")
                    Cells(j, 4).NumberFormat = "@"
                    Cells(j, 4) = innkpp(0)
                    Cells(j, 5).NumberFormat = "@"
                    If UBound(innkpp) > 0 Then Cells(j, 5) = innkpp(1)
                    Cells(j, 6) = DAT.Cells(i, 4)
                    Cells(j, 7).NumberFormat = "### ### ##0.00"
                    Cells(j, 7) = DAT.Cells(i, 7)
                    For c = 0 To 5
                        Cells(j, 8 + c).NumberFormat = "### ### ##0.00"
                        Cells(j, 8 + c) = DAT.Cells(i, 9 + c)
                    Next
                    'Временные колонки - индекс квартала и сумма НДС
                    Cells(j, 15) = DateToQIndex(DAT.Cells(i, 2))
                    Sum = 0
                    For j2 = 11 To 13
                        If IsNumeric(Cells(j, j2)) Then Sum = Sum + Cells(j, j2)
                    Next
                    Cells(j, 16) = Sum
                    j = j + 1
                End If
            End If
        End If
        i = i + 1
    Loop
    
    'Сортировка по периодам
    Cells(1, 15) = "Квартал"
    Cells(1, 16) = "НДС"
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("O2") 'Первый порядок сортировки
        .SortFields.Add Key:=Range("F2") 'Второй порядок сортировки
        .setRange Range("A2:P" + CStr(j - 1)) 'Диапазон сортируемой таблицы
        .Apply
    End With
    
    'Заполняем периоды НД
    PeriodND si
    
    'Удаление временных столбцов
    Columns(15).Delete
    Columns(15).Delete
    
    'Сохранение и закрытие документа
    On Error GoTo er
    Application.DisplayAlerts = False
    If j > firstEx Then ActiveWorkbook.SaveAs fileName:=fileName
    ActiveWorkbook.Close
    Exit Sub

er:
    ActiveWorkbook.Close
    MsgBox "Произошла ошибка при сохранении файла " + fileName

End Sub

'Расчёт периодов налоговой декларации
'sIndex - индекс продавца
Sub PeriodND(ByVal sIndex As Double)
    
    Dim oND As Integer
    oND = StupidQToQIndex(DIC.Cells(sIndex, cOPND))
    
    '******************** Первый этап ********************
    
    'Составляем список минимальных значений по каждому ИНН для периода oND
    Set ni = CreateObject("Scripting.Dictionary")   'Индексы по ИНН
    Set ns = CreateObject("Scripting.Dictionary")   'Суммы по ИНН
    i = 2
    Do While Cells(i, 1) <> ""
        INN = Cells(i, 4)
        If Cells(i, 15) = oND Then
            s = Cells(i, 16)
            If ns(INN) = 0 Or ns(INN) > s Then
                ns(INN) = s
                ni(INN) = i
            End If
        End If
        i = i + 1
    Loop
    
    'Проверяем сумму всех найденных записей на лимит оНД
    Do
        Sum = 0
        For Each i In ni
            Sum = Sum + ns(i)
        Next
        per = Sum - limitOND
        'И.. если сумма превышает лимит...
        If per > 0 Then
            'Находим запись, которая ближе всего к сумме превышения (per)
            Min = 0         'Минимальная разница
            isk = ""        'Запись, которую надо исключить
            plus = False    'Есть ли пололжительная разница?
            For Each i In ni
                If ns(i) <> 0 Then
                    r = ns(i) - per
                    If r >= 0 Then
                        If plus = False Or Min > r Then
                            Min = r
                            isk = i
                        End If
                        plus = True
                    End If
                    If r < 0 And Not plus Then
                        If Min = 0 Or Min < r Then
                            Min = r
                            isk = i
                        End If
                    End If
                End If
            Next
            'Переносим найденную запись в очередь
            ns.Remove (isk)
            ni.Remove (isk)
        End If
    Loop Until per <= 0
    
    'Расставим период НД оставшимся записям
    pnd = IndexToQYYYY(oND)
    For Each i In ni
        Cells(ni(i), 14) = pnd
    Next
    
    '******************** Второй этап ********************
    
    Dim tND As Integer  'Текущий период
    Dim Qi As Object    'Очередь записей (Сумма по индексу)
    'Соберём записи, которые "не влезли" в основном периоде
    Set Qi = CreateObject("Scripting.Dictionary")
    i = 2
    Do While Cells(i, 1) <> ""
        If Cells(i, 14) = "" And Cells(i, 15) = oND Then Qi(i) = Cells(i, 16)
        i = i + 1
    Loop
    tND = oND 'Основной период теперь текущий период, с него начнётся цикл второго этапа
    
    Do
        
        'Переходим к следующему периоду
        tND = tND + 1
        Dim Limit As Double
        Limit = DIC.Cells(sIndex, cLimND) - DIC.Cells(sIndex, cCorrect + tND)
        
        'Составляем список записей текущего периода
        Set ti = CreateObject("Scripting.Dictionary")  'Список записей текущего периода (Сумма по индексу)
        i = 2
        Do While Cells(i, 1) <> ""
            If Cells(i, 15) = tND Then ti(i) = Cells(i, 16)
            i = i + 1
        Loop
        
        'Проверяем сумму всех записей периода на лимит НД
        Do
            s = 0
            For Each i In ti
                s = s + ti(i)
            Next
            per = s - Limit
            If per > 0 Then
                If Limit < minLimit Then
                    'В этом случае период пропускаем
                    'Переносим все записи в очередь
                    For Each i In ti
                        Qi.Add i, ti(i)
                        ti.Remove (i)
                    Next
                Else
                    'Переносим запись с максимальным значением
                    maxs = 0    'Максимальное значение
                    msxi = 0    'Индекс записи с максимальным значением
                    For Each i In ti
                        If maxs < ti(i) Then
                            maxs = ti(i)
                            maxi = i
                        End If
                    Next
                    Qi.Add maxi, ti(maxi)
                    ti.Remove (maxi)
                End If
            End If
            
        Loop Until per <= 0
        
        'Расставим период НД оставшимся записям
        ost = -per
        pnd = IndexToQYYYY(tND)
        For Each i In ti
            Cells(i, 14) = pnd
        Next
        
        'Если очередь не пуста и осталось "место",
        'расставляем данные из неё, начиная с минимального значения
        If Qi.Count > 0 Then
            Do
                mins = 0    'Минимальное значение
                mini = 0    'Индекс минимального значения
                For Each i In Qi
                    If mins = 0 Or mins > Qi(i) Then
                        mins = Qi(i)
                        mini = i
                    End If
                Next
                Enter = ost >= mins
                If Enter Then
                    Cells(mini, 14) = pnd
                    Qi.Remove (mini)
                    ost = ost - mins
                End If
            Loop Until Not Enter Or Qi.Count = 0
        End If
        
        Debug.Print "Очередь на период " + CStr(tND)
        For Each i In Qi: Debug.Print i: Next
        
    Loop While tND < quartCount - 1
       
    'Сууууукаааа, вроде это всё работает!!!!!
       
    'Ещё надо подумать что делать, если все периоды обработаны, а данные в очереди остались.
    End
    
End Sub