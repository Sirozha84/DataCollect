VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormExport 
   Caption         =   "Выгрузка данных"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "FormExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sellers As Variant

'Инициализация выпадающих списков
Private Sub UserForm_Initialize()
    
    Message "Подготовка к выгрузке..."
    Main.Init
    Set Sellers = CreateObject("Scripting.Dictionary")
    Set Quartals = CreateObject("Scripting.Dictionary")
    Set Months = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    i = firstDat
    Do While Cells(i, 2) <> "" Or Cells(i, 15) <> ""
        Sellers(Cells(i, cSeller).text) = 1
        d = Cells(i, cDates)
        Months(YearAndMonth(d)) = 1
        Quartals(YearAndQuartal(d)) = 1
        i = i + 1
    Loop
    
    ComboBoxBuyers.AddItem "Все"
    For Each Seller In Sellers
        ComboBoxBuyers.AddItem Seller
    Next
    ComboBoxBuyers.ListIndex = 0
    
    For Each m In Months
        ComboBoxMonths.AddItem m
    Next
    ComboBoxMonths.ListIndex = 0
    
    For Each q In Quartals
        ComboBoxQuartals.AddItem q
    Next
    ComboBoxQuartals.ListIndex = 0
    
    Message "Готово!"
    
End Sub

Function YearAndMonth(ByVal d As Date)
    YearAndMonth = CStr(Year(d)) + " - "
    dy = Month(d)
    If dy = 1 Then YearAndMonth = YearAndMonth + "Январь"
    If dy = 2 Then YearAndMonth = YearAndMonth + "Февраль"
    If dy = 3 Then YearAndMonth = YearAndMonth + "Март"
    If dy = 4 Then YearAndMonth = YearAndMonth + "Апрель"
    If dy = 5 Then YearAndMonth = YearAndMonth + "Май"
    If dy = 6 Then YearAndMonth = YearAndMonth + "Июнь"
    If dy = 7 Then YearAndMonth = YearAndMonth + "Июль"
    If dy = 8 Then YearAndMonth = YearAndMonth + "Август"
    If dy = 9 Then YearAndMonth = YearAndMonth + "Сентябрь"
    If dy = 10 Then YearAndMonth = YearAndMonth + "Октябрь"
    If dy = 11 Then YearAndMonth = YearAndMonth + "Ноябрь"
    If dy = 12 Then YearAndMonth = YearAndMonth + "Декадрь"
End Function

Function YearAndQuartal(ByVal d As Date)
    YearAndQuartal = CStr(Year(d)) + " - " + CStr((Month(d) - 1) \ 3 + 1) + " квартал"
End Function

Private Sub OptionAll_Click()
    ComboBoxMonths.Enabled = False
    ComboBoxQuartals.Enabled = False
End Sub

Private Sub OptionMonth_Click()
    ComboBoxMonths.Enabled = True
    ComboBoxQuartals.Enabled = False
End Sub

Private Sub OptionQuartal_Click()
    ComboBoxMonths.Enabled = False
    ComboBoxQuartals.Enabled = True
End Sub

Private Sub CommandExit_Click()
    End
End Sub

Private Sub CommandExport_Click()
    If ComboBoxBuyers.ListIndex = 0 Then
        n = 1
        a = Sellers.Count
        For Each Seller In Sellers
            ExportFile Seller, CStr(n) + " из " + CStr(a) + ": "
            n = n + 1
        Next
    Else
        ExportFile ComboBoxBuyers.Value, ""
    End If
    Message "Готово!"
    End
End Sub

'Экспорт файла
Private Sub ExportFile(ByVal Seller As String, num As String)
    Message "Экспорт файла " + num + Seller
    
    'Определяемся с путём и именем файла
    patch = Cells(2, 3)
    fol = ""
    mnC = OptionMonth.Value
    mn = ComboBoxMonths.Value
    qrC = OptionQuartal.Value
    qr = ComboBoxQuartals.Value
    If mnC Then fol = "\" + mn
    If qrC Then fol = "\" + qr
    If fol <> "" Then folder (patch + fol)
    fileName = patch + fol + "\" + Seller + ".xlsx"
    
    'Создаём книгу
    'On Error GoTo er
    Workbooks.Add
    
    
    
    'Заполняем книгу
    Dim i As Long
    i = firstDat
    Dim j As Long
    j = 3
    Do While DAT.Cells(i, 2) <> "" Or DAT.Cells(i, 15) <> ""
        cp = True
        If DAT.Cells(i, cSeller) <> Seller Then cp = False
        d = DAT.Cells(i, cDates)
        If mnC Then If YearAndMonth(d) <> mn Then cp = False
        If qrC Then If YearAndQuartal(d) <> qr Then cp = False
        If cp Then
            For c = 1 To 14
                Cells(j, c) = DAT.Cells(i, c)
            Next
            j = j + 1
        End If
        i = i + 1
    Loop
    ActiveWorkbook.SaveAs fileName:=fileName
    ActiveWorkbook.Close

er:
End Sub