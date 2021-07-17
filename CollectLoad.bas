Attribute VB_Name = "CollectLoad"
'Last change: 17.07.2021 21:56

Dim LastRec As Long
Dim curFile As String
Dim curMark As String
Dim curProv As String
Dim curProvINN As String
Dim UINs As Object

'������� ������ � �������������� ������
Dim scRType As Byte     '����� �������: 0 - �����/����, ���/��� - � ����� ������, 1 - � ����������
Dim scFirst As Integer  '������ ������ � �������
Dim scKVO As Integer        '��� ���� ��������
Dim scND As Integer         '�����/����
Dim scSeller As Integer     '������������ ��������
Dim scSellerINN As Integer  '��� ��������
Dim scPrice As Integer      '���������
Dim scPWN20 As Integer      '��������� ���������� ������� 20%
Dim scPWN18 As Integer      '��������� ���������� ������� 18%
Dim scPWN10 As Integer      '��������� ���������� ������� 10%
Dim scNDS20 As Integer      '��� 20%
Dim scNDS18 As Integer      '��� 18%
Dim scNDS10 As Integer      '��� 20%
Dim scNDS As Integer        '��� �����

'������ �������� ����� ������
Sub Run()
    
    Message "����������..."
    Dictionary.Init
    Numerator.InitLoad
    Log.Init
    
    '������� ���� �� ������ ���������� �������
    Set UINs = CreateObject("Scripting.Dictionary")
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        If DTL.Cells(i, clAccept) = "OK" Then
            UINs(DTL.Cells(i, clUIN).text) = i
        Else
            DTL.Rows(i).Delete
            i = i - 1
        End If
        i = i + 1
    Loop
    LastRec = i
    
    '�������� ��������� ������ � ������ ���� �� ���
    Set files = Source.getFiles(DirImportLoad, False)
    n = 1
    s = 0
    e = 0
    For Each file In files
        curf = file
        If Len(curf) > 40 Then curf = "..." + Right(curf, 40)
        Message ("��������� ����� " + CStr(n) + " �� " + CStr(files.Count) + " (" + curf) + ")"
        er = AddFile(file)
        If er > 0 Then
            Log.Rec file, er
            e = e + 1
        Else
            s = s + 1
        End If
        n = n + 1
    Next

    '�������� �� ���������
    Message "�������� �� ���������..."
    ed = FindDuplicates

    '��������� ������ � �����������
    Message "������ ����������� �������..."
    Range(DIC.Cells(firstDic, cPBalance), DIC.Cells(maxRow, cPBalance + quartCount * 2 - 1)).Clear
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        If DTL.Cells(i, clAccept) = "OK" Then
            Qi = DateToQIndex(DTL.Cells(i, clDate))
            If Qi >= 0 Then
                Si = IndexByINN(DTL.Cells(i, clSaleINN).text)
                Qi = Qi * 2 + cPBalance
                Sum = DTL.Cells(i, clNDS)
                If DTL.Cells(i, 1).text = "�" Then Qi = Qi + 1
                DIC.Cells(Si, Qi) = DIC.Cells(Si, Qi) + Sum
            End If
        End If
        i = i + 1
    Loop

    '����������
    ActiveWorkbook.Save
    Message "������! ���� �������."
    Application.DisplayAlerts = True
    
    MsgBox ("��������� ���������!" + Chr(13) + _
            "������ ����������� �������: " + CStr(s) + Chr(13) + _
            "����� � ��������: " + CStr(e)) + Chr(13) + _
            "������� ������� ��: " + CStr(ed)
    
End Sub

'���������� ������ �� �����. ����������:
'0 - �� ������
'1 - ������ ��������
'2 - ������ � ������ (errors=true)
'7 - ��� �������, ��� �� �� ������
'8 - ���� �� ����������
Function AddFile(ByVal file As String) As Byte
    
    '����������
    Application.DisplayAlerts = False
    If Not TrySave(file) Then AddFile = 6: Exit Function
    errors = False
    Application.ScreenUpdating = False
    On Error GoTo er
    Set impBook = Nothing
    Set impBook = Workbooks.Open(file, False, False)
    Set SRC = impBook.Worksheets(1)
    On Error GoTo 0
        
    '������ � �������� �������
    curMark = UCase(SRC.Cells(1, 1).text)
    If curMark <> "�" And curMark <> "�" Then
        AddFile = 7
        impBook.Close False
        Exit Function
    End If
    
    '������ ������
    c = 60  '������� � �������� ������
    If SBFieldRecognition Then
        i = scFirst
        Do While SRC.Cells(i, scKVO).text <> ""
            If UINs(SRC.Cells(i, c).text) = "" And Len(SRC.Cells(i, scKVO).text) < 3 Then
                If Not copyRecord(i) Then
                    errors = True
                    DTL.Cells(LastRec, clAccept) = "fail"
                Else
                    DTL.Cells(LastRec, clDateCol) = DateTime.Now
                    uin = GenerateLoad
                    DTL.Cells(LastRec, clUIN) = uin
                    SRC.Cells(i, c) = uin
                    DTL.Cells(LastRec, clAccept) = "OK"
                End If
                DTL.Cells(LastRec, clFile) = file
                LastRec = LastRec + 1
            End If
            i = i + 1
        Loop
    Else
         AddFile = 8
    End If
    If errors Then AddFile = 2
    
    '����������
    On Error GoTo er
    impBook.Close True
    Application.ScreenUpdating = True
    DoEvents    '�� ����� ��� ���� ���, ����� ��� ��� ����� �� ��������, � ����� ����������� ����� ����
    Exit Function

er:
    AddFile = 1
    
End Function

'����������� ������ �� ����� ������. ���������� True, ���� ������ ��������� ��� ������
'Si - ������ � ����������
Function copyRecord(ByVal Si As Long) As Boolean
    
    SetFormates LastRec
    On Error GoTo er
    DTL.Cells(LastRec, clMark) = curMark
    kvo = SRC.Cells(Si, scKVO)
    DTL.Cells(LastRec, clKVO) = kvo
    DTL.Cells(LastRec, clSaleINN) = Left(SRC.Cells(Si, scSellerINN), 10)
    DTL.Cells(LastRec, clSaleName) = SRC.Cells(Si, scSeller)
    If kvo = "02" Then
        DTL.Cells(LastRec, clKVO) = "22"
        DTL.Cells(LastRec, clSaleINN) = curProvINN
        DTL.Cells(LastRec, clSaleName) = curProv
        kvochange = True
    End If
    If scRType = 0 Then
        nd = SRC.Cells(Si, scND).text
        DTL.Cells(LastRec, clNum) = NumFromND(nd)
        DTL.Cells(LastRec, clDate) = Right(nd, 10)
    Else
        DTL.Cells(LastRec, clNum) = SRC.Cells(Si, scND).text
        DTL.Cells(LastRec, clDate) = SRC.Cells(Si, scND + 1).text
    End If
    DTL.Cells(LastRec, clProvINN) = curProvINN
    DTL.Cells(LastRec, clProvName) = curProv
    DTL.Cells(LastRec, clPrice) = SRC.Cells(Si, scPrice)
    If scPWN20 <> 0 Then DTL.Cells(LastRec, clPrice + 1) = SRC.Cells(Si, scPWN20)
    If scPWN18 <> 0 Then DTL.Cells(LastRec, clPrice + 2) = SRC.Cells(Si, scPWN18)
    If scPWN10 <> 0 Then DTL.Cells(LastRec, clPrice + 3) = SRC.Cells(Si, scPWN10)
    If scNDS20 <> 0 Then DTL.Cells(LastRec, clNDS) = _
            WorksheetFunction.Sum(Range(SRC.Cells(Si, scNDS20), SRC.Cells(Si, scNDS10)))
    If scNDS20 = 0 And scNDS18 <> 0 Then DTL.Cells(LastRec, clNDS) = _
            WorksheetFunction.Sum(Range(SRC.Cells(Si, scNDS18), SRC.Cells(Si, scNDS10)))
    If scNDS <> 0 Then DTL.Cells(LastRec, clNDS) = SRC.Cells(Si, scNDS)
    
    copyRecord = VerifyLoad(LastRec)
    
    '��� ������� � 02 �� 22, ������ ��������� � ���� �������� ��������
    If kvochange Then
        i = IndexByINN(DTL.Cells(LastRec, clSaleINN).text)
        j = DateToQIndex(DTL.Cells(LastRec, clDate))
        If j >= 0 Then DIC.Cells(i, cSaleProtect + j) = "��"
    End If
    
    AddFormuls
    Exit Function
    
er:
    copyRecord = False
    
End Function

'��������� �������� �����
Sub SetFormates(ByVal i As Integer)
    DTL.Cells(LastRec, clKVO).NumberFormat = "@"
    DTL.Cells(LastRec, clNum).NumberFormat = "@"
    DTL.Cells(LastRec, clDate).NumberFormat = "dd.MM.yyyy"
    DTL.Cells(LastRec, clProvINN).NumberFormat = "@"
    DTL.Cells(LastRec, clSaleINN).NumberFormat = "@"
    For i = 0 To 4
        DTL.Cells(LastRec, clPrice + i).NumberFormat = numFormat
    Next
End Sub

'���������� ������ � �������� ������
Sub AddFormuls()
    s = CStr(LastRec)
    DTL.Cells(LastRec, clOst).Formula = "=M" + s + "-OneCellSum(P" + s + ")"
    'formul = "=R" + s + ">=0"
    'With DTL.Cells(LastRec, clRasp).Validation
    '    .Delete
    '    .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:=formul
    '    .ErrorMessage = "������������� ����� ��������� ����� ���"
    'End With
End Sub

'�������� ��������� ������� �� ���������� ������������� �������
Function FindDuplicates()
    e = 0
    Set numbers = CreateObject("Scripting.Dictionary")
    i = firstDtL
    Do While DTL.Cells(i, clAccept) <> ""
        NUM = DTL.Cells(i, clNum).text + "!" + _
              DTL.Cells(i, clProvINN).text + "!" + _
              DTL.Cells(i, clSaleINN).text
        If numbers(NUM) = Empty Then
            numbers(NUM) = i
        Else
            DTL.Cells(i, clCom) = "����� �� �����������"
            DTL.Cells(i, clCom).Interior.Color = colRed
            DTL.Cells(i, clAccept) = "fail"
            e = e + 1
        End If
        i = i + 1
    Loop
    FindDuplicates = e
End Function

'��������� ������ �� "����� � ����"
Function NumFromND(ByVal nd As String) As String
    ss = Split(nd, " ��")
    If UBound(ss) > 0 Then NumFromND = ss(0)
    ss = Split(nd, ";")
    If UBound(ss) > 0 Then NumFromND = ss(0)
End Function

'������������� ������� � ����� ������
Function SBFieldRecognition() As Boolean
    
    '����� ������, ��� 0
    scReset
    st = 7
    Do
        For i = 1 To 50
            If SRC.Cells(st, i) = "��� ���� �����-���" Then scKVO = i
            If Left(SRC.Cells(st, i), 12) = "����� � ����" Then If scND = 0 Then scND = i
            If Left(SRC.Cells(st, i), 12) = "������������" Then scSeller = i
            If Left(SRC.Cells(st, i), 7) = "���/���" Then scSellerINN = i
            If Left(SRC.Cells(st + 3, i), 10) = "� ������ �" Then scPrice = i
            If SRC.Cells(st + 3, i) = "20 ���������" Then If scPWN20 = 0 Then scPWN20 = i Else scNDS20 = i
            If SRC.Cells(st + 3, i) = "18 ���������" Then If scPWN18 = 0 Then scPWN18 = i Else scNDS18 = i
            If SRC.Cells(st + 3, i) = "10 ���������" Then If scPWN10 = 0 Then scPWN10 = i Else scNDS10 = i
        Next
        SBFieldRecognition = scKVO <> 0 And scND <> 0 And scSeller <> 0 And scSellerINN <> 0 And _
                scPrice <> 0 And scPWN18 <> 0 And scPWN10 <> 0 And scNDS18 <> 0 And scNDS10 <> 0
        st = st + 1
    Loop Until SBFieldRecognition Or st > 8
    If SBFieldRecognition Then
        scRType = 0
        scFirst = 13
        For i = 2 To 6
            If Left(SRC.Cells(i, 1).text, 8) = "��������" Then _
                    curProv = Replace(SRC.Cells(i, 1).text, "��������  ", "")
            If Left(SRC.Cells(i, 1).text, 17) = "�����������������" Then _
                    curProvINN = Right(SRC.Cells(i, 1).text, 20)
        Next
        Exit Function
    End If
        
    '����� ������, ��� 1
    scReset
    st = 2
    Do
        For i = 1 To 50
            If Left(SRC.Cells(4, i), 17) = "��� ���� ��������" Then scKVO = i
            If SRC.Cells(st, i) = "��" Then scND = i
            If Right(SRC.Cells(st, i), 10) = "(���. 020)" Then scND = i
            If SRC.Cells(st, i) = "�������� � ����������" Then scSeller = i + 2: scSellerINN = i
            If Right(SRC.Cells(st, i), 10) = "(���. 100)" Then scSeller = i + 2: scSellerINN = i
            If Left(SRC.Cells(4, i), 13) = "� ���. � ���." Then scPrice = i
            If Left(SRC.Cells(4, i), 9) = "���������" And scPrice = 0 Then scPrice = i
            If Right(SRC.Cells(st, i), 10) = "(���. 106)" And scPrice = 0 Then scPrice = i
            If Left(SRC.Cells(4, i), 9) = "����� ���" Then scNDS = i
            If Left(SRC.Cells(4, i), 3) = "20%" Then If scPWN20 = 0 Then scPWN20 = i Else scNDS20 = i
            If Left(SRC.Cells(4, i), 3) = "18%" Then If scPWN18 = 0 Then scPWN18 = i Else scNDS18 = i
            If Left(SRC.Cells(4, i), 3) = "10%" Then If scPWN10 = 0 Then scPWN10 = i Else scNDS10 = i
        Next
        SBFieldRecognition = scKVO <> 0 And scND <> 0 And scSeller <> 0 And scSellerINN <> 0 And _
                scPrice <> 0 And (scNDS <> 0 Or (scPWN18 <> 0 And scPWN10 <> 0 And scNDS18 <> 0 And scNDS10 <> 0))
        st = st + 1
    Loop Until SBFieldRecognition Or st > 4
    If SBFieldRecognition Then
        scRType = 1
        scFirst = 5
        curProv = SRC.Cells(2, 1).text
        curProvINN = SRC.Cells(3, 1).text
        Exit Function
    End If
    
    '������ �����
    scReset
    For i = 1 To 50
        If Left(SRC.Cells(9, i), 3) = "���" Then scKVO = i
        If Left(SRC.Cells(9, i), 12) = "����� � ����" And scND = 0 Then scND = i
        If SRC.Cells(9, i) = "��" And scND = 0 Then scND = i
        If Left(SRC.Cells(9, i), 12) = "������������" And acseller = 0 Then scSeller = i
        If Left(SRC.Cells(9, i), 7) = "���/���" Then scSellerINN = i
        If Left(SRC.Cells(9, i), 9) = "���������" Then scPrice = i
        If Left(SRC.Cells(9, i), 11) = "� ��� �����" Then scNDS = i
    Next
    SBFieldRecognition = scKVO <> 0 And scND <> 0 And scSeller <> 0 And _
            scSellerINN <> 0 And scPrice <> 0 And scNDS <> 0
    If SBFieldRecognition Then
        scRType = 0
        scFirst = 13
        curProv = Split(SRC.Cells(4, 2).text, ": ")(1)
        curProvINN = Right(SRC.Cells(5, 2).text, 20)
        Exit Function
    End If
        
End Function

'����� ����������� �������
Sub scReset()
    scKVO = 0
    scND = 0
    scSeller = 0
    scSellerINN = 0
    scPrice = 0
    scPWN20 = 0
    scPWN18 = 0
    scPWN10 = 0
    scNDS20 = 0
    scNDS18 = 0
    scNDS10 = 0
    scNDS = 0
End Sub

'******************** End of File ********************