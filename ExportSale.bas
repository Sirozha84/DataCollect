Attribute VB_Name = "ExportSale"
'��������� ������: 03.04.2021 21:33

'������� �����
Public Sub Run(ByVal inn As String, ByVal NUM As String, _
        ByVal FirstDate As Date, ByVal LastDate As Date)

    saler = SellFileName(inn)
    Message "������� ����� " + NUM + saler
    
    '�������� ������������ ������ � �����������
    Si = selIndexes(inn)
    Limit = DIC.Cells(Si, cLimND)
    ermsg = "� �������� " + DIC.Cells(Si, 1) + " � ��� " + inn + " "
    If Limit = Empty Then
        MsgBox ermsg + "�� ������ �����!"
        End
    End If
    If StupidQToQIndex(DIC.Cells(Si, cOPND)) < 0 Then
        MsgBox ermsg + "�� ������ ��� ������ �� ��������� �������� ������ ��!"
        End
    End If
    
    '������������ � ���� � ������ �����
    Patch = DirExport + "\��������"
    MakeDir Patch
    fileName = Patch + "\" + cutBadSymbols(saler) + ".xlsx"
    
    '������ �����
    Workbooks.Add
    i = 1
    Cells(i, 1) = "��� ����" + Chr(10) + "��������"
    Cells(i, 2) = "� ����" + Chr(10) + "�������"
    Cells(i, 3) = "���� ����" + Chr(10) + "�������"
    Cells(i, 4) = "���"
    Cells(i, 5) = "���"
    Cells(i, 6) = "������������"
    Cells(i, 7) = "����� � ���." + Chr(10) + "� ���."
    Cells(i, 8) = "�����" + Chr(10) + "��� ��� 20%"
    Cells(i, 9) = "�����" + Chr(10) + "��� ��� 18%"
    Cells(i, 10) = "�����" + Chr(10) + "��� ��� 10%"
    Cells(i, 11) = "��� 20%"
    Cells(i, 12) = "��� 18%"
    Cells(i, 13) = "��� 10%"
    Cells(i, 14) = "������ ��"
    Cells(i, 15) = "�������" '��������� �������
    Cells(i, 16) = "���"
    Cells(i, 17) = "������"
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
    
    '��������� �����
    i = firstDat
    j = firstEx
    Do While DAT.Cells(i, cAccept) <> ""
        If DAT.Cells(i, cAccept) = "OK" Then
            dc = DAT.Cells(i, cDateCol)
            If dc >= FirstDate And dc < LastDate + 1 Then
                cp = True
                If DAT.Cells(i, cSellINN).text <> inn Then cp = False
                d = DAT.Cells(i, cDates)
                If cp Then
                    '����������� ������ �� �����
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
                    '��������� ������� - ������ �������� � ����� ���
                    Cells(j, 15) = DateToQIndex(DAT.Cells(i, 2))
                    Sum = 0
                    For j2 = 11 To 13
                        If IsNumeric(Cells(j, j2)) Then Sum = Sum + Cells(j, j2)
                    Next
                    Cells(j, 16) = Sum
                    Cells(j, 17) = i
                    j = j + 1
                End If
            End If
        End If
        i = i + 1
    Loop
    
    '���������� �� �������� � ���������
    Columns("A:Q").Sort key1:=Range("O2"), order1:=xlDescending, _
                        key2:=Range("F2"), order2:=xlAscending
    
    '��������� ������� �� � �������� �� �� ���� �����
    SetProtect DAT
    PeriodND Si
    i = firstEx
    Do While Cells(i, 1) <> ""
        DAT.Cells(Cells(i, 17), cPND) = Cells(i, 14)
        i = i + 1
    Loop

    '�������� ��������� ��������
    Columns(15).Delete
    Columns(15).Delete
    Columns(15).Delete
    
    '���������� � �������� ���������
    On Error GoTo er
    Application.DisplayAlerts = False
    If j > firstEx Then ActiveWorkbook.SaveAs fileName:=fileName
    ActiveWorkbook.Close
    Exit Sub

er:
    ActiveWorkbook.Close
    MsgBox "��������� ������ ��� ���������� ����� " + fileName

End Sub

'������ �������� ��������� ����������
'Si - ������ ��������
Sub PeriodND(ByVal Si As Double)
    
    Dim oND As Integer
    oND = StupidQToQIndex(DIC.Cells(Si, cOPND))
    
    '******************** ������ ���� ********************
    
    '���������� ������ ����������� �������� �� ������� ��� ��� ������� oND
    Set ni = CreateObject("Scripting.Dictionary")   '������� �� ���
    Set ns = CreateObject("Scripting.Dictionary")   '����� �� ���
    i = 2
    Do While Cells(i, 1) <> ""
        inn = Cells(i, 4)
        If Cells(i, 15) = oND Then
            s = Cells(i, 16)
            If ns(inn) = 0 Or ns(inn) > s Then
                ns(inn) = s
                ni(inn) = i
            End If
        End If
        i = i + 1
    Loop
    
    '��������� ����� ���� ��������� ������� �� ����� ���
    Do
        Sum = 0
        For Each i In ni
            Sum = Sum + ns(i)
        Next
        per = Sum - limitOND
        '�.. ���� ����� ��������� �����...
        If per > 0 Then
            '������� ������, ������� ����� ����� � ����� ���������� (per)
            Min = 0         '����������� �������
            isk = ""        '������, ������� ���� ���������
            plus = False    '���� �� �������������� �������?
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
            '��������� ��������� ������ � �������
            ns.Remove (isk)
            ni.Remove (isk)
        End If
    Loop Until per <= 0
    
    '��������� ������ �� ���������� �������
    pnd = IndexToQYYYY(oND)
    For Each i In ni
        Cells(ni(i), 14) = pnd
    Next
    
    '******************** ������ ���� ********************
    
    Dim tND As Integer  '������� ������
    Dim Qi As Object    '������� ������� (����� �� �������)
    '������ ������, ������� "�� ������" � �������� �������
    Set Qi = CreateObject("Scripting.Dictionary")
    i = 2
    Do While Cells(i, 1) <> ""
        If Cells(i, 14) = "" And Cells(i, 15) = oND Then Qi(i) = Cells(i, 16)
        i = i + 1
    Loop
    tND = oND '�������� ������ ������ ������� ������, � ���� ������� ���� ������� �����
    
    Do
        
        '��������� � ���������� �������
        tND = tND + 1
        Dim Limit As Double
        Limit = DIC.Cells(Si, cLimND) - DIC.Cells(Si, cCorrect + tND)
        
        '���������� ������ ������� �������� �������
        Set ti = CreateObject("Scripting.Dictionary")  '������ ������� �������� ������� (����� �� �������)
        i = 2
        Do While Cells(i, 1) <> ""
            If Cells(i, 15) = tND Then ti(i) = Cells(i, 16)
            i = i + 1
        Loop
        
        '��������� ����� ���� ������� ������� �� ����� ��
        Do
            s = 0
            For Each i In ti
                s = s + ti(i)
            Next
            per = s - Limit
            If per > 0 Then
                If Limit < minLim Then
                    '� ���� ������ ������ ����������
                    '��������� ��� ������ � �������
                    For Each i In ti
                        Qi.Add i, ti(i)
                        ti.Remove (i)
                    Next
                Else
                    '��������� ������ � ������������ ���������
                    maxs = 0    '������������ ��������
                    msxi = 0    '������ ������ � ������������ ���������
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
        
        '��������� ������ �� ���������� �������
        ost = -per
        pnd = IndexToQYYYY(tND)
        For Each i In ti
            Cells(i, 14) = pnd
        Next
        
        '���� ������� �� ����� � �������� "�����",
        '����������� ������ �� ��, ������� � ������������ ��������
        If Qi.Count > 0 Then
            Do
                mins = 0    '����������� ��������
                mini = 0    '������ ������������ ��������
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
        
        'Debug.Print "������� �� ������ " + CStr(tND)
        'For Each i In Qi: Debug.Print i: Next
        
    Loop While tND < quartCount - 1

End Sub

'******************** End of File ********************