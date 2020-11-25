Attribute VB_Name = "Verify"
Const cComment = 15     '������� ��� �����������
Const startLimits = 3   '������ ������ � �������� �������

Dim Comment As String   '������ � �������������
Dim errors As Boolean   '���� ������� ������
Dim limitO As Variant   '������� ������� �� ��������
Dim summO As Variant    '�������� ���� �� ��������
Dim limitP As Variant   '������� ������� �� �������
Dim summP As Variant    '�������� ���� �� �������
Dim groups As Variant   '������� �����

'������������� �������� �������
Sub Init()

    '������ ������� ������� ��������
    Set dic = Sheets("������ ��������")
    Set limitO = CreateObject("Scripting.Dictionary")
    Set summO = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = startLimits
    Do While dic.Cells(i, 1) <> ""
        cmp = dic.Cells(i, 1).text
        lim = dic.Cells(i, 2)
        limitO(cmp) = lim
        i = i + 1
    Loop
    
    '������ ������� �����
    Set dic = Sheets("������ �������")
    Set groups = CreateObject("Scripting.Dictionary")
    i = startLimits
    Do While dic.Cells(i, 1) <> ""
        cmp = dic.Cells(i, 1).text
        grp = dic.Cells(i, 2).text
        groups(cmp) = grp
        i = i + 1
    Loop
    
    '������ ������� ������� ������
    Set limitP = CreateObject("Scripting.Dictionary")
    Set summP = CreateObject("Scripting.Dictionary")
    i = startLimits
    Do While dic.Cells(i, 2) <> ""
        grp = dic.Cells(i, 2).text
        If dic.Cells(i, 3).text <> "" And IsNumeric(dic.Cells(i, 3)) Then
            lim = dic.Cells(i, 3)
            limitP(grp) = lim
        End If
        i = i + 1
    Loop
    
End Sub

'�������� ������������ ������, ���������� true ���� ���� ������
'dat - ������� � �������
'src - ������� � �����������
'iC - ������ � ������
'iI - ������ � ����������
'changed - true ���� ������ ��� ���� ���������������� � ������ ����������� �� ���������
Function Verify(ByRef dat As Variant, ByRef src As Variant, ByVal iC As Long, ByVal iI As Long, _
    changed As Boolean) As Boolean
    
    red = RGB(255, 192, 192)
    grn = RGB(192, 255, 192)
    yel = RGB(255, 255, 192)
    
    Comment = ""
    errors = False
    Verify = True
    
    '2 - ����
    dat.Cells(iC, 2).NumberFormat = "dd.MM.yyyy"
    If Not IsDate(dat.Cells(iC, 2)) Then
        dat.Cells(iC, 2).Interior.Color = red
        src.Cells(iI, 2).Interior.Color = red
        AddCom "���� ������� �� ���������"
    End If
    
    '3 - ���
    If Not isINNKPP(dat.Cells(iC, 3).text) Then
        dat.Cells(iC, 3).Interior.Color = red
        src.Cells(iI, 3).Interior.Color = red
        AddCom "���/��� ������� �� ���������"
    End If
    
    '5 - ���
    If Not isINNKPP(dat.Cells(iC, 5).text) Then
        dat.Cells(iC, 5).Interior.Color = red
        src.Cells(iI, 5).Interior.Color = red
        AddCom "��� ����� �� ���������"
    End If
    
    '7 - ���������
    dat.Cells(iC, 7).NumberFormat = "### ### ##0.00"
    If Not isPrice(dat.Cells(iC, 7)) Then
        dat.Cells(iC, 7).Interior.Color = red
        src.Cells(iI, 7).Interior.Color = red
        AddCom "��������� ������� �� ���������"
    End If
    
    '8 - ������ ���
    If Not isNDS(dat.Cells(iC, 8).text) Then
        dat.Cells(iC, 8).Interior.Color = red
        src.Cells(iI, 8).Interior.Color = red
        AddCom "��� ����� �� ���������"
    End If
    
    '9-11 - ��������� ������ ���������� �������
    For i = 9 To 11
        dat.Cells(iC, i).NumberFormat = "### ### ##0.00"
        If Not isPriceNDS(dat.Cells(iC, i)) Then
            dat.Cells(iC, i).Interior.Color = red
            src.Cells(iI, i).Interior.Color = red
            AddCom "��������� ������ ���������� ������� ������� �� ���������"
        End If
    Next
    
    '12-14 - ����� ���
    e = False
    For i = 12 To 14
        dat.Cells(iC, i).NumberFormat = "### ### ##0.00"
        If Not isPriceNDS(dat.Cells(iC, i)) Then e = True
    Next
    If e Then
        dat.Cells(iC, i).Interior.Color = red
        src.Cells(iI, i).Interior.Color = red
        AddCom "����� ��� ������� �� ���������"
    Else
        Call LimitsTest(dat, iC)
    End If
    
    '����� ����������� � ������������� ���
    col = red
    If Not errors Then col = grn: Comment = "�������"
    dat.Cells(iC, cComment) = Comment
    dat.Cells(iC, cComment).Interior.Color = col
    src.Cells(iI, cComment) = Comment
    src.Cells(iI, cComment).Interior.Color = col
    
    Verify = errors
    
End Function

'�������� �������
Sub LimitsTest(ByRef dat As Variant, ByVal i As Long)
    sel = dat.Cells(i, 6)
    grp = groups(sel)
    buy = dat.Cells(i, 4) + "!" + grp
    Sum = 0
    For j = 12 To 14
        If IsNumeric(dat.Cells(i, j)) Then Sum = Sum + dat.Cells(i, j)
    Next
    summO(sel) = summO(sel) + Sum
    summP(buy) = summP(buy) + Sum
    If summO(sel) > limitO(sel) Then AddCom "�������� ����� ��������"
    If summP(buy) > limitP(grp) Then AddCom "�������� ����� �������"
End Sub

'���������� ����������� � ������
Sub AddCom(str As String)
    If Comment <> "" Then Comment = Comment + ", "
    Comment = Comment + str
    errors = True
End Sub

'�������� �� ������������ ����
Function isPrice(ByVal var As Variant)
    isPrice = False
    If IsNumeric(var) Then
        If var >= 0 And var <> "" Then isPrice = True
    End If
End Function

'�������� �� ������������ ���� ��� ��� � ��� �����
Function isPriceNDS(ByVal var As Variant)
    isPriceNDS = False
    If IsNumeric(var) Then
        If var >= 0 Then isPriceNDS = True
    Else
        If Not IsError(var) Then
            If var = "" Then isPriceNDS = True
        End If
    End If
End Function

'�������� �� ������������ ���/���
Function isINNKPP(ByVal str As String) As Boolean
    If str = "" Then isINNKPP = False: Exit Function
    Dim s() As String
    s = Split(str, "/")
    isINNKPP = True
    If Not IsNumeric(s(0)) Then isINNKPP = False
    If Len(s(0)) <> 10 And Len(s(0)) <> 12 Then isINNKPP = False
    If UBound(s) > 0 Then
        If Not IsNumeric(s(1)) Then isINNKPP = False
        If Len(s(1)) <> 9 Then isINNKPP = False
    End If
End Function

'�������� �� ������������ ���
Function isNDS(ByVal str As String) As Boolean
    isNDS = False
    If str = "10" Then isNDS = True
    If str = "18" Then isNDS = True
    If str = "20" Then isNDS = True
End Function