Attribute VB_Name = "Numerator"
'Last change: 12.07.2021 21:27

Dim Prefixes As Object  '������� ���������
Dim Liters As Object    '������� �������
Dim Codes As Object     '������� �����
Dim LNum As Long        '��������� ����� �����������

'������������� �������
Sub Init()
    
    '�������� ������� ����������
    Range(NUM.Cells(1, 1), NUM.Cells(3, 100)).Interior.Color = RGB(214, 214, 214)
    Set Prefixes = CreateObject("Scripting.Dictionary")
    i = firstNum
    Do While NUM.Cells(i, 1).text <> ""
        pref = NUM.Cells(i, 1).text
        Prefixes.Add pref, NUM.Cells(i, 2)
        i = i + 1
    Loop
    
    '�������� ��������� � �����
    Set Liters = CreateObject("Scripting.Dictionary")
    Set Codes = CreateObject("Scripting.Dictionary")
    i = firstDic
    Do While DIC.Cells(i, cINN) <> ""
        seller = DIC.Cells(i, cINN).text
        Liters.Add seller, GetLiter(seller, DIC.Cells(i, cPLiter).text)
        Codes.Add seller, DIC.Cells(i, cPCode).text
        i = i + 1
    Loop

End Sub

'������������� ���������� ������ ����������� (���������� ��������� � ������������)
Sub InitLoad()
    LNum = PRP.Cells(10, 2)
End Sub

'���������� ������� �� ��������
Sub Save()
    i = firstNum
    For Each Key In Prefixes.keys
        NUM.Cells(i, 1).NumberFormat = "@"
        NUM.Cells(i, 1) = Key
        NUM.Cells(i, 2) = Prefixes(Key)
        i = i + 1
    Next
    '�� ������ ������ ������ ��������� ������ ������
    NUM.Cells(i, 1) = ""
    NUM.Cells(i, 2) = ""
End Sub

'��������� ����������� ������ ��������
Public Function Generate(recDate As Date, seller As String) As String
    
    lit = Liters(seller)
    Dim d As Byte
    cod = GetCode(seller, recDate, d)
    
    pref = lit + cod
    If Not Prefixes.exists(pref) Then Prefixes.Add pref, 0
    Prefixes(pref) = Prefixes(pref) + 1

    Generate = pref + Right(CStr(Prefixes(pref) + 10 ^ d), d)

End Function

'��������� ����������� ������ �����������
Public Function GenerateLoad() As Long
    LNum = LNum + 1
    PRP.Cells(10, 2) = LNum
    GenerateLoad = LNum
End Function

'����� � ������� ��� ��������� ������
Function GetLiter(ByVal seller As String, lit As String) As String
    If lit <> "" Then
        GetLiter = UCase(lit)
    Else
        GetLiter = UCase(Left(seller, 1))
    End If
End Function

'����� � ������� ��� ��������� ����
'd - ���������� ���� ������
Function GetCode(ByVal seller As String, dateR As Date, ByRef d As Byte)
    GetCode = Codes(seller)
    d = 5
    If GetCode = "" Then
        GetCode = Right(CStr(Year(dateR)), 2) + _
            Right(CStr(Month(dateR) + 100), 2) + _
            Right(CStr(Day(dateR) + 100), 2)
        d = 3
    End If
End Function

'�������� ������������ �������� (������������ ������������ ��� � ������������ ��������)
Function CheckPrefix(UID As String, ByVal dateS As Date, seller As String) As Boolean
    Dim d As Byte
    pref = Liters(seller) + GetCode(seller, dateS, d)
    CheckPrefix = pref = Left(UID, Len(pref))
End Function

'******************** End of File ********************