Attribute VB_Name = "Log"
Dim err As Variant  '������� � ��������
Dim recN As Long    '������� ����� ������

Sub Init()
    '������ ������� (���� � ���) ��� ������ ������
    Call NewTab(tabErr, True)
    Set err = Sheets(tabErr)
    err.Columns(1).ColumnWidth = 100
    err.Columns(2).ColumnWidth = 20
    err.Cells(1, 1) = "����"
    err.Cells(1, 2) = "���������"
    Range(err.Cells(1, 1), err.Cells(1, 100)).Interior.Color = RGB(214, 214, 214)
    recN = 2
End Sub

Sub Rec(ByVal file As String, ByVal code As Integer)
    msg = "������������ ������"
    If code = 1 Then msg = "������ �������� �����"
    If code = 2 Then msg = "������ � ������"
    If code = 3 Then msg = "����������� ���"
    If code = 4 Then msg = "��������! ��������� ���������"
    err.Cells(recN, 1) = file
    err.Cells(recN, 2) = msg
    recN = recN + 1
End Sub