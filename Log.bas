Attribute VB_Name = "Log"
'Last change: 12.07.2021 15:42

Dim recN As Long    '������� ����� ������

'�������������
Sub Init()
    ERR.Cells.Clear
    ERR.Columns(1).ColumnWidth = 100
    ERR.Columns(2).ColumnWidth = 30
    ERR.Cells(1, 1) = "����"
    ERR.Cells(1, 2) = "���������"
    Range(ERR.Cells(1, 1), ERR.Cells(1, 100)).Interior.Color = colGray
    recN = firstErr
End Sub

'������ ������
Sub Rec(ByVal file As String, ByVal code As Integer)
    msg = "������������ ������"
    If code = 1 Then msg = "������ �������� �����"
    If code = 2 Then msg = "������ � ������"
    If code = 3 Then msg = "����������� ���"
    If code = 4 Then msg = "������ ����� �� ��������������"
    If code = 5 Then msg = "��������! ��������� ���������"
    If code = 6 Then msg = "���� ������������"
    If code = 7 Then msg = "����������� ������, ���� �� �� ������"
    If code = 8 Then msg = "���� �� ���������"
    ERR.Cells(recN, 1).Hyperlinks.Add Anchor:=ERR.Cells(recN, 1), _
        Address:="file:" + file, TextToDisplay:=file
    ERR.Cells(recN, 2) = msg
    recN = recN + 1
End Sub

'******************** End of File ********************