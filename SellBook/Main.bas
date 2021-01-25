Attribute VB_Name = "Main"
Public Const PatVersion = "20210108"

Public Patch As String
Public BookCount As Long

Public Sub ButtonGenerate()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    Patch = diag.SelectedItems(1)
    
    Set files = getFiles(Patch)
    Range(Cells(7, 1), Cells(1000000, 2)).Clear
    i = 7
    For Each file In files
        Cells(i, 1) = file
        er = ExportBook(file)
        If er = 0 Then Cells(i, 2) = "������ ��� ������ � ������"
        If er = 1 Then
            If BookCount > 0 Then
                Cells(i, 2) = "������� ����� ������ (" + CStr(BookCount) + ")"
            Else
                Cells(i, 2) = "������ ������"
            End If
        End If
        If er = 2 Then Cells(i, 2) = "������ ����� ������������ ������"
        i = i + 1
    Next
    Message "������!"
    MsgBox "������������ ���� ������ ���������!"
End Sub