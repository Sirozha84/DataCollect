VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProperties 
   Caption         =   "Настройки"
   ClientHeight    =   4088
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   7308
   OleObjectBlob   =   "FormProperties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
    TextBoxImportSale = PRP.Cells(pImportSale, 2).text
    TextBoxImportLoad = PRP.Cells(pImportLoad, 2).text
    TextBoxExport = PRP.Cells(pExport, 2).text
End Sub

Private Sub ButtonExploreImportSale_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxImportSale = diag.SelectedItems(1)
End Sub

Private Sub CommandImportLoad_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxImportLoad = diag.SelectedItems(1)
End Sub

Private Sub ButtonExport_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxExport = diag.SelectedItems(1)
End Sub

Private Sub CommandButton3_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxImportBuys = diag.SelectedItems(1)
End Sub

Private Sub CommandOK_Click()
    PRP.Cells(pImportSale, 2) = TextBoxImportSale
    PRP.Cells(pImportLoad, 2) = TextBoxImportLoad
    PRP.Cells(pExport, 2) = TextBoxExport
    End
End Sub

Private Sub CommandCancel_Click()
    End
End Sub