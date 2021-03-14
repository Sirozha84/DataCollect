VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProperties 
   Caption         =   "Настройки"
   ClientHeight    =   4200
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
    TextBoxImport = PRP.Cells(4, 2).text
    TextBoxExport = PRP.Cells(5, 2).text
    TextBoxImportBuys = PRP.Cells(6, 2).text
End Sub

Private Sub ButtonExploreImportSells_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxImport = diag.SelectedItems(1)
End Sub

Private Sub ButtonExport_Click()
    Set diag = Application.FileDialog(msoFileDialogFolderPicker)
    If diag.Show = 0 Then Exit Sub
    TextBoxExport = diag.SelectedItems(1)
End Sub

Private Sub CommandOK_Click()
    PRP.Cells(4, 2) = TextBoxImport
    PRP.Cells(5, 2) = TextBoxExport
    PRP.Cells(6, 2) = TextBoxImportBuys
    End
End Sub

Private Sub CommandCancel_Click()
    End
End Sub