Attribute VB_Name = "Verify"
Const cComment = 16
Dim Comment As String
Dim errors As Boolean

'Проверка корректности данных
Function Verify(ByRef cur As Variant, ByRef imSh As Variant, ByVal iC As Long, ByVal iI As Long) As Boolean
    Comment = ""
    errors = False
    red = RGB(255, 192, 192)
    Verify = True
    '2 - Дата
    cur.Cells(iC, 2).NumberFormat = "dd.MM.yyyy"
    If Not IsDate(cur.Cells(iC, 2)) Then
        cur.Cells(iC, 2).Interior.Color = red
        imSh.Cells(iI, 2).Interior.Color = red
        AddCom "Дата введена не корректно"
    End If
    
    '3 - ИНН
    If Not isINNKPP(cur.Cells(iC, 3)) Then
        cur.Cells(iC, 3).Interior.Color = red
        imSh.Cells(iI, 3).Interior.Color = red
        AddCom "ИНН/КПП введены не корректно"
    End If
    
    If errors Then
        cur.Cells(iC, cComment) = Comment
        cur.Cells(iC, cComment).Interior.Color = red
        imSh.Cells(iI, cComment) = Comment
        imSh.Cells(iI, cComment).Interior.Color = red
    End If
    Verify = errors
    
End Function

'Добавление комментария к строке
Sub AddCom(str As String)
    If Comment <> "" Then Comment = Comment + ", "
    Comment = Comment + str
    errors = True
End Sub

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