Attribute VB_Name = "ShortMacros"
Option Explicit


Sub ChangeCase()
' ChangeCase Macro
' vbUpperCase, vbLowerCase; vbProperCase
    Dim c As Object
    
    If Selection.Cells.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    
    If Selection(1).Value = StrConv(Selection(1), vbUpperCase) Then
        For Each c In Selection.Cells
        c.Value = StrConv(c, vbLowerCase)
        Next
    ElseIf Selection(1).Value = StrConv(Selection(1), vbLowerCase) Then
        For Each c In Selection.Cells
        c.Value = StrConv(c, vbProperCase)
        Next
    Else
        For Each c In Selection.Cells
        c.Value = StrConv(c, vbUpperCase)
        Next
    End If
End Sub



Sub TrimText()
' TrimText Macro
    Dim TrimCount As Integer
    Dim c As Object
    
    TrimCount = 0
    
    If Selection.Cells.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    
    For Each c In Selection.Cells
        If WorksheetFunction.IsText(c) And c.Value <> Trim(c) Then
            c.Value = Trim(c)
            TrimCount = TrimCount + 1
        End If
    Next
    MsgBox TrimCount
End Sub



Sub CopySheet_NewWB()
' Copy ActiveSheet to a new Workbook
    ActiveSheet.Copy
End Sub



Sub CenterAcrossSelection()
' CenterAcrossSelection macro
    Selection.HorizontalAlignment = xlCenterAcrossSelection
End Sub

