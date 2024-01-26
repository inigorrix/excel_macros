Attribute VB_Name = "SelectionInfo"
Option Explicit


Sub UniqueValues()
' UniqueValues macro
    Dim dict As Object
    Dim cell As Range
    Dim repeatedvalues, blankcells As Integer
    Dim msg As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    repeatedvalues = 0
    blankcells = 0
    
    If Selection.Cells.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    
    For Each cell In Selection.Cells
        If Not dict.exists(cell.Value) Then
            dict.Add cell.Value, 1
        Else
            dict(cell.Value) = dict(cell.Value) + 1
            If dict(cell.Value) = 2 Then
                repeatedvalues = repeatedvalues + 2
            Else
                repeatedvalues = repeatedvalues + 1
            End If
        End If
        If cell.Value = "" Then
            blankcells = blankcells + 1
        End If
    Next
    
    msg = _
        "Different values: " & dict.Count & vbCrLf & _
        "Unique values: " & Selection.Count - repeatedvalues & vbCrLf & vbCrLf & _
        "Selection count: " & Selection.Count & vbCrLf & _
        "Non-Blank cells: " & Selection.Count - blankcells & vbCrLf & _
        "Blank cells: " & blankcells
    MsgBox msg

End Sub



Sub Product()
' Product macro
    Dim msg As String
    Dim cell As Range
    Dim prod As Double
    
    If Selection.Cells.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    
    prod = 1
    For Each cell In Selection.Cells
        If cell.Value <> "" Then
            prod = prod * cell.Value
        End If
    Next
    
    MsgBox "Product:  " & Format(prod, "#,##0.0000")
End Sub



Sub Difference_and_Ratios()
' Difference_and_Ratios macro
    Dim msg As String
    Dim cell As Range
    Dim diff, ratio, nd1, nr1 As Double
    
    If Selection.Count = 2 Then
        diff = 0
        nd1 = 0
        ratio = 1
        nr1 = 1
        For Each cell In Selection.Cells
            diff = nd1 - cell.Value
            nd1 = cell.Value
            ratio = cell.Value / nr1
            nr1 = cell.Value
        Next
        msg = _
            "Absolute differece:  " & Format(Abs(diff), "#,##0.0000") & vbCrLf & vbCrLf & _
            "Ratio A:  " & Format(1 / ratio, "#,##0.0000") & vbCrLf & _
            "Ratio B:  " & Format(ratio, "#,##0.0000")
    Else
        msg = "Invalid selection"
    End If
    
    MsgBox msg
End Sub
