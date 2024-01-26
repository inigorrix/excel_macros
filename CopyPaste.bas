Attribute VB_Name = "CopyPaste"
Option Explicit


Sub CopyValues()
' Copy selected text as a list of values

    Dim cell As Object
    Dim values As String
    
    If Selection.Cells.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    
    For Each cell In Selection.Cells
        values = values & cell.Value & vbCrLf
    Next
    
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText Left(values, Len(values) - 2)
        .PutInClipboard
    End With
End Sub



Sub CopyUniqueValues()
' Copy selected text as a list of unique values

    Dim dict As Object
    Dim cell As Range
    Dim values As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    If Selection.Cells.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    
    For Each cell In Selection.Cells
        If Not dict.exists(cell.Value) And cell.Value <> "" Then
            dict.Add cell.Value, 1
            values = values & cell.Value & vbCrLf
        End If
    Next
    
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText Left(values, Len(values) - 2)
        .PutInClipboard
    End With

End Sub



Sub PasteInRegions()
'''' Macro in Beta ''''
' PasteInRegions macro
    Dim copied_data, split_data() As String
    Dim i As Integer
    Dim c As Object
    
    If Selection.Cells.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        copied_data = .GetText(1)
    End With
    
    copied_data = Replace(copied_data, ",", ".")
    
    ' if the last 2 characters of copied_data are a break-line character, remove them
    If Right(copied_data, 2) = vbCrLf Then
        copied_data = Left(copied_data, Len(copied_data) - 2)
    End If
    
    split_data = Split(copied_data, vbCrLf)
    i = 0
    
    'Debug.Print (UBound(split_data) - LBound(split_data))
    
    If UBound(split_data) - LBound(split_data) = Selection.Count - 1 Then
        For Each c In Selection.Cells
            c.Value = split_data(i)
            i = i + 1
        Next
    End If

End Sub
