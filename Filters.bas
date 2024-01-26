Attribute VB_Name = "Filters"
Option Explicit


Sub FilterByRange()
' FilteByRange Macro
    ActiveCell.Offset(0, 1).Select
    ActiveCell.EntireColumn.Insert Shift:=xlToRight
    Cells(ActiveSheet.AutoFilter.Range.Row, ActiveCell.Column).Select
    
    If InStr(Worksheets(ActiveSheet.Index + 1).Name, "Sheet") Then
        Selection.Value = "IS_IN_RANGE"
    Else
        Selection.Value = "IS_IN_" & Replace(Worksheets(ActiveSheet.Index + 1).Name, " ", "_")
    End If
    
    ActiveCell.Offset(1, 0).Select
    Selection.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "=AND(COUNTIF('" & Worksheets(ActiveSheet.Index + 1).Name & "'!C1,RC[-1])>0,RC[-1]<>"""")"
    ActiveCell.Copy
    Range(ActiveCell, Cells(ActiveSheet.AutoFilter.Range.Rows.Count + ActiveSheet.AutoFilter.Range.Row - 1, ActiveCell.Column)).PasteSpecial xlPasteFormulasAndNumberFormats
    ActiveCell.EntireColumn.AutoFit
    
    ActiveSheet.AutoFilter.Range.AutoFilter _
        Field:=ActiveCell.Column - ActiveSheet.AutoFilter.Range.Column + 1, _
        Criteria1:="TRUE"
    
    Cells(ActiveSheet.AutoFilter.Range.Row, ActiveCell.Column).Select
End Sub



Sub FilterByCopiedRange()
' FilterByCopiedRange Macro
    
    Dim copied_data, split_data() As String
    
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        copied_data = .GetText(1)
    End With
    
    ' if the last 2 characters of copied_data are a break-line character, remove them
    If Right(copied_data, 2) = vbCrLf Then
        copied_data = Left(copied_data, Len(copied_data) - 2)
    End If
    
    split_data = Split(copied_data, vbCrLf)
    
    ActiveSheet.AutoFilter.Range.AutoFilter _
        Field:=ActiveCell.Column - ActiveSheet.AutoFilter.Range.Column + 1, _
        Criteria1:=split_data, _
        Operator:=xlFilterValues
        
End Sub



Sub FilterByNotInCopiedRange()
' FilterByNotInCopiedRange Macro

    Dim dict_copied, dict_filter As Object
    Dim copied_data, split_data() As String
    Dim values_filter, split_filter() As String
    Dim act_cell As String
    Dim cell As Range
    Dim i As Integer
    
    Set dict_copied = CreateObject("Scripting.Dictionary")
    Set dict_filter = CreateObject("Scripting.Dictionary")
    
    act_cell = ActiveCell.Address
    
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        copied_data = .GetText(1)
    End With
    
    ' if the last 2 characters of copied_data are a break-line character, remove them
    If Right(copied_data, 2) = vbCrLf Then
        copied_data = Left(copied_data, Len(copied_data) - 2)
    End If
    
    split_data = Split(copied_data, vbCrLf)
    
    For i = LBound(split_data) To UBound(split_data)
        dict_copied.Add split_data(i), 1
    Next i
    
    Range( _
        Cells(ActiveSheet.AutoFilter.Range.Row + 1, ActiveCell.Column), _
        Cells(ActiveSheet.AutoFilter.Range.Rows.Count + ActiveSheet.AutoFilter.Range.Row - 1, ActiveCell.Column) _
        ).Select
    
    For Each cell In Selection.Cells
        If Not dict_filter.exists(cell.Value) And Not dict_copied.exists(cell.Value) And cell.Value <> "" Then
            dict_filter.Add cell.Value, 1
            values_filter = values_filter & cell.Value & vbCrLf
        End If
    Next

    values_filter = Left(values_filter, Len(values_filter) - 2)
    
    split_filter = Split(values_filter, vbCrLf)
    
    ActiveSheet.AutoFilter.Range.AutoFilter _
        Field:=ActiveCell.Column - ActiveSheet.AutoFilter.Range.Column + 1, _
        Criteria1:=split_filter, _
        Operator:=xlFilterValues
    
    Range(act_cell).Select
        
End Sub



Sub FilterByContainsCopiedRange()
' Only works for an array of max two values
' FilterByContainsCopiedRange Macro
    
    Dim copied_data, split_data() As String
    Dim i As Integer
    
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        copied_data = .GetText(1)
    End With
    
    ' if the last 2 characters of copied_data are a break-line character, remove them
    If Right(copied_data, 2) = vbCrLf Then
        copied_data = Left(copied_data, Len(copied_data) - 2)
    End If
    
    split_data = Split(copied_data, vbCrLf)
    
    For i = LBound(split_data) To UBound(split_data)
        split_data(i) = "*" & split_data(i) & "*"
        Debug.Print split_data(i)
    Next i
    
    ActiveSheet.AutoFilter.Range.AutoFilter _
        Field:=ActiveCell.Column - ActiveSheet.AutoFilter.Range.Column + 1, _
        Criteria1:=split_data, _
        Operator:=xlFilterValues
        
End Sub



Sub CellsWithFilter()
' CellsWithFilter Macro

' ActiveCell.ListObject is Nothing: False if ActiveCell is in a Table
' ActiveSheet.AutoFilterMode: True if there is filtermode somewhere in the sheet (even if there is no applied filter). Table doesn't count as filter mode (returns False)
' ActiveSheet.AutoFilter.FilterMode: True if there is an applied filter in the sheet
' Intersect(ActiveSheet.AutoFilter.Range, ActiveCell) Is Nothing: True if activecell is in filtermode region. Error if there is no filtermode region

    Dim i, f, i_start, i_end, active_col, filter_cols As Integer
    
    If Not (ActiveCell.ListObject Is Nothing) Then
        ActiveCell.Select
    ElseIf Not ActiveSheet.AutoFilterMode Then
        Exit Sub
    ElseIf Intersect(ActiveSheet.AutoFilter.Range, ActiveCell) Is Nothing Then
        Exit Sub
    End If
    
    active_col = ActiveCell.Column - ActiveSheet.AutoFilter.Range.Column + 1
    filter_cols = ActiveSheet.AutoFilter.Range.Columns.Count
    i_start = 0
    i_end = filter_cols - 1
    
    If ActiveCell.Row = ActiveSheet.AutoFilter.Range.Row Then
        i_start = 1
        i_end = i_end + 1
    End If

    For i = i_start To i_end
        If i > filter_cols - active_col Then
            f = i - filter_cols
        Else
            f = i
        End If

        If ActiveSheet.AutoFilter.Filters(active_col + f).On Then
            Cells(ActiveSheet.AutoFilter.Range.Row, active_col + f).Select
            Exit Sub
        End If
    Next
    
    Cells(ActiveSheet.AutoFilter.Range.Row, ActiveCell.Column).Select
    
End Sub

