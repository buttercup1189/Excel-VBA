Attribute VB_Name = "Module1"
Sub remove_empty_lines()
Dim start_time, arr(), rng, row_num As Long, column_num As Long, i, j
    Application.ScreenUpdating = False
    row_num = ActiveSheet.UsedRange.Rows.Count
    column_num = ActiveSheet.UsedRange.Columns.Count
    ReDim arr(1 To row_num)
    rng = ActiveSheet.UsedRange
    start_time = Timer
    For i = 1 To row_num
       For j = 1 To column_num
        If rng(i, j) <> "" Then arr(i) = i
     Next j, i
    With Cells(ActiveSheet.UsedRange.Row, ActiveSheet.UsedRange.Column + column_num).Resize(row_num, 1)
    .Value = WorksheetFunction.Transpose(arr)
    ActiveSheet.UsedRange.Sort Key1:=Cells(ActiveSheet.UsedRange.Row, ActiveSheet.UsedRange.Column + column_num)
    .Value = ""
   End With
    Application.ScreenUpdating = True
    MsgBox Format(Timer - start_time, "0.00s")
End Sub
