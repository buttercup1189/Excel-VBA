Attribute VB_Name = "Module2"
Sub KolumnKleaner()
    Dim N As Long, wf As WorksheetFunction, M As Long
    Dim i As Long, j As Long

    N = Columns.Count
    M = Rows.Count
    Set wf = Application.WorksheetFunction

    For i = N To 1 Step -1
        If wf.CountBlank(Columns(i)) <> M Then Exit For
    Next i

    For j = i To 1 Step -1
        If wf.CountBlank(Columns(j)) = M Then
            Cells(1, j).EntireColumn.Delete
        End If
    Next j
End Sub

