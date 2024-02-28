Attribute VB_Name = "Module2"
Sub Combine()
Dim J As Integer
On Error Resume Next
Sheets(1).Select
Worksheets.Add ' add a sheet in first place
Sheets(1).Name = "Combined"
Sheets(2).Activate
Range("A1").EntireRow.Select
Selection.Copy Destination:=Sheets(1).Range("A1")
For J = 2 To Sheets.Count ' from sheet 2 to last sheet
Sheets(J).Activate ' make the sheet active
Range("A1").Select
Selection.CurrentRegion.Select ' select all cells in this sheets
Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
Next
End Sub

