Attribute VB_Name = "Module1"
Sub GetSheets()
Path = "C:\Users\juby.lee\Documents\Fedexsheet\"
Filename = Dir(Path & "*.xls")
Do While Filename <> ""
Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
For Each Sheet In ActiveWorkbook.Sheets
Sheet.Copy After:=ThisWorkbook.Sheets(1)
Next Sheet
Workbooks(Filename).Close
Filename = Dir()
Loop
End Sub

