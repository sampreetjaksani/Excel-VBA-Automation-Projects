Attribute VB_Name = "Module3"
Sub DIVISION_001()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Sheets("RAW").Range("H:H").Copy
Sheets("RAW").Range("U1").PasteSpecial
Sheets("RAW").Range("U:U").RemoveDuplicates Columns:=1, Header:=xlYes
Dim IR As Integer, X As Integer, Y As String
IR = Sheets("RAW").Range("U" & Rows.Count).End(xlUp).Row
For X = 2 To IR
Y = Sheets("RAW").Range("U" & X)
Sheets.Add.Name = Y
Sheets("RAW").Range("A1").AutoFilter FIELD:=8, Criteria1:=Y
Sheets("RAW").Range("A1").CurrentRegion.Copy
Sheets(Y).Range("A1").PasteSpecial
Sheets("RAW").Range("A1").AutoFilter
Next
Sheets("RAW").Range("U:U").Delete
Sheets("RAW").Select
Range("A1").Select
End Sub

Sub DELETING_SHEETS()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
For X = Sheets.Count To 1 Step -1
If Sheets(X).Name = "RAW" _
Or Sheets(X).Name = "SHEET1" Then
Else
Sheets(X).Delete
End If
Next
End Sub
