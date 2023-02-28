Attribute VB_Name = "sp_End_RemoveEmptyRows"
Public Sub p_End_RemoveEmptyRows()

Dim wb(1 To 5) As Workbook
Set wb(1) = Workbooks("ResultsEndorsement")
wb(1).Activate

  Dim r As Range, rows As Long, i As Long
  Set r = wb(1).Worksheets(1).Range("A3:Q23")
  rows = r.rows.Count
  For i = rows To 1 Step (-1)
    If WorksheetFunction.CountA(r.rows(i)) = 0 Then r.rows(i).Delete
  Next

End Sub



