Attribute VB_Name = "sp_PopoulateFile_P1"
Public Sub p_PopulateFile_P1()

Dim wb(1 To 3) As Workbook
Dim sourceDoc As Excel.Range
Dim sourceRng(1 To 20) As Excel.Range

Call p_CreateTabs_P1
Call p_DB16_P1

Set wb(1) = Workbooks("ResultsSingle.xlsx")
wb(1).Activate

Set wb(2) = Workbooks("Datadump.xlsx")
wb(2).Activate

With wb(2).Worksheets("Response1")
Set sourceDoc = wb(1).Worksheets(1).Range("C3")
For Each sourceRng(1) In sourceDoc
    If Len(sourceRng(1).Text) > 0 Then
.Range("A15").Copy
wb(1).Worksheets(1).Range("N3").PasteSpecial
End If
Next sourceRng(1)

Set sourceDoc = wb(1).Worksheets(1).Range("C4")
For Each sourceRng(2) In sourceDoc
    If Len(sourceRng(2).Text) > 0 Then
.Range("A16").Copy
wb(1).Worksheets(1).Range("N4").PasteSpecial
End If
Next sourceRng(2)

Set sourceDoc = wb(1).Worksheets(1).Range("C5")
For Each sourceRng(3) In sourceDoc
    If Len(sourceRng(3).Text) > 0 Then
.Range("A17").Copy
wb(1).Worksheets(1).Range("N5").PasteSpecial
End If
Next sourceRng(3)

Set sourceDoc = wb(1).Worksheets(1).Range("C6")
For Each sourceRng(4) In sourceDoc
    If Len(sourceRng(4).Text) > 0 Then
.Range("A18").Copy
wb(1).Worksheets(1).Range("N6").PasteSpecial
End If
Next sourceRng(4)

Set sourceDoc = wb(1).Worksheets(1).Range("C7")
For Each sourceRng(5) In sourceDoc
    If Len(sourceRng(5).Text) > 0 Then
.Range("A19").Copy
wb(1).Worksheets(1).Range("N7").PasteSpecial
End If
Next sourceRng(5)

Set sourceDoc = wb(1).Worksheets(1).Range("C8")
For Each sourceRng(6) In sourceDoc
    If Len(sourceRng(6).Text) > 0 Then
.Range("A20").Copy
wb(1).Worksheets(1).Range("N8").PasteSpecial
End If
Next sourceRng(6)

Set sourceDoc = wb(1).Worksheets(1).Range("C9")
For Each sourceRng(7) In sourceDoc
    If Len(sourceRng(7).Text) > 0 Then
.Range("A21").Copy
wb(1).Worksheets(1).Range("N9").PasteSpecial
End If
Next sourceRng(7)

Set sourceDoc = wb(1).Worksheets(1).Range("C10")
For Each sourceRng(8) In sourceDoc
    If Len(sourceRng(8).Text) > 0 Then
.Range("A22").Copy
wb(1).Worksheets(1).Range("N10").PasteSpecial
Else
End If
Next sourceRng(8)

Set sourceDoc = wb(1).Worksheets(1).Range("C11")
For Each sourceRng(9) In sourceDoc
    If Len(sourceRng(9).Text) > 0 Then
.Range("A23").Copy
wb(1).Worksheets(1).Range("N11").PasteSpecial
Else
End If
Next sourceRng(9)

Set sourceDoc = wb(1).Worksheets(1).Range("C12")
For Each sourceRng(10) In sourceDoc
    If Len(sourceRng(10).Text) > 0 Then
.Range("A24").Copy
wb(1).Worksheets(1).Range("N12").PasteSpecial
Else
End If
Next sourceRng(10)
End With


Dim bIsEmpty As Boolean

With wb(1).Worksheets(1)
For Each Cell In .Range("G3")
  If IsEmpty(Cell) = True Then
  .Range("B3:F3").Value = "No Data Found"
  .Range("B3:F3").Merge
  .Range("B3:F3").Font.Size = 28
  .Range("B3:F3").Font.Bold = True
  .Range("B3:F3").HorizontalAlignment = xlCenter
  End If
  Next Cell
  For Each Cell In .Range("M3")
  If IsEmpty(Cell) = True Then
  .Range("J3:M3").Value = "No Data Found"
  .Range("J3:M3").Merge
  .Range("J3:M3").Font.Size = 28
  .Range("J3:M3").Font.Bold = True
  .Range("J3:M3").HorizontalAlignment = xlCenter
  End If
  Next Cell
  End With
 
  Call p_FormatFile_P1
  Call p_CompareResults_P1
  Call p_CheckMessages_P1
  Call p_StampFile_P1
  
End Sub

