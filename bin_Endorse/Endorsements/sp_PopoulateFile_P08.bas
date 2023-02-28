Attribute VB_Name = "sp_PopoulateFile_P08"

Public Sub p_PopulateFile_P08()

Dim wb(1 To 3) As Workbook
Dim sourceDoc As Excel.Range
Dim sourceRng(1 To 30) As Excel.Range

Call p_CreateTabs_P08
Call p_DB16_P08

Set wb(2) = Workbooks("Datadump.xlsx")
wb(2).Activate

Set wb(1) = Workbooks("ResultsEndorsement.xlsx")
wb(1).Activate

With wb(2).Worksheets("Response8")
Set sourceDoc = wb(1).Worksheets(1).Range("C3")
For Each sourceRng(1) In sourceDoc
    If Len(sourceRng(1).Text) > 0 Then
.Range("A30").Copy
wb(1).Worksheets(1).Range("N3").PasteSpecial
End If
Next sourceRng(1)

Set sourceDoc = wb(1).Worksheets(1).Range("C4")
For Each sourceRng(2) In sourceDoc
    If Len(sourceRng(2).Text) > 0 Then
.Range("A31").Copy
wb(1).Worksheets(1).Range("N4").PasteSpecial
End If
Next sourceRng(2)

Set sourceDoc = wb(1).Worksheets(1).Range("C5")
For Each sourceRng(3) In sourceDoc
    If Len(sourceRng(3).Text) > 0 Then
.Range("A32").Copy
wb(1).Worksheets(1).Range("N5").PasteSpecial
End If
Next sourceRng(3)

Set sourceDoc = wb(1).Worksheets(1).Range("C6")
For Each sourceRng(4) In sourceDoc
    If Len(sourceRng(4).Text) > 0 Then
.Range("A33").Copy
wb(1).Worksheets(1).Range("N6").PasteSpecial
End If
Next sourceRng(4)

Set sourceDoc = wb(1).Worksheets(1).Range("C7")
For Each sourceRng(5) In sourceDoc
    If Len(sourceRng(5).Text) > 0 Then
.Range("A34").Copy
wb(1).Worksheets(1).Range("N7").PasteSpecial
End If
Next sourceRng(5)

Set sourceDoc = wb(1).Worksheets(1).Range("C8")
For Each sourceRng(6) In sourceDoc
    If Len(sourceRng(6).Text) > 0 Then
.Range("A35").Copy
wb(1).Worksheets(1).Range("N8").PasteSpecial
End If
Next sourceRng(6)

Set sourceDoc = wb(1).Worksheets(1).Range("C9")
For Each sourceRng(7) In sourceDoc
    If Len(sourceRng(7).Text) > 0 Then
.Range("A36").Copy
wb(1).Worksheets(1).Range("N9").PasteSpecial
End If
Next sourceRng(7)

Set sourceDoc = wb(1).Worksheets(1).Range("C10")
For Each sourceRng(8) In sourceDoc
    If Len(sourceRng(8).Text) > 0 Then
.Range("A37").Copy
wb(1).Worksheets(1).Range("N10").PasteSpecial
Else
End If
Next sourceRng(8)

Set sourceDoc = wb(1).Worksheets(1).Range("C11")
For Each sourceRng(9) In sourceDoc
    If Len(sourceRng(9).Text) > 0 Then
.Range("A38").Copy
wb(1).Worksheets(1).Range("N11").PasteSpecial
Else
End If
Next sourceRng(9)

Set sourceDoc = wb(1).Worksheets(1).Range("C12")
For Each sourceRng(10) In sourceDoc
    If Len(sourceRng(10).Text) > 0 Then
.Range("A39").Copy
wb(1).Worksheets(1).Range("N12").PasteSpecial
Else
End If
Next sourceRng(10)

Set sourceDoc = wb(1).Worksheets(1).Range("C13")
For Each sourceRng(11) In sourceDoc
    If Len(sourceRng(11).Text) > 0 Then
.Range("A40").Copy
wb(1).Worksheets(1).Range("N13").PasteSpecial
Else
End If
Next sourceRng(12)

Set sourceDoc = wb(1).Worksheets(1).Range("C14")
For Each sourceRng(12) In sourceDoc
    If Len(sourceRng(12).Text) > 0 Then
.Range("A41").Copy
wb(1).Worksheets(1).Range("N14").PasteSpecial
Else
End If
Next sourceRng(12)

Set sourceDoc = wb(1).Worksheets(1).Range("C15")
For Each sourceRng(13) In sourceDoc
    If Len(sourceRng(13).Text) > 0 Then
.Range("A42").Copy
wb(1).Worksheets(1).Range("N15").PasteSpecial
Else
End If
Next sourceRng(13)

Set sourceDoc = wb(1).Worksheets(1).Range("C16")
For Each sourceRng(14) In sourceDoc
    If Len(sourceRng(14).Text) > 0 Then
.Range("A43").Copy
wb(1).Worksheets(1).Range("N16").PasteSpecial
Else
End If
Next sourceRng(14)

Set sourceDoc = wb(1).Worksheets(1).Range("C17")
For Each sourceRng(15) In sourceDoc
    If Len(sourceRng(15).Text) > 0 Then
.Range("A44").Copy
wb(1).Worksheets(1).Range("N17").PasteSpecial
Else
End If
Next sourceRng(15)

Set sourceDoc = wb(1).Worksheets(1).Range("C18")
For Each sourceRng(16) In sourceDoc
    If Len(sourceRng(16).Text) > 0 Then
.Range("A45").Copy
wb(1).Worksheets(1).Range("N18").PasteSpecial
Else
End If
Next sourceRng(16)

Set sourceDoc = wb(1).Worksheets(1).Range("C19")
For Each sourceRng(17) In sourceDoc
    If Len(sourceRng(17).Text) > 0 Then
.Range("A46").Copy
wb(1).Worksheets(1).Range("N19").PasteSpecial
Else
End If
Next sourceRng(17)

Set sourceDoc = wb(1).Worksheets(1).Range("C20")
For Each sourceRng(18) In sourceDoc
    If Len(sourceRng(18).Text) > 0 Then
.Range("A47").Copy
wb(1).Worksheets(1).Range("N20").PasteSpecial
Else
End If
Next sourceRng(18)

Set sourceDoc = wb(1).Worksheets(1).Range("C21")
For Each sourceRng(19) In sourceDoc
    If Len(sourceRng(19).Text) > 0 Then
.Range("A48").Copy
wb(1).Worksheets(1).Range("N21").PasteSpecial
Else
End If
Next sourceRng(19)

Set sourceDoc = wb(1).Worksheets(1).Range("C22")
For Each sourceRng(20) In sourceDoc
    If Len(sourceRng(20).Text) > 0 Then
.Range("A49").Copy
wb(1).Worksheets(1).Range("N22").PasteSpecial
Else
End If
Next sourceRng(20)
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
   
  'wb(3).Close
  
  Application.DisplayAlerts = False
Application.ScreenUpdating = False
  
  Call p_FormatFile_P08
  Call p_CompareResults_P08
  Call p_CheckMessages_P08

End Sub




