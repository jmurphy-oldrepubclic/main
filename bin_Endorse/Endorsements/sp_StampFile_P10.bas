Attribute VB_Name = "sp_StampFile_P10"
Public Sub p_StampFile_P10()

Dim wb(1 To 3) As Workbook


Set wb(2) = Workbooks("SourceData.xlsx")
wb(2).Activate

Set wb(1) = Workbooks("ResultsEndorsement")
wb(1).Activate

'Disable  alerts
With Application
.DisplayAlerts = False
.ScreenUpdating = False
End With


wb(2).Worksheets("Policy with Endor Inputs").Range("E2").Copy
wb(1).Worksheets(1).Range("G1:I1").PasteSpecial
wb(2).Worksheets("Policy with Endor Inputs").Range("B2").Copy
wb(1).Worksheets(1).Range("H3").PasteSpecial
wb(2).Worksheets("Policy with Endor Inputs").Range("K2").Copy
wb(1).Worksheets(1).Range("H4").PasteSpecial

wb(2).Worksheets("Policy with Endor Inputs").Range("M2").Copy
wb(1).Worksheets(1).Range("H5").PasteSpecial
wb(1).Worksheets(1).Columns("H").AutoFit
wb(1).Worksheets(1).Range("G1").Font.Bold = True
wb(1).Worksheets(1).Range("G1").Font.Size = 16
wb(1).Worksheets(1).Range("G1").Font.Color = vbWhite
wb(1).Worksheets(1).Range("G1:I1").HorizontalAlignment = xlCenter

End Sub




