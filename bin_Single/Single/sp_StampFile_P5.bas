Attribute VB_Name = "sp_StampFile_P5"
Public Sub p_StampFile_P5()

Dim wb(1 To 3) As Workbook

Set wb(2) = Workbooks("SourceData.xlsx")
wb(2).Activate

Set wb(1) = Workbooks("ResultsSingle")
wb(1).Activate


'Disable  alerts
With Application
.DisplayAlerts = False
.ScreenUpdating = False
End With


wb(2).Worksheets("Single Policy Inputs").Range("E6").Copy
wb(1).Worksheets(1).Range("G1:I1").PasteSpecial
wb(2).Worksheets("Single Policy Inputs").Range("B6").Copy
wb(1).Worksheets(1).Range("H3").PasteSpecial
wb(2).Worksheets("Single Policy Inputs").Range("K6").Copy
wb(1).Worksheets(1).Range("H4").PasteSpecial

wb(2).Worksheets("Single Policy Inputs").Range("M2").Copy
wb(1).Worksheets(1).Range("H5").PasteSpecial
wb(1).Worksheets(1).Columns("H").AutoFit
wb(1).Worksheets(1).Range("G1").Font.Bold = True
wb(1).Worksheets(1).Range("G1").Font.Size = 16
wb(1).Worksheets(1).Range("G1").Font.Color = vbWhite
wb(1).Worksheets(1).Range("G1:I1").HorizontalAlignment = xlCenter

End Sub

