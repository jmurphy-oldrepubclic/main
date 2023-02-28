Attribute VB_Name = "sp_FormatFile_P1"

Public Sub p_FormatFile_P1()
Dim wb(1 To 3) As Workbook

Set wb(1) = Workbooks("ResultsSingle.xlsx")
wb(1).Activate


'Insert columns with header names, make bold and add background color
On Error Resume Next

With wb(1).Worksheets(1)
.Range("A3:A4").Font.Bold = True
.Range("B1").Value = "Expected Calculation"
.Range("B1").Font.Bold = True
.Range("B1").Font.Size = 28
.Range("B1").Font.Color = vbWhite
.Range("J1").Value = "Actual Calculation"
.Range("J1").Font.Bold = True
.Range("J1").Font.Size = 28
.Range("J1").Font.Color = vbWhite
.Range("B2").Value = "OrderNumber"
.Range("B2").Font.Color = vbWhite
.Range("C2").Value = "TranCodeID"
.Range("C3:C10000").NumberFormat = "@"
.Range("C2").Font.Color = vbWhite
.Range("D2").Value = "Policy Date"
.Range("D2").Font.Color = vbWhite
.Range("E2").Value = "Liability"
.Range("E3:E10000").NumberFormat = "$#,##0.00"
.Range("E2").Font.Color = vbWhite
.Range("F2").Value = "Credit Liability"
.Range("F3:F10000").NumberFormat = "$#,##0.00"
.Range("F2").Font.Color = vbWhite
.Range("G2").Value = "Gross"
.Range("G3:G10000").NumberFormat = "###0.00"
.Range("G2").Font.Color = vbWhite
.Range("I2").Value = "OrderNumber"
.Range("I2").Font.Color = vbWhite
.Range("J2").Value = "TranCodeID"
.Range("J3:J10000").NumberFormat = "@"
.Range("J2").Font.Color = vbWhite
.Range("K2").Value = "Policy Date"
.Range("K2").Font.Color = vbWhite
.Range("L2").Value = "Liability"
.Range("L3:L10000").NumberFormat = "$#,##0.00"
.Range("L2").Font.Color = vbWhite
.Range("M2").Value = "Credit Liability"
.Range("M3:M10000").NumberFormat = "$#,##0.00"
.Range("M2").Font.Color = vbWhite
.Range("N2").Value = "Gross"
.Range("N3:N10000").NumberFormat = "###0.00"
.Range("N2").Font.Color = vbWhite
.Range("P2").Value = "TEST Results"
.Range("P2").Font.Color = vbWhite
End With


'Start formatting of data grid
Dim LastRow As Long

With wb(1).Worksheets(1)
.Range("B1:F1").Merge
.Range("J1:N1").Merge
.Range("G1:I1").Merge
.Range("B2:P2").EntireColumn.AutoFit
.Range("B1:F1").HorizontalAlignment = xlCenter
.Range("G1:I1").HorizontalAlignment = xlCenter
.Range("J1:N1").HorizontalAlignment = xlCenter
.Range("A1:Q1").Interior.ColorIndex = 11
.Range("A2:Q2").Interior.ColorIndex = 11
LastRow = .Range("G" & Rows.Count).End(xlUp).Row
.Range("A2").AutoFill Destination:=.Range("A2:A" & LastRow), Type:=xlFillCopy
.Range("H2").AutoFill Destination:=.Range("H2:H" & LastRow), Type:=xlFillCopy
.Range("O2").AutoFill Destination:=.Range("O2:O" & LastRow), Type:=xlFillCopy
.Range("Q2").AutoFill Destination:=.Range("Q2:Q" & LastRow), Type:=xlFillCopy
'.Range("P3").AutoFill Destination:=.Range("P3:P" & lastRow), Type:=xlFillCopy
End With


End Sub
