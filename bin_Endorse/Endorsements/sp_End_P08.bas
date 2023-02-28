Attribute VB_Name = "sp_End_P08"
Function allBlank(rg As Range) As Boolean
    allBlank = (rg.Cells.Count = WorksheetFunction.CountBlank(rg))
End Function

Public Sub p_End_P08()

Dim wb(1 To 5) As Workbook
Dim ws As Worksheet
Dim LastRow As Long
Dim StringOne As String
Dim StringTwo As String
Dim StringThree As String
Dim StringFour As String
Dim StringFive As String
Dim StringSix As String
Dim StringSeven As String
Dim FilePath As String
Dim adoDbConn(1 To 10) As New ADODB.Connection

Dim adoDbRs(1 To 10) As New ADODB.Recordset
Dim sourceDoc As Excel.Range
Dim sourceRng(1 To 20) As Excel.Range

Set wb(2) = Workbooks("SourceData.xlsx")
wb(2).Activate

    Set sourceDoc = wb(2).Sheets("Policy with Endor Inputs").Range("C3")
For Each sourceRng(1) In sourceDoc
        With sourceRng(1)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a State - See State Code(s) tab.", vbCritical
            Exit Sub
        End If
    End With
 Next sourceRng(1)
 
 
          If allBlank(wb(2).Sheets("Policy with Endor Inputs").Range("F3:G3")) Then
         MsgBox "Error: Enter a Trancode for Owners or Loan Policy.", vbCritical
        Exit Sub
    End If

 
 
     Set sourceDoc = wb(2).Sheets("Policy with Endor Inputs").Range("H3")
For Each sourceRng(4) In sourceDoc
        With sourceRng(4)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
            Exit Sub
        End If
    End With
 Next sourceRng(4)
 
 

  
    Set sourceDoc = wb(2).Sheets("Policy with Endor Inputs").Range("M3")
For Each sourceRng(6) In sourceDoc
        With sourceRng(6)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a value for Credit Liability of $0 or greater.", vbCritical
            Exit Sub
        End If
    End With
 Next sourceRng(6)
 

Set wb(1) = Workbooks.Add
Sheets("Sheet1").Name = "DataSet1"

wb(1).Activate

adoDbConn(1).Open "Provider=SQLOLEDB;Data Source=mn-qua-db16;Initial Catalog=RatesEngineTest_vNext;Trusted_connection=yes;"

With wb(2).Sheets("Policy with Endor Inputs")
s = .Range("C3").Value
c = .Range("D3").Value
tr1 = .Range("F3").Value
ad = .Range("H3").Value
ll1 = .Range("I3").Value
ul1 = .Range("J3").Value
pl1 = .Range("M3").Value
t = .Range("P3").Value
e1 = .Range("N4").Value
End With

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(1) = adoDbConn(1).Execute("DECLARE @s VARCHAR(2)  = '" & s & "', @ed DATETIME = '" & ad & "', @otr VARCHAR(10) = '" & tr1 & "', @oec VARCHAR(10) = '" & e1 & "', @cc VARCHAR(4) = '%" & c & "%', @ll1 Decimal(18,2) = '" & ll1 & "', @ul1 Decimal(18,2) = '" & ul1 & "', @cl Decimal(18,2) = '" & pl1 & "', @n VARCHAR(50) = '%" & t & "%' SELECT top 10 o.StateCode, o.CountyCode, o.OrderNumber, p.TranCode, p.EffectiveDate, p.Liability, p.CreditLiability, e.code FROM Tests te JOIN TestTags tt ON tt.Test_Id = te.Id JOIN Tags ta ON ta.Id = tt.Tag_Id JOIN OrderTags ot ON ot.Tag_Id = ta.Id JOIN Orders o ON o.Id = ot.Order_Id JOIN Policies p ON p.OrderId = o.Id JOIN Endorsements e on p.Id = e.PolicyId JOIN EndorsementResults er on e.id = er.EndorsementId WHERE o.StateCode = @s AND p.TranCode = @otr AND e.code = @oec AND p.EffectiveDate >= @ed and p.Liability >= @ll1 and p.Liability <= @ul1 and p.CreditLiability >= @cl")

  If Not adoDbRs(1).EOF Then
  'Transfer result.

  wb(1).Worksheets("DataSet1").Range("B2:I2").CopyFromRecordset adoDbRs(1)
  'Close the recordset
adoDbRs(1).Close
End If

  
'Insert columns with header names, make bold and add background color
With wb(1).Worksheets("DataSet1")
.Range("J2").Value = "AgencyNumber"
wb(2).Sheets("Policy with Endor Inputs").Range("B3").Copy
.Range("A2").PasteSpecial
.Range("K2").Value = "StateCode"
.Range("K2:K10000").NumberFormat = "@"
.Range("L2").Value = "CountyCode"
.Range("L2:L10000").NumberFormat = "@"
.Range("M2").Value = "TranCode"
.Range("M2:M10000").NumberFormat = "@"
.Range("N2").Value = "EffectiveDate"
.Range("N2:N10000").NumberFormat = "yyyy-mm-dd"
.Range("O2").Value = "Liability"
.Range("P2").Value = "CreditLiability"
.Range("Q2").Value = """"
.Range("R2").Value = ""","""
.Range("S2").Value = """:"""
.Range("T2").Value = "{"
.Range("U2").Value = "["
.Range("V2").Value = "}"
.Range("W2").Value = "]"
.Range("X2").Value = """"
.Range("Y2").Value = ":"
.Range("Z2").Value = ","
End With

With Application
.DisplayAlerts = False
.ScreenUpdating = False
End With
'
'Autofill columns to last row of columns with the copied source data
LastRow = wb(1).Worksheets("DataSet1").Cells(Rows.Count, "B").End(xlUp).Row
''Copy data down as per top cell
With wb(1).Worksheets("DataSet1")
If LastRow > 2 Then
.Range("A2").AutoFill Destination:=Range("A2:A" & LastRow), Type:=xlFillCopy
.Range("I2").AutoFill Destination:=Range("I2:I" & LastRow), Type:=xlFillCopy
.Range("J2").AutoFill Destination:=Range("J2:J" & LastRow), Type:=xlFillCopy
.Range("K2").AutoFill Destination:=Range("K2:K" & LastRow), Type:=xlFillCopy
.Range("L2").AutoFill Destination:=Range("L2:L" & LastRow), Type:=xlFillCopy
.Range("M2").AutoFill Destination:=Range("M2:M" & LastRow), Type:=xlFillCopy
.Range("N2").AutoFill Destination:=Range("N2:N" & LastRow), Type:=xlFillCopy
.Range("P2").AutoFill Destination:=Range("P2:P" & LastRow), Type:=xlFillCopy
.Range("O2").AutoFill Destination:=Range("O2:O" & LastRow), Type:=xlFillCopy
.Range("Q2").AutoFill Destination:=Range("Q2:Q" & LastRow), Type:=xlFillCopy
.Range("S2").AutoFill Destination:=Range("S2:S" & LastRow), Type:=xlFillCopy
.Range("T2").AutoFill Destination:=Range("T2:T" & LastRow), Type:=xlFillCopy
.Range("R2").AutoFill Destination:=Range("R2:R" & LastRow), Type:=xlFillCopy
.Range("V2").AutoFill Destination:=Range("V2:V" & LastRow), Type:=xlFillCopy
.Range("U2").AutoFill Destination:=Range("U2:U" & LastRow), Type:=xlFillCopy
.Range("W2").AutoFill Destination:=Range("W2:W" & LastRow), Type:=xlFillCopy
.Range("X2").AutoFill Destination:=Range("X2:X" & LastRow), Type:=xlFillCopy
.Range("Y2").AutoFill Destination:=Range("Y2:Y" & LastRow), Type:=xlFillCopy
.Range("Z2").AutoFill Destination:=Range("Z2:Z" & LastRow), Type:=xlFillCopy
End If
End With

wb(1).Sheets.Add
wb(1).Sheets("Sheet2").Name = "DataSet2"

wb(1).Activate

With wb(2).Sheets("Policy with Endor Inputs")
s = .Range("C3").Value
c = .Range("D3").Value
tr2 = .Range("G3").Value
ad = .Range("H3").Value
ll2 = .Range("K3").Value
ul2 = .Range("L3").Value
pl1 = .Range("M3").Value
t = .Range("P3").Value
e2 = .Range("O4").Value
End With

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(2) = adoDbConn(1).Execute("DECLARE @s VARCHAR(2)  = '" & s & "', @ed DATETIME = '" & ad & "', @ltr VARCHAR(10) = '" & tr2 & "', @lec VARCHAR(10) = '" & e2 & "', @cc VARCHAR(4) = '%" & c & "%', @ll2 Decimal(18,2) = '" & ll2 & "', @ul2 Decimal(18,2) = '" & ul2 & "', @cl Decimal(18,2) = '" & pl1 & "', @n VARCHAR(50) = '%" & t & "%' SELECT top 10 o.StateCode, o.CountyCode, o.OrderNumber, p.TranCode, p.EffectiveDate, p.Liability, p.CreditLiability, e.code FROM Tests te JOIN TestTags tt ON tt.Test_Id = te.Id JOIN Tags ta ON ta.Id = tt.Tag_Id JOIN OrderTags ot ON ot.Tag_Id = ta.Id JOIN Orders o ON o.Id = ot.Order_Id JOIN Policies p ON p.OrderId = o.Id JOIN Endorsements e on p.Id = e.PolicyId JOIN EndorsementResults er on e.id = er.EndorsementId WHERE o.StateCode = @s AND p.TranCode = @ltr AND e.code = @lec AND p.EffectiveDate >= @ed and p.Liability >= @ll2 and p.Liability <= @ul2 and p.CreditLiability >= @cl")

  If Not adoDbRs(2).EOF Then
  'Transfer result.

  wb(1).Worksheets("DataSet2").Range("B2:I2").CopyFromRecordset adoDbRs(2)
  'Close the recordset
adoDbRs(2).Close
End If
  adoDbConn(1).Close
  
'Insert columns with header names, make bold and add background color
With wb(1).Worksheets("DataSet2")
.Range("J2").Value = "AgencyNumber"
wb(2).Sheets("Policy with Endor Inputs").Range("B3").Copy
.Range("A2").PasteSpecial
.Range("K2").Value = "StateCode"
.Range("K2:K10000").NumberFormat = "@"
.Range("L2").Value = "CountyCode"
.Range("L2:L10000").NumberFormat = "@"
.Range("M2").Value = "TranCode"
.Range("M2:M10000").NumberFormat = "@"
.Range("N2").Value = "EffectiveDate"
.Range("N2:N10000").NumberFormat = "yyyy-mm-dd"
.Range("O2").Value = "Liability"
.Range("P2").Value = "CreditLiability"
.Range("Q2").Value = """"
.Range("R2").Value = ""","""
.Range("S2").Value = """:"""
.Range("T2").Value = "{"
.Range("U2").Value = "["
.Range("V2").Value = "}"
.Range("W2").Value = "]"
.Range("X2").Value = """"
.Range("Y2").Value = ":"
.Range("Z2").Value = ","
End With
'
'Autofill columns to last row of columns with the copied source data
LastRow = wb(1).Worksheets("DataSet2").Cells(Rows.Count, "B").End(xlUp).Row
''Copy data down as per top cell
With wb(1).Worksheets("DataSet2")
If LastRow > 2 Then
.Range("A2").AutoFill Destination:=Range("A2:A" & LastRow), Type:=xlFillCopy
.Range("I2").AutoFill Destination:=Range("I2:I" & LastRow), Type:=xlFillCopy
.Range("J2").AutoFill Destination:=Range("J2:J" & LastRow), Type:=xlFillCopy
.Range("K2").AutoFill Destination:=Range("K2:K" & LastRow), Type:=xlFillCopy
.Range("L2").AutoFill Destination:=Range("L2:L" & LastRow), Type:=xlFillCopy
.Range("M2").AutoFill Destination:=Range("M2:M" & LastRow), Type:=xlFillCopy
.Range("N2").AutoFill Destination:=Range("N2:N" & LastRow), Type:=xlFillCopy
.Range("P2").AutoFill Destination:=Range("P2:P" & LastRow), Type:=xlFillCopy
.Range("O2").AutoFill Destination:=Range("O2:O" & LastRow), Type:=xlFillCopy
.Range("Q2").AutoFill Destination:=Range("Q2:Q" & LastRow), Type:=xlFillCopy
.Range("S2").AutoFill Destination:=Range("S2:S" & LastRow), Type:=xlFillCopy
.Range("T2").AutoFill Destination:=Range("T2:T" & LastRow), Type:=xlFillCopy
.Range("R2").AutoFill Destination:=Range("R2:R" & LastRow), Type:=xlFillCopy
.Range("V2").AutoFill Destination:=Range("V2:V" & LastRow), Type:=xlFillCopy
.Range("U2").AutoFill Destination:=Range("U2:U" & LastRow), Type:=xlFillCopy
.Range("W2").AutoFill Destination:=Range("W2:W" & LastRow), Type:=xlFillCopy
.Range("X2").AutoFill Destination:=Range("X2:X" & LastRow), Type:=xlFillCopy
.Range("Y2").AutoFill Destination:=Range("Y2:Y" & LastRow), Type:=xlFillCopy
.Range("Z2").AutoFill Destination:=Range("Z2:Z" & LastRow), Type:=xlFillCopy
End If
End With

wb(1).Activate
Application.ScreenUpdating = False
wb(1).SaveAs "File8.xlsx"

End Sub





