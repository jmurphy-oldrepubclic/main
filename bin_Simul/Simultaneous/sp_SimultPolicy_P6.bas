Attribute VB_Name = "sp_SimultPolicy_P6"
Public Sub p_SimultPolicy_P6()

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

'Disable  alerts
With Application
.DisplayAlerts = False
.ScreenUpdating = False
End With

Set wb(2) = Workbooks("SourceData.xlsx")
wb(2).Activate

    Set sourceDoc = wb(2).Sheets("Simultanious Policy Inputs").Range("C3")
For Each sourceRng(1) In sourceDoc
        With sourceRng(1)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a State - See State Code(s) tab", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(1)
 
 
     Set sourceDoc = wb(2).Sheets("Simultanious Policy Inputs").Range("F3")
For Each sourceRng(2) In sourceDoc
        With sourceRng(2)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a Trancode for Owners policy", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(2)
 
   Set sourceDoc = wb(2).Sheets("Simultanious Policy Inputs").Range("G3")
For Each sourceRng(3) In sourceDoc
        With sourceRng(3)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a Trancode for Loan policy", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(3)
 

 
     Set sourceDoc = wb(2).Sheets("Simultanious Policy Inputs").Range("H3")
For Each sourceRng(4) In sourceDoc
        With sourceRng(4)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a Policy Date", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(4)
 
    Set sourceDoc = wb(2).Sheets("Simultanious Policy Inputs").Range("I3:J3")
For Each sourceRng(5) In sourceDoc
        With sourceRng(5)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a Lower and Upper Liability for Owners policy", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(5)
  
    Set sourceDoc = wb(2).Sheets("Simultanious Policy Inputs").Range("K3")
For Each sourceRng(6) In sourceDoc
        With sourceRng(6)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a value for Credit Liability of $0 or greater for Owners policy", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(6)
 
      Set sourceDoc = wb(2).Sheets("Simultanious Policy Inputs").Range("L3:M3")
For Each sourceRng(7) In sourceDoc
        With sourceRng(7)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a Lower and Upper Liability for Loan policy", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(7)
 
  Set sourceDoc = wb(2).Sheets("Simultanious Policy Inputs").Range("N3")
For Each sourceRng(8) In sourceDoc
        With sourceRng(8)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a value for Credit Liability of $0 or greater for Owners policy", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(8)
 
 
Workbooks.Add
Set wb(1) = ActiveWorkbook
Sheets("Sheet1").Name = "DataSet1"



wb(1).Activate
 

adoDbConn(1).Open "Provider=SQLOLEDB;Data Source=mn-qua-db16;Initial Catalog=RatesEngineTest_vNext;Trusted_connection=yes;"

s = wb(2).Sheets("Simultanious Policy Inputs").Range("C3").Value
c = wb(2).Sheets("Simultanious Policy Inputs").Range("D3").Value
tr1 = wb(2).Sheets("Simultanious Policy Inputs").Range("F3").Value
ad = wb(2).Sheets("Simultanious Policy Inputs").Range("H3").Value
ll1 = wb(2).Sheets("Simultanious Policy Inputs").Range("I3").Value
ul1 = wb(2).Sheets("Simultanious Policy Inputs").Range("J3").Value
pl1 = wb(2).Sheets("Simultanious Policy Inputs").Range("K3").Value
t = wb(2).Sheets("Simultanious Policy Inputs").Range("O3").Value

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(1) = adoDbConn(1).Execute("Declare @n VARCHAR(20) Declare @l Decimal(18,2) Declare @cl Decimal(18,2) Declare @ed Date Declare @tr VARCHAR(20) Declare @cc VARCHAR(4) select top 10 o.StateCode, o.CountyCode, o.OrderNumber, p.Trancode, p.EffectiveDate, p.Liability, p.CreditLiability from Orders o inner join Policies p on p.OrderId = o.Id inner join OrderTags ot on o.Id = ot.Order_Id inner join tags t on ot.Tag_Id = t.Id where o.StateCode = '" & s & "' and o.CountyCode like isnull(@cc, '%" & c & "%') and p.TranCode  = '" & tr1 & "' and p.EffectiveDate >= isnull(@ed,'" & ad & "') and p.Liability >= isnull(@l,'" & ll1 & "') and p.Liability <= isnull(@l, '" & ul1 & "') and p.CreditLiability >= isnull(@cl,'" & pl1 & "') and o.OrderNumber in (SELECT OrderNumber FROM Orders GROUP BY OrderNumber HAVING COUNT(OrderNumber)=1) Order By o.OrderNumber asc")

  If Not adoDbRs(1).EOF Then
  'Transfer result.
    On Error Resume Next
  wb(1).Worksheets("DataSet1").Range("B2:H2").CopyFromRecordset adoDbRs(1)
  'Close the recordset
  adoDbRs(1).Close
End If
adoDbConn(1).Close

'Insert columns with header names, make bold and add background color
wb(1).Worksheets("DataSet1").Range("I2").Value = "AgencyNumber"
wb(2).Sheets("Simultanious Policy Inputs").Range("B3").Copy

With wb(1).Worksheets("DataSet1")
.Range("A2").PasteSpecial
.Range("J2").Value = "StateCode"
.Range("J2:J10000").NumberFormat = "@"
.Range("K2").Value = "CountyCode"
.Range("K2:K10000").NumberFormat = "@"
.Range("L2").Value = "TranCode"
.Range("L2:L10000").NumberFormat = "@"
.Range("M2").Value = "EffectiveDate"
.Range("M2:M10000").NumberFormat = "yyyy-mm-dd"
.Range("N2").Value = "Liability"
.Range("O2").Value = "CreditLiability"
.Range("P2").Value = """"
.Range("Q2").Value = ""","""
.Range("R2").Value = """:"""
.Range("S2").Value = "{"
.Range("T2").Value = "["
.Range("U2").Value = "}"
.Range("V2").Value = "]"
.Range("W2").Value = """"
.Range("X2").Value = ":"
End With
'
'Autofill columns to last row of columns with the copied source data
LastRow = wb(1).Worksheets("DataSet1").Range("B" & Rows.Count).End(xlUp).Row
''Copy data down as per top cell
With wb(1).Worksheets("DataSet1")
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
End With

Sheets.Add
Sheets("Sheet2").Name = "DataSet2"

adoDbConn(2).Open "Provider=SQLOLEDB;Data Source=mn-qua-db16;Initial Catalog=RatesEngineTest_vNext;Trusted_connection=yes;"

tr2 = wb(2).Sheets("Simultanious Policy Inputs").Range("G3").Value
ll2 = wb(2).Sheets("Simultanious Policy Inputs").Range("L3").Value
ul2 = wb(2).Sheets("Simultanious Policy Inputs").Range("M3").Value
pl2 = wb(2).Sheets("Simultanious Policy Inputs").Range("N3").Value


'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(2) = adoDbConn(2).Execute("Declare @n VARCHAR(20) Declare @l Decimal(18,2) Declare @cl Decimal(18,2) Declare @ed Date Declare @tr VARCHAR(20) Declare @cc VARCHAR(4) select top 10 max(o.StateCode), max(o.CountyCode), o.OrderNumber, p.Trancode, max(p.EffectiveDate), max(p.Liability), max(p.CreditLiability) from Orders o inner join Policies p on p.OrderId = o.Id inner join OrderTags ot on o.Id = ot.Order_Id inner join tags t on ot.Tag_Id = t.Id where o.StateCode = '" & s & "' and o.CountyCode like isnull(@cc, '%" & c & "%') and p.TranCode = '" & tr2 & "' and p.EffectiveDate >= isnull(@ed,'" & ad & "') and p.Liability >= isnull(@l,'" & ll2 & "') and p.Liability <= isnull(@l, '" & ul2 & "') and p.CreditLiability >= isnull(@cl,'" & pl2 & "') and t.Name like isnull(@n, '%" & t & "%') Group by o.OrderNumber, p.TranCode Order by o.OrderNumber, p.TranCode desc")

  If Not adoDbRs(2).EOF Then
  'Transfer result.
    On Error Resume Next
   wb(1).Worksheets("DataSet2").Range("B2:H2").CopyFromRecordset adoDbRs(2)
  'Close the recordset
  adoDbRs(2).Close
End If

adoDbConn(2).Close

wb(1).Worksheets("DataSet2").Range("I2").Value = "AgencyNumber"
wb(2).Sheets("Simultanious Policy Inputs").Range("B3").Copy

With wb(1).Worksheets("DataSet2")
.Range("A2").PasteSpecial
.Range("J2").Value = "StateCode"
.Range("J2:J10000").NumberFormat = "@"
.Range("K2").Value = "CountyCode"
.Range("K2:K10000").NumberFormat = "@"
.Range("L2").Value = "TranCode"
.Range("L2:L10000").NumberFormat = "@"
.Range("M2").Value = "EffectiveDate"
.Range("M2:M10000").NumberFormat = "yyyy-mm-dd"
.Range("N2").Value = "Liability"
.Range("O2").Value = "CreditLiability"
.Range("P2").Value = """"
.Range("Q2").Value = ""","""
.Range("R2").Value = """:"""
.Range("S2").Value = "{"
.Range("T2").Value = "["
.Range("U2").Value = "}"
.Range("V2").Value = "]"
.Range("W2").Value = """"
.Range("X2").Value = ":"
.Range("Y2").Value = ","
End With
'
'Autofill columns to last row of columns with the copied source data
LastRow = wb(1).Worksheets("DataSet2").Range("B" & Rows.Count).End(xlUp).Row
'Copy data down as per top cell
With wb(1).Worksheets("DataSet2")
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
End With

wb(1).Activate
Application.ScreenUpdating = False
wb(1).SaveAs "File6.xlsx"


End Sub





