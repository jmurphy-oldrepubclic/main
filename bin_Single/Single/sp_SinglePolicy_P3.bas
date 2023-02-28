Attribute VB_Name = "sp_SinglePolicy_P3"
Public Sub p_SinglePolicy_P3()

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
Dim Command ' as new ADODB.Command

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

   Set sourceDoc = wb(2).Sheets("Single Policy Inputs").Range("E4")
For Each sourceRng(1) In sourceDoc
        With sourceRng(1)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(1)

'Open new DataSet3 document and set to active
Workbooks.Add
Set wb(1) = ActiveWorkbook
ActiveWindow.WindowState = xlMaximized

'Rename worksheet to DataSet3
Sheets("Sheet1").Name = "DataSet3"

wb(1).Activate

 On Error Resume Next

    Set sourceDoc = wb(2).Sheets("Single Policy Inputs").Range("C4")
For Each sourceRng(2) In sourceDoc
        With sourceRng(2)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a State - See State Code(s) tab", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(2)
 
 
     Set sourceDoc = wb(2).Sheets("Single Policy Inputs").Range("G4")
For Each sourceRng(4) In sourceDoc
        With sourceRng(4)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a Policy Date", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(4)
 
    Set sourceDoc = wb(2).Sheets("Single Policy Inputs").Range("H4:I4")
For Each sourceRng(5) In sourceDoc
        With sourceRng(5)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a Lower and Upper Liability", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(5)
 
    Set sourceDoc = wb(2).Sheets("Single Policy Inputs").Range("J4")
For Each sourceRng(6) In sourceDoc
        With sourceRng(6)
            If Right(.Value, 1) = "" And Len(.Offset(1).Value) = 0 Then
                MsgBox "Error: Enter a value for Credit Liability of $0 or greater", vbCritical
                wb(1).Close
            Exit Sub
        End If
    End With
 Next sourceRng(6)
 

adoDbConn(1).Open "Provider=SQLOLEDB;Data Source=mn-qua-db16;Initial Catalog=RatesEngineTest_vNext;Trusted_connection=yes;"

s = wb(2).Sheets("Single Policy Inputs").Range("C4").Value
c = wb(2).Sheets("Single Policy Inputs").Range("D4").Value
tr = wb(2).Sheets("Single Policy Inputs").Range("F4").Value
ad = wb(2).Sheets("Single Policy Inputs").Range("G4").Value
ll = wb(2).Sheets("Single Policy Inputs").Range("H4").Value
ul = wb(2).Sheets("Single Policy Inputs").Range("I4").Value
pl = wb(2).Sheets("Single Policy Inputs").Range("J4").Value
t = wb(2).Sheets("Single Policy Inputs").Range("K4").Value

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(1) = adoDbConn(1).Execute("Declare @n VARCHAR(50) Declare @l Decimal(18,2) Declare @cl Decimal(18,2) Declare @ed Date Declare @tr VARCHAR(20) Declare @cc VARCHAR(4) select top 10 o.StateCode, o.CountyCode, o.OrderNumber, p.Trancode, p.EffectiveDate, p.Liability, p.CreditLiability from Orders o inner join Policies p on p.OrderId = o.Id inner join OrderTags ot on o.Id = ot.Order_Id inner join tags t on ot.Tag_Id = t.Id where o.StateCode = '" & s & "' and o.CountyCode like isnull(@cc, '%" & c & "%') and p.TranCode like isnull(@tr,'%" & tr & "%') and p.EffectiveDate >= isnull(@ed,'" & ad & "') and p.Liability >= isnull(@l,'" & ll & "') and p.Liability <= isnull(@l, '" & ul & "') and p.CreditLiability >= isnull(@cl,'" & pl & "') and t.Name like isnull(@n, '%" & t & "%') and o.OrderNumber in (SELECT OrderNumber FROM Orders GROUP BY OrderNumber HAVING COUNT(OrderNumber)=1) Order By o.OrderNumber asc")

  If Not adoDbRs(1).EOF Then
  'Transfer result.
    On Error Resume Next
  wb(1).Sheets("DataSet3").Range("B2:H2").CopyFromRecordset adoDbRs(1)
  'Close the recordset
  adoDbRs(1).Close
End If

adoDbConn(1).Close

Application.CutCopyMode = False ' don't want an existing operation to interfere

'Insert columns with header names, make bold and add background color

wb(1).Sheets("DataSet3").Range("I2").Value = "AgencyNumber"
wb(2).Sheets("Single Policy Inputs").Range("B4").Copy

With wb(1).Sheets("DataSet3")
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
''Autofill columns to last row of columns with the copied source data
LastRow = wb(1).Sheets("DataSet3").Range("B" & Rows.Count).End(xlUp).Row
''Copy data down as per top cell
With wb(1).Sheets("DataSet3")
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

wb(1).Activate
Application.ScreenUpdating = False
wb(1).SaveAs Filename:="H:\ORT Projects\Rate Engine Rewrite\VBA Macros\Data_Processing\File3.xlsx"

End Sub





