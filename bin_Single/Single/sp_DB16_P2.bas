Attribute VB_Name = "sp_DB16_P2"
Public Sub p_DB16_P2()

Dim wb(1 To 3) As Workbook
Dim sourceDoc As Excel.Range
Dim sourceRng(1 To 20) As Excel.Range
Dim adoDbConn(1 To 20) As New ADODB.Connection
Dim adoDbRs(1 To 10) As New ADODB.Recordset

Set wb(1) = Workbooks("ResultsSingle")
wb(1).Activate

Workbooks("SourceData.xlsx").Activate


adoDbConn(1).Open "Provider=SQLOLEDB;Data Source=mn-qua-db16;Initial Catalog=RatesEngineTest_vNext;Trusted_connection=yes;"

With Workbooks("SourceData.xlsx").Worksheets("Single Policy Inputs")
s = .Range("C3").Value
c = .Range("D3").Value
tr = .Range("F3").Value
ad = .Range("G3").Value
ll = .Range("H3").Value
ul = .Range("I3").Value
pl = .Range("J3").Value
t = .Range("K3").Value

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(1) = adoDbConn(1).Execute("Declare @n VARCHAR(50) Declare @l Decimal(18,2) Declare @cl Decimal(18,2) Declare @ed Date Declare @tr VARCHAR(50) Declare @cc VARCHAR(4) select top 10 o.OrderNumber, p.Trancode, p.EffectiveDate, p.Liability, p.CreditLiability from Orders o inner join Policies p on p.OrderId = o.Id inner join OrderTags ot on o.Id = ot.Order_Id inner join tags t on ot.Tag_Id = t.Id where o.StateCode = '" & s & "' and o.CountyCode like isnull(@cc, '%" & c & "%') and p.TranCode like isnull(@tr,'%" & tr & "%') and p.EffectiveDate >= isnull(@ed,'" & ad & "') and p.Liability >= isnull(@l,'" & ll & "') and p.Liability <= isnull(@l, '" & ul & "') and p.CreditLiability >= isnull(@cl,'" & pl & "') and t.Name like isnull(@n, '%" & t & "%') and o.OrderNumber in (SELECT OrderNumber FROM Orders GROUP BY OrderNumber HAVING COUNT(OrderNumber)=1) Order By o.OrderNumber asc")
  
  If Not adoDbRs(1).EOF Then
  'Transfer result.
    On Error Resume Next
    wb(1).Worksheets(1).Range("B3:F3").CopyFromRecordset adoDbRs(1)
  'Close the recordset
  adoDbRs(1).Close
End If

s = .Range("C3").Value
c = .Range("D3").Value
tr = .Range("F3").Value
ad = .Range("G3").Value
ll = .Range("H3").Value
ul = .Range("I3").Value
pl = .Range("J3").Value
t = .Range("K3").Value

'RateEngineTest_vNext
Set adoDbRs(2) = adoDbConn(1).Execute("Declare @n VARCHAR(50) Declare @l Decimal(18,2) Declare @cl Decimal(18,2) Declare @ed Date Declare @tr VARCHAR(50) Declare @cc VARCHAR(4) select top 10 pr.CalculatedGrossPremium from OrderResults [or] inner join Orders o on [or].OrderId = o.Id inner join Policies p on p.OrderId = o.Id inner join PolicyResults pr on pr.OrderResultOrderId = [or].OrderId and p.TranCode = pr.TranCode inner join OrderTags ot on o.Id = ot.Order_Id inner join tags t on ot.Tag_Id = t.Id where o.StateCode = '" & s & "' and o.CountyCode like isnull(@cc, '%" & c & "%') and p.TranCode like isnull(@tr,'%" & tr & "%') and p.EffectiveDate >= isnull(@ed,'" & ad & "') and p.Liability >= isnull(@l,'" & ll & "') and p.Liability <= isnull(@l, '" & ul & "') and p.CreditLiability >= isnull(@cl,'" & pl & "') and t.Name like isnull(@n, '%" & t & "%') and o.OrderNumber in (SELECT OrderNumber FROM Orders GROUP BY OrderNumber HAVING COUNT(OrderNumber)=1) Order By o.OrderNumber asc")
  
  If Not adoDbRs(2).EOF Then
  'Transfer result.
    On Error Resume Next
    wb(1).Worksheets(1).Range("G3").CopyFromRecordset adoDbRs(2)
  'Close the recordset
  adoDbRs(2).Close
End If


s = .Range("C3").Value
c = .Range("D3").Value
tr = .Range("F3").Value
ad = .Range("G3").Value
ll = .Range("H3").Value
ul = .Range("I3").Value
pl = .Range("J3").Value
t = .Range("K3").Value

'RateEngineTest_vNext
Set adoDbRs(3) = adoDbConn(1).Execute("Declare @n VARCHAR(50)Declare @l Decimal(18,2) Declare @cl Decimal(18,2) Declare @ed Date Declare @tr VARCHAR(50) Declare @cc VARCHAR(4) select top 10 o.OrderNumber, p.Trancode, p.EffectiveDate, p.Liability, p.CreditLiability from Orders o inner join Policies p on p.OrderId = o.Id inner join OrderTags ot on o.Id = ot.Order_Id inner join tags t on ot.Tag_Id = t.Id where o.StateCode = '" & s & "' and o.CountyCode like isnull(@cc, '%" & c & "%') and p.TranCode like isnull(@tr,'%" & tr & "%') and p.EffectiveDate >= isnull(@ed,'" & ad & "') and p.Liability >= isnull(@l,'" & ll & "') and p.Liability <= isnull(@l, '" & ul & "') and p.CreditLiability >= isnull(@cl,'" & pl & "') and t.Name like isnull(@n, '%" & t & "%') and o.OrderNumber in (SELECT OrderNumber FROM Orders GROUP BY OrderNumber HAVING COUNT(OrderNumber)=1) Order By o.OrderNumber asc")
  
  If Not adoDbRs(3).EOF Then
  'Transfer result.
    On Error Resume Next
    wb(1).Worksheets(1).Range("I3:M3").CopyFromRecordset adoDbRs(3)
  'Close the recordset
  adoDbRs(3).Close
End If

adoDbConn(1).Close

End With
End Sub


