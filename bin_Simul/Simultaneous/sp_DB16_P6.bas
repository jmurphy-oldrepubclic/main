Attribute VB_Name = "sp_DB16_P6"
Public Sub p_DB16_P6()

Dim wb(1 To 3) As Workbook
Dim sourceDoc As Excel.Range
Dim sourceRng(1 To 20) As Excel.Range
Dim adoDbConn(1 To 20) As New ADODB.Connection
Dim adoDbRs(1 To 10) As New ADODB.Recordset

Set wb(1) = Workbooks("ResultsSimult")
wb(1).Activate

Set wb(2) = Workbooks("SourceData.xlsx")
wb(2).Activate

adoDbConn(1).Open "Provider=SQLOLEDB;Data Source=mn-qua-db16;Initial Catalog=RatesEngineTest_vNext;Trusted_connection=yes;"

With wb(2).Worksheets("Simultanious Policy Inputs")
s = .Range("C3").Value
c = .Range("D3").Value
tr1 = .Range("F3").Value
tr2 = .Range("G3").Value
ad = .Range("H3").Value
ll1 = .Range("I3").Value
ll2 = .Range("L3").Value
ul1 = .Range("J3").Value
ul2 = .Range("M3").Value
pl1 = .Range("K3").Value
pl2 = .Range("N3").Value
t = .Range("O3").Value

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(1) = adoDbConn(1).Execute("DECLARE @s VARCHAR(2)  = '" & s & "', @ed DATETIME = '" & ad & "', @otr VARCHAR(10) = '" & tr1 & "', @ltr VARCHAR(10) = '" & tr2 & "', @cc VARCHAR(4) = '%" & c & "%', @ll1 Decimal(18,2) = '" & ll1 & "', @ul1 Decimal(18,2) = '" & ul1 & "', @ll2 Decimal(18,2) = '" & ll2 & "', @ul2 Decimal(18,2) = '" & ul2 & "', @cl Decimal(18,2) = '" & pl1 & "', @n VARCHAR(50) = '%" & t & "%' SELECT top 10 o.OrderNumber, p.TranCode, p.EffectiveDate, p.Liability, p.CreditLiability, ta.Name FROM Tests te JOIN TestTags tt ON tt.Test_Id = te.Id JOIN Tags ta ON ta.Id = tt.Tag_Id JOIN OrderTags ot ON ot.Tag_Id = ta.Id JOIN Orders o ON o.Id = ot.Order_Id JOIN Policies p ON p.OrderId = o.Id" & _
 " WHERE o.Id IN (SELECT o.Id FROM Orders o JOIN Policies p ON p.OrderId = o.Id WHERE o.StateCode = @s AND p.TranCode = @otr AND p.EffectiveDate >= @ed and p.Liability >= @ll1 and p.Liability >= @ll2 and p.Liability <= @ul1 and p.Liability <= @ul2 and p.CreditLiability >= @cl) AND o.Id IN (SELECT o.ID FROM Orders o JOIN Policies p ON p.OrderId = o.Id WHERE o.StateCode = @s AND p.TranCode = @ltr AND p.EffectiveDate >= @ed and p.Liability >= @ll1 and p.Liability >= @ll2 and p.Liability <= @ul1 and p.Liability <= @ul2 and p.CreditLiability >= @cl) ORDER BY o.Id, p.TranCode")
  
  If Not adoDbRs(1).EOF Then
  'Transfer result.
    On Error Resume Next
    wb(1).Worksheets(1).Range("B3:F3").CopyFromRecordset adoDbRs(1)
  'Close the recordset
  adoDbRs(1).Close
End If

s = .Range("C3").Value
c = .Range("D3").Value
tr1 = .Range("F3").Value
ad = .Range("H3").Value
ll1 = .Range("L3").Value
ul1 = .Range("J3").Value
pl1 = .Range("K3").Value
t = .Range("O3").Value
tr2 = .Range("G3").Value
ll2 = .Range("L3").Value
ul2 = .Range("M3").Value
pl2 = .Range("N3").Value

'RateEngineTest_vNext
Set adoDbRs(2) = adoDbConn(1).Execute("DECLARE @s VARCHAR(2)  = '" & s & "', @ed DATETIME = '" & ad & "', @otr VARCHAR(10) = '" & tr1 & "', @ltr VARCHAR(10) = '" & tr2 & "', @cc VARCHAR(4) = ('%" & c & "%'), @ll1 Decimal(18,2) = '" & ll1 & "', @ul1 Decimal(18,2) = '" & ul1 & "', @ll2 Decimal(18,2) = '" & ll2 & "', @ul2 Decimal(18,2) = '" & ul2 & "', @cl Decimal(18,2) = '" & pl1 & "'  , @n VARCHAR(50) = '%" & t & "%' SELECT top 10 pr.CalculatedGrossPremium FROM Tests te JOIN TestTags tt ON tt.Test_Id = te.Id JOIN Tags ta ON ta.Id = tt.Tag_Id JOIN OrderTags ot ON ot.Tag_Id = ta.Id JOIN Orders o ON o.Id = ot.Order_Id JOIN Policies p ON p.OrderId = o.Id JOIN PolicyResults pr on p.Id = pr.PolicyId" _
& " WHERE o.Id IN (SELECT o.ID FROM Orders o JOIN Policies p ON p.OrderId = o.Id WHERE o.StateCode = @s AND p.TranCode = @otr AND p.EffectiveDate >= @ed and p.Liability >= @ll1 and p.Liability >= @ll2 and p.Liability <= @ul1 and p.Liability <= @ul2 and p.CreditLiability >= @cl) AND o.Id IN ( SELECT o.ID FROM Orders o JOIN Policies p ON p.OrderId = o.Id WHERE o.StateCode = @s AND p.TranCode = @ltr AND p.EffectiveDate >= @ed and p.Liability >= @ll1 and p.Liability >= @ll2 and p.Liability <= @ul1 and p.Liability <= @ul2 and p.CreditLiability >= @cl) ORDER BY o.Id, p.TranCode")


  If Not adoDbRs(2).EOF Then
  'Transfer result.
    On Error Resume Next
    wb(1).Worksheets(1).Range("G3").CopyFromRecordset adoDbRs(2)
  'Close the recordset
  adoDbRs(2).Close
End If

s = .Range("C3").Value
c = .Range("D3").Value
tr1 = .Range("F3").Value
ad = .Range("H3").Value
ll1 = .Range("L3").Value
ul1 = .Range("J3").Value
pl1 = .Range("K3").Value
t = .Range("O3").Value
tr2 = .Range("G3").Value
ll2 = .Range("L3").Value
ul2 = .Range("M3").Value
pl2 = .Range("N3").Value

'RateEngineTest_vNext
Set adoDbRs(3) = adoDbConn(1).Execute("DECLARE @s VARCHAR(2)  = '" & s & "', @ed DATETIME = '" & ad & "', @otr VARCHAR(10) = '" & tr1 & "', @ltr VARCHAR(10) = '" & tr2 & "', @cc VARCHAR(4) = ('%" & c & "%'), @ll1 Decimal(18,2) = '" & ll1 & "', @ul1 Decimal(18,2) = '" & ul1 & "', @ll2 Decimal(18,2) = '" & ll2 & "', @ul2 Decimal(18,2) = '" & ul2 & "', @cl Decimal(18,2) = '" & pl1 & "'  , @n VARCHAR(50) = '%" & t & "%' SELECT top 10 o.OrderNumber, p.TranCode, p.EffectiveDate, p.Liability, p.CreditLiability, ta.Name FROM Tests te JOIN TestTags tt ON tt.Test_Id = te.Id JOIN Tags ta ON ta.Id = tt.Tag_Id JOIN OrderTags ot ON ot.Tag_Id = ta.Id JOIN Orders o ON o.Id = ot.Order_Id JOIN Policies p ON p.OrderId = o.Id" _
& " WHERE o.Id IN (SELECT o.ID FROM Orders o JOIN Policies p ON p.OrderId = o.Id WHERE o.StateCode = @s AND p.TranCode = @otr AND p.EffectiveDate >= @ed and p.Liability >= @ll1 and p.Liability >= @ll2 and p.Liability <= @ul1 and p.Liability <= @ul2 and p.CreditLiability >= @cl) AND o.Id IN ( SELECT o.ID FROM Orders o JOIN Policies p ON p.OrderId = o.Id WHERE o.StateCode = @s AND p.TranCode = @ltr AND p.EffectiveDate >= @ed and p.Liability >= @ll1 and p.Liability >= @ll2 and p.Liability <= @ul1 and p.Liability <= @ul2 and p.CreditLiability >= @cl) ORDER BY o.Id, p.TranCode")
  
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






