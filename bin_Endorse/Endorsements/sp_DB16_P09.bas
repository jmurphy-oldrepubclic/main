Attribute VB_Name = "sp_DB16_P09"
Public Sub p_DB16_P09()

Dim wb(1 To 3) As Workbook
Dim sourceDoc As Excel.Range
Dim sourceRng(1 To 20) As Excel.Range
Dim adoDbConn(1 To 20) As New ADODB.Connection
Dim adoDbRs(1 To 10) As New ADODB.Recordset

Set wb(1) = Workbooks("ResultsEndorsement")
wb(1).Activate

Set wb(2) = Workbooks("SourceData.xlsx")
wb(2).Activate

adoDbConn(1).Open "Provider=SQLOLEDB;Data Source=mn-qua-db16;Initial Catalog=RatesEngineTest_vNext;Trusted_connection=yes;"

With wb(2).Sheets("Policy with Endor Inputs")
s = .Range("C3").Value
c = .Range("D3").Value
tr1 = .Range("F3").Value
tr2 = .Range("G3").Value
ad = .Range("H3").Value
ll1 = .Range("I3").Value
ul1 = .Range("J3").Value
ll2 = .Range("K3").Value
ul2 = .Range("L3").Value
pl1 = .Range("M3").Value
t = .Range("P3").Value
e1 = .Range("N5").Value

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(1) = adoDbConn(1).Execute("DECLARE @s VARCHAR(2)  = '" & s & "', @ed DATETIME = '" & ad & "', @otr VARCHAR(10) = '" & tr1 & "', @oec VARCHAR(10) = '" & e1 & "', @cc VARCHAR(4) = '%" & c & "%', @ll1 Decimal(18,2) = '" & ll1 & "', @ul1 Decimal(18,2) = '" & ul1 & "', @cl Decimal(18,2) = '" & pl1 & "', @n VARCHAR(50) = '%" & t & "%' SELECT TOP 10 o.OrderNumber, p.TranCode, e.Code, p.EffectiveDate, p.Liability, p.CreditLiability FROM Tests te JOIN TestTags tt ON tt.Test_Id = te.Id JOIN Tags ta ON ta.Id = tt.Tag_Id JOIN OrderTags ot ON ot.Tag_Id = ta.Id JOIN Orders o ON o.Id = ot.Order_Id JOIN Policies p ON p.OrderId = o.Id JOIN Endorsements e on p.Id = e.PolicyId JOIN EndorsementResults er on e.id = er.EndorsementId WHERE o.StateCode = @s AND p.TranCode = @otr AND e.Code = @oec AND p.EffectiveDate >= @ed and p.Liability >= @ll1 and p.Liability <= @ul1 and p.CreditLiability >= @cl Order By P.TranCode asc")
  
  If Not adoDbRs(1).EOF Then
  'Transfer result.
    On Error Resume Next
    wb(1).Worksheets(1).Range("B3").CopyFromRecordset adoDbRs(1)
  'Close the recordset
  adoDbRs(1).Close
End If

s = .Range("C3").Value
c = .Range("D3").Value
tr1 = .Range("F3").Value
tr2 = .Range("G3").Value
ad = .Range("H3").Value
ll1 = .Range("I3").Value
ul1 = .Range("J3").Value
ll2 = .Range("K3").Value
ul2 = .Range("L3").Value
pl1 = .Range("M3").Value
t = .Range("P3").Value
e2 = .Range("O5").Value

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(2) = adoDbConn(1).Execute("DECLARE @s VARCHAR(2)  = '" & s & "', @ed DATETIME = '" & ad & "', @ltr VARCHAR(10) = '" & tr2 & "', @lec VARCHAR(10) = '" & e2 & "', @cc VARCHAR(4) = '%" & c & "%', @ll2 Decimal(18,2) = '" & ll2 & "', @ul2 Decimal(18,2) = '" & ul2 & "', @cl Decimal(18,2) = '" & pl1 & "', @n VARCHAR(50) = '%" & t & "%' SELECT TOP 10 O.OrderNumber, p.TranCode, e.Code, p.EffectiveDate, p.Liability, p.CreditLiability FROM Tests te JOIN TestTags tt ON tt.Test_Id = te.Id JOIN Tags ta ON ta.Id = tt.Tag_Id JOIN OrderTags ot ON ot.Tag_Id = ta.Id JOIN Orders o ON o.Id = ot.Order_Id JOIN Policies p ON p.OrderId = o.Id JOIN Endorsements e on p.Id = e.PolicyId JOIN EndorsementResults er on e.id = er.EndorsementId WHERE o.StateCode = @s AND p.TranCode = @ltr AND e.Code = @lec AND p.EffectiveDate >= @ed and p.Liability >= @ll2  and p.Liability <= @ul2  and p.CreditLiability >= @cl order by p.tranCode asc")
  
  If Not adoDbRs(2).EOF Then
  'Transfer result.
    On Error Resume Next
    wb(1).Worksheets(1).Range("B13").CopyFromRecordset adoDbRs(2)
  'Close the recordset
  adoDbRs(2).Close
End If

s = .Range("C3").Value
c = .Range("D3").Value
tr1 = .Range("F3").Value
tr2 = .Range("G3").Value
ad = .Range("H3").Value
ll1 = .Range("I3").Value
ul1 = .Range("J3").Value
ll2 = .Range("K3").Value
ul2 = .Range("L3").Value
pl1 = .Range("M3").Value
t = .Range("P3").Value
e1 = .Range("N5").Value

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(3) = adoDbConn(1).Execute("DECLARE @s VARCHAR(2)  = '" & s & "', @ed DATETIME = '" & ad & "', @otr VARCHAR(10) = '" & tr1 & "', @oec VARCHAR(10) = '" & e1 & "', @cc VARCHAR(4) = '%" & c & "%', @ll1 Decimal(18,2) = '" & ll1 & "', @ul1 Decimal(18,2) = '" & ul1 & "', @cl Decimal(18,2) = '" & pl1 & "', @n VARCHAR(50) = '%" & t & "%' SELECT TOP 10 er.CalculatedGrossPremium FROM Tests te JOIN TestTags tt ON tt.Test_Id = te.Id JOIN Tags ta ON ta.Id = tt.Tag_Id JOIN OrderTags ot ON ot.Tag_Id = ta.Id JOIN Orders o ON o.Id = ot.Order_Id JOIN Policies p ON p.OrderId = o.Id JOIN Endorsements e on p.Id = e.PolicyId JOIN EndorsementResults er on e.id = er.EndorsementId WHERE o.StateCode = @s AND p.TranCode = @otr AND e.Code = @oec AND p.EffectiveDate >= @ed and p.Liability >= @ll1  and p.Liability <= @ul1  and p.CreditLiability >= @cl order by p.tranCode asc")
  
  If Not adoDbRs(3).EOF Then
  'Transfer result.
    On Error Resume Next
    wb(1).Worksheets(1).Range("G3").CopyFromRecordset adoDbRs(3)
  'Close the recordset
  adoDbRs(3).Close
End If

s = .Range("C3").Value
c = .Range("D3").Value
tr1 = .Range("F3").Value
tr2 = .Range("G3").Value
ad = .Range("H3").Value
ll1 = .Range("I3").Value
ul1 = .Range("J3").Value
ll2 = .Range("K3").Value
ul2 = .Range("L3").Value
pl1 = .Range("M3").Value
t = .Range("P3").Value
e2 = .Range("O5").Value

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(4) = adoDbConn(1).Execute("DECLARE @s VARCHAR(2)  = '" & s & "', @ed DATETIME = '" & ad & "', @ltr VARCHAR(10) = '" & tr2 & "', @lec VARCHAR(10) = '" & e2 & "', @cc VARCHAR(4) = '%" & c & "%', @ll2 Decimal(18,2) = '" & ll2 & "', @ul2 Decimal(18,2) = '" & ul2 & "', @cl Decimal(18,2) = '" & pl1 & "', @n VARCHAR(50) = '%" & t & "%' SELECT TOP 10 er.CalculatedGrossPremium FROM Tests te JOIN TestTags tt ON tt.Test_Id = te.Id JOIN Tags ta ON ta.Id = tt.Tag_Id JOIN OrderTags ot ON ot.Tag_Id = ta.Id JOIN Orders o ON o.Id = ot.Order_Id JOIN Policies p ON p.OrderId = o.Id JOIN Endorsements e on p.Id = e.PolicyId JOIN EndorsementResults er on e.id = er.EndorsementId WHERE o.StateCode = @s AND p.TranCode = @ltr AND e.Code = @lec AND p.EffectiveDate >= @ed and p.Liability >= @ll2  and p.Liability <= @ul2  and p.CreditLiability >= @cl order by p.tranCode asc")
  
  If Not adoDbRs(4).EOF Then
  'Transfer result.
    On Error Resume Next
    wb(1).Worksheets(1).Range("G13").CopyFromRecordset adoDbRs(4)
  'Close the recordset
  adoDbRs(4).Close
End If

s = .Range("C3").Value
c = .Range("D3").Value
tr1 = .Range("F3").Value
tr2 = .Range("G3").Value
ad = .Range("H3").Value
ll1 = .Range("I3").Value
ul1 = .Range("J3").Value
ll2 = .Range("K3").Value
ul2 = .Range("L3").Value
pl1 = .Range("M3").Value
t = .Range("P3").Value
e1 = .Range("N5").Value

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(5) = adoDbConn(1).Execute("DECLARE @s VARCHAR(2)  = '" & s & "', @ed DATETIME = '" & ad & "', @otr VARCHAR(10) = '" & tr1 & "', @oec VARCHAR(10) = '" & e1 & "', @cc VARCHAR(4) = '%" & c & "%', @ll1 Decimal(18,2) = '" & ll1 & "', @ul1 Decimal(18,2) = '" & ul1 & "', @cl Decimal(18,2) = '" & pl1 & "', @n VARCHAR(50) = '%" & t & "%' SELECT TOP 10 o.OrderNumber, p.TranCode, e.Code, p.EffectiveDate, p.Liability, p.CreditLiability FROM Tests te JOIN TestTags tt ON tt.Test_Id = te.Id JOIN Tags ta ON ta.Id = tt.Tag_Id JOIN OrderTags ot ON ot.Tag_Id = ta.Id JOIN Orders o ON o.Id = ot.Order_Id JOIN Policies p ON p.OrderId = o.Id JOIN Endorsements e on p.Id = e.PolicyId JOIN EndorsementResults er on e.id = er.EndorsementId WHERE o.StateCode = @s AND p.TranCode = @otr AND e.Code = @oec AND p.EffectiveDate >= @ed and p.Liability >= @ll1 and p.Liability <= @ul1 and p.CreditLiability >= @cl Order By P.TranCode asc")
  
  If Not adoDbRs(5).EOF Then
  'Transfer result.
    On Error Resume Next
    wb(1).Worksheets(1).Range("I3").CopyFromRecordset adoDbRs(5)
  'Close the recordset
  adoDbRs(5).Close
End If


s = .Range("C3").Value
c = .Range("D3").Value
tr1 = .Range("F3").Value
tr2 = .Range("G3").Value
ad = .Range("H3").Value
ll1 = .Range("I3").Value
ul1 = .Range("J3").Value
ll2 = .Range("K3").Value
ul2 = .Range("L3").Value
pl1 = .Range("M3").Value
t = .Range("P3").Value
e2 = .Range("O5").Value

'RateEngineTest_vNext
'adoDbConn(1).CommandTimeout = 120
Set adoDbRs(6) = adoDbConn(1).Execute("DECLARE @s VARCHAR(2)  = '" & s & "', @ed DATETIME = '" & ad & "', @ltr VARCHAR(10) = '" & tr2 & "', @lec VARCHAR(10) = '" & e2 & "', @cc VARCHAR(4) = '%" & c & "%', @ll2 Decimal(18,2) = '" & ll2 & "', @ul2 Decimal(18,2) = '" & ul2 & "', @cl Decimal(18,2) = '" & pl1 & "', @n VARCHAR(50) = '%" & t & "%' SELECT TOP 10 O.OrderNumber, p.TranCode, e.Code, p.EffectiveDate, p.Liability, p.CreditLiability FROM Tests te JOIN TestTags tt ON tt.Test_Id = te.Id JOIN Tags ta ON ta.Id = tt.Tag_Id JOIN OrderTags ot ON ot.Tag_Id = ta.Id JOIN Orders o ON o.Id = ot.Order_Id JOIN Policies p ON p.OrderId = o.Id JOIN Endorsements e on p.Id = e.PolicyId JOIN EndorsementResults er on e.id = er.EndorsementId WHERE o.StateCode = @s AND p.TranCode = @ltr AND e.Code = @lec AND p.EffectiveDate >= @ed and p.Liability >= @ll2  and p.Liability <= @ul2  and p.CreditLiability >= @cl order by p.tranCode asc")
  
  If Not adoDbRs(6).EOF Then
  'Transfer result.
    On Error Resume Next
    wb(1).Worksheets(1).Range("I13").CopyFromRecordset adoDbRs(6)
  'Close the recordset
  adoDbRs(6).Close
End If
adoDbConn(1).Close


End With
End Sub








