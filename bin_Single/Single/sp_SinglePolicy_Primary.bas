Attribute VB_Name = "sp_SinglePolicy_Primary"
Sub p_SinglePolicy_Primary()

'Disable  alerts


'Call sub procedures
On Error Resume Next
Call p_SinglePolicy_P1
On Error Resume Next
Call p_SinglePolicy_P2
On Error Resume Next
Call p_SinglePolicy_P3
On Error Resume Next
Call p_SinglePolicy_P4
On Error Resume Next
Call p_SinglePolicy_P5
End Sub

