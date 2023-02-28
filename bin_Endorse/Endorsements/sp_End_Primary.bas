Attribute VB_Name = "sp_End_Primary"
Sub p_End_Primary()

'Disable  alerts
With Application
.DisplayAlerts = False
.ScreenUpdating = False
End With

'Call sub procedures
On Error Resume Next
Call p_End_P07
On Error Resume Next
Call p_End_P08
On Error Resume Next
Call p_End_P09
On Error Resume Next
Call p_End_P10
On Error Resume Next
Call p_End_P11
End Sub

