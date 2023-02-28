Attribute VB_Name = "sp_PostData_P07"
Option Explicit

Public Sub p_PostData_P07()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate
Call p_RequestData701
Call p_RequestData702
Call p_RequestData703
Call p_RequestData704
Call p_RequestData705
Call p_RequestData706
Call p_RequestData707
Call p_RequestData708
Call p_RequestData709
Call p_RequestData710

wb(1).Save
End Sub





 
