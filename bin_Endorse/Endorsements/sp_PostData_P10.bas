Attribute VB_Name = "sp_PostData_P10"
Option Explicit

Public Sub p_PostData_P10()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate
Call p_RequestData1001
Call p_RequestData1002
Call p_RequestData1003
Call p_RequestData1004
Call p_RequestData1005
Call p_RequestData1006
Call p_RequestData1007
Call p_RequestData1008
Call p_RequestData1009
Call p_RequestData1010

wb(1).Save
End Sub


