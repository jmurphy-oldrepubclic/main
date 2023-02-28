Attribute VB_Name = "sp_PostData_P11"
Option Explicit

Public Sub p_PostData_P11()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate
Call p_RequestData1101
Call p_RequestData1102
Call p_RequestData1103
Call p_RequestData1104
Call p_RequestData1105
Call p_RequestData1106
Call p_RequestData1107
Call p_RequestData1108
Call p_RequestData1109
Call p_RequestData1110

wb(1).Save
End Sub
