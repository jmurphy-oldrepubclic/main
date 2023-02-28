Attribute VB_Name = "sp_PostData_P08"
Option Explicit

Public Sub p_PostData_P08()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate
Call p_RequestData801
Call p_RequestData802
Call p_RequestData803
Call p_RequestData804
Call p_RequestData805
Call p_RequestData806
Call p_RequestData807
Call p_RequestData808
Call p_RequestData809
Call p_RequestData810

wb(1).Save
End Sub
