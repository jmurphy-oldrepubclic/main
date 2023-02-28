Attribute VB_Name = "sp_PostData_P09"
Option Explicit

Public Sub p_PostData_P09()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate
Call p_RequestData901
Call p_RequestData902
Call p_RequestData903
Call p_RequestData904
Call p_RequestData905
Call p_RequestData906
Call p_RequestData907
Call p_RequestData908
Call p_RequestData909
Call p_RequestData910

wb(1).Save
End Sub

