Attribute VB_Name = "sp_PostData_P6"
Option Explicit

Public Sub p_PostData_P6()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate
Call p_RequestData601
Call p_RequestData602
Call p_RequestData603
Call p_RequestData604
Call p_RequestData605
Call p_RequestData606
Call p_RequestData607
Call p_RequestData608
Call p_RequestData609
Call p_RequestData610

wb(1).Save



End Sub





 
