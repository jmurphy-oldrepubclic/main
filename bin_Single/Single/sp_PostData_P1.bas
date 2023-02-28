Attribute VB_Name = "sp_PostData_P1"
Public Sub p_PostData_P1()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate

Call p_RequestData101
Call p_RequestData102
Call p_RequestData103
Call p_RequestData104
Call p_RequestData105
Call p_RequestData106
Call p_RequestData107
Call p_RequestData108
Call p_RequestData109
Call p_RequestData110

wb(1).Save

End Sub





 
  


