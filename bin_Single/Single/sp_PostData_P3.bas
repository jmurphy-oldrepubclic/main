Attribute VB_Name = "sp_PostData_P3"
Public Sub p_PostData_P3()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate

Call p_RequestData301
Call p_RequestData302
Call p_RequestData303
Call p_RequestData304
Call p_RequestData305
Call p_RequestData306
Call p_RequestData307
Call p_RequestData308
Call p_RequestData309
Call p_RequestData310

wb(1).Save
End Sub





 
  
