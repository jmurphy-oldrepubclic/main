Attribute VB_Name = "sp_PostData_P2"
Public Sub p_PostData_P2()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate


Call p_RequestData201
Call p_RequestData202
Call p_RequestData203
Call p_RequestData204
Call p_RequestData205
Call p_RequestData206
Call p_RequestData207
Call p_RequestData208
Call p_RequestData209
Call p_RequestData210

wb(1).Save
End Sub





 
  
