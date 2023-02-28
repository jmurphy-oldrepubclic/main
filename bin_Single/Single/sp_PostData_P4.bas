Attribute VB_Name = "sp_PostData_P4"
Public Sub p_PostData_P4()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate

Set wb(1) = ActiveWorkbook
Call p_RequestData401
Call p_RequestData402
Call p_RequestData403
Call p_RequestData404
Call p_RequestData405
Call p_RequestData406
Call p_RequestData407
Call p_RequestData408
Call p_RequestData409
Call p_RequestData410

wb(1).Save
End Sub





 
  
