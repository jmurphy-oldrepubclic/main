Attribute VB_Name = "sp_PostData_P5"
Public Sub p_PostData_P5()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate

Call p_RequestData501
Call p_RequestData502
Call p_RequestData503
Call p_RequestData504
Call p_RequestData505
Call p_RequestData506
Call p_RequestData507
Call p_RequestData508
Call p_RequestData509
Call p_RequestData510

wb(1).Save
End Sub





 
  
