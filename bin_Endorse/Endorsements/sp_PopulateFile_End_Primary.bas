Attribute VB_Name = "sp_PopulateFile_End_Primary"
Sub p_PopulateFile_End_Primary()

'Disable  alerts
With Application
.DisplayAlerts = False
.ScreenUpdating = False
End With

'Call sub procedures
Call p_PopulateFile_P07
Call p_PopulateFile_P08
Call p_PopulateFile_P09
Call p_PopulateFile_P10
Call p_PopulateFile_P11
End Sub


        
