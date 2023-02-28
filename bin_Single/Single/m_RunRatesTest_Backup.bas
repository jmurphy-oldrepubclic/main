Attribute VB_Name = "m_RunRatesTest_Backup"
Sub RunRatesTestBackup()

Dim st(1 To 50) As Worksheet
Dim wb(1 To 3) As Workbook
Dim LastRow As Long
Dim endRow As Long
Dim sourceDoc As Excel.Range
Dim rng(1 To 50) As Excel.Range
Dim sourceRng(1 To 20) As Excel.Range
Dim sdCount(1 To 50) As Long

'Disable  alerts
Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Call sub procedures
Call p_SinglePolicy_Primary
Call p_ConvertData_Primary
Call p_PostData_Primary
Call p_CreateTabs
Call p_PopulateFile_Primary
Call p_StampFile_Primary


Set wb(1) = ActiveWorkbook
wb(1).Activate
'wb(1).Save

End Sub
