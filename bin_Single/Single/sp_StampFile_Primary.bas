Attribute VB_Name = "sp_StampFile_Primary"
Public Sub p_StampFile_Primary()

On Error Resume Next
Call p_StampFile_P1
On Error Resume Next
Call p_StampFile_P2
On Error Resume Next
Call p_StampFile_P3
On Error Resume Next
Call p_StampFile_P4
On Error Resume Next
Call p_StampFile_P5

End Sub

