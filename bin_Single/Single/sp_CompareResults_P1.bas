Attribute VB_Name = "sp_CompareResults_P1"
Public Sub p_CompareResults_P1()

Call p_TrimResults_P1


On Error Resume Next
With Worksheets(1)
LastRow = .Range("N" & Rows.Count).End(xlUp).Row
For i = 3 To LastRow
    If .Range("G" & i).Value = .Range("N" & i).Value Then
    .Range("P" & i).Value = "Pass"
Else
   .Range("P" & i).Value = "Fail"
End If
Next i
'
On Error Resume Next
LastRow = .Range("N" & Rows.Count).End(xlUp).Row
For i = 3 To LastRow
    If .Range("P" & i).Value = "Pass" Then
    .Range("P" & i).Interior.Color = vbGreen
Else
    .Range("P" & i).Interior.Color = vbRed
    End If
Next i
End With

End Sub
