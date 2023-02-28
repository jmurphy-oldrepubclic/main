Attribute VB_Name = "sp_CreateTabs_P5"
Public Sub p_CreateTabs_P5()

Dim wb(1 To 3) As Workbook
Dim sourceDoc As Excel.Range
Dim sourceRng(1 To 20) As Excel.Range

Set wb(1) = Workbooks("SourceData.xlsx")

Set wb(2) = Workbooks("ResultsSingle")
wb(2).Activate

'
Set sourceDoc = wb(1).Sheets("Single Policy Inputs").Range("E6")
For Each sourceRng(1) In sourceDoc
    If Len(sourceRng(1).Text) > 0 Then
    wb(1).Activate
    wb(2).Sheets.Add().Name = sourceRng(1)
End If
Next sourceRng(1)

End Sub
