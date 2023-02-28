Attribute VB_Name = "sp_CreateTabs_P08"
Public Sub p_CreateTabs_P08()

Dim wb(1 To 3) As Workbook
Dim sourceDoc As Excel.Range
Dim sourceRng(1 To 20) As Excel.Range

Set wb(1) = Workbooks("SourceData.xlsx")

Set wb(2) = Workbooks("ResultsEndorsement")
wb(2).Activate

Set sourceDoc = wb(1).Sheets("Policy with Endor Inputs").Range("E4")
For Each sourceRng(1) In sourceDoc
    If Len(sourceRng(1).Text) > 0 Then
    wb(2).Activate
    wb(2).Sheets.Add().Name = sourceRng(1)
End If
Next sourceRng(1)

End Sub



