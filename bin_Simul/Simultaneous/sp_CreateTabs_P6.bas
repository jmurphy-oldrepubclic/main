Attribute VB_Name = "sp_CreateTabs_P6"
Public Sub p_CreateTabs_P6()

Dim wb(1 To 3) As Workbook
Dim sourceDoc As Excel.Range
Dim sourceRng(1 To 20) As Excel.Range

Set wb(1) = Workbooks("SourceData.xlsx")

Dim Path As String
Dim fName As String

Path = "H:\ORT Projects\Rate Engine Rewrite\Results\QA\"
fName = "ResultsSimult"

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Set wb(2) = Workbooks.Add

wb(2).SaveAs Filename:=Path & fName & ".xlsx", FileFormat:=51

Set sourceDoc = wb(1).Sheets("Simultanious Policy Inputs").Range("E3")
For Each sourceRng(1) In sourceDoc
    If Len(sourceRng(1).Text) > 0 Then
    wb(2).Activate
    wb(2).Sheets.Add().Name = sourceRng(1)
End If
Next sourceRng(1)

 On Error Resume Next
wb(2).Sheets("Sheet1").Delete

End Sub


