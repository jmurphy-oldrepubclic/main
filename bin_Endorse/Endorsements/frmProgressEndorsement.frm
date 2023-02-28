VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressEndorsement 
   Caption         =   "Calculating Rates"
   ClientHeight    =   825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   OleObjectBlob   =   "frmProgressEndorsement.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgressEndorsement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub UserForm_Activate()

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
Call p_End_Primary
Call p_ConvertData_End_Primary
Call p_PostData_End_Primary
Call p_PopulateFile_End_Primary

Set wb(1) = Workbooks("SourceData.xlsx")
wb(1).Activate
Set wb(2) = Workbooks("ResultsEndorsement.xlsx")
wb(2).Activate
Set wb(3) = Workbooks("Datadump.xlsx")
wb(3).Activate





Path = "H:\ORT Projects\Rate Testing\Results\QUA\"
NewfName = wb(1).Worksheets("Single Policy Inputs").Range("M6")
wb(2).SaveAs Filename:=Path & NewfName & ".xlsx", FileFormat:=51, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges

wb(3).Close

For i = 1 To 100
frmProgressEndorsement.lblProgress50.Caption = str(i) + "% Completed"
frmProgressEndorsement.lblProgress50.Width = i * 2
Next i


End Sub
