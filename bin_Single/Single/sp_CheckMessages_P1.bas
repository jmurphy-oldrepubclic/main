Attribute VB_Name = "sp_CheckMessages_P1"
Public Sub p_CheckMessages_P1()
Dim wb(1 To 5) As Workbook

Set wb(1) = Workbooks("ResultsSingle.xlsx")
wb(1).Activate


Set wb(2) = Workbooks("Datadump.xlsx")
wb(2).Activate

Dim a1 As Range
With wb(2).Worksheets("Response1")
For Each a1 In Range("A15")
        If InStr(a1.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("N3").Value = "Check Message"
    End If
    Next a1

    
    
Dim b1 As Range
For Each b1 In Range("A15")
     If InStr(b1.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("R3") = b1.Text
    End If
Next b1



    
    
Dim a2 As Range
For Each a2 In Range("A16")
        If InStr(a2.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("N4").Value = "Check Message"
    End If
    Next a2

    
    
Dim b2 As Range
For Each b2 In Range("A16")
     If InStr(b2.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("R4") = b2.Text
    End If
Next b2


Dim a3 As Range
For Each a3 In Range("A17")
        If InStr(a3.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("N5").Value = "Check Message"
    End If
    Next a3
    
    
Dim b3 As Range
For Each b3 In Range("A17")
     If InStr(b3.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("R5") = b3.Text
    End If
Next b3



Dim a4 As Range
For Each a4 In Range("A18")
        If InStr(a4.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("N6").Value = "Check Message"
    End If
    Next a4
    
    
Dim b4 As Range
For Each b4 In Range("A18")
     If InStr(b4.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("R6") = b4.Text
    End If
Next b4


Dim a5 As Range
For Each a5 In Range("A19")
        If InStr(a5.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("N7").Value = "Check Message"
    End If
    Next a5
    
    
Dim b5 As Range
For Each b5 In Range("A19")
     If InStr(b5.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("R7") = b5.Text
    End If
Next b5


Dim a6 As Range
For Each a6 In Range("A20")
        If InStr(a6.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("N8").Value = "Check Message"
    End If
    Next a6
    
    
Dim b6 As Range
For Each b6 In Range("A20")
     If InStr(b6.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("R8") = b6.Text
    End If
Next b6
    

Dim a7 As Range
For Each a7 In Range("A21")
        If InStr(a7.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("N9").Value = "Check Message"
    End If
    Next a7
    
    
Dim b7 As Range
For Each b7 In Range("A21")
     If InStr(b7.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("R9") = b7.Text
    End If
Next b7

Dim a8 As Range
For Each a8 In Range("A22")
        If InStr(a8.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("N10").Value = "Check Message"
    End If
    Next a8

    
Dim b8 As Range
For Each b8 In Range("A22")
     If InStr(b8.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("R10") = b8.Text
    End If
Next b8


Dim a9 As Range
For Each a9 In Range("A23")
        If InStr(a9.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("N11").Value = "Check Message"
    End If
    Next a9
    
    
Dim b9 As Range
For Each b9 In Range("A23")
     If InStr(b9.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("R11") = b9.Text
    End If
Next b9

Dim a10 As Range
For Each a10 In Range("A24")
        If InStr(a10.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("N12").Value = "Check Message"
    End If
    Next a10
    
    
Dim b10 As Range
For Each b10 In Range("A24")
     If InStr(b10.Value, "Message") > 0 Then
        wb(1).Worksheets(1).Range("R12") = b10.Text
    End If
Next b10
End With

wb(1).Worksheets(1).Range("N3").EntireColumn.AutoFit

End Sub
