Attribute VB_Name = "sp_RequestData1101"
Public Sub p_RequestData1101()

Dim wb(1 To 5) As Workbook
Dim ws As Worksheet
Dim env As String
Dim RegURL As String
Dim Endpoint As String
Dim newBody_Owners As String
Dim newBody_Loan As String
Dim allBody_Owner As String
Dim allBody_Loan As String
Dim response_Owner As String
Dim response_Loan As String
Dim objHTTP As Object
Set request = CreateObject("WinHttp.WinHttpRequest.5.1")

Set wb(1) = ActiveWorkbook
With wb(1).Sheets("Response11")
newBody_End_Owners = .Range("A1").Value
allBody_End_Owners = newBody_End_Owners
newBody_End_Loan = .Range("A11").Value
allBody_End_Loan = newBody_End_Loan
End With

RegURL = "https://quaratesws.oldrepublictitle.com/Calculator/CalculateOrder"
request.Open "POST", RegURL, False
request.setRequestHeader "Content-Type", "application/json; charset=UTF-8"

request.send allBody_End_Owners
response_End_Owners = request.responseText
request.send allBody_End_Loan
response_End_Loan = request.responseText

wb(1).Sheets("Response11").Range("A30:A50").Clear
'MsgBox request.responseText
wb(1).Sheets("Response11").Range("A30").Value = response_End_Owners
wb(1).Sheets("Response11").Range("A40").Value = response_End_Loan

Dim y1, y2, y3, y4 As Range

With wb(1).Sheets("Response11")
For Each y1 In .Range("A30")
     If InStr(y1.Value, "Endorsements") > 0 Then
        y1.Value = Right(y1.Value, InStr(y1.Value, "Endorsements") + 1)
    End If
Next y1

For Each y2 In .Range("A40")
     If InStr(y2.Value, "Endorsements") > 0 Then
        y2.Value = Right(y2.Value, InStr(y2.Value, "Endorsements") + 15)
    End If
Next y2

For Each y3 In .Range("A30")
     If InStr(y3.Value, "CalculatedNationalPremium") > 0 Then
        y3.Value = Left(y3.Value, InStr(y3.Value, "CalculatedNationalPremium") + 1)
    End If
Next y3

For Each y4 In .Range("A40")
     If InStr(y4.Value, "CalculatedNationalPremium") > 0 Then
        y4.Value = Left(y4.Value, InStr(y4.Value, "CalculatedNationalPremium") + 15)
    End If
Next y4
End With
End Sub








