Attribute VB_Name = "sp_RequestData610"
Public Sub p_RequestData610()

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
With wb(1).Sheets("Response6")
newBody_Simult = .Range("A10").Value
allBody_Simult = newBody_Simult
End With

RegURL = "https://quaratesws.oldrepublictitle.com/Calculator/CalculateOrder"
request.Open "POST", RegURL, False
request.setRequestHeader "Content-Type", "application/json; charset=UTF-8"

request.send allBody_Simult
response_Simult = request.responseText

Dim countTxt As String
Dim myTxt As String
Dim TranCodeCount As Long

'What do you want to count?
    countTxt = "},{"

'What do you want to analyze?
  myTxt = wb(1).Sheets("Response6").Range("A10").Value

'Count how many occurrences there are
  TranCodeCount = (Len(myTxt) - Len(Replace(myTxt, countTxt, ""))) / Len(countTxt)
  
'Report out results
  If TranCodeCount <> 1 Then
   Exit Sub
End If

'MsgBox request.responseText
wb(1).Sheets("Response6").Range("A48:A49").Value = response_Simult

Dim a1 As Range

With wb(1).Sheets("Response6")
For Each a1 In .Range("A48")
     If InStr(a1.Value, "Endorsements") > 0 Then
        a1.Value = Right(a1.Value, InStr(a1.Value, "Endorsements") + 190)
    End If
Next a1


'MsgBox request.responseText
wb(1).Sheets("Response6").Range("A48:A49").Value = response_Simult




For Each y1 In .Range("A48")
     If InStr(y1.Value, "Endorsements") > 0 Then
        y1.Value = Left(y1.Value, InStr(y1.Value, "Endorsements") - 15)
    End If
Next y1


Dim y2 As Range
Dim a2 As Range

For Each a2 In .Range("A49")
     If InStr(a2.Value, "PropertyTax") > 0 Then
        a2.Value = Right(a2.Value, InStr(a2.Value, "PropertyTax") + 143)
    End If
Next a2

For Each y2 In .Range("A49")
     If InStr(y2.Value, "Endorsements") > 0 Then
        y2.Value = Left(y2.Value, InStr(y2.Value, "Endorsements") - 15)
    End If
Next y2
End With
End Sub







