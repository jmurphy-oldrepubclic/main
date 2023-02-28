Attribute VB_Name = "sp_RequestData110"
Public Sub p_RequestData110()

Dim wb(1 To 5) As Workbook
Dim ws As Worksheet
Dim env As String
Dim RegURL As String
Dim Endpoint As String
Dim newBody As String
Dim allBody As String
Dim response As String
Dim objHTTP As Object
Set request = CreateObject("WinHttp.WinHttpRequest.5.1")

Set wb(1) = ActiveWorkbook
With wb(1).Sheets("Response1")
newBody = .Range("A10").Value
allBody = newBody

RegURL = "https://quaratesws.oldrepublictitle.com/Calculator/CalculateOrder"
request.Open "POST", RegURL, False
request.setRequestHeader "Content-Type", "application/json; charset=UTF-8"

request.send allBody
response = request.responseText
    
.Range("A24").Clear
'MsgBox request.responseText
.Range("A24").Value = response

Dim y As Range
Dim a As Range
Dim var As String

For Each a In .Range("A24")
     If InStr(a.Value, "PropertyTax") = 0 Then
       Exit Sub
    End If
Next a

For Each a In .Range("A24")
     If InStr(a.Value, "PropertyTax") > 0 Then
        a.Value = Right(a.Value, InStr(a.Value, "PropertyTax") + 155)
    End If
Next a

For Each y In .Range("A24")
     If InStr(y.Value, "Calculated") > 0 Then
        y.Value = Left(y.Value, InStr(y.Value, "Endorsements") - 23)
    End If
Next y

var1 = .Range("A24").Value
.Range("A24").Value = Replace(var1, ":null", "")

var2 = .Range("A24").Value
.Range("A24").Value = Replace(var2, "NetPremium", "")

var3 = .Range("A24").Value
.Range("A24").Value = Replace(var3, "Calculated", "")

var4 = .Range("A24").Value
.Range("A24").Value = Replace(var4, "Premium", "")

var6 = .Range("A24").Value
.Range("A24").Value = Replace(var6, ":", "")

For Each y In .Range("A15")
     If InStr(y.Value, """,""") > 0 Then
        y.Value = Left(y.Value, InStr(y.Value, """,""") - 2)
    End If
Next y
End With
End Sub


