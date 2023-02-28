Attribute VB_Name = "sp_RequestData304"
Public Sub p_RequestData304()

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
With wb(1).Sheets("Response3")
newBody = .Range("A4").Value
allBody = newBody

RegURL = "https://quaratesws.oldrepublictitle.com/Calculator/CalculateOrder"
request.Open "POST", RegURL, False
request.setRequestHeader "Content-Type", "application/json; charset=UTF-8"

request.send allBody
response = request.responseText
    
.Range("A18").Clear
'MsgBox request.responseText
.Range("A18").Value = response

Dim y As Range
Dim a As Range
Dim var As String

For Each a In .Range("A18")
     If InStr(a.Value, "PropertyTax") > 0 Then
        a.Value = Right(a.Value, InStr(a.Value, "PropertyTax") + 143)
    End If
Next a

For Each y In .Range("A18")
     If InStr(y.Value, "Endorsements") > 0 Then
        y.Value = Left(y.Value, InStr(y.Value, "Endorsements") - 15)
    End If
Next y

var1 = .Range("A18").Value
.Range("A18").Value = Replace(var1, ":null", "")

var2 = .Range("A18").Value
.Range("A18").Value = Replace(var2, "NetPremium", "")

var3 = .Range("A18").Value
.Range("A18").Value = Replace(var3, "Endorsements", "")

For Each y In .Range("A18")
     If InStr(y.Value, """,""") > 0 Then
        y.Value = Left(y.Value, InStr(y.Value, """,""") - 2)
    End If
Next y
End With
End Sub


