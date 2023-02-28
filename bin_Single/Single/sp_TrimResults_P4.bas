Attribute VB_Name = "sp_TrimResults_P4"
Public Function RemoveSpecialChars04(ByVal mfr As Range)

    Dim splChars As String
    Dim ch As Variant
    Dim splCharArray() As String

    splChars = ": , @ ; [ ] { } "" ium m u e N oal Calclat al"

    splCharArray = Split(splChars, " ")


    For Each ch In splCharArray
    mfr.Replace What:=ch, Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
    Next ch
    RemoveSpecialChars = mfr

End Function


Public Sub p_TrimResults_P4()

With Worksheets(1)
    p1 = RemoveSpecialChars04(.Range("N3"))
    p2 = RemoveSpecialChars04(.Range("N4"))
    p3 = RemoveSpecialChars04(.Range("N5"))
    p4 = RemoveSpecialChars04(.Range("N6"))
    p5 = RemoveSpecialChars04(.Range("N7"))
    p6 = RemoveSpecialChars04(.Range("N8"))
    p7 = RemoveSpecialChars04(.Range("N9"))
    p8 = RemoveSpecialChars04(.Range("N10"))
    p9 = RemoveSpecialChars04(.Range("N11"))
    p10 = RemoveSpecialChars04(.Range("N12"))
End With

End Sub


