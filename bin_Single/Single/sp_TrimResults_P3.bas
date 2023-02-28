Attribute VB_Name = "sp_TrimResults_P3"
Public Function RemoveSpecialChars03(ByVal mfr As Range)

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


Public Sub p_TrimResults_P3()

With Worksheets(1)
    p1 = RemoveSpecialChars03(.Range("N3"))
    p2 = RemoveSpecialChars03(.Range("N4"))
    p3 = RemoveSpecialChars03(.Range("N5"))
    p4 = RemoveSpecialChars03(.Range("N6"))
    p5 = RemoveSpecialChars03(.Range("N7"))
    p6 = RemoveSpecialChars03(.Range("N8"))
    p7 = RemoveSpecialChars03(.Range("N9"))
    p8 = RemoveSpecialChars03(.Range("N10"))
    p9 = RemoveSpecialChars03(.Range("N11"))
    p10 = RemoveSpecialChars03(.Range("N12"))
End With

End Sub



