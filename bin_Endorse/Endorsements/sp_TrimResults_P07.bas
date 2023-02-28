Attribute VB_Name = "sp_TrimResults_P07"
Public Function RemoveSpecialChars07(ByVal mfr As Range)

    Dim splChars As String
    Dim ch As Variant
    Dim splCharArray() As String

    splChars = ": , @ ; [ ] { } / # - | "" A a B b C c D d E e F f G g H h I i J j K k L l M m N n O o P p Q q R r S s T t U u V v W w X x Y y Z z"

    splCharArray = Split(splChars, " ")


    For Each ch In splCharArray
    mfr.Replace What:=ch, Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
    Next ch
    RemoveSpecialChars = mfr

End Function


Public Sub p_TrimResults_P07()

With Worksheets(1)
    p1 = RemoveSpecialChars07(.Range("N3"))
    p2 = RemoveSpecialChars07(.Range("N4"))
    p3 = RemoveSpecialChars07(.Range("N5"))
    p4 = RemoveSpecialChars07(.Range("N6"))
    p5 = RemoveSpecialChars07(.Range("N7"))
    p6 = RemoveSpecialChars07(.Range("N8"))
    p7 = RemoveSpecialChars07(.Range("N9"))
    p8 = RemoveSpecialChars07(.Range("N10"))
    p9 = RemoveSpecialChars07(.Range("N11"))
    p10 = RemoveSpecialChars07(.Range("N12"))
     p11 = RemoveSpecialChars07(.Range("N13"))
    p12 = RemoveSpecialChars07(.Range("N14"))
    p13 = RemoveSpecialChars07(.Range("N15"))
    p14 = RemoveSpecialChars07(.Range("N16"))
    p15 = RemoveSpecialChars07(.Range("N17"))
    p16 = RemoveSpecialChars07(.Range("N18"))
    p17 = RemoveSpecialChars07(.Range("N19"))
    p18 = RemoveSpecialChars07(.Range("N20"))
    p19 = RemoveSpecialChars07(.Range("N22"))
    p20 = RemoveSpecialChars07(.Range("N24"))
End With

End Sub




