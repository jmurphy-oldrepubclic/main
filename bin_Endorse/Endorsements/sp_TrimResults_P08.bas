Attribute VB_Name = "sp_TrimResults_P08"
Public Function RemoveSpecialChars08(ByVal mfr As Range)

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


Public Sub p_TrimResults_P08()

With Worksheets(1)
    p1 = RemoveSpecialChars08(.Range("N3"))
    p2 = RemoveSpecialChars08(.Range("N4"))
    p3 = RemoveSpecialChars08(.Range("N5"))
    p4 = RemoveSpecialChars08(.Range("N6"))
    p5 = RemoveSpecialChars08(.Range("N7"))
    p6 = RemoveSpecialChars08(.Range("N8"))
    p7 = RemoveSpecialChars08(.Range("N9"))
    p8 = RemoveSpecialChars08(.Range("N10"))
    p9 = RemoveSpecialChars08(.Range("N11"))
    p10 = RemoveSpecialChars08(.Range("N12"))
     p11 = RemoveSpecialChars08(.Range("N13"))
    p12 = RemoveSpecialChars08(.Range("N14"))
    p13 = RemoveSpecialChars08(.Range("N15"))
    p14 = RemoveSpecialChars08(.Range("N16"))
    p15 = RemoveSpecialChars08(.Range("N17"))
    p16 = RemoveSpecialChars08(.Range("N18"))
    p17 = RemoveSpecialChars08(.Range("N19"))
    p18 = RemoveSpecialChars08(.Range("N20"))
    p19 = RemoveSpecialChars08(.Range("N22"))
    p20 = RemoveSpecialChars08(.Range("N24"))
End With

End Sub






