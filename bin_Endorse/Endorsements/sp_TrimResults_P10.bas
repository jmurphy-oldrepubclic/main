Attribute VB_Name = "sp_TrimResults_P10"
Public Function RemoveSpecialChars10(ByVal mfr As Range)

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


Public Sub p_TrimResults_P10()

With Worksheets(1)
    p1 = RemoveSpecialChars10(.Range("N3"))
    p2 = RemoveSpecialChars10(.Range("N4"))
    p3 = RemoveSpecialChars10(.Range("N5"))
    p4 = RemoveSpecialChars10(.Range("N6"))
    p5 = RemoveSpecialChars10(.Range("N7"))
    p6 = RemoveSpecialChars10(.Range("N8"))
    p7 = RemoveSpecialChars10(.Range("N9"))
    p8 = RemoveSpecialChars10(.Range("N10"))
    p9 = RemoveSpecialChars10(.Range("N11"))
    p10 = RemoveSpecialChars10(.Range("N12"))
     p11 = RemoveSpecialChars10(.Range("N13"))
    p12 = RemoveSpecialChars10(.Range("N14"))
    p13 = RemoveSpecialChars10(.Range("N15"))
    p14 = RemoveSpecialChars10(.Range("N16"))
    p15 = RemoveSpecialChars10(.Range("N17"))
    p16 = RemoveSpecialChars10(.Range("N18"))
    p17 = RemoveSpecialChars10(.Range("N19"))
    p18 = RemoveSpecialChars10(.Range("N20"))
    p19 = RemoveSpecialChars10(.Range("N22"))
    p20 = RemoveSpecialChars10(.Range("N24"))
End With

End Sub






