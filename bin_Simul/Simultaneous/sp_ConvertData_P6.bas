Attribute VB_Name = "sp_ConvertData_P6"
Public Sub p_ConvertData_P6()

Dim wb(1 To 3) As Workbook
 
Workbooks("File6").Activate
 
With Workbooks("File6").Worksheets("DataSet1")
 a = .Range("A2")
 b = .Range("B2")
 c = .Range("C2")
 D = .Range("D2")
 e = .Range("E2")
 f = .Range("F2")
 g = .Range("G2")
 h = .Range("H2")
 i = .Range("I2")
 j = .Range("J2")
 k = .Range("K2")
 l = .Range("L2")
 m = .Range("M2")
 n = .Range("N2")
 o = .Range("O2")
 p = .Range("P2")
 q = .Range("Q2")
 r = .Range("R2")
 s = .Range("S2")
 t = .Range("T2")
 u = .Range("U2")
 v = .Range("V2")
 w = .Range("W2")
 x = .Range("X2")

 a2 = .Range("A3")
 b2 = .Range("B3")
 C2 = .Range("C3")
 D2 = .Range("D3")
 e2 = .Range("E3")
 f2 = .Range("F3")
 g2 = .Range("G3")
 h2 = .Range("H3")
 i2 = .Range("I3")
 j2 = .Range("J3")
 k2 = .Range("K3")
 l2 = .Range("L3")
 m2 = .Range("M3")
 n2 = .Range("N3")
 o2 = .Range("O3")
 p2 = .Range("P3")
 q2 = .Range("Q3")
 r2 = .Range("R3")
 s2 = .Range("S3")
 t2 = .Range("T3")
 u2 = .Range("U3")
 v2 = .Range("V3")
 w2 = .Range("W3")
 x2 = .Range("X3")

 a3 = .Range("A4")
 b3 = .Range("B4")
 C3 = .Range("C4")
 D3 = .Range("D4")
 e3 = .Range("E4")
 f3 = .Range("F4")
 g3 = .Range("G4")
 h3 = .Range("H4")
 i3 = .Range("I4")
 j3 = .Range("J4")
 k3 = .Range("K4")
 l3 = .Range("L4")
 m3 = .Range("M4")
 n3 = .Range("N4")
 o3 = .Range("O4")
 p3 = .Range("P4")
 q3 = .Range("Q4")
 r3 = .Range("R4")
 s3 = .Range("S4")
 t3 = .Range("T4")
 u3 = .Range("U4")
 v3 = .Range("V4")
 w3 = .Range("W4")
 x3 = .Range("X4")
 
 a4 = .Range("A5")
 b4 = .Range("B5")
 C4 = .Range("C5")
 D4 = .Range("D5")
 e4 = .Range("E5")
 f4 = .Range("F5")
 g4 = .Range("G5")
 h4 = .Range("H5")
 i4 = .Range("I5")
 j4 = .Range("J5")
 k4 = .Range("K5")
 l4 = .Range("L5")
 m4 = .Range("M5")
 n4 = .Range("N5")
 o4 = .Range("O5")
 p4 = .Range("P5")
 q4 = .Range("Q5")
 r4 = .Range("R5")
 s4 = .Range("S5")
 t4 = .Range("T5")
 u4 = .Range("U5")
 v4 = .Range("V5")
 w4 = .Range("W5")
 x4 = .Range("X5")
 
 a5 = .Range("A6")
 b5 = .Range("B6")
 C5 = .Range("C6")
 D5 = .Range("D6")
 e5 = .Range("E6")
 f5 = .Range("F6")
 g5 = .Range("G6")
 h5 = .Range("H6")
 i5 = .Range("I6")
 j5 = .Range("J6")
 k5 = .Range("K6")
 l5 = .Range("L6")
 m5 = .Range("M6")
 n5 = .Range("N6")
 o5 = .Range("O6")
 p5 = .Range("P6")
 q5 = .Range("Q6")
 r5 = .Range("R6")
 s5 = .Range("S6")
 t5 = .Range("T6")
 u5 = .Range("U6")
 v5 = .Range("V6")
 w5 = .Range("W6")
 x5 = .Range("X6")
 
  a6 = .Range("A7")
 b6 = .Range("B7")
 C6 = .Range("C7")
 D6 = .Range("D7")
 e6 = .Range("E7")
 f6 = .Range("F7")
 g6 = .Range("G7")
 h6 = .Range("H7")
 i6 = .Range("I7")
 j6 = .Range("J7")
 k6 = .Range("K7")
 l6 = .Range("L7")
 m6 = .Range("M7")
 n6 = .Range("N7")
 o6 = .Range("O7")
 p6 = .Range("P7")
 q6 = .Range("Q7")
 r6 = .Range("R7")
 s6 = .Range("S7")
 t6 = .Range("T7")
 u6 = .Range("U7")
 v6 = .Range("V7")
 w6 = .Range("W7")
 x6 = .Range("X7")
 
  a7 = .Range("A8")
 b7 = .Range("B8")
 c7 = .Range("C8")
 D7 = .Range("D8")
 e7 = .Range("E8")
 f7 = .Range("F8")
 g7 = .Range("G8")
 h7 = .Range("H8")
 i7 = .Range("I8")
 j7 = .Range("J8")
 k7 = .Range("K8")
 l7 = .Range("L8")
 m7 = .Range("M8")
 n7 = .Range("N8")
 o7 = .Range("O8")
 p7 = .Range("P8")
 q7 = .Range("Q8")
 r7 = .Range("R8")
 s7 = .Range("S8")
 t7 = .Range("T8")
 u7 = .Range("U8")
 v7 = .Range("V8")
 w7 = .Range("W8")
 x7 = .Range("X8")
 
  a8 = .Range("A9")
 b8 = .Range("B9")
 c8 = .Range("C9")
 D8 = .Range("D9")
 e8 = .Range("E9")
 f8 = .Range("F9")
 g8 = .Range("G9")
 h8 = .Range("H9")
 i8 = .Range("I9")
 j8 = .Range("J9")
 k8 = .Range("K9")
 l8 = .Range("L9")
 m8 = .Range("M9")
 n8 = .Range("N9")
 o8 = .Range("O9")
 p8 = .Range("P9")
 q8 = .Range("Q9")
 r8 = .Range("R9")
 s8 = .Range("S9")
 t8 = .Range("T9")
 u8 = .Range("U9")
 v8 = .Range("V9")
 w8 = .Range("W9")
 x8 = .Range("X9")
 
  a9 = .Range("A10")
 b9 = .Range("B10")
 c9 = .Range("C10")
 D9 = .Range("D10")
 e9 = .Range("E10")
 f9 = .Range("F10")
 g9 = .Range("G10")
 h9 = .Range("H10")
 i9 = .Range("I10")
 j9 = .Range("J10")
 k9 = .Range("K10")
 l9 = .Range("L10")
 m9 = .Range("M10")
 n9 = .Range("N10")
 o9 = .Range("O10")
 p9 = .Range("P10")
 q9 = .Range("Q10")
 r9 = .Range("R10")
 s9 = .Range("S10")
 t9 = .Range("T10")
 u9 = .Range("U10")
 v9 = .Range("V10")
 w9 = .Range("W10")
 x9 = .Range("X10")
 
 a10 = .Range("A11")
 b10 = .Range("B11")
 c10 = .Range("C11")
 D10 = .Range("D11")
 e10 = .Range("E11")
 f10 = .Range("F11")
 g10 = .Range("G11")
 h10 = .Range("H11")
 i10 = .Range("I11")
 j10 = .Range("J11")
 k10 = .Range("K11")
 l10 = .Range("L11")
 m10 = .Range("M11")
 n10 = .Range("N11")
 o10 = .Range("O11")
 p10 = .Range("P11")
 q10 = .Range("Q11")
 r10 = .Range("R11")
 s10 = .Range("S11")
 t10 = .Range("T11")
 u10 = .Range("U11")
 v10 = .Range("V11")
 w10 = .Range("W11")
 x10 = .Range("X11")
 End With
 
 With Workbooks("File6").Worksheets("DataSet2")
 a11 = .Range("A2")
 b11 = .Range("B2")
 c11 = .Range("C2")
 d11 = .Range("D2")
 e11 = .Range("E2")
 f11 = .Range("F2")
 g11 = .Range("G2")
 h11 = .Range("H2")
 i11 = .Range("I2")
 j11 = .Range("J2")
 k11 = .Range("K2")
 l11 = .Range("L2")
 m11 = .Range("M2")
 N11 = .Range("N2")
 o11 = .Range("O2")
 p11 = .Range("P2")
 q11 = .Range("Q2")
 r11 = .Range("R2")
 s11 = .Range("S2")
 t11 = .Range("T2")
 u11 = .Range("U2")
 v11 = .Range("V2")
 w11 = .Range("W2")
 x11 = .Range("X2")
 y11 = .Range("Y2")

 a12 = .Range("A3")
 b12 = .Range("B3")
 C12 = .Range("C3")
 D12 = .Range("D3")
 e12 = .Range("E3")
 f12 = .Range("F3")
 g12 = .Range("G3")
 h12 = .Range("H3")
 i12 = .Range("I3")
 j12 = .Range("J3")
 k12 = .Range("K3")
 l12 = .Range("L3")
 m12 = .Range("M3")
 n12 = .Range("N3")
 o12 = .Range("O3")
 p12 = .Range("P3")
 q12 = .Range("Q3")
 r12 = .Range("R3")
 s12 = .Range("S3")
 t12 = .Range("T3")
 u12 = .Range("U3")
 v12 = .Range("V3")
 w12 = .Range("W3")
 x12 = .Range("X3")
 y12 = .Range("Y3")
 
 a13 = .Range("A4")
 b13 = .Range("B4")
 C13 = .Range("C4")
 D13 = .Range("D4")
 e13 = .Range("E4")
 f13 = .Range("F4")
 g13 = .Range("G4")
 h13 = .Range("H4")
 i13 = .Range("I4")
 j13 = .Range("J4")
 k13 = .Range("K4")
 l13 = .Range("L4")
 m13 = .Range("M4")
 n13 = .Range("N4")
 o13 = .Range("O4")
 p13 = .Range("P4")
 q13 = .Range("Q4")
 r13 = .Range("R4")
 s13 = .Range("S4")
 t13 = .Range("T4")
 u13 = .Range("U4")
 v13 = .Range("V4")
 w13 = .Range("W4")
 x13 = .Range("X4")
 y13 = .Range("Y4")
 
 a14 = .Range("A5")
 b14 = .Range("B5")
 C14 = .Range("C5")
 D14 = .Range("D5")
 e14 = .Range("E5")
 f14 = .Range("F5")
 g14 = .Range("G5")
 h14 = .Range("H5")
 i14 = .Range("I5")
 j14 = .Range("J5")
 k14 = .Range("K5")
 l14 = .Range("L5")
 m14 = .Range("M5")
 n14 = .Range("N5")
 o14 = .Range("O5")
 p14 = .Range("P5")
 q14 = .Range("Q5")
 r14 = .Range("R5")
 s14 = .Range("S5")
 t14 = .Range("T5")
 u14 = .Range("U5")
 v14 = .Range("V5")
 w14 = .Range("W5")
 x14 = .Range("X5")
 y14 = .Range("Y5")
 
 a15 = .Range("A6")
 b15 = .Range("B6")
 C15 = .Range("C6")
 D15 = .Range("D6")
 e15 = .Range("E6")
 f15 = .Range("F6")
 g15 = .Range("G6")
 h15 = .Range("H6")
 i15 = .Range("I6")
 j15 = .Range("J6")
 k15 = .Range("K6")
 l15 = .Range("L6")
 m15 = .Range("M6")
 n15 = .Range("N6")
 o15 = .Range("O6")
 p15 = .Range("P6")
 q15 = .Range("Q6")
 r15 = .Range("R6")
 s15 = .Range("S6")
 t15 = .Range("T6")
 u15 = .Range("U6")
 v15 = .Range("V6")
 w15 = .Range("W6")
 x15 = .Range("X6")
 y15 = .Range("Y6")
 
  a16 = .Range("A7")
 b16 = .Range("B7")
 C16 = .Range("C7")
 D16 = .Range("D7")
 e16 = .Range("E7")
 f16 = .Range("F7")
 g16 = .Range("G7")
 h16 = .Range("H7")
 i16 = .Range("I7")
 j16 = .Range("J7")
 k16 = .Range("K7")
 l16 = .Range("L7")
 m16 = .Range("M7")
 n16 = .Range("N7")
 o16 = .Range("O7")
 p16 = .Range("P7")
 q16 = .Range("Q7")
 r16 = .Range("R7")
 s16 = .Range("S7")
 t16 = .Range("T7")
 u16 = .Range("U7")
 v16 = .Range("V7")
 w16 = .Range("W7")
 x16 = .Range("X7")
 y16 = .Range("Y7")
 
  a17 = .Range("A8")
 b17 = .Range("B8")
 c17 = .Range("C8")
 D17 = .Range("D8")
 e17 = .Range("E8")
 f17 = .Range("F8")
 g17 = .Range("G8")
 h17 = .Range("H8")
 i17 = .Range("I8")
 j17 = .Range("J8")
 k17 = .Range("K8")
 l17 = .Range("L8")
 m17 = .Range("M8")
 n17 = .Range("N8")
 o17 = .Range("O8")
 p17 = .Range("P8")
 q17 = .Range("Q8")
 r17 = .Range("R8")
 s17 = .Range("S8")
 t17 = .Range("T8")
 u17 = .Range("U8")
 v17 = .Range("V8")
 w17 = .Range("W8")
 x17 = .Range("X8")
 y17 = .Range("Y8")
 
  a18 = .Range("A9")
 b18 = .Range("B9")
 c18 = .Range("C9")
 D18 = .Range("D9")
 e18 = .Range("E9")
 f18 = .Range("F9")
 g18 = .Range("G9")
 h18 = .Range("H9")
 i18 = .Range("I9")
 j18 = .Range("J9")
 k18 = .Range("K9")
 l18 = .Range("L9")
 m18 = .Range("M9")
 n18 = .Range("N9")
 o18 = .Range("O9")
 p18 = .Range("P9")
 q18 = .Range("Q9")
 r18 = .Range("R9")
 s18 = .Range("S9")
 t18 = .Range("T9")
 u18 = .Range("U9")
 v18 = .Range("V9")
 w18 = .Range("W9")
 x18 = .Range("X9")
 y18 = .Range("Y9")
 
  a19 = .Range("A10")
 b19 = .Range("B10")
 c19 = .Range("C10")
 D19 = .Range("D10")
 e19 = .Range("E10")
 f19 = .Range("F10")
 g19 = .Range("G10")
 h19 = .Range("H10")
 i19 = .Range("I10")
 j19 = .Range("J10")
 k19 = .Range("K10")
 l19 = .Range("L10")
 m19 = .Range("M10")
 n19 = .Range("N10")
 o19 = .Range("O10")
 p19 = .Range("P10")
 q19 = .Range("Q10")
 r19 = .Range("R10")
 s19 = .Range("S10")
 t19 = .Range("T10")
 u19 = .Range("U10")
 v19 = .Range("V10")
 w19 = .Range("W10")
 x19 = .Range("X10")
 y19 = .Range("Y10")
 
 a20 = .Range("A11")
 b20 = .Range("B11")
 c20 = .Range("C11")
 D20 = .Range("D11")
 e20 = .Range("E11")
 f20 = .Range("F11")
 g20 = .Range("G11")
 h20 = .Range("H11")
 i20 = .Range("I11")
 j20 = .Range("J11")
 k20 = .Range("K11")
 l20 = .Range("L11")
 m20 = .Range("M11")
 n20 = .Range("N11")
 o20 = .Range("O11")
 p20 = .Range("P11")
 q20 = .Range("Q11")
 r20 = .Range("R11")
 s20 = .Range("S11")
 t20 = .Range("T11")
 u20 = .Range("U11")
 v20 = .Range("V11")
 w20 = .Range("W11")
 x20 = .Range("X11")
 y20 = .Range("Y11")
 End With
 
Workbooks.Open "H:\ORT Projects\Rate Engine Rewrite\VBA Macros\Data_Processing\" & "Datadump.xlsx"
Set wb(1) = Workbooks("Datadump.xlsx")
wb(1).Activate
wb(1).Sheets.Add.Name = "Response6"


        With wb(1).Sheets("Response6")
        .Range("A1:A20").Clear
        .Range("A1") = s & w & i & r & a & q & j & r & b & q & k & r & c & q & "Policies" & w & x & t & s & w & l & r & e & q & m & r & f & q & n & r & g & q & o & r & h & w & u + y11 + s11 & w11 & l11 & r11 & e11 & q11 & m11 & r11 & f11 & q11 & N11 & r11 & g11 & q11 & o11 & r11 & h11 & w11 & u11 & v11 & u11
        .Range("A2") = s2 & w2 & i2 & r2 & a2 & q2 & j2 & r2 & b2 & q2 & k2 & r2 & C2 & q2 & "Policies" & w2 & x2 & t2 & s2 & w2 & l2 & r2 & e2 & q2 & m2 & r2 & f2 & q2 & n2 & r2 & g2 & q2 & o2 & r2 & h2 & w2 & u2 + y12 + s12 & w12 & l12 & r12 & e12 & q12 & m12 & r12 & f12 & q12 & n12 & r12 & g12 & q12 & o12 & r12 & h12 & w12 & u12 & v12 & u12
        .Range("A3") = s3 & w3 & i3 & r3 & a3 & q3 & j3 & r3 & b3 & q3 & k3 & r3 & C3 & q3 & "Policies" & w3 & x3 & t3 & s3 & w3 & l3 & r3 & e3 & q3 & m3 & r3 & f3 & q3 & n3 & r3 & g3 & q3 & o3 & r3 & h3 & w3 & u3 + y13 + s13 & w13 & l13 & r13 & e13 & q13 & m13 & r13 & f13 & q13 & n13 & r13 & g13 & q13 & o13 & r13 & h13 & w13 & u13 & v13 & u13
        .Range("A4") = s4 & w4 & i4 & r4 & a4 & q4 & j4 & r4 & b4 & q4 & k4 & r4 & C4 & q4 & "Policies" & w4 & x4 & t4 & s4 & w4 & l4 & r4 & e4 & q4 & m4 & r4 & f4 & q4 & n4 & r4 & g4 & q4 & o4 & r4 & h4 & w4 & u4 + y14 + s14 & w14 & l14 & r14 & e14 & q14 & m14 & r14 & f14 & q14 & n14 & r14 & g14 & q14 & o14 & r14 & h14 & w14 & u14 & v14 & u14
        .Range("A5") = s5 & w5 & i5 & r5 & a5 & q5 & j5 & r5 & b5 & q5 & k5 & r5 & C5 & q5 & "Policies" & w5 & x5 & t5 & s5 & w5 & l5 & r5 & e5 & q5 & m5 & r5 & f5 & q5 & n5 & r5 & g5 & q5 & o5 & r5 & h5 & w5 & u5 + y15 + s15 & w15 & l15 & r15 & e15 & q15 & m15 & r15 & f15 & q15 & n15 & r15 & g15 & q15 & o15 & r15 & h15 & w15 & u15 & v15 & u15
        .Range("A6") = s6 & w6 & i6 & r6 & a6 & q6 & j6 & r6 & b6 & q6 & k6 & r6 & C6 & q6 & "Policies" & w6 & x6 & t6 & s6 & w6 & l6 & r6 & e6 & q6 & m6 & r6 & f6 & q6 & n6 & r6 & g6 & q6 & o6 & r6 & h6 & w6 & u6 + y16 + s16 & w16 & l16 & r16 & e16 & q16 & m16 & r16 & f16 & q16 & n16 & r16 & g16 & q16 & o16 & r16 & h16 & w16 & u16 & v16 & u16
        .Range("A7") = s7 & w7 & i7 & r7 & a7 & q7 & j7 & r7 & b7 & q7 & k7 & r7 & c7 & q7 & "Policies" & w7 & x7 & t7 & s7 & w7 & l7 & r7 & e7 & q7 & m7 & r7 & f7 & q7 & n7 & r7 & g7 & q7 & o7 & r7 & h7 & w7 & u7 + y17 + s17 & w17 & l17 & r17 & e17 & q17 & m17 & r17 & f17 & q17 & n17 & r17 & g17 & q17 & o17 & r17 & h17 & w17 & u17 & v17 & u17
        .Range("A8") = s8 & w8 & i8 & r8 & a8 & q8 & j8 & r8 & b8 & q8 & k8 & r8 & c8 & q8 & "Policies" & w8 & x8 & t8 & s8 & w8 & l8 & r8 & e8 & q8 & m8 & r8 & f8 & q8 & n8 & r8 & g8 & q8 & o8 & r8 & h8 & w8 & u8 + y18 + s18 & w18 & l18 & r18 & e18 & q18 & m18 & r18 & f18 & q18 & n18 & r18 & g18 & q18 & o18 & r18 & h18 & w18 & u18 & v18 & u18
        .Range("A9") = s9 & w9 & i9 & r9 & a9 & q9 & j9 & r9 & b9 & q9 & k9 & r9 & c9 & q9 & "Policies" & w9 & x9 & t9 & s9 & w9 & l9 & r9 & e9 & q9 & m9 & r9 & f9 & q9 & n9 & r9 & g9 & q9 & o9 & r9 & h9 & w9 & u9 + y19 + s19 & w19 & l19 & r19 & e19 & q19 & m19 & r19 & f19 & q19 & n19 & r19 & g19 & q19 & o19 & r19 & h19 & w19 & u19 & v19 & u19
        .Range("A10") = s10 & w10 & i10 & r10 & a10 & q10 & j10 & r10 & b10 & q10 & k10 & r10 & c10 & q10 & "Policies" & w10 & x10 & t10 & s10 & w10 & l10 & r10 & e10 & q10 & m10 & r10 & f10 & q10 & n10 & r10 & g10 & q10 & o10 & r10 & h10 & w10 & u10 + y20 + s20 & w20 & l20 & r20 & e20 & q20 & m20 & r20 & f20 & q20 & n20 & r20 & g20 & q20 & o20 & r20 & h20 & w20 & u20 & v20 & u20
        End With
        
'Disable  alerts
With Application
.DisplayAlerts = False
.ScreenUpdating = False
End With

'Workbooks("File6").Close

'Dim DirFile As String
'DirFile = "File6.xlsx"
'
'    If Len(Dir(DirFile)) <> 0 Then
'        SetAttr DirFile, vbNormal
'        Kill DirFile
'    End If
wb(1).Save

End Sub


