Attribute VB_Name = "Module1"
 Public wFlags As Long
    Public X As Long
    Public i As Long
    Declare Function waveOutGetNumDevs Lib "winmm" () As Long
    Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal wFlags As Long) As Long
    Declare Function ExitWindows Lib "User32.dll" (ByVal RestartCode As Long, ByVal DOSReturnCode As Integer) As Integer
Global Const SND_SYNC = &H0      'play sound first then continue
Global Const SND_ASYNC = &H1     'start playing sound and continue immediately
Global Const SND_NODEFAULT = &H2 'if file isn't found don't play a default sound
Global Const SND_LOOP = &H8      'continuously play sound must be accompanied by SND_ASYNC
Global Const SND_NOSTOP = &H10   'if another sound is playing do not play new sound
' End of Sound Crap
Public SOUND_ON As Boolean
Public aa, bb, cc, dd, ee, ff, gg, hh, ii, jj, kk, ll, mm, nn, oo, pp, qq, rr, ss, tt, uu, vv, ww, xx, yy, zz, sp
Public aaa, bbb, ccc, ddd, eee, fff, ggg, hhh, iii, jjj, kkk, lll, mmm, nnn, ooo, ppp, qqq, rrr, sss, ttt, uuu, vvv, www, xxx, yyy, zzz, _
a1, a2, a3, a4, a5, a6, a7, a8, a9, a0, _
ChrEvent As Boolean, _
CurLetter, _
CurLetPos, _
MessageDump, _
ExPro As Boolean, _
GenK As Boolean, _
EncDec As Integer, _
PassIf As Integer, _
PassLenoLen As Integer, _
PassLen As Integer, _
password, _
PassDoc As Boolean, _
Hedr, _
PassEnt As Boolean, _
Hedrlen As Integer, _
KeyInDoc As Boolean, _
CanUnload As Boolean




Public Sub KeyGen()
MDIForm1.ActiveForm.Label3.Caption = "Status: Generating Key": MDIForm1.ActiveForm.Label3.Refresh
MDIForm1.ActiveForm.Text2.Text = ""
MDIForm1.ActiveForm.Text2.Refresh
Dim curnum As String
Randomize Timer
If ChrEvent = False Then
1 aa = Int(Rnd * 999)
    If Len(aa) <> 3 Then GoTo 1
2 bb = Int(Rnd * 999)
    If Len(bb) <> 3 Then GoTo 2
3 cc = Int(Rnd * 999)
    If Len(cc) <> 3 Then GoTo 3
4 dd = Int(Rnd * 999)
    If Len(dd) <> 3 Then GoTo 4
5 ee = Int(Rnd * 999)
    If Len(ee) <> 3 Then GoTo 5
6 ff = Int(Rnd * 999)
    If Len(ff) <> 3 Then GoTo 6
7 gg = Int(Rnd * 999)
    If Len(gg) <> 3 Then GoTo 7
8 hh = Int(Rnd * 999)
    If Len(hh) <> 3 Then GoTo 8
9 ii = Int(Rnd * 999)
    If Len(ii) <> 3 Then GoTo 9
10 jj = Int(Rnd * 999)
    If Len(jj) <> 3 Then GoTo 10
11 kk = Int(Rnd * 999)
    If Len(kk) <> 3 Then GoTo 11
12 ll = Int(Rnd * 999)
    If Len(ll) <> 3 Then GoTo 12
13 mm = Int(Rnd * 999)
    If Len(mm) <> 3 Then GoTo 13
14 nn = Int(Rnd * 999)
    If Len(nn) <> 3 Then GoTo 14
15 oo = Int(Rnd * 999)
    If Len(oo) <> 3 Then GoTo 15
16 pp = Int(Rnd * 999)
    If Len(pp) <> 3 Then GoTo 16
17 qq = Int(Rnd * 999)
    If Len(qq) <> 3 Then GoTo 17
18 rr = Int(Rnd * 999)
    If Len(rr) <> 3 Then GoTo 18
19 ss = Int(Rnd * 999)
    If Len(ss) <> 3 Then GoTo 19
20 tt = Int(Rnd * 999)
    If Len(tt) <> 3 Then GoTo 20
21 uu = Int(Rnd * 999)
    If Len(uu) <> 3 Then GoTo 21
22 vv = Int(Rnd * 999)
    If Len(vv) <> 3 Then GoTo 22
23 ww = Int(Rnd * 999)
    If Len(ww) <> 3 Then GoTo 23
24 xx = Int(Rnd * 999)
    If Len(xx) <> 3 Then GoTo 24
25 yy = Int(Rnd * 999)
    If Len(yy) <> 3 Then GoTo 25
26 zz = Int(Rnd * 999)
    If Len(zz) <> 3 Then GoTo 26
27 sp = Int(Rnd * 999)
    If Len(sp) <> 3 Then GoTo 27
28 a1 = Int(Rnd * 999)
    If Len(a1) <> 3 Then GoTo 28
29 a2 = Int(Rnd * 999)
    If Len(a2) <> 3 Then GoTo 29
30 a3 = Int(Rnd * 999)
    If Len(a3) <> 3 Then GoTo 30
31 a4 = Int(Rnd * 999)
    If Len(a4) <> 3 Then GoTo 31
32 a5 = Int(Rnd * 999)
    If Len(a5) <> 3 Then GoTo 32
33 a6 = Int(Rnd * 999)
    If Len(a6) <> 3 Then GoTo 33
34 a7 = Int(Rnd * 999)
    If Len(a7) <> 3 Then GoTo 34
35 a8 = Int(Rnd * 999)
    If Len(a8) <> 3 Then GoTo 35
36 a9 = Int(Rnd * 999)
    If Len(a9) <> 3 Then GoTo 36
37 a0 = Int(Rnd * 999)
    If Len(a0) <> 3 Then GoTo 37
38 aaa = Int(Rnd * 999)
    If Len(aaa) <> 3 Then GoTo 38
39 bbb = Int(Rnd * 999)
    If Len(bbb) <> 3 Then GoTo 39
40 ccc = Int(Rnd * 999)
    If Len(ccc) <> 3 Then GoTo 40
41 ddd = Int(Rnd * 999)
    If Len(ddd) <> 3 Then GoTo 41
42 eee = Int(Rnd * 999)
    If Len(eee) <> 3 Then GoTo 42
43 fff = Int(Rnd * 999)
    If Len(fff) <> 3 Then GoTo 43
44 ggg = Int(Rnd * 999)
    If Len(ggg) <> 3 Then GoTo 44
45 hhh = Int(Rnd * 999)
    If Len(hhh) <> 3 Then GoTo 45
46 iii = Int(Rnd * 999)
    If Len(iii) <> 3 Then GoTo 46
47 jjj = Int(Rnd * 999)
    If Len(jjj) <> 3 Then GoTo 47
48 kkk = Int(Rnd * 999)
    If Len(kkk) <> 3 Then GoTo 48
49 lll = Int(Rnd * 999)
    If Len(lll) <> 3 Then GoTo 49
50 mmm = Int(Rnd * 999)
    If Len(mmm) <> 3 Then GoTo 50
51 nnn = Int(Rnd * 999)
    If Len(nnn) <> 3 Then GoTo 51
52 ooo = Int(Rnd * 999)
    If Len(ooo) <> 3 Then GoTo 52
53 ppp = Int(Rnd * 999)
    If Len(ppp) <> 3 Then GoTo 53
54 qqq = Int(Rnd * 999)
    If Len(qqq) <> 3 Then GoTo 54
55 rrr = Int(Rnd * 999)
    If Len(rrr) <> 3 Then GoTo 55
56 sss = Int(Rnd * 999)
    If Len(sss) <> 3 Then GoTo 56
57 ttt = Int(Rnd * 999)
    If Len(ttt) <> 3 Then GoTo 57
58 uuu = Int(Rnd * 999)
    If Len(uuu) <> 3 Then GoTo 58
59 vvv = Int(Rnd * 999)
    If Len(vvv) <> 3 Then GoTo 59
60 www = Int(Rnd * 999)
    If Len(www) <> 3 Then GoTo 60
61 xxx = Int(Rnd * 999)
    If Len(xxx) <> 3 Then GoTo 61
62 yyy = Int(Rnd * 999)
    If Len(yyy) <> 3 Then GoTo 62
63 zzz = Int(Rnd * 999)
    If Len(zzz) <> 3 Then GoTo 63
    
       MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1 & a2 & a3 & a4 & a5 & a6 & a7 & a8 & a9 & a0
Call KeyChk
End If
If ChrEvent = True Then
64 aa = Int(Rnd * 255)
        If Len(aa) <> 3 Then GoTo 64
    Dupl = False

65 bb = Int(Rnd * 255)
    If Len(bb) <> 3 Then GoTo 65
        MDIForm1.ActiveForm.Text2.Text = aa & bb
        Dupl = False
            j = -2
                curnum = bb
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 65
                    Next k
             
            
66 cc = Int(Rnd * 255)
    If Len(cc) <> 3 Then GoTo 66
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc
        Dupl = False
            j = -2
                curnum = cc
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 66
                    Next k
             

67 dd = Int(Rnd * 255)
    If Len(dd) <> 3 Then GoTo 67
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd
        Dupl = False
            j = -2
                curnum = dd
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 67
                    Next k
             

68 ee = Int(Rnd * 255)
    If Len(ee) <> 3 Then GoTo 68
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee
        Dupl = False
            j = -2
                curnum = ee
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 68
                    Next k
             

69 ff = Int(Rnd * 255)
    If Len(ff) <> 3 Then GoTo 69
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff
        Dupl = False
            j = -2
                curnum = ff
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 69
                    Next k
             

70 gg = Int(Rnd * 255)
    If Len(gg) <> 3 Then GoTo 70
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg
        Dupl = False
            j = -2
                curnum = gg
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 70
                    Next k
             

71 hh = Int(Rnd * 255)
    If Len(hh) <> 3 Then GoTo 71
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh
        Dupl = False
            j = -2
                curnum = hh
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 71
                    Next k
             

72 ii = Int(Rnd * 255)
    If Len(ii) <> 3 Then GoTo 72
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii
        Dupl = False
            j = -2
                curnum = ii
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 72
                    Next k
             

73 jj = Int(Rnd * 255)
    If Len(jj) <> 3 Then GoTo 73
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj
        Dupl = False
            j = -2
                curnum = jj
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 73
                    Next k
             

74 kk = Int(Rnd * 255)
    If Len(kk) <> 3 Then GoTo 74
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk
        Dupl = False
            j = -2
                curnum = kk
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 74
                    Next k
             

75 ll = Int(Rnd * 255)
    If Len(ll) <> 3 Then GoTo 75
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll
        Dupl = False
            j = -2
                curnum = ll
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 75
                    Next k
             

76 mm = Int(Rnd * 255)
    If Len(mm) <> 3 Then GoTo 76
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm
        Dupl = False
            j = -2
                curnum = mm
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 76
                    Next k
             

77 nn = Int(Rnd * 255)
    If Len(nn) <> 3 Then GoTo 77
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn
        Dupl = False
            j = -2
                curnum = nn
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 77
                    Next k
             

78 oo = Int(Rnd * 255)
    If Len(oo) <> 3 Then GoTo 78
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo
        Dupl = False
            j = -2
                curnum = oo
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 78
                    Next k
             

79 pp = Int(Rnd * 255)
    If Len(pp) <> 3 Then GoTo 79
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp
        Dupl = False
            j = -2
                curnum = pp
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 79
                    Next k
             

80 qq = Int(Rnd * 255)
    If Len(qq) <> 3 Then GoTo 80
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq
        Dupl = False
            j = -2
                curnum = qq
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 80
                    Next k
             

81 rr = Int(Rnd * 255)
    If Len(rr) <> 3 Then GoTo 81
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr
        Dupl = False
            j = -2
                curnum = rr
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 81
                    Next k
             

82 ss = Int(Rnd * 255)
    If Len(ss) <> 3 Then GoTo 82
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss
        Dupl = False
            j = -2
                curnum = ss
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 82
                    Next k
             

83 tt = Int(Rnd * 255)
    If Len(tt) <> 3 Then GoTo 83
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt
        Dupl = False
            j = -2
                curnum = tt
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1: j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 83
                    Next k
             
        
Call KeyGen2
    
End If

End Sub

Public Sub KeyDecode()
Dim FullKey
Dim CurKeyPos

    If MDIForm1.ActiveForm.Text2.Text <> "" Then
        For CurKeyPos = 1 To Len(MDIForm1.ActiveForm.Text2.Text) Step 3
            If CurKeyPos = 1 Then aa = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 4 Then bb = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 7 Then cc = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 10 Then dd = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 13 Then ee = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 16 Then ff = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 19 Then gg = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 22 Then hh = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 25 Then ii = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 28 Then jj = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 31 Then kk = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 34 Then ll = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 37 Then mm = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 40 Then nn = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 43 Then oo = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 46 Then pp = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 49 Then qq = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 52 Then rr = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 55 Then ss = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 58 Then tt = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 61 Then uu = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 64 Then vv = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 67 Then ww = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 70 Then xx = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 73 Then yy = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 76 Then zz = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 79 Then sp = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 82 Then aaa = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 85 Then bbb = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 88 Then ccc = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 91 Then ddd = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 94 Then eee = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 97 Then fff = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 100 Then ggg = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 103 Then hhh = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 106 Then iii = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 109 Then jjj = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 112 Then kkk = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 115 Then lll = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 118 Then mmm = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 121 Then nnn = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 124 Then ooo = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 127 Then ppp = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 130 Then qqq = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 133 Then rrr = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 136 Then sss = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 139 Then ttt = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 142 Then uuu = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 145 Then vvv = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 148 Then www = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 151 Then xxx = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 154 Then yyy = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 157 Then zzz = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 160 Then a1 = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 163 Then a2 = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 166 Then a3 = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 169 Then a4 = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 172 Then a5 = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 175 Then a6 = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 178 Then a7 = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 181 Then a8 = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 184 Then a9 = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
            If CurKeyPos = 187 Then a0 = Mid(MDIForm1.ActiveForm.Text2.Text, CurKeyPos, 3)
                  
         Next CurKeyPos
        Else
            MsgBox "There is no encryption key to decode the message with!", , "ERGH!!"
    End If
End Sub
Public Sub KeyChk()
Dim curnumchk
Dim curnum
Dim Dupl As Boolean
Dupl = False

For j = 1 To Len(MDIForm1.ActiveForm.Text2.Text) Step 3
    curnum = Mid(MDIForm1.ActiveForm.Text2.Text, j, 3)
    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) Step 3
        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3)
        If curnumchk = curnum Then Dupl = True: GoTo 33
    Next k
Next j

33
If Dupl = True Then Call KeyGen

End Sub
Public Sub KeyGen2()
Dim curnum As String
84 uu = Int(Rnd * 255)
    If Len(uu) <> 3 Then GoTo 84
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu
        Dupl = False
            j = -2
                curnum = uu
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                           curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 84
                    Next k
             

85 vv = Int(Rnd * 255)
    If Len(vv) <> 3 Then GoTo 85
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv
        Dupl = False
            j = -2
                curnum = vv
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 85
                    Next k
             

86 ww = Int(Rnd * 255)
    If Len(ww) <> 3 Then GoTo 86
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww
        Dupl = False
            j = -2
                curnum = ww
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 86
                    Next k
             

87 xx = Int(Rnd * 255)
    If Len(xx) <> 3 Then GoTo 87
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx
        Dupl = False
            j = -2
                curnum = xx
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 87
                    Next k
             

88 yy = Int(Rnd * 255)
    If Len(yy) <> 3 Then GoTo 88
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy
        Dupl = False
            j = -2
                curnum = yy
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 88
                    Next k
             

89 zz = Int(Rnd * 255)
    If Len(zz) <> 3 Then GoTo 89
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz
        Dupl = False
            j = -2
                curnum = zz
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 89
                    Next k
             

90 sp = Int(Rnd * 255)
    If Len(sp) <> 3 Then GoTo 90
            MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp
        Dupl = False
            j = -2
                curnum = sp
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 90
                    Next k
             



101 aaa = Int(Rnd * 255)
    If Len(aaa) <> 3 Then GoTo 101
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa
        Dupl = False
            j = -2
                curnum = aaa
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 101
                    Next k
             
    
102 bbb = Int(Rnd * 255)
    If Len(bbb) <> 3 Then GoTo 102
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb
        Dupl = False
            j = -2
                curnum = bbb
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 102
                    Next k
             

103 ccc = Int(Rnd * 255)
    If Len(ccc) <> 3 Then GoTo 103
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc
        Dupl = False
            j = -2
                curnum = ccc
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 103
                    Next k
             

104 ddd = Int(Rnd * 255)
    If Len(ddd) <> 3 Then GoTo 104
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd
        Dupl = False
            j = -2
                curnum = ddd
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 104
                    Next k
             

105 eee = Int(Rnd * 255)
    If Len(eee) <> 3 Then GoTo 105
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee
        Dupl = False
            j = -2
                curnum = eee
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 105
                    Next k
             

106 fff = Int(Rnd * 255)
    If Len(fff) <> 3 Then GoTo 106
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff
        Dupl = False
            j = -2
                curnum = fff
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 106
                    Next k
            

107 ggg = Int(Rnd * 255)
    If Len(ggg) <> 3 Then GoTo 107
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg
        Dupl = False
            j = -2
                curnum = ggg
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 107
                    Next k
             

108 hhh = Int(Rnd * 255)
    If Len(hhh) <> 3 Then GoTo 108
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh
        Dupl = False
            j = -2
                curnum = hhh
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 108
                    Next k
             

109 iii = Int(Rnd * 255)
    If Len(iii) <> 3 Then GoTo 109
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii
        Dupl = False
            j = -2
                curnum = iii
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 109
                    Next k
             

110 jjj = Int(Rnd * 255)
    If Len(jjj) <> 3 Then GoTo 110
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj
        Dupl = False
            j = -2
                curnum = jjj
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 110
                    Next k
             

111 kkk = Int(Rnd * 255)
    If Len(kkk) <> 3 Then GoTo 111
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk
        Dupl = False
            j = -2
                curnum = kkk
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 111
                    Next k
             

112 lll = Int(Rnd * 255)
    If Len(lll) <> 3 Then GoTo 112
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll
        Dupl = False
            j = -2
                curnum = lll
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 112
                    Next k
             

113 mmm = Int(Rnd * 255)
    If Len(mmm) <> 3 Then GoTo 113
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm
        Dupl = False
            j = -2
                curnum = mmm
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 113
                    Next k
             

114 nnn = Int(Rnd * 255)
    If Len(nnn) <> 3 Then GoTo 114
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn
        Dupl = False
            j = -2
                curnum = nnn
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 114
                    Next k
             

115 ooo = Int(Rnd * 255)
    If Len(ooo) <> 3 Then GoTo 115
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo
        Dupl = False
            j = -2
                curnum = ooo
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 115
                    Next k
             

116 ppp = Int(Rnd * 255)
    If Len(ppp) <> 3 Then GoTo 116
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp
        Dupl = False
            j = -2
                curnum = ppp
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 116
                    Next k
             

117 qqq = Int(Rnd * 255)
    If Len(qqq) <> 3 Then GoTo 117
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq
        Dupl = False
            j = -2
                curnum = qqq
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 117
                    Next k
             

118 rrr = Int(Rnd * 255)
    If Len(rrr) <> 3 Then GoTo 118
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr
        Dupl = False
            j = -2
                curnum = rrr
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 118
                    Next k
             

119 sss = Int(Rnd * 255)
    If Len(sss) <> 3 Then GoTo 119
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss
        Dupl = False
            j = -2
                curnum = sss
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 119
                    Next k
             

120 ttt = Int(Rnd * 255)
    If Len(ttt) <> 3 Then GoTo 120
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt
        Dupl = False
            j = -2
                curnum = ttt
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 120
                    Next k
             

121 uuu = Int(Rnd * 255)
    If Len(uuu) <> 3 Then GoTo 121
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu
        Dupl = False
            j = -2
                curnum = uuu
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 121
                    Next k
             

122 vvv = Int(Rnd * 255)
    If Len(vvv) <> 3 Then GoTo 122
            MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv
        Dupl = False
            j = -2
                curnum = vvv
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 122
                    Next k
             

123 www = Int(Rnd * 255)
    If Len(www) <> 3 Then GoTo 123
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www
        Dupl = False
            j = -2
                curnum = www
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 123
                    Next k
             

124 xxx = Int(Rnd * 255)
    If Len(xxx) <> 3 Then GoTo 124
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx
        Dupl = False
            j = -2
                curnum = xxx
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 124
                    Next k
             

125 yyy = Int(Rnd * 255)
    If Len(yyy) <> 3 Then GoTo 125
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy
        Dupl = False
            j = -2
                curnum = yyy
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 125
                    Next k
             

126 zzz = Int(Rnd * 255)
    If Len(zzz) <> 3 Then GoTo 126
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz
        Dupl = False
            j = -2
                curnum = zzz
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 126
                    Next k
             


91 a1 = Int(Rnd * 255)
    If Len(a1) <> 3 Then GoTo 91
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1
        Dupl = False
            j = -2
                curnum = a1
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 91
                    Next k
             

92 a2 = Int(Rnd * 255)
    If Len(a2) <> 3 Then GoTo 92
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1 & a2
        Dupl = False
            j = -2
                curnum = a2
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 92
                    Next k
             

93 a3 = Int(Rnd * 255)
    If Len(a3) <> 3 Then GoTo 93
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1 & a2 & a3
        Dupl = False
            j = -2
                curnum = a3
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                       If curnumchk = curnum Then Dupl = True: GoTo 93
                    Next k
             

94 a4 = Int(Rnd * 255)
    If Len(a4) <> 3 Then GoTo 94
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1 & a2 & a3 & a4
        Dupl = False
            j = -2
                curnum = a4
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 94
                    Next k
             

95 a5 = Int(Rnd * 255)
    If Len(a5) <> 3 Then GoTo 95
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1 & a2 & a3 & a4 & a5
        Dupl = False
            j = -2
                curnum = a5
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 95
                    Next k
             

96 a6 = Int(Rnd * 255)
    If Len(a6) <> 3 Then GoTo 96
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1 & a2 & a3 & a4 & a5 & a6
        Dupl = False
            j = -2
                curnum = a6
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 96
                    Next k
             

97 a7 = Int(Rnd * 255)
    If Len(a7) <> 3 Then GoTo 97
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1 & a2 & a3 & a4 & a5 & a6 & a7
        Dupl = False
            j = -2
                curnum = a7
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 97
                    Next k
             

98 a8 = Int(Rnd * 255)
    If Len(a8) <> 3 Then GoTo 98
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1 & a2 & a3 & a4 & a5 & a6 & a7 & a8
        Dupl = False
            j = -2
                curnum = a8
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 98
                    Next k
             

99 a9 = Int(Rnd * 255)
    If Len(a9) <> 3 Then GoTo 99
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1 & a2 & a3 & a4 & a5 & a6 & a7 & a8 & a9
        Dupl = False
            j = -2
                curnum = a9
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 99
                    Next k
             

100 a0 = Int(Rnd * 255)
    If Len(a0) <> 3 Then GoTo 100
        MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1 & a2 & a3 & a4 & a5 & a6 & a7 & a8 & a9 & a0
        Dupl = False
            j = -2
                curnum = a0
                    For k = j + 3 To Len(MDIForm1.ActiveForm.Text2.Text) - 3 Step 3
                        curnumchk = Mid(MDIForm1.ActiveForm.Text2.Text, k, 3): j = j + 1
                        If curnumchk = curnum Then Dupl = True: GoTo 100
                    Next k
             
MDIForm1.ActiveForm.Text2.Text = aa & bb & cc & dd & ee & ff & gg & hh & ii & jj & kk & ll & mm & nn & oo & pp & qq & rr & ss & tt & uu & vv & ww & xx & yy & zz & sp & aaa & bbb & ccc & ddd & eee & fff & ggg & hhh & iii & jjj & kkk & lll & mmm & nnn & ooo & ppp & qqq & rrr & sss & ttt & uuu & vvv & www & xxx & yyy & zzz & a1 & a2 & a3 & a4 & a5 & a6 & a7 & a8 & a9 & a0

End Sub
Public Sub EncryptMsg()
Dim MessageText
Dim progbnum As Integer
Dim lengthFile As Integer
 Dim progtotal As Integer

    
 If PassDoc = True And PassEnt = False Then
    Form3.Show
    Exit Sub
 End If
 If SOUND_ON = True Then
        strSoundName = "sounds\encryptingstandby.wav"  'or a valid .wav file
        wFlags = SND_SYNC
        X = sndPlaySound(strSoundName, wFlags)
    End If
 MessageText = "JoeCrypt 2.0" & PassIf & PassLenoLen & PassLen & password & Chr(13) + Chr(10) & MDIForm1.ActiveForm.Text1.Text
lengthFile = Len(MessageText)
     ExPro = False
     MessageDump = ""
     messagetext2 = MDIForm1.ActiveForm.Text1.Text
    If MDIForm1.ActiveForm.Text1.Text <> "" Then
    
            If GenK = True Then Call KeyGen Else Call KeyDecode
                If MDIForm1.ActiveForm.Text2.Text = "" Then
                    ans = MsgBox("There was no key found to encrypt with, Generate Key Now?", vbYesNo, "No Key!")
                        If ans = vbYes Then Call KeyGen
                        If ans = vbNo Then Exit Sub
                End If
            MDIForm1.ActiveForm.Label3.Caption = "Status: Encrypting...": MDIForm1.ActiveForm.Label3.Refresh
            Form2.Show
            
            
            Open "tempmsg" For Output As #1
            Print #1, MessageText
            Close #1
            Open "tempmsg" For Input As #1
        Do Until EOF(1)
          DoEvents
          If ExPro = True Then GoTo 24
         Line Input #1, MessageText
         For CurLetPos = 1 To Len(MessageText)
            curposfp = curposfp + 1
            CurLetter = Mid(MessageText, CurLetPos, 1)
                If ChrEvent = False Then
                If CurLetter = "a" Then MessageDump = MessageDump & aa:  GoTo 19
                If CurLetter = "b" Then MessageDump = MessageDump & bb: GoTo 19
                If CurLetter = "c" Then MessageDump = MessageDump & cc: GoTo 19
                If CurLetter = "d" Then MessageDump = MessageDump & dd:  GoTo 19
                If CurLetter = "e" Then MessageDump = MessageDump & ee: GoTo 19
                If CurLetter = "f" Then MessageDump = MessageDump & ff: GoTo 19
                If CurLetter = "g" Then MessageDump = MessageDump & gg: GoTo 19
                If CurLetter = "h" Then MessageDump = MessageDump & hh: GoTo 19
                If CurLetter = "i" Then MessageDump = MessageDump & ii: GoTo 19
                If CurLetter = "j" Then MessageDump = MessageDump & jj: GoTo 19
                If CurLetter = "k" Then MessageDump = MessageDump & kk: GoTo 19
                If CurLetter = "l" Then MessageDump = MessageDump & ll: GoTo 19
                If CurLetter = "m" Then MessageDump = MessageDump & mm: GoTo 19
                If CurLetter = "n" Then MessageDump = MessageDump & nn: GoTo 19
                If CurLetter = "o" Then MessageDump = MessageDump & oo: GoTo 19
                If CurLetter = "p" Then MessageDump = MessageDump & pp: GoTo 19
                If CurLetter = "q" Then MessageDump = MessageDump & qq: GoTo 19
                If CurLetter = "r" Then MessageDump = MessageDump & rr: GoTo 19
                If CurLetter = "s" Then MessageDump = MessageDump & ss: GoTo 19
                If CurLetter = "t" Then MessageDump = MessageDump & tt: GoTo 19
                If CurLetter = "u" Then MessageDump = MessageDump & uu: GoTo 19
                If CurLetter = "v" Then MessageDump = MessageDump & vv: GoTo 19
                If CurLetter = "w" Then MessageDump = MessageDump & ww: GoTo 19
                If CurLetter = "x" Then MessageDump = MessageDump & xx: GoTo 19
                If CurLetter = "y" Then MessageDump = MessageDump & yy: GoTo 19
                If CurLetter = "z" Then MessageDump = MessageDump & zz: GoTo 19
                If CurLetter = " " Then MessageDump = MessageDump & sp: GoTo 19
                If CurLetter = "A" Then MessageDump = MessageDump & aaa: GoTo 19
                If CurLetter = "B" Then MessageDump = MessageDump & bbb: GoTo 19
                If CurLetter = "C" Then MessageDump = MessageDump & ccc: GoTo 19
                If CurLetter = "D" Then MessageDump = MessageDump & ddd: GoTo 19
                If CurLetter = "E" Then MessageDump = MessageDump & eee: GoTo 19
                If CurLetter = "F" Then MessageDump = MessageDump & fff: GoTo 19
                If CurLetter = "G" Then MessageDump = MessageDump & ggg: GoTo 19
                If CurLetter = "H" Then MessageDump = MessageDump & hhh: GoTo 19
                If CurLetter = "I" Then MessageDump = MessageDump & iii: GoTo 19
                If CurLetter = "J" Then MessageDump = MessageDump & jjj: GoTo 19
                If CurLetter = "K" Then MessageDump = MessageDump & kkk: GoTo 19
                If CurLetter = "L" Then MessageDump = MessageDump & lll: GoTo 19
                If CurLetter = "M" Then MessageDump = MessageDump & mmm: GoTo 19
                If CurLetter = "N" Then MessageDump = MessageDump & nnn: GoTo 19
                If CurLetter = "O" Then MessageDump = MessageDump & ooo: GoTo 19
                If CurLetter = "P" Then MessageDump = MessageDump & ppp: GoTo 19
                If CurLetter = "Q" Then MessageDump = MessageDump & qqq: GoTo 19
                If CurLetter = "R" Then MessageDump = MessageDump & rrr: GoTo 19
                If CurLetter = "S" Then MessageDump = MessageDump & sss: GoTo 19
                If CurLetter = "T" Then MessageDump = MessageDump & ttt: GoTo 19
                If CurLetter = "U" Then MessageDump = MessageDump & uuu: GoTo 19
                If CurLetter = "V" Then MessageDump = MessageDump & vvv: GoTo 19
                If CurLetter = "W" Then MessageDump = MessageDump & www: GoTo 19
                If CurLetter = "X" Then MessageDump = MessageDump & xxx: GoTo 19
                If CurLetter = "Y" Then MessageDump = MessageDump & yyy: GoTo 19
                If CurLetter = "Z" Then MessageDump = MessageDump & zzz: GoTo 19
                If CurLetter = "1" Then MessageDump = MessageDump & a1: GoTo 19
                If CurLetter = "2" Then MessageDump = MessageDump & a2: GoTo 19
                If CurLetter = "3" Then MessageDump = MessageDump & a3: GoTo 19
                If CurLetter = "4" Then MessageDump = MessageDump & a4: GoTo 19
                If CurLetter = "5" Then MessageDump = MessageDump & a5: GoTo 19
                If CurLetter = "6" Then MessageDump = MessageDump & a6: GoTo 19
                If CurLetter = "7" Then MessageDump = MessageDump & a7: GoTo 19
                If CurLetter = "8" Then MessageDump = MessageDump & a8: GoTo 19
                If CurLetter = "9" Then MessageDump = MessageDump & a9: GoTo 19
                If CurLetter = "0" Then MessageDump = MessageDump & a0: GoTo 19
                End If
                    If ChrEvent = True Then
                If CurLetter = "a" Then MessageDump = MessageDump & Chr(aa): GoTo 19
                If CurLetter = "b" Then MessageDump = MessageDump & Chr(bb): GoTo 19
                If CurLetter = "c" Then MessageDump = MessageDump & Chr(cc): GoTo 19
                If CurLetter = "d" Then MessageDump = MessageDump & Chr(dd): GoTo 19
                If CurLetter = "e" Then MessageDump = MessageDump & Chr(ee): GoTo 19
                If CurLetter = "f" Then MessageDump = MessageDump & Chr(ff): GoTo 19
                If CurLetter = "g" Then MessageDump = MessageDump & Chr(gg): GoTo 19
                If CurLetter = "h" Then MessageDump = MessageDump & Chr(hh): GoTo 19
                If CurLetter = "i" Then MessageDump = MessageDump & Chr(ii): GoTo 19
                If CurLetter = "j" Then MessageDump = MessageDump & Chr(jj): GoTo 19
                If CurLetter = "k" Then MessageDump = MessageDump & Chr(kk): GoTo 19
                If CurLetter = "l" Then MessageDump = MessageDump & Chr(ll): GoTo 19
                If CurLetter = "m" Then MessageDump = MessageDump & Chr(mm): GoTo 19
                If CurLetter = "n" Then MessageDump = MessageDump & Chr(nn): GoTo 19
                If CurLetter = "o" Then MessageDump = MessageDump & Chr(oo): GoTo 19
                If CurLetter = "p" Then MessageDump = MessageDump & Chr(pp): GoTo 19
                If CurLetter = "q" Then MessageDump = MessageDump & Chr(qq): GoTo 19
                If CurLetter = "r" Then MessageDump = MessageDump & Chr(rr): GoTo 19
                If CurLetter = "s" Then MessageDump = MessageDump & Chr(ss): GoTo 19
                If CurLetter = "t" Then MessageDump = MessageDump & Chr(tt): GoTo 19
                If CurLetter = "u" Then MessageDump = MessageDump & Chr(uu): GoTo 19
                If CurLetter = "v" Then MessageDump = MessageDump & Chr(vv): GoTo 19
                If CurLetter = "w" Then MessageDump = MessageDump & Chr(ww): GoTo 19
                If CurLetter = "x" Then MessageDump = MessageDump & Chr(xx): GoTo 19
                If CurLetter = "y" Then MessageDump = MessageDump & Chr(yy): GoTo 19
                If CurLetter = "z" Then MessageDump = MessageDump & Chr(zz): GoTo 19
                If CurLetter = " " Then MessageDump = MessageDump & Chr(sp): GoTo 19
                If CurLetter = "A" Then MessageDump = MessageDump & Chr(aaa): GoTo 19
                If CurLetter = "B" Then MessageDump = MessageDump & Chr(bbb): GoTo 19
                If CurLetter = "C" Then MessageDump = MessageDump & Chr(ccc): GoTo 19
                If CurLetter = "D" Then MessageDump = MessageDump & Chr(ddd): GoTo 19
                If CurLetter = "E" Then MessageDump = MessageDump & Chr(eee): GoTo 19
                If CurLetter = "F" Then MessageDump = MessageDump & Chr(fff): GoTo 19
                If CurLetter = "G" Then MessageDump = MessageDump & Chr(ggg): GoTo 19
                If CurLetter = "H" Then MessageDump = MessageDump & Chr(hhh): GoTo 19
                If CurLetter = "I" Then MessageDump = MessageDump & Chr(iii): GoTo 19
                If CurLetter = "J" Then MessageDump = MessageDump & Chr(jjj): GoTo 19
                If CurLetter = "K" Then MessageDump = MessageDump & Chr(kkk): GoTo 19
                If CurLetter = "L" Then MessageDump = MessageDump & Chr(lll): GoTo 19
                If CurLetter = "M" Then MessageDump = MessageDump & Chr(mmm): GoTo 19
                If CurLetter = "N" Then MessageDump = MessageDump & Chr(nnn): GoTo 19
                If CurLetter = "O" Then MessageDump = MessageDump & Chr(ooo): GoTo 19
                If CurLetter = "P" Then MessageDump = MessageDump & Chr(ppp): GoTo 19
                If CurLetter = "Q" Then MessageDump = MessageDump & Chr(qqq): GoTo 19
                If CurLetter = "R" Then MessageDump = MessageDump & Chr(rrr): GoTo 19
                If CurLetter = "S" Then MessageDump = MessageDump & Chr(sss): GoTo 19
                If CurLetter = "T" Then MessageDump = MessageDump & Chr(ttt): GoTo 19
                If CurLetter = "U" Then MessageDump = MessageDump & Chr(uuu): GoTo 19
                If CurLetter = "V" Then MessageDump = MessageDump & Chr(vvv): GoTo 19
                If CurLetter = "W" Then MessageDump = MessageDump & Chr(www): GoTo 19
                If CurLetter = "X" Then MessageDump = MessageDump & Chr(xxx): GoTo 19
                If CurLetter = "Y" Then MessageDump = MessageDump & Chr(yyy): GoTo 19
                If CurLetter = "Z" Then MessageDump = MessageDump & Chr(zzz): GoTo 19
                If CurLetter = "1" Then MessageDump = MessageDump & Chr(a1): GoTo 19
                If CurLetter = "2" Then MessageDump = MessageDump & Chr(a2): GoTo 19
                If CurLetter = "3" Then MessageDump = MessageDump & Chr(a3): GoTo 19
                If CurLetter = "4" Then MessageDump = MessageDump & Chr(a4): GoTo 19
                If CurLetter = "5" Then MessageDump = MessageDump & Chr(a5): GoTo 19
                If CurLetter = "6" Then MessageDump = MessageDump & Chr(a6): GoTo 19
                If CurLetter = "7" Then MessageDump = MessageDump & Chr(a7): GoTo 19
                If CurLetter = "8" Then MessageDump = MessageDump & Chr(a8): GoTo 19
                If CurLetter = "9" Then MessageDump = MessageDump & Chr(a9): GoTo 19
                If CurLetter = "0" Then MessageDump = MessageDump & Chr(a0): GoTo 19
                End If
                     
                MessageDump = MessageDump & CurLetter
19
                        
            
                Form2.ProgressBar1.Max = lengthFile
                   progbnum = (curposfp)
                    
                     Form2.ProgressBar1.Value = progbnum
                     
            
                
                Next CurLetPos
         MessageDump = MessageDump & Chr(13) + Chr(10)
         Loop
         Close #1
         MDIForm1.ActiveForm.Text1.Text = MessageDump
         Else
            MsgBox "No text to encrypt!", , "ERGH!"
             
     End If
     MDIForm1.ActiveForm.Label3.Caption = "Status: Done": MDIForm1.ActiveForm.Label3.Refresh
24
    Close #1
On Error Resume Next
     Kill ("tempmsg")
     Form2.Hide
     PassIf = 0
     strSoundName = 0&
    wFlags = SND_ASYNC
    X = sndPlaySound(strSoundName, wFlags)
     If SOUND_ON = True Then
        strSoundName = "sounds\encryptioncomplete.wav"  'or a valid .wav file
        wFlags = SND_SYNC
        X = sndPlaySound(strSoundName, wFlags)
    End If
End Sub
Public Sub DecryptMsg()
If SOUND_ON = True Then
        strSoundName = "sounds\decryptingstandby.wav"  'or a valid .wav file
        wFlags = SND_SYNC
        X = sndPlaySound(strSoundName, wFlags)
    End If
Dim MessageText
Dim CurLetNum As String
Dim progtotal As Integer
  ExPro = False
  MessageDump = ""
     Hedrlen = 0
     Call KeyDecode
     If MDIForm1.ActiveForm.Text2.Text = "" Then Exit Sub
    If MDIForm1.ActiveForm.Text1.Text <> "" Then
           MDIForm1.ActiveForm.Label3.Caption = "Status: Decrypting...": MDIForm1.ActiveForm.Label3.Refresh
            Form2.Show
                        MessageText = MDIForm1.ActiveForm.Text1.Text
              messagetext2 = MDIForm1.ActiveForm.Text1.Text
              Open "tempmsg" For Output As #1
            Print #1, MessageText
            Close #1
            Open "tempmsg" For Input As #1
        
        
        Do Until EOF(1)
         Line Input #1, MessageText
         DoEvents
          If ExPro = True Then GoTo 23
         For CurLetPos = 1 To Len(MessageText)
             CurLetNum2 = Mid(MessageText, CurLetPos, 1)
            If ChrEvent = False Then
             If CurLetNum2 <> "1" And CurLetNum2 <> "2" And CurLetNum2 <> "3" And CurLetNum2 <> "4" And CurLetNum2 <> "5" And CurLetNum2 <> "6" And CurLetNum2 <> "7" And CurLetNum2 <> "8" And CurLetNum2 <> "9" And CurLetNum2 <> "0" Then
                    GoTo 21
                End If
             End If
                
            If ChrEvent = False Then
                    CurLetNum = Mid(MessageText, CurLetPos, 3)
                If CurLetNum = aa Then MessageDump = MessageDump & "a": GoTo 22
                If CurLetNum = bb Then MessageDump = MessageDump & "b": GoTo 22
                If CurLetNum = cc Then MessageDump = MessageDump & "c": GoTo 22
                If CurLetNum = dd Then MessageDump = MessageDump & "d": GoTo 22
                If CurLetNum = ee Then MessageDump = MessageDump & "e": GoTo 22
                If CurLetNum = ff Then MessageDump = MessageDump & "f": GoTo 22
                If CurLetNum = gg Then MessageDump = MessageDump & "g": GoTo 22
                If CurLetNum = hh Then MessageDump = MessageDump & "h": GoTo 22
                If CurLetNum = ii Then MessageDump = MessageDump & "i": GoTo 22
                If CurLetNum = jj Then MessageDump = MessageDump & "j": GoTo 22
                If CurLetNum = kk Then MessageDump = MessageDump & "k": GoTo 22
                If CurLetNum = ll Then MessageDump = MessageDump & "l": GoTo 22
                If CurLetNum = mm Then MessageDump = MessageDump & "m": GoTo 22
                If CurLetNum = nn Then MessageDump = MessageDump & "n": GoTo 22
                If CurLetNum = oo Then MessageDump = MessageDump & "o": GoTo 22
                If CurLetNum = pp Then MessageDump = MessageDump & "p": GoTo 22
                If CurLetNum = qq Then MessageDump = MessageDump & "q": GoTo 22
                If CurLetNum = rr Then MessageDump = MessageDump & "r": GoTo 22
                If CurLetNum = ss Then MessageDump = MessageDump & "s": GoTo 22
                If CurLetNum = tt Then MessageDump = MessageDump & "t": GoTo 22
                If CurLetNum = uu Then MessageDump = MessageDump & "u": GoTo 22
                If CurLetNum = vv Then MessageDump = MessageDump & "v": GoTo 22
                If CurLetNum = ww Then MessageDump = MessageDump & "w": GoTo 22
                If CurLetNum = xx Then MessageDump = MessageDump & "x": GoTo 22
                If CurLetNum = yy Then MessageDump = MessageDump & "y": GoTo 22
                If CurLetNum = zz Then MessageDump = MessageDump & "z": GoTo 22
                If CurLetNum = sp Then MessageDump = MessageDump & " ": GoTo 22
                If CurLetNum = aaa Then MessageDump = MessageDump & "A": GoTo 22
                If CurLetNum = bbb Then MessageDump = MessageDump & "B": GoTo 22
                If CurLetNum = ccc Then MessageDump = MessageDump & "C": GoTo 22
                If CurLetNum = ddd Then MessageDump = MessageDump & "D": GoTo 22
                If CurLetNum = eee Then MessageDump = MessageDump & "E": GoTo 22
                If CurLetNum = fff Then MessageDump = MessageDump & "F": GoTo 22
                If CurLetNum = ggg Then MessageDump = MessageDump & "G": GoTo 22
                If CurLetNum = hhh Then MessageDump = MessageDump & "H": GoTo 22
                If CurLetNum = iii Then MessageDump = MessageDump & "I": GoTo 22
                If CurLetNum = jjj Then MessageDump = MessageDump & "J": GoTo 22
                If CurLetNum = kkk Then MessageDump = MessageDump & "K": GoTo 22
                If CurLetNum = lll Then MessageDump = MessageDump & "L":  GoTo 22
                If CurLetNum = mmm Then MessageDump = MessageDump & "M":  GoTo 22
                If CurLetNum = nnn Then MessageDump = MessageDump & "N":  GoTo 22
                If CurLetNum = ooo Then MessageDump = MessageDump & "O": GoTo 22
                If CurLetNum = ppp Then MessageDump = MessageDump & "P": GoTo 22
                If CurLetNum = qqq Then MessageDump = MessageDump & "Q": GoTo 22
                If CurLetNum = rrr Then MessageDump = MessageDump & "R": GoTo 22
                If CurLetNum = sss Then MessageDump = MessageDump & "S":  GoTo 22
                If CurLetNum = ttt Then MessageDump = MessageDump & "T":  GoTo 22
                If CurLetNum = uuu Then MessageDump = MessageDump & "U":  GoTo 22
                If CurLetNum = vvv Then MessageDump = MessageDump & "V": GoTo 22
                If CurLetNum = www Then MessageDump = MessageDump & "W": GoTo 22
                If CurLetNum = xxx Then MessageDump = MessageDump & "X": GoTo 22
                If CurLetNum = yyy Then MessageDump = MessageDump & "Y": GoTo 22
                If CurLetNum = zzz Then MessageDump = MessageDump & "Z": GoTo 22
                If CurLetNum = a1 Then MessageDump = MessageDump & "1": GoTo 22
                If CurLetNum = a2 Then MessageDump = MessageDump & "2": GoTo 22
                If CurLetNum = a3 Then MessageDump = MessageDump & "3": GoTo 22
                If CurLetNum = a4 Then MessageDump = MessageDump & "4": GoTo 22
                If CurLetNum = a5 Then MessageDump = MessageDump & "5": GoTo 22
                If CurLetNum = a6 Then MessageDump = MessageDump & "6": GoTo 22
                If CurLetNum = a7 Then MessageDump = MessageDump & "7": GoTo 22
                If CurLetNum = s8 Then MessageDump = MessageDump & "8": GoTo 22
                If CurLetNum = a9 Then MessageDump = MessageDump & "9": GoTo 22
                If CurLetNum = a0 Then MessageDump = MessageDump & "0": GoTo 22
        End If
        If ChrEvent = True Then
                    CurLetNum = Mid(MessageText, CurLetPos, 1)
                If Asc(CurLetNum) = aa Then MessageDump = MessageDump & "a":: GoTo 22
                If Asc(CurLetNum) = bb Then MessageDump = MessageDump & "b":: GoTo 22
                If Asc(CurLetNum) = cc Then MessageDump = MessageDump & "c":: GoTo 22
                If Asc(CurLetNum) = dd Then MessageDump = MessageDump & "d":: GoTo 22
                If Asc(CurLetNum) = ee Then MessageDump = MessageDump & "e":: GoTo 22
                If Asc(CurLetNum) = ff Then MessageDump = MessageDump & "f":: GoTo 22
                If Asc(CurLetNum) = gg Then MessageDump = MessageDump & "g":: GoTo 22
                If Asc(CurLetNum) = hh Then MessageDump = MessageDump & "h":: GoTo 22
                If Asc(CurLetNum) = ii Then MessageDump = MessageDump & "i":: GoTo 22
                If Asc(CurLetNum) = jj Then MessageDump = MessageDump & "j":: GoTo 22
                If Asc(CurLetNum) = kk Then MessageDump = MessageDump & "k":: GoTo 22
                If Asc(CurLetNum) = ll Then MessageDump = MessageDump & "l":: GoTo 22
                If Asc(CurLetNum) = mm Then MessageDump = MessageDump & "m":: GoTo 22
                If Asc(CurLetNum) = nn Then MessageDump = MessageDump & "n":: GoTo 22
                If Asc(CurLetNum) = oo Then MessageDump = MessageDump & "o":: GoTo 22
                If Asc(CurLetNum) = pp Then MessageDump = MessageDump & "p":: GoTo 22
                If Asc(CurLetNum) = qq Then MessageDump = MessageDump & "q":: GoTo 22
                If Asc(CurLetNum) = rr Then MessageDump = MessageDump & "r":: GoTo 22
                If Asc(CurLetNum) = ss Then MessageDump = MessageDump & "s":: GoTo 22
                If Asc(CurLetNum) = tt Then MessageDump = MessageDump & "t":: GoTo 22
                If Asc(CurLetNum) = uu Then MessageDump = MessageDump & "u":: GoTo 22
                If Asc(CurLetNum) = vv Then MessageDump = MessageDump & "v":: GoTo 22
                If Asc(CurLetNum) = ww Then MessageDump = MessageDump & "w":: GoTo 22
                If Asc(CurLetNum) = xx Then MessageDump = MessageDump & "x":: GoTo 22
                If Asc(CurLetNum) = yy Then MessageDump = MessageDump & "y":: GoTo 22
                If Asc(CurLetNum) = zz Then MessageDump = MessageDump & "z":: GoTo 22
                If Asc(CurLetNum) = sp Then MessageDump = MessageDump & " ":: GoTo 22
                If Asc(CurLetNum) = aaa Then MessageDump = MessageDump & "A":: GoTo 22
                If Asc(CurLetNum) = bbb Then MessageDump = MessageDump & "B":: GoTo 22
                If Asc(CurLetNum) = ccc Then MessageDump = MessageDump & "C":: GoTo 22
                If Asc(CurLetNum) = ddd Then MessageDump = MessageDump & "D":: GoTo 22
                If Asc(CurLetNum) = eee Then MessageDump = MessageDump & "E":: GoTo 22
                If Asc(CurLetNum) = fff Then MessageDump = MessageDump & "F":: GoTo 22
                If Asc(CurLetNum) = ggg Then MessageDump = MessageDump & "G":: GoTo 22
                If Asc(CurLetNum) = hhh Then MessageDump = MessageDump & "H":: GoTo 22
                If Asc(CurLetNum) = iii Then MessageDump = MessageDump & "I":: GoTo 22
                If Asc(CurLetNum) = jjj Then MessageDump = MessageDump & "J":: GoTo 22
                If Asc(CurLetNum) = kkk Then MessageDump = MessageDump & "K":: GoTo 22
                If Asc(CurLetNum) = lll Then MessageDump = MessageDump & "L":: GoTo 22
                If Asc(CurLetNum) = mmm Then MessageDump = MessageDump & "M":: GoTo 22
                If Asc(CurLetNum) = nnn Then MessageDump = MessageDump & "N":: GoTo 22
                If Asc(CurLetNum) = ooo Then MessageDump = MessageDump & "O":: GoTo 22
                If Asc(CurLetNum) = ppp Then MessageDump = MessageDump & "P":: GoTo 22
                If Asc(CurLetNum) = qqq Then MessageDump = MessageDump & "Q":: GoTo 22
                If Asc(CurLetNum) = rrr Then MessageDump = MessageDump & "R":: GoTo 22
                If Asc(CurLetNum) = sss Then MessageDump = MessageDump & "S":: GoTo 22
                If Asc(CurLetNum) = ttt Then MessageDump = MessageDump & "T":: GoTo 22
                If Asc(CurLetNum) = uuu Then MessageDump = MessageDump & "U":: GoTo 22
                If Asc(CurLetNum) = vvv Then MessageDump = MessageDump & "V":: GoTo 22
                If Asc(CurLetNum) = www Then MessageDump = MessageDump & "W":: GoTo 22
                If Asc(CurLetNum) = xxx Then MessageDump = MessageDump & "X":: GoTo 22
                If Asc(CurLetNum) = yyy Then MessageDump = MessageDump & "Y":: GoTo 22
                If Asc(CurLetNum) = zzz Then MessageDump = MessageDump & "Z":: GoTo 22
                If Asc(CurLetNum) = a1 Then MessageDump = MessageDump & "1":: GoTo 22
                If Asc(CurLetNum) = a2 Then MessageDump = MessageDump & "2":: GoTo 22
                If Asc(CurLetNum) = a3 Then MessageDump = MessageDump & "3":: GoTo 22
                If Asc(CurLetNum) = a4 Then MessageDump = MessageDump & "4":: GoTo 22
                If Asc(CurLetNum) = a5 Then MessageDump = MessageDump & "5":: GoTo 22
                If Asc(CurLetNum) = a6 Then MessageDump = MessageDump & "6":: GoTo 22
                If Asc(CurLetNum) = a7 Then MessageDump = MessageDump & "7":: GoTo 22
                If Asc(CurLetNum) = a8 Then MessageDump = MessageDump & "8":: GoTo 22
                If Asc(CurLetNum) = a9 Then MessageDump = MessageDump & "9":: GoTo 22
                If Asc(CurLetNum) = a0 Then MessageDump = MessageDump & "0":: GoTo 22
                 MessageDump = MessageDump & CurLetNum
                 GoTo 22
            End If

21

        MessageDump = MessageDump & CurLetNum2
            CurLetPos = CurLetPos - 2
22
              
                        
                        
                  If ChrEvent = False Then CurLetPos = CurLetPos + 2
                  
            curposfp = curposfp + 1
            
            progtotal = (curposfp / (Len(messagetext2) / 3) * 100)
            
            If progtotal > 100 Then progtotal = 100
             Form2.ProgressBar1.Value = curposfp
              
             
            
            Next CurLetPos
             MessageDump = MessageDump & Chr(13) + Chr(10)
         Loop
         Close #1
         Else
            MsgBox "No text to decrypt!", , "ERGH!"
             
     End If
Open "tmpmsg2" For Output As #1
Print #1, MessageDump
Close #1
Open "tmpmsg2" For Input As #1
Line Input #1, Hedr
Do Until EOF(1)
Line Input #1, lineoftext
        alltext = alltext & lineoftext & Chr(13) + Chr(10)
    Loop
    Close #1
    MessageDump = alltext
Kill ("tmpmsg2")
On Error GoTo errfix
PassIf = Mid(Hedr, 13, 1)
If PassIf = 1 Then
    PassLenoLen = Mid(Hedr, 14, 1)
    PassLen = Mid(Hedr, 15, PassLenoLen)
    password = Mid(Hedr, 15 + PassLenoLen, PassLen)
    
form4.Show
Exit Sub
End If

         
23
    Call FinishDec
Exit Sub
errfix:
    MsgBox "There was an error, possibly the file is not encrypted(" & Err & ")"
    form4.Hide
End Sub
Public Sub FinishDec()
  MDIForm1.ActiveForm.Label3.Caption = "Status: Done": MDIForm1.ActiveForm.Label3.Refresh
Mid(MessageDump, 1, Hedrlen) = " "
MDIForm1.ActiveForm.Text1.Text = MessageDump


         
        strSoundName = 0&
    wFlags = SND_ASYNC
    X = sndPlaySound(strSoundName, wFlags)
        
        If SOUND_ON = True Then
            strSoundName = "sounds\decryptioncomplete.wav"  'or a valid .wav file
            wFlags = SND_SYNC
            X = sndPlaySound(strSoundName, wFlags)
        End If
            
            Form2.Hide
            Close #1
            On Error Resume Next
            Kill ("tempmsg")
End Sub
Sub ChatEnc()

Dim MessageText

MessageText = Chat1.DataSnd

lengthFile = Len(Chat1.DataSnd)
     ExPro = False
     MessageDump = ""
     messagetext2 = Chat1.DataSnd
    If Chat1.DataSnd <> "" Then
    
            
                If MDIForm1.ActiveForm.Text2.Text = "" Then
                    ans = MsgBox("There was no key found to encrypt with, Generate Key Now?(NOTE: Both users MUST use the SAME key.)", vbYesNo, "No Key!")
                        If ans = vbYes Then Call KeyGen
                        If ans = vbNo Then Exit Sub
                End If
            
          If ExPro = True Then GoTo 24
         For CurLetPos = 1 To Len(MessageText)
            curposfp = curposfp + 1
            CurLetter = Mid(MessageText, CurLetPos, 1)
                If ChrEvent = False Then
                If CurLetter = "a" Then MessageDump = MessageDump & aa:  GoTo 19
                If CurLetter = "b" Then MessageDump = MessageDump & bb: GoTo 19
                If CurLetter = "c" Then MessageDump = MessageDump & cc: GoTo 19
                If CurLetter = "d" Then MessageDump = MessageDump & dd:  GoTo 19
                If CurLetter = "e" Then MessageDump = MessageDump & ee: GoTo 19
                If CurLetter = "f" Then MessageDump = MessageDump & ff: GoTo 19
                If CurLetter = "g" Then MessageDump = MessageDump & gg: GoTo 19
                If CurLetter = "h" Then MessageDump = MessageDump & hh: GoTo 19
                If CurLetter = "i" Then MessageDump = MessageDump & ii: GoTo 19
                If CurLetter = "j" Then MessageDump = MessageDump & jj: GoTo 19
                If CurLetter = "k" Then MessageDump = MessageDump & kk: GoTo 19
                If CurLetter = "l" Then MessageDump = MessageDump & ll: GoTo 19
                If CurLetter = "m" Then MessageDump = MessageDump & mm: GoTo 19
                If CurLetter = "n" Then MessageDump = MessageDump & nn: GoTo 19
                If CurLetter = "o" Then MessageDump = MessageDump & oo: GoTo 19
                If CurLetter = "p" Then MessageDump = MessageDump & pp: GoTo 19
                If CurLetter = "q" Then MessageDump = MessageDump & qq: GoTo 19
                If CurLetter = "r" Then MessageDump = MessageDump & rr: GoTo 19
                If CurLetter = "s" Then MessageDump = MessageDump & ss: GoTo 19
                If CurLetter = "t" Then MessageDump = MessageDump & tt: GoTo 19
                If CurLetter = "u" Then MessageDump = MessageDump & uu: GoTo 19
                If CurLetter = "v" Then MessageDump = MessageDump & vv: GoTo 19
                If CurLetter = "w" Then MessageDump = MessageDump & ww: GoTo 19
                If CurLetter = "x" Then MessageDump = MessageDump & xx: GoTo 19
                If CurLetter = "y" Then MessageDump = MessageDump & yy: GoTo 19
                If CurLetter = "z" Then MessageDump = MessageDump & zz: GoTo 19
                If CurLetter = " " Then MessageDump = MessageDump & sp: GoTo 19
                If CurLetter = "A" Then MessageDump = MessageDump & aaa: GoTo 19
                If CurLetter = "B" Then MessageDump = MessageDump & bbb: GoTo 19
                If CurLetter = "C" Then MessageDump = MessageDump & ccc: GoTo 19
                If CurLetter = "D" Then MessageDump = MessageDump & ddd: GoTo 19
                If CurLetter = "E" Then MessageDump = MessageDump & eee: GoTo 19
                If CurLetter = "F" Then MessageDump = MessageDump & fff: GoTo 19
                If CurLetter = "G" Then MessageDump = MessageDump & ggg: GoTo 19
                If CurLetter = "H" Then MessageDump = MessageDump & hhh: GoTo 19
                If CurLetter = "I" Then MessageDump = MessageDump & iii: GoTo 19
                If CurLetter = "J" Then MessageDump = MessageDump & jjj: GoTo 19
                If CurLetter = "K" Then MessageDump = MessageDump & kkk: GoTo 19
                If CurLetter = "L" Then MessageDump = MessageDump & lll: GoTo 19
                If CurLetter = "M" Then MessageDump = MessageDump & mmm: GoTo 19
                If CurLetter = "N" Then MessageDump = MessageDump & nnn: GoTo 19
                If CurLetter = "O" Then MessageDump = MessageDump & ooo: GoTo 19
                If CurLetter = "P" Then MessageDump = MessageDump & ppp: GoTo 19
                If CurLetter = "Q" Then MessageDump = MessageDump & qqq: GoTo 19
                If CurLetter = "R" Then MessageDump = MessageDump & rrr: GoTo 19
                If CurLetter = "S" Then MessageDump = MessageDump & sss: GoTo 19
                If CurLetter = "T" Then MessageDump = MessageDump & ttt: GoTo 19
                If CurLetter = "U" Then MessageDump = MessageDump & uuu: GoTo 19
                If CurLetter = "V" Then MessageDump = MessageDump & vvv: GoTo 19
                If CurLetter = "W" Then MessageDump = MessageDump & www: GoTo 19
                If CurLetter = "X" Then MessageDump = MessageDump & xxx: GoTo 19
                If CurLetter = "Y" Then MessageDump = MessageDump & yyy: GoTo 19
                If CurLetter = "Z" Then MessageDump = MessageDump & zzz: GoTo 19
                If CurLetter = "1" Then MessageDump = MessageDump & a1: GoTo 19
                If CurLetter = "2" Then MessageDump = MessageDump & a2: GoTo 19
                If CurLetter = "3" Then MessageDump = MessageDump & a3: GoTo 19
                If CurLetter = "4" Then MessageDump = MessageDump & a4: GoTo 19
                If CurLetter = "5" Then MessageDump = MessageDump & a5: GoTo 19
                If CurLetter = "6" Then MessageDump = MessageDump & a6: GoTo 19
                If CurLetter = "7" Then MessageDump = MessageDump & a7: GoTo 19
                If CurLetter = "8" Then MessageDump = MessageDump & a8: GoTo 19
                If CurLetter = "9" Then MessageDump = MessageDump & a9: GoTo 19
                If CurLetter = "0" Then MessageDump = MessageDump & a0: GoTo 19
                End If
                    If ChrEvent = True Then
                If CurLetter = "a" Then MessageDump = MessageDump & Chr(aa): GoTo 19
                If CurLetter = "b" Then MessageDump = MessageDump & Chr(bb): GoTo 19
                If CurLetter = "c" Then MessageDump = MessageDump & Chr(cc): GoTo 19
                If CurLetter = "d" Then MessageDump = MessageDump & Chr(dd): GoTo 19
                If CurLetter = "e" Then MessageDump = MessageDump & Chr(ee): GoTo 19
                If CurLetter = "f" Then MessageDump = MessageDump & Chr(ff): GoTo 19
                If CurLetter = "g" Then MessageDump = MessageDump & Chr(gg): GoTo 19
                If CurLetter = "h" Then MessageDump = MessageDump & Chr(hh): GoTo 19
                If CurLetter = "i" Then MessageDump = MessageDump & Chr(ii): GoTo 19
                If CurLetter = "j" Then MessageDump = MessageDump & Chr(jj): GoTo 19
                If CurLetter = "k" Then MessageDump = MessageDump & Chr(kk): GoTo 19
                If CurLetter = "l" Then MessageDump = MessageDump & Chr(ll): GoTo 19
                If CurLetter = "m" Then MessageDump = MessageDump & Chr(mm): GoTo 19
                If CurLetter = "n" Then MessageDump = MessageDump & Chr(nn): GoTo 19
                If CurLetter = "o" Then MessageDump = MessageDump & Chr(oo): GoTo 19
                If CurLetter = "p" Then MessageDump = MessageDump & Chr(pp): GoTo 19
                If CurLetter = "q" Then MessageDump = MessageDump & Chr(qq): GoTo 19
                If CurLetter = "r" Then MessageDump = MessageDump & Chr(rr): GoTo 19
                If CurLetter = "s" Then MessageDump = MessageDump & Chr(ss): GoTo 19
                If CurLetter = "t" Then MessageDump = MessageDump & Chr(tt): GoTo 19
                If CurLetter = "u" Then MessageDump = MessageDump & Chr(uu): GoTo 19
                If CurLetter = "v" Then MessageDump = MessageDump & Chr(vv): GoTo 19
                If CurLetter = "w" Then MessageDump = MessageDump & Chr(ww): GoTo 19
                If CurLetter = "x" Then MessageDump = MessageDump & Chr(xx): GoTo 19
                If CurLetter = "y" Then MessageDump = MessageDump & Chr(yy): GoTo 19
                If CurLetter = "z" Then MessageDump = MessageDump & Chr(zz): GoTo 19
                If CurLetter = " " Then MessageDump = MessageDump & Chr(sp): GoTo 19
                If CurLetter = "A" Then MessageDump = MessageDump & Chr(aaa): GoTo 19
                If CurLetter = "B" Then MessageDump = MessageDump & Chr(bbb): GoTo 19
                If CurLetter = "C" Then MessageDump = MessageDump & Chr(ccc): GoTo 19
                If CurLetter = "D" Then MessageDump = MessageDump & Chr(ddd): GoTo 19
                If CurLetter = "E" Then MessageDump = MessageDump & Chr(eee): GoTo 19
                If CurLetter = "F" Then MessageDump = MessageDump & Chr(fff): GoTo 19
                If CurLetter = "G" Then MessageDump = MessageDump & Chr(ggg): GoTo 19
                If CurLetter = "H" Then MessageDump = MessageDump & Chr(hhh): GoTo 19
                If CurLetter = "I" Then MessageDump = MessageDump & Chr(iii): GoTo 19
                If CurLetter = "J" Then MessageDump = MessageDump & Chr(jjj): GoTo 19
                If CurLetter = "K" Then MessageDump = MessageDump & Chr(kkk): GoTo 19
                If CurLetter = "L" Then MessageDump = MessageDump & Chr(lll): GoTo 19
                If CurLetter = "M" Then MessageDump = MessageDump & Chr(mmm): GoTo 19
                If CurLetter = "N" Then MessageDump = MessageDump & Chr(nnn): GoTo 19
                If CurLetter = "O" Then MessageDump = MessageDump & Chr(ooo): GoTo 19
                If CurLetter = "P" Then MessageDump = MessageDump & Chr(ppp): GoTo 19
                If CurLetter = "Q" Then MessageDump = MessageDump & Chr(qqq): GoTo 19
                If CurLetter = "R" Then MessageDump = MessageDump & Chr(rrr): GoTo 19
                If CurLetter = "S" Then MessageDump = MessageDump & Chr(sss): GoTo 19
                If CurLetter = "T" Then MessageDump = MessageDump & Chr(ttt): GoTo 19
                If CurLetter = "U" Then MessageDump = MessageDump & Chr(uuu): GoTo 19
                If CurLetter = "V" Then MessageDump = MessageDump & Chr(vvv): GoTo 19
                If CurLetter = "W" Then MessageDump = MessageDump & Chr(www): GoTo 19
                If CurLetter = "X" Then MessageDump = MessageDump & Chr(xxx): GoTo 19
                If CurLetter = "Y" Then MessageDump = MessageDump & Chr(yyy): GoTo 19
                If CurLetter = "Z" Then MessageDump = MessageDump & Chr(zzz): GoTo 19
                If CurLetter = "1" Then MessageDump = MessageDump & Chr(a1): GoTo 19
                If CurLetter = "2" Then MessageDump = MessageDump & Chr(a2): GoTo 19
                If CurLetter = "3" Then MessageDump = MessageDump & Chr(a3): GoTo 19
                If CurLetter = "4" Then MessageDump = MessageDump & Chr(a4): GoTo 19
                If CurLetter = "5" Then MessageDump = MessageDump & Chr(a5): GoTo 19
                If CurLetter = "6" Then MessageDump = MessageDump & Chr(a6): GoTo 19
                If CurLetter = "7" Then MessageDump = MessageDump & Chr(a7): GoTo 19
                If CurLetter = "8" Then MessageDump = MessageDump & Chr(a8): GoTo 19
                If CurLetter = "9" Then MessageDump = MessageDump & Chr(a9): GoTo 19
                If CurLetter = "0" Then MessageDump = MessageDump & Chr(a0): GoTo 19
                End If
                     
                MessageDump = MessageDump & CurLetter
19
                        Chat1.DataSnd = MessageDump
                        Chat1.Label5.Caption = "Encrypted"
                
                Next CurLetPos
         Else
            MsgBox "No text to encrypt!", , "ERGH!"
             
     End If
24
    
End Sub
Sub DecChat()

Dim MessageText
Dim CurLetNum As String
Dim progtotal As Integer
  ExPro = False
  MessageDump = ""
      Call KeyDecode
     If Chat1.DataRec = "" Then Exit Sub
    If Chat1.DataRec <> "" Then
           MDIForm1.ActiveForm.Label3.Caption = "Status: Decrypting...": MDIForm1.ActiveForm.Label3.Refresh
             MessageText = Chat1.DataRec
              messagetext2 = Chat1.DataRec
              
        
          If ExPro = True Then GoTo 23
         For CurLetPos = 1 To Len(MessageText)
             CurLetNum2 = Mid(MessageText, CurLetPos, 1)
            If ChrEvent = False Then
             If CurLetNum2 <> "1" And CurLetNum2 <> "2" And CurLetNum2 <> "3" And CurLetNum2 <> "4" And CurLetNum2 <> "5" And CurLetNum2 <> "6" And CurLetNum2 <> "7" And CurLetNum2 <> "8" And CurLetNum2 <> "9" And CurLetNum2 <> "0" Then
                    GoTo 21
                End If
             End If
                
            If ChrEvent = False Then
                    CurLetNum = Mid(MessageText, CurLetPos, 3)
                If CurLetNum = aa Then MessageDump = MessageDump & "a": GoTo 22
                If CurLetNum = bb Then MessageDump = MessageDump & "b": GoTo 22
                If CurLetNum = cc Then MessageDump = MessageDump & "c": GoTo 22
                If CurLetNum = dd Then MessageDump = MessageDump & "d": GoTo 22
                If CurLetNum = ee Then MessageDump = MessageDump & "e": GoTo 22
                If CurLetNum = ff Then MessageDump = MessageDump & "f": GoTo 22
                If CurLetNum = gg Then MessageDump = MessageDump & "g": GoTo 22
                If CurLetNum = hh Then MessageDump = MessageDump & "h": GoTo 22
                If CurLetNum = ii Then MessageDump = MessageDump & "i": GoTo 22
                If CurLetNum = jj Then MessageDump = MessageDump & "j": GoTo 22
                If CurLetNum = kk Then MessageDump = MessageDump & "k": GoTo 22
                If CurLetNum = ll Then MessageDump = MessageDump & "l": GoTo 22
                If CurLetNum = mm Then MessageDump = MessageDump & "m": GoTo 22
                If CurLetNum = nn Then MessageDump = MessageDump & "n": GoTo 22
                If CurLetNum = oo Then MessageDump = MessageDump & "o": GoTo 22
                If CurLetNum = pp Then MessageDump = MessageDump & "p": GoTo 22
                If CurLetNum = qq Then MessageDump = MessageDump & "q": GoTo 22
                If CurLetNum = rr Then MessageDump = MessageDump & "r": GoTo 22
                If CurLetNum = ss Then MessageDump = MessageDump & "s": GoTo 22
                If CurLetNum = tt Then MessageDump = MessageDump & "t": GoTo 22
                If CurLetNum = uu Then MessageDump = MessageDump & "u": GoTo 22
                If CurLetNum = vv Then MessageDump = MessageDump & "v": GoTo 22
                If CurLetNum = ww Then MessageDump = MessageDump & "w": GoTo 22
                If CurLetNum = xx Then MessageDump = MessageDump & "x": GoTo 22
                If CurLetNum = yy Then MessageDump = MessageDump & "y": GoTo 22
                If CurLetNum = zz Then MessageDump = MessageDump & "z": GoTo 22
                If CurLetNum = sp Then MessageDump = MessageDump & " ": GoTo 22
                If CurLetNum = aaa Then MessageDump = MessageDump & "A": GoTo 22
                If CurLetNum = bbb Then MessageDump = MessageDump & "B": GoTo 22
                If CurLetNum = ccc Then MessageDump = MessageDump & "C": GoTo 22
                If CurLetNum = ddd Then MessageDump = MessageDump & "D": GoTo 22
                If CurLetNum = eee Then MessageDump = MessageDump & "E": GoTo 22
                If CurLetNum = fff Then MessageDump = MessageDump & "F": GoTo 22
                If CurLetNum = ggg Then MessageDump = MessageDump & "G": GoTo 22
                If CurLetNum = hhh Then MessageDump = MessageDump & "H": GoTo 22
                If CurLetNum = iii Then MessageDump = MessageDump & "I": GoTo 22
                If CurLetNum = jjj Then MessageDump = MessageDump & "J": GoTo 22
                If CurLetNum = kkk Then MessageDump = MessageDump & "K": GoTo 22
                If CurLetNum = lll Then MessageDump = MessageDump & "L":  GoTo 22
                If CurLetNum = mmm Then MessageDump = MessageDump & "M":  GoTo 22
                If CurLetNum = nnn Then MessageDump = MessageDump & "N":  GoTo 22
                If CurLetNum = ooo Then MessageDump = MessageDump & "O": GoTo 22
                If CurLetNum = ppp Then MessageDump = MessageDump & "P": GoTo 22
                If CurLetNum = qqq Then MessageDump = MessageDump & "Q": GoTo 22
                If CurLetNum = rrr Then MessageDump = MessageDump & "R": GoTo 22
                If CurLetNum = sss Then MessageDump = MessageDump & "S":  GoTo 22
                If CurLetNum = ttt Then MessageDump = MessageDump & "T":  GoTo 22
                If CurLetNum = uuu Then MessageDump = MessageDump & "U":  GoTo 22
                If CurLetNum = vvv Then MessageDump = MessageDump & "V": GoTo 22
                If CurLetNum = www Then MessageDump = MessageDump & "W": GoTo 22
                If CurLetNum = xxx Then MessageDump = MessageDump & "X": GoTo 22
                If CurLetNum = yyy Then MessageDump = MessageDump & "Y": GoTo 22
                If CurLetNum = zzz Then MessageDump = MessageDump & "Z": GoTo 22
                If CurLetNum = a1 Then MessageDump = MessageDump & "1": GoTo 22
                If CurLetNum = a2 Then MessageDump = MessageDump & "2": GoTo 22
                If CurLetNum = a3 Then MessageDump = MessageDump & "3": GoTo 22
                If CurLetNum = a4 Then MessageDump = MessageDump & "4": GoTo 22
                If CurLetNum = a5 Then MessageDump = MessageDump & "5": GoTo 22
                If CurLetNum = a6 Then MessageDump = MessageDump & "6": GoTo 22
                If CurLetNum = a7 Then MessageDump = MessageDump & "7": GoTo 22
                If CurLetNum = s8 Then MessageDump = MessageDump & "8": GoTo 22
                If CurLetNum = a9 Then MessageDump = MessageDump & "9": GoTo 22
                If CurLetNum = a0 Then MessageDump = MessageDump & "0": GoTo 22
        End If
        If ChrEvent = True Then
                    CurLetNum = Mid(MessageText, CurLetPos, 1)
                If Asc(CurLetNum) = aa Then MessageDump = MessageDump & "a":: GoTo 22
                If Asc(CurLetNum) = bb Then MessageDump = MessageDump & "b":: GoTo 22
                If Asc(CurLetNum) = cc Then MessageDump = MessageDump & "c":: GoTo 22
                If Asc(CurLetNum) = dd Then MessageDump = MessageDump & "d":: GoTo 22
                If Asc(CurLetNum) = ee Then MessageDump = MessageDump & "e":: GoTo 22
                If Asc(CurLetNum) = ff Then MessageDump = MessageDump & "f":: GoTo 22
                If Asc(CurLetNum) = gg Then MessageDump = MessageDump & "g":: GoTo 22
                If Asc(CurLetNum) = hh Then MessageDump = MessageDump & "h":: GoTo 22
                If Asc(CurLetNum) = ii Then MessageDump = MessageDump & "i":: GoTo 22
                If Asc(CurLetNum) = jj Then MessageDump = MessageDump & "j":: GoTo 22
                If Asc(CurLetNum) = kk Then MessageDump = MessageDump & "k":: GoTo 22
                If Asc(CurLetNum) = ll Then MessageDump = MessageDump & "l":: GoTo 22
                If Asc(CurLetNum) = mm Then MessageDump = MessageDump & "m":: GoTo 22
                If Asc(CurLetNum) = nn Then MessageDump = MessageDump & "n":: GoTo 22
                If Asc(CurLetNum) = oo Then MessageDump = MessageDump & "o":: GoTo 22
                If Asc(CurLetNum) = pp Then MessageDump = MessageDump & "p":: GoTo 22
                If Asc(CurLetNum) = qq Then MessageDump = MessageDump & "q":: GoTo 22
                If Asc(CurLetNum) = rr Then MessageDump = MessageDump & "r":: GoTo 22
                If Asc(CurLetNum) = ss Then MessageDump = MessageDump & "s":: GoTo 22
                If Asc(CurLetNum) = tt Then MessageDump = MessageDump & "t":: GoTo 22
                If Asc(CurLetNum) = uu Then MessageDump = MessageDump & "u":: GoTo 22
                If Asc(CurLetNum) = vv Then MessageDump = MessageDump & "v":: GoTo 22
                If Asc(CurLetNum) = ww Then MessageDump = MessageDump & "w":: GoTo 22
                If Asc(CurLetNum) = xx Then MessageDump = MessageDump & "x":: GoTo 22
                If Asc(CurLetNum) = yy Then MessageDump = MessageDump & "y":: GoTo 22
                If Asc(CurLetNum) = zz Then MessageDump = MessageDump & "z":: GoTo 22
                If Asc(CurLetNum) = sp Then MessageDump = MessageDump & " ":: GoTo 22
                If Asc(CurLetNum) = aaa Then MessageDump = MessageDump & "A":: GoTo 22
                If Asc(CurLetNum) = bbb Then MessageDump = MessageDump & "B":: GoTo 22
                If Asc(CurLetNum) = ccc Then MessageDump = MessageDump & "C":: GoTo 22
                If Asc(CurLetNum) = ddd Then MessageDump = MessageDump & "D":: GoTo 22
                If Asc(CurLetNum) = eee Then MessageDump = MessageDump & "E":: GoTo 22
                If Asc(CurLetNum) = fff Then MessageDump = MessageDump & "F":: GoTo 22
                If Asc(CurLetNum) = ggg Then MessageDump = MessageDump & "G":: GoTo 22
                If Asc(CurLetNum) = hhh Then MessageDump = MessageDump & "H":: GoTo 22
                If Asc(CurLetNum) = iii Then MessageDump = MessageDump & "I":: GoTo 22
                If Asc(CurLetNum) = jjj Then MessageDump = MessageDump & "J":: GoTo 22
                If Asc(CurLetNum) = kkk Then MessageDump = MessageDump & "K":: GoTo 22
                If Asc(CurLetNum) = lll Then MessageDump = MessageDump & "L":: GoTo 22
                If Asc(CurLetNum) = mmm Then MessageDump = MessageDump & "M":: GoTo 22
                If Asc(CurLetNum) = nnn Then MessageDump = MessageDump & "N":: GoTo 22
                If Asc(CurLetNum) = ooo Then MessageDump = MessageDump & "O":: GoTo 22
                If Asc(CurLetNum) = ppp Then MessageDump = MessageDump & "P":: GoTo 22
                If Asc(CurLetNum) = qqq Then MessageDump = MessageDump & "Q":: GoTo 22
                If Asc(CurLetNum) = rrr Then MessageDump = MessageDump & "R":: GoTo 22
                If Asc(CurLetNum) = sss Then MessageDump = MessageDump & "S":: GoTo 22
                If Asc(CurLetNum) = ttt Then MessageDump = MessageDump & "T":: GoTo 22
                If Asc(CurLetNum) = uuu Then MessageDump = MessageDump & "U":: GoTo 22
                If Asc(CurLetNum) = vvv Then MessageDump = MessageDump & "V":: GoTo 22
                If Asc(CurLetNum) = www Then MessageDump = MessageDump & "W":: GoTo 22
                If Asc(CurLetNum) = xxx Then MessageDump = MessageDump & "X":: GoTo 22
                If Asc(CurLetNum) = yyy Then MessageDump = MessageDump & "Y":: GoTo 22
                If Asc(CurLetNum) = zzz Then MessageDump = MessageDump & "Z":: GoTo 22
                If Asc(CurLetNum) = a1 Then MessageDump = MessageDump & "1":: GoTo 22
                If Asc(CurLetNum) = a2 Then MessageDump = MessageDump & "2":: GoTo 22
                If Asc(CurLetNum) = a3 Then MessageDump = MessageDump & "3":: GoTo 22
                If Asc(CurLetNum) = a4 Then MessageDump = MessageDump & "4":: GoTo 22
                If Asc(CurLetNum) = a5 Then MessageDump = MessageDump & "5":: GoTo 22
                If Asc(CurLetNum) = a6 Then MessageDump = MessageDump & "6":: GoTo 22
                If Asc(CurLetNum) = a7 Then MessageDump = MessageDump & "7":: GoTo 22
                If Asc(CurLetNum) = a8 Then MessageDump = MessageDump & "8":: GoTo 22
                If Asc(CurLetNum) = a9 Then MessageDump = MessageDump & "9":: GoTo 22
                If Asc(CurLetNum) = a0 Then MessageDump = MessageDump & "0":: GoTo 22
                 MessageDump = MessageDump & CurLetNum
                 GoTo 22
            End If

21

        MessageDump = MessageDump & CurLetNum2
            CurLetPos = CurLetPos - 2
22
              
                        
                        
                  If ChrEvent = False Then CurLetPos = CurLetPos + 2
                  
            curposfp = curposfp + 1
            
             
            
            Next CurLetPos
         
         Else
            MsgBox "No text to decrypt!", , "ERGH!"
             
     End If


         
23
    
    Chat1.DataRec = MessageDump
    Chat1.Label5.Caption = "Decrypted"
Exit Sub
errfix:
    MsgBox "There was an error, possibly the file is not encrypted(" & Err & ")"
End Sub
