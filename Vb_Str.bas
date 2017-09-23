Attribute VB_Name = "Vb_Str"
Option Explicit
Option Compare Database

Type S1S2
    S1 As String
    S2 As String
End Type
Type Map
    Sy1() As String
    Sy2() As String
End Type

Function AlignL$(S, W)
Dim L%
If IsNull(S) Then
    L = 0
Else
    L = Len(S)
End If
If W >= L Then
    AlignL = S & Space(W - L)
Else
    If W > 2 Then
        AlignL = Left(S, W - 2) + ".."
    Else
        AlignL = Left(S, W)
    End If
End If
End Function
Function StrAppSpc$(S, Optional App = "")
StrAppSpc = StrApp(S, App, " ")
End Function
Function AlignR$(S, W%)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
End Function

Function Brk(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then Err.Raise "Brk: Str[" & S & "] does not contains Sep[" & Sep & "]"
Brk = BrkAt(S, P, Len(Sep), NoTrim)
End Function

Function Brk1(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
Brk1 = Brk1__(S, P, Sep, NoTrim)
End Function

Function Brk1Rev(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(S, Sep)
Brk1Rev = Brk1__(S, P, Sep, NoTrim)
End Function

Function Brk2(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then
    Dim O As S1S2
    If NoTrim Then
        O.S2 = S
    Else
        O.S2 = Trim(S)
    End If
    Brk2 = O
    Exit Function
End If
Brk2 = BrkAt(S, P, Len(Sep), NoTrim)
End Function

Function BrkAt(S, P&, SepLen%, Optional NoTrim As Boolean) As S1S2
Dim O As S1S2
With O
    If NoTrim Then
        .S1 = Left(S, P - 1)
        .S2 = Mid(S, P + SepLen)
    Else
        .S1 = Trim(Left(S, P - 1))
        .S2 = Trim(Mid(S, P + SepLen))
    End If
End With
BrkAt = O
End Function

Function BrkBoth(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then
    Dim O As S1S2
    If NoTrim Then
        O.S1 = S
    Else
        O.S1 = Trim(S)
    End If
    O.S2 = O.S1
    BrkBoth = O
    Exit Function
End If
BrkBoth = BrkAt(S, P, Len(Sep), NoTrim)
End Function

Function BrkMapStr(MapStr$) As Map
Dim Ay$(): Ay = Split(MapStr, "|")
Dim Ay1$(), Ay2$()
    Dim I
    For Each I In Ay
        With BrkBoth(I, ":")
            Push Ay1, .S1
            Push Ay2, .S2
        End With
    Next
Dim O As Map
    O.Sy1 = Ay1
    O.Sy2 = Ay2
BrkMapStr = O
End Function

Function BrkQuote(QuoteStr$) As S1S2
Dim L%: L = Len(QuoteStr)
Dim O As S1S2
Select Case L
Case 0:
Case 1
    O.S1 = QuoteStr
    O.S2 = O.S1
Case 2
    O.S1 = Left(QuoteStr, 1)
    O.S2 = Right(QuoteStr, 1)
Case Else
    Dim P%
    If InStr(QuoteStr, "*") > 0 Then
        O = Brk(QuoteStr, "*", NoTrim:=True)
    End If
End Select
BrkQuote = O
End Function

Function BrkRev(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(S, Sep)
If P = 0 Then Err.Raise "BrkRev: Str[" & S & "] does not contains Sep[" & Sep & "]"
BrkRev = BrkAt(S, P, Len(Sep), NoTrim)
End Function

Function DftStr$(S, DftVal)
DftStr = Dft(S, DftVal)
End Function

Function FmtMacro$(MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
FmtMacro = FmtMacroAv(MacroStr, Av)
End Function

Function FmtMacroAv$(MacroStr$, Av())
Dim Ay$(): Ay = MacroStrNy(MacroStr)
Dim O$: O = MacroStr
Dim J%, I
For Each I In Ay
    O = Replace(O, I, Av(J))
    J = J + 1
Next
FmtMacroAv = O
End Function

Function FmtMacroDic$(MacroStr$, Dic As Dictionary)
Dim Ay$(): Ay = MacroStrNy(MacroStr)
If Not AyIsEmpty(Ay) Then
    Dim O$: O = MacroStr
    Dim I, K$
    For Each I In Ay
        K = RmvFstLasChr(I)
        If Dic.Exists(K) Then
            O = Replace(O, I, Dic(K))
        End If
    Next
End If
FmtMacroDic = O
End Function
Function FmtQQVBar$(QQStr$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQVBar = RplVBar(FmtQQAv(QQStr, Av))
End Function

Function FmtQQ$(QQStr$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQStr, Av)
End Function

Function FmtQQAv$(QQStr$, Av)
If AyIsEmpty(Av) Then FmtQQAv = QQStr: Exit Function
Dim O$
    Dim I, NeedUnEsc As Boolean
    O = QQStr
    For Each I In Av
        If InStr(I, "?") > 0 Then
            NeedUnEsc = True
            I = Replace(I, "?", Chr(255))
        End If
        O = Replace(O, "?", I, Count:=1)
    Next
    If NeedUnEsc Then O = Replace(O, Chr(255), "?")
FmtQQAv = O
End Function

Function FstChr$(S)
FstChr = Left(S, 1)
End Function

Function HasSubStr(S, SubStr) As Boolean
HasSubStr = InStr(S, SubStr) > 0
End Function

Function InstrN&(S, SubStr, N%)
Dim P&, J%
For J = 1 To N
    P = InStr(P + 1, S, SubStr)
    If P = 0 Then Exit Function
Next
InstrN = P
End Function

Function IsDigit(C) As Boolean
IsDigit = "0" <= C And C <= "9"
End Function

Function IsLetter(C) As Boolean
Dim C1$: C1 = UCase(C)
IsLetter = ("A" <= C1 And C1 <= "Z")
End Function

Function IsNmChr(C) As Boolean
IsNmChr = True
If IsLetter(C) Then Exit Function
If C = "_" Then Exit Function
If IsDigit(C) Then Exit Function
IsNmChr = False
End Function

Function IsOneOf(S, Ay) As Boolean
Dim I
For Each I In Ay
    If S = I Then IsOneOf = True: Exit Function
Next
End Function

Function IsPfx(S, Pfx) As Boolean
IsPfx = (Left(S, Len(Pfx)) = Pfx)
End Function

Function IsSfx(S, Sfx) As Boolean
IsSfx = (Right(S, Len(Sfx)) = Sfx)
End Function

Function JnComma$(Ay)
JnComma = Join(Ay, ",")
End Function

Function JnCrLf$(Ay)
JnCrLf = Join(Ay, vbCrLf)
End Function
Function JnDblCrLf$(Ay)
JnDblCrLf = Join(Ay, vbCrLf & vbCrLf)
End Function

Function JnSpc$(Ay)
JnSpc = Join(Ay, " ")
End Function

Function LasChr$(S)
LasChr = Right(S, 1)
End Function

Function LasLin$(S)
LasLin = AyLasEle(SplitCrLf(S))
End Function

Function LinesLasLin$(Lines)
LinesLasLin = AyLasEle(SplitCrLf(Lines))
End Function

Function LinesLinCnt&(Lines)
LinesLinCnt = Sz(SplitCrLf(Lines))
End Function

Function LvsJnComma$(Lvs$)
LvsJnComma = JnComma(SplitLvs(Lvs))
End Function

Function LvsJnQuoteComma$(Lvs$)
LvsJnQuoteComma = JnComma(AyQuote(SplitLvs(Lvs), "'"))
End Function

Function MacroStrNy(MacroStr$, Optional ExclBkt As Boolean) As String()
Dim Ay$(): Ay = Split(MacroStr, "{")
Dim O$(), J%
For J = 1 To UB(Ay)
    Push O, TakBef(Ay(J), "}")
Next
If Not ExclBkt Then
    O = AyAddPfxSfx(O, "{", "}")
End If
MacroStrNy = O
End Function

Function MapDic(A As Map) As Dictionary
Dim J&, O As New Dictionary
With A
    Dim U&: U = UB(.Sy1)
    For J = 0 To U
        O.Add .Sy1(J), .Sy2(J)
    Next
End With
Set MapDic = O
End Function

Function ParseTerm$(OStr)
OStr = Trim(OStr)
ParseTerm = FstTerm(OStr)
OStr = RmvFstTerm(OStr)
End Function

Function Quote$(S, QuoteStr$)
With BrkQuote(QuoteStr)
    Quote = .S1 & CStr(S) & .S2
End With
End Function

Function Rmv2Dash$(Lin)
Rmv2Dash = RTrim(RmvAft(Lin, "--"))
End Function

Function Rmv3Dash$(Lin)
Rmv3Dash = RTrim(RmvAft(Lin, "---"))
End Function

Function RmvAft$(S, Sep)
RmvAft = Brk1(S, Sep, NoTrim:=True).S1
End Function

Function RmvDblSpc$(S)
Dim O$: O = S
While HasSubStr(O, "  ")
    O = Replace(O, "  ", " ")
Wend
RmvDblSpc = O
End Function

Function RmvFstLasChr$(I)
RmvFstLasChr = RmvFstChr(RmvLasChr(I))
End Function

Function RmvFstNChr$(S, Optional N% = 1)
RmvFstNChr = Mid(S, N + 1)
End Function

Function RmvFstTerm$(S)
RmvFstTerm = Brk1(Trim(S), " ").S2
End Function

Function RmvLasChr$(S)
RmvLasChr = RmvLasNChr(S)
End Function

Function RmvLasNChr$(S, Optional N% = 1)
RmvLasNChr = Left(S, Len(S) - 1)
End Function

Function RmvPfx$(S, Pfx)
Dim L%: L = Len(Pfx)
If Left(S, L) = Pfx Then
    RmvPfx = Mid(S, L + 1)
Else
    RmvPfx = S
End If
End Function

Function RmvSfx$(S, Sfx)
Dim L%: L = Len(Sfx)
If Right(S, L) = Sfx Then
    RmvSfx = Left(S, Len(S) - L)
Else
    RmvSfx = S
End If
End Function

Function RplVBar$(S)
RplVBar = Replace(S, "|", vbCrLf)
End Function

Function S1S2(S1, S2) As S1S2
S1S2.S1 = S1
S1S2.S2 = S2
End Function

Function S1S2Ay(Ay1, Ay2) As S1S2()
If AyIsEmpty(Ay1) Then Exit Function
Dim U&: U = UB(Ay2)
If U <> UB(Ay1) Then Stop
Dim O() As S1S2
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = S1S2(Ay1(J), Ay2(J))
Next
S1S2Ay = O
End Function

Function S1S2AyDrs(A() As S1S2) As Drs
S1S2AyDrs.Fny = SplitSpc("S1 S2")
S1S2AyDrs.Dry = S1S2AyDry(A)
End Function

Function S1S2AyDry(A() As S1S2) As Variant()
Dim O()
Dim J%
For J = 0 To S1S2UB(A)
    With A(J)
        Push O, Array(.S1, .S2)
    End With
Next
S1S2AyDry = O
End Function

Function S1S2AyS1Ay(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To S1S2UB(A)
    Push O, A(J).S1
Next
S1S2AyS1Ay = O
End Function

Function S1S2AyS2Ay(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To S1S2UB(A)
    Push O, A(J).S2
Next
S1S2AyS2Ay = O
End Function

Sub S1S2Push(O() As S1S2, I As S1S2)
Dim N&: N = S1S2Sz(O)
ReDim Preserve O(N)
O(N) = I
End Sub

Function S1S2Sz&(A() As S1S2)
On Error Resume Next
S1S2Sz = UBound(A) + 1
End Function

Function S1S2UB&(A() As S1S2)
S1S2UB = S1S2Sz(A) - 1
End Function

Function SpcEsc$(S)
If InStr(S, "~") > 0 Then Debug.Print "SpcEsc: Warning: escaping a string-with-space is found with a [~].  The [~] before escape will be changed to space after unescape"
SpcEsc = Replace(S, " ", "~")
End Function

Function SpcUnE$(S)
SpcUnE = Replace(S, "~", " ")
End Function

Function SplitCrLf(S) As String()
SplitCrLf = Split(S, vbCrLf)
End Function

Function SplitLvs(Lvs) As String()
SplitLvs = Split(RmvDblSpc(Trim(Lvs)), " ")
End Function

Function SplitSpc(S) As String()
SplitSpc = Split(S, " ")
End Function

Function SplitVBar(S) As String()
SplitVBar = Split(S, "|")
End Function

Function StrAppCrLf$(S, App)
StrAppCrLf = StrApp(S, App, vbCrLf)
End Function

Function StrApp$(S, App, Sep)
If S = "" Then
    StrApp = App
Else
    StrApp = S & Sep & App
End If
End Function
Function StrAppVBar$(S, App)
StrAppVBar = StrApp(S, App, "|")
End Function

Sub StrBrw(S, Optional Fnn)
Dim T$: T = TmpFt("StrBrw", Fnn)
StrWrt S, T
FtBrw T
End Sub

Function StrDup$(N%, S)
Dim O$, J%
For J = 0 To N - 1
    O = O & S
Next
StrDup = O
End Function

Function SubStrCnt&(S, SubStr)
Dim P&: P = 1
Dim L%: L = Len(SubStr)
Dim O%
While P > 0
    P = InStr(P, S, SubStr)
    If P = 0 Then SubStrCnt = O: Exit Function
    O = O + 1
    P = P + L
Wend
SubStrCnt = O
End Function

Function TakAft$(S, Sep, Optional NoTrim As Boolean)
TakAft = Brk1(S, Sep, NoTrim).S2
End Function

Function TakAftRev$(S, Sep, Optional NoTrim As Boolean)
TakAftRev = Brk1Rev(S, Sep, NoTrim).S2
End Function

Function TakBef$(S, Sep, Optional NoTrim As Boolean)
TakBef = Brk2(S, Sep, NoTrim).S1
End Function

Function TakBefRev$(S, Sep, Optional NoTrim As Boolean)
TakBefRev = BrkRev(S, Sep, NoTrim).S1
End Function

Function TakBet$(S, S1, S2, Optional NoTrim As Boolean, Optional InclMarker As Boolean)
With Brk1(S, S1, NoTrim)
    If .S2 = "" Then Exit Function
    Dim O$: O = Brk1(.S2, S2, NoTrim).S1
    If InclMarker Then O = S1 & O & S2
    TakBet = O
End With
End Function

Function TmpFfn(Ext$, Optional Fdr, Optional Fnn)
Dim mFnn$
    mFnn = IIf(IsEmpty(Fnn), TmpNm, Fnn)
TmpFfn = TmpPth(Fdr) & mFnn & Ext
End Function


Private Function Brk1__(S, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    Dim O As S1S2
    If NoTrim Then
        O.S1 = S
    Else
        O.S1 = Trim(S)
    End If
    Brk1__ = O
    Exit Function
End If
Brk1__ = BrkAt(S, P, Len(Sep), NoTrim)
End Function

Private Sub Brk1Rev__Tst()
Dim S1$, S2$, ExpS1$, ExpS2$, S$
S = "aa --- bb --- cc"
ExpS1 = "aa --- bb"
ExpS2 = "cc"
With Brk1Rev(S, "---")
    S1 = .S1
    S2 = .S2
End With
Debug.Assert S1 = ExpS1
Debug.Assert S2 = ExpS2
End Sub

Private Sub BrkMapStr__Tst()
Dim MapStr$
MapStr = "aa:bb|cc|dd:ee"
Dim Act As Map: Act = BrkMapStr(MapStr)
Dim Exp1$(): Exp1 = SplitSpc("aa cc dd"): AyAssertEq Exp1, Act.Sy1
Dim Exp2$(): Exp2 = SplitSpc("aa cc dd"): AyAssertEq Exp2, Act.Sy1
End Sub

Private Sub InstrN__Tst()
Dim Act&, Exp&, S, SubStr, N%

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 1
Exp = 1
Act = InstrN(S, SubStr, N)
Debug.Assert Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 2
Exp = 6
Act = InstrN(S, SubStr, N)
Debug.Assert Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 3
Exp = 11
Act = InstrN(S, SubStr, N)
Debug.Assert Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 4
Exp = 0
Act = InstrN(S, SubStr, N)
Debug.Assert Exp = Act
End Sub

Private Sub RmvFstTerm__Tst()
Debug.Assert RmvFstTerm("  df dfdf  ") = "dfdf"
End Sub

Function SubStrCnt__Tst()
Debug.Assert SubStrCnt("aaaa", "aa") = 2
Debug.Assert SubStrCnt("aaaa", "a") = 4
End Function

Private Sub TakBet__Tst()
Const S1$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??"
Const S2$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX"
Debug.Assert TakBet(S1, "DATABASE=", ";") = "??"
Debug.Assert TakBet(S2, "DATABASE=", ";") = "??"
End Sub

Sub Tst()
BrkMapStr__Tst
End Sub
