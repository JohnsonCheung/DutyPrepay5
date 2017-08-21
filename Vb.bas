Attribute VB_Name = "Vb"
Option Compare Database
Option Explicit
Public Fso As New FileSystemObject
Public Tmp As New Tmp
Public Assert As New Assert
Public Ado As New Ado
Public Ide As New Ide
Public Tst As New Tst
Type AqlVal
    Sql As String
    Cn As ADODB.Connection
End Type
Enum eSimTy
    eTxt
    eNbr
    eDte
    eLgc
    eOth
End Enum
Function CurPj() As VBProject
Set CurPj = Application.VBE.ActiveVBProject
End Function
Function SimTyInsSqlValPhraseTp$(A As eSimTy)
Dim O$
Select Case A
Case eTxt: O = "'?'"
Case eNbr, eLgc: O = "?"
Case eDte: O = "#?#"
Case Else
    Err.Raise 1, , "Given SimTy[" & SimTyStr(A) & "] should be [eTxt eNbr eDte eLgc]"
End Select
SimTyInsSqlValPhraseTp = O
End Function
Function EmptyLngAy() As Ay
Dim O&()
Set EmptyLngAy = Ay(O)
End Function
Function EmptyStrAy() As Ay
Dim O$()
Set EmptyLngAy = Ay(O)
End Function
Function SimTyStr$(A As eSimTy)
Dim O$
Select Case A
Case eSimTy.eDte: O = "eDte"
Case eSimTy.eLgc: O = "eLgc"
Case eSimTy.eNbr: O = "eNbr"
Case eSimTy.eTxt: O = "eTxt"
Case eSimTy.eOth: O = "eOth"
Case Else: Stop
End Select
SimTyStr = O
End Function
Function VbTyToStr$(A As VbVarType)

End Function
Function VbTySimTy(A As VbVarType) As eSimTy
Select Case A
Case VbVarType.vbDate
Case Else
End Select
End Function

Function VbTyInsSqlValPhraseTp$(A As VbVarType)
VbTyInsSqlValPhraseTp = SimTyInsSqlValPhraseTp(VbTySimTy(A))
End Function

Function AyIsEmpty(Ay) As Boolean
AyIsEmpty = (Sz(Ay) = 0)
End Function

Function CurFb$()
CurFb = CurrentDb.Name
End Function

Function CurPth$()
CurPth = Nw.Ffn(CurFb).Pth
End Function

Function DtaDb() As Dao.Database
Set DtaDb = DBEngine.OpenDatabase(DtaFb)
End Function

Function DtaFb$()
DtaFb = Nw.Ffn(CurFb).AddFnSfxX("_Data").RplExt(".mdb")
End Function
Sub AyBrw(Ay)

End Sub
Function IsNmNeedQuote(S) As Boolean
IsNmNeedQuote = True
If HasSubStr(S, " ") Then Exit Function
If HasSubStr(S, "#") Then Exit Function
If HasSubStr(S, ".") Then Exit Function
IsNmNeedQuote = False
End Function
Sub AsrtEqAy(Ay1, Ay2)
Stop
End Sub
Sub PushNonEmpty(Ay, I)
If Vb.IsEmpty(I) Then Exit Sub
Push Ay, I
End Sub

Function PermitImpPth$()
Dim O$: O = CurPth & "Import - Permit\"
Nw.Pth(O).Ens
PermitImpPth = O
End Function
Function PermitImpPthX() As Pth
Set PermitImpPthX = Nw.Pth(PermitImpPth)
End Function

Function WrkPth$()
WrkPth = CurPth & "WorkingDir\"
End Function

Function DftDry(A As Dry) As Dry
If IsNothing(A) Then
    Set DftDry = New Dry
Else
    Set DftDry = A
End If
End Function
Function EmptySy() As String()
End Function
Function Tst_ResPth$()
Tst_ResPth = Nw.Pj.SrcPth & "TstRes\"
End Function

Sub PushAy(OAy, Ay)
If AyIsEmpty(Ay) Then Exit Sub
Dim V
For Each V In Ay
    Push OAy, V
Next
End Sub
Sub Push(Ay, I)
Dim N&: N = Sz(Ay)
ReDim Preserve Ay(N)
Asg I, Ay(N)
End Sub
Function Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Function
Function UB&(Ay)
On Error Resume Next
UB = Sz(Ay) - 1
End Function
Function VdtMth(M As Byte) As Boolean
If M > 12 Or M < 1 Then MsgBox "must between 1 and 12": VdtMth = True
End Function

Function VdtYr(Y As Byte) As Boolean
If Y = 0 Then MsgBox "Year cannot be ": VdtYr = True
End Function

Function RelItm(Nm$, Chd$(), Optional Dta) As RelItm
Dim O As RelItm
With O
    .Nm = Nm
    .Chd = Chd
    Asg Dta, .Dta
End With
RelItm = O
End Function
Function RelItmLvs(Nm$, ChdLvs$, Optional Dta) As RelItm
RelItmLvs = RelItm(Nm, SplitLvs(ChdLvs), Dta)
End Function


Function DaoTyStr$(T As Dao.DataTypeEnum)
Dim O$
Select Case T
Case Dao.DataTypeEnum.dbBoolean: O = "Boolean"
Case Dao.DataTypeEnum.dbDouble: O = "Double"
Case Dao.DataTypeEnum.dbText: O = "Text"
Case Dao.DataTypeEnum.dbDate: O = "Date"
Case Dao.DataTypeEnum.dbByte: O = "Byte"
Case Dao.DataTypeEnum.dbInteger: O = "Int"
Case Dao.DataTypeEnum.dbLong: O = "Long"
Case Dao.DataTypeEnum.dbDouble: O = "Doubld"
Case Dao.DataTypeEnum.dbDate: O = "Date"
Case Dao.DataTypeEnum.dbDecimal: O = "Decimal"
Case Dao.DataTypeEnum.dbCurrency: O = "Currency"
Case Dao.DataTypeEnum.dbSingle: O = "Single"

Case Else: Stop
End Select
DaoTyStr = O
End Function

Function BrkRev(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then Err.Raise "BrkRev: Str[" & S & "] does not contains Sep[" & Sep & "]"
BrkRev = BrkAt(S, P, Len(Sep), NoTrim)
End Function
Function Brk(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then Err.Raise "Brk: Str[" & S & "] does not contains Sep[" & Sep & "]"
Set Brk = BrkAt(S, P, Len(Sep), NoTrim)
End Function

Function Brk1(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
Brk1 = Brk1__(S, P, Sep, NoTrim)
End Function
Private Function Brk1__(S, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    Dim O As New S1S2
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
Function Brk1Rev(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(S, Sep)
Brk1Rev = Brk1__(S, P, Sep, NoTrim)
End Function

Function Brk2(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then
    Dim O As New S1S2
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
Dim O As New S1S2
With O
    If NoTrim Then
        .S1 = Left(S, P - 1)
        .S2 = Mid(S, P + SepLen)
    Else
        .S1 = Trim(Left(S, P - 1))
        .S2 = Trim(Mid(S, P + SepLen))
    End If
End With
Set BrkAt = O
End Function

Function BrkBoth(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then
    Dim O As New S1S2
    If NoTrim Then
        O.S1 = S
    Else
        O.S1 = Trim(S)
    End If
    O.S2 = O.S1
    BrkBoth = O
    Exit Function
End If
Set BrkBoth = BrkAt(S, P, Len(Sep), NoTrim)
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
Set BrkMapStr = Nw.Map(Ay1, Ay2)
End Function

Function BrkQuote(QuoteStr$) As S1S2
Dim L%: L = Len(QuoteStr)
Dim O As New S1S2
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
        Set O = Brk(QuoteStr, "*", NoTrim:=True)
    End If
End Select
Set BrkQuote = O
End Function

Function MapDic(A As Map) As Dic
Dim J&, O As New Dictionary
With A
    Dim U&: U = UB(.Sy1)
    For J = 0 To U
        O.Add .Sy1()(J), .Sy2()(J)
    Next
End With
Set MapDic = Nw.Dic(O)
End Function
Function NewWsX(Optional WsNm$, Optional Vis As Boolean) As Ws
Set NewWsX = Nw.Ws(NewWs(WsNm, Vis))
End Function
Function NewA1(Optional WsNm$, Optional Vis As Boolean) As Range
Set NewA1 = NewWs(WsNm, Vis).Range("A1")
End Function
Function NewWs(Optional WsNm$, Optional Vis As Boolean) As Worksheet
Dim Wb As Wb
Set Wb = NewWbX
Wb.DltWs "Sheet2"
Wb.DltWs "Sheet3"
If WsNm <> "" Then Wb.Ws("Sheet1").Name = WsNm
Set NewWs = Wb.Ws(1)
If Vis Then Wb.Vis
End Function

Private Sub BrkMapStr__Tst()
Dim MapStr$
MapStr = "aa:bb|cc|dd:ee"
Dim Act As Map: Act = BrkMapStr(MapStr)
Dim Exp1$(): Exp1 = SplitSpc("aa cc dd"): Ay(Exp1).AsrtEq Act.Sy1
Dim Exp2$(): Exp2 = SplitSpc("aa cc dd"): Ay(Exp1).AsrtEq Act.Sy1
End Sub

Function Xls() As Excel.Application
Static X As Excel.Application
On Error GoTo XX
Dim A$: A = X.Name
Set Xls = X
Exit Function
XX:
Set X = New Excel.Application
Set Xls = X
End Function
Function DftFb$(Optional Fb)
If IsMissing(Fb) Then
    DftFb = CurDb.Nm
Else
    DftFb = Fb
End If
End Function
Function IsPth(P) As Boolean
IsPth = Dir(P, vbDirectory) <> ""
End Function
Sub AsrtIsSy(V)
If Not IsSy(V) Then Stop
End Sub
Function IsSy(V) As Boolean
IsSy = VarType(V) = vbArray + vbString
End Function
Function IsAy(V) As Boolean
IsAy = VarType(V) And vbArray
End Function
Sub AsrtIsEq(V1, V2)
If V1 <> V2 Then Stop
End Sub
Sub AsrtIsAy(V)
If Not IsArray(V) Then Stop
End Sub
Sub AsrtIsStr(V)
If Not IsStr(V) Then Stop
End Sub

Function NewWb(Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = Xls.Workbooks.Add
If Vis Then O.Visible = True
Set NewWb = O
End Function
Function NewWbX(Optional Vis As Boolean) As Wb
Set NewWbX = Wb(NewWb(Vis))
End Function

Function TstResPth$()
TstResPth = Pj.SrcPth & "TstRes\"
End Function

Sub TstResPthBrw()
Pth(TstResPth).Brw
End Sub

Function IdxCnt(Idx&, Cnt&) As IdxCnt
IdxCnt.Idx = Idx
IdxCnt.Cnt = Cnt
End Function
Function DrLin$(Dr, Wdt%())
Dim UDr%
    UDr = UB(Dr)
Dim O$()
    Dim U1%: U1 = UB(Wdt)
    ReDim O(U1)
    Dim W, V
    Dim J%
    J = 0
    For Each W In Wdt
        If UDr >= J Then V = Dr(J) Else V = ""
        O(J) = DrLin__V(V, W)
        J = J + 1
    Next
DrLin = Quote(Join(O, " | "), "| * |")
End Function

Function JnComma$(Ay)
JnComma = Join(Ay, ",")
End Function

Function JnCrLf(Ay)
JnCrLf = Join(Ay, vbCrLf)
End Function

Function JnSpc(Ay)
JnSpc = Join(Ay, " ")
End Function
Function DftMdNm$(Nm$)
If Nm = "" Then
    DftMdNm = Md.Nm
Else
    DftMdNm = Nm
End If
End Function
Sub AsrtIsPth(P)
Pth(P).AsstIsExist
End Sub
Function DftPj(Optional A As VBProject) As VBProject
If IsNothing(A) Then
    Set DftPj = Application.VBE.ActiveVBProject
Else
    Set DftPj = A
End If
End Function
Function IsPrim(V) As Boolean

End Function
Function IsEq(V1, V2) As Boolean
IsEq = V1 = V2
End Function
Sub Er(QQMsg$, ParamArray Ap())
Dim Av(): Av = Ap
Err.Raise 1, , FmtQQAv(QQMsg, Av)
End Sub
Function CurMd() As CodeModule
Set CurMd = Application.VBE.ActiveCodePane.CodeModule
End Function
Function DftMd(Optional A) As CodeModule
If IsNothing(A) Then
    Set DftMd = CurMd
Else
    Set DftMd = A
End If
End Function
Sub AA()
Md.Src.Tst
End Sub
Private Function DrLin__V$(V, W)
Dim O$
If IsArray(V) Then
    If AyIsEmpty(V) Then
        O = AlignL("", W)
    Else
        O = AlignL(FmtQQ("Ay?:", UB(V)) & V(0), W)
    End If
Else
    O = Replace(V, vbCrLf, "|")
    O = AlignL(O, W)
End If
DrLin__V = O
End Function
Function SplitCrLf(S) As String()
SplitCrLf = Split(S, vbCrLf)
End Function

Function AySy(Ay) As String()
If AyIsEmpty(Ay) Then Exit Function
If IsStrAy(Ay) Then AySy = Ay: Exit Function
Dim U&, O$(), J&, I
J = 0
U = UB(Ay)
ReDim O(U)
For Each I In Ay
    O(J) = I
    J = J + 1
Next
AySy = O
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
Sub AAA()
Dim A%()
ReDim A(10)
A(0) = 11
AAAA A
Stop

End Sub
Function AAAA(A)
A(0) = 1
End Function

Function CollObjAy(ObjColl) As Object()
Dim O() As Object
Dim V
For Each V In ObjColl
    Push O, V
Next
CollObjAy = O
End Function
Sub Asg(V, OV)
If IsObject(V) Then
    Set OV = V
Else
    OV = V
End If
End Sub
Function IsNbr(V) As Boolean
Select Case VarType(V)
Case vbInteger, vbByte, vbLong, vbSingle, vbDouble, vbCurrency: IsNbr = True
End Select
End Function
Sub IsNoVal__Tst()
Debug.Assert IsNoVal(Nothing)
Debug.Assert IsNoVal(" ")
Debug.Assert IsNoVal(Empty)
Debug.Assert IsNoVal("")
End Sub
Function IsNoVal(V) As Boolean
IsNoVal = True
If IsNothing(V) Then Exit Function
If IsEmpty(V) Then Exit Function
If IsMissing(V) Then Exit Function
If IsEmptyStr(V) Then Exit Function
End Function
Function IsEmptyStr(V) As Boolean
If Not IsStr(V) Then Exit Function
If Trim(V) = "" Then Exit Function
IsEmptyStr = True
End Function
Function IsEmpty(V) As Boolean
IsEmpty = True
If IsMissing(V) Then Exit Function
If IsNothing(V) Then Exit Function
If VBA.IsEmpty(V) Then Exit Function
If IsStr(V) Then
    If V = "" Then Exit Function
End If
If IsArray(V) Then
    If AyIsEmpty(V) Then Exit Function
End If
IsEmpty = False
End Function

Function IsEmptyColl(ObjColl) As Boolean
IsEmptyColl = (ObjColl.Count = 0)
End Function

Function IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Function

Function IsStrAy(V) As Boolean
IsStrAy = VarType(V) = vbArray + vbString
End Function

Function Max(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
    If Av(J) > O Then O = Av(J)
Next
Max = O
End Function


Function Acs() As Access.Application
Static X As Access.Application
On Error GoTo XX
Dim A$: A = X.Name
Set Acs = X
Exit Function
XX:
Set X = New Access.Application
Set Acs = X
End Function

Sub FbBrw(Fb$)
Acs.OpenCurrentDatabase Fb
Acs.Visible = True
End Sub
Function NowStr$()
NowStr = Format(Now(), "YYYY-MM-DD HH:MM:SS")
End Function

Function Min(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
    If Av(J) < O Then O = Av(J)
Next
Min = O
End Function


Function VarLen%(V)
If IsNull(V) Then Exit Function
If IsArray(V) Then
    If AyIsEmpty(V) Then Exit Function
    VarLen = Len(V(0))
    Exit Function
End If
VarLen = Len(V)
End Function

Private Sub IsStrAy__Tst()
Dim A$()
Dim B: B = A
Dim C()
Dim D
Debug.Assert IsStrAy(A) = True
Debug.Assert IsStrAy(B) = True
Debug.Assert IsStrAy(C) = False
Debug.Assert IsStrAy(D) = False
End Sub

Function AlignL$(S, W)
Dim L%:
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
Function RplVBar$(S)
RplVBar = Replace(S, "|", vbCrLf)
End Function
Private Sub RmvFstTerm__Tst()
Debug.Assert RmvFstTerm("  df dfdf  ") = "dfdf"
End Sub
Function RestTerm$(S)
RestTerm = Brk1(Trim(S), " ").S2
End Function
Function RmvFstTerm(S)
RmvFstTerm = Brk1(Trim(S), " ").S2
End Function
Function FstTerm$(S)
FstTerm = Brk1(Trim(S), " ").S1
End Function
Sub StrDmp(S)
Dim J&, C$
For J = 1 To Len(S)
    C = Mid(S, J, 1)
    Debug.Print J, Asc(C), C
Next
End Sub
Function AlignR$(S, W%)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
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

Function IsSfx(S, Sfx) As Boolean
IsSfx = (Right(S, Len(Sfx)) = Sfx)
End Function

Function LasChr$(S)
LasChr = Right(S, 1)
End Function

Function LasLin$(S)
LasLin = Ay(SplitCrLf(S)).LasEle
End Function

Function LinesCnt&(Lines$)
LinesCnt = Sz(SplitCrLf(Lines))
End Function

Function Quote$(S, QuoteStr$)
With BrkQuote(QuoteStr)
    Quote = .S1 & CStr(S) & .S2
End With
End Function

Sub StrBrw(S)
Dim T$: T = Tmp.Ft
StrWrt S, T
Ft(T).Brw
End Sub

Function StrWrt(S, Ft)
Dim F%: F = FreeFile(1)
Open Ft For Output As #F
Print #F, S
Close #F
'Dim T As TextStream
'Set T = Fso.OpenTextFile(Ft, ForWriting, True)
'T.Write S
'T.Close
End Function

Function DaoTySimTy(T As Dao.DataTypeEnum) As eSimTy
Dim O As eSimTy
Select Case T
Case _
    Dao.DataTypeEnum.dbBigInt, _
    Dao.DataTypeEnum.dbByte, _
    Dao.DataTypeEnum.dbCurrency, _
    Dao.DataTypeEnum.dbDecimal, _
    Dao.DataTypeEnum.dbDouble, _
    Dao.DataTypeEnum.dbFloat, _
    Dao.DataTypeEnum.dbInteger, _
    Dao.DataTypeEnum.dbLong, _
    Dao.DataTypeEnum.dbNumeric, _
    Dao.DataTypeEnum.dbSingle
    O = eNbr
Case _
    Dao.DataTypeEnum.dbChar, _
    Dao.DataTypeEnum.dbGUID, _
    Dao.DataTypeEnum.dbMemo, _
    Dao.DataTypeEnum.dbText
    O = eTxt
Case _
    Dao.DataTypeEnum.dbBoolean
    O = eLgc
Case _
    Dao.DataTypeEnum.dbDate, _
    Dao.DataTypeEnum.dbTimeStamp, _
    Dao.DataTypeEnum.dbTime
    O = eDte
Case Else
    O = eOth
End Select
DaoTySimTy = O
End Function
Function DftCn(Cn) As ADODB.Connection
If IsNoVal(Cn) Then
    Set DftCn = CurCn
Else
    Set DftCn = Cn
End If
End Function
Sub AssertIsAy(V)
If Not IsAy(V) Then Stop
End Sub
Function CurCn() As ADODB.Connection

End Function
Function Dft(V, DftVal)
If IsNoVal(V) Then
    V = DftVal
Else
    V = V
End If
End Function

Function AyMinus(Ay1, Ay2)
Dim O: O = Ay1: Erase O
Dim A2 As Ay: Set A2 = Ay(Ay2)
Dim V
For Each V In Ay1
    If Not A2.Has(V) Then Push O, V
Next
AyMinus = O
End Function
Function ApSy(ParamArray Ap()) As String()
Dim Av(): Av = Ap
ApSy = AySy(Av)
End Function

Function SpcEsc$(S)
If InStr(S, "~") > 0 Then Debug.Print "SpcEsc: Warning: escaping a string-with-space is found with a [~].  The [~] before escape will be changed to space after unescape"
SpcEsc = Replace(S, " ", "~")
End Function
Function DftDb(D As Database) As Database
If IsNothing(D) Then
    Set DftDb = CurDb.Db
Else
    Set DftDb = D
End If
End Function

Function CurDb() As Db
Static X As Database
If IsNothing(X) Then Set X = CurrentDb
Set CurDb = Db(X)
End Function
Function SpcUnE$(S)
SpcUnE = Replace(S, "~", " ")
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
        O = Replace(O, "?", IIf(IsNull(I), "Null", I), Count:=1)
    Next
    If NeedUnEsc Then O = Replace(O, Chr(255), "?")
FmtQQAv = O
End Function

Function TakBet$(S, S1, S2, Optional NoTrim As Boolean, Optional InclMarker As Boolean)
With Brk1(S, S1, NoTrim)
    If .S2 = "" Then Exit Function
    Dim O$: O = Brk1(.S2, S2, NoTrim).S1
    If InclMarker Then O = S1 & O & S2
    TakBet = O
End With
End Function
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

Private Sub TakBet__Tst()
Const S1$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??"
Const S2$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX"
Debug.Assert TakBet(S1, "DATABASE=", ";") = "??"
Debug.Assert TakBet(S2, "DATABASE=", ";") = "??"
End Sub

Function RmvDblSpc$(S)
Dim O$: O = S
While HasSubStr(O, "  ")
    O = Replace(O, "  ", " ")
Wend
RmvDblSpc = O
End Function

Function RmvFstChr$(S)
RmvFstChr = RmvFstNChr(S)
End Function

Function RmvFstNChr$(S, Optional N% = 1)
RmvFstNChr = Mid(S, N + 1)
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

