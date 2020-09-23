Attribute VB_Name = "FormatSupport"
Option Explicit
'This Enum is a direct lift from Ulli's Code Formatter
'© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
'It allows the formating to preserve the hidden attributes (if any) of procedures
'which would otherwise be destroyed by the formatting process.
'This lift also includes the procedures RestoreMemberAttributes and SaveMemberAttributes
Private Enum MemAttrPtrs
 MemName = 0
 MemBind = 1
 MemBrws = 2
 MemCate = 3
 MemDfbd = 4
 MemDesc = 5
 MemDbnd = 6
 MemHelp = 7
 MemHidd = 8
 MemProp = 9
 MemRqed = 10
 MemStme = 11
 MemUide = 12
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private MemName, MemBind, MemBrws, MemCate, MemDfbd, MemDesc, MemDbnd, MemHelp, MemHidd, MemProp, MemRqed, MemStme, MemUide
#End If
Private Attributes()                    As Variant
'used by sorting code
Private SortElems()                     As Variant
''Private SortElem                       As Variant
'End of Lift: Thanks Ulli
Public Enum SilentFinds
 NoneFound
 OnlyOnce
 DeclarationOnly
 PrivateMod
 PublicMod
 CurProcOnly
 SelTextOnly
 ModuleOnly
 ModuleExempt
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private NoneFound, OnlyOnce, DeclarationOnly, PrivateMod, PublicMod, CurProcOnly, SelTextOnly, ModuleOnly, ModuleExempt
#End If
Private Type SelectionCoords
 startLine                             As Long
 EndLine                               As Long
 startCol                              As Long
 endCol                                As Long
End Type
''Private Sel                            As SelectionCoords
Public Enum InstrLocations
 IpNone
 IpExact
 IpMiddle
 IpLeft
 IpRight
 ip2nd
 ip3rd
 ipLeftOr2nd
 ip2ndOr3rd
 ipAny
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private IpNone, IpExact, IpMiddle, IpLeft, IpRight, ip2nd, ip3rd, ipLeftOr2nd, ip2ndOr3rd, ipAny
#End If
Public Enum WriteMode
 WMDelete
 WMInsert
 WMReplace
 WMReplaceUpDate
 WMMove
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private WMDelete, WMInsert, WMReplace, WMReplaceUpDate, WMMove
#End If
Public SuspendCF                        As Boolean
' the use of & in following protects them from detecion by the program itself
Private Const InCodeDontTouchOn         As String = "'SUSPEND " & "CODE FIXER ON"
Private Const InCodeDontTouchOff        As String = "'SUSPEND " & "CODE FIXER OFF"
'used to insert the comment marker
Public Const RGSignature                As String = vbNewLine & "'<" & ":-" & ")"
'used to detect and delete the comments with marker
Private Const RGSignatureDetector       As String = "'<" & ":-" & ")"
Private Const sIndentError              As String = RGSignature & "'Indent " & "Error"
Private Const sDeadCode                 As String = RGSignature & "'Dead " & "Code"
Private Const PrevCode                  As String = RGSignature & "PREVIOUS: "
Private Const PreserveConditComp        As String = "Private DummyToForceVBToRecogniseConditionalCompilationAtEndOfDeclarations As Boolean"
Private Const Apostrophe                As String = "'"
Private Const Colon                     As String = ":"
Private Const ContMark                  As String = " _"
Private Const StrPlus                   As String = " + "
Private Const Hash_If_False_Then        As String = "#If False Then"
Private Const Hash_End_If               As String = "#End If"
Private Const EnumCaseProtectorHead     As String = Hash_If_False_Then & " 'Enum Case Protection"
Private DisguiseStack                   As New ClsStackVB2TheMax
Public Type UnDoData
 strModulename                         As String
 strModule                             As String
 MemberData                            As Variant
End Type
Private UndoArray()                     As UnDoData
Public CompCount                        As Long

Public Sub AsDeclarationFormatting(cmpMod As CodeModule, _
                                   ByVal LTop As Long, _
                                   ByVal LEnd As Long)

  Dim I            As Long
  Dim CommentStore As String
  Dim DImCount     As Long
  Dim TmpLen       As Long
  Dim AsOffSet     As Long
  Dim EOLOffSet    As Long
  Dim L_Codeline   As String
  Dim Tpos         As Long
  Dim As_Pos       As Long
  Dim EndOfDec     As Long
  Dim StartOfDec   As Long

 If iRange = SelCode Then
  StartOfDec = LTop
  EndOfDec = LEnd
  Else
  StartOfDec = 1
  EndOfDec = cmpMod.CountOfDeclarationLines
 End If
 For I = StartOfDec To EndOfDec
  L_Codeline = cmpMod.Lines(I, 1)
  As_Pos = AsDeclarationTarget(L_Codeline)
  If As_Pos > 0 Then
   CommentStore = CommentClip(L_Codeline)
   DImCount = DImCount + 1
   If InCode(L_Codeline, As_Pos) Then
    If As_Pos > AsOffSet Then
     AsOffSet = As_Pos
    End If
   End If
   If LenB(CommentStore) Then
    TmpLen = Len(L_Codeline)
    If TmpLen > EOLOffSet Then
     EOLOffSet = TmpLen
    End If
   End If
  End If
 Next I
 If AsOffSet Then
  For I = StartOfDec To EndOfDec
   L_Codeline = cmpMod.Lines(I, 1)
   As_Pos = AsDeclarationTarget(L_Codeline)
   If As_Pos > 0 Then
    Tpos = Get_As_Pos(L_Codeline)
    If Tpos <> AsOffSet Then
     If InCode(L_Codeline, Tpos) Then
      L_Codeline = Safe_Replace(L_Codeline, " As ", Space$(Abs(1 + AsOffSet - Tpos)) & "As ", , 1)
      CodeLineWrite WMReplace, cmpMod, I, L_Codeline, Split(""), 0
      'cmpMod.ReplaceLine I, L_Codeline
     End If
    End If
   End If
  Next I
 End If
 If EOLOffSet Then
  For I = StartOfDec To EndOfDec
   L_Codeline = cmpMod.Lines(I, 1)
   As_Pos = AsDeclarationTarget(L_Codeline)
   If As_Pos > 0 Then
    CommentStore = CommentClip(L_Codeline)
    If LenB(CommentStore) Then
     L_Codeline = L_Codeline & Space$(EOLOffSet + AsOffSet - Len(L_Codeline)) & Trim$(CommentStore)
     cmpMod.ReplaceLine I, L_Codeline
    End If
   End If
  Next I
 End If

End Sub

Private Function AsDeclarationTarget(strCode As String) As Long

 If InStr(strCode, "Declare ") = 0 Then
  If CountSubString(strCode, " As ") = 1 Then
   If CountSubString(strCode, ", _") = 0 Then
    If CountSubString(strCode, ")") = CountSubString(strCode, "(") Then
     AsDeclarationTarget = Get_As_Pos(strCode)
    End If
   End If
  End If
 End If

End Function

Public Sub AsProcedureDo(cmpMod As CodeModule, _
                         ByVal LTop As Long, _
                         ByVal LEnd As Long)

  Dim I            As Long
  Dim CommentStore As String
  Dim DImCount     As Long
  Dim TmpLen       As Long
  Dim AsOffSet     As Long
  Dim EOLOffSet    As Long
  Dim L_Codeline   As String
  Dim Tpos         As Long
  Dim LastDimLine  As Long
  Dim FirstDimLine As Long

 FirstDimLine = -1
 If LTop < LEnd Then
  For I = LTop To LEnd
   If IsDimLine(Trim$(cmpMod.Lines(I, 1))) Then
    If FirstDimLine = -1 Then
     FirstDimLine = I
    End If
    LastDimLine = I
   End If
  Next I
  If FirstDimLine > -1 Then
   For I = FirstDimLine To LastDimLine
    L_Codeline = cmpMod.Lines(I, 1)
    If IsDimLine(Trim$(L_Codeline)) Then
     CommentStore = CommentClip(L_Codeline)
     DImCount = DImCount + 1
     TmpLen = Get_As_Pos(L_Codeline)
     If InCode(L_Codeline, TmpLen) Then
      If TmpLen > AsOffSet Then
       AsOffSet = TmpLen
      End If
     End If
     If LenB(CommentStore) Then
      TmpLen = Len(L_Codeline)
      If TmpLen > EOLOffSet Then
       EOLOffSet = TmpLen
      End If
     End If
    End If
   Next I
   If AsOffSet Then
    For I = LTop To LEnd
     L_Codeline = cmpMod.Lines(I, 1)
     If IsDimLine(Trim$(L_Codeline)) Then
      Tpos = Get_As_Pos(L_Codeline)
      If Tpos <> AsOffSet Then
       If InCode(L_Codeline, Tpos) Then
        L_Codeline = Safe_Replace(L_Codeline, " As ", Space$(Abs(1 + AsOffSet - Tpos)) & "As ", , 1)
        cmpMod.ReplaceLine I, L_Codeline
       End If
      End If
     End If
    Next I
   End If
   If EOLOffSet Then
    For I = LTop To LEnd
     L_Codeline = cmpMod.Lines(I, 1)
     If IsDimLine(Trim$(L_Codeline)) Then
      CommentStore = CommentClip(L_Codeline)
      If LenB(CommentStore) Then
       L_Codeline = L_Codeline & Space$(EOLOffSet + AsOffSet - Len(L_Codeline)) & Trim$(CommentStore)
       cmpMod.ReplaceLine I, L_Codeline
      End If
     End If
    Next I
   End If
  End If
 End If

End Sub

Public Sub AsProcedureFormatting(cmpMod As CodeModule, _
                                 cdeline As Long, _
                                 SelStartL As Long, _
                                 SelEndL As Long)

  Dim Pname      As String
  Dim PlineNo    As Long
  Dim PStartLine As Long
  Dim PEndLine   As Long
  Dim lJunk      As Long

 GetLineData cmpMod, cdeline, Pname, PlineNo, PStartLine, PEndLine, lJunk
 If iRange = SelCode Then
  ' only line in selection tested
  AsProcedureDo cmpMod, SelStartL, SelEndL
  Else
  ' check whole Procedure
  AsProcedureDo cmpMod, PStartLine, PEndLine
 End If

End Sub

Public Sub BlankDelete(cmpMod As CodeModule, _
                       ByVal KillLine As Long)

  'Allows code to remove either
  'current line if KillLine = currentLine
  'OR
  'next line if KillLine = currentLine + 1

 Do While LenB(Trim$(cmpMod.Lines(KillLine, 1))) = 0
  If KillLine <= cmpMod.CountOfLines Then
   cmpMod.DeleteLines KillLine, 1
   Else
   Exit Do
  End If
  If KillLine > cmpMod.CountOfLines Then
   Exit Do
  End If
 Loop

End Sub

Public Function BuildStrReplace(cmpMod As CodeModule, _
                                ByVal cdeline As Long, _
                                ByVal strCode As String, _
                                ByVal NewLineStartJump As Boolean, _
                                ByVal IndntLevel As Long, _
                                ByVal bIndentError As Boolean, _
                                ByVal bDeadCode As Boolean, _
                                ByVal NewLineEndJump As Boolean) As String

  Dim strtmp As String

 strtmp = IIf(NewLineStartJump, vbNewLine, vbNullString)
 strtmp = strtmp & IIf(isProcHead(strCode), vbNewLine, vbNullString)
 strtmp = strtmp & String$(IndntLevel, vbTab)
 strtmp = strtmp & strCode
 strtmp = strtmp & IIf(bIndentError, sIndentError, vbNullString)
 strtmp = strtmp & IIf(bDeadCode, sDeadCode, vbNullString)
 strtmp = strtmp & IIf(NewLineEndJump, vbNewLine, vbNullString)
 strtmp = strtmp & IIf(isProcEnd(strCode), vbNewLine, vbNullString)
 Do While InStr(strtmp, vbNewLine & vbNewLine) 'clean up excess blanklines
  strtmp = Replace$(strtmp, vbNewLine & vbNewLine, vbNewLine)
 Loop
 strtmp = FormatNewLine(cmpMod, cdeline, strtmp, IndntLevel)
 BuildStrReplace = strtmp

End Function

Public Sub CodeLineRead(cmpMod As CodeModule, _
                        cdeline As Long, _
                        strCode As String, _
                        LCArray As Variant)

  Dim I As Long

 strCode = Trim$(cmpMod.Lines(cdeline, 1))
 Do While HasLineCont(strCode)
  LCArray(I) = Len(strCode) - 1
  I = I + 1
  strCode = Left$(strCode, Len(strCode) - 1) & Trim$(cmpMod.Lines(cdeline + 1, 1))
  CodeLineWrite WMDelete, cmpMod, cdeline + 1, strCode, LCArray, 0
 Loop

End Sub

Public Sub CodeLineWrite(Mode As WriteMode, _
                         cmpMod As CodeModule, _
                         cdeline As Long, _
                         strCode As String, _
                         LCArray As Variant, _
                         ByVal IndntLevel As Long, _
                         Optional ByVal MoveTo As Long, _
                         Optional ByVal MoveFrom As Long)

  Dim I              As Long
  Dim MultiLineAddon As Long

 With cmpMod
  Select Case Mode
   Case WMDelete '
   .DeleteLines cdeline, 1
   Case WMInsert '.InsertLines
   If HasLineCont(strCode) Then
    FormatLineContinuation strCode
   End If
   .InsertLines cdeline, strCode
   Case WMReplace, WMReplaceUpDate '.ReplaceLine
   If UBound(LCArray) > -1 Then
    If LCArray(0) > 0 Then
     For I = 24 To 0 Step -1
      If LCArray(I) > 0 Then
       strCode = Left$(strCode, LCArray(I) + IndntLevel) & ContMark & vbNewLine & Mid$(strCode, LCArray(I) + IndntLevel + 1)
       MultiLineAddon = MultiLineAddon + 1
      End If
     Next I
    End If
   End If
   If HasLineCont(strCode) Then
    FormatLineContinuation strCode
   End If
   .ReplaceLine cdeline, strCode
   On Error Resume Next
   .CodePane.SetSelection cdeline, 1, cdeline, -1
   On Error GoTo 0
   If Mode = WMReplaceUpDate Then
    'update line data
    CodeLineRead cmpMod, cdeline, strCode, LCArray
   End If
   cdeline = cdeline + MultiLineAddon
   Case WMMove
   If MoveTo > 0 Then
    If MoveFrom > 0 Then
     strCode = .Lines(MoveFrom, 1)
     .DeleteLines cdeline, 1
     .InsertLines MoveTo, strCode
    End If
   End If
  End Select
  '.Lines
  '.AddFromString
  '.ProcBodyLine
  '.ProcCountLines
  '.ProcOfLine
  '.ProcStartLine
 End With

End Sub

Public Function ColonExpander(strCode As String) As Boolean

  Const DoneIt       As String = RGSignature & " Colon(s) Expanded"
  Dim strWork        As String
  Dim CommentStore   As String
  Dim GoToOnCodeLine As String

 strWork = strCode
 If InStr(strWork, ":") Then
  CommentStore = CommentClip(strWork)
  If InstrAtPosition(strWork, ":", ipAny, False) Then
   'and colon is in Code
   'GoTo targets need to retain their colon to keep label status
   If IsGotoLabel(strWork) Then
    'Other parts of code fail if Goto Target has comment on same line
    'so separate line but retain the colon
    If strCode <> strWork Then
     strCode = strWork & vbNewLine & Trim$(CommentStore)
     ColonExpander = True
    End If
    Exit Function
   End If
   'extremely rare but legal
   If IsGotoLabel(LeftWord(strWork)) Then
    If Not WordIsVBSingleWordCommand(LeftWord(strWork)) Then
     GoToOnCodeLine = LeftWord(strWork)
     strWork = Trim$(Mid$(strWork, Len(GoToOnCodeLine) + 1))
     CommentStore = Trim$(CommentStore)
    End If
   End If
   'deal with legal but unnecessary colons
   If InstrAtPosition(strWork, "Else:", ipAny, True) Then
    'If X then DoBarney Else: DoFred
    If InstrAtPosition(strWork, "Case Else:", ipAny, True) = False Then
     strWork = Safe_Replace(strWork, " Else: ", " Else ")
    End If
   End If
   strWork = Safe_Replace(strWork, " Then: ", " Then ")
   'If X Then: DoBarney Else DoFred
   '
   'All other code colons can be replaced with new line
   strWork = Safe_Replace(strWork, ": ", vbNewLine)
   If Right$(strWork, 1) = ":" Then
    'Just in case someone put an unnecessary colon on end of a line
    '            If InCode(StrCode, Len(StrCode)) Then
    strWork = Left$(strWork, Len(strCode) - 1)
    '            End If
   End If
   strWork = strWork & CommentStore
   Else
   'colon only appeared in comment so restore comment and exit
   strWork = strWork & CommentStore
  End If
  If LenB(GoToOnCodeLine) Then
   strWork = GoToOnCodeLine & vbNewLine & strWork
  End If
  If MultiRight(strWork, True, vbNewLine & "_") Then
   strWork = Left$(strWork, Len(strWork) - 3)
  End If
 End If
 If strWork <> strCode Then
  ColonExpander = True
  strCode = strWork & IIf(bShowFixComment, DoneIt, vbNullString)
  AddToSearchBox Mid$(DoneIt, 3), True
 End If

End Function

Public Sub CommentOut(cmpMod As CodeModule, _
                      strCode As String, _
                      cdeline As Long, _
                      IndntLevel As Long, _
                      LCArray As Variant, _
                      NewLineStartJump As Boolean, _
                       bIndentError As Boolean, _
                       bDeadCode As Boolean, _
                       NewLineEndJump As Boolean, _
                      ByVal StrEnd As String)

  'comments out all code between cdeLine and first line containing StrEnd (inclusive)
  'also adds a space before and after the commented out section
  'resets indentlevel to 0

 IndntLevel = 0
 
 strCode = BuildStrReplace(cmpMod, cdeline, strCode, NewLineStartJump, IndntLevel, bIndentError, bDeadCode, NewLineEndJump)
 strCode = Replace(strCode, vbNewLine, vbNewLine & "''")
 strCode = IIf(Left$(strCode, 2) <> "''", "''", vbNullString) & strCode
 Erase LCArray
 With cmpMod
CodeLineWrite WMReplace, cmpMod, cdeline, strCode, LCArray, IndntLevel
 ' .ReplaceLine cdeline, strCode
  Do
   cdeline = cdeline + 1
   strCode = Trim$(.Lines(cdeline, 1))
   CodeLineRead cmpMod, cdeline, strCode, LCArray
   strCode = IIf(Left$(strCode, 2) <> "''", "''", vbNullString) & strCode
   CodeLineWrite WMReplace, cmpMod, cdeline, strCode, LCArray, IndntLevel
   '.ReplaceLine cdeline, strCode
  Loop Until MultiLeft(Trim$(.Lines(cdeline, 1)), True, "''" & StrEnd)
  '.ReplaceLine cdeline, Trim$(.Lines(cdeline, 1)) & vbNewLine
  CodeLineWrite WMReplace, cmpMod, cdeline, Trim$(.Lines(cdeline, 1)) & vbNewLine, LCArray, IndntLevel
 End With 'cmpMod

End Sub

Public Function ConcealParameterCommas(ByVal varCode As Variant) As String

  Dim CommaSpacePos As Long

 'Replace any CommaSpace with comma in bracketed parameters
 'This allows CommaSpace delimited Dim,Public, Private to be safely detected without cutting in parameters
 'VB will automatically restore them
 CommaSpacePos = GetCommaSpacePos(varCode)
 If CommaSpacePos Then
  Do
   If InCode(varCode, CommaSpacePos) Then
    If EnclosedInBrackets(varCode, CommaSpacePos) Then
     varCode = Left$(varCode, CommaSpacePos - 1) & "," & Mid$(varCode, CommaSpacePos + 2) 'Replace$(varCode, ", ", ",")
    End If
   End If
   CommaSpacePos = GetCommaSpacePos(varCode, CommaSpacePos + 1)
  Loop While CommaSpacePos > 0
 End If
 ConcealParameterCommas = varCode

End Function

Public Function DealWithEndOfDeclarationsConditional(cmpMod As CodeModule, _
                                                     cdeline As Long, _
                                                     LCArray As Variant, _
                                                     ByVal Trigger As Boolean) As Boolean

  Dim EndOfCondition As Long
  Dim Msg            As String

 If Trigger Then
  If cdeline = cmpMod.CountOfDeclarationLines Then
   EndOfCondition = cdeline
   Do
    EndOfCondition = EndOfCondition + 1
    If EndOfCondition >= cmpMod.CountOfLines Then
     Msg = RGSignature & "Optional Compilation Structure not closed"
     'something is wrong
     Exit Do
    End If
   Loop Until MultiLeft(cmpMod.Lines(EndOfCondition, 1), True, "#End If")
   If LenB(Msg) Then
    CodeLineWrite WMReplace, cmpMod, cdeline, cmpMod.Lines(cdeline, 1) & vbNewLine & Msg, LCArray, 0
    AddToSearchBox Msg, True
    Else
    CodeLineWrite WMReplace, cmpMod, EndOfCondition, cmpMod.Lines(EndOfCondition, 1) & vbNewLine & PreserveConditComp, LCArray, 0
    DealWithEndOfDeclarationsConditional = True
   End If
  End If
 End If

End Function

Public Function DeclarationExpandMulti(strCode As String) As Boolean

  Const DoneIt     As String = RGSignature & "Multiple Declaration line expanded"
  Dim strWork      As String
  Dim Cline        As String
  Dim CommentStore As String
  Dim CommaPos     As Long
  Dim DimStatic    As String
  Dim Guard        As String
  Dim colonPos     As Long

 strWork = strCode
 DimStatic = LeftWord(strWork) & " "
 If InstrArray(strWork, ",", ":") Then
  Cline = strWork
  CommentStore = CommentClip(strWork)
  strWork = ConcealParameterCommas(strWork)
  Guard = strWork
  CommaPos = 0
  colonPos = 0
  Do
   'Very rare but possible Dim X: Dim Y
   colonPos = InStr(colonPos + 1, strWork, ": ")
   If colonPos Then
    If Not EnclosedInBrackets(strWork, colonPos) Then
     strWork = Left$(strWork, colonPos - 1) & vbNewLine & DimStatic & Mid$(strWork, CommaPos + 2)
    End If
   End If
  Loop While colonPos
  Do
   CommaPos = GetCommaSpacePos(strWork, CommaPos + 1)
   If CommaPos Then
    If Not EnclosedInBrackets(strWork, CommaPos) Then
     strWork = Left$(strWork, CommaPos - 1) & vbNewLine & DimStatic & Mid$(strWork, CommaPos + 2)
    End If
   End If
  Loop While CommaPos
  If Guard <> strWork Then
   strCode = strWork & CommentStore & IIf(bShowFixComment, DoneIt, vbNullString) & IIf(bShowPrevCode, PrevCode & strCode, vbNullString)
   AddToSearchBox Mid$(DoneIt, 3), True
   DeclarationExpandMulti = True
  End If
 End If

End Function

Public Function DeclarationMultiSingleTypeing(strCode As String) As Boolean

  Dim strWork      As String
  Dim strOrig      As String
  Dim CommentStore As String
  Dim TmpB         As Variant
  Const DoneIt     As String = RGSignature & "Multiple Declaration with single As Type repaired."
  Dim TypeDef      As String
  Dim J            As Long

 strWork = strCode
 CommentStore = CommentClip(strWork)
 If InStr(strWork, ",") Then
  If CountSubString(strWork, " As ") = 1 Then
   strWork = ConcealParameterCommas(strWork)
   strOrig = strWork
   TmpB = Split(strWork, ", ")
   If UBound(TmpB) > 0 Then
    TypeDef = TmpB(UBound(TmpB))
    If Get_As_Pos(TypeDef) > 0 Then
     TypeDef = Mid$(TypeDef, Get_As_Pos(TypeDef))
     For J = LBound(TmpB) To UBound(TmpB) - 1
      If Get_As_Pos(TmpB(J)) = 0 Then
       If InStr("!@#$%&", Right$(TmpB(J), 1)) = 0 Then
        TmpB(J) = TmpB(J) & TypeDef
       End If
      End If
     Next J
     If strOrig <> Join(TmpB, ", ") Then
      strCode = Join(TmpB, ", ") & CommentStore & IIf(bShowFixComment, DoneIt, vbNullString) & IIf(bShowPrevCode, PrevCode & strCode, vbNullString)
      DeclarationMultiSingleTypeing = True
      AddToSearchBox Mid$(DoneIt, 3), True
     End If
    End If
   End If
  End If
 End If

End Function

Public Sub DisguiseLiteral(StrSearch As Variant, _
                           ByVal HideMe As Variant, _
                           ByVal HideTShowF As Boolean)

  Dim LocalFind    As String
  Dim LocalReplace As String
  Dim FindPos      As Long
  Dim LDisguise    As String

 'Copyright 2003 Roger Gilchrist
 'e-mail: rojagilkrist@hotmail.com
 'Replaces literal test words with rubbish so that Split can't find them
 'Call second time with HideME and Disguise reversed
 'after the string has been reassembled by Join to restore literals
 'Disguise is regenerated each time you call the routine to hide something( HideTShowF = True)
 If HideTShowF Then
  ' create a masking value for string literals
  Do
   LDisguise = RandomString(48, 122, 3, 6)
  Loop While InStr(LDisguise, StrSearch) Or InStr(LDisguise, HideMe)
  DisguiseStack.Push HideMe
  DisguiseStack.Push LDisguise
  LocalFind = HideMe
  LocalReplace = LDisguise
  Else
  LocalFind = DisguiseStack.Pop
  LocalReplace = DisguiseStack.Pop
 End If
 FindPos = InStr(StrSearch, LocalFind)
 If FindPos Then
  Do
   If FindPos Then
    If InLiteral(StrSearch, FindPos) Then
     StrSearch = Left$(StrSearch, FindPos - 1) & LocalReplace & Mid$(StrSearch, FindPos + Len(LocalFind))
    End If
   End If
   FindPos = InStr(FindPos + 1, StrSearch, LocalFind)
  Loop While FindPos
 End If

End Sub

Public Sub DoCommentDelete()

  'Delete all Code Fixer style Comments
  
  Dim Code                      As String
  Dim compmod                   As CodeModule
  Dim Comp                      As VBComponent
  Dim Proj                      As VBProject
  Dim Pane                      As CodePane
  Dim CurProc                   As String
  Dim curModule                 As String
  Dim Procname                  As String
  Dim startLine                 As Long
  Dim EndLine                   As Long
  Dim startCol                  As Long
  Dim endCol                    As Long
  Dim SelstartLine              As Long
  Dim selendline                As Long
  Dim SelStartCol               As Long
  Dim SelEndCol                 As Long
  Dim AutoRangeRevert           As Boolean
  Dim PrevCurCodePane           As Long

 On Error Resume Next
 With mobjDoc
  .ShowWorking True, "Comment Removal..."
 End With 'mobjDoc
 DoEvents
 bCancel = False
 GetCounts
 CurProc = GetCurrentProcedure
 curModule = GetCurrentModule
 AutoSelectInitialize PrevCurCodePane, AutoRangeRevert
 If iRange = SelCode Then
  VBInstance.ActiveCodePane.GetSelection SelstartLine, SelStartCol, selendline, SelEndCol
 End If
 For Each Proj In VBInstance.VBProjects
  For Each Comp In Proj.VBComponents
   If SafeCompToProcess(Comp) Then
    If iRange > AllCode Then
     If Comp.Name <> curModule Then
      GoTo SkipComp
     End If
    End If
    Set compmod = Comp.CodeModule
    VisibleScroll_Init Pane, compmod
    startLine = 1
    If compmod.Find(RGSignatureDetector, startLine, 1, compmod.CountOfLines, -1, bWholeWordonly, bCaseSensitive, False) Then
     Do
      EndLine = -1
      startCol = 1
      endCol = -1
      If compmod.Find(RGSignatureDetector, startLine, startCol, EndLine, endCol, bWholeWordonly, bCaseSensitive, False) Then
       Procname = compmod.ProcOfLine(startLine, vbext_pk_Proc)
       If LenB(Procname) = 0 Then
        Procname = "(Declarations)"
       End If
       If iRange = ProcCode Then
        If Procname <> CurProc Then
         GoTo SkipProc
        End If
       End If
       Code$ = compmod.Lines(startLine, 1)
       ApplySelectedTextRestriction Code$, startLine, startCol, EndLine, endCol, SelstartLine, SelStartCol, selendline, SelEndCol
       If Len(Code$) Then
        Code$ = Left$(Code$, InStr(Code, RGSignatureDetector) - 1)
       End If
       If Len(Code$) Then
        compmod.ReplaceLine startLine, Code$
        Else
        compmod.DeleteLines startLine, 1
        startLine = startLine - 1 'cope with next line also being a comment
       End If
SkipProc:
      End If
      Code$ = vbNullString
      startLine = startLine + 1
      If mobjDoc.CancelSearch Then
       Exit Do
      End If
      If startLine > compmod.CountOfLines Then
       Exit Do
      End If
     Loop While compmod.Find(RGSignatureDetector, startLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) 'StartLine > 0 And StartLine <= CompMod.CountOfLines
    End If
   End If
SkipComp:
   Set Comp = Nothing
   If mobjDoc.CancelSearch Then
    Exit For
   End If
  Next Comp
  If mobjDoc.CancelSearch Then
   Exit For
  End If
 Next Proj
 'this turns off auto Selected text only
 If AutoRangeRevert Then
  iRange = PrevCurCodePane
  mobjDoc.ToggleButtonFaces
 End If
 mobjDoc.ComboDeleteText SearchB, RGSignature, True
 Set Proj = Nothing
 Set compmod = Nothing
 mobjDoc.ShowWorking False
 On Error GoTo 0

End Sub
Function GenerateEnumCapProtection(cmpMod As CodeModule, cdeline As Long, ByVal EnumStruct As String) As String
'"Public Enum fred vbnewline wilma vbnewline betty
Dim I As Long
Dim TmpA As Variant
Dim strtmp As String
Dim DoIt As Boolean
TmpA = Split(EnumStruct, vbNewLine)
TmpA(0) = ""
For I = LBound(TmpA) To UBound(TmpA)

TmpA(I) = LeftWord(TmpA(I))
'these remove any members that for whatever reason cannot be given Enum Case Protection
'becuase as long as you enclose them in [] you can use anything including reserved words, operators, and number started words
'(i.e. '3D-Left') would not be able to be used in Enum Case Protection because
'1. variables can't start with numbers,
'2. VB's auto formatting would make the - a minus sign (not acceptable in a declaration)
'3. Left would override the reserved word/command Left
TmpA(I) = Safe_Replace(TmpA(I), "[", vbNullString)
TmpA(I) = Trim$(Safe_Replace(TmpA(I), "]", vbNullString))
If InStr(TmpA(I), "-") Or GetSpacePos(TmpA(I)) Or IsNumeric(Left$(TmpA(I), 1)) Or InStr(TmpA(I), ".") Or IsInArray(Right$(TmpA(I), 1), TypeSuffixArray) Or IsInArray(TmpA(I), VBReservedWords) Then
TmpA(I) = vbNullString
End If
Next

EnumStruct = Join(TmpA, ", ")
'take out any blank members
Do While InStr(EnumStruct, ", ,")
EnumStruct = Replace(EnumStruct, ", ,", ", ")
Loop
'Mid takes out the initial blank member caused by deleting the Enum Declaration
EnumStruct = "Private " & Mid(EnumStruct, 3)
LongLineFormat EnumStruct, 1000, vbNewLine & "Private "
If cdeline + UBound(TmpA) + 1 = cmpMod.CountOfLines Then
'there can be no EnumCaseProtection so insert it
 DoIt = True
ElseIf Not MultiLeft(cmpMod.Lines(GetNextCodeLine(cmpMod, cdeline + UBound(TmpA) + 2), 1), True, Hash_If_False_Then) Then
'the next valid code line is not the head of an enum case protection so insert it
DoIt = True
ElseIf MultiLeft(cmpMod.Lines(GetNextCodeLine(cmpMod, cdeline + UBound(TmpA) + 2), 1), True, Hash_If_False_Then) Then
'It might be but check that the next line is the matching Enum Case Protection Code
If Not MultiLeft(EnumStruct, True, cmpMod.Lines(GetNextCodeLine(cmpMod, cdeline + UBound(TmpA) + 3), 1)) Then
'NOTE this line test code against generated string because the generated string may be multiple lines if the enum is long enough
DoIt = True
End If
End If
If DoIt Then
GenerateEnumCapProtection = EnumCaseProtectorHead & vbNewLine & EnumStruct & vbNewLine & "#End If" & RGSignature & "Enum Case Protection must follow the Enum Structure it protects"
End If
End Function
Private Function GetSafeCutPoint(strA As String, _
                                 ByVal BasePos As Long) As Long
  
  Do

    GetSafeCutPoint = GetSpacePos(strA, BasePos)
    BasePos = BasePos + 1
  Loop Until InCode(strA, GetSafeCutPoint) Or BasePos = Len(strA) Or GetSafeCutPoint = 0


End Function

Public Sub LongLineFormat(VarLongLine As Variant, _
                                            Optional ByVal BaseLength As Long = 1023, _
                                            Optional ByVal Sep As String = ContMark)
                                            
  
  'Copes with very long long lines by inserting designated separators
  'NOTE that if you use the ContMark (Line continuation characters= " _") there is a VB limit of 25
  'The other limit is that no single line (Without line continuations) can exceed a LenB value of 1023
  'if theither of these are reached you're in trouble but this is very unlikely.
  Dim initCut         As Long

  Dim strTemp         As String
  Dim strVeryLong     As String
  Dim LngCutPoint     As Long
  Dim LinContCount    As Long
  If BaseLength > 1023 Then
    BaseLength = 1023
  End If
  If InStr(Sep, ContMark) Then
  'this test makes sure that the resulting structure can not exceed the limits of VB
  Do Until LenB(VarLongLine) / BaseLength <= 24
  BaseLength = BaseLength + 1
  If BaseLength > 1023 Then
  MsgBox "The Code is too large to be used.", vbCritical
  'this message box should never hit if the original code worked
  Exit Sub
  End If
  Loop
  End If
  If LenB(VarLongLine) > BaseLength Then
    ' get initial cut point
    initCut = BaseLength
    strTemp = VarLongLine
    LngCutPoint = GetSafeCutPoint(strTemp, initCut)

    If LngCutPoint > 0 Then
      strVeryLong = Mid$(strTemp, LngCutPoint + 1)
      strTemp = Left$(strTemp, LngCutPoint - 1)

      Do While LenB(strVeryLong)
        LngCutPoint = GetSafeCutPoint(strVeryLong, initCut)

        If LngCutPoint = 0 Then

          If LenB(strVeryLong) > 0 Then
            strTemp = strTemp & Sep & strVeryLong
            Exit Do
          End If

        Else
          strTemp = strTemp & Sep & Left$(strVeryLong, LngCutPoint - 1)
          strVeryLong = Mid$(strVeryLong, LngCutPoint + 1)

'this test only needs to be done if using Line continuation characters
          If InStr(Sep, ContMark) Then
            LinContCount = LinContCount + 1
            If LinContCount = 24 Then
              'upperlimit is 25 .
              'This is set to 24 so that the last one can be used to clean up and exit.
              'The last line may  be excessively long and cause a Structural Failure.
              MsgBox "Too many Line Continuation Characters needed by current line", vbCritical
              strTemp = strTemp & Sep & strVeryLong
              Exit Do
            End If

          End If

        End If

      Loop

    End If

    VarLongLine = strTemp
  End If


End Sub

Public Sub DoIndent()

  Dim LineContArray(25)                As Long
  Dim Code                             As String
  Dim codeline                         As Long
  Dim ContLineCount                    As Long
  Dim StrTrigger                       As String
  Dim StrLastCode                      As String
  Dim strCodeOnly                      As String
  Dim StrRep                           As String
  Dim IndentLevel                      As Long
  Dim PReIndent                        As Long
  Dim targetproc                       As String
  Dim CurrentProc                      As String
  Dim TestNewProc                      As String
  Dim PrevProc                         As String
  Dim Proj                             As VBProject
  Dim Comp                             As VBComponent
  Dim compmod                          As CodeModule
  Dim Pane                             As CodePane
  Dim TargetModule                     As String
  Dim SelCodeLine                      As Long
  Dim SelstartLine                     As Long
  Dim selendline                       As Long
  Dim SelStartCol                      As Long
  Dim SelEndCol                        As Long
  Dim AutoRangeRevert                  As Boolean
  Dim PrevCurCodePane                  As Long
  Dim jumpDownVal                      As Boolean
  Dim jumpUpVal                        As Boolean
  Dim jumpUp2Val                       As Boolean
  Dim bDeadCode                        As Boolean
  Dim bIndentError                     As Boolean
  Dim NewLineStartJump                 As Boolean
  Dim NewLineEndJump                   As Boolean
  Dim HoldJump                         As Boolean
  Dim bDimSafety                       As Boolean
  Dim junk                             As String
  Dim bindentErrorRestrict             As Boolean                                                                                                                                                                                                                                                                                                                                                                                                                                                   ' keeps Indent error message within procedure with error
  Dim ReTestDim                        As Boolean                                                                                                                                                                                                                                                                                                                                                                                                                                                   ' Test for new End of Dims only when necessary
  Dim InNewProc                        As Boolean
  Dim InConditionalCompilation         As Boolean
  Dim DontWrite                        As Boolean
  Dim EnumTypeStruct                   As String
Dim strEnumCapProtection As String
 On Error Resume Next
 DoEvents
 bCancel = False
 GetCounts
 'Get limiters for Range
 targetproc = GetCurrentProcedure
 TargetModule = GetCurrentModule
 AutoSelectInitialize PrevCurCodePane, AutoRangeRevert
 If iRange = SelCode Then
  VBInstance.ActiveCodePane.GetSelection SelCodeLine, SelStartCol, selendline, SelEndCol
 End If
 '
 UnDoListInit
 DoCommentDelete
 mobjDoc.ShowWorking True, "Indenting..."
 For Each Proj In VBInstance.VBProjects
  For Each Comp In Proj.VBComponents
   PrevProc = vbNullString
   If SafeCompToProcess(Comp) Then
    bIndentError = False
    IndentLevel = 0
    'Test for Current Code Range
    If iRange > AllCode Then
     If Comp.Name <> TargetModule Then
      GoTo SkipComp
     End If
    End If
    '
    Set compmod = Comp.CodeModule
    'Indenting destroys Procedure Attributes (Tools menu) so store them. Thanks Ulli.
    SaveMemberAttributes compmod.Members
    If bSortModules Then
     If iRange < ProcCode Then
      ' don't sort for procedure or selected text modes
      UllisSort compmod
     End If
    End If
    '  If bVisibleIndenting Then
    VisibleScroll_Init Pane, compmod
    '   End If
    'Set Start of Search
    codeline = InitiateSelRange(compmod, targetproc, SelCodeLine)
    'Test for Procedure Range
    CurrentProc = GetCurrentProcedure
    If iRange = ProcCode Then
     If CurrentProc <> targetproc Then
      GoTo SkipProc
     End If
    End If
    ReTestDim = False
    'Indent loop start
    If iRange < SelCode Then
     If bProcDecl2Top Then
      MoveDimToTopOfProc compmod, codeline
     End If
    End If
    If bExpandIfThen Or bExpandColon Then
     MultipleInstructionLineExpansion compmod, Code, codeline, LineContArray, targetproc, selendline
    End If
    codeline = InitiateSelRange(compmod, targetproc, SelCodeLine)
    '    CodeLineRead CompMod, codeline, Code, LineContArray
    Do
ExpanderReDo:
     InNewProc = False
     TestNewProc = GetCurrentProcedure
     If PrevProc <> TestNewProc Then
      InNewProc = True
      PrevProc = TestNewProc
     End If
     bindentErrorRestrict = False
     jumpUpVal = False
     jumpDownVal = False
     jumpUp2Val = False
     bDimSafety = False
     NewLineStartJump = False
     NewLineEndJump = False
     DontWrite = False
     'Deal with Blank Lines
     BlankCleaners compmod, codeline
     'Safety exit if Deal with Blanks reaches end of code
     If codeline > compmod.CountOfLines Then
      Exit Do
     End If
     'Get line of code
     Erase LineContArray
     CodeLineRead compmod, codeline, Code, LineContArray
     StrRep = Code
     'Hard coded replaces comment marker 'Rem ' with more common '
     'Located here because JustACommentOrBlank will keep rest of procedure from fixing it.
     If MultiLeft(Code$, True, "Rem ") Then
      CodeLineWrite WMReplace, compmod, codeline, Replace$(Code$, "Rem ", "'", 1, 1), LineContArray, 0
     End If
     ContLineCount = codeline
     If bVisibleIndenting Then
      VisibleScroll_Do Pane, codeline
     End If
     If Code$ = sIndentError Then
      bIndentError = False
     End If
     'skip  blank and comment only lines
     If IsRgSignature(Code) Then
      'this keeps comments immediately after the code the refer to
      If LineisBlank(compmod, codeline - 1) Then
       CodeLineWrite WMDelete, compmod, codeline - 1, Code, LineContArray, 0
       GoTo ExpanderReDo
      End If
     End If
     If Not JustACommentOrBlank(Code$) Then
      'Apply Single Line Updates to code
      If bStrConcatenateUpdate Then
       If StringConcatenationUpDate(Code) Then
        CodeLineWrite WMReplace, compmod, codeline, Code, LineContArray, 0
        GoTo ExpanderReDo
       End If
      End If
      If bStrFunctionUpdate Then
       If DoStringFunctionsCorrect(Code) Then
        CodeLineWrite WMReplace, compmod, codeline, Code, LineContArray, 0
        GoTo ExpanderReDo
       End If
      End If
      If bChr2ConstFix Then
      Chr2ConstantDo Code
      End If
      'collect Code information for formatters
      StrTrigger = LeftWord(Code$)
      strCodeOnly = Code$
      junk$ = CommentClip(strCodeOnly)
      strCodeOnly = Trim$(strCodeOnly)
      StrLastCode = LastWord(strCodeOnly)
      'accumulate info from Line Continuation lines and allow jump across whole structure
      'this should preserve any indenting user placed on them
      'Expand Code fixes
      If IsDeclarationLine(Code) Then
       If Not InEnumCapProtection(codeline, compmod) Then
        If bDeclSingleTypeExpand Then
         If DeclarationMultiSingleTypeing(Code) Then
          CodeLineWrite WMReplace, compmod, codeline, Code, LineContArray, 0
          GoTo ExpanderReDo
         End If
        End If
        If bDeclExpand Then
         If DeclarationExpandMulti(Code) Then
          CodeLineWrite WMReplace, compmod, codeline, Code, LineContArray, 0
          GoTo ExpanderReDo
         End If
        End If
       End If
      End If
      If ArrayMember(StrTrigger, "#Const", "Event", "WithEvents", "Const", "Option", "Declare", "Global", "Dim", "Public", "Private", "Friend", "Static", "Function", "Sub", "Property", "Enum", "Type", "Deflng", "DefBool", "CefByte", "DefInt", "DefCur", "DefSng", "DefDbl", "DefDec", "DefDate", "DefStr", "DefObj", "Enum", "Type") Or (SecondWord(Code) = "As" And Not InTypeDef(codeline, compmod) And Not HasLineCont(Code) And CountSubString(Code, ")") = CountSubString(Code, "(")) Then
       If isProcHead(Code) Then
        'always separate procedures
        NewLineStartJump = Not LineisBlank(compmod, codeline - 1)
        IndentLevel = 0
        If bAddStructureSpace Then
         jumpUpVal = True
        End If
       End If
       If bTypeSuffixUpdate Then
        TypeSuffixExtender Code
       End If
       If bTestScope Then
        If ArrayMember(StrTrigger, "Enum", "Type") Or ArrayMember(SecondWord(Code$), "Enum", "Type") Then
         If Not MultiLeft(Code$, True, "Type As") Then
          EnumTypeStruct = GetEnumTypeStructure(compmod, codeline)
          If Not ScopeEnumTypeTest(EnumTypeStruct, Code, GetCurrentModule) Then
           If bCommentOutUnused Then
            CommentOut compmod, Code, codeline, IndentLevel, LineContArray, NewLineStartJump, bIndentError, bDeadCode, NewLineEndJump, "End " & StrTrigger
           End If
           
          End If
          If bEnumCapProtect Then
          If InstrAtPosition(EnumTypeStruct, "Enum", ipLeftOr2nd, True) Then
          
          strEnumCapProtection = GenerateEnumCapProtection(compmod, codeline, EnumTypeStruct)
          End If
         End If
         End If
        End If
        If Not InEnumCapProtection(codeline, compmod) Then
         If Not ScopeChangeTest(Code, Comp.Name, GetLineProcedure(compmod, codeline)) Then
          If ArrayMember(SecondWord(Code), "Function", "Sub", "Property") Then
           If bCommentOutUnused Then
            CommentOut compmod, Code, codeline, IndentLevel, LineContArray, NewLineStartJump, bIndentError, bDeadCode, NewLineEndJump, "End " & SecondWord(Code)
           End If
          End If
         End If
        End If
       End If
       If MultiLeft(Code$, True, "Type As") Then
        'guard for Enum/Type member named 'Type' (legal but irritating)
        Else
        If ArrayMember(StrTrigger, "Const", "Dim", "Static", "Public", "Private", "Global") Then
         ' if the 1st 3 occur within a procedure but not at top they muck up the indent flow
         ' so this preserves the indent level
         If ArrayMember(StrTrigger, "Public", "Private") Then
          If ArrayMember(SecondWord(Code), "Sub", "Function", "Property", "Declare") Then
           IndentLevel = 0
           GoTo NotForDeclarationFix
          End If
         End If
         bDimSafety = True
         PReIndent = IndentLevel
         If InDeclaration(compmod, codeline) Then
          If ArrayMember(SecondWord(Code$), "Enum", "Type") Then
           jumpUpVal = True
          End If
          If InConditionalCompilation Then
           IndentLevel = 1
           Else
           IndentLevel = 0
           If ArrayMember(StrTrigger, "Private") Then
            If InEnumCapProtection(codeline, compmod) Then
             IndentLevel = 1
            End If
           End If
          End If
          Else
          IndentLevel = 1
          jumpUpVal = True
         End If
         Else
         IndentLevel = 0
         jumpUpVal = True
        End If
       End If
       Else
      End If
NotForDeclarationFix:
      'these lines cope with single line structures
      'if you turn it on  'Expand If..Then... Structures' and 'Expand Colon Separators' will fix these
      SingleLineStructure strCodeOnly, StrTrigger, IndentLevel, jumpUpVal, jumpUp2Val, InConditionalCompilation
      '
      If ArrayMember(StrTrigger, "Exit") Then
       Select Case SecondWord(Code$)
        Case "Sub", "Function", "Property"
        If Not ExitIsInGoToStructure(compmod, codeline) Then
         '''an Exit at indent level 1 or less means that following code can never hit
         If Not UnNeededExit(compmod, codeline, Code) Then
          bDeadCode = IndentLevel <= 1
          '''this does not work yet so edited out
         End If
        End If
       End Select
      End If
      '
      bIndentError = IndentLevel < 0
      If IndentLevel < 0 Then
       IndentLevel = 0
      End If
      '
      If bPleonasmFix Then
      PleonasmCleaner Code
      End If
      '
      If bAddStructureSpace Then
       If ArrayMember(StrTrigger, "For", "Do", "While", "Select", "If") Then
        If Not LineisBlank(compmod, codeline - 1) Then
         NewLineStartJump = True
         ' it is a single line structure so needs blank line after as well
         If jumpUpVal = False Then
          NewLineEndJump = True
         End If
        End If
       End If
      End If
      If ArrayMember(StrTrigger, "Next", "Loop", "Wend", "End") Then
       If Not LineisBlank(compmod, codeline + 1) Then
        If StrTrigger = "End" Then
         Select Case SecondWord(Code)
          Case "If", "Select", "With", "Type", "Enum"
          NewLineEndJump = True
       
         End Select
         Else
         NewLineEndJump = True
        End If
       End If
       If bEnumCapProtect Then
          If SecondWord(Code) = "Enum" Then
          If Len(strEnumCapProtection) Then
          If compmod.Lines(codeline + 1, 1) <> EnumCaseProtectorHead Then
          Code = Code & vbNewLine & strEnumCapProtection
          End If
          strEnumCapProtection = ""
          End If
          End If
      End If
      End If
     End If
RetestEndofDeclarations:
     If isProcEnd(Code) Or isDeclarationEnd(compmod, codeline) Then
      If isDeclarationEnd(compmod, codeline) Then
       If DealWithEndOfDeclarationsConditional(compmod, codeline, LineContArray, InConditionalCompilation) Then
        GoTo RetestEndofDeclarations
       End If
      End If
      If bDeclAsFormat Then
       If isDeclarationEnd(compmod, codeline) Then
        AsDeclarationFormatting compmod, SelCodeLine, selendline
        ' DontWrite = True
        Else
        AsProcedureFormatting compmod, codeline, SelstartLine, selendline
       End If
      End If
      bindentErrorRestrict = True
      IndentLevel = 0
      bIndentError = IndentLevel > 0
      If bAddStructureSpace Then
       If Not codeline = compmod.CountOfDeclarationLines Then
        NewLineStartJump = True
       End If
      End If
      bDeadCode = False
      'this inserts a blank after the procedure if necessary
      If Not LineisBlank(compmod, codeline + 1) Then
       NewLineEndJump = True
      End If
     End If
     If HoldJump Then
      NewLineEndJump = True
      HoldJump = False
     End If
     If NewLineEndJump Then
      If Right$(strCodeOnly, 1) = "_" Then
       HoldJump = True
       NewLineEndJump = False
      End If
      If IsRgSignature(compmod.Lines(codeline + 1, 1)) Then
       HoldJump = True
       NewLineEndJump = False
      End If
     End If
     If isProcHead(Code) Then
      If bAddStructureSpace Then
       NewLineEndJump = True
      End If
      jumpUpVal = True
      If FormatVBStructures(Code) Then
       Erase LineContArray
       ContLineCount = codeline + CountSubString(Code, vbNewLine)
      End If
     End If
     If isAPIDeclare(Code) Then
      If FormatVBStructures(Code) Then
       Erase LineContArray
       ContLineCount = codeline + CountSubString(Code, vbNewLine)
      End If
     End If
     If bIndentError Then
      AddToSearchBox sIndentError, True
      bIndentError = False
     End If
     If bDeadCode Then
      AddToSearchBox sDeadCode, True
     End If
     '     If Not DontWrite Then
     If MultiRight(sDeadCode, True, Code) = False Then
      'don't pile dead code mesaages on each other
      StrRep = BuildStrReplace(compmod, codeline, Code, NewLineStartJump, IndentLevel, bIndentError, bDeadCode, NewLineEndJump)
      CodeLineWrite WMReplace, compmod, codeline, StrRep, LineContArray, IndentLevel
     End If
     ' End If
     If bDimSafety Then
      If IndentLevel > PReIndent Then
       IndentLevel = PReIndent
      End If
     End If
     If jumpUp2Val Then
      IndentLevel = IndentLevel + 2
     End If
     If jumpUpVal Then
      IndentLevel = IndentLevel + 1
     End If
     If jumpDownVal Then
      IndentLevel = IndentLevel - 1
     End If
     If ContLineCount > codeline Then
      codeline = ContLineCount + 1
      Else
      codeline = codeline + 1
     End If
     If NewLineEndJump Then
      codeline = codeline + 1
     End If
     If NewLineStartJump Then
      codeline = codeline + 1
     End If
     If mobjDoc.CancelSearch Then
      GoTo SkipComp
      'ensures that members are restored before exiting For structure
     End If
     If codeline > compmod.CountOfLines Then
      Exit Do
     End If
    Loop While InSRange(codeline, compmod, targetproc, selendline)
    'Indent loop Finish
SkipProc:
   End If
SkipComp:
   RestoreMemberAttributes compmod.Members
   If mobjDoc.CancelSearch Then
    Exit For
   End If
  Next Comp
  If mobjDoc.CancelSearch Then
   Exit For
  End If
 Next Proj
 mobjDoc.ShowWorking False
 Set Comp = Nothing
 Set Proj = Nothing
 Set compmod = Nothing
 On Error GoTo 0

End Sub

Private Function DoStringFunctionsCorrect(strCode As String) As Boolean

  Const DoneIt     As String = RGSignature & "Variable Function converted to String Function"
  Dim J            As Long
  Dim tmpstring2   As String
  Dim TmpString1   As String
  Dim strWork      As String
  Dim CommentStore As String

 'Ulli's code updated to use built-in array instead of Listbox
 'You don't get the option with me!!
 'but it is safe for code which make literal string references to these (Like the array that this uses)
 strWork = strCode
 CommentStore = CommentClip(strWork)
 For J = LBound(StrFuncArray) To UBound(StrFuncArray)
  ' This protects the String Type from being treated as the String function
  If Get_As_Pos(strWork) = 0 Then
   TmpString1 = StrFuncArray(J)
   If InstrAtPosition(strWork, TmpString1, ipAny, False) Then
    tmpstring2 = "(" & TmpString1
    TmpString1 = " " & TmpString1
    strWork = Replace$(strWork, TmpString1 & "(", TmpString1 & "$(")
    strWork = Replace$(strWork, tmpstring2 & "(", tmpstring2 & "$(")
   End If
  End If
 Next J
 If strCode <> strWork & CommentStore Then
  strCode = strWork & CommentStore & IIf(bShowFixComment, DoneIt, vbNullString) & IIf(bShowPrevCode, PrevCode & strCode, vbNullString)
  AddToSearchBox Mid$(DoneIt, 3), True
  DoStringFunctionsCorrect = True
 End If

End Function

Public Sub DoUnIndent()

  Dim codeline                         As Long
  Dim targetproc                       As String
  Dim CurrentProc                      As String
  Dim Proj                             As VBProject
  Dim Comp                             As VBComponent
  Dim compmod                          As CodeModule
  Dim Pane                             As CodePane
  Dim TargetModule                     As String
  Dim SelCodeLine                      As Long
  Dim selendline                       As Long
  Dim SelStartCol                      As Long
  Dim SelEndCol                        As Long
  Dim AutoRangeRevert                  As Boolean
  Dim PrevCurCodePane                  As Long
  Dim Code                             As String
  Dim LineContArray(25)                As Long

 On Error Resume Next
 DoEvents
 bCancel = False
 GetCounts
 'Get limiters for Range
 targetproc = GetCurrentProcedure
 TargetModule = GetCurrentModule
 AutoSelectInitialize PrevCurCodePane, AutoRangeRevert
 If iRange = SelCode Then
  VBInstance.ActiveCodePane.GetSelection SelCodeLine, SelStartCol, selendline, SelEndCol
 End If
 '
 DoCommentDelete
 mobjDoc.ShowWorking True, "UnIndenting..."
 For Each Proj In VBInstance.VBProjects
  For Each Comp In Proj.VBComponents
   If SafeCompToProcess(Comp) Then
    'Test for Current Code Range
    If iRange > AllCode Then
     If Comp.Name <> TargetModule Then
      GoTo SkipComp
     End If
    End If
    '
    Set compmod = Comp.CodeModule
    'Indenting destroys Procedure Attributes (Tools menu) so store them. Thanks Ulli.
    SaveMemberAttributes compmod.Members
    If bSortModules Then
     If iRange < ProcCode Then
      ' don't sort for procedure or selected text modes
      UllisSort compmod
     End If
    End If
    If bVisibleIndenting Then
     VisibleScroll_Init Pane, compmod
    End If
    'Set Start of Search
    codeline = InitiateSelRange(compmod, targetproc, SelCodeLine)
    'Test for Procedure Range
    CurrentProc = GetCurrentProcedure
    If iRange = ProcCode Then
     If CurrentProc <> targetproc Then
      GoTo SkipProc
     End If
    End If
    'Indent loop start
    Do
     'Deal with Blank Lines
     BlankCleaners compmod, codeline
     'Safety exit if Deal with Blanks reaches end of code
     If codeline > compmod.CountOfLines Then
      Exit Do
     End If
     'THIS IS IT
     'CompMod.ReplaceLine codeline, Trim$(CompMod.Lines(codeline, 1))
     Erase LineContArray
     CodeLineRead compmod, codeline, Code, LineContArray
     CodeLineWrite WMReplace, compmod, codeline, Code, LineContArray, 0
     '
     If bVisibleIndenting Then
      VisibleScroll_Do Pane, codeline
     End If
     codeline = codeline + 1
     If mobjDoc.CancelSearch Then
      GoTo SkipComp
      'ensures that members are restored before exiting For structure
     End If
     If codeline > compmod.CountOfLines Then
      Exit Do
     End If
    Loop While InSRange(codeline, compmod, targetproc, selendline)
    'Indent loop Finish
SkipProc:
   End If
SkipComp:
   RestoreMemberAttributes compmod.Members
   If mobjDoc.CancelSearch Then
    Exit For
   End If
  Next Comp
  If mobjDoc.CancelSearch Then
   Exit For
  End If
 Next Proj
 'this turns off auto Selected text only
 mobjDoc.ShowWorking False
 Set Comp = Nothing
 Set Proj = Nothing
 Set compmod = Nothing
 On Error GoTo 0

End Sub

Public Function EnclosedInBrackets(ByVal StrSearch As String, _
                                   ByVal ChrPos As Long) As Boolean

  Dim LBracketCount As Long
  Dim RBracketCount As Long
  Dim CommentStore  As String
  Dim MyStr         As String
  Dim LBit          As String
  Dim Rbit          As String

 'Copyright 2003 Roger Gilchrist
 'e-mail: rojagilkrist@hotmail.com
 'Detect whether a Character Position is between brackets
 If Not JustACommentOrBlank(StrSearch) Then
  If GetRightBracketPos(StrSearch) > 0 Then
   If GetLeftBracketPos(StrSearch) > 0 Then
    MyStr = StrSearch
    CommentStore = CommentClip(MyStr)
    DisguiseLiteral MyStr, "(", True
    DisguiseLiteral MyStr, ")", True
    LBracketCount = CountSubString(MyStr, "(")
    RBracketCount = CountSubString(MyStr, ")")
    If LBracketCount = LBracketCount Then
     If LBracketCount > 0 Then
      LBit = Left$(MyStr, ChrPos)
      Rbit = Mid$(MyStr, ChrPos)
      LBracketCount = Abs(CountSubString(LBit, "(") - CountSubString(LBit, ")"))
      RBracketCount = Abs(CountSubString(Rbit, "(") - CountSubString(Rbit, ")"))
      If RBracketCount = LBracketCount Then
       If LBracketCount > 0 Then
        EnclosedInBrackets = True
       End If
      End If
     End If
    End If
    DisguiseLiteral MyStr, "(", False
    DisguiseLiteral MyStr, ")", False
   End If
  End If
 End If

End Function

Public Function ExitIsInGoToStructure(compmod As CodeModule, _
                                      ByVal codeline As Long) As Boolean

  Dim CurProc  As String
  Dim GotoLine As Long
  Dim TmpLine  As Long

 CurProc = GetLineProcedure(compmod, codeline)
 TmpLine = codeline - 1
 Do While CurProc = GetLineProcedure(compmod, TmpLine)
  If InStr(compmod.Lines(TmpLine, 1), "GoTo") Then
   GotoLine = TmpLine
   Exit Do
  End If
  TmpLine = TmpLine - 1
 Loop
 If GotoLine > 0 Then
  GotoLine = 0
  TmpLine = codeline + 1
  Do While CurProc = GetLineProcedure(compmod, TmpLine)
   If IsGotoLabel(compmod.Lines(TmpLine, 1)) Then
    GotoLine = TmpLine
    Exit Do
   End If
   TmpLine = TmpLine + 1
  Loop
 End If
 ExitIsInGoToStructure = GotoLine > 0

End Function

Public Function ExpandForDetection(VarA As Variant) As String

  Dim CommentStore As String
  Dim L_Codeline   As String

 L_Codeline = Trim$(VarA)
 CommentStore = CommentClip(L_Codeline)
 L_Codeline = Replace$(L_Codeline, "(", " ( ")
 L_Codeline = Replace$(L_Codeline, ")", " ) ")
 L_Codeline = Replace$(L_Codeline, " -", " - ")
 L_Codeline = Replace$(L_Codeline, ",", " , ")
 L_Codeline = Replace$(L_Codeline, ":=", " := ")
 L_Codeline = Replace$(L_Codeline, "= -", "= - ") 'special case X = -FunctionName(y)
 Do While InStr(L_Codeline, "  ")
  L_Codeline = Replace$(L_Codeline, "  ", " ")
 Loop
 ExpandForDetection = L_Codeline

End Function

Public Sub FormatLineContinuation(strCode As String)

  '<STUB> Reason: not yet written


End Sub

Public Function FormatNewLine(compmod As CodeModule, _
                              ByVal codeline As Long, _
                              ByVal strCode As String, _
                              ByVal IndntLevel As Long) As String

  'formats new code strings by making sure that Format Comments are on separate line
  'and that there are no double newlines in string
  'and that Format comments are not indented

 strCode = Replace$(strCode, RGSignature, vbNewLine & RGSignature)
 If MultiLeft(strCode, True, String$(IndntLevel, vbTab) & vbNewLine & RGSignature) Then
  strCode = Mid$(strCode, InStr(strCode, RGSignature))
 End If
 FormatNewLine = Replace$(strCode, vbNewLine & vbNewLine, vbNewLine)
 ' Do While InStr(FormatNewLine, vbNewLine & vbNewLine & RGSignatureDetector)
 '  FormatNewLine = Replace$(FormatNewLine, vbNewLine & vbNewLine & String$(IndntLevel, vbTab) & RGSignatureDetector, vbNewLine & String$(IndntLevel, vbTab) & RGSignatureDetector & vbNewLine)
 ' Loop
 '  Do While InStr(FormatNewLine, vbNewLine & vbNewLine & RGSignatureDetector)
 '  FormatNewLine = Replace$(FormatNewLine, vbNewLine & vbNewLine & RGSignatureDetector, vbNewLine & RGSignatureDetector & vbNewLine)
 ' Loop
 If codeline > 1 Then
  If LenB(Trim$(compmod.Lines(codeline - 1, 1))) = 0 Then
   Do While MultiLeft(FormatNewLine, True, vbNewLine)
    FormatNewLine = Mid$(FormatNewLine, 3)
   Loop
  End If
 End If

End Function

Private Function FormatVBStructures(strCode As String) As Boolean

  'This routine can cope with (avoid) routine headers and Delcare lines embedded as literal strings.
  'Thanks Ulli for proving this was needed with the CodeProfiler code
  
  Dim SpaceOffSet As Long
  Dim CommaPos    As Long

 If Not InStr(strCode, ContMark) Then
  If CountSubString(strCode, ", ") Then
   SpaceOffSet = GetLeftBracketPos(strCode)
   CommaPos = GetCommaSpacePos(strCode)
   Do While CommaPos
    If InCode(strCode, CommaPos) Then
     strCode = Left$(strCode, CommaPos) & ContMark & vbNewLine & Space$(SpaceOffSet) & Mid$(strCode, CommaPos + 2)
     FormatVBStructures = True
    End If
    CommaPos = GetCommaSpacePos(strCode, CommaPos + 2 + SpaceOffSet)
   Loop
  End If
 End If

End Function

Public Function Get_As_Pos(VarSearch As Variant) As Long

  'gives a name to action making reading code more readable

 Get_As_Pos = InStr(1, VarSearch, " As ")

End Function

Public Function GetCommaSpacePos(VarSearch As Variant, _
                                 Optional StartAt As Long = 1) As Long

  'gives a name to action making reading code more readable

 GetCommaSpacePos = InStr(StartAt, VarSearch, ", ")

End Function

Public Function GetEnumTypeStructure(compmod As CodeModule, _
                                     ByVal codeline As Long) As String

  Dim StructEnd As Long

 StructEnd = codeline
 Do
  GetEnumTypeStructure = GetEnumTypeStructure & IIf(Len(GetEnumTypeStructure), vbNewLine, vbNullString) & Trim$(compmod.Lines(StructEnd, 1))
  StructEnd = StructEnd + 1
 Loop Until MultiLeft(Trim$(compmod.Lines(StructEnd, 1)), True, "End Enum", "End Type")

End Function

Public Function GetLeftBracketPos(VarSearch As Variant, _
                                  Optional StartAt As Long = 1) As Long

  'gives a name to action making reading code more readable

 GetLeftBracketPos = InStr(StartAt, VarSearch, "(")

End Function

Public Function GetModuleType(ByVal TargetMod As String) As Long

  Dim Proj                             As VBProject
  Dim Comp                             As VBComponent
  Dim GotOne                           As Boolean

 For Each Proj In VBInstance.VBProjects
  For Each Comp In Proj.VBComponents
   If SafeCompToProcess(Comp) Then
    If Comp.Name = TargetMod Then
     GetModuleType = Comp.Type
     GotOne = True
     Exit For
    End If
   End If
  Next Comp
  If GotOne Then
   Exit For
  End If
 Next Proj

End Function

Public Function GetNextCodeLine(cmpMod As CodeModule, _
                                FromLine As Long) As Long

  If FromLine < cmpMod.CountOfLines Then
 
 Do While JustACommentOrBlank(Trim$(cmpMod.Lines(FromLine, 1)))
  FromLine = FromLine + 1
  If FromLine >= cmpMod.CountOfLines Then
   Exit Do 'safety; should never hit
  End If
 Loop
 GetNextCodeLine = FromLine
End If
End Function

Public Function GetRightBracketPos(VarSearch As Variant, _
                                   Optional StartAt As Long = 1) As Long

  'gives a name to action making reading code more readable

 GetRightBracketPos = InStr(StartAt, VarSearch, ")")

End Function

Public Function getScopeTestWord(strCode As String, _
                                 Optional LeaveBrackets As Boolean = False) As String

  Dim I As Long

 For I = 1 To CountSubString(strCode, " ") + 1
  Select Case WordMember(strCode, I)
   Case "Global", "Private", "Public", "Friend", "Static", "Property", "Let", "Get", "Set", "Const", "Enum", "Event", "WithEvent", "Declare", "Sub", "Function", "Type", "Dim"
   Case Else
   getScopeTestWord = WordMember(strCode, I)
   If Not LeaveBrackets Then
    If InStr(getScopeTestWord, "(") Then
     getScopeTestWord = Left$(getScopeTestWord, InStr(getScopeTestWord, "(") - 1)
    End If
   End If
   Exit For
  End Select
 Next I

End Function

Public Function GetSpacePos(VarSearch As Variant, _
                            Optional StartAt As Long = 1) As Long

 GetSpacePos = InStr(StartAt, VarSearch, " ")

End Function

Public Function HasLineCont(ByVal VarTest As Variant) As Boolean

 HasLineCont = MultiRight(VarTest, True, ContMark)

End Function

Private Function HasScope(strCode As String) As Boolean

 HasScope = ArrayMember(WordMember(strCode, 1), "Global", "Dim", "Public", "Private", "Friend", "Static")

End Function

Public Sub PleonasmCleaner(strCode As String)
  
  Dim strWork        As String

  Dim CommentStore As String
  Dim SpaceOffSet  As String
  Dim arrTmp       As Variant
  Dim ThenPos As Long
Dim ElsePos As Long
Dim EqualTruePos As Long
Dim DoneIt As String
DoneIt = RGSignature & "Pleonasm Removed"

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Remove unnecessary '= True' from code
  ' remove end comments for restoring after changing Type Suffixes
  strWork = strCode
  CommentStore = CommentClip(strWork)

  If MultiRight(strWork, True, "= True") Then

    If Not MultiLeft(strWork, True, "While", "Do Until", "Do While", "Loop Until", "Loop While") Then
      'it's an assignment of True to a variable so leave it.
      Exit Sub
    End If

  End If

  arrTmp = Split(strWork)
  'rare case of 'A = True And <some Condition>'
  'being used to generate a bitwise calculation

  If UBound(arrTmp) > 2 Then
    If arrTmp(1) = "=" Then
      If arrTmp(2) = "True" Then
        If arrTmp(3) = "And" Then
          Exit Sub
        End If
      End If
    End If
  End If
'Rare case
'If X Then Control.property = True Else ......'
If MultiLeft(strWork, True, "If") Then
 ThenPos = InStr(strWork, " Then ")
 ElsePos = InStr(strWork, " Else")
 EqualTruePos = InStr(strWork, " = True ")
 If BetweenLng(ThenPos, EqualTruePos, ElsePos) Then
           Exit Sub
 End If
End If

  If InStr(strWork, "As Boolean = True") Then
    'protect optional parameters
    strWork = Replace$(strWork, "As Boolean = True", "As Boolean=True")
  End If

  If InStr(strWork, "= True") Then
    strWork = Safe_Replace(strWork, "= True) ", ") ")
    strWork = SpaceOffSet & Safe_Replace(strWork, "= True ", " ")
    strWork = SpaceOffSet & Safe_Replace(strWork, "= True", vbNullString)
  ElseIf InStr(strWork, " True = ") Then
    strWork = Safe_Replace(strWork, "(True = ", "( ")
    strWork = SpaceOffSet & Safe_Replace(strWork, " True = ", " ")
  End If

  If InStr(strWork, "As Boolean=True") Then
    'un-protect optional parameters so that no message is attached
    strWork = Replace$(strWork, "As Boolean=True", "As Boolean = True")
  End If

  If strWork & CommentStore <> strCode Then
    
        strCode = strWork & CommentStore & vbNewLine & DoneIt
        AddToSearchBox Mid$(DoneIt, 3), True
  End If


End Sub

Public Function IfThenStructureExpander(strCode As String) As Boolean

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Expands single line If Then (Else)  structures into multiple line version
  'Any end of line Comment is placed above the code
  
  Dim DoneIt       As String
  Dim strWork        As String
  Dim TmpA           As Variant
  Dim CommentStore   As String
  Dim ThenIfTest     As Long
  Dim I              As Long

 On Error GoTo BadError
 If HasLineCont(strCode) Then
  Exit Function
 End If
 strWork = strCode
 CommentStore = CommentClip(strWork)
 'colon test moved here to deal with unnecessary colons in If Then Structures
 If InstrAtPosition(strWork, "Else:", ipAny, True) Then
  strWork = Safe_Replace(strWork, "Else:", vbNewLine & "Else" & vbNewLine)
 End If
 If InstrAtPosition(strWork, "Then:", ipAny, True) Then
  strWork = Safe_Replace(strWork, "Then:", "Then")
 End If
 'For Code of format:    If pA < pX  Then pX = 1 Else If pX > pA Then pX = 0
 'changes Else If to ElseIf
 If InstrAtPosition(strWork, "Else If", ipAny, True) Then
  strWork = Safe_Replace(strWork, " Else If ", vbNewLine & "ElseIf ")
 End If
 'For Code of format:    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
 ' this need extra End Ifs attached to end of string
 If InstrAtPosition(strWork, "Then If", ipAny, True) Then
  DisguiseLiteral strWork, " Then If ", True
  ThenIfTest = CountSubString(strWork, " Then If ")
  DisguiseLiteral strWork, " Then If ", False
  If ThenIfTest > 0 Then
   strWork = Safe_Replace(strWork, " Then If ", " Then" & vbNewLine & "If ")
   For I = 1 To ThenIfTest
    strWork = strWork & vbNewLine & "End If"
   Next I
  End If
 End If
 DisguiseLiteral strWork, "Then", True
 If InStr(strWork, " Then ") Then
  TmpA = Split(strWork, "Then")
  strWork = Join(TmpA, "Then" & vbNewLine) & vbNewLine & "End If"
 End If
 DisguiseLiteral strWork, "Then", False
 DisguiseLiteral strWork, "Else", True
 If InstrAtPosition(strWork, "Else", ipAny, False) Then
  TmpA = Split(strWork, " Else ")
  strWork = Join(TmpA, vbNewLine & "Else" & vbNewLine)
 End If
 'special case probably poor coding but can be done
If InstrAtPosition(strWork, "Else" & vbNewLine & "End If", IpRight, True) Then
'    strWork = strWork & " "
    TmpA = Split(strWork, " Else" & vbNewLine)
    strWork = Join(TmpA, vbNewLine & "Else" & vbNewLine)
    If InStr(strWork, vbNewLine & "Else" & vbNewLine & "End If") Then
    DoneIt = RGSignature & "unneeded 'Else' statement"
    strWork = Replace(strWork, vbNewLine & "Else" & vbNewLine & "End If", vbNewLine & "Else" & DoneIt & vbNewLine & "End If")
    AddToSearchBox Mid$(DoneIt, 3), True
    End If
  End If

 If InstrAtPosition(strWork, "Case" & vbNewLine & "Else", ipAny, False) Then
  strWork = Safe_Replace(strWork, "Case" & vbNewLine & "Else", "Case Else")
 End If
 DisguiseLiteral strWork, "Else", False
 'comments are placed above the structure to leave space for my comment after
 If strWork & CommentStore <> strCode Then
 DoneIt = RGSignature & " Structure Expanded."
  strCode = CommentStore & vbNewLine & strWork & IIf(bShowFixComment, DoneIt, vbNullString)
  AddToSearchBox Mid$(DoneIt, 3), True
  IfThenStructureExpander = True
 End If

Exit Function

BadError:
 IfThenStructureExpander = False

End Function

Public Function InCode(ByVal VarSearch As Variant, _
                       ByVal TestPos As Long) As Boolean

 If TestPos Then
  If InLiteral(VarSearch, TestPos) Then
   Exit Function
   ElseIf InTimeLiteral(VarSearch) Then
   Exit Function
   ElseIf InComment(VarSearch, TestPos) Then
   Exit Function
   Else
   InCode = True
  End If
 End If

End Function

Public Function InComment(ByVal VarSearch As Variant, _
                          ByVal Tpos As Long) As Boolean

  Dim Possible As Long
  Dim arrTmp   As Variant
  Dim OPos     As Long
  Dim NPos     As Long
  Dim I        As Long

 Possible = InStr(VarSearch, "'")
 If Possible Then
  Do
   If Possible > Tpos Then
    Exit Function
   End If
   If InLiteral(VarSearch, Possible, False) Then
    Possible = InStr(Possible + 1, VarSearch, "'")
   End If
  Loop While InLiteral(VarSearch, Possible, False) And Possible > 0
  If Possible Then
   arrTmp = Split(VarSearch, "'")
   For I = LBound(arrTmp) To UBound(arrTmp)
    NPos = NPos + 1 + Len(arrTmp(I))
    If BetweenLng(OPos, Tpos, NPos) Then
     InComment = Not InLiteral(VarSearch, Tpos, False)
     Exit For
    End If
    OPos = NPos
    If OPos >= Tpos Then
     Exit For
    End If
   Next I
  End If
 End If

End Function

Public Function InEnumCapProtection(ByVal codeline As Long, _
                                    cmpMod As CodeModule) As Boolean

  Dim I            As Long
  Dim Possible     As Boolean

 'Copyright 2003 Roger Gilchrist
 'e-mail: rojagilkrist@hotmail.com
 'Protect Enum Capitalisation Protection from Declaration Formatter
 'by detecting that line is inside an Enum Capitalisation Protection structure
 With cmpMod
  For I = codeline To 1 Step -1
   If MultiLeft(.Lines(I, 1), True, Hash_If_False_Then) Then
    Possible = True
    Exit For
   End If
  Next I
  If Possible Then
   For I = codeline To .CountOfDeclarationLines
    If MultiLeft(.Lines(I, 1), True, Hash_End_If) Then
     InEnumCapProtection = True
     Exit For
    End If
   Next I
  End If
 End With

End Function

Public Function InstrArray(VarSearch As Variant, _
                           ParamArray varFind() As Variant) As Long

  Dim VarTmp As Variant

 For Each VarTmp In varFind
  If InStr(VarSearch, VarTmp) Then
   InstrArray = InStr(VarSearch, VarTmp)
   Exit Function
  End If
 Next VarTmp

End Function

Public Function InstrAtPosition(ByVal VarSearch As Variant, _
                                ByVal varFind As Variant, _
                                ByVal AtLocation As InstrLocations, _
                                Optional WholeWord As Boolean = True) As Boolean

  Dim TmpA         As Variant
  Dim WholeOffset  As String
  Dim SizeOfSearch As Long

 'Copyright 2003 Roger Gilchrist
 'e-mail: rojagilkrist@hotmail.com
 'Return True or False
 'Parameters:
 'varSearch, search in
 'varFInd, search for
 'AtLocation, test that varFind is in varSearch at this position
 '               Left, Right, LeftOr2nd, Middle(=exists but not at Left or Right)
 '               Exact(same as varSearch=varFind), None and Any present in no/any position
 'WholeWord, only if space delimited. While not safe for string literals with punctuation this is always true for code
 '
 'This routine only searches and finds in code, all literals and comments are excluded
 If LenB(VarSearch) Then
  If LenB(varFind) Then
   TmpA = Split(VarSearch)
   SizeOfSearch = UBound(TmpA)
   WholeOffset = IIf(WholeWord, " ", vbNullString)
   DisguiseLiteral VarSearch, varFind, True
   VarSearch = Trim$(VarSearch)
   If LenB(VarSearch) Then
    Select Case AtLocation
     Case IpExact
     InstrAtPosition = VarSearch = varFind
     Case IpLeft
     If (VarSearch = varFind) Then
      InstrAtPosition = True
      ElseIf MultiLeft(VarSearch, True, varFind & WholeOffset) Then
      InstrAtPosition = True
     End If
     Case IpMiddle
     InstrAtPosition = InStr(VarSearch, WholeOffset & varFind & WholeOffset) > 0
     Case IpRight
     If (VarSearch = varFind) Then
      InstrAtPosition = True
      ElseIf MultiRight(VarSearch, True, WholeOffset & varFind) Then
      InstrAtPosition = True
     End If
     Case ip2nd
     If SizeOfSearch > 0 Then
      TmpA(0) = vbNullString
      InstrAtPosition = MultiLeft(Trim$(Join(TmpA)), True, varFind & WholeOffset)
     End If
     Case ip3rd
     If SizeOfSearch > 1 Then
      TmpA(0) = vbNullString
      TmpA(1) = vbNullString
      InstrAtPosition = MultiLeft(Trim$(Join(TmpA)), True, varFind & WholeOffset)
     End If
     Case ipLeftOr2nd
     InstrAtPosition = MultiLeft(VarSearch, True, varFind & WholeOffset)
     If Not InstrAtPosition Then
      If SizeOfSearch > 0 Then
       TmpA(0) = vbNullString
       InstrAtPosition = MultiLeft(Trim$(Join(TmpA)), True, varFind & WholeOffset)
      End If
     End If
     Case ip2ndOr3rd
     If SizeOfSearch > 0 Then
      TmpA(0) = vbNullString
      InstrAtPosition = MultiLeft(Trim$(Join(TmpA)), True, varFind & WholeOffset)
     End If
     If Not InstrAtPosition Then
      If SizeOfSearch > 1 Then
       TmpA(1) = vbNullString
       InstrAtPosition = MultiLeft(Trim$(Join(TmpA)), True, varFind & WholeOffset)
      End If
     End If
     Case IpNone
     InstrAtPosition = True
     If (VarSearch = varFind) Then
      InstrAtPosition = False
      ElseIf InStr(VarSearch, WholeOffset & varFind & WholeOffset) > 0 Then
      InstrAtPosition = False
      ElseIf MultiLeft(VarSearch, True, varFind & WholeOffset) Then
      InstrAtPosition = False
      ElseIf MultiRight(VarSearch, True, WholeOffset & varFind) Then
      InstrAtPosition = False
     End If
     Case ipAny
     If (VarSearch = varFind) Then
      InstrAtPosition = True
      ElseIf InStr(VarSearch, WholeOffset & varFind & WholeOffset) > 0 Then
      InstrAtPosition = True
      ElseIf MultiLeft(VarSearch, True, varFind & WholeOffset) Then
      InstrAtPosition = True
      ElseIf MultiRight(VarSearch, True, WholeOffset & varFind) Then
      InstrAtPosition = True
     End If
    End Select
   End If
   DisguiseLiteral VarSearch, varFind, False
  End If
 End If

End Function

Private Function InTimeLiteral(ByVal VarSearch As Variant) As Boolean

  Dim Ps As Long
  Dim P1 As Long
  Dim P2 As Long

 If CountSubString(VarSearch, "#") > 1 Then
  Ps = InStr(VarSearch, "#")
  Do
   Do
    P1 = InStr(Ps, VarSearch, "#")
    P2 = InStr(P1 + 1, VarSearch, "#")
    Ps = P2
    If Ps = 0 Then
     Exit Do
    End If
   Loop While InLiteral(VarSearch, P1)
   If P1 > 0 Then
    If Not InComment(VarSearch, P1) Then
     If P2 > P1 Then
      If Not InComment(VarSearch, P2) Then
       InTimeLiteral = IsDate(Mid$(VarSearch, P1, P2))
      End If
     End If
    End If
   End If
   If Ps = 0 Then
    Exit Do
   End If
  Loop While P1 > 0
 End If

End Function

Public Function InTypeDef(ByVal codeline As Long, _
                          cmpMod As CodeModule) As Boolean

  Dim I            As Long
  Dim Possible     As Boolean

 'Copyright 2003 Roger Gilchrist
 'e-mail: rojagilkrist@hotmail.com
 'Protect Enum Capitalisation Protection from Declaration Formatter
 'by detecting that line is inside an Enum Capitalisation Protection structure
 With cmpMod
  For I = codeline To 1 Step -1
   If MultiLeft(.Lines(I, 1), True, "Type", "Dim Type", "Private Type", "Public Type") Then
    If MultiLeft(.Lines(I, 1), True, "Type") And MultiLeft(.Lines(I, 1), True, "Type As") = False And Trim$(.Lines(I, 1)) <> "Type" Then
     'because type can have a member called 'Type
     Possible = True
     Else
     Possible = True
    End If
    Exit For
   End If
  Next I
  If Possible Then
   For I = codeline To .CountOfDeclarationLines
    If MultiLeft(.Lines(I, 1), True, "End Type") Then
     InTypeDef = True
     Exit For
    End If
   Next I
  End If
 End With

End Function

Public Function isAPIDeclare(strCode As String) As Boolean

 isAPIDeclare = InstrAtPosition(strCode, "Declare", ipLeftOr2nd, True)

End Function

Public Function isDeclarationEnd(cmpMod As CodeModule, _
                                 cdeline As Long) As Boolean

 isDeclarationEnd = (cdeline = cmpMod.CountOfDeclarationLines)

End Function

Public Function IsDeclarationLine(strCode As String) As Boolean

 IsDeclarationLine = ArrayMember(LeftWord(strCode), "Dim", "Static", "Const", "Private", "Public", "Friend")
 If IsDeclarationLine Then
  IsDeclarationLine = isProcHead(strCode) = False And isAPIDeclare(strCode) = False
 End If

End Function

Public Function IsDimLine(ByVal VarTest As Variant) As Boolean

  'tests for Routine level declarations (Dim,Static,Const)

 IsDimLine = MultiLeft(VarTest, True, "Dim ", "Static ", "Const ")

End Function

Public Function IsGotoLabel(ByVal VarTest As Variant) As Boolean

  'Update ver 1.0.87
  'misrecognized if the original line was of format
  'If Y = 3 then VarTest: x = 3 Else sub_call2: x = 4
  'where VarTest is also a sub call

 If LenB(VarTest) Then
  If GetSpacePos(Trim$(VarTest)) = 0 Then
   ' If Not IsSubCall(Left$(VarTest, Len(VarTest) - 1), Modulenumber) Then
   IsGotoLabel = MultiRight(Trim$(VarTest), True, Colon)
   'End If
  End If
 End If

End Function

Public Function isProcEnd(strCode As String) As Boolean

 isProcEnd = InstrAtPositionArray(strCode, IpLeft, True, "End Sub", "End Function", "End Property")

End Function

Public Function isProcHead(strCode As String) As Boolean

  ' protects from detecting comments

 If Not JustACommentOrBlank(strCode) Then
  isProcHead = InstrAtPositionArray(strCode, ipLeftOr2nd, True, "Sub", "Function", "Property")
 End If

End Function

Public Function IsRgSignature(Code As String) As Boolean

 IsRgSignature = MultiLeft(Code, True, RGSignatureDetector)

End Function

Public Function IsUsedPrivate(strTest As String, _
                              strHomeModule As String) As Boolean

  'tests that a strTest is used inside its own module
  '>1 means that it ignores the calling line

 IsUsedPrivate = SilentSearch(strTest, ModuleOnly, strHomeModule) > 1

End Function

Public Function IsUsedProcedure(strTest As String, _
                                strHomeModule As String, _
                                StrHomeProc As String) As Boolean

  'tests that a strTest is used inside its own module
  '>1 means that it ignores the calling line

 IsUsedProcedure = SilentSearch(strTest, CurProcOnly, strHomeModule, StrHomeProc) > 1

End Function

Public Function IsUsedPublic(strTest As String, _
                             strHomeModule As String) As Boolean

  'tests that a strTest is used outside its own module
  '>0 becuase it won't find itself in other modules

 IsUsedPublic = SilentSearch(strTest, ModuleExempt, strHomeModule) > 0

End Function

Public Function JustACommentOrBlank(ByVal VarSearch As Variant) As Boolean

  'copright 2003 Roger Gilchrist
  'detect comments and empty strings

 TestLineSuspension VarSearch
 JustACommentOrBlank = MultiLeft(Trim$(VarSearch), True, Apostrophe, "Rem ") Or LenB(Trim$(VarSearch)) = 0 Or SuspendCF = True
 DoEvents

End Function

Public Function LineisBlank(cmpMod As CodeModule, _
                            cdeline As Long) As Boolean

 With cmpMod
  If cdeline <= .CountOfLines Then
   LineisBlank = LenB(Trim$(.Lines(cdeline, 1))) = 0
   Else
   LineisBlank = True
  End If
 End With

End Function

Private Sub MoveDimToTopOfProc(compmod As CodeModule, _
                               codeline As Long)

  Dim Pname              As String
  Dim PlineNo            As Long
  Dim PStartLine         As Long
  Dim PEndLine           As Long
  Dim prevHadLineCont    As Boolean
  Dim strtmp             As String
  Dim DimLineInsertPoint As Long
  Dim DimLineCount       As Long
  Dim I                  As Long
  Dim lJunk              As Long
  Dim strDims            As String

 GetLineData compmod, codeline, Pname, PlineNo, PStartLine, PEndLine, lJunk
 If Pname <> "(Declarations)" Then
  For I = PStartLine To PEndLine
   If MultiLeft(Trim$(compmod.Lines(I, 1)), True, "Dim ", "Static ", "Const ") Then
    DimLineCount = DimLineCount + 1
   End If
  Next I
  If DimLineCount Then
   DimLineInsertPoint = PStartLine
   strtmp = Trim$(compmod.Lines(DimLineInsertPoint, 1))
   Do While seek1stDim(strtmp, prevHadLineCont)
    DimLineInsertPoint = DimLineInsertPoint + 1
    ' this takes care of last line of line cont code
    prevHadLineCont = HasLineCont(strtmp)
    If DimLineInsertPoint = PEndLine Then
     DimLineInsertPoint = codeline ' this is a safe junk value
     Exit Do 'safety should never hit
    End If
    strtmp = Trim$(compmod.Lines(DimLineInsertPoint, 1))
   Loop
   'collect all Dim/Const/Static in procedure
   For I = DimLineInsertPoint To PEndLine
    strtmp = Trim$(compmod.Lines(I, 1))
    If MultiLeft(strtmp, True, "Dim ", "Static ", "Const ") Then
     strDims = strDims & IIf(Len(strDims), vbNewLine, vbNullString) & strtmp
     compmod.DeleteLines I, 1
     I = I - 1
    End If
   Next I
   compmod.InsertLines DimLineInsertPoint, strDims
  End If
 End If

End Sub

Public Sub MultipleInstructionLineExpansion(compmod As CodeModule, _
                                            Code As String, _
                                            codeline As Long, _
                                            LineContArray As Variant, _
                                            ByVal targetproc As String, _
                                            ByVal selendline As Long)

 Do
  Erase LineContArray
  CodeLineRead compmod, codeline, Code, LineContArray
  If Not JustACommentOrBlank(Code$) Then
   If ExpandCode(Code) Then
    'StrRep = FormatNewLine(CompMod, codeline, Code, 0)
    Erase LineContArray
    CodeLineWrite WMReplace, compmod, codeline, FormatNewLine(compmod, codeline, Code, 0), LineContArray, 0
    Else
    '       ' StrRep = FormatNewLine(CompMod, codeline, Code, 0)
    CodeLineWrite WMReplace, compmod, codeline, FormatNewLine(compmod, codeline, Code, 0), LineContArray, 0
   End If
  End If
  codeline = codeline + 1
  If codeline > compmod.CountOfLines Then
   Exit Do
  End If
 Loop While InSRange(codeline, compmod, targetproc, selendline)

End Sub

Private Function NotDuplicateHit(ByVal var1 As Variant, _
                                 ByVal var2 As Variant, _
                                 ByVal CurSize As Long) As Boolean

 If CurSize Then
  NotDuplicateHit = Not (var1(0) = var2(0))
  If Not NotDuplicateHit Then
   NotDuplicateHit = Not (var1(3) = var2(3))
  End If
  Else
  NotDuplicateHit = True
 End If

End Function

Private Function OkToDo(ByVal PropReplace As Boolean, _
                        ByVal Tval As String, _
                        ByVal PtargetPos As Long, _
                        ByVal Tlen As Long) As Boolean

  'This routine protects default properties from being over-applied
  'depending on what's being replaced this routine conducts different tests
  'If PropReplace is True then it looks for a space after the target word  or being at end of string
  'this stops it hitting the admittedly rare case of two controls having near identical names
  'ie Fred and Fred2
  'if False then it uses the other test which works for all other instances
  'ver 1.1.00 second test restructured as it was misfiring aand added Error Trap
  '

 On Error Resume Next
 If PropReplace Then
  OkToDo = Mid$(Tval, PtargetPos + Tlen, 1) = " " Or PtargetPos = Len(Tval) - Tlen + 1
  Else
  OkToDo = Mid$(Tval, PtargetPos + Tlen, 1) <> "." Or Mid$(Tval, PtargetPos + Tlen - 1, 2) = " ." Or PtargetPos = Len(Tval) - Tlen + 1
 End If
 On Error GoTo 0

End Function

Private Function RandomString(ByVal iLowerBoundAscii As Long, _
                              ByVal iUpperBoundAscii As Long, _
                              ByVal lLowerBoundLength As Long, _
                              ByVal lUpperBoundLength As Long) As String

  Dim sHoldString As String
  Dim LLength     As Long
  Dim LCount      As Long

 '      --Eric Lynn, Ballwin, Missouri
 '        VBPJ TechTips 7th Edition
 'Verify boundaries
 iLowerBoundAscii = KeepBetweenLng(0, iLowerBoundAscii, 255)
 iUpperBoundAscii = KeepBetweenLng(0, iUpperBoundAscii, 255)
 If lLowerBoundLength < 0 Then
  lLowerBoundLength = 0
 End If
 'Set a random length
 LLength = Int((CDbl(lUpperBoundLength) - CDbl(lLowerBoundLength) + 1) * Rnd + lLowerBoundLength)
 'Create the random string
 For LCount = 1 To LLength
  sHoldString = sHoldString & Chr$(Int((iUpperBoundAscii - iLowerBoundAscii + 1) * Rnd + iLowerBoundAscii))
 Next LCount
 RandomString = sHoldString

End Function

Public Sub ReduceScope(ByVal strTest As String, _
                       strCode As String, _
                       ByVal TargetMod As String, _
                       Hit As Boolean)

  Dim DoneIt As String

 If IsUsedPublic(strTest, TargetMod) Then
  Hit = True
  Else
  If IsUsedPrivate(strTest, TargetMod) Then
   If MultiLeft(strCode, True, "Public ") Then
    DoneIt = RGSignature & "Public Reduced to Private"
    strCode = Replace$(strCode, "Public ", "Private ", 1, 1) & DoneIt
   End If
   Hit = True
   Else
   If ArrayMember(SecondWord(strCode), "Function", "Sub", "Property") Then
    DoneIt = RGSignature & "Unused Procedure Detected"
    Else
    DoneIt = RGSignature & "Unused Variable commented out"
    strCode = "''" & strCode
   End If
   strCode = strCode & DoneIt
   AddToSearchBox Mid$(DoneIt, 3), True
   '    End If
   Hit = False
  End If
 End If

End Sub

Public Sub RestoreMemberAttributes(Membs As Members)

  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
  'restore the member attributes
  
  Dim I                                As Long
  Dim MemberAttributes()               As Variant

 'vbArray o
 MemberAttributes = Attributes
 For I = 1 To UBound(MemberAttributes)
  Err.Clear
  On Error Resume Next
  With Membs(MemberAttributes(I)(MemName))
   'may produce an error on undo when member attributes cannot be restored
   'because a new member was created after the last format scan and thats
   'now missing in the undo buffer but it's attributes have been saved
   If Err.Number = 0 Then
    If LenB(MemberAttributes(I)(MemCate)) Then
     .Category = MemberAttributes(I)(MemCate)
    End If
    If LenB(MemberAttributes(I)(MemDesc)) Then
     .Description = MemberAttributes(I)(MemDesc)
    End If
    If MemberAttributes(I)(MemHelp) Then
     .HelpContextID = MemberAttributes(I)(MemHelp)
    End If
    If LenB(MemberAttributes(I)(MemProp)) Then
     .PropertyPage = MemberAttributes(I)(MemProp)
    End If
    If MemberAttributes(I)(MemStme) <= 0 Then
     .StandardMethod = MemberAttributes(I)(MemStme)
    End If
    If MemberAttributes(I)(MemBind) Then
     .Bindable = True
    End If
    If MemberAttributes(I)(MemBrws) Then
     .Browsable = True
    End If
    If MemberAttributes(I)(MemDfbd) Then
     .DefaultBind = True
    End If
    If MemberAttributes(I)(MemDbnd) Then
     .DisplayBind = True
    End If
    If MemberAttributes(I)(MemHidd) Then
     .Hidden = True
    End If
    If MemberAttributes(I)(MemRqed) Then
     .RequestEdit = True
    End If
    If MemberAttributes(I)(MemUide) Then
     .UIDefault = True
    End If
   End If
  End With
  On Error GoTo 0
 Next I

End Sub

Private Function safe_InStr(ByVal StartPos As Long, _
                            ByVal VarSearch As Variant, _
                            ByVal varFind As Variant, _
                            ByVal Standard As Boolean) As Long

  'extends Instr so that it only finds real words

 safe_InStr = InStr(StartPos, VarSearch, varFind)
 If Not Standard Then
  Select Case safe_InStr
   Case 0
   'nothing
   Case 1
   'left edge on string
   If InStr(" .,!", Mid$(VarSearch, safe_InStr + Len(varFind), 1)) Then
    'its OK
    Else
    safe_InStr = 0
   End If
   Case Len(VarSearch) - Len(varFind) + 1
   'right edge of string
   If InStr(" .,!", Mid$(VarSearch, safe_InStr - 1, 1)) Then
    'its OK
    Else
    safe_InStr = 0
   End If
   Case Else
   'anywhere else in string
   If (InStr(" .,!", Mid$(VarSearch, safe_InStr - 1, 1)) And InStr(" .,!", Mid$(VarSearch, safe_InStr + Len(varFind), 1))) Then
    'its OK
    Else 'NOT INSTR(" .,!",...
    safe_InStr = 0
   End If
  End Select
 End If

End Function

Public Function Safe_Replace(Expression As Variant, _
                             Find As Variant, _
                             VarReplace As Variant, _
                             Optional Start As Long = 1, _
                             Optional Count As Integer = -1, _
                             Optional Standard As Boolean = True, _
                             Optional PropReplace As Boolean = False) As String

  Dim PossibleTarget As Long
  Dim LocalCount     As Long
  Dim RepOffSet      As Long

 'update PropReplace causes a different test to be conducted see OkToDo for details
 'Safe_Replace is designed to replace only CODE, Comments, Literal Strings and date literals cannot be touched
 Safe_Replace = Expression
 If LenB(Safe_Replace) > 0 Then
  'Ver 1.1.00 Speed up. Skips idiot case were the find and replace are the same
  If Find <> VarReplace Then
   'found coming from Dimformat (Fixed there but could be from other places so added safety here too)
   PossibleTarget = safe_InStr(Start, Safe_Replace, Find, Standard)
   If PossibleTarget Then
    RepOffSet = Len(Find)
    If RepOffSet < Len(VarReplace) Then
     RepOffSet = Len(VarReplace)
    End If
    DisguiseLiteral Expression, Find, True
    Do While PossibleTarget
     If OkToDo(PropReplace, Safe_Replace, PossibleTarget, Len(Find)) Then
      If InCode(Safe_Replace, PossibleTarget) Then
       Safe_Replace = Left$(Safe_Replace, PossibleTarget - 1) & VarReplace & Mid$(Safe_Replace, PossibleTarget + Len(Find))
       LocalCount = LocalCount + 1
      End If
     End If
     PossibleTarget = safe_InStr(PossibleTarget + RepOffSet + 1, Safe_Replace, Find, Standard)
     If Count > 0 Then
      If LocalCount >= Count Then
       Exit Do
      End If
     End If
    Loop
    DisguiseLiteral Expression, Find, False
   End If
   Else
  End If
 End If

End Function

Public Sub SaveMemberAttributes(Membs As Members)

  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
  
  Dim I                                As Long
  Dim Member                           As Member
  Dim MemberAttributes()               As Variant

 'vbArray o
 I = 0
 ReDim MemberAttributes(0 To Membs.Count)
 On Error Resume Next
 For Each Member In Membs
  I = I + 1
  Err.Clear
  With Member
   MemberAttributes(I) = Array(.Name, .Bindable, .Browsable, .Category, .DefaultBind, .Description, .DisplayBind, .HelpContextID, .Hidden, .PropertyPage, .RequestEdit, .StandardMethod, _
                               .UIDefault)
   If Err.Number Then
    I = I = 1
   End If
  End With
 Next Member
 On Error GoTo 0
 ReDim Preserve MemberAttributes(0 To I)
 Attributes = MemberAttributes

End Sub

Public Function ScopeChangeTest(strCode As String, _
                                ByVal TargetMod As String, _
                                ByVal targetproc As String) As Boolean

  Dim strTest As String
  Dim DoneIt  As String

 If ScopeTestable(strCode) Then
  If Not HasScope(strCode) Then
   'force maximum Scope and then reduce it
   If ScopeRestrictToPrivate(strCode, TargetMod) Then
    DoneIt = RGSignature & "UnScoped " & IIf(isProcHead(strCode), "Procedure", "Variable") & " changed to Private"
    strCode = "Private " & strCode & DoneIt
    Else
    DoneIt = RGSignature & "UnScoped " & IIf(isProcHead(strCode), "Procedure", "Variable") & " changed to Public"
    strCode = "Public " & strCode & DoneIt
   End If
   AddToSearchBox Mid$(DoneIt, 3), True
  End If
  If ArrayMember(SecondWord(strCode), "Enum", "Type") Then
   'exit function these are dealt with else where
   ScopeChangeTest = True
   Else
   Select Case LeftWord(strCode)
    Case "Global", "Public"
    If ScopeRestrictToPrivate(strCode, TargetMod) Then
     DoneIt = RGSignature & "Public Reduced to Private"
     strCode = Replace$(strCode, "Public ", "Private ", 1, 1)
     strCode = strCode & DoneIt
     AddToSearchBox Mid$(DoneIt, 3), True
    End If
    UpdateGlobal strCode, ScopeChangeTest
    strTest = getScopeTestWord(strCode)
    If StandardControlProcedure(strTest) Then
     'fake for Control Procedures
     'this will be improved soon
     ScopeChangeTest = True
     Else
     ReduceScope strTest, strCode, TargetMod, ScopeChangeTest
    End If
    Case "Private"
    strTest = getScopeTestWord(strCode)
    If StandardControlProcedure(strTest) Then
     'fake for Control Procedures
     'this will be improved soon
     ScopeChangeTest = True
     Else
     ReduceScope strTest, strCode, TargetMod, ScopeChangeTest
    End If
    Case "Friend"
    'just accept it
    ScopeChangeTest = True
    Case "Static"
    ScopeChangeTest = True
    If targetproc = "(Declarations)" Then
     ScopeChangeTest = True
     Else
     DoneIt = RGSignature & "Suggestion: Change to Module Level variable, Static uses about 3X more memory"
     strCode = strCode & DoneIt
     AddToSearchBox Mid$(DoneIt, 3), True
     ScopeChangeTest = True
    End If
    Case "Dim"
    strTest = getScopeTestWord(strCode)
    If targetproc = "(Declarations)" Then
     DoneIt = RGSignature & "Dim changed to Public"
     strCode = Replace$(strCode, "Dim ", "Public ", 1, 1) & DoneIt
     AddToSearchBox Mid$(DoneIt, 3), True
     ScopeChangeTest = True
     If Not StandardControlProcedure(strTest) Then
      ReduceScope strTest, strCode, TargetMod, ScopeChangeTest
     End If
     Else
     If IsUsedProcedure(strTest, TargetMod, targetproc) Then
      ScopeChangeTest = True
      Else
      DoneIt = RGSignature & "Unused Variable commented out"
      strCode = "''" & strCode & DoneIt
      AddToSearchBox Mid$(DoneIt, 3), True
     End If
    End If
   End Select
  End If
 End If

End Function

Public Function ScopeEnumTypeTest(strStruct As String, _
                                  strCode As String, _
                                  ByVal TargetMod As String) As Boolean

  Dim I             As Long
  Dim DoneIt        As String
  Dim TmpA          As Variant
  Dim strOrig       As String
  Dim strStructName As String
  Dim UsedLines     As Variant

 strOrig = strStruct
 strStructName = IIf(InstrAtPosition(strStruct, "Enum ", ipLeftOr2nd, True), "Enum", "Type")
 If Not HasScope(strCode) Then
  'force maximum Scope and then reduce it
  DoneIt = RGSignature & "UnScoped " & strStructName & " to Public"
  strStruct = "Public " & strStruct
  strCode = "Public " & strCode & DoneIt
  AddToSearchBox Mid$(DoneIt, 3), True
 End If
 TmpA = Split(Trim$(strStruct), vbNewLine)
 TmpA(UBound(TmpA)) = vbNullString 'ditch the end line
 TmpA(0) = getScopeTestWord(strCode)
 UsedLines = SearchAllUsage(getScopeTestWord(strCode))
 If UBound(UsedLines) > -1 Then
  For I = LBound(UsedLines) To UBound(UsedLines)
   If isProcHead(CStr(UsedLines(I))) Or InStr(UsedLines(I), " As " & TmpA(0)) Then
    ScopeEnumTypeTest = True
    If ArrayMember(LeftWord(strCode), "Private", getScopeTestWord(strCode)) Then
     DoneIt = RGSignature & "Type/Enum used as parameter must be Public"
     strCode = Replace$(strCode, "Private ", "Public ", 1, 1) & DoneIt
    End If
   End If
  Next I
  Else
 End If
 If Not ScopeEnumTypeTest Then
  For I = LBound(TmpA) To UBound(TmpA)
   TmpA(I) = LeftWord(TmpA(I))
  Next I
  If IsUsedPublic(CStr(TmpA(0)), TargetMod) Then
   If LeftWord(strCode) = "Private" Then
    DoneIt = RGSignature & strStructName & "made Public"
    strCode = Replace$(strCode, "Public ", "Private ", 1, 1) & DoneIt
   End If
   ScopeEnumTypeTest = True
   ElseIf IsUsedPrivate(CStr(TmpA(0)), TargetMod) Then
   If LeftWord(strCode) = "Public" Then
    DoneIt = RGSignature & strStructName & "made Public"
    strCode = Replace$(strCode, "Public ", "Private ", 1, 1) & DoneIt
   End If
   ScopeEnumTypeTest = True
   Else
   For I = 1 To UBound(TmpA) - 1
    If IsUsedPublic(CStr(TmpA(I)), TargetMod) Then
     If LeftWord(strCode) = "Private" Then
      DoneIt = RGSignature & strStructName & "made Private"
      strCode = Replace$(strCode, "Private ", "Public ", 1, 1) & DoneIt
      ScopeEnumTypeTest = True
      Exit For
     End If
     ElseIf IsUsedPrivate(CStr(TmpA(I)), TargetMod) Then
     If LeftWord(strCode) = "Public" Then
      DoneIt = RGSignature & strStructName & "made Private"
      strCode = Replace$(strCode, "Public ", "Private ", 1, 1) & DoneIt
      ScopeEnumTypeTest = True
      Exit For
     End If
    End If
   Next I
  End If
 End If
 If Not ScopeEnumTypeTest Then
  DoneIt = RGSignature & "Unused " & strStructName & " Structure"
  AddToSearchBox Mid$(DoneIt, 3), True
  strCode = strCode & DoneIt
 End If

End Function

Public Function ScopeRestrictToPrivate(ByVal strCode As String, _
                                       ByVal TargetMod As String) As Boolean

 If GetModuleType(TargetMod) = 5 Then
  If InStr(getScopeTestWord(strCode, True), "(") Then
   ScopeRestrictToPrivate = True
  End If
  If Not ScopeRestrictToPrivate Then
   If Get_As_Pos(strCode) Then
    strCode = Mid$(strCode, Get_As_Pos(strCode) + 4)
    If Not IsInArray(LeftWord(strCode), StandardTypes) Then
     ScopeRestrictToPrivate = True
    End If
   End If
  End If
  If Not ScopeRestrictToPrivate Then
   ScopeRestrictToPrivate = InstrAtPosition(strCode, "Const", ipLeftOr2nd, True)
  End If
 End If

End Function

Private Function ScopeTestable(strCode As String) As Boolean

 ScopeTestable = ArrayMember(WordMember(strCode, 1), "Const", "Declare", "Global", "Dim", "Public", "Private", "Friend", "Static", "Function", "Sub", "Property", "Enum", "Type")
 If Not ScopeTestable Then
  ScopeTestable = (SecondWord(strCode) = "As")
 End If

End Function

Public Function SearchAllUsage(ByVal strFind As String) As Variant

  'ver1.1.02 major rewrite to allow simple pattern searching
  
  Dim Code                      As String
  Dim Procname                  As String
  Dim ProcLineNo                As Long
  Dim compmod                   As CodeModule
  Dim Comp                      As VBComponent
  Dim Proj                      As VBProject
  Dim strStrComTest             As String
  Dim startLine                 As Long
  Dim startCol                  As Long
  Dim EndLine                   As Long
  Dim endCol                    As Long
  Dim CurProc                   As String
  Dim curModule                 As String
  Dim StartProcRow              As Long
  Dim SelstartLine              As Long
  Dim SeEndLine                 As Long
  Dim SelStartCol               As Long
  Dim SelEndCol                 As Long
  Dim TmpA()                    As Variant
  Dim bLocNoComments            As Boolean
  Dim bLocNoStrings             As Boolean
  Dim bLocCommentsOnly          As Boolean
  Dim bLocStringsOnly           As Boolean
  Dim iLocRange                 As Long
  Dim BLocPatternSearch         As Boolean

 On Error Resume Next
 'if strFind doesn't have any triggers and Pattern Search is on then turn it off
 If LenB(strFind) = 0 Then
  Exit Function
 End If
 ReDim TmpA(0) As Variant
 'save set filters
 bLocNoComments = bNoComments
 bLocNoStrings = bNoStrings
 bLocCommentsOnly = bCommentsOnly
 bLocStringsOnly = bStringsOnly
 iLocRange = iRange
 BLocPatternSearch = BPatternSearch
 'set relevant for this usage
 bNoComments = True
 bNoStrings = True
 bCommentsOnly = False
 bStringsOnly = False
 iRange = AllCode
 BPatternSearch = False
 StartProcRow = 1
ReTry:
 DoEvents
 bCancel = False
 GetCounts
 CurProc = GetCurrentProcedure
 curModule = GetCurrentModule
 For Each Proj In VBInstance.VBProjects
  For Each Comp In Proj.VBComponents
   Set compmod = Comp.CodeModule
   'Safety turns off filters if comment/double quote is actually in the search phrase
   startLine = 1 'initialize search range
   Do
    startCol = 1
    EndLine = -1
    endCol = -1
    If compmod.Find(strFind, startLine, startCol, EndLine, endCol, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch) Then
     Procname = GetProcName(compmod, startLine)
     ProcLineNo = GetProcLineNumber(compmod, startLine)
     Code$ = compmod.Lines(startLine, 1)
     'apply nostring no comment filters
     If BPatternSearch Then
      ' the string/comment filters cannot work on a PatternSearch StrFind
      'but you can get the actual string found and test that
      compmod.CodePane.SetSelection startLine, startCol, EndLine, endCol
      strStrComTest = Mid$(Code$, startCol, endCol - startCol)
      Else
      strStrComTest = strFind
     End If
     ApplyStringCommentFilters Code$, strStrComTest
     ApplySelectedTextRestriction Code$, startLine, startCol, EndLine, endCol, SelstartLine, SelStartCol, SeEndLine, SelEndCol
     If Len(Code) Then
      TmpA(UBound(TmpA)) = Code
      ReDim Preserve TmpA(UBound(TmpA) + 1) As Variant
      Code$ = vbNullString
     End If
SkipProc:
    End If
    startLine = startLine + 1
    If startLine >= compmod.CountOfLines Then
     Exit Do
    End If
   Loop While compmod.Find(strFind, startLine, 1, -1, -1, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch)
   ' End If
SkipComp:
   Set Comp = Nothing
  Next Comp
 Next Proj
 Set Proj = Nothing
 Set compmod = Nothing
 'automatically switch to pattern search if ordinary fails
 If Len(TmpA(0)) Then
  SearchAllUsage = TmpA
 End If
 'this turns auto switch to pattern search off if it was used
 On Error GoTo 0
 bNoComments = bLocNoComments
 bNoStrings = bLocNoStrings
 bCommentsOnly = bLocCommentsOnly
 bStringsOnly = bLocStringsOnly
 iRange = iLocRange
 BPatternSearch = BLocPatternSearch

End Function

Public Function seek1stDim(ByVal strCode As String, _
                           ByVal prevHadLineCont As Boolean) As Boolean

 If prevHadLineCont Then
  seek1stDim = True
  Exit Function
 End If
 If HasLineCont(strCode) Then
  seek1stDim = True
  Exit Function
 End If
 If isProcHead(strCode) Then
  seek1stDim = True
  Exit Function
 End If
 If JustACommentOrBlank(strCode) Then
  seek1stDim = True
  Exit Function
 End If
 If Not IsDimLine(strCode) Then
  seek1stDim = True
 End If

End Function

Public Function SilentSearch(ByVal strFind As String, _
                             Mode As SilentFinds, _
                             Optional ByVal TargetModule As String, _
                             Optional ByVal targetproc As String) As Long

  Dim Code                      As String
  Dim Procname                  As String
  Dim compmod                   As CodeModule
  Dim Comp                      As VBComponent
  Dim Proj                      As VBProject
  Dim strStrComTest             As String
  Dim startLine                 As Long
  Dim startCol                  As Long
  Dim EndLine                   As Long
  Dim endCol                    As Long
  Dim CurProc                   As String
  Dim curModule                 As String
  Dim StartProcRow              As Long
  Dim PrevCurCodePane           As Long
  Dim AutoRangeRevert           As Boolean
  Dim HiLitSelection            As String
  Dim SelstartLine              As Long
  Dim SeEndLine                 As Long
  Dim SelStartCol               As Long
  Dim SelEndCol                 As Long
  Dim TotalFinds                As Long
  Dim ModuleFinds               As Long
  Dim ProcCount                 As Long
  Dim DeclareCount              As Long
  Dim OrigNoComment             As Boolean
  Dim OrigNoString              As Boolean
  Dim OrigCommentOnly           As Boolean
  Dim OrigStringOnly            As Boolean
  Dim OrigBPatternSearch        As Boolean

 'store filters
 OrigNoComment = bNoComments
 OrigNoString = bNoStrings
 OrigCommentOnly = bCommentsOnly
 OrigStringOnly = bStringsOnly
 OrigBPatternSearch = BPatternSearch
 'force filters for SilentSearch
 bNoComments = True
 bNoStrings = True
 bCommentsOnly = False
 bStringsOnly = False
 BPatternSearch = False
 On Error Resume Next
 'if strFind doesn't have any triggers and Pattern Search is on then turn it off
 '  AutoPatternOff strFind
 '  If BPatternSearch Then
 '    If InStr(strFind, "\ASC(") Then
 '      strFind = ConvertAsciiSearch(strFind)
 '      AddToSearchBox strFind
 '    End If
 '  End If
 If LenB(strFind) = 0 Then
  Exit Function
 End If
 StartProcRow = 1
 SilentSearch = 0 'default value
ReTry:
 If LenB(strFind) > 0 Then
  bCancel = False
  GetCounts
  CurProc = GetCurrentProcedure
  curModule = GetCurrentModule
  'this does the auto switching if multiple code lines are selected
  AutoSelectInitialize PrevCurCodePane, AutoRangeRevert
  ' this sets search limits if there is selected code.
  If iRange = SelCode Then
   HiLitSelection$ = GetSelectedText(VBInstance)
   VBInstance.ActiveCodePane.GetSelection SelstartLine, SelStartCol, SeEndLine, SelEndCol
  End If
  For Each Proj In VBInstance.VBProjects
   For Each Comp In Proj.VBComponents
    If SafeCompToProcess(Comp) Then
     If Mode = ModuleOnly Or Mode = CurProcOnly Then
      If Comp.Name <> TargetModule Then
       GoTo SkipComp
      End If
     End If
     If Mode = ModuleExempt Then
      If Comp.Name = TargetModule Then
       GoTo SkipComp
      End If
     End If
     Set compmod = Comp.CodeModule
     startLine = 1 'initialize search range
     Do
      startCol = 1
      EndLine = -1
      endCol = -1
      If compmod.Find(strFind, startLine, startCol, EndLine, endCol, True, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch) Then
       DoEvents
       If Mode = CurProcOnly Then
        Procname = GetProcName(compmod, startLine)
        If Procname <> targetproc Then
         GoTo SkipProc
        End If
       End If
       Code$ = compmod.Lines(startLine, 1)
       'apply nostring no comment filters
       '                If BPatternSearch Then
       '                  ' the string/comment filters cannot work on a PatternSearch StrFind
       '                  'but you can get the actual string found and test that
       '                  CompMod.CodePane.SetSelection StartLine, StartCol, EndLine, EndCol
       '                  strStrComTest = Mid$(code$, StartCol, EndCol - StartCol)
       '                 Else
       '                  strStrComTest = strFind
       '                End If
       ApplyStringCommentFilters Code$, strStrComTest
       ApplySelectedTextRestriction Code$, startLine, startCol, EndLine, endCol, SelstartLine, SelStartCol, SeEndLine, SelEndCol
       If LenB(Code$) Then
        '''
        'Got one so determine type
        TotalFinds = TotalFinds + 1
        ModuleFinds = ModuleFinds + 1
        ProcCount = ProcCount + 1
        DeclareCount = DeclareCount + 1
        SilentSearch = TotalFinds
        ''None
        ''OnlyOnce
        ''DeclarationOnly
        ''PrivateMod
        ''PublicMod
        ''CurProcOnly
        ''SelTextOnly
        '''
       End If
SkipProc:
      End If
      startLine = startLine + 1
      If startLine >= compmod.CountOfLines Then
       Exit Do
      End If
     Loop While compmod.Find(strFind, startLine, 1, -1, -1, True, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch)
    End If
SkipComp:
    Set Comp = Nothing
   Next Comp
  Next Proj
  Set Proj = Nothing
  Set compmod = Nothing
 End If
 'Restore filters
 bNoComments = OrigNoComment
 bNoStrings = OrigNoString
 bCommentsOnly = OrigCommentOnly
 bStringsOnly = OrigStringOnly
 BPatternSearch = OrigBPatternSearch
 On Error GoTo 0

End Function

Private Sub SingleLineStructure(ByVal strCdeOnly As String, _
                                ByVal StrTrig As String, _
                                IndntLevel As Long, _
                                DoJump As Boolean, _
                                DoJump2 As Boolean, _
                                SetCondComp As Boolean)

  Dim EndStructure As Variant
  Dim PosA         As Long

 If ArrayMember(StrTrig, "#If", "#Else", "#ElseIf") Then
  If StrTrig = "#If" Then
   SetCondComp = True
  End If
  DoJump = InStr(strCdeOnly, ": #End If") = 0
  If DoJump Then
   SetCondComp = False
  End If
 End If
 If ArrayMember(StrTrig, "If") Then
  'do not increase indent if single line of format 'If X Then Y'
  DoJump = LastWord(strCdeOnly) = "Then"
 End If
 EndStructure = Array("End With", "Next", "Loop", "Wend")
 'If a line starts with strTrig then the next line should be indented
 'Unless the code line also contains the EndStructure Member
 PosA = ArrayMemberPosition(StrTrig, "With", "For", "Do", "While")
 If PosA > -1 Then
  DoJump = InStr(strCdeOnly, ": " & EndStructure(PosA)) = 0
 End If
 'if next line should be pushed out
 ' but strTrig is one of these sub-structure members then real indent is
 ' to first move back one step for current line
 If ArrayMemberPosition(StrTrig, "Else", "ElseIf", "Case") > -1 Then
  DoJump = True
  IndntLevel = IndntLevel - 1
 End If
 If ArrayMember(StrTrig, "Select") Then
  DoJump2 = InStr(strCdeOnly, ": End Select") = 0 And InStr(strCdeOnly, ": Case") = 0
 End If
 If ArrayMember(StrTrig, "#End", "End", "Next", "Loop", "Wend") Then
  If StrTrig = "#End If" Then
   SetCondComp = False
  End If
  If strCdeOnly <> "End" Then
   'protects the End command
   IndntLevel = IndntLevel - 1
   If ArrayMember(SecondWord(strCdeOnly), "Select") Then
    IndntLevel = IndntLevel - 1
   End If
  End If
 End If

End Sub

Public Function SortingTagExtraction(cmpMod As CodeModule) As Variant

  Dim TmpString1        As String
  Dim I                 As Long
  Dim J                 As Long
  Dim K                 As Long
  Dim CleanElems()      As Variant
  Dim CleanElem         As Variant
  Dim TmpA              As Variant
  Dim FakeTopofRoutine  As Long
  Dim FakeEndofRoutine  As Long
  Dim FakeNameofRoutine As String

 With cmpMod
  .CodePane.TopLine = 1
  ReDim CleanElems(0)
  'collect module descriptions -> (Name, StartingLine, Length)
  If .Members.Count Then
   For J = 1 To .Members.Count
    With .Members(J)
     TmpString1 = .Name
     I = (.Type = vbext_mt_Property Or .Type = vbext_mt_Method)
    End With
    If I Then
     For I = 1 To 4
      K = Choose(I, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set, vbext_pk_Proc)
      CleanElem = Null
      On Error Resume Next
      'IF you crash here first check that Error Trapping is not ON
      CleanElem = Array(TmpString1, cmpMod.PRocStartLine(TmpString1, K), cmpMod.ProcCountLines(TmpString1, K), K)
      On Error GoTo 0
      If Not IsNull(CleanElem) Then
       If NotDuplicateHit(CleanElem, CleanElems(UBound(CleanElems)), UBound(CleanElems)) Then
        ReDim Preserve CleanElems(UBound(CleanElems) + 1)
        CleanElems(UBound(CleanElems)) = CleanElem
       End If
      End If
     Next I
    End If
   Next J
   If UBound(CleanElems) > -1 Then
    If Not IsEmpty(CleanElems(UBound(CleanElems))) Then
     If CleanElems(UBound(CleanElems))(1) + CleanElems(UBound(CleanElems))(2) < .CountOfLines - .CountOfDeclarationLines Then
      CleanElems(UBound(CleanElems))(2) = CleanElems(UBound(CleanElems))(1) + CleanElems(UBound(CleanElems))(2) + .CountOfLines - CleanElems(UBound(CleanElems))(1) - .CountOfDeclarationLines
      CleanElems(UBound(CleanElems))(2) = CleanElems(UBound(CleanElems))(2) + 1
     End If
    End If
   End If
   Else
   'This is a slow way of doing the above for the special case that all the code
   'is enclosed by optional compilation structure '#If <var> Then' and '#End If'
   If .CountOfLines - .CountOfDeclarationLines Then
    For I = .CountOfDeclarationLines + 1 To .CountOfLines
     TmpString1 = .Lines(I, 1)
     If I = .CountOfLines Then
      If MultiLeft(TmpString1, True, "#End If") Then
       CleanElems(UBound(CleanElems)) = Array(FakeNameofRoutine, FakeTopofRoutine, I)
       Exit For '>---> Next
      End If
     End If
     If InstrAtPositionArray(TmpString1, ipLeftOr2nd, True, "Sub", "Function", "Property") Then
      FakeTopofRoutine = I
      TmpA = Split(ExpandForDetection(TmpString1))
      If InstrAtPositionArray(TmpString1, IpLeft, True, "Sub", "Function", "Property") Then
       FakeNameofRoutine = TmpA(1)
       Else
       FakeNameofRoutine = TmpA(2)
      End If
      FakeEndofRoutine = FakeTopofRoutine
      Do
       FakeEndofRoutine = FakeEndofRoutine + 1
      Loop Until MultiLeft(.Lines(FakeEndofRoutine, 1), True, "End Sub", "End Function", "End Property")
      CleanElem = Array(FakeNameofRoutine, FakeTopofRoutine, FakeEndofRoutine - FakeTopofRoutine)
      If Not IsNull(CleanElem) Then
       ReDim Preserve CleanElems(UBound(CleanElems) + 1)
       CleanElems(UBound(CleanElems)) = CleanElem
       I = FakeEndofRoutine
      End If
     End If
    Next I
    If FakeTopofRoutine + FakeEndofRoutine - FakeTopofRoutine < .CountOfLines - .CountOfDeclarationLines Then
     CleanElems(UBound(CleanElems))(2) = FakeEndofRoutine - FakeTopofRoutine + .CountOfLines - FakeTopofRoutine - .CountOfDeclarationLines
    End If
   End If
  End If
 End With
 SortingTagExtraction = CleanElems

End Function

Public Function StandardControlProcedure(strTst As String) As Boolean

  'this is just a dummy to stop single reference control procedures
  'attached to controls, classes or other standard VB stuff
  'are not seen as unused
  'will install more accurate test later

 StandardControlProcedure = InStr(strTst, "_") Or MultiLeft(LCase$(strTst), False, "main", "form", "class", "usercontrol", "userdocument", "addininstance")

End Function

Private Function StringConcatenationUpDate(strCode As String) As Boolean

  Dim strWork        As String
  Const DoneIt       As String = RGSignature & "Sting Concatenation fixed"
  Dim CommentStore   As String

 'Copyright 2003 Roger Gilchrist
 'e-mail: rojagilkrist@hotmail.com
 'fixes old style string concatenation '+' to use  safer '&'
 'NOTE This routine copes with line continuation if the " is before the line cont character
 'but NOT if form is:
 '
 '               non-String-Variable + '               "String Literal"
 '
 On Error GoTo BadError
 If InstrAtPositionArray(strCode, ipAny, False, Chr$(34) & StrPlus, StrPlus & Chr$(34)) Then
  strWork = strCode
  CommentStore = CommentClip(strWork)
  DisguiseLiteral strWork, StrPlus, True
  strWork = Replace$(strWork, Chr$(34) & StrPlus, Chr$(34) & " & ")
  strWork = Replace$(strWork, StrPlus & Chr$(34), " & " & Chr$(34))
  DisguiseLiteral strWork, StrPlus, False
  If strCode <> strWork & CommentStore Then
   strCode = strWork & CommentStore & IIf(bShowFixComment, DoneIt, vbNullString) & IIf(bShowPrevCode, PrevCode & strCode, vbNullString)
   AddToSearchBox Mid$(DoneIt, 3), True
   StringConcatenationUpDate = True
  End If
 End If

Exit Function

BadError:

End Function

Public Sub TestLineSuspension(VarSearch As Variant)

 If InStr(1, VarSearch, InCodeDontTouchOn, vbBinaryCompare) Then
  SuspendCF = True
  ElseIf InStr(1, VarSearch, InCodeDontTouchOff, vbBinaryCompare) Then
  SuspendCF = False
 End If
 'this turns Suspend back on if you reach the end of a routine
 If SuspendCF Then
  If MultiLeft(VarSearch, False, "End ") Then
   SuspendCF = False
  End If
 End If

End Sub

Public Sub TypeSuffixExtender(strCode As String)

  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'UPDATE 13 Jan 2003 total rewrite; no longer needs supporting routines
  
  Dim BracketPos     As Long
  Dim TSPos          As Long
  Dim I              As Long
  Dim ConstValue     As String
  Dim strWork        As String
  Dim CommentStore   As String
  Dim Num            As String
  Dim numcut         As Long
  Dim DoneIt         As String

 On Error GoTo BadError
 strWork = strCode
 CommentStore = CommentClip(strWork)
 'If constant the detach the value so that it can be reattached later
 'otherwise & in string values gets 'updated and destroyed
 If InStr(strWork, "Const") Then
  If InCode(strWork, InStr(strWork, "Const")) Then
   ConstValue = Trim$(Mid$(strWork, InStr(strWork, " =")))
   strWork = Left$(strWork, InStr(strWork, " =") - 1)
  End If
 End If
 For I = 0 To 5
  TSPos = InStr(strWork, TypeSuffixArray(I))
  Do While TSPos
   If InCode(strWork, TSPos) Then
    numcut = InStrRev(strWork, " ", TSPos)
    If numcut Then
     Num$ = Mid$(strWork, numcut, TSPos - numcut)
     If IsNumeric(Num$) Then
      Mid(strWork, TSPos, 1) = " "
     End If
    End If
   End If
   TSPos = InStr(TSPos + 1, strWork, TypeSuffixArray(I))
  Loop
 Next I
 For I = 0 To 5
  TSPos = InStr(strWork, TypeSuffixArray(I))
  If TSPos = Len(strWork) Or InStr(" ,()", Mid$(strWork, TSPos + 1, 1)) Then
   If InCode(strWork, TSPos) Then
    'Last character or followed by a space, comma or left bracket
    'will not attack older DB referencing style of DB!Table
    Do
     If TSPos Then
      If TSPos < Len(strWork) Then
       If Mid$(strWork, TSPos + 1, 1) = "(" Then
        If GetLeftBracketPos(Left$(strWork, TSPos)) = 0 Then
         BracketPos = InStrRev(strWork, ")")
         Else
         BracketPos = GetRightBracketPos(strWork)
        End If
        strWork = Left$(strWork, BracketPos) & " As " & AsTypeArray(I) & Mid$(strWork, BracketPos + 1)
        Mid$(strWork, TSPos) = " "
        ElseIf Mid$(strWork, TSPos + 1, 4) = " Lib" Then
        BracketPos = InStrRev(strWork, ")")
        strWork = Left$(strWork, BracketPos) & " As " & AsTypeArray(I) & Mid$(strWork, BracketPos + 1)
        Mid$(strWork, TSPos) = " "
        ElseIf Mid$(strWork, TSPos - 1, 1) = " " Then
        ' do nothing
        Else
        strWork = Left$(strWork, TSPos - 1) & " As " & AsTypeArray(I) & Mid$(strWork, TSPos + 1)
       End If
       Else
       strWork = Left$(strWork, TSPos - 1) & " As " & AsTypeArray(I) & Mid$(strWork, TSPos + 1)
      End If
     End If
     TSPos = InStr(TSPos + 1, strWork, TypeSuffixArray(I))
    Loop While TSPos
   End If
  End If
 Next I
 If LenB(ConstValue) Then
  strWork = strWork & " " & ConstValue
 End If
 If strCode <> strWork & CommentStore Then
  ' TypeSuffixExtender = True
  DoneIt = RGSignature & "Obsolete Type Suffix replaced."
  strCode = strWork & CommentStore & DoneIt
  AddToSearchBox Mid$(DoneIt, 3), True
 End If

Exit Sub

BadError:

End Sub

Public Sub UlliQuickSort(ixFrom As Long, _
                         ixThru As Long, _
                         KeyIsIn As Long)

  'Lifted from Ulli's Code Formatter
  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
  'Sorts a table of vbArrays
  
  Dim ixLeft   As Long
  Dim ixRite   As Long
  Dim TempElem As Variant

 'we have something to sort (@ least two elements)
 If ixFrom < ixThru Then
  ixLeft = ixFrom
  ixRite = ixThru
  'Get ref element and make room
  TempElem = SortElems(ixLeft)
  Do
   Do Until ixRite = ixLeft
    If LCase$(SortElems(ixRite)(KeyIsIn)) >= LCase$(TempElem(KeyIsIn)) Then
     ixRite = ixRite - 1
     Else 'is smaller than ref so move it to the left...
     SortElems(ixLeft) = SortElems(ixRite)
     '...and leave the item just moved alone for now
     ixLeft = ixLeft + 1
     Exit Do '>---> Loop
    End If
   Loop
   Do Until ixLeft = ixRite
    If LCase$(SortElems(ixLeft)(KeyIsIn)) <= LCase$(TempElem(KeyIsIn)) Then
     ixLeft = ixLeft + 1
     Else 'is greater than ref so move it to the right...
     SortElems(ixRite) = SortElems(ixLeft)
     '...and leave the item just moved alone for now
     ixRite = ixRite - 1
     Exit Do '>---> Loop
    End If
   Loop
  Loop Until ixLeft = ixRite
  'now the indexes have met and all bigger items are to the right and all smaller items are left
  SortElems(ixRite) = TempElem 'Insert ref elem in proper place and sort the two areas left and right of it
  'smaller part 1st to reduce recursion depth
  If ixLeft - ixFrom > ixThru - ixRite Then
   UlliQuickSort ixRite + 1, ixThru, KeyIsIn
   UlliQuickSort ixFrom, ixLeft - 1, KeyIsIn
   Else
   UlliQuickSort ixFrom, ixLeft - 1, KeyIsIn
   UlliQuickSort ixRite + 1, ixThru, KeyIsIn
  End If
 End If

End Sub

Private Sub UllisSort(cmpMod As CodeModule)

  'Lifted from Ulli's Code Formatter
  '© 2000/2002 UMGEDV GmbH  (umgedv@aol.com)
  
  Dim J            As Long
  Dim TmpString1   As String
  Dim I            As Long
  Dim K            As Long
  Dim TempElem     As Variant                                                                                                                                                                           'one element as temporary
  Dim strCode      As String

 ReDim SortElems(0)
 With cmpMod
  For J = 1 To .Members.Count
   With .Members(J)
    TmpString1 = .Name
    I = (.Type = vbext_mt_Property Or .Type = vbext_mt_Method)
   End With
   If I Then
    For I = 1 To 4
     K = Choose(I, vbext_pk_Proc, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set) 'determines seq of equal named modules
     TempElem = Null
     On Error Resume Next
     'If you have Error Checking set to Break on all Errors this will crash
     TempElem = Array(TmpString1, .PRocStartLine(TmpString1, K), .ProcCountLines(TmpString1, K))
     On Error GoTo 0
     If Not IsNull(TempElem) Then
      ReDim Preserve SortElems(UBound(SortElems) + 1)
      SortElems(UBound(SortElems)) = TempElem
     End If
    Next I
   End If
  Next J
  UlliQuickSort 1, UBound(SortElems), 0
  'build sorted component
  TmpString1 = vbNullString
  For I = 1 To UBound(SortElems)
   Select Case I
    Case 1
    'Sub or Function
    If SortElems(I)(1) > .CountOfDeclarationLines Then
     TmpString1 = TmpString1 & .Lines(SortElems(I)(1), SortElems(I)(2)) & vbNewLine
    End If
    Case Else
    'there's a quirk in VB: it returns Events as methods and if an
    'Event has the same name as a Sub/Function then this results in
    'duplicates, so here duplicates are filtered out
    If SortElems(I)(1) <> SortElems(I - 1)(1) Then
     'Sub or Function
     If SortElems(I)(1) > .CountOfDeclarationLines Then
      TmpString1 = TmpString1 & .Lines(SortElems(I)(1), SortElems(I)(2)) & vbNewLine
     End If
    End If
   End Select
  Next I
  'delete original modules
  .DeleteLines .CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines
  'add sorted modules
  .AddFromString TmpString1
  'remove trailing blank lines if any
  Do
   strCode = Trim$(.Lines(.CountOfLines, 1))
   If Len(strCode) = 0 Then
    .DeleteLines .CountOfLines
   End If
  Loop Until Len(strCode)
 End With

End Sub

Public Sub UnDoAction(ByVal Mode As Integer)

  Dim Proj         As VBProject
  Dim Comp         As VBComponent
  Dim compmod      As CodeModule
  Dim CurCompCount As Long
  Dim I            As Long
  Dim J            As Long

 mobjDoc.ShowWorking True, "Restoring Undo Data...", , False
 For Each Proj In VBInstance.VBProjects
  CompCount = Proj.VBComponents.Count * VBInstance.VBProjects.Count
  For Each Comp In Proj.VBComponents
   If SafeCompToProcess(Comp) Then
    CurCompCount = CurCompCount + 1
    Set compmod = Comp.CodeModule
    With frmUnDo.lstUnDo
     For I = .ListCount - 1 To 0 Step -1
      If .Selected(I) Or Mode = 1 Then
       For J = LBound(UndoArray) To UBound(UndoArray)
        If .List(I) = UndoArray(J).strModulename Then
         With compmod
          .CodePane.Show
          .DeleteLines 1, .CountOfLines
          .AddFromString UndoArray(J).strModule
         End With 'CompMod
         Attributes = UndoArray(J).MemberData
         RestoreMemberAttributes compmod.Members
         .RemoveItem I
         Exit For
        End If
       Next J
      End If
     Next I
    End With 'lstUndone
   End If
   Set compmod = Nothing
  Next Comp
 Next Proj
 Set Comp = Nothing
 Set Proj = Nothing
 mobjDoc.ShowWorking False

End Sub

Public Sub UnDoListInit()

  Dim Proj                             As VBProject
  Dim Comp                             As VBComponent
  Dim compmod                          As CodeModule
  Dim CurCount                         As Long
  Dim CurCompCount                     As Long

 frmUnDo.lstUnDo.Clear
 mobjDoc.ShowWorking True, "Saving Undo Data...", , False
 For Each Proj In VBInstance.VBProjects
  CompCount = Proj.VBComponents.Count * VBInstance.VBProjects.Count
  ReDim Preserve UndoArray(CompCount) As UnDoData
  For Each Comp In Proj.VBComponents
   If SafeCompToProcess(Comp) Then
    CurCompCount = CurCompCount + 1
    Set compmod = Comp.CodeModule
    With UndoArray(CurCompCount)
     .strModulename = Proj.Name & "-" & Comp.Name
     .strModule = compmod.Lines(1, compmod.CountOfLines)
     SaveMemberAttributes compmod.Members
     .MemberData = Attributes
     frmUnDo.lstUnDo.AddItem .strModulename
    End With
    CurCount = CurCount + 1
   End If
  Next Comp
 Next Proj
 mobjDoc.ShowWorking False

End Sub

Public Function UnNeededExit(cmpMod As CodeModule, _
                             ByVal cdeline As Long, _
                             strCode As String) As Boolean

  Dim DoneIt  As String
  Dim TmpLine As Long

 TmpLine = cdeline + 1
 DoneIt = RGSignature & "Unneeded Exit"
 TmpLine = GetNextCodeLine(cmpMod, TmpLine)
 If InstrAtPositionArray(cmpMod.Lines(TmpLine, 1), ipAny, True, "End Sub", "End Function", "End Property") Then
  strCode = strCode & DoneIt
  AddToSearchBox Mid$(DoneIt, 3), True
  UnNeededExit = True
  ElseIf InstrAtPositionArray(cmpMod.Lines(TmpLine, 1), ipAny, True, "End If") Then
  TmpLine = TmpLine + 1
  TmpLine = GetNextCodeLine(cmpMod, TmpLine)
  If InstrAtPositionArray(cmpMod.Lines(TmpLine, 1), IpLeft, True, "End Sub", "End Function", "End Property") Then
   strCode = strCode & DoneIt
   UnNeededExit = True
   AddToSearchBox Mid$(DoneIt, 3), True
  End If
 End If

End Function

Private Sub UpdateGlobal(strCode As String, _
                         Hit As Boolean)

 If LeftWord(strCode) = "Global" Then
  'This automatically upgrades Global to Public
  strCode = Replace$(strCode, "Global ", "Public ", 1, 1) & RGSignature & "Global changed to Public"
  Hit = True
 End If

End Sub

Private Function WordIsVBSingleWordCommand(ByVal strTest As String) As Boolean

  'This is a guard routine to stop certain VB commands being detected as GoTo Targets
  'It is used by DoSeparateCompoundLines to decide whether to leave or remove the colon

 WordIsVBSingleWordCommand = ArrayMember(strTest, "Do", "While", "Loop", "Wend", "Else", "Beep")
 'VB can in fact distinguish between most of these in the format 'X:' but Beep: for some reason
 'can be used as either and VB defaults to label.
 'The code 'Beep: Beep: Beep'  only sounds twice the first one is treated as a label.
 '

End Function
Private Sub Chr2ConstantDo(strCode As String)
Dim strWork As String
Dim DoneIt As String
  Dim ArrOld     As Variant
  Dim ArrNew     As Variant
  strWork = strCode
  'converts specific Chr$ to named variables for better readability
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  'Only has to check for the $ version because variant is fixed by UpDateStringFunctions
  ArrOld = Array("vbCrLf", "Chr$(13) & Chr$(10)", "Chr$(9)", "Chr$(13)", "Chr$(10)", "Chr$(0)", "Chr$(8)")
  ArrNew = Array("vbNewline", "vbNewline", "vbTab", "vbCr", "vbLf", "vbNullChar", "vbBack")
  UpdateStringArray strWork, ArrOld, ArrNew
If strWork <> strCode Then
DoneIt = RGSignature & "Chr updated to VBConstant"
strCode = strWork & DoneIt
AddToSearchBox Mid$(DoneIt, 3), True
End If
End Sub
Private Sub UpdateStringArray(strWork As String, _
                              ArrayOld As Variant, _
                              ArrayNew As Variant)
  
  Dim MyStr        As String

  Dim StrSafe      As String
  Dim I            As Long
  Dim CommentStore As String
  'General service for updating members of ArrayOld with equivalent member of ArrayNew
  'Copyright 2003 Roger Gilchrist
  'e-mail: rojagilkrist@hotmail.com
  StrSafe = strWork
  On Error GoTo BadError

  For I = LBound(ArrayOld) To UBound(ArrayOld)

    If InstrAtPosition(strWork, ArrayOld(I), ipAny, False) Then
      MyStr = strWork
      CommentStore = CommentClip(MyStr)
      DisguiseLiteral MyStr, ArrayOld(I), True
      MyStr = Safe_Replace(MyStr, ArrayOld(I), ArrayNew(I))
      DisguiseLiteral MyStr, ArrayOld(I), False

      If strWork <> MyStr & CommentStore Then
        strWork = MyStr & CommentStore
      End If

    End If

  Next I

  Exit Sub
BadError:
  strWork = StrSafe

End Sub

':) Roja's VB Code Fixer V1.1.4 (30/07/2003 1:25:07 PM) 101 + 3321 = 3422 Lines Thanks Ulli for inspiration and lots of code.

