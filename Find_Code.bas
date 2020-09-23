Attribute VB_Name = "Find_Code"
Option Explicit
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
'This is a very slightly modified version of routines I have used in earlier versions
'Slight changes to workaround UserDocument limits
Public Enum EnumMsg
 Search
 Complete
 inComplete
 Missing
 Found
 replaced
 replacing
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private Search, Complete, inComplete, Missing, Found, replaced, replacing
#End If
Public Enum Range
 AllCode
 ModCode
 ProcCode
 SelCode
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private AllCode, ModCode, ProcCode, SelCode
#End If
'Search Settings
Public Const LongLimit               As Single = 2147483646
'Prevent FlexGrid from over-flowing;
'very unlikely that anything could be found that often
'just a safety valve inherited from the old listbox/IntegerLimit
Public bWholeWordonly                As Boolean
Public bCaseSensitive                As Boolean
Public bNoComments                   As Boolean
Public bCommentsOnly                 As Boolean
Public bNoStrings                    As Boolean
Public bStringsOnly                  As Boolean
Public iRange                        As Integer
Public BPatternSearch                As Boolean
Public bTmpShowReplace               As Boolean
Public bPTmpWholeWordonly            As Boolean
Public bPTmpCaseSensitive            As Boolean
Public bPTmpNoComments               As Boolean
Public bPTmpCommentsOnly             As Boolean
Public bPTmpNoStrings                As Boolean
Public bPTmpStringsOnly              As Boolean
Public bPTmpFindSelectWholeLine      As Boolean
Public bGridlines                    As Boolean
Public bShowCompLineNo               As Boolean
Public bShowProcLineNo               As Boolean
Public bShowReplace                  As Boolean
Private bComplete                    As Boolean
'halt search before completion
Public bCancel                       As Boolean
Private Const Apostrophe             As String = "'"
Public bShowProject                  As Boolean
Public bShowComponent                As Boolean
Public bShowRoutine                  As Boolean
Public bFindSelectWholeLine          As Boolean
Private ReplaceCount                 As Long

Public Sub ApplySelectedTextRestriction(cde As String, _
                                        ByVal FStartR As Long, _
                                        ByVal FStartC As Long, _
                                        ByVal FEndR As Long, _
                                        ByVal FEndC As Long, _
                                        ByVal SStartR As Long, _
                                        ByVal SStartC As Long, _
                                        ByVal SEndR As Long, _
                                        ByVal SEndC As Long)

  'Code is needed in DoSearch and DoReplace so a separate Procedure
  
  Dim InSRange                  As Boolean

 If iRange = SelCode Then
  If Len(cde) Then
   InSRange = False
   If BetweenLng(SStartR, FStartR, SEndR) Then
    If BetweenLng(SStartR, FEndR, SEndR) Then
     InSRange = True
     'ver 2.2.1
     'Refinement of test; only checks column if in first or last line of selected text
     ' any other line must be in the range
     If SStartR = FStartR Then
      InSRange = SStartC >= FStartC
      ElseIf SEndR = FEndR Then
      InSRange = FEndC <= SEndC
     End If
     If FStartR = FEndR Then
      'special case single line selected
      InSRange = SStartC >= FStartC And FEndC <= SEndC
     End If
    End If
   End If
   If Not InSRange Then
    cde = vbNullString
   End If
  End If
 End If

End Sub

Public Sub ApplyStringCommentFilters(cde As String, _
                                     ByVal strTarget As String)

  Dim Codepos As Long

 Codepos = InStr(1, cde, strTarget, IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare))
 If bNoComments Then
  If InComment(cde, Codepos) Then
   cde = vbNullString
  End If
 End If
 If bCommentsOnly Then
  If Not InComment(cde, Codepos) Then
   cde = vbNullString
  End If
 End If
 If bStringsOnly Then
  If InQuotes(cde, Codepos) = False Then
   cde = vbNullString
  End If
 End If
 If bNoStrings Then
  If InQuotes(cde, Codepos) Then
   cde = vbNullString
  End If
 End If

End Sub

Public Function ArrayMember(ByVal Tval As Variant, _
                            ParamArray pMembers() As Variant) As Boolean

  Dim wal As Variant

 'returns true if any member of pMembers equals Tval
 For Each wal In pMembers
  If Tval = wal Then
   ArrayMember = True
   Exit Function
  End If
 Next wal

End Function

Public Function ArrayMemberPosition(ByVal Tval As Variant, _
                                    ParamArray pMembers() As Variant) As Long

  'this procedure returns a zero based array position
  'OR -1 if not a member
  
  Dim counter As Long
  Dim wal     As Variant

 'returns true if any member of pMembers equals Tval
 For Each wal In pMembers
  If Tval = wal Then
   ArrayMemberPosition = counter
   Exit Function
  End If
  counter = counter + 1
 Next wal
 ArrayMemberPosition = -1

End Function

Public Sub AutoPatternOff(ByVal strF As String)

 If BPatternSearch Then
  If instrAny(strF, "*", "!", "[", "]", "\") = False Then
   BPatternSearch = False
   mobjDoc.ClearForPattern
  End If
 End If

End Sub

Public Sub AutoSelectInitialize(PrevRange As Long, _
                                AutoRevert As Boolean)

  'Code is needed in DoSearch and DoReplace so a separate Procedure

 If bAutoSelectText Then
  If InStr(GetSelectedText(VBInstance), vbNewLine) > 0 Then
   PrevRange = iRange
   iRange = SelCode
   AutoRevert = True
   mobjDoc.ToggleButtonFaces
  End If
 End If

End Sub

Public Function BetweenLng(ByVal MinV As Long, _
                           ByVal Val As Long, _
                           ByVal MaxV As Long, _
                           Optional ByVal InClusive As Boolean = True) As Boolean

 If InClusive Then
  If Val >= MinV Then
   If Val <= MaxV Then
    BetweenLng = True
   End If
  End If
  Else
  If Val > MinV Then
   If Val < MaxV Then
    BetweenLng = True
   End If
  End If
 End If

End Function

Public Sub BlankCleaners(CmpMod As CodeModule, _
                         ByVal KillLine As Long)

 If bDeleteDoubleBlanks Then
  BlankDelete CmpMod, KillLine + 1
 End If
 If bDeleteAllBlanks Then
  BlankDelete CmpMod, KillLine
 End If

End Sub

Public Sub ClearFGrid(Fgrd As MSFlexGrid)

 With Fgrd
  .Rows = 2
  .TextMatrix(.Row, 0) = vbNullString
  .TextMatrix(.Row, 1) = vbNullString
  .TextMatrix(.Row, 2) = vbNullString
  .TextMatrix(.Row, 3) = vbNullString
 End With

End Sub

Public Function CommentClip(VarSearch As Variant) As String

  Dim MyStr       As String
  Dim CommentPos  As Long
  Dim SpaceTest   As String
  Dim SpaceOffSet As Long

 'This code clips end comments from VarSearch
 'NOTE also Modifies VarSearch
 'UPDATE now copes with literal embedded '
 On Error GoTo BadError
 MyStr = VarSearch
 CommentPos = InStr(1, MyStr, Apostrophe)
 If CommentPos > 0 Then
  Do While InLiteral(MyStr, CommentPos)
   CommentPos = InStr(CommentPos + 1, MyStr, Apostrophe)
   If CommentPos = 0 Then
    Exit Do
   End If
  Loop
  If CommentPos > 0 Then
   CommentClip = Mid$(MyStr, CommentPos)
   MyStr = Left$(MyStr, CommentPos - 1)
   'Preserve spaces with comment if comment is offset with them
   SpaceTest = Trim$(LTrim$(MyStr))
   SpaceOffSet = Len(MyStr) - Len(SpaceTest)
   CommentClip = String$(SpaceOffSet, 32) & CommentClip
   VarSearch = Trim$(MyStr)
  End If
 End If

Exit Function

BadError:
 CommentClip = vbNullString

End Function

Private Function ConvertAsciiSearch(ByVal strConv As String) As String

  'this routine adds the ability to search for Characters referred to by ascii value
  'ONLY works if pattern search is ON
  'other wise you are searching for the literal string
  
  Dim StartAsc As Long
  Dim EndAsc   As Long
  Dim AscVal   As Long

 StartAsc = 1
 Do While InStr(StartAsc, strConv, "\ASC(")
  StartAsc = InStr(StartAsc, strConv, "\ASC(")
  EndAsc = InStr(StartAsc, strConv, ")")
  If EndAsc > StartAsc Then
   AscVal = Val(Mid$(strConv, StartAsc + 5, EndAsc - 1))
   If AscVal > -1 And AscVal < 256 Then
    strConv = Left$(strConv, StartAsc - 1) & Chr$(AscVal) & Mid$(strConv, EndAsc + 1)
    Else
    StartAsc = EndAsc
    'this will jump you past an ivalid asc value and check for other "\ASC(" triggers
   End If
   Else
   Exit Do
   'this jumps out of Loop if the end bracket is missing
  End If
 Loop
 ConvertAsciiSearch = strConv

End Function

Public Sub DoFind(Fgrd As MSFlexGrid)

  Dim Proj           As VBProject
  Dim Comp           As VBComponent
  Dim Pane           As CodePane
  Dim StartText      As Long
  Dim EndText        As Long
  Dim StrProjName    As String
  Dim StrCompName    As String
  Dim strTarget      As String
  Dim strRoutine     As String
  Dim strFound       As String
  Dim startLine      As Long
  Dim startCol       As Long
  Dim EndLine        As Long
  Dim endCol         As Long
  Dim GotIT          As Boolean
  Dim ProcLineNo     As Long

 On Error Resume Next
ReTry:
 With Fgrd
  StrProjName = .TextMatrix(.Row, 0)
  StrCompName = .TextMatrix(.Row, 1)
  startLine = CLng(.TextMatrix(.Row, 2))
  strRoutine = .TextMatrix(.Row, 3)
  ProcLineNo = CLng(.TextMatrix(.Row, 4))
  strFound = .TextMatrix(.Row, 5)
 End With 'Fgrd
 strTarget = mobjDoc.ComboGetText(SearchB)
 'this is the fast Find used if no editing has been done
 For Each Proj In VBInstance.VBProjects
  If StrProjName = Proj.Name Then
   For Each Comp In Proj.VBComponents
    If Comp.Name = StrCompName Then
     If Comp.CodeModule.Find(strFound, startLine, 1, -1, -1) Then
      GotIT = True
      Exit For
     End If
    End If
   Next Comp
   If GotIT Then
    Exit For
   End If
  End If
 Next Proj
 'this will do Find if the line number has been changed by editing
 If Not GotIT Then
  startLine = 1
  For Each Proj In VBInstance.VBProjects
   If StrProjName = Proj.Name Then
    For Each Comp In Proj.VBComponents
     If Comp.Name = StrCompName Then
      startLine = 1
      If Comp.CodeModule.Find(strFound, startLine, 1, Comp.CodeModule.CountOfLines, -1) Then
       Do
        'this do loop takes care of the possibility of identical lines being present in different routines
        If strRoutine = GetProcName(Comp.CodeModule, startLine) Then
         GotIT = True
         'reset the Line data so fast Find will be used next time
         Fgrd.TextMatrix(Fgrd.Row, 2) = startLine
         Fgrd.TextMatrix(Fgrd.Row, 4) = GetProcLineNumber(Comp.CodeModule, startLine)
         Exit Do
        End If
        startLine = startLine + 1
       Loop While Comp.CodeModule.Find(strFound, startLine, 1, Comp.CodeModule.CountOfLines, -1)
      End If
     End If
     If GotIT Then
      Exit For
     End If
    Next Comp
    If GotIT Then
     Exit For
    End If
   End If
  Next Proj
 End If
 If GotIT Then
  If bFindSelectWholeLine Then
   'select the whole line
   StartText = InStr(1, Comp.CodeModule.Lines(startLine, 1), strFound, vbTextCompare)
   EndText = StartText + Len(strFound)
   Else
   'select the search word
   If Not BPatternSearch Then
    StartText = InStr(1, Comp.CodeModule.Lines(startLine, 1), strTarget, vbTextCompare)
    EndText = StartText + Len(strTarget)
    Else
    ' the string/comment filters cannot work on a PatternSearch StrFind
    'but you can get the actual string found and test that
    startCol = 1
    EndLine = -1
    endCol = -1
    If Comp.CodeModule.Find(strTarget, startLine, StartText, EndLine, EndText, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch) Then
    End If
   End If
  End If
  VisibleScroll_Init Pane, Comp.CodeModule
  If StartText > 0 Then
   ReportAction Fgrd, Found
   With Pane
    .SetSelection startLine, StartText, startLine, EndText
   End With
   Set Pane = Nothing
   Exit Sub
  End If
  Else
  'Line is missing (probably edited out) so re do whole search
  '(may be a pain if really large search was done),
  DoSearch Fgrd
 End If
 On Error GoTo 0

End Sub

Public Sub DoReplace(Fgrd As MSFlexGrid)

  Dim Msg                       As String
  Dim Code                      As String
  Dim CompMod                   As CodeModule
  Dim Comp                      As VBComponent
  Dim Proj                      As VBProject
  Dim strFind                   As String
  Dim strReplace                As String
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

 If bFilterWarning Then
  Msg = IIf(bWholeWordonly, "ON ", "OFF") & " Whole Word"
  Msg = Msg & vbNewLine & IIf(bCaseSensitive, "ON ", "OFF") & " Case Sensitive"
  Msg = Msg & vbNewLine & IIf(bNoComments, "ON ", "OFF") & " No Comments"
  Msg = Msg & vbNewLine & IIf(bCommentsOnly, "ON ", "OFF") & " Comments Only"
  Msg = Msg & vbNewLine & IIf(bNoStrings, "ON ", "OFF") & " No Strings"
  Msg = Msg & vbNewLine & IIf(bStringsOnly, "ON ", "OFF") & " Strings Only"
  If vbCancel = MsgBox(Msg & vbNewLine & vbNewLine & "Proceed with Replace anyway?", vbExclamation + vbOKCancel, "Filter Warning " & AppDetails) Then
   Exit Sub
  End If
 End If
 strFind = mobjDoc.ComboGetText(SearchB)
 strReplace = mobjDoc.ComboGetText(ReplaceB)
 If LenB(strFind) = 0 Then
  Exit Sub
 End If
 If bBlankWarning Then
  If LenB(strReplace) = 0 Then
   If vbCancel = MsgBox("Replace '" & strFind & "' with blank?", vbExclamation + vbOKCancel, "Blank Warning " & AppDetails) Then
    Exit Sub
   End If
  End If
 End If
 On Error Resume Next
 With mobjDoc
  .ShowWorking True, "Replacing..."
  .ComboBoxSave SearchB, , HistDeep
  .ComboBoxSave ReplaceB, , HistDeep
 End With 'mobjDoc
 If InStr(strReplace, "^N^") Then
  strReplace = Replace$(strReplace, "^N^", vbNewLine)
 End If
 If InStr(strReplace, "^T^") Then
  strReplace = Replace$(strReplace, "^T^", vbTab)
 End If
 DoEvents
 bCancel = False
 GetCounts
 CurProc = GetCurrentProcedure
 curModule = GetCurrentModule
 AutoSelectInitialize PrevCurCodePane, AutoRangeRevert
 If iRange = SelCode Then
  VBInstance.ActiveCodePane.GetSelection SelstartLine, SelStartCol, selendline, SelEndCol
 End If
 ReplaceCount = 0
 For Each Proj In VBInstance.VBProjects
  For Each Comp In Proj.VBComponents
   ReportAction Fgrd, replacing
   If SafeCompToProcess(Comp) Then
    If iRange > AllCode Then
     If Comp.Name <> curModule Then
      GoTo SkipComp
     End If
    End If
    Set CompMod = Comp.CodeModule
    startLine = 1
    If CompMod.Find(strFind, startLine, 1, CompMod.CountOfLines, -1, bWholeWordonly, bCaseSensitive, False) Then
     Do
      EndLine = -1
      startCol = 1
      endCol = -1
      If CompMod.Find(strFind, startLine, startCol, EndLine, endCol, bWholeWordonly, bCaseSensitive, False) Then
       Procname = CompMod.ProcOfLine(startLine, vbext_pk_Proc)
       If LenB(Procname) = 0 Then
        Procname = "(Declarations)"
       End If
       If iRange = ProcCode Then
        If Procname <> CurProc Then
         GoTo SkipProc
        End If
       End If
       Code$ = CompMod.Lines(startLine, 1)
       ApplyStringCommentFilters Code$, strFind
       ApplySelectedTextRestriction Code$, startLine, startCol, EndLine, endCol, SelstartLine, SelStartCol, selendline, SelEndCol
       If Len(Code$) Then
        If bWholeWordonly Then
         WholeWordReplacer Code$, strFind, strReplace
         Else
         Code$ = Replace$(Code$, strFind, strReplace, , , IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare))
        End If
        CompMod.ReplaceLine startLine, Code$
        ReplaceCount = ReplaceCount + 1
        mobjDoc.ShowWorking True, "Replacing...", "(" & ReplaceCount & ") Items"
       End If
SkipProc:
      End If
      Code$ = vbNullString
      startLine = startLine + 1
      If mobjDoc.CancelSearch Then
       Exit Do
      End If
      If startLine > CompMod.CountOfLines Then
       Exit Do
      End If
     Loop While CompMod.Find(strFind, startLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) 'StartLine > 0 And StartLine <= CompMod.CountOfLines
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
 Set Proj = Nothing
 Set CompMod = Nothing
 ReportAction Fgrd, replaced
 mobjDoc.ShowWorking False
 If bReplace2Search Then
  If Len(strReplace) Then
   mobjDoc.ComboBoxSave ReplaceB, strFind, HistDeep
   mobjDoc.ComboBoxSave SearchB, strReplace, HistDeep
   DoFind Fgrd
  End If
 End If
 On Error GoTo 0

End Sub

Public Sub DoSearch(Fgrd As MSFlexGrid)

  'ver1.1.02 major rewrite to allow simple pattern searching
  
  Dim Code                      As String
  Dim Procname                  As String
  Dim ProcLineNo                As Long
  Dim CompMod                   As CodeModule
  Dim Comp                      As VBComponent
  Dim Proj                      As VBProject
  Dim strFind                   As String
  Dim LongestPrj                As String
  Dim ResizeNeeded              As Boolean
  Dim strStrComTest             As String
  Dim startLine                 As Long
  Dim startCol                  As Long
  Dim EndLine                   As Long
  Dim endCol                    As Long
  Dim SecondRun                 As Boolean
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

 On Error Resume Next
 strFind = mobjDoc.ComboGetText(SearchB)
 'if strFind doesn't have any triggers and Pattern Search is on then turn it off
 AutoPatternOff strFind
 If BPatternSearch Then
  If InStr(strFind, "\ASC(") Then
   strFind = ConvertAsciiSearch(strFind)
   AddToSearchBox strFind
  End If
 End If
 If LenB(strFind) = 0 Then
  Exit Sub
 End If
 If strFind = " " Then
  MsgBox "Search for single spaces is cancelled, it overloads the system", vbInformation
  mobjDoc.ComboSetFocus SearchB
  Exit Sub
 End If
 DefaultGridSizes
 StartProcRow = 1
 mobjDoc.ShowWorking True, "Searching..."
ReTry:
 Fgrd.BackColorFixed = ColourHeadWork
 If LenB(strFind) > 0 Then
  LongestPrj$ = vbNullString
  bComplete = False
  ClearFGrid Fgrd
  mobjDoc.ComboBoxSave SearchB, , HistDeep
  mobjDoc.CancelButton True
  DoEvents
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
   If Len(Proj.Name) > Len(GridSizer(0)) Then
    GridSizer(0) = Proj.Name
    ResizeNeeded = True
   End If
   For Each Comp In Proj.VBComponents
    If Fgrd.Rows < LongLimit Then
     If SafeCompToProcess(Comp) Then
      If iRange > AllCode Then
       If Comp.Name <> curModule Then
        GoTo SkipComp
       End If
      End If
      If Len(Comp.Name) > Len(GridSizer(1)) Then
       GridSizer(1) = Comp.Name
       ResizeNeeded = True
      End If
      With Comp
       Set CompMod = .CodeModule
       '5hould I quit?
       ReportAction Fgrd, Search
       If LenB(.Name) = 0 Then
        bCancel = True
        bComplete = True
       End If
      End With
      If mobjDoc.CancelSearch Then
       Exit For
      End If
      'Safety turns off filters if comment/double quote is actually in the search phrase
      If bNoComments Then
       If InStr(strFind, Apostrophe) > 0 Then
        bNoComments = True
        mobjDoc.SetFilterButtons
       End If
      End If
      If bNoStrings Then
       If InStr(strFind, Chr$(34)) > 0 Then
        bNoStrings = False
        mobjDoc.SetFilterButtons
       End If
      End If
      startLine = 1 'initialize search range
      Do
       startCol = 1
       EndLine = -1
       endCol = -1
       If CompMod.Find(strFind, startLine, startCol, EndLine, endCol, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch) Then
        DoEvents
        If mobjDoc.CancelSearch Then
         Exit For
        End If
        Procname = GetProcName(CompMod, startLine)
        If iRange = ProcCode Then
         If Procname <> CurProc Then
          GoTo SkipProc
         End If
        End If
        ProcLineNo = GetProcLineNumber(CompMod, startLine)
        If Len(CStr(startLine)) > Len(GridSizer(2)) Then
         GridSizer(2) = CStr(startLine)
         ResizeNeeded = True
        End If
        If Len(Procname) > Len(GridSizer(3)) Then
         GridSizer(3) = Procname
         ResizeNeeded = True
        End If
        If Len(CStr(ProcLineNo)) > Len(GridSizer(4)) Then
         GridSizer(4) = CStr(ProcLineNo)
         ResizeNeeded = True
        End If
        Code$ = CompMod.Lines(startLine, 1)
        'apply nostring no comment filters
        If BPatternSearch Then
         ' the string/comment filters cannot work on a PatternSearch StrFind
         'but you can get the actual string found and test that
         CompMod.CodePane.SetSelection startLine, startCol, EndLine, endCol
         strStrComTest = Mid$(Code$, startCol, endCol - startCol)
         Else
         strStrComTest = strFind
        End If
        ApplyStringCommentFilters Code$, strStrComTest
        ApplySelectedTextRestriction Code$, startLine, startCol, EndLine, endCol, SelstartLine, SelStartCol, SeEndLine, SelEndCol
        If LenB(Code$) Then
         If ResizeNeeded Then
          'slight speed advantage of not doing this unless called for
          mobjDoc.GridReSize
          ResizeNeeded = False
         End If
         With Fgrd
          .Rows = .Rows + 1
          .TextMatrix(.Row, 0) = Proj.Name
          .TextMatrix(.Row, 1) = Comp.Name
          .TextMatrix(.Row, 2) = startLine
          .TextMatrix(.Row, 3) = Procname
          .TextMatrix(.Row, 4) = ProcLineNo
          .TextMatrix(.Row, 5) = Code$
          .Row = .Row + 1
          Code$ = vbNullString
         End With 'Fgrd
        End If
SkipProc:
       End If
       startLine = startLine + 1
       If mobjDoc.CancelSearch Then
        Exit Do
       End If
       If startLine >= CompMod.CountOfLines Then
        Exit Do
       End If
      Loop While CompMod.Find(strFind, startLine, 1, -1, -1, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch)
      ' End If
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
  Set Proj = Nothing
  Set CompMod = Nothing
  mobjDoc.CancelButton False
 End If
 With Fgrd
  .Rows = .Rows - 1
  If .Row = 0 Then
   'nothing found
   .BackColorSel = ColourHeadNoFind
   .ForeColorSel = ColourHeadFore
   Else
   .BackColorSel = ColourFindSelectBack
   .ForeColorSel = ColourFindSelectFore
  End If
 End With 'Fgrd
 'automatically switch to pattern search if ordinary fails
 If Fgrd.Row = 0 Then
  If Not BPatternSearch Then
   If instrAny(strFind, "*", "!", "[", "]", "\") Then
    BPatternSearch = Not BPatternSearch
    mobjDoc.ClearForPattern
    SecondRun = True
    GoTo ReTry
   End If
  End If
  If SecondRun Then
   'autoPattern Search is on
   If Fgrd.Row = 0 Then
    'turn it off it still no finds
    BPatternSearch = Not BPatternSearch
    mobjDoc.ClearForPattern
   End If
  End If
 End If
 If mobjDoc.CancelSearch Then
  ReportAction Fgrd, IIf(bComplete, Found, inComplete)
 End If
 Fgrd.Refresh
 If Fgrd.Rows > 1 Then
  ReportAction Fgrd, IIf(bComplete, Found, Complete)
  SetFocus_Safe Fgrd
  Else
  ReportAction Fgrd, IIf(bComplete, Found, Missing)
 End If
 'this turns off auto Selected text only
 If AutoRangeRevert Then
  iRange = PrevCurCodePane
  mobjDoc.ToggleButtonFaces
 End If
 mobjDoc.ShowWorking False
 'this turns auto switch to pattern search off if it was used
 If Fgrd.Rows = LongLimit Then
  ' as this is 2147483647 rows it is unlikely that this will ever hit but just in case :)
  MsgBox "Search halted because number of finds reached limit of Find ComboBox", vbCritical
 End If
 mobjDoc.GridReSize
 On Error GoTo 0

End Sub

Public Function ExpandCode(Code As String) As Boolean

 If bExpandIfThen Then
  ExpandCode = IfThenStructureExpander(Code)
 End If
 If Not ExpandCode Then
  If bExpandColon Then
   ExpandCode = ColonExpander(Code)
  End If
 End If

End Function

Private Function FilteredInStr(StrSearch As String, _
                               strFind As String) As Long

  'ver 1.1.02
  'improved word detection makes sure that the DoFind routine
  'highlights the correct (or at least first instance)
  'in string that matches all filters
  
  Dim LBit As String
  Dim Rbit As String

 Do
  FilteredInStr = InStr(FilteredInStr + 1, StrSearch, strFind, IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare))
  If StrSearch = strFind Then
   Exit Do
  End If
  If bWholeWordonly Then
   Select Case FilteredInStr
    Case 1
    Rbit$ = Mid$(StrSearch, Len(strFind) + 1, 1)
    If IsPunct(Rbit) Then
     Exit Do
    End If
    Case Len(StrSearch) - Len(strFind) + 1
    LBit$ = Mid$(StrSearch, Len(StrSearch) - Len(strFind), 1)
    If IsPunct(LBit) Then
     Exit Function
    End If
    Case 0
    Exit Do 'not found do nothing
    Case Else
    LBit$ = Mid$(StrSearch, FilteredInStr - 1, 1)
    Rbit$ = Mid$(StrSearch, FilteredInStr + Len(strFind), 1)
    If IsPunct(LBit) And IsPunct(Rbit) Then
     Exit Do
    End If
   End Select
   Else
   Exit Do
  End If
 Loop

End Function

Public Function GetCurrentModule() As String

 GetCurrentModule = VBInstance.ActiveCodePane.CodeModule.Name

End Function

Public Function GetCurrentProcedure() As String

  Dim sl    As Long
  Dim sc    As Long
  Dim el    As Long
  Dim ec    As Long
  Dim lJunk As Long

 VBInstance.ActiveCodePane.GetSelection sl, sc, el, ec
 GetLineData VBInstance.ActiveCodePane.CodeModule, sl, GetCurrentProcedure, lJunk, lJunk, lJunk, lJunk
 ' GetCurrentProcedure = VBInstance.ActiveCodePane.CodeModule.ProcOfLine(sl, vbext_pk_Proc)
 ' If Len(GetCurrentProcedure) = 0 Then
 '  GetCurrentProcedure = "(Declarations)"
 ' End If

End Function

Public Sub GetEndOfDimLines(ByVal CompMod As CodeModule, _
                            ByVal codeline As Long, _
                            EndOfDimLines As Long)

  Dim Pname           As String
  Dim PlineNo         As Long
  Dim PStartLine      As Long
  Dim PEndLine        As Long
  Dim lJunk           As Long
  Dim prevHadLineCont As Boolean
  Dim dimLineReached  As Long
  Dim strTmp          As String

 GetLineData CompMod, codeline, Pname, PlineNo, PStartLine, PEndLine, lJunk
 EndOfDimLines = PStartLine + 1
 strTmp = Trim$(CompMod.Lines(EndOfDimLines, 1))
 Do While HasLineCont(strTmp) Or IsDimLine(strTmp) Or (dimLineReached = False And JustACommentOrBlank(strTmp)) Or prevHadLineCont
  EndOfDimLines = EndOfDimLines + 1
  prevHadLineCont = HasLineCont(strTmp)
  'if this line has linecont then the next line must be counted if it does not have a line cont
  'as it is end of line cont code line
  If Not dimLineReached Then
   'once a dim line is reached then blanks are no longer avoided
   dimLineReached = IsDimLine(strTmp)
  End If
  ' this takes care of last line of line cont code
  If EndOfDimLines = PEndLine Then
   EndOfDimLines = codeline ' this is a safe junk value
   Exit Do 'safety should never hit
  End If
  strTmp = Trim$(CompMod.Lines(EndOfDimLines, 1))
 Loop

End Sub

Public Sub GetLineData(CmpMod As CodeModule, _
                       ByVal cdeline As Long, _
                       Procname As String, _
                       ProcLineNo As Long, _
                       PRocStartLine As Long, _
                       ProcEndLine As Long, _
                       ProcHeadLine As Long)

  Dim I         As Long
  Dim K         As Long
  Dim CleanElem As Variant

 For I = 1 To 4
  K = Choose(I, vbext_pk_Proc, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set)
  CleanElem = Null
  On Error Resume Next
  'IF you crash here first check that Error Trapping is not ON
  CleanElem = CmpMod.ProcOfLine(cdeline, K)
  On Error GoTo 0
  If Not IsNull(CleanElem) Then
   Procname = CleanElem
   If Len(Procname) = 0 Then
    Procname = "(Declarations)"
    ProcLineNo = cdeline
    PRocStartLine = 1
    ProcEndLine = CmpMod.CountOfDeclarationLines
    Else
    ProcLineNo = CmpMod.ProcBodyLine(Procname, K)
    PRocStartLine = CmpMod.PRocStartLine(Procname, K)
    ProcHeadLine = PRocStartLine
    Do Until Not JustACommentOrBlank(CmpMod.Lines(ProcHeadLine, 1))
     ProcHeadLine = ProcHeadLine + 1
    Loop
    ProcEndLine = PRocStartLine + CmpMod.ProcCountLines(Procname, K)
   End If
   Exit For
  End If
 Next I

End Sub

Public Function GetLineProcedure(CmpMod As CodeModule, _
                                 ByVal CurLine As Long) As String

  Dim I         As Long
  Dim K         As Long
  Dim CleanElem As Variant

 For I = 1 To 4
  K = Choose(I, vbext_pk_Proc, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set)
  CleanElem = Null
  On Error Resume Next
  'IF you crash here first check that Error Trapping is not ON
  CleanElem = CmpMod.ProcOfLine(CurLine, K)
  On Error GoTo 0
  If Not IsNull(CleanElem) Then
   GetLineProcedure = CleanElem
   If Len(GetLineProcedure) = 0 Then
    GetLineProcedure = "(Declarations)"
   End If
   Exit For
  End If
 Next I

End Function

Public Function GetProcLineNumber(CmpMod As CodeModule, _
                                  CodeLineNo As Long) As String

  Dim LProcName As String
  Dim I         As Long
  Dim K         As Long
  Dim CleanElem As Variant

 LProcName = GetProcName(CmpMod, CodeLineNo)
 If LProcName = "(Declarations)" Then
  GetProcLineNumber = CodeLineNo
  Else
  'The + 1 is because ProcBodyLine returns a 0 based count but most people like 1 based counts
  'Oddly CodeLineNo which is generated by VB's Find is 1 based
  For I = 1 To 4
   K = Choose(I, vbext_pk_Proc, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set)
   CleanElem = Null
   On Error Resume Next
   'IF you crash here first check that Error Trapping is not ON
   CleanElem = CodeLineNo - CmpMod.ProcBodyLine(LProcName, K) + 1
   On Error GoTo 0
   If Not IsNull(CleanElem) Then
    GetProcLineNumber = CleanElem
    Exit For
   End If
  Next I
 End If

End Function

Public Function GetProcName(CmpMod As CodeModule, _
                            ByVal CodeLineNo As Long) As String

  ' GetProcName = CmpMod.ProcOfLine(CodeLineNo, vbext_pk_Proc)
  ' If LenB(GetProcName) = 0 Then
  '  GetProcName = "(Declarations)"
  ' End If
  
  Dim I         As Long
  Dim K         As Long
  Dim CleanElem As Variant

 For I = 1 To 4
  K = Choose(I, vbext_pk_Proc, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set)
  CleanElem = Null
  On Error Resume Next
  'IF you crash here first check that Error Trapping is not ON
  CleanElem = CmpMod.ProcOfLine(CodeLineNo, K)
  On Error GoTo 0
  If Not IsNull(CleanElem) Then
   GetProcName = CleanElem
   If Len(GetProcName) = 0 Then
    GetProcName = "(Declarations)"
   End If
   Exit For
  End If
 Next I

End Function

Public Function InComment(ByVal Code As String, _
                          ByVal Codepos As Long) As Boolean

  Dim SQuotePos As Long

 On Error Resume Next
 SQuotePos = InStr(Code$, Apostrophe)
 If SQuotePos = 1 Then
  InComment = True
 End If
 If SQuotePos > 1 Then
  If Codepos > SQuotePos Then
   InComment = True
  End If
 End If
 On Error GoTo 0

End Function

Public Function InDeclaration(CmpMod As CodeModule, _
                              cdeline As Long) As Boolean

 InDeclaration = CmpMod.ProcOfLine(cdeline, vbext_pk_Proc) = vbNullString

End Function

Public Function InitiateSelRange(CmpMod As CodeModule, _
                                 ByVal CurProc As String, _
                                 ByVal SelstartLine As Long) As Long

 Select Case iRange
  Case AllCode, ModCode
  InitiateSelRange = 1
  Case ProcCode
  Do
   InitiateSelRange = InitiateSelRange + 1
  Loop Until (CurProc = GetProcName(CmpMod, InitiateSelRange))
  Case SelCode
  InitiateSelRange = SelstartLine
 End Select

End Function

Public Function InQuotes(ByVal Code As String, _
                         ByVal Codepos As Long) As Boolean

  Dim LQ As Long
  Dim FQ As Long

 On Error Resume Next
 LQ = InStr(StrReverse(Code$), Chr$(34))
 If LQ > 0 Then
  LQ = Len(Code$) - LQ + 1
 End If
 FQ = InStr(Code$, Chr$(34))
 If LQ = 0 Then
  If FQ = 0 Then
   Exit Function
  End If
 End If
 If LQ = FQ Then
  Exit Function
 End If
 If FQ < Codepos Then
  If Codepos < LQ Then
   InQuotes = True
  End If
 End If
 On Error GoTo 0

End Function

Public Function InSRange(LNum As Long, _
                         CmpMod As CodeModule, _
                         CurProc As String, _
                         selendline As Long) As Boolean

 Select Case iRange
  Case AllCode, ModCode
  InSRange = (LNum <= CmpMod.CountOfLines)
  Case ProcCode
  InSRange = (CurProc = GetProcName(CmpMod, LNum))
  Case SelCode
  InSRange = (LNum <= selendline)
 End Select

End Function

Public Function IsAlphaIntl(ByVal sChar As String) As Boolean

 IsAlphaIntl = Not (UCase$(sChar) = LCase$(sChar))

End Function

Public Function IsNumeral(ByVal strTest As String) As Boolean

 IsNumeral = InStr("1234567890", strTest) > 0

End Function

Public Function IsPunct(ByVal strTest As String) As Boolean

  'Detect punctuation

 If IsNumeral(strTest) Then
  IsPunct = False
  Else
  IsPunct = Not IsAlphaIntl(strTest)
 End If

End Function

Public Function KeepBetweenLng(ByVal MinV As Long, _
                               ByVal Val As Long, _
                               ByVal MaxV As Long) As Long

 If Val < MinV Then
  Val = MinV
  ElseIf Val > MaxV Then
  Val = MaxV
 End If
 KeepBetweenLng = Val

End Function

Public Function LastCodeWord(ByVal varChop As Variant) As String

  Dim TmpA As Variant
  Dim junk As String

 junk = CommentClip(varChop)
 If LenB(varChop) Then
  TmpA = Split(varChop)
  LastCodeWord = TmpA(UBound(TmpA))
 End If

End Function

Public Function LastWord(ByVal varChop As Variant) As String

  Dim TmpA As Variant

 If LenB(varChop) Then
  TmpA = Split(varChop)
  LastWord = TmpA(UBound(TmpA))
 End If

End Function

Public Function LeftWord(ByVal varChop As Variant) As String

 If LenB(varChop) Then
  LeftWord = Split(varChop)(0)
 End If

End Function

Public Sub ReportAction(Fgrd As MSFlexGrid, _
                        ByVal Act As EnumMsg, _
                        Optional ByVal AppendStr As String)

  Dim Msg                As String
  Dim StrItems           As String
  Dim StrFilterWarning   As String
  Dim strSearchEndStatus As String
  Dim StrPatternWarning  As String
  Dim strStatusMsg       As String
  Dim strStatusActMsg    As String

 StrItems = "(" & Fgrd.Rows - 1 & ") Item" & IIf(Fgrd.Rows - 1 <> 1, "s", vbNullString)
 StrFilterWarning = IIf(mobjDoc.AnyFilterOn, " <Filter>", vbNullString)
 StrPatternWarning = IIf(BPatternSearch, " <Pattern>", vbNullString)
 Select Case Act
  Case Search
  strSearchEndStatus = " Searching " & IIf(Len(AppendStr), " in " & AppendStr, "...")
  strStatusMsg = "Searching..."
  strStatusActMsg = StrItems & " Found" & StrFilterWarning & StrPatternWarning
  Fgrd.BackColorFixed = ColourHeadWork
  Case Complete
  strSearchEndStatus = " Search Complete."
  Fgrd.BackColorFixed = IIf(BPatternSearch, ColourHeadPattern, ColourHeadDefault)
  Case inComplete
  strSearchEndStatus = " Search Cancelled."
  Fgrd.BackColorFixed = IIf(BPatternSearch, ColourHeadPattern, ColourHeadDefault)
 End Select
 Select Case Act
  Case replacing
  strStatusMsg = "Replacing..."
  strStatusActMsg = vbNullString
  Msg = "Replacing.." & String$(Int(Rnd * 5 + 1), ".")
  With Fgrd
   .BackColorFixed = ColourHeadReplace
   If .Row = 0 Then
    'nothing found
    .BackColorSel = ColourHeadReplace
    .ForeColorSel = ColourHeadFore
    Else
    .BackColorSel = ColourFindSelectBack
    .ForeColorSel = ColourFindSelectFore
   End If
  End With 'Fgrd
  Case replaced
  Msg = "(" & ReplaceCount & ") Item" & IIf(ReplaceCount <> 1, "s", vbNullString) & " replaced"
  Fgrd.BackColorFixed = IIf(BPatternSearch, ColourHeadPattern, ColourHeadDefault)
  Case Missing
  Msg = StrItems & " Found" & StrFilterWarning & strSearchEndStatus & StrPatternWarning
  Fgrd.BackColorFixed = ColourHeadNoFind
  Case Else
  Msg = StrItems & " Found" & StrFilterWarning & strSearchEndStatus & StrPatternWarning
 End Select
 Select Case Act
  Case Search, replacing
  mobjDoc.ShowWorking True, strStatusMsg, strStatusActMsg
  Fgrd.TextMatrix(0, 5) = "Code: " & Msg
  Case Complete, inComplete, replaced, Missing
  Fgrd.TextMatrix(0, 5) = "Code: " & Msg
 End Select
 Fgrd.Refresh
 DoEvents

End Sub

Public Function SecondWord(ByVal varChop As Variant) As String

 If LenB(varChop) Then
  varChop = StripDoubleBlanks(varChop)
  If UBound(Split(varChop)) > 0 Then
   SecondWord = Split(varChop)(1)
  End If
 End If

End Function

Public Sub SelectedText(cmb As ComboBox, _
                        Cmd As CommandButton)

  Dim HiLitSelection As String

 HiLitSelection$ = GetSelectedText(VBInstance)
 If LenB(HiLitSelection$) Then
  If InStr(HiLitSelection$, vbNewLine) Then
   If (HiLitSelection$ <> vbNewLine) Then
    HiLitSelection$ = Left$(HiLitSelection$, InStr(HiLitSelection$, vbNewLine) - 1)
   End If
  End If
  If LenB(HiLitSelection$) Then
   cmb.SetFocus
   cmb.Text = HiLitSelection$
   Cmd = True
  End If
 End If

End Sub

Public Function StripDoubleBlanks(varChop As Variant) As String

 StripDoubleBlanks = varChop
 Do While InStr(StripDoubleBlanks, "  ")
  StripDoubleBlanks = Replace$(StripDoubleBlanks, "  ", " ")
 Loop

End Function

Public Sub VisibleScroll_Do(Pne As CodePane, _
                            ByVal cdeline As Long)

  Dim CurTopLine As Long

 With Pne
  CurTopLine = Abs(Int(.CountOfVisibleLines / 2) - cdeline) + 1
  If cdeline > CurTopLine Then
   .TopLine = CurTopLine
  End If
 End With 'Pane

End Sub

Public Sub VisibleScroll_Init(Pne As CodePane, _
                              CmpMod As CodeModule)

 Set Pne = CmpMod.CodePane
 With Pne
  'when docked only the first instance selected in GrdFound got highlighted
  'until I added next line, no idea why it works.
  .Window.Visible = False
  .Show
  .Window.SetFocus
 End With

End Sub

Private Sub WholeWordReplacer(cde As String, _
                              ByVal strF As String, _
                              ByVal StrRep As String)

  Dim FPos As Long
  Dim Cmpr As VbCompareMethod

 If Len(cde) Then
  If Len(strF) Then
   Cmpr = IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)
   FPos = InStr(1, cde, strF, Cmpr)
   Do While InStr(FPos, cde, strF, Cmpr)
    If IsPunct(Mid$(cde, FPos - 1, 1)) Or FPos = 1 Then
     If IsPunct(Mid$(cde, FPos + Len(strF), 1)) Or FPos = Len(cde) - Len(strF) + 1 Then
      cde = Left$(cde, FPos - 1) & StrRep & Mid$(cde, FPos + Len(strF))
     End If
    End If
    FPos = InStr(FPos + 1, cde, strF, Cmpr)
   Loop
  End If
 End If

End Sub

Public Function WordMember(ByVal varChop As Variant, _
                           WordNum As Long) As String

 If LenB(varChop) Then
  If UBound(Split(varChop)) >= WordNum - 1 Then
   WordMember = Split(varChop)(WordNum - 1)
  End If
 End If

End Function

':) Roja's VB Code Fixer V1.1.4 (30/07/2003 1:24:30 PM) 60 + 1293 = 1353 Lines Thanks Ulli for inspiration and lots of code.

