Attribute VB_Name = "support"
Option Explicit
Public mobjDoc                    As docfind
Public VBInstance                 As VBE
Public ProjCount                  As Long
Public bLaunchOnStart             As Boolean
Public bSaveHistory               As Boolean
Public bBlankWarning              As Boolean
Public bFilterWarning             As Boolean
Public bReplace2Search            As Boolean
Public bRemFilters                As Boolean
Public HistDeep                   As Long
Public bLoadingSettings           As Boolean
Public bAutoSelectText            As Boolean
'indent triggers
Public bVisibleIndenting          As Boolean
Public bDeleteDoubleBlanks        As Boolean
Public bDeleteAllBlanks           As Boolean
Public bAddStructureSpace         As Boolean
'fix triggers
Public bSortModules               As Boolean
Public bProcDecl2Top              As Boolean
Public bDeclSingleTypeExpand      As Boolean
Public bDeclExpand                As Boolean
Public bDeclAsFormat              As Boolean
Public bTypeSuffixUpdate          As Boolean
Public bExpandColon               As Boolean
Public bExpandIfThen              As Boolean
Public bStrConcatenateUpdate      As Boolean
Public bStrFunctionUpdate         As Boolean
Public bCommentOutUnused          As Boolean
Public bTestScope                 As Boolean
Public bPleonasmFix As Boolean
Public bChr2ConstFix As Boolean
Public bEnumCapProtect As Boolean
'fix comment triggers
Public bShowPrevCode              As Boolean
Public bShowFixComment            As Boolean
Public ColourTextFore             As Long
Public ColourTextBack             As Long
Public ColourFindSelectBack       As Long
Public ColourFindSelectFore       As Long
Public ColourHeadFore             As Long
Public ColourHeadDefault          As Long
Public ColourHeadWork             As Long
Public ColourHeadPattern          As Long
Public ColourHeadNoFind           As Long
Public ColourHeadReplace          As Long
Public GridSizer(4)               As String
Public Enum IndentEnum
 IndentOnly
 IndentFormat
 IndentFormatFixProp
 IndentFormatFixFull
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private IndentOnly, IndentFormat, IndentFormatFixProp, IndentFormatFixFull
#End If
Public Const offset               As Long = 120

Public Sub AddToSearchBox(ByVal strCom As String, _
                          Optional ByVal GenericCom As Boolean = False)

  'this adds any comments to the Search Combo
  'and the general comment marker if necessary

 mobjDoc.ComboSetText SearchB, strCom
 If GenericCom Then
  mobjDoc.ComboSetText SearchB, RGSignature
 End If

End Sub

Public Function AppDetails() As String

 With App
  AppDetails = .ProductName & " V" & .Major & "." & .Minor & "." & .Revision
 End With

End Function

Public Function Bool2Int(bValue As Boolean) As Integer

  'used to simplify settings

 Bool2Int = IIf(bValue, 1, 0)

End Function

Public Function Bool2Str(bValue As Boolean) As String

  'used to simplify settings

 Bool2Str = IIf(bValue, "1", "0")

End Function

Public Function CountSubString(VarSearch As Variant, _
                               varFind As Variant) As Long

  Dim TmpA As Variant

 TmpA = Split(VarSearch, varFind)
 CountSubString = UBound(TmpA)

End Function

Public Sub DefaultGridSizes()

 GridSizer(0) = "Project"
 GridSizer(1) = "Component"
 GridSizer(2) = "Line"
 GridSizer(3) = "Procedure"
 GridSizer(4) = "Line"
 mobjDoc.GridReSize

End Sub

Public Sub GetCounts()

  'the counts are used to control column visiblity
  
  Dim Proj      As VBProject

 On Error Resume Next
 ProjCount = VBInstance.VBProjects.Count
 'CompCount includes ProjCount just in case a group includes projects with only one component
 For Each Proj In VBInstance.VBProjects
  CompCount = CompCount + Proj.VBComponents.Count + IIf(ProjCount > 1, 1, 0)
 Next Proj
 On Error GoTo 0

End Sub

Public Function GetSelectedText(VBInstance As VBIDE.VBE) As String

  Dim startLine As Long
  Dim cmo       As VBIDE.CodeModule
  Dim codeText  As String
  Dim cpa       As VBIDE.CodePane
  Dim endCol    As Long
  Dim EndLine   As Long
  Dim startCol  As Long

 'Date: 4/27/1999
 'Versions: VB5 VB6 Level: Intermediate
 'Author: The VB2TheMax Team
 ' Return the string of code the is selected in the code window
 ' that is currently active.
 ' This function can only be used inside an add-in.
 On Error Resume Next
 ' get a reference to the active code window and the underlying module
 ' exit if no one is available
 Set cpa = VBInstance.ActiveCodePane
 Set cmo = cpa.CodeModule
 If Err.Number Then
  Exit Function
 End If
 ' get the current selection coordinates
 cpa.GetSelection startLine, startCol, EndLine, endCol
 ' exit if no text is highlighted
 If startLine = EndLine Then
  If startCol = endCol Then
   Exit Function
  End If
 End If
 ' get the code text
 If startLine = EndLine Then
  ' only one line is partially or fully highlighted
  codeText = Mid$(cmo.Lines(startLine, 1), startCol, endCol - startCol)
  Else
  ' the selection spans multiple lines of code
  ' first, get the selection of the first line
  codeText = Mid$(cmo.Lines(startLine, 1), startCol) & vbNewLine
  ' then get the lines in the middle, that are fully highlighted
  If startLine + 1 < EndLine Then
   codeText = codeText & cmo.Lines(startLine + 1, EndLine - startLine - 1)
  End If
  ' finally, get the highlighted portion of the last line
  codeText = codeText & Left$(cmo.Lines(EndLine, 1), endCol - 1)
 End If
 GetSelectedText = codeText
 On Error GoTo 0

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

Public Function InLiteral(ByVal VarSearch As Variant, _
                          ByVal Tpos As Long, _
                          Optional CommentTest As Boolean = True) As Boolean

  Dim Possible As Long
  Dim ArrTest  As Variant
  Dim I        As Long
  Dim OPos     As Long
  Dim NPos     As Long

 Possible = InStr(VarSearch, Chr$(34))
 If Possible Then
  If Possible = Tpos Then
   InLiteral = Not InComment(VarSearch, Tpos)
   Exit Function
  End If
  ArrTest = Split(VarSearch, Chr$(34))
  For I = LBound(ArrTest) To UBound(ArrTest)
   NPos = NPos + 1 + Len(ArrTest(I))
   If NPos > Tpos Then
    If IsOdd(I) Then
     If BetweenLng(OPos, Tpos, NPos) Then
      If CommentTest Then
       InLiteral = Not InComment(VarSearch, Tpos)
       Else
       ' this is only to stop nocomment creating recursive overflow
       InLiteral = True
      End If
      Exit For
     End If
    End If
   End If
   OPos = NPos
   If OPos > Tpos Then
    Exit For
   End If
  Next I
 End If

End Function

Public Function instrAny(ByVal StrSearch As String, _
                         ParamArray finds() As Variant) As Long

  Dim fnd    As Variant
  Dim TMpAny As Long

 instrAny = LongLimit
 For Each fnd In finds
  TMpAny = InStr(StrSearch, fnd)
  If TMpAny < instrAny Then
   If TMpAny > 0 Then
    instrAny = TMpAny
   End If
  End If
 Next fnd
 If instrAny = LongLimit Then
  instrAny = 0
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

Public Function IsInArray(ByVal FindValue As Variant, _
                          ByVal arrSearch As Variant) As Boolean

  'VERY FAST function to find a value in an Array
  'By Brian Gillham
  'Versions:  VB5, VB6
  'Posted:  4/18/2001

 On Error GoTo LocalError
 If Not IsArray(arrSearch) Then
  Exit Function
 End If
 'empty array
 If UBound(arrSearch) = -1 Then
  Exit Function
 End If
 IsInArray = InStr(1, vbNullChar & Join(arrSearch, vbNullChar) & vbNullChar, vbNullChar & FindValue & vbNullChar) > 0

Exit Function

LocalError:
 'Justin (just in case)

End Function

Public Function IsOdd(ByVal N As Variant) As Boolean

  'Here's a efficient IsEven function
  'By Sam Hills
  'shills@bbll.com
  'If you want an IsOdd function, just omit the Not.
  '        IsEven =not -(n And 1)

 IsOdd = -(N And 1)

End Function

Public Function MultiLeft(ByVal VarSearch As Variant, _
                          ByVal CaseSensitive As Boolean, _
                          ParamArray Afind() As Variant) As Boolean

  Dim FindIt As Variant

 'This routine was originally designed to test multiple possible left strings
 'BUT I also use it as a simple way of testing even a single left string
 'without having to separately code the length at every instance
 If Not CaseSensitive Then
  VarSearch = LCase$(VarSearch)
 End If
 For Each FindIt In Afind
  If CaseSensitive Then
   If Left$(VarSearch, Len(FindIt)) = FindIt Then
    MultiLeft = True
    Exit Function
   End If
   Else
   If Left$(VarSearch, Len(FindIt)) = LCase$(FindIt) Then
    MultiLeft = True
    Exit Function
   End If
  End If
 Next FindIt

End Function

Public Function MultiRight(ByVal VarSearch As Variant, _
                           ByVal CaseSensitive As Boolean, _
                           ParamArray Afind() As Variant) As Boolean

  Dim FindIt As Variant

 'This routine was originally designed to test multiple possible left strings
 'BUT I also use it as a simple way of testing even a single left string
 'without having to separately code the length at every instance
 'CaseSensitive was added to solve a problem with hand coding of standard VB routines with wrong case
 If Not CaseSensitive Then
  VarSearch = LCase$(VarSearch)
 End If
 For Each FindIt In Afind
  If CaseSensitive Then
   If Right$(VarSearch, Len(FindIt)) = FindIt Then
    MultiRight = True
    Exit Function
   End If
   Else
   If Right$(VarSearch, Len(FindIt)) = LCase$(FindIt) Then
    MultiRight = True
    Exit Function
   End If
  End If
 Next FindIt

End Function

Public Function SafeCompToProcess(ByVal cmp As VBComponent) As Boolean

  'returns True if the component is anything that can/should be processed by program
  'test that the component is one you can edit at all

 SafeCompToProcess = cmp.Type <> vbext_ct_ResFile And cmp.Type <> vbext_ct_RelatedDocument
 ' this routine is called at start of all rewrite code so
 ' this is a good spot to make sure that suspend is off
 SuspendCF = False
 'the counter has to increase whether or not it is true
 'If the call uses 'Dummy' to call function then the number
 'is not needed for that routine it is just thrown away
 'only test these if first test is pased

End Function

Public Sub SetFocus_Safe(CTL As Control)

  '*PURPOSE: protect SetFocus from any of the many conditions which can stuff it

 On Error Resume Next
 With CTL
  If .Visible Then
   If .Enabled Then
    .SetFocus
   End If
  End If
 End With
 On Error GoTo 0

End Sub

':) Roja's VB Code Fixer V1.1.4 (30/07/2003 1:24:21 PM) 55 + 377 = 432 Lines Thanks Ulli for inspiration and lots of code.

