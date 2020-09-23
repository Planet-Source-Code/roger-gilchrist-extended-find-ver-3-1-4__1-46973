Attribute VB_Name = "SillybitsMod"
Option Explicit
Public Enum CaseConvert
 vbUpperCase = 1
 vbLowerCase = 2
 vbProperCase = 3
 SimpleSentence = 4
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private vbUpperCase, vbLowerCase, vbProperCase, SimpleSentence
#End If
Private CurCase     As CaseConvert

Public Sub ConvertSelectedText(VBInstance As VBIDE.VBE, _
                               Optional Conversion As CaseConvert)

  'ConvertSelectedTex - Convert text selected in code window
  'Date: 7/31/1999
  'Versions: VB5 VB6 Level: Advanced
  'Author: The VB2TheMax Team
  ' Convert to uppercase, lowercase, or propercase the text that is
  ' currently selected in the active code window
  
  Dim startLine As Long
  Dim startCol  As Long
  Dim EndLine   As Long
  Dim endCol    As Long
  Dim codeText  As String
  Dim cpa       As VBIDE.CodePane
  Dim cmo       As VBIDE.CodeModule
  Dim I         As Long

 On Error Resume Next
 ' get a reference to the active code window and the underlying module
 ' exit if no one is available
 Set cpa = VBInstance.ActiveCodePane
 Set cmo = cpa.CodeModule
 If Err Then
  Exit Sub
 End If
 ' get the current selection coordinates
 cpa.GetSelection startLine, startCol, EndLine, endCol
 ' exit if no text is highlighted
 If startLine = EndLine Then
  If startCol = endCol Then
   Exit Sub
  End If
 End If
 ' get the code text
 If startLine = EndLine Then
  ' only one line is partially or fully highlighted
  codeText = cmo.Lines(startLine, 1)
  If Conversion = SimpleSentence Then
   SimpleSentenceCase codeText
   Else
   Mid$(codeText, startCol, endCol - startCol) = strConv(Mid$(codeText, startCol, endCol - startCol), Conversion)
   If Conversion = vbProperCase Then
    ProperProperCase codeText
   End If
  End If
  cmo.ReplaceLine startLine, codeText
  Else
  ' the selection spans multiple lines of code
  ' first, convert the highlighted text on the first line
  codeText = cmo.Lines(startLine, 1)
  If Conversion = SimpleSentence Then
   SimpleSentenceCase codeText
   Else
   Mid$(codeText, startCol, Len(codeText) + 1 - startCol) = strConv(Mid$(codeText, startCol, Len(codeText) + 1 - startCol), Conversion)
   If Conversion = vbProperCase Then
    ProperProperCase codeText
   End If
  End If
  cmo.ReplaceLine startLine, codeText
  ' then convert the lines in the middle, that are fully highlighted
  For I = startLine + 1 To EndLine - 1
   codeText = cmo.Lines(I, 1)
   If Conversion = SimpleSentence Then
    SimpleSentenceCase codeText
    Else
    codeText = strConv(codeText, Conversion)
    If Conversion = vbProperCase Then
     ProperProperCase codeText
    End If
   End If
   cmo.ReplaceLine I, codeText
  Next I
  ' finally, convert the highlighted portion of the last line
  codeText = cmo.Lines(EndLine, 1)
  If Conversion = SimpleSentence Then
   SimpleSentenceCase codeText
   Else
   Mid$(codeText, 1, endCol - 1) = strConv(Mid$(codeText, 1, endCol - 1), Conversion)
   If Conversion = vbProperCase Then
    ProperProperCase codeText
   End If
  End If
  cmo.ReplaceLine EndLine, codeText
 End If
 ' after replacing code we must restore the old selection
 ' this seems to be a side-effect of the ReplaceLine method
 cpa.SetSelection startLine, startCol, EndLine, endCol
 On Error GoTo 0

End Sub

Public Sub doCaseCycle()

 If CurCase = 0 Then
  CurCase = vbLowerCase
 End If
 ConvertSelectedText VBInstance, CurCase
 CurCase = CurCase + 1
 If CurCase = 5 Then
  CurCase = 1
 End If

End Sub

Public Sub ProperProperCase(strCode As String, _
                            Optional StrTrigger As String = vbNullString)

  Dim hasquotes As Long

 If LenB(StrTrigger) = 0 Then
  StrTrigger = Chr$(34)
 End If
 hasquotes = InStr(strCode, StrTrigger)
 Do While hasquotes
  strCode = Left$(strCode, hasquotes) & UCase$(Mid$(strCode, hasquotes + 1, 1)) & Mid$(strCode, hasquotes + 2)
  hasquotes = InStr(hasquotes + 1, strCode, StrTrigger)
 Loop

End Sub

Public Sub SimpleSentenceCase(strCode As String)

  Dim PunctSpace As Variant
  Dim Punct      As Variant
  Dim I          As Long

 PunctSpace = Array("! ", "? ", ". ", ", ", ": ", "; ")
 Punct = Array("[", "{", "(", "<", "'", Chr$(34))
 ProperProperCase strCode, Chr$(34)
 For I = 0 To UBound(Punct)
  ProperProperCase strCode, CStr(Punct(I))
 Next I
 For I = 0 To UBound(PunctSpace)
  ProperProperCase strCode, CStr(PunctSpace(I))
 Next I

End Sub

':) Roja's VB Code Fixer V1.1.4 (30/07/2003 1:25:30 PM) 11 + 141 = 152 Lines Thanks Ulli for inspiration and lots of code.

