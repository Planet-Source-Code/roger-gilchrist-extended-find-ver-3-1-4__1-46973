Attribute VB_Name = "ArraySupport"
Option Explicit
Public StrFuncArray         As Variant
Public TypeSuffixArray      As Variant
Public AsTypeArray          As Variant
Public StandardTypes        As Variant
Public VBReservedWords As Variant

Public Sub InitArrays()

  'These are Ulli's orignal list of VB functions whose efficency is greatly enhanced by using them as string functions rather than Variants
  'by adding a $ to calls program speed is increased and overhead reduced
  'I have added Replace to the list
  'thanks to Rudz for suggesting Dir

 StrFuncArray = Array("Chr", "ChrB", "ChrW", "Command", "CurDir", "Date", "Environ", "Error", "Format", "Hex", "LCase", "Left", "LeftB", "LTrim", "Mid", "MidB", "Oct", "Right", "RightB", _
                      "RTrim", "Space", "Str", "String", "Time", "Trim", "UCase", "Replace", "Dir")
 'These allow translating old style Type suffixes into As Type
 TypeSuffixArray = Array("%", "&", "!", "#", "@", "$")
 AsTypeArray = Array("Integer", "Long", "Single", "Double", "Currency", "String")
 StandardTypes = Array("Boolean", "Byte", "Integer", "Long", "Single", "Double", "Currency", "String", "Variant", "Date", "Object")
    VBReservedWords = Array("Alias", "And", "As", "Base", "Binary", "Boolean", "Byte", "ByVal", "Call", "Case", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng", "Close", _
                            "Compare", "Const", "CSng", "CStr", "Currency", "CVar", "CVErr", "Decimal", "Declare", "DefBool", "DefByte", "DefCur", "DefDate", "DefDbl", "DefDec", _
                            "DefInt", "DefLng", "DefObj", "DefSng", "DefStr", "DefVar", "Dim", "Do", "Double", "Each", "Else", "ElseIf", "End", "Enum", "Eqv", "Erase", "Error", _
                            "Exit", "Explicit", "False", "For", "Function", "Get", "Global", "GoSub", "GoTo", "If", "Imp", "In", "Input", "Input", "Integer", "Is", "LBound", "Let", _
                            "Lib", "Like", "Line", "Lock", "Long", "Loop", "LSet", "Name", "New", "Next", "Not", "Object", "Open", "Option", "On", "Or", "Output", "Preserve", "Print", _
                            "Private", "Property", "Public", "Put", "Random", "Read", "ReDim", "Resume", "Return", "RSet", "Seek", "Select", "Set", "Single", "Spc", "Static", "String", _
                            "Stop", "Sub", "Tab", "Then", "True", "UBound", "Variant", "While", "Wend", "With")

End Sub

Public Function InstrAtPositionArray(ByVal StrSearch As String, _
                                     ByVal AtLocation As InstrLocations, _
                                     ByVal WholeWord As Boolean, _
                                     ParamArray FindA() As Variant) As Boolean

  Dim findMember As Variant

 'Copyright 2003 Roger Gilchrist
 'e-mail: rojagilkrist@hotmail.com
 'check that any member of FindA is space delimited part of StrSearch at position AtLocation
 'See InstrAtPosition for parameter details
 For Each findMember In FindA
  If LenB(findMember) Then
   If InstrAtPosition(StrSearch, findMember, AtLocation, WholeWord) Then
    InstrAtPositionArray = True
    Exit For
   End If
  End If
 Next findMember

End Function

':) Roja's VB Code Fixer V1.1.4 (30/07/2003 1:25:29 PM) 5 + 37 = 42 Lines Thanks Ulli for inspiration and lots of code.

