Attribute VB_Name = "APISupport"
Option Explicit
'This module contains all the API stuff so it is easier to find/update
'Combo Find Consts
Private Const CB_FINDSTRING             As Long = &H14C
Private Const CB_FINDSTRINGEXACT        As Long = &H158
'TabWidth reader Consts
Private Const DefaultTabWidth           As Long = 4
Private FullTabWidth                    As Long
Private Const VBSettings                As String = "Software\Microsoft\VBA\Microsoft Visual Basic"
Private Const HKEY_CURRENT_USER         As Long = &H80000001
Private Const KEY_QUERY_VALUE           As Long = 1
Private Const REG_OPTION_RESERVED       As Long = 0
Private Const ERROR_NONE                As Long = 0
'Combo Find Declares
'TabWidth reader Declares
'EscPressed Declares
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal lpReserved As Long, _
                                                                                      lpType As Long, _
                                                                                      lpData As Any, _
                                                                                      lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Function EscPressed() As Boolean

  '*PURPOSE: detect Esc has been pressed

 EscPressed = (GetAsyncKeyState(vbKeyEscape) < 0)

End Function

Public Function GetFullTabWidth() As Long

  Dim K As Long

 If RegOpenKeyEx(HKEY_CURRENT_USER, VBSettings, REG_OPTION_RESERVED, KEY_QUERY_VALUE, 0) <> ERROR_NONE Then
  GetFullTabWidth = DefaultTabWidth
  Else
  K = Len(FullTabWidth)
  If RegQueryValueEx(0, "TabWidth", REG_OPTION_RESERVED, 0, FullTabWidth, K) <> ERROR_NONE Then
   GetFullTabWidth = DefaultTabWidth
  End If
 End If
 RegCloseKey 0

End Function

Public Function PosInCombo(ByVal strA As String, _
                           ByVal C As ComboBox, _
                           Optional CaseSensitive As Boolean = True) As Long

  'find if strA is in Combolist
  'returns -1 if not found

 PosInCombo = SendMessage(C.hWnd, IIf(CaseSensitive, CB_FINDSTRINGEXACT, CB_FINDSTRING), 0, ByVal strA)

End Function

':) Roja's VB Code Fixer V1.1.4 (30/07/2003 1:24:35 PM) 21 + 36 = 57 Lines Thanks Ulli for inspiration and lots of code.

