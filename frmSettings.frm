VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   14235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmProp 
      Caption         =   "Format"
      Height          =   3735
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   8775
      Begin VB.PictureBox picCFXPBugFix2 
         BorderStyle     =   0  'None
         Height          =   3485
         Left            =   120
         ScaleHeight     =   3480
         ScaleWidth      =   8580
         TabIndex        =   6
         Top             =   120
         Width           =   8580
         Begin VB.Frame Frame11 
            Caption         =   "Format"
            Height          =   1455
            Left            =   0
            TabIndex        =   64
            Top             =   1320
            Width           =   2895
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Proc Declation Move to Top"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   69
               Top             =   825
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Declaration Expand Multi"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   68
               Top             =   630
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Declaration Single Type FIx"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   67
               Top             =   435
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Declaration As/Comment Format"
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   66
               Top             =   1020
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Sort Procedures"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   65
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Fix Comments"
            Height          =   855
            Left            =   3120
            TabIndex        =   56
            Top             =   2520
            Width           =   3015
            Begin VB.CheckBox ChkFixCom 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Previous Code"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   58
               Top             =   480
               Width           =   2655
            End
            Begin VB.CheckBox ChkFixCom 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Fix Comment"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   57
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.CommandButton CmdFormat 
            Caption         =   "Format Now"
            Height          =   375
            Left            =   7200
            TabIndex        =   55
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Frame Frame9 
            Caption         =   "Fixes"
            Height          =   2415
            Left            =   3120
            TabIndex        =   52
            Top             =   0
            Width           =   3015
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Enum Case Protection"
               Height          =   195
               Index           =   14
               Left            =   120
               TabIndex        =   72
               Top             =   2160
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Chr$(#) to vbConstant"
               Height          =   195
               Index           =   13
               Left            =   120
               TabIndex        =   71
               Top             =   1920
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Pleonasm Fix (unneeded '= True')"
               Height          =   195
               Index           =   12
               Left            =   120
               TabIndex        =   70
               Top             =   1680
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Test Scope && Useage "
               Height          =   195
               Index           =   11
               Left            =   120
               TabIndex        =   63
               Top             =   1440
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Comment out Unused Structures"
               Height          =   195
               Index           =   10
               Left            =   120
               TabIndex        =   62
               Top             =   1230
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Type Suffix Update"
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   61
               Top             =   255
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "VB String Function Update"
               Height          =   195
               Index           =   9
               Left            =   120
               TabIndex        =   60
               Top             =   1035
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "String Concantation Update"
               Height          =   195
               Index           =   8
               Left            =   120
               TabIndex        =   59
               Top             =   840
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Expand If..Then... Structures"
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   54
               Top             =   645
               Width           =   2655
            End
            Begin VB.CheckBox ChkFix 
               Alignment       =   1  'Right Justify
               Caption         =   "Expand Colon Separators"
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   53
               Top             =   450
               Width           =   2655
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Indenting"
            Height          =   1215
            Left            =   20
            TabIndex        =   7
            Top             =   0
            Width           =   2895
            Begin VB.CheckBox ChkIndent 
               Alignment       =   1  'Right Justify
               Caption         =   "Show indenting (Slower)"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   2655
            End
            Begin VB.CheckBox ChkIndent 
               Alignment       =   1  'Right Justify
               Caption         =   "Remove Double Blanks"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   10
               Top             =   435
               Width           =   2655
            End
            Begin VB.CheckBox ChkIndent 
               Alignment       =   1  'Right Justify
               Caption         =   "Space out Structures"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   9
               Top             =   825
               Width           =   2655
            End
            Begin VB.CheckBox ChkIndent 
               Alignment       =   1  'Right Justify
               Caption         =   "Delete All Blanks"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   8
               Top             =   630
               Width           =   2655
            End
         End
      End
   End
   Begin VB.Frame frmProp 
      Caption         =   "Find"
      Height          =   3615
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   8775
      Begin VB.PictureBox picCFXPBugFix1 
         BorderStyle     =   0  'None
         Height          =   3365
         Left            =   100
         ScaleHeight     =   3360
         ScaleWidth      =   8580
         TabIndex        =   12
         Top             =   60
         Width           =   8580
         Begin VB.CheckBox ChkLaunchStartup 
            Caption         =   "Launch On Startup"
            Height          =   195
            Left            =   140
            TabIndex        =   51
            Top             =   3040
            Width           =   1935
         End
         Begin VB.Frame Frame2 
            Caption         =   "Search History Size "
            Height          =   1095
            Left            =   2775
            TabIndex        =   46
            Top             =   1845
            Width           =   2655
            Begin VB.PictureBox picCFXPBugFix0 
               BorderStyle     =   0  'None
               Height          =   840
               Left            =   100
               ScaleHeight     =   840
               ScaleWidth      =   2460
               TabIndex        =   47
               Top             =   175
               Width           =   2460
               Begin VB.CommandButton cmdClearHistory 
                  Caption         =   "Clear"
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   49
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.CheckBox ChkSaveHistory 
                  Caption         =   "Save"
                  Height          =   195
                  Left            =   20
                  TabIndex        =   48
                  Top             =   480
                  Width           =   735
               End
               Begin MSComctlLib.Slider SliderHistory 
                  Height          =   315
                  Left            =   15
                  TabIndex        =   50
                  Top             =   45
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   556
                  _Version        =   393216
                  LargeChange     =   36
                  Min             =   20
                  Max             =   200
                  SelStart        =   40
                  TickFrequency   =   20
                  Value           =   40
               End
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   " Found Grid Appearance"
            Height          =   1695
            Left            =   -100
            TabIndex        =   36
            Top             =   120
            Width           =   5535
            Begin VB.CheckBox ChkShow 
               Caption         =   "Project (If more than one)"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   2535
            End
            Begin VB.CheckBox ChkShow 
               Caption         =   "Component (If more than one)"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   44
               Top             =   480
               Width           =   2535
            End
            Begin VB.CheckBox ChkShow 
               Caption         =   "Procedure Name"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   43
               Top             =   720
               Width           =   2535
            End
            Begin VB.CheckBox ChkShow 
               Caption         =   "Show Grid Lines"
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   42
               Top             =   1080
               Width           =   2535
            End
            Begin VB.CheckBox ChkShow 
               Caption         =   "Show Procedure  Lines"
               Height          =   195
               Index           =   4
               Left            =   2640
               TabIndex        =   41
               Top             =   720
               Width           =   2535
            End
            Begin VB.CheckBox ChkShow 
               Caption         =   "Show Component Lines"
               Height          =   195
               Index           =   2
               Left            =   2640
               TabIndex        =   40
               Top             =   480
               Width           =   2535
            End
            Begin VB.CheckBox ChkAutoSelectedText 
               Caption         =   "Auto Selected Text"
               Height          =   195
               Left            =   2640
               TabIndex        =   39
               Top             =   1320
               Width           =   1935
            End
            Begin VB.CheckBox ChkSelectWhole 
               Caption         =   "Find select whole line"
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   1320
               Width           =   1935
            End
            Begin VB.CheckBox ChkRemFilters 
               Caption         =   "Remember Filters"
               Height          =   195
               Left            =   2640
               TabIndex        =   37
               Top             =   1080
               Width           =   1935
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Replace"
            Height          =   1095
            Left            =   20
            TabIndex        =   32
            Top             =   1845
            Width           =   2655
            Begin VB.CheckBox ChkReplace 
               Alignment       =   1  'Right Justify
               Caption         =   "  Show Filter Warning"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   2415
            End
            Begin VB.CheckBox ChkReplace 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Blank Warning"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   34
               Top             =   480
               Width           =   2415
            End
            Begin VB.CheckBox ChkReplace 
               Alignment       =   1  'Right Justify
               Caption         =   "Add Replace To Search"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   33
               Top             =   720
               Width           =   2415
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Colours"
            Height          =   3015
            Left            =   5535
            TabIndex        =   13
            Top             =   45
            Width           =   3015
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   2760
               Left            =   100
               ScaleHeight     =   2760
               ScaleWidth      =   2820
               TabIndex        =   14
               Top             =   175
               Width           =   2815
               Begin VB.PictureBox Picture2 
                  BorderStyle     =   0  'None
                  Height          =   2760
                  Left            =   0
                  ScaleHeight     =   2760
                  ScaleWidth      =   2820
                  TabIndex        =   18
                  Top             =   -25
                  Width           =   2820
                  Begin VB.Frame Frame5 
                     Caption         =   "Default"
                     Height          =   2175
                     Left            =   1440
                     TabIndex        =   22
                     Top             =   0
                     Width           =   1320
                     Begin VB.Label LblColour 
                        Alignment       =   2  'Center
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Standard"
                        Height          =   255
                        Index           =   5
                        Left            =   120
                        TabIndex        =   29
                        Top             =   720
                        Width           =   1095
                     End
                     Begin VB.Label LblColour 
                        Alignment       =   2  'Center
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Searching"
                        Height          =   255
                        Index           =   6
                        Left            =   120
                        TabIndex        =   28
                        Top             =   960
                        Width           =   1095
                     End
                     Begin VB.Label LblColour 
                        Alignment       =   2  'Center
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Pattern Find"
                        Height          =   255
                        Index           =   7
                        Left            =   120
                        TabIndex        =   27
                        Top             =   1230
                        Width           =   1095
                     End
                     Begin VB.Label LblColour 
                        Alignment       =   2  'Center
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "No Find"
                        Height          =   255
                        Index           =   8
                        Left            =   120
                        TabIndex        =   26
                        Top             =   1485
                        Width           =   1095
                     End
                     Begin VB.Label LblColour 
                        Alignment       =   2  'Center
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Text"
                        Height          =   255
                        Index           =   4
                        Left            =   120
                        TabIndex        =   25
                        Top             =   240
                        Width           =   1095
                     End
                     Begin VB.Label Label1 
                        Alignment       =   2  'Center
                        Caption         =   "Back Colours"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   24
                        Top             =   480
                        Width           =   1095
                     End
                     Begin VB.Label LblColour 
                        Alignment       =   2  'Center
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Replacing"
                        Height          =   255
                        Index           =   9
                        Left            =   120
                        TabIndex        =   23
                        Top             =   1800
                        Width           =   1095
                     End
                  End
                  Begin VB.Frame Frame7 
                     Caption         =   "General"
                     Height          =   855
                     Left            =   0
                     TabIndex        =   19
                     Top             =   0
                     Width           =   1335
                     Begin VB.Label LblColour 
                        Alignment       =   2  'Center
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Back"
                        Height          =   255
                        Index           =   1
                        Left            =   120
                        TabIndex        =   21
                        Top             =   480
                        Width           =   1095
                     End
                     Begin VB.Label LblColour 
                        Alignment       =   2  'Center
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Text"
                        Height          =   255
                        Index           =   0
                        Left            =   120
                        TabIndex        =   20
                        Top             =   240
                        Width           =   1095
                     End
                  End
                  Begin VB.Label LblColour 
                     Alignment       =   2  'Center
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Default"
                     Height          =   255
                     Index           =   11
                     Left            =   1560
                     TabIndex        =   31
                     Top             =   2400
                     Width           =   1095
                  End
                  Begin VB.Label LblColour 
                     Alignment       =   2  'Center
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Restore User"
                     Height          =   255
                     Index           =   10
                     Left            =   135
                     TabIndex        =   30
                     Top             =   2400
                     Width           =   1095
                  End
               End
               Begin VB.Frame Frame6 
                  Caption         =   "Selection"
                  Height          =   855
                  Left            =   20
                  TabIndex        =   15
                  Top             =   885
                  Width           =   1335
                  Begin VB.Label LblColour 
                     Alignment       =   2  'Center
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Back"
                     Height          =   255
                     Index           =   3
                     Left            =   120
                     TabIndex        =   17
                     Top             =   480
                     Width           =   1095
                  End
                  Begin VB.Label LblColour 
                     Alignment       =   2  'Center
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Text"
                     Height          =   255
                     Index           =   2
                     Left            =   120
                     TabIndex        =   16
                     Top             =   240
                     Width           =   1095
                  End
               End
            End
         End
      End
   End
   Begin VB.CommandButton CmdSetting 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Index           =   2
      Left            =   5160
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton CmdSetting 
      Caption         =   "Apply"
      Height          =   315
      Index           =   1
      Left            =   3960
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton CmdSetting 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4095
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7223
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Key             =   "find"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Format"
            Key             =   "format"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'track current frame
Private mintCurFrame                   As Long
'Safety switch for Apply and using control box to close
Private ApplyClicked                   As Boolean
'Preserves values at form load
Private OrigShowProj                   As Boolean
Private OrigShowComp                   As Boolean
Private OrigShowRout                   As Boolean
Private OrigLaunchStart                As Boolean
Private OrigRemFilt                    As Boolean
Private OrigSaveHist                   As Boolean
Private OrighistDeep                   As Long
Private OrigbFindSelectWholeLine       As Boolean
'Indet Originals
Private OrigVisibleIndenting           As Boolean
Private OrigAddStructureSpace          As Boolean
Private OrigDeleteDoubleBlanks         As Boolean
Private OrigDeleteAllBlanks            As Boolean
'fix Originals
Private OrigSortModules                As Boolean
Private OrigProcDecl2Top               As Boolean
Private OrigDeclSingleTypeExpand       As Boolean
Private OrigDeclAsFormat               As Boolean
Private OrigTypeSuffixUpdate           As Boolean
Private OrigDeclExpand                 As Boolean
Private OrigExpandColon                As Boolean
Private OrigExpandIfThen               As Boolean
Private OrigStrConcatenateUpdate       As Boolean
Private OrigStrFunctionUpdate          As Boolean
Private OrigCommentOutUnused           As Boolean
Private OrigTestScope                  As Boolean
Private OrigPleonasmFix As Boolean
Private OrigChr2ConstFix As Boolean
Private OrigEnumCapProtect As Boolean
'fix comment Originals
Private OrigShowPrevCode               As Boolean
Private OrigShowFixComment             As Boolean
Private ColourTextForeUnDo             As Long
Private ColourTextBackUnDo             As Long
Private ColourFindSelectForeUnDo       As Long
Private ColourFindSelectBackUnDo       As Long
Private ColourHeadDefaultUnDo          As Long
Private ColourHeadWorkUnDo             As Long
Private ColourHeadPatternUnDo          As Long
Private ColourHeadNoFindUnDo           As Long
Private ColourHeadForeUnDo             As Long
Private ColourHeadReplaceUnDo          As Long

Private Sub ChkAutoSelectedText_Click()

 If Not bLoadingSettings Then
  bAutoSelectText = ChkLaunchStartup.value = 1
 End If

End Sub

Private Sub ChkFix_Click(Index As Integer)

  'fix switches

 If Not bLoadingSettings Then
  bSortModules = ChkFix(0).value = 1
  bProcDecl2Top = ChkFix(1).value = 1
  bDeclSingleTypeExpand = ChkFix(2).value = 1
  bDeclExpand = ChkFix(3).value = 1
  bDeclAsFormat = ChkFix(4).value = 1
  bTypeSuffixUpdate = ChkFix(5).value = 1
  bExpandColon = ChkFix(6).value = 1
  bExpandIfThen = ChkFix(7).value = 1
  bStrConcatenateUpdate = ChkFix(8).value = 1
  bStrFunctionUpdate = ChkFix(9).value = 1
  bCommentOutUnused = ChkFix(10).value = 1
  bTestScope = ChkFix(11).value = 1
  bPleonasmFix = ChkFix(12).value = 1
  bChr2ConstFix = ChkFix(13).value = 1
  bEnumCapProtect = ChkFix(14).value = 1
 End If

End Sub

Private Sub ChkFixCom_Click(Index As Integer)

  'fix commenting switches

 If Not bLoadingSettings Then
  bShowFixComment = ChkFixCom(0).value = 1
  bShowPrevCode = ChkFixCom(1).value = 1
 End If

End Sub

Private Sub ChkIndent_Click(Index As Integer)

  'Indent switches

 If Not bLoadingSettings Then
  bVisibleIndenting = ChkIndent(0).value = 1
  bDeleteDoubleBlanks = ChkIndent(1).value = 1
  If bDeleteDoubleBlanks And Index = 1 Then
   ChkIndent(2).value = 0
  End If
  bDeleteAllBlanks = ChkIndent(2).value = 1
  If bDeleteAllBlanks And Index = 2 Then
   ChkIndent(1).value = 0
  End If
  bAddStructureSpace = ChkIndent(3).value = 1
 End If

End Sub

Private Sub ChkLaunchStartup_Click()

 If Not bLoadingSettings Then
  bLaunchOnStart = ChkLaunchStartup.value = 1
 End If

End Sub

Private Sub ChkRemFilters_Click()

 If Not bLoadingSettings Then
  bRemFilters = ChkRemFilters.value = 1
 End If

End Sub

Private Sub ChkReplace_Click(Index As Integer)

  'replace switches

 If Not bLoadingSettings Then
  bFilterWarning = ChkReplace(0).Index = 1
  bBlankWarning = ChkReplace(1).Index = 1
  bReplace2Search = ChkReplace(2).Index = 1
 End If

End Sub

Private Sub ChkSaveHistory_Click()

 If Not bLoadingSettings Then
  bSaveHistory = ChkSaveHistory.value = 1
 End If

End Sub

Private Sub ChkSelectWhole_Click()

 If Not bLoadingSettings Then
  bFindSelectWholeLine = ChkSelectWhole.value = 1
 End If

End Sub

Private Sub ChkShow_Click(Index As Integer)

  'Found Appearance switches

 If Not bLoadingSettings Then
  bShowProject = ChkShow(0).value = 1
  bShowComponent = ChkShow(1).value = 1
  bShowCompLineNo = ChkShow(2).value = 1
  bShowRoutine = ChkShow(3).value = 1
  bShowProcLineNo = ChkShow(4).value = 1
  bGridlines = ChkShow(5).value = 1
 End If

End Sub

Private Sub cmdClearHistory_Click()

 mobjDoc.ClearHistory

End Sub

Private Sub CmdFormat_Click()

 DoIndent

End Sub

Private Sub CmdSetting_Click(Index As Integer)

  ' OK/Apply/Cancel buttons

 ApplyClicked = False
 Select Case Index
  Case 0
  Me.Hide
  mobjDoc.ApplyChanges
  Case 1
  mobjDoc.ApplyChanges
  ApplyClicked = True
  Case 2
  RestoreOriginals
  Me.Hide
 End Select
 SaveFormPosition Me

End Sub

Private Sub Form_Load()

  Dim frm As Frame

 For Each frm In frmProp
  With frm
   .Visible = False
   .Caption = vbNullString
   .BorderStyle = 0
  End With 'frm
 Next frm
 LoadFormPosition Me
 Me.Width = TabStrip1.Width
 Me.Height = TabStrip1.Height + offset * 7
 Frame2Tab TabStrip1, frmProp, mintCurFrame
 Frame8.Caption = "Indenting (Size =" & GetFullTabWidth & ")"
 mobjDoc.ColoursApply
 ColourTextForeUnDo = ColourTextFore
 ColourTextBackUnDo = ColourTextBack
 ColourFindSelectForeUnDo = ColourFindSelectFore
 ColourFindSelectBackUnDo = ColourFindSelectBack
 ColourHeadDefaultUnDo = ColourHeadDefault
 ColourHeadWorkUnDo = ColourHeadWork
 ColourHeadPatternUnDo = ColourHeadPattern
 ColourHeadNoFindUnDo = ColourHeadNoFind
 ColourHeadForeUnDo = ColourHeadFore
 Me.Caption = "Properties " & AppDetails
 'set safety values for Cancel button
 OrigbFindSelectWholeLine = bFindSelectWholeLine
 OrigShowProj = bShowProject
 OrigShowComp = bShowComponent
 OrigShowRout = bShowRoutine
 OrigLaunchStart = bLaunchOnStart
 OrigVisibleIndenting = bVisibleIndenting
 OrigDeleteDoubleBlanks = bDeleteDoubleBlanks
 OrigDeleteAllBlanks = bDeleteAllBlanks
 OrigAddStructureSpace = bAddStructureSpace
 OrigExpandColon = bExpandColon
 OrigSortModules = bSortModules
 OrigProcDecl2Top = bProcDecl2Top
 OrigDeclSingleTypeExpand = bDeclSingleTypeExpand
 OrigDeclExpand = bDeclExpand
 OrigDeclAsFormat = bDeclAsFormat
 OrigCommentOutUnused = bCommentOutUnused
 OrigTestScope = bTestScope
 OrigPleonasmFix = bPleonasmFix
 OrigChr2ConstFix = bChr2ConstFix
 OrigEnumCapProtect = bEnumCapProtect
 OrigTypeSuffixUpdate = bTypeSuffixUpdate
 OrigExpandIfThen = bExpandIfThen
 OrigStrConcatenateUpdate = bStrConcatenateUpdate
 OrigStrFunctionUpdate = bStrFunctionUpdate
 OrigShowPrevCode = bShowPrevCode
 OrigShowFixComment = bShowFixComment
 OrigRemFilt = bRemFilters
 OrigSaveHist = bSaveHistory
 OrighistDeep = HistDeep

End Sub

Private Sub Form_Unload(Cancel As Integer)

 If Not ApplyClicked Then
  ' keeps changes if user clicks 'Apply' then uses CaptionBar 'X' button to close
  'otherwise restore
  RestoreOriginals
 End If
 SaveFormPosition Me
 Me.Hide

End Sub

Private Sub LblColour_Click(Index As Integer)

 Select Case Index
  Case 11 'Standard colours
  mobjDoc.ColoursStandard
  Case 10 'Undo
  ColourTextFore = ColourTextForeUnDo
  ColourTextBack = ColourTextBackUnDo
  ColourFindSelectFore = ColourFindSelectForeUnDo
  ColourFindSelectBack = ColourFindSelectBackUnDo
  ColourHeadDefault = ColourHeadDefaultUnDo
  ColourHeadWork = ColourHeadWorkUnDo
  ColourHeadPattern = ColourHeadPatternUnDo
  ColourHeadNoFind = ColourHeadNoFindUnDo
  ColourHeadFore = ColourHeadForeUnDo
  ColourHeadReplace = ColourHeadReplaceUnDo
  Case Else
  With CommonDialog1
   .Flags = cdlCCRGBInit Or cdlCCFullOpen
   'set current color as default
   Select Case Index
    Case 0
    .Color = ColourTextFore
    Case 1
    .Color = ColourTextBack
    Case 2
    .Color = ColourFindSelectFore
    Case 3
    .Color = ColourFindSelectBack
    Case 4
    .Color = ColourHeadFore
    Case 5
    .Color = ColourHeadDefault
    Case 6
    .Color = ColourHeadWork
    Case 7
    .Color = ColourHeadPattern
    Case 8
    .Color = ColourHeadNoFind
    Case 8
    .Color = ColourHeadReplace
   End Select
   .ShowColor
   'apply new or default colour
   If Not .CancelError Then
    Select Case Index
     Case 0
     ColourTextFore = .Color
     Case 1
     ColourTextBack = .Color
     Case 2
     ColourFindSelectFore = .Color
     Case 3
     ColourFindSelectBack = .Color
     Case 4
     ColourHeadFore = .Color
     Case 5
     ColourHeadDefault = .Color
     Case 6
     ColourHeadWork = .Color
     Case 7
     ColourHeadPattern = .Color
     Case 8
     ColourHeadNoFind = .Color
     Case 9
     ColourHeadReplace = .Color
    End Select
   End If
  End With
 End Select
 mobjDoc.ColoursApply

End Sub

Private Sub RestoreOriginals()

 ChkSelectWhole.value = Bool2Int(OrigbFindSelectWholeLine)
 ChkRemFilters.value = Bool2Int(OrigRemFilt)
 ChkLaunchStartup.value = Bool2Int(OrigLaunchStart)
 ChkShow(0).value = Bool2Int(OrigShowProj)
 ChkShow(1).value = Bool2Int(OrigShowComp)
 ChkShow(2).value = Bool2Int(OrigShowRout)
 ChkSaveHistory.value = Bool2Int(OrigSaveHist)
 '
 ChkIndent(0).value = Bool2Int(OrigVisibleIndenting)
 ChkIndent(1).value = Bool2Int(OrigDeleteDoubleBlanks)
 ChkIndent(2).value = Bool2Int(OrigDeleteAllBlanks)
 ChkIndent(3).value = Bool2Int(OrigAddStructureSpace)
 '
 ChkFix(0).value = Bool2Int(OrigSortModules)
 ChkFix(1).value = Bool2Int(OrigProcDecl2Top)
 ChkFix(2).value = Bool2Int(OrigDeclSingleTypeExpand)
 ChkFix(3).value = Bool2Int(OrigProcDecl2Top)
 ChkFix(4).value = Bool2Int(OrigDeclAsFormat)
 ChkFix(5).value = Bool2Int(OrigTypeSuffixUpdate)
 ChkFix(6).value = Bool2Int(OrigExpandColon)
 ChkFix(7).value = Bool2Int(OrigExpandIfThen)
 ChkFix(8).value = Bool2Int(OrigStrConcatenateUpdate)
 ChkFix(9).value = Bool2Int(OrigStrFunctionUpdate)
 ChkFix(10).value = Bool2Int(OrigTypeSuffixUpdate)
 ChkFix(11).value = Bool2Int(OrigTestScope)
ChkFix(12).value = Bool2Int(OrigPleonasmFix)
ChkFix(13).value = Bool2Int(OrigChr2ConstFix)
ChkFix(14).value = Bool2Int(OrigEnumCapProtect)
 '
 ChkFixCom(0).value = Bool2Int(OrigShowPrevCode)
 ChkFixCom(1).value = Bool2Int(OrigShowFixComment)
 SliderHistory.value = OrighistDeep

End Sub

Private Sub SliderHistory_Change()

 Frame2.Caption = "Search History Size (" & SliderHistory.Min & "-" & SliderHistory.Max & "):" & SliderHistory.value
 HistDeep = SliderHistory.value

End Sub

Private Sub TabStrip1_Click()

 Frame2Tab TabStrip1, frmProp, mintCurFrame

End Sub

':) Roja's VB Code Fixer V1.1.4 (30/07/2003 1:24:37 PM) 45 + 343 = 388 Lines Thanks Ulli for inspiration and lots of code.
