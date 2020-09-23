VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmHelp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "help"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelpClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5318
      _Version        =   393217
      BackColor       =   -2147483624
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      FileName        =   "C:\Program Files\Microsoft Visual Studio\VB98\QND Programs\ExtFindD3\ExtendedFindHelp.rtf"
      TextRTF         =   $"frmHelp.frx":0000
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   4920
      Picture         =   "frmHelp.frx":BA5E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   300
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHelpClose_Click()

 SaveFormPosition Me
 Me.Hide

End Sub

Private Sub Form_Load()

 LoadFormPosition Me

End Sub

Private Sub Form_Resize()

 With RichTextBox1
  .Top = 0
  .Left = 0
  .Width = frmHelp.ScaleWidth
  .Height = frmHelp.ScaleHeight - cmdHelpClose.Height - offset
 End With
 With cmdHelpClose
  .Left = frmHelp.Width - .Width - offset
  .Top = RichTextBox1.Height + offset / 2
 End With
 SaveFormPosition Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

 SaveFormPosition Me

End Sub

':) Roja's VB Code Fixer V1.1.4 (30/07/2003 1:24:19 PM) 1 + 38 = 39 Lines Thanks Ulli for inspiration and lots of code.
