VERSION 5.00
Begin VB.Form frmUnDo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "UnDo"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdUnDo 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.ListBox lstUnDo 
      Height          =   2985
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton CmdUnDo 
      Caption         =   "UnDo Selected"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton CmdUnDo 
      Caption         =   "UnDo All"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "frmUnDo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdUnDo_Click(Index As Integer)

 If Index = 2 Then
  Me.Hide
  Else
  UnDoAction Index
 End If

End Sub

':) Roja's VB Code Fixer V1.1.4 (30/07/2003 1:25:29 PM) 1 + 13 = 14 Lines Thanks Ulli for inspiration and lots of code.
