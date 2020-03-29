VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00D6D6D6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UnChefer file  | |"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000A&
      Caption         =   "⁄œ„ «ŸÂ«— Â–Â «·—”«·… „—… «Œ—Ï "
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5655
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   75
      TabIndex        =   2
      Top             =   720
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton UnCommand1 
      Caption         =   "Ok"
      Height          =   345
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ‰»ÌÂ «‰ ·„  ﬂ‰ ﬂ·„… «· ”ÃÌ· ’ÕÌÕ… ›·‰ Ì „ ›ﬂ «·„·› «·„‘›— »’Ê—… ’ÕÌÕ… "
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4605
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

Private Sub Check1_Click()
    ' Ì„ﬂ‰ﬂ  ⁄œÌ· Œ«’Ì… «·ﬂ«»‘‰ „‰ «·„·› «·Œ«—ÃÌ «œ« «—œ  ÷ÂÊ—Â« „—… «Œ—Ï »⁄œ «·€«∆Â«
End Sub

Public Sub UnCommand1_Click()
    Call Out_Byte
End Sub



