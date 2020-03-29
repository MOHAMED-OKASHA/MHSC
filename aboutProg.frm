VERSION 5.00
Begin VB.Form aboutProg 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9090
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "aboutProg.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "aboutProg.frx":030A
   ScaleHeight     =   264
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   606
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Image1 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   2715
      Width           =   60
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      ToolTipText     =   "«÷€ÿ Â‰« ·‰”Œ «·«Ì„Ì· ›Ì «·Õ«›Ÿ…"
      Top             =   3480
      Width           =   5415
   End
End
Attribute VB_Name = "aboutProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Hide
End Sub

Private Sub Image1_Click()
    Dim ret&
    ret = ShellExecute(Me.hwnd, "Open", "http://www.vb4arab.com", "", "", 1) '   › Õ —«»ÿ «·«‰ —‰  ⁄‰œ «·‰ﬁ— ⁄·Ï «·’Ê—…
End Sub

Private Sub Label1_Click()
    Clipboard.Clear
    Clipboard.SetText "Mohammad.okasha@gmail.com" ' ·‰”Œ «·«Ì„Ì· ›Ì «·–«ﬂ—…
End Sub

