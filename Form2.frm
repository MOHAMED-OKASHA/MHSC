VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00D6D6D6&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "login Program"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Akhbar MT"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Text            =   "                          Confuger   password"
      Top             =   1560
      Width           =   5340
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Akhbar MT"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Text            =   "                                 Enter password"
      Top             =   960
      Width           =   5340
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Text            =   "                Enter Your Name"
      Top             =   390
      Width           =   5325
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim Sname As String

Private Sub Command1_Click()
    Me.Visible = False
    Form1.Show
End Sub

Private Sub Command2_Click()
    If Text2(0).Text <> Text2(1).Text Then MsgBox "⁄–—« ÌÊÃœ Œÿ√ " & vbCrLf & "ﬂ·„… «·”— €Ì— „ ÿ«»ﬁ…", vbCritical + vbMsgBoxRight, "Œÿ√": Exit Sub
    Dim str As String
    Dim i As Integer
    For i = 1 To Len(Text2(0).Text)
        If IsNumeric(Mid(Text2(0).Text, i, 1)) Then
            sPasword = sPasword & Mid(Text2(0).Text, i, 1)
        Else
            Select Case Mid(Text2(0).Text, i, 1)
                Dim Tmp As String
            Case "a", "A"
                Tmp = 6
            Case "b", "B"
                Tmp = 3
            Case "c", "C"
                Tmp = 2
            Case "d", "D"
                Tmp = 9
            Case "e", "E"
                Tmp = 1
            Case "f", "F"
                Tmp = 7
            Case "g", "G"
                Tmp = 4
            Case "h", "H"
                Tmp = 8
            Case "i", "I"
                Tmp = 5
            Case "j", "J"
                Tmp = 8
            Case "k", "K"
                Tmp = 2
            Case "l", "L"
                Tmp = 6
            Case "m", "M"
                Tmp = 9
            Case "n", "N"
                Tmp = 3
            Case "o", "O"
                Tmp = 1
            Case "p", "P"
                Tmp = 5
            Case "q", "Q"
                Tmp = 7
            Case "r", "R"
                Tmp = 6
            Case "s", "S"
                Tmp = 7
            Case "t", "T"
                Tmp = 0
            Case "u", "U"
                Tmp = 6
            Case "v", "V"
                Tmp = 8
            Case "W", "w"
                Tmp = 0
            Case "x", "X"
                Tmp = 2
            Case "y", "Y"
                Tmp = 0
            Case "z", "Z"
                Tmp = 1
            Case Else
                MsgBox "[ " & Mid(Text2(0).Text, i, 1) & "  ·« Ì”Õ »«” Œœ«„ Â–« «·Õ—› «Ê «·—„“ [ " & vbCrLf & "›ﬁÿ «·Õ—Ê› «·«‰Ã·Ì“Ì… Ê«·«—ﬁ«„", vbCritical, " ‰»ÌÂ"
                sPasword = ""
                Text2(0).Text = ""
                Text2(1).Text = ""
                Exit Sub
            End Select
            sPasword = sPasword & Tmp
        End If
    Next
    '    Â‰« Œ«’ »ﬁ”„ Õ›Ÿ «”„ «·‘Œ’ «·–Ì ﬁ«„ »«· ”ÃÌ·
    Form1.Label2.Caption = " User name : " & Text1.Text
    Form1.Label1.Visible = False
    Me.Hide
    Form1.Show
    Form1.Command1.Enabled = True
    Form1.Command2.Enabled = True
    Form1.Command3.Enabled = True
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Trim(Text1.Text) = "Enter Your Name" Then
        Text1.Text = ""
    End If
End Sub

Private Sub Text2_Change(Index As Integer)
    If Text2(0).Text <> "" Then Text2(1).Enabled = True Else Text2(1).Enabled = False
End Sub

Private Sub Text2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Trim(Text2(0).Text) = "Enter password" Then
        Text2(0).PasswordChar = "l"
        Text2(0).Font = "Wingdings"
        Text2(1).PasswordChar = "l"
        Text2(1).Font = "Wingdings"
        Text2(0).Text = ""
        Text2(1).Text = ""
    End If

End Sub
