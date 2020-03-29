VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                    ŒÌ«—« "
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3720
   Icon            =   "FrmOption.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   " „"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ê÷⁄ «· ‘›Ì—"
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3495
      Begin VB.OptionButton Option2 
         Caption         =   "«·Ê÷⁄ «· «‰Ì (  ‘›Ì— „⁄ „‰⁄ «·Ê’Ê· )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "(«·Ê÷⁄ «·«Ê· ( ‘›Ì— ›ﬁÿ   "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "⁄„· «·»—‰«„Ã „⁄ »œ√ «· ‘€Ì·"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2880
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmOption.frx":83E2
      Left            =   120
      List            =   "FrmOption.frx":83EC
      TabIndex        =   0
      Text            =   "«Œ — «··€… „‰ Â‰« "
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "«··€…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
    Select Case Check1
    Case 1
        Set wsh = CreateObject("WScript.Shell")
        wsh.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe", "REG_SZ"
    Case 0
        Set wsh = CreateObject("WScript.Shell")
        wsh.Regdelete "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName
    End Select
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Call msave
    frmSplash.Show
    Me.Hide
    DoEvents
    Call SytemFormat
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Command1_Click
End Sub

Public Sub msave()
    On Error GoTo MsavErr
    Dim Tmdataout As String * 15
    Open App.Path & "\option.ini" For Random As #5 Len = 15
        Select Case Combo1.ListIndex
        Case 0
            Tmdataout = "Ms" & Combo1.List(Combo1.ListIndex)         '"Arabia"
            Put #5, 1, Tmdataout
        Case 1
            Tmdataout = "Ms" & Combo1.List(Combo1.ListIndex)
            Put #5, 1, Tmdataout
        Case Else
            Tmdataout = "Ms" & "Defult:1"
            Put #5, 1, Tmdataout
        End Select
        If Option1.Value = True Then
            Tmdataout = "Ms" & "Mode1"
            Put #5, 2, Tmdataout
        Else
            Tmdataout = "Ms" & "Mode2"
            Put #5, 2, Tmdataout
        End If
        If Form3.Check1.Value = vbChecked Then
            Tmdataout = "Ms" & "Checked"
            Put #5, 3, Tmdataout
        Else
            Tmdataout = "Ms" & "UnChecked"
            Put #5, 3, Tmdataout
        End If
        If FrmOption.Check1.Value = vbChecked Then
            Tmdataout = "Ms" & "Checked"
            Put #5, 4, Tmdataout
        Else
            Tmdataout = "Ms" & "UnChecked"
            Put #5, 4, Tmdataout
        End If
        Reset
        Exit Sub
MsavErr:
        MsgBox "ÕœÀ Œÿ√ «À‰«¡ ⁄„·Ì…  Œ“Ì‰ «⁄œ«œ«  «·»—‰«„Ã " & vbCrLf & "Ì„ﬂ‰ «‰  ﬂÊ‰ Â–Â «·„‘ﬂ·… »”»» " & vbCrLf & "ﬁÌ«„ﬂ »«·⁄»À ›Ì „·› «·«⁄œ«œ«   option «À‰«¡ ⁄„«Ì… «· Œ“Ì‰ «Ê " & vbCrLf & "ﬁÌ«„ﬂ » €Ì— «·«⁄œ«œ«  «À‰«¡ ⁄„·Ì…  ‘›Ì— «Ê ›ﬂ  ‘›Ì— „·› ", vbCritical + vbMsgBoxRight, " ‰»ÌÂ"
End Sub

Private Sub Option2_Click()
' "›Ì «·«’œ«— «· «‰Ì «‰ ‘«¡ «··Â"

End Sub

