VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D6D6D6&
   Caption         =   "                                                           Security center  (Beta)"
   ClientHeight    =   2655
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9990
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   2655
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   300
      Left            =   8550
      TabIndex        =   12
      Top             =   45
      Width           =   300
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   8880
      Top             =   1080
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "Form1.frx":83E2
      Left            =   5280
      List            =   "Form1.frx":83E4
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2550
      Left            =   7920
      Picture         =   "Form1.frx":83E6
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   9975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   4
      Top             =   2400
      Width           =   9975
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   9360
      Top             =   1080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "un cypher file  | |"
      Enabled         =   0   'False
      Height          =   375
      Left            =   585
      TabIndex        =   2
      Top             =   525
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cypher file #"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2085
      TabIndex        =   1
      Top             =   525
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add file +"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3585
      Picture         =   "Form1.frx":5B1A2
      TabIndex        =   0
      Top             =   525
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   9000
      MouseIcon       =   "Form1.frx":5B5F6
      TabIndex        =   11
      Top             =   0
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "about prog && programmer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   0
      MouseIcon       =   "Form1.frx":5BA38
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1995
      Width           =   2250
   End
   Begin VB.Label statePut 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "sign in"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   4485
      MouseIcon       =   "Form1.frx":5BD42
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1050
      Width           =   1020
   End
   Begin VB.Menu nmenu1 
      Caption         =   "menu1"
      Visible         =   0   'False
      Begin VB.Menu nshowProg 
         Caption         =   "«ŸÂ«— «·»—‰«„Ã"
      End
      Begin VB.Menu naboutprog 
         Caption         =   "ÕÊ· «·»—‰«„Ã"
      End
      Begin VB.Menu nexit 
         Caption         =   "Œ—ÊÃ"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'MMttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttkkMM
'MM
'MM kDDDDtt        kkDDDDt ttDDDDDDDDDDDDtt                DDDDMM            MMDDDDtt        kkDDDDDDDDDDDDkk              DDDDDDtt        ttDDDDDDDDDDDDtt                      kkMMMMMMMMkk                                                        ttMM
'MM  MMMMMM        MMMMMM  kkMMMMMMMMMMMMMMMM            MMMMMMMM          ttMMMMMMDD        DDMMMMMMMMMMMMMMMM          ttMMMMMMMM        kkMMMMMMMMMMMMMMDD                  MMMMMMMMMMMMMMMM                                                      ttMM
'MM  MMMMMM        MMMMMM  kkMMMMDDkkDDMMMMMM          kkMMMMMMMM          MMMMMMMMMM        DDMMMMDDkkkkMMMMMMkk        DDMMMMMMMM        kkMMMMDDkkDDMMMMMM                MMMMMMMMtt  kkMMMM                                                      ttMM
'MM  DDMMMMtt    kkMMMMkk  kkMMMMkk    MMMMMM          MMMMMMMMMM          MMMMMMMMMMtt      DDMMMMtt    kkMMMMDD        MMMMMMMMMMtt      kkMMMMkk    MMMMMM              ttMMMMMM          DD      DDMMMMMMDD        MMMMttttMMMMDD    MMMMMM      ttMM
'MM  ttMMMMDD    MMMMMM    kkMMMMkk    MMMMMM        MMMM  DDMMMM        ttMMMMttMMMMDD      DDMMMMtt    kkMMMMkk      ttMMMMttMMMMMM      kkMMMMkk    MMMMMM              DDMMMMkk                MMMMMMMMMMMMMM      MMMMMMMMMMMMMMDDMMMMMMMMMM    ttMM
'MM    MMMMMM    MMMMDD    kkMMMMMMMMMMMMMM        MMMMkk  MMMMMM        MMMMMM  MMMMMM      DDMMMMDDttkkMMMMMM        DDMMMM  DDMMMM      kkMMMMMMMMMMMMMM                MMMMMM                DDMMMMkk  ttMMMMMM    MMMMMMttkkMMMMMMkkttMMMMMM    ttMM
'MM    DDMMMMttkkMMMMtt    kkMMMMMMMMMMMMMMMMtt  kkMMDD    DDMMMM        MMMMDD  kkMMMMtt    DDMMMMMMMMMMMMMMtt        MMMMDD  ttMMMMkk    kkMMMMMMMMMMMMMMMM              MMMMMM                MMMMMM      MMMMMM    MMMMkk    MMMMMM    kkMMMM    ttMM
'MM      MMMMDDMMMMMM      kkMMMMkk    ttMMMMMM  MMMMMMMMMMMMMMMMMMDD  kkMMMMDDttkkMMMMDD    DDMMMMMMMMMMMMDD        ttMMMMDDttkkMMMMMM    kkMMMMkk    kkMMMMMM            MMMMMMtt              MMMMMM      DDMMMM    MMMMkk    MMMMMM    kkMMMM    ttMM
'MM      MMMMMMMMMMDD      kkMMMMkk      MMMMMM  MMMMMMMMMMMMMMMMMMMM  MMMMMMMMMMMMMMMMMM    DDMMMMtt  MMMMMMkk      DDMMMMMMMMMMMMMMMM    kkMMMMkk      MMMMMM            kkMMMMDD          tt  MMMMMM      DDMMMM    MMMMkk    MMMMMM    kkMMMM    ttMM
'MM      kkMMMMMMMMtt      kkMMMMkk    kkMMMMMM  ttttttttttMMMMMMtt    MMMMMMDDDDDDMMMMMMtt  DDMMMMtt    MMMMMMtt    MMMMMMDDDDDDMMMMMMkk  kkMMMMkk    DDMMMMMM    MMMMDD    MMMMMMkk      kkMM  MMMMMM      MMMMMM    MMMMkk    MMMMMM    kkMMMM    ttMM
'MM        MMMMMMMM        kkMMMMMMMMMMMMMMMMtt            DDMMMM    kkMMMMkk        MMMMDD  DDMMMMtt    kkMMMMMM  ttMMMMkk        MMMMMM  kkMMMMMMMMMMMMMMMMtt    MMMMkk    kkMMMMMMMMMMMMMMMM    MMMMMMDDMMMMMMtt    MMMMkk    MMMMMM    kkMMMM    ttMM
'MM        MMMMMMDD        kkMMMMMMMMMMMMMMtt              MMMMMM    MMMMMM          MMMMMM  DDMMMMkk      DDMMMMM MMMMMM          MMMMMM  kkMMMMMMMMMMMMMMtt      MMMMDD      ttMMMMMMMMMMMMkk    ttMMMMMMMMMMtt      MMMMDD    MMMMMM    DDMMMMtt  ttMM
'MM                                                                                                                                                                                  tttt                tt                                          ttMM
'MM-thes program use a many of Controls but No need to Copy it in the system
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'MM My Name Mohamed -okasha \ libya-Zliten
'MM data 2011\3\25 start , 2011\5\4 complete
'MM my program it can Use in Cooding and UnCooding files
'MM for comment or suggestion You Can connect With Me On
'MM  mohamed_ok_1993@yahoo.com  -----  mohammad.okasha@gmail.com
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    Option Explicit
    Dim sv As Variant
    Dim m As Integer
    Public sPath As String
    Dim Lstate As String
    Dim Tmdata As String * 15
    Dim Temp1 As String
    Dim StartProc As Boolean
    Public WithEvents skin_Xp As XtremeSkinFramework.SkinFramework
Attribute skin_Xp.VB_VarHelpID = -1
  Public WithEvents SysTray1 As SysTray_Control.SysTray
Attribute SysTray1.VB_VarHelpID = -1
    Public WithEvents label_URl As R99x5.R99x5_URL
Attribute label_URl.VB_VarHelpID = -1

Private Sub Command1_Click()
    CommonDialog1.Filter = "*.*"
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    If CommonDialog1.FileTitle <> "" Then
        If CommonDialog1.FileTitle <> Temp1 Then
            Temp1 = CommonDialog1.FileTitle
            sPath = CommonDialog1.FileName
            List1.Visible = True
            List1.AddItem Mid(sPath, 1, 3) & "....." & CommonDialog1.FileTitle
            FileCopy sPath, "C:\Program Files\MHSC\" & CommonDialog1.FileTitle
        Else
            MsgBox "Â–« «·„·›  „  «÷«› Â »«·›⁄·", vbCritical + vbMsgBoxRight, " ‰»ÌÂ"
        End If
    End If
End Sub

Private Sub Command2_Click()


    On Error GoTo Chfr
    If sPath <> "" Then
    Dim t
    
    StartProc = True
    sPath = "C:\Program Files\MHSC\" & Mid(List1.List(List1.ListIndex), 9, Len(List1.List(List1.ListIndex)) - 8)
    If FileLen(sPath) / 1024 > 2500 Then
    t = MsgBox("„‰ «·„Õ „· «‰ Ì√Œ– «·»—‰«„Ã »÷⁄ œﬁ«∆ﬁ ›Ì  ‘›Ì— Â–« «·„·› " & vbCrLf & _
    "Ì‰’Õ «·»—‰«„Ã »«·÷€ÿ ⁄·Ï «·“— ... " & vbCrLf & _
    "· ’€Ì— «·»—‰«„Ã ÊÊ÷⁄Â »ÃÊ«— «·”«⁄… Ê”Ì „ «‘⁄«—ﬂ ⁄‰œ «·«‰ Â«¡ „‰  ‘›Ì— «·„·›" & vbCrLf & _
    "«–« «—œ  «·ﬁÌ«„ »–«·ﬂ ›÷€ÿ ⁄·Ï “— ‰⁄„ ø", vbQuestion + vbMsgBoxRight + vbYesNo)
   If t = vbYes Then Call Command4_Click
   End If
        If List1.ListIndex = -1 Then MsgBox "Ì—ÃÏ «Œ Ì«— „·› „‰ «·ﬁ«∆„…", vbCritical + vbMsgBoxRight, "Œÿ√": Exit Sub
        Dim nbyte As Byte
        Dim i As Long
        statePut.Caption = "Ã«—Ì ‰”Œ «·„·› «·„ƒﬁ "
        '  MsgBox "C:\Program Files\MHSC\" & Mid(List1.List(List1.ListIndex), 9, Len(List1.List(List1.ListIndex)) - 8)
        
        Open "C:\Program Files\MHSC\" & Mid(List1.List(List1.ListIndex), 9, Len(List1.List(List1.ListIndex)) - 12) & ".mhsc" For Binary As #2
            Open "C:\Program Files\MHSC\" & Mid(List1.List(List1.ListIndex), 9, Len(List1.List(List1.ListIndex)) - 8) For Binary As #3
                ProgressBar1.Max = LOF(3)
                For i = 1 To LOF(3)
                    ProgressBar1 = i
                    nCounter = nCounter + 1: If nCounter = Len(sPasword) + 1 Then nCounter = 1
                    Get #3, , nbyte
                    DoEvents
                    Put #2, , ccbyte(nbyte, nCounter)
                    statePut.Caption = " %" & Int((100 / LOF(3)) * i) & "  ﬁœ„ ⁄„·Ì… «· ‘›Ì—  "
                    '  PutccByte (ccbyte(nbyte, nCounter))
                Next i
                Dim sType As String * 3
                sType = Right(List1.List(List1.ListIndex), 3)
                ' Seek #2, LOF(2)
                Put #2, , sType
                Reset
                statePut.Caption = "Ì „ «·«‰ „”Õ «·„·› «·„ƒﬁ  "
                Kill "C:\Program Files\MHSC\" & Mid(List1.List(List1.ListIndex), 9, Len(List1.List(List1.ListIndex)) - 8)
                List1.RemoveItem List1.ListIndex
                sPath = ""
                statePut.Caption = " „...."
                SysTray1.TrayType = Balloon
                SysTray1.TrayTitle = " „  ‘›Ì— «·»Ì«‰«  »‰Ã«Õ"
                Close
                MsgBox " „  ‘›Ì— «·»Ì«‰«  »‰Ã«Õ  Ì„ﬂ‰ﬂ «·«‰ „”Õ „·›« ﬂ «·–Ì «÷› Â« ", vbInformation + vbMsgBoxRight, "—”«·…"
                ProgressBar1.Value = 0
                nCounter = 0
            Else
                MsgBox "Ì—ÃÏ «÷«›… „·› «·Ï «·ﬁ«∆„… ", vbMsgBoxRight + vbCritical
            End If
            StartProc = False
            Exit Sub
Chfr:
MsgBox "ÕœÀ Œÿ√ «À‰«¡ ⁄„·Ì… «· ‘›Ì— ", vbMsgBoxRight + vbCritical
            statePut.Caption = "›‘· «·⁄„·Ì…"
            Kill "C:\Program Files\MHSC\" & Mid(List1.List(List1.ListIndex), 9, Len(List1.List(List1.ListIndex)) - 8)
StartProc = False
End Sub

Private Sub Command3_Click()
    CommonDialog1.FileName = "C:\Program Files\MHSC\*.mhsc"
    CommonDialog1.Filter = "*.MHSC"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileTitle <> "" Then
        If Form3.Check1.Value = vbChecked Then
            Call Out_Byte
        Else
            Form3.Visible = True
        End If
    Else
        MsgBox "Ì—ÃÏ «Œ Ì«— „·› «·»Ì«‰«  «·„‘›—  " & vbCrLf & "C:\Program Files\MHSC\", vbMsgBoxRight + vbExclamation, " ‰»ÌÂ"
    End If
    Exit Sub
End Sub

Private Sub Command4_Click()
SysTray1.AddTrayIcon
Me.Hide
Form3.Visible = False
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then MsgBox " „  ‘€Ì· «·»—‰«„Ã »«·›⁄·", vbCritical + vbMsgBoxRight: End
    Me.Hide
    StartProG = True
    StartProc = False
    Open App.Path & "\option.ini" For Random As #7 Len = 15
        Get #7, 1, Tmdata
        If Mid(Trim(Tmdata), 1, 2) <> "Ms" Then
            FrmOption.Visible = True
        Else
            frmSplash.Show
            Call SytemFormat
        End If
     Set SysTray1 = Form1.Controls.Add("SysTray_Control.SysTray", "SysTray1")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (X / Screen.TwipsPerPixelX) = &H205 And Form1.Visible = False Then
PopupMenu nmenu1
ElseIf (X / Screen.TwipsPerPixelX) = &H203 Then
SysTray1.DeleteTrayIcon

Me.Show
End If

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
Static iwidth, iheight, b
If b Then
Me.Width = iwidth
Me.Height = iheight
Else
iheight = Me.Height
 iwidth = Me.Width
 b = 1
 End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Form1
    Unload Form2
    Unload Form3
    Unload frmSplash
    Unload FrmOption
    Unload aboutProg
End
End Sub

Private Sub Label1_Click()
    '   sv = "4F6B61736861"
    Me.Hide
    Form2.Visible = True
End Sub

Private Sub Label3_Click()
    aboutProg.Show
End Sub

Private Sub Label4_Click()
If StartProc = False Then
    FrmOption.Show
    End If
End Sub

Private Sub naboutprog_Click()
aboutProg.Show
End Sub

Private Sub nexit_Click()
Unload Form1
Unload Form2
    Unload Form3
    Unload frmSplash
    Unload FrmOption
 Unload aboutProg
End
End Sub

Private Sub nshowProg_Click()
Form1.Show
End Sub

Private Sub SysTray2_DblClick(Button As MouseButtonConstants)
Unload Form2
    Unload Form3
    Unload frmSplash
    Unload FrmOption
End
End Sub

Private Sub Timer1_Timer()
    m = m + 17
    If m >= 170 Then m = 0
    Picture1.PaintPicture Picture2.Picture, 0, 0, Picture1.Width, 17, 0, m, Picture1.Width, 17

End Sub


