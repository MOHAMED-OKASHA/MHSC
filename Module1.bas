Attribute VB_Name = "Module1"
 ' Option Explicit   �� ���� ������� ��� ����� ���� ����� �� �������� ������� ����� �� �������
    Dim Temp, Temp1 As Single
    Dim Bit(8) As Byte
    Dim Cun As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Public sPasword As String
    Dim s, d, byteTemp As String
    Public nCounter As Byte
   
    Public StartProG As Boolean
Public Function ccbyte(ByVal sbyte As Byte, ByVal nCounter As Byte) As Byte
    s = sbyte
    d = ""
    For i = 1 To 8
        Bit(i) = Sgn(Fix(s) / 2 - Fix(Fix(s) / 2))
        s = Fix(s / 2)
    Next i
    s = ""
    For i = 8 To 1 Step -1
        s = s & Bit(i)
    Next i
    '********************************************************************   ��� ��� ��� *********************
    Call ButccByte(s, Mid(sPasword, nCounter, 1), "L")
    For i = 1 To 8
        If Mid(s, i, 1) = 1 Then Bit(i) = 0 Else Bit(i) = 1
        d = d & Bit(i)
    Next i
    '    byteTemp = str(Bit(8) * 2 ^ 7 + Bit(7) * 2 ^ 6 + Bit(6) * 2 ^ 5 + Bit(5) * 2 ^ 4 + Bit(4) * 2 ^ 3 + Bit(3) * 2 ^ 2 + Bit(2) * 2 ^ 1 + Bit(1) * 2 ^ 0)
    Bit(1) = Mid(d, 1, 1)
    Bit(2) = Mid(d, 2, 1)
    Bit(3) = Mid(d, 3, 1)
    Bit(4) = Mid(d, 4, 1)
    Bit(5) = Mid(d, 5, 1)
    Bit(6) = Mid(d, 6, 1)
    Bit(7) = Mid(d, 7, 1)
    Bit(8) = Mid(d, 8, 1)
    ccbyte = Bit(8) * 2 ^ 0 + Bit(7) * 2 ^ 1 + Bit(6) * 2 ^ 2 + Bit(5) * 2 ^ 3 + Bit(4) * 2 ^ 4 + Bit(3) * 2 ^ 5 + Bit(2) * 2 ^ 6 + Bit(1) * 2 ^ 7
End Function

Public Sub ButccByte(ByRef dd As Variant, ByVal pssW As Byte, ByVal t As String)
    Select Case t
    Case "L"
        For j = pssW To 1 Step -1
            dd = Mid(dd, Len(dd), 1) & Mid(dd, 1, Len(dd) - 1)
        Next
    Case "R"
        For j = 1 To pssW
            dd = Mid(dd, 2, Len(dd)) & Mid(dd, 1, 1)
        Next
    End Select
End Sub

Public Function ucbyte(ByVal sbyte As Byte, ByVal nCounter As Byte) As Byte
    s = sbyte
    d = ""
    For i = 1 To 8
        Bit(i) = Sgn(Fix(s) / 2 - Fix(Fix(s) / 2))
        s = Fix(s / 2)
    Next i
    s = ""
    For i = 8 To 1 Step -1
        s = s & Bit(i)
    Next i
    Call ButccByte(s, Mid(sPasword, nCounter, 1), "R")
    For i = 1 To 8
        If Mid(s, i, 1) Then Bit(i) = 0 Else Bit(i) = 1
        d = d & Bit(i)
    Next i
    '    byteTemp = str(Bit(8) * 2 ^ 7 + Bit(7) * 2 ^ 6 + Bit(6) * 2 ^ 5 + Bit(5) * 2 ^ 4 + Bit(4) * 2 ^ 3 + Bit(3) * 2 ^ 2 + Bit(2) * 2 ^ 1 + Bit(1) * 2 ^ 0)
    Bit(1) = Mid(d, 1, 1)
    Bit(2) = Mid(d, 2, 1)
    Bit(3) = Mid(d, 3, 1)
    Bit(4) = Mid(d, 4, 1)
    Bit(5) = Mid(d, 5, 1)
    Bit(6) = Mid(d, 6, 1)
    Bit(7) = Mid(d, 7, 1)
    Bit(8) = Mid(d, 8, 1)
    ucbyte = Bit(8) * 2 ^ 0 + Bit(7) * 2 ^ 1 + Bit(6) * 2 ^ 2 + Bit(5) * 2 ^ 3 + Bit(4) * 2 ^ 4 + Bit(3) * 2 ^ 5 + Bit(2) * 2 ^ 6 + Bit(1) * 2 ^ 7
End Function
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '123632642367896285242687
    'Public Sub PutccByte(ByVal ccbyte, ByVal NumperMotherboard As String)
    'Open "C:\Program Files\MHSC\" & Form1.List1.List(k) For Binary As #1 Len = 3
    'Close
    'End Sub

Public Sub SytemFormat()
    
    'On Error goto errSysFrmt
    '6 ��� ��� ����� ������� ������ �������� ����
    '     ���� ������ �� �������� ��� ���� ������  ���� ���� ������ ��������� �������
    If Dir$("C:\Program Files\MHSC", 16) = "" Then MkDir "C:\Program Files\MHSC"
    If Dir$("C:\Program Files\MHSC\Language", 16) = "" Then MkDir "C:\Program Files\MHSC\Language"
    If Dir$("C:\Program Files\MHSC\PlA", 16) = "" Then MkDir "C:\Program Files\MHSC\PlA"
    If Dir$("C:\Program Files\MHSC\PlB", 16) = "" Then MkDir "C:\Program Files\MHSC\PlB"
        
       
        '  ��� ������� ����� ������� � ����� ������ �� ��� �������
        Dim Tmp As String ' ��� ������� ���� ��� �� ������ �� ������ �� ���� ����� ����� ������� ���
        ''''''''''''''BUG''''''''''''''''''''''''''''''''  ��� ������ �������  ���� �� ����� ����� ��� ���� ���� �� ���� ���� IF
        Dim numFile  As Long
        numFile = FreeFile
        Tmp = Dir$("C:\WINDOWS\system32\R99x5.ocx")
        If Tmp = "" Then
        Dim a() As Byte ' ��� �������� ���� ����� �������� ��������� ������� LoadResData �� ��� ������� ����� �� ��� ���� �� ������ ������� ���
            Open "C:\WINDOWS\system32\R99x5.ocx" For Binary Access Write As #numFile
                
                a = LoadResData(101, "CUSTOM")
                DoEvents
                Put #numFile, , a
            End If
            numFile = FreeFile
            Tmp = Dir$("C:\WINDOWS\system32\SkinFramework.ocx")
            If Tmp = "" Then
            Dim b() As Byte
                Open "C:\WINDOWS\system32\SkinFramework.ocx" For Binary Access Write As #numFile
                   
                    b = LoadResData(102, "CUSTOM")
                    DoEvents
                    Put #numFile, , b
                    numFile = FreeFile
                    Dim c() As Byte
                    Open "C:\WINDOWS\system32\WinXP.Luna.cjstyles" For Binary Access Write As #numFile
                    
                        c = LoadResData(103, "CUSTOM")
                        DoEvents
                        Put #numFile, , c
                    End If
                
                numFile = FreeFile
                Tmp = Dir$("C:\WINDOWS\system32\SysTray.ocx")
                If Tmp = "" Then
                Dim d() As Byte
                    Open "C:\WINDOWS\system32\SysTray.ocx" For Binary Access Write As #numFile
                        DoEvents
                        d = LoadResData(104, "CUSTOM")
                        Put #numFile, , d
                    End If

                    Reset ' ����� ���� ������� ��������
                    '''''''''''''''''''''''''''''''''''''   ����� ����� ������� ������� ����� '''''''''''''''''''''''''
                   If StartProG = True Then
                    Set skin_Xp = Form1.Controls.Add("Codejock.SkinFramework.10.4.2", "skin_Xp")
                    Set label_URl = aboutProg.Controls.Add("R99x5.R99x5_URL", "label_URl")
                    
                    
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''' ������� ���� ��������
                    label_URl.Left = aboutProg.Label2.Left
                    label_URl.Top = aboutProg.Label2.Top
                    label_URl.Visible = True
                    
                 
                 '   ��� ���� ����� ��� ��� ��� ���� ����� ��� ������  ������ ������� ���� �� ����� ����� ����� ���� �� ��� �������
                  ' ���� �� ���� ������ ���� �� ���� �� ����� ����� ������� ����� ����
                  '  label_URl.UrlText = "���� �� �� ������� ����� ������ ���� �����"
                  '  label_URl.InactiveMouse = vbGreen
                  '  label_URl.BorderColor = &H444444
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   ���� ������� ��� Xp ��������
                    skin_Xp.LoadSkin "C:\WINDOWS\system32\WinXP.Luna.cjstyles", ""
                    skin_Xp.ApplyWindow Form1.hwnd
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                 '   ����� ���� ������ ����� ��� ����  ��
                    Form1.Label1.MousePointer = vbCustom
                       Form1.Label3.MousePointer = vbCustom
                       Form1.Label4.MousePointer = vbCustom
                       aboutProg.Image1.MousePointer = vbCustom
                       aboutProg.Label1.MousePointer = vbCustom
                       Form1.Label1.MouseIcon = LoadResPicture(101, vbResCursor)
                       Form1.Label3.MouseIcon = LoadResPicture(101, vbResCursor)
                       Form1.Label4.MouseIcon = LoadResPicture(101, vbResCursor)
                       aboutProg.Label1.MouseIcon = LoadResPicture(101, vbResCursor)
                       aboutProg.Image1.MouseIcon = LoadResPicture(101, vbResCursor)
                    StartProG = False
                    
                    
                    
                    End If
                    Open App.Path & "\option.ini" For Random As #1 Len = 15
                        Dim Tmdata1 As String * 15
                        Get #1, 1, Tmdata1
                        'If Trim(Tmdata) = "Arabia" Then FrmOption.Combo1.ListIndex = 0
                        'If Trim(Tmdata) = "English" ThenFrmOption.Combo1.ListIndex = 0   �����  �� ��� ������ ����� �� ���� ���� ����� ���� ������ ����� �������� ������� ������� ������� ��� ���
                        FrmOption.Combo1.ListIndex = Val(Mid(Trim(Tmdata1), 3, 1)) - 1
                        If FrmOption.Combo1.ListIndex = 0 Then
                            '
                        End If
                        Get #1, 2, Tmdata1
                        '   ��� ����� ���� ������� ������ ������ Mode1 �� Mode2  ����� ������� �������� �������
                        ' ��� �� ��� ����� 1 �� ��� ��� �� ������ �� ��� ����� ����� ����� ������ �� ������ �������� ������� ���� �� ���� ���� ��� ��� ������ ��� ��� ���� ��� ���� ����� ������� ��� �� ����
                        '����� �� ���� �� ����� ����� ���� ������ ���� ���� ���� �� ������ ���� ��� �� ���� ������ ����� ���� ������ ������ ����� ���� ������ ������� ������ ���� �����
                        FrmOption.Option2 = Right(Trim(Tmdata1), 1) - 1
                        Select Case FrmOption.Combo1.ListIndex
                            Dim strNameLang As String
                        Case "0"
                            strNameLang = "Arabia"
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''#
                            Dim Ctrl As Control     ' ������ ����� ������ ��� ��� ������� ��� �� �� ����� �� ������ ��� ������  ' #
                            On Error Resume Next
                            For Each Ctrl In Form1.Controls
                                Ctrl.Left = Ctrl.Container.ScaleWidth - Ctrl.Left - Ctrl.Width
                                If Ctrl.Alignment = 1 Then
                                    Ctrl.Alignment = 0
                                ElseIf Ctrl.Alignment = 0 Then
                                    Ctrl.Alignment = 1
                                End If
                                Ctrl.RightToLeft = True
                            Next
                            Form1.RightToLeft = True
                            Err.Clear
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Case "1"
                            strNameLang = "English"
                        Case Else
                            strNameLang = "Arabia"
                            MsgBox "�� ������ ��� �������� ����� ���������� ��� �������", vbInformation + vbMsgBoxRight, "�����"
                        End Select
                        Get #1, 3, Tmdata1
                        Select Case Mid(Tmdata1, 3, 1)
                        Case "C"
                            Form3.Check1.Value = vbChecked
                        Case "U"
                            Form3.Check1.Value = vbUnchecked
                        End Select
                        Get #1, 4, Tmdata1
                        Select Case Mid(Tmdata1, 3, 1)
                        Case "C"
                            FrmOption.Check1.Value = vbChecked
                        Case "U"
                            FrmOption.Check1.Value = vbUnchecked
                        End Select
                        Reset
                        Dim X As Single
                        X = Timer
                        Do
                            DoEvents
                        Loop Until Timer > X + 1
                        
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        
                        
                        
                       
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        frmSplash.Hide
                        Form1.Visible = True
                        
 
                        Exit Sub
                        
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' ����� �����
End Sub

Public Sub Out_Byte()
    Form3.CommonDialog1.ShowSave
    If Form1.CommonDialog1.FileName <> "" Then
        Dim i As Long
        Dim nbyte As Byte
        Dim sType As String * 3
        Close
        Open Form1.CommonDialog1.FileName For Binary As #3
            Get #3, LOF(3) - 2, sType
            Seek #3, 1
            Open Form3.CommonDialog1.FileName & "." & Trim(sType) For Binary As #4
                If Form3.Check1.Value = vbUnchecked Then
                    Form3.ProgressBar1.Max = LOF(3) - 3
                Else
                    Form1.ProgressBar1.Max = LOF(3) - 3
                End If
                For i = 1 To LOF(3) - 3
                    If Form3.Check1.Value = vbUnchecked Then Form3.ProgressBar1.Value = i Else Form1.ProgressBar1.Value = i
                    nCounter = nCounter + 1: If nCounter = Len(sPasword) + 1 Then nCounter = 1
                    Get #3, , nbyte
                    DoEvents
                    Put #4, , ucbyte(nbyte, nCounter)
                    '  getucByte (ucbyte(nbyte, nCounter))
                Next i
                MsgBox "�� �� ����� ����� ������ ������ ����� �� ������ �������", vbInformation + vbMsgBoxRight, "�����"
               Close
                Form3.ProgressBar1.Value = 0
            End If
            nCounter = 0

End Sub
