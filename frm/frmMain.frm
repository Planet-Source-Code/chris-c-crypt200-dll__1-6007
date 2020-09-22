VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CC1E317A-3102-11D1-816E-00A024E95548}#4.0#0"; "ACTIVEPR.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Security Beta 1 (Build 14012000)"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkOption 
      Caption         =   "Show Value?"
      Height          =   225
      Left            =   1260
      TabIndex        =   23
      Top             =   2580
      Width           =   1305
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   60
      Top             =   4230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ActiveProgressBar.AXProgress ProgressBar 
      Height          =   315
      Left            =   60
      Top             =   3660
      Visible         =   0   'False
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   556
      ForeColor       =   -2147483639
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":0ECA
      FillStyle       =   1
      Max             =   100
      CaptionStyle    =   1
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   345
      Index           =   2
      Left            =   3060
      TabIndex        =   11
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   "&Exit"
      Height          =   345
      Index           =   5
      Left            =   3060
      TabIndex        =   14
      Top             =   3240
      Width           =   1000
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   "&About..."
      Height          =   345
      Index           =   4
      Left            =   2040
      TabIndex        =   13
      Top             =   3240
      Width           =   1000
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   "&Help"
      Height          =   345
      Index           =   3
      Left            =   1020
      TabIndex        =   12
      Top             =   3240
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1845
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   3254
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "File"
      TabPicture(0)   =   "frmMain.frx":0EE6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDescrip(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDescrip(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDescrip(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtMyFile(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtMyFile(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdFile(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmdFile(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Word/Text"
      TabPicture(1)   =   "frmMain.frx":0F02
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDescrip(1)"
      Tab(1).Control(1)=   "lblDescrip(0)"
      Tab(1).Control(2)=   "lblDescrip(4)"
      Tab(1).Control(3)=   "txtMydata(1)"
      Tab(1).Control(4)=   "txtMydata(0)"
      Tab(1).ControlCount=   5
      Begin VB.CommandButton CmdFile 
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   3
         Top             =   1320
         Width           =   285
      End
      Begin VB.CommandButton CmdFile 
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   3720
         TabIndex        =   2
         Top             =   990
         Width           =   285
      End
      Begin VB.TextBox txtMyFile 
         Height          =   285
         Index           =   1
         Left            =   1215
         TabIndex        =   1
         Top             =   1320
         Width           =   2475
      End
      Begin VB.TextBox txtMyFile 
         Height          =   285
         Index           =   0
         Left            =   1215
         TabIndex        =   0
         Top             =   990
         Width           =   2475
      End
      Begin VB.TextBox txtMydata 
         Height          =   285
         Index           =   0
         Left            =   -73785
         MaxLength       =   255
         TabIndex        =   5
         Top             =   990
         Width           =   2790
      End
      Begin VB.TextBox txtMydata 
         Height          =   285
         Index           =   1
         Left            =   -73785
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   2790
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Please select the source and output file that you wish to encrypt ot decrypt inside the text box below:"
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   450
         Width           =   3945
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Output File:"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   16
         Top             =   1380
         Width           =   810
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Source File:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   15
         Top             =   1050
         Width           =   840
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Please enter the text that you wish to encrypt ot decrypt inside the Source String text box below:"
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   4
         Left            =   -74880
         TabIndex        =   21
         Top             =   450
         Width           =   3945
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Source String:"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   17
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Output String:"
         Height          =   195
         Index           =   1
         Left            =   -74820
         TabIndex        =   18
         Top             =   1380
         Width           =   975
      End
   End
   Begin VB.TextBox txtMydata 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1245
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   2250
      Width           =   2790
   End
   Begin VB.TextBox txtMydata 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1245
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1920
      Width           =   2790
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   "&Decryption"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   2040
      TabIndex        =   10
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   "E&ncryption"
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   1020
      TabIndex        =   9
      Top             =   2880
      Width           =   1000
   End
   Begin VB.Label lblDescrip 
      AutoSize        =   -1  'True
      Caption         =   "Ref. Key:"
      Height          =   195
      Index           =   3
      Left            =   210
      TabIndex        =   20
      Top             =   2310
      Width           =   660
   End
   Begin VB.Label lblDescrip 
      AutoSize        =   -1  'True
      Caption         =   "Seed Value:"
      Height          =   195
      Index           =   2
      Left            =   210
      TabIndex        =   19
      Top             =   1980
      Width           =   870
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents ObjData As clscrypt
Attribute ObjData.VB_VarHelpID = -1
Dim GoExit%
Dim SourceFileName$
Dim OutputFileName$

Dim nReturn&
Dim ReturnStr$
Dim WinTempPath$
Dim TempFile$
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Sub ChkOption_Click()
If ChkOption.Value = 1 Then
    txtMydata(2).PasswordChar = ""
    txtMydata(3).PasswordChar = ""
Else
    txtMydata(2).PasswordChar = "*"
    txtMydata(3).PasswordChar = "*"
End If
End Sub


Private Sub CmdAction_Click(Index As Integer)
Dim DataStr$

'Check for Max Seed Value:
If Len(txtMydata(2).Text) <> 0 Then
    If CLng(txtMydata(2).Text) > 32676 Then
        MsgBox "The seed value can not larger than 32,676.", vbExclamation + vbOKOnly, "Data Security"
        txtMydata(2).SelStart = Len(txtMydata(2).Text)
        txtMydata(2).SetFocus
        Exit Sub
    End If
End If

Select Case Index
Case 0 'Encrypt
    CmdAction(0).Enabled = False
    CmdAction(1).Enabled = False
    CmdAction(2).Enabled = True
    CmdAction(5).Enabled = False
    ChkOption.Enabled = False
    ProgressBar.Value = 0
    ProgressBar.Visible = True
    Select Case SSTab1.Tab
    Case 0 'File
        If ObjData.Get_File(SourceFileName$, TempFile$, CInt(txtMydata(2).Text), txtMydata(3).Text, 0) = 1 Then
            MsgBox "Data ecryption error.", vbExclamation + vbOKOnly, "Data Security"
            Kill TempFile$
        Else
            If ObjData.CancelFlag = 0 Then
                MsgBox "Data ecryption completed.", vbInformation + vbOKOnly, "Data Security"
                FileCopy TempFile$, OutputFileName$
                Kill TempFile$
            Else 'User cancel process.
                MsgBox "Process abort.", vbInformation + vbOKOnly, "Data Security"
            End If
        End If
    Case 1 'Word/Text
        DataStr$ = ObjData.Get_Data(txtMydata(0).Text, CInt(txtMydata(2).Text), txtMydata(3).Text, 0)
        txtMydata(1).Text = DataStr$
    End Select
    
    CmdAction(0).Enabled = True
    CmdAction(1).Enabled = True
    CmdAction(2).Enabled = False
    CmdAction(5).Enabled = True
    ChkOption.Enabled = True
    ProgressBar.Visible = False
Case 1 'Decrypt
    CmdAction(0).Enabled = False
    CmdAction(1).Enabled = False
    CmdAction(2).Enabled = True
    CmdAction(5).Enabled = False
    ChkOption.Enabled = False
    ProgressBar.Value = 0
    ProgressBar.Visible = True
    Select Case SSTab1.Tab
    Case 0 'File
        If ObjData.Get_File(SourceFileName$, TempFile$, CInt(txtMydata(2).Text), txtMydata(3).Text, 1) = 1 Then
            MsgBox "Data dencryption error.", vbExclamation + vbOKOnly, "Data Security"
            Kill TempFile$
        Else
            If ObjData.CancelFlag = 0 Then
                MsgBox "Data dencryption completed.", vbInformation + vbOKOnly, "Data Security"
                FileCopy TempFile$, OutputFileName$
                Kill TempFile$
            Else 'User cancel process.
                MsgBox "Process abort.", vbInformation + vbOKOnly, "Data Security"
            End If
        End If
    Case 1 'Word/Text
        DataStr$ = ObjData.Get_Data(txtMydata(0).Text, CInt(txtMydata(2).Text), txtMydata(3).Text, 1)
        txtMydata(1).Text = DataStr$
    End Select
    
    CmdAction(0).Enabled = True
    CmdAction(1).Enabled = True
    CmdAction(2).Enabled = False
    CmdAction(5).Enabled = True
    ChkOption.Enabled = True
    ProgressBar.Visible = False
Case 2 'Cancel
    ObjData.CancelFlag = 1
Case 3 'Help
    If Dir(App.Path & "\crypt2.hlp") = "" Then
        MsgBox "File not found. " & App.Path & "\crypt2.hlp", vbExclamation + vbOKOnly, "Data Security"
    Else
        CDlg.HelpFile = App.Path & "\crypt2.hlp"
        CDlg.HelpCommand = cdlHelpForceFile
        CDlg.ShowHelp
    End If
Case 4 'About
    frmAbout.Show vbModal
Case 5 'Exit
    GoExit% = 1
    Unload Me
End Select


End Sub


Private Sub CmdFile_Click(Index As Integer)
Dim MyFileName$, Ret
On Error Resume Next

If Index = 0 Then CDlg.DialogTitle = "Open Source File" 'Source
If Index = 1 Then CDlg.DialogTitle = "Save As" 'Output
    
'CDlg.Filter = "Text Files (*.txt)|*.txt|Document Files (*.doc)|*.doc|Rich Text Files (*.rtf)|*.rtf|Application Files (*.exe)|*.exe|All Files (*.*)|*.*"
CDlg.Filter = "Text Files (*.txt)|*.txt|Record Files (*.log)|*.log|All Files (*.*)|*.*"
    Do
        CDlg.CancelError = True
        CDlg.FileName = ""
        CDlg.ShowOpen
        If Err = cdlCancel Then Exit Sub
        MyFileName$ = CDlg.FileName
    
        ' If the file doesn't exist, go back.
        Ret = Len(Dir$(MyFileName$))
        If Err Then
            MsgBox Error$, 48, "Data Security"
            Exit Sub
        End If
        If Ret Then
            If Index = 0 Then
                Exit Do
            Else
                If MsgBox("Do you want to overwrite the " & MyFileName & "?", vbQuestion + vbYesNo, "Data Security") = vbYes Then
                    Exit Do
                End If
            End If
        
        Else
            If Index = 0 Then
                MsgBox MyFileName$ + " not found!", 48
            Else
                If MsgBox("Do you want to create " & MyFileName$ & "?", vbQuestion + vbYesNo, "Data Security") = vbYes Then
                    Exit Do
                End If
            End If
        End If
    Loop

txtMyFile(Index).Text = MyFileName$
txtMyFile(Index).SelStart = Len(MyFileName$)
If Index = 0 Then SourceFileName$ = MyFileName$
If Index = 1 Then OutputFileName$ = MyFileName$

End Sub


Private Sub Form_Load()
Set ObjData = New clscrypt

'Get Windows Temp path...
ReturnStr$ = String(255, " ")
nReturn& = GetTempPath(Len(ReturnStr$), ReturnStr$)
If nReturn& = 0 Then
    MsgBox "System Error."
    End
Else
    WinTempPath$ = Left(ReturnStr$, InStr(1, ReturnStr$, Chr(0), vbBinaryCompare) - 1)
    TempFile$ = WinTempPath$ & "\cry00036.dat"
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GoExit% = 1 Then
    Set ObjData = Nothing
    Cancel = 0
    End
Else
    Cancel = 1
End If
End Sub


Private Sub ObjData_PercentDone(ByVal Percent As Integer, ByVal TotalFileSize As Long, ProcessFileSize As Long)
ProgressBar.Value = Percent
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
Case 0 'File
    If Len(txtMyFile(0).Text) <> 0 And Len(txtMyFile(1).Text) And Len(txtMydata(2).Text) <> 0 And Len(txtMydata(3).Text) <> 0 Then
        CmdAction(0).Enabled = True
        CmdAction(1).Enabled = True
    Else
        CmdAction(0).Enabled = False
        CmdAction(1).Enabled = False
    End If
Case 1 'Word/Text
    If Len(txtMydata(0).Text) <> 0 And Len(txtMydata(2).Text) <> 0 And Len(txtMydata(3).Text) <> 0 Then
        CmdAction(0).Enabled = True
        CmdAction(1).Enabled = True
    Else
        CmdAction(0).Enabled = False
        CmdAction(1).Enabled = False
    End If
End Select

End Sub

Private Sub txtMyData_Change(Index As Integer)
If SSTab1.Tab = 1 Then 'Word/Text
    If Len(txtMydata(2).Text) <> 0 And Len(txtMydata(3).Text) <> 0 Then
        CmdAction(0).Enabled = True
        CmdAction(1).Enabled = True
    Else
        CmdAction(0).Enabled = False
        CmdAction(1).Enabled = False
    End If
Else
    'File
    If Len(txtMyFile(0).Text) <> 0 And Len(txtMyFile(1).Text) And Len(txtMydata(2).Text) <> 0 And Len(txtMydata(3).Text) <> 0 Then
        CmdAction(0).Enabled = True
        CmdAction(1).Enabled = True
    Else
        CmdAction(0).Enabled = False
        CmdAction(1).Enabled = False
    End If
End If
End Sub

Private Sub txtMyData_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 2 Then If (KeyAscii < 48 Or KeyAscii > 59) And KeyAscii <> 8 Then KeyAscii = 0
End Sub


Private Sub txtMyFile_Change(Index As Integer)
'File
If Len(txtMyFile(0).Text) <> 0 And Len(txtMyFile(1).Text) And Len(txtMydata(2).Text) <> 0 And Len(txtMydata(3).Text) <> 0 Then
    CmdAction(0).Enabled = True
    CmdAction(1).Enabled = True
Else
    CmdAction(0).Enabled = False
    CmdAction(1).Enabled = False
End If
End Sub




Private Sub txtMyFile_LostFocus(Index As Integer)
Dim Ret
Select Case Index
Case 0 'Source
    If Len(txtMyFile(0).Text) <> 0 Then
        Ret = Len(Dir$(txtMyFile(Index).Text))
        If Not Ret Then
            MsgBox "File not found. " & txtMyFile(0).Text & ".", vbExclamation + vbOKOnly, "Data Security"
            txtMyFile(0).SelStart = Len(txtMyFile(0).Text)
            txtMyFile(0).SetFocus
        End If
    End If
Case 1 'Output
    If Len(txtMyFile(1).Text) <> 0 Then
        Ret = Len(Dir$(txtMyFile(Index)))
        If Ret Then
            If MsgBox("Do you want to overwrite the " & txtMyFile(1).Text & "?", vbQuestion + vbYesNo, "Data Security") = vbNo Then
                txtMyFile(1).SelStart = Len(txtMyFile(1).Text)
                txtMyFile(1).SetFocus
            End If
        Else
            If MsgBox("Do you want to create " & txtMyFile(1).Text & "?", vbQuestion + vbYesNo, "Data Security") = vbNo Then
                txtMyFile(1).SelStart = Len(txtMyFile(1).Text)
                txtMyFile(1).SetFocus
            End If
        End If
    End If
End Select
End Sub


