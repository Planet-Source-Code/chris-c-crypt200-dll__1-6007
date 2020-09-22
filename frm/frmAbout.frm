VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Data Security..."
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdAction 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   1380
      TabIndex        =   2
      Top             =   1260
      Width           =   1000
   End
   Begin VB.Image imgPointer 
      Height          =   480
      Left            =   1920
      Picture         =   "frmAbout.frx":08CA
      Top             =   2100
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "ccthou@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   900
      Width           =   1440
   End
   Begin VB.Label lblAbout 
      Height          =   795
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   3015
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const email = "ccthou@yahoo.com"

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Sub SendEmail()
Dim Success As Long

Success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

Private Sub CmdAction_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblAbout(0).Caption = "Data Security Beta 1 (Build 14012000)" & vbCrLf & _
                      "Written By: Chris. C" & vbCrLf & vbCrLf & _
                      "To  contact me, please send email to:"


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault

End Sub


Private Sub lblAbout_Click(Index As Integer)
If Index = 1 Then SendEmail
End Sub

Private Sub lblAbout_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > 0 And X < 1441 And Y > 0 And Y < 196 And Index = 1 Then
    Me.MousePointer = vbCustom
    Me.MouseIcon = imgPointer.Picture
Else
    Me.MousePointer = vbDefault
End If
End Sub


