VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " E-Mail Sender - Idle"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrIdle 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3360
      Top             =   2040
   End
   Begin MSWinsockLib.Winsock tcpSmtp 
      Left            =   3960
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtTo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "To"
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox txtFrom 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "From"
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "Subject"
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      Height          =   1725
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmMain.frx":0442
      Top             =   1200
      Width           =   3615
   End
   Begin ComctlLib.ProgressBar pbLoading 
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label cmdAbout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label cmdSettings 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Settings"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   765
   End
   Begin VB.Label cmdExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   765
   End
   Begin VB.Label lblSubject 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblMessage 
      Caption         =   "Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblFrom 
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblTo 
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label cmdSend 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Send"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   765
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Strings And Integers
Private i250 As Integer
Public MailServer As String

Private Sub cmdAbout_Click()
MsgBox "Thanks for using this program made by ziller." & vbCrLf & "Please check my homepage for more VB stuff." & vbCrLf & "http://www.ziller.tk", vbInformation, "About Mail Sender"
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSend_Click()
If InStr(MailServer, ".") <= 0 Then cmdSettings_Click: Exit Sub
tcpSmtp.Close
tcpSmtp.Connect MailServer, 25
Me.Caption = "E-Mail Sender - Connecting.."
pbLoading.Visible = True
txtMessage.Height = 1395
pbLoading.Value = 10
End Sub

Private Sub cmdSettings_Click()
frmSmtpServer.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmSmtpServer
End Sub

Private Sub tcpSmtp_Close()
pbLoading.Visible = False
txtMessage.Height = 1365
pbLoading.Value = 0
Me.Caption = "E-Mail Sender - Idle"
End Sub

Private Sub tcpSmtp_DataArrival(ByVal bytesTotal As Long)
Dim InData As String
tcpSmtp.GetData InData
If Left(InData, 3) = "220" Then
  i250 = 1
  tcpSmtp.SendData "HELO " & tcpSmtp.LocalHostName & vbCrLf
  pbLoading.Value = pbLoading.Value + 20
  Me.Caption = "E-Mail Sender - Sending..."
End If
If Left(InData, 3) = "250" Then
  pbLoading.Value = pbLoading.Value + 10
  If i250 = 1 Then i250 = 2: tcpSmtp.SendData "MAIL FROM: " & txtFrom & vbCrLf: Exit Sub
  If i250 = 2 Then i250 = 3: tcpSmtp.SendData "RCPT TO: " & txtTo & vbCrLf: Exit Sub
  If i250 = 3 Then i250 = 4: tcpSmtp.SendData "DATA" & vbCrLf: Exit Sub
  If i250 = 4 Then
    tcpSmtp.SendData "QUIT" & vbCrLf
    Me.Caption = "E-Mail Sender - Sendt..."
    tmrIdle.Enabled = True
    tcpSmtp.Close: pbLoading.Value = pbLoading.Value + 10
    MsgBox "Mail sending is complete.", vbInformation, "Done"
    pbLoading.Visible = False
    txtMessage.Height = 1725
    Exit Sub
  End If
End If
If Left(InData, 3) = "354" Then
  tcpSmtp.SendData "FROM: " & txtFrom & vbCrLf
  tcpSmtp.SendData "TO: " & txtTo & vbCrLf
  tcpSmtp.SendData "SUBJECT: " & txtSubject & vbCrLf
  tcpSmtp.SendData "" & vbCrLf
  tcpSmtp.SendData txtMessage & vbCrLf
  tcpSmtp.SendData vbCrLf & "." & vbCrLf
  pbLoading.Value = pbLoading.Value + 20
End If
If Left(InData, 1) = "5" Then
  MsgBox "Error: " & Right(InData, Len(InData) - 1), vbCritical, "Error " & Left(InData, 3)
End If
End Sub

Private Sub tcpSmtp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Error: " & Description, vbCritical, "Error " & Number
tcpSmtp_Close
End Sub

Private Sub tmrIdle_Timer()
Me.Caption = "E-Mail Server - Idle"
End Sub
