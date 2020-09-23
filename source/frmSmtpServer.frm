VERSION 5.00
Begin VB.Form frmSmtpServer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Enter a smtp server"
   ClientHeight    =   540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbSmtpServers 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmSmtpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim SmtpServer As String
Open App.Path & "\servers.txt" For Input As #1
  Do
    Input #1, SmtpServer
    cmbSmtpServers.AddItem SmtpServer
  Loop Until EOF(1)
Close #1
End Sub

Private Sub lblSave_Click()
If InStr(MailServer, ".") >= 0 Then
  frmMain.MailServer = cmbSmtpServers.Text
Else
  MsgBox "Illegal SMTP Server. Server must be in format: x.x.x.x", vbCritical, "Error. Server not accepted.."
End If
Unload Me
End Sub
