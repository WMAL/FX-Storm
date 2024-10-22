VERSION 5.00
Object = "{33335233-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "SMTP34.OCX"
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comments - Feedback"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "tellfreind.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comments to us"
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin Project1.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   3720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "lvButtons_H1"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.PictureBox WindowsXPC1 
         Height          =   480
         Left            =   2880
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   9
         Top             =   2160
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "Your name so we know who is taking to us :)"
         Top             =   480
         Width           =   3255
      End
      Begin VB.PictureBox Command1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   4
         ToolTipText     =   "Send the comments or bugs to us"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         ToolTipText     =   "Your email so we can get back to you"
         Top             =   860
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Type your message here"
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Your Name:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin SMTPLib.SMTP SMTP1 
         Left            =   960
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         MailServer      =   ""
         From            =   ""
         To              =   ""
         Cc              =   ""
         BCc             =   ""
         ReplyTo         =   ""
         Date            =   ""
         Subject         =   ""
         MessageText     =   ""
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Your Email Address:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   860
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Message:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!
'Please read before making use of this code!
'Disclaimer: This is illegal if executed on real victims and could land you in prison for sure.
'This is intended for educational purposes only. We take no responsibility at all for your actions.
'This code is provided by EEEDS Eagle Eye Digital Security (Oman) for education purpose only.
'For more educational source codes please visit us http://www.digi77.com
'Author of this code W. Al Maawali Founder of  Eagle Eye Digital Solutions and Oman0.net can be reached via warith@digi77.com .

'Sharing knowledge is not about giving people something, or getting something from them.
'That is only valid for information sharing.
'Sharing knowledge occurs when people are genuinely interested in helping one another develop new capacities for action;
'it is about creating learning processes.
'Peter Senge
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!

Option Explicit
Private Sub Command1_Click()
If (Text1.Text = "" Or Text2 = "" Or Text3 = "") Then
MsgBox "Please fill all the input boxes"
Exit Sub
End If

If (Form1.msgcounter < 5) Then
Command1.Enabled = False
  On Error Resume Next
        SMTP1.WinsockLoaded = True

        SMTP1.MailServer = "mail.digi77.com"
        On Error Resume Next
        SMTP1.Action = 4 'Reset Headers
        On Error Resume Next
        SMTP1.From = "support@digi77.com"
        SMTP1.To = "support@digi77.com"
        SMTP1.Subject = "feedback from: " & Text3
        SMTP1.OtherHeaders = "Content-Type: text/html; charset=windows-1256"""
        SMTP1.MessageText = "Name : " & Text3 & vbCrLf & "Email : " & Text2 & vbCrLf & "Registered : " & Form1.demo & vbCrLf & "real name : " & Form1.realname & vbCrLf & "Message: " & Text1 & vbCrLf
        SMTP1.Action = 3 ' 'Send Message
        Form1.msgcounter = Form1.msgcounter + 1
        MsgBox "Msg Sent Thank You"
        Command1.Enabled = True
        Else
        MsgBox "You Can't send more than 5 messages in one session"
        End If
End Sub

Private Sub Form_Load()


If (settings.Check4.Value = 1) Then
WindowsXPC1.InitSubClassing
End If


'get info from registry

Text1.Text = GetSetting(Me.name, "feedbSettings", "text1")
Text2.Text = GetSetting(Me.name, "feedbSettings", "text2")
Text3.Text = GetSetting(Me.name, "feedbSettings", "text3")
End Sub

Private Sub Form_Terminate()
SaveSetting Me.name, "feedbSettings", "text1", Text1
 SaveSetting Me.name, "feedbSettings", "text2", Text2
SaveSetting Me.name, "feedbSettings", "text3", Text3
End Sub

Private Sub Form_Unload(Cancel As Integer)
 SaveSetting Me.name, "feedbSettings", "text1", Text1
 SaveSetting Me.name, "feedbSettings", "text2", Text2
SaveSetting Me.name, "feedbSettings", "text3", Text3
End Sub

