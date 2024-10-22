VERSION 5.00
Object = "{0E09D249-9012-48F2-9DA7-349A59F7CCB0}#1.0#0"; "Alink_Control.ocx"
Begin VB.Form Form3c 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About FX-Storm"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "Formabout.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin Project1.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      ToolTipText     =   "Read and acceptance of the Terms and conditions "
      Top             =   3120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Terms and Conditions"
      CapAlign        =   2
      BackStyle       =   2
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
   Begin Project1.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Send the pc code to us when you intend to buy the software"
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "Send Code"
      CapAlign        =   2
      BackStyle       =   2
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
      Left            =   3120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   11
      Top             =   0
      Width           =   1200
   End
   Begin Alink_Control.ALink ALink1 
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      ToolTipText     =   "Click to visit us"
      Top             =   240
      Width           =   1245
      _ExtentX        =   2064
      _ExtentY        =   344
      BackColor       =   14737632
      Caption         =   "www.digi77.com"
      FontName        =   "MS Sans Serif"
      ForeColor       =   12582912
      FontSize        =   8.25
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Licensed to:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4560
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "For sales questions contact sales@digi77.com"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "For support questions contact support@digi77.com"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "(C) Copyright 1997-2006 Dabdoob Soft and Modern Computers (Muscat Sabco Center)  All Rights Reserved."
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label3"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Version:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "FX-Storm"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form3c"
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

Private Sub Form_Load()

If (settings.Check4.Value = 1) Then
WindowsXPC1.InitSubClassing
End If
Label3.Caption = Form1.theversion
Label8.Caption = Form1.realname
End Sub

Private Sub lvButtons_H1_Click()
Load Form4
Form4.Show
Form4.SetFocus
End Sub

Private Sub lvButtons_H2_Click()
    Load Form4
    Form4.Show
    Form4.SetFocus
End Sub

Private Sub lvButtons_H3_Click()
    frmSysInfo1.Show

'Dim objShellHelper As New SHDocVw.ShellUIHelper
    
'objShellHelper.AddFavorite "http://www.oman70.net/", "Dabdoob Soft"
End Sub
