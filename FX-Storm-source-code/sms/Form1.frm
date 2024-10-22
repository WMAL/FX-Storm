VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Virtual FX Trader by Dr Jeeni"
   ClientHeight    =   10605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "USD/JPY"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   42
      Top             =   3840
      Width           =   8535
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         TabIndex        =   46
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   45
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   44
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6600
         TabIndex        =   43
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sell:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buy:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   49
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "low:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   48
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "High:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   47
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "GBP/USD"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   33
      Top             =   2880
      Width           =   8535
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6600
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   35
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "High:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   41
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "low:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   40
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buy:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   39
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sell:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "USD/CHF"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   24
      Top             =   1920
      Width           =   8535
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   27
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6600
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sell:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buy:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   31
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "low:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   30
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "High:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   29
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "USD/CAD"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Width           =   8535
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6600
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   17
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "High:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "low:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buy:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sell:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   0
      TabIndex        =   14
      Top             =   5160
      Width           =   8775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   8760
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   4920
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Logs"
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   6600
      Width           =   8655
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   8415
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "EUR/USD"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8535
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6600
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sell:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buy:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "low:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "High:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10230
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18177
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!
'Please read before doing anything with this code

'Disclaimer: This is illegal if excuted on real people and could land you in prison for sure.
'This is intended for educational purposes only. We take no responsibility at all for your actions.
'This code is provided by EEEDS Eagle Eye Digital Security (Oman) for education pupose only.
'For more educational source codes please visit us http://www.digi77.com/training.html
'Dr Jeeni Founder of www.oman0.net & www.digi77.com wishes you good luck :).

'Sharing knowledge is not about giving people something, or getting something from them.
'That is only valid for information sharing.
'Sharing knowledge occurs when people are genuinely interested in helping one another develop new capacities for action;
'it is about creating learning processes.
'Peter Senge
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!

Option Explicit
Dim Litem As ListItem
Dim thebuy As Integer
Dim thesell As Integer
Public theminusvalue As Double
Public theplusvalue As Double
Dim kickbuy As Integer
Dim kicksell As Integer
Dim hourcounter As Integer
Dim hourcounter2 As Integer
Dim hourcounter3 As Integer
Dim hourcounter4 As Integer
Dim orderprice As Double
Dim theversion As String
'Dim isess As Long



Private Sub Command1_Click()
Timer1.Enabled = True
    
End Sub

Private Sub Combo2_Change()
On Error Resume Next
ProgressBar1.Max = Val(Combo2.Text)
End Sub

Private Sub Command4_Click()
 ctxSysTray1.Popup "System bought at the price of ", Combo1.Text, Information
      
End Sub

Private Sub ctxSysTray1_DblClick(Button As Integer)
  Form1.Show
  Me.WindowState = 0
End Sub

Private Sub Form_Load()
livepivot ("http://qb.live.gftforex.com/quotes.jsp")

  
 
    
    Randomize
    theversion = "1.0"
   
    StatusBar1.Panels(1).Text = "Ready"
    
    
    
    Dim temp As String

    temp = GetSetting(Me.Name, "fxmo1", "check1")

    If (temp <> "") Then
        'Check1.Value = CLng(temp)
    End If
    
    
   
    
    
    
      'If Check2.Value = 1 Then
          '  ctxSysTray1.AddIconToSystray "FX-Virtual-Trader " & theversion

      'End If
    
thend:
        logit "Error in load"
End Sub

Public Sub getthevalues()
On Error GoTo thend


 If Combo1.Text = "AUD/JPY" Then
        getdily "http://digi77.com/software/forex/data/Text/AUDJPY.csv"
        calculate
        
        rounds2
        theminusvalue = 0.15
        theplusvalue = 0.05
        
     ElseIf Combo1.Text = "AUD/NZD" Then
        getdily "http://digi77.com/software/forex/data/Text/AUDNZD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
        
     ElseIf Combo1.Text = "AUD/USD" Then
        getdily "http://digi77.com/software/forex/data/Text/AUDUSD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "EUR/AUD" Then
        getdily "http://digi77.com/software/forex/data/Text/EURAUD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "EUR/CHF" Then
        getdily "http://digi77.com/software/forex/data/Text/EURCHF.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "EUR/CZK" Then
        getdily "http://digi77.com/software/forex/data/Text/EURCZK.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "EUR/DKK" Then
        getdily "http://digi77.com/software/forex/data/Text/EURDKK.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "EUR/GBP" Then
        getdily "http://digi77.com/software/forex/data/Text/EURGBP.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "EUR/HUF" Then
        getdily "http://digi77.com/software/forex/data/Text/EURHUF.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "EUR/JPY" Then
       
        getdily "http://digi77.com/software/forex/data/Text/EURJPY.csv"
        calculate
        rounds2
        theminusvalue = 0.15
        theplusvalue = 0.05
        
      ElseIf Combo1.Text = "EUR/NOK" Then
        getdily "http://digi77.com/software/forex/data/Text/EURNOK.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
     ElseIf Combo1.Text = "EUR/PLN" Then
        getdily "http://digi77.com/software/forex/data/Text/EURPLN.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
     ElseIf Combo1.Text = "EUR/SEK" Then
        getdily "http://digi77.com/software/forex/data/Text/EURSEK.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
     ElseIf Combo1.Text = "EUR/USD" Then
        getdily "http://digi77.com/software/forex/data/Text/EURUSD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
     ElseIf Combo1.Text = "GBP/CHF" Then
        getdily "http://digi77.com/software/forex/data/Text/GBPCHF.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
     ElseIf Combo1.Text = "GBP/JPY" Then
        getdily "http://digi77.com/software/forex/data/Text/GBPJPY.csv"
        calculate
        rounds2
        theminusvalue = 0.15
        theplusvalue = 0.05
        
     ElseIf Combo1.Text = "GBP/USD" Then
        getdily "http://digi77.com/software/forex/data/Text/GBPUSD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
     ElseIf Combo1.Text = "NZD/USD" Then
        getdily "http://digi77.com/software/forex/data/Text/NZDUSD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
     ElseIf Combo1.Text = "USD/CAD" Then
        getdily "http://digi77.com/software/forex/data/Text/USDCAD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
     ElseIf Combo1.Text = "USD/CHF" Then
        getdily "http://digi77.com/software/forex/data/Text/USDCHF.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
     ElseIf Combo1.Text = "USD/DKK" Then
        getdily "http://digi77.com/software/forex/data/Text/USDDKK.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
     ElseIf Combo1.Text = "USD/HKD" Then
        getdily "http://digi77.com/software/forex/data/Text/USDHKD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "USD/JPY" Then
        getdily "http://digi77.com/software/forex/data/Text/USDJPY.csv"
        calculate
        rounds2
        theminusvalue = 0.15
        theplusvalue = 0.05
        
      ElseIf Combo1.Text = "USD/MXN" Then
        getdily "http://digi77.com/software/forex/data/Text/USDMXN.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "GBP/NOK" Then
        getdily "http://digi77.com/software/forex/data/Text/GBPNOK.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "GBP/SAR" Then
        getdily "http://digi77.com/software/forex/data/Text/GBPSAR.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "GBP/SGD" Then
        getdily "http://digi77.com/software/forex/data/Text/GBPSGD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "GBP/USD" Then
        getdily "http://digi77.com/software/forex/data/Text/GBPUSD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "GBP/CHF" Then
        getdily "http://digi77.com/software/forex/data/Text/GBPCHF.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
      ElseIf Combo1.Text = "EUR/CAD" Then
        getdily "http://digi77.com/software/forex/data/Text/EURCAD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
        
      ElseIf Combo1.Text = "AUD/CAD" Then
        getdily "http://digi77.com/software/forex/data/Text/AUDCAD.csv"
        calculate
        rounds4
        theminusvalue = 0.0015
        theplusvalue = 0.0005
        
     End If
     
     
     
thend:
        logit "Error in getthevalues"
     
End Sub










'download pairs files
Public Sub getdily(url As String)


    On Error GoTo thend
    
    Dim l1 As String
    Dim theopen, thelow, thehigh, theclose, thepivot As Double
    Dim thesplit
    
    'random file name
     Dim filerandom As Integer
     Dim thefilename As String
     filerandom = (Rnd * 10000)
     thefilename = "\fxm-" & filerandom & ".txt"
    
    
    
    
    Open App.Path & thefilename For Output As #15
        Print #15, Trim(OpenURL(url))
    Close #15
    
    List1.Clear
    
    
    ' Dim getext As Boolean
        Open App.Path & thefilename For Input As #16
            Do While Not EOF(16)
                Input #16, l1
            
                
                If l1 <> "" Then
                    List1.AddItem Trim(l1)
                      
                    
                End If
                
            Loop
            
        Close #16
        Kill App.Path & thefilename
        
        thesplit = Split(List1.List(List1.ListCount - 1), vbTab)
        Text1.Text = thesplit(2)
        Text3.Text = thesplit(1)
        Text2.Text = thesplit(3)
        Text4.Text = thesplit(4)
        
        Form2.Text3 = thesplit(2)
        Form2.Text2 = thesplit(3)
        Form2.Text4 = thesplit(4)
        Form2.pivot

thend:
        logit "Error in getdily"
    
End Sub
'end get daily




'pivot calculate
'calculate all pivot values
Public Sub pivot()
    On Error GoTo thend
    
    
    Dim sum As Double, tlow As Double, thigh As Double, tclose As Double, temp1 As Double, temp2 As Double
    thigh = Text1.Text
    tlow = Text2.Text
    tclose = Text4.Text
      
    'get pivot ponit
    sum = thigh + tlow + tclose
    Text29.Text = sum / 3
    Text30.Text = sum / 3
    
    
    
    'get r2
    Text31.Text = (Text29.Text * 2) - tlow
    
    'get r1
    temp1 = Text29.Text
    temp2 = Text31.Text
    sum = temp1 + temp2
    Text34.Text = sum / 2
    
     
    'get s2
    Text32.Text = (Text29.Text * 2) - thigh
    
    
    'get s1
    temp1 = Text32.Text
    temp2 = Text29.Text
    Text33.Text = (temp1 + temp2) / 2
    
    
    
thend:
        logit "Error in pivot"
   
End Sub









Private Sub Form_Unload(Cancel As Integer)
    SaveSetting Me.Name, "fxmo1", "check1", Check1.Value
    SaveSetting Me.Name, "fxmo1", "check2", Check2.Value
    SaveSetting Me.Name, "fxmo1", "check3", Check3.Value
    SaveSetting Me.Name, "fxmo1", "combo2", Combo2.Text
    ctxSysTray1.RemoveIconFromSystray
    ForceQuit
End Sub

Private Sub lvButtons_H1_Click()
   Text11.Text = ""
   opstart
    
End Sub


Public Sub ForceQuit()
    ExitProcess 1
End Sub





'op start
Public Sub opstart()

 On Error GoTo thend
 
 
 If lvButtons_H1.Caption = "Start" Then
        
        
        thebuy = 0
        thesell = 0
        kickbuy = 0
        kicksell = 0
        hourcounter = 0
        hourcounter2 = 0
        ProgressBar1.Value = 0
        hourcounter3 = 0
        hourcounter4 = 0
        ProgressBar1.Max = Combo2.Text
        List1.Clear
         
        Me.Caption = "FX Virtual Trader" & " " & Combo1.Text & " " & Combo2.Text
        
        StatusBar1.Panels(1).Text = "Calculating..."
        
        getthevalues
        
        If Check3.Value = 0 Then
            Form2.dochecksum
        End If
        
        Text5.Text = Combo1.Text
        Text27.Text = Combo1.Text
        Timer1.Enabled = True
        StatusBar1.Panels(1).Text = "Waiting for signal..."
        lvButtons_H1.Caption = "Stop"
    Else
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        StatusBar1.Panels(1).Text = "Ready"
        lvButtons_H1.Caption = "Start"
    End If
    
thend:
        logit "Error in opstart"
   
End Sub

'call the live pibot
Public Sub calllivepivot()

On Error GoTo thend



If Combo1.Text = "EUR/USD" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=1&type=101")
       rounds4
      
       
                
     ElseIf Combo1.Text = "USD/JPY" Then
       
        Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=2&type=101")
       rounds2
       
       
        
     ElseIf Combo1.Text = "EUR/CHF" Then
    
       
        Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=3&type=101")
       rounds4
       
       
        
     ElseIf Combo1.Text = "GBP/USD" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=4&type=101")
       rounds4
      
      
    ElseIf Combo1.Text = "AUD/USD" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=5&type=101")
      
       
      
       rounds4
      
    ElseIf Combo1.Text = "AUD/USD" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=6&type=101")
      rounds4
       
      
    ElseIf Combo1.Text = "GBP/CHF" Then
    
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=7&type=101")
      
       rounds4
      
         
    ElseIf Combo1.Text = "GBP/JPY" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=9&type=101")
       rounds2
     
      
    ElseIf Combo1.Text = "CHF/JPY" Then
    
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=10&type=101")
       rounds2
       
      
      
    ElseIf Combo1.Text = "EUR/JPY" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=11&type=101")
       rounds2
       
       
    ElseIf Combo1.Text = "USD/CAD" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=12&type=101")
       rounds4
      
      
    ElseIf Combo1.Text = "USD/CHF" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=13&type=101")
       rounds4
       
      
    ElseIf Combo1.Text = "EUR/GBP" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=28&type=101")
       rounds4
     
    ElseIf Combo1.Text = "EUR/CAD" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=78&type=101")
       rounds4
     
       
    ElseIf Combo1.Text = "AUD/JPY" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=106&type=101")
      rounds2
      
      
       
    ElseIf Combo1.Text = "AUD/CAD" Then
      
       Call livepivot("http://64-145-12-164.client.dsl.net/DataService/FXcmgetLiveIntraday.asp?FX_id=139&type=101")
      rounds2
      
             
       
    End If
    
    
thend:
        logit "Error in calllivepivot"
   
End Sub



































'get livepivot
Public Sub livepivot(url As String)

    On Error GoTo thend
    
    Dim getext As Boolean, l1 As String
    Dim thespliter As String
    Dim temp As String
    Dim theopen, thelow, thehigh, theclose As Double
    Dim myString() As String
    Dim i As Long
    
    Dim thesplit
    
    'random file name
     Dim filerandom As Integer
     Dim thefilename As String
     filerandom = (Rnd * 10000)
     thefilename = "\fxmlp-" & filerandom & ".txt"
    
    'llls
        
    Open App.Path & thefilename For Output As #13
        Print #13, Trim(OpenURL(url))
    Close #13
    
    List1.Clear
    
  
    
    ' Dim getext As Boolean
        Open App.Path & thefilename For Input As #14
            Do While Not EOF(14)
                Input #14, l1
               
                'gettheday
                If Left(l1, 12) = "<currency id" Then
                    
                   If InStr(l1, "USD/JPY") > 0 Or InStr(l1, "USD/CHF") > 0 Or InStr(l1, "GBP/USD") > 0 Or InStr(l1, "EUR/USD") > 0 Or InStr(l1, "USD/CAD") > 0 Then
                    
                        List1.AddItem l1
                    'Call splitbysymbol(l1, "!")
                    End If
                End If
                
            Loop
            
        Close #14
          
        'currency id="1" symbol="USD/CHF" bid="1.2228" ask="1.2232" high="n/a" low="n/a" bidTrend="0" askTrend="0" />
        
         'eur/usd
        
        'get sell
        thesplit = Split(List1.List(3), "bid=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text9.Text = thesplit(0)
        Text9.Text = Trim(Replace(Text9.Text, Chr(34), Chr(160)))
        
       
        'get buy
        thesplit = Split(List1.List(3), "ask=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        
        Dim temp1, temp2 As Double
        temp1 = CDbl(Text8.Text)
        
        Text8.Text = thesplit(0)
        Text8.Text = Trim(Replace(Text8.Text, Chr(34), Chr(160)))
        
        temp2 = CDbl(Text8.Text)
        
        If temp1 > temp2 Then
            
            Text8.ForeColor = vbRed
            Text9.ForeColor = vbRed
        ElseIf temp1 < temp2 Then
        
            Text8.ForeColor = vbBlue
            Text9.ForeColor = vbBlue
        Else
            Text8.ForeColor = vbBlack
            Text9.ForeColor = vbBlack
        
        End If
        
        
        'get low
        thesplit = Split(List1.List(3), "low=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text7.Text = thesplit(0)
        Text7.Text = Trim(Replace(Text7.Text, Chr(34), Chr(160)))
        
        
        'get high
        thesplit = Split(List1.List(3), "high=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text6.Text = thesplit(0)
        Text6.Text = Trim(Replace(Text6.Text, Chr(34), Chr(160)))
        
        
        
        
        'usd/cad
        
        'get sell
        thesplit = Split(List1.List(4), "bid=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text1.Text = thesplit(0)
        Text1.Text = Trim(Replace(Text1.Text, Chr(34), Chr(160)))
        
       
        'get buy
        thesplit = Split(List1.List(4), "ask=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        
        'Dim temp1, temp2 As Double
        temp1 = CDbl(Text2.Text)
        
        Text2.Text = thesplit(0)
        Text2.Text = Trim(Replace(Text2.Text, Chr(34), Chr(160)))
        
        temp2 = CDbl(Text2.Text)
        
        If temp1 > temp2 Then
            
            Text2.ForeColor = vbRed
            Text1.ForeColor = vbRed
        ElseIf temp1 < temp2 Then
        
            Text2.ForeColor = vbBlue
            Text1.ForeColor = vbBlue
        Else
            Text2.ForeColor = vbBlack
            Text1.ForeColor = vbBlack
        
        End If
        
        
        'get low
        thesplit = Split(List1.List(4), "low=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text3.Text = thesplit(0)
        Text3.Text = Trim(Replace(Text3.Text, Chr(34), Chr(160)))
        
        
        'get high
        thesplit = Split(List1.List(4), "high=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text4.Text = thesplit(0)
        Text4.Text = Trim(Replace(Text4.Text, Chr(34), Chr(160)))
        
        
        
        
        
        
        
         'usd/chf
        
        'get sell
        thesplit = Split(List1.List(1), "bid=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text13.Text = thesplit(0)
        Text13.Text = Trim(Replace(Text13.Text, Chr(34), Chr(160)))
        
       
        'get buy
        thesplit = Split(List1.List(1), "ask=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        'Dim temp1, temp2 As Double
        temp1 = CDbl(Text12.Text)
        
        Text12.Text = thesplit(0)
        Text12.Text = Trim(Replace(Text12.Text, Chr(34), Chr(160)))
        
        temp2 = CDbl(Text12.Text)
        
        If temp1 > temp2 Then
            
            Text12.ForeColor = vbRed
            Text13.ForeColor = vbRed
        ElseIf temp1 < temp2 Then
        
            Text12.ForeColor = vbBlue
            Text13.ForeColor = vbBlue
        Else
            Text12.ForeColor = vbBlack
            Text13.ForeColor = vbBlack
        
        End If
        
        
        'get low
        thesplit = Split(List1.List(1), "low=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text10.Text = thesplit(0)
        Text10.Text = Trim(Replace(Text10.Text, Chr(34), Chr(160)))
        
        
        'get high
        thesplit = Split(List1.List(1), "high=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text5.Text = thesplit(0)
        Text5.Text = Trim(Replace(Text5.Text, Chr(34), Chr(160)))
        
        
        
        
        
         'gbp/usd
        
        'get sell
        thesplit = Split(List1.List(2), "bid=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text14.Text = thesplit(0)
        Text14.Text = Trim(Replace(Text14.Text, Chr(34), Chr(160)))
        
       
        'get buy
        thesplit = Split(List1.List(2), "ask=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        'Dim temp1, temp2 As Double
        temp1 = CDbl(Text15.Text)
        
        Text15.Text = thesplit(0)
        Text15.Text = Trim(Replace(Text15.Text, Chr(34), Chr(160)))
        
        temp2 = CDbl(Text15.Text)
        
        If temp1 > temp2 Then
            
            Text15.ForeColor = vbRed
            Text14.ForeColor = vbRed
        ElseIf temp1 < temp2 Then
        
            Text15.ForeColor = vbBlue
            Text14.ForeColor = vbBlue
        Else
            Text15.ForeColor = vbBlack
            Text14.ForeColor = vbBlack
        
        End If
        'get low
        thesplit = Split(List1.List(2), "low=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text16.Text = thesplit(0)
        Text16.Text = Trim(Replace(Text16.Text, Chr(34), Chr(160)))
        
        
        'get high
        thesplit = Split(List1.List(2), "high=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text17.Text = thesplit(0)
        Text17.Text = Trim(Replace(Text17.Text, Chr(34), Chr(160)))
        
        
        
         'usd/jpy
        
        'get sell
        thesplit = Split(List1.List(0), "bid=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text21.Text = thesplit(0)
        Text21.Text = Trim(Replace(Text21.Text, Chr(34), Chr(160)))
        
       
        'get buy
        thesplit = Split(List1.List(0), "ask=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        'Dim temp1, temp2 As Double
        temp1 = CDbl(Text20.Text)
        
        Text20.Text = thesplit(0)
        Text20.Text = Trim(Replace(Text20.Text, Chr(34), Chr(160)))
        
        temp2 = CDbl(Text20.Text)
        
        If temp1 > temp2 Then
            
            Text20.ForeColor = vbRed
            Text21.ForeColor = vbRed
        ElseIf temp1 < temp2 Then
        
            Text20.ForeColor = vbBlue
            Text21.ForeColor = vbBlue
        Else
            Text20.ForeColor = vbBlack
            Text21.ForeColor = vbBlack
        
        End If
        
        'get low
        thesplit = Split(List1.List(0), "low=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text19.Text = thesplit(0)
        Text19.Text = Trim(Replace(Text19.Text, Chr(34), Chr(160)))
        
        
        'get high
        thesplit = Split(List1.List(0), "high=")
        temp = thesplit(1)
        thesplit = Split(temp, " ")
        Text18.Text = thesplit(0)
        Text18.Text = Trim(Replace(Text18.Text, Chr(34), Chr(160)))
        
        
        
        
        
        
        Kill App.Path & thefilename
        
        
thend:
        logit "Error in livepivot"
       
End Sub





















'spliter by symbol
Public Sub splitbysymbol(theinput As String, theseperator As String)

   
    On Error GoTo erreur
    
    
    Dim tableau_a_remplir() As String
    Dim rep As Integer, nombre_elements  As Integer, i As Integer
    rep = splits(theinput, theseperator, tableau_a_remplir(), nombre_elements)
    If rep <> 0 Then GoTo erreur
        List1.Clear
    
    For i = 0 To nombre_elements - 1
        List1.AddItem tableau_a_remplir(i)
        List1.ListIndex = List1.ListCount - 1
    Next i
    'nb_elem_lbl.Caption = "This string contains " & nombre_elements & " elements."
Exit Sub



erreur:

        logit "Error in splitbysymbol"

End Sub

Private Sub lvButtons_H2_Click()

'For isess = 0 To Combo1.ListCount - 2
'    Dim ss(0 To 12) As New Form1

   ' ss(isess).Show
   ' ss(isess).Combo1 = ss(isess).Combo1.List(isess + 1)
   ' ss(isess).Combo2.Text = Combo2.Text
   ' ss(isess).Check1.Value = Check1.Value
   ' ss(isess).Check2.Value = Check2.Value
   ' ss(isess).Check3.Value = Check3.Value
   ' ss(isess).opstart
'Next isess




End Sub

Private Sub lvButtons_H3_Click()
    Form2.Show
End Sub

Private Sub Timer1_Timer()

   livepivot ("http://qb.live.gftforex.com/quotes.jsp")
End Sub



Public Sub checkprofit1()
   
   On Error GoTo thend
   
   
         Dim thepivot As Double
         Dim theopen As Double
         Dim thesum As Double
         Dim pivotbuyplus As Double
         Dim pivotbuymin As Double
         Dim pivotsellplus As Double
         Dim pivotsellmin As Double
         
         
         
         
         thepivot = Val(Text30.Text)
         theopen = Val(Text7.Text)
         
         pivotbuyplus = thepivot + theminusvalue
         pivotsellplus = thepivot + theplusvalue
         
         pivotbuymin = thepivot - theplusvalue
         pivotsellmin = thepivot - theminusvalue
         If theminusvalue = 0.15 Then
              
              pivotbuyplus = FormatNumber(Round(pivotbuyplus, 2), 2)
              pivotbuymin = FormatNumber(Round(pivotbuymin, 2), 2)
              pivotsellplus = FormatNumber(Round(pivotsellplus, 2), 2)
              pivotsellmin = FormatNumber(Round(pivotsellmin, 2), 2)
              
         Else
              pivotbuyplus = FormatNumber(Round(pivotbuyplus, 4), 4)
              pivotbuymin = FormatNumber(Round(pivotbuymin, 4), 4)
              pivotsellplus = FormatNumber(Round(pivotsellplus, 4), 4)
              pivotsellmin = FormatNumber(Round(pivotsellmin, 4), 4)
         End If
         
         thebuy = 0
         thesell = 0
         
         
         
         
        If Check3.Value = 0 Then
            'find pivot
            Dim i As Long
            Dim temppp As String
            Dim thesplit
            
            For i = 0 To 16
                Dim temp As String
                
                thesplit = Split(Form2.List2.List(i), "|")
                temp = thesplit(1)
                If Trim(temp) = Trim(thepivot) Then
                   temppp = thesplit(0)
                End If
                        
            Next i
            
         Else
            temppp = "pp"
         
         End If
          
        
         
         
         If theopen <= pivotbuyplus And theopen >= pivotbuymin And kickbuy <> 1 Then
            thebuy = 1
            logit "Current price is within the buy price| price:" & Text7.Text & " Level: " & Text29.Text & " " & temppp
         End If
        
         
         If theopen <= pivotsellplus And theopen >= pivotsellmin And kicksell <> 1 Then
            thesell = 1
            logit "Current price is within the sell price| price:" & Text7.Text & " Level: " & Text29.Text & " " & temppp
         End If
            
         
            
         If thebuy = 1 And thesell = 0 Then
            kicksell = 1
            
            'logit "Cureent price is out of the Sell range"
         ElseIf thebuy = 0 And thesell = 1 Then
            kickbuy = 1
             
            'logit "Cureent price is out of the Buy range"
         ElseIf thebuy = 0 And thesell = 0 Then
            logit hourcounter3 & " Current price is out of the sell and buy range restarting the engine again price:" & Text7.Text & " Level: " & Text29.Text & " " & temppp
           
            hourcounter = 0
            ProgressBar1.Value = 0
            If Check3.Value = 0 Then
                Form2.dochecksum
            End If
            
             If hourcounter3 = 5 Then
                 Timer2.Enabled = False
                 Timer1.Enabled = False
                 Timer3.Enabled = True
           
             End If
            hourcounter3 = hourcounter3 + 1
            Exit Sub
            
         ElseIf thebuy = 1 And thesell = 1 Then '***dont reset dont igore
            logit hourcounter4 & " Current price is with in both sell and buy range price:" & Text7.Text & " Level: " & Text29.Text & " " & temppp
             '***reset if even
             'hourcounter = 0
             'ProgressBar1.Value = 0
             '***end reset if even
             
             If hourcounter4 = 5 Then '***by 5 or by delay
                 Timer2.Enabled = False
                 Timer1.Enabled = False
                 Timer3.Enabled = True
           
             End If
            hourcounter4 = hourcounter4 + 1
            Exit Sub
            
         End If
             
             
             
        If hourcounter = Combo2.Text Then
            
            If thebuy = 1 And thesell = 0 Then
                orderprice = theopen
                logit "System bought at the price of :" & Text7.Text & " Level: " & Text29.Text & " " & temppp
                Me.Caption = "FX Virtual Trader" & " " & Combo1.Text & " " & Combo2.Text & " buy"
                If Check2.Value = 1 Then
                    ctxSysTray1.Popup "System bought at the price of " & theopen, Combo1.Text, Information
      
                End If
                
                
                 'play sound when done
                 If (Check1.Value = 1) Then
                        PlaySound ("ball.wav")
                 End If
                
                Timer1.Enabled = False
                Timer2.Enabled = True
                Exit Sub
            End If
            
            
            If thebuy = 0 And thesell = 1 Then
                orderprice = theopen
                logit "System sold at the price of :" & Text7.Text & " Level: " & Text29.Text & " " & temppp
                Me.Caption = "FX Virtual Trader" & " " & Combo1.Text & " " & Combo2.Text & " sell"
                If Check2.Value = 1 Then
                    ctxSysTray1.Popup "System sold at the price of " & theopen, Combo1.Text, Information
      
                End If
                
                  'play sound when done
                 If (Check1.Value = 1) Then
                        PlaySound ("ball.wav")
                 End If
                
                Timer1.Enabled = False
                Timer2.Enabled = True
                Exit Sub
            End If
        
        ElseIf hourcounter > Combo2.Text Then
             Timer2.Enabled = False
             Timer1.Enabled = False
             Timer3.Enabled = True
           
        
        End If
        
        On Error Resume Next
        ProgressBar1.ToolTipText = ProgressBar1.Value & "/" & Combo2.Text
        
        On Error Resume Next
        ProgressBar1.Value = ProgressBar1.Value + 1
        hourcounter = hourcounter + 1
        
        
thend:
        logit "Error in checkprofit1"
End Sub





'logger
Public Sub logit(thetext As String)
     
    If Left(thetext, 5) <> "Error" Then
    
        On Error Resume Next
        Text11.Text = Text11.Text & "--> " & Time & " " & Date & " " & thetext & vbCrLf & vbCrLf
        
        On Error Resume Next
        Text11.SelStart = Len(Text11.Text)
    End If

End Sub

Private Sub Timer2_Timer()
On Error GoTo thend
         
        Timer4.Enabled = True
       
        Dim thepivot As Double
         Dim theopen As Double
         Dim thesum As Double
         Dim pivotbuyplus As Double
         Dim pivotbuymin As Double
         Dim pivotsellplus As Double
         Dim pivotsellmin As Double
         
         
         
          
         
         
         thepivot = Val(Text30.Text)
         theopen = Val(Text7.Text)
         
         If Check3.Value = 0 Then
            'find pivot
            Dim i As Long
            Dim temppp As String
            Dim thesplit
            
            For i = 0 To 16
                Dim temp As String
                
                thesplit = Split(Form2.List2.List(i), "|")
                temp = thesplit(1)
                If Trim(temp) = Trim(thepivot) Then
                   temppp = thesplit(0)
                End If
                        
            Next i
            
         Else
            temppp = "pp"
         
         End If
            
            
          
         calllivepivot
         If thebuy = 1 Then
         
            If theopen >= Val(Text31.Text) Then
                logit " You won the buy deal"
                Me.Caption = "FX Virtual Trader" & " " & Combo1.Text & " " & Combo2.Text & " buy win"
                 If Check2.Value = 1 Then
                    ctxSysTray1.Popup "You won the buy deal", Combo1.Text, Information
      
                 End If
                Set Litem = Form1.ListView1.ListItems.Add(, , Date)
                Litem.ListSubItems.Add , , Time
                Litem.ListSubItems.Add , , Combo1.Text
                Litem.ListSubItems.Add , , Text29.Text
                Litem.ListSubItems.Add , , orderprice
                Litem.ListSubItems.Add , , "True"
                Litem.ListSubItems.Add , , "False"
                Litem.ListSubItems.Add , , Text7.Text
                Litem.ListSubItems.Add , , Combo2.Text & "m"
                Litem.ListSubItems.Add , , "Buy"
                 Litem.ListSubItems.Add , , temppp
                
                  
                
                
                
                
                
                
                Call LVW_ModifyLine(ListView1, ListView1.ListItems.Count, True, vbBlue, "Won", False, False)
                            
                'play sound when done
                 If (Check1.Value = 1) Then
                        PlaySound ("a58337.wav")
                 End If
                
                Timer1.Enabled = False
                Timer3.Enabled = False
                Timer4.Enabled = False
                lvButtons_H1.Caption = "Start"
                opstart
                Timer2.Enabled = False
                Exit Sub
                
           ElseIf theopen <= Val(Text33.Text) Then
                logit " You lost the buy deal"
                Me.Caption = "FX Virtual Trader" & " " & Combo1.Text & " " & Combo2.Text & " buy lose"
                If Check2.Value = 1 Then
                    ctxSysTray1.Popup "You lost the buy deal", Combo1.Text, Information
      
                 End If
                
                Set Litem = Form1.ListView1.ListItems.Add(, , Date)
                Litem.ListSubItems.Add , , Time
                Litem.ListSubItems.Add , , Combo1.Text
                Litem.ListSubItems.Add , , Text29.Text
                Litem.ListSubItems.Add , , orderprice
                Litem.ListSubItems.Add , , "False"
                Litem.ListSubItems.Add , , "True"
                Litem.ListSubItems.Add , , Text7.Text
                Litem.ListSubItems.Add , , Combo2.Text & "m"
                Litem.ListSubItems.Add , , "Buy"
                Litem.ListSubItems.Add , , temppp
                 
                Call LVW_ModifyLine(ListView1, ListView1.ListItems.Count, True, vbRed, "Lost", False, False)
                
                     'play sound when done
                 If (Check1.Value = 1) Then
                        PlaySound ("a7834.wav")
                 End If
                
                Timer1.Enabled = False
                Timer3.Enabled = False
                Timer4.Enabled = False
                lvButtons_H1.Caption = "Start"
                opstart
                Timer2.Enabled = False
                Exit Sub
                
            End If
            
         End If
        
        
        If thesell = 1 Then
         
            If theopen <= Val(Text32.Text) Then
                logit " You won the sell deal"
                Me.Caption = "FX Virtual Trader" & " " & Combo1.Text & " " & Combo2.Text & " sell win"
                If Check2.Value = 1 Then
                    ctxSysTray1.Popup "You won the sell deal", Combo1.Text, Information
      
                 End If
                
                Set Litem = Form1.ListView1.ListItems.Add(, , Date)
                Litem.ListSubItems.Add , , Time
                Litem.ListSubItems.Add , , Combo1.Text
                Litem.ListSubItems.Add , , Text29.Text
                Litem.ListSubItems.Add , , orderprice
                Litem.ListSubItems.Add , , "True"
                Litem.ListSubItems.Add , , "false"
                Litem.ListSubItems.Add , , Text7.Text
                Litem.ListSubItems.Add , , Combo2.Text & " m"
                Litem.ListSubItems.Add , , "Sell"
                Litem.ListSubItems.Add , , temppp
                
                Call LVW_ModifyLine(ListView1, ListView1.ListItems.Count, True, vbBlue, "Wond", False, False)
                
                
                'play sound when done
                 If (Check1.Value = 1) Then
                        PlaySound ("a58337.wav")
                 End If
                
                Timer2.Enabled = False
                'lvButtons_H1.Caption = "Start"
                opstart
                Exit Sub
            ElseIf theopen >= Val(Text34.Text) Then
                logit " You lost the sell deal"
                Me.Caption = "FX Virtual Trader" & " " & Combo1.Text & " " & Combo2.Text & " sell lose"
                 If Check2.Value = 1 Then
                    ctxSysTray1.Popup "You lost the sell deal", Combo1.Text, Information
      
                 End If
                
                Set Litem = Form1.ListView1.ListItems.Add(, , Date)
                Litem.ListSubItems.Add , , Time
                Litem.ListSubItems.Add , , Combo1.Text
                Litem.ListSubItems.Add , , Text29.Text
                Litem.ListSubItems.Add , , orderprice
                Litem.ListSubItems.Add , , "false"
                Litem.ListSubItems.Add , , "True"
                Litem.ListSubItems.Add , , Text7.Text
                Litem.ListSubItems.Add , , Combo2.Text & "m"
                Litem.ListSubItems.Add , , "sell"
                Litem.ListSubItems.Add , , temppp
                
                Call LVW_ModifyLine(ListView1, ListView1.ListItems.Count, True, vbRed, "Lost", False, False)
             
                'play sound when done
                 If (Check1.Value = 1) Then
                        PlaySound ("a7834.wav")
                 End If
                
                Timer2.Enabled = False
                'lvButtons_H1.Caption = "Start"
                opstart
                Exit Sub
                
            End If
            
        End If
        Timer4.Enabled = False
thend:
        logit "Error in Timer2"
        Timer4.Enabled = False
        
End Sub



'validate application and users
Public Sub validations()
    On Error GoTo thend
    
    Dim allow As String
    allow = Trim(OpenURL("http://digi77.com/software/forex/data/ver/1-0.pp"))
      
    If Trim(allow) = 0 Then
        End
    ElseIf (Trim(allow) = 1) Then
       
            
       
    Else
        End
    End If
    
    
thend:
        logit "Error in validations"
End Sub

Private Sub Timer3_Timer()
        
          If hourcounter2 = 10 Then
           'old before modification
            lvButtons_H1.Caption = "Start"
            Timer1.Enabled = False
            Timer2.Enabled = False
            Timer3.Enabled = False
            opstart
            
          End If
          
          logit "Paused " & hourcounter2 & "/10"
          hourcounter2 = hourcounter2 + 1
         
End Sub

Private Sub Timer4_Timer()

            'old before modification
            lvButtons_H1.Caption = "Start"
            Timer1.Enabled = False
            Timer2.Enabled = False
            Timer3.Enabled = False
            
            opstart
            Timer4.Enabled = False
End Sub
