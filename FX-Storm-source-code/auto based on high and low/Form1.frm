VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Virtual FX Trader By Dr Jeeni"
   ClientHeight    =   10605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   44
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
      Left            =   0
      TabIndex        =   42
      Top             =   5280
      Width           =   8655
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   240
         Width           =   8415
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Current Price"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   33
      Top             =   3960
      Width           =   8535
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   35
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "High:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Low:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Open:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   39
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Close:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   38
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   4335
      Left            =   8880
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   1080
         Top             =   3960
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   120
         Top             =   3120
      End
      Begin Project1.ctxSysTray ctxSysTray1 
         Left            =   0
         Top             =   2040
         _ExtentX        =   450
         _ExtentY        =   450
         TrayIcon        =   "Form1.frx":030A
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   480
         Top             =   3960
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   0
         Top             =   3960
      End
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Advice"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      TabIndex        =   15
      Top             =   2400
      Width           =   8535
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Text            =   "EUR/USD"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         Text            =   "Buy  Over"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Text            =   "EUR/USD"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Text            =   "Sell Bellow"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text29 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3240
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3240
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   4800
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6240
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6240
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Currency"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Instruction"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entry"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3240
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "T/P"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4800
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "S/L"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6240
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Privious Price"
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   8535
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Close:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Open:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Low:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "High:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   10230
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14579
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pair Information"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8535
      Begin Project1.lvButtons_H lvButtons_H3 
         Height          =   135
         Left            =   8280
         TabIndex        =   52
         Top             =   840
         Width           =   135
         _extentx        =   238
         _extenty        =   238
         caption         =   "."
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "Form1.frx":0624
         mode            =   0
         value           =   0
         cback           =   -2147483633
      End
      Begin Project1.lvButtons_H lvButtons_H2 
         Height          =   375
         Left            =   6720
         TabIndex        =   51
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
         _extentx        =   2355
         _extenty        =   661
         caption         =   "Monitor All"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "Form1.frx":0654
         mode            =   0
         value           =   0
         cback           =   -2147483633
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "PP only"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   50
         Top             =   960
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7200
         TabIndex        =   47
         Text            =   "60"
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable popups messages"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   46
         Top             =   960
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable Sounds"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin Project1.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _extentx        =   2355
         _extenty        =   661
         caption         =   "Start"
         capalign        =   2
         backstyle       =   2
         font            =   "Form1.frx":0684
         mode            =   0
         value           =   0
         cback           =   14737632
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "min."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   49
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Signal Delay"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Results"
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Width           =   8655
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
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
Timer2.Enabled = False
    
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


  
 
    validations
    Randomize
    theversion = "1.0"
    
    'fill lists
    Combo1.Text = "EUR/USD"
    Combo1.AddItem "EUR/USD" '1
    Combo1.AddItem "USD/JPY" '2
    'Combo1.AddItem "EUR/CHF" '3
    Combo1.AddItem "GBP/USD" '4
    Combo1.AddItem "AUD/USD" '5
    'Combo1.AddItem "GBP/CHF" '7
    Combo1.AddItem "USD/CAD" '8
    Combo1.AddItem "GBP/JPY" '9
   ' Combo1.AddItem "CHF/JPY" '10
    'Combo1.AddItem "EUR/JPY" '11
    Combo1.AddItem "USD/CHF" '13
    'Combo1.AddItem "EUR/GBP" '28
    'Combo1.AddItem "EUR/CAD" '78
    'Combo1.AddItem "AUD/JPY" '106
    'Combo1.AddItem "AUD/CAD" '139
    
    
    'mintues
    
    Combo2.Text = 60
    Combo2.AddItem 10
    Combo2.AddItem 15
    Combo2.AddItem 30
    Combo2.AddItem 45
    Combo2.AddItem 60
    Combo2.AddItem 120
    Combo2.AddItem 180
    
    'creat columns
    'for list cloumns
    ListView1.ColumnHeaders.Add , , "Date", ListView1.Width / 11
    ListView1.ColumnHeaders.Add , , "Time", (ListView1.Width / 11) + 40
    ListView1.ColumnHeaders.Add , , "Pair", ListView1.Width / 11 - 20
    ListView1.ColumnHeaders.Add , , "Entry", ListView1.Width / 11 - 20
    ListView1.ColumnHeaders.Add , , "Actual", ListView1.Width / 11 - 20
    ListView1.ColumnHeaders.Add , , "T/P", ListView1.Width / 11 - 20
    ListView1.ColumnHeaders.Add , , "S/L", ListView1.Width / 11 - 20
    ListView1.ColumnHeaders.Add , , "Ex Price", ListView1.Width / 11 - 20
    ListView1.ColumnHeaders.Add , , "Delay", ListView1.Width / 11 - 20
    ListView1.ColumnHeaders.Add , , "Deal", ListView1.Width / 11 - 20
    ListView1.ColumnHeaders.Add , , "level", ListView1.Width / 11 - 20
    
    
    
    StatusBar1.Panels(1).Text = "Ready"
    
    
    
    Dim temp As String

    temp = GetSetting(Me.Name, "fxmo1", "check1")

    If (temp <> "") Then
        Check1.Value = CLng(temp)
    End If
    
    
    temp = GetSetting(Me.Name, "fxmo1", "check2")

    If (temp <> "") Then
        Check2.Value = CLng(temp)
    End If
    
    
    temp = GetSetting(Me.Name, "fxmo1", "check3")

    If (temp <> "") Then
        Check3.Value = CLng(temp)
    End If
    
    
    temp = GetSetting(Me.Name, "fxmo1", "combo2")
    
    Combo2.Text = temp
    If Combo2.Text = "" Then
        Combo2.Text = 60
    End If
    
    
    
      If Check2.Value = 1 Then
            ctxSysTray1.AddIconToSystray "FX-Virtual-Trader " & theversion

      End If
    
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








Public Sub rounds4()

    On Error GoTo thend


    Text1.Text = FormatNumber(Round(Text1.Text, 4), 4)
    Text2.Text = FormatNumber(Round(Text2.Text, 4), 4)
    Text3.Text = FormatNumber(Round(Text3.Text, 4), 4)
    Text4.Text = FormatNumber(Round(Text4.Text, 4), 4)
    On Error Resume Next
    Text29.Text = FormatNumber(Round(Text29.Text, 4), 4)
    On Error Resume Next
    Text30.Text = FormatNumber(Round(Text30.Text, 4), 4)
    On Error Resume Next
    Text31.Text = FormatNumber(Round(Text31.Text, 4), 4)
    On Error Resume Next
    Text32.Text = FormatNumber(Round(Text32.Text, 4), 4)
    On Error Resume Next
    Text33.Text = FormatNumber(Round(Text33.Text, 4), 4)
    On Error Resume Next
    Text34.Text = FormatNumber(Round(Text34.Text, 4), 4)
    On Error Resume Next
    Text9.Text = FormatNumber(Round(Text9.Text, 4), 4)
    On Error Resume Next
    Text8.Text = FormatNumber(Round(Text8.Text, 4), 4)
    On Error Resume Next
    Text7.Text = FormatNumber(Round(Text7.Text, 4), 4)
    On Error Resume Next
    Text6.Text = FormatNumber(Round(Text6.Text, 4), 4)
    
thend:
        logit "Error in rounds4"
   

End Sub

Public Sub calculate()
On Error GoTo thend
    pivot
    calllivepivot
    
thend:
        logit "Error in calculate"
   
End Sub


Public Sub rounds2()

  On Error GoTo thend
  
    Text1.Text = FormatNumber(Round(Text1.Text, 2), 2)
    Text2.Text = FormatNumber(Round(Text2.Text, 2), 2)
    Text3.Text = FormatNumber(Round(Text3.Text, 2), 2)
    Text4.Text = FormatNumber(Round(Text4.Text, 2), 2)
    On Error Resume Next
    Text29.Text = FormatNumber(Round(Text29.Text, 2), 2)
    On Error Resume Next
    Text30.Text = FormatNumber(Round(Text30.Text, 2), 2)
    On Error Resume Next
    Text31.Text = FormatNumber(Round(Text31.Text, 2), 2)
    On Error Resume Next
    Text32.Text = FormatNumber(Round(Text32.Text, 2), 2)
    On Error Resume Next
    Text33.Text = FormatNumber(Round(Text33.Text, 2), 2)
    On Error Resume Next
    Text34.Text = FormatNumber(Round(Text34.Text, 2), 2)
    On Error Resume Next
    Text9.Text = FormatNumber(Round(Text9.Text, 2), 2)
    On Error Resume Next
    Text8.Text = FormatNumber(Round(Text8.Text, 2), 2)
    On Error Resume Next
    Text7.Text = FormatNumber(Round(Text7.Text, 2), 2)
    On Error Resume Next
    Text6.Text = FormatNumber(Round(Text6.Text, 2), 2)
    
thend:
        logit "Error in rounds2"
   
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting Me.Name, "fxmo1", "check1", Check1.Value
    SaveSetting Me.Name, "fxmo1", "check2", Check2.Value
    SaveSetting Me.Name, "fxmo1", "check3", Check3.Value
    SaveSetting Me.Name, "fxmo1", "combo2", Combo2.Text
    ctxSysTray1.RemoveIconFromSystray
    ForceQuit
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lvButtons_H1_Click()
   Text11.Text = ""
   opstart
    
End Sub


Public Sub ForceQuit()
    ExitProcess 1
End Sub


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
                If l1 <> "" And l1 <> "<!-- Close Connection -->" Then
                    List1.AddItem l1
                     Call splitbysymbol(l1, "!")
                     
                End If
                
            Loop
            
        Close #14
        
        'get close
        thesplit = Split(List1.List(List1.ListCount - 1), "|")
        temp = thesplit(1)
        thesplit = Split(temp, "*")
        theclose = thesplit(1)
        Text6.Text = theclose
        
        
        'get open
        thesplit = Split(List1.List(List1.ListCount - 1), "|")
        temp = thesplit(0)
        thesplit = Split(temp, "*")
        theopen = thesplit(1)
        Text7.Text = theopen
        
        
        
        'get high
        thesplit = Split(List1.List(List1.ListCount - 1), "|")
        temp = thesplit(2)
        thesplit = Split(temp, "*")
        thehigh = thesplit(1)
        Text9.Text = thehigh
        
        
        'get low
        thesplit = Split(List1.List(List1.ListCount - 1), "|")
        temp = thesplit(3)
        thesplit = Split(temp, "*")
        thelow = thesplit(1)
        Text8.Text = thelow
        
        
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

   On Error GoTo thend
    Timer4.Enabled = True
    
    calllivepivot
    'DoEvents
    checkprofit1
    Dim xcounter As Long
    Dim thesplit
    
    
    thesplit = Split(Text5.Text, "/")
    
    
    Open App.Path & "\" & thesplit(0) & thesplit(1) & "-History.csv" For Output As #21
         Print #21, "Date" & "," & "Time" & "," & "Pair" & "," & "Entry" & "," & "Actual" & "," & "T/P" & "," & "S/L" & "," & "Ex price" & "," & "Delay" & "," & "Deal" & "," & "Level"
             
         For xcounter = 1 To Form1.ListView1.ListItems.count
            Set Litem = ListView1.ListItems.Item(xcounter)
            Print #21, ListView1.ListItems.Item(xcounter) & "," & Litem.ListSubItems.Item(1) & "," & Litem.ListSubItems.Item(2) & "," & Litem.ListSubItems.Item(3) & "," & Litem.ListSubItems.Item(4) & "," & Litem.ListSubItems.Item(5) & "," & Litem.ListSubItems.Item(6) & "," & Litem.ListSubItems.Item(7) & "," & Litem.ListSubItems.Item(8) & "," & Litem.ListSubItems.Item(9) & "," & Litem.ListSubItems.Item(10)
            
          Next xcounter
        
        Close #21
        
     Timer4.Enabled = False
thend:
        logit "Error in Timer1"
      Timer4.Enabled = False
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
                
                  
                
                
                
                
                
                
                Call LVW_ModifyLine(ListView1, ListView1.ListItems.count, True, vbBlue, "Won", False, False)
                            
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
                 
                Call LVW_ModifyLine(ListView1, ListView1.ListItems.count, True, vbRed, "Lost", False, False)
                
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
                
                Call LVW_ModifyLine(ListView1, ListView1.ListItems.count, True, vbBlue, "Wond", False, False)
                
                
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
                
                Call LVW_ModifyLine(ListView1, ListView1.ListItems.count, True, vbRed, "Lost", False, False)
             
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
    allow = Trim(OpenURL("http://digi77.com/software/forex/data/test.pp"))
      
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
