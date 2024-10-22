VERSION 5.00
Begin VB.Form Form3b 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Settings"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin Project1.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   4080
      TabIndex        =   31
      ToolTipText     =   "Do not save changes"
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Cancel"
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
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   2520
      TabIndex        =   30
      ToolTipText     =   "Save changes"
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Save"
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
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "News Download Options"
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      TabIndex        =   20
      ToolTipText     =   "News Grabber sources and settings"
      Top             =   2160
      Width           =   4095
      Begin VB.CheckBox Check15 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Notify me for new head lines"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox Check14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "CNN Politics"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   28
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox Check13 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "CNN Top"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "CNN Economy"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox Check11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "CNN Markets"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reuters Politics"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox Check9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reuters Top"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox Check8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reuters Business"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox Check7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Daily Fx Currency"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "General "
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   4080
      TabIndex        =   15
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox WindowsXPC1 
         Height          =   480
         Left            =   2280
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   32
         Top             =   480
         Width           =   1200
      End
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable Sounds"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Allow sounds on some events"
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Start with Windows"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Start the application when windows starts"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable Thems"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Allow themed interface"
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable World Time Zone"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "View world time zone on the main application"
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "On Startup "
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      ToolTipText     =   "What to be done when the applications starts"
      Top             =   960
      Width           =   3975
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Download Calendar"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Download News"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Value           =   1  'Checked
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Information Retrieval"
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "Update news events every given minutes"
      Top             =   2160
      Width           =   4095
      Begin VB.VScrollBar VScroll1 
         Height          =   330
         Left            =   2760
         TabIndex        =   11
         Top             =   480
         Width           =   255
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   330
         Left            =   2760
         Min             =   1
         TabIndex        =   10
         Top             =   960
         Value           =   1
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         TabIndex        =   7
         Text            =   "10"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1920
         TabIndex        =   5
         Text            =   "1"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "minutes."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "minutes."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Update calendar Every:"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Update calendar events every given minutes"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Update news every:"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Update news events every given minitues"
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Currency Pairs Sorting"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "How should the currency pairs be sorted"
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Standard"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Alphabetical"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form3b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public actionCounter As Integer ' counts minutes till action is performed

Private Sub Check4_Click()
If (Check4.Value = 1) Then
    If Form1.loading = False Then
        Form1.WindowsXPC1.EngineStarted = True
        Form1.WindowsXPC1.InitSubClassing
        Me.WindowsXPC1.EngineStarted = True
        Me.WindowsXPC1.InitSubClassing
        Form1.Hide
        Form1.Show
    
        Form1.StatusBar1.Panels(1).Text = "Themed Gui Enabled"
      'add colors
    
    End If
    Else
        Form1.WindowsXPC1.EndWinXPCSubClassing
        Me.WindowsXPC1.EndWinXPCSubClassing
        Form1.StatusBar1.Panels(1).Text = "Themed Gui Disabled"
End If
End Sub

Private Sub Check5_Click()
    If (Check5.Value = 1) Then
    
        Dim MainKeyRoot As String
        Dim MainSubKey As String
    
        On Error Resume Next
        Form1.Reg.hKey = HKEY_CURRENT_USER
    
        MainKeyRoot = "Software\Microsoft\Windows\CurrentVersion\Run"
    
        Form1.Reg.KeyRoot = MainKeyRoot
        
        Form1.Reg.SetRegistryValue "FXBouncer", App.Path & "\" & App.EXEName & ".exe", REG_SZ
    Else
    
        Form1.Reg.hKey = HKEY_CURRENT_USER
        MainKeyRoot = "Software\Microsoft\Windows\CurrentVersion\Run"
        Form1.Reg.KeyRoot = MainKeyRoot
        Form1.Reg.DeleteValue "FXBouncer"
    End If
End Sub
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



Private Sub Form_Load()
    getsaved
    actionCounter = 0
    
    If (Form3b.Check4.Value = 1) Then
       
    End If
    
End Sub

'get saved data
Public Sub getsaved()
    
    'get saved data
    'get info from registry
    Dim tempregvalue As String
    tempregvalue = GetSetting(Me.name, "fx3", "Option1")
    If (tempregvalue = "True") Then
       Option1.Value = True
       alphcom
              
    Else
        Option1.Value = False
    End If
    
    tempregvalue = GetSetting(Me.name, "fx3", "Option2")
    If (tempregvalue = "True") Then
       Option2.Value = True
       rancombo
              
    Else
        Option2.Value = False
       
    End If
    
    If Option2.Value = False And Option1.Value = False Then
        Option2.Value = True
        rancombo
    End If
   
    
    Text1.Text = GetSetting(Me.name, "fx3", "text1")
    Text2.Text = GetSetting(Me.name, "fx3", "text2")
    
    If Text1.Text = "" Then
         Text1.Text = "1"
    End If
   
    If Text2.Text = "" Then
         Text2.Text = "10"
    End If
    
    
    
    
    Dim temp As String

    temp = GetSetting(Me.name, "fx3", "settingscheck1")

    If (temp <> "") Then
        Check1.Value = CLng(temp)
    End If
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck2")

    If (temp <> "") Then
        Check2.Value = CLng(temp)
    End If
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck3")

    If (temp <> "") Then
        Check3.Value = CLng(temp)
    End If
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck4")

    If (temp <> "") Then
        Check4.Value = CLng(temp)
    End If
    
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck5")

    If (temp <> "") Then
        Check5.Value = CLng(temp)
    End If
    
    
    
      
      
    temp = GetSetting(Me.name, "fx3", "settingscheck6")

    If (temp <> "") Then
        Check6.Value = CLng(temp)
    End If
    
    
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck7")

    If (temp <> "") Then
        Check7.Value = CLng(temp)
    End If
   
   
    temp = GetSetting(Me.name, "fx3", "settingscheck8")

    If (temp <> "") Then
        Check8.Value = CLng(temp)
    End If
    
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck9")

    If (temp <> "") Then
        Check9.Value = CLng(temp)
    End If
    
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck10")

    If (temp <> "") Then
        Check10.Value = CLng(temp)
    End If
    
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck11")

    If (temp <> "") Then
        Check11.Value = CLng(temp)
    End If
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck12")

    If (temp <> "") Then
        Check12.Value = CLng(temp)
    End If
    
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck13")

    If (temp <> "") Then
        Check13.Value = CLng(temp)
    End If
    
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck14")

    If (temp <> "") Then
        Check14.Value = CLng(temp)
    End If
    
    
    temp = GetSetting(Me.name, "fx3", "settingscheck15")

    If (temp <> "") Then
        Check15.Value = CLng(temp)
    End If
    
    
    
End Sub

Private Sub lvButtons_H1_Click()
    SaveSetting Me.name, "fx3", "option1", Option1.Value
    SaveSetting Me.name, "fx3", "option2", Option2.Value
    SaveSetting Me.name, "fx3", "text1", Text1.Text
    SaveSetting Me.name, "fx3", "text2", Text2.Text
    SaveSetting Me.name, "fx3", "settingscheck1", Check1.Value
    SaveSetting Me.name, "fx3", "settingscheck2", Check2.Value
    SaveSetting Me.name, "fx3", "settingscheck3", Check3.Value
    SaveSetting Me.name, "fx3", "settingscheck4", Check4.Value
    SaveSetting Me.name, "fx3", "settingscheck5", Check5.Value
    SaveSetting Me.name, "fx3", "settingscheck6", Check6.Value
    
    SaveSetting Me.name, "fx3", "settingscheck7", Check7.Value
    SaveSetting Me.name, "fx3", "settingscheck8", Check8.Value
    SaveSetting Me.name, "fx3", "settingscheck9", Check9.Value
    SaveSetting Me.name, "fx3", "settingscheck10", Check10.Value
    SaveSetting Me.name, "fx3", "settingscheck11", Check11.Value
    SaveSetting Me.name, "fx3", "settingscheck12", Check12.Value
    SaveSetting Me.name, "fx3", "settingscheck13", Check13.Value
    SaveSetting Me.name, "fx3", "settingscheck14", Check14.Value
    SaveSetting Me.name, "fx3", "settingscheck15", Check15.Value
    Unload Me
End Sub

Private Sub lvButtons_H2_Click()
    Me.Hide
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    Form1.Combo1.Clear
    'pairs combo box
    alphcom
    
End If
End Sub


Public Sub alphcom()
    Form1.Combo1.Text = "AUD/CAD"
    Form1.Combo1.AddItem "AUD/CAD"
    Form1.Combo1.AddItem "AUD/JPY"
    Form1.Combo1.AddItem "AUD/NZD"
    Form1.Combo1.AddItem "AUD/USD"
    Form1.Combo1.AddItem "EUR/AUD"
    Form1.Combo1.AddItem "EUR/CAD"
    Form1.Combo1.AddItem "EUR/CHF"
    Form1.Combo1.AddItem "EUR/GBP"
    Form1.Combo1.AddItem "EUR/JPY"
    Form1.Combo1.AddItem "EUR/SEK"
    Form1.Combo1.AddItem "EUR/USD"
    Form1.Combo1.AddItem "GBP/CHF"
    Form1.Combo1.AddItem "GBP/JPY"
    Form1.Combo1.AddItem "GBP/USD"
    Form1.Combo1.AddItem "GBP/SGD"
    Form1.Combo1.AddItem "GBP/CHF"
    Form1.Combo1.AddItem "GBP/USD"
    Form1.Combo1.AddItem "NZD/USD"
    Form1.Combo1.AddItem "USD/CAD"
    Form1.Combo1.AddItem "USD/CHF"
    Form1.Combo1.AddItem "USD/HKD"
    Form1.Combo1.AddItem "USD/JPY"
    Form1.Combo1.AddItem "USD/MXN"
End Sub



Public Sub rancombo()
        Form1.Combo1.Text = "EUR/USD"
        Form1.Combo1.AddItem "EUR/USD"
        Form1.Combo1.AddItem "USD/JPY"
        Form1.Combo1.AddItem "GBP/USD"
        Form1.Combo1.AddItem "USD/CHF"
        Form1.Combo1.AddItem "EUR/CHF"
        Form1.Combo1.AddItem "AUD/USD"
        Form1.Combo1.AddItem "USD/CAD"
        Form1.Combo1.AddItem "NZD/USD"
        Form1.Combo1.AddItem "EUR/GBP"
        Form1.Combo1.AddItem "EUR/JPY"
        Form1.Combo1.AddItem "GBP/JPY"
        Form1.Combo1.AddItem "GBP/CHF"
        Form1.Combo1.AddItem "EUR/AUD"
        Form1.Combo1.AddItem "EUR/CAD"
        Form1.Combo1.AddItem "AUD/CAD"
        Form1.Combo1.AddItem "AUD/JPY"
        Form1.Combo1.AddItem "AUD/NZD"
        Form1.Combo1.AddItem "EUR/SEK"
        Form1.Combo1.AddItem "GBP/CHF"
        Form1.Combo1.AddItem "GBP/SGD"
        Form1.Combo1.AddItem "GBP/USD"
        Form1.Combo1.AddItem "USD/HKD"
        Form1.Combo1.AddItem "USD/MXN"
        
End Sub
Private Sub Option2_Click()
    If Option2.Value = True Then
        Form1.Combo1.Clear
        rancombo
        
    End If
End Sub

Private Sub Text1_Change()
    If (Val(Text1.Text) <= 0) Then Text1.Text = "1"
    If (Val(Text1.Text) <= VScroll1.Max) Then VScroll1.Value = VScroll1.Max - Val(Text1.Text)

End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text2_Change()
    If (Val(Text2.Text) <= 0) Then Text2.Text = "1"
    If (Val(Text2.Text) <= VScroll2.Max) Then VScroll2.Value = VScroll2.Max - Val(Text2.Text)

End Sub

Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End Sub



Private Sub VScroll1_Change()
    Text1.Text = VScroll1.Max - VScroll1.Value
End Sub

Private Sub VScroll2_Change()
    Text2.Text = VScroll2.Max - VScroll2.Value
End Sub
