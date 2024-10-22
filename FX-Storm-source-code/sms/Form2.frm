VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form2"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   LinkTopic       =   "Form2"
   ScaleHeight     =   9510
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   4935
      Left            =   7680
      TabIndex        =   45
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6600
      TabIndex        =   44
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   4920
      TabIndex        =   41
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6480
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      TabIndex        =   36
      Text            =   "1.2065"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   35
      Text            =   "1.2169"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   34
      Text            =   "1.2047"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   600
      TabIndex        =   16
      Top             =   4680
      Width           =   1095
   End
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
      Left            =   1800
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
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
      Left            =   1560
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
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
      Left            =   1320
      TabIndex        =   13
      Top             =   3120
      Width           =   1095
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
      Left            =   1080
      TabIndex        =   12
      Top             =   3480
      Width           =   1095
   End
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
      Left            =   840
      TabIndex        =   11
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text10 
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
      TabIndex        =   10
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text35 
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
      Left            =   2040
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text37 
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
      Left            =   2280
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text12 
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
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text13 
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
      Left            =   840
      TabIndex        =   6
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox Text14 
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
      Left            =   1080
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox Text15 
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
      Left            =   1320
      TabIndex        =   4
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox Text16 
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
      Left            =   1560
      TabIndex        =   3
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Text17 
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
      Left            =   1800
      TabIndex        =   2
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox Text36 
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
      Left            =   2040
      TabIndex        =   1
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox Text38 
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
      Left            =   2280
      TabIndex        =   0
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label Label18 
      Caption         =   "Sorted price"
      Height          =   375
      Left            =   5040
      TabIndex        =   43
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "cureent price"
      Height          =   255
      Left            =   5160
      TabIndex        =   42
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Low:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   39
      Top             =   765
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "High:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   38
      Top             =   765
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Close:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   37
      Top             =   765
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "R6:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   33
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "R5:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   32
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "R4:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   31
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "R3:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   30
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "R2:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   29
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "R1:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label44 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "R7:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   27
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label46 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "R8:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   26
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "S1:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "S2:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "S3:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "S4:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   22
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "S5:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "S6:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   20
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label45 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "S7:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label47 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "S8:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   7800
      Width           =   255
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "PP:"
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
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   4800
      Width           =   255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doublearraysval(0 To 17) As Double

Dim thelevelvalue As String

Public Sub dochecksum()


        Text1.Text = Form1.Text7.Text
        
        
        
        Dim thevalue As Double
        If Form1.theminusvalue = 0.15 Then
                    thevalue = 0.11
              Else
              
                    thevalue = 0.0011
              End If
        
        If Val(Text1.Text) <= Val(Text5.Text) + thevalue And Val(Text1.Text) >= Val(Text17.Text) - thevalue Then
            'MsgBox Val(Text1.Text)
            doublearraysval(0) = Val(Text17.Text)
            doublearraysval(1) = Val(Text11.Text)
            doublearraysval(2) = Val(Text10.Text)
            doublearraysval(3) = Val(Text9.Text)
            doublearraysval(4) = Val(Text8.Text)
            doublearraysval(5) = Val(Text7.Text)
            doublearraysval(6) = Val(Text6.Text)
            doublearraysval(7) = Val(Text5.Text)
            doublearraysval(8) = Val(Text12.Text)
            doublearraysval(9) = Val(Text13.Text)
            doublearraysval(10) = Val(Text14.Text)
            doublearraysval(11) = Val(Text15.Text)
            doublearraysval(12) = Val(Text16.Text)
            doublearraysval(13) = Val(Text1.Text)
            
            doublearraysval(14) = Val(Text35.Text)
            doublearraysval(15) = Val(Text36.Text)
            doublearraysval(16) = Val(Text37.Text)
            doublearraysval(17) = Val(Text38.Text)
            
            
            'for lever info
            List2.Clear
            List2.AddItem "r8|" & Val(Text37.Text)
            List2.AddItem "r7|" & Val(Text35.Text)
            List2.AddItem "r6|" & Val(Text5.Text)
            List2.AddItem "r5|" & Val(Text6.Text)
            List2.AddItem "r4|" & Val(Text7.Text)
            List2.AddItem "r3|" & Val(Text8.Text)
            List2.AddItem "r2|" & Val(Text9.Text)
            List2.AddItem "r1|" & Val(Text10.Text)
            List2.AddItem "pp|" & Val(Text11.Text)
            List2.AddItem "s1|" & Val(Text12.Text)
            List2.AddItem "s2|" & Val(Text13.Text)
            List2.AddItem "s3|" & Val(Text14.Text)
            List2.AddItem "s4|" & Val(Text15.Text)
            List2.AddItem "s5|" & Val(Text16.Text)
            List2.AddItem "s6|" & Val(Text17.Text)
            List2.AddItem "s7|" & Val(Text36.Text)
            List2.AddItem "s8|" & Val(Text38.Text)
            
            


            
            
           
            List1.Clear
            Call Sort_num(doublearraysval(), 17)
            
            For i = 0 To 17
            
            If Form1.theminusvalue = 0.15 Then
            
                   List1.AddItem FormatNumber(Round(doublearraysval(i), 2), 2)
              
              Else
              
                   List1.AddItem FormatNumber(Round(doublearraysval(i), 4), 4)
              
              End If
               
            
            Next i
            
            Dim theabove As Double
            Dim thebelows As Double
            
            Dim btp As Double
            Dim bsl As Double
            Dim stp As Double
            Dim ssl As Double
            Dim pp As Double
            Dim s As Long
            For i = 0 To List1.ListCount
            
                If List1.List(i) = Text1.Text Then
                   ' MsgBox List1.List(i) - List1.List(i - 1) & "  " & List1.List(i + 1) - List1.List(i)
                
                
                        
                        If Form1.theminusvalue = 0.15 Then
                             theabove = Val(List1.List(i)) - Val(List1.List(i - 1))
                             theabove = FormatNumber(Round(theabove, 2), 2)
                             thebelows = Val(List1.List(i + 1)) - Val(List1.List(i))
                             thebelows = FormatNumber(Round(thebelows, 2), 2)
                        
                        Else
                             theabove = Val(List1.List(i)) - Val(List1.List(i - 1))
                             theabove = FormatNumber(Round(theabove, 4), 4)
                             thebelows = Val(List1.List(i + 1)) - Val(List1.List(i))
                            
                             thebelows = FormatNumber(Round(thebelows, 4), 4)
                        
                        End If
                           
                         If thebelows < theabove Then
                              pp = List1.List(i + 1)
                              btp = Val(List1.List(i + 3))
                              bsl = Val(List1.List(i - 1))
                              stp = Val(List1.List(i - 2))
                              ssl = Val(List1.List(i + 2))
                              
                              Form1.Text29.Text = pp
                              Form1.Text30.Text = pp
                              Form1.Text31.Text = btp
                              Form1.Text32.Text = stp
                              Form1.Text33.Text = bsl
                              Form1.Text34.Text = ssl
                               
                               
                                
                                    
                             ' MsgBox "below " & pp & " " & btp & " " & bsl & " " & stp & " " & ssl
                              Exit Sub
                         Else
                              pp = List1.List(i - 1)
                              stp = Val(List1.List(i - 3))
                              ssl = Val(List1.List(i + 1))
                              btp = Val(List1.List(i + 2))
                              bsl = Val(List1.List(i - 2))
                              
                              
                              Form1.Text29.Text = pp
                              Form1.Text30.Text = pp
                              Form1.Text31.Text = btp
                              Form1.Text32.Text = stp
                              Form1.Text33.Text = bsl
                              Form1.Text34.Text = ssl
                              
                                                         
                              'MsgBox "above " & pp & " " & btp & " " & bsl & " " & stp & " " & ssl
                              Exit Sub
                         End If
                 End If
            Next i
        
        End If
        

                    



                    
End Sub


Public Sub Sort_num(number() As Double, count As Integer)

Dim inner, outer As Integer
Dim temp As Double
Dim min As Double, minindex As Integer
For outer = 0 To count
   min = number(outer)
   minindex = outer
   
    For inner = outer To count
    
        If number(inner) < min Then
             min = number(inner)
             minindex = inner
         End If
    Next inner
    temp = number(outer)
    number(outer) = min
    number(minindex) = temp
Next outer


End Sub




'calculate all pivot values
Public Sub pivot()
    Dim sum As Double, tlow As Double, thigh As Double, tclose As Double, temp1 As Double, temp2 As Double
    thigh = Text3.Text
    tlow = Text2.Text
    tclose = Text4.Text
      
    'get pivot ponit
    sum = thigh + tlow + tclose
    Text11.Text = sum / 3
    
    'get r6
    Text5.Text = thigh - (2 * (tlow - Text11.Text))
    
    'get r2
    Text9.Text = (Text11.Text * 2) - tlow
    
    'get r1
    temp1 = Text11.Text
    temp2 = Text9.Text
    sum = temp1 + temp2
    Text10.Text = sum / 2
    
    
    'get s6
    Text17.Text = tlow - (2 * (thigh - Text11.Text))
    
    'get s2
    Text13.Text = (Text11.Text * 2) - thigh
    
    
    'get s1
    temp1 = Text13.Text
    temp2 = Text11.Text
    Text12.Text = (temp1 + temp2) / 2
    
    'get r4
    Text7.Text = (Text11.Text - Text13.Text) + Text9.Text
    
    
    'get r3
    temp1 = Text7.Text
    temp2 = Text9.Text
    Text8.Text = (temp1 + temp2) / 2
    
    'get r5
    temp1 = Text5.Text
    temp2 = Text7.Text
    Text6.Text = (temp1 + temp2) / 2
    
    'get s4
    Text15.Text = Text11.Text - (Text9.Text - Text13.Text)
    
    'get s3
    temp1 = Text15.Text
    temp2 = Text13.Text
    Text14 = (temp1 + temp2) / 2
    
    'get s5
    temp1 = Text15.Text
    temp2 = Text17.Text
    Text16.Text = (temp1 + temp2) / 2
    
     'get r7
     Text35.Text = (Text11.Text + ((2 * thigh) - (2 * tlow)))
     
    
    'get r8
     Text37.Text = ((2 * Text11.Text) + (2 * thigh)) - (3 * tlow)
     
     
     'get s7
     Text36.Text = (Text11.Text - ((2 * thigh) - (2 * tlow)))
     
     
     'get s 8
      Text38.Text = (2 * Text11.Text - ((3 * thigh) - (2 * tlow)))
      
      
      If Form1.theminusvalue = 0.15 Then
            rounds2
      
      Else
      
        rounds4
      
      End If
End Sub

' 2 rounds
Public Sub rounds2()
    
       Text11.Text = FormatNumber(Round(Text11.Text, 2), 2)
    Text5.Text = FormatNumber(Round(Text5.Text, 2), 2)
    Text9.Text = FormatNumber(Round(Text9.Text, 2), 2)
    Text10.Text = FormatNumber(Round(Text10.Text, 2), 2)
    Text17.Text = FormatNumber(Round(Text17.Text, 2), 2)
    Text13.Text = FormatNumber(Round(Text13.Text, 2), 2)
    Text12.Text = FormatNumber(Round(Text12.Text, 2), 2)
    Text7.Text = FormatNumber(Round(Text7.Text, 2), 2)
    Text8.Text = FormatNumber(Round(Text8.Text, 2), 2)
    Text6.Text = FormatNumber(Round(Text6.Text, 2), 2)
    Text15.Text = FormatNumber(Round(Text15.Text, 2), 2)
    Text14.Text = FormatNumber(Round(Text14.Text, 2), 2)
    Text16.Text = FormatNumber(Round(Text16.Text, 2), 2)
   

    Text35.Text = FormatNumber(Round(Text35.Text, 2), 2)
    Text36.Text = FormatNumber(Round(Text36.Text, 2), 2)
    Text37.Text = FormatNumber(Round(Text37.Text, 2), 2)
    Text38.Text = FormatNumber(Round(Text38.Text, 2), 2)
     
End Sub



'4 rounds
Public Sub rounds4()
    
    Text11.Text = FormatNumber(Round(Text11.Text, 4), 4)
    Text5.Text = FormatNumber(Round(Text5.Text, 4), 4)
    Text9.Text = FormatNumber(Round(Text9.Text, 4), 4)
    Text10.Text = FormatNumber(Round(Text10.Text, 4), 4)
    Text17.Text = FormatNumber(Round(Text17.Text, 4), 4)
    Text13.Text = FormatNumber(Round(Text13.Text, 4), 4)
    Text12.Text = FormatNumber(Round(Text12.Text, 4), 4)
    Text7.Text = FormatNumber(Round(Text7.Text, 4), 4)
    Text8.Text = FormatNumber(Round(Text8.Text, 4), 4)
    Text6.Text = FormatNumber(Round(Text6.Text, 4), 4)
    Text15.Text = FormatNumber(Round(Text15.Text, 4), 4)
    Text14.Text = FormatNumber(Round(Text14.Text, 4), 4)
    Text16.Text = FormatNumber(Round(Text16.Text, 4), 4)
   

    Text35.Text = FormatNumber(Round(Text35.Text, 4), 4)
    Text36.Text = FormatNumber(Round(Text36.Text, 4), 4)
    Text37.Text = FormatNumber(Round(Text37.Text, 4), 4)
    Text38.Text = FormatNumber(Round(Text38.Text, 4), 4)
   
 

End Sub


Private Sub Command1_Click()
    MsgBox thelevelvalue
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Hide
End Sub
