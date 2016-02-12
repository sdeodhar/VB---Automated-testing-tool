VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form4 
   Caption         =   "Form3"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   17850
   LinkTopic       =   "Form3"
   ScaleHeight     =   9630
   ScaleWidth      =   17850
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   360
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   11760
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      RThreshold      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton OK 
      Caption         =   "SELECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   4320
      Width           =   2295
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Run Time Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   2880
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Board Configuration Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Config Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current Mode :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If MSComm1.PortOpen = True Then
       MSComm1.PortOpen = False
    End If
    End
End Sub

Private Sub Form_Load()

cmdstr = ""
ackflag = 0

End Sub

Private Sub MSComm1_OnComm()
   
    Select Case MSComm1.CommEvent
    
    Case comEvReceive
    Timer1.Enabled = False
    Charval = MSComm1.Input
    
    If Charval = Chr(10) Then
    'Charval = Chr(13)
    'cmdstr = ""
    End If
    
    If Charval <> Chr(13) Then
    cmdstr = cmdstr + Charval
    
    ElseIf Len(cmdstr) <> 0 And Asc(cmdstr) = 90 Then         'Z
    Form2.Text1(8).Text = Right(cmdstr, Len(cmdstr) - 2)
    Form2.Label3(8).Caption = "Y"
    cmdstr = ""
    ackflag = 1
    Form4.Label2 = "Config Mode"
    Form4.Option2.Enabled = False
    Form4.Option1.Enabled = True
    
    ElseIf Asc(cmdstr) = 83 And Len(cmdstr) <> 0 Then        'S
    Form2.Text1(9).Text = Right(cmdstr, Len(cmdstr) - 2)
    Form2.Label3(9).Caption = "Y"
    ackflag = 1
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 80 And Len(cmdstr) <> 0 Then        'P
    Form2.Text1(0).Text = Right(cmdstr, Len(cmdstr) - 2)
    Form2.Label3(0).Caption = "Y"
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 73 And Len(cmdstr) <> 0 Then        'I
    Form2.Text1(1).Text = Right(cmdstr, Len(cmdstr) - 2)
    Form2.Label3(1).Caption = "Y"
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 68 And Len(cmdstr) <> 0 Then        'D
    Form2.Text1(2).Text = Right(cmdstr, Len(cmdstr) - 2)
    Form2.Label3(2).Caption = "Y"
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 84 And Len(cmdstr) <> 0 Then        'T
    Form2.Text1(3).Text = Right(cmdstr, Len(cmdstr) - 2)
    Form2.Label3(3).Caption = "Y"
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 65 And Len(cmdstr) <> 0 Then        'A
    Form2.Text1(4).Text = Right(cmdstr, Len(cmdstr) - 2)
    Form2.Label3(4).Caption = "Y"
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 70 And Len(cmdstr) <> 0 Then        'F
    Form2.Text1(5).Text = Right(cmdstr, Len(cmdstr) - 2)
    Form2.Label3(5).Caption = "Y"
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 66 And Len(cmdstr) <> 0 Then        'B
    Form2.Text1(6).Text = Right(cmdstr, Len(cmdstr) - 2)
    Form2.Label3(6).Caption = "Y"
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 82 And Len(cmdstr) <> 0 Then        'R
    Form2.Text1(7).Text = Right(cmdstr, Len(cmdstr) - 2)
    Form2.Label3(7).Caption = "Y"
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 87 And Len(cmdstr) <> 0 Then        'Set Speed --> W
    Form1.Label4.Caption = Right(cmdstr, Len(cmdstr) - 2)
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 88 And Len(cmdstr) <> 0 Then        'Actal Speed --> X
    Form1.Label5.Caption = Right(cmdstr, Len(cmdstr) - 2)
    cmdstr = ""
    Runflag = 1
    
    ElseIf Asc(cmdstr) = 89 And Len(cmdstr) <> 0 Then        'Current Angle --> Y
    Form1.Label6.Caption = Right(cmdstr, Len(cmdstr) - 2)
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 69 And Len(cmdstr) <> 0 Then        'RNGERR  --> E
    Form2.Label3(curlab).Caption = "N"
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 71 And Len(cmdstr) <> 0 Then        'Success  --> G
    Form2.Label3(curlab).Caption = "Y"
    cmdstr = ""
    
    'ElseIf Asc(cmdstr) = 74 And Len(cmdstr) <> 0 Then        'Enter Config  --> J
    'Load Form2
    'Form2.Show
    
    ElseIf Asc(cmdstr) = 72 And Len(cmdstr) <> 0 Then        'Fail  --> H
    Form2.Label3(curlab).Caption = "N"  'Right(cmdstr, Len(cmdstr) - 1)
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 67 And Len(cmdstr) <> 0 Then        'BCMD --> C
    'Form2.Label3(curlab).Caption = "Bad" 'Right(cmdstr, Len(cmdstr) - 1)
    cmdstr = ""
    
    ElseIf Asc(cmdstr) = 78 And Len(cmdstr) <> 0 Then        'no crank --> N
    Form1.Label6.Caption = "No Crank"
    cmdstr = ""
    
    Else
    cmdstr = ""
    
    End If
    
    If ackflag = 0 Then
    Form2.Text1(0) = ""
    Form2.Text1(1) = ""
    Form2.Text1(2) = ""
    Form2.Text1(3) = ""
    Form2.Text1(4) = ""
    Form2.Text1(5) = ""
    Form2.Text1(6) = ""
    Form2.Text1(7) = ""
    Form2.Text1(8) = ""
    Form2.Text1(9) = ""
    End If
    
    If Runflag = 1 Then
    ackflag = 0
    Form4.Label2 = "Run Mode"
    Form2.Label5 = "Run Mode"
    Form2.Command1.Enabled = False
    Form2.Command2.Enabled = False
    Form2.Command3.Enabled = False
    Form2.Command4.Enabled = False
    Form2.Command5.Enabled = False
    Form2.Command6.Enabled = False
    Form2.Command7.Enabled = False
    Form2.Command8.Enabled = False
    Form2.Command9.Enabled = False
    Form2.Command10.Enabled = False
    Form2.Command11.Enabled = False
    Form2.Command12.Enabled = False
    Form2.Command13.Enabled = False
    Form4.Option1.Enabled = False
    Form4.Option2.Enabled = True
    Runflag = 0
    End If
           
    Case comEventFrame, comEventOverrun, comEventBreak
    'Timer1.Enabled = False
    MSComm1.PortOpen = False
    MSComm1.PortOpen = True
    MSComm1.InBufferCount = 0
    'Timer1.Enabled = True
    End Select

End Sub


Private Sub OK_Click()
    
    If Option1.Value = True Then
        Load Form2
        Form2.Show
        
    ElseIf Option2.Value = True Then
        Load Form2
        Form2.Show
        Load Form1
        Form1.Show
    
    End If

End Sub

Private Sub Timer1_Timer()
    
Dim intmsg As Integer
Timer1.Enabled = False
'    comstate = 0
If ackflag = 2 Then
intmsg = MsgBox("No connection to PID Controller Board", vbOKOnly)
End If

MSComm1.PortOpen = False

MSComm1.PortOpen = True
Form4.MSComm1.Output = Chr(81)
Form4.MSComm1.Output = Chr(13)
Timer1.Enabled = True

End Sub
