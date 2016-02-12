VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "COMM SELECT"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optcomm 
      Caption         =   "COMM 8"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   7
      Left            =   4200
      TabIndex        =   10
      Top             =   4560
      Width           =   2220
   End
   Begin VB.OptionButton optcomm 
      Caption         =   "COMM 7"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   6
      Left            =   4200
      TabIndex        =   9
      Top             =   4080
      Width           =   2220
   End
   Begin VB.OptionButton optcomm 
      Caption         =   "COMM 6"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   5
      Left            =   4200
      TabIndex        =   8
      Top             =   3600
      Width           =   2220
   End
   Begin VB.OptionButton optcomm 
      Caption         =   "COMM 5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   4
      Left            =   4200
      TabIndex        =   7
      Top             =   3120
      Width           =   2220
   End
   Begin VB.OptionButton optcomm 
      Caption         =   "COMM 4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   3
      Left            =   4200
      TabIndex        =   6
      Top             =   2640
      Width           =   2220
   End
   Begin VB.OptionButton optcomm 
      Caption         =   "COMM 3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   4200
      TabIndex        =   5
      Top             =   2160
      Width           =   2220
   End
   Begin VB.OptionButton optcomm 
      Caption         =   "COMM 2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   4200
      TabIndex        =   4
      Top             =   1680
      Width           =   2220
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5520
      TabIndex        =   1
      Top             =   6000
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "APPLY"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2880
      TabIndex        =   0
      Top             =   6000
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   3540
      Begin VB.OptionButton optcomm 
         Caption         =   "COMM 1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   2220
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Form4.MSComm1.CommPort = selectedPort
    Form4.MSComm1.PortOpen = True
    Form4.OK.Enabled = True
    
    Form4.MSComm1.InputLen = 1
    Form4.MSComm1.RThreshold = 1
        
    ackflag = 2
    Form4.Timer1.Enabled = True
    Form4.MSComm1.Output = Chr(81)
    Form4.MSComm1.Output = Chr(13)
    
    Load Form4
    Form4.Show
    Unload Me
    
End Sub

Private Sub Command2_Click()
       If Form4.MSComm1.PortOpen = True Then
       Form4.MSComm1.PortOpen = False
    End If
    End
End Sub

Private Sub Form_Load()
CommCheck
End Sub

Private Sub optcomm_Click(Index As Integer)
    selectedPort = Index + 1
    
    For i = 0 To 7 '30
        If i <> Index Then
            Form3.optcomm(i).Value = False
        End If
    Next
    
End Sub
