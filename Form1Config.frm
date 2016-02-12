VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18420
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   18420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command14 
      Caption         =   "X"
      Height          =   495
      Left            =   12720
      TabIndex        =   46
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   33
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Write"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   32
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Write"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   31
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Write"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   30
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Write"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   29
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Write"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   28
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Write"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   27
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Write"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   26
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Write"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   25
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Write"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   24
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   23
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   8760
      TabIndex        =   22
      Text            =   "0"
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   2160
      TabIndex        =   21
      Text            =   "0"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   2160
      TabIndex        =   20
      Text            =   "0"
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   8760
      TabIndex        =   9
      Text            =   "0"
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   8760
      TabIndex        =   8
      Text            =   "0"
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   8760
      TabIndex        =   7
      Text            =   "0"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   8760
      TabIndex        =   6
      Text            =   "0"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Text            =   "0"
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Text            =   "0"
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Text            =   "0"
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   6120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Current Mode :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   44
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Config Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   45
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6120
      TabIndex        =   43
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   6120
      TabIndex        =   42
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   12000
      TabIndex        =   41
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   12000
      TabIndex        =   40
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   12000
      TabIndex        =   39
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   12000
      TabIndex        =   38
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   12000
      TabIndex        =   37
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6120
      TabIndex        =   36
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   35
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6120
      TabIndex        =   34
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Span"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   840
      TabIndex        =   19
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Zero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   840
      TabIndex        =   18
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Direction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6960
      TabIndex        =   17
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Dead Band"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   16
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Board Frequency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   6840
      TabIndex        =   15
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Max Angle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6840
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Number of Teeth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   6960
      TabIndex        =   13
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   12
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   10
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Configuration Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Form4.MSComm1.Output = Chr(90)
Form4.MSComm1.Output = Text1(8).Text
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
curlab = 8
End Sub

Private Sub Command13_Click()
   Unload Me
End Sub

Private Sub Command14_Click()
 If Form4.MSComm1.PortOpen = True Then
       Form4.MSComm1.PortOpen = False
    End If
    End
End Sub

Private Sub Command2_Click()

 'Form4.MSComm1.Output = Text1(0).Text
 Form4.MSComm1.Output = Chr(13)
 'Form4.MSComm1.Output = Chr(10)
   Load Form2
   Form2.Show
   Load Form1
   Form1.Show

End Sub

Private Sub Command3_Click()
Form4.MSComm1.Output = Chr(81)
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
End Sub

Private Sub Command4_Click()
Form4.MSComm1.Output = Chr(83)
Form4.MSComm1.Output = Text1(9).Text
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
curlab = 9
End Sub

Private Sub Command5_Click()
Form4.MSComm1.Output = Chr(80)
Form4.MSComm1.Output = Text1(0).Text
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
curlab = 0
End Sub

Private Sub Command6_Click()
Form4.MSComm1.Output = Chr(73)
Form4.MSComm1.Output = Text1(1).Text
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
curlab = 1
End Sub

Private Sub Command7_Click()
Form4.MSComm1.Output = Chr(68)
Form4.MSComm1.Output = Text1(2).Text
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
curlab = 2
End Sub

Private Sub Command8_Click()
Form4.MSComm1.Output = Chr(84)
Form4.MSComm1.Output = Text1(3).Text
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
curlab = 3
End Sub

Private Sub Command9_Click()
Form4.MSComm1.Output = Chr(65)
Form4.MSComm1.Output = Text1(4).Text
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
curlab = 4
End Sub
Private Sub Command10_Click()
Form4.MSComm1.Output = Chr(70)
Form4.MSComm1.Output = Text1(5).Text
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
curlab = 5
End Sub
Private Sub Command11_Click()
Form4.MSComm1.Output = Chr(66)
Form4.MSComm1.Output = Text1(6).Text
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
curlab = 6
End Sub
Private Sub Command12_Click()
Form4.MSComm1.Output = Chr(82)
Form4.MSComm1.Output = Text1(7).Text
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
curlab = 7
End Sub

Private Sub Form_Load()
Form4.MSComm1.Output = Chr(81)
Form4.MSComm1.Output = Chr(13)
'Form4.MSComm1.Output = Chr(10)
End Sub

