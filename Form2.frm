VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12585
   LinkTopic       =   "Form2"
   ScaleHeight     =   7380
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
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
      Left            =   5160
      TabIndex        =   23
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   7
      Left            =   8760
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   4560
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   9
      Left            =   2280
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   8
      Left            =   2280
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   6
      Left            =   8760
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3840
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   5
      Left            =   8760
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3120
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   4
      Left            =   8760
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   8760
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4440
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3840
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3240
      Width           =   3615
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
      Left            =   9000
      TabIndex        =   1
      Top             =   6120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Program"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Span"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   19
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Zero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   18
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Direction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7200
      TabIndex        =   17
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Dead Band"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6960
      TabIndex        =   16
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Board Frequency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   6960
      TabIndex        =   15
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Max Angle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6960
      TabIndex        =   14
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Number of Teeth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   7320
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   12
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Configure Board"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
