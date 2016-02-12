VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   11100
   ClientLeft      =   -15
   ClientTop       =   315
   ClientWidth     =   20400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11100
   ScaleWidth      =   20400
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   495
      Left            =   17040
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Height          =   2295
      Left            =   5880
      TabIndex        =   5
      Top             =   5160
      Width           =   6255
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   9360
      TabIndex        =   4
      Top             =   1440
      Width           =   7095
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   8055
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "ANGLE"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ACTUAL CURRENT SPEED"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10320
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SET SPEED"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MSComm1_OnComm()

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

