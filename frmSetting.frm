VERSION 5.00
Begin VB.Form frmSetting 
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optcomm 
      Caption         =   "COMM6"
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
      Left            =   5715
      TabIndex        =   6
      Top             =   4470
      Width           =   2220
   End
   Begin VB.OptionButton optcomm 
      Caption         =   "COMM 5"
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
      Left            =   5700
      TabIndex        =   5
      Top             =   3780
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
      Left            =   5700
      TabIndex        =   4
      Top             =   3105
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
      Left            =   7740
      TabIndex        =   1
      Top             =   5730
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
      Left            =   5145
      TabIndex        =   0
      Top             =   5760
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   4950
      TabIndex        =   2
      Top             =   2145
      Width           =   4380
      Begin VB.OptionButton optcomm 
         Caption         =   "COMM 1"
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
         Left            =   735
         TabIndex        =   3
         Top             =   285
         Width           =   2220
      End
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Select Case Port
        Case 0
            Form3.MSComm1.CommPort = 1
        Case 1
            Form3.MSComm1.CommPort = 2
        Case 2
            Form3.MSComm1.CommPort = 5
        Case 3
            Form3.MSComm1.CommPort = 6
    End Select
    Form3.OK.Enabled = True
    Unload Me
End Sub

Private Sub Command2_Click()
    Form3.OK.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    For i = 0 To 3
        If comm(i) = True Then
            optcomm(i).Enabled = True
        Else
            optcomm(i).Enabled = False
        End If
    Next i
    Select Case Form3.MSComm1.CommPort
        Case 1
            optcomm(0).Value = True
        Case 2
            optcomm(1).Value = True
        Case 5
            optcomm(2).Value = True
        Case 6
            optcomm(3).Value = True
    End Select
    frmSetting.Show
End Sub

Private Sub optcomm_Click(Index As Integer)
    Port = Index
End Sub
