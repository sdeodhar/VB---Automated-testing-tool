VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1935
      Left            =   3840
      TabIndex        =   5
      Top             =   5040
      Width           =   5775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1335
      Left            =   7080
      TabIndex        =   4
      Top             =   2400
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   960
      TabIndex        =   3
      Top             =   2280
      Width           =   4935
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2640
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "ANGLE"
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
      Left            =   4920
      TabIndex        =   2
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ACTUAL CURRENT SPEED"
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
      Left            =   6960
      TabIndex        =   1
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SET SPEED"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
