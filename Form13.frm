VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00C0FFFF&
   Caption         =   "About"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12225
   LinkTopic       =   "Form13"
   ScaleHeight     =   6180
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5280
      TabIndex        =   0
      Top             =   5280
      Width           =   1500
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Version              :- 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2520
      TabIndex        =   4
      Top             =   2520
      Width           =   5445
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Application Title :-  GROCERY       MANAGEMENT                                       SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   7845
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"Form13.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1890
      Left            =   2520
      TabIndex        =   2
      Top             =   3120
      Width           =   7725
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   12120
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "                                                                        ABOUT MY APP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12255
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub
