VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Login Form"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   3
      Top             =   4560
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   8280
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5640
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Height          =   1095
      Left            =   5880
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Height          =   1095
      Left            =   8040
      Picture         =   "Form1.frx":09A5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   1695
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      ScaleHeight     =   315
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   8280
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "                                                              GROCERY MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1800
      TabIndex        =   6
      Top             =   1920
      Width           =   11535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Username  :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   5
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Password   :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   4
      Top             =   5520
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If (Text1.Text = "ad" And Text2.Text = "ad") Then
Me.Hide
Form2.Show
Text1.Text = ""
Text2.Text = ""
Else
MsgBox ("Incorrect Password")
Text1.Text = " "
Text2.Text = " "
End If
End Sub

Private Sub Command2_Click()
End
End Sub

