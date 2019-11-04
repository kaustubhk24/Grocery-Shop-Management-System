VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Bill Form"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16935
   LinkTopic       =   "Form9"
   ScaleHeight     =   10950
   ScaleWidth      =   16935
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000B&
      Caption         =   "NEW"
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
      Left            =   2520
      MaskColor       =   &H00000000&
      TabIndex        =   94
      Top             =   8640
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000B&
      Caption         =   "SAVE"
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
      Left            =   6120
      MaskColor       =   &H00000000&
      TabIndex        =   93
      Top             =   8640
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000B&
      Caption         =   "CLOSE"
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
      Left            =   9360
      MaskColor       =   &H00000000&
      TabIndex        =   92
      Top             =   8640
      Width           =   2055
   End
   Begin VB.TextBox Text39 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   90
      Top             =   8760
      Width           =   2055
   End
   Begin VB.TextBox Text38 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   88
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox Text37 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   86
      Top             =   7560
      Width           =   2055
   End
   Begin VB.TextBox Text36 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   85
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox Text35 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   84
      Top             =   6360
      Width           =   2055
   End
   Begin VB.TextBox Text34 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   83
      Top             =   5760
      Width           =   2055
   End
   Begin VB.TextBox Text33 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   82
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox Text32 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   81
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text31 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   80
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text30 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   79
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text29 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   78
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text28 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   77
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text27 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   76
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text26 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   65
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox Text25 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   64
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox Text24 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   63
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text23 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   62
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text22 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   61
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   60
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text20 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   59
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   58
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   57
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   56
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text16 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   45
      Top             =   6960
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   44
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   43
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   42
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   41
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   40
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   39
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   38
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   37
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   36
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox Combo11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form9.frx":0000
      Left            =   9000
      List            =   "Form9.frx":0002
      TabIndex        =   34
      Top             =   6960
      Width           =   2295
   End
   Begin VB.ComboBox Combo10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form9.frx":0004
      Left            =   9000
      List            =   "Form9.frx":0006
      TabIndex        =   32
      Top             =   6360
      Width           =   2295
   End
   Begin VB.ComboBox Combo9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form9.frx":0008
      Left            =   9000
      List            =   "Form9.frx":000A
      TabIndex        =   30
      Top             =   5760
      Width           =   2295
   End
   Begin VB.ComboBox Combo8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form9.frx":000C
      Left            =   9000
      List            =   "Form9.frx":000E
      TabIndex        =   28
      Top             =   5160
      Width           =   2295
   End
   Begin VB.ComboBox Combo7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form9.frx":0010
      Left            =   9000
      List            =   "Form9.frx":0012
      TabIndex        =   26
      Top             =   4560
      Width           =   2295
   End
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form9.frx":0014
      Left            =   9000
      List            =   "Form9.frx":0016
      TabIndex        =   24
      Top             =   3960
      Width           =   2295
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form9.frx":0018
      Left            =   9000
      List            =   "Form9.frx":001A
      TabIndex        =   22
      Top             =   3360
      Width           =   2295
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form9.frx":001C
      Left            =   9000
      List            =   "Form9.frx":001E
      TabIndex        =   20
      Top             =   2760
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form9.frx":0020
      Left            =   9000
      List            =   "Form9.frx":0022
      TabIndex        =   18
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form9.frx":0024
      Left            =   9000
      List            =   "Form9.frx":0026
      TabIndex        =   16
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   14
      Top             =   6480
      Width           =   5175
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   5880
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "SEARCH"
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
      Left            =   3480
      MaskColor       =   &H00000000&
      TabIndex        =   11
      Top             =   5160
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form9.frx":0028
      Left            =   2160
      List            =   "Form9.frx":002A
      TabIndex        =   4
      Top             =   2760
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   1560
      Width           =   5175
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
      Height          =   495
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   5175
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3360
      Width           =   5175
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   4560
      Width           =   5175
   End
   Begin VB.Label Label42 
      BackColor       =   &H00C0FFFF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16320
      TabIndex        =   95
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label Label41 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Net Pay    ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   91
      Top             =   8760
      Width           =   1695
   End
   Begin VB.Label Label40 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Discount   ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   89
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label39 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Total       ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   87
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label38 
      BackColor       =   &H00C0FFFF&
      Caption         =   "="
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
      Left            =   14400
      TabIndex        =   75
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label37 
      BackColor       =   &H00C0FFFF&
      Caption         =   "="
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
      Left            =   14400
      TabIndex        =   74
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C0FFFF&
      Caption         =   "="
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
      Left            =   14400
      TabIndex        =   73
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label35 
      BackColor       =   &H00C0FFFF&
      Caption         =   "="
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
      Left            =   14400
      TabIndex        =   72
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label34 
      BackColor       =   &H00C0FFFF&
      Caption         =   "="
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
      Left            =   14400
      TabIndex        =   71
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00C0FFFF&
      Caption         =   "="
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
      Left            =   14400
      TabIndex        =   70
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0FFFF&
      Caption         =   "="
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
      Left            =   14400
      TabIndex        =   69
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0FFFF&
      Caption         =   "="
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
      Left            =   14400
      TabIndex        =   68
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0FFFF&
      Caption         =   "="
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
      Left            =   14400
      TabIndex        =   67
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0FFFF&
      Caption         =   "="
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
      Left            =   14400
      TabIndex        =   66
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   12480
      TabIndex        =   55
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   12480
      TabIndex        =   54
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   12480
      TabIndex        =   53
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   12480
      TabIndex        =   52
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   12480
      TabIndex        =   51
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   12480
      TabIndex        =   50
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   12480
      TabIndex        =   49
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   12480
      TabIndex        =   48
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   12480
      TabIndex        =   47
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   12480
      TabIndex        =   46
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Item 10:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   35
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Item 9:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   33
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Item 8:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   31
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Item 7:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   29
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Item 6:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   27
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Item 5:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   25
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Item 4:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   23
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Item 3:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   21
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Item 2:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   19
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Item 1:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   17
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Time :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   15
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Date :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   13
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   $"Form9.frx":002C
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
      Left            =   -1920
      TabIndex        =   10
      Top             =   0
      Width           =   20655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ID :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Name  :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gender :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Address :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Phone No :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo10_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM Item where [iname]='" & Combo10.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text15.Text = adoRS.Fields("price").Value
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing
End Sub

Private Sub Combo11_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM Item where [iname]='" & Combo11.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text16.Text = adoRS.Fields("price").Value
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing
End Sub

Private Sub Combo2_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM Item where [iname]='" & Combo2.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text7.Text = adoRS.Fields("price").Value
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing
End Sub

Private Sub Combo3_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM Item where [iname]='" & Combo3.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text8.Text = adoRS.Fields("price").Value
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing
End Sub

Private Sub Combo4_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM Item where [iname]='" & Combo4.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text9.Text = adoRS.Fields("price").Value
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing
End Sub

Private Sub Combo5_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM Item where [iname]='" & Combo5.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text10.Text = adoRS.Fields("price").Value
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing
End Sub

Private Sub Combo6_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM Item where [iname]='" & Combo6.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text11.Text = adoRS.Fields("price").Value
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing
End Sub

Private Sub Combo7_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM Item where [iname]='" & Combo7.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text12.Text = adoRS.Fields("price").Value
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing
End Sub

Private Sub Combo8_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM Item where [iname]='" & Combo8.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text13.Text = adoRS.Fields("price").Value
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing
End Sub

Private Sub Combo9_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM Item where [iname]='" & Combo9.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text14.Text = adoRS.Fields("price").Value
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing
End Sub

Private Sub Command1_Click()
Dim a As String
a = InputBox("Enter Customer ID", "GMS")

If a = "" Then
  MsgBox "Please Enter Customer ID", vbInformation, "GMS"
Else

 Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM customer where [id]='" & a & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text2.Text = adoRS.Fields("name").Value
     Text3.Text = adoRS.Fields("add").Value
     Text4.Text = adoRS.Fields("phno").Value
     Combo1.Text = adoRS.Fields("gender").Value
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing

End If
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text37.Text = "" Or Text38.Text = "" Or Text39.Text = "" Or Combo1.Text = "" Then
   MsgBox "Please Fill All Fields Properly", vbInformation, "GMS"
   
Else

 On Error GoTo ICanDealWithThis
 Dim cn As New ADODB.Connection
 Dim cmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
 cn.ConnectionString = strConn
 cn.Open

 strSQL = "INSERT INTO Bill([id],[name],[gender],[add],[phno],[date],[time],[total],[dis],[netpay]) VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Combo1.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text37.Text & "','" & Text38.Text & "','" & Text39.Text & "')"

 cmd.CommandText = strSQL
 cmd.CommandType = adCmdText
 cmd.ActiveConnection = cn
 cmd.Execute

 MsgBox "Record Added Successfully", vbInformation, "GMS"
 
 Set cmd = Nothing
 cn.close
 Set cn = Nothing
 
 Exit Sub
 
ICanDealWithThis:
 MsgBox "ID is Already Present", vbInformation, "GMS"
 
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Text1.Text = Val(Text1.Text) + 1
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""

Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Text21.Text = ""
Text22.Text = ""
Text23.Text = ""
Text24.Text = ""
Text25.Text = ""
Text26.Text = ""
Text27.Text = ""

Text28.Text = ""
Text29.Text = ""
Text30.Text = ""
Text31.Text = ""
Text32.Text = ""
Text33.Text = ""
Text34.Text = ""
Text35.Text = ""
Text36.Text = ""
Text37.Text = ""
Text38.Text = ""
Text39.Text = ""

Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo5.Text = ""
Combo6.Text = ""
Combo7.Text = ""
Combo8.Text = ""
Combo9.Text = ""
Combo10.Text = ""
Combo11.Text = ""



End Sub

Private Sub Form_Load()
Text5.Text = DateValue(Now)
Text6.Text = TimeValue(Now)
'Text5.Text = format(Now, "dd/MMM/yyyy")

Combo1.AddItem ("Male")
Combo1.AddItem ("Female")

Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM Item"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Combo2.AddItem (adoRS.Fields("iname").Value)
     Combo3.AddItem (adoRS.Fields("iname").Value)
     Combo4.AddItem (adoRS.Fields("iname").Value)
     Combo5.AddItem (adoRS.Fields("iname").Value)
     Combo6.AddItem (adoRS.Fields("iname").Value)
     Combo7.AddItem (adoRS.Fields("iname").Value)
     Combo8.AddItem (adoRS.Fields("iname").Value)
     Combo9.AddItem (adoRS.Fields("iname").Value)
     Combo10.AddItem (adoRS.Fields("iname").Value)
     Combo11.AddItem (adoRS.Fields("iname").Value)
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing

Dim adoConn1 As New ADODB.Connection
Dim adoRS1 As New ADODB.Recordset
Dim strConn1 As String
strConn1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn1.ConnectionString = strConn1
adoConn1.Open
adoRS1.Source = "SELECT * FROM Bill"
adoRS1.CursorType = adOpenForwardOnly
adoRS1.ActiveConnection = adoConn1
adoRS1.Open
Do While Not adoRS1.EOF
     Text1.Text = adoRS1.Fields("id").Value + 1
     adoRS1.MoveNext
Loop
adoRS1.close
Set adoRS1 = Nothing
adoConn1.close
Set adoConn1 = Nothing
End Sub

Private Sub Text17_Change()
Text27 = Val(Text7.Text) * Val(Text17.Text)
End Sub

Private Sub Text18_Change()
Text28 = Val(Text8.Text) * Val(Text18.Text)
End Sub

Private Sub Text19_Change()
Text29 = Val(Text9.Text) * Val(Text19.Text)
End Sub

Private Sub Text20_Change()
Text30 = Val(Text10.Text) * Val(Text20.Text)
End Sub

Private Sub Text21_Change()
Text31 = Val(Text11.Text) * Val(Text21.Text)
End Sub

Private Sub Text22_Change()
Text32 = Val(Text12.Text) * Val(Text22.Text)
End Sub

Private Sub Text23_Change()
Text33 = Val(Text13.Text) * Val(Text23.Text)
End Sub

Private Sub Text24_Change()
Text34 = Val(Text14.Text) * Val(Text24.Text)
End Sub

Private Sub Text25_Change()
Text35 = Val(Text15.Text) * Val(Text25.Text)
End Sub

Private Sub Text26_Change()
Text36 = Val(Text16.Text) * Val(Text26.Text)
End Sub

Private Sub Text27_Change()
Text37.Text = Val(Text37.Text) + Val(Text27.Text)
End Sub

Private Sub Text28_Change()
Text37.Text = Val(Text37.Text) + Val(Text28.Text)
End Sub

Private Sub Text29_Change()
Text37.Text = Val(Text37.Text) + Val(Text29.Text)
End Sub

Private Sub Text30_Change()
Text37.Text = Val(Text37.Text) + Val(Text30.Text)
End Sub

Private Sub Text31_Change()
Text37.Text = Val(Text37.Text) + Val(Text31.Text)
End Sub

Private Sub Text32_Change()
Text37.Text = Val(Text37.Text) + Val(Text32.Text)
End Sub

Private Sub Text33_Change()
Text37.Text = Val(Text37.Text) + Val(Text33.Text)
End Sub

Private Sub Text34_Change()
Text37.Text = Val(Text37.Text) + Val(Text34.Text)
End Sub

Private Sub Text35_Change()
Text37.Text = Val(Text37.Text) + Val(Text35.Text)
End Sub

Private Sub Text36_Change()
Text37.Text = Val(Text37.Text) + Val(Text36.Text)
End Sub

Private Sub Text38_Change()
Text39.Text = Val(Text37.Text) * Val(Text38.Text) / 100
Text39.Text = Val(Text37.Text) - Val(Text39.Text)
End Sub
