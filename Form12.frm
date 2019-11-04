VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Setup User"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12165
   LinkTopic       =   "Form12"
   ScaleHeight     =   6375
   ScaleWidth      =   12165
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
      Height          =   495
      Left            =   5160
      TabIndex        =   8
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
      Left            =   5160
      TabIndex        =   7
      Top             =   2280
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
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   3000
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
      IMEMode         =   3  'DISABLE
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3720
      Width           =   5175
   End
   Begin VB.TextBox Text5 
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
      IMEMode         =   3  'DISABLE
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4440
      Width           =   5175
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
      Left            =   8280
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000B&
      Caption         =   "DELETE"
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
      Left            =   5880
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   5400
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
      Left            =   3600
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
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
      Left            =   1320
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "                                                                                   CREATE USER"
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
      TabIndex        =   14
      Top             =   0
      Width           =   12135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "First Name  :-"
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
      Left            =   2880
      TabIndex        =   13
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Last Name  :-"
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
      Left            =   2880
      TabIndex        =   12
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "User Name  :-"
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
      Left            =   2760
      TabIndex        =   11
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Password  :-"
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
      Left            =   3000
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Confirm Password  :-"
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
      Left            =   1800
      TabIndex        =   9
      Top             =   4440
      Width           =   3255
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text1.SetFocus
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
   MsgBox "Please Fill All Fields Properly", vbInformation, "GMS"
   
   ElseIf Len(Text5.Text) < 3 Or Len(Text3.Text) < 3 Then
        MsgBox "Please enter minimum 3 character", vbExclamation, ""
        Cancel = True
        ElseIf Text4.Text <> Text5.Text Then
        MsgBox "Please enter correct Confirmation Password", vbInformation, "GMS"
Else

 On Error GoTo ICanDealWithThis
 Dim cn As New ADODB.Connection
 Dim cmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
 cn.ConnectionString = strConn
 cn.Open

 strSQL = "INSERT INTO login([fname],[lname],[username],[password]) VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text5.Text & "')"

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
 MsgBox "Username is Already Present", vbInformation, "GMS"
 
End If
End Sub

Private Sub Command3_Click()
Dim a As String
a = InputBox("Enter User Name", "GMS")

If a = "" Then
  MsgBox "Please Enter User Name", vbInformation, "GMS"
Else

 Dim adoConn As New ADODB.Connection
 Dim adoCmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 ' Open a Connection object
  strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"

  adoConn.ConnectionString = strConn

  adoConn.Open

  ' Define a query string
  strSQL = "DELETE FROM login WHERE[username]='" & a & "'"

 ' Set up the Command object
 adoCmd.CommandText = strSQL
 adoCmd.CommandType = adCmdText

 adoCmd.ActiveConnection = adoConn

 adoCmd.Execute
 MsgBox "Record Deleted Successfully", vbInformation, "GMS"

 ' Tidy up
 Set adoCmd = Nothing
 adoConn.close
 Set adoConn = Nothing

End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

