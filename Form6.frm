VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Grocery Profile"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form6"
   ScaleHeight     =   14790
   ScaleWidth      =   18960
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
      Left            =   5280
      TabIndex        =   9
      Top             =   2400
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
      Height          =   1575
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3360
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
      Left            =   5280
      TabIndex        =   7
      Top             =   5280
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
      Left            =   5280
      TabIndex        =   6
      Top             =   6120
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
      Left            =   5280
      TabIndex        =   5
      Top             =   6960
      Width           =   5175
   End
   Begin VB.TextBox Text6 
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
      Left            =   5280
      TabIndex        =   4
      Top             =   7800
      Width           =   5175
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
      Left            =   1800
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   9120
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
      Left            =   4080
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   9120
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
      Left            =   6360
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   9120
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
      Left            =   8640
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "                                                                 ADD GROCERY PROFILE"
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
      Left            =   1920
      TabIndex        =   16
      Top             =   360
      Width           =   10695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter Your Grocery Name :-"
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
      Left            =   840
      TabIndex        =   15
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Adddress :-"
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
      TabIndex        =   14
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Phone Number :-"
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
      TabIndex        =   13
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E-mail :-"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Country :-"
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
      Left            =   3120
      TabIndex        =   11
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Registration Number :-"
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
      Left            =   1440
      TabIndex        =   10
      Top             =   7800
      Width           =   3375
   End
End
Attribute VB_Name = "Form6"
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
Text6.Text = ""
Text1.SetFocus
End Sub

Private Sub Command2_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" And Text6.Text = "" Then
   MsgBox "Please Fill All Fields Properly", vbInformation, "GMS"
   
   ElseIf Len(Text3.Text) < 10 Then
        MsgBox "Enter the phone number in 10 digits!", vbExclamation, ""
        Cancel = True
Else

 On Error GoTo ICanDealWithThis
 Dim cn As New ADODB.Connection
 Dim cmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
 cn.ConnectionString = strConn
 cn.Open

 strSQL = "INSERT INTO Profile([rname],[add],[phno],[email],[country],[rno]) VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "')"

 cmd.CommandText = strSQL
 cmd.CommandType = adCmdText
 cmd.ActiveConnection = cn
 cmd.Execute

 MsgBox "Record Added Successfully", vbInformation, "GMS"
 
 Set cmd = Nothing
 cn.Close
 Set cn = Nothing
 
 Exit Sub
 
ICanDealWithThis:
 MsgBox "Registration Number is Already Present", vbInformation, "GMS"
 
End If
End Sub

Private Sub Command3_Click()
Dim a As String
a = InputBox("Enter Registration Number", "GMS")

If a = "" Then
  MsgBox "Please Enter Registration Number", vbInformation, "GMS"
Else

 Dim adoConn As New ADODB.Connection
 Dim adoCmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 ' Open a Connection object
  strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"

  adoConn.ConnectionString = strConn

  adoConn.Open

  ' Define a query string
  strSQL = "DELETE FROM Profile WHERE[rno]='" & a & "'"

 ' Set up the Command object
 adoCmd.CommandText = strSQL
 adoCmd.CommandType = adCmdText

 adoCmd.ActiveConnection = adoConn

 adoCmd.Execute
 MsgBox "Record Deleted Successfully", vbInformation, "GMS"

 ' Tidy up
 Set adoCmd = Nothing
 adoConn.Close
 Set adoConn = Nothing

End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Text1_Change()
If IsNumeric(Right(Text1, 1)) = True Then
' is a number
Text1 = Left(Text1, Len(Text1) - 1)
MsgBox "No numbers allowed"
Else
'anything but a number
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKey0 To vbKey9
  Case vbKeyBack, vbKeyClear, vbKeyDelete
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
    KeyAscii = 0
    Beep
End Select
End Sub

Private Sub Text5_Change()
If IsNumeric(Right(Text5, 1)) = True Then
' is a number
Text5 = Left(Text5, Len(Text5) - 1)
MsgBox "No numbers allowed"
Else
'anything but a number
End If
End Sub
