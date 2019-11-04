VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Customer Information"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12795
   LinkTopic       =   "Form7"
   ScaleHeight     =   5970
   ScaleWidth      =   12795
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   9120
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   7800
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
      Left            =   6840
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   7800
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
      Left            =   4560
      MaskColor       =   &H00000000&
      TabIndex        =   6
      Top             =   7800
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
      Left            =   2280
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   7800
      Width           =   2055
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
      Left            =   5640
      TabIndex        =   4
      Top             =   6360
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
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4920
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
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3240
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
      Left            =   5640
      TabIndex        =   1
      Top             =   2280
      Width           =   5175
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
      ItemData        =   "Form7.frx":0000
      Left            =   5640
      List            =   "Form7.frx":0002
      TabIndex        =   0
      Top             =   4080
      Width           =   5175
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
      Left            =   3360
      TabIndex        =   14
      Top             =   6360
      Width           =   1815
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
      Left            =   3360
      TabIndex        =   13
      Top             =   5400
      Width           =   1575
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
      Left            =   3480
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
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
      Left            =   3720
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
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
      Left            =   3960
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "                                                                     CUSTOMER MASTER"
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
      Left            =   1560
      TabIndex        =   9
      Top             =   360
      Width           =   12015
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = Val(Text1.Text) + 1
Text2.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text5.Text = ""
Text2.SetFocus
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text5.Text = "" Or Combo1.Text = "" Then
   MsgBox "Please Fill All Fields Properly", vbInformation, "GMS"
   
   ElseIf Len(Text5.Text) < 10 Then
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

 strSQL = "INSERT INTO customer([id],[name],[gender],[add],[phno]) VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Combo1.Text & "','" & Text3.Text & "','" & Text5.Text & "')"

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

Private Sub Command3_Click()
Dim a As String
a = InputBox("Enter ID", "GMS")

If a = "" Then
  MsgBox "Please Enter ID", vbInformation, "GMS"
Else

 Dim adoConn As New ADODB.Connection
 Dim adoCmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 ' Open a Connection object
  strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"

  adoConn.ConnectionString = strConn

  adoConn.Open

  ' Define a query string
  strSQL = "DELETE FROM customer WHERE[id]='" & a & "'"

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

Private Sub Form_Load()
Combo1.AddItem ("Male")
Combo1.AddItem ("Female")

Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GMSDB.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM customer"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text1.Text = adoRS.Fields("id").Value + 1
     adoRS.MoveNext
Loop
adoRS.close
Set adoRS = Nothing
adoConn.close
Set adoConn = Nothing
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKey0 To vbKey9
  Case vbKeyBack, vbKeyClear, vbKeyDelete
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
    KeyAscii = 0
    Beep
End Select
End Sub
