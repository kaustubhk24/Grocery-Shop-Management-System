VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Grocery Management System"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10785
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Wholesale Groceries Store Management System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   8415
      Left            =   5160
      TabIndex        =   0
      Top             =   600
      Width           =   8775
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu lo 
         Caption         =   "LogOut"
      End
      Begin VB.Menu close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu master 
      Caption         =   "Master"
      Begin VB.Menu as 
         Caption         =   "Add Supplier"
      End
      Begin VB.Menu ai 
         Caption         =   "Add Item"
      End
      Begin VB.Menu gp 
         Caption         =   "Add Grocery Profile"
      End
      Begin VB.Menu ac 
         Caption         =   "Add Customer"
      End
      Begin VB.Menu ae 
         Caption         =   "Add Employee"
      End
   End
   Begin VB.Menu transaction 
      Caption         =   "Transaction"
      Begin VB.Menu bf 
         Caption         =   "Bill Form"
      End
   End
   Begin VB.Menu report 
      Caption         =   "Report"
      Begin VB.Menu sr 
         Caption         =   "Supplier Report"
      End
      Begin VB.Menu ir 
         Caption         =   "Item Report"
      End
      Begin VB.Menu cr 
         Caption         =   "Customer Report"
      End
      Begin VB.Menu er 
         Caption         =   "Employee Report"
      End
   End
   Begin VB.Menu tools 
      Caption         =   "Tools"
      Begin VB.Menu cal 
         Caption         =   "Calculator"
      End
      Begin VB.Menu set 
         Caption         =   "Setup User"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub about_Click()
Form3.Show
End Sub

Private Sub ac_Click()
Form7.Show
End Sub

Private Sub ae_Click()
Form8.Show
End Sub

Private Sub ai_Click()
Form5.Show
End Sub

Private Sub as_Click()
Form4.Show
End Sub

Private Sub bf_Click()
Form9.Show
End Sub

Private Sub br_Click()
DataReport1.Show
End Sub

Private Sub cal_Click()
Form10.Show
End Sub

Private Sub close_Click()
End
End Sub

Private Sub cr_Click()
DataReport4.Show
End Sub

Private Sub db_Click()
Form3.Show
End Sub

Private Sub er_Click()
DataReport5.Show
End Sub

Private Sub Form_Load()
Form3.Show
End Sub

Private Sub gp_Click()
Form6.Show
End Sub

Private Sub ir_Click()
DataReport3.Show
End Sub

Private Sub lo_Click()
Form1.Show
End Sub

Private Sub note_Click()
Form11.Show
End Sub

Private Sub set_Click()
Form12.Show
End Sub

Private Sub sr_Click()
DataReport2.Show
End Sub

Private Sub ts_Click()
Form12.Show
End Sub
