VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form11 
   Caption         =   "Notepad"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   10995
   LinkTopic       =   "Form11"
   ScaleHeight     =   7455
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox Text1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   13361
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form11.frx":0000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11640
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu new 
         Caption         =   "New"
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
      Begin VB.Menu Open 
         Caption         =   "Open"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu find 
         Caption         =   "Find"
      End
   End
   Begin VB.Menu format 
      Caption         =   "Format"
      Begin VB.Menu font 
         Caption         =   "Font"
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String

Private Sub copy_Click()
a = Text1.SelText
End Sub

Private Sub cut_Click()
a = Text1.SelText
Text1.SelText = ""
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub find_Click()
a = (InputBox("find what", find))
Text1.find (a)
End Sub

Private Sub font_Click()
CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
CommonDialog1.ShowFont
Text1.SelFontName = CommonDialog1.FontName
Text1.SelFontSize = CommonDialog1.FontSize
End Sub

Private Sub new_Click()
Dim a As Integer
If Text1.Text <> "" Then
a = MsgBox("you wanna save this file", vbYesNoCancel + vbQuestion, "save")
If a = vbYes Then
CommonDialog1.Filter = "text files(*.txt,*.rtf)|*.txt;*rtf| All Files (*.*)|*.*"
CommonDialog1.ShowSave
Text1.SaveFile (CommonDialog1.FileName)
Text1.Text = ""
End If
If a = vbNo Then
Text1.Text = ""
End If
End If
End Sub

Private Sub Open_Click()
Dim a As Boolean
If Text1.Text <> "" Then

a = MsgBox("Do you want to save this file?", vbYesNoCancel + vbQuestion, "save")

If a = vbYes Then
CommonDialog1.Filter = "text files(*.txt,*.rtf)|*.txt;*rtf| All Files (*.*)|*.*"
CommonDialog1.ShowSave
Text1.SaveFile (CommonDialog1.FileName)
CommonDialog1.ShowOpen
Text1.LoadFile (CommonDialog1.FileName)
End If
If a = vbNo Then
CommonDialog1.Filter = "text files(*.txt,*.rtf)|*.txt;*rtf| All Files (*.*)|*.*"
CommonDialog1.ShowOpen
Text1.LoadFile (CommonDialog1.FileName)
End If
Else
  CommonDialog1.Filter = "text files(*.txt,*.rtf)|*.txt;*rtf| All Files (*.*)|*.*"
CommonDialog1.ShowOpen
Text1.LoadFile (CommonDialog1.FileName)
End If
End Sub

Private Sub paste_Click()
Text1.SelText = a
End Sub

Private Sub save_Click()
CommonDialog1.Filter = "text files(*.txt,*.rtf)|*.txt;*rtf| All Files (*.*)|*.*"
CommonDialog1.ShowSave
Text1.SaveFile (CommonDialog1.FileName)
End Sub

