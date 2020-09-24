VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RichTextBox3 
      Height          =   450
      Left            =   330
      TabIndex        =   2
      Top             =   2445
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   794
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0000
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   510
      Left            =   1995
      TabIndex        =   1
      Top             =   2310
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   900
      _Version        =   393217
      TextRTF         =   $"Form1.frx":00AE
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   930
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1640
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Form1.frx":015C
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Resize()
 
    RichTextBox1.Width = ScaleWidth
     RichTextBox1.Height = ScaleHeight
End Sub


Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then 'If right mouse button is clicked
        PopupMenu MDIForm1.menupopup 'show popup menu
    End If
    
End Sub
