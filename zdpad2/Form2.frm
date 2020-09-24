VERSION 5.00
Begin VB.Form form2 
   BackColor       =   &H00000000&
   Caption         =   "INFO"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillColor       =   &H00400040&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4185
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "COMMENTS"
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   4335
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         X1              =   1200
         X2              =   3120
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "zd2381@yahoo.com"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000012&
         Caption         =   "E-MAIL AT:"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "PHILIPPINE COMPUTER LEARNING CENTER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   " DESIGNED BY:           ZALDY P. DELFINO      PROGRAMMER"
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "ZD WORD PAD"
      BeginProperty Font 
         Name            =   "Magneto"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   2880
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5


Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.BackColor = vbWhite
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.BackColor = vbGreen
End Sub



Private Sub Timer1_Timer()

If Label4.ForeColor = vbRed Then
    Label4.ForeColor = vbGreen
Else
    Label4.ForeColor = vbRed
End If


End Sub



Private Sub Label5_Click()
    ShellExecute hwnd, "open", "mailto:zd2381@yahoo.com", vbNullString, vbNullString, SW_SHOW
End Sub
