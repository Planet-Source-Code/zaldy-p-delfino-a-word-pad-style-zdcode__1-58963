VERSION 5.00
Begin VB.Form formsplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5550
   ClientLeft      =   2415
   ClientTop       =   1995
   ClientWidth     =   7440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleMode       =   0  'User
   ScaleWidth      =   7380
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Timer Timer1 
         Interval        =   1500
         Left            =   360
         Top             =   4440
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ZALDY WORD PAD"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1335
         Left            =   840
         TabIndex        =   4
         Top             =   1320
         Width           =   5535
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copyright: Zaldy P. Delfino"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   3360
         Width           =   3855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   " WINDOWS 95/98/2000/NT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   3000
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   3480
         TabIndex        =   1
         Top             =   2040
         Width           =   2775
      End
   End
End
Attribute VB_Name = "formsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Label4.Caption = "ZALDY WORD PAD"
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
Unload Me
Form1.Show
End Sub
