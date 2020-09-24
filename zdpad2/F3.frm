VERSION 5.00
Begin VB.Form F3 
   Caption         =   "OPEN EXISTING FILE"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   LinkTopic       =   "Form3"
   ScaleHeight     =   6510
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "HOME"
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
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   0
      Picture         =   "F3.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "F3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
    f1.Show
    F3.Hide
End Sub
