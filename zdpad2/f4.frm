VERSION 5.00
Begin VB.Form f4 
   Caption         =   "INTRODUCTION"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   6240
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
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
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   6255
      Left            =   0
      Picture         =   "f4.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "f4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
    f4.Hide
    f1.Show
End Sub
