VERSION 5.00
Begin VB.Form closedocument 
   Caption         =   "Document close"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   5565
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Complex"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   0
      Picture         =   "closedocument.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "closedocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()
        f1.Show
        closedocument.Hide
End Sub
