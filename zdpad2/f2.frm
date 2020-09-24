VERSION 5.00
Begin VB.Form f2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "help"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "f2.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
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
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   0
      Picture         =   "f2.frx":0742
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "f2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
    f1.Show
    f2.Hide
End Sub
