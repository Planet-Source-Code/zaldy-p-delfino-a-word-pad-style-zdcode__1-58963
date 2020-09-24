VERSION 5.00
Begin VB.Form FORMOREINFO 
   BackColor       =   &H80000012&
   Caption         =   "Form3"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      BackColor       =   &H80000017&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000017&
      Caption         =   "JUST SEARCH"
      BeginProperty Font 
         Name            =   "Magneto"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000017&
      Caption         =   "FOR MORE INFO"
      BeginProperty Font 
         Name            =   "Magneto"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "FORMOREINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()
    f1.Show
    FORMOREINFO.Hide
End Sub
