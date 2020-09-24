VERSION 5.00
Begin VB.Form f1 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ZALDY WORD HELP!"
   ClientHeight    =   6495
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4980
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   0
      Picture         =   "Form3.frx":0442
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5055
   End
   Begin VB.Menu menucontent 
      Caption         =   "CONTENTS.."
      Begin VB.Menu cmdintro 
         Caption         =   "Introductios.."
      End
      Begin VB.Menu cmdwork 
         Caption         =   "Working With Documents"
         Begin VB.Menu cmdcreate 
            Caption         =   "Create a New Document"
         End
         Begin VB.Menu cmdexis 
            Caption         =   "Open an Existing Document"
         End
         Begin VB.Menu cmdinfo 
            Caption         =   "For More Info?"
         End
      End
   End
   Begin VB.Menu menuexit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "f1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdinfo_Click()
    FORMOREINFO.Show
    f1.Hide
End Sub



Private Sub cmd_Click()

End Sub

Private Sub cmdcreate_Click()
    f2.Show
    f1.Hide
End Sub

Private Sub cmdexis_Click()
    F3.Show
    f1.Hide
End Sub

Private Sub cmdintro_Click()
    MsgBox "Welcome to Zdpad2", vbOKOnly, "zdpad"
End Sub

Private Sub menuexit_Click()
    Unload Me
End Sub


