VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form progres 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pb 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "progres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim status As Integer

Private Sub Timer1_Timer()
x = x + 5
status = pb.Value
Label2.Caption = Str(status)
Label2.Refresh
pb.Value = pb.Value + 5
If x = 100 Then
Timer1.Enabled = False
Unload Me
wordPad.Show
End If

End Sub
