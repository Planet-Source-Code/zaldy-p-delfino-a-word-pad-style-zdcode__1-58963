VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000F&
   Caption         =   "ZD PAD"
   ClientHeight    =   4935
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6465
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1800
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0AD4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar3 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   1380
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox picRuler 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   120
         Picture         =   "MDIForm1.frx":0D66
         ScaleHeight     =   270
         ScaleWidth      =   11490
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   11490
         Begin VB.PictureBox Picture1 
            Height          =   255
            Left            =   1200
            ScaleHeight     =   195
            ScaleWidth      =   435
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   1799
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   40
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Undeline"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Center"
            Object.ToolTipText     =   "Align Center"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Justify"
            Object.ToolTipText     =   "Justify"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font Color"
            Object.ToolTipText     =   "Font Color"
            ImageIndex      =   16
            Style           =   5
         EndProperty
         BeginProperty Button40 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Page Color"
            Object.ToolTipText     =   "Page Color"
            ImageIndex      =   17
            Style           =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "MDIForm1.frx":AF60
         Left            =   2640
         List            =   "MDIForm1.frx":AF88
         TabIndex        =   4
         Text            =   "10"
         Top             =   0
         Width           =   810
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "MDIForm1.frx":AFBC
         Left            =   240
         List            =   "MDIForm1.frx":AFBE
         TabIndex        =   3
         Text            =   "Arial"
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4680
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "For Help, Press F1"
            TextSave        =   "For Help, Press F1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2/17/2005"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:16 AM"
            Key             =   "date"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Key             =   "time"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AFC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B502
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":BA44
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":BF86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C4C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":CA0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":CF4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D48E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D5A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D6B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D7C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D8D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D9E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":DAFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":DC0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":DD1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E27E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menufile 
      Caption         =   "&File"
      Begin VB.Menu cmdnew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu cmdopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu cmdsave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu cmdsaveas 
         Caption         =   "Save as..."
         Shortcut        =   +^{F12}
      End
      Begin VB.Menu cmdfileline1 
         Caption         =   "-"
      End
      Begin VB.Menu cmdprint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu cmdprintpreview 
         Caption         =   "Print Preview"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu cmdprinsetup 
         Caption         =   "&Print Set up"
      End
      Begin VB.Menu cmdfileline3 
         Caption         =   "-"
      End
      Begin VB.Menu cmdexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu menuedit 
      Caption         =   "&Edit"
      Begin VB.Menu cmdundo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu cmdfileline2 
         Caption         =   "-"
      End
      Begin VB.Menu cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu cmddelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu cmdselectall 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu meneview 
      Caption         =   "&View"
      Begin VB.Menu cmdtoolbar 
         Caption         =   "&Tollbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu cmdformat 
         Caption         =   "&Format Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu cmdstatus 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu menuinsert 
      Caption         =   "&Insert"
      Begin VB.Menu cmddate 
         Caption         =   "&Date"
      End
      Begin VB.Menu cmdpicture 
         Caption         =   "&Picture"
      End
   End
   Begin VB.Menu menuformat 
      Caption         =   "&Format"
      Begin VB.Menu cmdfont 
         Caption         =   "&Font"
         Shortcut        =   ^D
      End
      Begin VB.Menu cmdbullet 
         Caption         =   "&Bullet Style"
      End
      Begin VB.Menu cmdbold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu cmditalic 
         Caption         =   "&Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu cmdunderline 
         Caption         =   "&Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu cmduppercase 
         Caption         =   "&Uppercase"
      End
      Begin VB.Menu cmdlowercase 
         Caption         =   "&Lowercase"
      End
      Begin VB.Menu cmdsentencecase 
         Caption         =   "&Sentence Case"
      End
      Begin VB.Menu cmdpagecolor 
         Caption         =   "&Page Color"
      End
      Begin VB.Menu cmdfontcolor 
         Caption         =   "&Font Color"
      End
   End
   Begin VB.Menu menupopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu undo 
         Caption         =   "Undo"
      End
      Begin VB.Menu menufileline4 
         Caption         =   "-"
      End
      Begin VB.Menu cmdcopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu cmdpaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu cmdcut 
         Caption         =   "Cut"
      End
   End
   Begin VB.Menu menuhelp 
      Caption         =   "&Help"
      Begin VB.Menu cmdzaldy 
         Caption         =   "&About ZaldyWord"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu Menuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
   
Private Sub cmdbold_Click()
   Form1.RichTextBox1.SelBold = Not Form1.RichTextBox1.SelBold
End Sub

Private Sub cmdbullet_Click()
        Form1.RichTextBox1.TextRTF = Form1.RichTextBox1.TextRTF
        Form1.RichTextBox1.SelBullet = Not Form1.RichTextBox1.SelBullet
End Sub

Private Sub cmdcopy_Click()
    Clipboard.Clear
    Clipboard.SetText Screen.ActiveControl.SelText
End Sub

Private Sub cmdcut_Click()
    Clipboard.Clear
    Clipboard.SetText Screen.ActiveControl.SelText
    Screen.ActiveControl.SelText = ""
End Sub


Private Sub cmddelete_Click()
    Form1.RichTextBox3.TextRTF = Form1.RichTextBox1.TextRTF
    If Form1.RichTextBox1.SelText <> "" Then
        Form1.RichTextBox1.SelText = ""
    Else
        Form1.RichTextBox1.SelLength = 1
        Form1.RichTextBox1.SelText = ""
    End If
End Sub


Private Sub cmdexit_Click()
    If saved = 1 Then
        Clipboard.Clear
        Unload Me
        End
    ElseIf saved = 0 Then
        s = MsgBox("Document [" & Form1.Caption & "] Not Saved, save it.", vbInformation + vbYesNoCancel, "Zaldy's Wordpad Error")
        If s = vbYes Then
            cmdsave_Click
            cmdexit_Click
        ElseIf s = vbNo Then
            saved = 1
            cmdexit_Click
        End If
    End If
End Sub


'NOTE:cdlCFBoth, &H3, Causes the dialog box to list the available printer and screen fonts. The hDC property identifies thedevice context associated with the printer.
'NOTE:cdlCFApply, &H200, Enables the Apply button on the dialog box.
'NOTE:cdlCFEffects, &H100, Specifies that the dialog box enables strikethrough, underline, and color effects.
'NOTE:cdlCFForceFontExist, &H10000, Specifies that an error message box is displayed if the user attempts to select a font or style that doesn't exist.
Private Sub cmdfont_Click()
    cd.CancelError = True
    cd.Flags = cdlCFBoth Or cdlCFApply Or cdlCFEffects Or cdlCFForceFontExist
    cd.ShowFont
    Form1.RichTextBox1.SelFontName = cd.FontName
    Form1.RichTextBox1.SelFontSize = cd.FontSize
    Form1.RichTextBox1.SelItalic = cd.FontItalic
    Form1.RichTextBox1.SelBold = cd.FontBold
    Form1.RichTextBox1.SelUnderline = cd.FontUnderline
    Form1.RichTextBox1.SelStrikeThru = cd.FontStrikethru
    Form1.RichTextBox1.SelColor = cd.Color
End Sub


Private Sub cmdfontcolor_Click()
    cd.ShowColor
    Form1.RichTextBox1.SelColor = cd.Color
End Sub

Private Sub cmdformat_Click()
    cmdformat.Checked = Not cmdformat.Checked
    Toolbar2.Visible = cmdformat.Checked
End Sub

Private Sub cmditalic_Click()
    Form1.RichTextBox1.SelItalic = Not Form1.RichTextBox1.SelItalic
End Sub

Private Sub cmdlowercase_Click()
    Form1.RichTextBox3.TextRTF = Form1.RichTextBox1.TextRTF
    Form1.RichTextBox1.SelText = StrConv(Form1.RichTextBox1.SelText, vbLowerCase)
End Sub

Private Sub cmdnew_Click()
  windowCTR = windowCTR + 1
  ReDim zaldy(windowCTR)
  zaldy(windowCTR).Tag = windowCTR
  zaldy(windowCTR).Caption = "UNTITLED:"
  zaldy(windowCTR).Show
End Sub

Private Sub cmdobject_Click()
    
End Sub

Private Sub cmdopen_Click()
    cd.Filter = "Rich Text Documents|*.rtf|Text Files|*.txt|All Files|*.*"
    cd.ShowOpen
 If fname = cd.FileName Then
    MsgBox " File already loaded"
Exit Sub
End If
     fname = cd.FileName
     Form1.RichTextBox1.LoadFile fname
End Sub

Private Sub cmdpagecolor_Click()
    cd.ShowColor
    Form1.RichTextBox1.BackColor = cd.Color
End Sub

Private Sub cmdpaste_Click()
    Screen.ActiveControl.SelText = Clipboard.GetText
End Sub

Private Sub cmdpicture_Click()
     cd.DialogTitle = "Select Picture..."
     cd.Filter = "Bitmaps (*.bmp;*.dib)|*.bmp;*.dib|GIF Images (*.gif)|*.gif|JPEG Images (*.jpg)|*.jpg|"
    cd.ShowOpen
    
    'Load picture into picInsert
    Picture1.Picture = LoadPicture(cd.FileName)
    
    'Copy the picture into the clipboard.
    Clipboard.Clear
    Clipboard.SetData Picture1.Picture
    
    'Paste the picture into the RichTextBox.
    SendMessage ActiveForm.RichTextBox1.hwnd, zaldy_paste, 0, 0&
    
End Sub

Private Sub cmdprinsetup_Click()
     cd.Flags = cdlPDPrintSetup
     cd.ShowPrinter
     End Sub

Private Sub cmdprint_Click()
    cd.Flags = cdlPDReturnDC + cdlPDNoPageNums
    cd.ShowPrinter
    Form1.RichTextBox1.SelPrint cd.hDC

End Sub

Private Sub cmdsave_Click()
      cd.Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt
      cd.Filter = "Rich Text Documents|*.rtf|Text Files|*.txt|All Files|*.*"
      fname = cd.FileName
      Form1.RichTextBox1.SaveFile fname
  
End Sub

Private Sub cmdsaveas_Click()
     cd.Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt
     cd.Filter = "Rich Text Documents|*.rtf|Text Files|*.txt|All Files|*.*"
     cd.ShowSave
     fname = cd.FileName
    Form1.RichTextBox1.SaveFile fname
End Sub

Private Sub cmdselectall_Click()
    Form1.RichTextBox1.SelStart = 0
    Form1.RichTextBox1.SelLength = Len(Form1.RichTextBox1.Text)
End Sub

Private Sub cmdstatus_Click()
    cmdstatus.Checked = Not cmdstatus.Checked
    StatusBar1.Visible = cmdstatus.Checked
End Sub

Private Sub cmdtoolbar_Click()
    cmdtoolbar.Checked = Not cmdtoolbar.Checked
    Toolbar1.Visible = cmdtoolbar.Checked
End Sub

Private Sub cmdunderline_Click()
Form1.RichTextBox1.SelUnderline = Not Form1.RichTextBox1.SelUnderline
End Sub

Private Sub cmdundo_Click()
    SendMessage ActiveForm.RichTextBox1.hwnd, zaldy_undo, 0, 0&
End Sub

Private Sub cmduppercase_Click()
    Form1.RichTextBox3.TextRTF = Form1.RichTextBox1.TextRTF
    Form1.RichTextBox1.SelText = StrConv(Form1.RichTextBox1.SelText, vbUpperCase)
End Sub





Private Sub cmdzaldy_Click()
    f1.Show
End Sub

Private Sub Combo1_Click()
    ActiveForm.RichTextBox1.SelFontName = Combo1.Text 'Set selected font name
    Form1.RichTextBox1.SetFocus
End Sub

Private Sub Combo2_click()
    ActiveForm.RichTextBox1.SelFontSize = Combo2.Text 'Set selected font size
    Form1.RichTextBox1.SetFocus
    End Sub

Private Sub copy_Click()
       On Error Resume Next
    Clipboard.SetText Form1.RichTextBox1.SelRTF
End Sub

Private Sub cut_Click()
     On Error Resume Next
    Form1.RichTextBox2.TextRTF = Form1.RichTextBox1.TextRTF
    Clipboard.SetText Form1.RichTextBox1.SelRTF
    Form1.RichTextBox1.SelText = ""
End Sub

Private Sub MDIForm_Load()
For danish = 1 To Screen.FontCount - 1
    Combo1.AddItem Screen.Fonts(danish)
Next
End Sub





Private Sub MDIForm_Unload(Cancel As Integer)
    MsgBox "Thank You For Using Zdpad", vbOKOnly, "zdpad2"
End Sub

Private Sub Menuabout_Click()
    form2.Show
End Sub

Private Sub paste_Click()
        Form1.RichTextBox2.TextRTF = Form1.RichTextBox1.TextRTF
        Form1.RichTextBox1.SelRTF = Clipboard.GetText
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
        Case "Open"
            cmdopen_Click
        Case "Save"
            cmdsave_Click
        Case "Print"
            cmdprint_Click
        Case "Cut"
            cmdcut_Click
        Case "Copy"
            cmdcopy_Click
        Case "Paste"
            cmdpaste_Click
        Case "New"
            cmdnew_Click
    
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Toolbar2.Align = 0
    Select Case Button.Key
        Case "Bold"
            cmdbold_Click
        Case "Italic"
            cmditalic_Click
        Case "Underline"
            cmdunderline_Click
        Case "Align Left"
            Form1.RichTextBox1.SelAlignment = rtfLeft
        Case "Align Center"
            Form1.RichTextBox1.SelAlignment = rtfCenter
        Case "Align Right"
            Form1.RichTextBox1.SelAlignment = rtfRight
        Case "Justify"
            Form1.RichTextBox1.SelAlignment = rtfJustify
        Case "Font Color"
            cmdfontcolor_Click
        Case "Page Color"
            cmdpagecolor_Click
                       
    End Select
End Sub



Private Sub undo_Click()
    cmdundo_Click
End Sub
