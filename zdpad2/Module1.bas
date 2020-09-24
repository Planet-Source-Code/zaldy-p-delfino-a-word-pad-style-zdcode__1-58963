Attribute VB_Name = "Module1"

Public NewInstance As Boolean
Public zaldy() As New Form1
Public windowCTR As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long
Public Const zaldy_undo = &HC7
Public Const zaldy_paste = &H302
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public fname As String
 Public saved As Integer
