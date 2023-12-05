VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Menu mnu 
      Caption         =   "Item "
      Begin VB.Menu menu1 
         Caption         =   "Item Ledger"
         Index           =   1
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Private Sub Form_Load()
    ' Set a property of the UserControl
    UserControl11.Caption = "Hello from UserControl"

    ' Call a method of the UserControl
    UserControl11.ShowMessage "This is a message from the form"
End Sub
End Sub
