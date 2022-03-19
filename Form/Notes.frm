VERSION 5.00
Begin VB.Form frmNotes 
   BackColor       =   &H00FFFEF2&
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12000
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSave 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5633
      TabIndex        =   2
      Top             =   6735
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   11295
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   10815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   " Notes: Easy Info Solutions International"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   11285
      End
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NotesFlag As Long
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
       If Shift = 0 And KeyCode = vbKeyEscape Then
        Call btnSave_Click: KeyCode = 0
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        Call btnSave_Click: KeyCode = 0
        End If
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
    BusySystemIndicator True
    BusySystemIndicator False
    If NotesFlag = 1 Then Text1.Text = FrmAccountMaster.txtNotes.Text
    If NotesFlag = 2 Then Text1.Text = FrmBookMaster.txtNotes.Text
    If NotesFlag = 3 Then Text1.Text = frmDebitCreditVoucher.txtNotes.Text
    If NotesFlag = 4 Then Text1.Text = frmSalesChallanVoucher.txtNotes.Text
    If NotesFlag = 5 Then Text1.Text = frmSalesOrderVoucher.txtNotes.Text
    If NotesFlag = 6 Then Text1.Text = frmSalesVoucher.txtNotes.Text
    If NotesFlag = 7 Then Text1.Text = frmItemIssueReceiptVoucher.txtNotes.Text
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub btnSave_Click()
SaveFields
Call CloseForm(Me)
NotesFlag = 0
End Sub
Private Sub SaveFields()
If NotesFlag = 1 Then FrmAccountMaster.txtNotes.Text = Text1.Text
If NotesFlag = 2 Then FrmBookMaster.txtNotes.Text = Text1.Text
If NotesFlag = 3 Then frmDebitCreditVoucher.txtNotes.Text = Text1.Text
If NotesFlag = 4 Then frmSalesChallanVoucher.txtNotes.Text = Text1.Text
If NotesFlag = 5 Then frmSalesOrderVoucher.txtNotes.Text = Text1.Text
If NotesFlag = 6 Then frmSalesVoucher.txtNotes.Text = Text1.Text
If NotesFlag = 7 Then frmItemIssueReceiptVoucher.txtNotes.Text = Text1.Text
End Sub
