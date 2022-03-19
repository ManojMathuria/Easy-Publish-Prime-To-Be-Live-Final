VERSION 5.00
Begin VB.Form FrmDialog 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Select Printing Format"
      Height          =   2295
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton Command1 
         Caption         =   "Sheet"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Format-1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3000
         TabIndex        =   4
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Format-1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FrmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Flag As Long
Private Sub Command1_Click()
If Flag = 1 Then frmSalesVoucher.PtgType = 1: Me.ActiveControl.SetFocus: Unload Me
End Sub
Private Sub Command2_Click()
If Flag = 1 Then frmSalesVoucher.PtgType = 2: Me.ActiveControl.SetFocus: Unload Me
End Sub
Private Sub Command3_Click()
If Flag = 1 Then frmSalesVoucher.PtgType = 3: Me.ActiveControl.SetFocus: Unload Me
End Sub
Private Sub Command4_Click()
If Flag = 1 Then frmSalesVoucher.PtgType = 4: Me.ActiveControl.SetFocus: Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo ErrorHandler
If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
CenterForm Me
    BusySystemIndicator True
ErrorHandler:
    BusySystemIndicator False
End Sub
