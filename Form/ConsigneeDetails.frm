VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmConsigneeDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ship To Party Details"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   6485
      Picture         =   "ConsigneeDetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Proceed"
      Top             =   120
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   2155
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   6275
      _Version        =   65536
      _ExtentX        =   11068
      _ExtentY        =   3801
      _StockProps     =   77
      TintColor       =   16711935
      Alignment       =   0
      AutoSize        =   0   'False
      BevelSize       =   0
      BevelStyle      =   0
      BorderColor     =   -2147483642
      BorderStyle     =   1
      FillColor       =   -2147483633
      FontStyle       =   0
      FontTransparent =   0   'False
      LightColor      =   -2147483643
      ShadowColor     =   -2147483632
      TextColor       =   -2147483640
      WallPaper       =   0
      NoPrefix        =   0   'False
      FormatString    =   ""
      Caption         =   ""
      Picture         =   "ConsigneeDetails.frx":05BE
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1700
         Width           =   5190
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1380
         Width           =   5190
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1070
         Width           =   5190
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   40
         TabIndex        =   2
         Top             =   750
         Width           =   5190
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   40
         TabIndex        =   1
         Top             =   435
         Width           =   5190
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   40
         TabIndex        =   0
         Top             =   120
         Width           =   5190
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Party"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ConsigneeDetails.frx":05DA
         Picture         =   "ConsigneeDetails.frx":05F6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   1275
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   435
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   2249
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Address"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ConsigneeDetails.frx":0612
         Picture         =   "ConsigneeDetails.frx":062E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   10
         Top             =   1700
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " GSTIN"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ConsigneeDetails.frx":064A
         Picture         =   "ConsigneeDetails.frx":0666
      End
   End
End
Attribute VB_Name = "FrmConsigneeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    CenterForm Me
End Sub
Private Sub Form_Activate()
    Text1.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}", True
        KeyCode = 0
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2) Or KeyCode = vbKeyEscape Then
        cmdProceed_Click
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = True: cmdProceed_Click
End Sub
Private Sub cmdProceed_Click()
    Me.Hide
End Sub
