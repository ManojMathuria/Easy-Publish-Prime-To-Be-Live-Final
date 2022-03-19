VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBookPOPrintUtility 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Print Order Print Utility"
   ClientHeight    =   7365
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   6840
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mail"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BookPOPrintUtility.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BookPOPrintUtility.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BookPOPrintUtility.frx":0226
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6975
      Left            =   45
      TabIndex        =   9
      Top             =   345
      Width           =   6765
      _Version        =   65536
      _ExtentX        =   11933
      _ExtentY        =   12303
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
      Picture         =   "BookPOPrintUtility.frx":033A
      Begin VB.TextBox txtBookOrder 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   21
         Top             =   5230
         Width           =   4845
      End
      Begin VB.TextBox TextUnitCost 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   19
         Top             =   5545
         Width           =   4845
      End
      Begin VB.TextBox txtComboPrinting 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   17
         Top             =   3980
         Width           =   4845
      End
      Begin VB.TextBox txtAll 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   7
         Top             =   4915
         Width           =   4845
      End
      Begin VB.TextBox txtBinding 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   6
         Top             =   4600
         Width           =   4845
      End
      Begin VB.TextBox txtLamination 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   5
         Top             =   4290
         Width           =   4845
      End
      Begin VB.TextBox txtTitlePrinting 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   4
         Top             =   3660
         Width           =   4845
      End
      Begin VB.TextBox txtBookPrinting 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   3
         Top             =   3345
         Width           =   4845
      End
      Begin VB.TextBox MhRealInput16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
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
         Left            =   2670
         MaxLength       =   13
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox MhRealInput15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
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
         Left            =   840
         MaxLength       =   13
         TabIndex        =   0
         Top             =   0
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2865
         Left            =   0
         TabIndex        =   2
         Top             =   315
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   5054
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   0
         TabIndex        =   10
         Top             =   0
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
         Caption         =   " &From"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOPrintUtility.frx":0356
         Picture         =   "BookPOPrintUtility.frx":0372
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   11
         Top             =   0
         Width           =   765
         _Version        =   65536
         _ExtentX        =   1349
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
         Caption         =   " &To"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOPrintUtility.frx":038E
         Picture         =   "BookPOPrintUtility.frx":03AA
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   0
         TabIndex        =   12
         Top             =   3345
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
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
         Caption         =   " MF_Jobwork"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOPrintUtility.frx":03C6
         Picture         =   "BookPOPrintUtility.frx":03E2
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   0
         TabIndex        =   13
         Top             =   3660
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
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
         Caption         =   " SF_Jobwork"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOPrintUtility.frx":03FE
         Picture         =   "BookPOPrintUtility.frx":041A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   0
         TabIndex        =   14
         Top             =   4290
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
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
         Caption         =   " MO_Jobwork"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOPrintUtility.frx":0436
         Picture         =   "BookPOPrintUtility.frx":0452
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   0
         TabIndex        =   15
         Top             =   4605
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
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
         Caption         =   " BP_Jobwork"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOPrintUtility.frx":046E
         Picture         =   "BookPOPrintUtility.frx":048A
      End
      Begin Mh3dlblLib.Mh3dLabel lblBookOrder 
         Height          =   330
         Left            =   0
         TabIndex        =   16
         Top             =   4920
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
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
         Caption         =   " All-Job-Work"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOPrintUtility.frx":04A6
         Picture         =   "BookPOPrintUtility.frx":04C2
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   0
         TabIndex        =   18
         Top             =   3975
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
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
         Caption         =   " CF_Jobwork"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOPrintUtility.frx":04DE
         Picture         =   "BookPOPrintUtility.frx":04FA
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   0
         TabIndex        =   20
         Top             =   5550
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
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
         Caption         =   " FI-Unit Cost"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOPrintUtility.frx":0516
         Picture         =   "BookPOPrintUtility.frx":0532
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   0
         TabIndex        =   22
         Top             =   5235
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
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
         Caption         =   " Job-Work-UC"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOPrintUtility.frx":054E
         Picture         =   "BookPOPrintUtility.frx":056A
      End
      Begin MSForms.OptionButton OptionButton2 
         Height          =   375
         Left            =   5160
         TabIndex        =   24
         Top             =   15
         Width           =   1575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2778;661"
         Value           =   "1"
         Caption         =   "Purchase Order"
         FontName        =   "Comic Sans MS"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton OptionButton1 
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   15
         Width           =   1275
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2249;661"
         Value           =   "0"
         Caption         =   "Sales Order"
         FontName        =   "Comic Sans MS"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   6840
         Y1              =   3255
         Y2              =   3255
      End
   End
End
Attribute VB_Name = "FrmBookPOPrintUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchCode As String
Dim iCount As Long
Dim rstBookPrintOrder As New ADODB.Recordset, BookPOType As String
Private Sub Form_Load()
    Dim i As Integer
    On Error GoTo ErrorHandler
    Unload FrmBookPrintOrder
    CenterForm Me
    BusySystemIndicator True
    
    If VchCode = 1 Then
    Me.Caption = "Item Order Print Utility"
    ListView1.ColumnHeaders.Add 1, , "List of Order Types"
    For i = 1 To 8
        ListView1.ListItems.Add , , Choose(i, "Multi-Form-Format Order", "Multi-Element-Format Order", "Combo-Format Order", "Misc-Operations Order", "Binding-Process Order", "ALL Order (MF_ME_CF_MO_BP)", "Jobwork-Unit Cost", "Unit Cost")
        ListView1.ListItems(i).Checked = True
    Next
    ElseIf VchCode = 2 Then
    Me.Caption = "JobCard Print Utility"
    ListView1.ColumnHeaders.Add 1, , "List of JobCard Types"
    For i = 1 To 6
        ListView1.ListItems.Add , , Choose(i, "Multi-Form-Format Order", "Multi-Element-Format Order", "Combo-Format Order", "Misc-Operations Order", "Binding-Process Order", "ALL Order (MF_ME_CF_MO_BP)", "Jobwork-Unit Cost", "Unit Cost")
        ListView1.ListItems(i).Checked = True
    Next
    ElseIf VchCode = 3 Then
    Me.Caption = "Plate-Order Print Utility"
    ListView1.ColumnHeaders.Add 1, , "List of Plate-Order Types"
    For i = 1 To 4
        ListView1.ListItems.Add , , Choose(i, "Multi-Form-Format Order", "Multi-Element-Format Order", "Combo-Format Order", "ALL Order (MF_ME_CF)")
        ListView1.ListItems(i).Checked = True
    Next
    ElseIf VchCode = 4 Then
    Me.Caption = "Paper-Requisition-Slip Print Utility"
    ListView1.ColumnHeaders.Add 1, , "List of Paper-Requisition-Slip Types"
    For i = 1 To 4
        ListView1.ListItems.Add , , Choose(i, "Multi-Form-Format Order", "Multi-Element-Format Order", "Combo-Format Order", "ALL Order (MF_ME_CF)")
        ListView1.ListItems(i).Checked = True
    Next
    ElseIf VchCode = 5 Then
    Me.Caption = "Quotation Format Print Utility"
    ListView1.ColumnHeaders.Add 1, , "List of Quotation Format Types"
    For i = 1 To 7
        ListView1.ListItems.Add , , Choose(i, "Multi-Form-Format Order", "Multi-Element-Format Order", "Combo-Format Order", "Misc-Operations Order", "Binding-Process Order", "ALL Order (MF_ME_CF_MO_BP)", "Excel Format")
        ListView1.ListItems(i).Checked = True
    Next
    End If
    LockWindowUpdate ListView1.hwnd
    ListView1.ColumnHeaders(1).Width = ListView1.Width
    LockWindowUpdate 0
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Sendkeys "{TAB}", True
       KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyM Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Call CloseForm(Me)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstBookPrintOrder)
End Sub
Private Sub MhRealInput15_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput15_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput15, KeyAscii, 0
End Sub
Private Sub MhRealInput15_Validate(Cancel As Boolean)
    If Not ValidateNumber(Me.ActiveControl, 0) Then
        Cancel = True
    ElseIf Val(MhRealInput15.Text) <= 0 Then
        MhRealInput15.SetFocus
        FocusSelect Me.ActiveControl
        Cancel = True
    End If
End Sub
Private Sub MhRealInput16_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput16_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput16, KeyAscii, 0
End Sub
Private Sub MhRealInput16_Validate(Cancel As Boolean)
    If Not ValidateNumber(Me.ActiveControl, 0) Then
        Cancel = True
    ElseIf Val(MhRealInput16.Text) <= 0 Or Val(MhRealInput16.Text) < Val(MhRealInput15.Text) Then
        MhRealInput16.SetFocus
        FocusSelect Me.ActiveControl
        Cancel = True
    End If
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = True
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = False
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    If Button.Index = 1 Then
        Call PrintBookPrintOrder("M")
    ElseIf Button.Index = 2 Then
        Call PrintBookPrintOrder("P")
    ElseIf Button.Index = 3 Then
        Call CloseForm(Me)
    End If
End Sub
Sub Delay()
        iCount = 1
        For iCount = 1 To 1000000000
        iCount = iCount + 1
        Next
End Sub
Private Sub PrintBookPrintOrder(ByVal OutputType As String)
    If OptionButton2 = True Then BookPOType = "P" Else BookPOType = "S"
    Screen.MousePointer = vbHourglass
    If rstBookPrintOrder.State = adStateOpen Then rstBookPrintOrder.Close
    rstBookPrintOrder.Open "Select Code From BookPOParent Where Right(Type,1)= '" & BookPOType & "' And LEFT(Type,1)<>'O' And Left(Code,1)<>'*' And Name >= '" & Pad(Trim(MhRealInput15.Text), Space(1), 10, "L") & "' And Name <= '" & Pad(Trim(MhRealInput16.Text), Space(1), 10, "L") & "' AND FYCode='" & FYCode & " ' Order By BookPOParent.Name", cnDatabase, adOpenKeyset, adLockReadOnly
    Screen.MousePointer = vbNormal
    If rstBookPrintOrder.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    
    Do While Not rstBookPrintOrder.EOF
    If VchCode = 1 Then 'Jobwork Order
        If ListView1.ListItems(1).Checked Then Call FrmBookPrintOrder.PrintBookPrintOrder02(rstBookPrintOrder.Fields("Code").Value, txtBookPrinting.Text, OutputType, "BP", BookPOType)
        If ListView1.ListItems(2).Checked Then Call FrmBookPrintOrder.PrintBookPrintOrder02(rstBookPrintOrder.Fields("Code").Value, txtTitlePrinting.Text, OutputType, "TP", BookPOType)
        If ListView1.ListItems(3).Checked Then Call FrmBookPrintOrder.PrintBookPrintOrder02(rstBookPrintOrder.Fields("Code").Value, txtComboPrinting.Text, OutputType, "CB", BookPOType)
        If ListView1.ListItems(4).Checked Then Call FrmBookPrintOrder.PrintBookPrintOrder02(rstBookPrintOrder.Fields("Code").Value, txtLamination.Text, OutputType, "TL", BookPOType)
        If ListView1.ListItems(5).Checked Then Call FrmBookPrintOrder.PrintBookPrintOrder02(rstBookPrintOrder.Fields("Code").Value, txtBinding.Text, OutputType, "BB", BookPOType)
        If ListView1.ListItems(6).Checked Then Call FrmBookPrintOrder.PrintBookPrintOrder02(rstBookPrintOrder.Fields("Code").Value, txtAll.Text, OutputType, "ALL", BookPOType)
        If ListView1.ListItems(7).Checked Then Call FrmBookPrintOrder.PrintBookPrintOrder01(rstBookPrintOrder.Fields("Code").Value, txtBookOrder.Text, OutputType, "JUC", BookPOType)
        If ListView1.ListItems(8).Checked Then Call FrmBookPrintOrder.PrintBookPrintOrder01(rstBookPrintOrder.Fields("Code").Value, TextUnitCost.Text, OutputType, "UC", BookPOType)
    ElseIf VchCode = 2 Then 'JobCard
        If ListView1.ListItems(1).Checked Then Call FrmBookPrintOrder.JobCard(rstBookPrintOrder.Fields("Code").Value, txtBookPrinting.Text, OutputType, "BP", BookPOType)
        If ListView1.ListItems(2).Checked Then Call FrmBookPrintOrder.JobCard(rstBookPrintOrder.Fields("Code").Value, txtTitlePrinting.Text, OutputType, "TP", BookPOType)
        If ListView1.ListItems(3).Checked Then Call FrmBookPrintOrder.JobCard(rstBookPrintOrder.Fields("Code").Value, txtComboPrinting.Text, OutputType, "CB", BookPOType)
        If ListView1.ListItems(4).Checked Then Call FrmBookPrintOrder.JobCard(rstBookPrintOrder.Fields("Code").Value, txtLamination.Text, OutputType, "TL", BookPOType)
        If ListView1.ListItems(5).Checked Then Call FrmBookPrintOrder.JobCard(rstBookPrintOrder.Fields("Code").Value, txtBinding.Text, OutputType, "BB", BookPOType)
        If ListView1.ListItems(6).Checked Then Call FrmBookPrintOrder.JobCard(rstBookPrintOrder.Fields("Code").Value, txtAll.Text, OutputType, "ALL", BookPOType)
    ElseIf VchCode = 3 Then 'Plate Order
        'If ListView1.ListItems(1).Checked Then Call FrmBookPrintOrder.PrintTitlePlateOrder(rstBookPrintOrder.Fields("Code").Value, txtBookPrinting.Text, OutputType, "", BookPOType)
        If ListView1.ListItems(2).Checked Then Call FrmBookPrintOrder.PrintBookPrintOrder03(rstBookPrintOrder.Fields("Code").Value, txtTitlePrinting.Text, OutputType, "BP", BookPOType)
        If ListView1.ListItems(3).Checked Then Call FrmBookPrintOrder.PrintBookPrintOrder03(rstBookPrintOrder.Fields("Code").Value, txtComboPrinting.Text, OutputType, "TP", BookPOType)
        If ListView1.ListItems(4).Checked Then Call FrmBookPrintOrder.PrintBookPrintOrder03(rstBookPrintOrder.Fields("Code").Value, txtAll.Text, OutputType, "All", BookPOType)
    ElseIf VchCode = 4 Then 'Paper Slip
        If ListView1.ListItems(1).Checked Then Call FrmBookPrintOrder.PaperSlip(rstBookPrintOrder.Fields("Code").Value, txtBookPrinting.Text, OutputType, "BP", BookPOType)
        If ListView1.ListItems(2).Checked Then Call FrmBookPrintOrder.PaperSlip(rstBookPrintOrder.Fields("Code").Value, txtTitlePrinting.Text, OutputType, "TP", BookPOType)
        If ListView1.ListItems(3).Checked Then Call FrmBookPrintOrder.PaperSlip(rstBookPrintOrder.Fields("Code").Value, txtComboPrinting.Text, OutputType, "CB", BookPOType)
        If ListView1.ListItems(4).Checked Then Call FrmBookPrintOrder.PaperSlip(rstBookPrintOrder.Fields("Code").Value, txtAll.Text, OutputType, "All", BookPOType)
    ElseIf VchCode = 5 Then 'Quotation Format
        If ListView1.ListItems(1).Checked Then Call FrmBookPrintOrder.PrintQuotationFormat(rstBookPrintOrder.Fields("Code").Value, txtBookPrinting.Text, OutputType, "BP", BookPOType)
        If ListView1.ListItems(2).Checked Then Call FrmBookPrintOrder.PrintQuotationFormat(rstBookPrintOrder.Fields("Code").Value, txtTitlePrinting.Text, OutputType, "TP", BookPOType)
        If ListView1.ListItems(3).Checked Then Call FrmBookPrintOrder.PrintQuotationFormat(rstBookPrintOrder.Fields("Code").Value, txtComboPrinting.Text, OutputType, "CB", BookPOType)
        If ListView1.ListItems(4).Checked Then Call FrmBookPrintOrder.PrintQuotationFormat(rstBookPrintOrder.Fields("Code").Value, txtLamination.Text, OutputType, "TL", BookPOType)
        If ListView1.ListItems(5).Checked Then Call FrmBookPrintOrder.PrintQuotationFormat(rstBookPrintOrder.Fields("Code").Value, txtBinding.Text, OutputType, "BB", BookPOType)
        If ListView1.ListItems(6).Checked Then Call FrmBookPrintOrder.PrintQuotationFormat(rstBookPrintOrder.Fields("Code").Value, txtAll.Text, OutputType, "ALL", BookPOType)
    End If
'        If ListView1.ListItems(5).Checked Then Call FrmBookPrintOrder.PrintBookOrder(rstBookPrintOrder.Fields("Code").Value, txtBookOrder.Text, OutputType)
'        If ListView1.ListItems(6).Checked Then If OutputType = "P" Then Call FrmBookPrintOrder.PrintBookBoxLabel(rstBookPrintOrder.Fields("Code").Value, txtBinding.Text, OutputType)
        rstBookPrintOrder.MoveNext
    Loop
End Sub
