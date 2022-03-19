VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJobWorkOrderDuplication 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Work Order Duplication"
   ClientHeight    =   3105
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3840
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3840
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   2720
      Left            =   45
      TabIndex        =   2
      Top             =   345
      Width           =   3765
      _Version        =   65536
      _ExtentX        =   6641
      _ExtentY        =   4798
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
      Picture         =   "JobWorkOrderDuplication.frx":0000
      Begin MSComctlLib.ListView ListView1 
         Height          =   2395
         Left            =   0
         TabIndex        =   0
         Top             =   315
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   4233
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
         TabIndex        =   3
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
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
         Caption         =   " C&onvert To"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "JobWorkOrderDuplication.frx":001C
         Picture         =   "JobWorkOrderDuplication.frx":0038
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   330
         Left            =   1080
         TabIndex        =   4
         Top             =   0
         Width           =   2680
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4727;582"
         ListRows        =   3
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JobWorkOrderDuplication.frx":0054
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JobWorkOrderDuplication.frx":0596
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmJobWorkOrderDuplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    CenterForm Me
    ListView1.ColumnHeaders.Add 1, , "List of Job Work Types"
    Combo3.AddItem "Purchase Order", 0
    Combo3.AddItem "Sales Order", 1
    Combo3.AddItem "Cost Estimation", 2
    Combo3.ListIndex = 1
    Dim i As Integer
    For i = 1 To 6
        ListView1.ListItems.Add , , Choose(i, "Multi Form Format", "Spread Format", "Combo Format", "Misc Operations", "Binding Process", "BOM")
    Next
    LockWindowUpdate ListView1.hwnd
    ListView1.ColumnHeaders(1).Width = ListView1.Width
    LockWindowUpdate 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Sendkeys "{TAB}", True: KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2): KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = IIf(Shift = vbCtrlMask And KeyCode = vbKeyA, True, False)
        Next i
    ElseIf Shift = vbAltMask And KeyCode = vbKeyC Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1): KeyCode = 0
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        Me.Tag = "Copy"
        Me.Hide
    ElseIf Button.Index = 2 Then
        Me.Tag = "Cancel"
        Me.Hide
    End If
End Sub
