VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmItemSelectionList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Selection List...."
   ClientHeight    =   9630
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9705
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
   ScaleHeight     =   9630
   ScaleWidth      =   9705
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mail"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.CheckBox Check1 
         Caption         =   "Summarize"
         Height          =   255
         Left            =   8280
         TabIndex        =   13
         Top             =   40
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemSelectionList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemSelectionList.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemSelectionList.frx":0658
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemSelectionList.frx":076A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9250
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   9675
      _Version        =   65536
      _ExtentX        =   17066
      _ExtentY        =   16316
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
      Picture         =   "ItemSelectionList.frx":087C
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   300
         Index           =   2
         Left            =   105
         TabIndex        =   16
         Top             =   8880
         Width           =   9460
         _Version        =   65536
         _ExtentX        =   16686
         _ExtentY        =   529
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Ctrl+A->Select ALL  Ctrl+D->Deselect ALL  Ctrl+V->Print Preview  Ctrl+P->Print Preview"
         Alignment       =   0
         FillColor       =   8421504
         TextColor       =   16777215
         Picture         =   "ItemSelectionList.frx":0898
         Picture         =   "ItemSelectionList.frx":08B4
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   4020
         Left            =   4830
         TabIndex        =   8
         Top             =   4800
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   7091
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   11
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
         Caption         =   " From"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ItemSelectionList.frx":08D0
         Picture         =   "ItemSelectionList.frx":08EC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   2400
         TabIndex        =   12
         Top             =   0
         Width           =   885
         _Version        =   65536
         _ExtentX        =   1561
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
         Caption         =   "     To"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ItemSelectionList.frx":0908
         Picture         =   "ItemSelectionList.frx":0924
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   0
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "ItemSelectionList.frx":0940
         Caption         =   "ItemSelectionList.frx":0A58
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ItemSelectionList.frx":0AC4
         Keys            =   "ItemSelectionList.frx":0AE2
         Spin            =   "ItemSelectionList.frx":0B40
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd-mm-yyyy"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   1
         ForeColor       =   -2147483640
         Format          =   "dd-mm-yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "  -  -    "
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   39849
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   3270
         TabIndex        =   1
         Top             =   0
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "ItemSelectionList.frx":0B68
         Caption         =   "ItemSelectionList.frx":0C80
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ItemSelectionList.frx":0CEC
         Keys            =   "ItemSelectionList.frx":0D0A
         Spin            =   "ItemSelectionList.frx":0D68
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd-mm-yyyy"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   1
         ForeColor       =   -2147483640
         Format          =   "dd-mm-yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   " "
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "  -  -    "
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   39849
         CenturyMode     =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4500
         Left            =   0
         TabIndex        =   5
         Top             =   315
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   7938
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5325
         TabIndex        =   2
         Top             =   20
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6945
         TabIndex        =   3
         Top             =   20
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8295
         TabIndex        =   4
         Top             =   20
         Visible         =   0   'False
         Width           =   630
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   4020
         Left            =   0
         TabIndex        =   7
         Top             =   4800
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   7091
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4500
         Left            =   4830
         TabIndex        =   6
         Top             =   315
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   7938
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSForms.ComboBox ComboBox2 
         Height          =   400
         Left            =   3300
         TabIndex        =   15
         Top             =   8
         Width           =   1550
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2734;706"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   400
         Left            =   850
         TabIndex        =   14
         Top             =   8
         Width           =   1550
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2734;706"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "FrmItemSelectionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstMatCList As New ADODB.Recordset, rstSupplierList As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstItemGroupList As New ADODB.Recordset, rstPaperList As New ADODB.Recordset, rstVchSeriesList As New ADODB.Recordset
Public VchType As String, MC As String

Private Sub ComboBox1_Change()
Dim NewDate As Date
    Call AddDate(FinancialYearFrom, NewDate, ComboBox1.ListIndex)
    MhDateInput1.Value = NewDate
End Sub
Private Sub ComboBox2_Change()
Dim NewDate As Date
    Call AddDate(FinancialYearTo, NewDate, ComboBox2.ListIndex - 11)
    MhDateInput2.Value = NewDate
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If cnDatabase.State = 1 Then cnDatabase.Close: cnDatabase.Open
    MhDateInput1.ReadOnly = False:
    ComboBox1.Visible = False: ComboBox2.Visible = False
    If VchType <= 2 Or VchType = 33 Or VchType = 31 Or VchType = 49 Then
        ListView4.Visible = False: ListView1.BackColor = RGB(255, 255, 240): ListView2.BackColor = RGB(255, 255, 240): ListView3.BackColor = RGB(255, 255, 240): MhDateInput1.BackColor = RGB(255, 255, 240): MhDateInput2.BackColor = RGB(255, 255, 240):
        If VchType = 49 Then
            MhDateInput1.Visible = False: MhDateInput2.Visible = False
            ListView4.Visible = False
            ComboBox1.Visible = True: ComboBox2.Visible = True
            ComboBox1.FontSize = 10: ComboBox1.FontBold = True
            ComboBox1.Clear
            ComboBox1.AddItem " April", 0
            ComboBox1.AddItem " May", 1
            ComboBox1.AddItem " June", 2
            ComboBox1.AddItem " July", 3
            ComboBox1.AddItem " August", 4
            ComboBox1.AddItem " September", 5
            ComboBox1.AddItem " October", 6
            ComboBox1.AddItem " November", 7
            ComboBox1.AddItem " December", 8
            ComboBox1.AddItem " January", 9
            ComboBox1.AddItem " February", 10
            ComboBox1.AddItem " March", 11
            ComboBox1.ListIndex = 0
            
            ComboBox2.FontSize = 10: ComboBox2.FontBold = True
            ComboBox2.Clear
            ComboBox2.AddItem " April", 0
            ComboBox2.AddItem " May", 1
            ComboBox2.AddItem " June", 2
            ComboBox2.AddItem " July", 3
            ComboBox2.AddItem " August", 4
            ComboBox2.AddItem " September", 5
            ComboBox2.AddItem " October", 6
            ComboBox2.AddItem " November", 7
            ComboBox2.AddItem " December", 8
            ComboBox2.AddItem " January", 9
            ComboBox2.AddItem " February", 10
            ComboBox2.AddItem " March", 11
            ComboBox2.ListIndex = 11
        End If
    ElseIf (VchType >= 3 And VchType <= 6) Or (VchType >= 53 And VchType <= 56) Then
        ListView4.Visible = False: ListView1.BackColor = RGB(245, 255, 230): ListView2.BackColor = RGB(245, 255, 230): ListView3.BackColor = RGB(245, 255, 230): MhDateInput1.BackColor = RGB(245, 255, 230): MhDateInput2.BackColor = RGB(245, 255, 230):
    ElseIf (VchType >= 7 And VchType <= 10) Or (VchType >= 57 And VchType <= 60) Then
        ListView4.Visible = False: ListView1.BackColor = RGB(245, 250, 250): ListView2.BackColor = RGB(245, 250, 250): ListView3.BackColor = RGB(245, 250, 250): MhDateInput1.BackColor = RGB(245, 250, 250): MhDateInput2.BackColor = RGB(245, 250, 250):
    ElseIf (VchType >= 21 And VchType <= 24) Or (VchType >= 61 And VchType <= 64) Then
        ListView4.Visible = False: ListView1.BackColor = RGB(255, 250, 255): ListView2.BackColor = RGB(255, 250, 255): ListView3.BackColor = RGB(255, 250, 255): MhDateInput1.BackColor = RGB(255, 250, 255): MhDateInput2.BackColor = RGB(255, 250, 255):
    ElseIf (VchType >= 25 And VchType <= 28) Or (VchType >= 65 And VchType <= 69) Then
        ListView4.Visible = False: ListView1.BackColor = RGB(240, 255, 255): ListView2.BackColor = RGB(240, 255, 255): ListView3.BackColor = RGB(240, 255, 255): MhDateInput1.BackColor = RGB(240, 255, 255): MhDateInput2.BackColor = RGB(240, 255, 255):
    ElseIf VchType <= 20 And VchType >= 11 Then
        ListView4.Visible = False: ListView1.BackColor = RGB(245, 255, 230): ListView2.BackColor = RGB(245, 255, 230): ListView3.BackColor = RGB(245, 255, 230): MhDateInput1.BackColor = RGB(245, 255, 230): MhDateInput2.BackColor = RGB(245, 255, 230):
    ElseIf (VchType >= 35 And VchType <= 44) Or (VchType >= 46 And VchType <= 47) Then
        ListView4.Visible = False: ListView1.BackColor = RGB(240, 255, 255): ListView2.BackColor = RGB(240, 255, 255): ListView3.BackColor = RGB(240, 255, 255): MhDateInput1.BackColor = RGB(240, 255, 255): MhDateInput2.BackColor = RGB(240, 255, 255):
         ListView4.BackColor = RGB(240, 255, 255):
        Option3.Visible = True: Option2.Visible = True: Option1.Visible = True
        If VchType = 35 Or VchType = 36 Or VchType = 37 Or VchType = 38 Then ListView4.Visible = True
    ElseIf Right(VchType, 2) = 48 Then
        ListView4.Visible = True: ListView1.BackColor = RGB(240, 255, 255): ListView2.BackColor = RGB(240, 255, 255): ListView3.BackColor = RGB(240, 255, 255): MhDateInput1.BackColor = RGB(240, 255, 255): MhDateInput2.BackColor = RGB(240, 255, 255):
        ListView4.BackColor = RGB(240, 255, 255):
    ElseIf VchType = 103 Or VchType = 104 Or VchType = 105 Then
        ListView1.BackColor = RGB(240, 255, 255): ListView2.BackColor = RGB(240, 255, 255): ListView3.BackColor = RGB(240, 255, 255): MhDateInput1.BackColor = RGB(240, 255, 255): MhDateInput2.BackColor = RGB(240, 255, 255)
        ListView4.Visible = False: Option3.Visible = True: Option2.Visible = True: Option1.Visible = True: Check1.Visible = True
    End If

    If VchType <= 10 Then
        Me.Caption = "Selection List...." + Choose(IIf(VchType = 0, VchType + 1, VchType + 1), "Physical Stock Audit Item-Wise", "Inventory Movement Ledger Item-Wise", "Stock Status Item-Wise", "Sales Analysis Item-Wise", "Sales Return Analysis Item-Wise", "Sales And Sales Return Analysis Item-Wise", "Net Sales Analysis Item-Wise", "Sales Analysis One Party Item-Wise", "Sales Return Analysis One Party Item-Wise", "Sales And Sales Return Analysis One Party Item-Wise", "Net Sales Analysis One Party Item-Wise", "Paper Receipt Party-Wise")
        Me.Height = 9630
    ElseIf VchType <= 20 And VchType >= 11 Then
        Me.Height = 9630
        Me.Caption = "Selection List...." + Choose((VchType - 10), "Paper Receipt Party-Wise", "Paper Receipt Order-Wise", "Paer Receipt  Without-Order", "Paper Issue Party-Wise", "Paper Issue Order-Wise", "Paper Issue Without-Order", "Paper Transfer Party-Wise", "Paper Pending Order Party-Wise")
    ElseIf Right(VchType, 2) >= 21 And Right(VchType, 2) <= 48 And VchType <> 31 Then
        Me.Caption = "Selection List...." + Choose(Right(VchType, 2) - 20, "Sales Analysis Party-Wise", "Sales Return Analysis Party-Wise", "Sales And Sales Return Analysis Party-Wise", "Net Sales Analysis Party-Wise", "Sales Analysis One-Item Party-Wise", "Sales Return Analysis One-Item Party-Wise", "Sales And Sales Return Analysis One-Item Party-Wise", "Net Sales One-Item Party-Wise", "Sales Voucher-Wise", "30", "Sales Voucher-Wise", "32", "Short-Item Analysis Item-Wise", "34", "Purchase Orders-Party-Wise-Detailed", "Purchase Orders-Party-wise-Summarised", "Sales Orders-Party-Wise-Detailed", "Sales Orders-Party-wise-Summarised", "Purchase Orders Order-Wise", "Purchase Orders Party-wise", "Purchase Orders Item-wise", "Sale Orders Order-wise", "Sale Orders Party-wise", "Sale Orders Item-wise", "45", " Pending Sales Order", "Pending Sales Order Party-Wise", "Sales Voucher-Wise")
        Me.Height = 9630
    ElseIf VchType >= 53 And VchType <= 69 Then
        Me.Caption = "Selection List...." + Choose(VchType - 52, "Purchase Analysis Item-Wise", "Purchase Return Analysis Item-Wise", "Purchase And Purchase Return Analysis Item-Wise", "Net Purchase Analysis Item-Wise", "Purchase Analysis One Party Item-Wise", "Purchase Return Analysis One Party Item-Wise", "Purchase And Purchase Return Analysis One Party Item-Wise", "Net Purchase Analysis One Party Item-Wise", "Purchase Analysis Party-Wise", "Purchase Return Analysis Party-Wise", "Purchase And Purchase Return Analysis Party-Wise", "Net Purchase Analysis Party-Wise", "Purchase Analysis One-Item Party-Wise", "Purchase Return Analysis One-Item Party-Wise", "Purchase And Purchase Return Analysis One-Item Party-Wise", "Net Purchase One-Item Party-Wise", "Purchase Voucher-Wise")
        Me.Height = 9630
    ElseIf VchType = 103 Or VchType = 104 Or VchType = 105 Then
        Me.Caption = "Selection List...." + Choose(VchType - 102, "WIP Pending Ledger Item-Wise", "RM Pending Ledger Item-Wise", "RM AND WIP Pending Ledger Item-Wise")
        Me.Height = 9630
    ElseIf VchType = 49 Then
        Me.Caption = "Item Ledger Summarize": Me.Height = 10065 '9630
    Else
        Me.Caption = "Item Ledger Date-Wise": Me.Height = 9630
    End If
If (VchType <= 20 And VchType >= 18) Then: ListView1.Width = 9655
If VchType = 103 Or VchType = 104 Or VchType = 105 Then: ListView3.Width = 9655
    CenterForm Me
    BusySystemIndicator True
If VchType <= 2 Or VchType = 33 Then
    If rstSupplierList.State = adStateOpen Then rstSupplierList.Close
    rstSupplierList.Open "SELECT TOP 1 FinancialYearFrom  FROM CompanyMaster ORDER BY FinancialYearFrom", cnDatabase, adOpenForwardOnly, adLockReadOnly
Else
    If rstSupplierList.State = adStateOpen Then rstSupplierList.Close
    rstSupplierList.Open "SELECT TOP 1 FinancialYearFrom  FROM CompanyMaster WHERE FYCode='" & FYCode & "' ORDER BY FYCode", cnDatabase, adOpenForwardOnly, adLockReadOnly
End If
    MhDateInput1.Text = Format(rstSupplierList.Fields("FinancialYearFrom").Value, "dd-mm-yyyy")
    If VchType <> 49 Then
        MhDateInput2.Text = IIf(Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd"), Format(FinancialYearTo, "dd-mm-yyyy"), Format(Date, "dd-mm-yyyy"))
    ElseIf VchType = 49 Then
        MhDateInput2.Text = Format(FinancialYearTo, "dd-mm-yyyy")
    End If
    If rstSupplierList.State = adStateOpen Then rstSupplierList.Close
        rstSupplierList.Open "SELECT Name,Code FROM AccountMaster WHERE " & IIf(VchType <= 6, "[Group]='*99999'", IIf(VchType = 31, "[Group]='*99999'", "[Group]<>'*99999'")) & " ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
        rstAccountList.Open "SELECT Name,Code FROM AccountMaster WHERE " & IIf(VchType <= 2 Or VchType = 31 Or VchType = 49 Or VchType = 33, "[Group]='*99999'", "[Group]<>'*99999'") & " ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    If rstItemGroupList.State = adStateOpen Then rstItemGroupList.Close
        rstItemGroupList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type='5' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstItemGroupList.ActiveConnection = Nothing
    If rstPaperList.State = adStateOpen Then rstPaperList.Close
        rstPaperList.Open "SELECT Name,Code FROM PaperMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperList.ActiveConnection = Nothing
    If rstVchSeriesList.State = adStateOpen Then rstPaperList.Close
        rstVchSeriesList.Open "SELECT Name+'-'+VchName AS Name,Code FROM VchSeriesMaster Where Left(VchType,2)='" & Left(VchType, 2) & "' ORDER BY Name+'-'+VchName", cnDatabase, adOpenKeyset, adLockReadOnly
    rstVchSeriesList.ActiveConnection = Nothing
    If rstItemList.State = adStateOpen Then rstItemList.Close
        rstItemList.Open "SELECT Name,Code FROM BookMaster Where Type='F' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstItemList.ActiveConnection = Nothing
    If rstMatCList.State = adStateOpen Then rstMatCList.Close
        rstMatCList.Open "SELECT Name,Code FROM AccountMaster Where [Group]='*99999' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstMatCList.ActiveConnection = Nothing
    If Right(VchType, 2) = 48 Then
        Call FillList(ListView1, "List of Accounts...", rstAccountList)
        Call FillList(ListView2, "List of Voucher Series...", rstVchSeriesList)
        Call FillList(ListView3, "List of Items...", rstItemList)
        Call FillList(ListView4, "List of Material Centre...", rstMatCList)
    ElseIf VchType = 35 Or VchType = 36 Or VchType = 37 Or VchType = 38 Then
        Call FillList(ListView1, "List of Accounts...", rstAccountList)
        Call FillList(ListView2, "List of Item Groups...", rstItemGroupList)
        Call FillList(ListView3, "List of Items...", rstItemList)
        Call FillList(ListView4, "List of Material Centre...", rstMatCList)
    ElseIf VchType = 103 Or VchType = 104 Or VchType = 105 Or VchType = 31 Then
        Call FillList(ListView1, "List of Material Centre...", rstMatCList)
        Call FillList(ListView2, "List of Item Groups...", rstItemGroupList)
        Call FillList(ListView3, "List of Items...", rstItemList)
    ElseIf VchType = 49 Then
        Call FillList(ListView1, "List of Material Centre...", rstMatCList)
        Call FillList(ListView2, "List of Item Groups...", rstItemGroupList)
        Call FillList(ListView3, "List of Items...", rstItemList)
    ElseIf (Right(VchType, 2) >= 0 And Right(VchType, 2) <= 10) Or (Right(VchType, 2) >= 21 And Right(VchType, 2) <= 48) Or (Right(VchType, 2) >= 53 And Right(VchType, 2) <= 69 And Right(VchType, 2) <> 31) Then
        Call FillList(ListView1, "List of Accounts...", rstAccountList)
        Call FillList(ListView2, "List of Item Groups...", rstItemGroupList)
        Call FillList(ListView3, "List of Items...", rstItemList)
    ElseIf VchType <= 20 And VchType >= 11 Then
        Call FillList(ListView1, "List of Suppliers...", rstSupplierList)
        Call FillList(ListView2, "List of Accounts...", rstAccountList)
        Call FillList(ListView3, "List of Papers...", rstPaperList)
    End If
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}", True: KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4): KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyM Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3): KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2): KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1): KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then CloseForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstSupplierList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstItemGroupList)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstVchSeriesList)
    Call CloseRecordset(rstMatCList)
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then Cancel = True
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
        If VchType <= 20 And VchType >= 11 Then Else Call ItemSelection(False)
    End If
End Sub
Private Sub ListView3_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub ListView4_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView4.ListItems.Count
            ListView4.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 4 Then CloseForm Me: Exit Sub
    If VchType <= 10 Or VchType >= 21 Then
        PrintStockLedger
    ElseIf VchType <= 20 And VchType >= 11 Then
        PrintPaperLedger
    End If
End Sub
Private Sub PrintStockLedger()
Dim i As Integer

    On Error Resume Next
    
        FrmStockLedger.sDate = MhDateInput1.Text
        FrmStockLedger.eDate = MhDateInput2.Text
        FrmStockLedger.AccountList = SelectedItems(ListView1)
        
        If VchType = 0 And (Len(FrmStockLedger.AccountList) > 10 Or Len(FrmStockLedger.AccountList) < 8) Then
            If Len(FrmStockLedger.AccountList) < 8 Then MsgBox ("Please Select One-Godown"), vbCritical
            If Len(FrmStockLedger.AccountList) > 10 Then MsgBox ("Please Select One-Godown-Only"), vbCritical
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = False
            Next i
            ListView1.SetFocus
            Exit Sub
        Else
            MC = FrmStockLedger.AccountList
        End If
               
        If ((VchType >= 7 And VchType <= 10) Or (VchType >= 57 And VchType <= 60)) And (Len(FrmStockLedger.AccountList) > 10 Or Len(FrmStockLedger.AccountList) < 8) Then
            If Len(FrmStockLedger.AccountList) < 8 Then MsgBox ("Please Select One-Party Account"), vbCritical
            If Len(FrmStockLedger.AccountList) > 10 Then MsgBox ("Please Select One-Party Account Only"), vbCritical
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = False
            Next i
            ListView1.SetFocus
            Exit Sub
        End If
        FrmStockLedger.ItemGroupList = SelectedItems(ListView2)
        FrmStockLedger.ItemList = SelectedItems(ListView3)
        
        If ((VchType >= 25 And VchType <= 28) Or (VchType >= 65 And VchType <= 68)) And (Len(FrmStockLedger.ItemList) > 10 Or Len(FrmStockLedger.ItemList) < 8) Then
                If Len(FrmStockLedger.ItemList) < 8 Then MsgBox ("Please Select One-Item"), vbCritical
                If Len(FrmStockLedger.ItemList) > 10 Then MsgBox ("Please Select One-Item-Only"), vbCritical
            For i = 1 To ListView3.ListItems.Count
                ListView3.ListItems(i).Checked = False
            Next i
            ListView3.SetFocus
             Exit Sub
        End If
        
        If VchType = 35 Or VchType = 36 Or VchType = 37 Or VchType = 38 Then FrmStockLedger.MatCList = SelectedItems(ListView4)
        FrmStockLedger.sMcCode = "": FrmStockLedger.SCode = "": FrmStockLedger.oSCode = "":  FrmStockLedger.vtCode = "": FrmStockLedger.vDate = "":
        FrmStockLedger.VchType = VchType
        
        If (VchType = 31 Or VchType = 49) And (Len(FrmStockLedger.ItemList) > 10 Or Len(FrmStockLedger.ItemList) < 8) Then
            If Len(FrmStockLedger.ItemList) < 8 Then MsgBox ("Please Select One-Item"), vbCritical
            If Len(FrmStockLedger.ItemList) > 10 Then MsgBox ("Please Select One-Item-Only"), vbCritical
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Checked = False
        Next i
        ListView3.SetFocus
         Exit Sub
        Else
            FrmStockLedger.sMcCode = MC
        End If
        
        If (VchType = 2) And Len(FrmStockLedger.ItemList) < 8 Then MsgBox ("Please Select at Least One Item"), vbCritical: Exit Sub
        
        Load FrmStockLedger
        FrmStockLedger.Show
        Sendkeys "{TAB}", True:        Sendkeys "{TAB}", True
        CloseForm (Me)
End Sub
Private Sub PrintPaperLedger()
    On Error Resume Next
        FrmPaperLedger.sDate = GetDate(MhDateInput1.Text)
        FrmPaperLedger.eDate = GetDate(MhDateInput2.Text)
        FrmPaperLedger.SupplierList = SelectedItems(ListView1)
        FrmPaperLedger.AccountList = SelectedItems(ListView2)
        FrmPaperLedger.PaperList = SelectedItems(ListView3)
        FrmPaperLedger.VchType = VchType
        Load FrmPaperLedger
        FrmPaperLedger.Show
        CloseForm (Me)
End Sub
Private Sub ItemSelection(ByVal SelectAll As Boolean)
If Right(VchType, 2) = 48 Then

Else
    If rstItemList.State = adStateOpen Then rstItemList.Close
'    rstItemList.Open "SELECT Name,Code FROM BookMaster " & IIf(SelectAll, " Where Type='F' ", "Where Type='F' AND [Group] IN (" & SelectedItems(ListView2) & ")") & " ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstItemList.Open "WITH ItemGroupMaster AS (SELECT Name,Code FROM GeneralMaster WHERE Type IN ('5') AND Code IN (" & SelectedItems(ListView2) & ") UNION ALL SELECT P.Name,P.Code FROM GeneralMaster P INNER JOIN ItemGroupMaster C ON P.UnderGroup=C.Code) SELECT Name,Code FROM BookMaster WHERE [Group] IN (SELECT Code FROM ItemGroupMaster)  ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstItemList.ActiveConnection = Nothing
    ListView3.ListItems.Clear
    Call FillList(ListView3, "List of Items...", rstItemList)
End If
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If (VchType >= 11 And VchType <= 20) Or Right(VchType, 2) = 48 Then Else Call ItemSelection(False)
End Sub
