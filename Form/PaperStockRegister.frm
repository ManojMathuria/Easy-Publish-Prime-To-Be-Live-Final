VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPaperStockRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Stock Ledger"
   ClientHeight    =   9525
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   15045
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
   ScaleHeight     =   9525
   ScaleWidth      =   15045
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   4
         EndProperty
      EndProperty
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
            Picture         =   "PaperStockRegister.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PaperStockRegister.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PaperStockRegister.frx":0658
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PaperStockRegister.frx":0A33
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9060
      Left            =   45
      TabIndex        =   8
      Top             =   405
      Width           =   14955
      _Version        =   65536
      _ExtentX        =   26379
      _ExtentY        =   15981
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
      Picture         =   "PaperStockRegister.frx":0B45
      Begin VB.CheckBox Check8 
         Caption         =   "Show Total Party-wise"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   8280
         TabIndex        =   20
         Top             =   120
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Show Total By Paper UOM"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9720
         TabIndex        =   19
         Top             =   500
         Width           =   2535
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Show Total By Paper Size"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9720
         TabIndex        =   18
         Top             =   150
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Show Total By Paper GSM"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12360
         TabIndex        =   17
         Top             =   500
         Width           =   2535
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Show Total By Paper Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12360
         TabIndex        =   16
         Top             =   150
         Width           =   2535
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Centimeter-wise"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4080
         TabIndex        =   15
         Top             =   500
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Inclusive IN-Transit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5910
         TabIndex        =   14
         Top             =   50
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Exclusive IN-Transit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5910
         TabIndex        =   13
         Top             =   325
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Summarised Balance"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1900
         TabIndex        =   12
         Top             =   500
         Width           =   2040
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Only IN-Transit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5910
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Negative Balance"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   75
         TabIndex        =   2
         Top             =   500
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4080
         Left            =   0
         TabIndex        =   3
         Top             =   915
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   7197
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
         TabIndex        =   9
         Top             =   0
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
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
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "PaperStockRegister.frx":0B61
         Picture         =   "PaperStockRegister.frx":0B7D
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   2880
         TabIndex        =   10
         Top             =   0
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
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
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "PaperStockRegister.frx":0B99
         Picture         =   "PaperStockRegister.frx":0BB5
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4080
         Left            =   0
         TabIndex        =   4
         Top             =   4980
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   7197
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
      Begin MSComctlLib.ListView ListView3 
         Height          =   4080
         Left            =   6150
         TabIndex        =   5
         Top             =   915
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   7197
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
      Begin MSComctlLib.ListView ListView4 
         Height          =   4080
         Left            =   6150
         TabIndex        =   6
         Top             =   4980
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   7197
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
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Top             =   0
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calendar        =   "PaperStockRegister.frx":0BD1
         Caption         =   "PaperStockRegister.frx":0CE9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperStockRegister.frx":0D55
         Keys            =   "PaperStockRegister.frx":0D73
         Spin            =   "PaperStockRegister.frx":0DD1
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
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   4080
         TabIndex        =   1
         Top             =   0
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calendar        =   "PaperStockRegister.frx":0DF9
         Caption         =   "PaperStockRegister.frx":0F11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperStockRegister.frx":0F7D
         Keys            =   "PaperStockRegister.frx":0F9B
         Spin            =   "PaperStockRegister.frx":0FF9
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
   End
End
Attribute VB_Name = "FrmPaperStockRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset, rstPaperStockRegister As New ADODB.Recordset, rstPaperSizeList As New ADODB.Recordset, rstPaperGSMList As New ADODB.Recordset, rstPaperList As New ADODB.Recordset, rstAccountList As New ADODB.Recordset
Dim EMailID As String, Attachment As String, Message As String, OutputTo As String, PaperTbl
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    PaperTbl = "SELECT Code As Paper FROM PaperChild UNION " & _
               "SELECT Paper FROM PaperIOChild UNION " & _
               "SELECT Item As Paper FROM MaterialSVChild C INNER JOIN MaterialSVParent P ON P.Code=C.Code WHERE Category='2' AND ApprovedBy<>'' UNION " & _
               "SELECT Paper FROM PaperMVChild UNION " & _
               "SELECT Paper FROM PaperDNChild UNION " & _
               "SELECT Item As Paper FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE C.Category='2' AND LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' UNION " & _
               "SELECT Paper FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' UNION " & _
               "SELECT Paper FROM BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' UNION " & _
               "SELECT Paper1 As Paper FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' UNION " & _
               "SELECT Paper2 As Paper FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' UNION " & _
               "SELECT Paper4 As Paper FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*'"
    rstCompanyMaster.Open "SELECT PrintName,Phone,eMail FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    Check3_Click
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    MhDateInput2.Text = Format(IIf(Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd"), FinancialYearTo, Date), "dd-mm-yyyy")
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(FrmPaperStockRegister)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Sendkeys "{TAB}", True
       KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(FrmPaperStockRegister)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstPaperSizeList)
    Call CloseRecordset(rstPaperGSMList)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstPaperStockRegister)
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call LoadPaperGSMList
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (KeyCode = vbKeyA Or KeyCode = vbKeyD) And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
        Call LoadPaperGSMList
    End If
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call LoadPaperList
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (KeyCode = vbKeyA Or KeyCode = vbKeyD) And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
        Call LoadPaperList
    End If
End Sub
Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call LoadAccountList
End Sub
Private Sub ListView3_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (KeyCode = vbKeyA Or KeyCode = vbKeyD) And Shift = vbCtrlMask Then
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
        Call LoadAccountList
    End If
End Sub
Private Sub ListView4_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (KeyCode = vbKeyA Or KeyCode = vbKeyD) And Shift = vbCtrlMask Then
        For i = 1 To ListView4.ListItems.Count
            ListView4.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        OutputTo = "S"
        PrintPaperStockRegister
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintPaperStockRegister
    ElseIf Button.Index = 3 Then
        Call CloseForm(FrmPaperStockRegister)
    End If
End Sub
Private Sub LoadPaperGSMList()
    Dim SelectedPaperSizes
    If rstPaperGSMList.State = adStateOpen Then rstPaperGSMList.Close
    SelectedPaperSizes = SelectedItems(ListView1)
    If Check3.Value Then
        rstPaperGSMList.Open "SELECT DISTINCT GSM As Name,STR(GSM) As Code FROM PaperMaster P INNER JOIN (" & PaperTbl & ") As C ON P.Code=C.Paper WHERE " & IIf(SelectedPaperSizes = "''", "1=1", "IIF(Form='S',LTRIM(cmWidth)+'x'+LTRIM(cmLength)+'cm²',LTRIM(cmLength)+'cm-Reel') IN (" & SelectedPaperSizes & ")") & " ORDER BY GSM", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstPaperGSMList.Open "SELECT DISTINCT GSM As Name,STR(GSM) As Code FROM PaperMaster P INNER JOIN (" & PaperTbl & ") As C ON P.Code=C.Paper WHERE " & IIf(SelectedPaperSizes = "''", "1=1", "IIF(Form='S',LTRIM(inWidth)+'x'+LTRIM(inLength)+'in²',LTRIM(inLength)+'in-Reel') IN (" & SelectedPaperSizes & ")") & " ORDER BY GSM", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    rstPaperGSMList.ActiveConnection = Nothing
    ListView2.ListItems.Clear
    Call FillList(ListView2, "List of Paper GSMs...", rstPaperGSMList)
End Sub
Private Sub LoadAccountList()
    Dim SelectedPapers, AccountTbl
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    SelectedPapers = SelectedItems(ListView3)
    AccountTbl = "SELECT Account FROM PaperChild WHERE " & IIf(SelectedPapers = "''", "1=1", "Code IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT Account FROM PaperIOChild WHERE " & IIf(SelectedPapers = "''", "1=1", "Paper IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT Account FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE C.Category='2' AND ApprovedBy<>'' AND " & IIf(SelectedPapers = "''", "1=1", "C.Item IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT AccountFrom As Account FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE " & IIf(SelectedPapers = "''", "1=1", "C.Paper IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT AccountTo As Account FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE " & IIf(SelectedPapers = "''", "1=1", "C.Paper IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT Account FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE " & IIf(SelectedPapers = "''", "1=1", "C.Paper IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT Vendor As Account FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE C.Category='2' AND LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND " & IIf(SelectedPapers = "''", "1=1", "C.Item IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT RAccount As Account FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND " & IIf(SelectedPapers = "''", "1=1", "C.Paper IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT RAccount As Account FROM BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND " & IIf(SelectedPapers = "''", "1=1", "C.Paper IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT RAccount1 As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND " & IIf(SelectedPapers = "''", "1=1", "C.Paper1 IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT RAccount2 As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND " & IIf(SelectedPapers = "''", "1=1", "C.Paper2 IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT RAccount4 As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND " & IIf(SelectedPapers = "''", "1=1", "C.Paper4 IN (" & SelectedPapers & ")")
    rstAccountList.Open "SELECT Name,Code FROM AccountMaster P INNER JOIN (" & AccountTbl & ") As C ON P.Code=C.Account ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    ListView4.ListItems.Clear
    Call FillList(ListView4, "List of Godowns...", rstAccountList)
End Sub
Private Sub LoadPaperList()
    Dim SelectedPaperGSMs, SelectedPaperSizes
    If rstPaperList.State = adStateOpen Then rstPaperList.Close
    SelectedPaperSizes = SelectedItems(ListView1)
    SelectedPaperGSMs = SelectedItems(ListView2)
    If Check3.Value Then
        rstPaperList.Open "SELECT Name,Code FROM PaperMaster P INNER JOIN (" & PaperTbl & ") As C ON P.Code=C.Paper WHERE " & IIf(SelectedPaperSizes = "''" Or SelectedPaperGSMs = "''", "1=1", "IIF(Form='S',LTRIM(cmWidth)+'x'+LTRIM(cmLength)+'cm²',LTRIM(cmLength)+'cm-Reel') IN (" & SelectedPaperSizes & ") AND STR(GSM) IN (" & SelectedPaperGSMs & ")") & " ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstPaperList.Open "SELECT Name,Code FROM PaperMaster P INNER JOIN (" & PaperTbl & ") As C ON P.Code=C.Paper WHERE " & IIf(SelectedPaperSizes = "''" Or SelectedPaperGSMs = "''", "1=1", "IIF(Form='S',LTRIM(inWidth)+'x'+LTRIM(inLength)+'in²',LTRIM(inLength)+'in-Reel') IN (" & SelectedPaperSizes & ") AND STR(GSM) IN (" & SelectedPaperGSMs & ")") & " ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    rstPaperList.ActiveConnection = Nothing
    ListView3.ListItems.Clear
    Call FillList(ListView3, "List of Papers...", rstPaperList)
End Sub
Private Sub PrintPaperStockRegister()
    Dim OpBal As String, SQL As String, SelectedPapers As String, SelectedAccounts As String
    Dim oExcel As Object, StkIn As String, StkOut As String, i As Integer
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    Dim CRXParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim CRXParamDef As CRAXDRT.ParameterFieldDefinition
    rptPaperStockRegister.Text11.SetText "Paper Stock Register (" & IIf(Check2.Value, "Summarised", "Detailed") & ")"
    rptPaperStockRegister.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPaperStockRegister.Text13.SetText "From [" + Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy") + "] [" & IIf(Option1.Value, "Including In-Transit", IIf(Option2.Value, "Excluding In-Transit", "In-Transit Only")) & "]"
    If rstPaperStockRegister.State = adStateOpen Then rstPaperStockRegister.Close
    SelectedPapers = SelectedItems(ListView3)
    SelectedAccounts = SelectedItems(ListView4)
    If Option3.Value Then 'Only In-Transit
        OpBal = "(SELECT IIF(SUM(QuantitySheets) IS NULL,0,SUM(QuantitySheets)) FROM PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code WHERE Paper=M2.Code AND Account=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "' AND (P.DeliveryEndDate IS NULL AND P.BillNo=''))"
    Else
        OpBal = "(SELECT IIF(SUM(OpBalSheets) IS NULL,0,SUM(OpBalSheets)) FROM PaperChild WHERE Code=M2.Code AND Account=M1.Code)+" & _
                        "(SELECT IIF(SUM(QuantitySheets) IS NULL,0,SUM(QuantitySheets)) FROM PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code WHERE Paper=M2.Code AND Account=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "' AND " & IIf(Option1.Value, "1=1", "(P.DeliveryEndDate IS NOT NULL OR P.BillNo<>'')") & ")+" & _
                        "(SELECT IIF(SUM(Quantity) IS NULL,0,SUM(PARSENAME(Quantity,2)*1)*CONVERT(DECIMAL(12,3),M3.Value1)+SUM(PARSENAME(Quantity,1)*1)) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND ApprovedBy<>'' AND Item=M2.Code AND Quantity>=0 AND Account=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "')-" & _
                        "(SELECT IIF(SUM(Quantity) IS NULL,0,SUM(PARSENAME(0-Quantity,2)*1)*CONVERT(DECIMAL(12,3),M3.Value1)+SUM(PARSENAME(0-Quantity,1)*1)) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND ApprovedBy<>'' AND Item=M2.Code AND Quantity<0 AND Account=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "')-" & _
                        "(SELECT IIF(SUM(QuantitySheets) IS NULL,0,SUM(QuantitySheets)) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M2.Code AND AccountFrom=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "')+" & _
                        "(SELECT IIF(SUM(QuantitySheets) IS NULL,0,SUM(QuantitySheets)) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M2.Code AND AccountTo=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "')-" & _
                        "(SELECT IIF(SUM(Quantity) IS NULL,0,SUM(PARSENAME(0-Quantity,2)*1)*CONVERT(DECIMAL(12,3),M3.Value1)+SUM(PARSENAME(0-Quantity,1)*1)) FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE Paper=M2.Code AND Account=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "' AND Quantity<0)+" & _
                        "(SELECT IIF(SUM(Quantity) IS NULL,0,SUM(PARSENAME(Quantity,2)*1)*CONVERT(DECIMAL(12,3),M3.Value1)+SUM(PARSENAME(Quantity,1)*1)) FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE Paper=M2.Code AND Account=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "' AND Quantity>=0)-" & _
                        "(SELECT IIF(SUM(ROUND(C.TotalConsumption,0)) IS NULL,0,SUM(ROUND(C.TotalConsumption,0))) FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Category='2' AND Item=M2.Code AND Vendor=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "')-" & _
                        "(SELECT IIF(SUM(PaperConsumptionSheets) IS NULL,0,SUM(PaperConsumptionSheets)) FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=M2.Code AND RAccount=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "')-" & _
                        "(SELECT IIF(SUM(PaperConsumptionSheets) IS NULL,0,SUM(PaperConsumptionSheets)) FROM BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=M2.Code AND RAccount=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "')-" & _
                        "(SELECT IIF(SUM(PaperConsumptionSheets1) IS NULL,0,SUM(PaperConsumptionSheets1)) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper1=M2.Code AND RAccount1=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "')-" & _
                        "(SELECT IIF(SUM(PaperConsumptionSheets2) IS NULL,0,SUM(PaperConsumptionSheets2)) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper2=M2.Code AND RAccount2=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "')-" & _
                        "(SELECT IIF(SUM(PaperConsumptionSheets4) IS NULL,0,SUM(PaperConsumptionSheets4)) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper4=M2.Code AND RAccount4=M1.Code AND Date<'" & GetDate(MhDateInput1.Text) & "')"
    End If
    'VchNo,VchDate,VchType,Particulars,BookQuantity,Forms,Quantity,GSM,GodownName,SizeName,PaperName
    If Check3.Value Then    'cmwise
        If Option3.Value Then 'Only In-Transit
            SQL = SQL + "SELECT * FROM (SELECT '' As VchNo,'" & Format(CDate(GetDate(MhDateInput1.Text)) - 1, "dd-MMM-yyyy") & "' As VchDate,'OB' As VchType,'Opening Balance' As Particulars,0 As BookQuantity,0.00 As Forms," & OpBal & " As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM AccountMaster M1,(PaperMaster M2 INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code) WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ")) As Tbl WHERE Quantity<>0 UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'PI' As VchType,' IN (FROM : '+(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.Supplier)+' Challan No.:'+P.BiltyNo+')' As Particulars,0 As BookQuantity,0.00 As Forms,QuantitySheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code "
            SQL = SQL + "WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Option1.Value, "1=1", "(P.DeliveryEndDate IS NULL AND P.BillNo='')")
            SQL = SQL + "ORDER BY GodownName,SizeName,PaperName,VchDate,VchNo"
        Else
            SQL = SQL + "SELECT * FROM (SELECT '' As VchNo,'" & Format(CDate(GetDate(MhDateInput1.Text)) - 1, "dd-MMM-yyyy") & "' As VchDate,'OB' As VchType,'Opening Balance' As Particulars,0 As BookQuantity,0.00 As Forms," & OpBal & " As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM AccountMaster M1,(PaperMaster M2 INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code) WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ")) As Tbl WHERE Quantity<>0 UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'PI' As VchType,' IN (FROM : '+(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.Supplier)+' Challan No.:'+P.BiltyNo+')' As Particulars,0 As BookQuantity,0.00 As Forms,QuantitySheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Option1.Value, "1=1", "(P.DeliveryEndDate IS NOT NULL OR P.BillNo<>'')") & " UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'SI' As VchType,'Stock (Generated)' As Particulars,0 As BookQuantity,0.00 As Forms,PARSENAME(Quantity,2)*CONVERT(DECIMAL(12,3),M3.Value1)+PARSENAME(Quantity,1) As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Item=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE C.Category='2' AND P.ApprovedBy<>'' AND C.Quantity>=0 AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'SR' As VchType,'Stock (Consumed)' As Particulars,0 As BookQuantity,0.00 As Forms,PARSENAME(0-Quantity,2)*CONVERT(DECIMAL(12,3),M3.Value1)+PARSENAME(0-Quantity,1) As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Item=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE C.Category='2' AND P.ApprovedBy<>'' AND C.Quantity<0 AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'MO' As VchType,'Out (To : '+(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.AccountTo)+')' As Particulars,0 As BookQuantity,0.00 As Forms,QuantitySheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.AccountFrom=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'MI' As VchType,' IN (FROM : '+(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.AccountFrom)+')' As Particulars,0 As BookQuantity,0.00 As Forms,QuantitySheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.AccountTo=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'DN' As VchType,'Debit Note' As Particulars,0 As BookQuantity,0.00 As Forms,PARSENAME(0-Quantity,2)*CONVERT(DECIMAL(12,3),M3.Value1)+PARSENAME(0-Quantity,1) As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND Quantity<0 UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'CN' As VchType,'Credit Note' As Particulars,0 As BookQuantity,0.00 As Forms,PARSENAME(Quantity,2)*CONVERT(DECIMAL(12,3),M3.Value1)+PARSENAME(Quantity,1) As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND Quantity>=0 UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'PC' As VchType,'(UFG : '+(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+')' As Particulars,OrderQuantity As BookQuantity,0.00 As Forms,ROUND(C.TotalConsumption,0) As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Vendor=M1.Code) INNER JOIN PaperMaster M2 ON C.Item=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND C.Category='2' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,' (SF : '+(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') - Wastage-'+LTRIM([PaperWastage%])+'%/'+LTRIM(PaperWastageFinal)+'-Units Ref : '+Ref As Particulars,ActualQuantity As BookQuantity,0.00 As Forms,PaperConsumptionSheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.RAccount=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,'(CF:'+(SELECT STUFF((SELECT ','+(LTRIM(I.PrintName)+'-'+LTRIM([Ups/Plate])+'Ups-('+LTRIM(FrontPrintingColor)+'+'+LTRIM(BackPrintingColor)+')Col-'+LTRIM(ActualQuantity)) FROM BookPOChild0901 T INNER JOIN BookMaster I ON T.Book=I.Code WHERE T.Code=P.Code ORDER BY T.Code,I.PrintName FOR XML PATH('')),1,1,''))+' Wastage-'+LTRIM([PaperWastage%])+'%/'+LTRIM(PaperWastageFinal)+'-Units)' As Particulars,(SELECT MIN(ActualQuantity) FROM BookPOChild0901 WHERE Code=P.Code) As BookQuantity,0.00 As Forms,PaperConsumptionSheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'PC'+P.Code As VchCode " & _
                                    "FROM (((BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.RAccount=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,' (MF : '+(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') - Wastage-'+LTRIM([PaperWastage1%])+'%/'+LTRIM(PaperWastageFinal1)+'-Units Ref : '+Ref As Particulars,ActualQuantity As BookQuantity,Forms1 As Forms,PaperConsumptionSheets1 As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.RAccount1=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper1=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,' (MF : '+(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') - Wastage-'+LTRIM([PaperWastage2%])+'%/'+LTRIM(PaperWastageFinal2)+'-Units Ref : '+Ref As Particulars,ActualQuantity As BookQuantity,Forms2 As Forms,PaperConsumptionSheets2 As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.RAccount2=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper2=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,' (MF : '+(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') - Wastage-'+LTRIM([PaperWastage4%])+'%/'+LTRIM(PaperWastageFinal4)+'-Units Ref : '+Ref As Particulars,ActualQuantity As BookQuantity,Forms4 As Forms,PaperConsumptionSheets4 As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.cmWidth)+'x'+LTRIM(M2.cmLength)+'cm²',LTRIM(M2.cmLength)+'cm-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.RAccount4=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper4=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' "
            SQL = SQL + "ORDER BY GodownName,SizeName,PaperName,VchDate,VchNo"
        End If
    Else
        If Option3.Value Then 'Only In-Transit
            SQL = SQL + "SELECT * FROM (SELECT '' As VchNo,'" & Format(CDate(GetDate(MhDateInput1.Text)) - 1, "dd-MMM-yyyy") & "' As VchDate,'OB' As VchType,'Opening Balance' As Particulars,0 As BookQuantity,0.00 As Forms," & OpBal & " As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM AccountMaster M1,(PaperMaster M2 INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code) WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ")) As Tbl WHERE Quantity<>0 UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'PI' As VchType,' IN (FROM : '+(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.Supplier)+' Challan No.:'+P.BiltyNo+')' As Particulars,0 As BookQuantity,0.00 As Forms,QuantitySheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code "
            SQL = SQL + "WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Option1.Value, "1=1", "(P.DeliveryEndDate IS NULL AND P.BillNo='')")
            SQL = SQL + "ORDER BY GodownName,SizeName,PaperName,VchDate,VchNo"
        Else
            SQL = SQL + "SELECT * FROM (SELECT '' As VchNo,'" & Format(CDate(GetDate(MhDateInput1.Text)) - 1, "dd-MMM-yyyy") & "' As VchDate,'OB' As VchType,'Opening Balance' As Particulars,0 As BookQuantity,0.00 As Forms," & OpBal & " As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM AccountMaster M1,(PaperMaster M2 INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code) WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ")) As Tbl WHERE Quantity<>0 UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'PI' As VchType,' IN (FROM : '+(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.Supplier)+' Challan No.:'+P.BiltyNo+')' As Particulars,0 As BookQuantity,0.00 As Forms,QuantitySheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Option1.Value, "1=1", "(P.DeliveryEndDate IS NOT NULL OR P.BillNo<>'')") & " UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'SI' As VchType,'Stock (Generated)' As Particulars,0 As BookQuantity,0.00 As Forms,PARSENAME(Quantity,2)*CONVERT(DECIMAL(12,3),M3.Value1)+PARSENAME(Quantity,1) As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Item=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE C.Category='2' AND P.ApprovedBy<>'' AND C.Quantity>=0 AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'SR' As VchType,'Stock (Consumed)' As Particulars,0 As BookQuantity,0.00 As Forms,PARSENAME(0-Quantity,2)*CONVERT(DECIMAL(12,3),M3.Value1)+PARSENAME(0-Quantity,1) As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Item=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE C.Category='2' AND P.ApprovedBy<>'' AND C.Quantity<0 AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'MO' As VchType,'Out (To : '+(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.AccountTo)+')' As Particulars,0 As BookQuantity,0.00 As Forms,QuantitySheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.AccountFrom=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'MI' As VchType,' IN (FROM : '+(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.AccountFrom)+')' As Particulars,0 As BookQuantity,0.00 As Forms,QuantitySheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.AccountTo=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'DN' As VchType,'Debit Note' As Particulars,0 As BookQuantity,0.00 As Forms,PARSENAME(0-Quantity,2)*CONVERT(DECIMAL(12,3),M3.Value1)+PARSENAME(0-Quantity,1) As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND Quantity<0 UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'CN' As VchType,'Credit Note' As Particulars,0 As BookQuantity,0.00 As Forms,PARSENAME(Quantity,2)*CONVERT(DECIMAL(12,3),M3.Value1)+PARSENAME(Quantity,1) As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND Quantity>=0 UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,P.Date As VchDate,'PC' As VchType,'(UFG : '+(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+')' As Particulars,OrderQuantity As BookQuantity,0.00 As Forms,ROUND(C.TotalConsumption,0) As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Vendor=M1.Code) INNER JOIN PaperMaster M2 ON C.Item=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND C.Category='2' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,' (SF : '+(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') - Wastage-'+LTRIM([PaperWastage%])+'%/'+LTRIM(PaperWastageFinal)+'-Units Ref : '+Ref As Particulars,ActualQuantity As BookQuantity,0.00 As Forms,PaperConsumptionSheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.RAccount=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,'(CF:'+(SELECT STUFF((SELECT ','+(LTRIM(I.PrintName)+'-'+LTRIM([Ups/Plate])+'Ups-('+LTRIM(FrontPrintingColor)+'+'+LTRIM(BackPrintingColor)+')Col-'+LTRIM(ActualQuantity)) FROM BookPOChild0901 T INNER JOIN BookMaster I ON T.Book=I.Code WHERE T.Code=P.Code ORDER BY T.Code,I.PrintName FOR XML PATH('')),1,1,''))+' Wastage-'+LTRIM([PaperWastage%])+'%/'+LTRIM(PaperWastageFinal)+'-Units)' As Particulars,(SELECT MIN(ActualQuantity) FROM BookPOChild0901 WHERE Code=P.Code) As BookQuantity,0.00 As Forms,PaperConsumptionSheets As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'PC'+P.Code As VchCode " & _
                                    "FROM (((BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.RAccount=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,' (MF : '+(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') - Wastage-'+LTRIM([PaperWastage1%])+'%/'+LTRIM(PaperWastageFinal1)+'-Units Ref : '+Ref As Particulars,ActualQuantity As BookQuantity,Forms1 As Forms,PaperConsumptionSheets1 As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.RAccount1=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper1=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,' (MF : '+(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') - Wastage-'+LTRIM([PaperWastage2%])+'%/'+LTRIM(PaperWastageFinal2)+'-Units Ref : '+Ref As Particulars,ActualQuantity As BookQuantity,Forms2 As Forms,PaperConsumptionSheets2 As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.RAccount2=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper2=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' UNION ALL "
            SQL = SQL + "SELECT LTRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,' (MF : '+(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') - Wastage-'+LTRIM([PaperWastage4%])+'%/'+LTRIM(PaperWastageFinal4)+'-Units Ref : '+Ref As Particulars,ActualQuantity As BookQuantity,Forms4 As Forms,PaperConsumptionSheets4 As Quantity,M2.GSM,'Party Name : '+LTRIM(M1.PrintName) As GodownName,'Size Name : '+IIF(Form='S',LTRIM(M2.inWidth)+'x'+LTRIM(M2.inLength)+'in²',LTRIM(M2.inLength)+'in-Reel') As SizeName,'Paper Name : '+LTRIM(M2.PrintName) As PaperName,M3.Value1 As SPU,'UOM : '+LTRIM(M3.PrintName)+'='+LTRIM(M3.Value1) As UOM,'' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.RAccount4=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper4=M2.Code) INNER JOIN GeneralMaster M3 ON M2.UOM=M3.Code " & _
                                    "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' "
            SQL = SQL + "ORDER BY GodownName,SizeName,PaperName,VchDate,VchNo"
        End If
    End If
    rstPaperStockRegister.Open SQL, cnDatabase, adOpenKeyset, adLockOptimistic
    rstPaperStockRegister.ActiveConnection = Nothing
    Screen.MousePointer = vbNormal
    If rstPaperStockRegister.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rptPaperStockRegister.Database.SetDataSource rstPaperStockRegister, 3, 1
    rptPaperStockRegister.DiscardSavedData
    If Check4.Value = 0 Then rptPaperStockRegister.Section11.Suppress = True
    If Check5.Value = 0 Then rptPaperStockRegister.Section19.Suppress = True
    If Check6.Value = 0 Then rptPaperStockRegister.Section9.Suppress = True
    If Check7.Value = 0 Then rptPaperStockRegister.Section21.Suppress = True
    If Check8.Value = 0 Then rptPaperStockRegister.Section7.Suppress = True
    Set CRXParamDefs = rptPaperStockRegister.ParameterFields
    For Each CRXParamDef In CRXParamDefs
        If CRXParamDef.ParameterFieldName = "PF1" Then
            CRXParamDef.SetCurrentValue (IIf(Check1.Value, 0, 0.1))
        ElseIf CRXParamDef.ParameterFieldName = "PF2" Then
            CRXParamDef.SetCurrentValue (IIf(Check2.Value, "S", "D"))
        End If
    Next
    rptPaperStockRegister.EnableParameterPrompting = False
    EMailID = "xxxxxxxxxx"
    Attachment = "Paper Stock Register"
    Message = "Dear Sir,<Br>Please find attached herewith Paper Stock Register From [" & Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") & "] To [" & Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy") & "] for doing the needful at your end.<Br>Kindly inform us if you find any discrepancy in the same and acknowledge the receipt of mail.<Br><Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
    If OutputTo = "S" Then
        FrmReportViewer.EMailID = EMailID
        FrmReportViewer.Subject = "Paper Stock Register"
        FrmReportViewer.Attachment = Attachment
        FrmReportViewer.Message = Message
        Set FrmReportViewer.Report = rptPaperStockRegister
        FrmReportViewer.Show vbModal
    Else
        rptPaperStockRegister.PaperSource = crPRBinAuto
        rptPaperStockRegister.PrintOut
    End If
    Set rptPaperStockRegister = Nothing
    On Error GoTo 0
End Sub
Private Sub Check3_Click()
    If rstPaperSizeList.State = adStateOpen Then rstPaperSizeList.Close
    If Check3.Value Then
        rstPaperSizeList.Open "SELECT Name,Name As Code FROM (SELECT DISTINCT IIF(Form='S',LTRIM(cmWidth)+'x'+LTRIM(cmLength)+'cm²',LTRIM(cmLength)+'cm-Reel') As Name FROM PaperMaster P INNER JOIN (" & PaperTbl & ") As C ON P.Code=C.Paper WHERE Form<>'') As Tbl ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstPaperSizeList.Open "SELECT Name,Name As Code FROM (SELECT DISTINCT IIF(Form='S',LTRIM(inWidth)+'x'+LTRIM(inLength)+'in²',LTRIM(inLength)+'in-Reel') As Name FROM PaperMaster P INNER JOIN (" & PaperTbl & ") As C ON P.Code=C.Paper WHERE Form<>'') As Tbl ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    rstPaperSizeList.ActiveConnection = Nothing
    ListView1.ListItems.Clear
    Call FillList(ListView1, "List of Paper Sizes...", rstPaperSizeList)
    Call LoadPaperGSMList
    Call LoadPaperList
    Call LoadAccountList
End Sub
