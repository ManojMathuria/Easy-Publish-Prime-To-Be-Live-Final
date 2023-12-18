VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAccountSelectionList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Account Selection List...."
   ClientHeight    =   9150
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   9705
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
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
            Object.ToolTipText     =   "Print Preview [Alt+V]"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print [Alt+P]"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mail [Alt+E]"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit [Escape]"
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
            Picture         =   "AccountSelectionList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AccountSelectionList.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AccountSelectionList.frx":0658
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AccountSelectionList.frx":076A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   8830
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   9675
      _Version        =   65536
      _ExtentX        =   17066
      _ExtentY        =   15593
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
      Picture         =   "AccountSelectionList.frx":087C
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   7
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
         Picture         =   "AccountSelectionList.frx":0898
         Picture         =   "AccountSelectionList.frx":08B4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   2400
         TabIndex        =   8
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
         Caption         =   " To"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountSelectionList.frx":08D0
         Picture         =   "AccountSelectionList.frx":08EC
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
         Calendar        =   "AccountSelectionList.frx":0908
         Caption         =   "AccountSelectionList.frx":0A20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountSelectionList.frx":0A8C
         Keys            =   "AccountSelectionList.frx":0AAA
         Spin            =   "AccountSelectionList.frx":0B08
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
         Left            =   3270
         TabIndex        =   1
         Top             =   0
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "AccountSelectionList.frx":0B30
         Caption         =   "AccountSelectionList.frx":0C48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountSelectionList.frx":0CB4
         Keys            =   "AccountSelectionList.frx":0CD2
         Spin            =   "AccountSelectionList.frx":0D30
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
         Height          =   7980
         Left            =   0
         TabIndex        =   9
         Top             =   315
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   14076
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
         Height          =   7980
         Left            =   4830
         TabIndex        =   10
         Top             =   315
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   14076
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
         Top             =   30
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
         Top             =   30
         Visible         =   0   'False
         Width           =   750
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
         Top             =   30
         Visible         =   0   'False
         Width           =   630
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   420
         Index           =   2
         Left            =   80
         TabIndex        =   11
         Top             =   8340
         Width           =   9540
         _Version        =   65536
         _ExtentX        =   16828
         _ExtentY        =   741
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
         Caption         =   " Ctrl+A->Select ALL  Ctrl+D->Deselect ALL  Alt+V->Print Preview  Alt+P->Print Alt+M->Email  Escape->Exit"
         FillColor       =   8421504
         TextColor       =   16777215
         Picture         =   "AccountSelectionList.frx":0D58
         Picture         =   "AccountSelectionList.frx":0D74
      End
      Begin MSForms.ComboBox ComboBox2 
         Height          =   405
         Left            =   3280
         TabIndex        =   13
         Top             =   8
         Visible         =   0   'False
         Width           =   1545
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
         Height          =   405
         Left            =   860
         TabIndex        =   12
         Top             =   15
         Visible         =   0   'False
         Width           =   1545
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
Attribute VB_Name = "FrmAccountSelectionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstAccountList As New ADODB.Recordset, rstAccountGroupList As New ADODB.Recordset, rstCompList As New ADODB.Recordset
Public VchType As String, ComboFlag As Boolean, PrintFlag As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If VchType >= 0 Then ListView1.BackColor = RGB(255, 255, 240): ListView2.BackColor = RGB(255, 255, 240): MhDateInput1.BackColor = RGB(255, 255, 240): MhDateInput2.BackColor = RGB(255, 255, 240):
    If VchType >= 0 Then Me.Caption = " Selection List....Account Ledger"
    CenterForm Me
    If VchType = 1 Then
        MhDateInput1.Visible = False: MhDateInput2.Visible = False
        ComboBox1.Visible = True
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
        
        ComboBox2.Visible = True
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
    ComboFlag = True
    BusySystemIndicator True
    rstCompList.Open "SELECT TOP 1 FinancialYearFrom  FROM CompanyMaster WHERE FYCode='" & FYCode & "' ORDER BY FYCode", cnDatabase, adOpenForwardOnly, adLockReadOnly
    MhDateInput1.Text = Format(rstCompList.Fields("FinancialYearFrom").Value, "dd-mm-yyyy")
    MhDateInput2.Text = IIf(Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd"), Format(FinancialYearTo, "dd-mm-yyyy"), Format(Date, "dd-mm-yyyy"))
    rstAccountList.Open "SELECT Name,Code FROM AccountMaster WHERE [Group]<>'*99999' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountGroupList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type IN ('12','26') ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    If VchType >= 0 And VchType <= 30 Then
        Call FillList(ListView1, "List of Accounts Groups...", rstAccountGroupList)
        Call FillList(ListView2, "List of Accounts ...", rstAccountList)
    End If
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub ComboBox1_Change()
Dim NewDate As Date
If ComboFlag = True Then If ComboBox2.ListIndex < ComboBox1.ListIndex Then If MsgBox("  End Month can't be lower than the Start Month. ", vbInformation) = vbOK Then ComboBox1.ListIndex = ComboBox2.ListIndex: Exit Sub
    Call AddDate(FinancialYearFrom, NewDate, ComboBox1.ListIndex)
    MhDateInput1.Value = NewDate
End Sub
Private Sub ComboBox2_Change()
Dim NewDate As Date
If ComboFlag = True Then If ComboBox2.ListIndex < ComboBox1.ListIndex Then If MsgBox("  End Month can't be lower than the Start Month. ", vbInformation) = vbOK Then ComboBox2.ListIndex = ComboBox1.ListIndex: Exit Sub
    Call AddDate(FinancialYearTo, NewDate, ComboBox2.ListIndex - 11)
    MhDateInput2.Value = NewDate
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
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If VchType <= 29 Or VchType >= 0 Then
        If Button.Index = 1 Then    'Crystal Preview
            FrmAccountLedger.OutputTo = ("S")
        ElseIf Button.Index = 2 Then 'Crystal Print
            PrintFlag = True
            FrmAccountLedger.OutputTo = ("P")
        ElseIf Button.Index = 3 Then 'Crystal Email
            PrintFlag = True
            FrmAccountLedger.OutputTo = ("M")
        ElseIf Button.Index = 4 Then 'Exit
            CloseForm Me: Exit Sub
        End If
            PrintAccountLedger
    End If
End Sub
Private Sub PrintAccountLedger()
Dim i As Integer
    On Error Resume Next
        FrmAccountLedger.FromTo = ComboBox1.ListIndex + 4 & " AND " & ComboBox2.ListIndex + 4
        FrmAccountLedger.sDate = MhDateInput1.Text
        FrmAccountLedger.eDate = MhDateInput2.Text
        FrmAccountLedger.AccountList = SelectedItems(ListView2)
        If VchType >= 0 And VchType <= 30 And Len(FrmAccountLedger.AccountList) > 10 Then MsgBox ("Please Select One-Party Account Only"), vbCritical: Exit Sub
        FrmAccountLedger.AccountGroupList = SelectedItems(ListView1)
        FrmAccountLedger.AccountList = SelectedItems(ListView2)
        FrmAccountLedger.VchType = VchType
    If VchType >= 0 And VchType <= 30 And (Len(FrmAccountLedger.AccountList) > 10 Or Len(FrmAccountLedger.AccountList) < 8) Then
            If Len(FrmAccountLedger.AccountList) < 8 Then MsgBox ("Please Select One-Account"), vbCritical
            If Len(FrmAccountLedger.AccountList) > 10 Then MsgBox ("Please Select One-Account-Only"), vbCritical
            For i = 1 To ListView1.ListItems.Count
                ListView2.ListItems(i).Checked = False
            Next i
            ListView2.SetFocus
            Exit Sub
    End If
        Load FrmAccountLedger
        If PrintFlag = False Then FrmAccountLedger.Show
        PrintFlag = False
        CloseForm (Me)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then CloseForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstAccountGroupList)
    Call CloseRecordset(rstCompList)
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
        If VchType > 0 And VchType < 29 Then Call AccountSelection
    End If
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call AccountSelection
End Sub
Private Sub AccountSelection()
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "WITH AccountGroupMaster AS (SELECT Name,Code FROM GeneralMaster WHERE Type IN ('12','26') AND Code IN (" & SelectedItems(ListView1) & ") UNION ALL SELECT P.Name,P.Code FROM GeneralMaster P INNER JOIN AccountGroupMaster C ON P.UnderGroup=C.Code) SELECT Name,Code FROM AccountMaster WHERE [Group] IN (SELECT Code FROM AccountGroupMaster)  Order By Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    ListView2.ListItems.Clear
    Call FillList(ListView2, "List of Accounts...", rstAccountList)
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
