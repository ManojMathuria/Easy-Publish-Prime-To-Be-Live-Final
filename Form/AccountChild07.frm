VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmAccountChild07 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Misc Operation Rate Details"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
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
   ScaleHeight     =   2370
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   6330
      Picture         =   "AccountChild07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   6330
      Picture         =   "AccountChild07.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   2135
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   105
      Width           =   5610
      _Version        =   65536
      _ExtentX        =   9895
      _ExtentY        =   3766
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
      Picture         =   "AccountChild07.frx":0204
      Begin VB.TextBox Text4 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   1
         Top             =   730
         Width           =   3925
      End
      Begin VB.TextBox Text1 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   0
         Top             =   425
         Width           =   3925
      End
      Begin VB.TextBox Text2 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   100
         Width           =   3925
      End
      Begin VB.TextBox Text3 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1050
         Width           =   3925
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   1050
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Size"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild07.frx":0220
         Picture         =   "AccountChild07.frx":023C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   105
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Party Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild07.frx":0258
         Picture         =   "AccountChild07.frx":0274
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   425
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Operation"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild07.frx":0290
         Picture         =   "AccountChild07.frx":02AC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   730
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Calc Mode"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild07.frx":02C8
         Picture         =   "AccountChild07.frx":02E4
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   1560
         TabIndex        =   3
         Top             =   1360
         Width           =   3925
         _Version        =   65536
         _ExtentX        =   6923
         _ExtentY        =   582
         Calculator      =   "AccountChild07.frx":0300
         Caption         =   "AccountChild07.frx":0320
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild07.frx":038C
         Keys            =   "AccountChild07.frx":03AA
         Spin            =   "AccountChild07.frx":03F4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "########0.000"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "########0.000"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   13
         Top             =   1360
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild07.frx":041C
         Picture         =   "AccountChild07.frx":0438
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   1560
         TabIndex        =   4
         Top             =   1680
         Width           =   3930
         _Version        =   65536
         _ExtentX        =   6923
         _ExtentY        =   582
         Calculator      =   "AccountChild07.frx":0454
         Caption         =   "AccountChild07.frx":0474
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild07.frx":04E0
         Keys            =   "AccountChild07.frx":04FE
         Spin            =   "AccountChild07.frx":0548
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "########0.000"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "########0.000"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Range"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild07.frx":0570
         Picture         =   "AccountChild07.frx":058C
      End
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   930
      Index           =   2
      Left            =   5760
      TabIndex        =   15
      Top             =   1320
      Width           =   1440
      _Version        =   65536
      _ExtentX        =   2540
      _ExtentY        =   1640
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
      Caption         =   "Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D->Delete  Ctrl+S->Save"
      AutoSize        =   -1  'True
      FillColor       =   8421504
      TextColor       =   16777215
      Picture         =   "AccountChild07.frx":05A8
      Multiline       =   -1  'True
      GlobalMem       =   -1  'True
      Picture         =   "AccountChild07.frx":05C4
   End
End
Attribute VB_Name = "FrmAccountChild07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rstAccountChild As New ADODB.Recordset
Public AccountName As String
Dim rstSizeList As New ADODB.Recordset, rstOperationList As New ADODB.Recordset, rstCalcModeList As New ADODB.Recordset
Dim SizeCode As String, OperationCode As String, CalcModeCode As String
Private Sub Form_Load()
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
    Text2.Text = Trim(AccountName)
    rstOperationList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='7' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstOperationList.ActiveConnection = Nothing
    rstCalcModeList.Open "SELECT Name As Col0,Value1,Code FROM GeneralMaster WHERE Type='20' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstCalcModeList.ActiveConnection = Nothing
    rstSizeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type IN ('11', '1') ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstSizeList.ActiveConnection = Nothing
    ClearFields
    LoadFields
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        cmdProceed_Click
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        cmdCancel_Click
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set rstAccountChild = Nothing
    Call CloseRecordset(rstOperationList)
    Call CloseRecordset(rstCalcModeList)
    Call CloseRecordset(rstSizeList)
End Sub
Private Sub ClearFields()
    Text1.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    MhRealInput1.Value = 0
    MhRealInput2.Value = 0
    SizeCode = ""
End Sub
Private Sub LoadFields()
    Dim Cnt As Integer
    If rstAccountChild.RecordCount = 0 Then Exit Sub
    If Not IsNull(rstAccountChild.Fields("OperationName").Value) Then
        OperationCode = rstAccountChild.Fields("LaminationType").Value
        Text1.Text = rstAccountChild.Fields("OperationName").Value
        CalcModeCode = rstAccountChild.Fields("CalcMode").Value
        Text4.Text = rstAccountChild.Fields("CalcModeName").Value
        SizeCode = rstAccountChild.Fields("Size").Value
        Text3.Text = IIf(IsNull(rstAccountChild.Fields("SizeName").Value), "", rstAccountChild.Fields("SizeName").Value)
        MhRealInput1.Value = Val(rstAccountChild.Fields("Rate").Value)
        MhRealInput2.Value = Val(rstAccountChild.Fields("Range").Value)
    End If
End Sub
Private Sub SaveFields()
    rstAccountChild.Fields("LaminationType").Value = OperationCode
    rstAccountChild.Fields("OperationName").Value = Trim(Text1.Text)
    rstAccountChild.Fields("CalcMode").Value = CalcModeCode
    rstAccountChild.Fields("CalcModeName").Value = Trim(Text4.Text)
    rstAccountChild.Fields("Size").Value = SizeCode
    rstAccountChild.Fields("SizeName").Value = Trim(Text3.Text)
    rstAccountChild.Fields("Rate").Value = MhRealInput1.Value
    rstAccountChild.Fields("Range").Value = MhRealInput2.Value
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text1.Text)
        If rstOperationList.RecordCount = 0 Then DisplayError ("No Record in Operation Master"): Text1.SetFocus: Exit Sub Else rstOperationList.MoveFirst
        rstOperationList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": OperationCode = ""
        Call LoadSelectionList(rstOperationList, "List of Operation(s)...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text1, OperationCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(OperationCode) <> "" Then Sendkeys "{TAB}" Else Text1.Text = ""
    ElseIf KeyCode = vbKeyDelete Then
        Text1.Text = "": OperationCode = ""
    End If
End Sub
Private Sub Text1_Validate(Cancel As Boolean)
    If CheckEmpty(Text1.Text, False) Then Cancel = True
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text3.Text)
        If rstSizeList.RecordCount = 0 Then DisplayError ("No Record in Size Master"): Text3.SetFocus: Exit Sub Else rstSizeList.MoveFirst
        rstSizeList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": SizeCode = ""
        Call LoadSelectionList(rstSizeList, "List of Size(s)...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, SizeCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(SizeCode) <> "" Then Sendkeys "{TAB}" Else Text3.Text = ""
    ElseIf KeyCode = vbKeyDelete Then
        Text3.Text = "": SizeCode = ""
    End If
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text4.Text)
        rstCalcModeList.MoveFirst
        rstCalcModeList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": CalcModeCode = ""
        Call LoadSelectionList(rstCalcModeList, "List of Calc Mode(s)...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text4, CalcModeCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(CalcModeCode) <> "" Then Sendkeys "{TAB}" Else Text4.Text = ""
    ElseIf KeyCode = vbKeyDelete Then
        Text4.Text = "": CalcModeCode = ""
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If CheckEmpty(Text4.Text, False) Then Cancel = True
End Sub
Private Sub cmdProceed_Click()
    Dim Control As Object
    If CheckMandatoryFields Then Exit Sub
    SaveFields
    rstAccountChild.Update
    Call CloseForm(Me)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text1.Text, False) Then
        Text1.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text4.Text, False) Then
        Text4.SetFocus: CheckMandatoryFields = True
    End If
End Function
