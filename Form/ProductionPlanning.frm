VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmProductionPlanning 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production Planning"
   ClientHeight    =   6435
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7620
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
   ScaleHeight     =   6435
   ScaleWidth      =   7620
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
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
            ImageIndex      =   3
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ProductionPlanning.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ProductionPlanning.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ProductionPlanning.frx":0658
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6065
      Left            =   45
      TabIndex        =   7
      Top             =   345
      Width           =   7530
      _Version        =   65536
      _ExtentX        =   13282
      _ExtentY        =   10698
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
      Picture         =   "ProductionPlanning.frx":076C
      Begin MSComctlLib.ListView ListView4 
         Height          =   2875
         Left            =   0
         TabIndex        =   2
         Top             =   320
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   5080
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
         TabIndex        =   8
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
         Picture         =   "ProductionPlanning.frx":0788
         Picture         =   "ProductionPlanning.frx":07A4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   9
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
         Picture         =   "ProductionPlanning.frx":07C0
         Picture         =   "ProductionPlanning.frx":07DC
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3190
         Left            =   3750
         TabIndex        =   3
         Top             =   0
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   5636
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   2670
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "ProductionPlanning.frx":07F8
         Caption         =   "ProductionPlanning.frx":0910
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ProductionPlanning.frx":097C
         Keys            =   "ProductionPlanning.frx":099A
         Spin            =   "ProductionPlanning.frx":09F8
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
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "ProductionPlanning.frx":0A20
         Caption         =   "ProductionPlanning.frx":0B38
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ProductionPlanning.frx":0BA4
         Keys            =   "ProductionPlanning.frx":0BC2
         Spin            =   "ProductionPlanning.frx":0C20
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   2880
         Left            =   0
         TabIndex        =   4
         Top             =   3180
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   5080
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
         Height          =   2880
         Left            =   3750
         TabIndex        =   5
         Top             =   3180
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   5080
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
   End
End
Attribute VB_Name = "FrmProductionPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OrderType As String
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstProductionPlanning As New ADODB.Recordset
Dim rstBookList As New ADODB.Recordset
Dim rstBoardList As New ADODB.Recordset
Dim rstClassList As New ADODB.Recordset
Dim rstGroupList As New ADODB.Recordset
Dim cnProductionPlanning As New ADODB.Connection
Dim K As Integer
Dim OutputTo As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Me.Caption = "Production Planning (" + IIf(OrderType = "M", "Main", "Supplement") + ") Orders"
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "SELECT PrintName,MCRepair FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    '
    rstGroupList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type = '5' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstGroupList.ActiveConnection = Nothing
    Call FillList(ListView4, "List of Groups...", rstGroupList)
    '
    rstClassList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type='4' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstClassList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Classes...", rstClassList)
    '
    rstBoardList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type='2' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.ActiveConnection = Nothing
    Call FillList(ListView2, "List of Boards...", rstBoardList)
    '
    If OrderType = "M" Then MhDateInput1.Text = "01-10-" + Trim(Year(FinancialYearFrom) - 2) Else MhDateInput1.Text = Format(FinancialYearFrom, "dd-MM-yyyy")
    If Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd") Then MhDateInput2.Text = Format(FinancialYearTo, "dd-mm-yyyy") Else MhDateInput2.Text = Format(Date, "dd-mm-yyyy")
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{TAB}", True
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
    If UnloadMode = 0 Then CloseForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstBoardList)
    Call CloseRecordset(rstClassList)
    Call CloseRecordset(rstGroupList)
    Call CloseRecordset(rstProductionPlanning)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf OrderType = "M" And (Month(GetDate(MhDateInput1.Text)) <> 10 And Month(GetDate(MhDateInput1.Text)) <> 4) Or Day(GetDate(MhDateInput1.Text)) <> 1 Then
        Cancel = True
    ElseIf OrderType = "S" And Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput2.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
        FocusSelect Me.ActiveControl
        Cancel = True
    ElseIf OrderType = "M" And Year(GetDate(MhDateInput2.Text)) - Year(GetDate(MhDateInput1.Text)) < 2 Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next
    End If
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
     Call BookSelection
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next
        Call BookSelection
    End If
End Sub
Private Sub ListView3_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next
    End If
End Sub
Private Sub ListView4_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView4.ListItems.Count
            ListView4.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        OutputTo = "S"
        PrintProductionPlanning
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintProductionPlanning
    ElseIf Button.Index = 3 Then
        CloseForm Me
    End If
End Sub
Private Sub BookSelection()
    If rstBookList.State = adStateOpen Then rstBookList.Close
    rstBookList.Open "SELECT Name,LEFT(BusyCode,4) As Code FROM BookMaster WHERE [Group] IN (" & SelectedItems(ListView4) & ")  AND [Class] IN (" & SelectedItems(ListView1) & ") AND Board IN (" & SelectedItems(ListView2) & ") AND BusyCode<>'' AND LEN(BusyCode)>=6 AND Type='F' AND [Class]<>'' AND [Group]<>'' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstBookList.ActiveConnection = Nothing
    ListView3.ListItems.Clear
    Call FillList(ListView3, "List of Books...", rstBookList)
End Sub
Private Sub PrintProductionPlanning()
    Dim MatCentre
    MatCentre = GetChildGroup()
    If CheckEmpty(MatCentre, False) Then DisplayError ("Failed to connect to Database"): Exit Sub
    Dim DatabaseName As String
    Dim FromDate As String, ToDate As String, FromDate02 As String, ToDate02 As String
    Dim oExcel As Object
    Dim i As Long, Cnt As Long
    Dim Period01 As String, Period02 As String, Period03 As String
    On Error GoTo ErrorHandler
    DoEvents
    ShowProgressInStatusBar True
    MdiMainMenu.ProgressBar1.Value = 1
    DatabaseName = Trim(ReadFromFile("Busy Database Name"))
    If ServerName = "" Or DatabaseName = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    cnDatabase.Execute "UPDATE BookMaster SET SaleLY1003=0,SaleTY0409=0,StockTransferLY1003=0,StockTransferTY0409=0,SpecimenLY1003=0,SpecimenTY0409=0,PendingSO=0,SaleableStock=0,RepairableStock=0,POLTLY1003=0,POLY0409=0,POLY1003=0,POTY0409=0,PendingPO=0,ESO30=0,ESO60=0,ESO90=0,ESO150=0,PSO15=0,PSO30=0"
    i = 0: K = 0
    cnProductionPlanning.CursorLocation = adUseClient: cnProductionPlanning.CommandTimeout = 0
    Do While True
        i = InStr(1, DatabaseName, ",")
        If cnProductionPlanning.State = adStateOpen Then cnProductionPlanning.Close
        If i = 0 Then cnProductionPlanning.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=" & Mid(DatabaseName, 1) & ";Data Source=" & ServerName Else cnProductionPlanning.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=" & Mid(DatabaseName, 1, i - 1) & ";Data Source=" & ServerName
        K = K + 1
        If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
        If K = 1 Then   'Last Year Data Processing
            If OrderType = "M" Then
                If Month(GetDate(MhDateInput1.Text)) = 10 Then  'Oct-Mar
                    FromDate = "01-Oct-" + Trim(Year(GetDate(MhDateInput1.Text)) + 1): ToDate = "31-Mar-" + Trim(Year(GetDate(MhDateInput1.Text)) + 2)
                    rstProductionPlanning.Open "SELECT Alias," & _
                                                                  "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code) As Sale01," & _
                                                                  "0 As Sale02," & _
                                                                  "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen01," & _
                                                                  "0 As Specimen02 " & _
                                                                  "FROM Master1 M WHERE MasterType=6 AND Alias<>'' AND LEFT(Alias,4) IN (" & SelectedItems(ListView3, True) & ") ORDER BY Alias", cnProductionPlanning, adOpenKeyset, adLockReadOnly
                Else    'Apr-Sep
                    FromDate = "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text)) + 1): ToDate = "30-Sep-" + Trim(Year(GetDate(MhDateInput1.Text)) + 1)
                    FromDate02 = "01-Oct-" + Trim(Year(GetDate(MhDateInput1.Text)) + 1): ToDate02 = "31-Mar-" + Trim(Year(GetDate(MhDateInput1.Text)) + 2)
                    rstProductionPlanning.Open "SELECT Alias," & _
                                                                  "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code) As Sale01," & _
                                                                  "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate02 & "' AND Date<='" & ToDate02 & "' AND MasterCode1=M.Code) As Sale02," & _
                                                                  "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen01," & _
                                                                  "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate02 & "' AND Date<='" & ToDate02 & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen02 " & _
                                                                  "FROM Master1 M WHERE MasterType=6 AND Alias<>'' AND LEFT(Alias,4) IN (" & SelectedItems(ListView3, True) & ") ORDER BY Alias", cnProductionPlanning, adOpenKeyset, adLockReadOnly
                End If
            Else
                FromDate = Trim(Day(GetDate(MhDateInput2.Text))) + "-" + MonthName(Trim(Month(GetDate(MhDateInput2.Text))), True) + "-" + Trim(Year(GetDate(MhDateInput2.Text)) - 1)
                If Not IsDate(FromDate) Then FromDate = Trim(Day(GetDate(MhDateInput2.Text)) - 1) + "-" + MonthName(Trim(Month(GetDate(MhDateInput2.Text))), True) + "-" + Trim(Year(GetDate(MhDateInput2.Text)) - 1)
                rstProductionPlanning.Open "SELECT Alias," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale30," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale60," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale90," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale150," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen30," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen60," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen90," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen150," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text)) - 1) & "' AND Date<='" & "31-Mar-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND MasterCode1=M.Code) As LYSale," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text)) - 1) & "' AND Date<='" & "31-Mar-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As LYSpecimen," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As CSale15," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As CSale30," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As CSpecimen15," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As CSpecimen30 " & _
                                                              "FROM Master1 M WHERE MasterType=6 AND Alias<>'' AND LEFT(Alias,4) IN (" & SelectedItems(ListView3, True) & ") ORDER BY Alias", cnProductionPlanning, adOpenKeyset, adLockReadOnly
            End If
            rstProductionPlanning.ActiveConnection = Nothing
            Call UpdatePPFigures("1") 'Update Sale, Stock Transfer And Specimen Figures
            MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
        Else    'Current Year Data Processing
            If OrderType = "M" Then
                FromDate = "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text)) + 2): ToDate = "30-Sep-" + Trim(Year(GetDate(MhDateInput1.Text)) + 2)
                rstProductionPlanning.Open "SELECT Name,Alias," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code) As Sale," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=3 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As SaleReturn," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As SaleOrder," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE RecType=4 AND Method=2 AND RefCode IN (SELECT RefCode FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & ")))) As SaleOrderSupplied," & _
                                                              "(SELECT ISNULL(SUM(D1),0) FROM Tran4 WHERE MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As OpBal," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As NetSale," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=5 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As NetStockTransfer," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType IN (4,11) AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As NetPurchase," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=8 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As NetStockAdjustment," & _
                                                              "ISNULL((SELECT SUM(D1) FROM Tran4 WHERE MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE Name LIKE '%" & rstCompanyMaster.Fields("MCRepair").Value & "%')),0)+ISNULL((SELECT SUM(Value1) FROM Tran2 WHERE VchType IN (2,3,4,5,8,9,10,11) AND RecType=2 AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE Name LIKE '%" & rstCompanyMaster.Fields("MCRepair").Value & "%')),0) As RepairableStock " & _
                                                              "FROM Master1 M WHERE MasterType=6 AND Alias<>'' AND LEFT(Alias,4) IN (" & SelectedItems(ListView3, True) & ") ORDER BY Alias", cnProductionPlanning, adOpenKeyset, adLockReadOnly
            Else
                FromDate = Trim(Day(GetDate(MhDateInput2.Text))) + "-" + MonthName(Trim(Month(GetDate(MhDateInput2.Text))), True) + "-" + Trim(Year(GetDate(MhDateInput2.Text)) - 1)
                If Not IsDate(FromDate) Then FromDate = Trim(Day(GetDate(MhDateInput2.Text)) - 1) + "-" + MonthName(Trim(Month(GetDate(MhDateInput2.Text))), True) + "-" + Trim(Year(GetDate(MhDateInput2.Text)) - 1)
                rstProductionPlanning.Open "SELECT Name,Alias," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale30,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale60,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale90,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale150," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen30,(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen60," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen90,(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen150," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As CYSale," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As CYSpecimen," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As CSale15," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As CSale30," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As CSpecimen15," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE RecType=2 AND Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%SPECIMEN%')) As CSpecimen30," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As SaleOrder," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE RecType=4 AND Method=2 AND RefCode IN (SELECT RefCode FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & ")))) As SaleOrderSupplied," & _
                                                              "(SELECT ISNULL(SUM(D1),0) FROM Tran4 WHERE MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As OpBal," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As NetSale," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=5 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As NetStockTransfer," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType IN (4,11) AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As NetPurchase," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=8 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (" & MatCentre & "))) As NetStockAdjustment," & _
                                                              "ISNULL((SELECT SUM(D1) FROM Tran4 WHERE MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE Name LIKE '%" & rstCompanyMaster.Fields("MCRepair").Value & "%')),0)+ISNULL((SELECT SUM(Value1) FROM Tran2 WHERE VchType IN (2,3,4,5,8,9,10,11) AND RecType=2 AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE Name LIKE '%" & rstCompanyMaster.Fields("MCRepair").Value & "%')),0) As RepairableStock " & _
                                                              "FROM Master1 M WHERE MasterType=6 AND Alias<>'' AND LEFT(Alias,4) IN (" & SelectedItems(ListView3, True) & ") ORDER BY Alias", cnProductionPlanning, adOpenKeyset, adLockReadOnly
            End If
            rstProductionPlanning.ActiveConnection = Nothing
            Call UpdatePPFigures("2") 'Update Sale,Stock Transfer,Specimen, Stock & Pending Sales Order Figures
            MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
        End If
        If i = 0 Then Exit Do Else DatabaseName = Mid(DatabaseName, i + 1): i = 0
    Loop
    DatabaseName = Trim(ReadFromFile("Database Name")): If DatabaseName = "" Then Exit Sub
    i = 0: K = 0
    Do While True
        i = InStr(1, DatabaseName, ",")
        If cnProductionPlanning.State = adStateOpen Then cnProductionPlanning.Close
        If i = 0 Then cnProductionPlanning.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\" & Mid(DatabaseName, 1) & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#" Else cnProductionPlanning.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\" & Mid(DatabaseName, 1, i - 1) & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
        K = K + 1
        If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
        If K = 1 Then
            If OrderType = "M" Then
                FromDate = "01-Oct-" & Trim(Year(GetDate(MhDateInput1.Text))): ToDate = "31-Mar-" & Trim(Year(GetDate(MhDateInput1.Text)) + 1)
                rstProductionPlanning.Open "SELECT M.BusyCode,CLng(Sum(C.ActualQuantity)) As PrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & ToDate & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.BusyCode ORDER BY M.BusyCode", cnProductionPlanning, adOpenKeyset, adLockReadOnly
                rstProductionPlanning.ActiveConnection = Nothing
                Call UpdatePPFigures("3") 'Update Print Order Figures
            End If
            MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
        ElseIf K = 2 Then
            If OrderType = "M" Then
                FromDate = "01-Apr-" & Trim(Year(GetDate(MhDateInput1.Text)) + 1): ToDate = "30-Sep-" & Trim(Year(GetDate(MhDateInput1.Text)) + 1)
                rstProductionPlanning.Open "SELECT M.BusyCode,CLng(Sum(C.ActualQuantity)) As PrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & ToDate & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.BusyCode ORDER BY M.BusyCode", cnProductionPlanning, adOpenKeyset, adLockReadOnly
                rstProductionPlanning.ActiveConnection = Nothing
                Call UpdatePPFigures("4") 'Update Print Order Figures
                MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
                If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
                FromDate = "01-Oct-" & Trim(Year(GetDate(MhDateInput1.Text)) + 1): ToDate = "31-Mar-" & Trim(Year(GetDate(MhDateInput1.Text)) + 2)
                rstProductionPlanning.Open "SELECT M.BusyCode,CLng(Sum(C.ActualQuantity)) As PrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & ToDate & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.BusyCode ORDER BY M.BusyCode", cnProductionPlanning, adOpenKeyset, adLockReadOnly
                rstProductionPlanning.ActiveConnection = Nothing
                Call UpdatePPFigures("5") 'Update Print Order Figures
            Else
                FromDate = "01-Apr-" & Trim(Year(GetDate(MhDateInput1.Text)) - 1)
                rstProductionPlanning.Open "SELECT M.BusyCode,CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.QuantityReceived))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.QuantityReceived))))+(SELECT CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.QuantityReceived))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.QuantityReceived)))) FROM BookPOChild08 C INNER JOIN BookPOParent P ON P.Code=C.Code WHERE P.Type='R' AND P.Book=M.Code AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "#) As PendingPrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book " & _
                                                              "WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.BusyCode,M.Code ORDER BY M.BusyCode", cnProductionPlanning, adOpenKeyset, adLockReadOnly
                rstProductionPlanning.ActiveConnection = Nothing
                Call UpdatePPFigures("4") 'Update Pending Print Order Figures
            End If
            MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
        Else
            FromDate = "01-Apr-" & Trim(Year(GetDate(MhDateInput1.Text)) + IIf(OrderType = "M", 2, 0))
            rstProductionPlanning.Open "SELECT M.BusyCode,CLng(Sum(C.ActualQuantity)) As PrintOrder,CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.QuantityReceived))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.QuantityReceived))))+(SELECT CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.QuantityReceived))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.QuantityReceived)))) FROM BookPOChild08 C INNER JOIN BookPOParent P ON P.Code=C.Code WHERE P.Type='R' AND P.Book=M.Code AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "#) As PendingPrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book " & _
                                                          "WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.Code,M.BusyCode ORDER BY M.BusyCode", cnProductionPlanning, adOpenKeyset, adLockReadOnly
            rstProductionPlanning.ActiveConnection = Nothing
            Call UpdatePPFigures("6") 'Update Print Order & Pending Print Order Figures
            MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
        End If
        If i = 0 Then Exit Do Else DatabaseName = Mid(DatabaseName, i + 1): i = 0
    Loop
    Call CloseRecordset(rstProductionPlanning)
    Call CloseConnection(cnProductionPlanning)
    Screen.MousePointer = vbNormal
    On Error Resume Next
    If Not FileExist(App.Path & "\Template\Production Planning.xlsx") Then Exit Sub
    Screen.MousePointer = vbHourglass
    If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
    If OrderType = "M" Then
        rstProductionPlanning.Open "SELECT Code,PrintName,BusyCode As Alias,POLTLY1003,POLY0409,POLY1003,POTY0409,SaleLY1003,SaleTY0409,StockTransferLY1003,StockTransferTY0409,SpecimenLY1003,SpecimenTY0409,ESO30 As CYReturn,PendingPO,SaleableStock,RepairableStock,PendingSO,Remarks FROM BookMaster WHERE LEFT(BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") AND Type='F' AND BusyCode<>'' AND RIGHT(BusyCode,1)<>'S' ORDER BY PrintName", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstProductionPlanning.Open "SELECT Code,PrintName,BusyCode As Alias,POTY0409,PendingPO,SaleableStock,RepairableStock,PendingSO,ESO30,ESO60,ESO90,ESO150,PSO15,PSO30,SaleLY1003 As LYSale,SaleTY0409 As CYSale FROM BookMaster WHERE Type='F' AND LEFT(BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") AND BusyCode<>'' AND RIGHT(BusyCode,1)<>'S' ORDER BY PrintName", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstProductionPlanning.RecordCount = 0 Then
        DisplayError ("No Record Found")
        ShowProgressInStatusBar False
        Screen.MousePointer = vbNormal
        On Error GoTo 0
        Exit Sub
    End If
    DoEvents
    'Writing To Excel
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Production Planning"): oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Production Planning (" & CompCode & ")"): oExcel.DisplayAlerts = True
    oExcel.Sheets("Reorder Level Register").Visible = False: oExcel.Sheets("Production Planning (" & IIf(OrderType = "M", "SO", "MO") & ")").Visible = False: oExcel.Sheets("Production Planning (" & IIf(OrderType = "M", "MO", "SO") & ")").Select: oExcel.Visible = False
    oExcel.Cells(1, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Cells(2, "A").Value = "Production Planning (" & IIf(OrderType = "M", "Main", "Supplement") & " Orders) As On [" & Format(GetDate(MhDateInput2.Text), "dd-MMM-yyyy") & "]"
    If OrderType = "M" Then
        Period01 = "(" + Right(Year(GetDate(MhDateInput1.Text)), 2) + "-" + Right(Year(GetDate(MhDateInput1.Text)) + 1, 2) + ")"
        Period02 = "(" + Right(Year(GetDate(MhDateInput1.Text)) + 1, 2) + "-" + Right(Year(GetDate(MhDateInput1.Text)) + 2, 2) + ")"
        Period03 = "(" + Right(Year(GetDate(MhDateInput1.Text)) + 2, 2) + "-" + Right(Year(GetDate(MhDateInput1.Text)) + 3, 2) + ")"
        oExcel.Cells(5, "D").Value = Period01
        oExcel.Cells(5, "E").Value = Period02
        oExcel.Cells(5, "G").Value = Period03
        If Month(GetDate(MhDateInput1.Text)) = 10 Then
            oExcel.Cells(4, "H").Value = "Oct-Mar"
            oExcel.Cells(4, "I").Value = "Apr-Sep"
            oExcel.Cells(5, "H").Value = Period02
            oExcel.Cells(5, "I").Value = Period03
        Else
            oExcel.Cells(4, "H").Value = "Apr-Sep"
            oExcel.Cells(4, "I").Value = "Oct-Mar"
            oExcel.Cells(5, "H").Value = Period02
            oExcel.Cells(5, "I").Value = Period02
        End If
    End If
    i = IIf(OrderType = "M", 7, 5): Cnt = 1
    Do While Not rstProductionPlanning.EOF
        oExcel.Cells(i, "A").Value = Cnt
        oExcel.Application.Cells(i, "B").Value = Trim(rstProductionPlanning.Fields("PrintName").Value)
        oExcel.Application.Cells(i, "C").Value = Trim(rstProductionPlanning.Fields("Alias").Value)
        If OrderType = "M" Then
            'Print Order
            oExcel.Application.Cells(i, "D").Value = Val(rstProductionPlanning.Fields("POLTLY1003").Value)
            oExcel.Application.Cells(i, "E").Value = Val(rstProductionPlanning.Fields("POLY0409").Value)
            oExcel.Application.Cells(i, "F").Value = Val(rstProductionPlanning.Fields("POLY1003").Value)
            oExcel.Application.Cells(i, "G").Value = Val(rstProductionPlanning.Fields("POTY0409").Value)
            'Sale
            oExcel.Application.Cells(i, "H").Value = Val(rstProductionPlanning.Fields("SaleLY1003").Value)
            oExcel.Application.Cells(i, "I").Value = Val(rstProductionPlanning.Fields("SaleTY0409").Value)
            'Specimen
            oExcel.Application.Cells(i, "L").Value = Val(rstProductionPlanning.Fields("SpecimenLY1003").Value)
            oExcel.Application.Cells(i, "M").Value = Val(rstProductionPlanning.Fields("SpecimenTY0409").Value)
            'Current Return
            oExcel.Application.Cells(i, "N").Value = Val(rstProductionPlanning.Fields("CYReturn").Value)
            'Pending Print Order
            oExcel.Application.Cells(i, "O").Value = Val(rstProductionPlanning.Fields("PendingPO").Value)
            oExcel.Application.Cells(i, "P").Value = Val(rstProductionPlanning.Fields("SaleableStock").Value)
            oExcel.Application.Cells(i, "Q").Value = Val(rstProductionPlanning.Fields("RepairableStock").Value)
            oExcel.Application.Cells(i, "R").Value = Val(rstProductionPlanning.Fields("PendingSO").Value)
            If i > 7 Then oExcel.Range("S" & Trim(i)).FormulaR1C1 = oExcel.Range("S7").FormulaR1C1
            If Val(oExcel.Application.Cells(i, "S")) < 0 Then oExcel.Application.Cells(i, "S").Value = 0
            oExcel.Application.Cells(i, "T").Value = rstProductionPlanning.Fields("Remarks").Value
            oExcel.Application.Cells(i, "XFD").Value = rstProductionPlanning.Fields("Code").Value
        Else
            oExcel.Application.Cells(i, "D").Value = Val(rstProductionPlanning.Fields("POTY0409").Value)
            oExcel.Application.Cells(i, "E").Value = Val(rstProductionPlanning.Fields("SaleableStock").Value)
            oExcel.Application.Cells(i, "F").Value = Val(rstProductionPlanning.Fields("RepairableStock").Value)
            oExcel.Application.Cells(i, "G").Value = Val(rstProductionPlanning.Fields("PendingPO").Value)
            oExcel.Application.Cells(i, "H").Value = Val(rstProductionPlanning.Fields("PendingSO").Value)
            oExcel.Application.Cells(i, "I").Value = Val(oExcel.Application.Cells(i, "E")) + Val(oExcel.Application.Cells(i, "G")) - Val(oExcel.Application.Cells(i, "H"))
            oExcel.Application.Cells(i, "J").Value = Val(rstProductionPlanning.Fields("LYSale").Value)
            oExcel.Application.Cells(i, "K").Value = Val(rstProductionPlanning.Fields("CYSale").Value)
            oExcel.Application.Cells(i, "L").Value = Val(rstProductionPlanning.Fields("ESO30").Value)
            oExcel.Application.Cells(i, "M").Value = Val(rstProductionPlanning.Fields("ESO60").Value)
            oExcel.Application.Cells(i, "N").Value = Val(rstProductionPlanning.Fields("ESO90").Value)
            oExcel.Application.Cells(i, "O").Value = Val(rstProductionPlanning.Fields("ESO150").Value)
            oExcel.Application.Cells(i, "P").Value = Val(rstProductionPlanning.Fields("PSO15").Value)
            oExcel.Application.Cells(i, "Q").Value = Val(rstProductionPlanning.Fields("PSO30").Value)
        End If
        Cnt = Cnt + 1: i = i + 1
        rstProductionPlanning.MoveNext
    Loop
    oExcel.Columns("A:B").EntireColumn.AutoFit
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    MdiMainMenu.ProgressBar1.Value = 100
    If OutputTo = "S" Then oExcel.Range("A1").Activate: oExcel.Visible = True Else oExcel.Workbooks.Item(1).PrintOut
    ShowProgressInStatusBar False
    Set oExcel = Nothing
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to update Production Planning figures")
    ShowProgressInStatusBar False
    Call CloseRecordset(rstProductionPlanning)
    Call CloseConnection(cnProductionPlanning)
End Sub
Private Sub UpdatePPFigures(ByVal UpdationType As String)
    If rstProductionPlanning.RecordCount > 0 Then rstProductionPlanning.MoveFirst
    Do While Not rstProductionPlanning.EOF
        If UpdationType = "1" Then
            If OrderType = "M" Then
                cnDatabase.Execute "UPDATE BookMaster SET SaleLY1003=SaleLY1003+" & Val(rstProductionPlanning.Fields("Sale01").Value) & ",SpecimenLY1003=SpecimenLY1003+" & Val(rstProductionPlanning.Fields("Specimen01").Value) & ",SaleTY0409=SaleTY0409+" & Val(rstProductionPlanning.Fields("Sale02").Value) & ",SpecimenTY0409=SpecimenTY0409+" & Val(rstProductionPlanning.Fields("Specimen02").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"
            Else
               cnDatabase.Execute "UPDATE BookMaster SET ESO30=ESO30+" & Val(rstProductionPlanning.Fields("Sale30").Value) + Val(rstProductionPlanning.Fields("Specimen30").Value) & ",ESO60=ESO60+" & Val(rstProductionPlanning.Fields("Sale60").Value) + Val(rstProductionPlanning.Fields("Specimen60").Value) & ",ESO90=ESO90+" & Val(rstProductionPlanning.Fields("Sale90").Value) + Val(rstProductionPlanning.Fields("Specimen90").Value) & ",ESO150=ESO150+" & Val(rstProductionPlanning.Fields("Sale150").Value) + Val(rstProductionPlanning.Fields("Specimen150").Value) & "," & _
                                                    "PSO15=PSO15+" & Val(rstProductionPlanning.Fields("CSale15").Value) + Val(rstProductionPlanning.Fields("CSpecimen15").Value) & ",PSO30=PSO30+" & Val(rstProductionPlanning.Fields("CSale30").Value) + Val(rstProductionPlanning.Fields("CSpecimen30").Value) & ",SaleLY1003=SaleLY1003+" & Val(rstProductionPlanning.Fields("LYSale").Value) + Val(rstProductionPlanning.Fields("LYSpecimen").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"
            End If
        ElseIf UpdationType = "2" Then
            If OrderType = "M" Then
                If Month(GetDate(MhDateInput1.Text)) = 10 Then cnDatabase.Execute "UPDATE BookMaster SET SaleTY0409=SaleTY0409+" & Val(rstProductionPlanning.Fields("Sale").Value) & ",SpecimenTY0409=SpecimenTY0409+" & Val(rstProductionPlanning.Fields("Specimen").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"
                cnDatabase.Execute "UPDATE BookMaster SET ESO30=ESO30+" & Val(rstProductionPlanning.Fields("SaleReturn").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"  'Current Return
            Else
                cnDatabase.Execute "UPDATE BookMaster SET ESO30=ESO30+" & Val(rstProductionPlanning.Fields("Sale30").Value) + Val(rstProductionPlanning.Fields("Specimen30").Value) & ",ESO60=ESO60+" & Val(rstProductionPlanning.Fields("Sale60").Value) + Val(rstProductionPlanning.Fields("Specimen60").Value) & ",ESO90=ESO90+" & Val(rstProductionPlanning.Fields("Sale90").Value) + Val(rstProductionPlanning.Fields("Specimen90").Value) & ",ESO150=ESO150+" & Val(rstProductionPlanning.Fields("Sale150").Value) + Val(rstProductionPlanning.Fields("Specimen150").Value) & "," & _
                                                    "PSO15=PSO15+" & Val(rstProductionPlanning.Fields("CSale15").Value) + Val(rstProductionPlanning.Fields("CSpecimen15").Value) & ",PSO30=PSO30+" & Val(rstProductionPlanning.Fields("CSale30").Value) + Val(rstProductionPlanning.Fields("CSpecimen30").Value) & ",SaleTY0409=SaleTY0409+" & Val(rstProductionPlanning.Fields("CYSale").Value) + Val(rstProductionPlanning.Fields("CYSpecimen").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"
            End If
            If StrConv(Mid(rstProductionPlanning.Fields("Alias").Value, 6, 1), vbUpperCase) <> "Z" Then cnDatabase.Execute "UPDATE BookMaster SET PendingSO=PendingSO+" & Val(rstProductionPlanning.Fields("SaleOrder").Value) - Val(rstProductionPlanning.Fields("SaleOrderSupplied").Value) & ",SaleableStock=SaleableStock+" & Val(rstProductionPlanning.Fields("OpBal").Value) - Val(rstProductionPlanning.Fields("NetSale").Value) + Val(rstProductionPlanning.Fields("NetStockTransfer").Value) + Val(rstProductionPlanning.Fields("NetPurchase").Value) + Val(rstProductionPlanning.Fields("NetStockAdjustment").Value) & ",RepairableStock=RepairableStock+" & Val(rstProductionPlanning.Fields("RepairableStock").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"
        ElseIf UpdationType = "3" Then
            If OrderType = "M" Then cnDatabase.Execute "UPDATE BookMaster SET POLTLY1003=POLTLY1003+" & Val(CheckNull(rstProductionPlanning.Fields("PrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'"
        ElseIf UpdationType = "4" Then
            If OrderType = "M" Then cnDatabase.Execute "UPDATE BookMaster SET POLY0409=POLY0409+" & Val(CheckNull(rstProductionPlanning.Fields("PrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'" Else cnDatabase.Execute "UPDATE BookMaster SET PendingPO=PendingPO+" & Val(CheckNull(rstProductionPlanning.Fields("PendingPrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'"
        ElseIf UpdationType = "5" Then
            If OrderType = "M" Then cnDatabase.Execute "UPDATE BookMaster SET POLY1003=POLY1003+" & Val(CheckNull(rstProductionPlanning.Fields("PrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'"
        ElseIf UpdationType = "6" Then
            cnDatabase.Execute "UPDATE BookMaster SET POTY0409=POTY0409+" & Val(CheckNull(rstProductionPlanning.Fields("PrintOrder").Value)) & ",PendingPO=PendingPO+" & Val(CheckNull(rstProductionPlanning.Fields("PendingPrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'"
        End If
        rstProductionPlanning.MoveNext
    Loop
End Sub
