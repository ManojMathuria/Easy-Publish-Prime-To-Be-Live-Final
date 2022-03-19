VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPrintOrderStatusRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Order Status Register"
   ClientHeight    =   6435
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8385
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
   ScaleWidth      =   8385
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8385
      _ExtentX        =   14790
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
            Picture         =   "PrinterOrderStatusRegister.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrinterOrderStatusRegister.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrinterOrderStatusRegister.frx":0658
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6065
      Left            =   45
      TabIndex        =   10
      Top             =   345
      Width           =   8305
      _Version        =   65536
      _ExtentX        =   14649
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
      Picture         =   "PrinterOrderStatusRegister.frx":076C
      Begin VB.CheckBox Check3 
         Caption         =   "UFG"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5700
         TabIndex        =   4
         Top             =   53
         Value           =   1  'Checked
         Width           =   840
      End
      Begin VB.CheckBox Check2 
         Caption         =   "FG"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4920
         TabIndex        =   3
         Top             =   53
         Value           =   1  'Checked
         Width           =   750
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3840
         TabIndex        =   2
         Top             =   53
         Width           =   1095
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
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
         Caption         =   " &From"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "PrinterOrderStatusRegister.frx":0788
         Picture         =   "PrinterOrderStatusRegister.frx":07A4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   12
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
         Picture         =   "PrinterOrderStatusRegister.frx":07C0
         Picture         =   "PrinterOrderStatusRegister.frx":07DC
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2880
         Left            =   0
         TabIndex        =   6
         Top             =   315
         Width           =   8310
         _ExtentX        =   14658
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
            Weight          =   700
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
         Calendar        =   "PrinterOrderStatusRegister.frx":07F8
         Caption         =   "PrinterOrderStatusRegister.frx":0910
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PrinterOrderStatusRegister.frx":097C
         Keys            =   "PrinterOrderStatusRegister.frx":099A
         Spin            =   "PrinterOrderStatusRegister.frx":09F8
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
         Calendar        =   "PrinterOrderStatusRegister.frx":0A20
         Caption         =   "PrinterOrderStatusRegister.frx":0B38
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PrinterOrderStatusRegister.frx":0BA4
         Keys            =   "PrinterOrderStatusRegister.frx":0BC2
         Spin            =   "PrinterOrderStatusRegister.frx":0C20
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
      Begin MSComctlLib.ListView ListView3 
         Height          =   2880
         Left            =   0
         TabIndex        =   7
         Top             =   3180
         Width           =   4150
         _ExtentX        =   7329
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2880
         Left            =   4140
         TabIndex        =   8
         Top             =   3180
         Width           =   4170
         _ExtentX        =   7355
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Left            =   6580
         TabIndex        =   5
         Top             =   0
         Width           =   1725
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3043;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "FrmPrintOrderStatusRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnBusy As New ADODB.Connection
Dim rstPrintOrderStatusRegister As New ADODB.Recordset, rstCompanyMaster As New ADODB.Recordset, rstBusy As New ADODB.Recordset, rstBookList As New ADODB.Recordset, rstAccountList As New ADODB.Recordset
Dim OutputTo As String
Public OrderType As String
Private Sub Form_Load()
    '01-Bookwise 02-Print Orderwise 05-Book Printerwise 06-Title Printerwise 08-Book Binderwise XX-Busy YY-Debit Note
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    cnBusy.CursorLocation = adUseClient
    If InStr(1, "0102", OrderType) = 0 Then ListView3.Width = 8310: ListView4.Visible = False
    If OrderType = "XX" Then
        Me.Height = 3940: Mh3dFrame1.Height = 3200: ListView3.Visible = False
    Else
        rstAccountList.Open "SELECT Name,Code FROM AccountMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
        rstAccountList.ActiveConnection = Nothing
        Call FillList(ListView3, "List of " & IIf(InStr(1, "01020506", OrderType) > 0, "Printers", "Binders") & "...", rstAccountList)
        If InStr(1, "0102", OrderType) > 0 Then
            If rstAccountList.State = adStateOpen Then rstAccountList.Close
            rstAccountList.Open "SELECT Name,Code FROM AccountMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
            rstAccountList.ActiveConnection = Nothing
            Call FillList(ListView4, "List of Book Binders...", rstAccountList)
        End If
    End If
    If InStr(1, "XXYY", OrderType) > 0 Then
        Me.Caption = "Pending " & IIf(OrderType = "XX", "Print Order Register [Busy]", "Debit Notes Register"): Check1.Enabled = False: Combo1.Enabled = False
    Else
        Me.Caption = "Print Order Status Register [" & Choose(Val(OrderType), "Bookwise", "Print Orderwise", "", "", "Book Printerwise", "Title Printerwise", "", "Book Binderwise") & "]"
    End If
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    Call BookSelection(True)
    MhDateInput1.Text = Format(DateAdd("D", -365, FinancialYearFrom), "dd-mm-yyyy"): MhDateInput2.Text = Format(Date, "dd-mm-yyyy")
    Combo1.AddItem "Without Stock", 0: Combo1.AddItem "With Stock", 1: Combo1.AddItem "With Pending SO", 2: Combo1.ListIndex = 0
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
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstBusy)
    Call CloseRecordset(rstPrintOrderStatusRegister)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput2.Text)) Or Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then Cancel = True
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
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
    If Button.Index = 3 Then CloseForm Me: Exit Sub
    If Button.Index = 1 Then OutputTo = "S" Else OutputTo = "P"
    PrintOrderStatusRegister
End Sub
Private Sub BookSelection(ByVal SelectAll As Boolean)
    If rstBookList.State = adStateOpen Then rstBookList.Close
    rstBookList.Open "SELECT Name,Code FROM BookMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstBookList.ActiveConnection = Nothing
    ListView2.ListItems.Clear
    Call FillList(ListView2, "List of Books...", rstBookList)
End Sub
Private Function GetBusyOrder(ByVal xOrderNo As String)
    If rstBusy.State = adStateOpen Then rstBusy.Close
    rstBusy.Open "SELECT T1.VchNo,M.Alias,M.PrintName As BinderName,T1.Date,ABS(T3.Value1) As OrderedQuantity FROM (Tran1 T1 INNER JOIN Tran3 T3 ON T1.VchNo=T3.No) INNER JOIN Master1 M ON M.Code=T1.MasterCode1 WHERE T1.VchType=T3.VchType AND T1.VchType=13 AND T1.Date=T3.Date AND LTRIM(T3.No)='" & Trim(xOrderNo) & "'", cnBusy, adOpenKeyset, adLockReadOnly
    rstBusy.ActiveConnection = Nothing
End Function
Private Function ConnectToBusy() As Boolean
    On Error GoTo ErrHandler
    Dim DatabaseName
    DatabaseName = Trim(ReadFromFile("Busy Database Name"))
    DatabaseName = StrReverse(Left(StrReverse(DatabaseName), InStr(1, StrReverse(DatabaseName), ",") - 1))
    If cnBusy.State = adStateOpen Then cnBusy.Close
    cnBusy.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=" & DatabaseName & ";Data Source=" & ServerName
    ConnectToBusy = True
ErrHandler:
End Function
Private Sub GetAllItemStock()
    Dim SQL As String
    MdiMainMenu.StatusBar1.Panels(2).Text = "Processing !!! Please Wait....."
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    If rstBusy.State = adStateOpen Then rstBusy.Close
    SQL = "SELECT LEFT(Alias,4) As Alias,Alias As FullAlias," & _
              "(SELECT ISNULL(SUM(D1),0) FROM Tran4 WHERE MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As OpBal," & _
              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FinancialYearFrom & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetSale," & _
              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=5 AND Date>='" & FinancialYearFrom & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetStockTransfer," & _
              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType IN (4,11) AND Date>='" & FinancialYearFrom & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetPurchase," & _
              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=8 AND Date>='" & FinancialYearFrom & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetStockAdjustment "
    If Combo1.ListIndex = 1 Then
        SQL = SQL + "FROM Master1 M WHERE MasterType=6 AND Alias<>'' ORDER BY Left(Alias,4)"
    Else
        SQL = SQL + "," & _
                   "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As SaleOrder," & _
                   "ABS(ISNULL((SELECT ISNULL(SUM(Value1),0) FROM Tran3 WHERE RecType=4 AND Method=2 AND RefCode IN (SELECT RefCode FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code)),0)) As SaleOrderSupplied " & _
                   "FROM Master1 M WHERE MasterType=6 AND Alias<>'' ORDER BY Left(Alias,4)"
    End If
    rstBusy.Open SQL, cnBusy, adOpenKeyset, adLockReadOnly
    rstBusy.ActiveConnection = Nothing
ErrorHandler:
    On Error GoTo 0
    Screen.MousePointer = vbNormal
End Sub
Private Function GetStock(ByVal xItem As String) As String
    Dim EffStock As Long, PendingSO As Long
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    If rstBusy.RecordCount > 0 Then rstBusy.MoveFirst
    rstBusy.Find "[Alias]='" & Left(xItem, 4) & "'"
    Do While Not rstBusy.EOF
        If rstBusy.Fields("Alias") = Left(xItem, 4) Then
            If InStr(1, "Z-Z_", Mid(rstBusy.Fields("FullAlias").Value, 6, 2), vbTextCompare) = 0 Then
                EffStock = EffStock + Val(rstBusy.Fields("OpBal").Value) - Val(rstBusy.Fields("NetSale").Value) + Val(rstBusy.Fields("NetStockTransfer").Value) + Val(rstBusy.Fields("NetPurchase").Value) + Val(rstBusy.Fields("NetStockAdjustment").Value)
                If Combo1.ListIndex = 2 Then PendingSO = PendingSO + Val(rstBusy.Fields("SaleOrder").Value) - Val(rstBusy.Fields("SaleOrderSupplied").Value)
            End If
        Else
            Exit Do
        End If
        rstBusy.MoveNext
    Loop
    GetStock = Trim(Str(EffStock)) + "|" + Trim(Str(PendingSO))
ErrorHandler:
    On Error GoTo 0
    Screen.MousePointer = vbNormal
End Function
Private Sub PrintOrderStatusRegister()
    Dim oExcel As Object
    Dim i As Long, K As Long, Cnt As Long, T As String
    Dim SelectedBooks, SelectedPrinters, SelectedBinders, SQL, Path
    If OrderType = "XX" Or (Combo1.ListIndex > 0 And InStr(1, "0102050608", OrderType) > 0) Then If Not ConnectToBusy Then Screen.MousePointer = vbNormal: DisplayError ("Failed to connect to busy"): Exit Sub
    On Error Resume Next
    If Not FileExist(App.Path & "\Template\Print Order Status Register.xlsx") Then Exit Sub
    Screen.MousePointer = vbHourglass
    If rstPrintOrderStatusRegister.State = adStateOpen Then rstPrintOrderStatusRegister.Close
    SelectedBooks = SelectedItems(ListView2)
    If OrderType <> "XX" Then
        If Combo1.ListIndex > 0 And InStr(1, "0102050608", OrderType) > 0 Then GetAllItemStock
        SQL = "SELECT P.Code,P.Name As OrderNo,P.Date As OrderDate,M2.PrintName As BookName,M2.BusyCode As Alias,(SELECT PrintName FROM GeneralMaster WHERE Code=M2.[Size]) As BookSize,M2.FormType,(SELECT STR(Forms1) FROM BookPOChild05 WHERE Code=P.Code) As Forms1,(SELECT STR(Forms2) FROM BookPOChild05 WHERE Code=P.Code) As Forms2,(SELECT STR(Forms4) FROM BookPOChild05 WHERE Code=P.Code) As Forms4,(SELECT STR(FrontPrintingType) FROM BookPOChild06 WHERE Code=P.Code) As FrontPrintingType,(SELECT STR(BackPrintingType) FROM BookPOChild06 WHERE Code=P.Code) As BackPrintingType,FORMAT((SELECT ActualQuantity FROM BookPOChild05 WHERE Code=P.Code),'0') As TextQuantity,FORMAT((SELECT ActualQuantity FROM BookPOChild06 WHERE Code=P.Code),'0') As TitleQuantity,FORMAT((SELECT ActualQuantity FROM BookPOChild08 WHERE Code=P.Code),'0') As BookQuantity,(P.DeliveredQuantityC+P.DeliveredQuantityB) As QuantityReceived," & _
                  "(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinterName,(SELECT Status FROM BookPOChild05 WHERE Code=P.Code) As TextStatus,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinterName,(SELECT Status FROM BookPOChild06 WHERE Code=P.Code) As TitleStatus,(SELECT PrintName FROM AccountMaster WHERE Code=P.Binder) As BinderName,(SELECT Status FROM BookPOChild08 WHERE Code=P.Code) As BookStatus,(SELECT PrintName FROM BookPOChild05 T INNER JOIN PaperMaster M ON T.Paper1=M.Code WHERE T.Code=P.Code) As Paper1,(SELECT PrintName FROM BookPOChild05 T INNER JOIN PaperMaster M ON T.Paper2=M.Code WHERE T.Code=P.Code) As Paper2,(SELECT PrintName FROM BookPOChild05 T INNER JOIN PaperMaster M ON T.Paper4=M.Code WHERE T.Code=P.Code) As Paper4,(SELECT PrintName FROM BookPOChild06 T INNER JOIN PaperMaster M ON T.Paper=M.Code WHERE T.Code=P.Code) As Paper," & _
                  "'' As LaminationType,(SELECT PrintName FROM BookPOChild08 T INNER JOIN GeneralMaster M ON T.BindingType=M.Code WHERE T.Code=P.Code) As BindingType,C.Narration,C.BillNo,C.BillDate,C.BillAmount "
    End If
        If InStr(1, "0102", OrderType) > 0 Then 'Book/Print Orderwise
        SelectedPrinters = SelectedItems(ListView3): SelectedBinders = SelectedItems(ListView4)
        SQL = SQL + "FROM (((BookPOParent P LEFT JOIN BookPOChild08 C ON P.Code=C.Code) LEFT JOIN AccountMaster M1 ON P.BookPrinter=M1.Code) LEFT JOIN BookMaster M2 ON P.Book=M2.Code) " & _
                             "WHERE LEFT(P.Type,1) IN ('" & IIf(Check2.Value And Check3.Value, "F','R", IIf(Check2.Value, "F", "R")) & "') AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Check1.Value, "1=1", "C.Status NOT IN ('D','E','W')") & " AND " & IIf(SelectedBooks = "''", "1=1", "M2.Code IN (" & SelectedBooks & ")") & " AND " & IIf(SelectedPrinters = "''", "1=1", "P.BookPrinter IN (" & SelectedPrinters & ")") & " AND " & IIf(SelectedBinders = "''", "1=1", "M1.Code IN (" & SelectedBinders & ")") & Space(1) & _
                             "ORDER BY " & IIf(OrderType = "01", "M2.PrintName,CONVERT(INT,P.Name)", "CONVERT(INT,P.Name),M2.PrintName")
        rstPrintOrderStatusRegister.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
    ElseIf OrderType = "05" Then    'Text Printerwise
        SelectedPrinters = SelectedItems(ListView3)
        SQL = SQL + ",C.[TotalPlates1-�]+C.[TotalPlates1-�]+C.[TotalPlates1-1]+C.[RevisedPlates1]+((C.[TotalPlates2-�]+C.[TotalPlates2-�]+C.[TotalPlates2-1]+C.[RevisedPlates2])*2)+((C.[TotalPlates4-�]+C.[TotalPlates4-�]+C.[TotalPlates4-1]+C.[RevisedPlates4])*4) As TotalPlates FROM (((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.BookPrinter=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code) " & _
                             "WHERE LEFT(P.Type,1) IN ('" & IIf(Check2.Value And Check3.Value, "F','R", IIf(Check2.Value, "F", "R")) & "') AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Check1.Value, "1=1", "C.Status NOT IN ('D')") & " AND " & IIf(SelectedBooks = "''", "1=1", "M2.Code IN (" & SelectedBooks & ")") & " AND " & IIf(SelectedPrinters = "''", "1=1", "M1.Code IN (" & SelectedPrinters & ")") & Space(1) & _
                             "ORDER BY M1.PrintName,M2.PrintName,CONVERT(INT,P.Name)"
        rstPrintOrderStatusRegister.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
    ElseIf OrderType = "06" Then    'Title Printerwise
        SelectedPrinters = SelectedItems(ListView3)
        SQL = SQL + ",C.TotalPlates FROM (((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.TitlePrinter=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code) " & _
                             "WHERE LEFT(P.Type,1) IN ('" & IIf(Check2.Value And Check3.Value, "F','R", IIf(Check2.Value, "F", "R")) & "') AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Check1.Value, "1=1", "C.Status NOT IN ('D')") & " AND " & IIf(SelectedBooks = "''", "1=1", "M2.Code IN (" & SelectedBooks & ")") & " AND " & IIf(SelectedPrinters = "''", "1=1", "M1.Code IN (" & SelectedPrinters & ")") & Space(1) & _
                             "ORDER BY M1.PrintName,M2.PrintName,CONVERT(INT,P.Name)"
        rstPrintOrderStatusRegister.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
    ElseIf InStr(1, "08YY", OrderType) > 0 Then 'Book Binderwise/Pending Debit Note
        SelectedBinders = SelectedItems(ListView3)
        SQL = SQL + "FROM (((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Binder=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code) " & _
                             "WHERE LEFT(P.Type,1) IN ('" & IIf(Check2.Value And Check3.Value, "F','R", IIf(Check2.Value, "F", "R")) & "') AND C.OrderDate>='" & GetDate(MhDateInput1.Text) & "' AND C.OrderDate<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Check1.Value, "1=1", "C.Status NOT IN ('D','E','W')") & " AND " & IIf(SelectedBooks = "''", "1=1", "M2.Code IN (" & SelectedBooks & ")") & " AND " & IIf(SelectedBinders = "''", "1=1", "M1.Code IN (" & SelectedBinders & ")") & Space(1) & _
                             "ORDER BY M1.PrintName,M2.PrintName,CONVERT(INT,P.Name)"
        rstPrintOrderStatusRegister.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
    ElseIf OrderType = "XX" Then    'Pending Print Order (Busy)
        rstPrintOrderStatusRegister.Open "SELECT M2.PrintName+' ('+BusyCode+')' As BookName,P.Name As OrderNo,C.OrderDate As OrderDate,C.ActualQuantity As OrderedQuantity,(SELECT PrintName+' ('+Alias+')' FROM AccountMaster WHERE Code=P.BookPrinter) As BookPrinterName,M1.PrintName+' ('+Alias+')' As BinderName,(SELECT Alias FROM AccountMaster WHERE Code=P.Laminator) As LaminatorAlias " & _
                                                               "FROM (((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Binder=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code) " & _
                                                               "WHERE LEFT(P.Type,1) IN ('" & IIf(Check2.Value And Check3.Value, "F','R", IIf(Check2.Value, "F", "R")) & "') AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND P.QuantityReceived=0 AND " & IIf(SelectedBooks = "''", "1=1", "M2.Code IN (" & SelectedBooks & ")") & Space(1) & _
                                                               "ORDER BY CONVERT(INT,P.Name),M2.PrintName", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstPrintOrderStatusRegister.RecordCount = 0 Then Screen.MousePointer = vbNormal: On Error GoTo 0: Exit Sub
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Print Order Status Register")
    oExcel.DisplayAlerts = False
    If InStr(1, "XXYY", OrderType) > 0 Then
        Path = IIf(OrderType = "XX", "Pending Print Order Register (Busy)", "Pending Debit Note Register")
    Else
        Path = "Print Order Status Register (" & Choose(Val(OrderType), "Bookwise", "Print Orderwise", "", "", "Book Printerwise", "Title Printerwise", "", "Book Binderwise") & ")"
    End If
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\" & Path & " (" & CompCode & ")")
    oExcel.DisplayAlerts = True
    If OrderType <> "XX" Then
        For i = 1 To oExcel.Sheets.Count
            If InStr(1, "Sheet1", oExcel.Sheets(i).Name) = 0 Then oExcel.Sheets(i).Visible = False
        Next
        oExcel.Visible = False
        oExcel.Cells(1, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
        oExcel.Cells(2, "A").Value = IIf(OrderType = "YY", "Pending Debit Note Register", "Order Status Register (" & Choose(Val(OrderType), "Item-wise", "Order-wise", "", "", "Multi-Sheet-Party-wise", "Single-Sheet-Party-wise", "", "Item Party-wise") & ")") & " From " & Format(MhDateInput1, "dd-MMM-yyyy") & " To " & Format(MhDateInput2, "dd-MMM-yyyy")
        i = 5: Cnt = 1
        Do While Not rstPrintOrderStatusRegister.EOF
            If OrderType = "YY" Then If CheckEmpty(rstPrintOrderStatusRegister.Fields("BillNo").Value, False) Or Left(Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value), 1) = "*" Then GoTo Continue
            oExcel.Cells(i, "A").Value = Cnt
            oExcel.Application.Cells(i, "B").Value = Trim(rstPrintOrderStatusRegister.Fields("BookName").Value)
            oExcel.Application.Cells(i, "C").Value = ""
            oExcel.Application.Cells(i, "D").Value = Trim(rstPrintOrderStatusRegister.Fields("BookSize").Value) & "/" & Choose(Val(rstPrintOrderStatusRegister.Fields("FormType").Value), "08", "16", "04", "12", "24", "32", "64", "06", "02")
            If Val(CheckNull(rstPrintOrderStatusRegister.Fields("Forms1").Value)) <> 0 Then oExcel.Application.Cells(i, "E").Value = rstPrintOrderStatusRegister.Fields("Forms1").Value
            If Val(CheckNull(rstPrintOrderStatusRegister.Fields("Forms2").Value)) <> 0 Then oExcel.Application.Cells(i, "F").Value = rstPrintOrderStatusRegister.Fields("Forms2").Value
            If Val(CheckNull(rstPrintOrderStatusRegister.Fields("Forms4").Value)) <> 0 Then oExcel.Application.Cells(i, "G").Value = rstPrintOrderStatusRegister.Fields("Forms4").Value
            oExcel.Application.Cells(i, "H").Value = Trim(Str(rstPrintOrderStatusRegister.Fields("FrontPrintingType").Value))
            If Val(CheckNull(rstPrintOrderStatusRegister.Fields("BackPrintingType").Value)) <> 0 Then oExcel.Application.Cells(i, "H").Value = oExcel.Application.Cells(i, "H").Value & "+" & Trim(Str(rstPrintOrderStatusRegister.Fields("BackPrintingType").Value))
            oExcel.Application.Cells(i, "I").Value = IIf(Left(Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value), 1) = "*", Mid(Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value), 2), Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value))
            oExcel.Application.Cells(i, "J").Value = Format(rstPrintOrderStatusRegister.Fields("OrderDate").Value, "dd-MM-yy")
            oExcel.Application.Cells(i, "K").Value = rstPrintOrderStatusRegister.Fields(IIf(OrderType = "06", "TitleQuantity", IIf(OrderType = "05", "TextQuantity", "BookQuantity"))).Value
            oExcel.Application.Cells(i, "L").Value = rstPrintOrderStatusRegister.Fields("QuantityReceived").Value
            oExcel.Application.Cells(i, "M").Formula = "=K" & Trim(Str(i)) & "-L" & Trim(Str(i))
            oExcel.Application.Cells(i, "N").Value = Trim(rstPrintOrderStatusRegister.Fields("TextPrinterName").Value)
            For K = 22 To 36
                If IIf(IsNull(rstPrintOrderStatusRegister.Fields("TextStatus").Value), "N", rstPrintOrderStatusRegister.Fields("TextStatus").Value) = Trim(oExcel.Application.Cells(K, "XFC")) Then oExcel.Application.Cells(i, "O").Value = Trim(oExcel.Application.Cells(K, "XFD")): Exit For
            Next
            oExcel.Application.Cells(i, "P").Value = Trim(rstPrintOrderStatusRegister.Fields("TitlePrinterName").Value)
            For K = 39 To 53
                If IIf(IsNull(rstPrintOrderStatusRegister.Fields("TitleStatus").Value), "N", rstPrintOrderStatusRegister.Fields("TitleStatus").Value) = Trim(oExcel.Application.Cells(K, "XFC")) Then oExcel.Application.Cells(i, "R").Value = Trim(oExcel.Application.Cells(K, "XFD")): Exit For
            Next
            oExcel.Application.Cells(i, "R").Value = Trim(rstPrintOrderStatusRegister.Fields("BinderName").Value)
            For K = 6 To 13
                If IIf(IsNull(rstPrintOrderStatusRegister.Fields("BookStatus").Value), "N", rstPrintOrderStatusRegister.Fields("BookStatus").Value) = Trim(oExcel.Application.Cells(K, "XFC")) Then oExcel.Application.Cells(i, "S").Value = Trim(oExcel.Application.Cells(K, "XFD")): Exit For
            Next
            oExcel.Application.Cells(i, "T").Value = Trim(rstPrintOrderStatusRegister.Fields("Narration").Value)
            If Combo1.ListIndex > 0 Then
                T = GetStock(rstPrintOrderStatusRegister.Fields("Alias").Value)
                oExcel.Application.Cells(i, "U").Value = Val(Left(T, InStr(1, T, "|") - 1))
                If Combo1.ListIndex = 2 Then
                    oExcel.Application.Cells(i, "V").Value = Val(Mid(T, InStr(1, T, "|") + 1))
                    oExcel.Application.Cells(i, "W").Formula = "=V" & Trim(Str(i)) & "-U" & Trim(Str(i))
                End If
            End If
            If Not CheckEmpty(rstPrintOrderStatusRegister.Fields("Paper1").Value, False) Then
                oExcel.Application.Cells(i, "X").Value = rstPrintOrderStatusRegister.Fields("Paper1").Value
            ElseIf Not CheckEmpty(rstPrintOrderStatusRegister.Fields("Paper2").Value, False) Then
                oExcel.Application.Cells(i, "X").Value = rstPrintOrderStatusRegister.Fields("Paper2").Value
            ElseIf Not CheckEmpty(rstPrintOrderStatusRegister.Fields("Paper4").Value, False) Then
                oExcel.Application.Cells(i, "X").Value = rstPrintOrderStatusRegister.Fields("Paper4").Value
            End If
            oExcel.Application.Cells(i, "AC").Value = Trim(rstPrintOrderStatusRegister.Fields("BillNo").Value)
            oExcel.Application.Cells(i, "AD").Value = Format(rstPrintOrderStatusRegister.Fields("BillDate").Value, "dd-MM-yy")
            oExcel.Application.Cells(i, "AE").Value = Trim(rstPrintOrderStatusRegister.Fields("TotalPlates").Value)
            oExcel.Application.Cells(i, "AF").Value = Trim(rstPrintOrderStatusRegister.Fields("PBillAmount").Value)
            oExcel.Application.Cells(i, "XFB").Value = rstPrintOrderStatusRegister.Fields("Code").Value
            MdiMainMenu.StatusBar1.Panels(2).Text = "Processed record #" & Trim(Str(Cnt)) & " of " & Trim(Str(rstPrintOrderStatusRegister.RecordCount)) & " !!!"
            Cnt = Cnt + 1: i = i + 1
Continue:
            rstPrintOrderStatusRegister.MoveNext
        Loop
        MdiMainMenu.StatusBar1.Panels(2).Text = ""
        oExcel.Range("Y5:AB" & Trim(Str(i - 1))).Formula = oExcel.Range("Y5:AB5").Formula
        oExcel.Columns("A:AD").EntireColumn.AutoFit
        oExcel.Columns("C:C").Hidden = True
        If Combo1.ListIndex = 0 Then
            oExcel.Columns("U:W").Hidden = True
        ElseIf Combo1.ListIndex = 1 Then
            oExcel.Columns("V:W").Hidden = True
        End If
        oExcel.Columns("AC:AD").Hidden = True: oExcel.Columns("H").Hidden = True: oExcel.Columns("C").Hidden = True: oExcel.Columns("Y:AA").Hidden = True
        If OrderType = "06" Then oExcel.Columns("H").Hidden = False
        If OrderType = "YY" Then oExcel.Columns("O").Hidden = True: oExcel.Columns("Q").Hidden = True: oExcel.Columns("U:AB").Hidden = True: oExcel.Columns("AC:AD").Hidden = False
        oExcel.Range("A4:XFB4").AutoFilter
    Else
        For i = 1 To oExcel.Sheets.Count
            If InStr(1, "Sheet2", oExcel.Sheets(i).Name) = 0 Then oExcel.Sheets(i).Visible = False
        Next
        oExcel.Visible = False
        oExcel.Cells(1, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
        oExcel.Cells(2, "A").Value = "Pending Print Order Register (Busy) From " & Format(MhDateInput1, "dd-MMM-yyyy") & " To " & Format(MhDateInput2, "dd-MMM-yyyy")
        i = 5: Cnt = 1
        Do While Not rstPrintOrderStatusRegister.EOF
            If Left(Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value), 1) = "*" Then GoTo Skip
            GetBusyOrder (Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value))
            If rstBusy.RecordCount > 0 Then If Trim(rstBusy.Fields("VchNo").Value) = Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value) And Format(rstBusy.Fields("Date").Value, "yyyymmdd") = Format(rstPrintOrderStatusRegister.Fields("OrderDate").Value, "yyyymmdd") And Trim(rstBusy.Fields("Alias").Value) = Trim(rstPrintOrderStatusRegister.Fields("Alias").Value) And Val(rstBusy.Fields("OrderedQuantity").Value) = Val(rstPrintOrderStatusRegister.Fields("OrderedQuantity").Value) Then GoTo Skip
            oExcel.Cells(i, "A").Value = Cnt
            oExcel.Application.Cells(i, "B").Value = Trim(rstPrintOrderStatusRegister.Fields("BookName").Value)
            oExcel.Application.Cells(i, "C").Value = Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value)
            oExcel.Application.Cells(i, "D").Value = Format(rstPrintOrderStatusRegister.Fields("OrderDate").Value, "dd-MMM-yyyy")
            oExcel.Application.Cells(i, "E").Value = Format(rstBusy.Fields("Date").Value, "dd-MMM-yyyy")
            oExcel.Application.Cells(i, "F").Value = Val(rstPrintOrderStatusRegister.Fields("OrderedQuantity").Value)
            oExcel.Application.Cells(i, "G").Value = Val(rstBusy.Fields("OrderedQuantity").Value)
            oExcel.Application.Cells(i, "H").Value = Trim(CheckNull(rstPrintOrderStatusRegister.Fields("BookPrinterName").Value))
            oExcel.Application.Cells(i, "I").Value = Trim(CheckNull(rstPrintOrderStatusRegister.Fields("BinderName").Value))
            oExcel.Application.Cells(i, "J").Value = Trim(rstBusy.Fields("BinderName").Value)
            oExcel.Application.Cells(i, "K").Value = Trim(rstPrintOrderStatusRegister.Fields("LaminatorAlias").Value)
            Cnt = Cnt + 1: i = i + 1
Skip:
            MdiMainMenu.StatusBar1.Panels(2).Text = "Processed record #" & Trim(Str(Cnt)) & " of " & Trim(Str(rstPrintOrderStatusRegister.RecordCount)) & " !!!"
            rstPrintOrderStatusRegister.MoveNext
        Loop
        MdiMainMenu.StatusBar1.Panels(2).Text = ""
        oExcel.Columns("A:J").EntireColumn.AutoFit
    End If
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then oExcel.Range("A1").Activate: oExcel.Visible = True Else oExcel.Workbooks.Item(1).PrintOut
    Set oExcel = Nothing
    Call CloseConnection(cnBusy)
    On Error GoTo 0
End Sub
