VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBillRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Account Ledger"
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6540
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
   ScaleHeight     =   6660
   ScaleWidth      =   6540
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
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
            Picture         =   "BillRegister.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BillRegister.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BillRegister.frx":0658
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BillRegister.frx":076A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6255
      Left            =   45
      TabIndex        =   7
      Top             =   345
      Width           =   6450
      _Version        =   65536
      _ExtentX        =   11377
      _ExtentY        =   11033
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
      Picture         =   "BillRegister.frx":087C
      Begin VB.OptionButton Option3 
         Caption         =   "Open"
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
         Left            =   3765
         TabIndex        =   2
         Top             =   10
         Value           =   -1  'True
         Width           =   1020
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
         Left            =   5775
         TabIndex        =   4
         Top             =   10
         Width           =   630
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
         Left            =   4905
         TabIndex        =   3
         Top             =   10
         Width           =   750
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5940
         Left            =   0
         TabIndex        =   5
         Top             =   315
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   10478
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
         Caption         =   " From"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BillRegister.frx":0898
         Picture         =   "BillRegister.frx":08B4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   9
         Top             =   0
         Width           =   645
         _Version        =   65536
         _ExtentX        =   1138
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
         Picture         =   "BillRegister.frx":08D0
         Picture         =   "BillRegister.frx":08EC
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
         Calendar        =   "BillRegister.frx":0908
         Caption         =   "BillRegister.frx":0A20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BillRegister.frx":0A8C
         Keys            =   "BillRegister.frx":0AAA
         Spin            =   "BillRegister.frx":0B08
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
         Left            =   2550
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BillRegister.frx":0B30
         Caption         =   "BillRegister.frx":0C48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BillRegister.frx":0CB4
         Keys            =   "BillRegister.frx":0CD2
         Spin            =   "BillRegister.frx":0D30
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
Attribute VB_Name = "FrmBillRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstBillRegister As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim AccountType As String, ReportType As Byte
Dim oOutlook As New Outlook.Application
Dim OutputTo As String
Public VchCodeType As String
Public VchCode As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    Me.Caption = IIf(Left(VchCodeType, 2) = 2, " Purchase", " Sales") & IIf(Right(VchCodeType, 1) = 1, " Order Ledger Detailed", " Order Ledger Summarised")
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster Where FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    Option3.Value = True
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    MhDateInput2.Text = IIf(Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd"), Format(FinancialYearTo, "dd-mm-yyyy"), Format(Date, "dd-mm-yyyy"))
    rstAccountList.Open "SELECT Name,Code FROM AccountMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    Call FillList(ListView1, "List of Accounts...", rstAccountList)
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
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstBillRegister)
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
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 4 Then CloseForm Me: Exit Sub
    OutputTo = Choose(Button.Index, "S", "P", "M")
    PrintBillRegister
End Sub
Private Sub PrintBillRegister()
    On Error Resume Next
    Dim SQL As String
    Dim i As Integer, iCount As Integer
    For i = 1 To ListView1.ListItems.Count
       If ListView1.ListItems(i).Checked Then iCount = iCount + 1
    Next
    Screen.MousePointer = vbHourglass
    rptBillRegister.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptBillRegister.Text9.SetText IIf(Left(VchCodeType, 1) = 1, " Purchase Order ", " Sales Order ") + "" + IIf(Option3.Value, "Open ", IIf(Option1.Value, "Close ", "Both-Open & Close ")) + IIf(Right(VchCodeType, 1) = "1", "( Detailed ", "( Summarised ") + " Ledger ) From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]"
'    rptBillRegister.Text9.SetText IIf(Option3.Value, "Pending Payment ", IIf(Option1.Value, "Paid Payment ", "")) & "Account Ledger From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]"
    If rstBillRegister.State = adStateOpen Then rstBillRegister.Close
    VchCode = IIf(Left(VchCodeType, 1) = 1, "S", "P")
        If Right(VchCodeType, 1) = 1 Then  'rstBillRegister.Open
                        SQL = "SELECT LTRIM(M1.PrintName) As AccountName,'P'+'-PO/'+'" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+LTRIM(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,P.BillNo,P.BillDate,Format(C.Quantity,'0.000') As Quantity,P.BillAmount,P.PaidAmount,M1.Code FROM ((PaperPOParent P INNER JOIN PaperPOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Supplier) INNER JOIN PaperMaster M2 ON M2.Code=C.Paper WHERE 'P'<>'" & Left(VchCode, 1) & "' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "P.BillNo=''", IIf(Option1.Value, "P.BillNo<>''", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,'P'+'-GI/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+LTRIM(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,P.BillNo,P.BillDate,Format(C.Quantity,'0') As Quantity,P.BillAmount,P.PaidAmount,M1.Code FROM ((OutsourceItemPOParent P INNER JOIN OutsourceItemPOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Supplier) INNER JOIN OutsourceItemMaster M2 ON M2.Code=C.OutsourceItem WHERE 'P'<>'" & Left(VchCode, 1) & "' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "P.BillNo=''", IIf(Option1.Value, "P.BillNo<>''", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-MF-Ptg/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,IIf(Right(P.Type,1)='S',C.BillAmount,0) As BillAmount,IIf(Right(P.Type,1)='S',0,C.BillAmount) As PaidAmount,M1.Code FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.BookPrinter) INNER JOIN BookMaster M2 ON M2.Code=P.Book " & _
                               "WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-MF-Plate/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.PBillNo,C.PBillDate,Format(C.ActualQuantity,'0') As Quantity,IIf(Right(P.Type,1)='S',C.PBillAmount,0) As BillAmount,IIf(Right(P.Type,1)='S',0,C.BillAmount) As PPaidAmount,M1.Code FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=C.PlateMaker) INNER JOIN BookMaster M2 ON M2.Code=P.Book " & _
                               "WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-MF-Paper/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,IIf(Right(P.Type,1)='S',C.RBillAmount,0) As BillAmount,IIf(Right(P.Type,1)='S',0,C.RBillAmount) As PaidAmount,M1.Code FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.BookPrinter) INNER JOIN BookMaster M2 ON M2.Code=P.Book " & _
                               "WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-SF_Ptg/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,IIf(Right(P.Type,1)='S',C.BillAmount,0) As BillAmount,IIf(Right(P.Type,1)='S',0,C.BillAmount) As PaidAmount,M1.Code FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.TitlePrinter) INNER JOIN BookMaster M2 ON M2.Code=P.Book " & _
                               "WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-SF_Plate/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.PBillNo,C.PBillDate,Format(C.ActualQuantity,'0') As Quantity,IIf(Right(P.Type,1)='S',C.PBillAmount,0) As BillAmount,IIf(Right(P.Type,1)='S',0,C.PBillAmount) As PaidAmount,M1.Code FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=C.PlateMaker) INNER JOIN BookMaster M2 ON M2.Code=P.Book " & _
                               "WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-SF_Paper/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,IIf(Right(P.Type,1)='S',C.RBillAmount,0) As BillAmount,IIf(Right(P.Type,1)='S',0,C.RBillAmount) As PaidAmount,M1.Code FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.TitlePrinter) INNER JOIN BookMaster M2 ON M2.Code=P.Book " & _
                               "WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-CF_Ptg/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,IIf(Right(P.Type,1)='S',C.PrintAmount,0) As BillAmount,IIf(Right(P.Type,1)='S',0,C.PrintAmount) As PaidAmount,M1.Code FROM ((BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.TitlePrinter) INNER JOIN BookMaster M2 ON M2.Code=P.Book " & _
                               "WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-CF_Plate/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.PBillNo,C.PBillDate,Format(C.ActualQuantity,'0') As Quantity,IIf(Right(P.Type,1)='S',C.PlateAmount,0) As BillAmount,IIf(Right(P.Type,1)='S',0,C.PlateAmount) As PaidAmount,M1.Code FROM ((BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=C.PlateMaker) INNER JOIN BookMaster M2 ON M2.Code=P.Book " & _
                               "WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-CF_Paper/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,IIf(Right(P.Type,1)='S',C.PaperAmount,0) As BillAmount,IIf(Right(P.Type,1)='S',0,C.PaperAmount) As PaidAmount,M1.Code FROM ((BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=C.PlateMaker) INNER JOIN BookMaster M2 ON M2.Code=P.Book " & _
                               "WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-MO/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format((SELECT Sum(C2.Quantity) FROM BookPOChild07 C2 WHERE C2.Code=P.Code),'0') As Quantity,IIf(Right(P.Type,1)='S',(SELECT Sum(C2.BillAmount) FROM BookPOChild07 C2 WHERE C2.Code=P.Code),0) As BillAmount,IIf(Right(P.Type,1)='S',0,(SELECT Sum(C2.BillAmount) FROM BookPOChild07 C2 WHERE C2.Code=P.Code)) As PaidAmount,M1.Code " & _
                               "FROM ((BookPOParent P Left JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Laminator) INNER JOIN BookMaster M2 ON M2.Code=P.Book WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") UNION " & _
                               "SELECT LTRIM(M1.PrintName) As AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-BP/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,IIf(Right(P.Type,1)='S',C.BillAmount,0) As BillAmount,IIf(Right(P.Type,1)='S',0,C.BillAmount) As PaidAmount,M1.Code FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Binder) INNER JOIN BookMaster M2 ON M2.Code=P.Book " & _
                               "WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") " & _
                               " ORDER BY AccountName,VchDate,VchNo" ', cnDatabase, adOpenKeyset, adLockOptimistic
       Else
                    SQL = "SELECT  IIF(LTRIM(A1.PrintName) Is Not NULL,LTRIM(A1.PrintName),IIF(LTRIM(A2.PrintName) Is Not NULL,LTRIM(A2.PrintName),IIF(LTRIM(A3.PrintName) Is Not NULL,LTRIM(A3.PrintName),IIF(LTRIM(A4.PrintName) Is Not NULL,LTRIM(A4.PrintName),IIF(LTRIM(A5.PrintName) Is Not NULL,LTRIM(A5.PrintName),IIF(LTRIM(A6.PrintName) Is Not NULL,LTRIM(A6.PrintName),LTRIM(A7.PrintName))))))) AS AccountName,IIf(Right(P.Type,1)='S','S','P')+'-'+LTRIM(P.Name)+'-FG/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(P.EstQty01,'0') As Quantity," & _
                              "IIf(Right(P.Type,1)='S',IIF((SELECT C.BillAmount+C.PBillAmount+C.RBillAmount FROM BookPOChild05 C WHERE C.Code=P.Code) IS NULL,0,(SELECT C.BillAmount+C.PBillAmount+C.RBillAmount FROM BookPOChild05 C WHERE C.Code=P.Code))+IIF((SELECT C1.BillAmount+C1.PBillAmount+C1.RBillAmount FROM BookPOChild06 C1 WHERE C1.Code=P.Code) IS NULL,0,(SELECT C1.BillAmount+C1.PBillAmount+C1.RBillAmount FROM BookPOChild06 C1 WHERE C1.Code=P.Code))+IIF((SELECT Sum(C2.BillAmount) FROM BookPOChild07 C2 WHERE C2.Code=P.Code) IS NULL,0,(SELECT Sum(C2.BillAmount) FROM BookPOChild07 C2 WHERE C2.Code=P.Code))+IIF((SELECT C3.BillAmount FROM BookPOChild08 C3 WHERE C3.Code=P.Code) IS NULL,0,(SELECT C3.BillAmount FROM BookPOChild08 C3 WHERE C3.Code=P.Code))+IIF((SELECT C4.PrintAmount+C4.PlateAmount+C4.PaperAmount FROM BookPOChild09 C4 WHERE C4.Code=P.Code) IS NULL,0,(SELECT C4.PrintAmount+C4.PlateAmount+C4.PaperAmount FROM BookPOChild09 C4 WHERE C4.Code=P.Code)),0) As BillAmount," & _
                              "IIf(Right(P.Type,1)='S',0,IIF((SELECT C.BillAmount+C.PBillAmount+C.RBillAmount FROM BookPOChild05 C WHERE C.Code=P.Code) IS NULL,0,(SELECT C.BillAmount+C.PBillAmount+C.RBillAmount FROM BookPOChild05 C WHERE C.Code=P.Code))+IIF((SELECT C1.BillAmount+C1.PBillAmount+C1.RBillAmount FROM BookPOChild06 C1 WHERE C1.Code=P.Code) IS NULL,0,(SELECT C1.BillAmount+C1.PBillAmount+C1.RBillAmount FROM BookPOChild06 C1 WHERE C1.Code=P.Code))+IIF((SELECT Sum(C2.BillAmount) FROM BookPOChild07 C2 WHERE C2.Code=P.Code) IS NULL,0,(SELECT Sum(C2.BillAmount) FROM BookPOChild07 C2 WHERE C2.Code=P.Code))+IIF((SELECT C3.BillAmount FROM BookPOChild08 C3 WHERE C3.Code=P.Code) IS NULL,0,(SELECT C3.BillAmount FROM BookPOChild08 C3 WHERE C3.Code=P.Code))+IIF((SELECT C4.PrintAmount+C4.PlateAmount+C4.PaperAmount FROM BookPOChild09 C4 WHERE C4.Code=P.Code) IS NULL,0,(SELECT C4.PrintAmount+C4.PlateAmount+C4.PaperAmount FROM BookPOChild09 C4 WHERE C4.Code=P.Code))) As PaidAmount," & _
                              "IIF(LTRIM(A1.Code) Is Not NULL,LTRIM(A1.Code),IIF(LTRIM(A2.Code) Is Not NULL,LTRIM(A2.Code),IIF(LTRIM(A3.Code) Is Not NULL,LTRIM(A3.Code),IIF(LTRIM(A4.Code) Is Not NULL,LTRIM(A4.Code),IIF(LTRIM(A5.Code) Is Not NULL,LTRIM(A5.Code),IIF(LTRIM(A6.Code) Is Not NULL,LTRIM(A6.Code),LTRIM(A7.Code))))))) AS ACode,P.UnitRate,(SELECT Sum(C6.Quantity) FROM JobworkBVChild C6 WHERE C6.Ref=P.Code) As QtyBilled,(SELECT Max(C7.Rate) FROM JobworkBVChild C7 WHERE C7.Ref=P.Code) As Rate,(SELECT Sum(C8.Amount) FROM JobworkBVChild C8 WHERE C8.Ref=P.Code) As Billed,(BillAmount-Billed) As ToBeBilled " & _
                              "FROM (((((((((((BookPOParent AS P Left JOIN BookPOChild05 AS C ON P.Code=C.Code) LEFT JOIN BookPOChild06 AS C1 ON P.Code=C1.Code) LEFT JOIN BookPOChild08 AS C3 ON P.Code=C3.Code) LEFT JOIN BookPOChild09 AS C4 ON P.Code=C4.Code) Left JOIN AccountMaster AS A1 ON A1.Code=P.BookPrinter) Left JOIN AccountMaster AS A2 ON A2.Code=P.TitlePrinter) Left JOIN AccountMaster AS A3 ON A3.Code=P.Binder) Left JOIN AccountMaster AS A4 ON A4.Code=P.Laminator) Left JOIN AccountMaster AS A5 ON A5.Code=C.PlateMaker) Left JOIN AccountMaster AS A6 ON A6.Code=C1.PlateMaker)Left JOIN AccountMaster AS A7 ON A7.Code=C4.PlateMaker) INNER JOIN BookMaster AS M2 ON M2.Code=P.Book " & _
                              "WHERE LEFT(P.Type,1)<>'O' AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "'  AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)<P.EstQty01", IIf(Option1.Value, "IIf(Right(P.Type,1)='P',P.QuantityReceivedB-P.QuantityIssuedB,P.QuantityIssuedB-P.QuantityReceivedB)>=P.EstQty01", 1)) & " AND IIF(LTRIM(A1.Code) Is Not NULL,LTRIM(A1.Code),IIF(LTRIM(A2.Code) Is Not NULL,LTRIM(A2.Code),IIF(LTRIM(A3.Code) Is Not NULL,LTRIM(A3.Code),IIF(LTRIM(A4.Code) Is Not NULL,LTRIM(A4.Code),IIF(LTRIM(A5.Code) Is Not NULL,LTRIM(A5.Code),IIF(LTRIM(A6.Code) Is Not NULL,LTRIM(A6.Code),LTRIM(A7.Code))))))) IN (" & SelectedItems(ListView1) & ") " & _
                              " ORDER BY P.Code,IIF(LTRIM(A1.PrintName) Is Not NULL,LTRIM(A1.PrintName),IIF(LTRIM(A2.PrintName) Is Not NULL,LTRIM(A2.PrintName),IIF(LTRIM(A3.PrintName) Is Not NULL,LTRIM(A3.PrintName),IIF(LTRIM(A4.PrintName) Is Not NULL,LTRIM(A4.PrintName),IIF(LTRIM(A5.PrintName) Is Not NULL,LTRIM(A5.PrintName),IIF(LTRIM(A6.PrintName) Is Not NULL,LTRIM(A6.PrintName),LTRIM(A7.PrintName))))))),P.Date,P.Name" ', cnDatabase, adOpenKeyset, adLockOptimistic
    End If
    If DatabaseType = "MS SQL" Then SQL = Replace(SQL, "#", "'")
    Screen.MousePointer = vbNormal
    If rstBillRegister.State = adStateOpen Then rstBillRegister.Close
    rstBillRegister.Open SQL, cnDatabase, adOpenKeyset, adLockOptimistic
    rstBillRegister.ActiveConnection = Nothing
    If rstBillRegister.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rstBillRegister.MoveFirst
    rptBillRegister.Database.SetDataSource rstBillRegister, 3, 1
    rptBillRegister.Text6.SetText IIf(Right(VchCodeType, 1) = 1, "", "UnitRate")
    rptBillRegister.Text13.SetText IIf(Right(VchCodeType, 1) = 1, "", "Billed Amt.")
    rptBillRegister.Text14.SetText IIf(Right(VchCodeType, 1) = 1, "", "Pending Bill Amt.")
    rptBillRegister.DiscardSavedData
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptBillRegister: FrmReportViewer.Show vbModal
    ElseIf OutputTo = "P" Then
        rptBillRegister.PaperSource = crPRBinAuto
        rptBillRegister.PrintOut
    Else
        If iCount >= 1 Then
            Dim oOutlookMsg As Outlook.MailItem, FileName As String
            Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
            With oOutlookMsg
                '.To = rstPaperDebitNote.Fields("EMail").Value
                .Subject = IIf(Option3.Value, "Pending Payment  ", IIf(Option1.Value, "Paid Payment", "")) & "Account Ledger"
                .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith " & IIf(Option3.Value, "Pending Payment ", IIf(Option1.Value, "Paid Payment", "")) & "Account Ledger from " + Format(GetDate(MhDateInput1.Text), "dd-MMM-yyyy") + " to " + Format(GetDate(MhDateInput2.Text), "dd-MMM-yyyy") & " for doing the needful at your end.<Br><b>Kindly do acknowledge the receipt of the mail</b>.<Br><Br>Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
                rptBillRegister.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptBillRegister.ExportOptions.DestinationType = crEDTDiskFile
                FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
                rptBillRegister.ExportOptions.DiskFileName = FileName
                rptBillRegister.Export False
                .Attachments.Add (FileName)
                .Importance = olImportanceHigh
                .ReadReceiptRequested = True
                If CheckEmpty(.To, False) Then .Display Else .Send
            End With
            Set oOutlookMsg = Nothing
        End If
    End If
    Set rptBillRegister = Nothing
    On Error GoTo 0
End Sub
