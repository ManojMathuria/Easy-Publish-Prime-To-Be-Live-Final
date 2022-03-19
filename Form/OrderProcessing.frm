VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrderProcessing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Order Processing"
   ClientHeight    =   6510
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
   ScaleHeight     =   6510
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
            Picture         =   "OrderProcessing.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrderProcessing.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrderProcessing.frx":0658
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrderProcessing.frx":076A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6140
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   9675
      _Version        =   65536
      _ExtentX        =   17066
      _ExtentY        =   10830
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
      Picture         =   "OrderProcessing.frx":087C
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
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
         Picture         =   "OrderProcessing.frx":0898
         Picture         =   "OrderProcessing.frx":08B4
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
         Picture         =   "OrderProcessing.frx":08D0
         Picture         =   "OrderProcessing.frx":08EC
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   0
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "OrderProcessing.frx":0908
         Caption         =   "OrderProcessing.frx":0A20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "OrderProcessing.frx":0A8C
         Keys            =   "OrderProcessing.frx":0AAA
         Spin            =   "OrderProcessing.frx":0B08
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
         Calendar        =   "OrderProcessing.frx":0B30
         Caption         =   "OrderProcessing.frx":0C48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "OrderProcessing.frx":0CB4
         Keys            =   "OrderProcessing.frx":0CD2
         Spin            =   "OrderProcessing.frx":0D30
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
         Height          =   5820
         Left            =   0
         TabIndex        =   9
         Top             =   315
         Width           =   4850
         _ExtentX        =   8546
         _ExtentY        =   10266
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
         Height          =   5820
         Left            =   4830
         TabIndex        =   10
         Top             =   315
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   10266
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
         Left            =   5205
         TabIndex        =   2
         Top             =   30
         Value           =   -1  'True
         Width           =   1380
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
         Left            =   6705
         TabIndex        =   3
         Top             =   30
         Width           =   1380
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
         Left            =   8175
         TabIndex        =   4
         Top             =   30
         Width           =   1380
      End
   End
End
Attribute VB_Name = "FrmOrderProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstOrderProcessing As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstItemList As New ADODB.Recordset
Dim AccountType As String, ReportType As Byte
Dim oOutlook As New Outlook.Application
Dim OutputTo As String
Public VchCode As String
Public VchCodeType As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If VchCodeType = 11 Or VchCodeType = 12 Or VchCodeType = 13 Or VchCodeType = 14 Or VchCodeType = 15 Or VchCodeType = 21 Or VchCodeType = 22 Or VchCodeType = 23 Or VchCodeType = 24 Or VchCodeType = 25 Then
    Me.Caption = IIf(Left(VchCodeType, 1) = 1, " Purchase", " Sales") & IIf(Right(VchCodeType, 1) = 1, " Order Processing [ Jobwork-Detailed ]", IIf(Right(VchCodeType, 1) = 3, " Order Processing [ Order-Wise ]", IIf(Right(VchCodeType, 1) = 4, " Order Processing [ Party-Wise ]", IIf(Right(VchCodeType, 1) = 7, " Order Processing [ Party-Wise ]", IIf(Right(VchCodeType, 1) = 5, " Order Processing [ Item-Wise ]", IIf(Right(VchCodeType, 1) = 6, " Order Processing [ Item-Wise ]", " Order Processing [ Jobwork-Summarised ]"))))))
    ElseIf VchCodeType = 16 Or VchCodeType = 17 Or VchCodeType = 18 Or VchCodeType = 19 Or VchCodeType = 20 Then
    Me.Caption = " Issue-Receipt Detailed " & IIf(Right(VchCodeType, 1) = 6, " [ Item-Wise ]", IIf(Right(VchCodeType, 1) = 7, " [ Item Party-Wise ]", IIf(Right(VchCodeType, 1) = 8, " [ Item Group-Wise ]", IIf(Right(VchCodeType, 1) = 9, " [ Item Voucher-Wise ]", IIf(Right(VchCodeType, 1) = 0, " [ Item Date-Wise ]", " [ '' ]")))))
    Option1.Caption = "Receipt": Option3.Caption = "Issue": Option2.Caption = "Both":
    ElseIf VchCodeType = 26 Or VchCodeType = 27 Or VchCodeType = 28 Or VchCodeType = 29 Or VchCodeType = 30 Then
    Me.Caption = " Issue-Receipt Summarised " & IIf(Right(VchCodeType, 1) = 6, " [ Item-Wise ]", IIf(Right(VchCodeType, 1) = 7, " [ Item Party-Wise ]", IIf(Right(VchCodeType, 1) = 8, " [ Item Group-Wise ]", IIf(Right(VchCodeType, 1) = 9, " [ Item Voucher-Wise ]", IIf(Right(VchCodeType, 1) = 0, " [ Item Date-Wise ]", " [ '' ]")))))
    End If
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    Option3.Value = True
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    MhDateInput2.Text = IIf(Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd"), Format(FinancialYearTo, "dd-mm-yyyy"), Format(Date, "dd-mm-yyyy"))
    rstAccountList.Open "SELECT Name,Code FROM AccountMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstItemList.Open "SELECT Name,Code FROM BookMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    Call FillList(ListView1, "List of Accounts...", rstAccountList)
    Call FillList(ListView2, "List of Items...", rstItemList)
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
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstOrderProcessing)
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
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 4 Then CloseForm Me: Exit Sub
    OutputTo = Choose(Button.Index, "S", "P", "M")
    PrintOrderProcessing
End Sub
Private Sub PrintOrderProcessing()
    On Error Resume Next
    Dim SQL As String
    Dim i As Integer, iCount As Integer
    For i = 1 To ListView1.ListItems.Count
       If ListView1.ListItems(i).Checked Then iCount = iCount + 1
    Next
    Dim j As Integer, jCount As Integer
    For j = 1 To ListView2.ListItems.Count
       If ListView2.ListItems(j).Checked Then iCount = jCount + 1
    Next
    Screen.MousePointer = vbHourglass
    rptOrderProcessing.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    If VchCodeType = 11 Or VchCodeType = 12 Or VchCodeType = 21 Or VchCodeType = 22 Or VchCodeType = 13 Or VchCodeType = 14 Or VchCodeType = 15 Or VchCodeType = 23 Or VchCodeType = 24 Or VchCodeType = 25 Then
        rptOrderProcessing.Text9.SetText IIf(Left(VchCodeType, 1) = 1, " Purchase Order ", " Sales Order ") + IIf(Right(VchCodeType, 1) = 1, " Party-Wise", IIf(Right(VchCodeType, 1) = 2, " Party-Wise", IIf(Right(VchCodeType, 1) = 3, " Order-Wise", IIf(Right(VchCodeType, 1) = 4, " Party-Wise", IIf(Right(VchCodeType, 1) = 5, " Item-Wise", " ''"))))) + " " + IIf(Option3.Value, " ( Pending-", IIf(Option1.Value, " ( Close -", " ( Both-Pending & Close  ")) + IIf(Right(VchCodeType, 1) = 1, " Jobwork-Detailed )", IIf(Right(VchCodeType, 2) = 13, " Detailed )", IIf(Right(VchCodeType, 2) = 14, " Detailed )", IIf(Right(VchCodeType, 2) = 15, " Detailed )", IIf(Right(VchCodeType, 2) = 23, " Detailed )", IIf(Right(VchCodeType, 2) = 24, " Detailed )", IIf(Right(VchCodeType, 2) = 25, " Detailed )", " Jobwork-Summarised )"))))))) + "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]"
    ElseIf VchCodeType = 16 Or VchCodeType = 17 Or VchCodeType = 18 Or VchCodeType = 19 Or VchCodeType = 20 Then
        rptOrderProcessing.Text9.SetText IIf(Option1.Value, "  Item Receipt-", IIf(Option3.Value, "  Item Issue-", "  Item Issue-Receipt  ")) + IIf(Right(VchCodeType, 1) = 6, " Item-Wise", IIf(Right(VchCodeType, 1) = 7, " Party-Wise", IIf(Right(VchCodeType, 1) = 8, " Group-Wise", IIf(Right(VchCodeType, 1) = 9, " Voucher-Wise", IIf(Right(VchCodeType, 1) = 0, " Date-Wise", " ''"))))) + " " + IIf(Left(VchCodeType, 1) = 1, " ( Detailed )", IIf(Left(VchCodeType, 2) = 20, " ( Detailed )", " ( Summarised )")) + "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]"
    End If
        If rstOrderProcessing.State = adStateOpen Then rstOrderProcessing.Close
        VchCode = IIf(Left(VchCodeType, 1) = 1, "S", IIf(Left(VchCodeType, 1) = 1, "S", "P"))
        If Right(VchCodeType, 1) = 1 Then
            SQL = "SELECT IIF(B.PrintName IS NOT NULL,LTRIM(B.PrintName),IIF(A1.PrintName IS NOT NULL,LTRIM(A1.PrintName),IIF(A2.PrintName IS NOT NULL,A2.PrintName,IIF(A3.PrintName IS NOT NULL,LTRIM(A3.PrintName),LTRIM(A4.PrintName))))) As AccountName,RIGHT(T.Type,1)+'O/'+LTRIM(T.Name)+'/JW/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,T.Date As VchDate,I.PrintName As ItemName,P.ChallanNo,P.ChallanDate,T.EstQty01 As Quantity,IIF(LEFT(BOM,18) IN ('0310FI','0710XXXXXXXXXXXXFI','0110FI','0510XXXXXXXXXXXXFI','0000'),C.Quantity,0) As ReceivedQty,ABS(IIF(LEFT(BOM,18) IN ('0410FI','0810XXXXXXXXXXXXFI','0210FI','0610XXXXXXXXXXXXFI','0000'),C.Quantity,0)) As IssuedQty,IIF(B.Code IS NOT NULL,LTRIM(B.Code),IIF(A1.Code IS NOT NULL,LTRIM(A1.Code),IIF(A2.Code IS NOT NULL,LTRIM(A2.Code),IIF(A3.Code IS NOT NULL,LTRIM(A3.Code),LTRIM(A4.Code))))) As BCode," & _
                        "LTRIM(P.Name) As GRNNo,P.Date,A.Name As Consignee,B.Name As Party,P.Remarks,LTRIM(T.Name) As PO,I.Name As Book,C.Quantity As Qty,C.Rate,C.Amount,P.BOX,P.Freight,P.TYPE,RIGHT(P.Type,2)+'-'+LTRIM(P.Name) As MRNNo,RIGHT(T.Type,1) As OrderType,I.Code,IIF(LEFT(P.Type,4)='0110','Purchase',IIF(LEFT(P.Type,4)='0210','Purchase Return',IIF(LEFT(P.Type,4)='0510','Pur Challan IN',IIF(LEFT(P.Type,4)='0610','Pur Challan Out',IIF(LEFT(P.Type,4)='0310','Sales Return',IIF(LEFT(P.Type,4)='0410','Sales',IIF(LEFT(P.Type,4)='0710','Sales Challan IN',IIF(LEFT(P.Type,4)='0810','Sales Challan Out','Order Processing')))))))) As TypeRef,IIF(C.BOM IS NOT Null,C.BOM,'0000FI') As BOM,T.Code,P.Remarks As RemarkC " & _
                        "FROM ((((((((BookPOParent T INNER JOIN BookMaster As I ON T.Book=I.Code) LEFT JOIN JobworkBVChild C ON C.Ref=T.Code) LEFT JOIN JobworkBVParent P ON P.Code=C.Code) LEFT JOIN AccountMaster As B ON P.Party=B.Code) LEFT JOIN AccountMaster As A ON P.Consignee=A.Code) LEFT JOIN AccountMaster As A1 ON T.Binder=A1.Code) LEFT JOIN AccountMaster As A2 ON T.BookPrinter=A2.Code) LEFT JOIN AccountMaster As A3 ON T.TitlePrinter=A3.Code) LEFT JOIN AccountMaster As A4 ON T.Laminator=A4.Code WHERE LEFT(IIF(C.BOM IS NOT Null,C.BOM,'0000FI'),18) IN " & IIf(Left(VchCode, 1) = "S", "('0110FI','0210FI','0510XXXXXXXXXXXXFI','0610XXXXXXXXXXXXFI','0000FI')", "('0310FI','0410FI','0710XXXXXXXXXXXXFI','0810XXXXXXXXXXXXFI','0000FI')") & " AND LEFT(T.Type,1)<>'O' AND RIGHT(T.Type,1)<>'" & Left(VchCode, 1) & "' AND LEFT(T.Code,1)<>'*' AND T.Date>='" & GetDate(MhDateInput1.Text) & "' AND T.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & _
                        IIf(Option3.Value, "T.DeliveredQuantityC+T.DeliveredQuantityB<T.EstQty01", IIf(Option1.Value, "T.DeliveredQuantityC+T.DeliveredQuantityB>=T.EstQty01", "1=1")) & " AND IIF(B.Code IS NOT NULL,LTRIM(B.Code),IIF(A1.Code IS NOT NULL,LTRIM(A1.Code),IIF(A2.Code IS NOT NULL,LTRIM(A2.Code),IIF(A3.Code IS NOT NULL,LTRIM(A3.Code),LTRIM(A4.Code))))) IN (" & SelectedItems(ListView1) & ") AND I.Code IN (" & SelectedItems(ListView2) & ") ORDER BY AccountName,T.Name"
        ElseIf Right(VchCodeType, 1) = 2 Then
            SQL = "SELECT IIF(B.PrintName IS NOT NULL,LTRIM(B.PrintName),IIF(A1.PrintName IS NOT NULL,LTRIM(A1.PrintName),IIF(A2.PrintName IS NOT NULL,A2.PrintName,IIF(A3.PrintName IS NOT NULL,LTRIM(A3.PrintName),LTRIM(A4.PrintName))))) As AccountName,RIGHT(T.Type,1)+'O/'+LTRIM(T.Name)+'/JW/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "' As VchNo,T.Date As VchDate,I.PrintName As ItemName,'' As ChallanNo,'' As ChallanDate,T.EstQty01 As Quantity,SUM(IIF(LEFT(BOM,18) IN ('0310FI','0710XXXXXXXXXXXXFI','0110FI','0510XXXXXXXXXXXXFI','0000'),C.Quantity,0)) As ReceivedQty,SUM(ABS(IIF(LEFT(BOM,18) IN ('0410FI','0810XXXXXXXXXXXXFI','0210FI','0610XXXXXXXXXXXXFI','0000'),C.Quantity,0))) As IssuedQty,IIF(B.Code IS NOT NULL,LTRIM(B.Code),IIF(A1.Code IS NOT NULL,LTRIM(A1.Code),IIF(A2.Code IS NOT NULL,LTRIM(A2.Code),IIF(A3.Code IS NOT NULL,LTRIM(A3.Code),LTRIM(A4.Code))))) As BCode," & _
                        "'' As GRNNo,'' As Date,A.Name As Consignee,B.Name As Party,'' As Remarks,LTRIM(T.Name) As PO,I.Name As Book,SUM(C.Quantity) As Qty,0 As Rate,SUM(C.Amount),'' AS BOX,'' AS Freight,'' AS TYPE,'' AS MRNNo,RIGHT(T.Type,1) As OrderType,I.Code,IIF(LEFT(P.Type,4)='0110','Purchase',IIF(LEFT(P.Type,4)='0210','Purchase Return',IIF(LEFT(P.Type,4)='0510','Pur Challan IN',IIF(LEFT(P.Type,4)='0610','Pur Challan Out',IIF(LEFT(P.Type,4)='0310','Sales Return',IIF(LEFT(P.Type,4)='0410','Sales',IIF(LEFT(P.Type,4)='0710','Sales Challan IN',IIF(LEFT(P.Type,4)='0810','Sales Challan Out','Order Processing')))))))) As TypeRef,IIF(C.BOM IS NOT Null,C.BOM,'0000FI') As BOM,T.Code,'' AS RemarkC " & _
                        "FROM ((((((((BookPOParent T INNER JOIN BookMaster As I ON T.Book=I.Code) LEFT JOIN JobworkBVChild C ON C.Ref=T.Code) LEFT JOIN JobworkBVParent P ON P.Code=C.Code) LEFT JOIN AccountMaster As B ON P.Party=B.Code) LEFT JOIN AccountMaster As A ON P.Consignee=A.Code) LEFT JOIN AccountMaster As A1 ON T.Binder=A1.Code) LEFT JOIN AccountMaster As A2 ON T.BookPrinter=A2.Code) LEFT JOIN AccountMaster As A3 ON T.TitlePrinter=A3.Code) LEFT JOIN AccountMaster As A4 ON T.Laminator=A4.Code WHERE LEFT(IIF(C.BOM IS NOT Null,C.BOM,'0000FI'),18) IN " & IIf(Left(VchCode, 1) = "S", "('0110FI','0210FI','0510XXXXXXXXXXXXFI','0610XXXXXXXXXXXXFI','0000FI')", "('0310FI','0410FI','0710XXXXXXXXXXXXFI','0810XXXXXXXXXXXXFI','0000FI')") & " AND LEFT(T.Type,1)<>'O' AND RIGHT(T.Type,1)<>'" & Left(VchCode, 1) & "' AND LEFT(T.Code,1)<>'*' AND T.Date>='" & GetDate(MhDateInput1.Text) & "' AND T.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & _
                        IIf(Option3.Value, "T.DeliveredQuantityC+T.DeliveredQuantityB<T.EstQty01", IIf(Option1.Value, "T.DeliveredQuantityC+T.DeliveredQuantityB>=T.EstQty01", "1=1")) & " AND IIF(B.Code IS NOT NULL,LTRIM(B.Code),IIF(A1.Code IS NOT NULL,LTRIM(A1.Code),IIF(A2.Code IS NOT NULL,LTRIM(A2.Code),IIF(A3.Code IS NOT NULL,LTRIM(A3.Code),LTRIM(A4.Code))))) IN (" & SelectedItems(ListView1) & ") AND I.Code IN (" & SelectedItems(ListView2) & ") " & _
                        "Group BY IIF(B.PrintName IS NOT NULL,LTRIM(B.PrintName),IIF(A1.PrintName IS NOT NULL,LTRIM(A1.PrintName),IIF(A2.PrintName IS NOT NULL,A2.PrintName,IIF(A3.PrintName IS NOT NULL,LTRIM(A3.PrintName),LTRIM(A4.PrintName))))),RIGHT(T.Type,1)+'O/'+LTRIM(T.Name)+'/JW/21-22',T.Date,LTRIM(T.Name),P.TYPE,T.TYPE,T.Name,I.PrintName,T.EstQty01,C.BOM,IIF(B.Code IS NOT NULL,LTRIM(B.Code),IIF(A1.Code IS NOT NULL,LTRIM(A1.Code),IIF(A2.Code IS NOT NULL,LTRIM(A2.Code),IIF(A3.Code IS NOT NULL,LTRIM(A3.Code),LTRIM(A4.Code))))),A.Name,B.Name,I.Name,I.Code,T.Code " & _
                        "ORDER BY AccountName,T.Name"
        ElseIf Right(VchCodeType, 1) = 3 Or Right(VchCodeType, 1) = 4 Or Right(VchCodeType, 1) = 5 Then
        SQL = "SELECT DISTINCT (Select PrintName From AccountMaster Where Code=P.Party) AS AccountName,IIF(LEFT(BOM,6)IN ('1801FI'),'SO-T','PO-T')+'/'+LTRIM(P.Name)+'-FG/19-21' AS VchNo,P.Date AS VchDate,(Select PrintName From BookMaster Where Code=C.Item ) AS ItemName,P.ChallanNo,(P.ChallanDate),IIF(LEFT(BOM,6)IN ('1701FI','1801FI'),ABS(C.Quantity),0) AS Quantity,ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0) As Dispatched," & _
                  "(IIF(LEFT(BOM,6)IN ('1701FI','1801FI'),ABS(C.[Quantity]),0)-ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0)) As Balance,(P.Party) AS BCode,LTRIM(P.Name) AS GRNNo,P.Date,(Select Name From AccountMaster Where Code=P.Consignee) AS Consignee,(Select Name From AccountMaster Where Code=P.Party) AS Party,P.Remarks,LTRIM(P.Name) AS PO,(Select PrintName From BookMaster Where Code=C.Item ) AS Book,C.Quantity As Qty,C.Rate,C.Amount,P.BOX,P.Freight,P.TYPE, Right(P.Type,2)+'-'+LTRIM(P.Name) As MRNNo,Right(P.Type,2) As OrderType,C.Item,IIF(LEFT(P.Type,4)='0110','Purchase',IIF(LEFT(P.Type,4)='0210','Purchase Return',IIF(LEFT(P.Type,4)='0510','Pur Challan IN',IIF(LEFT(P.Type,4)='0610','Pur Challan Out',IIF(LEFT(P.Type,4)='0310','Sales Return',IIF(LEFT(P.Type,4)='0410','Sales',IIF(LEFT(P.Type,4)='0710','Sales Challan IN',IIF(LEFT(P.Type,4)='0810','Sales Challan Out','Order Processing')))))))) As TypeRef,C.BOM  As BOM,P.Code,P.Remarks AS RemarkC " & _
                  "FROM JobworkBVParent P INNER JOIN JobworkBVChild C ON P.Code=C.Code Left Join JobworkBVRef R ON R.VchCode=C.Code " & _
                  "WHERE LEFT((C.BOM),6) IN ('" & IIf(VchCode = "P", "1801FI", "1701FI") & "') AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & _
                  IIf(Option3.Value, "ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0)<ABS(C.Quantity)", IIf(Option1.Value, "ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0)>=ABS(C.Quantity)", IIf(Option2.Value, "IIf(Right(P.Type,1)='P',ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0),ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0))>=0", 1))) & "  " & _
                  "AND (P.Party) IN (" & SelectedItems(ListView1) & ") AND C.Item IN (" & SelectedItems(ListView2) & ") " & _
                  "ORDER BY '" & IIf(Right(VchCodeType, 1) = 3, "Code", IIf(Right(VchCodeType, 1) = 4, "AccountName", IIf(Right(VchCodeType, 1) = 5, "ItemName", "AccountName"))) & "'"
        ElseIf Right(VchCodeType, 1) = 6 Or Right(VchCodeType, 1) = 7 Or Right(VchCodeType, 1) = 8 Or Right(VchCodeType, 1) = 9 Or Right(VchCodeType, 1) = 0 Then
        SQL = "SELECT DISTINCT (Select PrintName From AccountMaster Where Code=P.Party) AS AccountName,IIF(LEFT(BOM,6)IN ('0110FI','0510FI','0101FI','0501FI','0310FI','0710FI','0301FI','0701FI'),'GRN','GDN')+'/'+LTRIM(P.Name)+'-FG/19-21' AS VchNo,P.Date AS VchDate,(Select PrintName From BookMaster Where Code=C.Item ) AS ItemName,P.ChallanNo,(P.ChallanDate),IIF(LEFT(BOM,6)IN ('0110FI','0510FI','0101FI','0501FI','0310FI','0710FI','0301FI','0701FI','0210FI','0610FI','0201FI','0601FI','0410FI','0810FI','0401FI','0801FI'),ABS(C.Quantity),0) AS Quantity,IIF(LEFT(BOM,6)IN ('0110FI','0510FI','0101FI','0501FI','0310FI','0710FI','0301FI','0701FI'),ABS(C.Quantity),0) AS Received," & _
                  "IIF(LEFT(BOM,6)IN ('0210FI','0610FI','0201FI','0601FI','0410FI','0810FI','0401FI','0801FI'),ABS(C.Quantity),0) AS Issued,(P.Party) AS BCode,LTRIM(P.Name) AS GRNNo,P.Date,(Select Name From AccountMaster Where Code=P.Consignee) AS Consignee,(Select Name From AccountMaster Where Code=P.Party) AS Party,P.Remarks,LTRIM(P.Name) AS PO,(Select PrintName From BookMaster Where Code=C.Item ) AS Book,C.Quantity As Qty,C.Rate,C.Amount,P.BOX,P.Freight,P.TYPE, Right(P.Type,2)+'-'+LTRIM(P.Name) As MRNNo,Right(P.Type,2) As OrderType,C.Item,IIF(LEFT(P.Type,4)='0110','Purchase',IIF(LEFT(P.Type,4)='0210','Purchase Return',IIF(LEFT(P.Type,4)='0510','Pur Challan IN',IIF(LEFT(P.Type,4)='0610','Pur Challan Out',IIF(LEFT(P.Type,4)='0310','Sales Return',IIF(LEFT(P.Type,4)='0410','Sales',IIF(LEFT(P.Type,4)='0710','Sales Challan IN',IIF(LEFT(P.Type,4)='0810','Sales Challan Out','Order Processing')))))))) As TypeRef,C.BOM  As BOM,P.Code,P.Remarks AS RemarkC " & _
                  "FROM JobworkBVParent P INNER JOIN JobworkBVChild C ON P.Code=C.Code Left Join JobworkBVRef R ON R.VchCode=C.Code " & _
                  "WHERE LEFT(BOM,6)IN ('0110FI','0510FI','0101FI','0501FI','0310FI','0710FI','0301FI','0701FI','0210FI','0610FI','0201FI','0601FI','0410FI','0810FI','0401FI','0801FI') AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & _
                  IIf(Option3.Value, "LEFT(BOM,6)IN ('0210FI','0610FI','0201FI','0601FI','0410FI','0810FI','0401FI','0801FI')", IIf(Option1.Value, "LEFT(BOM,6)IN ('0110FI','0510FI','0101FI','0501FI','0310FI','0710FI','0301FI','0701FI')", IIf(Option2.Value, "LEFT(BOM,6)IN ('0110FI','0510FI','0101FI','0501FI','0310FI','0710FI','0301FI','0701FI','0210FI','0610FI','0201FI','0601FI','0410FI','0810FI','0401FI','0801FI')", 1))) & "  " & _
                  "AND (P.Party) IN (" & SelectedItems(ListView1) & ") AND C.Item IN (" & SelectedItems(ListView2) & ") " & _
                  "ORDER BY '" & IIf(Right(VchCodeType, 1) = 9, "Code", IIf(Right(VchCodeType, 1) = 8, "ItemName", IIf(Right(VchCodeType, 1) = 7, "AccountName", IIf(Right(VchCodeType, 1) = 6, "ItemName", "VchDate")))) & "'"
        End If
        If DatabaseType = "MS SQL" Then SQL = Replace(SQL, "#", "'")
        Screen.MousePointer = vbNormal
        If rstOrderProcessing.State = adStateOpen Then rstOrderProcessing.Close
        rstOrderProcessing.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
        If rstOrderProcessing.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
        rstOrderProcessing.MoveFirst
        rptOrderProcessing.Database.SetDataSource rstOrderProcessing, 3, 1
    If Right(VchCodeType, 1) = 1 Then
'        If MsgBox("Do You wants to Print Order Remarks?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Quit !") = vbYes Then
        rptOrderProcessing.Section9.Suppress = False 'HeaderTop
        rptOrderProcessing.Section1.Suppress = False 'Header_Detailed
        rptOrderProcessing.Section16.Suppress = True 'Header_Summaried
        rptOrderProcessing.Section18.Suppress = True 'Header_SaleOrder
        rptOrderProcessing.Section23.Suppress = True 'Header_GRN_GDN
        
        rptOrderProcessing.Section6.Suppress = False 'Group1 Account
        rptOrderProcessing.Section10.Suppress = False 'Group2 VchNo
        rptOrderProcessing.Section8.Suppress = True 'Group3 Qty.
        rptOrderProcessing.Section26.Suppress = True 'Group4 Item
        
        rptOrderProcessing.Section5.Suppress = True 'Details_Summaried_Jw
        rptOrderProcessing.Section15.Suppress = False 'Details_Detailed_Jw
        rptOrderProcessing.Section19.Suppress = True 'Details_SaleOrder_T
        rptOrderProcessing.Section24.Suppress = True 'Details_GRN_GDN
        
        rptOrderProcessing.Section27.Suppress = True 'Footer_Item
        rptOrderProcessing.Section12.Suppress = True 'Footer_Qty
        rptOrderProcessing.Section11.Suppress = False 'Footer_Vch._Order-Total
        
        rptOrderProcessing.Section7.Suppress = False 'Party-wise-Total_Jobwork
        rptOrderProcessing.Section20.Suppress = True 'Party-wise-Total_SaleOrder
        rptOrderProcessing.Section22.Suppress = True 'Party-wise-Total_GRN_GDN
        
        rptOrderProcessing.Section4.Suppress = False 'Grand-Total
        rptOrderProcessing.Section21.Suppress = True 'Grand-Total
        rptOrderProcessing.Section25.Suppress = True 'Grand-Total
        
    ElseIf Right(VchCodeType, 1) = 2 Then
        rptOrderProcessing.Section9.Suppress = True 'HeaderTop
        rptOrderProcessing.Section1.Suppress = True 'Header_Detailed
        rptOrderProcessing.Section16.Suppress = False 'Header_Summaried
        rptOrderProcessing.Section18.Suppress = True 'Header_SaleOrder
        rptOrderProcessing.Section23.Suppress = True 'Header_GRN_GDN
        
        rptOrderProcessing.Section6.Suppress = False 'Group1 Account
        rptOrderProcessing.Section10.Suppress = True 'Group2 VchNo
        rptOrderProcessing.Section8.Suppress = True 'Group3 Qty.
        rptOrderProcessing.Section26.Suppress = True 'Group4 Item
        
        rptOrderProcessing.Section5.Suppress = False 'Details_Summaried_Jw
        rptOrderProcessing.Section15.Suppress = True 'Details_Detailed_Jw
        rptOrderProcessing.Section19.Suppress = True 'Details_SaleOrder_T
        rptOrderProcessing.Section24.Suppress = True 'Details_GRN_GDN
        
        rptOrderProcessing.Section27.Suppress = True 'Footer_Item
        rptOrderProcessing.Section12.Suppress = True 'Footer_Qty
        rptOrderProcessing.Section11.Suppress = True 'Footer_Vch._Order-Total
        
        rptOrderProcessing.Section7.Suppress = False 'Party-wise-Total_Jobwork
        rptOrderProcessing.Section20.Suppress = True 'Party-wise-Total_SaleOrder
        rptOrderProcessing.Section22.Suppress = True 'Party-wise-Total_Summaried_Jw
        
        rptOrderProcessing.Section4.Suppress = False 'Grand-Total
        rptOrderProcessing.Section21.Suppress = True 'Grand-Total
        rptOrderProcessing.Section25.Suppress = True 'Grand-Total
    ElseIf Right(VchCodeType, 1) = 4 Then
        
        rptOrderProcessing.Section26.Suppress = True 'Group3 Item
        rptOrderProcessing.Section4.Suppress = True
        rptOrderProcessing.Section7.Suppress = True
        rptOrderProcessing.Section9.Suppress = True
        rptOrderProcessing.Section1.Suppress = True
        rptOrderProcessing.Section16.Suppress = True
        rptOrderProcessing.Section18.Suppress = False
        rptOrderProcessing.Section20.Suppress = False
        rptOrderProcessing.Section10.Suppress = True
        rptOrderProcessing.Section11.Suppress = True
        rptOrderProcessing.Section19.Suppress = False
        rptOrderProcessing.Section5.Suppress = True
        rptOrderProcessing.Section15.Suppress = True
        rptOrderProcessing.Section22.Suppress = True
        rptOrderProcessing.Section23.Suppress = True 'Header
        rptOrderProcessing.Section24.Suppress = True 'Details
        rptOrderProcessing.Section25.Suppress = True 'Grand-Total
        rptOrderProcessing.Section27.Suppress = True
        rptOrderProcessing.Text34.SetText IIf(Left(VchCodeType, 1) = 1, "Qty. IN", "Qty. OUT")
    ElseIf Right(VchCodeType, 1) = 3 Or Right(VchCodeType, 1) = 5 Then
        rptOrderProcessing.Section4.Suppress = True
        rptOrderProcessing.Section6.Suppress = True
        rptOrderProcessing.Section26.Suppress = True 'Group3 Item
        rptOrderProcessing.Section7.Suppress = True
        rptOrderProcessing.Section9.Suppress = True
        rptOrderProcessing.Section1.Suppress = True
        rptOrderProcessing.Section16.Suppress = True
        rptOrderProcessing.Section18.Suppress = False
        rptOrderProcessing.Section20.Suppress = True
        rptOrderProcessing.Section10.Suppress = True
        rptOrderProcessing.Section11.Suppress = True
        rptOrderProcessing.Section19.Suppress = False
        rptOrderProcessing.Section5.Suppress = True
        rptOrderProcessing.Section15.Suppress = True
        rptOrderProcessing.Section22.Suppress = True
        rptOrderProcessing.Section23.Suppress = True 'Header
        rptOrderProcessing.Section24.Suppress = True 'Details
        rptOrderProcessing.Section27.Suppress = True
        rptOrderProcessing.Section25.Suppress = True 'Grand-Total
        rptOrderProcessing.Text34.SetText IIf(Left(VchCodeType, 1) = 1, "Qty. IN", "Qty. OUT")
    
    ElseIf Right(VchCodeType, 1) = 6 Or Right(VchCodeType, 1) = 8 Or Right(VchCodeType, 1) = 9 Or Right(VchCodeType, 1) = 0 Then
        rptOrderProcessing.Section9.Suppress = True 'Header
        rptOrderProcessing.Section1.Suppress = True 'Header
        rptOrderProcessing.Section16.Suppress = True 'Header
        rptOrderProcessing.Section18.Suppress = True 'Header
        rptOrderProcessing.Section23.Suppress = False 'Header
        rptOrderProcessing.Section6.Suppress = True 'Group1 Account
        rptOrderProcessing.Section10.Suppress = True 'Group2 Vch.
        rptOrderProcessing.Section8.Suppress = True 'Group3 Qty.
        rptOrderProcessing.Section26.Suppress = True 'Group3 Item
        rptOrderProcessing.Section5.Suppress = True 'Details
        rptOrderProcessing.Section15.Suppress = True 'Details
        rptOrderProcessing.Section19.Suppress = True 'Details
        rptOrderProcessing.Section24.Suppress = False 'Details
        rptOrderProcessing.Section26.Suppress = True 'Details
        rptOrderProcessing.Section27.Suppress = True 'Details
        rptOrderProcessing.Section11.Suppress = True 'Order-Total
        rptOrderProcessing.Section7.Suppress = True 'Party-wise-Total
        rptOrderProcessing.Section20.Suppress = True 'Party-wise-Total
        rptOrderProcessing.Section22.Suppress = True 'Party-wise-Total
        rptOrderProcessing.Section4.Suppress = True 'Grand-Total
        rptOrderProcessing.Section21.Suppress = True 'Grand-Total
        rptOrderProcessing.Section25.Suppress = False 'Grand-Total
    ElseIf Right(VchCodeType, 1) = 7 Then
        rptOrderProcessing.Section9.Suppress = True 'Header
        rptOrderProcessing.Section1.Suppress = True 'Header
        rptOrderProcessing.Section16.Suppress = True 'Header
        rptOrderProcessing.Section18.Suppress = True 'Header
        rptOrderProcessing.Section23.Suppress = False 'Header
        rptOrderProcessing.Section6.Suppress = False 'Group1 Account
        rptOrderProcessing.Section10.Suppress = True 'Group2 Vch.
        rptOrderProcessing.Section8.Suppress = True 'Group3 Qty.
        rptOrderProcessing.Section26.Suppress = True 'Group3 Item
        rptOrderProcessing.Section5.Suppress = True 'Details
        rptOrderProcessing.Section15.Suppress = True 'Details
        rptOrderProcessing.Section19.Suppress = True 'Details
        rptOrderProcessing.Section24.Suppress = False 'Details
        rptOrderProcessing.Section26.Suppress = True 'Details
        rptOrderProcessing.Section27.Suppress = True 'Details
        rptOrderProcessing.Section11.Suppress = True 'Order-Total
        rptOrderProcessing.Section7.Suppress = True 'Party-wise-Total
        rptOrderProcessing.Section20.Suppress = True 'Party-wise-Total
        rptOrderProcessing.Section22.Suppress = False 'Party-wise-Total
        rptOrderProcessing.Section4.Suppress = True 'Grand-Total
        rptOrderProcessing.Section21.Suppress = True 'Grand-Total
        rptOrderProcessing.Section25.Suppress = False 'Grand-Total
'        rptOrderProcessing.Section9.Suppress = True 'Header
'        rptOrderProcessing.Section1.Suppress = True 'Header
'        rptOrderProcessing.Section16.Suppress = True 'Header
'        rptOrderProcessing.Section18.Suppress = True 'Header
'        rptOrderProcessing.Section23.Suppress = False 'Header
'        rptOrderProcessing.Section6.Suppress = False 'Group1 Account
'
'        rptOrderProcessing.Section10.Suppress = True 'Group2 Vch.
'        rptOrderProcessing.Section8.Suppress = True 'Group3 Qty.
'        rptOrderProcessing.Section26.Suppress = False 'Group3 Item
'        rptOrderProcessing.Section5.Suppress = True 'Details
'        rptOrderProcessing.Section15.Suppress = True 'Details
'        rptOrderProcessing.Section19.Suppress = True 'Details
'        rptOrderProcessing.Section24.Suppress = True 'Details
'        rptOrderProcessing.Section27.Suppress = True 'Details
'        rptOrderProcessing.Section11.Suppress = True 'Order-Total
'        rptOrderProcessing.Section7.Suppress = True 'Party-wise-Total
'        rptOrderProcessing.Section20.Suppress = True 'Party-wise-Total
'        rptOrderProcessing.Section22.Suppress = True 'Party-wise-Total
'        rptOrderProcessing.Section4.Suppress = True 'Grand-Total
'        rptOrderProcessing.Section21.Suppress = True 'Grand-Total
'        rptOrderProcessing.Section25.Suppress = False 'Grand-Total

    End If

    rptOrderProcessing.DiscardSavedData
If Right(VchCodeType, 1) = 1 Or Right(VchCodeType, 1) = 2 Then
    rptOrderProcessing.PaperOrientation = crLandscape
ElseIf Right(VchCodeType, 1) = 3 Or Right(VchCodeType, 1) = 4 Or Right(VchCodeType, 1) = 5 Then
    rptOrderProcessing.PaperOrientation = crLandscape
ElseIf Right(VchCodeType, 1) = 6 Or Right(VchCodeType, 1) = 7 Or Right(VchCodeType, 1) = 8 Or Right(VchCodeType, 1) = 9 Or Right(VchCodeType, 1) = 0 Then
    rptOrderProcessing.PaperOrientation = crLandscape 'crPortrait
End If
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptOrderProcessing: FrmReportViewer.Show vbModal
    ElseIf OutputTo = "P" Then
        rptOrderProcessing.PaperSource = crPRBinAuto
        rptOrderProcessing.PrintOut
    Else
        If iCount >= 1 Then
            Dim oOutlookMsg As Outlook.MailItem, FileName As String
            Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
            With oOutlookMsg
                '.To = rstPaperDebitNote.Fields("EMail").Value
                .Subject = IIf(Option3.Value, "Pending  ", IIf(Option1.Value, "Close ", "")) & "Sale Orders "
                .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith " & IIf(Option3.Value, "Pending ", IIf(Option1.Value, "Close ", "")) & "Sale Order from " + Format(GetDate(MhDateInput1.Text), "dd-MMM-yyyy") + " to " + Format(GetDate(MhDateInput2.Text), "dd-MMM-yyyy") & " for doing the needful at your end.<Br><b>Kindly do acknowledge the receipt of the mail</b>.<Br><Br>Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
                rptOrderProcessing.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptOrderProcessing.ExportOptions.DestinationType = crEDTDiskFile
                FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
                rptOrderProcessing.ExportOptions.DiskFileName = FileName
                rptOrderProcessing.Export False
                .Attachments.Add (FileName)
                .Importance = olImportanceHigh
                .ReadReceiptRequested = True
                If CheckEmpty(.To, False) Then .Display Else .Send
            End With
            Set oOutlookMsg = Nothing
        End If
    End If
    Set rptOrderProcessing = Nothing
    On Error GoTo 0
End Sub
