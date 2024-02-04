VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOutsourceItemSupplierRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outsource Item Purchase Order Status Register"
   ClientHeight    =   7050
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10020
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
   ScaleHeight     =   7050
   ScaleWidth      =   10020
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10020
      _ExtentX        =   17674
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
            Picture         =   "OutsourceItemSupplierRegister.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OutsourceItemSupplierRegister.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OutsourceItemSupplierRegister.frx":0658
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6675
      Left            =   45
      TabIndex        =   6
      Top             =   345
      Width           =   9930
      _Version        =   65536
      _ExtentX        =   17515
      _ExtentY        =   11774
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
      Picture         =   "OutsourceItemSupplierRegister.frx":076C
      Begin VB.OptionButton Option2 
         Caption         =   "Summarised"
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
         Left            =   8400
         TabIndex        =   2
         Top             =   10
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detailed"
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
         Left            =   6840
         TabIndex        =   1
         Top             =   10
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Without Nil"
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
         Left            =   5040
         TabIndex        =   0
         Top             =   53
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6360
         Left            =   0
         TabIndex        =   3
         Top             =   315
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   11218
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
         TabIndex        =   7
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
         Caption         =   " &From"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "OutsourceItemSupplierRegister.frx":0788
         Picture         =   "OutsourceItemSupplierRegister.frx":07A4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   2520
         TabIndex        =   8
         Top             =   0
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   8421376
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
         Picture         =   "OutsourceItemSupplierRegister.frx":07C0
         Picture         =   "OutsourceItemSupplierRegister.frx":07DC
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6360
         Left            =   4950
         TabIndex        =   4
         Top             =   315
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   11218
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
         Left            =   1080
         TabIndex        =   9
         Top             =   0
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "OutsourceItemSupplierRegister.frx":07F8
         Caption         =   "OutsourceItemSupplierRegister.frx":0910
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "OutsourceItemSupplierRegister.frx":097C
         Keys            =   "OutsourceItemSupplierRegister.frx":099A
         Spin            =   "OutsourceItemSupplierRegister.frx":09F8
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
         Left            =   3510
         TabIndex        =   10
         Top             =   0
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "OutsourceItemSupplierRegister.frx":0A20
         Caption         =   "OutsourceItemSupplierRegister.frx":0B38
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "OutsourceItemSupplierRegister.frx":0BA4
         Keys            =   "OutsourceItemSupplierRegister.frx":0BC2
         Spin            =   "OutsourceItemSupplierRegister.frx":0C20
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
Attribute VB_Name = "FrmOutsourceItemSupplierRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstOutsourceItemSupplierRegister As New ADODB.Recordset
Dim rstOutsourceItemList As New ADODB.Recordset
Dim rstSupplierList As New ADODB.Recordset
Dim OutputTo As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "Select PrintName FROM CompanyMaster Where FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    rstSupplierList.Open "Select Name, Code From AccountMaster Order By Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstSupplierList.ActiveConnection = Nothing
    rstOutsourceItemList.Open "Select Name, Code From OutsourceItemMaster Order By Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstOutsourceItemList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Suppliers...", rstSupplierList)
    Call FillList(ListView2, "List of Outsource Items...", rstOutsourceItemList)
    Option1.Value = True
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    If Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd") Then
        MhDateInput2.Text = Format(FinancialYearTo, "dd-mm-yyyy")
    Else
        MhDateInput2.Text = Format(Date, "dd-mm-yyyy")
    End If
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
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
    If UnloadMode = 0 Then
        CloseForm Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstSupplierList)
    Call CloseRecordset(rstOutsourceItemSupplierRegister)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput2.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
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
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = True
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = False
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    If Button.Index = 1 Then
        OutputTo = "S"
        PrintOutsourceItemSupplierRegister
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintOutsourceItemSupplierRegister
    ElseIf Button.Index = 3 Then
        CloseForm Me
    End If
End Sub
Private Sub PrintOutsourceItemSupplierRegister()
    Dim CRXParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim CRXParamDef As CRAXDRT.ParameterFieldDefinition
    Dim SelectedOutsourceItems As String
    Dim SelectedSuppliers As String
    Dim SQL As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptOutsourceItemSupplierRegister.Text11.SetText "Outsource Item Purchase Order Status Register (" & IIf(Option1.Value, "Detailed", "Summarised") & ")"
    rptOutsourceItemSupplierRegister.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptOutsourceItemSupplierRegister.Text13.SetText "From [" + Format(MhDateInput1.Text, "dd-mm-yyyy") + "] To [" + Format(MhDateInput2.Text, "dd-mm-yyyy") + "]"
    If rstOutsourceItemSupplierRegister.State = adStateOpen Then
        rstOutsourceItemSupplierRegister.Close
    End If
    SelectedOutsourceItems = SelectedItems(ListView2)
    SelectedSuppliers = SelectedItems(ListView1)
    'rstOutsourceItemSupplierRegister.Open
                                             SQL = "Select 'Supplier Name : '+LTRIM(AccountMaster.PrintName) As SupplierName,'Item Name : '+LTRIM(OutsourceItemMaster.PrintName) As OutsourceItemName,'' As GodownName,LTRIM(OutsourceItemPOParent.Name) As VchNo,'' As OrderNo,OutsourceItemPOParent.Date As VchDate,'PO' As VchType,Quantity From AccountMaster,OutsourceItemMaster,OutsourceItemPOParent,OutsourceItemPOChild Where OutsourceItemMaster.Code=OutsourceItemPOChild.OutsourceItem And OutsourceItemPOParent.Code=OutsourceItemPOChild.Code And OutsourceItemPOParent.Supplier=AccountMaster.Code And OutsourceItemMaster.Code In (" & SelectedOutsourceItems & ")  And AccountMaster.Code In (" & SelectedSuppliers & ")  And OutsourceItemPOParent.Date>=#" & GetDate(MhDateInput1.Text) & "# And OutsourceItemPOParent.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION " & _
                                                       "Select 'Supplier Name : '+LTRIM(AccountMaster.PrintName) As SupplierName,'Item Name : '+LTRIM(OutsourceItemMaster.PrintName) As OutsourceItemName,(Select LTRIM(PrintName) From AccountMaster Where Code=MaterialIOChild.Godown) As GodownName,LTRIM(OutsourceItemPOParent.Name) As VchNo,LTRIM(MaterialIOParent.Name) As OrderNo,MaterialIOParent.Date As VchDate,'PS' As VchType,Quantity From AccountMaster,OutsourceItemMaster,MaterialIOParent,MaterialIOChild,OutsourceItemPOParent Where OutsourceItemMaster.Code=MaterialIOChild.Item And MaterialIOChild.Category='1' AND MaterialIOParent.Code=MaterialIOChild.Code And MaterialIOChild.Ref=OutsourceItemPOParent.Code And MaterialIOParent.Source=AccountMaster.Code And OutsourceItemMaster.Code In (" & SelectedOutsourceItems & ")  And AccountMaster.Code In (" & SelectedSuppliers & ")  And MaterialIOParent.Date>=#" & GetDate(MhDateInput1.Text) & "# And " & _
                                                       "MaterialIOParent.Date<=#" & GetDate(MhDateInput2.Text) & "# Order By SupplierName,OutsourceItemName,VchNo,OrderNo" ', cnDatabase, adOpenKeyset, adLockOptimistic
    If DatabaseType = "MS SQL" Then SQL = Replace(SQL, "#", "'")
    If rstOutsourceItemSupplierRegister.State = adStateOpen Then rstOutsourceItemSupplierRegister.Close
    rstOutsourceItemSupplierRegister.Open SQL, cnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    If rstOutsourceItemSupplierRegister.RecordCount = 0 Then 'Screen.MousePointer = vbNormal: Exit Sub
        On Error GoTo 0
        Exit Sub
    End If
    rptOutsourceItemSupplierRegister.Database.SetDataSource rstOutsourceItemSupplierRegister, 3, 1
    rptOutsourceItemSupplierRegister.DiscardSavedData
    Set CRXParamDefs = rptOutsourceItemSupplierRegister.ParameterFields
    For Each CRXParamDef In CRXParamDefs
        If CRXParamDef.ParameterFieldName = "PF1" Then
            CRXParamDef.SetCurrentValue (IIf(Check1.Value, 0, 0.1))
        ElseIf CRXParamDef.ParameterFieldName = "PF2" Then
            CRXParamDef.SetCurrentValue (IIf(Option1.Value, "D", "S"))
        End If
    Next
    rptOutsourceItemSupplierRegister.EnableParameterPrompting = False
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptOutsourceItemSupplierRegister
        FrmReportViewer.Show vbModal
    Else
        rptOutsourceItemSupplierRegister.PaperSource = crPRBinAuto
        rptOutsourceItemSupplierRegister.PrintOut
    End If
    Set rptOutsourceItemSupplierRegister = Nothing
    On Error GoTo 0
End Sub
