VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMaterialStockRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Stock Register"
   ClientHeight    =   8235
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13620
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
   ScaleHeight     =   8235
   ScaleWidth      =   13620
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13620
      _ExtentX        =   24024
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
            Picture         =   "MaterialStockRegister.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaterialStockRegister.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaterialStockRegister.frx":0658
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7860
      Left            =   45
      TabIndex        =   7
      Top             =   345
      Width           =   13530
      _Version        =   65536
      _ExtentX        =   23865
      _ExtentY        =   13864
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
      Picture         =   "MaterialStockRegister.frx":076C
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
         Left            =   7080
         TabIndex        =   1
         Top             =   53
         Width           =   1455
      End
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
         Left            =   11100
         TabIndex        =   3
         Top             =   10
         Width           =   1350
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
         Left            =   9120
         TabIndex        =   2
         Top             =   10
         Value           =   -1  'True
         Width           =   1425
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   0
         TabIndex        =   8
         Top             =   0
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
         Caption         =   " &From"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "MaterialStockRegister.frx":0788
         Picture         =   "MaterialStockRegister.frx":07A4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   3360
         TabIndex        =   9
         Top             =   0
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
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
         Picture         =   "MaterialStockRegister.frx":07C0
         Picture         =   "MaterialStockRegister.frx":07DC
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3600
         Left            =   6630
         TabIndex        =   4
         Top             =   315
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   6350
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         Height          =   3960
         Left            =   0
         TabIndex        =   5
         Top             =   3900
         Width           =   13530
         _ExtentX        =   23865
         _ExtentY        =   6985
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         Left            =   1440
         TabIndex        =   0
         Top             =   0
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   582
         Calendar        =   "MaterialStockRegister.frx":07F8
         Caption         =   "MaterialStockRegister.frx":0910
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "MaterialStockRegister.frx":097C
         Keys            =   "MaterialStockRegister.frx":099A
         Spin            =   "MaterialStockRegister.frx":09F8
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
         Height          =   3600
         Left            =   0
         TabIndex        =   10
         Top             =   315
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   6350
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         Left            =   4830
         TabIndex        =   11
         Top             =   0
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Calendar        =   "MaterialStockRegister.frx":0A20
         Caption         =   "MaterialStockRegister.frx":0B38
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "MaterialStockRegister.frx":0BA4
         Keys            =   "MaterialStockRegister.frx":0BC2
         Spin            =   "MaterialStockRegister.frx":0C20
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
Attribute VB_Name = "FrmMaterialStockRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstMaterialStockRegister As New ADODB.Recordset
Dim rstBoardList As New ADODB.Recordset
Dim rstBookList As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim OutputTo As String
Public ReportType As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    If ReportType = "1" Then
        Me.Caption = "Bill of Materials [BOM] Stock Register [Godownwise/Itemwise/BOM Itemwise]"
    Else
        Me.Caption = "Bill of Materials [BOM] Stock Register [Godownwise/Itemwise]"
    End If
    rstCompanyMaster.Open "Select PrintName FROM CompanyMaster Where FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountList.Open "Select Name As Col0, Code From AccountMaster Order By Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    Call FillList(ListView3, "List of Godowns...", rstAccountList)
    If ReportType = "1" Then
        rstBoardList.Open "Select Name,Code From GeneralMaster Where Type = '5' AND Code='000000' Order by Name", cnDatabase, adOpenKeyset, adLockReadOnly
        rstBoardList.ActiveConnection = Nothing
        Call FillList(ListView1, IIf(ReportType = "1", "List of Boards...", "List of Item Types"), rstBoardList)
        Call BookSelection(True)
        ListView2.MultiSelect = False
    Else
        rstBoardList.Open "Select Name,Code From GeneralMaster Where Type = '0' Order by Name", cnDatabase, adOpenKeyset, adLockOptimistic
        rstBoardList.ActiveConnection = Nothing
        rstBoardList.AddNew
        rstBoardList.Fields("Name").Value = "Bill of Materials [BOM]"
        rstBoardList.Fields("Code").Value = "000001"
        rstBoardList.Update
        rstBoardList.AddNew
        rstBoardList.Fields("Name").Value = "Bill of Materials [FG]"
        rstBoardList.Fields("Code").Value = "000003"
        rstBoardList.Update
        rstBoardList.AddNew
        rstBoardList.Fields("Name").Value = "Bill of Materials [UFG] Items"
        rstBoardList.Fields("Code").Value = "000004"
        rstBoardList.Update
        rstBoardList.AddNew
        rstBoardList.Fields("Name").Value = "Bill of Materials [UFG] Elements"
        rstBoardList.Fields("Code").Value = "000005"
        rstBoardList.Update
        Call FillList(ListView1, IIf(ReportType = "1", "List of Bill of Materials [BOM]...", "List of Bill of Materials [BOM] Types"), rstBoardList)
        ListView1.ListItems(1).Selected = True
        Call BookSelection(False)
        ListView1.MultiSelect = False
    End If
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
    Call CloseForm(Me)
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
        Call CloseForm(Me)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBoardList)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstMaterialStockRegister)
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
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call BookSelection(False)
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Selected = True
        Next i
        If ReportType = "1" Then
            Call BookSelection(True)
        Else
            Call BookSelection(False)
        End If
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Selected = False
        Next i
        If ReportType = "2" Then
            ListView1.ListItems(4).Selected = True
        End If
        Call BookSelection(False)
    End If
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Selected = True
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Selected = False
        Next i
    End If
End Sub
Private Sub ListView3_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Selected = True
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Selected = False
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    If Button.Index = 1 Then
        OutputTo = "S"
        PrintMaterialStockRegister
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintMaterialStockRegister
    ElseIf Button.Index = 3 Then
        Call CloseForm(Me)
    End If
End Sub
Private Sub BookSelection(ByVal SelectAll As Boolean)
    If rstBookList.State = adStateOpen Then
        rstBookList.Close
    End If
    If ReportType = "1" Then
        rstBookList.Open "Select Name, Code From BookMaster " & IIf(SelectAll, "", "") & " Order By Name", cnDatabase, adOpenKeyset, adLockReadOnly
        'rstBookList.Open "Select Name, Code From BookMaster " & IIf(SelectAll, "BookMaster WHERE Board='000000'", "Where [Group] In (" & SelectedItems(ListView1, False) & ")") & " Order By Name", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstBookList.Open "Select " & IIf(Val(ListView1.SelectedItem.SubItems(1)) = 1, "Name, Code FROM OutsourceItemMaster", IIf(Val(ListView1.SelectedItem.SubItems(1)) = 3, "Name, Code FROM BookMaster Where Type='F'", IIf(Val(ListView1.SelectedItem.SubItems(1)) = 4, "Name, Code FROM BookMaster WHERE Type='R'", "I.Name+' ['+(select Name From ElementMaster Where Code=Element)+']' AS NAme ,C.Code+Element From BookMaster I INNER JOIN BookChild06 C ON C.Code=I.Code Where I.Type='F'"))) & " ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
        'rstBookList.Open "Select Name, Code FROM " & IIf(Val(ListView1.SelectedItem.SubItems(1)) = 1, "OutsourceItemMaster", IIf(Val(ListView1.SelectedItem.SubItems(1)) = 3, "BookMaster WHERE Board='000000'", IIf(Val(ListView1.SelectedItem.SubItems(1)) = 4, "BookMaster WHERE Type='R'", "BookMaster WHERE Board<>'000000' AND Type='F'"))) & " ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    rstBookList.ActiveConnection = Nothing
    ListView2.ListItems.Clear
    Call FillList(ListView2, "List of [BOM] Items...", rstBookList)
End Sub
Private Sub PrintMaterialStockRegister()
    Dim CRXParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim CRXParamDef As CRAXDRT.ParameterFieldDefinition
    Dim OutsourceItemQuantity As String
    Dim FreshBookQuantity As String
    Dim RepairBookQuantity As String
    Dim TitleQuantity As String
    Dim SelectedBoards As String
    Dim SelectedBooks As String
    Dim SelectedAccounts As String
    Dim SQL As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptMaterialStockRegister.Text11.SetText "UFG Stock Register (" & IIf(Option1.Value, "Detailed", "Summarised") & ")"
    rptMaterialStockRegister.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptMaterialStockRegister.Text13.SetText "From [" + Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy") + "]"
    If rstMaterialStockRegister.State = adStateOpen Then
        rstMaterialStockRegister.Close
    End If
    If ReportType = "1" Then
        rptMaterialStockRegister.Text11.Width = IIf(Option1.Value, 10800, 15780)
        rptMaterialStockRegister.Text12.Width = IIf(Option1.Value, 10800, 15780)
        rptMaterialStockRegister.Text13.Width = IIf(Option1.Value, 10800, 15780)
        rptMaterialStockRegister.Text9.Left = IIf(Option1.Value, 8880, 13860)
        rptMaterialStockRegister.Field17.Left = IIf(Option1.Value, 9840, 14820)
        rptMaterialStockRegister.Line5.Right = IIf(Option1.Value, 10800, 15780)
        SelectedBoards = SelectedItems(ListView1, False)
        SelectedBooks = SelectedItems(ListView2, False)
        SelectedAccounts = SelectedItems(ListView3, False)
        OutsourceItemQuantity = "(SELECT ISNULL(SUM(OpBal),0) FROM AccountChild0801 WHERE Category=C.Category AND Item=C.Item AND Code=A.Code)+" & _
                                                  "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                  "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                  "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                  "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                  "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                  "(SELECT ISNULL(SUM(Round(I.TotalConsumption,0)),0) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND Category=C.Category AND Item=C.Item AND I.Vendor=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
        FreshBookQuantity = "(SELECT ISNULL(SUM(OpBal),0) FROM AccountChild0801 WHERE Category=C.Category AND Item=C.Item AND Code=A.Code)+" & _
                                           "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                           "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                           "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                           "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                           "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                           "(SELECT ISNULL(SUM(Round(I.TotalConsumption,0)),0) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND Category=C.Category AND Item=C.Item AND I.Vendor=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
        RepairBookQuantity = "(SELECT ISNULL(SUM(OpBal),0) FROM AccountChild0801 WHERE Category='4' AND Item=O.Code AND Code=A.Code)+" & _
                                            "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                            "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                            "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                            "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                            "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                            "(SELECT ISNULL(SUM(Round(I.TotalConsumption,0)),0) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND Category='4' AND Item=O.Code AND I.Vendor=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
        TitleQuantity = "(SELECT ISNULL(SUM(OpBal),0,) FROM AccountChild0801 WHERE Category='5' AND Item=O.Code AND Code=A.Code)+" & _
                                 "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                 "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                 "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                 "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                 "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                 "(SELECT ISNULL(SUM(Round(I.TotalConsumption,0)),0)  FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND Category='5' AND Item=O.Code AND I.Vendor=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
        SQL = SQL + "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & OutsourceItemQuantity & " As Quantity,'Board Name : '+Trim(G.PrintName) As BoardName," & _
                            "'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A WHERE B.Code=C.Code AND B.Board=G.Code AND C.Item=O.Code AND C.Category='1' AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") AND (" & OutsourceItemQuantity & ") <> 0 UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND I.Godown=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "Trim(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.Account=A.Code AND I.Quantity>=0 AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "Trim(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.Account=A.Code AND I.Quantity<0 AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.AccountFrom=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.AccountTo=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,Round(I.TotalConsumption,0),'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category=C.Category AND I.Item=C.Item) AND I.Vendor=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL "
        SQL = SQL + "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & FreshBookQuantity & " As Quantity,'Board Name : '+Trim(G.PrintName) As BoardName," & _
                            "'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A WHERE B.Code=C.Code AND B.Board=G.Code AND C.Item=O.Code AND C.Category='3' AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") AND (" & FreshBookQuantity & ") <> 0 UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND I.Godown=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.Account=A.Code AND I.Quantity>=0 AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.Account=A.Code AND I.Quantity<0 AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.AccountFrom=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.AccountTo=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,Round(I.TotalConsumption,0),'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category=C.Category AND I.Item=C.Item) AND I.Vendor=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL "
        SQL = SQL + "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & RepairBookQuantity & " As Quantity,'Board Name : '+Trim(G.PrintName) As BoardName," & _
                            "'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A WHERE O.Type='R' AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") AND (" & RepairBookQuantity & ") <> 0 UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND I.Godown=A.Code AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.AccountTo=A.Code AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,Round(I.TotalConsumption,0),'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='4' AND I.Item=O.Code) AND I.Vendor=A.Code AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL "
        SQL = SQL + "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & TitleQuantity & " As Quantity,'Board Name : '+Trim(G.PrintName) As BoardName," & _
                            "'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A WHERE B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") AND (" & TitleQuantity & ") <> 0 UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND I.Godown=A.Code AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.AccountTo=A.Code AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,Round(I.TotalConsumption,0),'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='5' AND I.Item=O.Code) AND I.Vendor=A.Code AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# "
        If DatabaseType = "MS SQL" Then SQL = Replace(SQL, "#", "'")
        rstMaterialStockRegister.Open SQL & "ORDER BY GodownName,BoardName,BookName,ItemType,ItemName,VchDate,VchNo", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
        SelectedBooks = SelectedItems(ListView2, False)
        SelectedAccounts = SelectedItems(ListView3, False)
        If Val(ListView1.SelectedItem.SubItems(1)) = 1 Then
        If DatabaseType = "MS SQL" Then
                   OutsourceItemQuantity = "ISNULL((SELECT SUM(OpBal) FROM AccountChild0801 WHERE Category='1' AND Item=O.Code AND Code=A.Code),0)+" & _
                                                      "ISNULL((SELECT SUM(I.Quantity) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#),0)+" & _
                                                      "ISNULL((SELECT SUM(I.Quantity) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#),0)-" & _
                                                      "ISNULL((SELECT SUM(I.Quantity) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#),0)-" & _
                                                      "ISNULL((SELECT SUM(I.Quantity) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#),0)+" & _
                                                      "ISNULL((SELECT SUM(I.Quantity) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#),0)-" & _
                                                      "ISNULL((SELECT SUM(Round(I.TotalConsumption,0)) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND Category='1' AND Item=O.Code AND I.Vendor=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#),0)"
        Else
            OutsourceItemQuantity = "(SELECT ISNULL(SUM(OpBal),0) FROM AccountChild0801 WHERE Category='1' AND Item=O.Code AND Code=A.Code)+" & _
                                                      "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                      "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                      "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                      "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                      "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                      "(SELECT ISNULL(SUM(Round(I.TotalConsumption,0)),0)  FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND Category='1' AND Item=O.Code AND I.Vendor=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
        End If
     If DatabaseType = "MS SQL" Then OutsourceItemQuantity = Replace(OutsourceItemQuantity, "#", "'")
            SQL = "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & OutsourceItemQuantity & " As Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A WHERE A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") AND (" & OutsourceItemQuantity & ") <> 0 UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='1' AND I.Item=O.Code) AND I.Godown=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='1' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='1' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='1' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='1' AND I.Item=O.Code) AND M.AccountTo=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed :'+(Select Name From Bookmaster Where Code=(Select Book From BookPOParent Where Code= I.code)) As Particulars,Round(I.TotalConsumption,0),'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (BOM)' As ItemName,'1' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='1' AND I.Item=O.Code) AND I.Vendor=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# "
        ElseIf Val(ListView1.SelectedItem.SubItems(1)) = 3 Then
            FreshBookQuantity = "(SELECT ISNULL(SUM(OpBal),0) FROM AccountChild0801 WHERE Category='3' AND Item=O.Code AND Code=A.Code)+" & _
                                               "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='3' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                               "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='3' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                               "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='3' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                               "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='3' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                               "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='3' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                               "(SELECT ISNULL(SUM(Round(I.TotalConsumption,0)),0)  FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND Category='3' AND Item=O.Code AND I.Vendor=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
            SQL = "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & FreshBookQuantity & " As Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A WHERE A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") AND (" & FreshBookQuantity & ") <> 0 UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='3' AND I.Item=O.Code) AND I.Godown=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='3' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='3' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='3' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='3' AND I.Item=O.Code) AND M.AccountTo=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed :'+(Select Name From Bookmaster Where Code=(Select Book From BookPOParent Where Code= I.code)) As Particulars,Round(I.TotalConsumption,0),'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='3' AND I.Item=O.Code) AND I.Vendor=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# "
        ElseIf Val(ListView1.SelectedItem.SubItems(1)) = 4 Then
            RepairBookQuantity = "(SELECT ISNULL(SUM(OpBal),0) FROM AccountChild0801 WHERE Category='4' AND Item=O.Code AND Code=A.Code)+" & _
                                                "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                "(SELECT ISNULL(SUM(Round(I.TotalConsumption,0)),0)  FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND Category='4' AND Item=O.Code AND I.Vendor=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
            SQL = "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & RepairBookQuantity & " As Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A WHERE O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") AND (" & RepairBookQuantity & ") <> 0 UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND I.Godown=A.Code AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.AccountTo=A.Code AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed :'+(Select Name From Bookmaster Where Code=(Select Book From BookPOParent Where Code= I.code)) As Particulars,Round(I.TotalConsumption,0),'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='4' AND I.Item=O.Code) AND I.Vendor=A.Code AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# "
        ElseIf Val(ListView1.SelectedItem.SubItems(1)) = 5 Then
            TitleQuantity = "(SELECT ISNULL(SUM(OpBal),0) FROM AccountChild0801 WHERE Category='5' AND Item=O.Code AND Code=A.Code)+" & _
                                     "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                     "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                     "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                     "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                     "(SELECT ISNULL(SUM(I.Quantity),0) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                     "(SELECT SNULL(SUM(Round(I.TotalConsumption,0)),0))  FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND Category='5' AND Item=O.Code AND I.Vendor=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
            SQL = "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & TitleQuantity & " As Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A WHERE A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") AND (" & TitleQuantity & ") <> 0 UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND I.Godown=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT LTRIM(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.AccountTo=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT LTRIM(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed :'+(Select Name From Bookmaster Where Code=(Select Book From BookPOParent Where Code= I.code)) As Particulars,Round(I.TotalConsumption,0),'' As BoardName,'' As BookName,LTRIM(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+LTRIM(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='5' AND I.Item=O.Code) AND I.Vendor=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# "
        End If
        If DatabaseType = "MS SQL" Then SQL = Replace(SQL, "#", "'")
        rstMaterialStockRegister.Open SQL & "ORDER BY GodownName,ItemType,ItemName,VchDate,VchNo", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    Screen.MousePointer = vbNormal
    If rstMaterialStockRegister.RecordCount = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    rptMaterialStockRegister.Database.SetDataSource rstMaterialStockRegister, 3, 1
    rptMaterialStockRegister.DiscardSavedData
    Set CRXParamDefs = rptMaterialStockRegister.ParameterFields
    For Each CRXParamDef In CRXParamDefs
        If CRXParamDef.ParameterFieldName = "PF1" Then
            CRXParamDef.SetCurrentValue (IIf(Check1.Value, 0, 0.1))
        ElseIf CRXParamDef.ParameterFieldName = "PF2" Then
            CRXParamDef.SetCurrentValue (IIf(Option1.Value, "D", "S"))
        ElseIf CRXParamDef.ParameterFieldName = "PF3" Then
            CRXParamDef.SetCurrentValue (ReportType)
        End If
    Next
    rptMaterialStockRegister.EnableParameterPrompting = False
    If ReportType = "1" Then
        If Option2.Value Then
            rptMaterialStockRegister.PaperOrientation = crLandscape
        Else
            rptMaterialStockRegister.PaperOrientation = crPortrait
        End If
    Else
        rptMaterialStockRegister.PaperOrientation = crPortrait
    End If
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptMaterialStockRegister
        FrmReportViewer.Show vbModal
    Else
        rptMaterialStockRegister.PaperSource = crPRBinAuto
        rptMaterialStockRegister.PrintOut
    End If
    Set rptMaterialStockRegister = Nothing
    On Error GoTo 0
End Sub
