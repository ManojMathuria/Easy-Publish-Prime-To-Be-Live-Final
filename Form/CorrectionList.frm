VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCorrectionList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Status Report [Itemwise]"
   ClientHeight    =   7365
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9420
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
   ScaleHeight     =   7365
   ScaleWidth      =   9420
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Export"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Itemwise"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Memberwise"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Previous Report"
               EndProperty
            EndProperty
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
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CorrectionList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CorrectionList.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CorrectionList.frx":0658
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CorrectionList.frx":076C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CorrectionList.frx":0B47
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6975
      Left            =   45
      TabIndex        =   7
      Top             =   345
      Width           =   9330
      _Version        =   65536
      _ExtentX        =   16457
      _ExtentY        =   12303
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
      Picture         =   "CorrectionList.frx":0CE1
      Begin VB.CheckBox Check1 
         Caption         =   "All Entries"
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
         Left            =   8130
         TabIndex        =   3
         Top             =   53
         Width           =   1140
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Closed Entries"
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
         Left            =   6480
         TabIndex        =   2
         Top             =   53
         Width           =   1500
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3655
         Left            =   0
         TabIndex        =   4
         Top             =   320
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   6456
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   0
         TabIndex        =   5
         Top             =   3960
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   5318
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
         Picture         =   "CorrectionList.frx":0CFD
         Picture         =   "CorrectionList.frx":0D19
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
         Picture         =   "CorrectionList.frx":0D35
         Picture         =   "CorrectionList.frx":0D51
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
         Calendar        =   "CorrectionList.frx":0D6D
         Caption         =   "CorrectionList.frx":0E85
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "CorrectionList.frx":0EF1
         Keys            =   "CorrectionList.frx":0F0F
         Spin            =   "CorrectionList.frx":0F6D
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
         Calendar        =   "CorrectionList.frx":0F95
         Caption         =   "CorrectionList.frx":10AD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "CorrectionList.frx":1119
         Keys            =   "CorrectionList.frx":1137
         Spin            =   "CorrectionList.frx":1195
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
Attribute VB_Name = "FrmCorrectionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstBookList As New ADODB.Recordset
Dim rstMemberList As New ADODB.Recordset
Dim rstCorrectionList As New ADODB.Recordset
Dim RptType As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    RptType = "I"
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    If Format(FinancialYearTo, "yyyymmdd") > Format(Date, "yyyymmdd") Then MhDateInput2.Text = (Format(FinancialYearTo, "dd-mm-yyyy")) Else MhDateInput2.Text = Format(Date, "dd-mm-yyyy") + 90
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    rstBookList.Open "SELECT Name,Code FROM BookMaster WHERE Type='F' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstBookList.ActiveConnection = Nothing
    Call FillList(ListView2, "List of Items...", rstBookList)
    rstMemberList.Open "SELECT M.Name+' ('+D.Name+')' As Name,M.Code FROM TeamMemberMaster M INNER JOIN GeneralMaster D ON M.Department=D.Code ORDER BY M.Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstMemberList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Editorial Team Members...", rstMemberList)
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
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
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
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstMemberList)
    Call CloseRecordset(rstCorrectionList)
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (KeyCode = vbKeyA Or KeyCode = vbKeyD) And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (KeyCode = vbKeyA Or KeyCode = vbKeyD) And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 4 Then CloseForm Me: Exit Sub
    Call PrintCorrectionList(RptType, Choose(Button.Index, "S", "P", "E"))
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error Resume Next
    Me.Caption = "Project Status Report [" & Choose(ButtonMenu.Index, "Itemwise", "Memberwise", "Previous Report") & "]"
    RptType = Choose(ButtonMenu.Index, "I", "M", "R")
End Sub
Private Sub PrintCorrectionList(ByVal RptType As String, ByVal OutputTo As String)
On Error GoTo ErrHandler
    If rstCorrectionList.State = adStateOpen Then rstCorrectionList.Close
    Screen.MousePointer = vbHourglass
    If RptType = "I" Then
        rstCorrectionList.Open "SELECT M2.PrintName+' ('+D.PrintName+')' As ItemName,C.Correction,C.ArrivedOn,C.TargetDate,C.StartDate,C.RectifiedOn,C.Remarks,M1.PrintName As MemberName FROM ((BookMaster M1 INNER JOIN BookChild02 C ON M1.Code=C.Code) INNER JOIN TeamMemberMaster M2 ON C.[Member]=M2.Code) INNER JOIN GeneralMaster D ON M2.Designation=D.Code " & _
                                                        "WHERE " & IIf(Check1.Value, "1=1", IIf(Check2.Value, "Status='Done'", "Status<>'Done'")) & " AND M1.Code In (" & SelectedItems(ListView2) & ") AND M2.Code In (" & SelectedItems(ListView1) & ") AND TargetDate>='" & GetDate(MhDateInput1.Text) & "' AND TargetDate<='" & GetDate(MhDateInput2.Text) & "' ORDER BY M2.PrintName,M1.PrintName,C.SNo", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf RptType = "M" Then
        rstCorrectionList.Open "SELECT M1.PrintName As ItemName,C.Correction,C.ArrivedOn,C.TargetDate,C.StartDate,C.RectifiedOn,C.Remarks,M2.PrintName+' ('+D.PrintName+')' As MemberName FROM ((BookMaster M1 INNER JOIN BookChild02 C ON M1.Code=C.Code) INNER JOIN TeamMemberMaster M2 ON C.[Member]=M2.Code) INNER JOIN GeneralMaster D ON M2.Designation=D.Code " & _
                                                       "WHERE " & IIf(Check1.Value, "1=1", IIf(Check2.Value, "Status='Done'", "Status<>'Done'")) & " AND M1.Code In (" & SelectedItems(ListView2) & ") AND M2.Code In (" & SelectedItems(ListView1) & ") AND TargetDate>='" & GetDate(MhDateInput1.Text) & "' AND TargetDate<='" & GetDate(MhDateInput2.Text) & "' ORDER BY M2.PrintName,M1.PrintName,C.SNo", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstCorrectionList.Open "SELECT M1.PrintName As ItemName,C.Correction,C.ArrivedOn,C.TargetDate,C.ArrivedOn,C.RectifiedOn,C.Remarks,M2.PrintName As MemberName FROM ((BookMaster M1 INNER JOIN BookChild0201 C ON M1.Code=C.Code) INNER JOIN GeneralMaster M2 ON C.[Member]=M2.Code) " & _
                                                       "WHERE " & IIf(Check1.Value, "1=1", IIf(Check2.Value, "RectifiedOn<>''", "RectifiedOn<>''")) & " AND M1.Code In (" & SelectedItems(ListView2) & ") AND TargetDate>='" & GetDate(MhDateInput1.Text) & "' AND TargetDate<='" & GetDate(MhDateInput2.Text) & "' ORDER BY M2.PrintName,M1.PrintName,C.SNo", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
    Screen.MousePointer = vbNormal
    If InStr(1, "S_P", OutputTo) > 0 Then
        On Error Resume Next
        rptCorrectionList.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptCorrectionList.Text5.SetText "From [" + Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy") + "]"
        If rstCorrectionList.RecordCount = 0 Then On Error GoTo 0: Exit Sub
        rptCorrectionList.Database.SetDataSource rstCorrectionList, 3, 1
        rptCorrectionList.DiscardSavedData
        If OutputTo = "S" Then Set FrmReportViewer.Report = rptCorrectionList: FrmReportViewer.Show vbModal Else rptCorrectionList.PaperSource = crPRBinAuto: rptCorrectionList.PrintOut
        Set rptCorrectionList = Nothing
        On Error GoTo 0
    Else
        Dim oExcel As Object, i As Integer
        On Error Resume Next
        If Not FileExist(App.Path & "\Template\Manuscript Status Register.xlsx") Then Exit Sub
        DoEvents
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open (App.Path & "\Template\Manuscript Status Register")
        oExcel.DisplayAlerts = False
        oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Manuscript Status Register (" & CompCode & ")")
        oExcel.DisplayAlerts = True
        oExcel.Sheets("Sheet1").Select
        oExcel.Visible = False
        oExcel.Cells(1, 1).Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
        oExcel.Cells(2, 1).Value = "From [" & Format(MhDateInput1.Text, "dd-mm-yyyy") & "] To [" & Format(MhDateInput2.Text, "dd-mm-yyyy") & "]"
        i = 4
        Do While Not rstCorrectionList.EOF
            oExcel.Application.Cells(i, 1).Value = IIf(RptType = "I", Trim(rstCorrectionList.Fields("ItemName").Value), Trim(rstCorrectionList.Fields("MemberName").Value))
            oExcel.Application.Cells(i, 2).Value = IIf(RptType = "I", Trim(rstCorrectionList.Fields("MemberName").Value), Trim(rstCorrectionList.Fields("ItemName").Value))
            oExcel.Application.Cells(i, 3).Value = Format(rstCorrectionList.Fields("ArrivedOn").Value, "dd-MMM-yyyy")
            oExcel.Application.Cells(i, 4).Value = rstCorrectionList.Fields("Correction").Value
            oExcel.Application.Cells(i, 5).Value = Format(rstCorrectionList.Fields("TargetDate").Value, "dd-MMM-yyyy")
            oExcel.Application.Cells(i, 6).Value = Format(rstCorrectionList.Fields("StartDate").Value, "dd-MMM-yyyy")
            oExcel.Application.Cells(i, 7).Value = Format(rstCorrectionList.Fields("RectifiedOn").Value, "dd-MMM-yyyy")
            oExcel.Application.Cells(i, 8).Value = rstCorrectionList.Fields("Remarks").Value
            oExcel.Application.Cells(i, 9).Value = DateDiff("D", rstCorrectionList.Fields("StartDate").Value, rstCorrectionList.Fields("RectifiedOn").Value)
            i = i + 1
            rstCorrectionList.MoveNext
        Loop
        oExcel.Sheets("Sheet1").Activate
        oExcel.Columns("A:H").EntireColumn.AutoFit
        oExcel.Workbooks.Item(1).Save
        Screen.MousePointer = vbNormal
        oExcel.Range("A1").Activate
        oExcel.Visible = True
        Set oExcel = Nothing
        On Error GoTo 0
    End If
Screen.MousePointer = vbNormal
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
