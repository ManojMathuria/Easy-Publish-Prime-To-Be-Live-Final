VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFinishSizeMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finish Size Master"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6750
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   4870
      Left            =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   8590
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Picture         =   "FinishSizeMaster.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   4630
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8176
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&List"
         TabPicture(0)   =   "FinishSizeMaster.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Mh3dLabel1(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "DataGrid1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Text1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "FinishSizeMaster.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).ControlCount=   1
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   600
            TabIndex        =   10
            Top             =   4160
            Width           =   5775
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3640
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   450
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   6429
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   9164542
            HeadLines       =   1
            RowHeight       =   18
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   1
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "Name"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               ScrollBars      =   3
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   5940.284
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   3780
            Left            =   -74880
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   6667
            _StockProps     =   77
            Enabled         =   0   'False
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
            Picture         =   "FinishSizeMaster.frx":0054
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
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   1
               Top             =   425
               Width           =   4695
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
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   0
               Top             =   100
               Width           =   4695
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   5
               Top             =   425
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "FinishSizeMaster.frx":0070
               Picture         =   "FinishSizeMaster.frx":008C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   4
               Top             =   100
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "FinishSizeMaster.frx":00A8
               Picture         =   "FinishSizeMaster.frx":00C4
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   2715
               Left            =   120
               TabIndex        =   2
               Top             =   960
               Width           =   6030
               _Version        =   524288
               _ExtentX        =   10636
               _ExtentY        =   4789
               _StockProps     =   64
               ButtonDrawMode  =   8
               EditEnterAction =   5
               EditModeReplace =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               GridColor       =   4227327
               MaxCols         =   6
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "FinishSizeMaster.frx":00E0
            End
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
               Left            =   1320
               MaxLength       =   40
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   2055
               Width           =   4695
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   6300
               Y1              =   850
               Y2              =   850
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   285
            Index           =   2
            Left            =   2040
            TabIndex        =   13
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   503
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
            Caption         =   "Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D->Delete  Ctrl+S->Save"
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "FinishSizeMaster.frx":08D4
            Picture         =   "FinishSizeMaster.frx":08F0
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H008BD6FE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Find"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   120
            TabIndex        =   12
            Top             =   4160
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filter"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print Preview"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Mail"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "First"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Previous"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Next"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Last"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4
      Left            =   2760
      Top             =   2280
   End
End
Attribute VB_Name = "FrmFinishSizeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SL As Boolean 'Selection List
Public MasterCode As String  'Master to Modify
Dim cnFinishSizeMaster As New ADODB.Connection
Dim rstFinishSizeList As New ADODB.Recordset
Dim rstFinishSizeMaster As New ADODB.Recordset
Dim rstFinishSizeChild As New ADODB.Recordset
Dim rstSizeList As New ADODB.Recordset
Dim SizeCode As String, EditMode As Boolean
Dim SortOrder, PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstFinishSizeList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstFinishSizeList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstFinishSizeList.EOF Or rstFinishSizeList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub Form_Load()
    If Not SL Then MasterCode = ""
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
    BusySystemIndicator True
    cnFinishSizeMaster.CursorLocation = adUseClient: cnFinishSizeMaster.Open cnDatabase.ConnectionString
    rstFinishSizeList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type='11' ORDER BY Name", cnDatabase, adOpenKeyset, adLockOptimistic
    rstSizeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='1' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstFinishSizeMaster.CursorLocation = adUseClient
    rstFinishSizeList.Filter = adFilterNone
    If rstFinishSizeList.RecordCount > 0 Then
        If CheckEmpty(MasterCode, False) Then
            rstFinishSizeList.MoveLast
        Else
            rstFinishSizeList.MoveFirst
            rstFinishSizeList.Find "[Code]='" & MasterCode & "'"
        End If
    End If
    Set DataGrid1.DataSource = rstFinishSizeList
    SSTab1.Tab = 0
    If Not (rstFinishSizeList.EOF Or rstFinishSizeList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstSizeList.ActiveConnection = Nothing
    rstFinishSizeList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    MdiMainMenu.mnuFinishSizeMaster.Enabled = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If SSTab1.Tab = 0 Then
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If Not EditMode Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then Me.ActiveControl.SetFocus Else Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                End If
            End If
        End If
        If Not EditMode Then KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF8 And Toolbar1.Buttons.Item(3).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(13)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(14)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyN And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(15)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(16)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        If Toolbar1.Buttons.Item(1).Enabled Then
            If SL Then
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstFinishSizeList.Fields("Code").Value: slName = rstFinishSizeList.Fields("Name").Value: KeyCode = 0: Unload Me: Exit Sub
            Else
                SSTab1.Tab = 1: SSTab1.SetFocus
            End If
        Else
           If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
    Else
        If Me.Tag <> "S" Then slCode = "": slName = ""
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstFinishSizeList)
    Call CloseRecordset(rstFinishSizeMaster)
    Call CloseRecordset(rstFinishSizeChild)
    Call CloseConnection(cnFinishSizeMaster)
    ShowProgressInStatusBar False
    MdiMainMenu.mnuFinishSizeMaster.Enabled = True
End Sub
Private Sub Text1_Change()
    On Error Resume Next
    With rstFinishSizeList
    If SortOrder = "" Then SortOrder = "Name"
        If .RecordCount = 0 Then Exit Sub
              .MoveFirst
        If Not CheckEmpty(Text1.Text, False) Then
            If SortOrder = "Name" Then .Filter = "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'" Else .Filter = "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
            If .EOF Then
            .Filter = adFilterNone
                .MoveFirst
                If PrevStr <> "" And Len(Text1.Text) > 1 Then If dblBookMark <> 0 Then .Bookmark = dblBookMark Else PrevStr = ""
                Beep
                DisplayError ("Spelling Error")
                Text1.Text = PrevStr
                Sendkeys "{End}"
            Else
                PrevStr = Text1.Text
                dblBookMark = DataGrid1.Bookmark
            End If
        Else
            .Filter = adFilterNone
            PrevStr = ""
        End If
        If Not (.EOF Or .BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
    End With
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    If rstFinishSizeList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstFinishSizeList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstFinishSizeList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstFinishSizeList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstFinishSizeList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstFinishSizeList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstFinishSizeList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstFinishSizeList
            .MoveLast
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    End If
    If KeyProcessed Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
        KeyProcessed = False
        KeyCode = 0
    End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If Toolbar1.Buttons.Item(1).Enabled Then
        If SSTab1.Tab = 1 Then
           ViewRecord
        Else
            If Not (rstFinishSizeList.EOF Or rstFinishSizeList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            Text1.SetFocus
        End If
        SSTab1.TabEnabled(0) = True
    Else
       SSTab1.TabEnabled(0) = False
       Text2.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim UpdateFlag As Integer, i As Integer, SizeCode As Variant
    Dim HiLiteRecord As Boolean
    If Button.Index = 1 Then
        If rstFinishSizeMaster.State = adStateOpen Then rstFinishSizeMaster.Close
        rstFinishSizeMaster.Open "SELECT * FROM GeneralMaster WHERE Code=''", cnDatabase, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstFinishSizeMaster) Then
           Call SetButtons(False)
           SSTab1.Tab = 1
           Text2.SetFocus
           blnRecordExist = False
           cnFinishSizeMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstFinishSizeList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstFinishSizeList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Master"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnDatabase.Execute "DELETE FROM GeneralMaster WHERE Code = '" & rstFinishSizeList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstFinishSizeList.Delete
                rstFinishSizeList.MoveNext
                If rstFinishSizeList.RecordCount > 0 And rstFinishSizeList.EOF Then rstFinishSizeList.MoveLast
                Call UpdateUserAction("Size Group Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), cnDatabase)
                ShowProgressInStatusBar True
                Timer1.Enabled = True
            Else
                DisplayError ("Failed to delete the record")
            End If
            On Error GoTo 0
        End If
        SetButtons (True)
        SetButtonsForNoRecord
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 4 Then
        If CheckMandatoryFields Then Exit Sub
        If blnRecordExist And AllowMastersModification = 0 Then Call DisplayError("You don't have the rights to Edit this Master"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(5): Exit Sub
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstFinishSizeMaster) Then
            If UpdateSizeList("D") Then
                UpdateFlag = 1
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.SetActiveCell 5, i
                    fpSpread1.GetText 5, i, SizeCode
                    If Not CheckEmpty(SizeCode, False) Then
                        If Not UpdateSizeList("I") Then UpdateFlag = 0: Exit For
                    End If
                Next
            End If
        End If
        If UpdateFlag Then
            Call UpdateUserAction("Size Group Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), cnDatabase)
            AddToList
            cnFinishSizeMaster.CommitTrans
            If rstFinishSizeMaster.State = adStateOpen Then rstFinishSizeMaster.Close
            rstFinishSizeMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            Call MsgBox("Record updated !!!", vbInformation, App.Title)
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstFinishSizeMaster) Then
            cnFinishSizeMaster.RollbackTrans
            If rstFinishSizeMaster.State = adStateOpen Then rstFinishSizeMaster.Close
            rstFinishSizeMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstFinishSizeList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstFinishSizeList): Loop
        Set DataGrid1.DataSource = rstFinishSizeList
        rstFinishSizeList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Name", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstFinishSizeList.RecordCount > 0 Then rstFinishSizeList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstFinishSizeList.RecordCount > 0 Then
           rstFinishSizeList.MovePrevious
           If rstFinishSizeList.BOF Then rstFinishSizeList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstFinishSizeList.RecordCount > 0 Then
           rstFinishSizeList.MoveNext
           If rstFinishSizeList.EOF Then rstFinishSizeList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstFinishSizeList.RecordCount > 0 Then rstFinishSizeList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstFinishSizeList.EOF Or rstFinishSizeList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        Text1.SetFocus
    End If
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
End Sub
Private Sub SetButtons(bVal As Boolean)
    Toolbar1.Buttons.Item(1).Enabled = bVal
    Toolbar1.Buttons.Item(2).Enabled = bVal
    Toolbar1.Buttons.Item(3).Enabled = bVal
    Toolbar1.Buttons.Item(4).Enabled = Not bVal
    Toolbar1.Buttons.Item(5).Enabled = Not bVal
    Toolbar1.Buttons.Item(6).Enabled = bVal
    Toolbar1.Buttons.Item(7).Enabled = bVal
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstFinishSizeList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstFinishSizeMaster.EOF Or rstFinishSizeMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnDatabase, "GeneralMaster", "Code", "Name+Type", Trim(Text2.Text) & "11", rstFinishSizeMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = Text2.Text
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then fpSpread1.DeleteRows fpSpread1.ActiveRow, 1: fpSpread1.SetFocus
    ElseIf KeyCode = vbKeySpace Then
        Dim Size As Variant
        With fpSpread1
            If .ActiveCol = 1 Then
                .GetText .ActiveCol, .ActiveRow, Size
                Text4.Text = FixQuote(Size)
                If rstSizeList.RecordCount = 0 Then DisplayError ("No Record in Size Master"): .SetActiveCell 1, .ActiveRow: .SetFocus: Exit Sub Else rstSizeList.MoveFirst
                rstSizeList.Find "[Col0] = '" & FixQuote(Trim(Size)) & "'"
                SelectionType = "S": SizeCode = ""
                Call LoadSelectionList(rstSizeList, "List of Sizes...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text4, SizeCode)
                Call CloseForm(FrmSelectionList)
                If SizeCode = "" Then
                    .SetText 1, .ActiveRow, "": .SetActiveCell 1, .ActiveRow
                Else
                    .SetText 1, .ActiveRow, Text4.Text
                    .SetText 5, .ActiveRow, SizeCode
                    .SetFocus
                    Sendkeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 4 Then
                .GetText .ActiveCol, .ActiveRow, Size
                Text4.Text = FixQuote(Size)
                If rstSizeList.RecordCount = 0 Then DisplayError ("No Record in Size Master"): .SetActiveCell 1, .ActiveRow: .SetFocus: Exit Sub Else rstSizeList.MoveFirst
                rstSizeList.Find "[Col0] = '" & FixQuote(Trim(Size)) & "'"
                SelectionType = "S": SizeCode = ""
                Call LoadSelectionList(rstSizeList, "List of Sizes...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text4, SizeCode)
                Call CloseForm(FrmSelectionList)
                If SizeCode = "" Then
                    .SetText 4, .ActiveRow, "": .SetActiveCell 1, .ActiveRow
                Else
                    .SetText 4, .ActiveRow, Text4.Text
                    .SetText 6, .ActiveRow, SizeCode
                    .SetFocus
                    Sendkeys "{ENTER}"
                End If
            End If
        End With
    End If
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstFinishSizeList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstFinishSizeMaster.State = adStateOpen Then rstFinishSizeMaster.Close
    rstFinishSizeMaster.Open "SELECT * FROM GeneralMaster WHERE Code='" & FixQuote(rstFinishSizeList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    If rstFinishSizeMaster.RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    fpSpread1.SetActiveCell 1, 1
End Sub
Private Sub LoadFields()
    If rstFinishSizeMaster.EOF Or rstFinishSizeMaster.BOF Then Exit Sub
    Text2.Text = rstFinishSizeMaster.Fields("Name").Value
    Text3.Text = rstFinishSizeMaster.Fields("PrintName").Value
    Call LoadSizeList(rstFinishSizeMaster.Fields("Code").Value)
End Sub
Private Sub LoadSizeList(ByVal strFinishSizeCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstFinishSizeChild.State = adStateOpen Then rstFinishSizeChild.Close
    rstFinishSizeChild.Open "SELECT M1.Name As TextSizeName,[Ups/Form],[Ups/BdgForm],M2.Name As TitleSizeName,C.[TextSize] As TextSizeCode,C.TitleSize As TitleSizeCode FROM (FinishSizeChild C INNER JOIN GeneralMaster M1 ON C.[TextSize]=M1.Code) INNER JOIN GeneralMaster M2 ON C.TitleSize=M2.Code WHERE C.Code='" & strFinishSizeCode & "' ORDER BY M1.Name", cnFinishSizeMaster, adOpenKeyset, adLockReadOnly
    rstFinishSizeChild.ActiveConnection = Nothing
    If rstFinishSizeChild.RecordCount > 0 Then rstFinishSizeChild.MoveFirst
    i = 0
    Do While Not rstFinishSizeChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstFinishSizeChild.Fields("TextSizeName").Value
            .SetText 2, i, Val(rstFinishSizeChild.Fields("Ups/Form").Value)
            .SetText 3, i, Val(rstFinishSizeChild.Fields("Ups/BdgForm").Value)
            .SetText 4, i, rstFinishSizeChild.Fields("TitleSizeName").Value
            .SetText 5, i, rstFinishSizeChild.Fields("TextSizeCode").Value
            .SetText 6, i, rstFinishSizeChild.Fields("TitleSizeCode").Value
        End With
        rstFinishSizeChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Size List")
End Sub
Private Function UpdateSizeList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 4) As Variant
    On Error GoTo ErrorHandler
    UpdateSizeList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        cnFinishSizeMaster.Execute "DELETE FROM FinishSizeChild WHERE Code='" & rstFinishSizeMaster.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 5, .ActiveRow, CellVal(1)
            .GetText 2, .ActiveRow, CellVal(2)
            .GetText 3, .ActiveRow, CellVal(3)
            .GetText 6, .ActiveRow, CellVal(4)
        End With
        cnFinishSizeMaster.Execute "INSERT INTO FinishSizeChild VALUES ('" & rstFinishSizeMaster.Fields("Code").Value & "','" & CellVal(1) & "'," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ",'" & CellVal(4) & "')"
    End If
    Exit Function
ErrorHandler:
    UpdateSizeList = False
End Function
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstFinishSizeMaster.RecordCount = 0 Then Exit Sub
    If rstFinishSizeMaster.State = adStateOpen Then rstFinishSizeMaster.Close
    rstFinishSizeMaster.CursorLocation = adUseServer
    rstFinishSizeMaster.Open "SELECT * FROM GeneralMaster WHERE Code='" & FixQuote(rstFinishSizeList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstFinishSizeMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    cnFinishSizeMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstFinishSizeMaster.EOF Or rstFinishSizeMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstFinishSizeMaster.Fields("Code").Value = GenerateCode(cnDatabase, "SELECT MAX(Code) FROM GeneralMaster", 6, "0")
        rstFinishSizeMaster.Fields("CreatedBy").Value = UserCode
        rstFinishSizeMaster.Fields("CreatedOn").Value = Now()
        rstFinishSizeMaster.Fields("Recordstatus").Value = "N"
    Else
        rstFinishSizeMaster.Fields("ModifiedBy").Value = UserCode
        rstFinishSizeMaster.Fields("ModifiedOn").Value = Now()
        rstFinishSizeMaster.Fields("Recordstatus").Value = "M"
    End If
    rstFinishSizeMaster.Fields("Name").Value = Trim(Text2.Text)
    rstFinishSizeMaster.Fields("PrintName").Value = Trim(Text3.Text)
    rstFinishSizeMaster.Fields("Type").Value = "11"
    rstFinishSizeMaster.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstFinishSizeList.MoveFirst
    rstFinishSizeList.Find "[Code] = '" & rstFinishSizeMaster.Fields("Code").Value & "'"
    If rstFinishSizeList.EOF Then rstFinishSizeList.AddNew: rstFinishSizeList.Fields("Code").Value = rstFinishSizeMaster.Fields("Code").Value
    rstFinishSizeList.Fields("Name").Value = rstFinishSizeMaster.Fields("Name").Value
    rstFinishSizeList.Update
    rstFinishSizeList.Sort = "Name Asc"
    rstFinishSizeList.Find "[Code] = '" & rstFinishSizeMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        Text2.SetFocus: CheckMandatoryFields = True
    ElseIf CheckDuplicate(cnDatabase, "GeneralMaster", "Code", "Name+Type", Trim(Text2.Text) & "11", rstFinishSizeMaster.Fields("Code").Value, False) Then
        Text2.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus: CheckMandatoryFields = True
    ElseIf InStr(1, Text2.Text, "X") = 0 Then
        DisplayError ("Size Format must be 00.00X00.00")
        Text2.SetFocus: CheckMandatoryFields = True
    ElseIf InStr(1, UCase(Text2.Text), "X") > 0 Then
        If Len(Left(UCase(Text2.Text), InStr(1, UCase(Text2.Text), "X") - 1)) <> 5 Or Len(Mid(UCase(Text2.Text), InStr(1, UCase(Text2.Text), "X") + 1, 5)) <> 5 Or (Not IsNumeric(Left(UCase(Text2.Text), InStr(1, UCase(Text2.Text), "X") - 1))) Or (Not IsNumeric(Mid(UCase(Text2.Text), InStr(1, UCase(Text2.Text), "X") + 1, 5))) Then
            DisplayError ("Size Format must be 00.00X00.00")
            Text2.SetFocus: CheckMandatoryFields = True
        End If
    Else
        Dim i As Integer, SizeCode As Variant
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.SetActiveCell 5, i
            fpSpread1.GetText 5, i, SizeCode
            If CheckEmpty(SizeCode, False) Then CheckMandatoryFields = True: DisplayError "Data incomplete in row #" & Trim(Str(i)): Exit For
            fpSpread1.GetText 2, i, SizeCode
            If Val(SizeCode) = 0 Then CheckMandatoryFields = True: DisplayError "Data incomplete in row #" & Trim(Str(i)): Exit For
            fpSpread1.GetText 3, i, SizeCode
            If Val(SizeCode) = 0 Then CheckMandatoryFields = True: DisplayError "Data incomplete in row #" & Trim(Str(i)): Exit For
            fpSpread1.GetText 6, i, SizeCode
            If CheckEmpty(SizeCode, False) Then CheckMandatoryFields = True: DisplayError "Data incomplete in row #" & Trim(Str(i)): Exit For
        Next
    End If
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstFinishSizeList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then Timer1.Enabled = False: ShowProgressInStatusBar False
End Sub
