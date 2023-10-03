VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOutsourceItemMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Item (BOM)  Master"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
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
   ScaleHeight     =   5160
   ScaleWidth      =   7950
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   5160
      Left            =   15
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   7935
      _Version        =   65536
      _ExtentX        =   13996
      _ExtentY        =   9102
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
      Picture         =   "OutsourceItemMaster.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   4935
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   8705
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
         TabPicture(0)   =   "OutsourceItemMaster.frx":001C
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
         TabPicture(1)   =   "OutsourceItemMaster.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).Control(0).Enabled=   0   'False
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
            Left            =   840
            TabIndex        =   6
            Top             =   4450
            Width           =   6735
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3930
            Left            =   120
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   480
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   6932
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "Name"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Unit"
               Caption         =   "Unit"
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
                  ColumnWidth     =   5174.929
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   6870.047
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   1215
            Left            =   -74880
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   480
            Width           =   7455
            _Version        =   65536
            _ExtentX        =   13150
            _ExtentY        =   2143
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
            Picture         =   "OutsourceItemMaster.frx":0054
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
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   11
               Top             =   720
               Width           =   5655
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
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   1
               Top             =   420
               Width           =   5655
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
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   0
               Top             =   105
               Width           =   5655
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   9
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
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
               Picture         =   "OutsourceItemMaster.frx":0070
               Picture         =   "OutsourceItemMaster.frx":008C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   10
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
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
               Picture         =   "OutsourceItemMaster.frx":00A8
               Picture         =   "OutsourceItemMaster.frx":00C4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   120
               TabIndex        =   12
               Top             =   720
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
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
               Caption         =   " UOM"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemMaster.frx":00E0
               Picture         =   "OutsourceItemMaster.frx":00FC
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   3120
            TabIndex        =   13
            Top             =   0
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
            _ExtentY        =   582
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
            Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D->Delete  Ctrl+S->Save"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "OutsourceItemMaster.frx":0118
            Picture         =   "OutsourceItemMaster.frx":0134
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
            ForeColor       =   &H80000009&
            Height          =   330
            Left            =   120
            TabIndex        =   8
            Top             =   4455
            Width           =   735
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
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
Attribute VB_Name = "FrmOutsourceItemMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SL As Boolean 'Selection List
Public MasterCode As String  'Master to Modify
Dim CxnOutsourceItemMaster As New ADODB.Connection
Dim rstOutsourceItemList As New ADODB.Recordset
Dim rstOutsourceItemMaster As New ADODB.Recordset
Dim rstUOMList As New ADODB.Recordset
Dim UOMCode As String
Dim SortOrder, PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    If Not SL Then MasterCode = ""
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    CxnOutsourceItemMaster.CursorLocation = adUseClient
    CxnOutsourceItemMaster.Open cnDatabase.ConnectionString
    rstOutsourceItemList.Open "Select Name,Code,ISNULL((Select Name From GeneralMaster Where Code=UOM),'') As UOM From OutsourceItemMaster Order By Name", CxnOutsourceItemMaster, adOpenKeyset, adLockOptimistic
    rstUOMList.Open "SELECT Name As Col0,Value1,Code FROM GeneralMaster WHERE Type='25' ORDER BY Name", CxnOutsourceItemMaster, adOpenKeyset, adLockReadOnly
    rstOutsourceItemMaster.CursorLocation = adUseClient
    rstOutsourceItemList.Filter = adFilterNone
    If rstOutsourceItemList.RecordCount > 0 Then
        If CheckEmpty(MasterCode, False) Then
            rstOutsourceItemList.MoveLast
        Else
            rstOutsourceItemList.MoveFirst
            rstOutsourceItemList.Find "[Code]='" & MasterCode & "'"
        End If
    End If
    Set DataGrid1.DataSource = rstOutsourceItemList
    BusySystemIndicator False
    SSTab1.Tab = 0
    If Not (rstOutsourceItemList.EOF Or rstOutsourceItemList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstOutsourceItemList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    SortOrder = "Name"
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_Activate()
    MdiMainMenu.mnuOutsourceItemMaster.Enabled = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If SSTab1.Tab = 0 Then
            CloseForm Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                    Me.ActiveControl.SetFocus
                Else
                    Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                End If
            End If
            KeyCode = 0
        End If
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
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
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
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstOutsourceItemList.Fields("Code").Value: slName = rstOutsourceItemList.Fields("Name").Value: KeyCode = 0: Unload Me: Exit Sub
            Else
                SSTab1.Tab = 1
                SSTab1.SetFocus
            End If
        Else
            Sendkeys "{TAB}"
        End If
        KeyCode = 0
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
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstOutsourceItemMaster)
    Call CloseConnection(CxnOutsourceItemMaster)
    ShowProgressInStatusBar False
    MdiMainMenu.mnuOutsourceItemMaster.Enabled = True
End Sub
Private Sub Text1_Change()
On Error Resume Next
    With rstOutsourceItemList
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
    
    If rstOutsourceItemList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstOutsourceItemList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstOutsourceItemList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstOutsourceItemList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstOutsourceItemList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstOutsourceItemList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstOutsourceItemList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstOutsourceItemList
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
    On Error Resume Next
    
    If Toolbar1.Buttons.Item(1).Enabled Then
        If SSTab1.Tab >= 1 Then
            ViewRecord
        Else
            If Not (rstOutsourceItemList.EOF Or rstOutsourceItemList.BOF) Then
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
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text4.Text)
        If rstUOMList.RecordCount = 0 Then DisplayError ("No Record in UOM Master"): Text4.SetFocus: Exit Sub Else rstUOMList.MoveFirst
        rstUOMList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": UOMCode = ""
        Call LoadSelectionList(rstUOMList, "List of UOMs...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text4, UOMCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(UOMCode) <> "" Then rstUOMList.MoveFirst: rstUOMList.Find "[Code] = '" & UOMCode & "'"  ':CalcDependents: SendKeys "{TAB}" Else Text4.Text = ""      'MhRealInput10.Value = Val(rstUOMList.Fields("Value1").Value):
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If CheckEmpty(Text4.Text, False) Then Cancel = True
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    
    If Button.Index = 1 Then
        If rstOutsourceItemMaster.State = adStateOpen Then
           rstOutsourceItemMaster.Close
        End If
        rstOutsourceItemMaster.Open "Select * From OutsourceItemMaster Where Code = ''", CxnOutsourceItemMaster, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstOutsourceItemMaster) Then
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            CxnOutsourceItemMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstOutsourceItemList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstOutsourceItemList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Master")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnOutsourceItemMaster.Execute "DELETE FROM OutsourceItemMaster WHERE Code = '" & rstOutsourceItemList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstOutsourceItemList.Delete
                rstOutsourceItemList.MoveNext
                If rstOutsourceItemList.RecordCount > 0 And rstOutsourceItemList.EOF Then rstOutsourceItemList.MoveLast
                Call UpdateUserAction("Outsource Item Master", "D", Trim(Text2.Text), cnDatabase)
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
        If blnRecordExist And AllowMastersModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Master")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        SaveFields
        If UpdateRecord(rstOutsourceItemMaster) Then
            Call UpdateUserAction("Outsource Item Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), cnDatabase)
            AddToList
            CxnOutsourceItemMaster.CommitTrans
            If rstOutsourceItemMaster.State = adStateOpen Then
                rstOutsourceItemMaster.Close
            End If
            rstOutsourceItemMaster.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstOutsourceItemMaster) Then
            CxnOutsourceItemMaster.RollbackTrans
            If rstOutsourceItemMaster.State = adStateOpen Then
                rstOutsourceItemMaster.Close
            End If
            rstOutsourceItemMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstOutsourceItemList.ActiveConnection = CxnOutsourceItemMaster
        Do While Not RefreshRecord(rstOutsourceItemList)
        Loop
        Set DataGrid1.DataSource = rstOutsourceItemList
        rstOutsourceItemList.ActiveConnection = Nothing
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
        If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstOutsourceItemList.RecordCount > 0 Then
           rstOutsourceItemList.MovePrevious
           If rstOutsourceItemList.BOF Then
              rstOutsourceItemList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstOutsourceItemList.RecordCount > 0 Then
           rstOutsourceItemList.MoveNext
           If rstOutsourceItemList.EOF Then
              rstOutsourceItemList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        CloseForm Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstOutsourceItemList.EOF Or rstOutsourceItemList.BOF) Then
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
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstOutsourceItemList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstOutsourceItemList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstOutsourceItemList.EOF Or rstOutsourceItemList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
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
    If rstOutsourceItemList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstOutsourceItemMaster.EOF Or rstOutsourceItemMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnOutsourceItemMaster, "OutsourceItemMaster", "Code", "Name", Text2.Text, rstOutsourceItemMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = Text2.Text
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstOutsourceItemList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstOutsourceItemMaster.State = adStateOpen Then
       rstOutsourceItemMaster.Close
    End If
    rstOutsourceItemMaster.Open "Select * From OutsourceItemMaster Where Code = '" & FixQuote(rstOutsourceItemList.Fields("Code").Value) & "'", CxnOutsourceItemMaster, adOpenKeyset, adLockOptimistic
    If rstOutsourceItemMaster.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text4.Enabled = True
End Sub
Private Sub LoadFields()
    If rstOutsourceItemMaster.EOF Or rstOutsourceItemMaster.BOF Then Exit Sub
    Text2.Text = rstOutsourceItemMaster.Fields("Name").Value
    Text3.Text = rstOutsourceItemMaster.Fields("PrintName").Value
    UOMCode = rstOutsourceItemMaster.Fields("UOM").Value
    rstUOMList.MoveFirst
    rstUOMList.Find "[Code] = '" & UOMCode & "'"
    Text4.Text = rstUOMList.Fields("Col0").Value
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstOutsourceItemMaster.RecordCount = 0 Then Exit Sub
    If rstOutsourceItemMaster.State = adStateOpen Then
       rstOutsourceItemMaster.Close
    End If
    rstOutsourceItemMaster.CursorLocation = adUseServer
    rstOutsourceItemMaster.Open "Select * From OutsourceItemMaster Where Code = '" & FixQuote(rstOutsourceItemList.Fields("Code").Value) & "'", CxnOutsourceItemMaster, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstOutsourceItemMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    CxnOutsourceItemMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstOutsourceItemMaster.EOF Or rstOutsourceItemMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstOutsourceItemMaster.Fields("Code").Value = GenerateCode(CxnOutsourceItemMaster, "Select Max(Code) From OutsourceItemMaster", 6, "0")
        rstOutsourceItemMaster.Fields("CreatedBy").Value = UserCode
        rstOutsourceItemMaster.Fields("CreatedOn").Value = Now()
        rstOutsourceItemMaster.Fields("Recordstatus").Value = "N"
    Else
        rstOutsourceItemMaster.Fields("ModifiedBy").Value = UserCode
        rstOutsourceItemMaster.Fields("ModifiedOn").Value = Now()
        rstOutsourceItemMaster.Fields("Recordstatus").Value = "M"
    End If
    rstOutsourceItemMaster.Fields("Name").Value = Trim(Text2.Text)
    rstOutsourceItemMaster.Fields("PrintName").Value = Trim(Text3.Text)
    rstOutsourceItemMaster.Fields("UOM").Value = UOMCode
    rstOutsourceItemMaster.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    
    rstOutsourceItemList.MoveFirst
    rstOutsourceItemList.Find "[Code] = '" & rstOutsourceItemMaster.Fields("Code").Value & "'"
    If rstOutsourceItemList.EOF Then
       rstOutsourceItemList.AddNew
       rstOutsourceItemList.Fields("Code").Value = rstOutsourceItemMaster.Fields("Code").Value
    End If
    rstOutsourceItemList.Fields("Name").Value = rstOutsourceItemMaster.Fields("Name").Value
    rstOutsourceItemList.Update
    rstOutsourceItemList.Sort = "Name Asc"
    rstOutsourceItemList.Find "[Code] = '" & rstOutsourceItemMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnOutsourceItemMaster, "OutsourceItemMaster", "Code", "Name", Text2.Text, rstOutsourceItemMaster.Fields("Code").Value, False) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text4.Text, False) Then   'UOM
        SSTab1.Tab = 1: Text4.SetFocus: CheckMandatoryFields = True
    End If
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstOutsourceItemList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
