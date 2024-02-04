VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBindingTypeMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binding Type Master"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   390
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
      Picture         =   "BindingTypeMaster.frx":0000
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
         TabPicture(0)   =   "BindingTypeMaster.frx":001C
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
         TabPicture(1)   =   "BindingTypeMaster.frx":0038
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
            Height          =   3750
            Left            =   -74880
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   6615
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
            Picture         =   "BindingTypeMaster.frx":0054
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
               Picture         =   "BindingTypeMaster.frx":0070
               Picture         =   "BindingTypeMaster.frx":008C
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
               Picture         =   "BindingTypeMaster.frx":00A8
               Picture         =   "BindingTypeMaster.frx":00C4
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   2690
               Left            =   120
               TabIndex        =   2
               Top             =   960
               Width           =   6030
               _Version        =   524288
               _ExtentX        =   10636
               _ExtentY        =   4745
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
               MaxCols         =   2
               MaxRows         =   1000
               SpreadDesigner  =   "BindingTypeMaster.frx":00E0
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
            Picture         =   "BindingTypeMaster.frx":07C3
            Picture         =   "BindingTypeMaster.frx":07DF
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
Attribute VB_Name = "FrmBindingTypeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SL As Boolean, MasterCode As String 'SL-Selection List, MasterCode-Master to Modify
Dim cnBindingTypeMaster As New ADODB.Connection 'Connection
Dim rstBindingTypeList As New ADODB.Recordset, rstBindingTypeMaster As New ADODB.Recordset, rstBindingTypeChild As New ADODB.Recordset 'Recordset
Dim BinderyProcessCode  As Variant
Dim EditMode As Boolean, PrevStr As String, dblBookMark As Double, blnRecordExist As Boolean
Private Sub Form_Load()
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    If Not SL Then MasterCode = ""
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
'    Me.Top = (MdiMainMenu.ScaleHeight - Me.Height) \ 2 + 1000
 '   Me.Left = (MdiMainMenu.ScaleWidth - Me.Width) \ 2
    BusySystemIndicator True
    cnBindingTypeMaster.CursorLocation = adUseClient: cnBindingTypeMaster.Open cnDatabase.ConnectionString
    rstBindingTypeList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type='6' ORDER BY Name", cnDatabase, adOpenKeyset, adLockOptimistic
    rstBindingTypeMaster.CursorLocation = adUseClient
    rstBindingTypeList.Filter = adFilterNone
    If rstBindingTypeList.RecordCount > 0 Then
        If CheckEmpty(MasterCode, False) Then
            rstBindingTypeList.MoveFirst
        Else
            rstBindingTypeList.MoveFirst
            rstBindingTypeList.Find "[Code]='" & MasterCode & "'"
        End If
    End If
    Set DataGrid1.DataSource = rstBindingTypeList
    SSTab1.Tab = 0
    If Not (rstBindingTypeList.EOF Or rstBindingTypeList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstBindingTypeList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    MdiMainMenu.mnuBindingTypeMaster.Enabled = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        EditMode = False
        If SSTab1.Tab = 0 Then
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If InStr(1, "fpSpread1", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True
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
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF8)) And Toolbar1.Buttons.Item(3).Enabled Then 'Delete
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2)) And Toolbar1.Buttons.Item(4).Enabled Then 'Save
        EditMode = False
        If InStr(1, "fpSpread1", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True
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
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstBindingTypeList.Fields("Code").Value: slName = rstBindingTypeList.Fields("Name").Value: KeyCode = 0: Unload Me: Exit Sub
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
    Call CloseRecordset(rstBindingTypeList)
    Call CloseRecordset(rstBindingTypeMaster)
    Call CloseRecordset(rstBindingTypeChild)
    Call CloseConnection(cnBindingTypeMaster)
    ShowProgressInStatusBar False
    MdiMainMenu.mnuBindingTypeMaster.Enabled = True
End Sub
Private Sub Text1_Change()
    If rstBindingTypeList.RecordCount = 0 Then Exit Sub
    rstBindingTypeList.MoveFirst
    If Len(Text1.Text) > 0 Then
        rstBindingTypeList.Filter = "[Name] Like '%" & FixQuote(Text1.Text) & "%'"
        If rstBindingTypeList.EOF Then  'if Spelling mistake
            rstBindingTypeList.Filter = adFilterNone
            rstBindingTypeList.MoveFirst
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            Sendkeys "{End}"
        Else    'if Spelling alright
            PrevStr = Text1.Text
        End If
    Else
        rstBindingTypeList.Filter = adFilterNone
        rstBindingTypeList.MoveFirst
        Set DataGrid1.DataSource = rstBindingTypeList
        PrevStr = ""
    End If
    If Not (rstBindingTypeList.EOF Or rstBindingTypeList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    If rstBindingTypeList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstBindingTypeList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstBindingTypeList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstBindingTypeList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstBindingTypeList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstBindingTypeList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstBindingTypeList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstBindingTypeList
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
            If Not (rstBindingTypeList.EOF Or rstBindingTypeList.BOF) Then
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
    Dim UpdateFlag As Integer, i As Integer
    Dim HiLiteRecord As Boolean
    If Button.Index = 1 Then
        If rstBindingTypeMaster.State = adStateOpen Then rstBindingTypeMaster.Close
        rstBindingTypeMaster.Open "SELECT * FROM GeneralMaster WHERE Code=''", cnDatabase, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstBindingTypeMaster) Then
           Call SetButtons(False)
           SSTab1.Tab = 1
           Text2.SetFocus
           blnRecordExist = False
           cnBindingTypeMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstBindingTypeList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstBindingTypeList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Master"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnDatabase.Execute "DELETE FROM GeneralMaster WHERE Code = '" & rstBindingTypeList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstBindingTypeList.Delete
                rstBindingTypeList.MoveNext
                If rstBindingTypeList.RecordCount > 0 And rstBindingTypeList.EOF Then rstBindingTypeList.MoveLast
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
        If UpdateRecord(rstBindingTypeMaster) Then
            If UpdateBinderyProcessList("D") Then
                UpdateFlag = 1
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.SetActiveCell 1, i
                    fpSpread1.GetText 2, i, BinderyProcessCode
                    If Not CheckEmpty(BinderyProcessCode, False) Then
                        If Not UpdateBinderyProcessList("I") Then UpdateFlag = 0: Exit For
                    End If
                Next
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnBindingTypeMaster.CommitTrans
            If rstBindingTypeMaster.State = adStateOpen Then rstBindingTypeMaster.Close
            rstBindingTypeMaster.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstBindingTypeMaster) Then
            cnBindingTypeMaster.RollbackTrans
            If rstBindingTypeMaster.State = adStateOpen Then rstBindingTypeMaster.Close
            rstBindingTypeMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        RefreshData rstBindingTypeList
        Set DataGrid1.DataSource = rstBindingTypeList
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
        If rstBindingTypeList.RecordCount > 0 Then rstBindingTypeList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstBindingTypeList.RecordCount > 0 Then
           rstBindingTypeList.MovePrevious
           If rstBindingTypeList.BOF Then rstBindingTypeList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstBindingTypeList.RecordCount > 0 Then
           rstBindingTypeList.MoveNext
           If rstBindingTypeList.EOF Then rstBindingTypeList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstBindingTypeList.RecordCount > 0 Then rstBindingTypeList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstBindingTypeList.EOF Or rstBindingTypeList.BOF) Then
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
    If rstBindingTypeList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstBindingTypeMaster.EOF Or rstBindingTypeMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnDatabase, "GeneralMaster", "Code", "Name+Type", Trim(Text2.Text) & "6", rstBindingTypeMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = Text2.Text
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    With fpSpread1
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf .ActiveCol = 1 Then
            If KeyCode = vbKeySpace Then
                With fpSpread1
                    .GetText .ActiveCol + 1, .ActiveRow, BinderyProcessCode
                    On Error Resume Next
                    FrmGeneralMaster.SL = True
                    FrmGeneralMaster.MasterType = "7"
                    FrmGeneralMaster.MasterCode = BinderyProcessCode
                    Load FrmGeneralMaster
                    If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                    On Error GoTo 0
                    .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 1, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                End With
            End If
        End If
    End With
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstBindingTypeList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstBindingTypeMaster.State = adStateOpen Then rstBindingTypeMaster.Close
    rstBindingTypeMaster.Open "SELECT * FROM GeneralMaster WHERE Code='" & FixQuote(rstBindingTypeList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    If rstBindingTypeMaster.RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    fpSpread1.SetActiveCell 1, 1
End Sub
Private Sub LoadFields()
    If rstBindingTypeMaster.EOF Or rstBindingTypeMaster.BOF Then Exit Sub
    Text2.Text = rstBindingTypeMaster.Fields("Name").Value
    Text3.Text = rstBindingTypeMaster.Fields("PrintName").Value
    Call LoadBinderyProcessList(rstBindingTypeMaster.Fields("Code").Value)
End Sub
Private Sub LoadBinderyProcessList(ByVal strBinderyProcessCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstBindingTypeChild.State = adStateOpen Then rstBindingTypeChild.Close
    rstBindingTypeChild.Open "SELECT Name As BinderyProcessName,BinderyProcess  As BinderyProcessCode FROM BindingTypeChild C INNER JOIN GeneralMaster M ON C.BinderyProcess=M.Code WHERE C.Code='" & strBinderyProcessCode & "' ORDER BY M.Name", cnBindingTypeMaster, adOpenKeyset, adLockReadOnly
    rstBindingTypeChild.ActiveConnection = Nothing
    If rstBindingTypeChild.RecordCount > 0 Then rstBindingTypeChild.MoveFirst
    i = 0
    Do While Not rstBindingTypeChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstBindingTypeChild.Fields("BinderyProcessName").Value
            .SetText 2, i, rstBindingTypeChild.Fields("BinderyProcessCode").Value
        End With
        rstBindingTypeChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Bindery Process List")
End Sub
Private Function UpdateBinderyProcessList(ByVal ActionType As String) As Boolean
    Dim CellVal(1) As Variant
    On Error GoTo ErrorHandler
    UpdateBinderyProcessList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        cnBindingTypeMaster.Execute "DELETE FROM BindingTypeChild WHERE Code='" & rstBindingTypeMaster.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 2, .ActiveRow, CellVal(1)
        End With
        cnBindingTypeMaster.Execute "INSERT INTO BindingTypeChild VALUES ('" & rstBindingTypeMaster.Fields("Code").Value & "','" & CellVal(1) & "')"
    End If
    Exit Function
ErrorHandler:
    UpdateBinderyProcessList = False
End Function
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstBindingTypeMaster.RecordCount = 0 Then Exit Sub
    If rstBindingTypeMaster.State = adStateOpen Then rstBindingTypeMaster.Close
    rstBindingTypeMaster.CursorLocation = adUseServer
    rstBindingTypeMaster.Open "SELECT * FROM GeneralMaster WHERE Code='" & FixQuote(rstBindingTypeList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstBindingTypeMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    cnBindingTypeMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstBindingTypeMaster.EOF Or rstBindingTypeMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstBindingTypeMaster.Fields("Code").Value = GenerateCode(cnDatabase, "SELECT MAX(Code) FROM GeneralMaster", 6, "0")
        rstBindingTypeMaster.Fields("CreatedBy").Value = UserCode
        rstBindingTypeMaster.Fields("CreatedOn").Value = Now()
        rstBindingTypeMaster.Fields("Recordstatus").Value = "N"
    Else
        rstBindingTypeMaster.Fields("ModifiedBy").Value = UserCode
        rstBindingTypeMaster.Fields("ModifiedOn").Value = Now()
        rstBindingTypeMaster.Fields("Recordstatus").Value = "M"
    End If
    rstBindingTypeMaster.Fields("Name").Value = Trim(Text2.Text)
    rstBindingTypeMaster.Fields("PrintName").Value = Trim(Text3.Text)
    rstBindingTypeMaster.Fields("Type").Value = "6"
    rstBindingTypeMaster.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstBindingTypeList.MoveFirst
    rstBindingTypeList.Find "[Code] = '" & rstBindingTypeMaster.Fields("Code").Value & "'"
    If rstBindingTypeList.EOF Then rstBindingTypeList.AddNew: rstBindingTypeList.Fields("Code").Value = rstBindingTypeMaster.Fields("Code").Value
    rstBindingTypeList.Fields("Name").Value = rstBindingTypeMaster.Fields("Name").Value
    rstBindingTypeList.Update
    rstBindingTypeList.Sort = "Name Asc"
    rstBindingTypeList.Find "[Code] = '" & rstBindingTypeMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        Text2.SetFocus: CheckMandatoryFields = True
    ElseIf CheckDuplicate(cnDatabase, "GeneralMaster", "Code", "Name+Type", Trim(Text2.Text) & "6", rstBindingTypeMaster.Fields("Code").Value, False) Then
        Text2.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus: CheckMandatoryFields = True
    End If
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstBindingTypeList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then Timer1.Enabled = False: ShowProgressInStatusBar False
End Sub
