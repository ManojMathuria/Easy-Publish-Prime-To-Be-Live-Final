VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGeneralMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Master"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GeneralMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7590
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   5595
      Left            =   15
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13361
      _ExtentY        =   9869
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
      Picture         =   "GeneralMaster.frx":000C
      Begin TabDlg.SSTab SSTab1 
         Height          =   5355
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   9446
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
         TabPicture(0)   =   "GeneralMaster.frx":0028
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
         TabPicture(1)   =   "GeneralMaster.frx":0044
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dLabel1(1)"
         Tab(1).Control(1)=   "Mh3dFrame2"
         Tab(1).ControlCount=   2
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
            TabIndex        =   11
            Top             =   4875
            Width           =   6615
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4485
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   330
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   7911
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
                  ColumnWidth     =   6494.74
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   4770
            Left            =   -74880
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   480
            Width           =   7095
            _Version        =   65536
            _ExtentX        =   12515
            _ExtentY        =   8414
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
            Picture         =   "GeneralMaster.frx":0060
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
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   120
               TabIndex        =   2
               Top             =   740
               Width           =   5535
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
               Left            =   1440
               MaxLength       =   60
               TabIndex        =   1
               Top             =   425
               Width           =   5535
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
               MaxLength       =   60
               TabIndex        =   0
               Top             =   100
               Width           =   5535
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   6
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
               Picture         =   "GeneralMaster.frx":007C
               Picture         =   "GeneralMaster.frx":0098
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   5
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
               Picture         =   "GeneralMaster.frx":00B4
               Picture         =   "GeneralMaster.frx":00D0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   14
               Top             =   740
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
               Caption         =   " Group (s)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "GeneralMaster.frx":00EC
               Picture         =   "GeneralMaster.frx":0108
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
               Height          =   330
               Left            =   1440
               TabIndex        =   4
               ToolTipText     =   "One Color"
               Top             =   740
               Width           =   5535
               _Version        =   65536
               _ExtentX        =   9763
               _ExtentY        =   582
               Calculator      =   "GeneralMaster.frx":0124
               Caption         =   "GeneralMaster.frx":0144
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "GeneralMaster.frx":01B0
               Keys            =   "GeneralMaster.frx":01CE
               Spin            =   "GeneralMaster.frx":0218
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###########0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###########0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   3600
               Left            =   120
               TabIndex        =   3
               Top             =   1050
               Visible         =   0   'False
               Width           =   6855
               _Version        =   524288
               _ExtentX        =   12091
               _ExtentY        =   6350
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   4
               ScrollBars      =   2
               SelectBlockOptions=   2
               SpreadDesigner  =   "GeneralMaster.frx":0240
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   300
            Index           =   2
            Left            =   2760
            TabIndex        =   15
            Top             =   0
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
            _ExtentY        =   529
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
            Picture         =   "GeneralMaster.frx":07D7
            Picture         =   "GeneralMaster.frx":07F3
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   300
            Index           =   1
            Left            =   -69840
            TabIndex        =   16
            Top             =   0
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   529
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
            Caption         =   " Ctrl+E->Edit  Ctrl+S->Save"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "GeneralMaster.frx":080F
            Picture         =   "GeneralMaster.frx":082B
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
            TabIndex        =   13
            Top             =   4875
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7590
      _ExtentX        =   13388
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
Attribute VB_Name = "FrmGeneralMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SL As Boolean 'Selection List
Public MasterCode As String  'Master to Modify
Public MasterType As String
Dim rstGeneralList As New ADODB.Recordset
Dim rstAccountGroup As New ADODB.Recordset
Dim rstGeneralMaster As New ADODB.Recordset
Dim rstCheckRef As New ADODB.Recordset
Dim SortOrder, PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim UnderGroupCode As Variant
Dim UnderGroup As Variant, LR As Integer
Private Sub Form_Load()
    If Not SL Then MasterCode = ""
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    If MasterType <> "1" And MasterType <> "5" And MasterType <> "12" And MasterType <> "15" Then Mh3dFrame2.Height = 860: Text4.Visible = False: MhRealInput1.Visible = False: Mh3dLabel2.Visible = False
    If MasterType = "1" Then MhRealInput1.Visible = False: Mh3dLabel2.Caption = " Group (s)": Mh3dFrame2.Height = 1170: fpSpread1.Visible = False
    If MasterType = "15" Then Text4.Visible = False: Mh3dLabel2.Caption = " Sheets/Unit": Mh3dFrame2.Height = 1170: fpSpread1.Visible = False: Mh3dLabel2.Visible = True
    If MasterType = "5" Or MasterType = "12" Then MhRealInput1.Visible = False: Mh3dLabel2.Caption = " Under Group ": Mh3dFrame2.Height = 1170: fpSpread1.Visible = False: Text4.Locked = False
    
    CenterForm Me
    BusySystemIndicator True
    If rstGeneralList.State Then rstGeneralList.Close
    rstGeneralList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type IN ('" & IIf(MasterType = 12, "12" & "','" & "26", MasterType) & "') ORDER BY Name", cnDatabase, adOpenKeyset, adLockOptimistic
    rstGeneralMaster.CursorLocation = adUseClient
    rstGeneralList.Filter = adFilterNone
    rstGeneralList.Filter = adFilterNone
    If rstGeneralList.RecordCount Then
        If CheckEmpty(MasterCode, False) Then
            rstGeneralList.MoveFirst
        Else
            rstGeneralList.MoveFirst
            rstGeneralList.Find "[Code]='" & MasterCode & "'"
        End If
    End If
    Set DataGrid1.DataSource = rstGeneralList
    BusySystemIndicator False
    SSTab1.Tab = 0
    If Not (rstGeneralList.EOF Or rstGeneralList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstGeneralList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    SortOrder = "Name"
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    SetMenuOptions False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If SSTab1.Tab = 0 Then
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
                Mh3dFrame2.Height = 1170
                fpSpread1.Visible = False
            Else
                If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then Me.ActiveControl.SetFocus Else Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                Mh3dFrame2.Height = 1170
                fpSpread1.Visible = False
            End If
        End If
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        slCode = rstGeneralList.Fields("Code").Value: slName = rstGeneralList.Fields("Name").Value: KeyCode = 0:
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF8 And Toolbar1.Buttons.Item(3).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        Mh3dFrame2.Height = 1170
        fpSpread1.Visible = False
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
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstGeneralList.Fields("Code").Value: slName = rstGeneralList.Fields("Name").Value: KeyCode = 0: Unload Me: Exit Sub
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
    Call CloseRecordset(rstGeneralList)
    Call CloseRecordset(rstGeneralMaster)
    Call CloseRecordset(rstCheckRef)
    Call CloseRecordset(rstAccountGroup)
    ShowProgressInStatusBar False
    SetMenuOptions True
End Sub

Private Sub Text1_Change()
    On Error Resume Next
With rstGeneralList
        If .RecordCount = 0 Then Exit Sub
        If SortOrder = "" Then SortOrder = "Name"
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
    If rstGeneralList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstGeneralList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstGeneralList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstGeneralList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstGeneralList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstGeneralList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstGeneralList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstGeneralList
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
            If Not (rstGeneralList.EOF Or rstGeneralList.BOF) Then
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
    Dim HiLiteRecord As Boolean
    If Button.Index = 1 Then
        If rstGeneralMaster.State = adStateOpen Then rstGeneralMaster.Close
        rstGeneralMaster.Open "SELECT * FROM GeneralMaster WHERE Code=''", cnDatabase, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstGeneralMaster) Then
           Call SetButtons(False)
           SSTab1.Tab = 1
           Text2.SetFocus
           blnRecordExist = False
        End If
    ElseIf Button.Index = 2 Then
        If rstGeneralList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstGeneralList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Master"): Exit Sub
        SSTab1.Tab = 1
        If CheckRef Or Left(rstGeneralList.Fields("Code").Value, 1) = "*" Then
            DisplayError ("Failed to delete the record")
        ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnDatabase.Execute "DELETE FROM GeneralMaster WHERE Code = '" & rstGeneralList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstGeneralList.Delete
                rstGeneralList.MoveNext
                If rstGeneralList.RecordCount > 0 And rstGeneralList.EOF Then rstGeneralList.MoveLast
'                Call UpdateUserAction(Choose(Val(MasterType), "Size", "", "", "", "Item Group", "Binding Type", "Lamination Type", "Correction Team", "Output Type", "", "", "Account Group", "Department", "Designation", "Paper Unit", "", "Billing Narration", "HSN Code") + " Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), cnDatabase)
                ShowProgressInStatusBar True
                Timer1.Enabled = True
            Else
                DisplayError (Err.Description)
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
        If UpdateRecord(rstGeneralMaster) Then
'            Call UpdateUserAction(Choose(Val(MasterType), "Size", "", "", "", "Item Group", "Binding Type", "Lamination Type", "Correction Team", "Output Type", "", "", "Account Group", "Department", "Designation", "Paper Unit", "", "Billing Narration", "HSN Code") + " Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), cnDatabase)
            AddToList
            If rstGeneralMaster.State = adStateOpen Then rstGeneralMaster.Close
            rstGeneralMaster.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstGeneralMaster) Then
           If rstGeneralMaster.State = adStateOpen Then rstGeneralMaster.Close
           rstGeneralMaster.CursorLocation = adUseClient
           Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstGeneralList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstGeneralList): Loop
        Set DataGrid1.DataSource = rstGeneralList
        rstGeneralList.ActiveConnection = Nothing
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
        If rstGeneralList.RecordCount > 0 Then rstGeneralList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstGeneralList.RecordCount > 0 Then
           rstGeneralList.MovePrevious
           If rstGeneralList.BOF Then rstGeneralList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstGeneralList.RecordCount > 0 Then
           rstGeneralList.MoveNext
           If rstGeneralList.EOF Then rstGeneralList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstGeneralList.RecordCount > 0 Then rstGeneralList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstGeneralList.EOF Or rstGeneralList.BOF) Then
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
    If rstGeneralList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstGeneralList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstGeneralList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstGeneralList.EOF Or rstGeneralList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstGeneralMaster.EOF Or rstGeneralMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnDatabase, "GeneralMaster", "Code", "Name+Type", Trim(Text2.Text) & MasterType, rstGeneralMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = Text2.Text
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstGeneralList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstGeneralMaster.State = adStateOpen Then rstGeneralMaster.Close
    rstGeneralMaster.Open "SELECT * FROM GeneralMaster WHERE Code='" & FixQuote(rstGeneralList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    If rstGeneralMaster.RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    MhRealInput1.Value = 0
End Sub
Private Sub LoadFields()
    If rstGeneralMaster.EOF Or rstGeneralMaster.BOF Then Exit Sub
    Text2.Text = rstGeneralMaster.Fields("Name").Value
    Text3.Text = rstGeneralMaster.Fields("PrintName").Value
    MhRealInput1.Value = rstGeneralMaster.Fields("Value1").Value
    UnderGroupCode = rstGeneralMaster.Fields("UnderGroup").Value
    If MasterType = "1" Then    'Size Master
        With rstCheckRef
            If .State = adStateOpen Then .Close
            .Open "SELECT M.Name FROM SizeGroupChild C INNER JOIN GeneralMaster M ON C.Code=M.Code WHERE [Size]='" & rstGeneralMaster.Fields("Code").Value & "' ORDER BY M.Name", cnDatabase, adOpenKeyset, adLockReadOnly
            If .RecordCount > 0 Then
                Do While Not .EOF
                    Text4.Text = Text4.Text & IIf(Text4.Text = "", "", ", ") & .Fields("Name").Value
                    .MoveNext
                Loop
            End If
        End With
    End If
    If MasterType = "12" Or MasterType = "5" Then    'Account Group Master
        With rstAccountGroup
            If .State = adStateOpen Then rstAccountGroup.Close 'Else rstAccountGroup.Open
            .Open "SELECT M.Name FROM GeneralMaster M WHERE M.Code='" & UnderGroupCode & "' ORDER BY M.Name", cnDatabase, adOpenKeyset, adLockReadOnly
            If .RecordCount > 0 Then
                Do While Not .EOF
                    Text4.Text = .Fields("Name").Value
                    .MoveNext
                Loop
            End If
        End With
    End If
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstGeneralMaster.RecordCount = 0 Then Exit Sub
    If rstGeneralMaster.State = adStateOpen Then rstGeneralMaster.Close
    rstGeneralMaster.CursorLocation = adUseServer
    rstGeneralMaster.Open "SELECT * FROM GeneralMaster WHERE Code='" & FixQuote(rstGeneralList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstGeneralMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer) 'Item And Account Group
UnderGroup = ""
If Shift = 0 And KeyCode = 144 Then
With fpSpread1
.GetText 1, .ActiveRow, UnderGroup: Text4.Text = UnderGroup
.GetText 4, .ActiveRow, UnderGroupCode
End With
End If
End Sub
Private Sub SaveFields()
    If rstGeneralMaster.EOF Or rstGeneralMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstGeneralMaster.Fields("Code").Value = GenerateCode(cnDatabase, "SELECT MAX(Code) FROM GeneralMaster", 6, "0")
        rstGeneralMaster.Fields("CreatedBy").Value = UserCode
        rstGeneralMaster.Fields("CreatedOn").Value = Now()
        rstGeneralMaster.Fields("Recordstatus").Value = "N"
    Else
        rstGeneralMaster.Fields("ModifiedBy").Value = UserCode
        rstGeneralMaster.Fields("ModifiedOn").Value = Now()
        rstGeneralMaster.Fields("Recordstatus").Value = "M"
    End If
    rstGeneralMaster.Fields("Name").Value = Trim(Text2.Text)
    rstGeneralMaster.Fields("PrintName").Value = Trim(Text3.Text)
    rstGeneralMaster.Fields("Type").Value = MasterType
    rstGeneralMaster.Fields("Value1").Value = MhRealInput1.Value
    rstGeneralMaster.Fields("PrintStatus").Value = "N"
    rstGeneralMaster.Fields("UnderGroup").Value = UnderGroupCode
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstGeneralList.MoveFirst
    rstGeneralList.Find "[Code] = '" & rstGeneralMaster.Fields("Code").Value & "'"
    If rstGeneralList.EOF Then rstGeneralList.AddNew: rstGeneralList.Fields("Code").Value = rstGeneralMaster.Fields("Code").Value
    rstGeneralList.Fields("Name").Value = rstGeneralMaster.Fields("Name").Value
    rstGeneralList.Update
    rstGeneralList.Sort = "Name Asc"
    rstGeneralList.Find "[Code] = '" & rstGeneralMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        Text2.SetFocus: CheckMandatoryFields = True
    ElseIf CheckDuplicate(cnDatabase, "GeneralMaster", "Code", "Name+Type", Trim(Text2.Text) & MasterType, rstGeneralMaster.Fields("Code").Value, False) Then
        Text2.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus: CheckMandatoryFields = True
    End If
    If MasterType = "1" Then
        If InStr(1, Text2.Text, "X") > 0 Then
            If Len(Left(Text2.Text, InStr(1, Text2.Text, "X") - 1)) <> 5 Or Len(Mid(Text2.Text, InStr(1, Text2.Text, "X") + 1, 5)) <> 5 Or (Not IsNumeric(Left(Text2.Text, InStr(1, Text2.Text, "X") - 1))) Or (Not IsNumeric(Mid(Text2.Text, InStr(1, Text2.Text, "X") + 1, 5))) Then
                DisplayError ("Size Format must be 00.00X00.00")
                Text2.SetFocus: CheckMandatoryFields = True
            End If
        Else
            DisplayError ("Size Format must be 00.00X00.00")
            Text2.SetFocus: CheckMandatoryFields = True
        End If
    End If
    If MasterType = "15" Then 'Paper Unit
        If MhRealInput1.Value = 0 Then DisplayError ("Sheets/Unit cann't be zero"): MhRealInput1.SetFocus: CheckMandatoryFields = True
    End If
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstGeneralList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
Private Function CheckRef() As Boolean
    On Error GoTo ErrorHandler
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "SELECT Board FROM BookMaster WHERE Board='" & rstGeneralList.Fields("Code").Value & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True: Exit Function
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "SELECT [Group] FROM BookMaster WHERE [Group]='" & rstGeneralList.Fields("Code").Value & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True: Exit Function
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "SELECT BindingType FROM BookMaster WHERE BindingType='" & rstGeneralList.Fields("Code").Value & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "SELECT LaminationType FROM BookMaster WHERE LaminationType='" & rstGeneralList.Fields("Code").Value & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "SELECT Member FROM BookChild02 WHERE Member='" & rstGeneralList.Fields("Code").Value & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True
    Exit Function
ErrorHandler:
    CheckRef = True
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then Timer1.Enabled = False: ShowProgressInStatusBar False
End Sub
Private Sub SetMenuOptions(bVal As Boolean)
    MdiMainMenu.mnuAccountGroupMaster.Enabled = bVal
    MdiMainMenu.mnuItemGroupMaster.Enabled = bVal
    MdiMainMenu.mnuBindingTypeMaster.Enabled = bVal
    MdiMainMenu.mnuOperationMaster.Enabled = bVal
    MdiMainMenu.mnuSizeMaster.Enabled = bVal
    MdiMainMenu.mnuPaperUnitMaster.Enabled = bVal
    MdiMainMenu.mnuHSNCodeMaster.Enabled = bVal
    MdiMainMenu.mnuBillingNarrationMaster.Enabled = bVal
    MdiMainMenu.mnuProjectManagement(1).Enabled = bVal
    MdiMainMenu.mnuProjectManagement(2).Enabled = bVal
    MdiMainMenu.mnuMachineMaster.Enabled = bVal
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer) 'Item And Account Group
    Dim i As Long
    If MasterType = "12" Or MasterType = "5" Then    'Item And Account Group
        If KeyCode = vbKeySpace Or KeyCode = 40 Or KeyCode = 144 Then
        On Error Resume Next
            Mh3dFrame2.Height = 4770
            fpSpread1.Visible = True
       
        Screen.MousePointer = vbNormal
        If rstAccountGroup.State = adStateOpen Then rstAccountGroup.Close
        If MasterType = "12" Then rstAccountGroup.Open "SELECT Name As Col0, Code,Name,Value1,(Select Name From GeneralMaster Where Code=G.UnderGroup) As UGroup FROM GeneralMaster G  WHERE Type = '12' OR Type = '26' AND Code NOT IN ('" & slCode & "','*26001','*26002','*26003') ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly 'AND Code < ='*99999' AND Code > = '*99001'
         If MasterType = "5" Then rstAccountGroup.Open "SELECT Name As Col0, Code,Name,Value1,(Select Name From GeneralMaster Where Code=G.UnderGroup) As UGroup FROM GeneralMaster G  WHERE Type = '5' OR Type = '5' And Code NOT IN ('','','')ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
         If rstAccountGroup.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
         
         With fpSpread1
         .ClearRange 1, 1, .MaxCols, .MaxRows, False
         .MaxRows = rstAccountGroup.RecordCount + 1
         rstAccountGroup.MoveFirst
         Do While Not rstAccountGroup.EOF
                     i = i + 1
                    .SetText 1, i, rstAccountGroup.Fields("Name").Value
                    .SetText 2, i, IIf(Val(rstAccountGroup.Fields("Value1").Value) = 1, "Y", "N")
                    .SetText 3, i, rstAccountGroup.Fields("UGroup").Value
                    .SetText 4, i, rstAccountGroup.Fields("Code").Value
         rstAccountGroup.MoveNext
        Loop
        i = rstAccountGroup.RecordCount + 1
       If Text4.Text = "" Then fpSpread1.SetActiveCell 1, i
        End With
        End If
    End If
    Call Text4_Change
End Sub
Private Sub Text4_Change() 'Item And Account Group
  Dim i As Integer, cVal As Variant, R As Long
  'On Error Resume Next
    With fpSpread1
            If .DataRowCnt = 0 Then Exit Sub: i = rstAccountGroup.RecordCount + 1: .SetActiveCell 1, i
            fpSpread1.MaxCols = 4
            For i = 1 To .DataRowCnt
                .GetText 1, i, cVal
                If InStr(StrConv(cVal, vbUpperCase), StrConv(Text4.Text, vbUpperCase)) = 0 Then
                
                ElseIf Text4.Text = " " Or Text4.Text = "" Then
                .SetActiveCell 1, rstAccountGroup.RecordCount + 1
                Else
                    .SetActiveCell 1, i
                End If
            Next
    End With
End Sub
