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
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   390
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
   ScaleHeight     =   5730
   ScaleWidth      =   7590
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   5725
      Left            =   15
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13361
      _ExtentY        =   10098
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
         Height          =   5485
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   9684
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
            TabIndex        =   12
            Top             =   5015
            Width           =   6615
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4495
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   450
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   7938
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
               DataField       =   "UGroupName"
               Caption         =   "Group"
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
                  ColumnWidth     =   6510.047
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   6524.788
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   4680
            Left            =   -74880
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   480
            Width           =   7095
            _Version        =   65536
            _ExtentX        =   12515
            _ExtentY        =   8255
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   120
               TabIndex        =   19
               Top             =   735
               Visible         =   0   'False
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
               Caption         =   " Conversion Units"
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
               TabIndex        =   6
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
               Picture         =   "GeneralMaster.frx":00B4
               Picture         =   "GeneralMaster.frx":00D0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   7
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
               Picture         =   "GeneralMaster.frx":00EC
               Picture         =   "GeneralMaster.frx":0108
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   15
               Top             =   735
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
               Caption         =   " Group (s)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "GeneralMaster.frx":0124
               Picture         =   "GeneralMaster.frx":0140
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
               Left            =   1680
               MaxLength       =   120
               TabIndex        =   2
               Top             =   735
               Visible         =   0   'False
               Width           =   5295
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
               MaxLength       =   60
               TabIndex        =   1
               Top             =   425
               Width           =   5295
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
               MaxLength       =   60
               TabIndex        =   0
               Top             =   100
               Width           =   5295
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
               Height          =   330
               Left            =   1680
               TabIndex        =   4
               ToolTipText     =   "One Color"
               Top             =   735
               Visible         =   0   'False
               Width           =   5295
               _Version        =   65536
               _ExtentX        =   9340
               _ExtentY        =   582
               Calculator      =   "GeneralMaster.frx":015C
               Caption         =   "GeneralMaster.frx":017C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "GeneralMaster.frx":01E8
               Keys            =   "GeneralMaster.frx":0206
               Spin            =   "GeneralMaster.frx":0250
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
               Height          =   3525
               Left            =   120
               TabIndex        =   3
               Top             =   1050
               Visible         =   0   'False
               Width           =   6855
               _Version        =   524288
               _ExtentX        =   12091
               _ExtentY        =   6218
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   3
               ScrollBars      =   2
               SelectBlockOptions=   2
               SpreadDesigner  =   "GeneralMaster.frx":0278
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
               Height          =   330
               Left            =   1680
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   735
               Visible         =   0   'False
               Width           =   5295
               _Version        =   65536
               _ExtentX        =   9340
               _ExtentY        =   582
               _StockProps     =   77
               TintColor       =   16711935
               Alignment       =   0
               AutoSize        =   0   'False
               BevelSize       =   0
               BevelStyle      =   0
               BorderColor     =   -2147483642
               BorderStyle     =   1
               FillColor       =   16777215
               FontStyle       =   0
               FontTransparent =   0   'False
               LightColor      =   -2147483643
               ShadowColor     =   -2147483632
               TextColor       =   -2147483640
               WallPaper       =   0
               NoPrefix        =   0   'False
               FormatString    =   ""
               Caption         =   ""
               Picture         =   "GeneralMaster.frx":07DD
               Begin VB.CheckBox cbValue 
                  Caption         =   "Check1"
                  Height          =   210
                  Left            =   90
                  TabIndex        =   18
                  Top             =   80
                  Width           =   210
               End
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   300
            Index           =   2
            Left            =   2520
            TabIndex        =   16
            Top             =   0
            Width           =   4815
            _Version        =   65536
            _ExtentX        =   8493
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
            Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D/F8->Delete  Ctrl+S->Save"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "GeneralMaster.frx":07F9
            Picture         =   "GeneralMaster.frx":0815
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   300
            Index           =   1
            Left            =   -69840
            TabIndex        =   17
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
            Picture         =   "GeneralMaster.frx":0831
            Picture         =   "GeneralMaster.frx":084D
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
            TabIndex        =   14
            Top             =   5015
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
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
Public oAccountGroup As String
Dim rstGeneralList As New ADODB.Recordset, rstAccountGroup As New ADODB.Recordset, rstGeneralMaster As New ADODB.Recordset, rstchkRef As New ADODB.Recordset
Dim UnderGroupCode As Variant, oKeyCode As Variant
Dim UnderGroup As Variant
Dim SortCol, PrevStr As String, dblBookMark As Double, blnRecordExist As Boolean
Private Sub Form_Load()
    If Not SL Then MasterCode = ""
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    Mh3dFrame2.Height = 1170
    If MasterType = "1" Then Me.Caption = "Size Master"
    If MasterType = "5" Then Me.Caption = "Item Group Master"
    If MasterType = "11" Then Me.Caption = "Finish Size Master"
    If MasterType = "18" Then Me.Caption = "HSN Code Master"
    If MasterType = "25" Then Me.Caption = "General Unit"
    If MasterType = "56" Then Me.Caption = "State Code Master"
    If MasterType = "1201" Then Me.Caption = "Selection List of Accounts"
    
    If MasterType = "1" Then 'Size Master
        Mh3dLabel2.Caption = " Group (s)": Text4.Visible = True: Text4.Locked = False
        DataGrid1.Columns(0).Caption = " Name "
        DataGrid1.Columns(1).Caption = " Under Group "
        DataGrid1.Columns(0).Width = 3525: DataGrid1.Columns(1).Width = 3000
    ElseIf MasterType = "5" Or MasterType = "12" Or MasterType = "26" Or MasterType = "12','26" Then 'Item/Account Group Master
        Mh3dLabel2.Caption = " Under Group ": Text4.Visible = True
        DataGrid1.Columns(0).Caption = " Name "
        DataGrid1.Columns(1).Caption = " Under Group "
        DataGrid1.Columns(0).Width = 3525: DataGrid1.Columns(1).Width = 3000
    ElseIf MasterType = "1201" Then  'List of Accounts
        Mh3dLabel2.Caption = " Under Group "
        DataGrid1.Columns(0).Caption = " Name "
        DataGrid1.Columns(1).Caption = " Under Group "
        DataGrid1.Columns(0).Width = 6525: DataGrid1.Columns(1).Width = 3000
        SSTab1.TabVisible(1) = False
        Mh3dLabel1(2).Visible = False
    ElseIf MasterType = "7" Then 'Operation Master
        Mh3dLabel1(0).Width = 2400: Text2.Left = 2500: Text2.Width = 4470
        Mh3dLabel3.Width = 2400: Text3.Left = 2500: Text3.Width = 4470
        Mh3dLabel2.Caption = " Don't Use No in Calculation": Mh3dFrame4.Visible = True
        Mh3dLabel2.Width = 2400: Mh3dFrame4.Left = 2500: Mh3dFrame4.Width = 4470 '5295
        Text4.Visible = False: Text4.Locked = True
    ElseIf MasterType = "15" Or MasterType = "20" Or MasterType = "23" Or MasterType = "25" Or MasterType = "56" Then  'Paper Unit/Calc Mode/Color Master
        Mh3dLabel2.Caption = IIf(MasterType = "25", "Unit Qty.", IIf(MasterType = "15", " Sheets/Unit", IIf(MasterType = "20", " Value (0 if varies)", " Color"))): MhRealInput1.Visible = True
        DataGrid1.Columns(0).Caption = " Name "
        DataGrid1.Columns(1).Caption = IIf(MasterType = "15" Or MasterType = "25", " Quantity/Unit", IIf(MasterType = "56", " State Code", IIf(MasterType = "20", " Value (0 if varies)", " Color")))
        DataGrid1.Columns(0).Width = 3525: DataGrid1.Columns(1).Width = 3000
        If MasterType = "56" Then Mh3dLabel3.Caption = " Code": MhRealInput1.Visible = False: Mh3dLabel2.Visible = False: Mh3dFrame2.Height = 860 'State Master
        If MasterType = "25" Then
            Mh3dFrame2.Height = 1470
            Mh3dLabel2.Caption = " Conversion Unit"
            Text4.Visible = True
            Text4.Top = 735
            Mh3dLabel4.Visible = True
            Mh3dLabel4.Top = 1035
            Mh3dLabel4.Caption = " Con. Unit Value"
            MhRealInput1.Top = 1035
        End If
    Else
        Mh3dLabel2.Visible = False: Mh3dFrame2.Height = 860
        If MasterType = "56" Then Mh3dLabel3.Caption = " Code" 'State Master
    End If
    CenterForm Me
'    Me.Left = (MdiMainMenu.ScaleWidth - Me.Width) \ 2
    WheelHook DataGrid1
    BusySystemIndicator True
    If cnDatabase.State Then cnDatabase.Close: cnDatabase.Open
    If rstGeneralList.State Then rstGeneralList.Close
    'rstGeneralList.Open "SELECT Name,Code,Value1,UnderGroup As UGroupName,(Select Name From GeneralMaster Where UnderGroup=Code) As UGroupName FROM GeneralMaster WHERE Type IN ('" & IIf(MasterType = 12, "12" & "','" & "26", MasterType) & "') ORDER BY Name", cnDatabase, adOpenKeyset, adLockOptimistic
    If oAccountGroup > "*99999" Then
        rstGeneralList.Open "SELECT G.Name,G.Code,G.Value1,ISNULL(G.UnderGroup,'') As UGroupCode,ISNULL(G1.Name,'') As UGroupName,ISNULL(G1.Value1,0) As UGroupValue1 FROM GeneralMaster G Left Join GeneralMaster G1 on G.UnderGroup=G1.Code WHERE G.Type IN ('" & IIf(MasterType = "12", "12" & "','" & "26", MasterType) & "') AND (G.UnderGroup='" & oAccountGroup & "' OR G.Code='" & oAccountGroup & "') ORDER BY G.Name", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf MasterType = "1201" Then   'Export Rate
        rstGeneralList.Open "SELECT G.Name,G.Code,0 As Value1,ISNULL([Group],'') As UGroupCode,ISNULL(G1.Name,'') As UGroupName,0 As UGroupValue1 FROM AccountMaster G Left Join GeneralMaster G1 on [Group]=G1.Code WHERE G1.Type IN ('12','26') ORDER BY G.Name", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf MasterType = "15" Or MasterType = "20" Or MasterType = "23" Or MasterType = "25" Then  'Paper Unit/Calc Mode/Color Master
        rstGeneralList.Open "SELECT G.Name,G.Code,G.Value1,ISNULL(G.UnderGroup,'') As UGroupCode,ISNULL(G.Value1,'') As UGroupName,ISNULL(G1.Value1,0) As UGroupValue1 FROM GeneralMaster G Left Join GeneralMaster G1 on G.UnderGroup=G1.Code WHERE G.Type IN ('" & IIf(MasterType = "12", "12" & "','" & "26", MasterType) & "') ORDER BY G.Name", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf MasterType = "56" Then  'Paper Unit/Calc Mode/Color Master
        rstGeneralList.Open "SELECT G.Name,G.Code,G.Value1,ISNULL(G.UnderGroup,'') As UGroupCode,ISNULL(G.PrintName,'') As UGroupName,ISNULL(G1.Value1,0) As UGroupValue1 FROM GeneralMaster G Left Join GeneralMaster G1 on G.UnderGroup=G1.Code WHERE G.Type IN ('" & IIf(MasterType = "12", "12" & "','" & "26", MasterType) & "') ORDER BY G.Name", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstGeneralList.Open "SELECT G.Name,G.Code,G.Value1,ISNULL(G.UnderGroup,'') As UGroupCode,ISNULL(G1.Name,'') As UGroupName,ISNULL(G1.Value1,0) As UGroupValue1 FROM GeneralMaster G Left Join GeneralMaster G1 on G.UnderGroup=G1.Code WHERE G.Type IN ('" & IIf(MasterType = "12", "12" & "','" & "26", MasterType) & "') ORDER BY G.Name", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
    rstGeneralMaster.CursorLocation = adUseClient
    rstGeneralList.Filter = adFilterNone
    If rstGeneralList.RecordCount > 0 Then
        rstGeneralList.MoveFirst
        If Not CheckEmpty(MasterCode, False) Then rstGeneralList.Find "[Code]='" & MasterCode & "'"
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
    SortCol = "Name"
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
            Else
                If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then Me.ActiveControl.SetFocus Else Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            End If
        End If
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        slCode = rstGeneralList.Fields("Code").Value
        KeyCode = 0
    ElseIf ((Shift = 0 And KeyCode = vbKeyF8) Or (Shift = vbCtrlMask And KeyCode = vbKeyD)) And Toolbar1.Buttons.Item(3).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2) And Toolbar1.Buttons.Item(4).Enabled Then
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
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstGeneralList.Fields("Code").Value: slName = rstGeneralList.Fields("Name").Value: slValue1 = rstGeneralList.Fields("Value1").Value: slUGroupName = rstGeneralList.Fields("UGroupName").Value: slUGroupCode = rstGeneralList.Fields("UGroupCode").Value: slUGroupValue1 = rstGeneralList.Fields("UGroupValue1").Value: KeyCode = 0: Unload Me: Exit Sub
            Else
                SSTab1.Tab = 1: SSTab1.SetFocus
            End If
        Else
            'Sendkeys "{TAB}"
            If Me.ActiveControl.Name = "Text4" Then
                fpSpread1.GetText 1, fpSpread1.ActiveRow, UnderGroup: If MasterType <> "1" And UnderGroup <> "" Then Text4.Text = UnderGroup: fpSpread1.GetText 3, fpSpread1.ActiveRow, UnderGroupCode
                If MasterType <> "25" Then Sendkeys "{TAB}": Sendkeys "{TAB}": Text4.SetFocus
                If MasterType = "25" Then MhRealInput1.SetFocus
            ElseIf Me.ActiveControl.Name <> "fpSpread1" Then
                Sendkeys "{TAB}":
            End If
        End If
         oKeyCode = KeyCode: KeyCode = 0
        If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1 Else If Me.Tag <> "S" Then slCode = "": slName = "": slValue1 = 0: slUGroupName = "": slUGroupCode = "": slUGroupValue1 = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstGeneralList)
    Call CloseRecordset(rstGeneralMaster)
    Call CloseRecordset(rstchkRef)
    Call CloseRecordset(rstAccountGroup)
    ShowProgressInStatusBar False
    SetMenuOptions True
End Sub
Private Sub Text1_Change()
    On Error Resume Next
    With rstGeneralList
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        If Not CheckEmpty(Text1.Text, False) Then
            .Filter = "[" & SortCol & "] Like '%" & FixQuote(Text1.Text) & "%'"
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
    With rstGeneralList
        If .RecordCount = 0 Then Exit Sub
        If Shift = 0 And KeyCode = vbKeyUp Then
            .MovePrevious
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyBack Then
            .MoveFirst
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyDown Then
            .MoveNext
            If .EOF Then .MoveLast
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
            .MoveFirst
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
            KeyProcessed = True
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
            .MoveLast
            If .EOF Then .MoveLast
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
    End With
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
            MasterType = Left(MasterType, 2)
            If MasterType = 5 Or MasterType = 12 Then Mh3dFrame2.Height = 1170: fpSpread1.Visible = False
            Text1.SetFocus
        End If
        SSTab1.TabEnabled(0) = True
    Else
       SSTab1.TabEnabled(0) = False
       Text2.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean, UpdateFlag As Integer
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
        If chkRef Or Left(rstGeneralList.Fields("Code").Value, 1) = "*" Then
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
        
        UpdateFlag = 0 '
        If UpdateRecord(rstGeneralMaster) Then
            If UpdateSizeList("D") Then '
                UpdateFlag = 1 '
                    If Not CheckEmpty(UnderGroupCode, False) Then '
                        If Not UpdateSizeList("I") Then UpdateFlag = 0
                    End If '
            End If '
        End If '
        If UpdateFlag Then '
            Call UpdateUserAction("Size Group Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), cnDatabase) '
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
        Do Until RefreshRecord(rstGeneralList): Loop
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
On Error Resume Next
    Static SortOrder As String
    SortCol = DataGrid1.Columns(ColIndex).DataField
    If SortOrder = "A" Then
        rstGeneralList.Sort = "[" + SortCol & "] Desc"
        SortOrder = "D"
    Else
        rstGeneralList.Sort = "[" + SortCol & "] Asc"
        SortOrder = "A"
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
    MasterType = Left(MasterType, 2)
        If MasterType <> 56 Then Text3.Text = Text2.Text
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
    rstGeneralMaster.Open "SELECT *,(Select Name From GeneralMaster UG Where UG.Code=G.UnderGroup) AS UnderGroupName FROM GeneralMaster G WHERE Code='" & FixQuote(rstGeneralList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    If rstGeneralMaster.RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    MhRealInput1.Value = 0
    cbValue.Value = 0
End Sub
Private Sub LoadFields()
    With rstGeneralMaster
        If .EOF Or .BOF Then Exit Sub
        Text2.Text = .Fields("Name").Value
        Text3.Text = .Fields("PrintName").Value
        MhRealInput1.Value = .Fields("Value1").Value
        
        If .Fields("Value1").Value > 1 Then cbValue.Value = 1 Else cbValue.Value = .Fields("Value1").Value
        UnderGroup = .Fields("UnderGroupName").Value
        UnderGroupCode = .Fields("UnderGroup").Value
        If MasterType = "1" Then    'Size Master
            With rstchkRef
                If .State = adStateOpen Then .Close
                .Open "SELECT STUFF((SELECT ', '+(LTRIM(M.Name)) FROM SizeGroupChild C INNER JOIN GeneralMaster M ON C.Code=M.Code WHERE [Size]='" & rstGeneralMaster.Fields("Code").Value & "' ORDER BY M.Name FOR XML PATH('')),1,1,'') As Name", cnDatabase, adOpenKeyset, adLockReadOnly
                If .RecordCount > 0 Then Text4.Text = CheckNull(Trim(.Fields("Name").Value))
            End With
        ElseIf MasterType = "12" Or MasterType = "25" Or MasterType = "5" Or MasterType = "12','26" Then  'Account Group Master
            With rstAccountGroup
                If .State = adStateOpen Then rstAccountGroup.Close
                .Open "SELECT M.Name FROM GeneralMaster M WHERE M.Code='" & UnderGroupCode & "' ORDER BY M.Name", cnDatabase, adOpenKeyset, adLockReadOnly
                If .RecordCount > 0 Then Text4.Text = CheckNull(Trim(.Fields("Name").Value))
            End With
        End If
    End With
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstGeneralMaster.RecordCount = 0 Then Exit Sub
    If rstGeneralMaster.State = adStateOpen Then rstGeneralMaster.Close
    rstGeneralMaster.CursorLocation = adUseServer
    rstGeneralMaster.Open "SELECT * ,(Select Name From GeneralMaster UG Where UG.Code=G.UnderGroup) AS UnderGroupName FROM GeneralMaster G WHERE Code='" & FixQuote(rstGeneralList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstGeneralMaster.Fields("PrintStatus") = "N"
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
    If Shift = 0 And KeyCode = vbKeyReturn Then
        With fpSpread1
            .GetText 3, .ActiveRow, UnderGroupCode
            .GetText 1, .ActiveRow, UnderGroup: Text4.Text = UnderGroup
        End With
    End If
End Sub
Private Sub SaveFields()
    With rstGeneralMaster
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnDatabase, "SELECT MAX(Code) FROM GeneralMaster", 6, "0")
            .Fields("CreatedBy").Value = UserCode
            .Fields("CreatedOn").Value = Now()
            .Fields("RecordStatus").Value = "N"
        Else
            .Fields("ModifiedBy").Value = UserCode
            .Fields("ModifiedOn").Value = Now()
            .Fields("RecordStatus").Value = "M"
        End If
        .Fields("Name").Value = Trim(Text2.Text)
        .Fields("PrintName").Value = Trim(Text3.Text)
        .Fields("Type").Value = Left(MasterType, 2)
        .Fields("Value1").Value = IIf(MasterType = "7", cbValue.Value, MhRealInput1.Value)
        .Fields("PrintStatus").Value = "N"
        .Fields("UnderGroup").Value = IIf(Trim(Text4.Text) <> "", UnderGroupCode, "")
    End With
End Sub
Private Sub AddToList()
    On Error Resume Next
    With rstGeneralList
        .MoveFirst
        .Find "[Code] = '" & rstGeneralMaster.Fields("Code").Value & "'"
        If .EOF Then .AddNew: .Fields("Code").Value = rstGeneralMaster.Fields("Code").Value
        .Fields("Name").Value = rstGeneralMaster.Fields("Name").Value
        .Update
        .Sort = "Name Asc"
        .Find "[Code] = '" & rstGeneralMaster.Fields("Code").Value & "'"
    End With
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
    If MhRealInput1.Value = 0 Then
        If MasterType = "15" Then 'Paper Unit Master
            DisplayError ("Sheets/Unit cann't be zero"): MhRealInput1.SetFocus: CheckMandatoryFields = True
        ElseIf MasterType = "23" Then 'Color Master
            DisplayError ("Color cann't be zero"): MhRealInput1.SetFocus: CheckMandatoryFields = True
        ElseIf Trim(Text4.Text) <> "" And MasterType = "25" Then 'Color Master
            DisplayError ("Conversion Units Value cann't be zero"): MhRealInput1.SetFocus: CheckMandatoryFields = True
        End If
    End If
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstGeneralList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
Private Function chkRef() As Boolean
'    On Error GoTo ErrorHandler
'    If rstchkRef.State = adStateOpen Then rstchkRef.Close
'    rstchkRef.Open "SELECT Board FROM BookMaster WHERE Board='" & rstGeneralList.Fields("Code").Value & "'", cnDatabase, adOpenKeyset, adLockReadOnly
'    If rstchkRef.RecordCount > 0 Then chkRef = True: Exit Function
'    If rstchkRef.State = adStateOpen Then rstchkRef.Close
'    rstchkRef.Open "SELECT [Group] FROM BookMaster WHERE [Group]='" & rstGeneralList.Fields("Code").Value & "'", cnDatabase, adOpenKeyset, adLockReadOnly
'    If rstchkRef.RecordCount > 0 Then chkRef = True: Exit Function
'    If rstchkRef.State = adStateOpen Then rstchkRef.Close
'    rstchkRef.Open "SELECT BindingType FROM BookMaster WHERE BindingType='" & rstGeneralList.Fields("Code").Value & "'", cnDatabase, adOpenKeyset, adLockReadOnly
'    If rstchkRef.RecordCount > 0 Then chkRef = True
'    If rstchkRef.State = adStateOpen Then rstchkRef.Close
'    rstchkRef.Open "SELECT LaminationType FROM BookMaster WHERE LaminationType='" & rstGeneralList.Fields("Code").Value & "'", cnDatabase, adOpenKeyset, adLockReadOnly
'    If rstchkRef.RecordCount > 0 Then chkRef = True
'    If rstchkRef.State = adStateOpen Then rstchkRef.Close
'    rstchkRef.Open "SELECT Member FROM BookChild02 WHERE Member='" & rstGeneralList.Fields("Code").Value & "'", cnDatabase, adOpenKeyset, adLockReadOnly
'    If rstchkRef.RecordCount > 0 Then chkRef = True
'    Exit Function
'ErrorHandler:
'    chkRef = True
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
    Dim i As Long, cndn As String
    If Not (MasterType = "1" Or MasterType = "25" Or MasterType = "5" Or MasterType = "7" Or MasterType = "12" Or MasterType = "12','26") Then Exit Sub
    If (Shift = 0 And KeyCode = vbKeySpace) Or (Shift = 0 And KeyCode = vbKeyDown) Or (Shift = 0 And KeyCode = vbKeyReturn) Then
        On Error Resume Next
        Mh3dFrame2.Height = 4680
        fpSpread1.Visible = True
        Screen.MousePointer = vbNormal
        With rstAccountGroup
            If .State = adStateOpen Then .Close
            If MasterType = "5" Then
                cndn = "Type = '" & MasterType & "'"
            ElseIf MasterType = "7" Then
                cndn = "Type = 20"
            ElseIf MasterType = "12" Or MasterType = "12','26" Then
                cndn = "Type IN ('" & MasterType & "','26') AND Code NOT IN ('" & slCode & "','*26001','*26002','*26003')"
            ElseIf MasterType = "1" Then
                cndn = "Type = 10"
            ElseIf MasterType = "25" Then
                cndn = "Type = '" & MasterType & "'"
            End If
    If oAccountGroup > "*99000" Then
            .Open "SELECT Name As Col0,Code,Name,Value1,(SELECT Name FROM GeneralMaster WHERE Code=G.UnderGroup) As UGroupName FROM GeneralMaster G  WHERE (G.UnderGroup='" & oAccountGroup & "' OR G.Code='" & oAccountGroup & "') ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
            .Open "SELECT Name As Col0,Code,Name,Value1,(SELECT Name FROM GeneralMaster WHERE Code=G.UnderGroup) As UGroupName FROM GeneralMaster G  WHERE " & cndn & " ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
            If .RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
        End With
        With fpSpread1
            .ClearRange 1, 1, .MaxCols, .MaxRows, False
            .MaxRows = rstAccountGroup.RecordCount + 1
            If MasterType = "1" Then .ColWidth(1) = 50.5
            rstAccountGroup.MoveFirst
            Do Until rstAccountGroup.EOF
                i = i + 1
                .SetText 1, i, rstAccountGroup.Fields("Name").Value
                .SetText 2, i, rstAccountGroup.Fields("UGroupName").Value
                .SetText 3, i, rstAccountGroup.Fields("Code").Value
                rstAccountGroup.MoveNext
            Loop
            If Text4.Text = "" Then fpSpread1.SetActiveCell 1, 1
        End With
    End If
    Call Text4_Change
End Sub
Private Sub Text4_Change()
    Dim i As Integer, cVal As Variant
    If oKeyCode <> vbKeyReturn Then
    With fpSpread1
        For i = 1 To .DataRowCnt
            .Row = i: .RowHidden = False
        Next
        For i = 1 To .DataRowCnt
            .GetText 1, i, cVal
            If CheckEmpty(Text4.Text, False) Then
                .SetActiveCell 1, 1
            ElseIf InStr(StrConv(cVal, vbUpperCase), StrConv(Trim(Text4.Text), vbUpperCase)) > 0 Then
                '.SetActiveCell 1, i: Exit Sub
                .SetActiveCell 1, i
            ElseIf InStr(StrConv(cVal, vbUpperCase), StrConv(Trim(Text4.Text), vbUpperCase)) < 0 Then
            .Row = i: .RowHidden = True
            End If
        Next
    End With
    End If
    oKeyCode = 0
End Sub
Private Function UpdateSizeList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 4) As Variant
    On Error GoTo ErrorHandler
    UpdateSizeList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        cnDatabase.Execute "DELETE FROM SizeGroupChild WHERE Size='" & rstGeneralMaster.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
'            .GetText 5, .ActiveRow, CellVal(1)
'            .GetText 2, .ActiveRow, CellVal(2)
'            .GetText 3, .ActiveRow, CellVal(3)
'            .GetText 6, .ActiveRow, CellVal(4)
        End With
        cnDatabase.Execute "INSERT INTO SizeGroupChild VALUES ('" & UnderGroupCode & "','" & rstGeneralMaster.Fields("Code").Value & "')"
    End If
    Exit Function
ErrorHandler:
    UpdateSizeList = False
End Function

