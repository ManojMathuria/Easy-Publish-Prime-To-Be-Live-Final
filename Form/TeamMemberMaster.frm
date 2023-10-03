VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTeamMemberMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Member Master"
   ClientHeight    =   5160
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   6750
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   5160
      Left            =   15
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
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
      Picture         =   "TeamMemberMaster.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   4930
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
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
         TabPicture(0)   =   "TeamMemberMaster.frx":001C
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
         TabPicture(1)   =   "TeamMemberMaster.frx":0038
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
            TabIndex        =   11
            Top             =   4450
            Width           =   5775
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3930
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   450
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   6932
            _Version        =   393216
            AllowUpdate     =   0   'False
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
               DataField       =   "Designation"
               Caption         =   "Designation"
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
                  ColumnWidth     =   3404.977
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   2280.189
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   2430
            Left            =   -74880
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   4286
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
            Picture         =   "TeamMemberMaster.frx":0054
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
               Left            =   1320
               MaxLength       =   40
               TabIndex        =   6
               Top             =   2000
               Width           =   4815
            End
            Begin VB.TextBox txtReportingToName 
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
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   5
               Top             =   1680
               Width           =   4815
            End
            Begin VB.TextBox txtLoginIdName 
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
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   4
               Top             =   1370
               Width           =   4815
            End
            Begin VB.TextBox txtDesignationName 
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
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   3
               Top             =   1050
               Width           =   4815
            End
            Begin VB.TextBox txtDepartmentName 
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
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   2
               Top             =   740
               Width           =   4815
            End
            Begin VB.TextBox txtPrintName 
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
               TabIndex        =   1
               Top             =   420
               Width           =   4815
            End
            Begin VB.TextBox txtName 
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
               TabIndex        =   0
               Top             =   105
               Width           =   4815
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   14
               Top             =   105
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Picture         =   "TeamMemberMaster.frx":0070
               Picture         =   "TeamMemberMaster.frx":008C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   15
               Top             =   420
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Picture         =   "TeamMemberMaster.frx":00A8
               Picture         =   "TeamMemberMaster.frx":00C4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   120
               TabIndex        =   16
               Top             =   740
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Caption         =   " Department"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "TeamMemberMaster.frx":00E0
               Picture         =   "TeamMemberMaster.frx":00FC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   17
               Top             =   1050
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Caption         =   " Designation"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "TeamMemberMaster.frx":0118
               Picture         =   "TeamMemberMaster.frx":0134
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   18
               Top             =   1370
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Caption         =   " Login Id"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "TeamMemberMaster.frx":0150
               Picture         =   "TeamMemberMaster.frx":016C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   19
               Top             =   1680
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Caption         =   " Reporting To"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "TeamMemberMaster.frx":0188
               Picture         =   "TeamMemberMaster.frx":01A4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   120
               TabIndex        =   21
               Top             =   2000
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Caption         =   " E-Mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "TeamMemberMaster.frx":01C0
               Picture         =   "TeamMemberMaster.frx":01DC
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   1920
            TabIndex        =   20
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
            Picture         =   "TeamMemberMaster.frx":01F8
            Picture         =   "TeamMemberMaster.frx":0214
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
            TabIndex        =   13
            Top             =   4450
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
Attribute VB_Name = "FrmTeamMemberMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SL As Boolean, MasterCode As String
Dim rstTeamMemberList As New ADODB.Recordset, rstTeamMemberMaster As New ADODB.Recordset, rstDepartmentList As New ADODB.Recordset, rstDesignationList As New ADODB.Recordset, rstLoginIdList As New ADODB.Recordset, rstReportingToList As New ADODB.Recordset
Dim DepartmentCode As String, DesignationCode As String, LoginIdCode As String, ReportingToCode As String
Dim SortOrder, PrevStr As String, dblBookMark As Double, blnRecordExist As Boolean
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstTeamMemberList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstTeamMemberList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstTeamMemberList.EOF Or rstTeamMemberList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    rstTeamMemberList.Open "SELECT Name,Code,(Select Name From GeneralMaster Where Code = Designation) As Designation FROM TeamMemberMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockOptimistic
    rstDepartmentList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='13' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstDesignationList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='14' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstLoginIdList.Open "SELECT Name As Col0,Code FROM UserMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstReportingToList.Open "SELECT Name As Col0,Code FROM TeamMemberMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstTeamMemberMaster.CursorLocation = adUseClient
    rstTeamMemberList.Filter = adFilterNone
    Set DataGrid1.DataSource = rstTeamMemberList
    BusySystemIndicator False
    SSTab1.Tab = 0
    If Not (rstTeamMemberList.EOF Or rstTeamMemberList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstTeamMemberList.ActiveConnection = Nothing
    rstDepartmentList.ActiveConnection = Nothing
    rstDesignationList.ActiveConnection = Nothing
    rstLoginIdList.ActiveConnection = Nothing
    rstReportingToList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu
    Text1.SetFocus
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
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
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD And Toolbar1.Buttons.Item(3).Enabled Then
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
           SSTab1.Tab = 1
           SSTab1.SetFocus
        Else
            Sendkeys "{TAB}"
        End If
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    Else
        CloseForm Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstTeamMemberList)
    Call CloseRecordset(rstTeamMemberMaster)
    Call CloseRecordset(rstDepartmentList)
    Call CloseRecordset(rstDesignationList)
    Call CloseRecordset(rstLoginIdList)
    Call CloseRecordset(rstReportingToList)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
On Error Resume Next
With rstTeamMemberList
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
    
    If rstTeamMemberList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstTeamMemberList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstTeamMemberList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstTeamMemberList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstTeamMemberList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstTeamMemberList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstTeamMemberList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstTeamMemberList
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
            If Not (rstTeamMemberList.EOF Or rstTeamMemberList.BOF) Then
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
        txtName.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    
    If Button.Index = 1 Then
        If rstTeamMemberMaster.State = adStateOpen Then rstTeamMemberMaster.Close
        rstTeamMemberMaster.Open "SELECT * FROM TeamMemberMaster WHERE Code=''", cnDatabase, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstTeamMemberMaster) Then
            Call SetButtons(False)
            SSTab1.Tab = 1
            txtName.SetFocus
            blnRecordExist = False
        End If
    ElseIf Button.Index = 2 Then
        If rstTeamMemberList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstTeamMemberList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Master")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnDatabase.Execute "DELETE FROM TeamMemberMaster WHERE Code='" & rstTeamMemberList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstTeamMemberList.Delete
                rstTeamMemberList.MoveNext
                If rstTeamMemberList.RecordCount > 0 And rstTeamMemberList.EOF Then rstTeamMemberList.MoveLast
                Call UpdateUserAction("Team Member Master", "D", Trim(txtName.Text), cnDatabase)
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
        If UpdateRecord(rstTeamMemberMaster) Then
            Call UpdateUserAction("Team Member Master", IIf(blnRecordExist, "M", "A"), Trim(txtName.Text), cnDatabase)
            AddToList
            If rstTeamMemberMaster.State = adStateOpen Then rstTeamMemberMaster.Close
            rstTeamMemberMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstTeamMemberMaster) Then
            If rstTeamMemberMaster.State = adStateOpen Then rstTeamMemberMaster.Close
            rstTeamMemberMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstTeamMemberList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstTeamMemberList): Loop
        Set DataGrid1.DataSource = rstTeamMemberList
        rstTeamMemberList.ActiveConnection = Nothing
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
        If rstTeamMemberList.RecordCount > 0 Then rstTeamMemberList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstTeamMemberList.RecordCount > 0 Then
           rstTeamMemberList.MovePrevious
           If rstTeamMemberList.BOF Then
              rstTeamMemberList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstTeamMemberList.RecordCount > 0 Then
           rstTeamMemberList.MoveNext
           If rstTeamMemberList.EOF Then
              rstTeamMemberList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstTeamMemberList.RecordCount > 0 Then rstTeamMemberList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        CloseForm Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstTeamMemberList.EOF Or rstTeamMemberList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        Text1.SetFocus
    End If
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
    End If
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
    If rstTeamMemberList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    If rstTeamMemberMaster.EOF Or rstTeamMemberMaster.BOF Then Exit Sub
    If CheckEmpty(txtName, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnDatabase, "TeamMemberMaster", "Code", "Name", txtName.Text, rstTeamMemberMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(txtPrintName, False) Then
        txtPrintName.Text = txtName.Text
    End If
End Sub
Private Sub txtDepartmentName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(txtDepartmentName.Text)
        If rstDepartmentList.RecordCount = 0 Then DisplayError ("No Record in Department Master"): txtDepartmentName.SetFocus: Exit Sub Else rstDepartmentList.MoveFirst
        rstDepartmentList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": DepartmentCode = ""
        Call LoadSelectionList(rstDepartmentList, "List of Departments...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(txtDepartmentName, DepartmentCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(DepartmentCode) <> "" Then Sendkeys "{TAB}" Else txtDepartmentName.Text = ""
    End If
End Sub
Private Sub txtDepartmentName_Validate(Cancel As Boolean)
    If CheckEmpty(txtDepartmentName.Text, False) Then Cancel = True
End Sub
Private Sub txtDesignationName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(txtDesignationName.Text)
        If rstDesignationList.RecordCount = 0 Then DisplayError ("No Record in Designation Master"): txtDesignationName.SetFocus: Exit Sub Else rstDesignationList.MoveFirst
        rstDesignationList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": DesignationCode = ""
        Call LoadSelectionList(rstDesignationList, "List of Designations...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(txtDesignationName, DesignationCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(DesignationCode) <> "" Then Sendkeys "{TAB}" Else txtDesignationName.Text = ""
    End If
End Sub
Private Sub txtDesignationName_Validate(Cancel As Boolean)
    If CheckEmpty(txtDesignationName.Text, False) Then Cancel = True
End Sub
Private Sub txtLoginIdName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(txtLoginIdName.Text)
        If rstLoginIdList.RecordCount = 0 Then DisplayError ("No Record in Login Id Master"): txtLoginIdName.SetFocus: Exit Sub Else rstLoginIdList.MoveFirst
        rstLoginIdList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": LoginIdCode = ""
        Call LoadSelectionList(rstLoginIdList, "List of Login Ids...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(txtLoginIdName, LoginIdCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(LoginIdCode) <> "" Then Sendkeys "{TAB}" Else txtLoginIdName.Text = ""
    End If
End Sub
Private Sub txtLoginIdName_Validate(Cancel As Boolean)
    If CheckEmpty(txtLoginIdName.Text, False) Then Cancel = True
End Sub
Private Sub txtReportingToName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(txtReportingToName.Text)
        If rstReportingToList.RecordCount = 0 Then DisplayError ("No Record in Team Member Master"): txtReportingToName.SetFocus: Exit Sub Else rstReportingToList.MoveFirst
        rstReportingToList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": ReportingToCode = ""
        Call LoadSelectionList(rstReportingToList, "List of Team Members...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(txtReportingToName, ReportingToCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(ReportingToCode) <> "" Then Sendkeys "{TAB}" Else txtReportingToName.Text = ""
    ElseIf KeyCode = vbKeyDelete Then
        txtReportingToName.Text = "": ReportingToCode = ""
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstTeamMemberList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstTeamMemberMaster.State = adStateOpen Then
       rstTeamMemberMaster.Close
    End If
    rstTeamMemberMaster.Open "SELECT * FROM TeamMemberMaster WHERE Code='" & FixQuote(rstTeamMemberList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    If rstTeamMemberMaster.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    txtName.Text = ""
    txtPrintName.Text = ""
    txtDepartmentName.Text = ""
    txtDesignationName.Text = ""
    txtLoginIdName.Text = ""
    txtReportingToName.Text = ""
    ReportingToCode = ""
    Text2.Text = ""
End Sub
Private Sub LoadFields()
    If rstTeamMemberMaster.EOF Or rstTeamMemberMaster.BOF Then Exit Sub
    txtName.Text = rstTeamMemberMaster.Fields("Name").Value
    txtPrintName.Text = rstTeamMemberMaster.Fields("PrintName").Value
    DepartmentCode = rstTeamMemberMaster.Fields("Department").Value
    rstDepartmentList.MoveFirst
    rstDepartmentList.Find "[Code] = '" & DepartmentCode & "'"
    txtDepartmentName.Text = rstDepartmentList.Fields("Col0").Value
    DesignationCode = rstTeamMemberMaster.Fields("Designation").Value
    rstDesignationList.MoveFirst
    rstDesignationList.Find "[Code] = '" & DesignationCode & "'"
    txtDesignationName.Text = rstDesignationList.Fields("Col0").Value
    LoginIdCode = rstTeamMemberMaster.Fields("LoginId").Value
    rstLoginIdList.MoveFirst
    rstLoginIdList.Find "[Code] = '" & LoginIdCode & "'"
    txtLoginIdName.Text = rstLoginIdList.Fields("Col0").Value
    ReportingToCode = rstTeamMemberMaster.Fields("ReportingTo").Value
    If rstReportingToList.RecordCount > 0 Then rstReportingToList.MoveFirst
    rstReportingToList.Find "[Code] = '" & ReportingToCode & "'"
    If Not rstReportingToList.EOF Then txtReportingToName.Text = rstReportingToList.Fields("Col0").Value
    Text2.Text = CheckNull(rstTeamMemberMaster.Fields("eMail").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstTeamMemberMaster.RecordCount = 0 Then Exit Sub
    If rstTeamMemberMaster.State = adStateOpen Then rstTeamMemberMaster.Close
    rstTeamMemberMaster.CursorLocation = adUseServer
    rstTeamMemberMaster.Open "SELECT * FROM TeamMemberMaster WHERE Code = '" & FixQuote(rstTeamMemberList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstTeamMemberMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    txtName.SetFocus
    blnRecordExist = True
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstTeamMemberMaster.EOF Or rstTeamMemberMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstTeamMemberMaster.Fields("Code").Value = GenerateCode(cnDatabase, "SELECT MAX(Code) FROM TeamMemberMaster", 6, "0")
        rstTeamMemberMaster.Fields("CreatedBy").Value = UserCode
        rstTeamMemberMaster.Fields("CreatedOn").Value = Now()
        rstTeamMemberMaster.Fields("Recordstatus").Value = "N"
    Else
        rstTeamMemberMaster.Fields("ModifiedBy").Value = UserCode
        rstTeamMemberMaster.Fields("ModifiedOn").Value = Now()
        rstTeamMemberMaster.Fields("Recordstatus").Value = "M"
    End If
    rstTeamMemberMaster.Fields("Name").Value = Trim(txtName.Text)
    rstTeamMemberMaster.Fields("PrintName").Value = Trim(txtPrintName.Text)
    rstTeamMemberMaster.Fields("Department").Value = DepartmentCode
    rstTeamMemberMaster.Fields("Designation").Value = DesignationCode
    rstTeamMemberMaster.Fields("LoginId").Value = LoginIdCode
    rstTeamMemberMaster.Fields("ReportingTo").Value = ReportingToCode
    rstTeamMemberMaster.Fields("eMail").Value = Text2.Text
    rstTeamMemberMaster.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstTeamMemberList.MoveFirst
    rstTeamMemberList.Find "[Code] = '" & rstTeamMemberMaster.Fields("Code").Value & "'"
    If rstTeamMemberList.EOF Then rstTeamMemberList.AddNew: rstTeamMemberList.Fields("Code").Value = rstTeamMemberMaster.Fields("Code").Value
    rstTeamMemberList.Fields("Name").Value = rstTeamMemberMaster.Fields("Name").Value
    rstTeamMemberList.Update
    rstTeamMemberList.Sort = "Name Asc"
    rstTeamMemberList.Find "[Code] = '" & rstTeamMemberMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(txtName.Text, False) Then
        txtName.SetFocus: CheckMandatoryFields = True
    ElseIf CheckDuplicate(cnDatabase, "TeamMemberMaster", "Code", "Name", txtName.Text, rstTeamMemberMaster.Fields("Code").Value, False) Then
        txtName.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(txtPrintName.Text, False) Then
        txtPrintName.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(txtDepartmentName.Text, False) Then
        txtDepartmentName.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(txtDesignationName.Text, False) Then
        txtDesignationName.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(txtLoginIdName.Text, False) Then
        txtLoginIdName.SetFocus: CheckMandatoryFields = True
    End If
    If Not CheckEmpty(Text2.Text, False) Then If InStr(1, Text2.Text, "@") = 0 Then Text2.SetFocus: CheckMandatoryFields = True
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then Timer1.Enabled = False: ShowProgressInStatusBar False
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstTeamMemberList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
