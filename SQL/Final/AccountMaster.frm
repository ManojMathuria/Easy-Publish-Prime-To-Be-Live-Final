VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmAccountMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Master"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13230
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
   ScaleHeight     =   9150
   ScaleWidth      =   13230
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9150
      Left            =   15
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   0
      Width           =   13215
      _Version        =   65536
      _ExtentX        =   23310
      _ExtentY        =   16140
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
      Picture         =   "AccountMaster.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8910
         Left            =   120
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   120
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   15716
         _Version        =   393216
         Style           =   1
         Tabs            =   10
         TabsPerRow      =   8
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
         TabPicture(0)   =   "AccountMaster.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "AccountMaster.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2(0)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&Details"
         TabPicture(2)   =   "AccountMaster.frx":0054
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "&Details"
         TabPicture(3)   =   "AccountMaster.frx":0070
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         TabCaption(4)   =   "&Details"
         TabPicture(4)   =   "AccountMaster.frx":008C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Mh3dFrame2(3)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "&Details"
         TabPicture(5)   =   "AccountMaster.frx":00A8
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Mh3dFrame2(4)"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "&Details"
         TabPicture(6)   =   "AccountMaster.frx":00C4
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Mh3dFrame2(5)"
         Tab(6).Control(0).Enabled=   0   'False
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "&Details"
         TabPicture(7)   =   "AccountMaster.frx":00E0
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Mh3dFrame2(6)"
         Tab(7).Control(0).Enabled=   0   'False
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "&Details"
         TabPicture(8)   =   "AccountMaster.frx":00FC
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "Mh3dFrame2(7)"
         Tab(8).Control(0).Enabled=   0   'False
         Tab(8).ControlCount=   1
         TabCaption(9)   =   "&Op.Bal."
         TabPicture(9)   =   "AccountMaster.frx":0118
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "Mh3dFrame2(8)"
         Tab(9).ControlCount=   1
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
            TabIndex        =   66
            Top             =   8445
            Width           =   12255
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7915
            Left            =   120
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   450
            Width           =   12735
            _ExtentX        =   22463
            _ExtentY        =   13970
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "AccountGroup"
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
            BeginProperty Column01 
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
            BeginProperty Column02 
               DataField       =   "Alias"
               Caption         =   "Alias"
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
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   2564.788
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   7409.764
               EndProperty
               BeginProperty Column02 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   2190.047
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   3065
            Index           =   0
            Left            =   -74880
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   480
            Width           =   12735
            _Version        =   65536
            _ExtentX        =   22463
            _ExtentY        =   5406
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
            Picture         =   "AccountMaster.frx":0134
            Begin VB.TextBox Text10 
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
               Index           =   0
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   3
               Top             =   740
               Width           =   11415
            End
            Begin VB.TextBox Text12 
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
               Index           =   0
               Left            =   10200
               MaxLength       =   40
               TabIndex        =   9
               Top             =   2310
               Width           =   2415
            End
            Begin VB.TextBox Text11 
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
               Index           =   0
               Left            =   1200
               MaxLength       =   80
               TabIndex        =   10
               Top             =   2625
               Width           =   8295
            End
            Begin VB.TextBox Text9 
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
               Index           =   0
               Left            =   10200
               MaxLength       =   40
               TabIndex        =   11
               Top             =   2625
               Width           =   2415
            End
            Begin VB.TextBox Text8 
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
               Index           =   0
               Left            =   1200
               MaxLength       =   40
               TabIndex        =   8
               Top             =   2310
               Width           =   8295
            End
            Begin VB.TextBox Text7 
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
               Index           =   0
               Left            =   1200
               MaxLength       =   40
               TabIndex        =   7
               Top             =   2000
               Width           =   11415
            End
            Begin VB.TextBox Text6 
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
               Index           =   0
               Left            =   1200
               MaxLength       =   40
               TabIndex        =   6
               Top             =   1680
               Width           =   11415
            End
            Begin VB.TextBox Text5 
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
               Index           =   0
               Left            =   1200
               MaxLength       =   40
               TabIndex        =   5
               Top             =   1370
               Width           =   11415
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
               Index           =   0
               Left            =   1200
               MaxLength       =   40
               TabIndex        =   4
               Top             =   1050
               Width           =   11415
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
               Index           =   0
               Left            =   1200
               MaxLength       =   40
               TabIndex        =   0
               Top             =   105
               Width           =   11415
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   61
               Top             =   420
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0150
               Picture         =   "AccountMaster.frx":016C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   60
               Top             =   105
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0188
               Picture         =   "AccountMaster.frx":01A4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1275
               Index           =   0
               Left            =   120
               TabIndex        =   69
               Top             =   1050
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   2249
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":01C0
               Picture         =   "AccountMaster.frx":01DC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   70
               Top             =   2310
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":01F8
               Picture         =   "AccountMaster.frx":0214
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   71
               Top             =   2625
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0230
               Picture         =   "AccountMaster.frx":024C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   0
               Left            =   9480
               TabIndex        =   72
               Top             =   2625
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0268
               Picture         =   "AccountMaster.frx":0284
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   0
               Left            =   9480
               TabIndex        =   73
               Top             =   2310
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":02A0
               Picture         =   "AccountMaster.frx":02BC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   13
               Left            =   9480
               TabIndex        =   110
               Top             =   420
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":02D8
               Picture         =   "AccountMaster.frx":02F4
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
               Index           =   0
               Left            =   1200
               MaxLength       =   40
               TabIndex        =   1
               Top             =   420
               Width           =   8295
            End
            Begin VB.TextBox Text13 
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
               Index           =   0
               Left            =   10200
               MaxLength       =   40
               TabIndex        =   2
               Top             =   420
               Width           =   2415
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   113
               Top             =   740
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
               Caption         =   " Group"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0310
               Picture         =   "AccountMaster.frx":032C
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8325
            Index           =   4
            Left            =   -74880
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   480
            Width           =   12735
            _Version        =   65536
            _ExtentX        =   22463
            _ExtentY        =   14684
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
            Picture         =   "AccountMaster.frx":0348
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
               Index           =   4
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   105
               Width           =   11175
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
               Index           =   4
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   420
               Width           =   8055
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
               Index           =   4
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   740
               Width           =   11175
            End
            Begin VB.TextBox Text5 
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
               Index           =   4
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   1050
               Width           =   11175
            End
            Begin VB.TextBox Text6 
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
               Index           =   4
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   1365
               Width           =   5590
            End
            Begin VB.TextBox Text7 
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
               Index           =   4
               Left            =   7025
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   1365
               Width           =   5590
            End
            Begin VB.TextBox Text8 
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
               Index           =   4
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   1680
               Width           =   5590
            End
            Begin VB.TextBox Text9 
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
               Index           =   4
               Left            =   7740
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   2000
               Width           =   4875
            End
            Begin VB.TextBox Text11 
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
               Index           =   4
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   2000
               Width           =   5590
            End
            Begin VB.TextBox Text12 
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
               Index           =   4
               Left            =   7740
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   1680
               Width           =   4875
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   75
               Top             =   420
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
               Picture         =   "AccountMaster.frx":0364
               Picture         =   "AccountMaster.frx":0380
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   76
               Top             =   105
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
               Picture         =   "AccountMaster.frx":039C
               Picture         =   "AccountMaster.frx":03B8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   960
               Index           =   4
               Left            =   120
               TabIndex        =   77
               Top             =   740
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   1693
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":03D4
               Picture         =   "AccountMaster.frx":03F0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   78
               Top             =   1680
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":040C
               Picture         =   "AccountMaster.frx":0428
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   79
               Top             =   2000
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0444
               Picture         =   "AccountMaster.frx":0460
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   4
               Left            =   7025
               TabIndex        =   80
               Top             =   2000
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":047C
               Picture         =   "AccountMaster.frx":0498
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   4
               Left            =   7025
               TabIndex        =   81
               Top             =   1680
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":04B4
               Picture         =   "AccountMaster.frx":04D0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   11
               Left            =   9480
               TabIndex        =   108
               Top             =   420
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":04EC
               Picture         =   "AccountMaster.frx":0508
            End
            Begin VB.TextBox Text13 
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
               Index           =   4
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   420
               Width           =   2415
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
               Height          =   330
               Left            =   1440
               TabIndex        =   114
               TabStop         =   0   'False
               Top             =   2310
               Width           =   11175
               _Version        =   65536
               _ExtentX        =   19711
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
               Picture         =   "AccountMaster.frx":0524
               Begin VB.CheckBox chkRound 
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
                  Height          =   290
                  Left            =   120
                  TabIndex        =   34
                  Top             =   30
                  Width           =   255
               End
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   115
               Top             =   2310
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
               Caption         =   " Round Off Qty"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0540
               Picture         =   "AccountMaster.frx":055C
            End
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   5385
               Left            =   120
               TabIndex        =   35
               Top             =   2835
               Width           =   12495
               _Version        =   524288
               _ExtentX        =   22040
               _ExtentY        =   9499
               _StockProps     =   64
               ButtonDrawMode  =   1
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
               MaxCols         =   10
               MaxRows         =   100
               SpreadDesigner  =   "AccountMaster.frx":0578
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   12720
               Y1              =   2730
               Y2              =   2730
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8325
            Index           =   6
            Left            =   -74880
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   480
            Width           =   12735
            _Version        =   65536
            _ExtentX        =   22463
            _ExtentY        =   14684
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
            Picture         =   "AccountMaster.frx":0F2F
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
               Index           =   6
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   105
               Width           =   11415
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
               Index           =   6
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   425
               Width           =   8295
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
               Index           =   6
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   735
               Width           =   11415
            End
            Begin VB.TextBox Text5 
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
               Index           =   6
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   1050
               Width           =   11415
            End
            Begin VB.TextBox Text6 
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
               Index           =   6
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   1365
               Width           =   11415
            End
            Begin VB.TextBox Text7 
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
               Index           =   6
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   1680
               Width           =   11415
            End
            Begin VB.TextBox Text8 
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
               Index           =   6
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   2000
               Width           =   8295
            End
            Begin VB.TextBox Text9 
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
               Index           =   6
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   2310
               Width           =   2415
            End
            Begin VB.TextBox Text11 
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
               Index           =   6
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   2310
               Width           =   8295
            End
            Begin VB.TextBox Text12 
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
               Index           =   6
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   2000
               Width           =   2415
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   83
               Top             =   425
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0F4B
               Picture         =   "AccountMaster.frx":0F67
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   84
               Top             =   105
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0F83
               Picture         =   "AccountMaster.frx":0F9F
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   6
               Left            =   120
               TabIndex        =   85
               Top             =   735
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   2240
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0FBB
               Picture         =   "AccountMaster.frx":0FD7
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   86
               Top             =   2000
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":0FF3
               Picture         =   "AccountMaster.frx":100F
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   87
               Top             =   2310
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":102B
               Picture         =   "AccountMaster.frx":1047
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   6
               Left            =   9480
               TabIndex        =   88
               Top             =   2310
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":1063
               Picture         =   "AccountMaster.frx":107F
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   6
               Left            =   9480
               TabIndex        =   89
               Top             =   2000
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":109B
               Picture         =   "AccountMaster.frx":10B7
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   9
               Left            =   9480
               TabIndex        =   107
               Top             =   425
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":10D3
               Picture         =   "AccountMaster.frx":10EF
            End
            Begin VB.TextBox Text13 
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
               Index           =   6
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   425
               Width           =   2415
            End
            Begin FPSpreadADO.fpSpread fpSpread4 
               Height          =   5385
               Left            =   120
               TabIndex        =   47
               Top             =   2835
               Width           =   12495
               _Version        =   524288
               _ExtentX        =   22040
               _ExtentY        =   9499
               _StockProps     =   64
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
               MaxCols         =   9
               MaxRows         =   100
               SpreadDesigner  =   "AccountMaster.frx":110B
            End
            Begin VB.Line Line6 
               X1              =   0
               X2              =   12720
               Y1              =   2730
               Y2              =   2730
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8325
            Index           =   7
            Left            =   -74880
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   480
            Width           =   12735
            _Version        =   65536
            _ExtentX        =   22463
            _ExtentY        =   14684
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
            Picture         =   "AccountMaster.frx":199B
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
               Index           =   7
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   105
               Width           =   11415
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
               Index           =   7
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   425
               Width           =   8295
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
               Index           =   7
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   735
               Width           =   11415
            End
            Begin VB.TextBox Text5 
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
               Index           =   7
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   1050
               Width           =   11415
            End
            Begin VB.TextBox Text6 
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
               Index           =   7
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   1365
               Width           =   11415
            End
            Begin VB.TextBox Text7 
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
               Index           =   7
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   1680
               Width           =   11415
            End
            Begin VB.TextBox Text8 
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
               Index           =   7
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   2000
               Width           =   8295
            End
            Begin VB.TextBox Text9 
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
               Index           =   7
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   2310
               Width           =   2415
            End
            Begin VB.TextBox Text11 
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
               Index           =   7
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   2310
               Width           =   8295
            End
            Begin VB.TextBox Text12 
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
               Index           =   7
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   2000
               Width           =   2415
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   91
               Top             =   425
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":19B7
               Picture         =   "AccountMaster.frx":19D3
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   92
               Top             =   105
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":19EF
               Picture         =   "AccountMaster.frx":1A0B
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   7
               Left            =   120
               TabIndex        =   93
               Top             =   735
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   2240
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":1A27
               Picture         =   "AccountMaster.frx":1A43
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   94
               Top             =   2000
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":1A5F
               Picture         =   "AccountMaster.frx":1A7B
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   95
               Top             =   2310
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":1A97
               Picture         =   "AccountMaster.frx":1AB3
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   7
               Left            =   9480
               TabIndex        =   96
               Top             =   2310
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":1ACF
               Picture         =   "AccountMaster.frx":1AEB
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   7
               Left            =   9480
               TabIndex        =   97
               Top             =   2000
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":1B07
               Picture         =   "AccountMaster.frx":1B23
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   3
               Left            =   9480
               TabIndex        =   106
               Top             =   425
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":1B3F
               Picture         =   "AccountMaster.frx":1B5B
            End
            Begin VB.TextBox Text13 
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
               Index           =   7
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   425
               Width           =   2415
            End
            Begin FPSpreadADO.fpSpread fpSpread3 
               Height          =   5385
               Left            =   120
               TabIndex        =   59
               Top             =   2835
               Width           =   12495
               _Version        =   524288
               _ExtentX        =   22040
               _ExtentY        =   9499
               _StockProps     =   64
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
               MaxCols         =   13
               MaxRows         =   100
               SpreadDesigner  =   "AccountMaster.frx":1B77
            End
            Begin VB.Line Line7 
               X1              =   0
               X2              =   12720
               Y1              =   2730
               Y2              =   2730
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8325
            Index           =   3
            Left            =   -74880
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   480
            Width           =   12735
            _Version        =   65536
            _ExtentX        =   22463
            _ExtentY        =   14684
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
            Picture         =   "AccountMaster.frx":271E
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
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   105
               Width           =   11415
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
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   420
               Width           =   8295
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
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   735
               Width           =   11415
            End
            Begin VB.TextBox Text5 
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
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   1050
               Width           =   11415
            End
            Begin VB.TextBox Text6 
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
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   1365
               Width           =   11415
            End
            Begin VB.TextBox Text7 
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
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   1680
               Width           =   11415
            End
            Begin VB.TextBox Text8 
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
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   2000
               Width           =   8295
            End
            Begin VB.TextBox Text9 
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
               Index           =   3
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   2310
               Width           =   2415
            End
            Begin VB.TextBox Text11 
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
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   2310
               Width           =   8295
            End
            Begin VB.TextBox Text12 
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
               Index           =   3
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   2000
               Width           =   2415
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   99
               Top             =   420
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":273A
               Picture         =   "AccountMaster.frx":2756
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   100
               Top             =   105
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":2772
               Picture         =   "AccountMaster.frx":278E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   1
               Left            =   120
               TabIndex        =   101
               Top             =   735
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   2240
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":27AA
               Picture         =   "AccountMaster.frx":27C6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   102
               Top             =   2000
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":27E2
               Picture         =   "AccountMaster.frx":27FE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   103
               Top             =   2310
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":281A
               Picture         =   "AccountMaster.frx":2836
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   1
               Left            =   9480
               TabIndex        =   104
               Top             =   2310
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":2852
               Picture         =   "AccountMaster.frx":286E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   1
               Left            =   9480
               TabIndex        =   105
               Top             =   2000
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":288A
               Picture         =   "AccountMaster.frx":28A6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   12
               Left            =   9480
               TabIndex        =   109
               Top             =   420
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":28C2
               Picture         =   "AccountMaster.frx":28DE
            End
            Begin VB.TextBox Text13 
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
               Index           =   3
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   420
               Width           =   2415
            End
            Begin FPSpreadADO.fpSpread fpSpread6 
               Height          =   5385
               Left            =   120
               TabIndex        =   138
               Top             =   2835
               Width           =   12495
               _Version        =   524288
               _ExtentX        =   22040
               _ExtentY        =   9499
               _StockProps     =   64
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
               MaxCols         =   7
               MaxRows         =   100
               SpreadDesigner  =   "AccountMaster.frx":28FA
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   12720
               Y1              =   2730
               Y2              =   2730
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8300
            Index           =   8
            Left            =   -74880
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   480
            Width           =   12735
            _Version        =   65536
            _ExtentX        =   22463
            _ExtentY        =   14640
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
            Picture         =   "AccountMaster.frx":31DE
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   8085
               Left            =   120
               TabIndex        =   112
               Top             =   105
               Width           =   12495
               _Version        =   524288
               _ExtentX        =   22040
               _ExtentY        =   14261
               _StockProps     =   64
               ButtonDrawMode  =   1
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
               MaxCols         =   5
               MaxRows         =   100
               SpreadDesigner  =   "AccountMaster.frx":31FA
            End
            Begin VB.TextBox Text99 
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
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   137
               TabStop         =   0   'False
               Top             =   3870
               Width           =   5775
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8325
            Index           =   5
            Left            =   -74880
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   480
            Width           =   12735
            _Version        =   65536
            _ExtentX        =   22463
            _ExtentY        =   14684
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
            Picture         =   "AccountMaster.frx":38D8
            Begin FPSpreadADO.fpSpread fpSpread5 
               Height          =   5385
               Left            =   120
               TabIndex        =   125
               Top             =   2835
               Width           =   12495
               _Version        =   524288
               _ExtentX        =   22040
               _ExtentY        =   9499
               _StockProps     =   64
               ButtonDrawMode  =   1
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
               MaxRows         =   100
               SpreadDesigner  =   "AccountMaster.frx":38F4
            End
            Begin VB.TextBox Text13 
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
               Index           =   5
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   136
               TabStop         =   0   'False
               Top             =   425
               Width           =   2415
            End
            Begin VB.TextBox Text12 
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
               Index           =   5
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   127
               TabStop         =   0   'False
               Top             =   2000
               Width           =   2415
            End
            Begin VB.TextBox Text11 
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
               Index           =   5
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   126
               TabStop         =   0   'False
               Top             =   2310
               Width           =   8295
            End
            Begin VB.TextBox Text9 
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
               Index           =   5
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   124
               TabStop         =   0   'False
               Top             =   2310
               Width           =   2415
            End
            Begin VB.TextBox Text8 
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
               Index           =   5
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   123
               TabStop         =   0   'False
               Top             =   2000
               Width           =   8295
            End
            Begin VB.TextBox Text7 
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
               Index           =   5
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   122
               TabStop         =   0   'False
               Top             =   1680
               Width           =   11415
            End
            Begin VB.TextBox Text6 
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
               Index           =   5
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   1365
               Width           =   11415
            End
            Begin VB.TextBox Text5 
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
               Index           =   5
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   120
               TabStop         =   0   'False
               Top             =   1050
               Width           =   11415
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
               Index           =   5
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   119
               TabStop         =   0   'False
               Top             =   735
               Width           =   11415
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
               Index           =   5
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   118
               TabStop         =   0   'False
               Top             =   425
               Width           =   8295
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
               Index           =   5
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   117
               TabStop         =   0   'False
               Top             =   105
               Width           =   11415
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   3
               Left            =   120
               TabIndex        =   128
               Top             =   425
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":4025
               Picture         =   "AccountMaster.frx":4041
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   129
               Top             =   105
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":405D
               Picture         =   "AccountMaster.frx":4079
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   2
               Left            =   120
               TabIndex        =   130
               Top             =   735
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   2240
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":4095
               Picture         =   "AccountMaster.frx":40B1
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   3
               Left            =   120
               TabIndex        =   131
               Top             =   2000
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":40CD
               Picture         =   "AccountMaster.frx":40E9
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   3
               Left            =   120
               TabIndex        =   132
               Top             =   2310
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":4105
               Picture         =   "AccountMaster.frx":4121
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   2
               Left            =   9480
               TabIndex        =   133
               Top             =   2310
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " GST No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":413D
               Picture         =   "AccountMaster.frx":4159
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   2
               Left            =   9480
               TabIndex        =   134
               Top             =   2000
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":4175
               Picture         =   "AccountMaster.frx":4191
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   5
               Left            =   9480
               TabIndex        =   135
               Top             =   425
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "AccountMaster.frx":41AD
               Picture         =   "AccountMaster.frx":41C9
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   12720
               Y1              =   2730
               Y2              =   2730
            End
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
            TabIndex        =   68
            Top             =   8445
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   13230
      _ExtentX        =   23336
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
Attribute VB_Name = "FrmAccountMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SL As Boolean, MasterCode As String, AccountType As String, AccountGroup As String, RateType As String 'SL-Selection List, MasterCode-Master to Modify
Dim cnAccountMaster As New ADODB.Connection
Dim rstAccountList As New ADODB.Recordset, rstAccountMaster As New ADODB.Recordset, rstAccountGroupList As New ADODB.Recordset, rstAccountChild As New ADODB.Recordset, rstElementList As New ADODB.Recordset, rstSizeList As New ADODB.Recordset
Dim AccountGroupCode As String, SizeGroupCode As Variant, ColorCode As Variant, SizeCode As String, OperationCode As Variant, CalcModeCode As Variant, BindingTypeCode As Variant, BinderyProcessCode As Variant, BOMItemCode As Variant, PaperCode As Variant, FinishedItemCode As Variant, UnfinishedItemCode As Variant, PlateCode As Variant, ElementCode As String, AccountGroupName As String
Dim PrevStr As String, blnRecordExist As Boolean, SortCol As String, SortOrder As String, EditMode As Boolean
Private Sub Form_Load()
    If Not SL Then MasterCode = ""
    On Error GoTo ErrorHandler
    CenterForm Me
    Me.Top = (MdiMainMenu.ScaleHeight - Me.Height) \ 2 + 1000
    BusySystemIndicator True
    Dim Cnt As Integer
    For Cnt = 1 To 8
        If Cnt <> Val(AccountType) Then SSTab1.TabVisible(Cnt) = False
    Next
    AccountGroup = IIf(CheckEmpty(AccountGroup, False), "Type IN ('12','26') AND P.Code<'*99995'", "[Group]='" & AccountGroup & "'")
    If AccountType <> "01" Then SSTab1.TabVisible(9) = False
    Me.Caption = Choose(Val(AccountType), "Account", , , "Processing Rate", "Printing", "Plate Rate", "Miscellaneous Operation Rate", "Binding Rate") & " Master" & IIf(RateType = "P", " [Purchase]", IIf(RateType = "S", " [Sale]", ""))
    cnAccountMaster.CursorLocation = adUseClient: cnAccountMaster.Open cnDatabase.ConnectionString
    rstAccountList.Open "SELECT P.Name,Alias,C.Name As AccountGroup,P.Code FROM AccountMaster P INNER JOIN GeneralMaster C ON P.[Group]=C.Code WHERE " & AccountGroup & " ORDER BY P.Name", cnAccountMaster, adOpenKeyset, adLockPessimistic
    LoadMasterList
    rstAccountMaster.CursorLocation = adUseClient
    rstAccountList.Filter = adFilterNone
    If rstAccountList.RecordCount > 0 Then
        rstAccountList.MoveFirst
        If Not CheckEmpty(MasterCode, False) Then rstAccountList.Find "[Code]='" & MasterCode & "'"
    End If
    Set DataGrid1.DataSource = rstAccountList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortCol = "Name"
    If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstAccountList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        EditMode = False
        If SSTab1.Tab = 0 Then
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then 'Add button enabled
                SSTab1.Tab = 0
            Else
                If InStr(1, "fpSpread1_fpSpread2_fpSpread3_fpSpread4_fpSpread5_fpSpread6", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True
                If Not EditMode Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then Me.ActiveControl.SetFocus Else Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                End If
            End If
        End If
        If Not EditMode Then KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then 'Add
        If AccountType = "01" Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then 'Edit
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2): KeyCode = 0
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF8)) And Toolbar1.Buttons.Item(3).Enabled Then 'Delete
        If AccountType = "01" Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF12 And Toolbar1.Buttons.Item(1).Enabled Then 'Duplicate
        If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then DuplicateRecord
        KeyCode = 0
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2)) And Toolbar1.Buttons.Item(4).Enabled Then 'Save
        EditMode = False
        If InStr(1, "fpSpread1_fpSpread2_fpSpread3_fpSpread4_fpSpread5_fpSpread6", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then 'Refresh
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF And Toolbar1.Buttons.Item(1).Enabled Then 'First
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(13): KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then 'Previous
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(14): KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyN And Toolbar1.Buttons.Item(1).Enabled Then 'Next
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(15): KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL And Toolbar1.Buttons.Item(1).Enabled Then 'Last
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(16): KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        If Toolbar1.Buttons.Item(1).Enabled Then
            If SL Then
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstAccountList.Fields("Code").Value: slName = rstAccountList.Fields("Name").Value: KeyCode = 0: Unload Me: Exit Sub
            Else
                SSTab1.Tab = Val(AccountType): SSTab1.SetFocus
            End If
        Else 'Move to next control
            If InStr(1, "fpSpread1_fpSpread2_fpSpread3_fpSpread4_fpSpread5_fpSpread6", Me.ActiveControl.Name) = 0 Then SendKeys "{TAB}"
        End If
        If InStr(1, "fpSpread1_fpSpread2_fpSpread3_fpSpread4_fpSpread5_fpSpread6", Me.ActiveControl.Name) = 0 Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1 Else If Me.Tag <> "S" Then slCode = "": slName = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstAccountMaster)
    Call CloseRecordset(rstAccountGroupList)
    Call CloseRecordset(rstAccountChild)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstElementList)
    Call CloseConnection(cnAccountMaster)
    ShowProgressInStatusBar False
End Sub
Private Sub Text1_Change()
    If rstAccountList.RecordCount = 0 Then Exit Sub
    rstAccountList.MoveFirst
    If Len(Text1.Text) > 0 Then
        rstAccountList.Filter = "[" & SortCol & "] Like '%" & FixQuote(Text1.Text) & "%'"
        If rstAccountList.EOF Then  'if Spelling mistake
            rstAccountList.Filter = adFilterNone
            rstAccountList.MoveFirst
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            SendKeys "{End}"
        Else    'if Spelling alright
            PrevStr = Text1.Text
        End If
    Else
        rstAccountList.Filter = adFilterNone
        rstAccountList.MoveFirst
        Set DataGrid1.DataSource = rstAccountList
        PrevStr = ""
    End If
    If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    If rstAccountList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstAccountList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstAccountList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstAccountList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstAccountList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstAccountList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstAccountList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstAccountList
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
    With SSTab1
        If Toolbar1.Buttons.Item(1).Enabled Then 'Add Button Enabled
            If IIf(.Tab = 9, 1, .Tab) = Val(AccountType) Then
                ViewRecord
            Else
                If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
                    With DataGrid1.SelBookmarks
                        If .Count <> 0 Then .Remove 0
                        .Add DataGrid1.Bookmark
                    End With
                End If
                Text1.SetFocus
            End If
            .TabEnabled(0) = True
        Else
            .TabEnabled(0) = False
            Mh3dFrame2(0).Enabled = IIf(.Tab = 1, True, False): Mh3dFrame2(8).Enabled = IIf(.Tab = 1, False, True): IIf(.Tab = 1, Text2(0), fpSpread1).SetFocus
        End If
    End With
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim CellVal(1 To 4) As Variant, HiLiteRecord As Boolean, UpdateFlag As Integer, i As Integer
    If Button.Index = 1 Then 'Add Button
        If rstAccountMaster.State = adStateOpen Then rstAccountMaster.Close
        rstAccountMaster.Open "SELECT * FROM AccountMaster WHERE Code=''", cnAccountMaster, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstAccountMaster) Then
            Call SetButtons(False): SSTab1.Tab = 1: Text2(0).SetFocus: blnRecordExist = False
            cnAccountMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then 'Edit Button
        If rstAccountList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = Val(AccountType)
        EditRecord
    ElseIf Button.Index = 3 Then 'Delete Button
        If rstAccountList.RecordCount = 0 Or Left(rstAccountList.Fields("Code").Value, 1) = "*" Then Exit Sub
        If AllowMastersDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Master"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            cnAccountMaster.BeginTrans
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnAccountMaster.Execute "DELETE FROM AccountMaster WHERE Code='" & rstAccountList.Fields("Code").Value & "'"
            If Err.Number = 0 Then
                cnAccountMaster.CommitTrans
                rstAccountList.Delete
                rstAccountList.MoveNext
                If rstAccountList.RecordCount > 0 And rstAccountList.EOF Then rstAccountList.MoveLast
                ShowProgressInStatusBar True
                Timer1.Enabled = True
                Text1.Text = ""
                rstAccountList.Filter = adFilterNone
            Else
                DisplayError (Err.Description)
                cnAccountMaster.RollbackTrans
            End If
            MdiMainMenu.MousePointer = vbNormal
            On Error GoTo 0
        End If
        SetButtons (True)
        SetButtonsForNoRecord
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 4 Then 'Save Button
        If ValidateForm Then Exit Sub
        If blnRecordExist And AllowMastersModification = 0 Then Call DisplayError("You don't have the rights to Edit this Master"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(5): Exit Sub
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstAccountMaster) Then
            UpdateFlag = 1
            If UpdateRateList("D") Then
                If AccountType = "01" Then
                    With fpSpread1
                        For i = 1 To .DataRowCnt
                            .SetActiveCell 1, i
                            .GetText 1, i, CellVal(1) 'Category
                            .GetText 4, i, CellVal(2) 'Item
                            .GetText 5, i, CellVal(3) 'Imported
                            If Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False)) And (CellVal(3) = "N" Or CellVal(3) = "") Then If Not UpdateRateList("I") Then UpdateFlag = 0: Exit For
                        Next
                    End With
                ElseIf AccountType = "04" Then
                    With fpSpread6
                        For i = 1 To .DataRowCnt
                            .SetActiveCell 1, i
                            If Not UpdateRateList("I") Then UpdateFlag = 0: Exit For
                        Next
                    End With
                ElseIf AccountType = "05" Then
                    With fpSpread2
                        For i = 1 To .DataRowCnt
                            .SetActiveCell 1, i
                            .GetText 9, i, CellVal(1) 'Size group
                            .GetText 10, i, CellVal(2) 'Color
                            If Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False)) Then If Not UpdateRateList("I") Then UpdateFlag = 0: Exit For
                        Next
                    End With
                ElseIf AccountType = "06" Then
                    With fpSpread5
                        For i = 1 To .DataRowCnt
                            .SetActiveCell 1, i
                            .GetText 5, i, CellVal(1) 'Size group
                            If Not CheckEmpty(CellVal(1), False) Then If Not UpdateRateList("I") Then UpdateFlag = 0: Exit For
                        Next
                    End With
                ElseIf AccountType = "07" Then
                    With fpSpread4
                        For i = 1 To .DataRowCnt
                            .SetActiveCell 1, i
                            .GetText 7, i, CellVal(1) 'Operation
                            .GetText 8, i, CellVal(2) 'Size
                            .GetText 9, i, CellVal(3) 'Calc Mode
                            If IIf(CheckEmpty(CellVal(2), False), Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(3), False)), Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False) Or CheckEmpty(CellVal(3), False))) Then If Not UpdateRateList("I") Then UpdateFlag = 0: Exit For
                        Next
                    End With
                ElseIf AccountType = "08" Then
                    With fpSpread3
                        For i = 1 To .DataRowCnt
                            .SetActiveCell 1, i
                            .GetText 10, i, CellVal(1) 'Binding Type
                            .GetText 11, i, CellVal(2) 'Bindery Process
                            .GetText 12, i, CellVal(3) 'Calc Mode
                            .GetText 13, i, CellVal(4) 'Size Group
                            If IIf(CheckEmpty(CellVal(4), False), Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False) Or CheckEmpty(CellVal(3), False)), Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False) Or CheckEmpty(CellVal(3), False) Or CheckEmpty(CellVal(4), False))) Then If Not UpdateRateList("I") Then UpdateFlag = 0: Exit For
                        Next
                    End With
                End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnAccountMaster.CommitTrans
            If rstAccountMaster.State = adStateOpen Then rstAccountMaster.Close
            rstAccountMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            Call MsgBox("Record updated !!!", vbInformation, App.Title)
        Else
            DisplayError ("Failed to Save the Record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then 'Cancel Button
        If CancelRecordUpdate(rstAccountMaster) Then
            cnAccountMaster.RollbackTrans
            If rstAccountMaster.State = adStateOpen Then rstAccountMaster.Close
            rstAccountMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then 'Refresh Button
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        RefreshData rstAccountList
        Set DataGrid1.DataSource = rstAccountList
        LoadMasterList
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then 'Filter Button
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Name", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then 'First Record Button
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then 'Previous Record Button
        If rstAccountList.RecordCount > 0 Then
           rstAccountList.MovePrevious
           If rstAccountList.BOF Then rstAccountList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then 'Next Record Button
        If rstAccountList.RecordCount > 0 Then
           rstAccountList.MoveNext
           If rstAccountList.EOF Then
              rstAccountList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then 'Last Record Button
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
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
    SortCol = DataGrid1.Columns(ColIndex).DataField
    SortOrder = IIf(SortOrder = "Asc", "Desc", "Asc")
    rstAccountList.Sort = "[" + SortCol & "] " & SortOrder
    DataGrid1.ClearSelCols
    If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub Text2_Validate(Index As Integer, Cancel As Boolean) 'Account Name
    If rstAccountMaster.EOF Or rstAccountMaster.BOF Or AccountType <> "01" Then Exit Sub
    If CheckEmpty(Text2(0), True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnAccountMaster, "AccountMaster", "Code", "Name", Trim(Text2(0).Text), rstAccountMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3(0), False) Then
        Text3(0).Text = Text2(0).Text
    End If
End Sub
Private Sub Text10_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer) 'Account Group
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "12,26"
        FrmGeneralMaster.MasterCode = AccountGroupCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        AccountGroupCode = slCode: Text10(0).Text = slName
        If Not CheckEmpty(AccountGroupCode, False) Then LoadMasterList: SendKeys "{TAB}"
    End If
End Sub
Private Sub Text10_Validate(Index As Integer, Cancel As Boolean) 'Account Group
    If CheckEmpty(Text10(0).Text, False) Then Cancel = True
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstAccountList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    With rstAccountMaster
        If .State = adStateOpen Then .Close
        .Open "SELECT * FROM AccountMaster WHERE Code='" & FixQuote(rstAccountList.Fields("Code").Value) & "'", cnAccountMaster, adOpenKeyset, adLockReadOnly
        If .RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End With
End Sub
Private Sub ClearFields()
    Text2(Val(AccountType) - 1).Text = ""
    Text3(Val(AccountType) - 1).Text = ""
    Text13(Val(AccountType) - 1).Text = ""
    Text4(Val(AccountType) - 1).Text = ""
    Text5(Val(AccountType) - 1).Text = ""
    Text6(Val(AccountType) - 1).Text = ""
    Text7(Val(AccountType) - 1).Text = ""
    Text8(Val(AccountType) - 1).Text = ""
    Text12(Val(AccountType) - 1).Text = ""
    Text9(Val(AccountType) - 1).Text = ""
    Text11(Val(AccountType) - 1).Text = ""
    If AccountType = "01" Then
        Text10(Val(AccountType) - 1).Text = "": fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    ElseIf AccountType = "04" Then
        fpSpread6.ClearRange 1, 1, fpSpread6.MaxCols, fpSpread6.MaxRows, True: fpSpread6.SetActiveCell 1, 1
    ElseIf AccountType = "05" Then
        chkRound.Value = 0: fpSpread2.ClearRange 1, 1, fpSpread2.MaxCols, fpSpread2.MaxRows, True: fpSpread2.SetActiveCell 1, 1
    ElseIf AccountType = "06" Then
        fpSpread5.ClearRange 1, 1, fpSpread5.MaxCols, fpSpread5.MaxRows, True: fpSpread5.SetActiveCell 1, 1
    ElseIf AccountType = "07" Then
        fpSpread4.ClearRange 1, 1, fpSpread4.MaxCols, fpSpread4.MaxRows, True: fpSpread4.SetActiveCell 1, 1
    ElseIf AccountType = "08" Then
        fpSpread3.ClearRange 1, 1, fpSpread3.MaxCols, fpSpread3.MaxRows, True:  fpSpread3.SetActiveCell 1, 1
    End If
End Sub
Private Sub LoadFields()
    With rstAccountMaster
        If .EOF Or .BOF Then Exit Sub
        Text2(Val(AccountType) - 1).Text = .Fields("Name").Value
        Text3(Val(AccountType) - 1).Text = .Fields("PrintName").Value
        Text13(Val(AccountType) - 1).Text = .Fields("Alias").Value
        Text4(Val(AccountType) - 1).Text = .Fields("Address1").Value
        Text5(Val(AccountType) - 1).Text = .Fields("Address2").Value
        Text6(Val(AccountType) - 1).Text = .Fields("Address3").Value
        Text7(Val(AccountType) - 1).Text = .Fields("Address4").Value
        Text8(Val(AccountType) - 1).Text = .Fields("Phone").Value
        Text12(Val(AccountType) - 1).Text = .Fields("Mobile").Value
        Text9(Val(AccountType) - 1).Text = .Fields("TIN").Value
        Text11(Val(AccountType) - 1).Text = .Fields("EMail").Value
        AccountGroupCode = .Fields("Group").Value
        If rstAccountGroupList.RecordCount > 0 Then rstAccountGroupList.MoveFirst
        rstAccountGroupList.Find "[Code] = '" & AccountGroupCode & "'"
        If AccountType = "01" Then
            If Not rstAccountGroupList.EOF Then Text10(Val(AccountType) - 1).Text = rstAccountGroupList.Fields("Col0").Value
        ElseIf AccountType = "05" Then
            chkRound.Value = IIf(.Fields("RoundOffQty").Value, 1, 0)
        End If
        If Not rstAccountGroupList.EOF Then AccountGroupName = rstAccountGroupList.Fields("Col0").Value
        Call LoadRateList(.Fields("Code").Value)
    End With
End Sub
Private Sub LoadRateList(ByVal MasterCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    With rstAccountChild
        If .State = adStateOpen Then .Close
        If AccountType = "01" Then 'Account Master
            .Open "SELECT * FROM (SELECT Category,IIF(Category='4',Item+'-'+SubItem,Item) As Item,SubItem,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=C.Item),IIF(Category='2',(SELECT Name FROM PaperMaster WHERE Code=C.Item),IIF(Category='3',(SELECT Name FROM BookMaster WHERE Code=C.Item),(SELECT Name FROM ElementMaster WHERE Code=C.Item)+'-'+(SELECT Name FROM BookMaster WHERE Code=C.SubItem)))) As ItemName,OpBal,Imported FROM AccountChild0801 C WHERE C.Code='" & MasterCode & "') As Tbl ORDER BY Category,ItemName", cnAccountMaster, adOpenKeyset, adLockReadOnly
            i = 0
            Do Until .EOF
                i = i + 1
                fpSpread1.SetText 1, i, Choose(Val(.Fields("Category").Value), "BOM Item", "Paper", "Unfinished Item", "Multi Element Format"): fpSpread1.SetText 4, i, .Fields("Item").Value
                fpSpread1.SetText 2, i, .Fields("ItemName").Value
                fpSpread1.SetText 3, i, Val(.Fields("OpBal").Value)
                fpSpread1.SetText 5, i, .Fields("Imported").Value
                .MoveNext
            Loop
            fpSpread1.SetActiveCell 1, 1
        ElseIf AccountType = "04" Then 'Processing Rate List
            .Open "SELECT NegativeOnePcRate,NegativeCutPcRate,NegativePastingRate,PositiveOnePcRate,PositiveCutPcRate,PositivePastingRate,wef FROM AccountChild04 WHERE Code='" & MasterCode & "' AND Type='" & RateType & "' ORDER BY wef DESC", cnAccountMaster, adOpenKeyset, adLockReadOnly
            i = 0
            Do Until .EOF
                i = i + 1
                fpSpread6.SetText 1, i, Val(.Fields("NegativeCutPcRate").Value)
                fpSpread6.SetText 2, i, Val(.Fields("PositiveCutPcRate").Value)
                fpSpread6.SetText 3, i, Val(.Fields("NegativeOnePcRate").Value)
                fpSpread6.SetText 4, i, Val(.Fields("PositiveOnePcRate").Value)
                fpSpread6.SetText 5, i, Val(.Fields("NegativePastingRate").Value)
                fpSpread6.SetText 6, i, Val(.Fields("PositivePastingRate").Value)
                fpSpread6.SetText 7, i, Format(.Fields("wef").Value, "dd-MM-yyyy")
                .MoveNext
            Loop
            fpSpread6.SetActiveCell 1, 1
        ElseIf AccountType = "05" Then 'Printing Rate List
            .Open "SELECT SizeGroup,S.Name As SizeGroupName,[Color],R.Name As ColorName,Range,PrintingRate,PaperWastageRate,PaperWastageMin,PaperWastageMax,wef FROM (AccountChild05 C INNER  JOIN GeneralMaster S ON C.SizeGroup=S.Code) INNER JOIN GeneralMaster R ON C.Color=R.Code WHERE C.Code='" & MasterCode & "' AND C.Type='" & RateType & "' ORDER BY wef DESC,S.Name,R.Name,Range", cnAccountMaster, adOpenKeyset, adLockReadOnly
            i = 0
            Do Until .EOF
                i = i + 1
                fpSpread2.SetText 1, i, .Fields("SizeGroupName").Value: fpSpread2.SetText 9, i, .Fields("SizeGroup").Value
                fpSpread2.SetText 2, i, .Fields("ColorName").Value: fpSpread2.SetText 10, i, .Fields("Color").Value
                fpSpread2.SetText 3, i, Val(.Fields("Range").Value)
                fpSpread2.SetText 4, i, Val(.Fields("PrintingRate").Value)
                fpSpread2.SetText 5, i, Val(.Fields("PaperWastageRate").Value)
                fpSpread2.SetText 6, i, Val(.Fields("PaperWastageMin").Value)
                fpSpread2.SetText 7, i, Val(.Fields("PaperWastageMax").Value)
                fpSpread2.SetText 8, i, Format(.Fields("wef").Value, "dd-MM-yyyy")
                .MoveNext
            Loop
            fpSpread2.SetActiveCell 1, 1
        ElseIf AccountType = "06" Then 'Plate Rate List
            .Open "SELECT SizeGroup,S.Name As SizeGroupName,Rate,Plate,P.Name As PlateName,wef FROM (AccountChild06 C INNER  JOIN GeneralMaster S ON C.SizeGroup=S.Code) INNER JOIN GeneralMaster P ON C.Plate=P.Code WHERE C.Code='" & MasterCode & "' AND C.Type='" & RateType & "' ORDER BY wef DESC,S.Name,P.Name", cnAccountMaster, adOpenKeyset, adLockReadOnly
            i = 0
            Do Until .EOF
                i = i + 1
                fpSpread5.SetText 1, i, .Fields("SizeGroupName").Value: fpSpread5.SetText 5, i, .Fields("SizeGroup").Value
                fpSpread5.SetText 2, i, .Fields("PlateName").Value: fpSpread5.SetText 6, i, .Fields("Plate").Value
                fpSpread5.SetText 3, i, Val(.Fields("Rate").Value)
                fpSpread5.SetText 4, i, Format(.Fields("wef").Value, "dd-MM-yyyy")
                .MoveNext
            Loop
            fpSpread5.SetActiveCell 1, 1
        ElseIf AccountType = "07" Then 'Operation Rate List
            .Open "SELECT LaminationType As Operation,O.Name As OperationName,CalcMode,M.Name As CalcModeName,[Size],S.Name As SizeName,Range,Rate,wef FROM ((AccountChild07 C INNER JOIN GeneralMaster O ON C.LaminationType=O.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code) LEFT JOIN GeneralMaster S ON C.[Size]=S.Code WHERE C.Code='" & MasterCode & "' AND C.Type='" & RateType & "' ORDER BY wef DESC,O.Name,M.Name,S.Name,Range", cnAccountMaster, adOpenKeyset, adLockReadOnly
            i = 0
            Do Until .EOF
                i = i + 1
                fpSpread4.SetText 1, i, .Fields("OperationName").Value: fpSpread4.SetText 7, i, .Fields("Operation").Value
                fpSpread4.SetText 2, i, .Fields("SizeName").Value: fpSpread4.SetText 8, i, .Fields("Size").Value
                fpSpread4.SetText 3, i, .Fields("CalcModeName").Value: fpSpread4.SetText 9, i, .Fields("CalcMode").Value
                fpSpread4.SetText 4, i, Val(.Fields("Range").Value)
                fpSpread4.SetText 5, i, Val(.Fields("Rate").Value)
                fpSpread4.SetText 6, i, Format(.Fields("wef").Value, "dd-MM-yyyy")
                .MoveNext
            Loop
            fpSpread4.SetActiveCell 1, 1
        ElseIf AccountType = "08" Then 'Binding Rate List
            .Open "SELECT BindingType,B.Name As BindingTypeName,BinderyProcess,P.Name As BinderyProcessName,CalcMode,M.Name As CalcModeName,SizeGroup,S.Name As SizeGroupName,Fraction,Range,Rate,AddOnRate,wef FROM (((AccountChild08 C INNER JOIN GeneralMaster B ON C.BindingType=B.Code) INNER JOIN GeneralMaster P ON C.BinderyProcess=P.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code) LEFT JOIN GeneralMaster S ON C.SizeGroup=S.Code WHERE C.Code='" & MasterCode & "' AND C.Type='" & RateType & "' ORDER BY B.Name,P.Name,M.Name,S.Name,Range", cnAccountMaster, adOpenKeyset, adLockReadOnly
            i = 0
            Do Until .EOF
                i = i + 1
                fpSpread3.SetText 1, i, .Fields("BindingTypeName").Value: fpSpread3.SetText 10, i, .Fields("BindingType").Value
                fpSpread3.SetText 2, i, .Fields("BinderyProcessName").Value: fpSpread3.SetText 11, i, .Fields("BinderyProcess").Value
                fpSpread3.SetText 3, i, .Fields("CalcModeName").Value: fpSpread3.SetText 12, i, .Fields("CalcMode").Value
                fpSpread3.SetText 4, i, .Fields("SizeGroupName").Value: fpSpread3.SetText 13, i, .Fields("SizeGroup").Value
                fpSpread3.SetText 5, i, Val(.Fields("Fraction").Value)
                fpSpread3.SetText 6, i, Val(.Fields("Range").Value)
                fpSpread3.SetText 7, i, Val(.Fields("Rate").Value)
                fpSpread3.SetText 8, i, Val(.Fields("AddOnRate").Value)
                fpSpread3.SetText 9, i, Format(.Fields("wef").Value, "dd-MM-yyyy")
                .MoveNext
            Loop
            fpSpread3.SetActiveCell 1, 1
        End If
    End With
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    With rstAccountMaster
        If .RecordCount = 0 Then Exit Sub
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseServer
        .Open "SELECT * FROM AccountMaster WHERE Code='" & FixQuote(rstAccountList.Fields("Code").Value) & "'", cnAccountMaster, adOpenKeyset, adLockPessimistic
        MdiMainMenu.MousePointer = vbHourglass
        .Fields("Printstatus") = "N"
    End With
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Choose(Val(AccountType), Text2(0), , , fpSpread6, fpSpread2, fpSpread5, fpSpread4, fpSpread3).SetFocus
    blnRecordExist = True
    cnAccountMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    With rstAccountMaster
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnAccountMaster, "SELECT MAX(Code) FROM AccountMaster", 6, "0")
            .Fields("CreatedBy").Value = UserCode: .Fields("CreatedOn").Value = Now()
            .Fields("RecordStatus").Value = "N"
        Else
            .Fields("ModifiedBy").Value = UserCode: .Fields("ModifiedOn").Value = Now()
            .Fields("RecordStatus").Value = "M"
        End If
        .Fields("Name").Value = Trim(Text2(Val(AccountType) - 1).Text)
        .Fields("PrintName").Value = Trim(Text3(Val(AccountType) - 1).Text)
        .Fields("Alias").Value = Trim(Text13(Val(AccountType) - 1).Text)
        .Fields("Group").Value = AccountGroupCode
        .Fields("Address1").Value = Trim(Text4(Val(AccountType) - 1).Text)
        .Fields("Address2").Value = Trim(Text5(Val(AccountType) - 1).Text)
        .Fields("Address3").Value = Trim(Text6(Val(AccountType) - 1).Text)
        .Fields("Address4").Value = Trim(Text7(Val(AccountType) - 1).Text)
        .Fields("Phone").Value = Trim(Text8(Val(AccountType) - 1).Text)
        .Fields("Mobile").Value = Trim(Text12(Val(AccountType) - 1).Text)
        .Fields("TIN").Value = Trim(Text9(Val(AccountType) - 1).Text)
        .Fields("EMail").Value = Trim(Text11(Val(AccountType) - 1).Text)
        .Fields("PrintStatus").Value = "N"
        If AccountType = "05" Then .Fields("RoundOffQty").Value = chkRound.Value
    End With
End Sub
Private Function UpdateRateList(ByVal ActionType As String) As Boolean
    On Error GoTo ErrorHandler
    Dim CellVal(1 To 9) As Variant
    UpdateRateList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType <> "I" Then
        cnAccountMaster.Execute "DELETE FROM AccountChild" & IIf(AccountType = "01", "08", "") & AccountType & " WHERE Code='" & rstAccountMaster.Fields("Code").Value & "'" & IIf(AccountType = "01", " AND Imported='N'", "")
    Else
        If AccountType = "01" Then
            With fpSpread1
                .GetText 1, .ActiveRow, CellVal(1) 'Category
                .GetText 4, .ActiveRow, CellVal(2) 'Item
                .GetText 3, .ActiveRow, CellVal(3) 'Op Bal
                CellVal(1) = IIf(CellVal(1) = "BOM Item", "1", IIf(CellVal(1) = "Paper", "2", IIf(CellVal(1) = "Unfinished Item", "3", "4")))
            End With
            cnAccountMaster.Execute "INSERT INTO AccountChild0801 VALUES ('" & rstAccountMaster.Fields("Code").Value & "','" & CellVal(1) & "','" & Left(CellVal(2), 6) & "'," & Val(CellVal(3)) & ",'N'," & IIf(CellVal(1) = "4", "'" & Right(CellVal(2), 6) & "'", "Null") & ")"
        ElseIf AccountType = "04" Then
            With fpSpread6
                .GetText 1, .ActiveRow, CellVal(1) 'Cut piece rate - Negative
                .GetText 2, .ActiveRow, CellVal(2) 'Cut piece rate - Positive
                .GetText 3, .ActiveRow, CellVal(3) 'One piece rate - Negative
                .GetText 4, .ActiveRow, CellVal(4) 'One piece rate - Positive
                .GetText 5, .ActiveRow, CellVal(5) 'Pasting rate - Negative
                .GetText 6, .ActiveRow, CellVal(6) 'Pasting rate - Positive
                .GetText 7, .ActiveRow, CellVal(7) 'wef
            End With
            cnAccountMaster.Execute "INSERT INTO AccountChild04 VALUES ('" & rstAccountMaster.Fields("Code").Value & "'," & Val(CellVal(3)) & "," & Val(CellVal(1)) & "," & Val(CellVal(5)) & "," & Val(CellVal(4)) & "," & Val(CellVal(2)) & "," & Val(CellVal(6)) & ",'" & Format(CellVal(7), "dd-MMM-yyyy") & "','" & RateType & "')"
        ElseIf AccountType = "05" Then
            With fpSpread2
                .GetText 9, .ActiveRow, CellVal(1) 'Size group
                .GetText 10, .ActiveRow, CellVal(2) 'Color
                .GetText 3, .ActiveRow, CellVal(3) 'Range
                .GetText 4, .ActiveRow, CellVal(4) 'Printing rate
                .GetText 5, .ActiveRow, CellVal(5) 'Paper wastage rate
                .GetText 6, .ActiveRow, CellVal(6) 'Paper wastage min
                .GetText 7, .ActiveRow, CellVal(7) 'Paper wastage max
                .GetText 8, .ActiveRow, CellVal(8) 'wef
            End With
            cnAccountMaster.Execute "INSERT INTO AccountChild05 VALUES ('" & rstAccountMaster.Fields("Code").Value & "','" & CellVal(1) & "'," & Val(CellVal(3)) & "," & Val(CellVal(4)) & "," & Val(CellVal(5)) & "," & Val(CellVal(6)) & "," & Val(CellVal(7)) & ",'" & CellVal(2) & "','" & Format(CellVal(8), "dd-MMM-yyyy") & "','" & RateType & "')"
        ElseIf AccountType = "06" Then
            With fpSpread5
                .GetText 5, .ActiveRow, CellVal(1) 'Size group
                .GetText 6, .ActiveRow, CellVal(2) 'Plate
                .GetText 3, .ActiveRow, CellVal(3) 'rate
                .GetText 4, .ActiveRow, CellVal(4) 'wef
            End With
            cnAccountMaster.Execute "INSERT INTO AccountChild06 VALUES ('" & rstAccountMaster.Fields("Code").Value & "','" & CellVal(1) & "'," & Val(CellVal(3)) & ",'" & CellVal(2) & "','" & Format(CellVal(4), "dd-MMM-yyyy") & "','" & RateType & "')"
        ElseIf AccountType = "07" Then
            With fpSpread4
                .GetText 7, .ActiveRow, CellVal(1) 'Operation
                .GetText 8, .ActiveRow, CellVal(2) 'Size
                .GetText 9, .ActiveRow, CellVal(3) 'Calc Mode
                .GetText 4, .ActiveRow, CellVal(4) 'Range
                .GetText 5, .ActiveRow, CellVal(5) 'Rate
                .GetText 6, .ActiveRow, CellVal(6) 'wef
            End With
            cnAccountMaster.Execute "INSERT INTO AccountChild07 VALUES ('" & rstAccountMaster.Fields("Code").Value & "'," & IIf(CheckEmpty(CellVal(2), False), "Null", "'" & CellVal(2) & "'") & ",'" & CellVal(1) & "','" & CellVal(3) & "'," & Val(CellVal(5)) & "," & Val(CellVal(4)) & ",'" & Format(CellVal(6), "dd-MMM-yyyy") & "','" & RateType & "')"
        ElseIf AccountType = "08" Then
            With fpSpread3
                .GetText 10, .ActiveRow, CellVal(1) 'Binding Type
                .GetText 11, .ActiveRow, CellVal(2) 'Bindery Process
                .GetText 12, .ActiveRow, CellVal(3) 'Calc Mode
                .GetText 13, .ActiveRow, CellVal(4) 'Size Group
                .GetText 5, .ActiveRow, CellVal(5) 'Fraction
                .GetText 6, .ActiveRow, CellVal(6) 'Range
                .GetText 7, .ActiveRow, CellVal(7) 'Rate
                .GetText 8, .ActiveRow, CellVal(8) 'Add-on Rate
                .GetText 9, .ActiveRow, CellVal(9) 'wef
            End With
            cnAccountMaster.Execute "INSERT INTO AccountChild08 VALUES ('" & rstAccountMaster.Fields("Code").Value & "','" & CellVal(1) & "','" & CellVal(2) & "','" & CellVal(3) & "','" & CellVal(4) & "'," & Val(CellVal(5)) & "," & Val(CellVal(6)) & "," & Val(CellVal(7)) & "," & Val(CellVal(8)) & ",'" & Format(CellVal(9), "dd-MMM-yyyy") & "','" & RateType & "')"
        End If
    End If
    Exit Function
ErrorHandler:
    UpdateRateList = False
End Function
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CurVal As Variant
    With fpSpread1
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            Dim Imported As Variant
            .GetText 5, .ActiveRow, Imported
            If Imported = "Y" Then Exit Sub
            If MsgBox("Are you sure to delete the record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 2 Then
                .GetText 1, .ActiveRow, CurVal
                If CheckEmpty(CurVal, False) Then Exit Sub
                If CurVal = "BOM Item" Then
                    .GetText 4, .ActiveRow, BOMItemCode
                    On Error Resume Next
                    FrmOutsourceItemMaster.SL = True
                    FrmOutsourceItemMaster.MasterCode = BOMItemCode
                    Load FrmOutsourceItemMaster
                    If Err.Number <> 364 Then FrmOutsourceItemMaster.Show vbModal
                    On Error GoTo 0
                    .SetText 2, .ActiveRow, slName: .SetText 4, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
                ElseIf CurVal = "Paper" Then
                    .GetText 4, .ActiveRow, PaperCode
                    On Error Resume Next
                    FrmPaperMaster.SL = True
                    FrmPaperMaster.MasterCode = PaperCode
                    Load FrmPaperMaster
                    If Err.Number <> 364 Then FrmPaperMaster.Show vbModal
                    On Error GoTo 0
                    .SetText 2, .ActiveRow, slName: .SetText 4, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
                ElseIf CurVal = "Unfinished Item" Then
                    .GetText 4, .ActiveRow, UnfinishedItemCode
                    On Error Resume Next
                    FrmBookMaster.SL = True
                    FrmBookMaster.ItemType = "R"
                    FrmBookMaster.MasterCode = UnfinishedItemCode
                    Load FrmBookMaster
                    If Err.Number <> 364 Then FrmBookMaster.Show vbModal
                    On Error GoTo 0
                    .SetText 2, .ActiveRow, slName: .SetText 4, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
                ElseIf CurVal = "Multi Element Format" Then
                    If rstElementList.RecordCount = 0 Then DisplayError ("No Record in Element Master"): .SetActiveCell 2, .ActiveRow: Exit Sub Else rstElementList.MoveFirst
                    .GetText 4, .ActiveRow, FinishedItemCode
                    Text99.Text = ""
                    If Not CheckEmpty(FinishedItemCode, False) Then
                        ElementCode = Left(FinishedItemCode, 6): FinishedItemCode = Right(FinishedItemCode, 6)
                        rstElementList.Find "[Code] = '" & RTrim(ElementCode) & "'"
                        Text99.Text = rstElementList.Fields("Col0").Value
                    End If
                    SelectionType = "S": ElementCode = ""
                    Call LoadSelectionList(rstElementList, "List of Elements...", "Name")
                    SearchOrder = 0
                    Call DisplaySelectionList(Text99, ElementCode)
                    Call CloseForm(FrmSelectionList)
                    If Not CheckEmpty(ElementCode, False) Then
                        On Error Resume Next
                        FrmBookMaster.SL = True
                        FrmBookMaster.ItemType = "F"
                        FrmBookMaster.MasterCode = FinishedItemCode
                        Load FrmBookMaster
                        If Err.Number <> 364 Then FrmBookMaster.Show vbModal
                        On Error GoTo 0
                        If Not CheckEmpty(slCode, False) Then .SetText 2, .ActiveRow, Text99.Text + "-" + slName: .SetText 4, .ActiveRow, ElementCode + "-" + slCode: SendKeys "{ENTER}"
                    End If
                    If CheckEmpty(ElementCode, False) Or CheckEmpty(slCode, False) Then .SetActiveCell 2, .ActiveRow: .SetText 2, .ActiveRow, "": .SetText 4, .ActiveRow, ""
                End If
            End If
        End If
    End With
End Sub
Private Sub fpSpread1_BeforeEditMode(ByVal Col As Long, ByVal Row As Long, ByVal UserAction As FPSpreadADO.BeforeEditModeActionConstants, CursorPos As Variant, Cancel As Variant)
    Dim Imported As Variant
    fpSpread1.GetText 5, fpSpread1.ActiveRow, Imported
    If Imported = "Y" Then Cancel = True
End Sub
Private Sub fpSpread6_KeyDown(KeyCode As Integer, Shift As Integer)
    With fpSpread6
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        End If
    End With
End Sub
Private Sub fpSpread2_KeyDown(KeyCode As Integer, Shift As Integer)
    With fpSpread2
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then
                .GetText 9, .ActiveRow, SizeGroupCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "10"
                FrmGeneralMaster.MasterCode = SizeGroupCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 1, .ActiveRow, slName: .SetText 9, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
            ElseIf .ActiveCol = 2 Then
                .GetText 10, .ActiveRow, ColorCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "23"
                FrmGeneralMaster.MasterCode = ColorCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 2, .ActiveRow, slName: .SetText 10, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
            End If
        End If
    End With
End Sub
Private Sub fpSpread4_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CurVal As Variant
    With fpSpread4
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf Shift = 0 And KeyCode = vbKeyDelete Then
            If .ActiveCol = 2 Then .SetText 2, .ActiveRow, "":  .SetText 8, .ActiveRow, ""
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then
                .GetText 7, .ActiveRow, OperationCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "7"
                FrmGeneralMaster.MasterCode = OperationCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 1, .ActiveRow, slName: .SetText 7, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
            ElseIf .ActiveCol = 2 Then
                .GetText 2, .ActiveRow, CurVal
                Text99.Text = FixQuote(CurVal)
                If rstSizeList.RecordCount = 0 Then DisplayError ("No Record in Size Master"): .SetActiveCell 2, .ActiveRow: Exit Sub Else rstSizeList.MoveFirst
                rstSizeList.Find "[Col0] = '" & RTrim(CurVal) & "'"
                SelectionType = "S": SizeCode = ""
                Call LoadSelectionList(rstSizeList, "List of Size(s)...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text99, SizeCode)
                Call CloseForm(FrmSelectionList)
                .SetText 8, .ActiveRow, SizeCode
                If CheckEmpty(SizeCode, False) Then
                    .SetActiveCell 2, .ActiveRow: .SetText 2, .ActiveRow, ""
                Else
                    .SetText 2, .ActiveRow, Text99.Text: SendKeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 3 Then
                .GetText 9, .ActiveRow, CalcModeCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "20"
                FrmGeneralMaster.MasterCode = CalcModeCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 3, .ActiveRow, slName: .SetText 9, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
            End If
        End If
    End With
End Sub
Private Sub fpSpread5_KeyDown(KeyCode As Integer, Shift As Integer)
    With fpSpread5
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then
                .GetText 5, .ActiveRow, SizeGroupCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "10"
                FrmGeneralMaster.MasterCode = SizeGroupCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 1, .ActiveRow, slName: .SetText 5, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
            ElseIf .ActiveCol = 2 Then
                .GetText 6, .ActiveRow, PlateCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "24"
                FrmGeneralMaster.MasterCode = PlateCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 2, .ActiveRow, slName: .SetText 6, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
            End If
        End If
    End With
End Sub
Private Sub fpSpread3_KeyDown(KeyCode As Integer, Shift As Integer)
    With fpSpread3
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf Shift = 0 And KeyCode = vbKeyDelete Then
            If .ActiveCol = 4 Then .SetText 4, .ActiveRow, "":  .SetText 13, .ActiveRow, ""
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then
                .GetText 10, .ActiveRow, BindingTypeCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "6"
                FrmGeneralMaster.MasterCode = BindingTypeCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 1, .ActiveRow, slName: .SetText 10, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then
                    .GetText 11, .ActiveRow, BinderyProcessCode
                    If CheckEmpty(BinderyProcessCode, False) Then
                        Dim i As Integer
                        i = .DataRowCnt
                        With rstAccountChild
                            If .State = adStateOpen Then .Close
                            .Open "SELECT B.Code,B.Name FROM BindingTypeChild C INNER JOIN GeneralMaster B ON C.BinderyProcess=B.Code WHERE C.Code='" & slCode & "' ORDER BY B.Name", cnAccountMaster, adOpenKeyset, adLockReadOnly
                            If .RecordCount > 0 Then
                                If MsgBox("Want to load all bindery process for this binding type?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Load !") = vbYes Then
                                    Do Until .EOF
                                        fpSpread3.SetText 1, i, slName: fpSpread3.SetText 10, i, slCode 'Binding Type
                                        fpSpread3.SetText 2, i, .Fields("Name").Value: fpSpread3.SetText 11, i, .Fields("Code").Value 'Bindery Process
                                        i = i + 1
                                        .MoveNext
                                    Loop
                                End If
                            End If
                        End With
                    End If
                    SendKeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 2 Then
                .GetText 11, .ActiveRow, BinderyProcessCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "7"
                FrmGeneralMaster.MasterCode = BinderyProcessCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 2, .ActiveRow, slName: .SetText 11, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
            ElseIf .ActiveCol = 3 Then
                .GetText 12, .ActiveRow, CalcModeCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "20"
                FrmGeneralMaster.MasterCode = CalcModeCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 3, .ActiveRow, slName: .SetText 12, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
            ElseIf .ActiveCol = 4 Then
                .GetText 13, .ActiveRow, SizeGroupCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "10"
                FrmGeneralMaster.MasterCode = SizeGroupCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 4, .ActiveRow, slName: .SetText 13, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then SendKeys "{ENTER}"
            End If
        End If
    End With
End Sub
Private Sub AddToList()
    On Error Resume Next
    With rstAccountList
        .MoveFirst
        .Find "[Code] = '" & rstAccountMaster.Fields("Code").Value & "'"
        If .EOF Then .AddNew
        .Fields("Code").Value = rstAccountMaster.Fields("Code").Value
        .Fields("Name").Value = rstAccountMaster.Fields("Name").Value
        .Fields("AccountGroup").Value = IIf(AccountType = "01", Text10(0).Text, AccountGroupName)
        .Update
        .Sort = SortCol & " " & SortOrder
        .Find "[Code] = '" & rstAccountMaster.Fields("Code").Value & "'"
    End With
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstAccountList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then Timer1.Enabled = False: ShowProgressInStatusBar False
End Sub
Private Function ValidateForm() As Boolean
    If AccountType = "01" Then
        If CheckEmpty(Text2(0).Text, False) Then 'Name
            SSTab1.Tab = 1: Text2(0).SetFocus: ValidateForm = True
        ElseIf CheckDuplicate(cnAccountMaster, "AccountMaster", "Code", "Name", Trim(Text2(0).Text), rstAccountMaster.Fields("Code").Value, False) Then
            SSTab1.Tab = 1: Text2(0).SetFocus: ValidateForm = True
        ElseIf CheckEmpty(Text3(0).Text, False) Then 'Print Name
           SSTab1.Tab = 1: Text3(0).SetFocus: ValidateForm = True
        ElseIf CheckEmpty(Text10(0).Text, False) Then 'Account Group
            SSTab1.Tab = 1: Text10(0).SetFocus: ValidateForm = True
        End If
    End If
End Function
Private Sub LoadMasterList()
    If rstAccountGroupList.State = adStateOpen Then rstAccountGroupList.Close
    rstAccountGroupList.Open "SELECT Name As Col0, Code FROM GeneralMaster WHERE Type IN ('12','26') AND Code<'*99995' ORDER BY Name", cnAccountMaster, adOpenKeyset, adLockReadOnly
    rstAccountGroupList.ActiveConnection = Nothing
    If AccountType = "01" Then
        If rstElementList.State = adStateOpen Then rstElementList.Close
        rstElementList.Open "SELECT Name As Col0,Code FROM ElementMaster ORDER BY Name", cnAccountMaster, adOpenKeyset, adLockReadOnly
        rstElementList.ActiveConnection = Nothing
    ElseIf AccountType = "07" Then
        If rstSizeList.State = adStateOpen Then rstSizeList.Close
        rstSizeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type IN ('1','11') ORDER BY Name", cnAccountMaster, adOpenKeyset, adLockReadOnly
        rstSizeList.ActiveConnection = Nothing
    End If
End Sub
Private Sub DuplicateRecord()
    Dim AccountCode As String, AccountName As String
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    AccountCode = GenerateCode(cnAccountMaster, "SELECT MAX(Code) FROM AccountMaster", 6, "0"): AccountName = Trim(Left(rstAccountList.Fields("Name").Value, 36)) + " (D)"
    With cnAccountMaster
        .BeginTrans
        .Execute "IF OBJECT_ID('tempdb.dbo.#Tbl', 'U') IS NOT NULL  DROP TABLE #Tbl"
        .Execute "SELECT * INTO #Tbl FROM AccountMaster WHERE Code='" & rstAccountList.Fields("Code").Value & "'"
        .Execute "UPDATE  #Tbl SET Code='" & AccountCode & "',Name='" & AccountName & "',PrintName='" & AccountName & "'"
        .Execute "INSERT INTO AccountMaster SELECT * FROM #Tbl"
        .Execute "DROP TABLE #Tbl"
        .Execute "SELECT * INTO #Tbl FROM AccountChild04 WHERE Code='" & rstAccountList.Fields("Code").Value & "'"
        .Execute "UPDATE  #Tbl SET Code='" & AccountCode & "'"
        .Execute "INSERT INTO AccountChild04 SELECT * FROM #Tbl"
        .Execute "DROP TABLE #Tbl"
        .Execute "SELECT * INTO #Tbl FROM AccountChild05 WHERE Code='" & rstAccountList.Fields("Code").Value & "'"
        .Execute "UPDATE  #Tbl SET Code='" & AccountCode & "'"
        .Execute "INSERT INTO AccountChild05 SELECT * FROM #Tbl"
        .Execute "DROP TABLE #Tbl"
        .Execute "SELECT * INTO #Tbl FROM AccountChild07 WHERE Code='" & rstAccountList.Fields("Code").Value & "'"
        .Execute "UPDATE  #Tbl SET Code='" & AccountCode & "'"
        .Execute "INSERT INTO AccountChild07 SELECT * FROM #Tbl"
        .Execute "DROP TABLE #Tbl"
        .Execute "SELECT * INTO #Tbl FROM AccountChild08 WHERE Code='" & rstAccountList.Fields("Code").Value & "'"
        .Execute "UPDATE  #Tbl SET Code='" & AccountCode & "'"
        .Execute "INSERT INTO AccountChild08 SELECT * FROM #Tbl"
        .Execute "DROP TABLE #Tbl"
        .CommitTrans
    End With
    Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    Text1.Text = Trim(AccountName): SendKeys "{END}"
    MdiMainMenu.MousePointer = vbNormal
    Call MsgBox("Successfully duplicated the record !", vbInformation, App.Title)
    Exit Sub
ErrorHandler:
    MdiMainMenu.MousePointer = vbNormal
    DisplayError ("Failed to duplicate the record")
    cnAccountMaster.RollbackTrans
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
    Mh3dFrame2(Val(AccountType) - 1).Enabled = Not bVal
    Mh3dFrame2(8).Enabled = False
End Sub
Private Sub SetButtonsForNoRecord()
    If rstAccountList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
