VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVchSeriesMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Voucher Series Master"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8670
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
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8670
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   5760
      Left            =   -105
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
      _ExtentY        =   10160
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
      Picture         =   "VchSeriesMaster.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   5655
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   9975
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
         TabPicture(0)   =   "VchSeriesMaster.frx":001C
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
         TabPicture(1)   =   "VchSeriesMaster.frx":0038
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
            Left            =   960
            TabIndex        =   13
            Top             =   5175
            Width           =   7455
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4650
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   8202
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
            ColumnCount     =   5
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
               DataField       =   "Prefix"
               Caption         =   "Prefix"
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
            BeginProperty Column02 
               DataField       =   "Suffix"
               Caption         =   "Suffix"
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
            BeginProperty Column03 
               DataField       =   "VchNumbering"
               Caption         =   "Numbering"
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
            BeginProperty Column04 
               DataField       =   "VchName"
               Caption         =   "Vch. Name"
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
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               ScrollBars      =   3
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  ColumnWidth     =   2085.166
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   1484.787
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1980.284
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   5055
            Left            =   -74880
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   480
            Width           =   8295
            _Version        =   65536
            _ExtentX        =   14631
            _ExtentY        =   8916
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
            Picture         =   "VchSeriesMaster.frx":0054
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               Enabled         =   0   'False
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
               Left            =   6360
               MaxLength       =   40
               TabIndex        =   1
               TabStop         =   0   'False
               Top             =   220
               Width           =   1815
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               Enabled         =   0   'False
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
               TabStop         =   0   'False
               Top             =   220
               Width           =   1815
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
               MaxLength       =   40
               TabIndex        =   8
               Top             =   2160
               Width           =   6495
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
               TabIndex        =   7
               Top             =   1840
               Width           =   6495
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
               TabIndex        =   6
               Top             =   1525
               Width           =   6495
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   16
               Top             =   1525
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
               Caption         =   " Series Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "VchSeriesMaster.frx":0070
               Picture         =   "VchSeriesMaster.frx":008C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   17
               Top             =   1840
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
               Caption         =   " Prefix"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "VchSeriesMaster.frx":00A8
               Picture         =   "VchSeriesMaster.frx":00C4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   19
               Top             =   220
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
               Caption         =   " Code"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "VchSeriesMaster.frx":00E0
               Picture         =   "VchSeriesMaster.frx":00FC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   5040
               TabIndex        =   20
               Top             =   225
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
               Caption         =   "  Vch Type Code"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "VchSeriesMaster.frx":0118
               Picture         =   "VchSeriesMaster.frx":0134
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   21
               Top             =   2160
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
               Caption         =   " Suffix"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "VchSeriesMaster.frx":0150
               Picture         =   "VchSeriesMaster.frx":016C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   360
               Left            =   120
               TabIndex        =   23
               Top             =   1190
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   635
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
               Caption         =   " Voucher Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "VchSeriesMaster.frx":0188
               Picture         =   "VchSeriesMaster.frx":01A4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   360
               Left            =   120
               TabIndex        =   22
               Top             =   840
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   635
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
               Caption         =   " Numbering Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "VchSeriesMaster.frx":01C0
               Picture         =   "VchSeriesMaster.frx":01DC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   375
               Left            =   4320
               TabIndex        =   24
               Top             =   840
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   661
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
               Caption         =   "  Starting No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "VchSeriesMaster.frx":01F8
               Picture         =   "VchSeriesMaster.frx":0214
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
               Height          =   375
               Left            =   5400
               TabIndex        =   2
               Top             =   840
               Width           =   2790
               _Version        =   65536
               _ExtentX        =   4921
               _ExtentY        =   661
               Calculator      =   "VchSeriesMaster.frx":0230
               Caption         =   "VchSeriesMaster.frx":0250
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "VchSeriesMaster.frx":02BC
               Keys            =   "VchSeriesMaster.frx":02DA
               Spin            =   "VchSeriesMaster.frx":0324
               AlignHorizontal =   1
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#####0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#####0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999
               MinValue        =   1
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   1
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   375
               Left            =   4320
               TabIndex        =   25
               Top             =   1200
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   661
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
               Caption         =   " Type Series"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "VchSeriesMaster.frx":034C
               Picture         =   "VchSeriesMaster.frx":0368
            End
            Begin MSForms.ComboBox ComboBox2 
               Height          =   375
               Left            =   1665
               TabIndex        =   4
               ToolTipText     =   "Voucher "
               Top             =   1185
               Width           =   2700
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "4762;661"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontName        =   "Arial"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox ComboBox3 
               Height          =   375
               Left            =   5380
               TabIndex        =   5
               ToolTipText     =   "Catagary"
               Top             =   1185
               Width           =   2840
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "5009;661"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontName        =   "Arial"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   8280
               Y1              =   720
               Y2              =   720
            End
            Begin MSForms.ComboBox ComboBox1 
               Height          =   375
               Left            =   1665
               TabIndex        =   3
               Top             =   840
               Width           =   2700
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "4762;661"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontName        =   "Arial"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   3720
            TabIndex        =   18
            Top             =   0
            Width           =   4695
            _Version        =   65536
            _ExtentX        =   8281
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
            Picture         =   "VchSeriesMaster.frx":0384
            Picture         =   "VchSeriesMaster.frx":03A0
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
            TabIndex        =   15
            Top             =   5175
            Width           =   855
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
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
Attribute VB_Name = "FrmVchSeriesMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SL As Boolean 'Selection List
Public MasterCode As String  'Master to Modify
Dim CxnVchSeriesMaster As New ADODB.Connection
Dim rstVchSeriesList As New ADODB.Recordset
Dim rstVchSeriesMaster As New ADODB.Recordset
Dim rstCompanyList As New ADODB.Recordset
Dim SortOrder, PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim VchCode As String, VchType As String, ComboFlag As Boolean, EditMode As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    If Not SL Then MasterCode = ""
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    CxnVchSeriesMaster.CursorLocation = adUseClient
    CxnVchSeriesMaster.Open cnDatabase.ConnectionString
    rstVchSeriesList.Open "Select Name,Code,Prefix,Suffix,VchNumbering,VchName,StartNo From VchSeriesMaster Order By VchType,Name", CxnVchSeriesMaster, adOpenKeyset, adLockOptimistic
    rstCompanyList.Open "SELECT * FROM CompanyMaster Where FYCode='" & FYCode & "'", CxnVchSeriesMaster, adOpenKeyset, adLockReadOnly
    rstVchSeriesMaster.CursorLocation = adUseClient
    rstVchSeriesList.Filter = adFilterNone
    If rstVchSeriesList.RecordCount > 0 Then
        If CheckEmpty(MasterCode, False) Then
            rstVchSeriesList.MoveLast
        Else
            rstVchSeriesList.MoveFirst
            rstVchSeriesList.Find "[Code]='" & MasterCode & "'"
        End If
    End If
    ComboFlag = False
'Numbering Type
    ComboBox1.Clear
    ComboBox1.FontSize = 9
    ComboBox1.AddItem "Automatic", 0
    ComboBox1.AddItem "Manual", 1
'    ComboBox1.ListIndex = 0
'Voucher Type
    ComboBox2.Clear
    ComboBox2.FontSize = 9
    ComboBox2.AddItem "Purchase", 0 '1
    ComboBox2.AddItem "Purchase Return", 1 '2
    ComboBox2.AddItem "Sale Return", 2 '3
    ComboBox2.AddItem "Sales", 3 '4
    ComboBox2.AddItem "Purchase Challan IN", 4 '5
    ComboBox2.AddItem "Purchase Challan Out", 5 '6
    ComboBox2.AddItem "Sale Challan IN", 6 '7
    ComboBox2.AddItem "Sale Challan Out", 7 '8
    ComboBox2.AddItem "Purchase Order", 8 '17
    ComboBox2.AddItem "Sale Order", 9 '18
    ComboBox2.AddItem "Stock Tranfer", 10 '19
    ComboBox2.AddItem "Stock Genral", 11 '20
    ComboBox2.AddItem "Promotional Sale Challan Out", 12 '21
    ComboBox2.AddItem "Promotional Purchase Challan Out", 13 '22
    ComboBox2.AddItem "Purchase Quotation", 14 '23
    ComboBox2.AddItem "Sales Quotation", 15 '24
    ComboBox2.AddItem "Payments", 16 '51
    ComboBox2.AddItem "Receipts", 17 '52
    ComboBox2.AddItem "Journal", 18 '53
    ComboBox2.AddItem "Countra", 19 '54
    ComboBox2.AddItem "Credit Note", 20 '55
    ComboBox2.AddItem "Debit Note", 21 '56
'    ComboBox2.ListIndex = 0
    
    Set DataGrid1.DataSource = rstVchSeriesList
    DataGrid1.Columns(0).Width = 1860
    DataGrid1.Columns(1).Width = 855
    DataGrid1.Columns(2).Width = 810
    DataGrid1.Columns(3).Width = 1110
    DataGrid1.Columns(4).Width = 3040
    BusySystemIndicator False
    SSTab1.Tab = 0
    If Not (rstVchSeriesList.EOF Or rstVchSeriesList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstVchSeriesList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    SortOrder = "VchName"
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub LoadFields()
Dim i As Integer
    If rstVchSeriesMaster.EOF Or rstVchSeriesMaster.BOF Then Exit Sub
    Text5.Text = rstVchSeriesMaster.Fields("Code").Value
    MhRealInput1 = rstVchSeriesMaster.Fields("StartNo").Value
    Text2.Text = rstVchSeriesMaster.Fields("Name").Value
    Text6.Text = rstVchSeriesMaster.Fields("VchType").Value
    Text3.Text = rstVchSeriesMaster.Fields("Prefix").Value
    Text4.Text = rstVchSeriesMaster.Fields("Suffix").Value
    If rstVchSeriesMaster.Fields("VchNumbering").Value = "A" Then
        ComboBox1.ListIndex = 0
    Else
        ComboBox1.ListIndex = 1
    End If
    VchCode = Left(rstVchSeriesMaster.Fields("VchType").Value, 2)
    If VchCode >= 1 And VchCode <= 7 Then
        ComboBox2.ListIndex = Choose(VchCode, 0, 1, 2, 3, 4, 6, 7)
    ElseIf VchCode >= 17 And VchCode <= 24 Then
        ComboBox2.ListIndex = Choose(VchCode - 16, 8, 9, 10, 11, 12, 13, 14, 15)
    ElseIf VchCode >= 51 And VchCode <= 56 Then
        ComboBox2.ListIndex = Choose(VchCode - 50, 16, 17, 18, 19, 20, 21)
    End If
    rstVchSeriesMaster.MoveFirst
    LoadVchType
    For i = 0 To ComboBox3.ListCount
       ComboBox3.ListIndex = i
        If ComboBox3.Text = rstVchSeriesMaster.Fields("VchName").Value Then Exit For
    Next
    'EditMode
    If chkVchSeries Then
    MsgBox "You can Edit only [Prefix,Suffix,Numbering Type] " & vbCrLf & "Due to Transaction Generated Against this Voucher Series", vbInformation + vbDefaultButton1
        ComboBox1.Enabled = False: ComboBox2.Enabled = False: ComboBox3.Enabled = False: Text5.Enabled = False: Text6.Enabled = False
    End If
End Sub
Private Sub Form_Activate()
    MdiMainMenu.mnuVchSeriesMaster.Enabled = False
End Sub
Private Sub ComboBox2_Validate(Cancel As Boolean)
    If CheckEmpty(ComboBox2, True) Then
        Cancel = True
    Else
        VchTypeUpdate
    End If
End Sub
Private Sub VchTypeUpdate() 'ComboBox2
If ComboFlag = False Then Exit Sub
If ComboBox2.ListIndex >= 0 And ComboBox2.ListIndex <= 7 Then
    VchCode = ComboBox2.ListIndex + 1
ElseIf ComboBox2.ListIndex >= 8 And ComboBox2.ListIndex <= 15 Then
    VchCode = ComboBox2.ListIndex + 9
ElseIf ComboBox2.ListIndex >= 16 And ComboBox2.ListIndex <= 21 Then
    VchCode = ComboBox2.ListIndex + 35
End If
    VchCode = Format(VchCode, "00")
    Text6.Text = VchCode
Call LoadVchType
End Sub
Private Sub LoadVchType()
        ComboBox3.Clear
        ComboBox3.FontSize = 9
    If VchCode = "01" Then
        ComboBox3.AddItem "Purchase", 0
        ComboBox3.AddItem "Purchase Unit Cost", 1
        ComboBox3.AddItem "Purchase Jobwork Cost", 2
        ComboBox3.AddItem "Purchase Jobwork", 3
    ElseIf VchCode = "02" Then
        ComboBox3.AddItem "Purchase Return", 0
        ComboBox3.AddItem "Purchase Return Unit Cost", 1
        ComboBox3.AddItem "Purchase Return Jobwork Cost", 2
        ComboBox3.AddItem "Purchase Return Jobwork", 3
    ElseIf VchCode = "03" Then
        ComboBox3.AddItem "Sale Return", 0
        ComboBox3.AddItem "Sale Return Unit Cost", 1
        ComboBox3.AddItem "Sale Return Jobwork Cost", 2
        ComboBox3.AddItem "Sale Return Jobwork", 3
    ElseIf VchCode = "04" Then
        ComboBox3.AddItem "Sales", 0
        ComboBox3.AddItem "Sales Unit Cost", 1
        ComboBox3.AddItem "Sales Jobwork Cost", 2
        ComboBox3.AddItem "Sales Jobwork", 3
    ElseIf VchCode = "05" Then
        ComboBox3.AddItem "Purchase Challan IN", 0
        ComboBox3.AddItem "Purchase Challan IN (Jobwork)", 1
    ElseIf VchCode = "06" Then
        ComboBox3.AddItem "Purchase Challan Out", 0
        ComboBox3.AddItem "Purchase Challan Out (Jobworj)", 1
    ElseIf VchCode = "07" Then
        ComboBox3.AddItem "Sale Challan IN", 0
        ComboBox3.AddItem "Sale Challan IN (Jobwork)", 1
    ElseIf VchCode = "08" Then
        ComboBox3.AddItem "Sale Challan Out", 0
        ComboBox3.AddItem "Sale Challan Out (Jobwork)", 1
    ElseIf VchCode = "17" Then
        ComboBox3.AddItem "Purchase Order", 0
    ElseIf VchCode = "18" Then
        ComboBox3.AddItem "Sale Order", 0
    ElseIf VchCode = "19" Then
        ComboBox3.AddItem "Stock Tranfer", 0
    ElseIf VchCode = "20" Then
        ComboBox3.AddItem "Stock Genral", 0
    ElseIf VchCode = "21" Then
        ComboBox3.AddItem "Promotional Sale Challan Out", 0
    ElseIf VchCode = "22" Then
        ComboBox3.AddItem "Promotional Purchase Challan Out", 0
    ElseIf VchCode = "23" Then
        ComboBox3.AddItem "Purchase Quotation", 0
        ComboBox3.AddItem "Purchase Quotation Unit Cost", 1
        ComboBox3.AddItem "Purchase Quotation Jobwork Cost", 2
        ComboBox3.AddItem "Purchase Quotation Jobwork", 3
    ElseIf VchCode = "24" Then
        ComboBox3.AddItem "Sales Quotation", 0
        ComboBox3.AddItem "Sales Quotation Unit Cost", 1
        ComboBox3.AddItem "Sales Quotation Jobwork Cost", 2
        ComboBox3.AddItem "Sales Quotation Jobwork", 3
    ElseIf VchCode = "51" Then
        ComboBox3.AddItem "Payments", 0
    ElseIf VchCode = "52" Then
        ComboBox3.AddItem "Receipts", 0
    ElseIf VchCode = "53" Then
        ComboBox3.AddItem "Journal", 0
    ElseIf VchCode = "54" Then
        ComboBox3.AddItem "Countra", 0
    ElseIf VchCode = "55" Then
        ComboBox3.AddItem "Credit Note", 0
     ElseIf VchCode = "56" Then
        ComboBox3.AddItem "Debit Note", 0
    End If
End Sub
Private Sub ComboBox3_Validate(Cancel As Boolean)
    If CheckEmpty(ComboBox3, True) Then
'        Cancel = True
    Else
        VchSeriseUpdate
    End If
End Sub
Private Sub VchSeriseUpdate()
If ComboFlag = False Or ComboBox3.ListIndex = -1 Then Exit Sub
    If VchCode = 1 Then
        VchType = Choose(ComboBox3.ListIndex + 1, "PF", "PU", "PC", "PJ")
    ElseIf VchCode = 2 Then
        VchType = Choose(ComboBox3.ListIndex + 1, "OF", "OU", "OC", "OJ")
    ElseIf VchCode = 3 Then
        VchType = Choose(ComboBox3.ListIndex + 1, "TF", "TU", "TC", "TJ")
    ElseIf VchCode = 4 Then
        VchType = Choose(ComboBox3.ListIndex + 1, "SF", "SU", "SC", "SJ")
    ElseIf VchCode = 5 Then
        VchType = Choose(ComboBox3.ListIndex + 1, "RF", "FR")
    ElseIf VchCode = 6 Then
        VchType = Choose(ComboBox3.ListIndex + 1, "IF", "FI")
    ElseIf VchCode = 7 Then
        VchType = Choose(ComboBox3.ListIndex + 1, "RF", "FR")
    ElseIf VchCode = 8 Then
        VchType = Choose(ComboBox3.ListIndex + 1, "IF", "FI")
    ElseIf VchCode = 17 Then
        VchType = "PO"
    ElseIf VchCode = 18 Then
        VchType = "SO"
    ElseIf VchCode = 19 Then
        VchType = "ST"
    ElseIf VchCode = 20 Then
        VchType = "JR"
    ElseIf VchCode = 21 Then
        VchType = "RF"
    ElseIf VchCode = 22 Then
        VchType = "IF"
    ElseIf VchCode = 23 Then
        VchType = Choose(ComboBox3.ListIndex + 1, "PQ", "ZU", "ZC", "ZJ")
    ElseIf VchCode = 24 Then
        VchType = Choose(ComboBox3.ListIndex + 1, "SQ", "QU", "QC", "QJ")
    ElseIf VchCode = 51 Then
        VchType = "PI"
    ElseIf VchCode = 52 Then
        VchType = "PR"
    ElseIf VchCode = 53 Then
        VchType = "JE"
    ElseIf VchCode = 54 Then
        VchType = "CE"
    ElseIf VchCode = 55 Then
        VchType = "CN"
    ElseIf VchCode = 56 Then
        VchType = "DN"
    End If
    Text6.Text = Format(VchCode, "00") + VchType
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
                    ComboBox1.Enabled = True: ComboBox2.Enabled = True: ComboBox3.Enabled = True: Text5.Enabled = True: Text6.Enabled = True
                    Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                End If
            End If
            KeyCode = 0: ComboFlag = False
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        ComboFlag = True
        ComboBox1.Enabled = True: ComboBox2.Enabled = True: ComboBox3.Enabled = True: Text5.Enabled = True: Text6.Enabled = True
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
       ComboFlag = False
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
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstVchSeriesList.Fields("Code").Value: slName = rstVchSeriesList.Fields("Name").Value: KeyCode = 0: Unload Me: Exit Sub
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
Private Sub ViewRecord()
    ClearFields
    If rstVchSeriesList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstVchSeriesMaster.State = adStateOpen Then
       rstVchSeriesMaster.Close
    End If
    rstVchSeriesMaster.Open "Select * From VchSeriesMaster Where Code = '" & FixQuote(rstVchSeriesList.Fields("Code").Value) & "'", CxnVchSeriesMaster, adOpenKeyset, adLockOptimistic
    If rstVchSeriesMaster.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    MhRealInput1 = 1
    ComboBox3.Clear
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstVchSeriesMaster.RecordCount = 0 Then Exit Sub
    If rstVchSeriesMaster.State = adStateOpen Then rstVchSeriesMaster.Close
    rstVchSeriesMaster.CursorLocation = adUseServer
    rstVchSeriesMaster.Open "Select * From VchSeriesMaster Where Code = '" & FixQuote(rstVchSeriesList.Fields("Code").Value) & "'", CxnVchSeriesMaster, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
'    rstVchSeriesMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    MhRealInput1.SetFocus
    blnRecordExist = True
    CxnVchSeriesMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstVchSeriesMaster.EOF Or rstVchSeriesMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        Text5.Text = GenerateCode(CxnVchSeriesMaster, "Select Max(Code) From VchSeriesMaster", 6, "0")
        rstVchSeriesMaster.Fields("Code").Value = Text5.Text
        rstVchSeriesMaster.Fields("CreatedBy").Value = UserCode
        rstVchSeriesMaster.Fields("CreatedOn").Value = Now()
        rstVchSeriesMaster.Fields("Recordstatus").Value = "N"
    Else
        rstVchSeriesMaster.Fields("ModifiedBy").Value = UserCode
        rstVchSeriesMaster.Fields("ModifiedOn").Value = Now()
        rstVchSeriesMaster.Fields("Recordstatus").Value = "M"
    End If
    rstVchSeriesMaster.Fields("Name").Value = Trim(Text2.Text)
    rstVchSeriesMaster.Fields("VchType").Value = Trim(Text6.Text)
    rstVchSeriesMaster.Fields("Prefix").Value = Trim(Text3.Text)
    rstVchSeriesMaster.Fields("Suffix").Value = Trim(Text4.Text)
    rstVchSeriesMaster.Fields("VchNumbering").Value = IIf(ComboBox1.ListIndex = 0, "A", "M")
    rstVchSeriesMaster.Fields("VchName").Value = ComboBox3.Text
    rstVchSeriesMaster.Fields("FYCode").Value = FYCode
    rstVchSeriesMaster.Fields("StartNo").Value = MhRealInput1.Value
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstVchSeriesList.MoveFirst
    rstVchSeriesList.Find "[Code] = '" & rstVchSeriesMaster.Fields("Code").Value & "'"
    If rstVchSeriesList.EOF Then
       rstVchSeriesList.AddNew
       rstVchSeriesList.Fields("Code").Value = rstVchSeriesMaster.Fields("Code").Value
    End If
    rstVchSeriesList.Fields("Name").Value = rstVchSeriesMaster.Fields("Name").Value
    rstVchSeriesList.Update
    rstVchSeriesList.Sort = "Name Asc"
    rstVchSeriesList.Find "[Code] = '" & rstVchSeriesMaster.Fields("Code").Value & "'"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
    Else
        If Me.Tag <> "S" Then slCode = "": slName = ""
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstVchSeriesList)
    Call CloseRecordset(rstVchSeriesMaster)
    Call CloseRecordset(rstCompanyList)
    Call CloseConnection(CxnVchSeriesMaster)
    ShowProgressInStatusBar False
    MdiMainMenu.mnuVchSeriesMaster.Enabled = True
End Sub

Private Sub Text1_Change()
On Error Resume Next
    With rstVchSeriesList
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
    If rstVchSeriesList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstVchSeriesList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstVchSeriesList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstVchSeriesList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstVchSeriesList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstVchSeriesList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstVchSeriesList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstVchSeriesList
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
            If Not (rstVchSeriesList.EOF Or rstVchSeriesList.BOF) Then
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
        MhRealInput1.SetFocus
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
'    If CheckEmpty(Text4.Text, False) Then Cancel = True
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Dim HiLiteRecord As Boolean
    If Button.Index = 1 Then
        If rstVchSeriesMaster.State = adStateOpen Then
           rstVchSeriesMaster.Close
        End If
        rstVchSeriesMaster.Open "Select * From VchSeriesMaster Where Code = ''", CxnVchSeriesMaster, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstVchSeriesMaster) Then
            Call SetButtons(False)
            SSTab1.Tab = 1
            MhRealInput1.SetFocus
            blnRecordExist = False
            CxnVchSeriesMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstVchSeriesList.RecordCount = 0 Then Exit Sub
        'If chkVchSeries Then DisplayError ("You can Edit only [Prefix,Suffix,Numbering Type] ,Due to Transaction Generated Against this Voucher Series")
        SSTab1.Tab = 1
        EditRecord
'        If Not chkVchSeries Then
'            ComboBox1.Enabled = True: ComboBox2.Enabled = True: ComboBox3.Enabled = True
'        Else
'            ComboBox1.Enabled = True: ComboBox2.Enabled = False: ComboBox3.Enabled = False: MhRealInput1.Visible = False
'        End If
    ElseIf Button.Index = 3 Then
        If rstVchSeriesList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Master")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            If chkVchSeries Then DisplayError ("Deletion Failed Due to Transaction Generated Against this Voucher Series"): Exit Sub
            MdiMainMenu.MousePointer = vbHourglass
            CxnVchSeriesMaster.Execute "DELETE FROM VchSeriesMaster WHERE Code = '" & rstVchSeriesList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstVchSeriesList.Delete
                rstVchSeriesList.MoveNext
                If rstVchSeriesList.RecordCount > 0 And rstVchSeriesList.EOF Then rstVchSeriesList.MoveLast
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
        If UpdateRecord(rstVchSeriesMaster) Then
            Call UpdateUserAction("Vch Series Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), cnDatabase)
            AddToList
            CxnVchSeriesMaster.CommitTrans
            If rstVchSeriesMaster.State = adStateOpen Then
                rstVchSeriesMaster.Close
            End If
            rstVchSeriesMaster.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstVchSeriesMaster) Then
            CxnVchSeriesMaster.RollbackTrans
            If rstVchSeriesMaster.State = adStateOpen Then
                rstVchSeriesMaster.Close
            End If
            rstVchSeriesMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstVchSeriesList.ActiveConnection = CxnVchSeriesMaster
        Do While Not RefreshRecord(rstVchSeriesList)
        Loop
        Set DataGrid1.DataSource = rstVchSeriesList
        rstVchSeriesList.ActiveConnection = Nothing
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
        If rstVchSeriesList.RecordCount > 0 Then rstVchSeriesList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstVchSeriesList.RecordCount > 0 Then
           rstVchSeriesList.MovePrevious
           If rstVchSeriesList.BOF Then
              rstVchSeriesList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstVchSeriesList.RecordCount > 0 Then
           rstVchSeriesList.MoveNext
           If rstVchSeriesList.EOF Then
              rstVchSeriesList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstVchSeriesList.RecordCount > 0 Then rstVchSeriesList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        CloseForm Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstVchSeriesList.EOF Or rstVchSeriesList.BOF) Then
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
        rstVchSeriesList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstVchSeriesList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstVchSeriesList.EOF Or rstVchSeriesList.BOF) Then
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
    If rstVchSeriesList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstVchSeriesMaster.EOF Or rstVchSeriesMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
'        Cancel = True
    ElseIf CheckDuplicate(CxnVchSeriesMaster, "VchSeriesMaster", "Code", "Name+VchType", Text2.Text + Text6.Text, rstVchSeriesMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = rstCompanyList.Fields("Alias").Value + "/"
        Text4.Text = "/" + Right(Text6.Text, 2)  '
    End If
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(ComboBox1.Text, False) Then
        ComboBox1.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(ComboBox2.Text, False) Then
        ComboBox2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(ComboBox3.Text, False) Then
        ComboBox3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text2.Text, False) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnVchSeriesMaster, "VchSeriesMaster", "Code", "Name+VchType", Text2.Text + Text6.Text, rstVchSeriesMaster.Fields("Code").Value, False) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
'        Text3.SetFocus
'        CheckMandatoryFields = True
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
    If SrchFor = "Name" Then rstVchSeriesList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
Private Function chkVchSeries() As Boolean
    On Error GoTo ErrHandler
    Dim rstChkVchSeries As New ADODB.Recordset
    rstChkVchSeries.Open "SELECT Top (1) VchSeries FROM JobworkBVParent P WHERE VchSeries='" & rstVchSeriesList.Fields("Code").Value & "' ", CxnVchSeriesMaster, adOpenKeyset, adLockReadOnly
    If rstChkVchSeries.RecordCount > 0 Then chkVchSeries = True
    Call CloseRecordset(rstChkVchSeries)
    Exit Function
ErrHandler:
    Call CloseRecordset(rstChkVchSeries)
End Function
