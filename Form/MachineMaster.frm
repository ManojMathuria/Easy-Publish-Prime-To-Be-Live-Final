VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMachineMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Master"
   ClientHeight    =   7050
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   7590
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7410
      Left            =   15
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13361
      _ExtentY        =   13070
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
      Picture         =   "MachineMaster.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   6810
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   12012
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
         TabPicture(0)   =   "MachineMaster.frx":001C
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
         TabPicture(1)   =   "MachineMaster.frx":0038
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
            Left            =   600
            TabIndex        =   16
            Top             =   6330
            Width           =   6615
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5820
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   450
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   10266
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
                  ColumnWidth     =   6524.788
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6180
            Left            =   -74880
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   480
            Width           =   7095
            _Version        =   65536
            _ExtentX        =   12515
            _ExtentY        =   10901
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
            Picture         =   "MachineMaster.frx":0054
            Begin TDBTime6Ctl.TDBTime MhTimeInput1 
               Height          =   330
               Left            =   2280
               TabIndex        =   9
               Top             =   2310
               Width           =   2355
               _Version        =   65536
               _ExtentX        =   4154
               _ExtentY        =   582
               Caption         =   "MachineMaster.frx":0070
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "MachineMaster.frx":00DC
               Spin            =   "MachineMaster.frx":012C
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "hh:nn"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "hh:nn"
               HighlightText   =   0
               Hour12Mode      =   1
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxTime         =   0.999988425925926
               MidnightMode    =   0
               MinTime         =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "00:00"
               ValidateMode    =   0
               ValueVT         =   1926037511
               Value           =   0
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
               Left            =   2280
               MaxLength       =   40
               TabIndex        =   1
               Top             =   420
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
               Left            =   2280
               MaxLength       =   40
               TabIndex        =   0
               Top             =   105
               Width           =   4695
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   19
               Top             =   105
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
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
               Picture         =   "MachineMaster.frx":0154
               Picture         =   "MachineMaster.frx":0170
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   20
               Top             =   420
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
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
               Picture         =   "MachineMaster.frx":018C
               Picture         =   "MachineMaster.frx":01A8
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
               Height          =   330
               Left            =   2280
               TabIndex        =   2
               Top             =   740
               Width           =   4695
               _Version        =   65536
               _ExtentX        =   8281
               _ExtentY        =   582
               Calculator      =   "MachineMaster.frx":01C4
               Caption         =   "MachineMaster.frx":01E4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "MachineMaster.frx":0250
               Keys            =   "MachineMaster.frx":026E
               Spin            =   "MachineMaster.frx":02B8
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   740
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
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
               Caption         =   " Units"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "MachineMaster.frx":02E0
               Picture         =   "MachineMaster.frx":02FC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   23
               Top             =   1050
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
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
               Caption         =   " Make ready Time (Min)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "MachineMaster.frx":0318
               Picture         =   "MachineMaster.frx":0334
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   3
               Left            =   120
               TabIndex        =   24
               Top             =   1370
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
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
               Caption         =   " Efficiency/Hr"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "MachineMaster.frx":0350
               Picture         =   "MachineMaster.frx":036C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   25
               Top             =   1680
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
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
               Caption         =   " Min. Area  (W && L)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "MachineMaster.frx":0388
               Picture         =   "MachineMaster.frx":03A4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   26
               Top             =   1995
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
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
               Caption         =   " Max. Area  (W && L)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "MachineMaster.frx":03C0
               Picture         =   "MachineMaster.frx":03DC
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   330
               Left            =   2280
               TabIndex        =   4
               Top             =   1370
               Width           =   4695
               _Version        =   65536
               _ExtentX        =   8281
               _ExtentY        =   582
               Calculator      =   "MachineMaster.frx":03F8
               Caption         =   "MachineMaster.frx":0418
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "MachineMaster.frx":0484
               Keys            =   "MachineMaster.frx":04A2
               Spin            =   "MachineMaster.frx":04EC
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1902837765
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   330
               Left            =   2280
               TabIndex        =   5
               Top             =   1680
               Width           =   2355
               _Version        =   65536
               _ExtentX        =   4145
               _ExtentY        =   582
               Calculator      =   "MachineMaster.frx":0514
               Caption         =   "MachineMaster.frx":0534
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "MachineMaster.frx":05A0
               Keys            =   "MachineMaster.frx":05BE
               Spin            =   "MachineMaster.frx":0608
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1902837765
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   330
               Left            =   4620
               TabIndex        =   6
               Top             =   1680
               Width           =   2355
               _Version        =   65536
               _ExtentX        =   4154
               _ExtentY        =   582
               Calculator      =   "MachineMaster.frx":0630
               Caption         =   "MachineMaster.frx":0650
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "MachineMaster.frx":06BC
               Keys            =   "MachineMaster.frx":06DA
               Spin            =   "MachineMaster.frx":0724
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99.99
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   330
               Left            =   2280
               TabIndex        =   7
               Top             =   1995
               Width           =   2355
               _Version        =   65536
               _ExtentX        =   4154
               _ExtentY        =   582
               Calculator      =   "MachineMaster.frx":074C
               Caption         =   "MachineMaster.frx":076C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "MachineMaster.frx":07D8
               Keys            =   "MachineMaster.frx":07F6
               Spin            =   "MachineMaster.frx":0840
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99.99
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
               Height          =   330
               Left            =   4620
               TabIndex        =   8
               Top             =   1995
               Width           =   2355
               _Version        =   65536
               _ExtentX        =   4154
               _ExtentY        =   582
               Calculator      =   "MachineMaster.frx":0868
               Caption         =   "MachineMaster.frx":0888
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "MachineMaster.frx":08F4
               Keys            =   "MachineMaster.frx":0912
               Spin            =   "MachineMaster.frx":095C
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99.99
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   27
               Top             =   2630
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
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
               Caption         =   " Category"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "MachineMaster.frx":0984
               Picture         =   "MachineMaster.frx":09A0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   28
               Top             =   2310
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
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
               Caption         =   " Start && End Time"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "MachineMaster.frx":09BC
               Picture         =   "MachineMaster.frx":09D8
            End
            Begin TDBTime6Ctl.TDBTime MhTimeInput2 
               Height          =   330
               Left            =   4620
               TabIndex        =   10
               Top             =   2310
               Width           =   2355
               _Version        =   65536
               _ExtentX        =   4154
               _ExtentY        =   582
               Caption         =   "MachineMaster.frx":09F4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "MachineMaster.frx":0A60
               Spin            =   "MachineMaster.frx":0AB0
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "hh:nn"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "hh:nn"
               HighlightText   =   0
               Hour12Mode      =   1
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxTime         =   0.999988425925926
               MidnightMode    =   0
               MinTime         =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "00:00"
               ValidateMode    =   0
               ValueVT         =   1926037511
               Value           =   0
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
               Height          =   330
               Left            =   2280
               TabIndex        =   3
               Top             =   1050
               Width           =   4695
               _Version        =   65536
               _ExtentX        =   8281
               _ExtentY        =   582
               Calculator      =   "MachineMaster.frx":0AD8
               Caption         =   "MachineMaster.frx":0AF8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "MachineMaster.frx":0B64
               Keys            =   "MachineMaster.frx":0B82
               Spin            =   "MachineMaster.frx":0BCC
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99
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
               Height          =   3000
               Left            =   120
               TabIndex        =   29
               Top             =   3060
               Width           =   6855
               _Version        =   524288
               _ExtentX        =   12091
               _ExtentY        =   5292
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
               MaxCols         =   5
               MaxRows         =   50
               ScrollBars      =   2
               SpreadDesigner  =   "MachineMaster.frx":0BF4
            End
            Begin MSForms.ComboBox cmbCategory 
               Height          =   330
               Left            =   2280
               TabIndex        =   11
               Top             =   2630
               Width           =   4695
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "8281;582"
               ListWidth       =   4762
               MatchEntry      =   0
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Calibri"
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   2520
            TabIndex        =   22
            Top             =   0
            Width           =   4815
            _Version        =   65536
            _ExtentX        =   8493
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
            Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D/F8->Delete  Ctrl+S->Save"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "MachineMaster.frx":12C1
            Picture         =   "MachineMaster.frx":12DD
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
            TabIndex        =   18
            Top             =   6330
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   13
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
Attribute VB_Name = "FrmMachineMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SL As Boolean 'Selection List
Public MasterCode As String  'Master to Modify
Dim cnMachineMaster As New ADODB.Connection
Dim rstMachineList As New ADODB.Recordset
Dim rstMachineMaster As New ADODB.Recordset
Dim rstMachineChild As New ADODB.Recordset
Dim EditMode As Boolean
Dim SortCol, PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static SortOrder  As String
    SortCol = DataGrid1.Columns(ColIndex).DataField
    If SortOrder = "A" Then
        rstMachineList.Sort = "[" + SortCol & "] Desc"
        SortOrder = "D"
    Else
        rstMachineList.Sort = "[" + SortCol & "] Asc"
        SortOrder = "A"
    End If
    DataGrid1.ClearSelCols
    If Not (rstMachineList.EOF Or rstMachineList.BOF) Then
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
    WheelHook DataGrid1
    BusySystemIndicator True
    cnMachineMaster.CursorLocation = adUseClient: cnMachineMaster.Open cnDatabase.ConnectionString
    If rstMachineList.State Then rstMachineList.Close
    rstMachineList.Open "SELECT Name,Code FROM MachineMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockOptimistic
    rstMachineMaster.CursorLocation = adUseClient
    rstMachineList.Filter = adFilterNone
    If rstMachineList.RecordCount > 0 Then
        rstMachineList.MoveFirst
        If Not CheckEmpty(MasterCode, False) Then rstMachineList.Find "[Code]='" & MasterCode & "'"
    End If
    Set DataGrid1.DataSource = rstMachineList
    BusySystemIndicator False
    SSTab1.Tab = 0
    If Not (rstMachineList.EOF Or rstMachineList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstMachineList.ActiveConnection = Nothing
    cmbCategory.AddItem "Printing", 0
    cmbCategory.AddItem "Binding", 1
    cmbCategory.AddItem "Cutting", 2
    cmbCategory.AddItem "Fabrication", 3
    cmbCategory.AddItem "Finishing", 4
    SetButtonsForNoRecord
    SortCol = "Name"
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    MdiMainMenu.mnuMachineMaster.Enabled = False
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
        KeyCode = 0
    ElseIf ((Shift = 0 And KeyCode = vbKeyF8) Or (Shift = vbCtrlMask And KeyCode = vbKeyD)) And Toolbar1.Buttons.Item(3).Enabled Then
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
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstMachineList.Fields("Code").Value: slName = rstMachineList.Fields("Name").Value: KeyCode = 0: Unload Me: Exit Sub
            Else
                SSTab1.Tab = 1
                SSTab1.SetFocus
            End If
        Else
           If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}" '    Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0 'KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1 Else If Me.Tag <> "S" Then slCode = "": slName = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstMachineList)
    Call CloseRecordset(rstMachineChild)
    Call CloseRecordset(rstMachineMaster)
    Call CloseConnection(cnMachineMaster)
    ShowProgressInStatusBar False
    MdiMainMenu.mnuMachineMaster.Enabled = True
End Sub
'Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Dim Qty As Variant, Sets As Variant, Hours As Variant, Efficiency As Variant
'        With fpSpread1
''            If .ActiveCol = 3 Then
''                .GetText 1, .ActiveRow, Qty
''                .GetText 2, .ActiveRow, Sets
''                .GetText 3, .ActiveRow, Hours
''                If Qty = "" Or Sets = "" Or Hours = "" Then Exit Sub
''                Efficiency = Round((Qty * Sets) / (((Hours * 60) - (MhRealInput2 * Sets)) / 60))
''                .SetText 4, .ActiveRow, Efficiency
''                .SetText 5, .ActiveRow, Hours * Efficiency
''            End If
'        End With
'End Sub
Private Sub Text1_Change()
    On Error Resume Next
    With rstMachineList
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
    With rstMachineList
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
            If Not (rstMachineList.EOF Or rstMachineList.BOF) Then
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
    Dim HiLiteRecord As Boolean, UpdateFlag As Integer, i As Integer, EfficiencyCode As Variant
    If Button.Index = 1 Then
        If rstMachineMaster.State = adStateOpen Then rstMachineMaster.Close
        rstMachineMaster.Open "SELECT * FROM MachineMaster WHERE Code=''", cnDatabase, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstMachineMaster) Then
           Call SetButtons(False)
           SSTab1.Tab = 1
           Text2.SetFocus
           blnRecordExist = False
           cnMachineMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstMachineList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstMachineList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Master"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnDatabase.Execute "DELETE FROM MachineMaster WHERE Code = '" & rstMachineList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstMachineList.Delete
                rstMachineList.MoveNext
                If rstMachineList.RecordCount > 0 And rstMachineList.EOF Then rstMachineList.MoveLast
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
        If UpdateRecord(rstMachineMaster) Then
            If UpdateEfficiencyList("D") Then
                UpdateFlag = 1
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.SetActiveCell 1, i
                    fpSpread1.GetText 4, i, EfficiencyCode
                    If Val(EfficiencyCode) <> 0 Then
                        If Not UpdateEfficiencyList("I") Then
                            UpdateFlag = 0
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        If UpdateFlag Then
            Call UpdateUserAction("Machine Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), cnDatabase)
            AddToList
            cnMachineMaster.CommitTrans
            If rstMachineMaster.State = adStateOpen Then rstMachineMaster.Close
            rstMachineMaster.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstMachineMaster) Then
        cnMachineMaster.RollbackTrans
           If rstMachineMaster.State = adStateOpen Then rstMachineMaster.Close
           rstMachineMaster.CursorLocation = adUseClient
           Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstMachineList.ActiveConnection = cnDatabase
        Do Until RefreshRecord(rstMachineList): Loop
        Set DataGrid1.DataSource = rstMachineList
        rstMachineList.ActiveConnection = Nothing
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
        If rstMachineList.RecordCount > 0 Then rstMachineList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstMachineList.RecordCount > 0 Then
           rstMachineList.MovePrevious
           If rstMachineList.BOF Then rstMachineList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstMachineList.RecordCount > 0 Then
           rstMachineList.MoveNext
           If rstMachineList.EOF Then rstMachineList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstMachineList.RecordCount > 0 Then rstMachineList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstMachineList.EOF Or rstMachineList.BOF) Then
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
    If rstMachineList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstMachineMaster.EOF Or rstMachineMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnDatabase, "MachineMaster", "Code", "Name", Trim(Text2.Text), rstMachineMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = Text2.Text
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstMachineList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstMachineMaster.State = adStateOpen Then rstMachineMaster.Close
    rstMachineMaster.Open "SELECT * FROM MachineMaster WHERE Code='" & FixQuote(rstMachineList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    If rstMachineMaster.RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    If rstMachineChild.State = adStateOpen Then rstMachineChild.Close
    rstMachineChild.Open "SELECT * FROM MachineChild WHERE Code='" & FixQuote(rstMachineList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    'If rstMachineChild.RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    MhRealInput1.Value = 0
    MhRealInput2.Value = 0
    MhRealInput3.Value = 0
    MhRealInput4.Value = 0
    MhRealInput5.Value = 0
    MhRealInput6.Value = 0
    MhRealInput7.Value = 0
    cmbCategory.ListIndex = 0
    MhTimeInput1.Value = 0
    MhTimeInput2.Value = 0
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    fpSpread1.SetActiveCell 1, 1
End Sub
Private Sub LoadFields()
Dim EfficiencyCode As Variant, Qty As Variant, Sets As Variant, Hours As Variant, Efficiency As Variant, i As Integer
    With rstMachineMaster
        If .EOF Or .BOF Then Exit Sub
        Text2.Text = .Fields("Name").Value
        Text3.Text = .Fields("PrintName").Value
        MhRealInput1.Value = Val(.Fields("Units").Value)
        MhRealInput2.Value = Val(.Fields("MakeReadyTime").Value)
        MhRealInput3.Value = Val(.Fields("Efficiency").Value)
        MhRealInput4.Value = Val(.Fields("MinSizeWidth").Value)
        MhRealInput5.Value = Val(.Fields("MinSizeLength").Value)
        MhRealInput6.Value = Val(.Fields("MaxSizeWidth").Value)
        MhRealInput7.Value = Val(.Fields("MaxSizeLength").Value)
        cmbCategory.ListIndex = Val(.Fields("Category").Value) - 1
        MhTimeInput1.Value = .Fields("StartTime").Value
        MhTimeInput2.Value = .Fields("EndTime").Value
    End With
    If rstMachineChild.State = adStateOpen Then rstMachineChild.Close
    rstMachineChild.Open "SELECT * FROM MachineChild WHERE Code='" & FixQuote(rstMachineList.Fields("Code").Value) & "'", cnMachineMaster, adOpenKeyset, adLockOptimistic
    rstMachineChild.ActiveConnection = Nothing
    If rstMachineChild.RecordCount = 0 Then Exit Sub
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    With rstMachineChild
    i = 0
        If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                With fpSpread1
                    i = i + 1
                    Qty = Val(rstMachineChild.Fields("QTY").Value)
                    Sets = Val(rstMachineChild.Fields("Sets").Value)
                    Hours = Val(rstMachineChild.Fields("Hours").Value)
                    Efficiency = Val(rstMachineChild.Fields("Efficiency").Value)
                    EfficiencyCode = Val(rstMachineChild.Fields("Code").Value)
'Load Data
                    .SetText 1, i, Qty
                    .SetText 2, i, Sets
                    .SetText 3, i, Hours
                    .SetText 4, i, Efficiency
                    .SetText 5, i, Hours * Efficiency
                End With
                    .MoveNext
        Loop
    End With
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstMachineMaster.RecordCount = 0 Then Exit Sub
    If rstMachineMaster.State = adStateOpen Then rstMachineMaster.Close
    rstMachineMaster.CursorLocation = adUseServer
    rstMachineMaster.Open "SELECT * FROM MachineMaster WHERE Code='" & FixQuote(rstMachineList.Fields("Code").Value) & "'", cnDatabase, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstMachineMaster.Fields("PrintStatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    cnMachineMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 1
End Sub
Private Sub SaveFields()
    With rstMachineMaster
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnDatabase, "SELECT MAX(Code) FROM MachineMaster", 6, "0")
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
        .Fields("Units").Value = MhRealInput1.Value
        .Fields("MakeReadyTime").Value = MhRealInput2.Value
        .Fields("Efficiency").Value = MhRealInput3.Value
        .Fields("MinSizeWidth").Value = MhRealInput4.Value
        .Fields("MinSizeLength").Value = MhRealInput5.Value
        .Fields("MaxSizeWidth").Value = MhRealInput6.Value
        .Fields("MaxSizeLength").Value = MhRealInput7.Value
        .Fields("StartTime").Value = Format(MhTimeInput1.Value, "hh:mm")
        .Fields("EndTime").Value = Format(MhTimeInput2.Value, "hh:mm")
        .Fields("Category").Value = cmbCategory.ListIndex + 1
        .Fields("PrintStatus").Value = "N"
    End With
End Sub
Private Sub AddToList()
    On Error Resume Next
    With rstMachineList
        .MoveFirst
        .Find "[Code] = '" & rstMachineMaster.Fields("Code").Value & "'"
        If .EOF Then .AddNew: .Fields("Code").Value = rstMachineMaster.Fields("Code").Value
        .Fields("Name").Value = rstMachineMaster.Fields("Name").Value
        .Update
        .Sort = "Name Asc"
        .Find "[Code] = '" & rstMachineMaster.Fields("Code").Value & "'"
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        Text2.SetFocus: CheckMandatoryFields = True
    ElseIf CheckDuplicate(cnDatabase, "MachineMaster", "Code", "Name", Trim(Text2.Text), rstMachineMaster.Fields("Code").Value, False) Then
        Text2.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus: CheckMandatoryFields = True
    End If
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstMachineList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then Timer1.Enabled = False: ShowProgressInStatusBar False
End Sub
Private Function UpdateEfficiencyList(ByVal ActionType As String) As Boolean
    Dim EfficiencyCode As Variant, Qty As Variant, Sets As Variant, Hours As Variant, Efficiency As Variant, i As Integer
    On Error GoTo ErrorHandler
    UpdateEfficiencyList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        If rstMachineChild.RecordCount <> 0 Then
            cnDatabase.Execute "DELETE FROM MachineChild WHERE Code='" & rstMachineMaster.Fields("Code").Value & "'"
        End If
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 1, .ActiveRow, Qty
            .GetText 2, .ActiveRow, Sets
            .GetText 3, .ActiveRow, Hours
            .GetText 4, .ActiveRow, Efficiency
            cnDatabase.Execute "INSERT INTO MachineChild VALUES ('" & rstMachineMaster.Fields("Code").Value & "'," & Qty & "," & Sets & "," & Hours & "," & Efficiency & ")"
            End With
    End If
    Exit Function
ErrorHandler:
    UpdateEfficiencyList = False
End Function
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then fpSpread1.DeleteRows fpSpread1.ActiveRow, 1: fpSpread1.SetFocus
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        Dim Qty As Variant, Sets As Variant, Hours As Variant, Efficiency As Variant
        With fpSpread1
            If .ActiveCol = 3 Then
                .GetText 1, .ActiveRow, Qty
                .GetText 2, .ActiveRow, Sets
                .GetText 3, .ActiveRow, Hours
                If Qty = "" Or Sets = "" Or Hours = "" Then Exit Sub
                            Efficiency = Round((Qty * Sets) / (((Hours * 60) - (MhRealInput2 * Sets)) / 60))
                .SetText 4, .ActiveRow, Efficiency
                .SetText 5, .ActiveRow, Hours * Efficiency
                .SetFocus
                Sendkeys "{ENTER}"
            End If
        End With
    End If
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
