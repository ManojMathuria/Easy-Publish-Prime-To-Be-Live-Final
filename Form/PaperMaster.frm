VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPaperMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Master"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
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
   ScaleWidth      =   7740
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   5160
      Left            =   15
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Width           =   7695
      _Version        =   65536
      _ExtentX        =   13573
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
      Picture         =   "PaperMaster.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   4935
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   120
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   8705
         _Version        =   393216
         Style           =   1
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
         TabPicture(0)   =   "PaperMaster.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Mh3dLabel1(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Mh3dLabel1(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "DataGrid1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Text1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "PaperMaster.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&Op.Bal."
         TabPicture(2)   =   "PaperMaster.frx":0054
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Mh3dFrame3"
         Tab(2).ControlCount=   1
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
            TabIndex        =   53
            Top             =   4450
            Width           =   4215
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3930
            Left            =   120
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   450
            Width           =   7215
            _ExtentX        =   12726
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
                  ColumnWidth     =   6600.189
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   3370
            Left            =   -74880
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   480
            Width           =   7215
            _Version        =   65536
            _ExtentX        =   12726
            _ExtentY        =   5944
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
            Picture         =   "PaperMaster.frx":0070
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   330
               Left            =   4920
               TabIndex        =   4
               ToolTipText     =   "Width"
               Top             =   420
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":008C
               Caption         =   "PaperMaster.frx":00AC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":0118
               Keys            =   "PaperMaster.frx":0136
               Spin            =   "PaperMaster.frx":0180
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.99
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
               Left            =   4920
               TabIndex        =   5
               ToolTipText     =   "Length"
               Top             =   740
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":01A8
               Caption         =   "PaperMaster.frx":01C8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":0234
               Keys            =   "PaperMaster.frx":0252
               Spin            =   "PaperMaster.frx":029C
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.99
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
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   12
               Top             =   2000
               Width           =   1215
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
               Left            =   4920
               MaxLength       =   40
               TabIndex        =   16
               Top             =   2315
               Width           =   2175
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
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   15
               Top             =   2315
               Width           =   2175
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
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   6
               Top             =   1055
               Width           =   2175
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
               MaxLength       =   80
               TabIndex        =   18
               Top             =   2945
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
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   17
               Top             =   2630
               Width           =   5655
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   25
               Top             =   2945
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
               Picture         =   "PaperMaster.frx":02C4
               Picture         =   "PaperMaster.frx":02E0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   23
               Top             =   2630
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
               Picture         =   "PaperMaster.frx":02FC
               Picture         =   "PaperMaster.frx":0318
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   3600
               TabIndex        =   32
               Top             =   1370
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
               Caption         =   " Wt/Unit (Kg)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":0334
               Picture         =   "PaperMaster.frx":0350
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   3600
               TabIndex        =   33
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
               Caption         =   " Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":036C
               Picture         =   "PaperMaster.frx":0388
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
               Height          =   330
               Left            =   4920
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   100
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   582
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
               Picture         =   "PaperMaster.frx":03A4
               Begin VB.OptionButton Option4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Board"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   250
                  Left            =   1125
                  TabIndex        =   1
                  Top             =   50
                  Width           =   790
               End
               Begin VB.OptionButton Option3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Paper"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   250
                  Left            =   165
                  TabIndex        =   0
                  Top             =   50
                  Width           =   775
               End
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   35
               Top             =   1370
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
               Caption         =   " GSM"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":03C0
               Picture         =   "PaperMaster.frx":03DC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Left            =   120
               TabIndex        =   36
               Top             =   2315
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
               Caption         =   " Paper Make"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":03F8
               Picture         =   "PaperMaster.frx":0414
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   3600
               TabIndex        =   39
               Top             =   2315
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
               Caption         =   " Sub-Make"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":0430
               Picture         =   "PaperMaster.frx":044C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   120
               TabIndex        =   40
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
               Caption         =   " Units/Bundle"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":0468
               Picture         =   "PaperMaster.frx":0484
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   330
               Left            =   1440
               TabIndex        =   8
               Top             =   1370
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":04A0
               Caption         =   "PaperMaster.frx":04C0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":052C
               Keys            =   "PaperMaster.frx":054A
               Spin            =   "PaperMaster.frx":0594
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   330
               Left            =   4920
               TabIndex        =   9
               Top             =   1370
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":05BC
               Caption         =   "PaperMaster.frx":05DC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":0648
               Keys            =   "PaperMaster.frx":0666
               Spin            =   "PaperMaster.frx":06B0
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.000"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.000"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.999
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   330
               Left            =   1440
               TabIndex        =   10
               Top             =   1680
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":06D8
               Caption         =   "PaperMaster.frx":06F8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":0764
               Keys            =   "PaperMaster.frx":0782
               Spin            =   "PaperMaster.frx":07CC
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.99
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   41
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
               Caption         =   " Form"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":07F4
               Picture         =   "PaperMaster.frx":0810
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame5 
               Height          =   330
               Left            =   1440
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   100
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   582
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
               Picture         =   "PaperMaster.frx":082C
               Begin VB.OptionButton Option2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Reel"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   250
                  Left            =   1140
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   50
                  Width           =   765
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Sheet"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   250
                  Left            =   165
                  TabIndex        =   20
                  TabStop         =   0   'False
                  Top             =   50
                  Width           =   775
               End
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
               Height          =   330
               Left            =   1440
               TabIndex        =   3
               ToolTipText     =   "Length"
               Top             =   740
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":0848
               Caption         =   "PaperMaster.frx":0868
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":08D4
               Keys            =   "PaperMaster.frx":08F2
               Spin            =   "PaperMaster.frx":093C
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.99
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   120
               TabIndex        =   43
               Top             =   1055
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
               Caption         =   " UOM"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":0964
               Picture         =   "PaperMaster.frx":0980
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   3600
               TabIndex        =   44
               Top             =   1055
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
               Caption         =   " Sheets/Unit"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":099C
               Picture         =   "PaperMaster.frx":09B8
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
               Height          =   330
               Left            =   4920
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   1055
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":09D4
               Caption         =   "PaperMaster.frx":09F4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":0A60
               Keys            =   "PaperMaster.frx":0A7E
               Spin            =   "PaperMaster.frx":0AC8
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "##0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   120
               TabIndex        =   45
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
               Caption         =   " Quality && Bulk"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":0AF0
               Picture         =   "PaperMaster.frx":0B0C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   3600
               TabIndex        =   46
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
               Caption         =   " Grade"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":0B28
               Picture         =   "PaperMaster.frx":0B44
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   330
               Left            =   2625
               TabIndex        =   13
               ToolTipText     =   "Bulk"
               Top             =   2000
               Width           =   990
               _Version        =   65536
               _ExtentX        =   1746
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":0B60
               Caption         =   "PaperMaster.frx":0B80
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":0BEC
               Keys            =   "PaperMaster.frx":0C0A
               Spin            =   "PaperMaster.frx":0C54
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   1440
               TabIndex        =   2
               ToolTipText     =   "Width"
               Top             =   420
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":0C7C
               Caption         =   "PaperMaster.frx":0C9C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":0D08
               Keys            =   "PaperMaster.frx":0D26
               Spin            =   "PaperMaster.frx":0D70
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.99
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   120
               TabIndex        =   48
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
               Caption         =   " Width in (inch)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":0D98
               Picture         =   "PaperMaster.frx":0DB4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
               Height          =   330
               Left            =   120
               TabIndex        =   49
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
               Caption         =   " Length  (inch)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":0DD0
               Picture         =   "PaperMaster.frx":0DEC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   330
               Left            =   3600
               TabIndex        =   50
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
               Caption         =   " Rate/Kg"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":0E08
               Picture         =   "PaperMaster.frx":0E24
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   330
               Left            =   4920
               TabIndex        =   11
               Top             =   1680
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":0E40
               Caption         =   "PaperMaster.frx":0E60
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":0ECC
               Keys            =   "PaperMaster.frx":0EEA
               Spin            =   "PaperMaster.frx":0F34
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999.99
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Left            =   3600
               TabIndex        =   51
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
               Caption         =   " Width  (cm)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":0F5C
               Picture         =   "PaperMaster.frx":0F78
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   330
               Left            =   3600
               TabIndex        =   52
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
               Caption         =   " Length  (cm)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperMaster.frx":0F94
               Picture         =   "PaperMaster.frx":0FB0
            End
            Begin MSForms.ComboBox Combo1 
               Height          =   330
               Left            =   4920
               TabIndex        =   14
               Top             =   2000
               Width           =   2175
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "3836;582"
               MatchEntry      =   0
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Calibri"
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame3 
            Height          =   4320
            Left            =   -74880
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   480
            Width           =   7215
            _Version        =   65536
            _ExtentX        =   12726
            _ExtentY        =   7620
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
            Picture         =   "PaperMaster.frx":0FCC
            Begin VB.TextBox MhRealInput1 
               Alignment       =   1  'Right Justify
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
               Left            =   5365
               MaxLength       =   13
               TabIndex        =   38
               Text            =   "0"
               Top             =   590
               Visible         =   0   'False
               Width           =   1520
            End
            Begin VB.TextBox MhRealInput2 
               Alignment       =   1  'Right Justify
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
               Left            =   3870
               MaxLength       =   13
               TabIndex        =   24
               Text            =   "0.000"
               Top             =   590
               Visible         =   0   'False
               Width           =   1515
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
               Left            =   430
               MaxLength       =   40
               TabIndex        =   21
               Top             =   590
               Visible         =   0   'False
               Width           =   3450
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   4095
               Left            =   120
               TabIndex        =   19
               Top             =   105
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   7223
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   9164542
               HeadLines       =   1
               RowHeight       =   20
               TabAction       =   2
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
               Caption         =   "Opening Balance"
               ColumnCount     =   3
               BeginProperty Column00 
                  DataField       =   "GodownName"
                  Caption         =   "Godown Name"
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
                  DataField       =   "OpBalOther"
                  Caption         =   "    Op.Bal. (UOM)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.000"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "OpBalTat"
                  Caption         =   "              Tat"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
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
                     ColumnWidth     =   3435.024
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1500.095
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     ColumnWidth     =   1544.882
                  EndProperty
               EndProperty
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   2880
            TabIndex        =   47
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
            Picture         =   "PaperMaster.frx":0FE8
            Picture         =   "PaperMaster.frx":1004
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   1
            Left            =   4800
            TabIndex        =   54
            Top             =   4450
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
            Caption         =   " F12-> Duplicate Paper Master"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "PaperMaster.frx":1020
            Picture         =   "PaperMaster.frx":103C
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
            TabIndex        =   31
            Top             =   4450
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
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
Attribute VB_Name = "FrmPaperMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterCode As String  'Master to Modify
Public SL As Boolean 'Selection List
Public FormType As String 'Paper/Reel
Dim cnPaperMaster As New ADODB.Connection
Dim rstPaperList As New ADODB.Recordset
Dim rstPaperMaster As New ADODB.Recordset
Dim rstPaperChild As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstUOMList As New ADODB.Recordset, rstQualityList As New ADODB.Recordset
Dim rstCheckRef As New ADODB.Recordset
Dim AccountCode As String, UOMCode As String, QualityCode As String
Dim SortOrder, PrevStr As String, UpdateFlag1 As Integer
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstPaperList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstPaperList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstPaperList.EOF Or rstPaperList.BOF) Then
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
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    cnPaperMaster.CursorLocation = adUseClient
    cnPaperMaster.Open cnDatabase.ConnectionString
    If FormType = "" Then FormType = "S" 'UpdateFlag1 = 1: Load FrmDialog: FrmDialog.Show vbModal
    If SL = True Then Me.Caption = "Paper Master " Else Me.Caption = "Paper Master " & IIf(FormType = "S", "(Sheet)", "(Reel)")
    If FormType <> "S" Then Mh3dLabel8.Caption = "Reel Width(in)": Mh3dLabel18.Caption = "Reel cut-off(in)": Mh3dLabel12.Caption = "Reel Width(cm)": Mh3dLabel19.Caption = "Reel cut-off(cm)": Mh3dLabel6.Caption = "Units/Reel"
    rstPaperList.Open "SELECT Name,Code FROM PaperMaster WHERE " & IIf(SL, "1=1", "[Form]='" & FormType & "'") & " ORDER BY Name", cnPaperMaster, adOpenKeyset, adLockOptimistic
    rstUOMList.Open "SELECT Name As Col0,Value1,Code FROM GeneralMaster WHERE Type='15' ORDER BY Name", cnPaperMaster, adOpenKeyset, adLockReadOnly
    rstQualityList.Open "SELECT Name As Col0,Value1,Code FROM GeneralMaster WHERE Type='16' ORDER BY Name", cnPaperMaster, adOpenKeyset, adLockReadOnly
    rstAccountList.Open "SELECT Name As Col0, Code FROM AccountMaster ORDER BY Name", cnPaperMaster, adOpenKeyset, adLockReadOnly
    rstPaperMaster.CursorLocation = adUseClient
    rstPaperList.Filter = adFilterNone
    If rstPaperList.RecordCount Then
        If CheckEmpty(MasterCode, False) Then
            rstPaperList.MoveFirst
        Else
            rstPaperList.MoveFirst
            rstPaperList.Find "[Code]='" & MasterCode & "'"
        End If
    End If
    Set DataGrid1.DataSource = rstPaperList
    BusySystemIndicator False
    SSTab1.Tab = 0
    If Not (rstPaperList.EOF Or rstPaperList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Combo1.AddItem "A", 0
    Combo1.AddItem "B", 1
    Combo1.AddItem "C", 2
    Combo1.AddItem "D", 3
    rstPaperList.ActiveConnection = Nothing
    rstUOMList.ActiveConnection = Nothing
    rstQualityList.ActiveConnection = Nothing
    rstAccountList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(FrmPaperMaster)
End Sub
Private Sub Form_Activate()
    MdiMainMenu.mnuPaperMaster.Enabled = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
       If SSTab1.Tab = 0 Then
            Unload Me
       Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "Text12" Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
            If Not Me.ActiveControl Is Nothing Then
                If Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "Text12" Then KeyCode = 0
            Else    'if Form Unloaded in case of Add
                KeyCode = 0
            End If
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
       If Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "Text12" Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
       KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
       KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF12 Then
        If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then DuplicateRecord
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
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstPaperList.Fields("Code").Value: slName = rstPaperList.Fields("Name").Value: KeyCode = 0: Unload Me: Exit Sub
            Else
                SSTab1.Tab = 1
                SSTab1.SetFocus
            End If
        Else
            If Me.ActiveControl.Name <> "MhRealInput1" Then Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "MhRealInput1" Then KeyCode = 0
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
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstPaperMaster)
    Call CloseRecordset(rstPaperChild)
    Call CloseRecordset(rstUOMList)
    Call CloseRecordset(rstQualityList)
    Call CloseRecordset(rstAccountList)
    Call CloseConnection(cnPaperMaster)
    Call CloseRecordset(rstCheckRef)
    ShowProgressInStatusBar False
    MdiMainMenu.mnuPaperMaster.Enabled = True
    FormType = ""
    End Sub
Private Sub MhRealInput3_Validate(Cancel As Boolean) 'GSM
    If MhRealInput4.Value = 0 Or (Not blnRecordExist) Then MhRealInput4.Value = (((MhRealInput6.Value * MhRealInput7.Value * MhRealInput3.Value) / 20000) / 500) * MhRealInput10.Value
    CalcDependents
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean) 'Wt/Unit
    CalcDependents
End Sub
Private Sub MhRealInput6_Validate(Cancel As Boolean)
    MhRealInput8.Value = MhRealInput6.Value / 2.54
    If MhRealInput4.Value = 0 Or (Not blnRecordExist) Then MhRealInput4.Value = (((MhRealInput6.Value * MhRealInput7.Value * MhRealInput3.Value) / 20000) / 500) * MhRealInput10.Value
End Sub
Private Sub MhRealInput7_Validate(Cancel As Boolean)
    MhRealInput9.Value = MhRealInput7.Value / 2.54
    If MhRealInput4.Value = 0 Or (Not blnRecordExist) Then MhRealInput4.Value = (((MhRealInput6.Value * MhRealInput7.Value * MhRealInput3.Value) / 20000) / 500) * MhRealInput10.Value
End Sub
Private Sub MhRealInput8_Validate(Cancel As Boolean)
    MhRealInput6.Value = MhRealInput8.Value * 2.54
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean)
    MhRealInput7.Value = MhRealInput9.Value * 2.54
End Sub
Private Sub Option1_Click()
    If Option1.Value Then CalcDependents
End Sub
Private Sub Option2_Click()
    If Option2.Value Then CalcDependents
End Sub
Private Sub Text1_Change()
    If rstPaperList.RecordCount = 0 Then Exit Sub
    rstPaperList.MoveFirst
    If Len(Text1.Text) > 0 Then
        rstPaperList.Filter = "[Name] Like '%" & FixQuote(Text1.Text) & "%'"
        If rstPaperList.EOF Then  'if Spelling mistake
            rstPaperList.Filter = adFilterNone
            rstPaperList.MoveFirst
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            Sendkeys "{End}"
        Else    'if Spelling alright
            PrevStr = Text1.Text
        End If
    Else
        rstPaperList.Filter = adFilterNone
        rstPaperList.MoveFirst
        Set DataGrid1.DataSource = rstPaperList
        PrevStr = ""
    End If
    If Not (rstPaperList.EOF Or rstPaperList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstPaperList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstPaperList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstPaperList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstPaperList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstPaperList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstPaperList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstPaperList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstPaperList
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
            If Not (rstPaperList.EOF Or rstPaperList.BOF) Then
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
        If SSTab1.Tab = 1 Then
            Mh3dFrame2.Enabled = True
            Mh3dFrame3.Enabled = False
            If Option3.Value Then Option3.SetFocus Else Option4.SetFocus
        Else
            Mh3dFrame2.Enabled = False
            Mh3dFrame3.Enabled = True
            DataGrid2.SetFocus
        End If
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    If Button.Index = 1 Then
        If rstPaperMaster.State = adStateOpen Then rstPaperMaster.Close
        rstPaperMaster.Open "Select * From PaperMaster Where Code = ''", cnPaperMaster, adOpenKeyset, adLockOptimistic
        ClearFields ("P")
        ClearFields ("C")
        Call LoadOpBalList("")
        If rstPaperChild.State = adStateClosed Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        If AddRecord(rstPaperMaster) Then
            Call SetButtons(False)
            SSTab1.Tab = 1
            If Option3.Value Then Option3.SetFocus Else Option4.SetFocus
            blnRecordExist = False
            cnPaperMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstPaperList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstPaperList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Master")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If CheckRef Then
            DisplayError ("Failed to delete the record")
        ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnPaperMaster.Execute "DELETE FROM PaperMaster WHERE Code = '" & rstPaperList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstPaperList.Delete
                rstPaperList.MoveNext
                If rstPaperList.RecordCount > 0 And rstPaperList.EOF Then rstPaperList.MoveLast
                Call UpdateUserAction("Paper Master", "D", Trim(Text2.Text), cnPaperMaster)
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
        MakeTextBoxInvisible (False)
        If blnRecordExist And AllowMastersModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Master")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstPaperMaster) Then
            If UpdateOpBalList("D") Then
                 UpdateFlag = 1
                 If rstPaperChild.RecordCount <> 0 Then
                      rstPaperChild.MoveFirst
                      Do While Not rstPaperChild.EOF
                          If (Val(rstPaperChild.Fields("OpBalOther").Value) <> 0 Or Val(rstPaperChild.Fields("OpBalTat").Value) <> 0) And rstPaperChild.Fields("Imported").Value = "N" Then
                               If Not UpdateOpBalList("U") Then
                                    UpdateFlag = 0
                                    Exit Do
                               End If
                          End If
                          rstPaperChild.MoveNext
                      Loop
                 End If
            End If
        End If
        If UpdateFlag Then
            Call UpdateUserAction("Paper Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), cnPaperMaster)
            AddToList
            cnPaperMaster.CommitTrans
            If rstPaperMaster.State = adStateOpen Then rstPaperMaster.Close
            rstPaperMaster.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstPaperMaster) Then
            cnPaperMaster.RollbackTrans
            If rstPaperMaster.State = adStateOpen Then rstPaperMaster.Close
            rstPaperMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstPaperList.ActiveConnection = cnPaperMaster
        Do While Not RefreshRecord(rstPaperList): Loop
        rstPaperList.ActiveConnection = Nothing
        Set DataGrid1.DataSource = rstPaperList
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
        If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstPaperList.RecordCount > 0 Then
           rstPaperList.MovePrevious
           If rstPaperList.BOF Then
              rstPaperList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstPaperList.RecordCount > 0 Then
           rstPaperList.MoveNext
           If rstPaperList.EOF Then
              rstPaperList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstPaperList.RecordCount > 0 Then rstPaperList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Call CloseForm(FrmPaperMaster)
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstPaperList.EOF Or rstPaperList.BOF) Then
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
    If rstPaperList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstPaperMaster.EOF Or rstPaperMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnPaperMaster, "PaperMaster", "Code", "Name", Text2.Text, rstPaperMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = Text2.Text
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
        If RTrim(UOMCode) <> "" Then rstUOMList.MoveFirst: rstUOMList.Find "[Code] = '" & UOMCode & "'": MhRealInput10.Value = Val(rstUOMList.Fields("Value1").Value): CalcDependents: Sendkeys "{TAB}" Else Text4.Text = ""
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If CheckEmpty(Text4.Text, False) Then Cancel = True
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text7.Text)
        rstQualityList.MoveFirst
        rstQualityList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": QualityCode = ""
        Call LoadSelectionList(rstQualityList, "List of Qualities...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text7, QualityCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(QualityCode) <> "" Then
            rstQualityList.MoveFirst: rstQualityList.Find "[Code] = '" & QualityCode & "'"
            If MhRealInput11.Value = 0 Then MhRealInput11.Value = Val(rstQualityList.Fields("Value1").Value)
            Sendkeys "{TAB}"
        Else
            Text7.Text = ""
        End If
    End If
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
    If CheckEmpty(Text7.Text, False) Then Cancel = True
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    If CheckEmpty(Text5.Text, True) Then Cancel = True Else CalcDependents
End Sub
Private Sub Text6_Validate(Cancel As Boolean)
    If CheckEmpty(Text6.Text, True) Then Cancel = True Else CalcDependents
End Sub
Private Sub ViewRecord()
    ClearFields ("P")
    ClearFields ("C")
    If rstPaperList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstPaperMaster.State = adStateOpen Then rstPaperMaster.Close
    rstPaperMaster.Open "Select * From PaperMaster Where Code = '" & FixQuote(rstPaperList.Fields("Code").Value) & "'", cnPaperMaster, adOpenKeyset, adLockOptimistic
    If rstPaperMaster.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields(ByVal strType As String)
    If strType = "P" Then
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        MhRealInput3.Value = 0
        MhRealInput4.Value = 0
        MhRealInput5.Value = 0
        MhRealInput6.Value = 0
        MhRealInput7.Value = 0
        MhRealInput8.Value = 0
        MhRealInput9.Value = 0
        MhRealInput10.Value = 0
        MhRealInput11.Value = 0
        MhRealInput12.Value = 0
    If FormType = "S" Then
        Option1.Value = True
        Option2.Value = False
    ElseIf FormType = "R" Then
        Option1.Value = False
        Option2.Value = True
    End If
        Option3.Value = True
        Option4.Value = False
        Combo1.ListIndex = 0
    ElseIf strType = "C" Then
        Text12.Text = ""
        MhRealInput1.Text = "0"
        MhRealInput2.Text = "0.000"
    End If
    Text4.Enabled = True: MhRealInput3.Enabled = True: MhRealInput4.Enabled = True
End Sub
Private Sub LoadFields()
    If rstPaperMaster.EOF Or rstPaperMaster.BOF Then Exit Sub
    If rstPaperMaster.Fields("Form").Value = "S" Then Option1.Value = True Else Option2.Value = True
    If rstPaperMaster.Fields("Type").Value = "P" Then Option3.Value = True Else Option4.Value = True
    UOMCode = rstPaperMaster.Fields("UOM").Value
    rstUOMList.MoveFirst
    rstUOMList.Find "[Code] = '" & UOMCode & "'"
    Text4.Text = rstUOMList.Fields("Col0").Value
    MhRealInput10.Value = Val(rstUOMList.Fields("Value1").Value)
    QualityCode = rstPaperMaster.Fields("Quality").Value
    rstQualityList.MoveFirst
    rstQualityList.Find "[Code] = '" & QualityCode & "'"
    Text7.Text = rstQualityList.Fields("Col0").Value
    MhRealInput3.Value = Val(rstPaperMaster.Fields("GSM").Value)
    Text5.Text = rstPaperMaster.Fields("Make").Value
    Text6.Text = rstPaperMaster.Fields("SubMake").Value
    MhRealInput4.Value = Val(rstPaperMaster.Fields("Weight/Unit").Value)
    MhRealInput5.Value = Val(rstPaperMaster.Fields("Units/Bundle").Value)
    MhRealInput6.Value = Val(rstPaperMaster.Fields("cmWidth").Value)
    MhRealInput7.Value = Val(rstPaperMaster.Fields("cmLength").Value)
    MhRealInput8.Value = Val(rstPaperMaster.Fields("inWidth").Value)
    MhRealInput9.Value = Val(rstPaperMaster.Fields("inLength").Value)
    MhRealInput11.Value = Val(rstPaperMaster.Fields("Bulk").Value)
    MhRealInput12.Value = Val(rstPaperMaster.Fields("Rate/Kg").Value)
    Combo1.ListIndex = Asc(rstPaperMaster.Fields("Grade").Value) - 65
    Text2.Text = rstPaperMaster.Fields("Name").Value
    Text3.Text = rstPaperMaster.Fields("PrintName").Value
    Call LoadOpBalList(rstPaperMaster.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstPaperMaster.RecordCount = 0 Then Exit Sub
    If rstPaperChild.State = adStateClosed Then SSTab1.Tab = 0: Exit Sub
    If rstPaperMaster.State = adStateOpen Then rstPaperMaster.Close
    rstPaperMaster.CursorLocation = adUseServer
    rstPaperMaster.Open "Select * From PaperMaster Where Code = '" & FixQuote(rstPaperList.Fields("Code").Value) & "'", cnPaperMaster, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstPaperMaster.Fields("Printstatus") = "N"
    If chkRef("SELECT Code As Paper FROM PaperChild WHERE Code='" & rstPaperList.Fields("Code").Value & "' UNION SELECT Paper1 As Paper FROM BookPOChild05 WHERE Paper1='" & rstPaperList.Fields("Code").Value & "' UNION SELECT Paper2 As Paper FROM BookPOChild05 WHERE Paper2='" & rstPaperList.Fields("Code").Value & "' UNION SELECT Paper4 As Paper FROM BookPOChild05 WHERE Paper4='" & rstPaperList.Fields("Code").Value & "' UNION SELECT Paper FROM BookPOChild06 WHERE Paper='" & rstPaperList.Fields("Code").Value & "' UNION SELECT Paper FROM BookPOChild09 WHERE Paper='" & rstPaperList.Fields("Code").Value & "' UNION SELECT Paper FROM PaperPOChild WHERE Paper='" & rstPaperList.Fields("Code").Value & "' UNION SELECT Paper FROM PaperMVChild WHERE Paper='" & rstPaperList.Fields("Code").Value & "' UNION SELECT Paper FROM PaperDNChild WHERE Paper='" & rstPaperList.Fields("Code").Value & "' UNION SELECT Item FROM MaterialSVChild WHERE Category='2' AND Item='" & rstPaperList.Fields("Code").Value & "'") Then
        Text4.Enabled = False: MhRealInput3.Enabled = False: MhRealInput4.Enabled = False
    End If
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    If Option3.Value Then Option3.SetFocus Else Option4.SetFocus
    blnRecordExist = True
    cnPaperMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstPaperMaster.EOF Or rstPaperMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstPaperMaster.Fields("Code").Value = GenerateCode(cnPaperMaster, "SELECT MAX(Code) FROM PaperMaster", 6, "0")
        rstPaperMaster.Fields("CreatedBy").Value = UserCode
        rstPaperMaster.Fields("CreatedOn").Value = Now()
        rstPaperMaster.Fields("Recordstatus").Value = "N"
    Else
        rstPaperMaster.Fields("ModifiedBy").Value = UserCode
        rstPaperMaster.Fields("ModifiedOn").Value = Now()
        rstPaperMaster.Fields("Recordstatus").Value = "M"
    End If
    rstPaperMaster.Fields("Name").Value = Trim(Text2.Text)
    rstPaperMaster.Fields("PrintName").Value = Trim(Text3.Text)
    If Option1.Value Then rstPaperMaster.Fields("Form").Value = "S" Else rstPaperMaster.Fields("Form").Value = "R"
    If Option3.Value Then rstPaperMaster.Fields("Type").Value = "P" Else rstPaperMaster.Fields("Type").Value = "B"
    rstPaperMaster.Fields("cmWidth").Value = MhRealInput6.Value
    rstPaperMaster.Fields("cmLength").Value = MhRealInput7.Value
    rstPaperMaster.Fields("inWidth").Value = MhRealInput8.Value
    rstPaperMaster.Fields("inLength").Value = MhRealInput9.Value
    rstPaperMaster.Fields("UOM").Value = UOMCode
    rstPaperMaster.Fields("Quality").Value = QualityCode
    rstPaperMaster.Fields("GSM").Value = MhRealInput3.Value
    rstPaperMaster.Fields("Make").Value = Trim(Text5.Text)
    rstPaperMaster.Fields("SubMake").Value = Trim(Text6.Text)
    rstPaperMaster.Fields("Weight/Unit").Value = MhRealInput4.Value
    rstPaperMaster.Fields("Units/Bundle").Value = MhRealInput5.Value
    rstPaperMaster.Fields("Rate/Kg").Value = MhRealInput12.Value
    rstPaperMaster.Fields("Quality").Value = QualityCode
    rstPaperMaster.Fields("Grade").Value = Combo1.Text
    rstPaperMaster.Fields("Bulk").Value = MhRealInput11.Value
    rstPaperMaster.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstPaperList.MoveFirst
    rstPaperList.Find "[Code] = '" & rstPaperMaster.Fields("Code").Value & "'"
    If rstPaperList.EOF Then
       rstPaperList.AddNew
       rstPaperList.Fields("Code").Value = rstPaperMaster.Fields("Code").Value
    End If
    rstPaperList.Fields("Name").Value = rstPaperMaster.Fields("Name").Value
    rstPaperList.Update
    rstPaperList.Sort = "Name Asc"
    rstPaperList.Find "[Code] = '" & rstPaperMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If Option1.Value Then
        If MhRealInput6.Value = 0 Then
            SSTab1.Tab = 1: MhRealInput6.SetFocus: CheckMandatoryFields = True
        ElseIf MhRealInput7.Value = 0 Then
            SSTab1.Tab = 1: MhRealInput7.SetFocus: CheckMandatoryFields = True
        ElseIf MhRealInput8.Value = 0 Then
            SSTab1.Tab = 1: MhRealInput8.SetFocus: CheckMandatoryFields = True
        ElseIf MhRealInput9.Value = 0 Then
            SSTab1.Tab = 1: MhRealInput9.SetFocus: CheckMandatoryFields = True
        End If
    Else
        If MhRealInput6.Value = 0 Then
            SSTab1.Tab = 1: MhRealInput6.SetFocus: CheckMandatoryFields = True
        ElseIf MhRealInput8.Value = 0 Then
            SSTab1.Tab = 1: MhRealInput8.SetFocus: CheckMandatoryFields = True
        End If
    End If
    If CheckDuplicate(cnPaperMaster, "PaperMaster", "Code", "Name", Text2.Text, rstPaperMaster.Fields("Code").Value, False) Then
        SSTab1.Tab = 1: Text2.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then   'Print Name
        SSTab1.Tab = 1: Text3.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text4.Text, False) Then   'UOM
        SSTab1.Tab = 1: Text4.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text7.Text, False) Then   'Quality
        SSTab1.Tab = 1: Text7.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Combo1.Text, False) Then  'Grade
        SSTab1.Tab = 1: Combo1.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text5.Text, False) Then   'Make
        SSTab1.Tab = 1: Text5.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text6.Text, False) Then   'Sub-Make
        SSTab1.Tab = 1: Text6.SetFocus: CheckMandatoryFields = True
    End If
End Function
Private Sub LoadOpBalList(ByVal strPaperCode As String)
    On Error GoTo ErrorHandler
    If rstPaperChild.State = adStateOpen Then rstPaperChild.Close
    rstPaperChild.Open "Select P.Account, A.Name As GodownName, P.OpBalOther, P.OpBalTat, P.Imported From PaperChild P, AccountMaster A Where P.Account = A.Code And P.Code = '" & FixQuote(strPaperCode) & "' Order By A.Name", cnPaperMaster, adOpenKeyset, adLockOptimistic
    rstPaperChild.ActiveConnection = Nothing
    Set DataGrid2.DataSource = rstPaperChild
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Opening Balance")
End Sub
Private Function UpdateOpBalList(ByVal strOption As String) As Boolean
    Dim Sheets As Long
    On Error GoTo ErrorHandler
    UpdateOpBalList = True
    If strOption = "D" Then
        cnPaperMaster.Execute "DELETE FROM PaperChild WHERE Code ='" & rstPaperMaster.Fields("Code").Value & "' And Imported = 'N'"
    Else
        Sheets = (Fix(Val(rstPaperChild.Fields("OpBalOther").Value)) * MhRealInput10.Value) + ((Val(rstPaperChild.Fields("OpBalOther").Value) - Fix(Val(rstPaperChild.Fields("OpBalOther").Value))) * 1000)
        cnPaperMaster.Execute "INSERT INTO PaperChild VALUES ('" & rstPaperMaster.Fields("Code").Value & "','" & rstPaperChild.Fields("Account").Value & "'," & rstPaperChild.Fields("OpBalOther").Value & "," & Sheets & "," & rstPaperChild.Fields("OpBalTat").Value & ",'N')"
    End If
    Exit Function
ErrorHandler:
    UpdateOpBalList = False
End Function
Private Function CheckRef() As Boolean
    On Error GoTo ErrorHandler
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "Select Paper1 From BookPOChild05 Where Paper1 = '" & rstPaperList.Fields("Code").Value & "'", cnPaperMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True: Exit Function
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "Select Paper2 From BookPOChild05 Where Paper2 = '" & rstPaperList.Fields("Code").Value & "'", cnPaperMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True: Exit Function
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "Select Paper4 From BookPOChild05 Where Paper4 = '" & rstPaperList.Fields("Code").Value & "'", cnPaperMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True: Exit Function
    If rstCheckRef.State = adStateOpen Then rstCheckRef.Close
    rstCheckRef.Open "Select Paper From BookPOChild06 Where Paper = '" & rstPaperList.Fields("Code").Value & "'", cnPaperMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then CheckRef = True: Exit Function
    Exit Function
ErrorHandler:
    CheckRef = True
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Private Sub DataGrid2_DblClick()
    Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
End Sub
Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyE Then
        If rstPaperChild.RecordCount = 0 Or rstPaperChild.Fields("Imported").Value = "Y" Then Exit Sub
        If Val(CheckNull(rstPaperChild.Fields("OpBalOther").Value)) <> 0 Or Val(CheckNull(rstPaperChild.Fields("OpBalTat").Value)) <> 0 Then
            AccountCode = rstPaperChild.Fields("Account").Value
            Text12.Text = rstPaperChild.Fields("GodownName").Value
            MhRealInput2.Text = Format(Val(rstPaperChild.Fields("OpBalOther").Value), "0.000")
            MhRealInput1.Text = Format(Val(rstPaperChild.Fields("OpBalTat").Value), "0")
        End If
        With DataGrid2
            Text12.Visible = True
            Text12.Move .Left + .Columns(0).Left, .Top + .RowTop(.Row), .Columns(0).Width + 10, .RowHeight + 30
            MhRealInput2.Visible = True
            MhRealInput2.Move .Left + .Columns(1).Left, .Top + .RowTop(.Row), .Columns(1).Width + 10, .RowHeight + 30
            MhRealInput1.Visible = True
            MhRealInput1.Move .Left + .Columns(2).Left, .Top + .RowTop(.Row), .Columns(2).Width + 10, .RowHeight + 30
        End With
        DataGrid2.Enabled = False
        Text12.SetFocus
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Sendkeys "^"
        Call AddRecord(rstPaperChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If rstPaperChild.RecordCount = 0 Or rstPaperChild.Fields("Imported").Value = "Y" Then Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            Set DataGrid2.DataSource = Nothing
            rstPaperChild.Delete
            rstPaperChild.MoveNext
            Set DataGrid2.DataSource = rstPaperChild
            DataGrid2.SetFocus
        End If
        If rstPaperChild.RecordCount = 0 Then
            Call ClearFields("C")
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
    End If
End Sub
Private Sub DataGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim menusel As String
    If Button = vbRightButton Then
       menusel = DisplayPopupMenu(Me.hwnd)
        Select Case menusel
            Case 1
                Call DataGrid2_KeyDown(vbKeyA, vbCtrlMask)
            Case 2
                Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
            Case 3
                Call DataGrid2_KeyDown(vbKeyD, vbCtrlMask)
            Case Else
        End Select
    End If
End Sub
Private Sub Text12_Change()
    If Text12.Text = " " Then
        Text12.Text = "?"
        Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub Text12_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text12.Text)
    If rstAccountList.RecordCount = 0 Then
        DisplayError ("No Record in Godown Master")
        Cancel = True
        Exit Sub
    Else
        rstAccountList.MoveFirst
    End If
    rstAccountList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstAccountList.EOF Then
        SelectionType = "S"
        AccountCode = ""
        Call LoadSelectionList(rstAccountList, "List of Godowns...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text12, AccountCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text12.Text, False) Then
            Text12.Text = "?"
        End If
        If RTrim(AccountCode) <> "" Then
            Sendkeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf (rstPaperChild.Fields("GodownName").Value <> Text12.Text) Or (CheckEmpty(rstPaperChild.Fields("GodownName").Value, False)) Then
        If CheckDuplicateGodown Then
            Call DisplayError("Duplicate Entry")
            Text12.SelStart = 0
            Text12.SelLength = Len(Text12.Text)
            Cancel = True
            Exit Sub
        End If
    End If
    AccountCode = rstAccountList.Fields("Code").Value
End Sub
Private Sub MhRealInput2_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput2_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput2, KeyAscii, 3
End Sub
Private Sub MhRealInput2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    Dim RPB As Double
    If Not ValidateNumber(Me.ActiveControl, 3) Then
        Cancel = True
    Else
        If Val(CheckNull(rstPaperChild.Fields("OpBalTat").Value)) = 0 Then
            If MhRealInput5.Value <> 0 Then
                RPB = MhRealInput5.Value
                If Val(MhRealInput2.Text) * 1000 Mod RPB * 1000 > 0 Then
                    MhRealInput1.Text = Format(Int(Val(MhRealInput2.Text) / RPB) + 1, "0")
                Else
                    MhRealInput1.Text = Format(Int(Val(MhRealInput2.Text) / RPB), "0")
                End If
            End If
        End If
    End If
End Sub
Private Sub MhRealInput1_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput1_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput2, KeyAscii, 0
End Sub
Private Sub MhRealInput1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Not ValidateNumber(Me.ActiveControl, 0) Then Exit Sub
        rstPaperChild.Fields("Account").Value = AccountCode
        rstPaperChild.Fields("GodownName").Value = Trim(Text12.Text)
        rstPaperChild.Fields("OpBalOther").Value = Format(Val(MhRealInput2.Text), "0.000")
        rstPaperChild.Fields("OpBalTat").Value = Format(Val(MhRealInput1.Text), "0")
        rstPaperChild.Fields("Imported").Value = "N"
        rstPaperChild.Update
        MakeTextBoxInvisible (False)
        If rstPaperChild.AbsolutePosition = rstPaperChild.RecordCount Then
            Call DataGrid2_KeyDown(vbKeyA, vbCtrlMask)
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
       MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    Cancel = True
End Sub
Private Sub MakeTextBoxInvisible(ByVal KeyEscPressed As Boolean)
    If KeyEscPressed Then
        If Not (rstPaperChild.EOF Or rstPaperChild.BOF) Then
            If Val(CheckNull(rstPaperChild.Fields("OpBalOther").Value)) = 0 And Val(CheckNull(rstPaperChild.Fields("OpBalTat").Value)) = 0 Then
                rstPaperChild.Delete
                rstPaperChild.MoveNext
                If rstPaperChild.RecordCount > 0 Then rstPaperChild.MoveFirst
            End If
        End If
    End If
    Text12.Visible = False
    MhRealInput2.Visible = False
    MhRealInput1.Visible = False
    DataGrid2.Enabled = True
    If Mh3dFrame3.Enabled Then
        DataGrid2.SetFocus
    End If
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then
        rstPaperList.Filter = "[Name] Like '%" & SrchText & "%'"
    End If
End Sub
Private Function CheckDuplicateGodown() As Boolean
    Dim dblBookMark As Double
    If rstPaperChild.RecordCount = 0 Then Exit Function
    If Not (rstPaperChild.EOF Or rstPaperChild.BOF) Then dblBookMark = rstPaperChild.Bookmark
    rstPaperChild.MoveFirst
    Do While Not rstPaperChild.EOF
        If rstPaperChild.Fields("GodownName").Value = Trim(Text12.Text) Then CheckDuplicateGodown = True: Exit Do
        rstPaperChild.MoveNext
    Loop
    If dblBookMark <> 0 Then rstPaperChild.Bookmark = dblBookMark Else rstPaperChild.MoveLast
End Function
Private Function CalcDependents(Optional ByVal BeforeUpdate As Boolean) As Boolean
    Dim Value1 As Double
    If MhRealInput4.Value > 0 Then 'Wt/Unit
        Value1 = IIf(Option1.Value, 50, 400) / MhRealInput4.Value
        If MhRealInput5.Value = 0 Or (Not blnRecordExist) Then MhRealInput5.Value = IIf(Value1 = Int(Value1), Int(Value1), Int(Value1) + 1)
        If MhRealInput6.Value > 0 And MhRealInput7.Value > 0 And MhRealInput10.Value > 0 Then
            If MhRealInput3.Value = 0 Or (Not blnRecordExist) Then MhRealInput3.Value = (MhRealInput4.Value * 20000 * 500) / (MhRealInput6.Value * MhRealInput7.Value * MhRealInput10.Value)
        End If
    End If
    Text2.Text = Trim(Text5.Text) + "-" + Trim(MhRealInput3.Value) + "gsm-"
    If Option1.Value Then Text2.Text = Text2.Text + Trim(MhRealInput8.Text) + "X" + Trim(MhRealInput9.Text) + "in²" + "-(" + Trim(MhRealInput6.Text) + "X" + Trim(MhRealInput7.Text) + "cm²)" + "-" + Trim(MhRealInput4.Text) + "kg" + "-" + Trim(Text6.Text) Else Text2.Text = Text2.Text + Trim(MhRealInput8.Text) + "in-(" + Trim(MhRealInput6.Text) + "cm)-Reel-" + Trim(Text6.Text)
End Function
Private Sub DuplicateRecord()
    Dim TmpTbl As String
    TmpTbl = "T" & GetFileNameFromPath(GetTemporaryFileName()): TmpTbl = Left(TmpTbl, InStr(1, TmpTbl, ".", vbTextCompare) - 1)
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    Dim PaperCode As String, PaperName As String
    PaperCode = GenerateCode(cnPaperMaster, "SELECT MAX(Code) FROM PaperMaster", 6, "0")
    PaperName = Trim(Left(rstPaperList.Fields("Name").Value, 76)) + " (D)"
    cnPaperMaster.BeginTrans
    cnPaperMaster.Execute "SELECT * INTO [" & TmpTbl & "] FROM PaperMaster Where Code = '" & rstPaperList.Fields("Code").Value & "'"
    cnPaperMaster.Execute "UPDATE  [" & TmpTbl & "] SET Code='" & PaperCode & "',Name='" & PaperName & "',PrintName='" & PaperName & "'"
    cnPaperMaster.Execute "INSERT INTO PaperMaster SELECT * FROM " & TmpTbl
    cnPaperMaster.Execute "DROP TABLE " & TmpTbl
    cnPaperMaster.CommitTrans
    Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    Text1.Text = Trim(PaperName): Sendkeys "{END}"
    MdiMainMenu.MousePointer = vbNormal
    Call MsgBox("Successfully Duplicated the Record !", vbInformation, App.Title)
    Exit Sub
ErrorHandler:
    MdiMainMenu.MousePointer = vbNormal
    DisplayError ("Failed to Duplicate the Record")
    cnPaperMaster.RollbackTrans
End Sub
