VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPackingSlip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing Slip"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14505
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   8865
   ScaleWidth      =   14505
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   8850
      Left            =   15
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   14475
      _Version        =   65536
      _ExtentX        =   25532
      _ExtentY        =   15610
      _StockProps     =   77
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
      Picture         =   "PackingSlip.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8625
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   120
         Width           =   14250
         _ExtentX        =   25135
         _ExtentY        =   15214
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
         TabPicture(0)   =   "PackingSlip.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Combo7"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Mh3dLabel18"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Mh3dLabel4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "DataGrid1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Text1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdChange"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "PackingSlip.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).ControlCount=   1
         Begin VB.CommandButton cmdChange 
            Height          =   375
            Left            =   13770
            Picture         =   "PackingSlip.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Exit"
            Top             =   8130
            Width           =   375
         End
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
            Left            =   1200
            TabIndex        =   22
            Top             =   8160
            Width           =   9465
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7635
            Left            =   120
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   450
            Width           =   14025
            _ExtentX        =   24739
            _ExtentY        =   13467
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   9164542
            HeadLines       =   1
            RowHeight       =   18
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
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7995
            Left            =   -74880
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   480
            Width           =   13995
            _Version        =   65536
            _ExtentX        =   24686
            _ExtentY        =   14102
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
            Picture         =   "PackingSlip.frx":0612
            Begin VB.TextBox Text20 
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
               Left            =   9075
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   420
               Width           =   4815
            End
            Begin VB.TextBox Text19 
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
               Left            =   11235
               MaxLength       =   80
               TabIndex        =   9
               Top             =   2530
               Width           =   2655
            End
            Begin VB.TextBox Text18 
               Alignment       =   2  'Center
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
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   420
               Width           =   5610
            End
            Begin VB.TextBox Text17 
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
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   2220
               Width           =   7290
            End
            Begin VB.TextBox Text15 
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
               Left            =   10755
               MaxLength       =   80
               TabIndex        =   10
               Top             =   2850
               Width           =   3135
            End
            Begin VB.TextBox Text16 
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
               Left            =   1800
               MaxLength       =   80
               TabIndex        =   12
               Top             =   3475
               Width           =   7290
            End
            Begin VB.TextBox Text14 
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
               Left            =   10755
               MaxLength       =   40
               TabIndex        =   7
               Top             =   2220
               Width           =   3135
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
               Left            =   10755
               MaxLength       =   40
               TabIndex        =   11
               Top             =   3160
               Width           =   3135
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
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   1900
               Width           =   7290
            End
            Begin VB.TextBox Text11 
               Alignment       =   2  'Center
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
               Left            =   10755
               MaxLength       =   20
               TabIndex        =   5
               ToolTipText     =   "GR No."
               Top             =   1590
               Width           =   1575
            End
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
               Left            =   10755
               MaxLength       =   40
               TabIndex        =   3
               Top             =   960
               Width           =   3135
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
               Left            =   1800
               MaxLength       =   20
               TabIndex        =   1
               Top             =   2850
               Width           =   7290
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
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
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   105
               Width           =   2130
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   3765
               Left            =   120
               TabIndex        =   14
               Top             =   4005
               Width           =   13770
               _Version        =   524288
               _ExtentX        =   24289
               _ExtentY        =   6641
               _StockProps     =   64
               EditEnterAction =   2
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
               GridColor       =   33023
               MaxCols         =   7
               MaxRows         =   2000
               RowHeaderDisplay=   0
               SpreadDesigner  =   "PackingSlip.frx":062E
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
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   2
               TabStop         =   0   'False
               Top             =   3160
               Width           =   7290
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
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   960
               Width           =   7290
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   24
               Top             =   105
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Invoice No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":0E82
               Picture         =   "PackingSlip.frx":0E9E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   3915
               TabIndex        =   25
               Top             =   105
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Invoice Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":0EBA
               Picture         =   "PackingSlip.frx":0ED6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   26
               Top             =   960
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Party Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":0EF2
               Picture         =   "PackingSlip.frx":0F0E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   27
               Top             =   3165
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Narration"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":0F2A
               Picture         =   "PackingSlip.frx":0F46
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   5595
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   105
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               Calendar        =   "PackingSlip.frx":0F62
               Caption         =   "PackingSlip.frx":107A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PackingSlip.frx":10E6
               Keys            =   "PackingSlip.frx":1104
               Spin            =   "PackingSlip.frx":1162
               AlignHorizontal =   2
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mm-yyyy"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   1
               ForeColor       =   -2147483640
               Format          =   "dd-mm-yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "  -  -    "
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   39849
               CenturyMode     =   0
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
               Left            =   7290
               MaxLength       =   20
               TabIndex        =   0
               Top             =   2530
               Width           =   1800
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
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   1590
               Width           =   7290
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
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   1280
               Width           =   7290
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
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   2530
               Width           =   3825
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   1275
               Left            =   120
               TabIndex        =   33
               Top             =   1275
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Picture         =   "PackingSlip.frx":118A
               VAlignment      =   0
               Picture         =   "PackingSlip.frx":11A6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   0
               Left            =   9075
               TabIndex        =   34
               Top             =   960
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Handed Over To"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":11C2
               Picture         =   "PackingSlip.frx":11DE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   9075
               TabIndex        =   35
               Top             =   2535
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Paid GR"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":11FA
               Picture         =   "PackingSlip.frx":1216
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   9075
               TabIndex        =   36
               Top             =   1590
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " GR No. && Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":1232
               Picture         =   "PackingSlip.frx":124E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   9075
               TabIndex        =   37
               Top             =   1905
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Date && Time"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":126A
               Picture         =   "PackingSlip.frx":1286
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Left            =   120
               TabIndex        =   38
               Top             =   2850
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Transport"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":12A2
               Picture         =   "PackingSlip.frx":12BE
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput4 
               Height          =   330
               Left            =   10755
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   1905
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "PackingSlip.frx":12DA
               Caption         =   "PackingSlip.frx":13F2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PackingSlip.frx":145E
               Keys            =   "PackingSlip.frx":147C
               Spin            =   "PackingSlip.frx":14DA
               AlignHorizontal =   2
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mm-yyyy"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   1
               ForeColor       =   -2147483640
               Format          =   "dd-mm-yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "  -  -    "
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   39849
               CenturyMode     =   0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Left            =   120
               TabIndex        =   39
               Top             =   2535
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Mat.Centre"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":1502
               Picture         =   "PackingSlip.frx":151E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   5610
               TabIndex        =   40
               Top             =   2535
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Station"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":153A
               Picture         =   "PackingSlip.frx":1556
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   7395
               TabIndex        =   41
               Top             =   105
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Vocher Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":1572
               Picture         =   "PackingSlip.frx":158E
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput2 
               Height          =   330
               Left            =   9075
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "PackingSlip.frx":15AA
               Caption         =   "PackingSlip.frx":16C2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PackingSlip.frx":172E
               Keys            =   "PackingSlip.frx":174C
               Spin            =   "PackingSlip.frx":17AA
               AlignHorizontal =   2
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mm-yyyy"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   1
               ForeColor       =   -2147483640
               Format          =   "dd-mm-yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "  -  -    "
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   39849
               CenturyMode     =   0
            End
            Begin TDBNumber6Ctl.TDBNumber MhTimeInput1 
               Height          =   330
               Left            =   12315
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "PackingSlip.frx":17D2
               Caption         =   "PackingSlip.frx":17F2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PackingSlip.frx":185E
               Keys            =   "PackingSlip.frx":187C
               Spin            =   "PackingSlip.frx":18C6
               AlignHorizontal =   2
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "00.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "00.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   23.59
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   190251013
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhTimeInput2 
               Height          =   330
               Left            =   12315
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   1275
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "PackingSlip.frx":18EE
               Caption         =   "PackingSlip.frx":190E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PackingSlip.frx":197A
               Keys            =   "PackingSlip.frx":1998
               Spin            =   "PackingSlip.frx":19E2
               AlignHorizontal =   2
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "00.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "00.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   23.59
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   183828485
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhTimeInput3 
               Height          =   330
               Left            =   12315
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   1905
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "PackingSlip.frx":1A0A
               Caption         =   "PackingSlip.frx":1A2A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PackingSlip.frx":1A96
               Keys            =   "PackingSlip.frx":1AB4
               Spin            =   "PackingSlip.frx":1AFE
               AlignHorizontal =   2
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "00.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "00.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   23.59
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   -1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   36765701
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput3 
               Height          =   330
               Left            =   11235
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   1275
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "PackingSlip.frx":1B26
               Caption         =   "PackingSlip.frx":1C3E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PackingSlip.frx":1CAA
               Keys            =   "PackingSlip.frx":1CC8
               Spin            =   "PackingSlip.frx":1D26
               AlignHorizontal =   0
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mm-yyyy"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   1
               ForeColor       =   -2147483640
               Format          =   "dd-mm-yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "  -  -    "
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   39849
               CenturyMode     =   0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   9075
               TabIndex        =   47
               Top             =   3165
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Packed By"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":1D4E
               Picture         =   "PackingSlip.frx":1D6A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   9075
               TabIndex        =   48
               Top             =   2220
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " GR Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":1D86
               Picture         =   "PackingSlip.frx":1DA2
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
               Height          =   330
               Left            =   10755
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   1275
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               Picture         =   "PackingSlip.frx":1DBE
               Begin VB.CheckBox Check1 
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
                  Height          =   270
                  Left            =   120
                  TabIndex        =   4
                  Top             =   45
                  Width           =   345
               End
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   330
               Left            =   9075
               TabIndex        =   50
               Top             =   1275
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Bundles Picked"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":1DDA
               Picture         =   "PackingSlip.frx":1DF6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   330
               Left            =   120
               TabIndex        =   55
               Top             =   3480
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":1E12
               Picture         =   "PackingSlip.frx":1E2E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
               Height          =   330
               Left            =   9075
               TabIndex        =   56
               Top             =   3480
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Bundles"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":1E4A
               Picture         =   "PackingSlip.frx":1E66
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
               Height          =   330
               Left            =   10755
               TabIndex        =   13
               Top             =   3480
               Width           =   3135
               _Version        =   65536
               _ExtentX        =   5530
               _ExtentY        =   582
               Calculator      =   "PackingSlip.frx":1E82
               Caption         =   "PackingSlip.frx":1EA2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PackingSlip.frx":1F0E
               Keys            =   "PackingSlip.frx":1F2C
               Spin            =   "PackingSlip.frx":1F76
               AlignHorizontal =   1
               AlignVertical   =   0
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
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   84738053
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel lblBilledQty 
               Height          =   240
               Left            =   9120
               TabIndex        =   57
               Top             =   7500
               Width           =   2190
               _Version        =   65536
               _ExtentX        =   3863
               _ExtentY        =   423
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
               Caption         =   ""
               Alignment       =   1
               BorderStyle     =   0
               Picture         =   "PackingSlip.frx":1F9E
               Picture         =   "PackingSlip.frx":1FBA
            End
            Begin Mh3dlblLib.Mh3dLabel lblPkdQty 
               Height          =   240
               Left            =   11535
               TabIndex        =   58
               Top             =   7500
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   423
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
               Caption         =   ""
               Alignment       =   1
               BorderStyle     =   0
               Picture         =   "PackingSlip.frx":1FD6
               Picture         =   "PackingSlip.frx":1FF2
            End
            Begin Mh3dlblLib.Mh3dLabel lblVchCreator 
               Height          =   240
               Left            =   120
               TabIndex        =   59
               Top             =   7500
               Width           =   8895
               _Version        =   65536
               _ExtentX        =   15690
               _ExtentY        =   423
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
               Caption         =   ""
               Alignment       =   0
               BorderStyle     =   0
               Picture         =   "PackingSlip.frx":200E
               Picture         =   "PackingSlip.frx":202A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
               Height          =   330
               Left            =   9075
               TabIndex        =   60
               Top             =   2850
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Booking Route"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":2046
               Picture         =   "PackingSlip.frx":2062
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
               Height          =   330
               Left            =   120
               TabIndex        =   62
               Top             =   420
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Clubbed Invoice"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":207E
               Picture         =   "PackingSlip.frx":209A
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame3 
               Height          =   330
               Left            =   10755
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   2535
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               Picture         =   "PackingSlip.frx":20B6
               Begin VB.CheckBox Check2 
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
                  Height          =   270
                  Left            =   120
                  TabIndex        =   8
                  Top             =   40
                  Width           =   345
               End
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   1
               Left            =   7395
               TabIndex        =   65
               Top             =   420
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Clubbed GR"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":20D2
               Picture         =   "PackingSlip.frx":20EE
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput5 
               Height          =   330
               Left            =   12315
               TabIndex        =   6
               Top             =   1590
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "PackingSlip.frx":210A
               Caption         =   "PackingSlip.frx":2222
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PackingSlip.frx":228E
               Keys            =   "PackingSlip.frx":22AC
               Spin            =   "PackingSlip.frx":230A
               AlignHorizontal =   2
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mm-yyyy"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   1
               ForeColor       =   -2147483640
               Format          =   "dd-mm-yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   " "
               ReadOnly        =   0
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "  -  -    "
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   39849
               CenturyMode     =   0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
               Height          =   330
               Left            =   10635
               TabIndex        =   67
               Top             =   105
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
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
               Caption         =   " Vocher  Time"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PackingSlip.frx":2332
               Picture         =   "PackingSlip.frx":234E
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   14000
               Y1              =   850
               Y2              =   850
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   14000
               Y1              =   3900
               Y2              =   3900
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
            Height          =   330
            Left            =   120
            TabIndex        =   29
            Top             =   8160
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
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
            Caption         =   " Find"
            Alignment       =   0
            FillColor       =   9164542
            TextColor       =   0
            Picture         =   "PackingSlip.frx":236A
            Picture         =   "PackingSlip.frx":2386
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
            Height          =   330
            Left            =   10650
            TabIndex        =   53
            Top             =   8160
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
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
            Caption         =   " Vch Series"
            Alignment       =   0
            FillColor       =   9164542
            TextColor       =   0
            Picture         =   "PackingSlip.frx":23A2
            Picture         =   "PackingSlip.frx":23BE
         End
         Begin MSForms.ComboBox Combo7 
            Height          =   330
            Left            =   11730
            TabIndex        =   52
            Top             =   8160
            Width           =   2040
            VariousPropertyBits=   545282075
            BackColor       =   16777215
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3598;582"
            MatchEntry      =   0
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   14505
      _ExtentX        =   25585
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
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Preview"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
Attribute VB_Name = "FrmPackingSlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchType As String
Dim cnPackingSlip As New ADODB.Connection
'
Dim rstPackingSlipList As New ADODB.Recordset, rstPackingSlipChild As New ADODB.Recordset
Dim rstBillInfo As New ADODB.Recordset, rstVchSeriesList As New ADODB.Recordset
Dim rstPackerList As New ADODB.Recordset, rstDelivererList As New ADODB.Recordset, rstTransporterList As New ADODB.Recordset, rstBookingRouteList As New ADODB.Recordset
'
Dim PackerCode As String, DelivererCode As String, TransporterCode As String, BookingRouteCode As String
Dim VchSeries As Variant, Bundles As Integer, SaveAndContinue As Boolean, VchCreator As String
Dim ClubbedVch As String, ClubbedGRVch As String, OldClubbedGRVch As String 'OldClubbedGRVch-For bilty clearance of all bills if the bilty of one bill is cleared
Dim LRange As Long, URange As Long
Dim EditMode As Boolean, SortOrder As String, PrevStr As String, dblBookMark As Double
'
Private Const E_POINTER As Long = &H80004003
Private Const S_OK As Long = 0
Private Const URL_ESCAPE_PERCENT As Long = &H1000&
Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeA" (ByVal pszURL As String, ByVal pszEscaped As String, ByRef pcchEscaped As Long, ByVal dwFlags As Long) As Long
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    cnPackingSlip.CursorLocation = adUseClient
    If cnPackingSlip.State = adStateOpen Then cnPackingSlip.Close
    cnPackingSlip.Open cnDatabase.ConnectionString
    rstVchSeriesList.Open "SELECT Name As Col0,Code FROM VchSeriesMaster WHERE Left(FYCode,2)='" & Left(FYCode, 2) & "' AND VchType= '" & VchType & "' ORDER BY Name", cnPackingSlip, adOpenKeyset, adLockReadOnly
    If rstVchSeriesList.RecordCount > 0 Then
        rstVchSeriesList.MoveFirst
        Dim i As Integer
        i = 0
        Do Until rstVchSeriesList.EOF
            Combo7.AddItem Trim(rstVchSeriesList.Fields(0).Value), i: i = i + 1
            rstVchSeriesList.MoveNext
        Loop
        Combo7.ListIndex = 0: cmdChange_Click
        rstDelivererList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='*99998' ORDER BY Name", cnPackingSlip, adOpenKeyset, adLockReadOnly
        rstTransporterList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='*99996' ORDER BY Name", cnPackingSlip, adOpenKeyset, adLockReadOnly
        rstBookingRouteList.Open "SELECT Name As Col0,Code FROM BookingRouteMaster ORDER BY Name", cnPackingSlip, adOpenKeyset, adLockReadOnly
        rstPackerList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='*99997' ORDER BY Name", cnPackingSlip, adOpenKeyset, adLockReadOnly
        rstDelivererList.ActiveConnection = Nothing
        rstTransporterList.ActiveConnection = Nothing
        rstBookingRouteList.ActiveConnection = Nothing
        rstVchSeriesList.ActiveConnection = Nothing
'        SetButtonsForNoRecord
    End If
    BusySystemIndicator False
    rstPackerList.ActiveConnection = Nothing
    With fpSpread1
        .Col = 6
        .ColHidden = True
    End With
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True
    MdiMainMenu.mnuPackingSlip.Enabled = False
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
                If Not EditMode Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
            If Not EditMode Then KeyCode = 0
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD And Toolbar1.Buttons.Item(3).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF2 And Toolbar1.Buttons.Item(4).Enabled Then
        If Not EditMode Then SaveAndContinue = True: Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        If Not EditMode Then SaveAndContinue = False: Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF9 And Toolbar1.Buttons.Item(6).Enabled Then
        Call UpdatePktPick(rstPackingSlipList.Fields(0).Value)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF10 And Toolbar1.Buttons.Item(4).Enabled Then
        If Me.ActiveControl.Name = "fpSpread1" Then AddSlip: UpdateBalQty
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF11 And Toolbar1.Buttons.Item(4).Enabled Then
        If Me.ActiveControl.Name = "fpSpread1" Then If Not chkPacked Then LoadUnpackedBillsList 'Clubbed Packing
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(9)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(10)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyF And Toolbar1.Buttons.Item(1).Enabled Then
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
            SSTab1.Tab = 1: SSTab1.SetFocus
        Else
           If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstPackingSlipList)
    Call CloseRecordset(rstPackingSlipChild)
    Call CloseRecordset(rstBillInfo)
    Call CloseRecordset(rstDelivererList)
    Call CloseRecordset(rstTransporterList)
    Call CloseRecordset(rstBookingRouteList)
    Call CloseRecordset(rstPackerList)
    Call CloseRecordset(rstVchSeriesList)
    Call CloseConnection(cnPackingSlip)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuPackingSlip.Enabled = True
End Sub
Private Sub Text1_Change()
On Error Resume Next
With rstPackingSlipList
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
    If rstPackingSlipList.State = adStateClosed Then Exit Sub
    If rstPackingSlipList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstPackingSlipList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstPackingSlipList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstPackingSlipList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstPackingSlipList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstPackingSlipList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstPackingSlipList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstPackingSlipList
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
            Me.Width = 14595 '9045
            Mh3dFrame1.Width = 14475 '8930
            SSTab1.Width = 14250 '8700
            CenterForm Me
            ViewRecord
        Else
            Me.Width = 14595 '13155
            Mh3dFrame1.Width = 14475 '13040
            SSTab1.Width = 14250 '12810
            CenterForm Me
            If Not (rstPackingSlipList.EOF Or rstPackingSlipList.BOF) Then
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
        Me.Width = 14595 '9045
        Mh3dFrame1.Width = 14475 '8930
        SSTab1.Width = 15000 '8700
        CenterForm Me
        Text8.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean, Reply As String
    Dim UpdateFlag As Integer, i As Integer
    Dim CellVal As Variant
    If Button.Index = 2 Then
        If rstPackingSlipList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        fpSpread1.Col = 5: fpSpread1.Lock = False
        If UserLevel <> 1 Then If VchCreator <> UserCode Then fpSpread1.Col = 5: fpSpread1.Lock = True
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstPackingSlipList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnPackingSlip.Execute "DELETE FROM PackingSlipParent WHERE Code = '" & rstPackingSlipList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstPackingSlipList.Delete
                rstPackingSlipList.MoveNext
                If rstPackingSlipList.RecordCount > 0 And rstPackingSlipList.EOF Then rstPackingSlipList.MoveLast
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
        If rstPackingSlipList.Fields("GR No").Value <> "" And AllowTransactionsModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Voucher")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        AddSlip 'Delete rows with Packed Qty 0
        Bundles = 0
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.GetText 5, i, CellVal
            If Val(CellVal) <> 0 Then
                fpSpread1.GetText 1, i, CellVal
                If Val(CellVal) <> 0 Then Bundles = Bundles + 1
            End If
        Next
        If Bundles <> 0 Then MhRealInput13.Text = Format(Bundles, "0") Else Bundles = Val(MhRealInput13.Text)
        UpdateFlag = 0
        If UpdateItemList("D") Then
            UpdateFlag = 1
            For i = 1 To fpSpread1.DataRowCnt
                fpSpread1.SetActiveCell 1, i
                fpSpread1.GetText 5, i, CellVal
                If Val(CellVal) <> 0 Then
                    If Not UpdateItemList("I") Then
                        UpdateFlag = 0
                        Exit For
                    End If
                End If
            Next
        End If
        If UpdateFlag = 1 Then
'            If Not CheckEmpty(Text11.Text, False) Then
'                If UserLevel <> 3 And Not SaveAndContinue Then If MsgBox("SMS Consignment Details?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm SMS !") = vbYes Then SMSConsignmentDetails
'            End If
        End If
        If UpdateFlag Then
            cnPackingSlip.CommitTrans
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            If SaveAndContinue Then
                Form_KeyDown vbKeyE, vbCtrlMask
                If MsgBox("Print Packing Slip?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Print !") = vbYes Then
                    GetPackingSlipNoRange
                    PrintPackingSlip rstPackingSlipList.Fields("Code").Value, "P"
                End If
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 5, i
                        .GetText 5, i, CellVal
                        If Val(CellVal) = 0 Then Exit For
                    Next
                    .SetFocus
                End With
            End If
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        cnPackingSlip.RollbackTrans
        Call SetButtons(True)
        SetButtonsForNoRecord
        SSTab1.Tab = 0
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        cmdChange_Click
        rstDelivererList.ActiveConnection = cnPackingSlip
        Do While Not RefreshRecord(rstDelivererList)
        Loop
        rstDelivererList.ActiveConnection = Nothing
        rstTransporterList.ActiveConnection = cnPackingSlip
        Do While Not RefreshRecord(rstTransporterList)
        Loop
        rstTransporterList.ActiveConnection = Nothing
        rstBookingRouteList.ActiveConnection = cnPackingSlip
        Do While Not RefreshRecord(rstBookingRouteList)
        Loop
        rstBookingRouteList.ActiveConnection = Nothing
        rstPackerList.ActiveConnection = cnPackingSlip
        Do While Not RefreshRecord(rstPackerList)
        Loop
        rstPackerList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Party Name", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstPackingSlipList.RecordCount = 0 Then Exit Sub
        DisplayMenu "P"
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstPackingSlipList.RecordCount = 0 Then Exit Sub
        DisplayMenu "S"
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstPackingSlipList.RecordCount > 0 Then rstPackingSlipList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstPackingSlipList.RecordCount > 0 Then
            rstPackingSlipList.MovePrevious
            If rstPackingSlipList.BOF Then rstPackingSlipList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstPackingSlipList.RecordCount > 0 Then
            rstPackingSlipList.MoveNext
            If rstPackingSlipList.EOF Then rstPackingSlipList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstPackingSlipList.RecordCount > 0 Then rstPackingSlipList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        CloseForm Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstPackingSlipList.EOF Or rstPackingSlipList.BOF) Then
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
        rstPackingSlipList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstPackingSlipList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstPackingSlipList.EOF Or rstPackingSlipList.BOF) Then
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
    Toolbar1.Buttons.Item(9).Enabled = bVal
    Toolbar1.Buttons.Item(10).Enabled = bVal
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstPackingSlipList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(9).Enabled = False
        Toolbar1.Buttons.Item(10).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text10_Change()
    If Text10.Text = " " Then Text10.Text = "?": Sendkeys "{TAB}"
End Sub
Private Sub Text10_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text10.Text)
    If rstDelivererList.RecordCount = 0 Then
        DisplayError ("No Record in Deliverer Master")
        Cancel = True
        Exit Sub
    Else
        rstDelivererList.MoveFirst
    End If
    rstDelivererList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstDelivererList.EOF Then
        SelectionType = "S"
        DelivererCode = ""
        Call LoadSelectionList(rstDelivererList, "List of Deliverers...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text10, DelivererCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text10.Text, False) Then
            Text10.Text = "?"
        End If
        If RTrim(DelivererCode) <> "" Then Sendkeys "{TAB}"
        Cancel = True
    Else
        DelivererCode = rstDelivererList.Fields("Code").Value
    End If
End Sub
Private Sub Text19_Change()
    If Text19.Text = " " Then
        Text19.Text = "?": Sendkeys "{TAB}"
    ElseIf CheckEmpty(Text19, False) Then
        TransporterCode = ""
    End If
End Sub
Private Sub Text19_Validate(Cancel As Boolean)
    Dim SearchString As String
    If CheckEmpty(Text19, False) Then Exit Sub
    SearchString = FixQuote(Text19.Text)
    If rstTransporterList.RecordCount = 0 Then
        DisplayError ("No Record in Transporter Master")
        Cancel = True
        Exit Sub
    Else
        rstTransporterList.MoveFirst
    End If
    rstTransporterList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstTransporterList.EOF Then
        SelectionType = "S"
        TransporterCode = ""
        Call LoadSelectionList(rstTransporterList, "List of Transporters...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text19, TransporterCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text19.Text, False) Then Text19.Text = "?"
        If RTrim(TransporterCode) <> "" Then Sendkeys "{TAB}"
        Cancel = True
    Else
        TransporterCode = rstTransporterList.Fields("Code").Value
    End If
End Sub
Private Sub Text15_Change()
    If Text15.Text = " " Then Text15.Text = "?": Sendkeys "{TAB}"
End Sub
Private Sub Text15_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text15.Text)
    If rstBookingRouteList.RecordCount = 0 Then
        DisplayError ("No Record in Booking Route Master")
        Cancel = True
        Exit Sub
    Else
        rstBookingRouteList.MoveFirst
    End If
    rstBookingRouteList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstBookingRouteList.EOF Then
        SelectionType = "S"
        BookingRouteCode = ""
        Call LoadSelectionList(rstBookingRouteList, "List of Booking Routes...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text15, BookingRouteCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text15.Text, False) Then Text15.Text = "?"
        If RTrim(BookingRouteCode) <> "" Then Sendkeys "{TAB}"
        Cancel = True
    Else
        BookingRouteCode = rstBookingRouteList.Fields("Code").Value
    End If
End Sub
Private Sub Check1_Click()
    If Check1.Value Then
        If MhDateInput3.ValueIsNull Then MhDateInput3.Text = Format(Date, "dd-MM-yyyy"): MhTimeInput2.Text = Mid(Format(Time, "hh:mm"), 1, 2) + "." + Mid(Format(Time, "hh:mm"), 4, 2)
        Text11.Enabled = True: MhDateInput5.Enabled = True
    Else
        MhDateInput3.Text = "  -  -    ": MhTimeInput2.Text = "00.00"
        Text11.Enabled = False: MhDateInput5.Enabled = False
    End If
End Sub
Private Sub Check2_Click()
    If Check2.Value Then Text19.Enabled = True Else Text19.Text = "": TransporterCode = "": Text19.Enabled = False
End Sub
Private Sub Text13_Change()
    If Text13.Text = " " Then
        Text13.Text = "?"
        Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text13_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text13.Text)
    If rstPackerList.RecordCount = 0 Then
        DisplayError ("No Record in Packer Master")
        Cancel = True
        Exit Sub
    Else
        rstPackerList.MoveFirst
    End If
    rstPackerList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstPackerList.EOF Then
        SelectionType = "S"
        PackerCode = ""
        Call LoadSelectionList(rstPackerList, "List of Packers...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text13, PackerCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text13.Text, False) Then
            Text13.Text = "?"
        End If
        If RTrim(PackerCode) <> "" Then
            Sendkeys "{TAB}"
        End If
        Cancel = True
    Else
        PackerCode = rstPackerList.Fields("Code").Value
    End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
    If UserLevel = 3 Then If KeyAscii >= 48 And KeyAscii <= 57 Then KeyAscii = 0
End Sub
Private Sub Text11_Validate(Cancel As Boolean)
    If Not CheckEmpty(Text11.Text, False) Then
         If MhDateInput4.Text = "  -  -    " Then
            MhDateInput4.Text = Format(Date, "dd-MM-yyyy")
            MhTimeInput3.Text = Mid(Format(Time, "hh:mm"), 1, 2) + "." + Mid(Format(Time, "hh:mm"), 4, 2)
            MhDateInput5.Text = Format(Date - 1, "dd-MM-yyyy")
        End If
    Else
        MhDateInput4.Text = "  -  -    "
        MhTimeInput3.Text = "00.00"
        MhDateInput5.Text = "  -  -    "
    End If
End Sub
Private Sub Text20_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call LoadUnpackedBillsList("G") 'Clubbed Bilty
    ElseIf KeyCode = vbKeyDelete Then
        Text20.Text = "": ClubbedGRVch = ""
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstPackingSlipList.EOF Then
        If rstPackingSlipChild.State = adStateOpen Then rstPackingSlipChild.Close
        Exit Sub
    End If
    LoadFields
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text15.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text11.Enabled = False
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    Text20.Text = ""
    MhRealInput13.Text = "0"
    MhDateInput1.Text = "  -  -    "
    MhDateInput2.Text = "  -  -    "
    MhDateInput3.Text = "  -  -    "
    MhDateInput4.Text = "  -  -    "
    MhDateInput5.Text = "  -  -    "
    MhDateInput5.Enabled = False
    MhTimeInput1.Text = "00.00"
    MhTimeInput2.Text = "00.00"
    MhTimeInput3.Text = "00.00"
    Check1.Value = 0
    Check2.Value = 0
    Text19.Text = ""
    Text19.Enabled = False
    DelivererCode = ""
    TransporterCode = ""
    BookingRouteCode = ""
    PackerCode = ""
    ClubbedVch = ""
    ClubbedGRVch = "": OldClubbedGRVch = ""
    VchCreator = ""
    BookingRouteCode = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True:    fpSpread1.SetActiveCell 1, 1
End Sub
Private Sub LoadFields()
    Dim GRType As String
    Text5.Text = rstPackingSlipList.Fields("Bill No").Value
    MhDateInput2.Text = Format(Date, "dd-MM-yyyy")
    MhTimeInput1.Text = Mid(Format(Time, "hh:mm"), 1, 2) + "." + Mid(Format(Time, "hh:mm"), 4, 2)
    VchCreator = UserCode:    lblVchCreator = "Created by : " & Username
    With rstPackingSlipChild
        If .State = adStateOpen Then .Close
        .Open "SELECT *,(SELECT Name FROM UserMaster WHERE Code=CreatedBy) As VchCreator,(SELECT Name FROM UserMaster WHERE Code=ModifiedBy) As VchEditor,(SELECT STUFF((SELECT '+'+(LTRIM(Name)) FROM JobWorkBVParent T WHERE CHARINDEX(''''+Code+'''',ClubbedVch)>0 FOR XML PATH('')),1,1,'')) As VchClubbed,(SELECT STUFF((SELECT '+'+(LTRIM(Name)) FROM JobWorkBVParent T WHERE CHARINDEX(''''+Code+'''',ClubbedGRVch)>0 FOR XML PATH('')),1,1,'')) As VchGRClubbed FROM PackingSlipParent WHERE Code='" & rstPackingSlipList.Fields("Code").Value & "'", cnPackingSlip, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        If .RecordCount > 0 Then
            MhDateInput2.Text = Format(.Fields("CreatedOn").Value, "dd-MM-yyyy")
            MhTimeInput1.Text = Format(.Fields("CreatedOn").Value, "hh.mm")
            MhDateInput3.Text = Format(.Fields("HandedOverOn").Value, "dd-MM-yyyy")
            MhTimeInput2.Text = Format(.Fields("HandedOverOn").Value, "hh.mm")
            MhDateInput4.Text = Format(.Fields("GREntryDate").Value, "dd-MM-yyyy")
            MhTimeInput3.Text = Format(.Fields("GREntryDate").Value, "hh.mm")
            Check2.Value = IIf(.Fields("PaidGR").Value, 1, 0)
            If Check2.Value Then Text19.Enabled = True
            Text16.Text = CheckNull(.Fields("Remarks").Value)
            DelivererCode = CheckNull(.Fields("Deliverer").Value)
            If rstDelivererList.RecordCount > 0 Then
                rstDelivererList.MoveFirst
                rstDelivererList.Find "[Code] = '" & DelivererCode & "'"
                If Not rstDelivererList.EOF Then
                    Text10.Text = rstDelivererList.Fields("Col0").Value
                End If
            End If
            TransporterCode = CheckNull(.Fields("Transporter").Value)
            If rstTransporterList.RecordCount > 0 Then
                rstTransporterList.MoveFirst
                rstTransporterList.Find "[Code] = '" & TransporterCode & "'"
                If Not rstTransporterList.EOF Then Text19.Text = rstTransporterList.Fields("Col0").Value
            End If
            BookingRouteCode = CheckNull(.Fields("BookingRoute").Value)
            If rstBookingRouteList.RecordCount > 0 Then
                rstBookingRouteList.MoveFirst
                rstBookingRouteList.Find "[Code] = '" & BookingRouteCode & "'"
                If Not rstBookingRouteList.EOF Then Text15.Text = rstBookingRouteList.Fields("Col0").Value
            End If
            PackerCode = CheckNull(.Fields("Packer").Value)
            If rstPackerList.RecordCount > 0 Then
                rstPackerList.MoveFirst
                rstPackerList.Find "[Code] = '" & PackerCode & "'"
                If Not rstPackerList.EOF Then Text13.Text = rstPackerList.Fields("Col0").Value
            End If
            ClubbedVch = CheckNull(.Fields("ClubbedVch").Value)
            ClubbedGRVch = CheckNull(.Fields("ClubbedGRVch").Value)
            OldClubbedGRVch = ClubbedGRVch
            VchCreator = CheckNull(.Fields("CreatedBy").Value): lblVchCreator = " Created by : " & StrConv(CheckNull(.Fields("VchCreator").Value), vbUpperCase)
            If Not CheckEmpty(CheckNull(.Fields("VchCreator").Value), False) Then lblVchCreator = lblVchCreator & " Modified by : " & StrConv(CheckNull(.Fields("VchEditor").Value), vbUpperCase)
            Text18.Text = CheckNull(.Fields("VchClubbed").Value)
            Text20.Text = CheckNull(.Fields("VchGRClubbed").Value)
            GRType = CheckNull(.Fields("GRType").Value)
        End If
    End With
    If ClubbedVch = "" Then ClubbedVch = "'" & rstPackingSlipList.Fields("Code").Value & "'"
    If Text18.Text = "" Then Text18.Text = Text5.Text
    Call LoadItemList(rstPackingSlipList.Fields("Code").Value)
    If Not CheckEmpty(GRType, False) Then Text14.Text = GRType
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstPackingSlipChild.State = adStateClosed Then SSTab1.Tab = 0: Exit Sub
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text8.SetFocus
    cnPackingSlip.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text10.Text, False) Then
       Text10.SetFocus
       CheckMandatoryFields = True
       Exit Function
    ElseIf Not CheckExists(Text10, "Col0", rstDelivererList, DelivererCode) Then
        Text10.SetFocus
        CheckMandatoryFields = True
        Exit Function
    End If
    If Text19.Enabled Then
        If Not CheckEmpty(Text19.Text, False) Then If Not CheckExists(Text19, "Col0", rstTransporterList, TransporterCode) Then Text19.SetFocus: CheckMandatoryFields = True: Exit Function
    End If
    If CheckEmpty(Text15.Text, False) Then
       Text15.SetFocus
       CheckMandatoryFields = True
       Exit Function
    ElseIf Not CheckExists(Text15, "Col0", rstBookingRouteList, BookingRouteCode) Then
        Text15.SetFocus
        CheckMandatoryFields = True
        Exit Function
    End If
    If CheckEmpty(Text13.Text, False) Then
       Text13.SetFocus
       CheckMandatoryFields = True
       Exit Function
    ElseIf Not CheckExists(Text13, "Col0", rstPackerList, PackerCode) Then
        Text13.SetFocus
        CheckMandatoryFields = True
        Exit Function
    End If
    If CheckPkdQty() Then
        fpSpread1.SetFocus
        CheckMandatoryFields = True
        Exit Function
    End If
End Function
Private Function CheckPkdQty() As Boolean
    Dim i As Integer, BilledQty As Variant, PackedQty As Variant, SNo As Variant, SlipNo As Variant
    CheckPkdQty = False
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.SetActiveCell 5, i
        fpSpread1.GetText 4, i, BilledQty
        fpSpread1.GetText 5, i, PackedQty
        fpSpread1.GetText 7, i, SlipNo
        fpSpread1.GetText 2, i, SNo
        If PackedQty > BilledQty Then
            DisplayError "Billed Qty is more than Packed Qty in Slip #" & Trim(Str(SlipNo)) & " Item #" & Trim(Str(SNo))
            CheckPkdQty = True
            Exit For
        End If
    Next
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
    If SrchFor = "Party Name" Then rstPackingSlipList.Filter = "[Party Name] Like '%" & SrchText & "%'"
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Name As String, i As Integer, BookName As Variant, PkdQty As Variant
    If Shift = 0 And KeyCode = vbKeySpace Then
        Dim Qty As Variant
        fpSpread1.GetText 4, fpSpread1.ActiveRow, Qty
        If Val(Qty) <> 0 Then fpSpread1.SetText 5, fpSpread1.ActiveRow, Val(Qty)
    ElseIf Shift = 0 And KeyCode = vbKeyF3 Then
        Name = InputBox("Enter the Name of the Book to be searched", "Book Search !")
        If Name <> "" Then
            For i = 1 To fpSpread1.DataRowCnt
                fpSpread1.GetText 3, i, BookName
                fpSpread1.GetText 5, i, PkdQty
                If InStr(1, StrConv(BookName, vbUpperCase), StrConv(Name, vbUpperCase)) > 0 And Val(PkdQty) = 0 Then fpSpread1.SetActiveCell 5, i: Exit For
            Next
        End If
        fpSpread1.SetFocus
    End If
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If Col = 5 Then
        Dim PackedQty As Variant, BilledQty As Variant
        fpSpread1.GetText 5, Row, PackedQty
        fpSpread1.GetText 4, Row, BilledQty
        If Val(PackedQty) > Val(BilledQty) Then Cancel = True Else UpdateBalQty
    End If
End Sub
Private Function UpdateItemList(ByVal ActionType As String) As Boolean
    On Error GoTo ErrorHandler
    UpdateItemList = True
    If ActionType = "D" Then
        cnPackingSlip.Execute "DELETE FROM PackingSlipParent WHERE Code = '" & rstPackingSlipList.Fields("Code").Value & "'"
        If Not CheckEmpty(OldClubbedGRVch, False) Then cnPackingSlip.Execute "UPDATE PackingSlipParent SET ClubbedGRVch='',ClubbedGRVchNo='',GREntryDate=NULL WHERE Code IN (" & OldClubbedGRVch & ")"
        cnPackingSlip.Execute "INSERT INTO PackingSlipParent VALUES('" & rstPackingSlipList.Fields("Code").Value & "','" & IIf(chkPacked, FixQuote(ClubbedVch), FixQuote("'" & rstPackingSlipList.Fields("Code").Value & "'")) & "','','" & IIf(chkPacked, FixQuote(ClubbedGRVch), "") & "','','" & DelivererCode & "'," & IIf(MhDateInput3.Text = "  -  -    ", "Null", "'" & GetDate(MhDateInput3.Text) + Space(1) + Mid(MhTimeInput2.Text, 1, 2) + ":" + Mid(MhTimeInput2.Text, 4, 2) + ":00'") & ",'" & Trim(Text14.Text) & "'," & IIf(MhDateInput4.Text = "  -  -    ", "Null", "'" & GetDate(MhDateInput4.Text) + Space(1) + Mid(MhTimeInput3.Text, 1, 2) + ":" + Mid(MhTimeInput3.Text, 4, 2) + ":00'") & "," & Val(Check2.Value) & ",'" & TransporterCode & "','" & PackerCode & "','" & BookingRouteCode & "','" & Text16.Text & "','" & VchCreator & "','" & GetDate(MhDateInput2.Text) + Space(1) + Mid(MhTimeInput1.Text, 1, 2) + ":" + Mid(MhTimeInput1.Text, 4, 2) + ":00" & "','" & UserCode & "','" & FYCode & "' )"
'        cnPackingSlip.Execute "UPDATE PackingSlipParent SET ClubbedVchNo=(SELECT STUFF((SELECT '+'+LTRIM(AutoVchNo) FROM JobWorkBVParent T WHERE CHARINDEX(''''+Code+'''',ClubbedVch)>0 FOR XML PATH('')),1,1,'')) WHERE Code='" & rstPackingSlipList.Fields("Code").Value & "'"
        cnPackingSlip.Execute "UPDATE PackingSlipParent SET ClubbedVchNo=(SELECT STUFF((SELECT '+'+LTRIM(Name) FROM JobWorkBVParent T WHERE CHARINDEX(''''+Code+'''',ClubbedVch)>0 FOR XML PATH('')),1,1,'')) WHERE Code='" & rstPackingSlipList.Fields("Code").Value & "'"
        'Update Bilty Details
        cnPackingSlip.Execute "DELETE FROM JobWorkBVOthInf WHERE Code='" & rstPackingSlipList.Fields("Code").Value & "'"
        If Not CheckEmpty(OldClubbedGRVch, False) Then cnPackingSlip.Execute "UPDATE JobWorkBVOthInf SET BiltyNo='',BiltyDate=NULL WHERE Code IN (" & OldClubbedGRVch & ")"
        cnPackingSlip.Execute "INSERT INTO JobWorkBVOthInf VALUES ('" & rstPackingSlipList.Fields("Code").Value & "','" & Trim(Text11.Text) & "'," & IIf(MhDateInput5.Text = "  -  -    ", "Null", "'" & GetDate(MhDateInput5.Text) & "'") & ",'" & IIf(StrConv(Trim(Text14.Text), vbUpperCase) = "SELF" Or StrConv(Trim(Text14.Text), vbUpperCase) = "BANK", FixQuote(StrConv(Trim(Text14.Text), vbUpperCase)), "DIRECT") & "'," & IIf(Bundles = 0, 0, MhRealInput13.Text) & ",'" & Trim(Text8.Text) & "','" & Trim(Text2.Text) & "'," & IIf(Check1.Value, 1, 0) & ")"
        If Not CheckEmpty(ClubbedGRVch, False) Then cnPackingSlip.Execute "UPDATE JobWorkBVOthInf SET BiltyNo='" & Trim(Text11.Text) & "',BiltyDate=" & IIf(MhDateInput5.Text = "  -  -    ", "Null", "'" & GetDate(MhDateInput5.Text) & "'") & " WHERE Code IN (" & ClubbedGRVch & ")"
    Else
        Dim CellVal(1 To 4) As Variant
        With fpSpread1
            .GetText 7, .ActiveRow, CellVal(1)  'Slip No.
            .GetText 2, .ActiveRow, CellVal(2)  'Serial No.
            .GetText 6, .ActiveRow, CellVal(3)  'Item Code
            .GetText 5, .ActiveRow, CellVal(4)  'Quantity
        End With
        cnPackingSlip.Execute "INSERT INTO PackingSlipChild VALUES ('" & rstPackingSlipList.Fields("Code").Value & "'," & CellVal(1) & "," & CellVal(2) & ",'" & CellVal(3) & "'," & Val(CellVal(4)) & ")"
    End If
    Exit Function
ErrorHandler:
    UpdateItemList = False
End Function
Private Sub UpdatePktPick(ByVal VchCode As String)
    On Error GoTo ErrorHandler
    cnPackingSlip.Execute "UPDATE PackingSlipParent SET ModifiedBy='" & UserCode & "' WHERE CONVERT(BIGINT,Code)=" & Val(VchCode) & " AND (HandedOverOn IS NULL OR HandedOverOn='')"
    cnPackingSlip.Execute "UPDATE PackingSlipParent SET HandedOverOn='" & Format(Now(), "dd-MMM-yyyy hh:mm") & "' WHERE CONVERT(BIGINT,Code)=" & Val(VchCode) & " AND HandedOverOn IS NULL" 'Despatch Database
    cnDatabase.Execute "UPDATE VchOtherInfo SET OF9='Y' WHERE VchCode=" & Val(VchCode) & " AND (OF9='' OR OF9 IS NULL)"   'Busy Database
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to update bundle pick")
End Sub
Public Function PrintPackingSlip(ByVal VchNo As String, ByVal OutputTo As String)
    Dim rstPrintPackingSlip As New ADODB.Recordset
    If LRange = 0 Then LRange = 1
    If URange = 0 Then URange = 9999
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rstPrintPackingSlip.Open "SELECT ISNULL(Station,'') As Station,SlipNo,(SELECT STUFF((SELECT '+'+(LTRIM(Name)) FROM JobworkBVParent T WHERE CHARINDEX(''''+Code+'''',ClubbedVch)>0 FOR XML PATH('')),1,1,'')) As PvtMark,SrNo,(SELECT PrintName FROM BookMaster WHERE Code=ItemCode) As ItemName,Quantity,(SELECT PrintName FROM AccountMaster WHERE Code=Packer) As Packer,(SELECT COUNT(*) FROM PackingSlipChild WHERE Code=P.Code AND SlipNo=C1.SlipNo GROUP BY SlipNo) As ItemInSlip,'' As Remarks,ISNULL(Transport,'') As Tranport,(SELECT Weight FROM BookMaster WHERE Code=ItemCode)*Quantity As ItemWt FROM (PackingSlipParent P INNER JOIN PackingSlipChild C1 ON P.Code=C1.Code) LEFT JOIN JobworkBVOthInf C2 ON P.Code=C2.Code WHERE P.Code='" & Trim(VchNo) & "' AND SlipNo>=" & LRange & " AND SlipNo<=" & URange & " ORDER BY CONVERT(INT,SlipNo),CONVERT(INT,SrNo)", cnPackingSlip, adOpenKeyset, adLockOptimistic
    rstPrintPackingSlip.ActiveConnection = Nothing
    Screen.MousePointer = vbNormal
    rptPackingSlip.Database.SetDataSource rstPrintPackingSlip, 3, 1
    rptPackingSlip.DiscardSavedData
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptPackingSlip
        FrmReportViewer.Show vbModal
    Else
        rptPackingSlip.PrintOut True
    End If
    Set rptPackingSlip = Nothing
    Call CloseRecordset(rstPrintPackingSlip)
    On Error GoTo 0
    LRange = 0: URange = 0
End Function
Public Function PrintPackingSlipWithPrivateMark(ByVal VchNo As String, ByVal OutputTo As String)
    Dim rstPrintPackingSlip As New ADODB.Recordset, rstCompanyMaster As New ADODB.Recordset, rstPrivateMark As New ADODB.Recordset
    If LRange = 0 Then LRange = 1
    If URange = 0 Then URange = 9999
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rstCompanyMaster.Open "Select * FROM CompanyMaster Where FYCode='" & FYCode & "' ", cnDatabase, adOpenKeyset, adLockReadOnly
'    rstPrivateMark.Open "SELECT (SELECT TOP 1 LTRIM(Name) FROM JobWorkBVParent T WHERE CHARINDEX(''''+Code+'''',ClubbedVch)>0 ORDER BY AutoVchNo DESC) As BillNo,Pkt As Bundles,UPPER(Station) As Station,CASE WHEN UPPER(BiltyType)='SELF' OR UPPER(BiltyType)='BANK' THEN 'SELF' ELSE GrType END As GRType,Transport,M.PrintName As Booking FROM (PackingSlipParent P LEFT JOIN JobWorkBVOthInf C ON P.Code=C.Code) LEFT JOIN BookingRouteMaster M ON P.BookingRoute=M.Code WHERE P.Code='" & Trim(VchNo) & "'", cnPackingSlip, adOpenKeyset, adLockOptimistic
    rstPrivateMark.Open "SELECT (SELECT TOP 1 LTRIM(Name) FROM JobWorkBVParent T WHERE CHARINDEX(''''+Code+'''',ClubbedVch)>0 ORDER BY Name DESC) As BillNo,Pkt As Bundles,UPPER(Station) As Station,CASE WHEN UPPER(BiltyType)='SELF' OR UPPER(BiltyType)='BANK' THEN 'SELF' ELSE GrType END As GRType,Transport,M.PrintName As Booking," & _
                                "(Select Name From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As Consignee,(Select Address1 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress1,(Select Address2 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress2,(Select Address3 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress3,(Select Address4 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress4,(Select Address4 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress4,(Select Mobile From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeMobile " & _
                                "FROM (PackingSlipParent P LEFT JOIN JobWorkBVOthInf C ON P.Code=C.Code) LEFT JOIN BookingRouteMaster M ON P.BookingRoute=M.Code WHERE P.Code='" & Trim(VchNo) & "'", cnPackingSlip, adOpenKeyset, adLockOptimistic
    rstPrivateMark.ActiveConnection = Nothing
    With rstPrintPackingSlip
        .Open "SELECT Station,SlipNo,(SELECT STUFF((SELECT '+'+(LTRIM(Name)) FROM JobWorkBVParent T WHERE CHARINDEX(''''+Code+'''',ClubbedVch)>0 FOR XML PATH('')),1,1,'')) As PvtMark,SrNo,(SELECT PrintName FROM BookMaster WHERE Code=ItemCode) As ItemName,Quantity,(SELECT PrintName FROM AccountMaster WHERE Code=Packer) As Packer,(SELECT COUNT(*) FROM PackingSlipChild WHERE Code=P.Code AND SlipNo=C1.SlipNo GROUP BY SlipNo) As ItemInSlip,'' As Remarks,Remarks As Transport,Remarks As Booking,(SELECT Weight FROM BookMaster WHERE Code=ItemCode)*Quantity As ItemWt FROM (PackingSlipParent P INNER JOIN PackingSlipChild C1 ON P.Code=C1.Code) LEFT JOIN JobWorkBVOthInf C2 ON P.Code=C2.Code WHERE P.Code='" & Trim(VchNo) & "' AND SlipNo>=" & LRange & " AND SlipNo<=" & URange & " ORDER BY CONVERT(INT,SlipNo),CONVERT(INT,SrNo)", cnPackingSlip, adOpenKeyset, adLockOptimistic
        .ActiveConnection = Nothing
        Do While Not .EOF
            .Fields("Booking").Value = rstPrivateMark.Fields("Booking").Value
            .Fields("Transport").Value = rstPrivateMark.Fields("Transport").Value
            .Update
            .MoveNext
        Loop
        .MoveFirst
    End With
    rptPackingSlipWithPrivateMark.Text17.SetText Trim(rstPrivateMark.Fields("BillNo").Value) + "/" + Trim(rstPrivateMark.Fields("Bundles").Value)
    rptPackingSlipWithPrivateMark.Text12.SetText StrConv(Trim(rstPrivateMark.Fields("GRType").Value), vbUpperCase)
    rptPackingSlipWithPrivateMark.Text14.SetText StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress1").Value), vbUpperCase) & " " & StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress2").Value), vbUpperCase) & " " & StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress3").Value), vbUpperCase) & " " & StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress4").Value), vbUpperCase) & ", Mobile : " & StrConv(Trim(rstPrivateMark.Fields("ConsigneeMobile").Value), vbUpperCase)
    rptPackingSlipWithPrivateMark.Text13.SetText StrConv(Trim(rstPrivateMark.Fields("Station").Value), vbUpperCase)
    rptPackingSlipWithPrivateMark.Text8.SetText "From: " & StrConv(Trim(rstCompanyMaster.Fields("PrintName").Value), vbUpperCase)
    rptPackingSlipWithPrivateMark.Text10.SetText StrConv(Trim(rstCompanyMaster.Fields("Address1").Value), vbUpperCase) & " " & StrConv(Trim(rstCompanyMaster.Fields("Address2").Value), vbUpperCase) & " " & StrConv(Trim(rstCompanyMaster.Fields("Address3").Value), vbUpperCase) & " " & StrConv(Trim(rstCompanyMaster.Fields("Address4").Value), vbUpperCase)
    Screen.MousePointer = vbNormal
    rptPackingSlipWithPrivateMark.Database.SetDataSource rstPrintPackingSlip, 3, 1
    rptPackingSlipWithPrivateMark.DiscardSavedData
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptPackingSlipWithPrivateMark
        FrmReportViewer.Show vbModal
    Else
        rptPackingSlipWithPrivateMark.PrintOut False
    End If
    Set rptPackingSlipWithPrivateMark = Nothing
    Call CloseRecordset(rstPrintPackingSlip): Call CloseRecordset(rstCompanyMaster)
    On Error GoTo 0
    LRange = 0: URange = 0
End Function
Public Function PrintForwardingSlip(ByVal VchNo As String, ByVal OutputTo As String)
    Dim rstCompanyMaster As New ADODB.Recordset, rstForwardingSlip As New ADODB.Recordset, rstForwardingSlipChild As New ADODB.Recordset
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rstCompanyMaster.Open "Select PrintName,Address1,Address2,Address3,Address4,Phone,Fax,Mobile FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    rstForwardingSlip.Open "SELECT CASE WHEN ClubbedGRVchNo='' THEN ClubbedVchNo ELSE ClubbedGRVchNo END As BillNo,P.CreatedOn,(SELECT SUM(Pkt) FROM JobworkBVOthInf WHERE CHARINDEX(''''+Code+'''',CASE WHEN ClubbedGRVch='' THEN ClubbedVch ELSE ClubbedGRVch END)>0) As Bundles,UPPER(Transport) As TransportName,Station,CASE WHEN UPPER(BiltyType)='SELF' OR UPPER(BiltyType)='BANK' THEN 'SELF' ELSE GrType END As GRType,(SELECT SUM(Amount) FROM JobworkBVParent T WHERE CHARINDEX(''''+Code+'''',CASE WHEN ClubbedGRVch='' THEN ClubbedVch ELSE ClubbedGRVch END)>0) As VchAmt,Remarks,PaidGR,ISNULL(M.Name,'') As Transporter,M1.PrintName As Booking,'' As Weight," & _
                                           "(Select PrintName From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As Consignee,(Select Address1 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress1,(Select Address2 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress2,(Select Address3 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress3,(Select Address4 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress4,(Select Mobile From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeMobile " & _
                                           "FROM ((PackingSlipParent P LEFT JOIN JobWorkBVOthInf C ON CONVERT(INT,P.Code)=C.Code) LEFT JOIN AccountMaster M ON P.Transporter=M.Code) LEFT JOIN BookingRouteMaster M1 ON P.BookingRoute=M1.Code WHERE P.Code='" & Trim(VchNo) & "'", cnPackingSlip, adOpenKeyset, adLockOptimistic
'    rstForwardingSlipChild.Open "SELECT CASE WHEN ClubbedVchNo IS NULL THEN LTRIM(AutoVchNo) ELSE ClubbedVchNo END As BillNo,Date As BillDate,Station,Transport As TransportName,Pkt As Bundles,ISNULL((SELECT Name FROM AccountMaster WHERE Code=C2.Packer),'') As Packer FROM (JobWorkBVParent P LEFT JOIN JobWorkBVOthInf C ON P.Code=C.Code) LEFT JOIN PackingSlipParent C2 ON P.Code=C2.Code WHERE LEFT(P.Type,2)='" & Left(VchType, 2) & "' AND RIGHT(P.Type,2)='" & Right(VchType, 2) & "' AND P.MaterialCentre IN (SELECT MaterialCentre FROM JobWorkBVParent WHERE Code=" & Trim(VchNo) & ") AND P.Party IN (SELECT Party FROM JobWorkBVParent WHERE Code=" & Trim(VchNo) & ") AND ISNULL(Station,'')='' AND PktPicked=0 AND P.Code<>'" & Trim(VchNo) & "' ORDER BY Date,P.Name", cnDatabase, adOpenKeyset, adLockOptimistic
    rstForwardingSlipChild.Open "SELECT CASE WHEN ClubbedVchNo IS NULL THEN LTRIM(Name) ELSE ClubbedVchNo END As BillNo,Date As BillDate,P.Station,P.Transport As TransportName,Pkt As Bundles,ISNULL((SELECT Name FROM AccountMaster WHERE Code=C2.Packer),'') As Packer FROM (JobWorkBVParent P LEFT JOIN JobWorkBVOthInf C ON P.Code=C.Code) LEFT JOIN PackingSlipParent C2 ON P.Code=C2.Code WHERE LEFT(P.Type,2)='" & Left(VchType, 2) & "' AND RIGHT(P.Type,2)='" & Right(VchType, 2) & "' AND P.MaterialCentre IN (SELECT MaterialCentre FROM JobWorkBVParent WHERE Code=" & Trim(VchNo) & ") AND P.Party IN (SELECT Party FROM JobWorkBVParent WHERE Code=" & Trim(VchNo) & ") AND ISNULL(P.Station,'')='' AND PktPicked=0 AND P.Code<>'" & Trim(VchNo) & "' ORDER BY Date,P.Name", cnDatabase, adOpenKeyset, adLockOptimistic
    rstForwardingSlip.ActiveConnection = Nothing: rstForwardingSlipChild.ActiveConnection = Nothing
'    If rstForwardingSlip.RecordCount > 0 Then rstForwardingSlip.Fields("Remarks").Value = "<Font Size='2' Face='Calibri' Align='LEFT'>Kindly arrange to book <Font Size='3'><b>" + Trim(rstForwardingSlip.Fields("Bundles").Value) + " Bundle(s) " & IIf(Val(rstForwardingSlip.Fields("Weight").Value) = 0, "", "(Weight : " & Trim(rstForwardingSlip.Fields("Weight").Value) & " Kg)") & " <Font Size='2'></b>of Printed Books (<Font Size='3'><b>Value : Rs. " + Format(rstForwardingSlip.Fields("VchAmt").Value, "0.00") + "<Font Size='2'></b>) to <Font Size='3'><b>" + StrConv(rstForwardingSlip.Fields("Station").Value, vbUpperCase) + "<Font Size='2'></b> and oblige." Else rstForwardingSlip.Fields("Remarks").Value = ""
    If rstForwardingSlip.RecordCount > 0 Then rstForwardingSlip.Fields("Remarks").Value = "<Font Size='2' Face='Calibri' Align='LEFT'>Kindly arrange to book <Font Size='3'><b>" + Trim(rstForwardingSlip.Fields("Bundles").Value) + " Bundle(s) <Font Size='2'></b>of Printed Books (<Font Size='3'><b>Value : Rs. " + Format(rstForwardingSlip.Fields("VchAmt").Value, "0.00") + "<Font Size='2'></b>) to <Font Size='3'><b>" + StrConv(rstForwardingSlip.Fields("Station").Value, vbUpperCase) + "<Font Size='2'></b> and oblige." Else rstForwardingSlip.Fields("Remarks").Value = ""
    If rstForwardingSlipChild.RecordCount = 0 Then rptForwardingSlip.Subreport1.BottomLineStyle = crLSNoLine: rptForwardingSlip.Subreport1.TopLineStyle = crLSNoLine: rptForwardingSlip.Subreport1.LeftLineStyle = crLSNoLine: rptForwardingSlip.Subreport1.RightLineStyle = crLSNoLine
    Screen.MousePointer = vbNormal
    'Header
    rptForwardingSlip.Text1.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptForwardingSlip.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptForwardingSlip.Text2.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    
    rptForwardingSlip.Text12.SetText StrConv(Trim(rstCompanyMaster.Fields("PrintName").Value), vbProperCase)
    rptForwardingSlip.Text13.SetText StrConv(Trim(rstCompanyMaster.Fields("Address1").Value), vbProperCase)
    rptForwardingSlip.Text17.SetText StrConv(Trim(rstCompanyMaster.Fields("Address2").Value), vbProperCase)
    rptForwardingSlip.Text18.SetText StrConv(Trim(rstCompanyMaster.Fields("Address3").Value), vbProperCase) & " " & StrConv(Trim(rstCompanyMaster.Fields("Address4").Value), vbProperCase) & " Mobile : " & Trim(rstCompanyMaster.Fields("Mobile").Value)
    rptForwardingSlip.Text9.SetText "for " + StrConv(Trim(rstCompanyMaster.Fields("PrintName").Value), vbProperCase)
    
    rptForwardingSlip.Text5.SetText "Address : " & StrConv(Trim(rstForwardingSlip.Fields("ConsigneeAddress1").Value), vbProperCase)
    rptForwardingSlip.Text11.SetText StrConv(Trim(rstForwardingSlip.Fields("ConsigneeAddress2").Value), vbProperCase)
    rptForwardingSlip.Text15.SetText StrConv(Trim(rstForwardingSlip.Fields("ConsigneeAddress3").Value), vbProperCase)
    rptForwardingSlip.Text16.SetText StrConv(Trim(rstForwardingSlip.Fields("ConsigneeAddress4").Value), vbProperCase)
    rptForwardingSlip.Text19.SetText "Mobile : " & Trim(rstForwardingSlip.Fields("ConsigneeMobile").Value)
    rptForwardingSlip.DiscardSavedData
    rptForwardingSlip.Database.SetDataSource rstForwardingSlip, 3, 1
    rptForwardingSlip.Subreport1.OpenSubreport.Database.SetDataSource rstForwardingSlipChild, 3, 1
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptForwardingSlip
        FrmReportViewer.Show vbModal
    Else
        rptForwardingSlip.PrintOut False    'Print Report Without Prompt
    End If
    Set rptForwardingSlip = Nothing
    Call CloseRecordset(rstForwardingSlip): Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstForwardingSlipChild)
    On Error GoTo 0
End Function
Public Function PrintPrivateMark(ByVal VchNo As String, ByVal OutputTo As String)
    Dim rstCompanyMaster As New ADODB.Recordset, rstPrivateMark As New ADODB.Recordset
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rstCompanyMaster.Open "Select * FROM CompanyMaster Where FYCode='" & FYCode & "' ", cnDatabase, adOpenKeyset, adLockReadOnly
'    rstPrivateMark.Open "SELECT (SELECT TOP 1 LTRIM(Name) FROM JobWorkBVParent T WHERE CHARINDEX(''''+Code+'''',ClubbedVch)>0 ORDER BY AutoVchNo DESC) As BillNo,Pkt As Bundles,UPPER(Station) As Station,CASE WHEN UPPER(BiltyType)='SELF' OR UPPER(BiltyType)='BANK' THEN 'SELF' ELSE GrType END As GRType,Transport,M.PrintName As Booking FROM (PackingSlipParent P LEFT JOIN JobWorkBVOthInf C ON P.Code=C.Code) LEFT JOIN BookingRouteMaster M ON P.BookingRoute=M.Code WHERE P.Code='" & Trim(VchNo) & "'", cnPackingSlip, adOpenKeyset, adLockOptimistic
    rstPrivateMark.Open "SELECT (SELECT TOP 1 LTRIM(Name) FROM JobWorkBVParent T WHERE CHARINDEX(''''+Code+'''',ClubbedVch)>0 ORDER BY Name DESC) As BillNo,Pkt As Bundles,UPPER(Station) As Station,CASE WHEN UPPER(BiltyType)='SELF' OR UPPER(BiltyType)='BANK' THEN 'SELF' ELSE GrType END As GRType,Transport,M.PrintName As Booking,  " & _
                                      "(Select PrintName From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As Consignee,(Select Address1 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress1,(Select Address2 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress2,(Select Address3 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress3,(Select Address4 From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeAddress4,(Select Mobile From AccountMaster Where Code= (Select Consignee From JobworkBVParent Where Code='" & Trim(VchNo) & "')) As ConsigneeMobile " & _
                                      "FROM (PackingSlipParent P LEFT JOIN JobWorkBVOthInf C ON P.Code=C.Code) LEFT JOIN BookingRouteMaster M ON P.BookingRoute=M.Code WHERE P.Code='" & Trim(VchNo) & "'", cnPackingSlip, adOpenKeyset, adLockOptimistic
    rstPrivateMark.ActiveConnection = Nothing
    Screen.MousePointer = vbNormal
    rptPrivateMark.Text5.SetText StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress1").Value), vbUpperCase) & " " & StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress2").Value), vbUpperCase) & " " & StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress3").Value), vbUpperCase) & " " & StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress4").Value), vbUpperCase)
    rptPrivateMark.Text3.SetText "From: " & StrConv(Trim(rstCompanyMaster.Fields("PrintName").Value), vbUpperCase)
    rptPrivateMark.Text4.SetText StrConv(Trim(rstCompanyMaster.Fields("Address1").Value), vbUpperCase) & " " & StrConv(Trim(rstCompanyMaster.Fields("Address2").Value), vbUpperCase) & " " & StrConv(Trim(rstCompanyMaster.Fields("Address3").Value), vbUpperCase) & " " & StrConv(Trim(rstCompanyMaster.Fields("Address4").Value), vbUpperCase)
    rptPrivateMark.Text10.SetText StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress1").Value), vbUpperCase) & " " & StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress2").Value), vbUpperCase) & " " & StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress3").Value), vbUpperCase) & " " & StrConv(Trim(rstPrivateMark.Fields("ConsigneeAddress4").Value), vbUpperCase)
    rptPrivateMark.Text7.SetText "From: " & StrConv(Trim(rstCompanyMaster.Fields("PrintName").Value), vbUpperCase)
    rptPrivateMark.Text9.SetText StrConv(Trim(rstCompanyMaster.Fields("Address1").Value), vbUpperCase) & " " & StrConv(Trim(rstCompanyMaster.Fields("Address2").Value), vbUpperCase) & " " & StrConv(Trim(rstCompanyMaster.Fields("Address3").Value), vbUpperCase) & " " & StrConv(Trim(rstCompanyMaster.Fields("Address4").Value), vbUpperCase)
    
    'rptPrivateMark.Section4.Suppress = True
    If MsgBox("Are You want to Print in A5 Paper Size in Place of A4 .. ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Paper Size !") = vbYes Then
    rptPrivateMark.PaperSize = crPaperA5: rptPrivateMark.PaperOrientation = crLandscape
    Else
    rptPrivateMark.PaperSize = crDefaultPaperSize: rptPrivateMark.PaperOrientation = crDefaultPaperOrientation
    End If
    rptPrivateMark.Database.SetDataSource rstPrivateMark, 3, 1
    rptPrivateMark.DiscardSavedData
    'rptPrivateMark.Areas.Item("D").CopiesToPrint = 2
    rptPrivateMark.BottomMargin = 60
    rptPrivateMark.TopMargin = 100
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptPrivateMark
        FrmReportViewer.Show vbModal
    Else
        rptPrivateMark.PrintOut True
    End If
    Set rptPrivateMark = Nothing
    Call CloseRecordset(rstPrivateMark):     Call CloseRecordset(rstCompanyMaster)
    On Error GoTo 0
End Function
Private Sub DisplayMenu(ByVal OutputTo As String)
    Dim menusel As String
    If rstPackingSlipList.RecordCount = 0 Then Exit Sub
    menusel = DisplayPopupMenu(Me.hwnd, 3)
    Select Case menusel
        Case 1
            GetPackingSlipNoRange
            PrintPackingSlip rstPackingSlipList.Fields("Code").Value, OutputTo
        Case 2
            PrintForwardingSlip rstPackingSlipList.Fields("Code").Value, OutputTo
        Case 3
            PrintPrivateMark rstPackingSlipList.Fields("Code").Value, OutputTo
        Case 4
            GetPackingSlipNoRange
            PrintPackingSlipWithPrivateMark rstPackingSlipList.Fields("Code").Value, OutputTo
    End Select
    If Not (rstPackingSlipList.EOF Or rstPackingSlipList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.SetFocus
End Sub
Private Sub SetGridColumnsWidth()
    DataGrid1.Columns(0).Visible = False
'    DataGrid1.Columns(1).Visible = False
    DataGrid1.Columns(1).Width = 1230
    DataGrid1.Columns(2).Width = 1065
    DataGrid1.Columns(3).Width = 1365
    DataGrid1.Columns(4).Width = 1665
    DataGrid1.Columns(5).Width = 1065
    DataGrid1.Columns(6).Visible = False
    DataGrid1.Columns(7).Width = 705
    DataGrid1.Columns(8).Width = 3100
    DataGrid1.Columns(9).Width = 1350
    DataGrid1.Columns(10).Width = 1379
    DataGrid1.Columns(11).Width = 1250
    DataGrid1.Columns(7).Alignment = dbgRight
End Sub
Private Function GetPackingSlipNoRange()
    Load FrmNumberRange
    FrmNumberRange.Show vbModal
    If FrmNumberRange.Text1.Text <> "" Then LRange = Val(FrmNumberRange.Text1.Text): URange = Val(FrmNumberRange.Text2.Text)
    Call CloseForm(FrmNumberRange)
End Function
Private Sub cmdChange_Click()
    On Error GoTo ErrorHandler
    With rstVchSeriesList
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        .Find "[Col0] = '" & Trim(Combo7.SelText) & "'"
        VchSeries = .Fields("Code").Value
    End With
    BusySystemIndicator True
    Set DataGrid1.DataSource = Nothing
    With rstPackingSlipList
        If .State = adStateOpen Then .Close
        .Open "SELECT P.Code,LTRIM(P.Name) As VchNo,LTRIM(P.Name) As [Bill No],CONVERT(VARCHAR,Date,105) As [Bill Date],G.Name As [Material Centre],C.BiltyNo As [GR No],C.BiltyDate As [GR Date],C.BiltyType As [GR Type],C.Pkt As Bundles,A.Name As [Party Name],C.Station,C.Transport As [Transport Name],CASE WHEN C.PktPicked IS NULL OR  C.PktPicked=0 THEN 'No' ELSE 'Yes' END As [Bundles Lifted],P.Remarks As Narration FROM ((JobWorkBVParent P LEFT JOIN JobWorkBVOthInf C ON P.Code=C.Code) INNER JOIN AccountMaster G ON P.MaterialCentre=G.Code) INNER JOIN AccountMaster A ON P.Party=A.Code WHERE LEFT(P.Type,2)='" & Left(VchType, 2) & "' AND RIGHT(P.Type,2)='" & Right(VchType, 2) & "' AND FYCode='" & FYCode & "' ORDER BY Date,P.Name", cnPackingSlip, adOpenKeyset, adLockReadOnly
'        .Open "SELECT P.Code,LTRIM(P.AutoVchNo) As VchNo,LTRIM(P.Name) As [Bill No],CONVERT(VARCHAR,Date,105) As [Bill Date],G.Name As [Material Centre],C.BiltyNo As [GR No],C.BiltyDate As [GR Date],C.BiltyType As [GR Type],C.Pkt As Bundles,A.Name As [Party Name],C.Station,C.Transport As [Transport Name],CASE WHEN C.PktPicked IS NULL OR  C.PktPicked=0 THEN 'No' ELSE 'Yes' END As [Bundles Lifted],P.Remarks As Narration FROM ((JobWorkBVParent P LEFT JOIN JobWorkBVOthInf C ON P.Code=C.Code) INNER JOIN AccountMaster G ON P.MaterialCentre=G.Code) INNER JOIN AccountMaster A ON P.Party=A.Code WHERE LEFT(P.Type,2)='" & Left(VchType, 2) & "' AND RIGHT(P.Type,2)='" & Right(VchType, 2) & "' ORDER BY Date,AutoVchNo", cnPackingSlip, adOpenKeyset, adLockReadOnly
        .Filter = adFilterNone
        .ActiveConnection = Nothing
        If .RecordCount > 0 Then .MoveLast
        Set DataGrid1.DataSource = rstPackingSlipList
        SetGridColumnsWidth
        SSTab1.Tab = 0
        SortOrder = "Bill No"
        If Not (.EOF Or .BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        'If (Not .EOF) And (Not .EOF) Then Text1.Text = .Fields("Bill No").Value
    End With
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Bill List")
    BusySystemIndicator False
End Sub
Private Sub LoadItemList(ByVal VchCode As Variant) 'Done
    On Error GoTo ErrorHandler
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    With rstBillInfo
        If .State = adStateOpen Then .Close
        .Sort = ""
        .Open "SELECT DISTINCT A.Name As PartyName,A.PrintName,A.Address1,A.Address2,A.Address3,A.Address4,P.Date As BillDate,ISNULL(P.Remarks,'') As Narration,G.Name As MaterialCentre,ISNULL(C.Transport,'') As TransportName,ISNULL(C.Station,'') As Station,ISNULL(C.BiltyType,'') As GRType,ISNULL(C.BiltyNo,'') As GRNo,C.BiltyDate As GRDate,ISNULL(C.Pkt,'') As Bundles,ISNULL(C.PktPicked,0) As InTransit FROM ((JobWorkBVParent P LEFT JOIN JobWorkBVOthInf C ON P.Code=C.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster G ON P.MaterialCentre=G.Code WHERE P.Code='" & VchCode & "'", cnPackingSlip, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        MhDateInput1.Text = Format(.Fields("BillDate").Value, "dd-MM-yyyy")
        Text3.Text = Trim(.Fields("PartyName").Value)
        Text6.Text = CheckNull(Trim(.Fields("Address1").Value))
        Text7.Text = CheckNull(Trim(.Fields("Address2").Value))
        Text12.Text = CheckNull(Trim(.Fields("Address3").Value))
        Text17.Text = CheckNull(Trim(.Fields("Address4").Value))
        If Trim(.Fields("Station").Value) <> "" Then Text8.Text = Trim(.Fields("Station").Value)
        Text9.Text = Trim(.Fields("MaterialCentre").Value)
        If Trim(.Fields("TransportName").Value) <> "" Then Text2.Text = Trim(.Fields("TransportName").Value)
        Text4.Text = Trim(.Fields("Narration").Value)
        Text14.Text = IIf(StrConv(Trim(.Fields("GRType").Value), vbUpperCase) = "SELF" Or StrConv(Trim(.Fields("GRType").Value), vbUpperCase) = "BANK", StrConv(Trim(.Fields("GRType").Value), vbUpperCase), StrConv(Trim(.Fields("PrintName").Value), vbUpperCase))
        If Trim(.Fields("GRNo").Value) <> "" Then Text11.Text = Trim(.Fields("GRNo").Value)
        If Not IsNull(.Fields("GRDate").Value) Then MhDateInput5.Text = Format(.Fields("GRDate").Value, "dd-MM-yyyy")
        MhRealInput13.Value = Val(.Fields("Bundles").Value)
        Check1.Value = IIf(.Fields("InTransit").Value, 1, 0)
        'Unique items in the bill
        If .State = adStateOpen Then .Close
        .Open "SELECT DISTINCT I.Code As ItemCode,PrintName+' [Price : '+CONVERT(VARCHAR,T.Rate)+']' As ItemName,Quantity As Qty,Quantity As BalQty,Quantity As SrNo FROM JobWorkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code WHERE T.Code IN (" & ClubbedVch & ")", cnPackingSlip, adOpenKeyset, adLockOptimistic
        .ActiveConnection = Nothing
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do Until .EOF
            .Fields("Qty").Value = 0: .Fields("BalQty").Value = 0: .Fields("SrNo").Value = 0
            .Update
            .MoveNext
        Loop
    End With
    With rstPackingSlipChild 'For Totaling of Qty of Duplicate Items - Start (if case of duplicate items in the bill)
        If .State = adStateOpen Then .Close
        .Open "SELECT Item As ItemCode,SrNo,Quantity FROM JobWorkBVChild WHERE Code IN (" & ClubbedVch & ") ORDER BY Item,SrNo", cnPackingSlip, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        Dim BilledQty As Long
        .MoveFirst
        Do Until .EOF
            With rstBillInfo
                .MoveFirst
                .Find "[ItemCode]='" & rstPackingSlipChild.Fields("ItemCode").Value & "'"
                If Not .EOF Then
                    .Fields("Qty").Value = .Fields("Qty").Value + Val(rstPackingSlipChild.Fields("Quantity").Value)
                    .Fields("BalQty").Value = .Fields("BalQty").Value + Val(rstPackingSlipChild.Fields("Quantity").Value)
                    .Fields("SrNo").Value = Val(rstPackingSlipChild.Fields("SrNo").Value)
                    .Update
                    BilledQty = BilledQty + Val(Abs(rstPackingSlipChild.Fields("Quantity").Value))
                End If
            End With
            .MoveNext
        Loop
        lblBilledQty.Caption = BilledQty
        If InStr(1, ClubbedVch, ",'") = 0 Then rstBillInfo.Sort = "[SrNo] Asc" Else rstBillInfo.Sort = "[ItemName] Asc"
        If .State = adStateOpen Then .Close
        .Open "SELECT T.*,I.PrintName As ItemName FROM PackingSlipChild T INNER JOIN BookMaster I ON T.ItemCode=I.Code WHERE T.Code='" & CheckNull(VchCode) & "' ORDER BY SlipNo,SrNo", cnPackingSlip, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        If .RecordCount > 0 Then
            Dim i As Integer
            i = 0
            .MoveFirst
            Do Until .EOF
                i = i + 1
                fpSpread1.SetText 7, i, Val(.Fields("SlipNo").Value)
                fpSpread1.SetText 2, i, Val(.Fields("SrNo").Value)
                fpSpread1.SetText 3, i, .Fields("ItemName").Value
                fpSpread1.SetText 4, i, 0
                fpSpread1.SetText 5, i, Val(.Fields("Quantity").Value)
                fpSpread1.SetText 6, i, .Fields("ItemCode").Value
                .MoveNext
            Loop
            UpdateBalQty
        End If
    End With
    AddSlip
    UpdateBalQty
    fpSpread1.SetActiveCell 5, 1
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Item List")
End Sub
Private Sub AddSlip()
    Dim i As Integer, n As Integer, SlipNo As Variant, SrNo As Variant, ItemCode As Variant, PackedQty As Variant, TSlipNo As Variant
    Dim SortKeys, SortKeyOrder As Variant, Item01 As Variant, Item02 As Variant
    With fpSpread1
        .GetText 3, .ActiveRow, Item01  'Item on the Current Cursor position
        SortKeys = Array(7, 3): SortKeyOrder = Array(1, 1): .Sort 1, 1, 7, .DataRowCnt, SortByRow, SortKeys, SortKeyOrder   'Sort on SlipNo+ItemName in Asc Order
        i = 1: TSlipNo = 0
        'Display SlipNo only on the first row of the Slip
        Do While i <= .DataRowCnt
            .GetText 5, i, PackedQty
            If Val(PackedQty) = 0 Then
                .DeleteRows i, 1: i = i - 1
            Else
                .GetText 7, i, SlipNo
                If TSlipNo <> SlipNo Then
                    SrNo = 1: TSlipNo = SlipNo: .SetText 1, i, SlipNo
                Else
                    .SetText 1, i, ""
                End If
                .SetText 2, i, SrNo: SrNo = SrNo + 1
            End If
            i = i + 1
        Loop
        .GetText 7, .DataRowCnt, SlipNo 'Get Last SlipNo
        SlipNo = Val(SlipNo) + 1
        i = .DataRowCnt 'Get Last row No.
        n = i + 1   'Get row No. next to Last row No.
    End With
    'Add Items whose Bal Qty is not zero.
    SrNo = 0
    rstBillInfo.MoveFirst
    Do While Not rstBillInfo.EOF
        If Abs(Val(rstBillInfo.Fields("BalQty").Value)) <> 0 Then
            i = i + 1
            With fpSpread1
                SrNo = Val(SrNo) + 1
                .SetText 2, i, SrNo
                .SetText 3, i, rstBillInfo.Fields("ItemName").Value
                .SetText 4, i, 0
                .SetText 5, i, 0
                .SetText 6, i, rstBillInfo.Fields("ItemCode").Value
                .SetText 7, i, SlipNo
            End With
        End If
        rstBillInfo.MoveNext
    Loop
    'Moved the Cursor to Item saved earlier
    Dim PointerMoved As Boolean
    With fpSpread1
        For i = n To fpSpread1.DataRowCnt
            .GetText 3, i, Item02
            If Item01 = Item02 Then .SetActiveCell 5, i: PointerMoved = True: Exit Sub
        Next
    End With
    If Not PointerMoved Then fpSpread1.SetActiveCell 5, n
End Sub
Private Sub UpdateBalQty()
    Dim i As Integer, ItemCode As Variant, PackedQty As Variant, SlipNo As Variant, n As Integer, BilledQty As Variant
    With rstBillInfo
        .MoveFirst
        Do While Not .EOF
            .Fields("BalQty").Value = .Fields("Qty").Value 'Bal Qty=Billing Qty
            .Update
            .MoveNext
        Loop
    End With
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 6, i, ItemCode
            If Not CheckEmpty(ItemCode, False) Then
                .GetText 7, i, SlipNo
                If n <> Val(SlipNo) Then    'Display SlipNo only on the first row of the Slip
                    n = Val(SlipNo)
                    .SetText 1, i, Val(SlipNo)
                End If
                .GetText 5, i, PackedQty
                rstBillInfo.MoveFirst
                rstBillInfo.Find "[ItemCode]='" & ItemCode & "'"
                If Not rstBillInfo.EOF Then
                    rstBillInfo.Fields("BalQty").Value = Val(rstBillInfo.Fields("BalQty").Value) + Val(PackedQty)
                    rstBillInfo.Update
                    .SetText 4, i, Abs(Val(rstBillInfo.Fields("BalQty").Value)) + Val(PackedQty)    'Abs(BalQty)+PackedQty
                End If
            End If
        Next
        i = 1
        Dim PkdQty As Long
        Do While i <= .DataRowCnt  'Delete rows with Billed+Packed=0
            .GetText 4, i, BilledQty
            .GetText 5, i, PackedQty
            PkdQty = PkdQty + Val(PackedQty)
            If Val(PackedQty) + Val(BilledQty) = 0 Then
                .DeleteRows i, 1
                i = i - 1
            End If
            i = i + 1
        Loop
        lblPkdQty.Caption = PkdQty
    End With
End Sub
Private Function chkPacked() As Boolean 'Done
    Dim i As Integer, PackedQty As Variant
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 5, i, PackedQty
            If Val(PackedQty) <> 0 Then chkPacked = True: Exit Function
        Next
    End With
End Function
Private Sub LoadUnpackedBillsList(Optional ByVal ListType As String) 'Done
    With rstPackingSlipChild
        If .State = adStateOpen Then .Close
'        .Open "SELECT P.Code,LTRIM(Name) As BillNo,Date FROM JobWorkBVParent P LEFT JOIN JobWorkBVOthInf C ON P.Code=C.Code WHERE P.Code<>'" & rstPackingSlipList.Fields("Code").Value & "' AND ISNULL(C.BiltyNo,'')='' AND VchSeries=" & VchSeries & " AND P.MaterialCentre=(SELECT MaterialCentre FROM JobWorkBVParent WHERE Code='" & rstPackingSlipList.Fields("Code").Value & "') AND LEFT(Type,2)='" & Left(VchType, 2) & "' AND RIGHT(Type,2)='" & Right(VchType, 2) & "' ORDER BY Date DESC,AutoVchNo DESC", cnDatabase, adOpenKeyset, adLockReadOnly 'List of Unpacked Bills
        .Open "SELECT P.Code,LTRIM(Name) As BillNo,Date FROM JobWorkBVParent P LEFT JOIN JobWorkBVOthInf C ON P.Code=C.Code WHERE P.Code<>'" & rstPackingSlipList.Fields("Code").Value & "' AND ISNULL(C.BiltyNo,'')='' AND P.MaterialCentre=(SELECT MaterialCentre FROM JobWorkBVParent WHERE Code='" & rstPackingSlipList.Fields("Code").Value & "') AND LEFT(Type,2)='" & Left(VchType, 2) & "' AND RIGHT(Type,2)='" & Right(VchType, 2) & "' ORDER BY Date DESC,Name DESC", cnDatabase, adOpenKeyset, adLockReadOnly 'List of Unpacked Bills
        .ActiveConnection = Nothing
        If .RecordCount = 0 Then DisplayError ("No Unpacked Bill"): fpSpread1.SetFocus: Exit Sub
        Load FrmUnpackedBillsList
        FrmUnpackedBillsList.Text2 = Text3.Text
        Dim i As Integer
        For i = 1 To .RecordCount 'Add blank rows
            With FrmUnpackedBillsList.fpSpread1
                .MaxRows = .MaxRows + 1
                .InsertRows i, 1
            End With
        Next
        i = 0
        Do While Not .EOF 'Load Unpacked Items List
            i = i + 1
            FrmUnpackedBillsList.fpSpread1.SetText 1, i, .Fields("BillNo").Value
            FrmUnpackedBillsList.fpSpread1.SetText 2, i, Format(.Fields("Date").Value, "dd-MM-yy")
            FrmUnpackedBillsList.fpSpread1.SetText 4, i, .Fields("Code").Value
            If InStr(1, IIf(ListType = "G", ClubbedGRVch, ClubbedVch), "'" & .Fields("Code").Value & "'") > 0 Then FrmUnpackedBillsList.fpSpread1.SetText 3, i, 1 'Check previously selected bills
            .MoveNext
        Loop
        FrmUnpackedBillsList.fpSpread1.SetActiveCell 3, 1
    End With
    FrmUnpackedBillsList.Show vbModal
    If FrmUnpackedBillsList.VchCodeList <> "" Then 'Something Selected
        If ListType = "G" Then 'Bilty
            ClubbedGRVch = FrmUnpackedBillsList.VchCodeList & ",'" & rstPackingSlipList.Fields("Code").Value & "'" 'Selected+Current
            OldClubbedGRVch = ClubbedGRVch
            Text20.Text = Text5.Text & "+" & FrmUnpackedBillsList.VchNoList
        Else
            ClubbedVch = FrmUnpackedBillsList.VchCodeList & ",'" & rstPackingSlipList.Fields("Code").Value & "'"
            Text18.Text = Text5.Text & "+" & FrmUnpackedBillsList.VchNoList
            Call LoadItemList(rstPackingSlipList.Fields("Code").Value)
        End If
    Else 'Nothing Selected
        If ListType = "G" Then
            ClubbedGRVch = "": OldClubbedGRVch = ClubbedGRVch: Text20.Text = ""
        Else
            ClubbedVch = "'" & rstPackingSlipList.Fields("Code").Value & "'"
            Text18.Text = Text5.Text
            Call LoadItemList(rstPackingSlipList.Fields("Code").Value)
        End If
    End If
    CloseForm FrmUnpackedBillsList
End Sub
Private Sub SMSConsignmentDetails() 'Done
    Dim WinHttpReq As Object
    Dim Response As String, URL As String, VchAmt As Double, AccountName As String, VchQty As Long
    Set WinHttpReq = CreateObject("Msxml2.XMLHTTP")
    With rstPackingSlipChild
        If .State = adStateOpen Then .Close
        .Open "SELECT Mobile FROM AccountMaster WHERE Code=(SELECT Party FROM JobWorkBVParent WHERE Code='" & rstPackingSlipList.Fields("Code").Value & "')", cnPackingSlip, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        If .RecordCount = 0 Then Call MsgBox("Failed to SMS Consignment details !!!", vbInformation, App.Title): Exit Sub
        If CheckEmpty(Trim(CheckNull(.Fields("Mobile").Value)), False) Then Call MsgBox("Mobile No. is blank !!!", vbInformation, App.Title): Exit Sub
        Response = .Fields("Mobile").Value
        If .State = adStateOpen Then .Close
        .Open "SELECT Amount As VchAmt,LTRIM(M.Name) As AccountName,(SELECT ABS(SUM(Quantity)) FROM JobWorkBVChild WHERE Code=T.Code) As VchQty FROM JobWorkBVParent T INNER JOIN AccountMaster M ON T.Party=M.Code WHERE T.Code='" & rstPackingSlipList.Fields("Code").Value & "'", cnPackingSlip, adOpenKeyset, adLockReadOnly
        AccountName = .Fields("AccountName").Value: VchQty = Val(.Fields("VchQty").Value): VchAmt = Val(.Fields("VchAmt").Value)
        If .State = adStateOpen Then .Close
        .Open "SELECT URL,UserIDField,UserIDValue,PasswordField,PasswordValue,SenderIDField,SenderIDValue,MobileNoField,MessageField FROM SMSConfig ", cnPackingSlip, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        If .RecordCount = 0 Then Call MsgBox("No SMS API found !!!", vbInformation, App.Title): Exit Sub
        URL = .Fields("URL").Value
        URL = Replace(URL, .Fields("UserIDField").Value, .Fields("UserIDValue").Value)
        URL = Replace(URL, .Fields("PasswordField").Value, .Fields("PasswordValue").Value)
        URL = Replace(URL, .Fields("SenderIDField").Value, .Fields("SenderIDValue").Value)
        URL = Replace(URL, .Fields("MobileNoField").Value, Response)
        Response = "Name:-" & URLEncode(Trim(AccountName)) & " Bill No:-" & Trim(Text5.Text) & " Qty:-" & Trim(Format(VchQty, "#0")) & " Amt:-Rs." & Trim(Format(VchAmt, "#0.00"))
        Response = Response & " Tpt:-" & Trim(Text2.Text) & " Stn:-" & Trim(Text8.Text) & " GR No:-" & Trim(Text11.Text) & " Dt:-" & MhDateInput5.Text
        URL = Replace(URL, .Fields("MessageField").Value, Response)
    End With
    With WinHttpReq
        .Open "GET", URL, False
        .Send
        Response = .responseText
    End With
    If InStr(LCase(Response), "true") > 0 Then Call MsgBox("Successfully SMSed the Consignment details !!!", vbInformation, App.Title) Else Call MsgBox(Response, vbInformation, App.Title)
End Sub
Private Function URLEncode(ByVal URL As String) As String 'Done
    Dim cchEscaped As Long
    Dim HRESULT As Long
    cchEscaped = Len(URL) * 1.5
    URLEncode = String(cchEscaped, 0)
    HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
    If HRESULT = E_POINTER Then URLEncode = String$(cchEscaped, 0): HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
    If HRESULT <> S_OK Then DisplayError ("System error")
    URLEncode = Left$(URLEncode, cchEscaped): URLEncode = Replace$(URLEncode, "+", "%2B"): URLEncode = Replace$(URLEncode, " ", "+")
End Function
'Vch Series
