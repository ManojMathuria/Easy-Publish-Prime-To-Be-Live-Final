VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockJournalVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Journal-Finished Goods"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15675
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
   ScaleHeight     =   8880
   ScaleWidth      =   15675
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   8865
      Left            =   15
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   15650
      _Version        =   65536
      _ExtentX        =   27605
      _ExtentY        =   15637
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
      Picture         =   "StockJournalVoucher.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8645
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   15420
         _ExtentX        =   27199
         _ExtentY        =   15240
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
         TabPicture(0)   =   "StockJournalVoucher.frx":001C
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
         TabPicture(1)   =   "StockJournalVoucher.frx":0038
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
            Left            =   605
            TabIndex        =   12
            Top             =   8115
            Width           =   10155
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7590
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   450
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   13388
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
               Caption         =   "  Vch No."
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
               DataField       =   "VchSeriesName"
               Caption         =   "Vch Series"
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
               DataField       =   "Date"
               Caption         =   "Vch Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd-MM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "igMaterialCentreName"
               Caption         =   "Mat Centre-Item Generated"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "icMaterialCentreName"
               Caption         =   "Mat Centre-Item Consumed"
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
                  ColumnWidth     =   2085.166
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column02 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column03 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   5114.835
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   5084.788
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7995
            Left            =   -74880
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   480
            Width           =   15195
            _Version        =   65536
            _ExtentX        =   26802
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
            Picture         =   "StockJournalVoucher.frx":0054
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
               Left            =   1800
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   0
               Top             =   105
               Width           =   1770
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
               Left            =   1800
               MaxLength       =   40
               TabIndex        =   4
               Top             =   945
               Width           =   13290
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   2955
               Left            =   120
               TabIndex        =   6
               Top             =   1785
               Width           =   14970
               _Version        =   524288
               _ExtentX        =   26405
               _ExtentY        =   5212
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
               MaxCols         =   6
               MaxRows         =   2000
               SpreadDesigner  =   "StockJournalVoucher.frx":0070
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7320
               MaxLength       =   25
               TabIndex        =   1
               Top             =   105
               Width           =   2370
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
               MaxLength       =   40
               TabIndex        =   5
               Top             =   1265
               Width           =   13290
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
               MaxLength       =   60
               TabIndex        =   3
               Top             =   630
               Width           =   13290
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   6480
               TabIndex        =   15
               Top             =   105
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
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
               Caption         =   " Vch. No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "StockJournalVoucher.frx":07E7
               Picture         =   "StockJournalVoucher.frx":0803
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   11955
               TabIndex        =   16
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
               Caption         =   " Vch. Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "StockJournalVoucher.frx":081F
               Picture         =   "StockJournalVoucher.frx":083B
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   17
               Top             =   630
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
               Caption         =   " MC-Item Generated"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "StockJournalVoucher.frx":0857
               Picture         =   "StockJournalVoucher.frx":0873
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   18
               Top             =   945
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
               Caption         =   " MC-Item Consumed"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "StockJournalVoucher.frx":088F
               Picture         =   "StockJournalVoucher.frx":08AB
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   13515
               TabIndex        =   2
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "StockJournalVoucher.frx":08C7
               Caption         =   "StockJournalVoucher.frx":09DF
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "StockJournalVoucher.frx":0A4B
               Keys            =   "StockJournalVoucher.frx":0A69
               Spin            =   "StockJournalVoucher.frx":0AC7
               AlignHorizontal =   0
               AlignVertical   =   0
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
            Begin VB.TextBox Text9 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3840
               TabIndex        =   19
               Top             =   2520
               Width           =   2535
            End
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   2955
               Left            =   120
               TabIndex        =   7
               Top             =   4935
               Width           =   14970
               _Version        =   524288
               _ExtentX        =   26405
               _ExtentY        =   5212
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
               MaxCols         =   6
               MaxRows         =   2000
               SpreadDesigner  =   "StockJournalVoucher.frx":0AEF
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   20
               Top             =   1265
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
               Picture         =   "StockJournalVoucher.frx":126A
               Picture         =   "StockJournalVoucher.frx":1286
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   120
               TabIndex        =   22
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
               Caption         =   " Vch Series"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "StockJournalVoucher.frx":12A2
               Picture         =   "StockJournalVoucher.frx":12BE
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   15200
               Y1              =   4830
               Y2              =   4830
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   15200
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   15200
               Y1              =   1680
               Y2              =   1680
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   10740
            TabIndex        =   21
            Top             =   8115
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
            Picture         =   "StockJournalVoucher.frx":12DA
            Picture         =   "StockJournalVoucher.frx":12F6
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
            Top             =   8115
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
      Width           =   15675
      _ExtentX        =   27649
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
Attribute VB_Name = "frmStockJournalVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchCode As String  'Vch to Modify
Public VchType As String 'JR-Stock Journal-Finished Goods
Dim cnStockJournalVoucher As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstMaterialCentreList As New ADODB.Recordset, rstItemList As New ADODB.Recordset
Dim rstStockJournalVoucherList As New ADODB.Recordset, rstStockJournalVoucherParent As New ADODB.Recordset, rstStockJournalVoucherChild As New ADODB.Recordset, rstVchSeriesList As New ADODB.Recordset
Dim igMaterialCentreCode As String, icMaterialCentreCode As String, MaterialCentreCode As String, VchPrefix As String, VchNumbering As String, VchSeriesCode As String, oVchSeriesCode As String, oVchNo As String, AutoVchNo As String
Dim SortOrder, PrevStr, dblBookMark As Double, blnRecordExist As Boolean, EditMode As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    'If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
   ' Me.Left = (MdiMainMenu.ScaleWidth - Me.Width) \ 2
    Me.Top = 1200
    WheelHook DataGrid1
    BusySystemIndicator True
    VchPrefix = "2010" '10-Stock affected
    cnStockJournalVoucher.CursorLocation = adUseClient: cnStockJournalVoucher.Open cnDatabase.ConnectionString
    rstStockJournalVoucherParent.CursorLocation = adUseClient
    LoadMasterList
    With rstStockJournalVoucherList
        .Open "SELECT T.Code,T.Name,V.Name As VchSeriesName,Date,RIGHT(T.Type,2) As Type,M1.Name As igMaterialCentreName,M2.Name As icMaterialCentreName FROM ((JobworkBVParent T INNER JOIN AccountMaster M1 ON T.Party=M1.Code) INNER JOIN AccountMaster M2 ON T.MaterialCentre=M2.Code) INNER JOIN VchSeriesMaster V ON T.VchSeries=V.Code WHERE RIGHT(Type,2)='" & VchType & "' AND T.FYCode='" & FYCode & "' ORDER BY T.Name", cnStockJournalVoucher, adOpenKeyset, adLockPessimistic
        .Filter = adFilterNone
        If .RecordCount > 0 Then
            .MoveLast
            If Not CheckEmpty(VchCode, False) Then .MoveFirst: .Find "[Code]='" & VchCode & "'"
        End If
        Set DataGrid1.DataSource = rstStockJournalVoucherList
        BusySystemIndicator False
        SSTab1.Tab = 0
    If FrmStockLedger.dSortBy = True Then
    SortOrder = "Code"
    Else
    SortOrder = "AutoVchNo"
    End If
        If Not (.EOF Or .BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        .ActiveConnection = Nothing
    End With
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    With MdiMainMenu
        .mnuStockJournalFinishedGoods.Enabled = False
    End With
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    With Toolbar1.Buttons
        If Shift = 0 And KeyCode = vbKeyEscape Then
            If SSTab1.Tab = 0 Then  'List
                Unload Me
            Else
                If .Item(1).Enabled Then    'Add Button Enabled
                    SSTab1.Tab = 0
                Else
                    If Not EditMode Then
                        If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                            Me.ActiveControl.SetFocus
                        Else
                            Toolbar1_ButtonClick .Item(5)
                        End If
                    End If
                End If
                If Not EditMode Then KeyCode = 0
            End If
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And .Item(1).Enabled Then
            Toolbar1_ButtonClick .Item(1)
            KeyCode = 0
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And .Item(2).Enabled Then
            Toolbar1_ButtonClick .Item(2)
            KeyCode = 0
        ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF8)) And .Item(3).Enabled Then
            Toolbar1_ButtonClick .Item(3)
            KeyCode = 0
        ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2)) And .Item(4).Enabled Then 'Save
            If Not EditMode Then Toolbar1_ButtonClick .Item(4)
            KeyCode = 0
        ElseIf Shift = 0 And KeyCode = vbKeyF5 And .Item(6).Enabled Then
            Toolbar1_ButtonClick .Item(6)
            KeyCode = 0
        ElseIf Shift = 0 And KeyCode = vbKeyF12 And .Item(1).Enabled Then 'Duplicate
            If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then DuplicateRecord
            KeyCode = 0
        ElseIf Shift = vbAltMask And KeyCode = vbKeyP And .Item(1).Enabled Then
            Toolbar1_ButtonClick .Item(9)
            KeyCode = 0
        ElseIf Shift = vbAltMask And KeyCode = vbKeyV And .Item(1).Enabled Then
            Toolbar1_ButtonClick .Item(10)
            KeyCode = 0
        ElseIf Shift = vbAltMask And KeyCode = vbKeyM And .Item(1).Enabled Then
            Toolbar1_ButtonClick .Item(11)
            KeyCode = 0
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF And .Item(1).Enabled Then
            Toolbar1_ButtonClick .Item(13)
            KeyCode = 0
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP And .Item(1).Enabled Then
            Toolbar1_ButtonClick .Item(14)
            KeyCode = 0
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyN And .Item(1).Enabled Then
            Toolbar1_ButtonClick .Item(15)
            KeyCode = 0
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL And .Item(1).Enabled Then
            Toolbar1_ButtonClick .Item(16)
            KeyCode = 0
        ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
            If .Item(1).Enabled Then 'Add Button Enabled
                SSTab1.Tab = 1: SSTab1.SetFocus
            Else
               If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then Sendkeys "{TAB}"
            End If
            If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then KeyCode = 0
        End If
    End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstStockJournalVoucherList)
    Call CloseRecordset(rstStockJournalVoucherParent)
    Call CloseRecordset(rstStockJournalVoucherChild)
    Call CloseRecordset(rstMaterialCentreList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstVchSeriesList)
    Call CloseConnection(cnStockJournalVoucher)
    Call CloseRecordset(rstCompanyMaster)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuStockJournalFinishedGoods.Enabled = True
End Sub
Private Sub Text1_Change()
On Error Resume Next
    With rstStockJournalVoucherList
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
    With rstStockJournalVoucherList
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
    If Toolbar1.Buttons.Item(1).Enabled Then 'Add Button Enabled
        If SSTab1.Tab = 1 Then
            ViewRecord
        Else
            If Not (rstStockJournalVoucherList.EOF Or rstStockJournalVoucherList.BOF) Then
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
        Text8.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean, UpdateFlag As Integer, CellVal01 As Variant, CellVal02 As Variant, i As Integer
    With rstStockJournalVoucherList
        If Button.Index = 1 Then
            If rstStockJournalVoucherParent.State = adStateOpen Then rstStockJournalVoucherParent.Close
            rstStockJournalVoucherParent.Open "SELECT * FROM JobworkBVParent WHERE Code=''", cnStockJournalVoucher, adOpenKeyset, adLockOptimistic
            ClearFields
            If AddRecord(rstStockJournalVoucherParent) Then
                Text2.Text = GenerateCode(cnStockJournalVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
                Call SetButtons(False)
                SSTab1.Tab = 1
                Text8.SetFocus
                blnRecordExist = False
                cnStockJournalVoucher.BeginTrans
            End If
        ElseIf Button.Index = 2 Then
            If .RecordCount = 0 Then Exit Sub
            SSTab1.Tab = 1
            EditRecord
        ElseIf Button.Index = 3 Then
            If .RecordCount = 0 Then Exit Sub
            If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
            SSTab1.Tab = 1
            If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
                On Error Resume Next
                MdiMainMenu.MousePointer = vbHourglass
                cnStockJournalVoucher.BeginTrans
                cnStockJournalVoucher.Execute "DELETE FROM JobworkBVParent WHERE Code='" & .Fields("Code").Value & "'"
                MdiMainMenu.MousePointer = vbNormal
                If Err.Number = 0 Then
                    .Delete
                    .MoveNext
                    If .RecordCount > 0 And .EOF Then .MoveLast
                    cnStockJournalVoucher.CommitTrans
                    ShowProgressInStatusBar True
                    Timer1.Enabled = True
                    Text1.Text = ""
                    .Filter = adFilterNone
                Else
                    DisplayError (Err.Description)
                    cnStockJournalVoucher.RollbackTrans
                End If
                On Error GoTo 0
            End If
            SetButtons (True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
            HiLiteRecord = True
        ElseIf Button.Index = 4 Then
            If CheckMandatoryFields Then Exit Sub
            SaveFields
            UpdateFlag = 0
            If UpdateRecord(rstStockJournalVoucherParent) Then
                If UpdateItemList("D", 0) Then
                    UpdateFlag = 1
                   With fpSpread1
                       For i = 1 To .DataRowCnt
                           .SetActiveCell 2, i
                           .GetText 2, i, CellVal01 'Quantity
                           .GetText 3, i, CellVal02 'Item Code
                           If Val(CellVal01) <> 0 And Not CheckEmpty(CellVal02, False) Then If Not UpdateItemList("I1", i) Then UpdateFlag = 0: Exit For
                       Next
                   End With
                End If
                If UpdateFlag = 1 Then
                    With fpSpread2
                        For i = 1 To .DataRowCnt
                           .SetActiveCell 2, i
                           .GetText 2, i, CellVal01 'Quantity
                           .GetText 3, i, CellVal02 'Item Code
                           If Val(CellVal01) <> 0 And Not CheckEmpty(CellVal02, False) Then If Not UpdateItemList("I2", i) Then UpdateFlag = 0: Exit For
                        Next
                    End With
                End If
            End If
            If UpdateFlag Then
                AddToList
                cnStockJournalVoucher.CommitTrans
                If rstStockJournalVoucherParent.State = adStateOpen Then rstStockJournalVoucherParent.Close
                rstStockJournalVoucherParent.CursorLocation = adUseClient
                Call SetButtons(True)
                ShowProgressInStatusBar True
                Timer1.Enabled = True
                Call MsgBox("Record updated !!!", vbInformation, App.Title)
                SSTab1.Tab = 0
            Else
                DisplayError ("Failed to save the record")
                Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            End If
        ElseIf Button.Index = 5 Then
            If CancelRecordUpdate(rstStockJournalVoucherParent) Then
                cnStockJournalVoucher.RollbackTrans
                If rstStockJournalVoucherParent.State = adStateOpen Then rstStockJournalVoucherParent.Close
                rstStockJournalVoucherParent.CursorLocation = adUseClient
                Call SetButtons(True)
                SetButtonsForNoRecord
                SSTab1.Tab = 0
            End If
        ElseIf Button.Index = 6 Then
            SSTab1.Tab = 0
            Set DataGrid1.DataSource = Nothing
            .Filter = adFilterNone
            RefreshData rstStockJournalVoucherList
            Set DataGrid1.DataSource = rstStockJournalVoucherList
            If .RecordCount > 0 Then .MoveLast
            LoadMasterList
            HiLiteRecord = True
        ElseIf Button.Index = 7 Then
            SSTab1.Tab = 0
            With FrmFilter
                .Combo1.AddItem "Material Centre-Item Generated", 0
                .Combo1.AddItem "Material Centre-Item Consumed", 1
                .Combo1.ListIndex = 0
                Set .srcForm = Me
                .Show vbModal
            End With
            HiLiteRecord = True
        ElseIf Button.Index = 9 Then
            If .RecordCount = 0 Then Exit Sub
            Call PrintStockJournalVoucher(.Fields("Code").Value, .Fields("Type").Value, "P")
            HiLiteRecord = True
        ElseIf Button.Index = 10 Then
            If .RecordCount = 0 Then Exit Sub
            Call PrintStockJournalVoucher(.Fields("Code").Value, .Fields("Type").Value, "S")
            HiLiteRecord = True
        ElseIf Button.Index = 13 Then
            If .RecordCount > 0 Then .MoveFirst
            HiLiteRecord = True
            ViewRecord
        ElseIf Button.Index = 14 Then
            If .RecordCount > 0 Then
                .MovePrevious
                If .BOF Then .MoveNext
            End If
            HiLiteRecord = True
            ViewRecord
        ElseIf Button.Index = 15 Then
            If .RecordCount > 0 Then
                .MoveNext
                If .EOF Then .MovePrevious
            End If
            HiLiteRecord = True
            ViewRecord
        ElseIf Button.Index = 16 Then
            If .RecordCount > 0 Then .MoveLast
            HiLiteRecord = True
            ViewRecord
        ElseIf Button.Index = 18 Then
            Unload Me
            HiLiteRecord = False
        End If
        If HiLiteRecord Then
            If Not (.EOF Or .BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            Text1.SetFocus
        End If
    End With
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstStockJournalVoucherList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstStockJournalVoucherList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstStockJournalVoucherList.EOF Or rstStockJournalVoucherList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub SetButtons(bVal As Boolean)
    With Toolbar1.Buttons
        .Item(1).Enabled = bVal
        .Item(2).Enabled = bVal
        .Item(3).Enabled = bVal
        .Item(4).Enabled = Not bVal
        .Item(5).Enabled = Not bVal
        .Item(6).Enabled = bVal
        .Item(7).Enabled = bVal
        .Item(9).Enabled = bVal
        .Item(10).Enabled = bVal
        .Item(11).Enabled = bVal
        .Item(13).Enabled = bVal
        .Item(14).Enabled = bVal
        .Item(15).Enabled = bVal
        .Item(16).Enabled = bVal
        .Item(18).Enabled = bVal
    End With
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstStockJournalVoucherList.RecordCount = 0 Then
        With Toolbar1.Buttons
            .Item(2).Enabled = False
            .Item(3).Enabled = False
            .Item(9).Enabled = False
            .Item(10).Enabled = False
            .Item(11).Enabled = False
            .Item(13).Enabled = False
            .Item(14).Enabled = False
            .Item(15).Enabled = False
            .Item(16).Enabled = False
        End With
    End If
End Sub
Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text8.Text)
        If rstVchSeriesList.RecordCount = 0 Then DisplayError ("No Record in Voucher Series Master"): Text8.SetFocus: Exit Sub Else rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": VchSeriesCode = ""
        Call LoadSelectionList(rstVchSeriesList, "List of Voucher Series...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text8, VchSeriesCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(VchSeriesCode) <> "" Then Sendkeys "{TAB}" Else Text8.Text = ""
    End If
End Sub
Private Sub Text8_Validate(Cancel As Boolean)
    If CheckEmpty(Text8.Text, False) Then
        Cancel = True
    Else
        rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        VchNumbering = rstVchSeriesList.Fields("VchNumbering").Value
        If VchNumbering = "A" Then Text2.Locked = True Else Text2.Locked = False
        If Not blnRecordExist Then 'Vch-New
            If VchNumbering = "A" Then
                AutoVchNo = GenerateCode(cnStockJournalVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
            End If
        Else 'Vch-Old
            If VchSeriesCode = oVchSeriesCode Then
                Text2.Text = oVchNo
            Else
                If VchNumbering = "A" Then
                    AutoVchNo = GenerateCode(cnStockJournalVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                    Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                End If
            End If
        End If
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean) 'Vch No.
    With rstStockJournalVoucherParent
        If .EOF Or .BOF Then Exit Sub
        If CheckEmpty(Text2, True) Then
            Cancel = True
        ElseIf CheckDuplicate(cnStockJournalVoucher, "JobworkBVParent", "Code", "[Name]+RIGHT(Type,2)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, .Fields("Code").Value, False, FYCode) Then
            Cancel = True
        End If
    End With
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)    'Vch Date
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = "*99999"
        FrmAccountMaster.MasterCode = igMaterialCentreCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        igMaterialCentreCode = slCode: Text3.Text = slName
        If Not CheckEmpty(igMaterialCentreCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        igMaterialCentreCode = "": Text3.Text = ""
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If CheckEmpty(Text3.Text, False) Then Cancel = True
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = "*99999"
        FrmAccountMaster.MasterCode = icMaterialCentreCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        icMaterialCentreCode = slCode: Text5.Text = slName
        If Not CheckEmpty(icMaterialCentreCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        icMaterialCentreCode = "": Text5.Text = ""
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    If CheckEmpty(Text5.Text, False) Then Cancel = True
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstStockJournalVoucherList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    With rstStockJournalVoucherParent
        If .State = adStateOpen Then .Close
        .Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstStockJournalVoucherList.Fields("Code").Value) & "'", cnStockJournalVoucher, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End With
End Sub
Private Sub ClearFields()
    Text8.Text = "" 'Vch Series
    Text2.Text = "" 'Vch No.
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    Text3.Text = "" 'Mat Centre-Item Generated
    Text5.Text = "" 'Mat Centre-Item Consumed
    Text4.Text = "" 'Remarks
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    fpSpread2.ClearRange 1, 1, fpSpread2.MaxCols, fpSpread2.MaxRows, True: fpSpread2.SetActiveCell 1, 1
    igMaterialCentreCode = "": icMaterialCentreCode = "":  VchSeriesCode = "": oVchSeriesCode = "": oVchNo = "": AutoVchNo = ""
End Sub
Private Sub LoadFields()
    With rstStockJournalVoucherParent
        If .EOF Or .BOF Then Exit Sub
        VchSeriesCode = .Fields("VchSeries").Value: oVchSeriesCode = VchSeriesCode
        If rstVchSeriesList.RecordCount > 0 Then rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        If Not rstVchSeriesList.EOF Then Text8.Text = rstVchSeriesList.Fields("Col0").Value
        AutoVchNo = Trim(.Fields("AutoVchNo").Value)
        Text2.Text = .Fields("Name").Value
        oVchNo = Trim(Text2.Text)
        MhDateInput1.Text = Format(.Fields("Date").Value, "dd-MM-yyyy")
        igMaterialCentreCode = .Fields("Party").Value
        If rstMaterialCentreList.RecordCount > 0 Then rstMaterialCentreList.MoveFirst
        rstMaterialCentreList.Find "[Code] = '" & igMaterialCentreCode & "'"
        If Not rstMaterialCentreList.EOF Then Text3.Text = rstMaterialCentreList.Fields("Col0").Value
        icMaterialCentreCode = .Fields("MaterialCentre").Value
        If rstMaterialCentreList.RecordCount > 0 Then rstMaterialCentreList.MoveFirst
        rstMaterialCentreList.Find "[Code] = '" & icMaterialCentreCode & "'"
        If Not rstMaterialCentreList.EOF Then Text5.Text = rstMaterialCentreList.Fields("Col0").Value
        Text4.Text = .Fields("Remarks").Value
        Call LoadItemList(.Fields("Code").Value)
    End With
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    With rstStockJournalVoucherParent
        If .RecordCount = 0 Then Exit Sub
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseServer
        .Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstStockJournalVoucherList.Fields("Code").Value) & "'", cnStockJournalVoucher, adOpenKeyset, adLockPessimistic
        MdiMainMenu.MousePointer = vbHourglass
        .Fields("RecordStatus") = "N"
    End With
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text8.SetFocus
    blnRecordExist = True
    cnStockJournalVoucher.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    With rstStockJournalVoucherParent
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnStockJournalVoucher, "SELECT MAX(Code) FROM JobworkBVParent", 6, "0")
            .Fields("CreatedBy").Value = UserCode
            .Fields("CreatedOn").Value = Now()
            .Fields("Recordstatus").Value = "N"
        Else
            .Fields("ModifiedBy").Value = UserCode
            .Fields("ModifiedOn").Value = Now()
            .Fields("Recordstatus").Value = "M"
        End If
        .Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
        .Fields("VchSeries").Value = VchSeriesCode
        .Fields("AutoVchNo").Value = Pad(Trim(AutoVchNo), Space(1), 10, "L")
        .Fields("Date").Value = GetDate(MhDateInput1.Text)
        .Fields("Party").Value = igMaterialCentreCode
        .Fields("Consignee").Value = igMaterialCentreCode
        .Fields("MaterialCentre").Value = icMaterialCentreCode
        .Fields("Tax").Value = Null
        .Fields("Remarks").Value = Trim(Text4.Text)
        .Fields("Type").Value = VchPrefix & VchType
        .Fields("FYCode").Value = FYCode
        .Fields("RecordStatus").Value = "N"
    End With
End Sub
Private Sub AddToList()
    On Error Resume Next
    With rstStockJournalVoucherList
        .MoveFirst
        .Find "[Code] = '" & rstStockJournalVoucherParent.Fields("Code").Value & "'"
        If .EOF Then .AddNew
        .Fields("Code").Value = rstStockJournalVoucherParent.Fields("Code").Value
        .Fields("Name").Value = Pad(rstStockJournalVoucherParent.Fields("Name").Value, Space(1), 10, "L")
        .Fields("Date").Value = rstStockJournalVoucherParent.Fields("Date").Value
        .Fields("VchSeriesName").Value = Text8.Text
        .Fields("igMaterialCentreName").Value = Trim(Text3.Text)
        .Fields("icMaterialCentreName").Value = Trim(Text5.Text)
        .Fields("Type").Value = Right(rstStockJournalVoucherParent.Fields("Type").Value, 2)
        .Update
        .Sort = SortOrder & " Asc"
        .Find "[Code] = '" & rstStockJournalVoucherParent.Fields("Code").Value & "'"
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    Dim i As Integer, x As Integer, Item01 As Variant, Item02 As Variant
    If CheckEmpty(Text8.Text, False) Then
        Text8.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text2.Text, False) Then
        DisplayError ("Voucher No. cannot be blank"): Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnStockJournalVoucher, "JobworkBVParent", "Code", "[Name]+RIGHT(Type,2)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, rstStockJournalVoucherParent.Fields("Code").Value, False, FYCode) Then
        Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text3.Text, False) Then ''Mat Centre-Item Generated
        Text3.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text5.Text, False) Then ''Mat Centre-Item Consumed
        Text5.SetFocus:   CheckMandatoryFields = True: Exit Function
    Else
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.SetActiveCell 1, i
            fpSpread1.GetText 5, i, Item01
            For x = 1 To fpSpread2.DataRowCnt
                fpSpread2.SetActiveCell 1, x
                fpSpread2.GetText 5, x, Item02
                If Item02 = Item01 Then CheckMandatoryFields = True: Exit For
            Next
            If CheckMandatoryFields Then DisplayError "Same item cann't be generated (row #" & Trim(Str(i)) & ") and consumed (row #" & Trim(Str(x)) & ") simultaneously": Exit For
        Next
    End If
End Function
Private Sub LoadItemList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    With rstStockJournalVoucherChild
        If .State = adStateOpen Then .Close
        .Open "SELECT I.Code As ItemCode,I.Name As ItemName,T.HSNCode,ABS(T.Quantity) As Quantity,T.Rate,T.Amount,SrNo FROM JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code WHERE T.Code='" & strOrderCode & "' AND Quantity>0 ORDER BY SrNo", cnStockJournalVoucher, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        If .RecordCount > 0 Then .MoveFirst
        i = 0
        Do While Not .EOF
            i = i + 1
            fpSpread1.SetText 1, i, .Fields("ItemName").Value
            fpSpread1.SetText 2, i, Val(.Fields("Quantity").Value)
            fpSpread1.SetText 3, i, Val(.Fields("Rate").Value)
            fpSpread1.SetText 4, i, Val(.Fields("Amount").Value)
            fpSpread1.SetText 5, i, .Fields("ItemCode").Value
            fpSpread1.SetText 6, i, .Fields("HSNCode").Value
            .MoveNext
        Loop
        i = 0
        If .State = adStateOpen Then .Close
        .Open "SELECT I.Code As ItemCode,I.Name As ItemName,T.HSNCode,ABS(T.Quantity) As Quantity,T.Rate,T.Amount,SrNo FROM JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code WHERE T.Code='" & strOrderCode & "' AND Quantity<0 ORDER BY SrNo", cnStockJournalVoucher, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        If .RecordCount > 0 Then .MoveFirst
        i = 0
        Do While Not .EOF
            i = i + 1
            fpSpread2.SetText 1, i, .Fields("ItemName").Value
            fpSpread2.SetText 2, i, Val(.Fields("Quantity").Value)
            fpSpread2.SetText 3, i, Val(.Fields("Rate").Value)
            fpSpread2.SetText 4, i, Val(.Fields("Amount").Value)
            fpSpread2.SetText 5, i, .Fields("ItemCode").Value
            fpSpread2.SetText 6, i, .Fields("HSNCode").Value
            .MoveNext
        Loop
    End With
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Item List")
End Sub
Private Function UpdateItemList(ByVal ActionType As String, ByVal SrNo As Integer) As Boolean
    Dim CellVal(1 To 5) As Variant
    On Error GoTo ErrorHandler
    UpdateItemList = True
    If ActionType = "D" Then
        If Not blnRecordExist Then Exit Function
        cnStockJournalVoucher.Execute "DELETE FROM JobworkBVChild WHERE Code='" & rstStockJournalVoucherParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I1" Then
        With fpSpread1
            .GetText 2, .ActiveRow, CellVal(1)  'Quantity
            .GetText 3, .ActiveRow, CellVal(2)  'Rate
            .GetText 4, .ActiveRow, CellVal(3)  'Amount
            .GetText 5, .ActiveRow, CellVal(4)  'Item Code
            .GetText 6, .ActiveRow, CellVal(5)  'HSN Code
        End With
        cnStockJournalVoucher.Execute "INSERT INTO JobworkBVChild VALUES ('" & rstStockJournalVoucherParent.Fields("Code").Value & "','','" & VchPrefix & "FI" & "','" & CellVal(4) & "','" & CellVal(5) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ",Null," & SrNo & ",'','','','','',0,'')"
    ElseIf ActionType = "I2" Then
        With fpSpread2
            .GetText 2, .ActiveRow, CellVal(1)  'Quantity
            .GetText 3, .ActiveRow, CellVal(2)  'Rate
            .GetText 4, .ActiveRow, CellVal(3)  'Amount
            .GetText 5, .ActiveRow, CellVal(4)  'Item Code
            .GetText 6, .ActiveRow, CellVal(5)  'HSN Code
        End With
        cnStockJournalVoucher.Execute "INSERT INTO JobworkBVChild VALUES ('" & rstStockJournalVoucherParent.Fields("Code").Value & "','','" & VchPrefix & "FI" & "','" & CellVal(4) & "','" & CellVal(5) & "'," & 0 - Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ",Null," & SrNo & ",'','','','','',0,'')"
    End If
    Exit Function
ErrorHandler:
    UpdateItemList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Material Centre-Item Generated" Then
        rstStockJournalVoucherList.Filter = "[igMaterialCentreName] Like '%" & SrchText & "%'"
    ElseIf SrchFor = "Material Centre-Item Consumed" Then
        rstStockJournalVoucherList.Filter = "[icMaterialCentreName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Item As Variant, i As Integer, x As Integer, cVal(1 To 6) As Variant
    With fpSpread1
        If Shift = 0 And KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf KeyCode = vbKeyF3 Then
            If .ActiveCol = 1 Then
                .GetText 5, .ActiveRow, Item
                On Error Resume Next
                FrmBookMaster.SL = True
                FrmBookMaster.ItemType = "F"
                FrmBookMaster.MasterCode = Item
                Load FrmBookMaster
                If Err.Number <> 364 Then FrmBookMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText 5, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then
                    rstItemList.MoveFirst: rstItemList.Find "[Code] ='" & slCode & "'"
                    .GetText 3, .ActiveRow, Item 'Price
                    If Val(Item) = 0 Then
                        .SetText 3, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                    ElseIf Val(Item) <> Val(rstItemList.Fields("Price").Value) Then
                        If MsgBox("Variation in Current (" & Format(Item, "#0.00") & ") and Master (" & Format(rstItemList.Fields("Price").Value, "#0.00") & ") Rate ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 3, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                    End If
                    .SetText 6, .ActiveRow, rstItemList.Fields("HSNCode").Value
                    LoadMasterList
                    .SetFocus
                    Sendkeys "{ENTER}"
                End If
            End If
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then
                MaterialCentreCode = igMaterialCentreCode
                LoadMasterList True
                With FrmItemSearchList
                    Set .rstItemSearchList = rstItemList
                    Load FrmItemSearchList
                    .Show vbModal
                    If .LoadItems Then
                        For i = 1 To .fpSpread1.DataRowCnt
                            .fpSpread1.GetText 1, i, cVal(1) 'Item
                            .fpSpread1.GetText 3, i, cVal(2) 'Quantity
                            .fpSpread1.GetText 4, i, cVal(3) 'Price
                            .fpSpread1.GetText 5, i, cVal(4) 'Item Code
                            .fpSpread1.GetText 6, i, cVal(5) 'HSN Code
                            .fpSpread1.GetText 7, i, cVal(6) 'HSN
                            If Val(cVal(2)) > 0 Then
                                x = fpSpread1.DataRowCnt + 1
                                fpSpread1.SetText 1, x, cVal(1)
                                fpSpread1.SetText 2, x, Val(cVal(2))
                                fpSpread1.SetText 3, x, Val(cVal(3))
                                fpSpread1.SetText 4, x, Val(cVal(2)) * Val(cVal(3))
                                fpSpread1.SetText 5, x, cVal(4)
                                fpSpread1.SetText 6, x, cVal(5)
                            End If
                        Next
                    End If
                End With
                Call CloseForm(FrmItemSearchList)
                .SetFocus
            End If
        End If
    End With
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Item As Variant, Qty As Variant, Rate As Variant
    With fpSpread1
        If Col = 2 Or Col = 3 Then 'Qty & Rate
            .GetText 5, Row, Item
            .GetText 2, Row, Qty
            .GetText 3, Row, Rate
            If Not CheckEmpty(Item, False) Then .SetText 4, Row, Qty * Rate Else .SetText 2, Row, "": .SetText 3, Row, "": .SetText 4, Row, ""
        End If
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub fpSpread2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Item As Variant, i As Integer, x As Integer, cVal(1 To 6) As Variant
    With fpSpread2
        If Shift = 0 And KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf KeyCode = vbKeyF3 Then
            If .ActiveCol = 1 Then
                .GetText 5, .ActiveRow, Item
                On Error Resume Next
                FrmBookMaster.SL = True
                FrmBookMaster.ItemType = "F"
                FrmBookMaster.MasterCode = Item
                Load FrmBookMaster
                If Err.Number <> 364 Then FrmBookMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText 5, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then
                    rstItemList.MoveFirst: rstItemList.Find "[Code] ='" & slCode & "'"
                    .GetText 3, .ActiveRow, Item 'Price
                    If Val(Item) = 0 Then
                        .SetText 3, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                    ElseIf Val(Item) <> Val(rstItemList.Fields("Price").Value) Then
                        If MsgBox("Variation in Current (" & Format(Item, "#0.00") & ") and Master (" & Format(rstItemList.Fields("Price").Value, "#0.00") & ") Rate ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 3, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                    End If
                    .SetText 6, .ActiveRow, rstItemList.Fields("HSNCode").Value
                    LoadMasterList
                    .SetFocus
                    Sendkeys "{ENTER}"
                End If
            End If
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then
                MaterialCentreCode = icMaterialCentreCode
                LoadMasterList True
                With FrmItemSearchList
                    Set .rstItemSearchList = rstItemList
                    Load FrmItemSearchList
                    .Show vbModal
                    If .LoadItems Then
                        For i = 1 To .fpSpread1.DataRowCnt
                            .fpSpread1.GetText 1, i, cVal(1) 'Item
                            .fpSpread1.GetText 3, i, cVal(2) 'Quantity
                            .fpSpread1.GetText 4, i, cVal(3) 'Price
                            .fpSpread1.GetText 5, i, cVal(4) 'Item Code
                            .fpSpread1.GetText 6, i, cVal(5) 'HSN Code
                            .fpSpread1.GetText 7, i, cVal(6) 'HSN
                            If Val(cVal(2)) > 0 Then
                                x = fpSpread2.DataRowCnt + 1
                                fpSpread2.SetText 1, x, cVal(1)
                                fpSpread2.SetText 2, x, Val(cVal(2))
                                fpSpread2.SetText 3, x, Val(cVal(3))
                                fpSpread2.SetText 4, x, Val(cVal(2)) * Val(cVal(3))
                                fpSpread2.SetText 5, x, cVal(4)
                                fpSpread2.SetText 6, x, cVal(5)
                            End If
                        Next
                    End If
                End With
                Call CloseForm(FrmItemSearchList)
                .SetFocus
            End If
        End If
    End With
End Sub
Private Sub fpSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Item As Variant, Qty As Variant, Rate As Variant
    With fpSpread2
        If Col = 2 Or Col = 3 Then 'Qty & Rate
            .GetText 5, Row, Item
            .GetText 2, Row, Qty
            .GetText 3, Row, Rate
            If Not CheckEmpty(Item, False) Then .SetText 4, Row, Qty * Rate Else .SetText 2, Row, "": .SetText 3, Row, "": .SetText 4, Row, ""
        End If
    End With
End Sub
Private Sub fpSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub LoadMasterList(Optional ByVal LoadSelected As Boolean)
    If rstMaterialCentreList.State = adStateOpen Then rstMaterialCentreList.Close
    rstMaterialCentreList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='*99999' ORDER BY Name", cnStockJournalVoucher, adOpenKeyset, adLockReadOnly
    rstMaterialCentreList.ActiveConnection = Nothing
    If rstItemList.State = adStateOpen Then rstItemList.Close
    If LoadSelected Then
        'rstItemList.Open "SELECT I.Name As Col0,FORMAT(dbo.ufnGetItemStock('" & MaterialCentreCode & "',I.Code,'" & Left(VchPrefix, 2) & "','" & CheckNull(rstStockJournalVoucherParent.Fields("Code").Value) & "','" & GetDate(MhDateInput1.Text) & "'),'#0') As Col1,0 As Quantity,I.Price,I.Code,H.Code As HSNCode,H.Name As HSNName FROM BookMaster I INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE I.Type='F' ORDER BY I.Name", cnStockJournalVoucher, adOpenKeyset, adLockReadOnly
        rstItemList.Open "SELECT * FROM(SELECT I.Name As Col0," & _
                "FORMAT((ISNULL((SELECT SUM(OPBAL) FROM BookChild C WHERE C.MaterialCentre ='" & MaterialCentreCode & "' AND C.Item=I.Code),0) " & _
                "+ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND P.MaterialCentre ='" & MaterialCentreCode & "' AND C.Item=I.Code And SubString(P.Type,3,2)='10'),0)" & _
                "+ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND P.MaterialCentre ='" & MaterialCentreCode & "' AND C.Item=I.Code),0)" & _
                "+ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='05' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND P.MaterialCentre ='" & MaterialCentreCode & "' AND C.Item=I.Code),0)" & _
                "+ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='08' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND P.MaterialCentre ='" & MaterialCentreCode & "' AND C.Item=I.Code),0)" & _
                "+ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND MaterialCentre ='" & MaterialCentreCode & "' AND C.Item=I.Code AND C.Quantity>0),0)" & _
                "+ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND Party ='" & MaterialCentreCode & "' AND C.Item=I.Code AND C.Quantity>0),0)" & _
                "-ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND P.MaterialCentre ='" & MaterialCentreCode & "' AND C.Item=I.Code),0)" & _
                "-ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND P.MaterialCentre ='" & MaterialCentreCode & "' AND C.Item=I.Code And SubString(P.Type,3,2)='10'),0)" & _
                "-ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='06' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND P.MaterialCentre ='" & MaterialCentreCode & "' AND C.Item=I.Code),0)" & _
                "-ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='07' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND P.MaterialCentre ='" & MaterialCentreCode & "' AND C.Item=I.Code),0)" & _
                "-ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND Party ='" & MaterialCentreCode & "' AND C.Item=I.Code AND C.Quantity<0),0)" & _
                "-ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date <='" & GetDate(MhDateInput1.Text) & "' AND MaterialCentre ='" & MaterialCentreCode & "' AND C.Item=I.Code AND C.Quantity<0),0)" & _
                "),'#0') As Col1,0 As Quantity,I.Price,I.Code As code,H.Code As HSNCode,H.Name As HSNName " & _
                " FROM (BookMaster I INNER Join GeneralMaster H ON H.Code=I.HSNCode)" & _
                "WHERE I.Type='F') As Tbl ORDER BY Col0 ASC", cnStockJournalVoucher, adOpenKeyset, adLockReadOnly
    
    Else
        rstItemList.Open "SELECT I.Name As Col0,FORMAT(0,'#0') As Col1,0 As Quantity,I.Price,I.Code,H.Code As HSNCode FROM BookMaster I INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE I.Type='F' ORDER BY I.Name", cnStockJournalVoucher, adOpenKeyset, adLockReadOnly
    End If
    rstItemList.ActiveConnection = Nothing
    If rstVchSeriesList.State = adStateOpen Then rstVchSeriesList.Close
    rstVchSeriesList.Open "SELECT Name As Col0,Prefix,Suffix,VchNumbering,Code FROM VchSeriesMaster WHERE Left(FYCode,2)='" & Left(FYCode, 2) & "' AND VchType ='20" & VchType & "' ORDER BY Name", cnStockJournalVoucher, adOpenKeyset, adLockReadOnly
    rstVchSeriesList.ActiveConnection = Nothing
End Sub
Private Sub DuplicateRecord()
    Dim Tbl As String
    Tbl = "T" & GetFileNameFromPath(GetTemporaryFileName()): Tbl = Left(Tbl, InStr(1, Tbl, ".", vbTextCompare) - 1)
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    Dim VchCode As String, VchNo As String
    VchCode = GenerateCode(cnStockJournalVoucher, "SELECT MAX(Code) FROM JobworkBVParent", 6, "0")
    VchNo = GenerateCode(cnStockJournalVoucher, "SELECT MAX(VAL(Name)) FROM JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "'", 10, Space(1))
    With cnStockJournalVoucher
        .BeginTrans
        .Execute "SELECT * INTO " & Tbl & " FROM JobworkBVParent Where Code = '" & rstStockJournalVoucherList.Fields("Code").Value & "'"
        .Execute "UPDATE " & Tbl & " SET Code='" & VchCode & "',Name='" & Pad(Trim(VchNo), Space(1), 10, "L") & "',[Date]=NOW()"
        .Execute "INSERT INTO JobworkBVParent SELECT * FROM " & Tbl
        .Execute "DROP TABLE " & Tbl
        .Execute "SELECT * INTO " & Tbl & " FROM JobworkBVChild Where Code = '" & rstStockJournalVoucherList.Fields("Code").Value & "'"
        .Execute "UPDATE " & Tbl & " SET Code='" & VchCode & "'"
        .Execute "UPDATE " & Tbl & " SET RefCode=''"
        .Execute "INSERT INTO JobworkBVChild SELECT * FROM " & Tbl
        .Execute "DROP TABLE " & Tbl
        .CommitTrans
        Me.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(6)
        Me.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(2)
        Me.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(4)
    End With
    MdiMainMenu.MousePointer = vbNormal
    Call MsgBox("Successfully Duplicated the Record !", vbInformation, App.Title)
    Exit Sub
ErrorHandler:
    cnStockJournalVoucher.RollbackTrans
    MdiMainMenu.MousePointer = vbNormal
    DisplayError ("Failed to Duplicate the Record")
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Public Sub PrintStockJournalVoucher(ByVal VchCode As String, ByVal VchType As String, Optional ByVal OutputType As String)
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstStockJournalVoucherChild.State = adStateOpen Then rstStockJournalVoucherChild.Close
    If DatabaseType = "MS SQL" Then
        rstStockJournalVoucherChild.Open "SELECT LTRIM(P.Name) As VchNo,P.[Date] As VchDate,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.[PARTY]) As Godown,'FG ' AS Category,(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=C.Item)  As ItemName,CASE WHEN Quantity>=0 THEN 'Items Generated' ELSE 'Items Consumed' END As ItemType,Quantity,Remarks,'Piece' As UOMName, 'FG' As CategoryName FROM JobworkBVParent P INNER JOIN JobworkBVChild C ON P.Code=C.Code WHERE FYCode= " & FYCode & " AND P.Code='" & rstStockJournalVoucherList.Fields("Code").Value & "' ", cnStockJournalVoucher, adOpenKeyset, adLockOptimistic
        rstCompanyMaster.Open "SELECT * FROM CompanyMaster WHERE FYCode= " & FYCode & " FYCode='" & FYCode & "'", cnStockJournalVoucher, adOpenKeyset, adLockOptimistic
    Else
        rstStockJournalVoucherChild.Open "SELECT LTRIM(Name) As VchNo,[Date] As VchDate,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.Party) As Godown,'FG' as Category,(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=C.Item) As ItemName,IIF(C.Quantity>=0,'Items Generated','Items Consumed') As ItemType,C.Quantity,P.Remarks,'Piece' As UOMName, 'FG' As CategoryName FROM JobworkBVParent P Left Join JobworkBVChild C On (P.Code=C.Code And P.Code='" & rstStockJournalVoucherList.Fields("Code").Value & "' )", cnStockJournalVoucher, adOpenKeyset, adLockOptimistic
    End If
    rptStockJournal.Text1.SetText "Stock Journal Voucher"
    rptStockJournal.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptStockJournal.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False)) Then
        rptStockJournal.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False) Then
        rptStockJournal.Text24.SetText "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptStockJournal.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value)
    Else
        rptStockJournal.Section5.Suppress = True
    End If
    rptStockJournal.Text27.SetText "for " & Trim(rstStockJournalVoucherChild.Fields("Godown").Value)
    rptStockJournal.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptStockJournal.Database.SetDataSource rstStockJournalVoucherChild, 3, 1
    Screen.MousePointer = vbNormal
    If OutputType = "S" Then
        Set FrmReportViewer.Report = rptStockJournal
        FrmReportViewer.Show vbModal
    Else
        rptStockJournal.PaperSource = crPRBinAuto
        rptStockJournal.PrintOut
    End If
    Set rptStockJournal = Nothing
    On Error GoTo 0
End Sub
