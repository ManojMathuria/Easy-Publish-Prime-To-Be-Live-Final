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
Begin VB.Form frmSalesVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Voucher"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SalesVoucher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   13740
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9060
      Left            =   15
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Width           =   13710
      _Version        =   65536
      _ExtentX        =   24183
      _ExtentY        =   15981
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
      Picture         =   "SalesVoucher.frx":000C
      Begin TabDlg.SSTab SSTab1 
         Height          =   8835
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   120
         Width           =   13485
         _ExtentX        =   23786
         _ExtentY        =   15584
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
         TabPicture(0)   =   "SalesVoucher.frx":0028
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
         TabPicture(1)   =   "SalesVoucher.frx":0044
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).Control(1)=   "Mh3dLabel1(1)"
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
            Left            =   720
            MaxLength       =   40
            TabIndex        =   26
            Top             =   8310
            Width           =   8100
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7785
            Left            =   120
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   450
            Width           =   13260
            _ExtentX        =   23389
            _ExtentY        =   13732
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "Vch No."
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
               DataField       =   "PartyName"
               Caption         =   "Party"
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
               DataField       =   "ConsigneeName"
               Caption         =   "Consignee"
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
            BeginProperty Column05 
               DataField       =   "Amount"
               Caption         =   "          Amount"
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
            BeginProperty Column06 
               DataField       =   "IntegrationStatus"
               Caption         =   "Integration Status"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "Integrated"
                  FalseValue      =   "Not Integrated"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
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
                  ColumnWidth     =   1800
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
                  ColumnWidth     =   3449.764
               EndProperty
               BeginProperty Column04 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   2039.811
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
                  DividerStyle    =   3
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1544.882
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8175
            Left            =   -74880
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   480
            Width           =   13260
            _Version        =   65536
            _ExtentX        =   23389
            _ExtentY        =   14420
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
            Picture         =   "SalesVoucher.frx":0060
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
               Left            =   11400
               MaxLength       =   25
               TabIndex        =   3
               Top             =   105
               Width           =   1755
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
               Left            =   9855
               MaxLength       =   40
               TabIndex        =   48
               Top             =   630
               Width           =   3300
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
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   0
               Top             =   105
               Width           =   2130
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
               Left            =   5820
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   5
               Top             =   630
               Width           =   2730
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
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   6
               Top             =   945
               Width           =   3435
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   32
               Top             =   6330
               Width           =   13035
               _Version        =   65536
               _ExtentX        =   22992
               _ExtentY        =   494
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
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":007C
               Picture         =   "SalesVoucher.frx":0098
               Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
                  Height          =   285
                  Left            =   9210
                  TabIndex        =   10
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   930
                  _Version        =   65536
                  _ExtentX        =   1640
                  _ExtentY        =   503
                  Calculator      =   "SalesVoucher.frx":00B4
                  Caption         =   "SalesVoucher.frx":00D4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "SalesVoucher.frx":0140
                  Keys            =   "SalesVoucher.frx":015E
                  Spin            =   "SalesVoucher.frx":01A8
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "#########0"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "#########0"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   9999999999
                  MinValue        =   0
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  NegativeColor   =   255
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  ReadOnly        =   1
                  Separator       =   ""
                  ShowContextMenu =   1
                  ValueVT         =   5
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
                  Height          =   285
                  Left            =   11600
                  TabIndex        =   11
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1185
                  _Version        =   65536
                  _ExtentX        =   2090
                  _ExtentY        =   503
                  Calculator      =   "SalesVoucher.frx":01D0
                  Caption         =   "SalesVoucher.frx":01F0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "SalesVoucher.frx":025C
                  Keys            =   "SalesVoucher.frx":027A
                  Spin            =   "SalesVoucher.frx":02C4
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "#########0.00"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "#########0.00"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   9999999999.99
                  MinValue        =   0
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  NegativeColor   =   255
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  ReadOnly        =   1
                  Separator       =   ""
                  ShowContextMenu =   1
                  ValueVT         =   5
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
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
               Left            =   4620
               MaxLength       =   25
               TabIndex        =   1
               Top             =   105
               Width           =   2490
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
               Left            =   9855
               MaxLength       =   40
               TabIndex        =   8
               Top             =   950
               Width           =   3300
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
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   4
               Top             =   630
               Width           =   3435
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   3420
               TabIndex        =   29
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
               Caption         =   " Vch No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":02EC
               Picture         =   "SalesVoucher.frx":0308
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   30
               Top             =   630
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
               Caption         =   " Bill To Party"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":0324
               Picture         =   "SalesVoucher.frx":0340
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   8535
               TabIndex        =   31
               Top             =   945
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":035C
               Picture         =   "SalesVoucher.frx":0378
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   8535
               TabIndex        =   2
               Top             =   105
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   582
               Calendar        =   "SalesVoucher.frx":0394
               Caption         =   "SalesVoucher.frx":04AC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesVoucher.frx":0518
               Keys            =   "SalesVoucher.frx":0536
               Spin            =   "SalesVoucher.frx":0594
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   7215
               TabIndex        =   33
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
               Caption         =   " Vch Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":05BC
               Picture         =   "SalesVoucher.frx":05D8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   34
               Top             =   945
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
               Caption         =   " Tax Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":05F4
               Picture         =   "SalesVoucher.frx":0610
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   9940
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   6810
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "SalesVoucher.frx":062C
               Caption         =   "SalesVoucher.frx":064C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesVoucher.frx":06B8
               Keys            =   "SalesVoucher.frx":06D6
               Spin            =   "SalesVoucher.frx":0720
               AlignHorizontal =   1
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#########0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   -1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
               Height          =   330
               Left            =   9390
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   6810
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   582
               Calculator      =   "SalesVoucher.frx":0748
               Caption         =   "SalesVoucher.frx":0768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesVoucher.frx":07D4
               Keys            =   "SalesVoucher.frx":07F2
               Spin            =   "SalesVoucher.frx":083C
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
               ForeColor       =   255
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   100
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel35 
               Height          =   645
               Left            =   10785
               TabIndex        =   35
               Top             =   6810
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   1138
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
               Caption         =   " Post-Tax Amt"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":0864
               Picture         =   "SalesVoucher.frx":0880
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   645
               Left            =   11985
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   6810
               Width           =   1170
               _Version        =   65536
               _ExtentX        =   2055
               _ExtentY        =   1147
               Calculator      =   "SalesVoucher.frx":089C
               Caption         =   "SalesVoucher.frx":08BC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesVoucher.frx":0928
               Keys            =   "SalesVoucher.frx":0946
               Spin            =   "SalesVoucher.frx":0990
               AlignHorizontal =   1
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#########0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   -1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   330498053
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   650
               Left            =   1200
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   6810
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1147
               Calculator      =   "SalesVoucher.frx":09B8
               Caption         =   "SalesVoucher.frx":09D8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesVoucher.frx":0A44
               Keys            =   "SalesVoucher.frx":0A62
               Spin            =   "SalesVoucher.frx":0AAC
               AlignHorizontal =   1
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#########0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   -1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel37 
               Height          =   650
               Left            =   120
               TabIndex        =   36
               Top             =   6810
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1147
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
               Caption         =   " Pre-Tax Amt"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":0AD4
               Picture         =   "SalesVoucher.frx":0AF0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
               Height          =   330
               Left            =   8425
               TabIndex        =   37
               Top             =   7130
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
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
               Caption         =   " SGST"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":0B0C
               Picture         =   "SalesVoucher.frx":0B28
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
               Height          =   330
               Left            =   9940
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   7130
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "SalesVoucher.frx":0B44
               Caption         =   "SalesVoucher.frx":0B64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesVoucher.frx":0BD0
               Keys            =   "SalesVoucher.frx":0BEE
               Spin            =   "SalesVoucher.frx":0C38
               AlignHorizontal =   1
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#########0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   -1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
               Height          =   330
               Left            =   9390
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   7130
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   582
               Calculator      =   "SalesVoucher.frx":0C60
               Caption         =   "SalesVoucher.frx":0C80
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesVoucher.frx":0CEC
               Keys            =   "SalesVoucher.frx":0D0A
               Spin            =   "SalesVoucher.frx":0D54
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
               ForeColor       =   255
               Format          =   "##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   100
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   645
               Left            =   5235
               TabIndex        =   15
               Top             =   6810
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1147
               Calculator      =   "SalesVoucher.frx":0D7C
               Caption         =   "SalesVoucher.frx":0D9C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesVoucher.frx":0E08
               Keys            =   "SalesVoucher.frx":0E26
               Spin            =   "SalesVoucher.frx":0E70
               AlignHorizontal =   1
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#########0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   645
               Left            =   4515
               TabIndex        =   38
               Top             =   6810
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
               _ExtentY        =   1138
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
               Caption         =   " Freight"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":0E98
               Picture         =   "SalesVoucher.frx":0EB4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   650
               Left            =   2280
               TabIndex        =   39
               Top             =   6810
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   1147
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
               Caption         =   " Discount"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":0ED0
               Picture         =   "SalesVoucher.frx":0EEC
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   650
               Left            =   3120
               TabIndex        =   13
               Top             =   6810
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   1147
               Calculator      =   "SalesVoucher.frx":0F08
               Caption         =   "SalesVoucher.frx":0F28
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesVoucher.frx":0F94
               Keys            =   "SalesVoucher.frx":0FB2
               Spin            =   "SalesVoucher.frx":0FFC
               AlignHorizontal =   1
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#########0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   8425
               TabIndex        =   40
               Top             =   6810
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
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
               Caption         =   " IGST/CGST"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":1024
               Picture         =   "SalesVoucher.frx":1040
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   650
               Left            =   3675
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   6810
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   1147
               Calculator      =   "SalesVoucher.frx":105C
               Caption         =   "SalesVoucher.frx":107C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesVoucher.frx":10E8
               Keys            =   "SalesVoucher.frx":1106
               Spin            =   "SalesVoucher.frx":1150
               AlignHorizontal =   1
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#########0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   645
               Left            =   6315
               TabIndex        =   41
               Top             =   6810
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1138
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
               Caption         =   " Round Off"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":1178
               Picture         =   "SalesVoucher.frx":1194
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   645
               Left            =   7395
               TabIndex        =   16
               Top             =   6810
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   1138
               Calculator      =   "SalesVoucher.frx":11B0
               Caption         =   "SalesVoucher.frx":11D0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesVoucher.frx":123C
               Keys            =   "SalesVoucher.frx":125A
               Spin            =   "SalesVoucher.frx":12A4
               AlignHorizontal =   1
               AlignVertical   =   2
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#########0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
               MinValue        =   -9999999999.99
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   -2147483640
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
               Left            =   4620
               TabIndex        =   42
               Top             =   630
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
               Caption         =   " Mat Centre"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":12CC
               Picture         =   "SalesVoucher.frx":12E8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   4620
               TabIndex        =   43
               Top             =   945
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
               Caption         =   " Bill Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":1304
               Picture         =   "SalesVoucher.frx":1320
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   8535
               TabIndex        =   44
               Top             =   630
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
               Caption         =   " Ship To Party"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":133C
               Picture         =   "SalesVoucher.frx":1358
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   4875
               Left            =   120
               TabIndex        =   9
               Top             =   1470
               Width           =   13035
               _Version        =   524288
               _ExtentX        =   22992
               _ExtentY        =   8599
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
               MaxCols         =   17
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "SalesVoucher.frx":1374
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   120
               TabIndex        =   45
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
               Caption         =   " Vch Series"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":21E6
               Picture         =   "SalesVoucher.frx":2202
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   9960
               TabIndex        =   47
               Top             =   105
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
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
               Caption         =   " Purchase Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesVoucher.frx":221E
               Picture         =   "SalesVoucher.frx":223A
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame5 
               Height          =   525
               Left            =   120
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   7560
               Width           =   13020
               _Version        =   65536
               _ExtentX        =   22966
               _ExtentY        =   926
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
               Picture         =   "SalesVoucher.frx":2256
               Begin VB.CommandButton btnNotes 
                  Caption         =   " Notes"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   120
                  TabIndex        =   51
                  Top             =   105
                  Width           =   1455
               End
               Begin VB.TextBox txtNotes 
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
                  Left            =   1560
                  MaxLength       =   40
                  MultiLine       =   -1  'True
                  TabIndex        =   50
                  ToolTipText     =   "Open Notes"
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   1455
               End
               Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
                  Height          =   330
                  Left            =   10080
                  TabIndex        =   52
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
                  Caption         =   " Integration Status"
                  Alignment       =   0
                  FillColor       =   9164542
                  TextColor       =   0
                  Picture         =   "SalesVoucher.frx":2272
                  Picture         =   "SalesVoucher.frx":228E
               End
               Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
                  Height          =   330
                  Left            =   11760
                  TabIndex        =   53
                  TabStop         =   0   'False
                  Top             =   105
                  Width           =   1170
                  _Version        =   65536
                  _ExtentX        =   2064
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
                  Picture         =   "SalesVoucher.frx":22AA
                  Begin VB.CheckBox chkIntegrate 
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
                     Height          =   225
                     Left            =   465
                     TabIndex        =   54
                     Top             =   60
                     Width           =   180
                  End
               End
            End
            Begin MSForms.ComboBox cmbBillType 
               Height          =   330
               Left            =   5820
               TabIndex        =   7
               Top             =   945
               Width           =   2730
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "4815;582"
               ListWidth       =   4056
               MatchEntry      =   0
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Calibri"
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   13240
               Y1              =   6710
               Y2              =   6710
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   13240
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   13240
               Y1              =   1370
               Y2              =   1370
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   8805
            TabIndex        =   46
            Top             =   8310
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
            Picture         =   "SalesVoucher.frx":22C6
            Picture         =   "SalesVoucher.frx":22E2
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   1
            Left            =   -67320
            TabIndex        =   55
            Top             =   0
            Width           =   5895
            _Version        =   65536
            _ExtentX        =   10398
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
            Caption         =   " Ctrl+E->Edit  Ctrl+S OR F2->Save F11->Get Reference  F9->Delete Row"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "SalesVoucher.frx":22FE
            Picture         =   "SalesVoucher.frx":231A
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   120
            TabIndex        =   27
            Top             =   8310
            Width           =   615
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   13740
      _ExtentX        =   24236
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
Attribute VB_Name = "frmSalesVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchCode As String 'Vch to Modify
Public VchType As String, oVchType As String 'SF-Sales Voucher PF-Purchase Voucher TF-Sales Return Voucher OF-Purchase Return Voucher
Public PtgType As String 'IIf(PtgType = 1, "Sales Invoice", IIf(PtgType = 2, "Tax Invoice", IIf(PtgType = 3, "Speciman Challan", "Delivery Challan")))
Dim cnSalesVoucher As New ADODB.Connection, cnTally As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset, rstTransportList As New ADODB.Recordset
Dim rstPartyList As New ADODB.Recordset, rstMaterialCentreList As New ADODB.Recordset, rstTaxList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstHSNCodeList As New ADODB.Recordset, rstVchSeriesList As New ADODB.Recordset, rstSalesTypeList As New ADODB.Recordset
Dim rstSalesVoucherList As New ADODB.Recordset, rstSalesVoucherParent As New ADODB.Recordset, rstSalesVoucherChild As New ADODB.Recordset, rstOrderList As New ADODB.Recordset
Dim PartyCode As String, PartyStateCode As String, ConsigneeCode As String, MaterialCentreCode As String, TaxCode As String, VchPrefix As String, VchNumbering As String, VchSeriesCode As String, oVchSeriesCode As String, oVchNo As String, AutoVchNo As String, StartNo As String, oVchDate As Date, SalesTypeCode As String
Dim SortOrder, PrevStr, dblBookMark As Double, blnRecordExist As Boolean, EditMode As Boolean, VchSeries As String
Dim frmSalesTptDetails As New FrmDespatchDetails
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
    Me.Top = 500
    WheelHook DataGrid1
    BusySystemIndicator True
    oVchType = VchType
    Mh3dLabel15.Caption = IIf(InStr(1, "SF_TF", VchType) > 0, " Sales", " Purchase") + " Type"
    Me.Caption = IIf(VchType = "SF", "Sales", IIf(VchType = "PF", "Purchase", IIf(VchType = "TF", "Sales Return", "Purchase Return"))) & "-Supply " & IIf(VchType = "SF", "Outward", IIf(VchType = "PF", "Inward", IIf(VchType = "TF", "Outward Return", "Inward Return"))) & "-Finished Goods"
    cnSalesVoucher.CursorLocation = adUseClient: cnSalesVoucher.Open cnDatabase.ConnectionString: cnTally.CursorLocation = adUseClient
    rstSalesVoucherParent.CursorLocation = adUseClient
    LoadMasterList
    With rstSalesVoucherList
        .Open "SELECT T.Code,T.Name,V.Code As VchSeriesCode,V.Name As VchSeriesName,Date,T.Type,P.Name As PartyName,C.Name As ConsigneeName,Amount,IntegrationStatus FROM ((JobworkBVParent T INNER JOIN AccountMaster P ON T.Party=P.Code) INNER JOIN AccountMaster C ON T.Consignee=C.Code) INNER JOIN VchSeriesMaster V ON T.VchSeries=V.Code WHERE RIGHT(Type,2)='" & VchType & "' AND T.FYCode='" & FYCode & "' ORDER BY T.Name", cnSalesVoucher, adOpenKeyset, adLockPessimistic
        .Filter = adFilterNone
        If .RecordCount > 0 Then
            .MoveLast
            If Not CheckEmpty(VchCode, False) Then .MoveFirst: .Find "[Code]='" & VchCode & "'"
        End If
        Set DataGrid1.DataSource = rstSalesVoucherList
        Me.DataGrid1.Columns(6).Caption = IIf(TallyIntegration = True, "Tally Integration", "Busy Integration")
        BusySystemIndicator False
        SSTab1.Tab = 0
    If FrmStockLedger.dSortBy = True Or FrmAccountLedger.dSortBy = True Then
        SortOrder = "Code"
    ElseIf FrmAccountLedger.dSortBy = True Then
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
    cmbBillType.AddItem "Direct", 0 'Includes against Sales/Purchase Order in case of Sales/Purchase Voucher
    cmbBillType.AddItem "Against Challan", 1
    SetButtonsForNoRecord
    fpSpread1.TextTip = TextTipFloating
    Load frmSalesTptDetails
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    With MdiMainMenu
        .mnuSalesSupplyOutwardFinishedItem.Enabled = False: .mnuSalesReturnSupplyOutwardReturnFinishedItem.Enabled = False: .mnuPurchaseSupplyInwardFinishedItem.Enabled = False: .mnuPurchaseReturnSupplyInwardReturnFinishedItem.Enabled = False
    End With
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    With Toolbar1.Buttons
        If InStr(1, "fpSpread1", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True Else EditMode = False
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
               If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}"
            End If
            If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0
        End If
    End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstSalesVoucherList)
    Call CloseRecordset(rstSalesVoucherParent)
    Call CloseRecordset(rstSalesVoucherChild)
    Call CloseRecordset(rstPartyList)
    Call CloseRecordset(rstMaterialCentreList)
    Call CloseRecordset(rstTaxList)
    Call CloseRecordset(rstHSNCodeList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstVchSeriesList)
    Call CloseRecordset(rstOrderList)
    Call CloseConnection(cnSalesVoucher)
    Call CloseConnection(cnTally)
    Call CloseForm(frmSalesTptDetails)
    ShowProgressInStatusBar False
    DisableChildMenu
    With MdiMainMenu
        .mnuSalesSupplyOutwardFinishedItem.Enabled = True: .mnuSalesReturnSupplyOutwardReturnFinishedItem.Enabled = True: .mnuPurchaseSupplyInwardFinishedItem.Enabled = True: .mnuPurchaseReturnSupplyInwardReturnFinishedItem.Enabled = True
    End With
End Sub
Private Sub Text1_Change()
On Error Resume Next
    With rstSalesVoucherList
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
    With rstSalesVoucherList
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
            If Not (rstSalesVoucherList.EOF Or rstSalesVoucherList.BOF) Then
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
        Text6.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean, UpdateFlag As Integer, CellVal01 As Variant, CellVal02 As Variant, CellVal03 As Variant, i As Integer
    With rstSalesVoucherList
        If Button.Index = 1 Then
            If rstSalesVoucherParent.State = adStateOpen Then rstSalesVoucherParent.Close
            rstSalesVoucherParent.Open "SELECT * FROM JobworkBVParent WHERE Code=''", cnSalesVoucher, adOpenKeyset, adLockOptimistic
            ClearFields
            If AddRecord(rstSalesVoucherParent) Then
                MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
                Call SetButtons(False)
                SSTab1.Tab = 1
                Text6.SetFocus
                blnRecordExist = False
                cnSalesVoucher.BeginTrans
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
                cnSalesVoucher.BeginTrans
                cnSalesVoucher.Execute "DELETE FROM JobworkBVRef WHERE VchCode='" & .Fields("Code").Value & "'"
                cnSalesVoucher.Execute "DELETE FROM JobworkBVParent WHERE Code='" & .Fields("Code").Value & "'"
                MdiMainMenu.MousePointer = vbNormal
                If Err.Number = 0 Then
                    .Delete
                    .MoveNext
                    If .RecordCount > 0 And .EOF Then .MoveLast
                    cnSalesVoucher.CommitTrans
                    ShowProgressInStatusBar True
                    Timer1.Enabled = True
                    Text1.Text = ""
                    .Filter = adFilterNone
                    If BusyIntegration Or TallyIntegration Then DelOldVch True
                Else
                    DisplayError (Err.Description)
                    cnSalesVoucher.RollbackTrans
                End If
                On Error GoTo 0
            End If
            SetButtons (True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
            HiLiteRecord = True
        ElseIf Button.Index = 4 Then
            Sendkeys "{TAB}"
            If CheckMandatoryFields Then Exit Sub
            Load_TransportList
            frmSalesTptDetails.Show vbModal
            If MsgBox("Are you sure to save the voucher?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Save !") = vbNo Then Exit Sub
            SaveFields
            UpdateFlag = 0
            If UpdateRecord(rstSalesVoucherParent) Then
                If UpdateItemList("D", 0, 0) Then
                    UpdateFlag = 1
                   With fpSpread1
                       For i = 1 To .DataRowCnt
                           .SetActiveCell 3, i
                           .GetText 4, i, CellVal01 'Quantity
                           .GetText 8, i, CellVal02 'Item Code
                           .GetText 12, i, CellVal03 'Bal Qty
                           If Val(CellVal01) <> 0 And Not CheckEmpty(CellVal02, False) Then If Not UpdateItemList("I", i, Val(CellVal03)) Then UpdateFlag = 0: Exit For
                       Next
                   End With
                End If
            End If
            If UpdateFlag Then
                AddToList
                cnSalesVoucher.CommitTrans
                If rstSalesVoucherParent.State = adStateOpen Then rstSalesVoucherParent.Close
                rstSalesVoucherParent.CursorLocation = adUseClient
                Call SetButtons(True)
                ShowProgressInStatusBar True
                Timer1.Enabled = True
                Call MsgBox("Record updated !!!", vbInformation, App.Title)
                'If VchType = "SF" Then
                    If BusyIntegration Or TallyIntegration Then
                        If MsgBox("Are you sure to export the Voucher?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Export !") = vbYes Then PushVch
                    End If
                'End If
                SSTab1.Tab = 0
            Else
                DisplayError ("Failed to save the record")
                Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            End If
        ElseIf Button.Index = 5 Then
            If CancelRecordUpdate(rstSalesVoucherParent) Then
                cnSalesVoucher.RollbackTrans
                If rstSalesVoucherParent.State = adStateOpen Then rstSalesVoucherParent.Close
                rstSalesVoucherParent.CursorLocation = adUseClient
                Call SetButtons(True)
                SetButtonsForNoRecord
                SSTab1.Tab = 0
            End If
        ElseIf Button.Index = 6 Then
            SSTab1.Tab = 0
            Set DataGrid1.DataSource = Nothing
            .Filter = adFilterNone
            RefreshData rstSalesVoucherList
            Set DataGrid1.DataSource = rstSalesVoucherList
            If .RecordCount > 0 Then .MoveLast
            LoadMasterList
            HiLiteRecord = True
        ElseIf Button.Index = 7 Then
            SSTab1.Tab = 0
            With FrmFilter
                .Combo1.AddItem "Party", 0
                .Combo1.AddItem "Material Centre", 1
                .Combo1.ListIndex = 0
                Set .srcForm = Me
                .Show vbModal
            End With
            HiLiteRecord = True
        ElseIf Button.Index = 9 Then
            If .RecordCount = 0 Then Exit Sub
            Call PrintSalesVoucher(.Fields("Code").Value, Right(.Fields("Type").Value, 2), "P")
            HiLiteRecord = True
        ElseIf Button.Index = 10 Then
            If .RecordCount = 0 Then Exit Sub
            Call PrintSalesVoucher(.Fields("Code").Value, Right(.Fields("Type").Value, 2), "S")
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
Private Sub Load_TransportList()
        Dim i As Integer, j As Integer
'ComboBox1
        If rstTransportList.State = adStateOpen Then rstTransportList.Close
            rstTransportList.Open "SELECT Transport      As Transporter From JobworkBVParent Where Party='" & PartyCode & "' UNION " & _
                                                "SELECT Transporter   As Transporter FROM AccountMaster Where Code='" & PartyCode & "' UNION " & _
                                                "SELECT Transporter2 As Transporter  FROM AccountMaster Where Code='" & PartyCode & "' UNION " & _
                                                "SELECT Transporter3 As Transporter  FROM AccountMaster Where Code='" & PartyCode & "' UNION " & _
                                                "SELECT Transporter4 As Transporter  FROM AccountMaster Where Code='" & PartyCode & "'", cnSalesVoucher, adOpenKeyset, adLockReadOnly
        On Error GoTo ErrorHandler
        frmSalesTptDetails.ComboFlag = False
        rstTransportList.ActiveConnection = Nothing
    With rstTransportList

            frmSalesTptDetails.ComboBox1.Clear
            If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If Not IsNull(.Fields("Transporter").Value) Then If Trim(.Fields("Transporter").Value) <> "" Then frmSalesTptDetails.ComboBox1.AddItem .Fields("Transporter").Value, i: i = i + 1
            .MoveNext
        Loop
    End With
'ComboBox2
        If rstTransportList.State = adStateOpen Then rstTransportList.Close
            rstTransportList.Open "SELECT Station  AS City From JobworkBVParent Where Party='" & PartyCode & "' UNION " & _
                                                "SELECT City                    FROM AccountMaster Where Code='" & PartyCode & "' UNION " & _
                                                "SELECT City                    FROM AccountMaster Where Code='" & PartyCode & "' UNION " & _
                                                "SELECT City                    FROM AccountMaster Where Code='" & PartyCode & "' UNION " & _
                                                "SELECT City                    FROM AccountMaster Where Code='" & PartyCode & "'", cnSalesVoucher, adOpenKeyset, adLockReadOnly
        On Error GoTo ErrorHandler
        rstTransportList.ActiveConnection = Nothing
    With rstTransportList
            frmSalesTptDetails.ComboBox2.Clear
            If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If Not IsNull(.Fields("City").Value) Then If Trim(.Fields("City").Value) <> "" Then frmSalesTptDetails.ComboBox2.AddItem .Fields("City").Value, 0: j = j + 1
            .MoveNext
        Loop
    End With
            frmSalesTptDetails.ComboFlag = True
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Transport List")
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstSalesVoucherList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstSalesVoucherList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstSalesVoucherList.EOF Or rstSalesVoucherList.BOF) Then
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
    If rstSalesVoucherList.RecordCount = 0 Then
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
Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountGroup = IIf(InStr(1, "SF_TF", VchType) > 0, "*26027", "*26025")
        FrmAccountMaster.AccountType = "01"
        FrmAccountMaster.MasterCode = SalesTypeCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        SalesTypeCode = slCode: Text10.Text = slName
        If Not CheckEmpty(SalesTypeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text10_Validate(Cancel As Boolean)
    If CheckEmpty(Text10.Text, False) Then Cancel = True
End Sub
Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text6.Text)
        If rstVchSeriesList.RecordCount = 0 Then DisplayError ("No Record in Voucher Series Master"): Text6.SetFocus: Exit Sub Else rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": VchSeriesCode = ""
        Call LoadSelectionList(rstVchSeriesList, "List of Voucher Series...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text6, VchSeriesCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(VchSeriesCode) <> "" Then Sendkeys "{TAB}" Else Text6.Text = ""
    End If
End Sub
Private Sub Text6_Validate(Cancel As Boolean)
    If CheckEmpty(Text6.Text, False) Then
        Cancel = True
    Else
        rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        VchNumbering = rstVchSeriesList.Fields("VchNumbering").Value
        StartNo = rstVchSeriesList.Fields("StartNo").Value
        If VchNumbering = "A" Then Text2.Locked = True Else Text2.Locked = False
        If Not blnRecordExist Then 'Vch-New
            If VchNumbering = "A" Then
                AutoVchNo = GenerateCode(cnSalesVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
            Else
                AutoVchNo = GenerateCode(cnSalesVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                If Trim(AutoVchNo) > StartNo Then
                    Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                Else
                    AutoVchNo = StartNo
                    AutoVchNo = Pad(AutoVchNo, " ", 10, "L")
                    Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                End If
            End If
        Else 'Vch-Old
            If VchNumbering = "A" Then
                If VchSeriesCode = oVchSeriesCode Then
                    Text2.Text = oVchNo
                Else
                    If VchNumbering = "A" Then
                        AutoVchNo = GenerateCode(cnSalesVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                        Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                    End If
                End If
            Else
                If VchSeriesCode = oVchSeriesCode Then
                    AutoVchNo = GetNumValue(Trim(Text2.Text))
                    AutoVchNo = Pad(AutoVchNo, " ", 10, "L")
                    Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                Else
                    AutoVchNo = GenerateCode(cnSalesVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                    If Trim(AutoVchNo) > StartNo Then
                        Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                    Else
                        AutoVchNo = StartNo
                        AutoVchNo = Pad(AutoVchNo, " ", 10, "L")
                        Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean) 'Vch No.
    With rstSalesVoucherParent
        If .EOF Or .BOF Then Exit Sub
        If CheckEmpty(Text2, True) Then
            Cancel = True
        ElseIf CheckDuplicate(cnSalesVoucher, "JobworkBVParent", "Code", "[Name]+RIGHT(Type,2)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, .Fields("Code").Value, False, FYCode) Then
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
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = PartyCode
        FrmAccountMaster.StateCode = PartyStateCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        PartyCode = slCode: Text3.Text = slName
        If Not IsNull(slStateCode) Then
        PartyStateCode = slStateCode
        End If
        If Not CheckEmpty(PartyCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        PartyCode = "": Text3.Text = ""
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If CheckEmpty(Text3.Text, False) Then Cancel = True
    If CheckEmpty(Text8.Text, False) Then Text8.Text = Text3.Text: ConsigneeCode = PartyCode
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = "*99999"
        FrmAccountMaster.MasterCode = MaterialCentreCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        MaterialCentreCode = slCode: Text7.Text = slName
        If Not CheckEmpty(MaterialCentreCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        MaterialCentreCode = "": Text7.Text = ""
    End If
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
    If CheckEmpty(Text7.Text, False) Then Cancel = True
End Sub
Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = ConsigneeCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        ConsigneeCode = slCode: Text8.Text = slName
        If Not CheckEmpty(ConsigneeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        ConsigneeCode = "": Text8.Text = ""
    End If
End Sub
Private Sub Text8_Validate(Cancel As Boolean)
    If CheckEmpty(Text8.Text, False) Then Cancel = True
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        slStateCode = PartyStateCode
        FrmTaxMaster.SL = True
        FrmTaxMaster.MasterCode = TaxCode
        Load FrmTaxMaster
        If Err.Number <> 364 Then FrmTaxMaster.Show vbModal
        On Error GoTo 0
        TaxCode = slCode: Text5.Text = slName
        If Not CheckEmpty(TaxCode, False) Then
            rstTaxList.MoveFirst: rstTaxList.Find "[Code] = '" & TaxCode & "'"
            If Val(rstTaxList.Fields("SGST%").Value) > 0 Then   'Intra-State GST
                MhRealInput7.Value = Val(rstTaxList.Fields("CGST%").Value)
                MhRealInput9.Value = Val(rstTaxList.Fields("SGST%").Value)
            Else    'Inter-State GST
                MhRealInput7.Value = Val(rstTaxList.Fields("IGST%").Value)
                MhRealInput9.Value = 0
            End If
            CalculateTotal
            LoadMasterList
            Sendkeys "{TAB}"
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TaxCode = "": Text5.Text = ""
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    If CheckEmpty(Text5.Text, False) Then Cancel = True
End Sub
Private Sub cmbBillType_Click()
    '01-Purchase 02-Purchase Return 03-Sales Return 04-Sales
    VchPrefix = IIf(VchType = "SF", "04", IIf(VchType = "PF", "01", IIf(VchType = "TF", "03", "02"))) & Choose(cmbBillType.ListIndex + 1, "10", "01") '10-Stock affected 01-Stock not affected
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)    'Discount
    CalculateTotal
End Sub
Private Sub MhRealInput6_Validate(Cancel As Boolean)    'Freight
    CalculateTotal
End Sub
Private Sub MhRealInput12_Validate(Cancel As Boolean)   'Adjustment
    CalculateTotal
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstSalesVoucherList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    With rstSalesVoucherParent
        If .State = adStateOpen Then .Close
        .Open "SELECT *, (Select State From AccountMaster Where Code=Party) AS State FROM JobworkBVParent WHERE Code='" & FixQuote(rstSalesVoucherList.Fields("Code").Value) & "'", cnSalesVoucher, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End With
End Sub
Private Sub ClearFields()
    Text6.Text = "" 'Vch Series
    Text2.Text = "" 'Vch No.
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    Text3.Text = "" 'Party
    Text7.Text = "" 'Material Centre
    Text8.Text = "" 'Consignee
    Text10.Text = "" 'SalesType
    Text5.Text = "" 'Tax Name
    cmbBillType.ListIndex = 0: cmbBillType.Enabled = True: cmbBillType_Click
    Text4.Text = "" 'Remarks
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    MhRealInput1.Value = 0
    MhRealInput2.Value = 0
    MhRealInput3.Value = 0
    MhRealInput4.Value = 0
    MhRealInput5.Value = 0
    MhRealInput6.Value = 0
    MhRealInput12.Value = 0
    MhRealInput7.Value = 0
    MhRealInput9.Value = 0
    MhRealInput8.Value = 0
    MhRealInput10.Value = 0
    MhRealInput11.Value = 0
    chkIntegrate.Value = 0
    PartyStateCode = ""
    PartyCode = "": ConsigneeCode = "": MaterialCentreCode = "": TaxCode = "": VchSeriesCode = "": oVchSeriesCode = "": oVchNo = "": AutoVchNo = ""
    frmSalesTptDetails.Text1.Text = "": frmSalesTptDetails.Text2.Text = "": frmSalesTptDetails.Text3.Text = "": frmSalesTptDetails.Text4.Text = "": frmSalesTptDetails.Text5.Text = "": frmSalesTptDetails.MhDateInput1.Value = Null: frmSalesTptDetails.MhDateInput2.Value = Null
End Sub
Private Sub LoadFields()
    With rstSalesVoucherParent
        If .EOF Or .BOF Then Exit Sub
        VchSeriesCode = .Fields("VchSeries").Value: oVchSeriesCode = VchSeriesCode
        If rstVchSeriesList.RecordCount > 0 Then rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        If Not rstVchSeriesList.EOF Then Text6.Text = rstVchSeriesList.Fields("Col0").Value
        Text2.Text = Trim(.Fields("Name").Value)
        AutoVchNo = Trim(.Fields("AutoVchNo").Value)
        oVchNo = Trim(Text2.Text)
        oVchDate = Format(.Fields("Date").Value, "dd-MMM-yyyy")
        MhDateInput1.Text = Format(.Fields("Date").Value, "dd-MM-yyyy")
        PartyCode = .Fields("Party").Value
        PartyStateCode = .Fields("State").Value
        If rstPartyList.RecordCount > 0 Then rstPartyList.MoveFirst
        rstPartyList.Find "[Code] = '" & PartyCode & "'"
        If Not rstPartyList.EOF Then Text3.Text = rstPartyList.Fields("Col0").Value
        MaterialCentreCode = .Fields("MaterialCentre").Value
        If rstMaterialCentreList.RecordCount > 0 Then rstMaterialCentreList.MoveFirst
        rstMaterialCentreList.Find "[Code] = '" & MaterialCentreCode & "'"
        If Not rstMaterialCentreList.EOF Then Text7.Text = rstMaterialCentreList.Fields("Col0").Value
        ConsigneeCode = .Fields("Consignee").Value
        If rstPartyList.RecordCount > 0 Then rstPartyList.MoveFirst
        rstPartyList.Find "[Code] = '" & ConsigneeCode & "'"
        If Not rstPartyList.EOF Then Text8.Text = rstPartyList.Fields("Col0").Value
        SalesTypeCode = .Fields("SalesType").Value
        If rstSalesTypeList.RecordCount > 0 Then rstSalesTypeList.MoveFirst
        rstSalesTypeList.Find "[Code] = '" & SalesTypeCode & "'"
        If Not rstSalesTypeList.EOF Then Text10.Text = rstSalesTypeList.Fields("Col0").Value
        TaxCode = .Fields("Tax").Value
        If rstTaxList.RecordCount > 0 Then rstTaxList.MoveFirst
        rstTaxList.Find "[Code] = '" & TaxCode & "'"
        If Not rstTaxList.EOF Then Text5.Text = rstTaxList.Fields("Col0").Value
        cmbBillType.ListIndex = IIf(Mid(.Fields("Type").Value, 3, 2) = "10", 0, 1)
        Text4.Text = .Fields("Remarks").Value
        Call LoadItemList(.Fields("Code").Value)
        MhRealInput4.Value = Val(.Fields("Rebate%").Value)
        MhRealInput5.Value = Val(.Fields("Rebate").Value)
        MhRealInput6.Value = Val(.Fields("Freight").Value)
        MhRealInput12.Value = Val(.Fields("Adjustment").Value)
        If Val(.Fields("SGST%").Value) > 0 Then  'Intra-State Supply
            MhRealInput7.Value = Val(.Fields("CGST%").Value)
            MhRealInput8.Value = Val(.Fields("CGST").Value)
            MhRealInput9.Value = Val(.Fields("SGST%").Value)
            MhRealInput10.Value = Val(.Fields("SGST").Value)
        Else    'Inter-State Supply
            MhRealInput7.Value = Val(.Fields("IGST%").Value)
            MhRealInput8.Value = Val(.Fields("IGST").Value)
        End If
        MhRealInput11.Value = Val(.Fields("Amount").Value)
        chkIntegrate.Value = IIf(.Fields("IntegrationStatus").Value = "True", 1, 0)
        txtNotes.Text = .Fields("Notes").Value
        frmSalesTptDetails.Text1.Text = CheckNull(.Fields("Transport").Value): frmSalesTptDetails.Text2.Text = CheckNull(.Fields("GRNo").Value): frmSalesTptDetails.Text3.Text = CheckNull(.Fields("VehicleNo").Value): frmSalesTptDetails.Text4.Text = CheckNull(.Fields("Station").Value): frmSalesTptDetails.Text5.Text = CheckNull(.Fields("eWayBill").Value): If Not IsNull(.Fields("GRDate").Value) Then frmSalesTptDetails.MhDateInput1.Value = .Fields("GRDate").Value: If Not IsNull(.Fields("eWayBill").Value) Then frmSalesTptDetails.MhDateInput2.Value = .Fields("eWayBillDate").Value
    End With
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    With rstSalesVoucherParent
        If .RecordCount = 0 Then Exit Sub
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseServer
        .Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstSalesVoucherList.Fields("Code").Value) & "'", cnSalesVoucher, adOpenKeyset, adLockPessimistic
        MdiMainMenu.MousePointer = vbHourglass
        .Fields("RecordStatus") = "N"
    End With
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    If fpSpread1.DataRowCnt > 0 Then cmbBillType.Enabled = False
    Text6.SetFocus
    blnRecordExist = True
    cnSalesVoucher.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    With rstSalesVoucherParent
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnSalesVoucher, "SELECT MAX(Code) FROM JobworkBVParent", 6, "0")
            .Fields("CreatedBy").Value = UserCode
            .Fields("CreatedOn").Value = Now()
            .Fields("Recordstatus").Value = "N"
        Else
            .Fields("ModifiedBy").Value = UserCode
            .Fields("ModifiedOn").Value = Now()
            .Fields("Recordstatus").Value = "M"
        End If
        .Fields("VchSeries").Value = VchSeriesCode
        .Fields("AutoVchNo").Value = Pad(Trim(AutoVchNo), Space(1), 10, "L")
        .Fields("Name").Value = Trim(Text2.Text)
        .Fields("Date").Value = GetDate(MhDateInput1.Text)
        .Fields("Party").Value = PartyCode
        .Fields("MaterialCentre").Value = MaterialCentreCode
        .Fields("Consignee").Value = ConsigneeCode
        .Fields("Tax").Value = TaxCode
        .Fields("Remarks").Value = Trim(Text4.Text)
        .Fields("Rebate%").Value = MhRealInput4.Value
        .Fields("Rebate").Value = MhRealInput5.Value
        .Fields("Freight").Value = MhRealInput6.Value
        .Fields("Adjustment").Value = MhRealInput12.Value
        .Fields("TaxableAmount").Value = MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value 'Pre-Tax Amt - Discount + Freight
        If MhRealInput9.Value > 0 Then  'Intra-State Supply
            .Fields("CGST%").Value = MhRealInput7.Value
            .Fields("CGST").Value = MhRealInput8.Value
            .Fields("SGST%").Value = MhRealInput9.Value
            .Fields("SGST").Value = MhRealInput10.Value
            .Fields("IGST%").Value = 0
            .Fields("IGST").Value = 0
        Else    'Inter-State Supply
            .Fields("CGST%").Value = 0
            .Fields("CGST").Value = 0
            .Fields("SGST%").Value = 0
            .Fields("SGST").Value = 0
            .Fields("IGST%").Value = MhRealInput7.Value
            .Fields("IGST").Value = MhRealInput8.Value
        End If
        .Fields("Amount").Value = MhRealInput11.Value
        .Fields("Type").Value = VchPrefix & VchType
        .Fields("FYCode").Value = FYCode
        .Fields("RecordStatus").Value = "N"
        .Fields("Notes").Value = txtNotes.Text
        .Fields("SalesType").Value = SalesTypeCode
        .Fields("Transport").Value = frmSalesTptDetails.Text1.Text
        .Fields("GRNo").Value = frmSalesTptDetails.Text2.Text
        If frmSalesTptDetails.MhDateInput1.ValueIsNull Then .Fields("GRDate").Value = Null Else .Fields("GRDate").Value = GetDate(frmSalesTptDetails.MhDateInput1.Text)
        .Fields("eWayBill").Value = frmSalesTptDetails.Text5.Text
        If frmSalesTptDetails.MhDateInput2.ValueIsNull Then .Fields("eWayBillDate").Value = Null Else .Fields("eWayBillDate").Value = GetDate(frmSalesTptDetails.MhDateInput2.Text)
        .Fields("VehicleNo").Value = frmSalesTptDetails.Text3.Text
        .Fields("IntegrationStatus").Value = IIf(chkIntegrate.Value = 1, "True", "False")
        .Fields("Station").Value = frmSalesTptDetails.Text4.Text
    End With
End Sub
Private Sub AddToList()
    On Error Resume Next
    With rstSalesVoucherList
        .MoveFirst
        .Find "[Code] = '" & rstSalesVoucherParent.Fields("Code").Value & "'"
        If .EOF Then .AddNew
        .Fields("Code").Value = rstSalesVoucherParent.Fields("Code").Value
        .Fields("Name").Value = Trim(rstSalesVoucherParent.Fields("Name").Value)
        .Fields("VchSeriesCode").Value = VchSeriesCode
        .Fields("VchSeriesName").Value = Text6.Text
        .Fields("Date").Value = rstSalesVoucherParent.Fields("Date").Value
        .Fields("PartyName").Value = Trim(Text3.Text)
        .Fields("ConsigneeName").Value = Trim(Text8.Text)
        .Fields("ChallanNo").Value = rstSalesVoucherParent.Fields("ChallanNo").Value
        .Fields("ChallanDate").Value = rstSalesVoucherParent.Fields("ChallanDate").Value
        .Fields("Amount").Value = MhRealInput11.Value
        .Fields("IntegrationStatus").Value = IIf(chkIntegrate.Value = 1, "True", "False")
        .Fields("SalesType").Value = Trim(Text10.Text)
        .Fields("Type").Value = rstSalesVoucherParent.Fields("Type").Value
        .Update
        .Sort = SortOrder & " Asc"
        .Find "[Code] = '" & rstSalesVoucherParent.Fields("Code").Value & "'"
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text6.Text, False) Then
        Text6.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text2.Text, False) Then
        DisplayError ("Voucher No. cannot be blank"): Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnSalesVoucher, "JobworkBVParent", "Code", "[Name]+RIGHT(Type,2)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, rstSalesVoucherParent.Fields("Code").Value, False, FYCode) Then
        Dim VchNo As String
        If Not blnRecordExist Then 'Vch-New
            If VchNumbering = "A" Then
                AutoVchNo = GenerateCode(cnSalesVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                VchNo = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                If Trim(VchNo) <> Trim(Text2.Text) Then DisplayError ("Vch No. changed from " & Trim(Text2.Text) & " to " & Trim(VchNo))
                Text2.Text = VchNo: Exit Function
            End If
        Else 'Vch-Old
            If VchSeriesCode = oVchSeriesCode Then
                Text2.Text = oVchNo
            Else
                If VchNumbering = "A" Then
                    AutoVchNo = GenerateCode(cnSalesVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                    VchNo = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                    If Trim(VchNo) <> Trim(Text2.Text) Then DisplayError ("Vch No. changed from " & Trim(Text2.Text) & " to " & Trim(VchNo))
                    Text2.Text = VchNo: Exit Function
                End If
            End If
'        Else
            Text2.SetFocus: CheckMandatoryFields = True: Exit Function
        End If
    ElseIf CheckEmpty(Text3.Text, False) Then 'Party
        Text3.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text7.Text, False) Then 'Material Centre
        Text7.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text8.Text, False) Then 'Consignee
        Text8.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text10.Text, False) Then 'SalesType
        Text10.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text5.Text, False) Then 'Tax
        Text5.SetFocus:   CheckMandatoryFields = True: Exit Function
    End If
        Call Text6_Validate(False)
End Function
Private Sub LoadItemList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    With rstSalesVoucherChild
        If .State = adStateOpen Then .Close
        If cmbBillType.ListIndex = 0 Then 'Direct
            .Open "SELECT I.Code As ItemCode,I.Name As ItemName,H.Code As HSNCode,H.Name As HSNName,T.Ref As RefOrderCode,(SELECT LTRIM(VchNo) FROM JobworkBVRef WHERE RefCode=T.Ref AND RIGHT(VchType,2)='" & IIf(VchType = "SF", "SO", IIf(VchType = "PF", "PO", "")) & "') As RefOrderNo,ABS(T.Quantity) As Quantity,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.Ref AND VchCode<>'" & strOrderCode & "')*1) As BalQty,T.Rate,T.[Disc%],T.Amount,T.RefCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM (JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) LEFT JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & strOrderCode & "' ORDER BY T.SrNo", cnSalesVoucher, adOpenKeyset, adLockReadOnly
        Else
            .Open "SELECT I.Code As ItemCode,I.Name As ItemName,H.Code As HSNCode,H.Name As HSNName,T.Ref As RefOrderCode,(SELECT LTRIM(VchNo) FROM JobworkBVRef WHERE RefCode=T.Ref AND LEFT(VchType,2)+RIGHT(VchType,2)='" & IIf(VchType = "SF", "08IF", IIf(VchType = "PF", "05RF", IIf(VchType = "TF", "07RF", "06IF"))) & "') As RefOrderNo,ABS(T.Quantity) As Quantity,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.Ref AND VchCode<>'" & strOrderCode & "')*1) As BalQty,T.Rate,T.[Disc%],T.Amount,T.RefCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM (JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) LEFT JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & strOrderCode & "' ORDER BY T.SrNo", cnSalesVoucher, adOpenKeyset, adLockReadOnly
        End If
        .ActiveConnection = Nothing
        If .RecordCount > 0 Then .MoveFirst
        i = 0
        Do While Not .EOF
            i = i + 1
            fpSpread1.SetText 1, i, .Fields("ItemName").Value
            fpSpread1.SetText 2, i, .Fields("HSNName").Value
            fpSpread1.SetText 3, i, .Fields("RefOrderNo").Value
            fpSpread1.SetText 4, i, Val(.Fields("Quantity").Value)
            fpSpread1.SetText 5, i, Val(.Fields("Rate").Value)
            fpSpread1.SetText 6, i, Val(.Fields("Disc%").Value)
            fpSpread1.SetText 7, i, Val(.Fields("Amount").Value)
            fpSpread1.SetText 8, i, .Fields("ItemCode").Value
            fpSpread1.SetText 9, i, .Fields("HSNCode").Value
            fpSpread1.SetText 10, i, .Fields("RefOrderCode").Value
            fpSpread1.SetText 11, i, .Fields("RefCode").Value
            fpSpread1.SetText 12, i, Val(CheckNull(.Fields("BalQty").Value))
            fpSpread1.SetText 13, i, .Fields("LongNarration01").Value
            fpSpread1.SetText 14, i, .Fields("LongNarration02").Value
            fpSpread1.SetText 15, i, .Fields("LongNarration03").Value
            fpSpread1.SetText 16, i, .Fields("LongNarration04").Value
            fpSpread1.SetText 17, i, .Fields("LongNarration05").Value
            .MoveNext
        Loop
    End With
    CalculateTotal
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Item List")
End Sub
Private Function UpdateItemList(ByVal ActionType As String, ByVal SrNo As Integer, ByVal BalQty As Long) As Boolean
    Dim CellVal(1 To 12) As Variant
    On Error GoTo ErrorHandler
    UpdateItemList = True
    If ActionType = "D" Then
        If Not blnRecordExist Then Exit Function
        cnSalesVoucher.Execute "DELETE FROM JobworkBVRef WHERE VchCode='" & rstSalesVoucherParent.Fields("Code").Value & "'"
        cnSalesVoucher.Execute "DELETE FROM JobworkBVChild WHERE Code='" & rstSalesVoucherParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 4, .ActiveRow, CellVal(1)  'Quantity
            .GetText 5, .ActiveRow, CellVal(2)  'Rate
            .GetText 6, .ActiveRow, CellVal(3)  'Disc %
            .GetText 7, .ActiveRow, CellVal(4)  'Amount
            .GetText 8, .ActiveRow, CellVal(5)  'Item Code
            .GetText 9, .ActiveRow, CellVal(6)  'HSN Code
            .GetText 10, .ActiveRow, CellVal(7)  'Ref Order Code
            .GetText 13, .ActiveRow, CellVal(8) 'Long Narration I
            .GetText 14, .ActiveRow, CellVal(9) 'Long Narration II
            .GetText 15, .ActiveRow, CellVal(10) 'Long Narration III
            .GetText 16, .ActiveRow, CellVal(11) 'Long Narration IV
            .GetText 17, .ActiveRow, CellVal(12) 'Long Narration V
        End With
        'For child
        CellVal(1) = IIf(InStr(1, "SF_OF", VchType) > 0, 0 - Val(CellVal(1)), Val(CellVal(1))) '-ve for SF/OF & +ve for PF/TF
        cnSalesVoucher.Execute "INSERT INTO JobworkBVChild VALUES ('" & rstSalesVoucherParent.Fields("Code").Value & "','" & CellVal(7) & "','" & VchPrefix & "FI" & "','" & CellVal(5) & "','" & CellVal(6) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(4)) & ",Null," & SrNo & ",'" & CellVal(8) & "','" & CellVal(9) & "','" & CellVal(10) & "','" & CellVal(11) & "','" & CellVal(12) & "'," & Val(CellVal(3)) & ",'')"
        If Not CheckEmpty(CellVal(7), False) Then 'for ref
            CellVal(1) = IIf(Abs(Val(CellVal(1))) > BalQty, BalQty, Abs(Val(CellVal(1))))
            CellVal(1) = IIf(cmbBillType.ListIndex = 0, IIf(VchType = "SF", 0 - Val(CellVal(1)), Val(CellVal(1))), IIf(InStr(1, "PF_TF", VchType) > 0, 0 - Val(CellVal(1)), Val(CellVal(1)))) 'Direct : -ve for SF & +ve for PF, Against Challan : -ve for PF/TF & +ve for SF/OF
            cnSalesVoucher.Execute "INSERT INTO JobworkBVRef VALUES ('" & CellVal(7) & "',2,'" & VchPrefix & VchType & "','" & rstSalesVoucherParent.Fields("Code").Value & "','" & rstSalesVoucherParent.Fields("Name").Value & "','" & Format(rstSalesVoucherParent.Fields("Date").Value, "dd-MMM-yyyy") & "','" & PartyCode & "','" & CellVal(5) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ")"
        End If
    End If
    Exit Function
ErrorHandler:
    UpdateItemList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Party" Then
        rstSalesVoucherList.Filter = "[PartyName] Like '%" & SrchText & "%'"
    ElseIf SrchFor = "Material Centre" Then
        rstSalesVoucherList.Filter = "[MaterialCentreName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Item As Variant, i As Integer, x As Integer, cVal(1 To 6) As Variant, Disc As Double
    With fpSpread1
        If .EditMode Then Exit Sub
        If Shift = 0 And KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus: CalculateTotal
        ElseIf KeyCode = vbKeyF3 Then
            If .ActiveCol = 1 Then
                If cmbBillType.ListIndex = 1 Then Exit Sub 'Against Challan-New Item addition not allowed
                .GetText 10, .ActiveRow, Item 'Ref Order Code
                If Not CheckEmpty(Item, False) Then Exit Sub 'Old Item change not allowed
                .GetText 8, .ActiveRow, Item
                On Error Resume Next
                With FrmBookMaster
                    .SL = True
                    .ItemType = "F"
                    .MasterCode = Item
                    Load FrmBookMaster
                    If Err.Number <> 364 Then .Show vbModal
                End With
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText 8, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then
                    rstItemList.MoveFirst: rstItemList.Find "[Code] ='" & slCode & "'"
                    .GetText 5, .ActiveRow, Item 'Price
                    If Val(Item) = 0 Then
                        If Left(VchType, 1) = "S" Then
                            .SetText 5, .ActiveRow, Val(rstItemList.Fields("SalePrice").Value)
                        ElseIf Left(VchType, 1) = "P" Then
                            .SetText 5, .ActiveRow, Val(rstItemList.Fields("PurPrice").Value)
                        Else
                            .SetText 5, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                        End If
                    ElseIf Val(Item) <> Val(rstItemList.Fields("Price").Value) Then
                        If MsgBox("Variation in Current (" & Format(Item, "#0.00") & ") and Master (" & Format(rstItemList.Fields("Price").Value, "#0.00") & ") Rate ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
    '                            .SetText 5, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                            If Left(VchType, 1) = "S" Then
                                .SetText 5, .ActiveRow, Val(rstItemList.Fields("SalePrice").Value)
                            ElseIf Left(VchType, 1) = "P" Then
                                .SetText 5, .ActiveRow, Val(rstItemList.Fields("PurPrice").Value)
                            Else
                                .SetText 5, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                            End If
                        End If
                    End If
                    .GetText 9, .ActiveRow, Item 'HSN Code
                    If CheckEmpty(Item, False) Then .SetText 2, .ActiveRow, rstItemList.Fields("HSNName").Value: .SetText 9, .ActiveRow, rstItemList.Fields("HSNCode").Value
                    LoadMasterList
                    .SetFocus
                    cmbBillType.Enabled = False
                    Sendkeys "{ENTER}"
                End If
            End If
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then
                If cmbBillType.ListIndex = 1 Then Exit Sub 'Against Challan-New Item addition not allowed
                LoadMasterList True
                With FrmItemSearchList
                    Set .rstItemSearchList = rstItemList
                    Load FrmItemSearchList
                    .fpSpread1.SetActiveCell 3, 1
                    .Show vbModal
                    If .LoadItems Then
                        For i = 1 To .fpSpread1.DataRowCnt
                            .fpSpread1.GetText 1, i, cVal(1) 'Item
                            .fpSpread1.GetText 3, i, cVal(2) 'Quantity
                        If Left(VchType, 1) = "S" Then
                            .fpSpread1.GetText 9, i, cVal(3) 'Price
                        ElseIf Left(VchType, 1) = "P" Then
                            .fpSpread1.GetText 8, i, cVal(3) 'Price
                        Else
                            .fpSpread1.GetText 4, i, cVal(3) 'Price
                        End If
                            .fpSpread1.GetText 5, i, cVal(4) 'Item Code
                            .fpSpread1.GetText 6, i, cVal(5) 'HSN Code
                            .fpSpread1.GetText 7, i, cVal(6) 'HSN
                            If Val(cVal(2)) > 0 Then
                                x = fpSpread1.DataRowCnt + 1
                                fpSpread1.SetText 1, x, cVal(1)
                                fpSpread1.SetText 2, x, cVal(6)
                                fpSpread1.SetText 4, x, Val(cVal(2))
                                fpSpread1.SetText 5, x, Val(cVal(3))
                                fpSpread1.SetText 6, x, 0
                                fpSpread1.SetText 7, x, Val(cVal(2)) * Val(cVal(3))
                                fpSpread1.SetText 8, x, cVal(4)
                                fpSpread1.SetText 9, x, cVal(5)
                            End If
                        Next
                        CalculateTotal
                    End If
                End With
                Call CloseForm(FrmItemSearchList)
                .SetFocus
            ElseIf .ActiveCol = 2 Then
                .GetText 8, .ActiveRow, Item 'Item Code
                If CheckEmpty(Item, False) Then Exit Sub
                .GetText 9, .ActiveRow, Item 'HSN Code
                On Error Resume Next
                With FrmGeneralMaster
                    .SL = True
                    .MasterType = "18"
                    .MasterCode = Item
                    Load FrmGeneralMaster
                    If Err.Number <> 364 Then .Show vbModal
                End With
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText 9, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then LoadMasterList: Sendkeys "{ENTER}"
            End If
        ElseIf KeyCode = vbKeyReturn Then
            If .ActiveCol = 5 Then
                .GetText 8, .ActiveRow, Item
                Disc = FetchDiscount(PartyCode, Item)
                .GetText 6, .ActiveRow, Item 'Disc %
                If Val(Item) = 0 Then
                    .SetText 6, .ActiveRow, Disc
                ElseIf Val(Item) <> Disc Then
                    If MsgBox("Variation in Current (" & Format(Item, "#0.00") & ") and Master (" & Format(Disc, "#0.00") & ") Disc % ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 6, .ActiveRow, Disc
                    .SetFocus
                End If
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .ActiveCol >= 13 Then .SetText .ActiveCol, .ActiveRow, ""
        ElseIf KeyCode = vbKeyF11 Then
            If fpSpread1.DataRowCnt = 0 Then LoadOrderList
        End If
        If fpSpread1.DataRowCnt > 0 Then cmbBillType.Enabled = False Else cmbBillType.Enabled = True
    End With
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Item As Variant, Qty As Variant, Rate As Variant, Disc As Variant
    With fpSpread1
        If Col = 4 Or Col = 5 Or Col = 6 Then 'Qty, Disc % & Rate
            .GetText 8, Row, Item
            .GetText 4, Row, Qty
            .GetText 5, Row, Rate
            .GetText 6, Row, Disc
            Disc = (Rate * Disc) / 100
            If Not CheckEmpty(Item, False) Then .SetText 7, Row, Qty * Round((Rate - Disc), 2): CalculateTotal Else .SetText 4, Row, "": .SetText 5, Row, "": .SetText 6, Row, "": .SetText 6, Row, ""
        End If
    End With
End Sub
Private Sub fpSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As FPSpreadADO.TextTipFetchMultilineConstants, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim BalQty As Variant
    fpSpread1.GetText 12, Row, BalQty
    If Val(BalQty) = 0 Then Exit Sub
    If Col = 4 Then
        fpSpread1.SetTextTipAppearance "Calibri", 10, False, False, &HC0FFFF, &H80000008
        TipText = "Pending : " & Trim(BalQty)
        ShowTip = True
    End If
End Sub
Private Sub CalculateTotal()
    Dim i As Integer, Qty As Variant, Amt As Variant
    MhRealInput1.Value = 0: MhRealInput2.Value = 0
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 4, i, Qty
            .GetText 7, i, Amt
            MhRealInput1.Value = MhRealInput1.Value + Val(Qty)
            MhRealInput2.Value = MhRealInput2.Value + Val(Amt)
        Next
        MhRealInput3.Value = MhRealInput2.Value 'Pre-Tax Amount
        MhRealInput5.Value = (MhRealInput3.Value * MhRealInput4.Value) / 100 'Discount
        MhRealInput8.Value = ((MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value) * MhRealInput7.Value) / 100 'IGST/CGST
        MhRealInput10.Value = ((MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value) * MhRealInput9.Value) / 100 'SGST
        MhRealInput11.Value = MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput8.Value + MhRealInput10.Value + MhRealInput12.Value 'Post-Tax Amount
    End With
End Sub
Private Sub LoadMasterList(Optional ByVal LoadSelected As Boolean)
    If rstPartyList.State = adStateOpen Then rstPartyList.Close
    rstPartyList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnSalesVoucher, adOpenKeyset, adLockReadOnly
    rstPartyList.ActiveConnection = Nothing
    If rstMaterialCentreList.State = adStateOpen Then rstMaterialCentreList.Close
    rstMaterialCentreList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='*99999' ORDER BY Name", cnSalesVoucher, adOpenKeyset, adLockReadOnly
    rstMaterialCentreList.ActiveConnection = Nothing
    'SF-Sales Voucher PF-Purchase Voucher TF-Sales Return Voucher OF-Purchase Return Voucher
    If rstSalesTypeList.State = adStateOpen Then rstSalesTypeList.Close
    rstSalesTypeList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='" & IIf(InStr(1, "SF_TF", VchType) > 0, "*26027", "*26025") & "' ORDER BY Name", cnSalesVoucher, adOpenKeyset, adLockReadOnly
    If rstTaxList.State = adStateOpen Then rstTaxList.Close
    If PartyStateCode = "" Or PartyStateCode = Null Then
    rstTaxList.Open "SELECT Name As Col0,[IGST%],[SGST%],[CGST%],Region,Code FROM TaxMaster ORDER BY Name", cnSalesVoucher, adOpenKeyset, adLockReadOnly
    Else
    rstTaxList.Open "SELECT Name As Col0,[IGST%],[SGST%],[CGST%],Region,Code FROM TaxMaster Where Region='" & IIf(CompStateCode = PartyStateCode, "L", "I") & "' ORDER BY Name", cnSalesVoucher, adOpenKeyset, adLockReadOnly
    End If
    rstTaxList.ActiveConnection = Nothing
    If rstHSNCodeList.State = adStateOpen Then rstHSNCodeList.Close
    rstHSNCodeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='18' ORDER BY Name", cnSalesVoucher, adOpenKeyset, adLockReadOnly
    rstHSNCodeList.ActiveConnection = Nothing
    If rstItemList.State = adStateOpen Then rstItemList.Close
    If LoadSelected Then
    On Error Resume Next
'        rstItemList.Open "SELECT I.Name As Col0,FORMAT(dbo.ufnGetItemStock('" & MaterialCentreCode & "',I.Code,'" & Left(VchPrefix, 2) & "','" & CheckNull(rstSalesVoucherParent.Fields("Code").Value) & "','" & GetDate(MhDateInput1.Text) & "'),'#0') As Col1,0 As Quantity,I.Price,I.Code,H.Code As HSNCode,H.Name As HSNName FROM BookMaster I INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE I.Type='F' ORDER BY I.Name", cnSalesVoucher, adOpenKeyset, adLockReadOnly
If MsgBox("Do you want's Item List Display With closing Stock", vbInformation + vbDefaultButton2 + vbYesNo) = vbYes Then
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
                "),'#0') As Col1,0 As Quantity,I.Price,I.Code As code,H.Code As HSNCode,H.Name As HSNName ,I.PurPrice,I.SalePrice " & _
                " FROM (BookMaster I INNER Join GeneralMaster H ON H.Code=I.HSNCode)" & _
                "WHERE I.Type='F') As Tbl ORDER BY Col0 ASC", cnSalesVoucher, adOpenKeyset, adLockReadOnly
Else
        rstItemList.Open "SELECT * FROM(SELECT I.Name As Col0," & _
                "FORMAT(0,'#0') As Col1,0 As Quantity,I.Price,I.Code As code,H.Code As HSNCode,H.Name As HSNName,I.PurPrice,I.SalePrice " & _
                " FROM (BookMaster I INNER Join GeneralMaster H ON H.Code=I.HSNCode)" & _
                "WHERE I.Type='F') As Tbl ORDER BY Col0 ASC", cnSalesVoucher, adOpenKeyset, adLockReadOnly
End If
    If Err.Number = -2147217871 Then MsgBox "Due To Query Timeout. Unable To Fetch Stock !!!", vbInformation: rstItemList.Open "SELECT * FROM(SELECT I.Name As Col0,FORMAT(0,'#0') As Col1,0 As Quantity,I.Price,I.Code As code,H.Code As HSNCode,H.Name As HSNName  FROM (BookMaster I INNER Join GeneralMaster H ON H.Code=I.HSNCode)WHERE I.Type='F') As Tbl ORDER BY Col0 ASC", cnSalesVoucher, adOpenKeyset, adLockReadOnly
    Else
        rstItemList.Open "SELECT I.Name As Col0,FORMAT(0,'#0') As Col1,0 As Quantity,I.Price,I.Code,H.Name As HSNName,H.Code As HSNCode,I.PurPrice,I.SalePrice FROM BookMaster I INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE I.Type='F' ORDER BY I.Name", cnSalesVoucher, adOpenKeyset, adLockReadOnly
    End If
    rstItemList.ActiveConnection = Nothing
    If rstVchSeriesList.State = adStateOpen Then rstVchSeriesList.Close
    rstVchSeriesList.Open "SELECT Name As Col0,Prefix,Suffix,VchNumbering,Code,StartNo FROM VchSeriesMaster WHERE Left(FYCode,2)='" & Left(FYCode, 2) & "' AND VchType='" & IIf(VchType = "SF", "04", IIf(VchType = "PF", "01", IIf(VchType = "TF", "03", "02"))) & VchType & "' ORDER BY Name", cnSalesVoucher, adOpenKeyset, adLockReadOnly
    rstVchSeriesList.ActiveConnection = Nothing
End Sub
Private Sub LoadOrderList()
    If cmbBillType.ListIndex = 0 And InStr(1, "TF_OF", VchType) > 0 Then Exit Sub
    If rstOrderList.State = adStateOpen Then rstOrderList.Close
    If cmbBillType.ListIndex = 0 Then 'Direct
        rstOrderList.Open "SELECT VchCode,VchNo,VchDate,SUM(Quantity) As Ordered,SUM(Bal) As Bal FROM (SELECT VchCode,LTRIM(VchNo) As VchNo,VchDate,ABS(Quantity) As Quantity,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.RefCode)*1) As Bal FROM JobworkBVRef T WHERE RIGHT(VchType,2)='" & IIf(VchType = "SF", "SO", "PO") & "' AND Method=1 AND Party='" & PartyCode & "') As Tbl WHERE Bal>0 GROUP BY VchCode,VchNo,VchDate ORDER BY VchNo", cnSalesVoucher, adOpenKeyset, adLockReadOnly
    Else
        rstOrderList.Open "SELECT VchCode,VchNo,VchDate,SUM(Quantity) As Ordered,SUM(Bal) As Bal FROM (SELECT VchCode,LTRIM(VchNo) As VchNo,VchDate,ABS(Quantity) As Quantity,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.RefCode)*1) As Bal FROM JobworkBVRef T WHERE LEFT(VchType,2)+RIGHT(VchType,2)='" & IIf(VchType = "SF", "08IF", IIf(VchType = "PF", "05RF", IIf(VchType = "TF", "07RF", "06IF"))) & "' AND Method=1 AND Party='" & PartyCode & "') As Tbl WHERE Bal>0 GROUP BY VchCode,VchNo,VchDate ORDER BY VchNo", cnSalesVoucher, adOpenKeyset, adLockReadOnly ''Sales/Purchase Challan can have Method=1 & 2 both
    End If
    rstOrderList.ActiveConnection = Nothing
    If rstOrderList.RecordCount = 0 Then DisplayError ("No Pending Order Exists"): fpSpread1.SetFocus: Exit Sub
    With FrmOrderList.fpSpread1
        .ClearRange 1, 1, .MaxCols, .MaxRows, True
        .ColWidth(1) = 15.9
        .Col = 1: .Row = SpreadHeader: .Text = "Ref No."
        .Col = 2: .ColHidden = True 'Order No.
        .Col = 5: .Row = SpreadHeader + 1: .Text = " Delivered"
        .Col = 6: .ColHidden = True 'Unbilled
        .Col = 8: .ColHidden = True 'Status
        .Width = 6350
    End With
    FrmOrderList.Text1.Width = 6090
    Load FrmOrderList
    FrmOrderList.Text2 = Text3.Text
    With FrmOrderList.fpSpread1
        .ClearRange 1, 1, .MaxCols, .MaxRows, True
        .ColWidth(1) = 15.9
        .Col = 1: .Row = SpreadHeader: .Text = "Ref No."
        .Col = 2: .ColHidden = True 'Order No.
        .Col = 5: .Row = SpreadHeader + 1: .Text = " Delivered"
        .Col = 6: .ColHidden = True 'Unbilled
        .Col = 8: .ColHidden = True 'Status
        .Width = 6350
    End With
    FrmOrderList.Text1.Width = 6090
    Load FrmOrderList
    FrmOrderList.Text2 = Text3.Text
    Dim i As Integer, Disc As Double
    With rstOrderList
        For i = 1 To .RecordCount
            With FrmOrderList.fpSpread1
                .MaxRows = .MaxRows + 1
                .InsertRows i, 1
            End With
        Next
        i = 0
        Do Until .EOF
            i = i + 1
            FrmOrderList.fpSpread1.SetText 1, i, .Fields("VchNo").Value: FrmOrderList.fpSpread1.SetText 10, i, .Fields("VchCode").Value
            FrmOrderList.fpSpread1.SetText 3, i, Format(.Fields("VchDate").Value, "dd-MMM-yy")
            FrmOrderList.fpSpread1.SetText 4, i, Val(.Fields("Ordered").Value)
            FrmOrderList.fpSpread1.SetText 5, i, Val(.Fields("Ordered").Value) - Val(.Fields("Bal").Value) 'Delivered
            FrmOrderList.fpSpread1.SetText 7, i, Val(.Fields("Bal").Value) 'Pending
            FrmOrderList.fpSpread1.SetText 9, i, 0
            .MoveNext
        Loop
        FrmOrderList.fpSpread1.SetActiveCell 9, 1
    End With
    With FrmOrderList
        .Check1.Visible = False
        .Text2.Width = 4080
        .Check2 = 0: .Check2.Left = 5480
        .fpSpread1.Width = 15255 - 8905
        .Mh3dFrame2.Width = 15495 - 8925
        .cmdExit.Left = 15720 - 8935
        .Width = 16290 - 8935
        CenterForm FrmOrderList
        .Show vbModal
    End With
    If Not CheckEmpty(FrmOrderList.VchCodeList, False) Then
        If rstOrderList.State = adStateOpen Then rstOrderList.Close
        If cmbBillType.ListIndex = 0 Then 'Direct
            rstOrderList.Open "SELECT ItemCode,ItemName,HSNCode,OrderNo,OrderCode,Rate,[Disc%],SUM(Bal) As Bal FROM (SELECT I.Code As ItemCode,I.Name As ItemName,H.Code+'-'+H.Name As HSNCode,LTRIM(T.VchNo) As OrderNo,T.RefCode As OrderCode,T.Rate,T.[Disc%],ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.RefCode)*1) As Bal FROM (JobworkBVRef T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE RIGHT(VchType,2)='" & IIf(VchType = "SF", "SO", "PO") & "' AND Method=1 AND Party='" & PartyCode & "' AND T.Method=1 AND T.VchCode IN (" & FrmOrderList.VchCodeList & ")) As Tbl WHERE Bal>0 GROUP BY ItemCode,ItemName,HSNCode,OrderNo,OrderCode,Rate,[Disc%] ORDER BY ItemName,OrderNo", cnSalesVoucher, adOpenKeyset, adLockReadOnly
        Else
            rstOrderList.Open "SELECT ItemCode,ItemName,HSNCode,OrderNo,OrderCode,Rate,[Disc%],SUM(Bal) As Bal FROM (SELECT I.Code As ItemCode,I.Name As ItemName,H.Code+'-'+H.Name As HSNCode,LTRIM(T.VchNo) As OrderNo,T.RefCode As OrderCode,T.Rate,T.[Disc%],ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.RefCode)*1) As Bal FROM (JobworkBVRef T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE LEFT(VchType,2)+RIGHT(VchType,2)='" & IIf(VchType = "SF", "08IF", IIf(VchType = "PF", "05RF", IIf(VchType = "TF", "07RF", "06IF"))) & "' AND T.Method=1 AND T.VchCode IN (" & FrmOrderList.VchCodeList & ")) As Tbl WHERE Bal>0 GROUP BY ItemCode,ItemName,HSNCode,OrderNo,OrderCode,Rate,[Disc%] ORDER BY ItemName,OrderNo", cnSalesVoucher, adOpenKeyset, adLockReadOnly
        End If
        If rstOrderList.RecordCount > 0 Then
            i = 0
            With fpSpread1
                Do Until rstOrderList.EOF
                    i = i + 1
                    .SetText 1, i, rstOrderList.Fields("ItemName").Value
                    .SetText 2, i, Mid(rstOrderList.Fields("HSNCode").Value, InStr(1, rstOrderList.Fields("HSNCode").Value, "-") + 1, 40)
                    .SetText 3, i, rstOrderList.Fields("OrderNo").Value
                    .SetText 4, i, Val(rstOrderList.Fields("Bal").Value)
                    .SetText 5, i, Val(rstOrderList.Fields("Rate").Value)
                    .SetText 6, i, Val(rstOrderList.Fields("Disc%").Value)
                    Disc = (Val(rstOrderList.Fields("Rate").Value) * Val(rstOrderList.Fields("Disc%").Value)) / 100
                    .SetText 7, i, Val(rstOrderList.Fields("Bal").Value) * Round((Val(rstOrderList.Fields("Rate").Value) - Disc), 2)
                    .SetText 8, i, rstOrderList.Fields("ItemCode").Value
                    .SetText 9, i, Left(rstOrderList.Fields("HSNCode").Value, InStr(1, rstOrderList.Fields("HSNCode").Value, "-") - 1)
                    .SetText 10, i, rstOrderList.Fields("OrderCode").Value
                    .SetText 12, i, Val(rstOrderList.Fields("Bal").Value)
                    rstOrderList.MoveNext
                Loop
                Call CalculateTotal
            End With
            With rstOrderList
                If .State = adStateOpen Then .Close
                .Open "SELECT TOP 1 Transport,GRNo,GRDate,VehicleNo,Station FROM JobWorkBVParent WHERE Code IN (" & FrmOrderList.VchCodeList & ") AND (Transport<>'' AND Transport IS NOT NULL) ORDER BY Name", cnSalesVoucher, adOpenKeyset, adLockReadOnly
                If .RecordCount > 0 Then
                    If MsgBox("Do u want to update Transport Details from order ref ('" & CheckNull(.Fields("Transport").Value) & "','" & CheckNull(.Fields("GRNo").Value) & "','" & CheckNull(.Fields("VehicleNo").Value) & "','" & CheckNull(.Fields("Station").Value) & "')?", vbYesNo + vbQuestion + vbDefaultButton1, "Update Transport Details !") = vbYes Then frmSalesTptDetails.Text1.Text = CheckNull(.Fields("Transport").Value): frmSalesTptDetails.Text2.Text = CheckNull(.Fields("GRNo").Value): frmSalesTptDetails.Text3.Text = CheckNull(.Fields("VehicleNo").Value): frmSalesTptDetails.Text4.Text = CheckNull(.Fields("Station").Value): If Not IsNull(.Fields("GRDate").Value) Then frmSalesTptDetails.MhDateInput1.Value = .Fields("GRDate").Value
                End If
            End With
        End If
    End If
    fpSpread1.SetFocus
    CloseForm FrmOrderList
End Sub
Private Sub DuplicateRecord()
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    Dim VchCode As String, VchNo As String
    VchCode = GenerateCode(cnSalesVoucher, "SELECT MAX(Code) FROM JobworkBVParent", 6, "0")
    rstVchSeriesList.MoveFirst
    rstVchSeriesList.Find "[Code] = '" & rstSalesVoucherList.Fields("VchSeriesCode").Value & "'"
    AutoVchNo = GenerateCode(cnSalesVoucher, "SELECT MAX(CONVERT(INT,AutoVchNo))  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & rstSalesVoucherList.Fields("VchSeriesCode").Value & "' AND FYCode='" & FYCode & "'", 10, Space(1))
    VchNo = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
    With cnSalesVoucher
        .BeginTrans
        .Execute "SELECT * INTO #Tbl FROM JobworkBVParent WHERE Code = '" & rstSalesVoucherList.Fields("Code").Value & "'"
        .Execute "UPDATE #Tbl SET Code='" & VchCode & "',Name='" & Trim(VchNo) & "',AutoVchNo='" & Pad(Trim(AutoVchNo), Space(1), 10, "L") & "',[Date]=GETDATE()"
        .Execute "INSERT INTO JobworkBVParent SELECT * FROM #Tbl"
        .Execute "DROP TABLE #Tbl"
        .Execute "SELECT * INTO #Tbl FROM JobworkBVChild Where Code = '" & rstSalesVoucherList.Fields("Code").Value & "'"
        .Execute "UPDATE #Tbl SET Code='" & VchCode & "'"
        .Execute "UPDATE #Tbl SET Ref='',RefCode=''"
        .Execute "INSERT INTO JobworkBVChild SELECT * FROM #Tbl"
        .Execute "DROP TABLE #Tbl"
        .CommitTrans
        Me.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(6)
        Me.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(2)
        Me.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(4)
    End With
    MdiMainMenu.MousePointer = vbNormal
    Call MsgBox("Successfully Duplicated the Record !", vbInformation, App.Title)
    Exit Sub
ErrorHandler:
    cnSalesVoucher.RollbackTrans
    MdiMainMenu.MousePointer = vbNormal
    DisplayError ("Failed to Duplicate the Record")
End Sub
Private Sub btnNotes_Click()
    frmNotes.NotesFlag = 6
    frmNotes.Label1.Caption = "Notes : Voucher No. : " & Text2.Text
    frmNotes.Show (vbModal)
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Public Sub PrintSalesVoucher(ByVal VchCode As String, ByVal VchType As String, Optional ByVal OutputType As String)
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstSalesVoucherChild.State = adStateOpen Then rstSalesVoucherChild.Close
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,EMail,Website,GSTIN,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,Prefix,Suffix FROM CompanyMaster P INNER JOIN CompChild C ON P.Code=C.Code WHERE VchType= " & IIf(Left(VchType, 2) = "SF", 4, IIf(Left(VchType, 2) = "PF", 1, IIf(Left(VchType, 2) = "TF", 3, 2))), cnSalesVoucher, adOpenKeyset, adLockReadOnly
    rstCompanyMaster.ActiveConnection = Nothing
    rstSalesVoucherChild.Open "SELECT LTrim(P1.Name)+'/' +'" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "'  As BillNo,P1.Date As BillDate,A.PrintName As Party,A.Address1 As PartyAddress1,A.Address2 As PartyAddress2,A.Address3 As PartyAddress3,A.Address4 As PartyAddress4,A.TIN As PartyGSTIN,A.Mobile As Mobile,A.eMail As eMail,IIf(Right(P1.Type, 2) = 'SF',C.PrintName, IIf(Right(P1.Type, 2) = 'PF',C.PrintName, IIf(Right(P1.Type, 2) = 'TF',C.PrintName,C.PrintName))) As Consignee,IIf(Right(P1.Type, 2) = 'SF',C.Address1, IIf(Right(P1.Type, 2) = 'PF',C.Address1,IIf(Right(P1.Type, 2) = 'TF',C.Address1,C.Address1))) As ConsigneeAddress1,IIf(Right(P1.Type, 2) = 'SF',C.Address2, IIf(Right(P1.Type, 2) = 'PF',C.Address2, IIf(Right(P1.Type, 2) = 'TF',C.Address2,C.Address2))) As ConsigneeAddress2," & _
                                                "IIf(Right(P1.Type, 2) = 'SF',C.Address3, IIf(Right(P1.Type, 2) = 'PF',C.Address3, IIf(Right(P1.Type, 2) = 'TF',C.Address3,C.Address3))) As ConsigneeAddress3,IIf(Right(P1.Type, 2) = 'SF',C.Address4, IIf(Right(P1.Type, 2) = 'PF',C.Address4,IIf(Right(P1.Type, 2) = 'TF',C.Address4,C.Address4))) As ConsigneeAddress4,IIf(Right(P1.Type, 2) = 'SF',C.TIN, IIf(Right(P1.Type, 2) = 'PF',C.TIN, IIf(Right(P1.Type, 2) = 'TF',C.TIN,C.TIN))) As ConsigneeGSTIN,C.Mobile As CMobile,C.eMail As CeMail,P1.[Rebate%],P1.Rebate,P1.Freight,P1.Adjustment,P1.TaxableAmount,P1.[IGST%],P1.IGST,P1.[SGST%],P1.SGST,P1.[CGST%],P1.CGST,P1.Amount As TotalAmount,P1.Remarks,'' As Narration,I.PrintName As Item,H.PrintName As HSNCode," & _
                                                "C1.Quantity,C1.Rate,C1.Amount,N.Name As SrNo,'' As cmbTitle,LTRIM(C1.Code)+LTRIM(C1.SrNo) As Ref,C1.[Disc%] AS Disc,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,M.PrintName As MC,Transport,Station,ISNULL(eWayBill +'dt.'+ Convert(nvarchar,eWayBillDate),'') as eWayBill,ISNULL(GRNo +'dt.'+ Convert(nvarchar,GRDate),'') as GRNO FROM (((((((JobworkBVParent P1 INNER JOIN JobworkBVChild C1 ON P1.Code=C1.Code)INNER JOIN BookMaster I ON C1.Item=I.Code)INNER JOIN AccountMaster A ON P1.Party=A.Code)INNER JOIN AccountMaster C ON P1.Consignee=C.Code)LEFT JOIN AccountMaster M ON P1.MaterialCentre=M.Code)LEFT JOIN GeneralMaster N ON C1.Narration=N.Code)LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code)LEFT JOIN GeneralMaster S ON I.FinishSize=S.Code WHERE P1.Code='" + Left(VchCode, 6) + "' ORDER BY I.PrintName,N.Name", cnSalesVoucher, adOpenKeyset, adLockOptimistic
    If rstSalesVoucherChild.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rstSalesVoucherChild.ActiveConnection = Nothing
        With rptSalesOrderVoucher
        If Logo = "S" Then
            .Picture1.Width = LogoW
            .Picture1.Height = LogoH
        End If
        If Len(LTrim(rstCompanyMaster.Fields("PrintName").Value)) <= 30 Then
            .Text2.Font.Size = 20
        ElseIf Len(LTrim(rstCompanyMaster.Fields("PrintName").Value)) <= 40 Then
            .Text2.Font.Size = 18
        ElseIf Len(LTrim(rstCompanyMaster.Fields("PrintName").Value)) <= 50 Then
            .Text2.Font.Size = 16
        ElseIf Len(LTrim(rstCompanyMaster.Fields("PrintName").Value)) <= 60 Then
            .Text2.Font.Size = 14
        End If
        If LogoLine = "N" Then
            .Picture1.LeftLineStyle = crLSNoLine
            .Picture1.RightLineStyle = crLSNoLine
            .Picture1.TopLineStyle = crLSNoLine
            .Picture1.BottomLineStyle = crLSNoLine
        End If
    End With
    If Left(VchType, 2) = "SF" Then Load FrmDialog: Screen.MousePointer = vbNormal: FrmDialog.Flag = 1: FrmDialog.Caption = "":  FrmDialog.Command1.Caption = "Sales Invoice": FrmDialog.Command2.Caption = "Tax Invoice": FrmDialog.Command3.Caption = "Speciman Challan": FrmDialog.Command4.Caption = "Delivery Challan": FrmDialog.Command5.Visible = False: FrmDialog.Show vbModal
    If FrmDialog.Flag = 1 And PtgType > 0 Then
        rptSalesOrderVoucher.Text1.SetText IIf(PtgType = 1, "Sales Invoice", IIf(PtgType = 2, "Tax Invoice", IIf(PtgType = 3, "Speciman Challan", "Delivery Challan")))
        FrmDialog.Flag = 0: PtgType = 0
    Else
        rptSalesOrderVoucher.Text1.SetText IIf(Left(VchType, 2) = "SF", "Sales Invoice", IIf(Left(VchType, 2) = "PF", "Purchase Invoice ", IIf(Left(VchType, 2) = "TF", "Sales Returns ", "Purchase Return ")))
    End If
    rptSalesOrderVoucher.Text13.SetText IIf(Left(VchType, 2) = "SF", "Buyer :", IIf(Left(VchType, 2) = "PF", "Supplier :", IIf(Left(VchType, 2) = "TF", "Buyer :", "Supplier :")))
    rptSalesOrderVoucher.Text7.SetText IIf(Left(VchType, 2) = "SF", "Consignee : ", IIf(Left(VchType, 2) = "PF", "Consignee : ", IIf(Left(VchType, 2) = "TF", "Consignee : ", "Consignee :")))
    rptSalesOrderVoucher.Text35.SetText "Printed on " & Format(Now, "dd-MMM-yyyy") & " at " & Format(Now, "hh:mm")
    'rptSalesOrderVoucher.Text40.SetText IIf(BillType = "O", "(ORIGINAL FOR RECIPIENT)", IIf(BillType = "D", "(DUPLICATE FOR SUPPLIER)", "(TRIPLICATE FOR SUPPLIER)"))
    rptSalesOrderVoucher.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptSalesOrderVoucher.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False)) Then
        rptSalesOrderVoucher.Text4.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value) & Space(1) & "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptSalesOrderVoucher.Text4.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False) Then
        rptSalesOrderVoucher.Text4.SetText "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    End If
    rptSalesOrderVoucher.Text8.SetText "GSTIN/UIN : " & Trim(rstCompanyMaster.Fields("GSTIN").Value)
    rptSalesOrderVoucher.Text10.SetText "(" & UCase(Trim(NumberToWords(rstSalesVoucherChild.Fields("TotalAmount").Value, False))) & ")"
    rptSalesOrderVoucher.Text11.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptSalesOrderVoucher.Text26.SetText CheckNull(rstCompanyMaster.Fields("Declaration01").Value)
    rptSalesOrderVoucher.Text25.SetText CheckNull(rstCompanyMaster.Fields("Declaration02").Value)
    rptSalesOrderVoucher.Text22.SetText CheckNull(rstCompanyMaster.Fields("Declaration03").Value)
    rptSalesOrderVoucher.Text12.SetText CheckNull(rstCompanyMaster.Fields("Declaration04").Value)
    rptSalesOrderVoucher.Text9.SetText CheckNull(rstCompanyMaster.Fields("Declaration05").Value)
    rptSalesOrderVoucher.Text30.SetText CheckNull(rstCompanyMaster.Fields("Declaration06").Value)
    rptSalesOrderVoucher.Text31.SetText CheckNull(rstCompanyMaster.Fields("Declaration07").Value)
    'If Left(VchType, 2) <> "SF" Then rptSalesOrderVoucher.Text33.SetText ""
    'If Left(VchType, 2) <> "OF" Then
    rptSalesOrderVoucher.Text36.SetText (rstSalesVoucherChild.Fields("MC").Value)
    rptSalesOrderVoucher.Text44.SetText CheckNull(rstSalesVoucherChild.Fields("Transport").Value)
    rptSalesOrderVoucher.Text49.SetText CheckNull(rstSalesVoucherChild.Fields("Station").Value)
    rptSalesOrderVoucher.Text46.SetText CheckNull(rstSalesVoucherChild.Fields("eWayBill").Value)
    rptSalesOrderVoucher.Text51.SetText CheckNull(rstSalesVoucherChild.Fields("GRNO").Value)
    rptSalesOrderVoucher.Database.SetDataSource rstSalesVoucherChild, 3, 1
    rptSalesOrderVoucher.DiscardSavedData
    Screen.MousePointer = vbNormal
    If OutputType = "S" Then
        Set FrmReportViewer.Report = rptSalesOrderVoucher
        FrmReportViewer.Show vbModal
    Else
        If rstSalesVoucherList.State = adStateClosed Then  'For Print Utility
            rptSalesOrderVoucher.PaperSource = crPRBinAuto
            rptSalesOrderVoucher.PrintOut False
        Else
            rptSalesOrderVoucher.PaperSource = crPRBinAuto
            rptSalesOrderVoucher.PrintOut
        End If
    End If
    Set rptSalesOrderVoucher = Nothing
    If rstSalesVoucherList.State = adStateClosed Then  'For Print Utility
        Call CloseRecordset(rstCompanyMaster)
    End If
    Call CloseRecordset(rstSalesVoucherChild)
    On Error GoTo 0
End Sub
Private Sub PushVch()
    Dim XMLStr, SaleAccount, UOM, i
    With rstCompanyMaster
        If .State = adStateOpen Then .Close
        '.Open "SELECT VchSeries,Account,UOM FROM AppConfig WHERE VchType='SF'", cnSalesVoucher, adOpenKeyset, adLockReadOnly
        .Open "SELECT Name+'-'+VchNAme+'-'+Right(VchType,2) As VchSeries,VchNAme As Account,'No.' As UOM FROM VchSeriesMaster WHERE Right(VchType,2)='" & VchType & "'", cnSalesVoucher, adOpenKeyset, adLockReadOnly
        VchSeries = .Fields("VchSeries").Value: SaleAccount = .Fields("Account").Value: UOM = .Fields("UOM").Value
    End With
    With rstCompanyMaster
        If .State = adStateOpen Then .Close
        .Open "SELECT * FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnSalesVoucher, adOpenKeyset, adLockReadOnly
    End With
    With rstSalesVoucherChild
        If .State = adStateOpen Then .Close
        XMLStr = "SELECT LTRIM(H.Name) As BillNo,H.Date As BillDate,M.PrintName As MatCentre," + _
                        "B.PrintName As Buyer,Party As AccountCode,B.Address1 As bAddress1,B.Address2 As bAddress2,B.Address3 As bAddress3,B.Address4 As bAddress4,B.TIN As bGSTIN,C.PrintName As Consignee,C.Address1 As cAddress1,C.Address2 As cAddress2,C.Address3 As cAddress3,C.Address4 As cAddress4,C.TIN As cGSTIN," + _
                        "H.TaxableAmount,H.[Rebate%],H.Rebate,H.Freight,H.Adjustment,H.Tax,H.[IGST%],H.IGST,H.[SGST%],H.SGST,H.[CGST%],H.CGST,H.Amount As FinalAmount,H.Remarks," + _
                        "I.Name As ItemName,I.ItemIntegrationName As Item,I.BusyCode As ItemAlias,D.Rate,D.[Disc%],ABS(D.Quantity) As Quantity,D.Amount,H.Name As DeliveryNoteNo,H.GRNo As DispatchDocNo,H.Transport As DishpatchThrough,H.Station As Destination,H.Transport As [CarrierName/Agent],H.GRNo As [BillofLoading/LR-RRNO],H.GRDate,H.VehicleNo As MotorVehicleNo,H.eWayBill,H.eWayBillDate,IIF((Select Name From AccountMaster Where Code=H.SalesType)='Sales','Bill of Supply','Bill of Supply') As DocType,'Supply' As SubType,'Generated by me' As BILLSTATUS,D.Item As ItemCode,D.LongNarration01,D.LongNarration02,D.LongNarration03,D.LongNarration04,D.LongNarration05,ISNULL((SELECT LTRIM(VchNo) FROM JobworkBVRef WHERE RefCode=D.Ref AND RIGHT(VchType,2)='" & IIf(oVchType = "SF", "SO", IIf(oVchType = "PF", "PO", "")) & "'),'') As RefOrderNo " & _
                        "FROM ((((JobWorkBVParent H INNER JOIN AccountMaster B ON H.Party=B.Code) INNER JOIN AccountMaster C ON H.Consignee=C.Code) INNER JOIN AccountMaster M ON H.MaterialCentre=M.Code) INNER JOIN JobWorkBVChild D ON H.Code=D.Code) INNER JOIN BookMaster I ON D.Item=I.Code " + _
                        "WHERE H.Code='" + rstSalesVoucherList.Fields("Code").Value + "'"
        .Open XMLStr, cnSalesVoucher, adOpenKeyset, adLockReadOnly
        XMLStr = ""
        If TallyIntegration Then
            Dim Dom As Object
            Set Dom = CreateObject("MSXML2.DomDocument")
            Dom.async = False
            XMLStr = XMLStr + "<ENVELOPE>"
            XMLStr = XMLStr + "<HEADER>"
            XMLStr = XMLStr + "<TALLYREQUEST>Import Data</TALLYREQUEST>"
            XMLStr = XMLStr + "</HEADER>"
            XMLStr = XMLStr + "<BODY>"
            XMLStr = XMLStr + "<IMPORTDATA>"
            XMLStr = XMLStr + "<REQUESTDESC>"
            XMLStr = XMLStr + "<REPORTNAME>Vouchers</REPORTNAME>"
            XMLStr = XMLStr + "<STATICVARIABLES>"
            XMLStr = XMLStr + "<SVCURRENTCOMPANY>##SVCURRENTCOMPANY</SVCURRENTCOMPANY>" '##SVCURRENTCOMPANY-Current Open Company
            XMLStr = XMLStr + "</STATICVARIABLES>"
            XMLStr = XMLStr + "</REQUESTDESC>"
            XMLStr = XMLStr + "<REQUESTDATA>"
            XMLStr = XMLStr + "<TALLYMESSAGE xmlns:UDF=""TallyUDF"">"
            XMLStr = XMLStr + "<VOUCHER ACTION=""Create"">"
            XMLStr = XMLStr + "<VOUCHERTYPENAME>" + Replace(Trim(VchSeries), "&", "&amp;") + "</VOUCHERTYPENAME>" 'VchSeries
            XMLStr = XMLStr + "<VOUCHERNUMBER>" + Replace(Trim(.Fields("BillNo").Value), "&", "&amp;") + "</VOUCHERNUMBER>" 'Vch No.
            XMLStr = XMLStr + "<DATE>" + Format(.Fields("BillDate").Value, "yyyyMMdd") + "</DATE>" 'Vch Date
            If Not CheckEmpty(Trim(.Fields("Remarks").Value), False) Then XMLStr = XMLStr + "<NARRATION>" + Replace(Trim(.Fields("Remarks").Value), "&", "&amp;") + "</NARRATION>" 'Narration
            XMLStr = XMLStr + "<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>"
            XMLStr = XMLStr + "<ISINVOICE>Yes</ISINVOICE>"
            XMLStr = XMLStr + "<HASDISCOUNTS>Yes</HASDISCOUNTS>"
            'Buyer Info
            XMLStr = XMLStr + "<PARTYNAME>" + Replace(Trim(.Fields("Buyer").Value), "&", "&amp;") + "</PARTYNAME>"
            If Not CheckEmpty(Trim(.Fields("bAddress1").Value) + Trim(.Fields("bAddress2").Value) + Trim(.Fields("bAddress3").Value) + Trim(.Fields("bAddress4").Value), False) Then
                XMLStr = XMLStr + "<ADDRESS.LIST TYPE=""String"">"
                If Not CheckEmpty(Trim(.Fields("bAddress1").Value), False) Then XMLStr = XMLStr + "<ADDRESS>" + Replace(Trim(.Fields("bAddress1").Value), "&", "&amp;") + "</ADDRESS>"
                If Not CheckEmpty(Trim(.Fields("bAddress2").Value), False) Then XMLStr = XMLStr + "<ADDRESS>" + Replace(Trim(.Fields("bAddress2").Value), "&", "&amp;") + "</ADDRESS>"
                If Not CheckEmpty(Trim(.Fields("bAddress3").Value), False) Then XMLStr = XMLStr + "<ADDRESS>" + Replace(Trim(.Fields("bAddress3").Value), "&", "&amp;") + "</ADDRESS>"
                If Not CheckEmpty(Trim(.Fields("bAddress4").Value), False) Then XMLStr = XMLStr + "<ADDRESS>" + Replace(Trim(.Fields("bAddress4").Value), "&", "&amp;") + "</ADDRESS>"
                XMLStr = XMLStr + "</ADDRESS.LIST>"
            End If
            If Not CheckEmpty(Trim(.Fields("bGSTIN").Value), False) Then XMLStr = XMLStr + "<PARTYGSTIN>" + Replace(Trim(.Fields("bGSTIN").Value), "&", "&amp;") + "</PARTYGSTIN>"
            'Consignee Info
            XMLStr = XMLStr + "<BASICBUYERNAME>" + Replace(Trim(.Fields("Consignee").Value), "&", "&amp;") + "</BASICBUYERNAME>"
            If Not CheckEmpty(Trim(.Fields("cAddress1").Value) + Trim(.Fields("cAddress2").Value) + Trim(.Fields("cAddress3").Value) + Trim(.Fields("cAddress4").Value), False) Then
                XMLStr = XMLStr + "<BASICBUYERADDRESS.LIST TYPE=""String"">"
                If Not CheckEmpty(Trim(.Fields("cAddress1").Value), False) Then XMLStr = XMLStr + "<BASICBUYERADDRESS>" + Replace(Trim(.Fields("cAddress1").Value), "&", "&amp;") + "</BASICBUYERADDRESS>"
                If Not CheckEmpty(Trim(.Fields("cAddress2").Value), False) Then XMLStr = XMLStr + "<BASICBUYERADDRESS>" + Replace(Trim(.Fields("cAddress2").Value), "&", "&amp;") + "</BASICBUYERADDRESS>"
                If Not CheckEmpty(Trim(.Fields("cAddress3").Value), False) Then XMLStr = XMLStr + "<BASICBUYERADDRESS>" + Replace(Trim(.Fields("cAddress3").Value), "&", "&amp;") + "</BASICBUYERADDRESS>"
                If Not CheckEmpty(Trim(.Fields("cAddress4").Value), False) Then XMLStr = XMLStr + "<BASICBUYERADDRESS>" + Replace(Trim(.Fields("cAddress4").Value), "&", "&amp;") + "</BASICBUYERADDRESS>"
                XMLStr = XMLStr + "</BASICBUYERADDRESS.LIST>"
            End If
            
            If Not CheckEmpty(Trim(.Fields("cGSTIN").Value), False) Then XMLStr = XMLStr + "<CONSIGNEEGSTIN>" + Replace(Trim(.Fields("cGSTIN").Value), "&", "&amp;") + "</CONSIGNEEGSTIN>"
            'Dishpatch Details
            'Delivery Note No (s)
            If Not CheckEmpty(Trim(.Fields("DeliveryNoteNo").Value), False) Then XMLStr = XMLStr + "<BASICSHIPDELIVERYNOTE>" + Replace(Trim(.Fields("DeliveryNoteNo").Value), "&", "&amp;") + "</BASICSHIPDELIVERYNOTE>"
            'Delivery Note Date
            If Not CheckEmpty(Trim(.Fields("BillDate").Value), False) Then XMLStr = XMLStr + "<BASICSHIPPINGDATE>" + Replace(Trim(.Fields("BillDate").Value), "&", "&amp;") + "</BASICSHIPPINGDATE>"
            'Despatch Doc No
            If Not CheckEmpty(Trim(.Fields("BillNo").Value), False) Then XMLStr = XMLStr + "<BASICSHIPDOCUMENTNO>" + Replace(Trim(.Fields("BillNo").Value), "&", "&amp;") + "</BASICSHIPDOCUMENTNO>"
            'Despatched Through
            If Not CheckEmpty(Trim(.Fields("DishpatchThrough").Value), False) Then XMLStr = XMLStr + "<BASICSHIPPEDBY>" + Replace(Trim(.Fields("DishpatchThrough").Value), "&", "&amp;") + "</BASICSHIPPEDBY>"
            'Destination
            If Not CheckEmpty(Trim(.Fields("Destination").Value), False) Then XMLStr = XMLStr + "<BASICFINALDESTINATION>" + Replace(Trim(.Fields("Destination").Value), "&", "&amp;") + "</BASICFINALDESTINATION>"
            'CarrierName/Agent
            If Not CheckEmpty(Trim(.Fields("CarrierName/Agent").Value), False) Then XMLStr = XMLStr + "<EICHECKPOST>" + Replace(Trim(.Fields("CarrierName/Agent").Value), "&", "&amp;") + "</EICHECKPOST>"
            'BillofLoading/LR-RRNO
            If Not CheckEmpty(Trim(.Fields("BillofLoading/LR-RRNO").Value), False) Then XMLStr = XMLStr + "<BILLOFLADINGNO>" + Replace(Trim(.Fields("BillofLoading/LR-RRNO").Value), "&", "&amp;") + "</BILLOFLADINGNO>"
            'BILL OF LADING DATE
            If Not CheckEmpty(Trim(.Fields("GRDate").Value), False) Then XMLStr = XMLStr + "<BILLOFLADINGDATE>" + Replace(Trim(.Fields("GRDate").Value), "&", "&amp;") + "</BILLOFLADINGDATE>"
            'CarrierName/Agent
            If Not CheckEmpty(Trim(.Fields("MotorVehicleNo").Value), False) Then XMLStr = XMLStr + "<BASICSHIPVESSELNO>" + Replace(Trim(.Fields("MotorVehicleNo").Value), "&", "&amp;") + "</BASICSHIPVESSELNO>"
            XMLStr = XMLStr + "<REFERENCE>" + Replace(Trim(.Fields("BillNo").Value), "&", "&amp;") + "</REFERENCE>"
            'e-Way Bill Details
            If Not CheckEmpty(Trim(.Fields("eWayBill").Value), False) Then
            XMLStr = XMLStr + "<EWAYBILLDETAILS.LIST>"
            
            'CONSIGNOR Info
            XMLStr = XMLStr + "<CONSIGNORNAME>" + Replace(Trim(rstCompanyMaster.Fields("Name").Value), "&", "&amp;") + "</CONSIGNORNAME>" 'CONSIGNOR Info
            XMLStr = XMLStr + "<CONSIGNORPINCODE>" + Replace(Trim(rstCompanyMaster.Fields("Address4").Value), "&", "&amp;") + "</CONSIGNORPINCODE>"
            XMLStr = XMLStr + "<CONSIGNORGSTIN>" + Replace(Trim(rstCompanyMaster.Fields("GSTIN").Value), "&", "&amp;") + "</CONSIGNORGSTIN>"
            'xmlstr = xmlstr + "<CONSIGNORSTATENAME>" + Replace(Trim(rstCompanyMaster.Fields("Address3").Value), "&", "&amp;") + "</CONSIGNORSTATENAME>"
            If Not CheckEmpty(Trim(rstCompanyMaster.Fields("Address1").Value) + Trim(rstCompanyMaster.Fields("Address2").Value) + Trim(rstCompanyMaster.Fields("Address3").Value) + Trim(rstCompanyMaster.Fields("Address4").Value), False) Then
                XMLStr = XMLStr + "<CONSIGNORADDRESS.LIST TYPE=""String"">"
                If Not CheckEmpty(Trim(rstCompanyMaster.Fields("Address1").Value), False) Then XMLStr = XMLStr + "<CONSIGNORADDRESS>" + Replace(Trim(rstCompanyMaster.Fields("Address1").Value), "&", "&amp;") + "</CONSIGNORADDRESS>"
                If Not CheckEmpty(Trim(rstCompanyMaster.Fields("Address2").Value), False) Then XMLStr = XMLStr + "<CONSIGNORADDRESS>" + Replace(Trim(rstCompanyMaster.Fields("Address2").Value), "&", "&amp;") + "</CONSIGNORADDRESS>"
                If Not CheckEmpty(Trim(rstCompanyMaster.Fields("Address3").Value), False) Then XMLStr = XMLStr + "<CONSIGNORADDRESS>" + Replace(Trim(rstCompanyMaster.Fields("Address3").Value), "&", "&amp;") + "</CONSIGNORADDRESS>"
                If Not CheckEmpty(Trim(rstCompanyMaster.Fields("Address4").Value), False) Then XMLStr = XMLStr + "<CONSIGNORADDRESS>" + Replace(Trim(rstCompanyMaster.Fields("Address4").Value), "&", "&amp;") + "</CONSIGNORADDRESS>"
                XMLStr = XMLStr + "</CONSIGNORADDRESS.LIST>"
            End If
            
            'CONSIGNEE
            XMLStr = XMLStr + "<CONSIGNEENAME>" + Replace(Trim(.Fields("Buyer").Value), "&", "&amp;") + "</CONSIGNEENAME>" 'CONSIGNEE Info
            XMLStr = XMLStr + "<CONSIGNEEPINCODE>" + Replace(Trim(.Fields("bAddress4").Value), "&", "&amp;") + "</CONSIGNEEPINCODE>"
            XMLStr = XMLStr + "<CONSIGNEEGSTIN>" + Replace(Trim(.Fields("bGSTIN").Value), "&", "&amp;") + "</CONSIGNEEGSTIN>"
            'xmlstr = xmlstr + "<CONSIGNEESTATENAME>" + Replace(Trim(.Fields("bAddress3").Value), "&", "&amp;") + "</CONSIGNEESTATENAME>"

            If Not CheckEmpty(Trim(.Fields("bAddress1").Value) + Trim(.Fields("bAddress2").Value) + Trim(.Fields("bAddress3").Value) + Trim(.Fields("bAddress4").Value), False) Then
                XMLStr = XMLStr + "<CONSIGNEEADDRESS.LIST TYPE=""String"">"
                If Not CheckEmpty(Trim(.Fields("bAddress1").Value), False) Then XMLStr = XMLStr + "<CONSIGNEEADDRESS>" + Replace(Trim(.Fields("bAddress1").Value), "&", "&amp;") + "</CONSIGNEEADDRESS>"
                If Not CheckEmpty(Trim(.Fields("bAddress2").Value), False) Then XMLStr = XMLStr + "<CONSIGNEEADDRESS>" + Replace(Trim(.Fields("bAddress2").Value), "&", "&amp;") + "</CONSIGNEEADDRESS>"
                If Not CheckEmpty(Trim(.Fields("bAddress3").Value), False) Then XMLStr = XMLStr + "<CONSIGNEEADDRESS>" + Replace(Trim(.Fields("bAddress3").Value), "&", "&amp;") + "</CONSIGNEEADDRESS>"
                If Not CheckEmpty(Trim(.Fields("bAddress4").Value), False) Then XMLStr = XMLStr + "<CONSIGNEEADDRESS>" + Replace(Trim(.Fields("bAddress4").Value), "&", "&amp;") + "</CONSIGNEEADDRESS>"
                XMLStr = XMLStr + "</CONSIGNEEADDRESS.LIST>"
            End If
            
                XMLStr = XMLStr + "<DOCUMENTTYPE>" + Replace(Trim(.Fields("DocType").Value), "&", "&amp;") + "</DOCUMENTTYPE>"
                XMLStr = XMLStr + "<SUBTYPE>" + Replace(Trim(.Fields("SubType").Value), "&", "&amp;") + "</SUBTYPE>"
                XMLStr = XMLStr + "<BILLSTATUS>" + Replace(Trim(.Fields("BILLSTATUS").Value), "&", "&amp;") + "</BILLSTATUS>"
                
            If Not CheckEmpty(Trim(.Fields("eWayBill").Value), False) Then XMLStr = XMLStr + "<BILLNUMBER>" + Replace(Trim(.Fields("eWayBill").Value), "&", "&amp;") + "</BILLNUMBER>"
            If Not CheckEmpty(Trim(.Fields("eWayBillDate").Value), False) Then XMLStr = XMLStr + "<BILLDATE>" + Replace(Trim(.Fields("eWayBillDate").Value), "&", "&amp;") + "</BILLDATE>"
            
            XMLStr = XMLStr + "<TRANSPORTDETAILS.LIST>"
            
            XMLStr = XMLStr + "<DOCUMENTDATE>" + Replace(Trim(.Fields("BillDate").Value), "&", "&amp;") + "</DOCUMENTDATE>"
            XMLStr = XMLStr + "<BASICSHIPPEDBY>" + Replace(Trim(.Fields("CarrierName/Agent").Value), "&", "&amp;") + "</BASICSHIPPEDBY>"
            XMLStr = XMLStr + "<TRANSPORTERID>" + Replace(Trim(.Fields("CarrierName/Agent").Value), "&", "&amp;") + "</TRANSPORTERID>"
            'xmlstr = xmlstr + "<TRANSPORTMODE>" + Replace(Trim(.Fields("TRANSPORTMODE").Value), "&", "&amp;") + "</TRANSPORTMODE>" 'Air,Rail,Road,Ship
            XMLStr = XMLStr + "<VEHICLENUMBER>" + Replace(Trim(.Fields("MotorVehicleNo").Value), "&", "&amp;") + "</VEHICLENUMBER>"
            XMLStr = XMLStr + "<DOCUMENTNUMBER>" + Replace(Trim(.Fields("BillofLoading/LR-RRNO").Value), "&", "&amp;") + "</DOCUMENTNUMBER>"
            'xmlstr = xmlstr + "<VEHICLETYPE>" + Replace(Trim(.Fields("VEHICLETYPE").Value), "&", "&amp;") + "</VEHICLETYPE>"   'Regular
            'xmlstr = xmlstr + "<DISTANCE>" + Replace(Trim(.Fields("DISTANCE").Value), "&", "&amp;") + "</DISTANCE>"
            
            XMLStr = XMLStr + "</TRANSPORTDETAILS.LIST>"
            
            XMLStr = XMLStr + "</EWAYBILLDETAILS.LIST>"
            End If
            .MoveFirst
            Do Until .EOF
                XMLStr = XMLStr + "<INVENTORYENTRIES.LIST>"
                
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION.LIST TYPE=""String"">"
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + "StockItem:" + Replace(Trim(.Fields("ItemName").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'LongNarration02
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + Replace(Trim(.Fields("LongNarration01").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'LongNarration02
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + Replace(Trim(.Fields("LongNarration02").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'LongNarration02
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + Replace(Trim(.Fields("LongNarration03").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'LongNarration03
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + Replace(Trim(.Fields("LongNarration04").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'LongNarration04
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + Replace(Trim(.Fields("LongNarration05").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'LongNarration05
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + "Ref. No.:" + Replace(Trim(.Fields("RefOrderNo").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'RefOrderNo
                XMLStr = XMLStr + "</BASICUSERDESCRIPTION.LIST>"
                
                XMLStr = XMLStr + "<STOCKITEMNAME>" + Replace(Trim(.Fields("Item").Value), "&", "&amp;") + "</STOCKITEMNAME>"
                
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<RATE>" + Format(Val(.Fields("Rate").Value), "0.00") + "/" + Replace(UOM, "&", "&amp;") + "</RATE>"
                XMLStr = XMLStr + "<DISCOUNT>" + Format(Val(.Fields("Disc%").Value), "0.00") + "</DISCOUNT>"
                XMLStr = XMLStr + "<AMOUNT>" + Format(Val(.Fields("Amount").Value), "0.00") + "</AMOUNT>"
                XMLStr = XMLStr + "<ACTUALQTY>" + Format(Val(.Fields("Quantity").Value), "0.00") + " " + Replace(UOM, "&", "&amp;") + "</ACTUALQTY>"
                XMLStr = XMLStr + "<BILLEDQTY>" + Format(Val(.Fields("Quantity").Value), "0.00") + " " + Replace(UOM, "&", "&amp;") + "</BILLEDQTY>"
                XMLStr = XMLStr + "<BATCHALLOCATIONS.LIST>"
                XMLStr = XMLStr + "<GODOWNNAME>" + Replace(Trim(.Fields("MatCentre").Value), "&", "&amp;") + "</GODOWNNAME>"
                XMLStr = XMLStr + "<DESTINATIONGODOWNNAME>" + Replace(Trim(.Fields("MatCentre").Value), "&", "&amp;") + "</DESTINATIONGODOWNNAME>"
                XMLStr = XMLStr + "</BATCHALLOCATIONS.LIST>"
                XMLStr = XMLStr + "<ACCOUNTINGALLOCATIONS.LIST>"
                XMLStr = XMLStr + "<LEDGERNAME>" + Replace(SaleAccount, "&", "&amp;") + "</LEDGERNAME>"
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                XMLStr = XMLStr + "<AMOUNT>" + Format(Val(.Fields("Amount").Value), "0.00") + "</AMOUNT>"
                XMLStr = XMLStr + "</ACCOUNTINGALLOCATIONS.LIST>"
                XMLStr = XMLStr + "</INVENTORYENTRIES.LIST>"
                
                .MoveNext
            Loop
            .MoveFirst
            XMLStr = XMLStr + "<LEDGERENTRIES.LIST>"
            XMLStr = XMLStr + "<LEDGERNAME>" + Replace(Trim(.Fields("Buyer").Value), "&", "&amp;") + "</LEDGERNAME>" 'Buyer Name
            XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>"
            XMLStr = XMLStr + "<ISPARTYLEDGER>Yes</ISPARTYLEDGER>"
            XMLStr = XMLStr + "<AMOUNT>" + Trim(0 - Val(.Fields("FinalAmount").Value)) + "</AMOUNT>" 'Vch Amount
            XMLStr = XMLStr + "<BILLALLOCATIONS.LIST>"
            XMLStr = XMLStr + "<NAME>" + Replace(Trim(.Fields("BillNo").Value), "&", "&amp;") + "</NAME>" 'Vch No.
            XMLStr = XMLStr + "<BILLTYPE></BILLTYPE>"
        If VchType <> "PF" Then
            XMLStr = XMLStr + "<AMOUNT>" + Trim(0 - Val(.Fields("FinalAmount").Value)) + "</AMOUNT>" 'Vch Amount
        ElseIf VchType = "PF" Then
            XMLStr = XMLStr + "<AMOUNT>" + Trim(Val(.Fields("FinalAmount").Value)) + "</AMOUNT>" 'Vch Amount
        End If
            XMLStr = XMLStr + "</BILLALLOCATIONS.LIST>"
            XMLStr = XMLStr + "</LEDGERENTRIES.LIST>"
            If Val(.Fields("Rebate").Value) > 0 Then
                XMLStr = XMLStr + "<LEDGERENTRIES.LIST>"
                XMLStr = XMLStr + "<BASICRATEOFINVOICETAX.LIST TYPE=""Number"">"
                XMLStr = XMLStr + "<BASICRATEOFINVOICETAX> " + Trim(Format(0 - Val(.Fields("Rebate%").Value), "0.00")) + "</BASICRATEOFINVOICETAX>"
                XMLStr = XMLStr + "</BASICRATEOFINVOICETAX.LIST>"
                XMLStr = XMLStr + "<LEDGERNAME>Discount</LEDGERNAME>"
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                XMLStr = XMLStr + "<AMOUNT>" + Trim(Format(0 - Val(.Fields("Rebate").Value), "0.00")) + "</AMOUNT>"
                XMLStr = XMLStr + "<VATEXPAMOUNT>" + Trim(0 - Format(Val(.Fields("Rebate").Value), "0.00")) + "</VATEXPAMOUNT>"
                XMLStr = XMLStr + "</LEDGERENTRIES.LIST>"
            End If
            If Val(.Fields("Freight").Value) > 0 Then
                XMLStr = XMLStr + "<LEDGERENTRIES.LIST>"
                XMLStr = XMLStr + "<LEDGERNAME>" + Replace("Packing & Forwarding Charges", "&", "&amp;") + "</LEDGERNAME>"
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                XMLStr = XMLStr + "<AMOUNT>" + Format(Val(.Fields("Freight").Value), "0.00") + "</AMOUNT>"
                XMLStr = XMLStr + "<VATEXPAMOUNT>" + Format(Val(.Fields("Freight").Value), "0.00") + "</VATEXPAMOUNT>"
                XMLStr = XMLStr + "</LEDGERENTRIES.LIST>"
            End If
            rstTaxList.MoveFirst
            rstTaxList.Find "[Code]='" & .Fields("Tax").Value & "'"
            If rstTaxList.Fields("Region").Value = "I" Then
                XMLStr = XMLStr + "<LEDGERENTRIES.LIST>"
                XMLStr = XMLStr + "<BASICRATEOFINVOICETAX.LIST TYPE=""Number"">"
                XMLStr = XMLStr + "<BASICRATEOFINVOICETAX> " + Trim(Format(Val(.Fields("IGST%").Value), "0.00")) + "</BASICRATEOFINVOICETAX>"
                XMLStr = XMLStr + "</BASICRATEOFINVOICETAX.LIST>"
                XMLStr = XMLStr + "<LEDGERNAME>IGST-" + IIf(Val(.Fields("IGST%").Value) = 0, "Exempted", Trim(Format(Val(.Fields("IGST%").Value), "0.00")) + "%") + "</LEDGERNAME>"
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                XMLStr = XMLStr + "<AMOUNT>" + Trim(Format(Val(.Fields("IGST").Value), "0.00")) + "</AMOUNT>"
                XMLStr = XMLStr + "<VATEXPAMOUNT>" + Trim(Format(Val(.Fields("IGST").Value), "0.00")) + "</VATEXPAMOUNT>"
                XMLStr = XMLStr + "</LEDGERENTRIES.LIST>"
            Else
                XMLStr = XMLStr + "<LEDGERENTRIES.LIST>"
                XMLStr = XMLStr + "<BASICRATEOFINVOICETAX.LIST TYPE=""Number"">"
                XMLStr = XMLStr + "<BASICRATEOFINVOICETAX> " + Trim(Format(Val(.Fields("CGST%").Value), "0.00")) + "</BASICRATEOFINVOICETAX>"
                XMLStr = XMLStr + "</BASICRATEOFINVOICETAX.LIST>"
                XMLStr = XMLStr + "<LEDGERNAME>CGST-" + IIf(Val(.Fields("CGST%").Value) = 0, "Exempted", Trim(Format(Val(.Fields("CGST%").Value), "0.00")) + "%") + "</LEDGERNAME>"
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                XMLStr = XMLStr + "<AMOUNT>" + Trim(Format(Val(.Fields("CGST").Value), "0.00")) + "</AMOUNT>"
                XMLStr = XMLStr + "<VATEXPAMOUNT>" + Trim(Format(Val(.Fields("CGST").Value), "0.00")) + "</VATEXPAMOUNT>"
                XMLStr = XMLStr + "</LEDGERENTRIES.LIST>"
                XMLStr = XMLStr + "<LEDGERENTRIES.LIST>"
                XMLStr = XMLStr + "<BASICRATEOFINVOICETAX.LIST TYPE=""Number"">"
                XMLStr = XMLStr + "<BASICRATEOFINVOICETAX> " + Trim(Format(Val(.Fields("SGST%").Value), "0.00")) + "</BASICRATEOFINVOICETAX>"
                XMLStr = XMLStr + "</BASICRATEOFINVOICETAX.LIST>"
                XMLStr = XMLStr + "<LEDGERNAME>SGST-" + IIf(Val(.Fields("SGST%").Value) = 0, "Exempted", Trim(Format(Val(.Fields("SGST%").Value), "0.00")) + "%") + "</LEDGERNAME>"
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                XMLStr = XMLStr + "<AMOUNT>" + Trim(Format(Val(.Fields("SGST").Value), "0.00")) + "</AMOUNT>"
                XMLStr = XMLStr + "<VATEXPAMOUNT>" + Trim(Format(Val(.Fields("SGST").Value), "0.00")) + "</VATEXPAMOUNT>"
                XMLStr = XMLStr + "</LEDGERENTRIES.LIST>"
            End If
            If Val(.Fields("Adjustment").Value) <> 0 Then
                XMLStr = XMLStr + "<LEDGERENTRIES.LIST>"
                XMLStr = XMLStr + "<LEDGERNAME>Round Off</LEDGERNAME>"
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                XMLStr = XMLStr + "<AMOUNT>" + Format(Val(.Fields("Adjustment").Value), "0.00") + "</AMOUNT>"
                XMLStr = XMLStr + "<VATEXPAMOUNT>" + Format(Val(.Fields("Adjustment").Value), "0.00") + "</VATEXPAMOUNT>"
                XMLStr = XMLStr + "</LEDGERENTRIES.LIST>"
            End If
            XMLStr = XMLStr + "</VOUCHER>"
            XMLStr = XMLStr + "</TALLYMESSAGE>"
            XMLStr = XMLStr + "</REQUESTDATA>"
            XMLStr = XMLStr + "</IMPORTDATA>"
            XMLStr = XMLStr + "</BODY>"
            XMLStr = XMLStr + "</ENVELOPE>"
            Dim WinHttpReq As Object
            Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
            With WinHttpReq
                .Open "POST", "http://localhost:" + ReadFromFile("Tally Port"), False
                Do While True
                    On Error Resume Next
                    DelOldVch False
                    .Send XMLStr
                    If Err.Number = -2147012867 Then
                        If MsgBox(Err.Description + "Tally is not open. Would you like to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Proceed !") = vbNo Then Exit Do
                    Else
                        .waitForResponse 4000
                        Dom.loadXML .responseText
                        If Dom.selectSingleNode("//CREATED").Text = "1" Then
                            MsgBox "Voucher Exported to Tally !!!", vbInformation, App.Title: UpdateIntegration: Exit Do
                        Else
                            If MsgBox(Dom.selectSingleNode("//LINEERROR").Text + " Would you like to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Proceed !") = vbNo Then Exit Do
                        End If
                    End If
                Loop
            End With
            
        ElseIf BusyIntegration Then
            Dim FI  As CFixedInterface
            Dim rstBusy As Dao.Recordset
            Dim cnBusy As New ADODB.Connection
            Dim VchSeriesName, VchDate, VchNo, STName, AccountCode, AccountName, MCName
            Dim ItemCode, ItemName, Qty, Price
            Set FI = CreateObject("Busy2L16.CFixedInterface")
            Dim BusyServerPassword, BusyServerUser, BusyDatabaseName, BusyServerName, BusyConnectionString As String
            BusyServerPassword = Trim(ReadFromFile("Busy Server Password"))
            BusyServerUser = "sa"
            BusyDatabaseName = "BusyComp0007_db"
            BusyServerName = "182.71.145.139"
            cnBusy.CursorLocation = adUseClient
        If DatabaseType = "MS SQL" Then
        If cnBusy.State = adStateOpen Then cnBusy.Close
            cnBusy.CommandTimeout = 300
            BusyConnectionString = "Provider=SQLOLEDB.1;Password=" & BusyServerPassword & ";Persist Security Info=True;User ID=" & BusyServerUser & ";Initial Catalog=" & BusyDatabaseName & ";Data Source=" & BusyServerName
        Else
            BusyConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Printwell Printers International\PPI Busy win Data\DATA\COMP0006\db12022.bds;Persist Security Info=True;Jet OLEDB:Database Password=ILoveMyINDIA"
        End If
            If cnBusy.State = adStateOpen Then cnBusy.Close
            cnBusy.Open BusyConnectionString
                        
            If DatabaseType = "MS SQL" Then
            
            FI.OpenCSDB Trim(ReadFromFile("Busy Path")), Trim(ReadFromFile("Busy Server Name")), "sa", Trim(ReadFromFile("Busy Server Password")), cnBusy.DefaultDatabase 'Mid(cnBusy.DefaultDatabase, 8, 12)
            ElseIf DatabaseType = "MS Access" Then
            FI.OpenDB (ReadFromFile("Busy Path")), Trim(ReadFromFile("Busy Data Path")), Trim(ReadFromFile("Busy CompCode")), Trim(ReadFromFile("Busy FY"))
            End If
            AccountCode = Trim(.Fields("AccountCode").Value)
            If CheckEmpty(AccountCode, False) Then Set FI = Nothing: Exit Sub
            VchSeriesName = VchSeries: MCName = Trim(.Fields("MatCentre").Value): VchNo = Trim(Text2.Text): VchDate = FI.FormatDate(MhDateInput1.Value): Qty = Val(MhRealInput1.Value)
            Set rstBusy = FI.GetRecordset("SELECT Name,GSTNo FROM Master1 P INNER JOIN MasterAddressInfo C ON P.Code=C.MasterCode WHERE Code=" & AccountCode)
            Set rstBusy = FI.GetRecordset("SELECT Name,GSTNo FROM Master1 P INNER JOIN MasterAddressInfo C ON P.Code=C.MasterCode")
            Set rstBusy = FI.GetRecordset("SELECT * FROM Master1  ")
            
            AccountName = Replace(rstBusy.Fields("Name").Value, "&", "&amp;", 1)
            STName = IIf(Left(rstBusy.Fields("GSTNo").Value, 2) = "07", "L/GST-Exempt", "I/GST-Exempt")
            ItemCode = Mid(ItemCode, 2, 6)
            Set rstBusy = FI.GetRecordset("SELECT Name,D3 As Price FROM Master1 WHERE Code=" & ItemCode)
            ItemName = Replace(rstBusy.Fields("Name").Value, "&", "&amp;", 1): Price = Val(rstBusy.Fields("Price").Value)
            XMLStr = "<PurchaseOrder>"
                XMLStr = XMLStr & "<VchSeriesName>" & VchSeriesName & "</VchSeriesName><Date>" & VchDate & "</Date><VchType>13</VchType><VchNo>" & VchNo & "</VchNo><STPTName>" & STName & "</STPTName><MasterName1>" & AccountName & "</MasterName1><MasterName2>" & MCName & "</MasterName2>"
                XMLStr = XMLStr & "<ItemEntries>"
                XMLStr = XMLStr & "<ItemDetail><SrNo>1</SrNo><ItemName>" & ItemName & "</ItemName><UnitName>Nos</UnitName><Qty>" & Trim(Qty) & "</Qty><QtyMainUnit>" & Trim(Qty) & "</QtyMainUnit><QtyAltUnit>" & Trim(Qty) & "</QtyAltUnit><Price>" & Trim(Price) & "</Price><Amt>" & Trim(Qty * Price) & "</Amt><STAmount>0</STAmount><STPercent>0</STPercent><TaxBeforeSurcharge>0</TaxBeforeSurcharge><MC>" & MCName & "</MC></ItemDetail>"
                XMLStr = XMLStr & "</ItemEntries>"
                XMLStr = XMLStr & "<PendingOrders>"
                    XMLStr = XMLStr & "<OrderDetail><MasterName1>" & ItemName & "</MasterName1><MasterName2>" & AccountName & "</MasterName2>"
                    XMLStr = XMLStr & "<OrderRefs><Method>1</Method><SrNo>1</SrNo><RefNo>" & VchNo & "</RefNo><Date>" & VchDate & "</Date><DueDate>" & VchDate & "</DueDate><Value1>" & Trim(0 - Qty) & "</Value1><Value2>" & Trim(0 - Qty) & "</Value2><ItemSrNo>1</ItemSrNo><tmpMasterCode1>" & Trim(ItemCode) & "</tmpMasterCode1><tmpMasterCode2>" & Trim(AccountCode) & "</tmpMasterCode2></OrderRefs>"
                    XMLStr = XMLStr & "</OrderDetail>"
                XMLStr = XMLStr & "</PendingOrders>"
            XMLStr = XMLStr & "</PurchaseOrder>"
            If Not FI.SaveVchFromXML(13, XMLStr, Err, True, 2) Then DisplayError (Err)
'            Dim VchSeriesName, VchDate, VchNo, STName, AccountCode, AccountName, MCName, XMLStr
'            Dim ItemCode, ItemName, Qty, Price
'            AccountCode = IIf(Not CheckEmpty(BookPrinterCode, False), BookPrinterCode, BinderCode)
'            If CheckEmpty(AccountCode, False) Then Set FI = Nothing: Exit Sub
'            VchSeriesName = IIf(BookPOType = "F", "Main", "Repair"): MCName = "Noida Godown": VchNo = Trim(Text2.Text): VchDate = FI.FormatDate(rstBookPOChild08.Fields("OrderDate").Value): Qty = Val(rstBookPOChild08.Fields("ActualQuantity").Value)
'            Set rstSaral = FI.GetRecordset("SELECT Name,GSTNo FROM Master1 P INNER JOIN MasterAddressInfo C ON P.Code=C.MasterCode WHERE Code=" & AccountCode)
'            AccountName = Replace(rstSaral.Fields("Name").Value, "&", "&amp;", 1)
'            STName = IIf(Left(rstSaral.Fields("GSTNo").Value, 2) = "07", "L/GST-Exempt", "I/GST-Exempt")
'            ItemCode = Mid(BookCode, 2, 6)
'            Set rstSaral = FI.GetRecordset("SELECT Name,D3 As Price FROM Master1 WHERE Code=" & ItemCode)
'            ItemName = Replace(rstSaral.Fields("Name").Value, "&", "&amp;", 1): Price = Val(rstSaral.Fields("Price").Value)
'            XMLStr = "<PurchaseOrder>"
'                XMLStr = XMLStr & "<VchSeriesName>" & VchSeriesName & "</VchSeriesName><Date>" & VchDate & "</Date><VchType>13</VchType><VchNo>" & VchNo & "</VchNo><STPTName>" & STName & "</STPTName><MasterName1>" & AccountName & "</MasterName1><MasterName2>" & MCName & "</MasterName2>"
'                XMLStr = XMLStr & "<ItemEntries>"
'                XMLStr = XMLStr & "<ItemDetail><SrNo>1</SrNo><ItemName>" & ItemName & "</ItemName><UnitName>Nos</UnitName><Qty>" & Trim(Qty) & "</Qty><QtyMainUnit>" & Trim(Qty) & "</QtyMainUnit><QtyAltUnit>" & Trim(Qty) & "</QtyAltUnit><Price>" & Trim(Price) & "</Price><Amt>" & Trim(Qty * Price) & "</Amt><STAmount>0</STAmount><STPercent>0</STPercent><TaxBeforeSurcharge>0</TaxBeforeSurcharge><MC>" & MCName & "</MC></ItemDetail>"
'                XMLStr = XMLStr & "</ItemEntries>"
'                XMLStr = XMLStr & "<PendingOrders>"
'                    XMLStr = XMLStr & "<OrderDetail><MasterName1>" & ItemName & "</MasterName1><MasterName2>" & AccountName & "</MasterName2>"
'                    XMLStr = XMLStr & "<OrderRefs><Method>1</Method><SrNo>1</SrNo><RefNo>" & VchNo & "</RefNo><Date>" & VchDate & "</Date><DueDate>" & VchDate & "</DueDate><Value1>" & Trim(0 - Qty) & "</Value1><Value2>" & Trim(0 - Qty) & "</Value2><ItemSrNo>1</ItemSrNo><tmpMasterCode1>" & Trim(ItemCode) & "</tmpMasterCode1><tmpMasterCode2>" & Trim(AccountCode) & "</tmpMasterCode2></OrderRefs>"
'                    XMLStr = XMLStr & "</OrderDetail>"
'                XMLStr = XMLStr & "</PendingOrders>"
'            XMLStr = XMLStr & "</PurchaseOrder>"
'            If Not FI.SaveVchFromXML(13, XMLStr, ErrMsg, True, 2) Then DisplayError (ErrMsg)
        End If
    End With
End Sub
Private Sub DelOldVch(ByVal dspMsg As Boolean)
    Dim XMLStr
    If TallyIntegration Then
        Dim Dom As Object
        Set Dom = CreateObject("MSXML2.DomDocument")
        Dom.async = False
        XMLStr = XMLStr + "<ENVELOPE>"
        XMLStr = XMLStr + "<HEADER>"
        XMLStr = XMLStr + "<TALLYREQUEST>Import Data</TALLYREQUEST>"
        XMLStr = XMLStr + "</HEADER>"
        XMLStr = XMLStr + "<BODY>"
        XMLStr = XMLStr + "<IMPORTDATA>"
        XMLStr = XMLStr + "<REQUESTDESC>"
        XMLStr = XMLStr + "<REPORTNAME>Vouchers</REPORTNAME>"
        XMLStr = XMLStr + "<STATICVARIABLES>"
        XMLStr = XMLStr + "<SVCURRENTCOMPANY>##SVCURRENTCOMPANY</SVCURRENTCOMPANY>" '##SVCURRENTCOMPANY-Current Open Company
        XMLStr = XMLStr + "</STATICVARIABLES>"
        XMLStr = XMLStr + "</REQUESTDESC>"
        XMLStr = XMLStr + "<REQUESTDATA>"
        XMLStr = XMLStr + "<TALLYMESSAGE xmlns:UDF=""TallyUDF"">"
        XMLStr = XMLStr + "<VOUCHER DATE='" & Format(oVchDate, "yyyyMMdd") & "' TAGNAME = ""Voucher Number"" TAGVALUE='" & oVchNo & "' ACTION=""Delete"">"
        XMLStr = XMLStr + "<VOUCHERTYPENAME>" + VchSeries + "</VOUCHERTYPENAME>"
        XMLStr = XMLStr + "</VOUCHER>"
        XMLStr = XMLStr + "</TALLYMESSAGE>"
        XMLStr = XMLStr + "</REQUESTDATA>"
        XMLStr = XMLStr + "</IMPORTDATA>"
        XMLStr = XMLStr + "</BODY>"
        XMLStr = XMLStr + "</ENVELOPE>"
        Dim WinHttpReq As Object
        Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
        With WinHttpReq
            .Open "POST", "http://localhost:" + ReadFromFile("Tally Port"), False
            Do While True
                On Error Resume Next
                .Send XMLStr
                If Err.Number = -2147012867 Then
                    If MsgBox(Err.Description + "Tally is not open. Would you like to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Proceed !") = vbNo Then Exit Do
                Else
                    .waitForResponse 4000
                    Dom.loadXML .responseText
                    If Dom.selectSingleNode("//DELETED").Text = "1" Then
                        If dspMsg Then MsgBox "Voucher from Tally Deleted !!!", vbInformation, App.Title: Exit Do
                    Else
                        If Dom.selectSingleNode("//LINEERROR").Text = "Voucher does not exist!" Then
                            Exit Do
                        Else
                            If MsgBox(Dom.selectSingleNode("//LINEERROR").Text + " Would you like to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Proceed !") = vbNo Then Exit Do
                        End If
                    End If
                End If
            Loop
        End With
    End If
End Sub
Private Sub UpdateIntegration()
cnSalesVoucher.Execute "Update JobworkBVParent Set IntegrationStatus='True' WHERE Code='" & rstSalesVoucherList.Fields("Code").Value & "' AND RIGHT(Type,2)='" & VchType & "' AND FYCode='" & FYCode & "'"
End Sub
