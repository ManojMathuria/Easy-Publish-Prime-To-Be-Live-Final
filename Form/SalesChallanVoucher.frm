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
Begin VB.Form frmSalesChallanVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Challan Voucher"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
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
      TabIndex        =   25
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
      Picture         =   "SalesChallanVoucher.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8835
         Left            =   120
         TabIndex        =   27
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
         TabPicture(0)   =   "SalesChallanVoucher.frx":001C
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
         TabPicture(1)   =   "SalesChallanVoucher.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtNotes"
         Tab(1).Control(1)=   "btnNotes"
         Tab(1).Control(2)=   "Mh3dFrame2"
         Tab(1).ControlCount=   3
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
            Height          =   345
            Left            =   -73440
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   54
            ToolTipText     =   "Open Notes"
            Top             =   8310
            Visible         =   0   'False
            Width           =   1455
         End
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
            Height          =   345
            Left            =   -74880
            TabIndex        =   53
            Top             =   8310
            Width           =   1455
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
            Left            =   600
            MaxLength       =   40
            TabIndex        =   29
            Top             =   8310
            Width           =   8220
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7785
            Left            =   120
            TabIndex        =   28
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
            ColumnCount     =   8
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
               DataField       =   "ChallanNo"
               Caption         =   "Challan No."
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
            BeginProperty Column06 
               DataField       =   "ChallanDate"
               Caption         =   "Challan Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd-MM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "Amount"
               Caption         =   "          Amount"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
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
                  ColumnWidth     =   3734.929
               EndProperty
               BeginProperty Column04 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   2025.071
               EndProperty
               BeginProperty Column05 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1725.165
               EndProperty
               BeginProperty Column06 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1200.189
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7570
            Left            =   -74880
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   480
            Width           =   13260
            _Version        =   65536
            _ExtentX        =   23389
            _ExtentY        =   13353
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
            Picture         =   "SalesChallanVoucher.frx":0054
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
               Width           =   1770
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
               Height          =   330
               Left            =   12135
               TabIndex        =   48
               Top             =   105
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   32896
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
               Caption         =   " Box"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesChallanVoucher.frx":0070
               Picture         =   "SalesChallanVoucher.frx":008C
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
               Height          =   330
               Left            =   9940
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   7130
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "SalesChallanVoucher.frx":00A8
               Caption         =   "SalesChallanVoucher.frx":00C8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":0134
               Keys            =   "SalesChallanVoucher.frx":0152
               Spin            =   "SalesChallanVoucher.frx":019C
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   8175
               TabIndex        =   47
               Top             =   105
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   32896
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
               Caption         =   " Challan No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesChallanVoucher.frx":01C4
               Picture         =   "SalesChallanVoucher.frx":01E0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   5220
               TabIndex        =   36
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
               Caption         =   " Vch Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesChallanVoucher.frx":01FC
               Picture         =   "SalesChallanVoucher.frx":0218
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
               Left            =   9375
               MaxLength       =   40
               TabIndex        =   8
               Top             =   630
               Width           =   3775
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
               Left            =   9380
               MaxLength       =   255
               TabIndex        =   3
               Top             =   105
               Width           =   1090
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
               Left            =   6420
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   7
               Top             =   630
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
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   9
               Top             =   945
               Width           =   4035
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   35
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
               Picture         =   "SalesChallanVoucher.frx":0234
               Picture         =   "SalesChallanVoucher.frx":0250
               Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
                  Height          =   285
                  Left            =   9210
                  TabIndex        =   13
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   930
                  _Version        =   65536
                  _ExtentX        =   1640
                  _ExtentY        =   503
                  Calculator      =   "SalesChallanVoucher.frx":026C
                  Caption         =   "SalesChallanVoucher.frx":028C
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "SalesChallanVoucher.frx":02F8
                  Keys            =   "SalesChallanVoucher.frx":0316
                  Spin            =   "SalesChallanVoucher.frx":0360
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
                  TabIndex        =   14
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1185
                  _Version        =   65536
                  _ExtentX        =   2090
                  _ExtentY        =   503
                  Calculator      =   "SalesChallanVoucher.frx":0388
                  Caption         =   "SalesChallanVoucher.frx":03A8
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "SalesChallanVoucher.frx":0414
                  Keys            =   "SalesChallanVoucher.frx":0432
                  Spin            =   "SalesChallanVoucher.frx":047C
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
               Left            =   3680
               MaxLength       =   25
               TabIndex        =   1
               Top             =   105
               Width           =   1555
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
               Left            =   9375
               MaxLength       =   40
               TabIndex        =   11
               Top             =   950
               Width           =   3775
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
               TabIndex        =   6
               Top             =   630
               Width           =   4035
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   2960
               TabIndex        =   32
               Top             =   105
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
               Caption         =   " Vch No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesChallanVoucher.frx":04A4
               Picture         =   "SalesChallanVoucher.frx":04C0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   33
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
               Picture         =   "SalesChallanVoucher.frx":04DC
               Picture         =   "SalesChallanVoucher.frx":04F8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   8175
               TabIndex        =   34
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesChallanVoucher.frx":0514
               Picture         =   "SalesChallanVoucher.frx":0530
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   6420
               TabIndex        =   2
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "SalesChallanVoucher.frx":054C
               Caption         =   "SalesChallanVoucher.frx":0664
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":06D0
               Keys            =   "SalesChallanVoucher.frx":06EE
               Spin            =   "SalesChallanVoucher.frx":074C
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   37
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
               Picture         =   "SalesChallanVoucher.frx":0774
               Picture         =   "SalesChallanVoucher.frx":0790
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   9940
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   6810
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "SalesChallanVoucher.frx":07AC
               Caption         =   "SalesChallanVoucher.frx":07CC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":0838
               Keys            =   "SalesChallanVoucher.frx":0856
               Spin            =   "SalesChallanVoucher.frx":08A0
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
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   6810
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   582
               Calculator      =   "SalesChallanVoucher.frx":08C8
               Caption         =   "SalesChallanVoucher.frx":08E8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":0954
               Keys            =   "SalesChallanVoucher.frx":0972
               Spin            =   "SalesChallanVoucher.frx":09BC
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
               TabIndex        =   38
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
               Picture         =   "SalesChallanVoucher.frx":09E4
               Picture         =   "SalesChallanVoucher.frx":0A00
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   645
               Left            =   11985
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   6810
               Width           =   1170
               _Version        =   65536
               _ExtentX        =   2055
               _ExtentY        =   1147
               Calculator      =   "SalesChallanVoucher.frx":0A1C
               Caption         =   "SalesChallanVoucher.frx":0A3C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":0AA8
               Keys            =   "SalesChallanVoucher.frx":0AC6
               Spin            =   "SalesChallanVoucher.frx":0B10
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   650
               Left            =   1200
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   6810
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1147
               Calculator      =   "SalesChallanVoucher.frx":0B38
               Caption         =   "SalesChallanVoucher.frx":0B58
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":0BC4
               Keys            =   "SalesChallanVoucher.frx":0BE2
               Spin            =   "SalesChallanVoucher.frx":0C2C
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
               TabIndex        =   39
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
               Picture         =   "SalesChallanVoucher.frx":0C54
               Picture         =   "SalesChallanVoucher.frx":0C70
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
               Height          =   330
               Left            =   8425
               TabIndex        =   40
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
               Picture         =   "SalesChallanVoucher.frx":0C8C
               Picture         =   "SalesChallanVoucher.frx":0CA8
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
               Height          =   330
               Left            =   9390
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   7130
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   582
               Calculator      =   "SalesChallanVoucher.frx":0CC4
               Caption         =   "SalesChallanVoucher.frx":0CE4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":0D50
               Keys            =   "SalesChallanVoucher.frx":0D6E
               Spin            =   "SalesChallanVoucher.frx":0DB8
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
               TabIndex        =   18
               Top             =   6810
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1147
               Calculator      =   "SalesChallanVoucher.frx":0DE0
               Caption         =   "SalesChallanVoucher.frx":0E00
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":0E6C
               Keys            =   "SalesChallanVoucher.frx":0E8A
               Spin            =   "SalesChallanVoucher.frx":0ED4
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
               TabIndex        =   41
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
               Picture         =   "SalesChallanVoucher.frx":0EFC
               Picture         =   "SalesChallanVoucher.frx":0F18
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   650
               Left            =   2280
               TabIndex        =   42
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
               Picture         =   "SalesChallanVoucher.frx":0F34
               Picture         =   "SalesChallanVoucher.frx":0F50
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   650
               Left            =   3120
               TabIndex        =   16
               Top             =   6810
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   1147
               Calculator      =   "SalesChallanVoucher.frx":0F6C
               Caption         =   "SalesChallanVoucher.frx":0F8C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":0FF8
               Keys            =   "SalesChallanVoucher.frx":1016
               Spin            =   "SalesChallanVoucher.frx":1060
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
               TabIndex        =   43
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
               Picture         =   "SalesChallanVoucher.frx":1088
               Picture         =   "SalesChallanVoucher.frx":10A4
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   650
               Left            =   3675
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   6810
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   1147
               Calculator      =   "SalesChallanVoucher.frx":10C0
               Caption         =   "SalesChallanVoucher.frx":10E0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":114C
               Keys            =   "SalesChallanVoucher.frx":116A
               Spin            =   "SalesChallanVoucher.frx":11B4
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
               TabIndex        =   44
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
               Caption         =   " Adjustment"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesChallanVoucher.frx":11DC
               Picture         =   "SalesChallanVoucher.frx":11F8
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   645
               Left            =   7395
               TabIndex        =   19
               Top             =   6810
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   1138
               Calculator      =   "SalesChallanVoucher.frx":1214
               Caption         =   "SalesChallanVoucher.frx":1234
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":12A0
               Keys            =   "SalesChallanVoucher.frx":12BE
               Spin            =   "SalesChallanVoucher.frx":1308
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
               Left            =   5220
               TabIndex        =   45
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
               Picture         =   "SalesChallanVoucher.frx":1330
               Picture         =   "SalesChallanVoucher.frx":134C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   10455
               TabIndex        =   46
               Top             =   105
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
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
               Caption         =   " Dated"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesChallanVoucher.frx":1368
               Picture         =   "SalesChallanVoucher.frx":1384
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput2 
               Height          =   330
               Left            =   11055
               TabIndex        =   4
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "SalesChallanVoucher.frx":13A0
               Caption         =   "SalesChallanVoucher.frx":14B8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":1524
               Keys            =   "SalesChallanVoucher.frx":1542
               Spin            =   "SalesChallanVoucher.frx":15A0
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
               Height          =   330
               Left            =   12610
               TabIndex        =   5
               Top             =   105
               Width           =   540
               _Version        =   65536
               _ExtentX        =   952
               _ExtentY        =   582
               Calculator      =   "SalesChallanVoucher.frx":15C8
               Caption         =   "SalesChallanVoucher.frx":15E8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesChallanVoucher.frx":1654
               Keys            =   "SalesChallanVoucher.frx":1672
               Spin            =   "SalesChallanVoucher.frx":16BC
               AlignHorizontal =   1
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   5220
               TabIndex        =   49
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
               Caption         =   " Challan Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesChallanVoucher.frx":16E4
               Picture         =   "SalesChallanVoucher.frx":1700
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   8175
               TabIndex        =   50
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
               Caption         =   " Ship To Party"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesChallanVoucher.frx":171C
               Picture         =   "SalesChallanVoucher.frx":1738
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   4875
               Left            =   120
               TabIndex        =   12
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
               SpreadDesigner  =   "SalesChallanVoucher.frx":1754
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   120
               TabIndex        =   52
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
               Picture         =   "SalesChallanVoucher.frx":25B0
               Picture         =   "SalesChallanVoucher.frx":25CC
            End
            Begin MSForms.ComboBox cmbChallanType 
               Height          =   330
               Left            =   6420
               TabIndex        =   10
               Top             =   945
               Width           =   1770
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "3122;582"
               ListWidth       =   4762
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
            TabIndex        =   51
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
            Picture         =   "SalesChallanVoucher.frx":25E8
            Picture         =   "SalesChallanVoucher.frx":2604
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
            TabIndex        =   30
            Top             =   8310
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   26
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
Attribute VB_Name = "frmSalesChallanVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchCode As String  'Vch to Modify
Public VchType As String 'RF-Material In IF-Material Out
Dim cnSalesChallanVoucher As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstPartyList As New ADODB.Recordset, rstMaterialCentreList As New ADODB.Recordset, rstTaxList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstHSNCodeList As New ADODB.Recordset, rstVchSeriesList As New ADODB.Recordset
Dim rstSalesChallanVoucherList As New ADODB.Recordset, rstSalesChallanVoucherParent As New ADODB.Recordset, rstSalesChallanVoucherChild As New ADODB.Recordset, rstOrderList As New ADODB.Recordset
Dim PartyCode As String, ConsigneeCode As String, MaterialCentreCode As String, TaxCode As String, VchPrefix As String, VchNumbering As String, VchSeriesCode As String, oVchSeriesCode As String, oVchNo As String, AutoVchNo As String
Dim SortOrder, PrevStr, dblBookMark As Double, blnRecordExist As Boolean, EditMode As Boolean
Dim frmSalesChallanTptDetails As New FrmDespatchDetails
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    Me.Caption = "Material " & IIf(VchType = "IF", "Out", "In") & "-Supply " & IIf(VchType = "IF", "Outward", "Inward") & "-Finished Goods"
    cnSalesChallanVoucher.CursorLocation = adUseClient: cnSalesChallanVoucher.Open cnDatabase.ConnectionString
    rstSalesChallanVoucherParent.CursorLocation = adUseClient
    LoadMasterList
    With rstSalesChallanVoucherList
        .Open "SELECT T.Code,T.Name,V.Code As VchSeriesCode,V.Name As VchSeriesName,Date,T.Type,P.Name As PartyName,C.Name As ConsigneeName,ChallanNo,ChallanDate,Amount FROM ((JobworkBVParent T INNER JOIN AccountMaster P ON T.Party=P.Code) INNER JOIN AccountMaster C ON T.Consignee=C.Code) INNER JOIN VchSeriesMaster V ON T.VchSeries=V.Code  WHERE RIGHT(Type,2)='" & VchType & "' AND FYCode='" & FYCode & "' ORDER BY T.Name", cnSalesChallanVoucher, adOpenKeyset, adLockPessimistic
        .Filter = adFilterNone
        If .RecordCount > 0 Then
            .MoveLast
            If Not CheckEmpty(VchCode, False) Then .MoveFirst: .Find "[Code]='" & VchCode & "'"
        End If
        Set DataGrid1.DataSource = rstSalesChallanVoucherList
        BusySystemIndicator False
        SSTab1.Tab = 0
    If FrmStockLedger.dSortBy = True Then
    SortOrder = "Code"
    Else
    SortOrder = "AutoVchNo"
    End If
'        SortOrder = "Name"
        If Not (.EOF Or .BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        .ActiveConnection = Nothing
    End With
    cmbChallanType.AddItem IIf(VchType = "IF", "Sales", "Purchase") & " Challan", 0
    cmbChallanType.AddItem IIf(VchType = "IF", "Purchase Return", "Sales Return") & " Challan", 1
    cmbChallanType.AddItem "Promotional " & IIf(VchType = "IF", "Sales", "Purchase") & " Challan", 2
    SetButtonsForNoRecord
    fpSpread1.TextTip = TextTipFloating
    Load frmSalesChallanTptDetails
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    With MdiMainMenu
        .mnuMaterialOutSupplyOutward.Enabled = False: .mnuMaterialInSupplyInward.Enabled = False
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
    Call CloseRecordset(rstSalesChallanVoucherList)
    Call CloseRecordset(rstSalesChallanVoucherParent)
    Call CloseRecordset(rstSalesChallanVoucherChild)
    Call CloseRecordset(rstPartyList)
    Call CloseRecordset(rstMaterialCentreList)
    Call CloseRecordset(rstTaxList)
    Call CloseRecordset(rstHSNCodeList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstVchSeriesList)
    Call CloseRecordset(rstOrderList)
    Call CloseConnection(cnSalesChallanVoucher)
    Call CloseForm(frmSalesChallanTptDetails)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuMaterialOutSupplyOutward.Enabled = True: MdiMainMenu.mnuMaterialInSupplyInward.Enabled = True
End Sub
Private Sub Text1_Change()
On Error Resume Next
    With rstSalesChallanVoucherList
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
    With rstSalesChallanVoucherList
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
            If Not (rstSalesChallanVoucherList.EOF Or rstSalesChallanVoucherList.BOF) Then
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
    Dim HiLiteRecord As Boolean, UpdateFlag As Integer, CellVal01 As Variant, CellVal02 As Variant, CellVal03 As Variant, CellVal04 As Variant, i As Integer
    With rstSalesChallanVoucherList
        If Button.Index = 1 Then
            If rstSalesChallanVoucherParent.State = adStateOpen Then rstSalesChallanVoucherParent.Close
            rstSalesChallanVoucherParent.Open "SELECT * FROM JobworkBVParent WHERE Code=''", cnSalesChallanVoucher, adOpenKeyset, adLockOptimistic
            ClearFields
            If AddRecord(rstSalesChallanVoucherParent) Then
                MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
                Call SetButtons(False)
                SSTab1.Tab = 1
                Text6.SetFocus
                blnRecordExist = False
                cnSalesChallanVoucher.BeginTrans
            End If
        ElseIf Button.Index = 2 Then
            If .RecordCount = 0 Then Exit Sub
            SSTab1.Tab = 1
            EditRecord
        ElseIf Button.Index = 3 Then
            If .RecordCount = 0 Then Exit Sub
            If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
            SSTab1.Tab = 1
            If chkRef("SELECT RefCode FROM JobworkBVRef WHERE VchCode='" & .Fields("Code").Value & "' AND Method=1 AND RefCode IN (SELECT RefCode FROM JobworkBVRef WHERE VchCode<>'" & .Fields("Code").Value & "')") Then 'Sales/Purchase Challan can have Method=1 & 2. 1-Challan Item Ref 2-Referring Sales/Purchase Order Item Ref
                DisplayError ("Failed to delete the record")
            ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
                On Error Resume Next
                MdiMainMenu.MousePointer = vbHourglass
                cnSalesChallanVoucher.BeginTrans
                cnSalesChallanVoucher.Execute "DELETE FROM JobworkBVRef WHERE VchCode='" & .Fields("Code").Value & "'"
                cnSalesChallanVoucher.Execute "DELETE FROM JobworkBVParent WHERE Code='" & .Fields("Code").Value & "'"
                MdiMainMenu.MousePointer = vbNormal
                If Err.Number = 0 Then
                    .Delete
                    .MoveNext
                    If .RecordCount > 0 And .EOF Then .MoveLast
                    cnSalesChallanVoucher.CommitTrans
                    ShowProgressInStatusBar True
                    Timer1.Enabled = True
                    Text1.Text = ""
                    .Filter = adFilterNone
                Else
                    DisplayError (Err.Description)
                    cnSalesChallanVoucher.RollbackTrans
                End If
                On Error GoTo 0
            End If
            SetButtons (True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
            HiLiteRecord = True
        ElseIf Button.Index = 4 Then
            If CheckMandatoryFields Then Exit Sub
            frmSalesChallanTptDetails.Show vbModal
            If MsgBox("Are you sure to save the voucher?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Save !") = vbNo Then Exit Sub
            SaveFields
            UpdateFlag = 0
            If UpdateRecord(rstSalesChallanVoucherParent) Then
                If UpdateItemList("D", 0, "", 0) Then
                    UpdateFlag = 1
                   With fpSpread1
                       For i = 1 To .DataRowCnt
                           .SetActiveCell 3, i
                           .GetText 4, i, CellVal01 'Quantity
                           .GetText 8, i, CellVal02 'Item Code
                           .GetText 11, i, CellVal03 'Ref Code
                           .GetText 12, i, CellVal04 'Bal Qty
                           If Val(CellVal01) <> 0 And Not CheckEmpty(CellVal02, False) Then If Not UpdateItemList("I", i, CellVal03, Val(CellVal04)) Then UpdateFlag = 0: Exit For
                       Next
                   End With
                End If
            End If
            If UpdateFlag Then
                AddToList
                cnSalesChallanVoucher.CommitTrans
                If rstSalesChallanVoucherParent.State = adStateOpen Then rstSalesChallanVoucherParent.Close
                rstSalesChallanVoucherParent.CursorLocation = adUseClient
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
            If CancelRecordUpdate(rstSalesChallanVoucherParent) Then
                cnSalesChallanVoucher.RollbackTrans
                If rstSalesChallanVoucherParent.State = adStateOpen Then rstSalesChallanVoucherParent.Close
                rstSalesChallanVoucherParent.CursorLocation = adUseClient
                Call SetButtons(True)
                SetButtonsForNoRecord
                SSTab1.Tab = 0
            End If
        ElseIf Button.Index = 6 Then
            SSTab1.Tab = 0
            Set DataGrid1.DataSource = Nothing
            .Filter = adFilterNone
            RefreshData rstSalesChallanVoucherList
            Set DataGrid1.DataSource = rstSalesChallanVoucherList
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
            Call PrintSalesChallanVoucher(.Fields("Code").Value, Right(.Fields("Type").Value, 2), "P")
            HiLiteRecord = True
        ElseIf Button.Index = 10 Then
            If .RecordCount = 0 Then Exit Sub
            Call PrintSalesChallanVoucher(.Fields("Code").Value, Right(.Fields("Type").Value, 2), "S")
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
        rstSalesChallanVoucherList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstSalesChallanVoucherList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstSalesChallanVoucherList.EOF Or rstSalesChallanVoucherList.BOF) Then
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
    If rstSalesChallanVoucherList.RecordCount = 0 Then
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
        If VchNumbering = "A" Then Text2.Locked = True Else Text2.Locked = False
        If Not blnRecordExist Then 'Vch-New
            If VchNumbering = "A" Then
                AutoVchNo = GenerateCode(cnSalesChallanVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
            End If
        Else 'Vch-Old
            If VchSeriesCode = oVchSeriesCode Then
                Text2.Text = oVchNo
            Else
                If VchNumbering = "A" Then
                    AutoVchNo = GenerateCode(cnSalesChallanVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                    Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                End If
            End If
        End If
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean) 'Vch No.
    With rstSalesChallanVoucherParent
        If .EOF Or .BOF Then Exit Sub
        If CheckEmpty(Text2, True) Then
            Cancel = True
        ElseIf CheckDuplicate(cnSalesChallanVoucher, "JobworkBVParent", "Code", "[Name]+RIGHT(Type,2)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, .Fields("Code").Value, False, FYCode) Then
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
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        PartyCode = slCode: Text3.Text = slName
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
Private Sub cmbChallanType_Click()
    '08-Sales Challan 06-Purchase Return Challan 05-Purchase Challan 07-Sales Return Challan 21-Promotional Sales Challan 22-Promotional Purchase Challan
    VchPrefix = IIf(VchType = "IF", Choose(cmbChallanType.ListIndex + 1, "08", "06", "22"), Choose(cmbChallanType.ListIndex + 1, "05", "07", "21")) & "10" '10-Stock affected
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
    If rstSalesChallanVoucherList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    With rstSalesChallanVoucherParent
        If .State = adStateOpen Then .Close
        .Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstSalesChallanVoucherList.Fields("Code").Value) & "'", cnSalesChallanVoucher, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End With
End Sub
Private Sub ClearFields()
    Text6.Text = "" 'Vch Series
    Text2.Text = "" 'Vch No.
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    Text9.Text = "" 'Challan No
    MhDateInput2.Text = "  -  -    " 'Challan Date
    MhRealInput13.Value = 0 'No. of box
    Text3.Text = "" 'Party
    Text7.Text = "" 'Material Centre
    Text8.Text = "" 'Consignee
    Text5.Text = "" 'Tax Name
    cmbChallanType.ListIndex = 0: cmbChallanType.Enabled = True: cmbChallanType_Click
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
    PartyCode = "": ConsigneeCode = "": MaterialCentreCode = "": TaxCode = "":  VchSeriesCode = "": oVchSeriesCode = "": oVchNo = "": AutoVchNo = ""
    frmSalesChallanTptDetails.Text1.Text = "": frmSalesChallanTptDetails.Text2.Text = "": frmSalesChallanTptDetails.Text3.Text = "": frmSalesChallanTptDetails.Text4.Text = "": frmSalesChallanTptDetails.MhDateInput1.Value = Null
End Sub
Private Sub LoadFields()
    With rstSalesChallanVoucherParent
        If .EOF Or .BOF Then Exit Sub
        VchSeriesCode = .Fields("VchSeries").Value: oVchSeriesCode = VchSeriesCode
        If rstVchSeriesList.RecordCount > 0 Then rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        If Not rstVchSeriesList.EOF Then Text6.Text = rstVchSeriesList.Fields("Col0").Value
        AutoVchNo = Trim(.Fields("AutoVchNo").Value)
        Text2.Text = .Fields("Name").Value
        oVchNo = Trim(Text2.Text)
        MhDateInput1.Text = Format(.Fields("Date").Value, "dd-MM-yyyy")
        Text9.Text = .Fields("ChallanNo").Value
        If Not IsNull(.Fields("ChallanDate").Value) Then MhDateInput2.Text = Format(.Fields("ChallanDate").Value, "dd-MM-yyyy")
        MhRealInput13.Value = Val(.Fields("Box").Value)
        PartyCode = .Fields("Party").Value
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
        TaxCode = .Fields("Tax").Value
        If rstTaxList.RecordCount > 0 Then rstTaxList.MoveFirst
        rstTaxList.Find "[Code] = '" & TaxCode & "'"
        If Not rstTaxList.EOF Then Text5.Text = rstTaxList.Fields("Col0").Value
        cmbChallanType.ListIndex = IIf(VchType = "IF", IIf(Left(.Fields("Type").Value, 2) = "08", 0, IIf(Left(.Fields("Type").Value, 2) = "06", 1, 2)), IIf(Left(.Fields("Type").Value, 2) = "05", 0, IIf(Left(.Fields("Type").Value, 2) = "07", 1, 2)))
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
        txtNotes.Text = .Fields("Notes").Value
        frmSalesChallanTptDetails.Text1.Text = CheckNull(.Fields("Transport").Value): frmSalesChallanTptDetails.Text2.Text = CheckNull(.Fields("GRNo").Value): frmSalesChallanTptDetails.Text3.Text = CheckNull(.Fields("VehicleNo").Value): frmSalesChallanTptDetails.Text4.Text = CheckNull(.Fields("Station").Value): If Not IsNull(.Fields("GRDate").Value) Then frmSalesChallanTptDetails.MhDateInput1.Value = .Fields("GRDate").Value
    End With
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    With rstSalesChallanVoucherParent
        If .RecordCount = 0 Then Exit Sub
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseServer
        .Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstSalesChallanVoucherList.Fields("Code").Value) & "'", cnSalesChallanVoucher, adOpenKeyset, adLockPessimistic
        MdiMainMenu.MousePointer = vbHourglass
        .Fields("RecordStatus") = "N"
    End With
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    If fpSpread1.DataRowCnt > 0 Then cmbChallanType.Enabled = False
    Text6.SetFocus
    blnRecordExist = True
    cnSalesChallanVoucher.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    With rstSalesChallanVoucherParent
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnSalesChallanVoucher, "SELECT MAX(Code) FROM JobworkBVParent", 6, "0")
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
        .Fields("ChallanNo").Value = Text9.Text
        If MhDateInput2.ValueIsNull Then .Fields("ChallanDate").Value = Null Else .Fields("ChallanDate").Value = GetDate(MhDateInput2.Text)
        .Fields("Box").Value = MhRealInput13.Value
        .Fields("Party").Value = PartyCode
        .Fields("MaterialCentre").Value = MaterialCentreCode
        .Fields("Consignee").Value = ConsigneeCode
        .Fields("Tax").Value = TaxCode
        .Fields("Remarks").Value = Trim(Text4.Text)
        .Fields("Rebate%").Value = MhRealInput4.Value
        .Fields("Rebate").Value = MhRealInput5.Value
        .Fields("Freight").Value = MhRealInput6.Value
        .Fields("Adjustment").Value = MhRealInput12.Value
        .Fields("TaxableAmount").Value = MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput12.Value 'Pre-Tax Amt - Discount + Freight + Adjustment
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
        .Fields("Transport").Value = frmSalesChallanTptDetails.Text1.Text
        .Fields("GRNo").Value = frmSalesChallanTptDetails.Text2.Text
        If frmSalesChallanTptDetails.MhDateInput1.ValueIsNull Then .Fields("GRDate").Value = Null Else .Fields("GRDate").Value = GetDate(frmSalesChallanTptDetails.MhDateInput1.Text)
        .Fields("VehicleNo").Value = frmSalesChallanTptDetails.Text3.Text
        .Fields("Station").Value = frmSalesChallanTptDetails.Text4.Text
    End With
End Sub
Private Sub AddToList()
    On Error Resume Next
    With rstSalesChallanVoucherList
        .MoveFirst
        .Find "[Code] = '" & rstSalesChallanVoucherParent.Fields("Code").Value & "'"
        If .EOF Then .AddNew
        .Fields("Code").Value = rstSalesChallanVoucherParent.Fields("Code").Value
        .Fields("Name").Value = Trim(rstSalesChallanVoucherParent.Fields("Name").Value)
        .Fields("VchSeriesCode").Value = VchSeriesCode
        .Fields("VchSeriesName").Value = Text6.Text
        .Fields("Date").Value = rstSalesChallanVoucherParent.Fields("Date").Value
        .Fields("PartyName").Value = Trim(Text3.Text)
        .Fields("ConsigneeName").Value = Trim(Text8.Text)
        .Fields("ChallanNo").Value = rstSalesChallanVoucherParent.Fields("ChallanNo").Value
        .Fields("ChallanDate").Value = rstSalesChallanVoucherParent.Fields("ChallanDate").Value
        .Fields("Amount").Value = MhRealInput11.Value
        .Fields("Type").Value = rstSalesChallanVoucherParent.Fields("Type").Value
        .Update
        .Sort = SortOrder & " Asc"
        .Find "[Code] = '" & rstSalesChallanVoucherParent.Fields("Code").Value & "'"
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text6.Text, False) Then
        Text6.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text2.Text, False) Then
        DisplayError ("Voucher No. cannot be blank"): Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnSalesChallanVoucher, "JobworkBVParent", "Code", "[Name]+RIGHT(Type,2)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, rstSalesChallanVoucherParent.Fields("Code").Value, False, FYCode) Then
        Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text3.Text, False) Then 'Party
        Text3.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text7.Text, False) Then 'Material Centre
        Text7.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text8.Text, False) Then 'Consignee
        Text8.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text5.Text, False) Then 'Tax
        Text5.SetFocus:   CheckMandatoryFields = True: Exit Function
    End If
End Function
Private Sub LoadItemList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    With rstSalesChallanVoucherChild
        If .State = adStateOpen Then .Close
        .Open "SELECT I.Code As ItemCode,I.Name As ItemName,H.Code As HSNCode,H.Name As HSNName,T.Ref As RefOrderCode," & IIf(InStr(1, "0_2", Trim(cmbChallanType.ListIndex)) > 0, "(SELECT LTRIM(VchNo) FROM JobworkBVRef WHERE RefCode=T.Ref AND RIGHT(VchType,2)='" & IIf(VchType = "IF", "SO", "PO") & "')", "''") & " As RefOrderNo,ABS(T.Quantity) As Quantity,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.Ref AND VchCode<>'" & strOrderCode & "')*1) As BalQty,T.Rate,T.[Disc%],T.Amount,T.RefCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM (JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & strOrderCode & "' ORDER BY T.SrNo", cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly
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
Private Function UpdateItemList(ByVal ActionType As String, ByVal SrNo As Integer, ByVal RefCode As String, ByVal BalQty As Long) As Boolean
    Dim CellVal(1 To 12) As Variant
    On Error GoTo ErrorHandler
    UpdateItemList = True
    If ActionType = "D" Then
        If Not blnRecordExist Then Exit Function
        cnSalesChallanVoucher.Execute "DELETE FROM JobworkBVRef WHERE VchCode='" & rstSalesChallanVoucherParent.Fields("Code").Value & "'"
        cnSalesChallanVoucher.Execute "DELETE FROM JobworkBVChild WHERE Code='" & rstSalesChallanVoucherParent.Fields("Code").Value & "'"
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
        If CheckEmpty(RefCode, False) Then RefCode = GenerateCode(cnSalesChallanVoucher, "SELECT MAX(RefCode) FROM JobworkBVRef", 6, "0")
        CellVal(1) = IIf(VchType = "IF", 0 - Val(CellVal(1)), Val(CellVal(1))) '-ve/-ve for IF & +ve/+ve for RF (child/ref)
        cnSalesChallanVoucher.Execute "INSERT INTO JobworkBVChild VALUES ('" & rstSalesChallanVoucherParent.Fields("Code").Value & "','" & CellVal(7) & "','" & VchPrefix & "FI" & "','" & CellVal(5) & "','" & CellVal(6) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(4)) & ",Null," & SrNo & ",'" & CellVal(8) & "','" & CellVal(9) & "','" & CellVal(10) & "','" & CellVal(11) & "','" & CellVal(12) & "'," & Val(CellVal(3)) & ",'" & RefCode & "')"
        cnSalesChallanVoucher.Execute "INSERT INTO JobworkBVRef VALUES ('" & RefCode & "',1,'" & VchPrefix & VchType & "','" & rstSalesChallanVoucherParent.Fields("Code").Value & "','" & rstSalesChallanVoucherParent.Fields("Name").Value & "','" & Format(rstSalesChallanVoucherParent.Fields("Date").Value, "dd-MMM-yyyy") & "','" & PartyCode & "','" & CellVal(5) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ")"
        If Not CheckEmpty(CellVal(7), False) Then
            CellVal(1) = IIf(Abs(Val(CellVal(1))) > BalQty, BalQty, Abs(Val(CellVal(1))))
            CellVal(1) = IIf(VchType = "IF", 0 - Val(CellVal(1)), Val(CellVal(1)))
            cnSalesChallanVoucher.Execute "INSERT INTO JobworkBVRef VALUES ('" & CellVal(7) & "',2,'" & VchPrefix & VchType & "','" & rstSalesChallanVoucherParent.Fields("Code").Value & "','" & rstSalesChallanVoucherParent.Fields("Name").Value & "','" & Format(rstSalesChallanVoucherParent.Fields("Date").Value, "dd-MMM-yyyy") & "','" & PartyCode & "','" & CellVal(5) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ")"
        End If
    End If
    Exit Function
ErrorHandler:
    UpdateItemList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Party" Then
        rstSalesChallanVoucherList.Filter = "[PartyName] Like '%" & SrchText & "%'"
    ElseIf SrchFor = "Material Centre" Then
        rstSalesChallanVoucherList.Filter = "[MaterialCentreName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Item As Variant, i As Integer, x As Integer, cVal(1 To 6) As Variant, Disc As Variant
    With fpSpread1
        If .EditMode Then Exit Sub
        If Shift = 0 And KeyCode = vbKeyF9 Then
            .GetText 11, .ActiveRow, Item  'Ref Code
            If Not CheckEmpty(Item, False) Then
                If chkRef("SELECT RefCode FROM JobworkBVRef WHERE RefCode='" & Item & "' AND VchCode<>'" & rstSalesChallanVoucherParent.Fields("Code").Value & "'") Then DisplayError ("Failed to delete the record"): .SetFocus
            ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
                .DeleteRows .ActiveRow, 1: .SetFocus: CalculateTotal
            End If
        ElseIf KeyCode = vbKeyF3 Then
            If .ActiveCol = 1 Then
                .GetText 10, .ActiveRow, Item 'Ref Order Code
                If Not CheckEmpty(Item, False) Then Exit Sub
                .GetText 11, .ActiveRow, Item 'Ref Code
                If Not CheckEmpty(Item, False) Then If chkRef("SELECT RefCode FROM JobworkBVRef WHERE RefCode='" & Item & "' AND VchCode<>'" & rstSalesChallanVoucherParent.Fields("Code").Value & "'") Then Exit Sub
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
                        .SetText 5, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                    ElseIf Val(Item) <> Val(rstItemList.Fields("Price").Value) Then
                        If MsgBox("Variation in Current (" & Format(Item, "#0.00") & ") and Master (" & Format(rstItemList.Fields("Price").Value, "#0.00") & ") Rate ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 5, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                    End If
                    .GetText 9, .ActiveRow, Item 'HSN Code
                    If CheckEmpty(Item, False) Then .SetText 2, .ActiveRow, rstItemList.Fields("HSNName").Value: .SetText 9, .ActiveRow, rstItemList.Fields("HSNCode").Value
                    LoadMasterList
                    .SetFocus
                    cmbChallanType.Enabled = False
                    Sendkeys "{ENTER}"
                End If
            End If
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then
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
        If fpSpread1.DataRowCnt > 0 Then cmbChallanType.Enabled = False Else cmbChallanType.Enabled = True
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
        MhRealInput8.Value = ((MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput12.Value) * MhRealInput7.Value) / 100 'IGST/CGST
        MhRealInput10.Value = ((MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput12.Value) * MhRealInput9.Value) / 100 'SGST
        MhRealInput11.Value = Round(MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput8.Value + MhRealInput10.Value + MhRealInput12.Value, 0) 'Post-Tax Amount
    End With
End Sub
Private Sub LoadMasterList(Optional ByVal LoadSelected As Boolean)
    If rstPartyList.State = adStateOpen Then rstPartyList.Close
    rstPartyList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly
    rstPartyList.ActiveConnection = Nothing
    If rstMaterialCentreList.State = adStateOpen Then rstMaterialCentreList.Close
    rstMaterialCentreList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='*99999' ORDER BY Name", cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly
    rstMaterialCentreList.ActiveConnection = Nothing
    If rstTaxList.State = adStateOpen Then rstTaxList.Close
    rstTaxList.Open "SELECT Name As Col0,[IGST%],[SGST%],[CGST%],Code FROM TaxMaster ORDER BY Name", cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly
    rstTaxList.ActiveConnection = Nothing
    If rstHSNCodeList.State = adStateOpen Then rstHSNCodeList.Close
    rstHSNCodeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='18' ORDER BY Name", cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly
    rstHSNCodeList.ActiveConnection = Nothing
    If rstItemList.State = adStateOpen Then rstItemList.Close
    If LoadSelected Then
        rstItemList.Open "SELECT I.Name As Col0,FORMAT(dbo.ufnGetItemStock('" & MaterialCentreCode & "',I.Code,'" & Left(VchPrefix, 2) & "','" & CheckNull(rstSalesChallanVoucherParent.Fields("Code").Value) & "','" & GetDate(MhDateInput1.Text) & "'),'#0') As Col1,0 As Quantity,I.Price,I.Code,H.Code As HSNCode,H.Name As HSNName FROM BookMaster I INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE I.Type='F' ORDER BY I.Name", cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly
    Else
        rstItemList.Open "SELECT I.Name As Col0,FORMAT(0,'#0') As Col1,0 As Quantity,I.Price,I.Code,H.Name As HSNName,H.Code As HSNCode FROM BookMaster I INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE I.Type='F' ORDER BY I.Name", cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly
    End If
    rstItemList.ActiveConnection = Nothing
    If rstVchSeriesList.State = adStateOpen Then rstVchSeriesList.Close
    rstVchSeriesList.Open "SELECT Name As Col0,Prefix,Suffix,VchNumbering,Code FROM VchSeriesMaster WHERE VchType='" & IIf(VchType = "IF", "08", "05") & VchType & "' ORDER BY Name", cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly
    rstVchSeriesList.ActiveConnection = Nothing
End Sub
Private Sub LoadOrderList()
    If cmbChallanType.ListIndex = 1 Then Exit Sub 'Sale/Purchase Return
    If rstOrderList.State = adStateOpen Then rstOrderList.Close
    rstOrderList.Open "SELECT VchCode,VchNo,VchDate,SUM(Quantity) As Ordered,SUM(Bal) As Bal FROM (SELECT VchCode,LTRIM(VchNo) As VchNo,VchDate,ABS(Quantity) As Quantity,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.RefCode)*1) As Bal FROM JobworkBVRef T WHERE RIGHT(VchType,2)='" & IIf(VchType = "IF", "SO", "PO") & "' AND Party='" & PartyCode & "' AND Method=1) As Tbl WHERE Bal>0 GROUP BY VchCode,VchNo,VchDate ORDER BY VchNo", cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly 'Sales/Purchase Order can have Method=1 only
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
        rstOrderList.Open "SELECT ItemCode,ItemName,HSNCode,OrderNo,OrderCode,Rate,[Disc%],SUM(Bal) As Bal FROM (SELECT I.Code As ItemCode,I.Name As ItemName,H.Code+'-'+H.Name As HSNCode,LTRIM(T.VchNo) As OrderNo,T.RefCode As OrderCode,T.Rate,T.[Disc%],ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.RefCode)*1) As Bal FROM (JobworkBVRef T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE RIGHT(T.VchType,2)='" & IIf(VchType = "IF", "SO", "PO") & "' AND Method=1 AND T.VchCode IN (" & FrmOrderList.VchCodeList & ")) As Tbl WHERE Bal>0 GROUP BY ItemCode,ItemName,HSNCode,OrderNo,OrderCode,Rate,[Disc%] ORDER BY ItemName,OrderNo", cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly
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
                .Open "SELECT TOP 1 Transport,GRNo,GRDate,VehicleNo,Station FROM JobWorkBVParent WHERE Code IN (" & FrmOrderList.VchCodeList & ") AND (Transport<>'' AND Transport IS NOT NULL) ORDER BY Name", cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly
                If .RecordCount > 0 Then
                    If MsgBox("Do u want to update Transport Details from order ref ('" & CheckNull(.Fields("Transport").Value) & "','" & CheckNull(.Fields("GRNo").Value) & "','" & CheckNull(.Fields("VehicleNo").Value) & "','" & CheckNull(.Fields("Station").Value) & "')?", vbYesNo + vbQuestion + vbDefaultButton1, "Update Transport Details !") = vbYes Then frmSalesChallanTptDetails.Text1.Text = CheckNull(.Fields("Transport").Value): frmSalesChallanTptDetails.Text2.Text = CheckNull(.Fields("GRNo").Value): frmSalesChallanTptDetails.Text3.Text = CheckNull(.Fields("VehicleNo").Value): frmSalesChallanTptDetails.Text4.Text = CheckNull(.Fields("Station").Value): If Not IsNull(.Fields("GRDate").Value) Then frmSalesChallanTptDetails.MhDateInput1.Value = .Fields("GRDate").Value
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
    VchCode = GenerateCode(cnSalesChallanVoucher, "SELECT MAX(Code) FROM JobworkBVParent", 6, "0")
    rstVchSeriesList.MoveFirst
    rstVchSeriesList.Find "[Code] = '" & rstSalesChallanVoucherList.Fields("VchSeriesCode").Value & "'"
    AutoVchNo = GenerateCode(cnSalesChallanVoucher, "SELECT MAX(CONVERT(INT,AutoVchNo))  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & rstSalesChallanVoucherList.Fields("VchSeriesCode").Value & "' AND FYCode='" & FYCode & "'", 10, Space(1))
    VchNo = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
    With cnSalesChallanVoucher
        .BeginTrans
        .Execute "SELECT * INTO #Tbl FROM JobworkBVParent WHERE Code = '" & rstSalesChallanVoucherList.Fields("Code").Value & "'"
        .Execute "UPDATE #Tbl SET Code='" & VchCode & "',Name='" & Trim(VchNo) & "',AutoVchNo='" & Pad(Trim(AutoVchNo), Space(1), 10, "L") & "',[Date]=GETDATE()"
        .Execute "INSERT INTO JobworkBVParent SELECT * FROM #Tbl"
        .Execute "DROP TABLE #Tbl"
        .Execute "SELECT * INTO #Tbl FROM JobworkBVChild Where Code = '" & rstSalesChallanVoucherList.Fields("Code").Value & "'"
        .Execute "UPDATE #Tbl SET Code='" & VchCode & "'"
        .Execute "UPDATE #Tbl SET Ref='',RefCode=''"
        .Execute "INSERT INTO JobworkBVChild SELECT * FROM #Tbl"
        .Execute "DROP TABLE #Tbl"
        .CommitTrans
        Me.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(6)
        Me.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(2)
        Me.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(4)
        Text1.Text = VchNo
    End With
    MdiMainMenu.MousePointer = vbNormal
    Call MsgBox("Successfully Duplicated the Record !", vbInformation, App.Title)
    Exit Sub
ErrorHandler:
    cnSalesChallanVoucher.RollbackTrans
    MdiMainMenu.MousePointer = vbNormal
    DisplayError ("Failed to Duplicate the Record")
End Sub
Private Sub btnNotes_Click()
    frmNotes.NotesFlag = 4
    frmNotes.Label1.Caption = "Notes : Voucher No.: " & Text2.Text
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
Public Sub PrintSalesChallanVoucher(ByVal VchCode As String, ByVal VchType As String, Optional ByVal OutputType As String)
Dim ChallanType As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstSalesChallanVoucherParent.State = adStateOpen Then rstSalesChallanVoucherParent.Close
    rstSalesChallanVoucherParent.Open "SELECT TYPE FROM JobworkBVParent WHERE Code='" + Left(VchCode, 6) + "' ", cnSalesChallanVoucher, adOpenKeyset, adLockOptimistic
    If rstSalesChallanVoucherParent.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    ChallanType = (rstSalesChallanVoucherParent.Fields("TYPE").Value)
    rstSalesChallanVoucherParent.ActiveConnection = Nothing
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,EMail,Website,GSTIN,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,Prefix,Suffix FROM CompanyMaster P INNER JOIN CompChild C ON P.Code=C.Code WHERE VchType= " & IIf(ChallanType = "0510RF", 5, IIf(ChallanType = "0710RF", 7, IIf(ChallanType = "0610IF", 6, IIf(ChallanType = "0810IF", 8, IIf(ChallanType = "2110RF", 21, 0))))), cnSalesChallanVoucher, adOpenKeyset, adLockReadOnly
    If rstSalesChallanVoucherChild.State = adStateOpen Then rstSalesChallanVoucherChild.Close
    rstSalesChallanVoucherChild.Open "SELECT LTrim(P1.Name) As BillNo,P1.Date As BillDate,A.PrintName As Party,A.Address1 As PartyAddress1,A.Address2 As PartyAddress2,A.Address3 As PartyAddress3,A.Address4 As PartyAddress4,A.TIN As PartyGSTIN,A.Mobile As Mobile,A.eMail As eMail,IIf(P1.Type= '0810IF',C.PrintName,IIf(P1.Type= '0610IF',C.PrintName,IIf(P1.Type= '0510RF',C.PrintName,IIf(P1.Type= '0710RF',C.PrintName,IIf(P1.Type= '2110RF',C.PrintName,''))))) As Consignee,IIf(P1.Type= '0810IF',C.Address1,IIf(P1.Type= '0610IF',C.Address1,IIf(P1.Type= '0510RF',C.Address1,IIf(P1.Type= '0710RF',C.PrintName,IIf(P1.Type= '2110RF',C.Address1,''))))) As ConsigneeAddress1,IIf(P1.Type= '0810IF',C.Address2,IIf(P1.Type= '0610IF',C.Address2,IIf(P1.Type= '0510RF',C.Address2,IIf(P1.Type= '0710RF',C.Address2,IIf(P1.Type= '2110RF',C.Address2,''))))) As ConsigneeAddress2," & _
                                                "IIf(P1.Type= '0810IF',C.Address3,IIf(P1.Type= '0610IF',C.Address3,IIf(P1.Type= '0510RF',C.Address3,IIf(P1.Type= '0710RF',C.Address3,IIf(P1.Type= '2110RF',C.Address3,''))))) As ConsigneeAddress3,IIf(P1.Type= '0810IF',C.Address4,IIf(P1.Type= '0610IF',C.Address4,IIf(P1.Type= '0510RF',C.Address4,IIf(P1.Type= '0710RF',C.Address4,IIf(P1.Type= '2110RF',C.Address4,''))))) As ConsigneeAddress4,IIf(P1.Type= '0810IF',C.TIN,IIf(P1.Type= '0610IF',C.TIN,IIf(P1.Type= '0510RF',C.TIN,IIf(P1.Type= '0710RF',C.TIN,IIf(P1.Type= '2110RF',C.TIN,''))))) As ConsigneeGSTIN,C.Mobile As Mobile,C.eMail As CeMail,P1.[Rebate%],P1.Rebate,P1.Freight,P1.Adjustment,P1.TaxableAmount,P1.[IGST%],P1.IGST,P1.[SGST%],P1.SGST,P1.[CGST%],P1.CGST,P1.Amount As TotalAmount,P1.Remarks,'' As Narration,I.PrintName As Item,H.PrintName As HSNCode," & _
                                                "C1.Quantity,C1.Rate,C1.Amount,N.Name As SrNo,'' As cmbTitle,LTrim(C1.Code)+LTrim(C1.SrNo) As Ref,C1.[Disc%] AS Disc,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,P1.ChallanNo,P1.ChallanDate,P1.Transport,P1.GRNo,P1.GRDate,P1.VehicleNo,P1.Station,eWayBill,eWayBillDate,M.PrintName As MC FROM (((((((JobworkBVParent P1 INNER JOIN JobworkBVChild C1 ON P1.Code=C1.Code)INNER JOIN AccountMaster A ON P1.Party=A.Code)INNER JOIN AccountMaster C ON P1.Consignee=C.Code)INNER JOIN BookMaster I ON C1.Item=I.Code)LEFT JOIN AccountMaster M ON P1.MaterialCentre=M.Code)LEFT JOIN GeneralMaster N ON C1.Narration=N.Code)LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code)LEFT JOIN GeneralMaster S ON I.FinishSize=S.Code WHERE P1.Code='" + Left(VchCode, 6) + "' ORDER BY I.PrintName,N.Name", cnSalesChallanVoucher, adOpenKeyset, adLockOptimistic
    If rstSalesChallanVoucherChild.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rstSalesChallanVoucherChild.ActiveConnection = Nothing
With rptSalesOrderVoucher
    rptSalesOrderVoucher.Text1.SetText IIf(ChallanType = "0710RF", "Sales Return", IIf(ChallanType = "0510RF", "Purchase", IIf(ChallanType = "0810IF", "Sales ", IIf(ChallanType = "0610IF", "Purchase Return", IIf(ChallanType = "2110RF", "Promotional Sales", "Stock Transfer"))))) & " Challan"
    rptSalesOrderVoucher.Text13.SetText IIf(ChallanType = "0710RF", "Buyer :", IIf(ChallanType = "0510RF", "Supplier :", IIf(ChallanType = "0810IF", "Buyer :", IIf(ChallanType = "0610IF", "Supplier :", IIf(ChallanType = "2110RF", "Buyer :", "From: Material Centre")))))
    
    If Left(ChallanType, 2) = "01" Or Left(ChallanType, 2) = "02" Or Left(ChallanType, 2) = "03" Or Left(ChallanType, 2) = "04" Or Left(ChallanType, 2) = "05" Or Left(ChallanType, 2) = "06" Or Left(ChallanType, 2) = "07" Or Left(ChallanType, 2) = "08" Or Left(ChallanType, 2) = "21" Then
        If Right(ChallanType, 2) = "RF" Or Right(ChallanType, 2) = "IF" Then rptSalesOrderVoucher.Text7.SetText "Consignee :"
    Else
        rptSalesOrderVoucher.Text7.SetText "TO: Material Centre :"
    End If
    'rptSalesOrderVoucher.Text7.SetText IIf(ChallanType = "0710RF", "Consignee :", IIf(ChallanType = "0510RF", "Consignee :", IIf(ChallanType = "0810IF", "Consignee :", IIf(ChallanType = "0610IF", "Consignee :", "Consignee :", "TO: Material Centre"))))
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
    
    If Trim(rstSalesChallanVoucherChild.Fields("ChallanNo").Value) <> "" Then .Text42.SetText Trim(rstSalesChallanVoucherChild.Fields("ChallanNo").Value) + " Dt : " & Format(rstSalesChallanVoucherChild.Fields("ChallanDate").Value, "dd-MM-yy") Else .Text41.SetText ""
    If Trim(rstSalesChallanVoucherChild.Fields("GRNo").Value) = "" And rstSalesChallanVoucherChild.Fields("VehicleNo").Value <> "" Then .Text46.SetText Trim(rstSalesChallanVoucherChild.Fields("VehicleNo").Value): .Text45.SetText "Vehicle NO.  :"
    If Trim(rstSalesChallanVoucherChild.Fields("GRNo").Value) = "" And rstSalesChallanVoucherChild.Fields("VehicleNo").Value = "" Then .Text45.SetText ""
    If Trim(rstSalesChallanVoucherChild.Fields("GRNo").Value) <> "" Then .Text46.SetText Trim(rstSalesChallanVoucherChild.Fields("GRNo").Value) + " Dt : " & Format(rstSalesChallanVoucherChild.Fields("GRDate").Value, "dd-MM-yy")
    If rstSalesChallanVoucherChild.Fields("Transport").Value Then .Text44.SetText Trim(rstSalesChallanVoucherChild.Fields("Transport").Value) Else .Text43.SetText ""
    If Trim(rstSalesChallanVoucherChild.Fields("Station").Value) <> "" Then .Text49.SetText Trim(rstSalesChallanVoucherChild.Fields("Station").Value) Else .Text47.SetText ""
    If Trim(rstSalesChallanVoucherChild.Fields("eWayBill").Value) <> "" Then .Text51.SetText Trim(rstSalesChallanVoucherChild.Fields("eWayBill").Value) + " Dt : " & Format(rstSalesChallanVoucherChild.Fields("eWayBillDate").Value, "dd-MM-yy") Else .Text50.SetText ""
    
    rptSalesOrderVoucher.Text10.SetText "(" & UCase(Trim(NumberToWords(rstSalesChallanVoucherChild.Fields("TotalAmount").Value, False))) & ")"
    rptSalesOrderVoucher.Text11.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptSalesOrderVoucher.Text26.SetText CheckNull(rstCompanyMaster.Fields("Declaration01").Value)
    rptSalesOrderVoucher.Text25.SetText CheckNull(rstCompanyMaster.Fields("Declaration02").Value)
    rptSalesOrderVoucher.Text22.SetText CheckNull(rstCompanyMaster.Fields("Declaration03").Value)
    rptSalesOrderVoucher.Text12.SetText CheckNull(rstCompanyMaster.Fields("Declaration04").Value)
    rptSalesOrderVoucher.Text9.SetText CheckNull(rstCompanyMaster.Fields("Declaration05").Value)
    rptSalesOrderVoucher.Text30.SetText CheckNull(rstCompanyMaster.Fields("Declaration06").Value)
    rptSalesOrderVoucher.Text31.SetText CheckNull(rstCompanyMaster.Fields("Declaration07").Value)
    If Len(LTrim((rstSalesChallanVoucherChild.Fields("MC").Value))) <> 0 Then rptSalesOrderVoucher.Section11.Suppress = False: rptSalesOrderVoucher.Section5.Suppress = False
    rptSalesOrderVoucher.Text36.SetText CheckNull(rstSalesChallanVoucherChild.Fields("MC").Value)
    rptSalesOrderVoucher.Database.SetDataSource rstSalesChallanVoucherChild, 3, 1
    rptSalesOrderVoucher.DiscardSavedData
    Screen.MousePointer = vbNormal
    If OutputType = "S" Then
        Set FrmReportViewer.Report = rptSalesOrderVoucher
        FrmReportViewer.Show vbModal
    Else
        If rstSalesChallanVoucherList.State = adStateClosed Then  'For Print Utility
            rptSalesOrderVoucher.PaperSource = crPRBinAuto
            rptSalesOrderVoucher.PrintOut False
        Else
            rptSalesOrderVoucher.PaperSource = crPRBinAuto
            rptSalesOrderVoucher.PrintOut
        End If
    End If
    Set rptSalesOrderVoucher = Nothing
    If rstSalesChallanVoucherList.State = adStateClosed Then  'For Print Utility
        Call CloseRecordset(rstCompanyMaster)
    End If
    Call CloseRecordset(rstSalesChallanVoucherChild)
End With
    On Error GoTo 0
End Sub
'ALTER TABLE [JobWorkBVRef] ALTER COLUMN [VchNo] NVARCHAR(25)
