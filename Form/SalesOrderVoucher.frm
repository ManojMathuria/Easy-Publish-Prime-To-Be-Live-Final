VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSalesOrderVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Order Voucher"
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
      TabIndex        =   13
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
      Picture         =   "SalesOrderVoucher.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8835
         Left            =   120
         TabIndex        =   15
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
         TabPicture(0)   =   "SalesOrderVoucher.frx":001C
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
         TabPicture(1)   =   "SalesOrderVoucher.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtNotes"
         Tab(1).Control(1)=   "Mh3dFrame2"
         Tab(1).Control(2)=   "btnNotes"
         Tab(1).ControlCount=   3
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
            TabIndex        =   46
            Top             =   8280
            Width           =   855
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
            Left            =   605
            MaxLength       =   40
            TabIndex        =   17
            Top             =   8310
            Width           =   8220
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7785
            Left            =   120
            TabIndex        =   16
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
            ColumnCount     =   6
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
               DataField       =   "MaterialCentreName"
               Caption         =   "Material Centre"
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
                  ColumnWidth     =   4080.189
               EndProperty
               BeginProperty Column04 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   3284.788
               EndProperty
               BeginProperty Column05 
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
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   480
            Width           =   13260
            _Version        =   65536
            _ExtentX        =   23389
            _ExtentY        =   13353
            _StockProps     =   77
            BackColor       =   16777215
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
            Picture         =   "SalesOrderVoucher.frx":0054
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
               Left            =   6420
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   4
               Top             =   630
               Width           =   4095
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
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   0
               Top             =   105
               Width           =   1770
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
               Left            =   11700
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   5
               Top             =   630
               Width           =   1455
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
               Width           =   4035
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   23
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
               Picture         =   "SalesOrderVoucher.frx":0070
               Picture         =   "SalesOrderVoucher.frx":008C
               Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
                  Height          =   285
                  Left            =   9210
                  TabIndex        =   24
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   930
                  _Version        =   65536
                  _ExtentX        =   1640
                  _ExtentY        =   503
                  Calculator      =   "SalesOrderVoucher.frx":00A8
                  Caption         =   "SalesOrderVoucher.frx":00C8
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "SalesOrderVoucher.frx":0134
                  Keys            =   "SalesOrderVoucher.frx":0152
                  Spin            =   "SalesOrderVoucher.frx":019C
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
                  TabIndex        =   27
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1185
                  _Version        =   65536
                  _ExtentX        =   2090
                  _ExtentY        =   503
                  Calculator      =   "SalesOrderVoucher.frx":01C4
                  Caption         =   "SalesOrderVoucher.frx":01E4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "SalesOrderVoucher.frx":0250
                  Keys            =   "SalesOrderVoucher.frx":026E
                  Spin            =   "SalesOrderVoucher.frx":02B8
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
               Left            =   6420
               MaxLength       =   25
               TabIndex        =   1
               Top             =   105
               Width           =   2730
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
               Left            =   6420
               MaxLength       =   40
               TabIndex        =   7
               Top             =   950
               Width           =   6735
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
               TabIndex        =   3
               Top             =   630
               Width           =   4035
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   5220
               TabIndex        =   20
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
               Picture         =   "SalesOrderVoucher.frx":02E0
               Picture         =   "SalesOrderVoucher.frx":02FC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   21
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
               Picture         =   "SalesOrderVoucher.frx":0318
               Picture         =   "SalesOrderVoucher.frx":0334
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   5220
               TabIndex        =   22
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
               Picture         =   "SalesOrderVoucher.frx":0350
               Picture         =   "SalesOrderVoucher.frx":036C
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   11700
               TabIndex        =   2
               Top             =   105
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   582
               Calendar        =   "SalesOrderVoucher.frx":0388
               Caption         =   "SalesOrderVoucher.frx":04A0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesOrderVoucher.frx":050C
               Keys            =   "SalesOrderVoucher.frx":052A
               Spin            =   "SalesOrderVoucher.frx":0588
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
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   4875
               Left            =   120
               TabIndex        =   8
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
               MaxCols         =   17
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "SalesOrderVoucher.frx":05B0
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
               Left            =   960
               MaxLength       =   100
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   2160
               Width           =   11715
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   10500
               TabIndex        =   26
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
               Picture         =   "SalesOrderVoucher.frx":12E7
               Picture         =   "SalesOrderVoucher.frx":1303
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   28
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
               Picture         =   "SalesOrderVoucher.frx":131F
               Picture         =   "SalesOrderVoucher.frx":133B
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   9940
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   6810
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "SalesOrderVoucher.frx":1357
               Caption         =   "SalesOrderVoucher.frx":1377
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesOrderVoucher.frx":13E3
               Keys            =   "SalesOrderVoucher.frx":1401
               Spin            =   "SalesOrderVoucher.frx":144B
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
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   6810
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   582
               Calculator      =   "SalesOrderVoucher.frx":1473
               Caption         =   "SalesOrderVoucher.frx":1493
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesOrderVoucher.frx":14FF
               Keys            =   "SalesOrderVoucher.frx":151D
               Spin            =   "SalesOrderVoucher.frx":1567
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
               TabIndex        =   31
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
               Picture         =   "SalesOrderVoucher.frx":158F
               Picture         =   "SalesOrderVoucher.frx":15AB
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   645
               Left            =   11985
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   6810
               Width           =   1170
               _Version        =   65536
               _ExtentX        =   2055
               _ExtentY        =   1147
               Calculator      =   "SalesOrderVoucher.frx":15C7
               Caption         =   "SalesOrderVoucher.frx":15E7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesOrderVoucher.frx":1653
               Keys            =   "SalesOrderVoucher.frx":1671
               Spin            =   "SalesOrderVoucher.frx":16BB
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
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   6810
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1147
               Calculator      =   "SalesOrderVoucher.frx":16E3
               Caption         =   "SalesOrderVoucher.frx":1703
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesOrderVoucher.frx":176F
               Keys            =   "SalesOrderVoucher.frx":178D
               Spin            =   "SalesOrderVoucher.frx":17D7
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
               TabIndex        =   34
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
               Picture         =   "SalesOrderVoucher.frx":17FF
               Picture         =   "SalesOrderVoucher.frx":181B
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
               Height          =   330
               Left            =   8425
               TabIndex        =   35
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
               Picture         =   "SalesOrderVoucher.frx":1837
               Picture         =   "SalesOrderVoucher.frx":1853
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
               Height          =   330
               Left            =   9940
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   7130
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "SalesOrderVoucher.frx":186F
               Caption         =   "SalesOrderVoucher.frx":188F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesOrderVoucher.frx":18FB
               Keys            =   "SalesOrderVoucher.frx":1919
               Spin            =   "SalesOrderVoucher.frx":1963
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
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   7130
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   582
               Calculator      =   "SalesOrderVoucher.frx":198B
               Caption         =   "SalesOrderVoucher.frx":19AB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesOrderVoucher.frx":1A17
               Keys            =   "SalesOrderVoucher.frx":1A35
               Spin            =   "SalesOrderVoucher.frx":1A7F
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
               TabIndex        =   10
               Top             =   6810
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1147
               Calculator      =   "SalesOrderVoucher.frx":1AA7
               Caption         =   "SalesOrderVoucher.frx":1AC7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesOrderVoucher.frx":1B33
               Keys            =   "SalesOrderVoucher.frx":1B51
               Spin            =   "SalesOrderVoucher.frx":1B9B
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
               Picture         =   "SalesOrderVoucher.frx":1BC3
               Picture         =   "SalesOrderVoucher.frx":1BDF
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
               Picture         =   "SalesOrderVoucher.frx":1BFB
               Picture         =   "SalesOrderVoucher.frx":1C17
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   650
               Left            =   3120
               TabIndex        =   9
               Top             =   6810
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   1147
               Calculator      =   "SalesOrderVoucher.frx":1C33
               Caption         =   "SalesOrderVoucher.frx":1C53
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesOrderVoucher.frx":1CBF
               Keys            =   "SalesOrderVoucher.frx":1CDD
               Spin            =   "SalesOrderVoucher.frx":1D27
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
               Picture         =   "SalesOrderVoucher.frx":1D4F
               Picture         =   "SalesOrderVoucher.frx":1D6B
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   650
               Left            =   3675
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   6810
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   1147
               Calculator      =   "SalesOrderVoucher.frx":1D87
               Caption         =   "SalesOrderVoucher.frx":1DA7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesOrderVoucher.frx":1E13
               Keys            =   "SalesOrderVoucher.frx":1E31
               Spin            =   "SalesOrderVoucher.frx":1E7B
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
               Caption         =   " Adjustment"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "SalesOrderVoucher.frx":1EA3
               Picture         =   "SalesOrderVoucher.frx":1EBF
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   645
               Left            =   7395
               TabIndex        =   11
               Top             =   6810
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   1138
               Calculator      =   "SalesOrderVoucher.frx":1EDB
               Caption         =   "SalesOrderVoucher.frx":1EFB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "SalesOrderVoucher.frx":1F67
               Keys            =   "SalesOrderVoucher.frx":1F85
               Spin            =   "SalesOrderVoucher.frx":1FCF
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
               Left            =   10500
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
               Picture         =   "SalesOrderVoucher.frx":1FF7
               Picture         =   "SalesOrderVoucher.frx":2013
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   120
               TabIndex        =   44
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
               Picture         =   "SalesOrderVoucher.frx":202F
               Picture         =   "SalesOrderVoucher.frx":204B
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   5220
               TabIndex        =   47
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
               Picture         =   "SalesOrderVoucher.frx":2067
               Picture         =   "SalesOrderVoucher.frx":2083
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
            TabIndex        =   43
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
            Picture         =   "SalesOrderVoucher.frx":209F
            Picture         =   "SalesOrderVoucher.frx":20BB
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
            Height          =   345
            Left            =   -68280
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   45
            ToolTipText     =   "Open Notes"
            Top             =   3360
            Visible         =   0   'False
            Width           =   1455
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
            TabIndex        =   18
            Top             =   8310
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   14
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
Attribute VB_Name = "frmSalesOrderVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchCode As String  'Vch to Modify
Public VchType As String 'SO-Sales Order PO-Purchase Order ST-Stock Tranfer SQ-Sales Quotation PQ-Purchase Quotation
Dim cnSalesOrderVoucher As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstPartyList As New ADODB.Recordset, rstMaterialCentreList As New ADODB.Recordset, rstTaxList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstHSNCodeList As New ADODB.Recordset, rstVchSeriesList As New ADODB.Recordset
Dim rstSalesOrderVoucherList As New ADODB.Recordset, rstSalesOrderVoucherParent As New ADODB.Recordset, rstSalesOrderVoucherChild As New ADODB.Recordset, rstOrderList As New ADODB.Recordset
Dim PartyCode As String, PartyStateCode As String, ConsigneeCode As String, MaterialCentreCode As String, TaxCode As String, VchPrefix As String, VchNumbering As String, VchSeriesCode As String, oVchSeriesCode As String, oVchNo As String, AutoVchNo As String
Dim SortOrder, PrevStr, dblBookMark As Double, blnRecordExist As Boolean, EditMode As Boolean
Dim frmSalesOrderTptDetails As New FrmDespatchDetails
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    Me.Caption = Switch(VchType = "PQ", "Quotation-Supply Inward", VchType = "SQ", "Quotation-Supply Outward", VchType = "SO", "Sales Order-Supply Outward", VchType = "PO", "Purchase Order-Supply Inward", VchType = "ST", "Stock Transfer") & "-Finished Goods"
    VchPrefix = Switch(VchType = "SQ", "24", VchType = "PQ", "23", VchType = "SO", "18", VchType = "PO", "17", VchType = "ST", "19") & IIf(VchType = "ST", "10", "01") '01-Stock not affected 10-Stock affected
    Mh3dLabel3.Caption = IIf(VchType = "ST", " From", " Party"): Mh3dLabel12.Caption = IIf(VchType = "ST", " To", " Mat Centre")
    DataGrid1.Columns(3).Caption = IIf(VchType = "ST", " From", " Party"): DataGrid1.Columns(4).Caption = IIf(VchType = "ST", " To", " Mat Centre")
    If VchType = "ST" Then Text9.Visible = False: Mh3dLabel6.Visible = False: Mh3dLabel12.Left = 5220: Text7.Left = 6420: Text7.Width = 6735
    cnSalesOrderVoucher.CursorLocation = adUseClient: cnSalesOrderVoucher.Open cnDatabase.ConnectionString
    rstSalesOrderVoucherParent.CursorLocation = adUseClient
    LoadMasterList
    With rstSalesOrderVoucherList
        .Open "SELECT T.Code,T.Name,V.Code As VchSeriesCode,V.Name As VchSeriesName,Date,T.Type,P.Name As PartyName,M.Name As MaterialCentreName,Amount FROM ((JobworkBVParent T INNER JOIN AccountMaster P ON T.Party=P.Code) INNER JOIN AccountMaster M ON T.MaterialCentre=M.Code) INNER JOIN VchSeriesMaster V ON T.VchSeries=V.Code WHERE RIGHT(Type,2)='" & VchType & "' AND T.FYCode='" & FYCode & "' ORDER BY T.Name", cnSalesOrderVoucher, adOpenKeyset, adLockPessimistic
        .Filter = adFilterNone
        If .RecordCount > 0 Then
            .MoveLast
            If Not CheckEmpty(VchCode, False) Then .MoveFirst: .Find "[Code]='" & VchCode & "'"
        End If
        Set DataGrid1.DataSource = rstSalesOrderVoucherList
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
    SetButtonsForNoRecord
    fpSpread1.TextTip = TextTipFloating
    Load frmSalesOrderTptDetails
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    With MdiMainMenu
        .mnuSalesOrderSupplyOutwardFinishedItem.Enabled = False: .mnuPurchaseOrderSupplyInwardFinishedItem.Enabled = False: .mnuStockTranferFinishedItem.Enabled = False
    End With
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If InStr(1, "fpSpread1", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True Else EditMode = False
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
    Call CloseRecordset(rstSalesOrderVoucherList)
    Call CloseRecordset(rstSalesOrderVoucherParent)
    Call CloseRecordset(rstSalesOrderVoucherChild)
    Call CloseRecordset(rstPartyList)
    Call CloseRecordset(rstMaterialCentreList)
    Call CloseRecordset(rstTaxList)
    Call CloseRecordset(rstHSNCodeList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstVchSeriesList)
    Call CloseRecordset(rstOrderList)
    Call CloseConnection(cnSalesOrderVoucher)
    Call CloseForm(frmSalesOrderTptDetails)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuSalesOrderSupplyOutwardFinishedItem.Enabled = True: MdiMainMenu.mnuPurchaseOrderSupplyInwardFinishedItem.Enabled = True: MdiMainMenu.mnuStockTranferFinishedItem.Enabled = True
End Sub
Private Sub Text1_Change()
On Error Resume Next
    With rstSalesOrderVoucherList
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
    With rstSalesOrderVoucherList
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
            If Not (rstSalesOrderVoucherList.EOF Or rstSalesOrderVoucherList.BOF) Then
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
    With rstSalesOrderVoucherList
        If Button.Index = 1 Then
            If rstSalesOrderVoucherParent.State = adStateOpen Then rstSalesOrderVoucherParent.Close
            rstSalesOrderVoucherParent.Open "SELECT * FROM JobworkBVParent WHERE Code=''", cnSalesOrderVoucher, adOpenKeyset, adLockOptimistic
            ClearFields
            If AddRecord(rstSalesOrderVoucherParent) Then
                MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
                Call SetButtons(False)
                SSTab1.Tab = 1
                Text8.SetFocus
                blnRecordExist = False
                cnSalesOrderVoucher.BeginTrans
            End If
        ElseIf Button.Index = 2 Then
            If .RecordCount = 0 Then Exit Sub
            SSTab1.Tab = 1
            EditRecord
        ElseIf Button.Index = 3 Then
            If .RecordCount = 0 Then Exit Sub
            If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
            SSTab1.Tab = 1
            If chkRef("SELECT RefCode FROM JobworkBVRef WHERE VchCode='" & .Fields("Code").Value & "' AND Method=1 AND RefCode IN (SELECT RefCode FROM JobworkBVRef WHERE VchCode<>'" & .Fields("Code").Value & "')") Then
                DisplayError ("Failed to delete the record")
            ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
                On Error Resume Next
                MdiMainMenu.MousePointer = vbHourglass
                cnSalesOrderVoucher.BeginTrans
                cnSalesOrderVoucher.Execute "DELETE FROM JobworkBVRef WHERE VchCode='" & .Fields("Code").Value & "'"
                cnSalesOrderVoucher.Execute "DELETE FROM JobworkBVParent WHERE Code='" & .Fields("Code").Value & "'"
                MdiMainMenu.MousePointer = vbNormal
                If Err.Number = 0 Then
                    .Delete
                    .MoveNext
                    If .RecordCount > 0 And .EOF Then .MoveLast
                    cnSalesOrderVoucher.CommitTrans
                    ShowProgressInStatusBar True
                    Timer1.Enabled = True
                    Text1.Text = ""
                    .Filter = adFilterNone
                Else
                    DisplayError (Err.Description)
                    cnSalesOrderVoucher.RollbackTrans
                End If
                On Error GoTo 0
            End If
            SetButtons (True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
            HiLiteRecord = True
        ElseIf Button.Index = 4 Then
            If CheckMandatoryFields Then Exit Sub
            frmSalesOrderTptDetails.Show vbModal
            If MsgBox("Are you sure to save the voucher?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Save !") = vbNo Then Exit Sub
            SaveFields
            UpdateFlag = 0
            If UpdateRecord(rstSalesOrderVoucherParent) Then
                If UpdateItemList("D", 0) Then
                    UpdateFlag = 1
                   With fpSpread1
                       For i = 1 To .DataRowCnt
                           .SetActiveCell 3, i
                           .GetText 4, i, CellVal01 'Quantity
                           .GetText 8, i, CellVal02 'Item Code
                           If Val(CellVal01) <> 0 And Not CheckEmpty(CellVal02, False) Then If Not UpdateItemList("I", i) Then UpdateFlag = 0: Exit For
                       Next
                   End With
                End If
            End If
            If UpdateFlag Then
                AddToList
                cnSalesOrderVoucher.CommitTrans
                If rstSalesOrderVoucherParent.State = adStateOpen Then rstSalesOrderVoucherParent.Close
                rstSalesOrderVoucherParent.CursorLocation = adUseClient
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
            If CancelRecordUpdate(rstSalesOrderVoucherParent) Then
                cnSalesOrderVoucher.RollbackTrans
                If rstSalesOrderVoucherParent.State = adStateOpen Then rstSalesOrderVoucherParent.Close
                rstSalesOrderVoucherParent.CursorLocation = adUseClient
                Call SetButtons(True)
                SetButtonsForNoRecord
                SSTab1.Tab = 0
            End If
        ElseIf Button.Index = 6 Then
            SSTab1.Tab = 0
            Set DataGrid1.DataSource = Nothing
            .Filter = adFilterNone
            RefreshData rstSalesOrderVoucherList
            Set DataGrid1.DataSource = rstSalesOrderVoucherList
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
            If VchType = "SO" Then DisplayMenu ("P") Else Call PrintSalesOrderVoucher(.Fields("Code").Value, Right(.Fields("Type").Value, 2), "P")
        ElseIf Button.Index = 10 Then
            If .RecordCount = 0 Then Exit Sub
            If VchType = "SO" Then DisplayMenu ("S") Else Call PrintSalesOrderVoucher(.Fields("Code").Value, Right(.Fields("Type").Value, 2), "S")
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
        rstSalesOrderVoucherList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstSalesOrderVoucherList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstSalesOrderVoucherList.EOF Or rstSalesOrderVoucherList.BOF) Then
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
    If rstSalesOrderVoucherList.RecordCount = 0 Then
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
                AutoVchNo = GenerateCode(cnSalesOrderVoucher, "SELECT MAX(CONVERT(INT,AutoVchNo)) FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
            End If
        Else 'Vch-Old
            If VchSeriesCode = oVchSeriesCode Then
                Text2.Text = oVchNo
            Else
                If VchNumbering = "A" Then
                    AutoVchNo = GenerateCode(cnSalesOrderVoucher, "SELECT MAX(CONVERT(INT,AutoVchNo)) FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                    Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                End If
            End If
        End If
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean) 'Vch No.
    With rstSalesOrderVoucherParent
        If .EOF Or .BOF Then Exit Sub
        If CheckEmpty(Text2, True) Then
            Cancel = True
        ElseIf CheckDuplicate(cnSalesOrderVoucher, "JobworkBVParent", "Code", "[Name]+RIGHT(Type,2)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, .Fields("Code").Value, False, FYCode) Then
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
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = IIf(VchType = "ST", "*99999", "")
        FrmAccountMaster.MasterCode = PartyCode
        FrmAccountMaster.StateCode = PartyStateCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        PartyCode = slCode: Text3.Text = slName
        If Not IsNull(slStateCode) Then
        PartyStateCode = slStateCode
        End If
        If Not CheckEmpty(PartyCode, False) Then
            If Not blnRecordExist Then Text9.Text = Text3.Text: ConsigneeCode = PartyCode
            LoadMasterList
            Sendkeys "{TAB}"
        End If
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If CheckEmpty(Text3.Text, False) Then Cancel = True
End Sub
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = IIf(VchType = "ST", "*99999", "")
        FrmAccountMaster.MasterCode = ConsigneeCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        ConsigneeCode = slCode: Text9.Text = slName
        If Not CheckEmpty(ConsigneeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
    If CheckEmpty(Text9.Text, False) Then Cancel = True
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
    End If
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
    If CheckEmpty(Text7.Text, False) Then Cancel = True
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
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    If CheckEmpty(Text5.Text, False) Then Cancel = True
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
    If rstSalesOrderVoucherList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    With rstSalesOrderVoucherParent
        If .State = adStateOpen Then .Close
        .Open "SELECT * , (Select State From AccountMaster Where Code=Party) AS State FROM JobworkBVParent WHERE Code='" & FixQuote(rstSalesOrderVoucherList.Fields("Code").Value) & "'", cnSalesOrderVoucher, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End With
End Sub
Private Sub ClearFields()
    Text8.Text = "" 'Vch Series
    Text2.Text = "" 'Vch No.
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    Text3.Text = "" 'Party
    Text9.Text = "" 'Consignee
    Text7.Text = "" 'Material Centre
    Text5.Text = "" 'Tax Name
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
    PartyStateCode = ""
    PartyCode = "": ConsigneeCode = "": MaterialCentreCode = "": TaxCode = "": VchSeriesCode = "": oVchSeriesCode = "": oVchNo = "": AutoVchNo = ""
    frmSalesOrderTptDetails.Text1.Text = "": frmSalesOrderTptDetails.Text2.Text = "": frmSalesOrderTptDetails.Text3.Text = "": frmSalesOrderTptDetails.Text4.Text = "": frmSalesOrderTptDetails.MhDateInput1.Value = Null
End Sub
Private Sub LoadFields()
    With rstSalesOrderVoucherParent
        If .EOF Or .BOF Then Exit Sub
        VchSeriesCode = .Fields("VchSeries").Value: oVchSeriesCode = VchSeriesCode
        If rstVchSeriesList.RecordCount > 0 Then rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        If Not rstVchSeriesList.EOF Then Text8.Text = rstVchSeriesList.Fields("Col0").Value
        AutoVchNo = Trim(.Fields("AutoVchNo").Value)
        Text2.Text = .Fields("Name").Value
        oVchNo = Trim(Text2.Text)
        MhDateInput1.Text = Format(.Fields("Date").Value, "dd-MM-yyyy")
        PartyCode = .Fields("Party").Value
        PartyStateCode = .Fields("State").Value
        If rstPartyList.RecordCount > 0 Then rstPartyList.MoveFirst
        rstPartyList.Find "[Code] = '" & PartyCode & "'"
        If Not rstPartyList.EOF Then Text3.Text = rstPartyList.Fields("Col0").Value
        ConsigneeCode = .Fields("Consignee").Value
        If rstPartyList.RecordCount > 0 Then rstPartyList.MoveFirst
        rstPartyList.Find "[Code] = '" & ConsigneeCode & "'"
        If Not rstPartyList.EOF Then Text9.Text = rstPartyList.Fields("Col0").Value
        MaterialCentreCode = .Fields("MaterialCentre").Value
        If rstMaterialCentreList.RecordCount > 0 Then rstMaterialCentreList.MoveFirst
        rstMaterialCentreList.Find "[Code] = '" & MaterialCentreCode & "'"
        If Not rstMaterialCentreList.EOF Then Text7.Text = rstMaterialCentreList.Fields("Col0").Value
        TaxCode = .Fields("Tax").Value
        If rstTaxList.RecordCount > 0 Then rstTaxList.MoveFirst
        rstTaxList.Find "[Code] = '" & TaxCode & "'"
        If Not rstTaxList.EOF Then Text5.Text = rstTaxList.Fields("Col0").Value
        Text4.Text = .Fields("Remarks").Value
        Call LoadItemList(.Fields("Code").Value)
        MhRealInput4.Value = Val(.Fields("Rebate%").Value)
        MhRealInput5.Value = Val(.Fields("Rebate").Value)
        MhRealInput6.Value = Val(.Fields("Freight").Value)
        MhRealInput12.Value = Val(.Fields("Adjustment").Value)
        If Val(.Fields("SGST%").Value) > 0 Then 'Intra-State Supply
            MhRealInput7.Value = Val(.Fields("CGST%").Value)
            MhRealInput8.Value = Val(.Fields("CGST").Value)
            MhRealInput9.Value = Val(.Fields("SGST%").Value)
            MhRealInput10.Value = Val(.Fields("SGST").Value)
        Else 'Inter-State Supply
            MhRealInput7.Value = Val(.Fields("IGST%").Value)
            MhRealInput8.Value = Val(.Fields("IGST").Value)
        End If
        MhRealInput11.Value = Val(.Fields("Amount").Value)
        txtNotes.Text = .Fields("Notes").Value
        frmSalesOrderTptDetails.Text1.Text = CheckNull(.Fields("Transport").Value): frmSalesOrderTptDetails.Text2.Text = CheckNull(.Fields("GRNo").Value): frmSalesOrderTptDetails.Text3.Text = CheckNull(.Fields("VehicleNo").Value): frmSalesOrderTptDetails.Text4.Text = CheckNull(.Fields("Station").Value): If Not IsNull(.Fields("GRDate").Value) Then frmSalesOrderTptDetails.MhDateInput1.Value = .Fields("GRDate").Value
    End With
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    With rstSalesOrderVoucherParent
        If .RecordCount = 0 Then Exit Sub
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseServer
        .Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstSalesOrderVoucherList.Fields("Code").Value) & "'", cnSalesOrderVoucher, adOpenKeyset, adLockPessimistic
        MdiMainMenu.MousePointer = vbHourglass
        .Fields("RecordStatus") = "N"
    End With
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text8.SetFocus
    blnRecordExist = True
    cnSalesOrderVoucher.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    With rstSalesOrderVoucherParent
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnSalesOrderVoucher, "SELECT MAX(Code) FROM JobworkBVParent", 6, "0")
            .Fields("CreatedBy").Value = UserCode
            .Fields("CreatedOn").Value = Now()
            .Fields("Recordstatus").Value = "N"
        Else
            .Fields("ModifiedBy").Value = UserCode
            .Fields("ModifiedOn").Value = Now()
            .Fields("Recordstatus").Value = "M"
        End If
        .Fields("Name").Value = Trim(Text2.Text)
        .Fields("VchSeries").Value = VchSeriesCode
        .Fields("AutoVchNo").Value = Pad(Trim(AutoVchNo), Space(1), 10, "L")
        .Fields("Date").Value = GetDate(MhDateInput1.Text)
        .Fields("Box").Value = 0
        .Fields("Party").Value = PartyCode
        .Fields("Consignee").Value = ConsigneeCode
        .Fields("MaterialCentre").Value = MaterialCentreCode
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
        .Fields("Transport").Value = frmSalesOrderTptDetails.Text1.Text
        .Fields("GRNo").Value = frmSalesOrderTptDetails.Text2.Text
        If frmSalesOrderTptDetails.MhDateInput1.ValueIsNull Then .Fields("GRDate").Value = Null Else .Fields("GRDate").Value = GetDate(frmSalesOrderTptDetails.MhDateInput1.Text)
        .Fields("VehicleNo").Value = frmSalesOrderTptDetails.Text3.Text
        .Fields("Station").Value = frmSalesOrderTptDetails.Text4.Text
    End With
End Sub
Private Sub AddToList()
    On Error Resume Next
    With rstSalesOrderVoucherList
        .MoveFirst
        .Find "[Code] = '" & rstSalesOrderVoucherParent.Fields("Code").Value & "'"
        If .EOF Then .AddNew
        .Fields("Code").Value = rstSalesOrderVoucherParent.Fields("Code").Value
        .Fields("Name").Value = Trim(rstSalesOrderVoucherParent.Fields("Name").Value)
        .Fields("VchSeriesCode").Value = VchSeriesCode
        .Fields("VchSeriesName").Value = Text8.Text
        .Fields("Date").Value = rstSalesOrderVoucherParent.Fields("Date").Value
        .Fields("PartyName").Value = Trim(Text3.Text)
        .Fields("MaterialCentreName").Value = Trim(Text7.Text)
        .Fields("Type").Value = rstSalesOrderVoucherParent.Fields("Type").Value
        .Fields("Amount").Value = MhRealInput11.Value
        .Update
        .Sort = SortOrder & " Asc"
        .Find "[Code] = '" & rstSalesOrderVoucherParent.Fields("Code").Value & "'"
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text8.Text, False) Then
        Text8.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text2.Text, False) Then
        DisplayError ("Voucher No. cannot be blank"): Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnSalesOrderVoucher, "JobworkBVParent", "Code", "[Name]+RIGHT(Type,2)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, rstSalesOrderVoucherParent.Fields("Code").Value, False, FYCode) Then
        Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text3.Text, False) Then 'Party/From Mat Centre
        Text3.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text7.Text, False) Then 'Mat Centre
        Text7.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text5.Text, False) Then 'Tax
        Text5.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf VchType = "ST" Then
        If Text3.Text = Text7.Text Then DisplayError ("Source & Target Material Centres cann't be same"): Text3.SetFocus: CheckMandatoryFields = True: Exit Function
    End If
End Function
Private Sub LoadItemList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    With rstSalesOrderVoucherChild
        If .State = adStateOpen Then .Close
        .Open "SELECT I.Code As ItemCode,I.Name As ItemName,H.Code As HSNCode,H.Name As HSNName,T.Ref As AgRef,(SELECT LTRIM(VchNo) FROM JobworkBVRef WHERE RefCode=T.Ref AND RIGHT(VchType,2)='" & IIf(VchType = "SO", "SQ", "PQ") & "') As AgRefNo,ABS(T.Quantity) As Quantity,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.Ref AND VchCode<>'" & strOrderCode & "')*1) As BalQty,T.Rate,T.[Disc%],T.Amount,RefCode As CTRef,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM (JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & strOrderCode & "' AND " & IIf(VchType = "ST", "Quantity>0", "1=1") & " ORDER BY SrNo", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        If .RecordCount > 0 Then .MoveFirst
        i = 0
        Do Until .EOF
            i = i + 1
            fpSpread1.SetText 1, i, .Fields("ItemName").Value
            fpSpread1.SetText 2, i, .Fields("HSNName").Value
            fpSpread1.SetText 3, i, CheckNull(.Fields("AgRefNo").Value)
            fpSpread1.SetText 4, i, Val(.Fields("Quantity").Value)
            fpSpread1.SetText 5, i, Val(.Fields("Rate").Value)
            fpSpread1.SetText 6, i, Val(.Fields("Disc%").Value)
            fpSpread1.SetText 7, i, Val(.Fields("Amount").Value)
            fpSpread1.SetText 8, i, .Fields("ItemCode").Value
            fpSpread1.SetText 9, i, .Fields("HSNCode").Value
            fpSpread1.SetText 10, i, CheckNull(.Fields("AgRef").Value)
            fpSpread1.SetText 11, i, .Fields("CTRef").Value
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
    DisplayError (Err.Description)
End Sub
Private Function UpdateItemList(ByVal ActionType As String, ByVal SrNo As Integer) As Boolean
    Dim CellVal(1 To 14) As Variant
    On Error GoTo ErrorHandler
    UpdateItemList = True
    If ActionType = "D" Then
        If Not blnRecordExist Then Exit Function
        cnSalesOrderVoucher.Execute "DELETE FROM JobworkBVRef WHERE VchCode='" & rstSalesOrderVoucherParent.Fields("Code").Value & "'"
        cnSalesOrderVoucher.Execute "DELETE FROM JobworkBVChild WHERE Code='" & rstSalesOrderVoucherParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 4, .ActiveRow, CellVal(1) 'qnty
            .GetText 5, .ActiveRow, CellVal(2) 'rate
            .GetText 6, .ActiveRow, CellVal(3) 'disc %
            .GetText 7, .ActiveRow, CellVal(4) 'amount
            .GetText 8, .ActiveRow, CellVal(5) 'item code
            .GetText 9, .ActiveRow, CellVal(6) 'hsn code
            .GetText 10, .ActiveRow, CellVal(7) 'against ref
            .GetText 11, .ActiveRow, CellVal(8) 'current Tran ref
            .GetText 12, .ActiveRow, CellVal(9) 'bal qnty
            .GetText 13, .ActiveRow, CellVal(10) 'Long Narration I
            .GetText 14, .ActiveRow, CellVal(11) 'Long Narration II
            .GetText 15, .ActiveRow, CellVal(12) 'Long Narration III
            .GetText 16, .ActiveRow, CellVal(13) 'Long Narration IV
            .GetText 17, .ActiveRow, CellVal(14) 'Long Narration V
        End With
        With rstSalesOrderVoucherParent
            If VchType = "ST" Then
                cnSalesOrderVoucher.Execute "INSERT INTO JobworkBVChild VALUES ('" & .Fields("Code").Value & "',Null,'" & VchPrefix & "FI" & "','" & CellVal(5) & "','" & CellVal(6) & "'," & 0 - Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(4)) & ",Null," & SrNo & ",'" & CellVal(10) & "','" & CellVal(11) & "','" & CellVal(12) & "','" & CellVal(13) & "','" & CellVal(14) & "'," & Val(CellVal(3)) & ",'')" 'stock transfer out
            Else
                CellVal(1) = IIf(InStr(1, "PQ_SO", VchType) > 0, Val(CellVal(1)), 0 - Val(CellVal(1))) '+ve/+ve for PQ & SO & -ve/-ve for SQ & PO (child/ref)
                If CheckEmpty(CellVal(8), False) Then CellVal(8) = GenerateCode(cnSalesOrderVoucher, "SELECT MAX(RefCode) FROM JobworkBVRef", 6, "0")
                cnSalesOrderVoucher.Execute "INSERT INTO JobworkBVRef VALUES ('" & CellVal(8) & "',1,'" & VchPrefix & VchType & "','" & .Fields("Code").Value & "','" & .Fields("Name").Value & "','" & Format(.Fields("Date").Value, "dd-MMM-yyyy") & "','" & PartyCode & "','" & CellVal(5) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ")" 'current Tran ref entry
                If Not CheckEmpty(CellVal(7), False) Then 'against ref
                    CellVal(1) = IIf(Abs(Val(CellVal(1))) > Val(CellVal(9)), Val(CellVal(9)), Abs(Val(CellVal(1))))
                    CellVal(1) = IIf(InStr(1, "PO", VchType) > 0, 0 - Val(CellVal(1)), Val(CellVal(1)))
                    cnSalesOrderVoucher.Execute "INSERT INTO JobworkBVRef VALUES ('" & CellVal(7) & "',2,'" & VchPrefix & VchType & "','" & .Fields("Code").Value & "','" & .Fields("Name").Value & "','" & Format(.Fields("Date").Value, "dd-MMM-yyyy") & "','" & PartyCode & "','" & CellVal(5) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ")" 'reduce against ref qnty
                End If
            End If
            cnSalesOrderVoucher.Execute "INSERT INTO JobworkBVChild VALUES ('" & .Fields("Code").Value & "'," & IIf(CheckEmpty(CellVal(7), False), "Null", "'" & CellVal(7) & "'") & ",'" & VchPrefix & "FI" & "','" & CellVal(5) & "','" & CellVal(6) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(4)) & ",Null," & SrNo & ",'" & CellVal(10) & "','" & CellVal(11) & "','" & CellVal(12) & "','" & CellVal(13) & "','" & CellVal(14) & "'," & Val(CellVal(3)) & ",'" & CellVal(8) & "')"
        End With
    End If
    Exit Function
ErrorHandler:
    UpdateItemList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Party" Then
        rstSalesOrderVoucherList.Filter = "[PartyName] Like '%" & SrchText & "%'"
    ElseIf SrchFor = "Material Centre" Then
        rstSalesOrderVoucherList.Filter = "[MaterialCentreName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Item As Variant, i As Integer, x As Integer, cVal(1 To 6) As Variant, Disc As Double
    With fpSpread1
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF9) Then
            .GetText 11, .ActiveRow, Item 'current Tran ref
            If Not CheckEmpty(Item, False) Then
                If chkRef("SELECT RefCode FROM JobworkBVRef WHERE RefCode='" & Item & "' AND VchCode<>'" & rstSalesOrderVoucherParent.Fields("Code").Value & "'") Then
                    DisplayError ("Failed to delete the record"): .SetFocus
                 ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
                    .DeleteRows .ActiveRow, 1: .SetFocus: CalculateTotal
                End If
            End If
        ElseIf KeyCode = vbKeyF3 Then
            If .ActiveCol = 1 Then
                .GetText 10, .ActiveRow, Item 'against ref
                If Not CheckEmpty(Item, False) Then Exit Sub
                .GetText 11, .ActiveRow, Item 'current Tran ref
                If Not CheckEmpty(Item, False) Then If chkRef("SELECT RefCode FROM JobworkBVRef WHERE RefCode='" & Item & "' AND VchCode<>'" & rstSalesOrderVoucherParent.Fields("Code").Value & "'") Then Exit Sub
                .GetText 9, .ActiveRow, Item
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
            If fpSpread1.DataRowCnt = 0 And InStr(1, "SO_PO", VchType) > 0 Then LoadOrderList
        End If
    End With
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Item As Variant, Qty As Variant, Rate As Variant, Disc As Variant
    With fpSpread1
        If Col = 4 Or Col = 5 Or Col = 6 Then 'qnty,disc,rate
            .GetText 8, Row, Item
            .GetText 4, Row, Qty
            .GetText 5, Row, Rate
            .GetText 6, Row, Disc
            Disc = (Rate * Disc) / 100
            If Not CheckEmpty(Item, False) Then .SetText 7, Row, Qty * Round((Rate - Disc), 2): CalculateTotal Else .SetText 4, Row, "": .SetText 5, Row, "": .SetText 6, Row, "": .SetText 7, Row, ""
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
    rstPartyList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
    rstPartyList.ActiveConnection = Nothing
    If rstMaterialCentreList.State = adStateOpen Then rstMaterialCentreList.Close
    rstMaterialCentreList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='*99999' ORDER BY Name", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
    rstMaterialCentreList.ActiveConnection = Nothing
    If rstTaxList.State = adStateOpen Then rstTaxList.Close
    If PartyStateCode = "" Or PartyStateCode = Null Then
    rstTaxList.Open "SELECT Name As Col0,[IGST%],[SGST%],[CGST%],Region,Code FROM TaxMaster ORDER BY Name", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
    Else
    rstTaxList.Open "SELECT Name As Col0,[IGST%],[SGST%],[CGST%],Region,Code FROM TaxMaster Where Region='" & IIf(CompStateCode = PartyStateCode, "L", "I") & "' ORDER BY Name", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
    End If
    rstTaxList.ActiveConnection = Nothing
    If rstHSNCodeList.State = adStateOpen Then rstHSNCodeList.Close
    rstHSNCodeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='18' ORDER BY Name", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
    rstHSNCodeList.ActiveConnection = Nothing
    If rstItemList.State = adStateOpen Then rstItemList.Close
    If LoadSelected Then
        'rstItemList.Open "SELECT I.Name As Col0,FORMAT(dbo.ufnGetItemStock('" & MaterialCentreCode & "',I.Code,'" & Left(VchPrefix, 2) & "','" & CheckNull(rstSalesOrderVoucherParent.Fields("Code").Value) & "','" & GetDate(MhDateInput1.Text) & "'),'#0') As Col1,0 As Quantity,I.Price,I.Code,H.Code As HSNCode,H.Name As HSNName FROM BookMaster I INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE I.Type='F' ORDER BY I.Name", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
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
                "WHERE I.Type='F') As Tbl ORDER BY Col0 ASC", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
    Else
        rstItemList.Open "SELECT I.Name As Col0,FORMAT(0,'#0') As Col1,0 As Quantity,I.Price,I.Code,H.Name As HSNName,H.Code As HSNCode FROM BookMaster I INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE I.Type='F' ORDER BY I.Name", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
    End If
    rstItemList.ActiveConnection = Nothing
    If rstVchSeriesList.State = adStateOpen Then rstVchSeriesList.Close
    rstVchSeriesList.Open "SELECT Name As Col0,Prefix,Suffix,VchNumbering,Code FROM VchSeriesMaster WHERE Left(FYCode,2)='" & Left(FYCode, 2) & "' AND VchType ='" & Switch(VchType = "SQ", "24", VchType = "PQ", "23", VchType = "SO", "18", VchType = "PO", "17", VchType = "ST", "19") & VchType & "' ORDER BY Name", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
    rstVchSeriesList.ActiveConnection = Nothing
End Sub
Private Sub LoadOrderList()
    If rstOrderList.State = adStateOpen Then rstOrderList.Close
    rstOrderList.Open "SELECT VchCode,VchNo,VchDate,SUM(Quantity) As Ordered,SUM(Bal) As Bal FROM (SELECT VchCode,LTRIM(VchNo) As VchNo,VchDate,ABS(Quantity) As Quantity,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.RefCode)*1) As Bal FROM JobworkBVRef T WHERE RIGHT(VchType,2)='" & IIf(VchType = "SO", "SQ", "PQ") & "' AND Party='" & PartyCode & "') As Tbl WHERE Bal>0 GROUP BY VchCode,VchNo,VchDate ORDER BY VchNo", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
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
        rstOrderList.Open "SELECT ItemCode,ItemName,HSNCode,RefNo,RefId,Rate,[Disc%],SUM(Bal) As Bal FROM (SELECT I.Code As ItemCode,I.Name As ItemName,H.Code+'-'+H.Name As HSNCode,LTRIM(T.VchNo) As RefNo,T.RefCode As RefId,T.Rate,T.[Disc%],ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=T.RefCode)*1) As Bal FROM (JobworkBVRef T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE RIGHT(T.VchType,2)='" & IIf(VchType = "SO", "SQ", "PQ") & "' AND Method=1 AND T.VchCode IN (" & FrmOrderList.VchCodeList & ")) As Tbl WHERE Bal>0 GROUP BY ItemCode,ItemName,HSNCode,RefNo,RefId,Rate,[Disc%] ORDER BY ItemName,RefNo", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
        If rstOrderList.RecordCount > 0 Then
            i = 0
            With fpSpread1
                Do Until rstOrderList.EOF
                    i = i + 1
                    .SetText 1, i, rstOrderList.Fields("ItemName").Value
                    .SetText 2, i, Mid(rstOrderList.Fields("HSNCode").Value, InStr(1, rstOrderList.Fields("HSNCode").Value, "-") + 1, 40): .SetText 9, i, Left(rstOrderList.Fields("HSNCode").Value, InStr(1, rstOrderList.Fields("HSNCode").Value, "-") - 1)
                    .SetText 3, i, rstOrderList.Fields("RefNo").Value
                    .SetText 4, i, Val(rstOrderList.Fields("Bal").Value)
                    .SetText 5, i, Val(rstOrderList.Fields("Rate").Value)
                    .SetText 6, i, Val(rstOrderList.Fields("Disc%").Value)
                    Disc = (Val(rstOrderList.Fields("Rate").Value) * Val(rstOrderList.Fields("Disc%").Value)) / 100
                    .SetText 7, i, Val(rstOrderList.Fields("Bal").Value) * Round((Val(rstOrderList.Fields("Rate").Value) - Disc), 2)
                    .SetText 8, i, rstOrderList.Fields("ItemCode").Value
                    .SetText 10, i, rstOrderList.Fields("RefId").Value
                    .SetText 12, i, Val(rstOrderList.Fields("Bal").Value)
                    rstOrderList.MoveNext
                Loop
                Call CalculateTotal
            End With
            With rstOrderList
                If .State = adStateOpen Then .Close
                .Open "SELECT TOP 1 Transport,GRNo,GRDate,VehicleNo,Station FROM JobWorkBVParent WHERE Code IN (" & FrmOrderList.VchCodeList & ") AND (Transport<>'' AND Transport IS NOT NULL) ORDER BY Name", cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
                If .RecordCount > 0 Then
                    If MsgBox("Do u want to update Transport Details from order ref ('" & CheckNull(.Fields("Transport").Value) & "','" & CheckNull(.Fields("GRNo").Value) & "','" & CheckNull(.Fields("VehicleNo").Value) & "','" & CheckNull(.Fields("Station").Value) & "')?", vbYesNo + vbQuestion + vbDefaultButton1, "Update Transport Details !") = vbYes Then frmSalesOrderTptDetails.Text1.Text = CheckNull(.Fields("Transport").Value): frmSalesOrderTptDetails.Text2.Text = CheckNull(.Fields("GRNo").Value): frmSalesOrderTptDetails.Text3.Text = CheckNull(.Fields("VehicleNo").Value): frmSalesOrderTptDetails.Text4.Text = CheckNull(.Fields("Station").Value): If Not IsNull(.Fields("GRDate").Value) Then frmSalesOrderTptDetails.MhDateInput1.Value = .Fields("GRDate").Value
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
    VchCode = GenerateCode(cnSalesOrderVoucher, "SELECT MAX(Code) FROM JobworkBVParent", 6, "0")
    rstVchSeriesList.MoveFirst
    rstVchSeriesList.Find "[Code] = '" & rstSalesOrderVoucherList.Fields("VchSeriesCode").Value & "'"
    AutoVchNo = GenerateCode(cnSalesOrderVoucher, "SELECT MAX(CONVERT(INT,AutoVchNo))  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & rstSalesOrderVoucherList.Fields("VchSeriesCode").Value & "' AND FYCode='" & FYCode & "'", 10, Space(1))
    VchNo = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
    With cnSalesOrderVoucher
        .BeginTrans
        .Execute "SELECT * INTO #Tbl FROM JobworkBVParent Where Code = '" & rstSalesOrderVoucherList.Fields("Code").Value & "'"
        .Execute "UPDATE #Tbl SET Code='" & VchCode & "',Name='" & Trim(VchNo) & "',AutoVchNo='" & Pad(Trim(AutoVchNo), Space(1), 10, "L") & "',[Date]=GETDATE()"
        .Execute "INSERT INTO JobworkBVParent SELECT * FROM #Tbl"
        .Execute "DROP TABLE #Tbl"
        .Execute "SELECT * INTO #Tbl FROM JobworkBVChild Where Code = '" & rstSalesOrderVoucherList.Fields("Code").Value & "'"
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
    cnSalesOrderVoucher.RollbackTrans
    MdiMainMenu.MousePointer = vbNormal
    DisplayError ("Failed to Duplicate the Record")
End Sub
Private Sub btnNotes_Click()
    frmNotes.NotesFlag = 5
    frmNotes.Label1.Caption = "Notes : Voucher No. : " & Text2.Text
    frmNotes.Show vbModal
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Private Sub DisplayMenu(ByVal OutputTo As String)
    Dim menusel As String
    If rstSalesOrderVoucherList.RecordCount = 0 Then Exit Sub
    menusel = DisplayPopupMenu(Me.hwnd, 4)
    Select Case menusel
        Case 1
            Call PrintSalesOrderVoucher(rstSalesOrderVoucherList.Fields("Code").Value, Right(rstSalesOrderVoucherList.Fields("Type").Value, 2), OutputTo)
        Case 2
            Call PrintSalesOrderVoucher(rstSalesOrderVoucherList.Fields("Code").Value, Right(rstSalesOrderVoucherList.Fields("Type").Value, 2), OutputTo, False)
        End Select
    If Not (rstSalesOrderVoucherList.EOF Or rstSalesOrderVoucherList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.SetFocus
End Sub
Public Sub PrintSalesOrderVoucher(ByVal VchCode As String, ByVal VchType As String, Optional ByVal OutputType As String, Optional ByVal Original As Boolean = True)
    Dim TaxableVal As Double, RebateVal As Double, IGSTVal As Double, CGSTVal As Double, SGSTVal As Double
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    With rstSalesOrderVoucherChild
        If .State = adStateOpen Then .Close
        If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
        rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,EMail,Website,GSTIN,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,Prefix,Suffix FROM CompanyMaster P INNER JOIN CompChild C ON P.Code=C.Code WHERE VchType= " & Switch(Right(VchType, 2) = "SQ", "24", Right(VchType, 2) = "PQ", "23", Right(VchType, 2) = "SO", "18", Right(VchType, 2) = "PO", "17", Right(VchType, 2) = "ST", "19"), cnSalesOrderVoucher, adOpenKeyset, adLockReadOnly
        rstCompanyMaster.ActiveConnection = Nothing
        If Original Then
        If rstSalesOrderVoucherChild.State = adStateOpen Then rstSalesOrderVoucherChild.Close
            .Open "SELECT Distinct LTrim(P.Name)+'/' +'" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "' As BillNo,P.Date As BillDate,A.PrintName As Party,A.Address1 As PartyAddress1,A.Address2 As PartyAddress2,A.Address3 As PartyAddress3,A.Address4 As PartyAddress4,A.TIN As PartyGSTIN,A.Mobile As Mobile,A.EMail As EMail,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.PrintName,M.PrintName) As Consignee,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address1,M.Address1) As ConsigneeAddress1,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address2,M.Address2) As ConsigneeAddress2,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address3,M.Address3) As ConsigneeAddress3,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address4,M.Address4) As ConsigneeAddress4,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.TIN,M.TIN) As ConsigneeGSTIN," & _
            "IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Mobile,M.Mobile) As CMobile,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.EMail,M.EMail) As CEMail," & _
                         "P.[Rebate%],P.Rebate,P.Freight,P.Adjustment,P.TaxableAmount,P.[IGST%],P.IGST,P.[SGST%],P.SGST,P.[CGST%],P.CGST,P.Amount As TotalAmount,P.Remarks,'' As Narration,I.PrintName As Item,H.PrintName As HSNCode,ABS(C.Quantity) AS Quantity,C.Rate,C.Amount,'' As SrNo,'' As cmbTitle,LTRIM(C.Code)+LTRIM(C.SrNo) As Ref,C.[Disc%] As Disc,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM (((((JobworkBVParent P INNER JOIN JobworkBVChild C ON P.Code=C.Code) INNER JOIN BookMaster I ON C.Item=I.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster S ON P.Consignee=S.Code) LEFT JOIN AccountMaster M ON P.MaterialCentre=M.Code) LEFT JOIN GeneralMaster H ON C.HSNCode=H.Code WHERE P.Code='" + Left(VchCode, 6) + "' ORDER BY I.PrintName", cnSalesOrderVoucher, adOpenKeyset, adLockOptimistic
            If .RecordCount = 0 Then On Error GoTo 0: Screen.MousePointer = vbNormal: Exit Sub
        Else
            .Open "SELECT * FROM (SELECT LTrim(P.Name)+'/' +'" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "' As BillNo,P.Date As BillDate,A.PrintName As Party,A.Address1 As PartyAddress1,A.Address2 As PartyAddress2,A.Address3 As PartyAddress3,A.Address4 As PartyAddress4,A.TIN As PartyGSTIN,A.Mobile As Mobile,A.EMail As EMail,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.PrintName,M.PrintName) As Consignee,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address1,M.Address1) As ConsigneeAddress1,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address2,M.Address2) As ConsigneeAddress2,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address3,M.Address3) As ConsigneeAddress3,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address4,M.Address4) As ConsigneeAddress4,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.TIN,M.TIN) As ConsigneeGSTIN," & _
            "IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Mobile,M.Mobile) As CMobile,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.EMail,M.EMail) As CEMail," & _
                          "P.[Rebate%],0 As Rebate,0 As Freight,0 As Adjustment,0 As TaxableAmount,P.[IGST%],0 As IGST,P.[SGST%],0 As SGST,P.[CGST%],0 As CGST,0 As TotalAmount,P.Remarks,'' As Narration,I.PrintName As Item,H.PrintName As HSNCode,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=C.RefCode)*1) As Quantity,C.Rate,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=C.RefCode)*(C.Amount/C.Quantity)) As Amount,'' As SrNo,'' As cmbTitle,LTRIM(C.Code)+LTRIM(C.SrNo) As Ref,C.[Disc%] AS Disc FROM (((((JobworkBVParent P INNER JOIN JobworkBVChild C ON P.Code=C.Code) INNER JOIN BookMaster I ON C.Item=I.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster S ON P.Consignee=S.Code) LEFT JOIN AccountMaster M ON P.MaterialCentre=M.Code) LEFT JOIN GeneralMaster H ON C.HSNCode=H.Code WHERE P.Code='" + Left(VchCode, 6) + "') As Tbl WHERE Quantity>0 ORDER BY Item", cnSalesOrderVoucher, adOpenKeyset, adLockOptimistic
            If .RecordCount = 0 Then On Error GoTo 0: Screen.MousePointer = Normal: Exit Sub
            Do Until .EOF
                TaxableVal = TaxableVal + Val(.Fields("Amount").Value)
                .MoveNext
            Loop
            .MoveFirst
            RebateVal = Val(TaxableVal * Val(.Fields("Rebate%").Value) / 100)
            TaxableVal = TaxableVal - RebateVal
            IGSTVal = TaxableVal * Val(.Fields("IGST%").Value) / 100: CGSTVal = TaxableVal * Val(.Fields("CGST%").Value) / 100: SGSTVal = TaxableVal * Val(.Fields("SGST%").Value) / 100
            If .State = adStateOpen Then .Close
            .Open "SELECT * FROM (SELECT LTrim(P.Name)+'/' +'" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "' As BillNo,P.Date As BillDate,A.PrintName As Party,A.Address1 As PartyAddress1,A.Address2 As PartyAddress2,A.Address3 As PartyAddress3,A.Address4 As PartyAddress4,A.TIN As PartyGSTIN,A.Mobile As Mobile,A.EMail As EMail,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.PrintName,M.PrintName) As Consignee,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address1,M.Address1) As ConsigneeAddress1,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address2,M.Address2) As ConsigneeAddress2,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address3,M.Address3) As ConsigneeAddress3,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Address4,M.Address4) As ConsigneeAddress4,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.TIN,M.TIN) As ConsigneeGSTIN," & _
            "IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.Mobile,M.Mobile) As CMobile,IIF(CHARINDEX(RIGHT(P.Type,2),'SO_SQ_PO_PQ')>0,S.EMail,M.EMail) As CEMail," & _
                          "P.[Rebate%]," & RebateVal & " As Rebate,0 As Freight,0 As Adjustment," & TaxableVal & " As TaxableAmount,P.[IGST%]," & IGSTVal & " As IGST,P.[SGST%]," & SGSTVal & " As SGST,P.[CGST%]," & CGSTVal & " As CGST," & TaxableVal + IGSTVal + SGSTVal + CGSTVal & " As TotalAmount,P.Remarks,'' As Narration,I.PrintName As Item,H.PrintName As HSNCode,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=C.RefCode)*1) As Quantity,C.Rate,ABS((SELECT SUM(Quantity) FROM JobworkBVRef WHERE RefCode=C.RefCode)*(C.Amount/C.Quantity)) As Amount,'' As SrNo,'' As cmbTitle,LTRIM(C.Code)+LTRIM(C.SrNo) As Ref,C.[Disc%] AS Disc FROM (((((JobworkBVParent P INNER JOIN JobworkBVChild C ON P.Code=C.Code) INNER JOIN BookMaster I ON C.Item=I.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster S ON P.Consignee=S.Code) LEFT JOIN AccountMaster M ON P.MaterialCentre=M.Code) LEFT JOIN GeneralMaster H ON C.HSNCode=H.Code WHERE P.Code='" + Left(VchCode, 6) + "'" & _
                          ") As Tbl WHERE Quantity>0 ORDER BY Item", cnSalesOrderVoucher, adOpenKeyset, adLockOptimistic
        End If
        .ActiveConnection = Nothing
    End With
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
        .Text1.SetText IIf(Right(VchType, 2) = "SQ", "Sales Quotation", IIf(Right(VchType, 2) = "PQ", "Purchase Quotation", IIf(Right(VchType, 2) = "SO", "Sales", IIf(Right(VchType, 2) = "PO", "Purchase", "Stock Transfer")))) & IIf(Right(VchType, 2) = "PO", " Order", IIf(Right(VchType, 2) = "SO", " Order", ""))
        .Text13.SetText IIf(Right(VchType, 2) = "SO", "Buyer :", IIf(Right(VchType, 2) = "PO", "Supplier :", IIf(Right(VchType, 2) = "PQ", "Supplier :", IIf(Right(VchType, 2) = "SQ", "Buyer :", "From: Material Centre"))))
        .Text7.SetText IIf(Right(VchType, 2) = "SO" Or Right(VchType, 2) = "PO", "Consignee :", IIf(Right(VchType, 2) = "SQ" Or Right(VchType, 2) = "PQ", "Consignee :", "TO: Material Centre"))
        .Text35.SetText "Printed on " & Format(Now, "dd-MMM-yyyy") & " at " & Format(Now, "hh:mm")
        'If Len(LTrim(rstCompanyMaster.Fields("PrintName").Value)) <> 25 Then .Text2.Font.Size = 48
        .Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        .Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
        If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False)) Then
            .Text4.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value) & Space(1) & "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
        ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
            .Text4.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value)
        ElseIf Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False) Then
            .Text4.SetText "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
        End If
        .Text8.SetText "GSTIN/UIN : " & Trim(rstCompanyMaster.Fields("GSTIN").Value)
        .Text10.SetText "(" & UCase(Trim(NumberToWords(rstSalesOrderVoucherChild.Fields("TotalAmount").Value, False))) & ")"
        .Text11.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
        .Text26.SetText CheckNull(rstCompanyMaster.Fields("Declaration01").Value)
        .Text25.SetText CheckNull(rstCompanyMaster.Fields("Declaration02").Value)
        .Text22.SetText CheckNull(rstCompanyMaster.Fields("Declaration03").Value)
        .Text12.SetText CheckNull(rstCompanyMaster.Fields("Declaration04").Value)
        .Text9.SetText CheckNull(rstCompanyMaster.Fields("Declaration05").Value)
        .Text30.SetText CheckNull(rstCompanyMaster.Fields("Declaration06").Value)
        .Text33.SetText ""
        .Text31.SetText CheckNull(rstCompanyMaster.Fields("Declaration07").Value)
        .Database.SetDataSource rstSalesOrderVoucherChild, 3, 1
        .DiscardSavedData
        Screen.MousePointer = vbNormal
        If OutputType = "S" Then
            Set FrmReportViewer.Report = rptSalesOrderVoucher
            FrmReportViewer.Show vbModal
        Else
            If rstSalesOrderVoucherList.State = adStateClosed Then  'For Print Utility
                .PaperSource = crPRBinAuto
                .PrintOut False
            Else
                .PaperSource = crPRBinAuto
                .PrintOut
            End If
        End If
        Set rptSalesOrderVoucher = Nothing
    End With
    If rstSalesOrderVoucherList.State = adStateClosed Then  'For Print Utility
        Call CloseRecordset(rstCompanyMaster)
    End If
    Call CloseRecordset(rstSalesOrderVoucherChild)
    On Error GoTo 0
    Screen.MousePointer = vbNormal
End Sub
'Sales Quotation : -ve/Sales Order : +ve/Purchase Quotation : +ve/Purchase Order : -ve
