VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPaperIssueReceiptVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Receipt Voucher"
   ClientHeight    =   8265
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
   ScaleHeight     =   8265
   ScaleWidth      =   13740
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   8250
      Left            =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   13715
      _Version        =   65536
      _ExtentX        =   24192
      _ExtentY        =   14552
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
      Picture         =   "PaperIssueReceiptVoucher.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8030
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Width           =   13485
         _ExtentX        =   23786
         _ExtentY        =   14155
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
         TabPicture(0)   =   "PaperIssueReceiptVoucher.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Mh3dLabel1(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "DataGrid1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Text1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Option1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Option2"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "PaperIssueReceiptVoucher.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).ControlCount=   1
         Begin VB.OptionButton Option2 
            Caption         =   " Reel"
            Height          =   210
            Left            =   2640
            TabIndex        =   31
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   " Sheet"
            Height          =   210
            Left            =   1680
            TabIndex        =   30
            Top             =   20
            Value           =   -1  'True
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
            TabIndex        =   13
            Top             =   7590
            Width           =   8220
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7070
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   450
            Width           =   13260
            _ExtentX        =   23389
            _ExtentY        =   12462
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
               Caption         =   "        Vch No."
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
            BeginProperty Column02 
               DataField       =   "Account1Name"
               Caption         =   "Party Name"
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
            BeginProperty Column03 
               DataField       =   "Account2Name"
               Caption         =   "Supplier Name"
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
            BeginProperty Column05 
               DataField       =   "ChallanDate"
               Caption         =   "Challan Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd-MMM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "Ref"
               Caption         =   "Ref"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "dd-MM-yyyy"
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
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   4140.284
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   2940.095
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
                  ColumnWidth     =   1590.236
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1154.835
               EndProperty
               BeginProperty Column06 
                  Locked          =   -1  'True
                  ColumnWidth     =   764.787
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7340
            Left            =   -74880
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   480
            Width           =   13260
            _Version        =   65536
            _ExtentX        =   23389
            _ExtentY        =   12947
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
            Picture         =   "PaperIssueReceiptVoucher.frx":0054
            Begin VB.TextBox Text7 
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
               Left            =   7175
               MaxLength       =   10
               TabIndex        =   5
               Top             =   950
               Width           =   1500
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
               Left            =   7175
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   3
               Top             =   630
               Width           =   5980
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   19
               Top             =   6870
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
               Picture         =   "PaperIssueReceiptVoucher.frx":0070
               Picture         =   "PaperIssueReceiptVoucher.frx":008C
               Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
                  Height          =   285
                  Left            =   8175
                  TabIndex        =   20
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
                  _ExtentY        =   503
                  Calculator      =   "PaperIssueReceiptVoucher.frx":00A8
                  Caption         =   "PaperIssueReceiptVoucher.frx":00C8
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperIssueReceiptVoucher.frx":0134
                  Keys            =   "PaperIssueReceiptVoucher.frx":0152
                  Spin            =   "PaperIssueReceiptVoucher.frx":019C
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "#######0.000"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "#######0.000"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   99999999.999
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
               Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
                  Height          =   285
                  Left            =   9150
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1035
                  _Version        =   65536
                  _ExtentX        =   1834
                  _ExtentY        =   503
                  Calculator      =   "PaperIssueReceiptVoucher.frx":01C4
                  Caption         =   "PaperIssueReceiptVoucher.frx":01E4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperIssueReceiptVoucher.frx":0250
                  Keys            =   "PaperIssueReceiptVoucher.frx":026E
                  Spin            =   "PaperIssueReceiptVoucher.frx":02B8
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "######0"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "######0"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   99999999
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
               Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
                  Height          =   285
                  Left            =   11000
                  TabIndex        =   23
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   660
                  _Version        =   65536
                  _ExtentX        =   1173
                  _ExtentY        =   503
                  Calculator      =   "PaperIssueReceiptVoucher.frx":02E0
                  Caption         =   "PaperIssueReceiptVoucher.frx":0300
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperIssueReceiptVoucher.frx":036C
                  Keys            =   "PaperIssueReceiptVoucher.frx":038A
                  Spin            =   "PaperIssueReceiptVoucher.frx":03D4
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "######0"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "######0"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   9999999
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
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   0
               Top             =   105
               Width           =   1530
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
               Left            =   1200
               MaxLength       =   40
               TabIndex        =   4
               Top             =   950
               Width           =   4660
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
               TabIndex        =   2
               Top             =   630
               Width           =   4660
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   16
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
               Caption         =   " Vch No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperIssueReceiptVoucher.frx":03FC
               Picture         =   "PaperIssueReceiptVoucher.frx":0418
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   17
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
               Caption         =   " Party Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperIssueReceiptVoucher.frx":0434
               Picture         =   "PaperIssueReceiptVoucher.frx":0450
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   18
               Top             =   950
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
               Caption         =   " Location"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperIssueReceiptVoucher.frx":046C
               Picture         =   "PaperIssueReceiptVoucher.frx":0488
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   12060
               TabIndex        =   1
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "PaperIssueReceiptVoucher.frx":04A4
               Caption         =   "PaperIssueReceiptVoucher.frx":05BC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperIssueReceiptVoucher.frx":0628
               Keys            =   "PaperIssueReceiptVoucher.frx":0646
               Spin            =   "PaperIssueReceiptVoucher.frx":06A4
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
               Height          =   5415
               Left            =   120
               TabIndex        =   8
               Top             =   1470
               Width           =   13035
               _Version        =   524288
               _ExtentX        =   22992
               _ExtentY        =   9551
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
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "PaperIssueReceiptVoucher.frx":06CC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   10860
               TabIndex        =   21
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
               Picture         =   "PaperIssueReceiptVoucher.frx":1284
               Picture         =   "PaperIssueReceiptVoucher.frx":12A0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   5850
               TabIndex        =   24
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
               Caption         =   " Supplier Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperIssueReceiptVoucher.frx":12BC
               Picture         =   "PaperIssueReceiptVoucher.frx":12D8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   5850
               TabIndex        =   25
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
               Caption         =   " Challan No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperIssueReceiptVoucher.frx":12F4
               Picture         =   "PaperIssueReceiptVoucher.frx":1310
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput2 
               Height          =   330
               Left            =   9780
               TabIndex        =   6
               Top             =   945
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "PaperIssueReceiptVoucher.frx":132C
               Caption         =   "PaperIssueReceiptVoucher.frx":1444
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperIssueReceiptVoucher.frx":14B0
               Keys            =   "PaperIssueReceiptVoucher.frx":14CE
               Spin            =   "PaperIssueReceiptVoucher.frx":152C
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   8655
               TabIndex        =   26
               Top             =   945
               Width           =   1135
               _Version        =   65536
               _ExtentX        =   2002
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
               Caption         =   " Challan Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperIssueReceiptVoucher.frx":1554
               Picture         =   "PaperIssueReceiptVoucher.frx":1570
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   10860
               TabIndex        =   27
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
               Caption         =   " Cartage"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperIssueReceiptVoucher.frx":158C
               Picture         =   "PaperIssueReceiptVoucher.frx":15A8
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
               Height          =   330
               Left            =   12060
               TabIndex        =   7
               Top             =   945
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calculator      =   "PaperIssueReceiptVoucher.frx":15C4
               Caption         =   "PaperIssueReceiptVoucher.frx":15E4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperIssueReceiptVoucher.frx":1650
               Keys            =   "PaperIssueReceiptVoucher.frx":166E
               Spin            =   "PaperIssueReceiptVoucher.frx":16B8
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
            Begin VB.TextBox Text6 
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
               Left            =   5160
               MaxLength       =   100
               TabIndex        =   29
               Top             =   3600
               Width           =   1500
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
            TabIndex        =   28
            Top             =   7590
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
            Picture         =   "PaperIssueReceiptVoucher.frx":16E0
            Picture         =   "PaperIssueReceiptVoucher.frx":16FC
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
            TabIndex        =   14
            Top             =   7590
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   10
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
Attribute VB_Name = "frmPaperIssueReceiptVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnPaperIRVch As New ADODB.Connection
Dim rstPaperIRVList As New ADODB.Recordset, rstPaperIRVParent As New ADODB.Recordset, rstPaperIRVChild As New ADODB.Recordset, rstOrderList As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstPaperList As New ADODB.Recordset, rstCompanyMaster As New ADODB.Recordset
Dim Account1Code As String, Account2Code As String, SortOrder, PrevStr, dblBookMark As Double, blnRecordExist As Boolean, EditMode As Boolean, PaperCode As String
Public VchType As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    cnPaperIRVch.CursorLocation = adUseClient
    cnPaperIRVch.Open cnDatabase.ConnectionString
    Me.Caption = "Paper " & IIf(VchType = "T", "Transfer", IIf(VchType = "I", "Issue", "Receipt")) & " Voucher"
    If VchType <> "R" Then DataGrid1.Columns(2).Caption = " Source": DataGrid1.Columns(3).Caption = " Destination": Mh3dLabel3.Caption = " Source": Mh3dLabel2.Caption = " Destination"
    If VchType = "R" Then '(Select Name From PaperPOParent Where Code=Ref) As Ref
        rstPaperIRVList.Open "SELECT DISTINCT P.Code,P.Name,Date,M1.Name As Account1Name,M2.Name As Account2Name,BiltyNo As ChallanNo,BiltyDate As ChallanDate,'' As Ref FROM ((PaperPOParent P Left JOIN PaperIOChild C ON P.Code=C.Code) Left JOIN AccountMaster M1 ON C.Account=M1.Code) Left JOIN AccountMaster M2 ON P.Supplier=M2.Code WHERE OrderType='" & VchType & "' AND FYCode='" & FYCode & "' ORDER BY P.Name", cnPaperIRVch, adOpenKeyset, adLockOptimistic
    Else
        rstPaperIRVList.Open "SELECT DISTINCT T.Code,T.Name,Date,M1.Name As Account1Name,M2.Name As Account2Name,BiltyNo As ChallanNo,BiltyDate As ChallanDate,'' As Ref  FROM (PaperMVParent T INNER JOIN AccountMaster M1 ON T.AccountFrom=M1.Code) INNER JOIN AccountMaster M2 ON T.AccountTo=M2.Code WHERE [Type]='" & VchType & "' AND FYCode='" & FYCode & "' ORDER BY T.Name", cnPaperIRVch, adOpenKeyset, adLockOptimistic
    End If
    rstPaperIRVParent.CursorLocation = adUseClient
    rstPaperIRVList.Filter = adFilterNone
    If rstPaperIRVList.RecordCount > 0 Then rstPaperIRVList.MoveLast
    Set DataGrid1.DataSource = rstPaperIRVList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstPaperIRVList.EOF Or rstPaperIRVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstPaperIRVList.ActiveConnection = Nothing
    LoadMasterList
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    MdiMainMenu.mnuPaperModule(2).Enabled = False: MdiMainMenu.mnuPaperModule(3).Enabled = False: MdiMainMenu.mnuPaperModule(4).Enabled = False
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If SSTab1.Tab = 0 Then
            Unload Me
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
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(9)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(10)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyM And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(11)
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
            SSTab1.Tab = 1: SSTab1.SetFocus
        Else
           If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstPaperIRVList)
    Call CloseRecordset(rstPaperIRVParent)
    Call CloseRecordset(rstPaperIRVChild)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstOrderList)
    Call CloseConnection(cnPaperIRVch)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuPaperModule(2).Enabled = True: MdiMainMenu.mnuPaperModule(3).Enabled = True: MdiMainMenu.mnuPaperModule(4).Enabled = True
End Sub
Private Sub Text1_Change()
On Error Resume Next
With rstPaperIRVList
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
    If rstPaperIRVList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstPaperIRVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstPaperIRVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstPaperIRVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstPaperIRVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstPaperIRVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstPaperIRVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstPaperIRVList
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
            ViewRecord
        Else
            If Not (rstPaperIRVList.EOF Or rstPaperIRVList.BOF) Then
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
        Text3.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    Dim CellVal01 As Variant, CellVal02 As Variant, i As Integer
    If Button.Index = 1 Then
        If rstPaperIRVParent.State = adStateOpen Then rstPaperIRVParent.Close
        rstPaperIRVParent.Open "SELECT * FROM " & IIf(VchType = "R", "PaperPOParent", "PaperMVParent") & " WHERE Code=''", cnPaperIRVch, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstPaperIRVParent) Then
            Text2.Text = GenerateCode(cnPaperIRVch, "SELECT MAX(CONVERT(INT,Name)) FROM  " & IIf(VchType = "R", "PaperPOParent", "PaperMVParent") & " WHERE " & IIf(VchType = "R", "[OrderType]", "[Type]") & "='" & VchType & "' AND FYCode='" & FYCode & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text3.SetFocus
            blnRecordExist = False
            cnPaperIRVch.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstPaperIRVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstPaperIRVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnPaperIRVch.BeginTrans
            cnPaperIRVch.Execute "DELETE FROM " & IIf(VchType = "R", "PaperPOParent", "PaperMVParent") & " WHERE Code='" & rstPaperIRVList.Fields("Code").Value & "'"
            cnPaperIRVch.Execute "DELETE FROM " & IIf(VchType = "R", "PaperIOChild", "PaperIOChild") & " WHERE Code='" & rstPaperIRVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstPaperIRVList.Delete
                rstPaperIRVList.MoveNext
                If rstPaperIRVList.RecordCount > 0 And rstPaperIRVList.EOF Then rstPaperIRVList.MoveLast
                cnPaperIRVch.CommitTrans
                ShowProgressInStatusBar True
                Timer1.Enabled = True
            Else
                DisplayError ("Failed to delete the record")
                cnPaperIRVch.RollbackTrans
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
            Call DisplayError("You don't have the rights to Edit this Voucher")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstPaperIRVParent) Then
            If UpdatePaperList("D") Then
                UpdateFlag = 1
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 3, i
                        .GetText 2, i, CellVal01    'Quantity
                        .GetText 10, i, CellVal02    'Paper Code
                        If CellVal02 <> "" Then
                            If Not UpdatePaperList("I") Then UpdateFlag = 0: Exit For
                        End If
                    Next
                End With
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnPaperIRVch.CommitTrans
            If rstPaperIRVParent.State = adStateOpen Then rstPaperIRVParent.Close
            rstPaperIRVParent.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstPaperIRVParent) Then
            cnPaperIRVch.RollbackTrans
            If rstPaperIRVParent.State = adStateOpen Then rstPaperIRVParent.Close
            rstPaperIRVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstPaperIRVList.Filter = adFilterNone
        rstPaperIRVList.ActiveConnection = cnPaperIRVch
        Do While Not RefreshRecord(rstPaperIRVList): Loop
        Set DataGrid1.DataSource = rstPaperIRVList
        rstPaperIRVList.ActiveConnection = Nothing
        If rstPaperIRVList.RecordCount > 0 Then rstPaperIRVList.MoveLast
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Party", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstPaperIRVList.RecordCount = 0 Then Exit Sub
        Call PrintPaperIRVch(rstPaperIRVList.Fields("Code").Value, "P")
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstPaperIRVList.RecordCount = 0 Then Exit Sub
        Call PrintPaperIRVch(rstPaperIRVList.Fields("Code").Value, "S")
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstPaperIRVList.RecordCount > 0 Then rstPaperIRVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstPaperIRVList.RecordCount > 0 Then
            rstPaperIRVList.MovePrevious
            If rstPaperIRVList.BOF Then rstPaperIRVList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstPaperIRVList.RecordCount > 0 Then
            rstPaperIRVList.MoveNext
            If rstPaperIRVList.EOF Then rstPaperIRVList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstPaperIRVList.RecordCount > 0 Then rstPaperIRVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstPaperIRVList.EOF Or rstPaperIRVList.BOF) Then
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
        rstPaperIRVList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstPaperIRVList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstPaperIRVList.EOF Or rstPaperIRVList.BOF) Then
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
    Toolbar1.Buttons.Item(11).Enabled = bVal
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstPaperIRVList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(9).Enabled = False
        Toolbar1.Buttons.Item(10).Enabled = False
        Toolbar1.Buttons.Item(11).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstPaperIRVParent.EOF Or rstPaperIRVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnPaperIRVch, IIf(VchType = "R", "PaperPOParent", "PaperMVParent"), "Code", IIf(VchType = "R", "[Name]+[OrderType]", "[Name]+[Type]"), Trim(Text2.Text) & VchType, rstPaperIRVParent.Fields("Code").Value, False, FYCode) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If MhDateInput2.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput2.Text)) Then Cancel = True
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = Account1Code
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        Account1Code = slCode: Text3.Text = slName
        If Not CheckEmpty(Account1Code, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If CheckEmpty(Text3.Text, False) Then Cancel = True
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = Account2Code
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        Account2Code = slCode: Text5.Text = slName
        If Not CheckEmpty(Account2Code, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    If CheckEmpty(Text5.Text, False) Then Cancel = True
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstPaperIRVList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstPaperIRVParent.State = adStateOpen Then rstPaperIRVParent.Close
    If VchType = "R" Then
        rstPaperIRVParent.Open "SELECT DISTINCT P.*,C.Account As Party FROM PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code WHERE P.Code='" & FixQuote(rstPaperIRVList.Fields("Code").Value) & "'", cnPaperIRVch, adOpenKeyset, adLockOptimistic
    Else
        rstPaperIRVParent.Open "SELECT * FROM PaperMVParent WHERE Code='" & FixQuote(rstPaperIRVList.Fields("Code").Value) & "'", cnPaperIRVch, adOpenKeyset, adLockOptimistic
    End If
    If rstPaperIRVParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = "" 'Vch No.
    Text3.Text = "" 'Party Name/From Party
    Text5.Text = "" 'Supplier Name/To Party
    Text7.Text = "" 'Challan No.
    Text4.Text = "" 'Remarks
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput2.Text = "  -  -    "    'Challan Date
    MhRealInput19.Value = 0 'Total Quantity (Kg)
    MhRealInput20.Value = 0 'Total Quantity (Sheets)
    MhRealInput21.Value = 0 'Total Bundles
    MhRealInput14.Value = 0 'Cartage
    With fpSpread1
    .ClearRange 1, 1, .MaxCols, .MaxRows, True: .SetActiveCell 1, 1
        If Option2.Value Then
            .ColWidth(1) = 64.125
            .Col = 2: .ColHidden = True
            .Col = 3: .ColHidden = True
            .Col = 4: .ColHidden = True
            .ColWidth(5) = 8
            .ColWidth(6) = 11
            .Col = 7: .ColHidden = True
            .ColWidth(8) = 9.625
            .SetText 8, 0, "Reels"
            .ColWidth(9) = 9.25
            MhRealInput21.Width = 1200
            MhRealInput21.Left = 10490
            MhRealInput20.Width = 1350
        ElseIf Option1.Value Then
            .ColWidth(1) = 41
            .Col = 2: .ColHidden = False
            .Col = 3: .ColHidden = False
            .Col = 4: .ColHidden = False
            .ColWidth(5) = 8
            .ColWidth(6) = 8.375
            .ColWidth(7) = 6.625
            .Col = 7: .ColHidden = False
            .SetText 8, 0, "Bundles"
            MhRealInput21.Width = 660
            MhRealInput21.Left = 11000
            MhRealInput20.Width = 1200
            .ColWidth(9) = 9.25
        End If
    End With
    Account1Code = "": Account2Code = ""
End Sub
Private Sub LoadFields()
    If rstPaperIRVParent.EOF Or rstPaperIRVParent.BOF Then Exit Sub
    Text2.Text = rstPaperIRVParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstPaperIRVParent.Fields("Date").Value, "dd-MM-yyyy")
    If VchType = "R" Then
        Account1Code = rstPaperIRVParent.Fields("Party").Value
        Account2Code = rstPaperIRVParent.Fields("Supplier").Value
    Else
        Account1Code = rstPaperIRVParent.Fields("AccountFrom").Value
        Account2Code = rstPaperIRVParent.Fields("AccountTo").Value
    End If
    If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & Account1Code & "'"
    If Not rstAccountList.EOF Then Text3.Text = rstAccountList.Fields("Col0").Value
    If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & Account2Code & "'"
    If Not rstAccountList.EOF Then Text5.Text = rstAccountList.Fields("Col0").Value
    Text7.Text = CheckNull(rstPaperIRVParent.Fields("BiltyNo").Value)
    If Not IsNull(rstPaperIRVParent.Fields("BiltyDate").Value) Then MhDateInput2.Text = Format(rstPaperIRVParent.Fields("BiltyDate").Value, "dd-MM-yyyy")
    MhRealInput14.Value = Val(rstPaperIRVParent.Fields("BiltyAmount").Value)
    Text4.Text = rstPaperIRVParent.Fields("Remarks").Value
    Call LoadPaperList(rstPaperIRVParent.Fields("Code").Value)
    CalculateTotal
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstPaperIRVParent.RecordCount = 0 Then Exit Sub
    If rstPaperIRVParent.State = adStateOpen Then rstPaperIRVParent.Close
    rstPaperIRVParent.CursorLocation = adUseServer
    rstPaperIRVParent.Open "SELECT * FROM " & IIf(VchType = "R", "PaperPOParent", "PaperMVParent") & " WHERE Code='" & FixQuote(rstPaperIRVList.Fields("Code").Value) & "'", cnPaperIRVch, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstPaperIRVParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text3.SetFocus
    blnRecordExist = True
    cnPaperIRVch.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstPaperIRVParent.EOF Or rstPaperIRVParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstPaperIRVParent.Fields("Code").Value = GenerateCode(cnPaperIRVch, "SELECT MAX(Code) FROM " & IIf(VchType = "R", "PaperPOParent", "PaperMVParent"), 6, "0")
        If VchType = "R" Then rstPaperIRVParent.Fields("DeliveryDate").Value = Now()
        rstPaperIRVParent.Fields("CreatedBy").Value = UserCode
        rstPaperIRVParent.Fields("CreatedOn").Value = Now()
        rstPaperIRVParent.Fields("Recordstatus").Value = "N"
    Else
        rstPaperIRVParent.Fields("ModifiedBy").Value = UserCode
        rstPaperIRVParent.Fields("ModifiedOn").Value = Now()
        rstPaperIRVParent.Fields("Recordstatus").Value = "M"
    End If
    rstPaperIRVParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstPaperIRVParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstPaperIRVParent.Fields("BiltyNo").Value = Trim(Text7.Text)
    If MhDateInput2.ValueIsNull Then rstPaperIRVParent.Fields("BiltyDate").Value = Null Else rstPaperIRVParent.Fields("BiltyDate").Value = GetDate(MhDateInput2.Text)
    rstPaperIRVParent.Fields("BiltyAmount").Value = MhRealInput14.Value
    If VchType = "R" Then
        rstPaperIRVParent.Fields("Supplier").Value = Account2Code
        rstPaperIRVParent.Fields("GST%").Value = 0
        rstPaperIRVParent.Fields("GST").Value = 0
        rstPaperIRVParent.Fields("Cartage/Kg").Value = 0
        rstPaperIRVParent.Fields("Cartage").Value = 0
        rstPaperIRVParent.Fields("Adjustment").Value = 0
        rstPaperIRVParent.Fields("BillAmount").Value = 0
        rstPaperIRVParent.Fields("PaidAmount").Value = 0
        rstPaperIRVParent.Fields("DeliveryEndDate").Value = GetDate(MhDateInput2.Text)
        rstPaperIRVParent.Fields("OrderType").Value = VchType
    Else
        rstPaperIRVParent.Fields("AccountFrom").Value = Account1Code
        rstPaperIRVParent.Fields("AccountTo").Value = Account2Code
        rstPaperIRVParent.Fields("Type").Value = VchType
    End If
    rstPaperIRVParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstPaperIRVParent.Fields("FYCode").Value = FYCode
    rstPaperIRVParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstPaperIRVList.MoveFirst
    rstPaperIRVList.Find "[Code] = '" & rstPaperIRVParent.Fields("Code").Value & "'"
    If rstPaperIRVList.EOF Then rstPaperIRVList.AddNew
    rstPaperIRVList.Fields("Code").Value = rstPaperIRVParent.Fields("Code").Value
    rstPaperIRVList.Fields("Name").Value = Pad(rstPaperIRVParent.Fields("Name").Value, Space(1), 10, "L")
    rstPaperIRVList.Fields("Date").Value = rstPaperIRVParent.Fields("Date").Value
    rstPaperIRVList.Fields("Account1Name").Value = Trim(Text3.Text)
    rstPaperIRVList.Fields("Account2Name").Value = Trim(Text5.Text)
    rstPaperIRVList.Fields("ChallanNo").Value = Trim(Text7.Text)
    rstPaperIRVList.Fields("ChallanDate").Value = Format(MhDateInput2.Text, "dd-MM-yyyy")
    rstPaperIRVList.Update
    rstPaperIRVList.Sort = SortOrder & " Asc"
    rstPaperIRVList.Find "[Code] = '" & rstPaperIRVParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        DisplayError ("Voucher No. cannot be blank"): Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnPaperIRVch, IIf(VchType = "R", "PaperPOParent", "PaperMVParent"), "Code", IIf(VchType = "R", "[Name]+[OrderType]", "[Name]+[Type]"), Trim(Text2.Text) & VchType, rstPaperIRVParent.Fields("Code").Value, False, FYCode) Then
        Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf fpSpread1.DataRowCnt = 0 Then
        DisplayError ("Blank Voucher cannot be saved"): fpSpread1.SetFocus
        CheckMandatoryFields = True: Exit Function
    End If
    If VchType = "R" Then
        If CheckEmpty(Text7.Text, False) Then
            Text7.SetFocus:   CheckMandatoryFields = True
        ElseIf MhDateInput2.ValueIsNull Then
            MhDateInput2.SetFocus:   CheckMandatoryFields = True
        End If
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
Private Sub LoadPaperList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstPaperIRVChild.State = adStateOpen Then rstPaperIRVChild.Close
    If VchType = "R" Then
        rstPaperIRVChild.Open "SELECT LTRIM(I.Name)+IIF(I.Form='S',' (UOM : '+LTRIM(U.Name)+'='+LTRIM(CONVERT(INT,U.Value1))+')','')  As PaperName,Paper As PaperCode,Quantity,U.Name As UOMName,T.[Weight/Unit],QuantityKg,QuantitySheets,T.[Units/Bundle],TotalBundles,ISNULL(Ref,'') As RefCode,ISNULL((SELECT LTRIM(Name) FROM PaperPOParent WHERE Code=Ref),'') As RefNo,U.Value1 As SPU,ISNULL((SELECT QuantitySheets-ISNULL((SELECT SUM(QuantitySheets) FROM PaperIOChild WHERE Ref+Paper=T.Ref+T.Paper AND Code<>T.Code),0) FROM PaperPOChild C WHERE C.Code+C.Paper=T.Ref+T.Paper),0) As Bal FROM (PaperIOChild T INNER JOIN PaperMaster I ON T.Paper=I.Code) INNER JOIN GeneralMaster U ON I.UOM=U.Code WHERE T.Code='" & strOrderCode & "' ORDER BY I.Name", cnPaperIRVch, adOpenKeyset, adLockOptimistic
    Else
        rstPaperIRVChild.Open "SELECT LTRIM(M.Name)+IIF(M.Form='S',' (UOM : '+LTRIM(U.Name)+'='+LTRIM(CONVERT(INT,U.Value1))+')','') As PaperName,Paper As PaperCode,Quantity,U.Name As UOMName,T.[Weight/Unit],QuantityKg,QuantitySheets,T.[Units/Bundle],TotalBundles,'' As RefCode,'' As RefNo,U.Value1 As SPU,0 As Bal FROM (PaperMVChild T INNER JOIN PaperMaster M ON T.Paper=M.Code) INNER JOIN GeneralMaster U ON M.UOM=U.Code WHERE T.Code='" & strOrderCode & "' ORDER BY M.Name", cnPaperIRVch, adOpenKeyset, adLockOptimistic
    End If
    rstPaperIRVChild.ActiveConnection = Nothing
    If rstPaperIRVChild.RecordCount > 0 Then rstPaperIRVChild.MoveFirst
    i = 0
    Do While Not rstPaperIRVChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstPaperIRVChild.Fields("PaperName").Value
            .SetText 2, i, Val(rstPaperIRVChild.Fields("Quantity").Value)
            .SetText 3, i, rstPaperIRVChild.Fields("UOMName").Value
            .SetText 4, i, Val(rstPaperIRVChild.Fields("Weight/Unit").Value)
            .SetText 5, i, Val(rstPaperIRVChild.Fields("QuantityKg").Value)
            .SetText 6, i, Val(rstPaperIRVChild.Fields("QuantitySheets").Value)
            .SetText 7, i, Val(rstPaperIRVChild.Fields("Units/Bundle").Value)
            .SetText 8, i, Val(rstPaperIRVChild.Fields("TotalBundles").Value)
            .SetText 9, i, rstPaperIRVChild.Fields("RefNo").Value
            .SetText 10, i, rstPaperIRVChild.Fields("PaperCode").Value
            .SetText 11, i, Val(rstPaperIRVChild.Fields("SPU").Value)
            .SetText 12, i, rstPaperIRVChild.Fields("RefCode").Value
            .SetText 13, i, Val(rstPaperIRVChild.Fields("Bal").Value)
        End With
        rstPaperIRVChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Paper List")
End Sub
Private Sub CalculateTotal()
    Dim i As Integer, QtyKg As Variant, QtySheets As Variant, Bdl As Variant
    MhRealInput19.Value = 0: MhRealInput20.Value = 0: MhRealInput21.Value = 0
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 5, i, QtyKg
            .GetText 6, i, QtySheets
            .GetText 8, i, Bdl
            MhRealInput19.Value = Val(MhRealInput19.Text) + Val(QtyKg)
            MhRealInput20.Value = Val(MhRealInput20.Text) + Val(QtySheets)
            MhRealInput21.Value = Val(MhRealInput21.Text) + Val(Bdl)
        Next
    End With
End Sub
Private Function UpdatePaperList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 8) As Variant
    On Error GoTo ErrorHandler
    UpdatePaperList = True
    If ActionType = "D" Then
        If Not blnRecordExist Then Exit Function
        cnPaperIRVch.Execute "DELETE FROM " & IIf(VchType = "R", "PaperIOChild", "PaperMVChild") & " WHERE Code='" & rstPaperIRVParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 2, .ActiveRow, CellVal(1)  'Quantity
            .GetText 4, .ActiveRow, CellVal(2)  'Wt/Unit
            .GetText 5, .ActiveRow, CellVal(3)  'Quantity (Kg)
            .GetText 6, .ActiveRow, CellVal(4)  'Quantity (Sheets)
            .GetText 7, .ActiveRow, CellVal(5)  'Unit/Bdl
            .GetText 8, .ActiveRow, CellVal(6)  'Total Bundles
            .GetText 10, .ActiveRow, CellVal(7)  'Paper Code
            .GetText 12, .ActiveRow, CellVal(8) 'Ref Code
        End With
        If VchType = "R" Then
            cnPaperIRVch.Execute "INSERT INTO PaperIOChild VALUES ('" & rstPaperIRVParent.Fields("Code").Value & "','" & CellVal(7) & "','" & Account1Code & "'," & Val(CellVal(2)) & "," & Val(CellVal(1)) & "," & Val(CellVal(4)) & "," & Val(CellVal(3)) & "," & Val(CellVal(5)) & "," & Val(CellVal(6)) & ",Null," & IIf(CheckEmpty(CellVal(8), False), "Null", "'" & CellVal(8) & "'") & ")"
        Else
            cnPaperIRVch.Execute "INSERT INTO PaperMVChild VALUES ('" & rstPaperIRVParent.Fields("Code").Value & "','" & CellVal(7) & "'," & Val(CellVal(2)) & "," & Val(CellVal(1)) & "," & Val(CellVal(4)) & "," & Val(CellVal(3)) & "," & Val(CellVal(5)) & "," & Val(CellVal(6)) & ")"
        End If
    End If
    Exit Function
ErrorHandler:
    UpdatePaperList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Party" Then rstPaperIRVList.Filter = "[Account1Name] Like '%" & SrchText & "%'"
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RefCode As Variant, Qty As Variant, BalQty As Double, SPU As Variant
    If KeyCode = vbKeyF9 Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then fpSpread1.DeleteRows fpSpread1.ActiveRow, 1: fpSpread1.SetFocus: CalculateTotal
    ElseIf KeyCode = vbKeySpace Then
        With fpSpread1
            If .ActiveCol = 1 Then
                If VchType = "R" Then
                    .GetText 12, .ActiveRow, RefCode 'Ref Code
                    If Not CheckEmpty(RefCode, False) Then Exit Sub
                    .GetText 10, .ActiveRow, RefCode 'Paper Code
                    On Error Resume Next
                    FrmPaperMaster.SL = True
                    FrmPaperMaster.MasterCode = RefCode
                    Load FrmPaperMaster
                    If Err.Number <> 364 Then FrmPaperMaster.Show vbModal
                    On Error GoTo 0
                    .SetText 1, .ActiveRow, slName
                    .SetText 10, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then
                        LoadMasterList
                        rstPaperList.MoveFirst: rstPaperList.Find "[Code] ='" & slCode & "'"
                        .SetText 3, .ActiveRow, rstPaperList.Fields("UOMName").Value
                        .SetText 4, .ActiveRow, Val(rstPaperList.Fields("Weight/Unit").Value)
                        .SetText 7, .ActiveRow, Val(rstPaperList.Fields("Units/Bundle").Value)
                        .SetText 11, .ActiveRow, Val(rstPaperList.Fields("SPU").Value)
                        Sendkeys "{ENTER}"
                    Else
                        .SetActiveCell 1, .ActiveRow
                    End If
                Else
                    .GetText 1, .ActiveRow, RefCode: Text6.Text = FixQuote(RefCode) 'Paper Name
                    LoadMasterList True
                    If rstPaperList.RecordCount = 0 Then DisplayError ("No Paper Exists"): .SetActiveCell 1, .ActiveRow: .SetFocus: Exit Sub Else rstPaperList.MoveFirst
                    rstPaperList.Find "[Col0] = '" & Trim(Text6.Text) & "'"
                    SelectionType = "S": PaperCode = ""
                    Call LoadSelectionList(rstPaperList, "List of Papers...", "Name", "       Stock")
                    SearchOrder = 0
                    Call DisplaySelectionList(Text6, PaperCode)
                    Call CloseForm(FrmSelectionList)
                    If Not CheckEmpty(Trim(PaperCode), False) Then
                        .SetText 1, .ActiveRow, Text6.Text
                        .SetText 10, .ActiveRow, PaperCode
                        rstPaperList.MoveFirst: rstPaperList.Find "[Code] ='" & PaperCode & "'"
                        .SetText 3, .ActiveRow, rstPaperList.Fields("UOMName").Value
                        .SetText 4, .ActiveRow, Val(rstPaperList.Fields("Weight/Unit").Value)
                        .SetText 7, .ActiveRow, Val(rstPaperList.Fields("Units/Bundle").Value)
                        .SetText 11, .ActiveRow, Val(rstPaperList.Fields("SPU").Value)
                        .SetText 13, .ActiveRow, U2S(Val(rstPaperList.Fields("Col1").Value), Val(rstPaperList.Fields("SPU").Value))
                        Sendkeys "{ENTER}"
                    Else
                        .SetActiveCell 1, .ActiveRow
                    End If
                End If
            End If
        End With
    ElseIf KeyCode = vbKeyReturn Then
        With fpSpread1
            If .ActiveCol = 2 Then
                .GetText 12, .ActiveRow, RefCode 'Ref Code
                If VchType = "R" And CheckEmpty(RefCode, False) Then Exit Sub
                .GetText 13, .ActiveRow, Qty: BalQty = Val(Qty)
                If BalQty > 0 Then
                    .GetText 11, .ActiveRow, SPU
                    .GetText 2, .ActiveRow, Qty
                    If U2S(Val(Qty), SPU) > BalQty Then DisplayError ("Received quantity (" & Format(Val(Qty), "0.000") & ") is more than Balance quantity (" & Format(S2U(BalQty, Val(SPU)), "0.000") & ")")
                    EditMode = False
                    .SetFocus
                End If
            End If
        End With
    ElseIf KeyCode = vbKeyF11 Then
        If fpSpread1.DataRowCnt = 0 And VchType = "R" Then LoadOrderList
    End If
End Sub
Private Sub LoadOrderList()
    If rstOrderList.State = adStateOpen Then rstOrderList.Close
'    ISNULL((SELECT SUM(QuantitySheets) FROM PaperIOChild WHERE Ref+Paper=P.Code+C.Paper OR Code+Paper=P.Code+C.Paper),0) As Received
    rstOrderList.Open "SELECT * FROM (SELECT P.Code+C.Paper As UniqCode,LTRIM(P.Name) As VchNo,P.Date As VchDate,LTRIM(I.Name)+' (UOM : '+LTRIM(U.Name)+'='+LTRIM(CONVERT(INT,U.Value1))+')' As Paper,U.Value1 As SPU,QuantitySheets As Ordered,ISNULL((SELECT SUM(QuantitySheets) FROM PaperIOChild WHERE Ref+Paper=P.Code+C.Paper),0) As Received FROM ((PaperPOParent P INNER JOIN PaperPOChild C ON P.Code=C.Code) INNER JOIN PaperMaster I ON C.Paper=I.Code) INNER JOIN GeneralMaster U ON  I.UOM=U.Code WHERE P.Supplier='" & Account2Code & "') As Tbl WHERE Ordered-Received>0 ORDER BY VchDate,VchNo,Paper", cnPaperIRVch, adOpenKeyset, adLockReadOnly
    rstOrderList.ActiveConnection = Nothing
    If rstOrderList.RecordCount = 0 Then DisplayError ("No Pending Order Exists"): fpSpread1.SetFocus: Exit Sub
    With FrmOrderList.fpSpread1
        .ClearRange 1, 1, .MaxCols, .MaxRows, True
        .MaxCols = 13
        .Col = 1: .Row = SpreadHeader: .Text = "Paper Name"
        .Col = 5: .Row = SpreadHeader + 1: .Text = "Delivered" '.ColHidden = True 'Billed
        .Col = 6: .ColHidden = True 'Unbilled
        .Col = 7: '.ColHidden = True 'Challan
        .Col = 8: .ColHidden = True 'Direct Sale
        .Col = 11: .ColHidden = True '
        .Col = 12: .ColHidden = True '
        .Col = 13: .ColHidden = True '
        .ColWidth(1) = 75
        .ColWidth(9) = 7.25
        
    End With
    Load FrmOrderList
    FrmOrderList.Text2 = Text3.Text
    Dim i As Integer
    With rstOrderList
        For i = 1 To .RecordCount
            With FrmOrderList.fpSpread1
                .MaxRows = .MaxRows + 1
                .InsertRows i, 1
            End With
        Next
        i = 0
        Do While Not .EOF
            i = i + 1
            FrmOrderList.fpSpread1.SetText 1, i, .Fields("Paper").Value 'Paper Name
            FrmOrderList.fpSpread1.SetText 2, i, .Fields("VchNo").Value 'Order No
            FrmOrderList.fpSpread1.SetText 3, i, Format(.Fields("VchDate").Value, "dd-MMM-yy") 'Date
            FrmOrderList.fpSpread1.SetText 4, i, S2U(Val(.Fields("Ordered").Value), Val(.Fields("SPU").Value)) 'Ordered
            FrmOrderList.fpSpread1.SetText 5, i, S2U(Val(.Fields("Received").Value), Val(.Fields("SPU").Value)) 'Delivered 9
            FrmOrderList.fpSpread1.SetText 7, i, S2U(Val(.Fields("Ordered").Value) - Val(.Fields("Received").Value), Val(.Fields("SPU").Value)) 'Pending 10
            FrmOrderList.fpSpread1.SetText 9, i, 1 'Check Box
            FrmOrderList.fpSpread1.SetText 10, i, .Fields("UniqCode").Value 'UniqueCode
            .MoveNext
        Loop
        FrmOrderList.fpSpread1.SetActiveCell 1, 1
    End With
    With FrmOrderList
        .Check1.Visible = False
        .Check2.Value = 1
        .Text2.Width = 13005
        CenterForm FrmOrderList
        .Show vbModal
    End With
    If Not CheckEmpty(FrmOrderList.VchCodeList, False) Then
        If rstOrderList.State = adStateOpen Then rstOrderList.Close
        rstOrderList.Open "SELECT P.Code As VchCode,LTRIM(P.Name) As VchNo,I.Code As PaperCode,LTRIM(I.Name)+' (UOM : '+LTRIM(U.Name)+'='+LTRIM(CONVERT(INT,U.Value1))+')' As PaperName,LTRIM(U.Name) As UOM,U.Value1 As SPU,I.[Weight/Unit],I.[Units/Bundle],QuantitySheets-ISNULL((SELECT SUM(QuantitySheets) FROM PaperIOChild WHERE Ref+Paper=P.Code+C.Paper),0) As Bal FROM ((PaperPOParent P INNER JOIN PaperPOChild C ON P.Code=C.Code) INNER JOIN PaperMaster I ON C.Paper=I.Code) INNER JOIN GeneralMaster U ON  I.UOM=U.Code WHERE P.Code+C.Paper IN (" & FrmOrderList.VchCodeList & ") ORDER BY Paper,VchNo", cnPaperIRVch, adOpenKeyset, adLockReadOnly
        If rstOrderList.RecordCount > 0 Then
            Dim TotalBdl As Double, Wt As Double, QtyWt As Double
            i = 0
            With fpSpread1
                Do While Not rstOrderList.EOF
                    i = i + 1
                    .SetText 1, i, rstOrderList.Fields("PaperName").Value
                    .SetText 2, i, S2U(Val(rstOrderList.Fields("Bal").Value), Val(rstOrderList.Fields("SPU").Value))
                    .SetText 3, i, rstOrderList.Fields("UOM").Value
                    .SetText 4, i, Val(rstOrderList.Fields("Weight/Unit").Value)
                    QtyWt = Round((Val(rstOrderList.Fields("Bal").Value) / Val(rstOrderList.Fields("SPU").Value)) * Val(rstOrderList.Fields("Weight/Unit").Value), 3)
                    .SetText 5, i, QtyWt
                    .SetText 6, i, Val(rstOrderList.Fields("Bal").Value)
                    .SetText 7, i, Val(rstOrderList.Fields("Units/Bundle").Value)
                    If Val(rstOrderList.Fields("Units/Bundle").Value) > 0 Then TotalBdl = QtyWt / (Val(rstOrderList.Fields("Weight/Unit").Value) * Val(rstOrderList.Fields("Units/Bundle").Value)): TotalBdl = Fix(TotalBdl) + IIf(TotalBdl - Fix(TotalBdl) > 0, 1, 0)
                    .SetText 8, i, TotalBdl
                    .SetText 9, i, rstOrderList.Fields("VchNo").Value
                    .SetText 10, i, rstOrderList.Fields("PaperCode").Value
                    .SetText 11, i, Val(rstOrderList.Fields("SPU").Value)
                    .SetText 12, i, rstOrderList.Fields("VchCode").Value
                    .SetText 13, i, Val(rstOrderList.Fields("Bal").Value)
                    rstOrderList.MoveNext
                Loop
                Call CalculateTotal
            End With
        End If
    End If
    CloseForm FrmOrderList
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Paper As Variant, QtyUnit As Variant, QtySheet As Variant, Wt As Variant, QtyWt As Variant, UPB As Variant, SPU As Variant 'UPB-Unit/Bundle SPU-Sheets/Unit
    Dim TotalBdl As Double, CalcQtyWt As Boolean, CalcQtySheet As Boolean
    With fpSpread1
        .GetText 1, Row, Paper
        If Not CheckEmpty(Paper, False) Then
            .GetText 2, Row, QtyUnit
            .GetText 4, Row, Wt
            .GetText 5, Row, QtyWt
            .GetText 6, Row, QtySheet
            .GetText 7, Row, UPB
            .GetText 11, Row, SPU
            If Col = 2 Then 'Qty-Unit
                CalcQtySheet = True: CalcQtyWt = True
            ElseIf Col = 4 Then 'Wt/Unit
                CalcQtyWt = True
            ElseIf Col = 5 Then 'Qty-Wt
                If Val(Wt) > 0 Then QtyUnit = S2U(Round((QtyWt / Wt) * SPU, 0), SPU): CalcQtySheet = True
            ElseIf Col = 6 Then 'Qty-Sheets
                If Val(SPU) > 0 Then QtyUnit = S2U(QtySheet, SPU): CalcQtyWt = True
            End If
            If CalcQtySheet Then QtySheet = U2S(QtyUnit, SPU)
            If CalcQtyWt Then If Val(SPU) > 0 Then QtyWt = Round(QtySheet * (Wt / SPU), 3)
            If UPB > 0 Then TotalBdl = QtyWt / (Wt * UPB): TotalBdl = Fix(TotalBdl) + IIf(TotalBdl - Fix(TotalBdl) > 0, 1, 0)
            .SetText 2, Row, QtyUnit: .SetText 5, Row, QtyWt: .SetText 6, Row, QtySheet:
            If Option1.Value Then .SetText 8, Row, TotalBdl
            CalculateTotal
        Else
                .SetText 2, Row, "": .SetText 4, Row, "": .SetText 5, Row, "": .SetText 6, Row, "": .SetText 7, Row, ""
        End If
    End With
    Exit Sub
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Public Sub PrintPaperIRVch(ByVal OrderCode As String, Optional ByVal OutputType As String)
    On Error Resume Next
    Dim SQL As String
    Screen.MousePointer = vbHourglass
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
'    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,EMail,Website,GSTIN,Prefix,Suffix FROM CompanyMaster P INNER JOIN CompChild C ON P.Code=C.Code WHERE VchType=" & IIf(VchType = "R", 6, 5), cnPaperIRVch, adOpenKeyset, adLockReadOnly
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,EMail,Website,GSTIN,Prefix,Suffix,Alias FROM CompanyMaster P INNER JOIN CompChild C ON P.Code=C.Code WHERE VchType=" & IIf(VchType = "R", 6, 5), cnPaperIRVch, adOpenKeyset, adLockReadOnly
    rptPaperIssueReceiptOrder.Text1.SetText "Paper " & IIf(VchType = "T", "Transfer", IIf(VchType = "I", "Issue", "Receipt")) & " Voucher"
    rptPaperIssueReceiptOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPaperIssueReceiptOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False)) Then
        rptPaperIssueReceiptOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value) & Space(1) & "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptPaperIssueReceiptOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False) Then
        rptPaperIssueReceiptOrder.Text24.SetText "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    End If
    
    If rstPaperIRVChild.State = adStateOpen Then rstPaperIRVChild.Close
    SQL = "SELECT '" & LTrim(rstCompanyMaster.Fields("Alias").Value) & "'+'/'+'" & VchType & "'+'/' +LTRIM(P.Name)+'/' +'" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "' As VchNo,[Date] As VchDate,LTRIM(M1.PrintName) As Account1Name,LTRIM(M2.PrintName) As Account2Name,Remarks,LTRIM(M3.PrintName)+IIF(M3.Form='S',' (UOM : '+LTRIM(G.PrintName)+'='+LTRIM(CONVERT(INT,G.Value1))+')','') As PaperName,Quantity,C.[Weight/Unit],QuantityKg,C.[Units/Bundle],TotalBundles,LTRIM(G.PrintName) As Unit,QuantitySheets "
    If VchType = "R" Then
        SQL = SQL + ",M2.Address1 As FromAddress1,M2.Address2 As FromAddress2,M2.Address3 As FromAddress3,M2.Address4 As FromAddress4,M1.Address1 As ToAddress1,M1.Address2 As ToAddress2,M1.Address3 As ToAddress3,M1.Address4 As ToAddress4,M3.Form FROM ((((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=C.Account) INNER JOIN AccountMaster M2 ON M2.Code=P.Supplier) INNER JOIN PaperMaster M3 ON M3.Code=C.Paper) INNER JOIN GeneralMaster G ON G.Code=M3.UOM WHERE P.Code='" & OrderCode & "' ORDER BY M3.PrintName"
    Else
        SQL = SQL + ",M1.Address1 As FromAddress1,M1.Address2 As FromAddress2,M1.Address3 As FromAddress3,M1.Address4 As FromAddress4,M2.Address1 As ToAddress1,M2.Address2 As ToAddress2,M2.Address3 As ToAddress3,M2.Address4 As ToAddress4,M3.Form FROM ((((PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.AccountFrom) INNER JOIN AccountMaster M2 ON M2.Code=P.AccountTo) INNER JOIN PaperMaster M3 ON M3.Code=C.Paper) INNER JOIN GeneralMaster G ON G.Code=M3.UOM WHERE P.Code='" & OrderCode & "' ORDER BY M3.PrintName"
    End If
    rstPaperIRVChild.Open SQL, cnPaperIRVch, adOpenKeyset, adLockOptimistic
    rptPaperIssueReceiptOrder.Database.SetDataSource rstPaperIRVChild, 3, 1
    Screen.MousePointer = vbNormal
    
    If VchType = "R" Then
        rptPaperIssueReceiptOrder.Text27.SetText Trim(rstPaperIRVChild.Fields("Account1Name").Value)
        rptPaperIssueReceiptOrder.Text9.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        
        rptPaperIssueReceiptOrder.Section14.Suppress = True
        rptPaperIssueReceiptOrder.Section20.Suppress = True
        rptPaperIssueReceiptOrder.Section21.Suppress = True
        rptPaperIssueReceiptOrder.Section22.Suppress = True
        rptPaperIssueReceiptOrder.Section23.Suppress = True
        
        rptPaperIssueReceiptOrder.Section15.Suppress = True
        rptPaperIssueReceiptOrder.Section24.Suppress = True
        rptPaperIssueReceiptOrder.Section25.Suppress = True
        rptPaperIssueReceiptOrder.Section26.Suppress = True
        rptPaperIssueReceiptOrder.Section27.Suppress = True
    Else
        rptPaperIssueReceiptOrder.Text27.SetText Trim(rstPaperIRVChild.Fields("Account2Name").Value)
        rptPaperIssueReceiptOrder.Text9.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        
        rptPaperIssueReceiptOrder.Section17.Suppress = True
        rptPaperIssueReceiptOrder.Section28.Suppress = True
        rptPaperIssueReceiptOrder.Section29.Suppress = True
        rptPaperIssueReceiptOrder.Section30.Suppress = True
        rptPaperIssueReceiptOrder.Section31.Suppress = True
        
        rptPaperIssueReceiptOrder.Section12.Suppress = True
        rptPaperIssueReceiptOrder.Section32.Suppress = True
        rptPaperIssueReceiptOrder.Section33.Suppress = True
        rptPaperIssueReceiptOrder.Section34.Suppress = True
        rptPaperIssueReceiptOrder.Section35.Suppress = True
    
    End If
    With rptPaperIssueReceiptOrder
    If Logo = "S" Then
    .Picture1.Width = LogoW
    .Picture1.Height = LogoH
    End If
'    .Text2.Width = Header '9000 '7800
'    .Text2.Left = HeaderL '1000 '1680
    If LogoLine = "N" Then
    .Picture1.LeftLineStyle = crLSNoLine
    .Picture1.RightLineStyle = crLSNoLine
    .Picture1.TopLineStyle = crLSNoLine
    .Picture1.BottomLineStyle = crLSNoLine
    End If
    End With
    
    If OutputType = "S" Then
        Set FrmReportViewer.Report = rptPaperIssueReceiptOrder
        FrmReportViewer.Show vbModal
    Else
        rptPaperIssueReceiptOrder.PaperSource = crPRBinAuto
        rptPaperIssueReceiptOrder.PrintOut
    End If
    Set rptPaperIssueReceiptOrder = Nothing
    On Error GoTo 0
End Sub
Private Sub LoadMasterList(Optional ByVal LoadSelected As Boolean)
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "SELECT LTRIM(Name) As Col0,Code FROM AccountMaster ORDER BY Name", cnPaperIRVch, adOpenKeyset, adLockReadOnly
    If rstPaperList.State = adStateOpen Then rstPaperList.Close
    If LoadSelected Then
        rstPaperList.Open "SELECT * FROM (SELECT LTRIM(P.Name)+IIF(P.Form='S',' (UOM : '+LTRIM(C.Name)+'='+LTRIM(CONVERT(INT,C.Value1))+')','') As Col0,FORMAT(dbo.ufnGetPaperStock('" & Account1Code & "',P.Code,'PMV','" & CheckNull(rstPaperIRVParent.Fields("Code").Value) & "','" & GetDate(MhDateInput1.Text) & "'),'#0.000') As Col1,[Weight/Unit],[Units/Bundle],C.Value1 As SPU,C.Name As UOMName,P.Code FROM PaperMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code) As Tbl WHERE CONVERT(DECIMAL(12,3),Col1)<>0 ORDER BY Col0", cnPaperIRVch, adOpenKeyset, adLockReadOnly
    Else
        rstPaperList.Open "SELECT M2.Name As UOMName,[Weight/Unit],[Units/Bundle],M2.Value1 As SPU,M1.Code FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code ORDER BY M1.Name", cnPaperIRVch, adOpenKeyset, adLockReadOnly
    End If
    rstAccountList.ActiveConnection = Nothing
    rstPaperList.ActiveConnection = Nothing
End Sub
