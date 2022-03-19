VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmOutsourceItemPurchaseOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Item (BOM) Purchase Voucher"
   ClientHeight    =   8400
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
   MDIChild        =   -1  'True
   ScaleHeight     =   8979.311
   ScaleMode       =   0  'User
   ScaleWidth      =   15675
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   8370
      Left            =   15
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   15630
      _Version        =   65536
      _ExtentX        =   27570
      _ExtentY        =   14764
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
      Picture         =   "OutsourceItemPurchaseOrder.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8145
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Width           =   15405
         _ExtentX        =   27173
         _ExtentY        =   14367
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
         TabPicture(0)   =   "OutsourceItemPurchaseOrder.frx":001C
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
         TabPicture(1)   =   "OutsourceItemPurchaseOrder.frx":0038
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
            MaxLength       =   40
            TabIndex        =   17
            Top             =   7710
            Width           =   9780
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7185
            Left            =   120
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   480
            Width           =   15180
            _ExtentX        =   26776
            _ExtentY        =   12674
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
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "    Voucher No."
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
               Caption         =   "   Voucher Date"
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
               DataField       =   "SupplierName"
               Caption         =   "   Supplier Name"
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
            BeginProperty Column03 
               DataField       =   "BillAmount"
               Caption         =   "     Voucher Amount"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
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
                  ColumnWidth     =   1425.26
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1470.047
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   9689.953
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   2009.764
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7485
            Left            =   -74880
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   480
            Width           =   15180
            _Version        =   65536
            _ExtentX        =   26776
            _ExtentY        =   13203
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
            Picture         =   "OutsourceItemPurchaseOrder.frx":0054
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   35
               Top             =   3370
               Width           =   14955
               _Version        =   65536
               _ExtentX        =   26379
               _ExtentY        =   503
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
               Picture         =   "OutsourceItemPurchaseOrder.frx":0070
               Picture         =   "OutsourceItemPurchaseOrder.frx":008C
               Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
                  Height          =   285
                  Left            =   13100
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1590
                  _Version        =   65536
                  _ExtentX        =   2805
                  _ExtentY        =   503
                  Calculator      =   "OutsourceItemPurchaseOrder.frx":00A8
                  Caption         =   "OutsourceItemPurchaseOrder.frx":00C8
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "OutsourceItemPurchaseOrder.frx":0134
                  Keys            =   "OutsourceItemPurchaseOrder.frx":0152
                  Spin            =   "OutsourceItemPurchaseOrder.frx":019C
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "######0.00"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "######0.00"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   9999999.99
                  MinValue        =   0
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  NegativeColor   =   255
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  ReadOnly        =   1
                  Separator       =   ""
                  ShowContextMenu =   1
                  ValueVT         =   1638405
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
                  Height          =   285
                  Left            =   9950
                  TabIndex        =   38
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1185
                  _Version        =   65536
                  _ExtentX        =   2090
                  _ExtentY        =   503
                  Calculator      =   "OutsourceItemPurchaseOrder.frx":01C4
                  Caption         =   "OutsourceItemPurchaseOrder.frx":01E4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "OutsourceItemPurchaseOrder.frx":0250
                  Keys            =   "OutsourceItemPurchaseOrder.frx":026E
                  Spin            =   "OutsourceItemPurchaseOrder.frx":02B8
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "######0.000"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "######0.000"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   9999999.999
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
            End
            Begin VB.CommandButton cmdUpload 
               Height          =   375
               Left            =   14310
               Picture         =   "OutsourceItemPurchaseOrder.frx":02E0
               Style           =   1  'Graphical
               TabIndex        =   42
               TabStop         =   0   'False
               ToolTipText     =   "Upload Bill"
               Top             =   105
               Width           =   375
            End
            Begin VB.CommandButton cmdView 
               Height          =   375
               Left            =   14700
               Picture         =   "OutsourceItemPurchaseOrder.frx":0622
               Style           =   1  'Graphical
               TabIndex        =   41
               TabStop         =   0   'False
               ToolTipText     =   "View Bill"
               Top             =   105
               Width           =   375
            End
            Begin VB.TextBox TxtAdNar 
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
               Left            =   7130
               MaxLength       =   40
               TabIndex        =   9
               Top             =   6525
               Width           =   4695
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
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   10
               Top             =   7050
               Width           =   1575
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
               Left            =   1680
               MaxLength       =   10
               TabIndex        =   0
               Top             =   105
               Width           =   1575
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
               TabIndex        =   4
               Top             =   950
               Width           =   13395
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
               TabIndex        =   3
               Top             =   630
               Width           =   13395
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   20
               Top             =   7050
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
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
               Caption         =   " Bill No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0B54
               Picture         =   "OutsourceItemPurchaseOrder.frx":0B70
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   21
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
               Caption         =   " Order No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0B8C
               Picture         =   "OutsourceItemPurchaseOrder.frx":0BA8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   5565
               TabIndex        =   22
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
               Caption         =   " Order Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0BC4
               Picture         =   "OutsourceItemPurchaseOrder.frx":0BE0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   120
               TabIndex        =   23
               Top             =   6210
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
               Caption         =   " GST (%)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0BFC
               Picture         =   "OutsourceItemPurchaseOrder.frx":0C18
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   24
               Top             =   630
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
               Caption         =   " Supplier Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0C34
               Picture         =   "OutsourceItemPurchaseOrder.frx":0C50
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   10260
               TabIndex        =   25
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
               Caption         =   " Delivery Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0C6C
               Picture         =   "OutsourceItemPurchaseOrder.frx":0C88
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   26
               Top             =   945
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0CA4
               Picture         =   "OutsourceItemPurchaseOrder.frx":0CC0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   11940
               TabIndex        =   27
               Top             =   6525
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
               Caption         =   " Net Amount"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0CDC
               Picture         =   "OutsourceItemPurchaseOrder.frx":0CF8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   120
               TabIndex        =   28
               Top             =   6525
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
               Caption         =   " Adjustment"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0D14
               Picture         =   "OutsourceItemPurchaseOrder.frx":0D30
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   11940
               TabIndex        =   29
               Top             =   6210
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
               Caption         =   " GST"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0D4C
               Picture         =   "OutsourceItemPurchaseOrder.frx":0D68
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
               Height          =   330
               Left            =   11940
               TabIndex        =   30
               Top             =   7050
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
               Caption         =   " Paid Amount"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0D84
               Picture         =   "OutsourceItemPurchaseOrder.frx":0DA0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
               Height          =   330
               Left            =   5565
               TabIndex        =   31
               Top             =   7050
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
               Caption         =   " Bill Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0DBC
               Picture         =   "OutsourceItemPurchaseOrder.frx":0DD8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
               Height          =   330
               Left            =   5565
               TabIndex        =   32
               Top             =   6525
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
               Caption         =   " Adj.Remarks"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "OutsourceItemPurchaseOrder.frx":0DF4
               Picture         =   "OutsourceItemPurchaseOrder.frx":0E10
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput3 
               Height          =   330
               Left            =   11820
               TabIndex        =   2
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "OutsourceItemPurchaseOrder.frx":0E2C
               Caption         =   "OutsourceItemPurchaseOrder.frx":0F44
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "OutsourceItemPurchaseOrder.frx":0FB0
               Keys            =   "OutsourceItemPurchaseOrder.frx":0FCE
               Spin            =   "OutsourceItemPurchaseOrder.frx":102C
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
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   7130
               TabIndex        =   1
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "OutsourceItemPurchaseOrder.frx":1054
               Caption         =   "OutsourceItemPurchaseOrder.frx":116C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "OutsourceItemPurchaseOrder.frx":11D8
               Keys            =   "OutsourceItemPurchaseOrder.frx":11F6
               Spin            =   "OutsourceItemPurchaseOrder.frx":1254
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
               Height          =   1920
               Left            =   120
               TabIndex        =   5
               Top             =   1485
               Width           =   14955
               _Version        =   524288
               _ExtentX        =   26379
               _ExtentY        =   3387
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
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "OutsourceItemPurchaseOrder.frx":127C
            End
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   1905
               Left            =   120
               TabIndex        =   6
               Top             =   3870
               Width           =   14955
               _Version        =   524288
               _ExtentX        =   26379
               _ExtentY        =   3360
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
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "OutsourceItemPurchaseOrder.frx":1AA9
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput2 
               Height          =   330
               Left            =   7130
               TabIndex        =   11
               Top             =   7050
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "OutsourceItemPurchaseOrder.frx":2280
               Caption         =   "OutsourceItemPurchaseOrder.frx":2398
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "OutsourceItemPurchaseOrder.frx":2404
               Keys            =   "OutsourceItemPurchaseOrder.frx":2422
               Spin            =   "OutsourceItemPurchaseOrder.frx":2480
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
               Height          =   330
               Left            =   13500
               TabIndex        =   12
               Top             =   7050
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "OutsourceItemPurchaseOrder.frx":24A8
               Caption         =   "OutsourceItemPurchaseOrder.frx":24C8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "OutsourceItemPurchaseOrder.frx":2534
               Keys            =   "OutsourceItemPurchaseOrder.frx":2552
               Spin            =   "OutsourceItemPurchaseOrder.frx":259C
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   330
               Left            =   1680
               TabIndex        =   7
               Top             =   6210
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "OutsourceItemPurchaseOrder.frx":25C4
               Caption         =   "OutsourceItemPurchaseOrder.frx":25E4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "OutsourceItemPurchaseOrder.frx":2650
               Keys            =   "OutsourceItemPurchaseOrder.frx":266E
               Spin            =   "OutsourceItemPurchaseOrder.frx":26B8
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
               ValueVT         =   2004746245
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   330
               Left            =   13500
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   6210
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "OutsourceItemPurchaseOrder.frx":26E0
               Caption         =   "OutsourceItemPurchaseOrder.frx":2700
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "OutsourceItemPurchaseOrder.frx":276C
               Keys            =   "OutsourceItemPurchaseOrder.frx":278A
               Spin            =   "OutsourceItemPurchaseOrder.frx":27D4
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
               Height          =   330
               Left            =   1680
               TabIndex        =   8
               Top             =   6525
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "OutsourceItemPurchaseOrder.frx":27FC
               Caption         =   "OutsourceItemPurchaseOrder.frx":281C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "OutsourceItemPurchaseOrder.frx":2888
               Keys            =   "OutsourceItemPurchaseOrder.frx":28A6
               Spin            =   "OutsourceItemPurchaseOrder.frx":28F0
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
               MinValue        =   -9999999.99
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   2004549637
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
               Height          =   330
               Left            =   13500
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   6525
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "OutsourceItemPurchaseOrder.frx":2918
               Caption         =   "OutsourceItemPurchaseOrder.frx":2938
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "OutsourceItemPurchaseOrder.frx":29A4
               Keys            =   "OutsourceItemPurchaseOrder.frx":29C2
               Spin            =   "OutsourceItemPurchaseOrder.frx":2A0C
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   285
               Left            =   120
               TabIndex        =   36
               Top             =   5760
               Width           =   14955
               _Version        =   65536
               _ExtentX        =   26379
               _ExtentY        =   503
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
               Picture         =   "OutsourceItemPurchaseOrder.frx":2A34
               Picture         =   "OutsourceItemPurchaseOrder.frx":2A50
               Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
                  Height          =   285
                  Left            =   11650
                  TabIndex        =   39
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1470
                  _Version        =   65536
                  _ExtentX        =   2593
                  _ExtentY        =   503
                  Calculator      =   "OutsourceItemPurchaseOrder.frx":2A6C
                  Caption         =   "OutsourceItemPurchaseOrder.frx":2A8C
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "OutsourceItemPurchaseOrder.frx":2AF8
                  Keys            =   "OutsourceItemPurchaseOrder.frx":2B16
                  Spin            =   "OutsourceItemPurchaseOrder.frx":2B60
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "######0.000"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "######0.000"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   9999999.999
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
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   2160
               Width           =   11715
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   15200
               Y1              =   6945
               Y2              =   6945
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
               Y1              =   1370
               Y2              =   1370
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   15200
               Y1              =   3765
               Y2              =   3765
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   15200
               Y1              =   6120
               Y2              =   6120
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   10725
            TabIndex        =   43
            Top             =   7710
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
            Picture         =   "OutsourceItemPurchaseOrder.frx":2B88
            Picture         =   "OutsourceItemPurchaseOrder.frx":2BA4
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
            TabIndex        =   18
            Top             =   7710
            Width           =   855
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   14
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
   Begin MSComDlg.CommonDialog cdUpload 
      Left            =   6600
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmOutsourceItemPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnBOMPurchaseOrder As New ADODB.Connection
Dim rstBOMPOList As New ADODB.Recordset, rstBOMPOParent As New ADODB.Recordset, rstBOMPOChild As New ADODB.Recordset, rstSupplierList As New ADODB.Recordset, rstBOMList As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstLastPurchaseRate As New ADODB.Recordset
Dim SupplierCode As String, AccountCode As String, BOMCode As String
Dim SortOrder, PrevStr
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim oOutlook As New Outlook.Application
Dim EditMode As Boolean
Dim EMailID, Attachment, Message
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    cnBOMPurchaseOrder.CursorLocation = adUseClient
    cnBOMPurchaseOrder.Open cnDatabase.ConnectionString
    rstBOMList.Open "SELECT P.Name As Col0,C.Name As UOMName,P.Code FROM OutsourceItemMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code ORDER BY P.Name", cnBOMPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstSupplierList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnBOMPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstAccountList.Open "SELECT LTRIM(Name) As Col0,Code FROM AccountMaster ORDER BY Name", cnBOMPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstBOMPOList.Open "SELECT T.Code,T.Name,Date,M.Name As SupplierName,BillAmount FROM OutsourceItemPOParent T INNER JOIN AccountMaster M ON T.Supplier=M.Code WHERE FYCode='" & FYCode & "' ORDER BY T.Name", cnBOMPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstBOMPOParent.CursorLocation = adUseClient
    rstBOMPOList.Filter = adFilterNone
    If rstBOMPOList.RecordCount > 0 Then rstBOMPOList.MoveLast
    Set DataGrid1.DataSource = rstBOMPOList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstBOMPOList.EOF Or rstBOMPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
        End If
    rstBOMPOList.ActiveConnection = Nothing
    rstBOMList.ActiveConnection = Nothing
    rstSupplierList.ActiveConnection = Nothing
    rstAccountList.ActiveConnection = Nothing
    LoadMasterList
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    MdiMainMenu.mnuPurchaseOrderSupplyInwardBOMItem.Enabled = False
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
           If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstBOMPOList)
    Call CloseRecordset(rstBOMPOParent)
    Call CloseRecordset(rstBOMPOChild)
    Call CloseRecordset(rstBOMList)
    Call CloseRecordset(rstSupplierList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstLastPurchaseRate)
    Call CloseConnection(cnBOMPurchaseOrder)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuPurchaseOrderSupplyInwardBOMItem.Enabled = True
End Sub
Private Sub Text1_Change()
On Error Resume Next
With rstBOMPOList
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
    If rstBOMPOList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstBOMPOList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstBOMPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstBOMPOList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstBOMPOList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstBOMPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstBOMPOList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstBOMPOList
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
            If Not (rstBOMPOList.EOF Or rstBOMPOList.BOF) Then
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
    Dim CellVal01 As Variant, CellVal02 As Variant, CellVal03 As Variant, i As Integer
    If Button.Index = 1 Then
        If rstBOMPOParent.State = adStateOpen Then rstBOMPOParent.Close
        rstBOMPOParent.Open "SELECT * FROM OutsourceItemPOParent WHERE Code=''", cnBOMPurchaseOrder, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstBOMPOParent) Then
            Text2.Text = GenerateCode(cnBOMPurchaseOrder, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,Name))", "VAL(Name))") & "  FROM OutsourceItemPOParent WHERE FYCode='" & FYCode & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text3.SetFocus
            blnRecordExist = False
            cnBOMPurchaseOrder.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstBOMPOList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstBOMPOList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnBOMPurchaseOrder.Execute "DELETE FROM OutsourceItemPOParent WHERE Code='" & rstBOMPOList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstBOMPOList.Delete
                rstBOMPOList.MoveNext
                If rstBOMPOList.RecordCount > 0 And rstBOMPOList.EOF Then rstBOMPOList.MoveLast
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
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstBOMPOParent) Then
            If UpdateBOMList("D") Then
                UpdateFlag = 1
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 5, i
                        .GetText 5, i, CellVal01    'Amount
                        .GetText 6, i, CellVal02    'BOMCODE
                        If Val(CellVal01) <> 0 And CellVal02 <> "" Then
                            If Not UpdateBOMList("I1") Then UpdateFlag = 0: Exit For
                        End If
                    Next
                End With
                If UpdateFlag = 1 Then
                    If Not UpdateBOMList("I0") Then UpdateFlag = 0
                    With fpSpread2
                        If UpdateFlag = 1 Then
                            For i = 1 To .DataRowCnt
                                .SetActiveCell 3, i
                                .GetText 3, i, CellVal01    'Quantity
                                .GetText 5, i, CellVal02    'Account
                                .GetText 6, i, CellVal03    'Bill of Material
                                If Val(CellVal01) <> 0 And CellVal02 <> "" And CellVal03 <> "" Then
                                    If Not UpdateBOMList("I2") Then UpdateFlag = 0: Exit For
                                End If
                            Next
                        End If
                    End With
                End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnBOMPurchaseOrder.CommitTrans
            If rstBOMPOParent.State = adStateOpen Then rstBOMPOParent.Close
            rstBOMPOParent.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstBOMPOParent) Then
            cnBOMPurchaseOrder.RollbackTrans
            If rstBOMPOParent.State = adStateOpen Then rstBOMPOParent.Close
            rstBOMPOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstBOMPOList.Filter = adFilterNone
        rstBOMPOList.ActiveConnection = cnBOMPurchaseOrder
        Do While Not RefreshRecord(rstBOMPOList)
        Loop
        Set DataGrid1.DataSource = rstBOMPOList
        rstBOMPOList.ActiveConnection = Nothing
        If rstBOMPOList.RecordCount > 0 Then rstBOMPOList.MoveLast
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Supplier", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstBOMPOList.RecordCount = 0 Then Exit Sub
        Call PrintBOMPurchaseOrder(rstBOMPOList.Fields("Code").Value, "", "P")
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstBOMPOList.RecordCount = 0 Then Exit Sub
        Call PrintBOMPurchaseOrder(rstBOMPOList.Fields("Code").Value, "", "S")
        HiLiteRecord = True
    ElseIf Button.Index = 11 Then
        If rstBOMPOList.RecordCount = 0 Then Exit Sub
        Call PrintBOMPurchaseOrder(rstBOMPOList.Fields("Code").Value, "", "M")
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstBOMPOList.RecordCount > 0 Then rstBOMPOList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstBOMPOList.RecordCount > 0 Then
            rstBOMPOList.MovePrevious
            If rstBOMPOList.BOF Then rstBOMPOList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstBOMPOList.RecordCount > 0 Then
            rstBOMPOList.MoveNext
            If rstBOMPOList.EOF Then rstBOMPOList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstBOMPOList.RecordCount > 0 Then rstBOMPOList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstBOMPOList.EOF Or rstBOMPOList.BOF) Then
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
        rstBOMPOList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstBOMPOList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstBOMPOList.EOF Or rstBOMPOList.BOF) Then
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
    If rstBOMPOList.RecordCount = 0 Then
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
    If rstBOMPOParent.EOF Or rstBOMPOParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnBOMPurchaseOrder, "OutsourceItemPOParent", "Code", "[Name]", Trim(Text2.Text), rstBOMPOParent.Fields("Code").Value, False, FYCode) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf Not blnRecordExist Then
        MhDateInput3.Text = Format(DateAdd("d", 1, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
    End If
End Sub
Private Sub Text3_Change()
    If Text3.Text = " " Then Text3.Text = "?": Sendkeys "{TAB}"
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text3.Text)
    If rstSupplierList.RecordCount = 0 Then DisplayError ("No Record in Supplier Master"): Cancel = True: Exit Sub Else rstSupplierList.MoveFirst
    rstSupplierList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstSupplierList.EOF Then
        SelectionType = "S"
        SupplierCode = ""
        Call LoadSelectionList(rstSupplierList, "List of Suppliers...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, SupplierCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then Text3.Text = "?"
        If RTrim(SupplierCode) <> "" Then Sendkeys "{TAB}"
        Cancel = True
    Else
        SupplierCode = rstSupplierList.Fields("Code").Value
    End If
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If MhDateInput2.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput2.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput3_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput3.Text)) Then Cancel = True
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)   'GST (%)
    Call CalculateTotal("N")    'GST Changed
End Sub
Private Sub MhRealInput14_Validate(Cancel As Boolean)   'Adjustment
    MhRealInput11_Validate (False)
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstBOMPOList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstBOMPOParent.State = adStateOpen Then rstBOMPOParent.Close
    rstBOMPOParent.Open "SELECT * FROM OutsourceItemPOParent WHERE Code='" & FixQuote(rstBOMPOList.Fields("Code").Value) & "'", cnBOMPurchaseOrder, adOpenKeyset, adLockOptimistic
    If rstBOMPOParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text8.Text = ""
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput2.Text = "  -  -    "    'Bill Date
    MhDateInput3.Text = Format(DateAdd("d", 1, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")    'Delivery Date
    MhRealInput18.Value = 0 'Total Quantity (Kg)
    MhRealInput19.Value = 0 'Total Gross Amount
    MhRealInput11.Value = 12 'GST (%)
    MhRealInput12.Value = 0 'GST
    MhRealInput14.Value = 0 'Adjustment
    MhRealInput15.Value = 0 'Net Amount
    MhRealInput16.Value = 0 'Paid Amount
    MhRealInput20.Value = 0 'Total Quantity (Ream) - To be issued
    TxtAdNar.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    fpSpread2.ClearRange 1, 1, fpSpread2.MaxCols, fpSpread2.MaxRows, True: fpSpread2.SetActiveCell 1, 1
End Sub
Private Sub LoadFields()
    If rstBOMPOParent.EOF Or rstBOMPOParent.BOF Then Exit Sub
    Text2.Text = rstBOMPOParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstBOMPOParent.Fields("Date").Value, "dd-MM-yyyy")
    MhDateInput3.Text = Format(rstBOMPOParent.Fields("DeliveryDate").Value, "dd-MM-yyyy")
    SupplierCode = rstBOMPOParent.Fields("Supplier").Value
    If rstSupplierList.RecordCount > 0 Then rstSupplierList.MoveFirst
    rstSupplierList.Find "[Code] = '" & SupplierCode & "'"
    If Not rstSupplierList.EOF Then Text3.Text = rstSupplierList.Fields("Col0").Value
    Text4.Text = rstBOMPOParent.Fields("Remarks").Value
    MhRealInput11.Value = Val(rstBOMPOParent.Fields("VAT%").Value)
    MhRealInput12.Value = Val(rstBOMPOParent.Fields("VAT").Value)
    MhRealInput14.Value = Val(rstBOMPOParent.Fields("Adjustment").Value)
    MhRealInput15.Value = Val(rstBOMPOParent.Fields("BillAmount").Value)
    Text8.Text = rstBOMPOParent.Fields("BillNo").Value
    If Not IsNull(rstBOMPOParent.Fields("BillDate").Value) Then MhDateInput2.Text = Format(rstBOMPOParent.Fields("BillDate").Value, "dd-MM-yyyy")
    MhRealInput16.Value = Val(rstBOMPOParent.Fields("PaidAmount").Value)
    TxtAdNar.Text = rstBOMPOParent.Fields("AdjustmentRemarks").Value
    Call LoadBOMList(rstBOMPOParent.Fields("Code").Value)
    CalculateTotal ("G")
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstBOMPOParent.RecordCount = 0 Then Exit Sub
    If rstBOMPOParent.State = adStateOpen Then rstBOMPOParent.Close
    rstBOMPOParent.CursorLocation = adUseServer
    rstBOMPOParent.Open "SELECT * FROM OutsourceItemPOParent WHERE Code='" & FixQuote(rstBOMPOList.Fields("Code").Value) & "'", cnBOMPurchaseOrder, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstBOMPOParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text3.SetFocus
    blnRecordExist = True
    If AllowTransactionsModification = 0 Then
        If Not CheckEmpty(Text8.Text, False) Then LockFields (True)
        Text1.Locked = False
    End If
    cnBOMPurchaseOrder.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstBOMPOParent.EOF Or rstBOMPOParent.BOF Then Exit Sub
    Dim lpBuff As String * 1024
    GetComputerName lpBuff, Len(lpBuff)
    If Not blnRecordExist Then
        rstBOMPOParent.Fields("Code").Value = GenerateCode(cnBOMPurchaseOrder, "SELECT MAX(Code) FROM OutsourceItemPOParent", 6, "0")
        rstBOMPOParent.Fields("IssueOrder").Value = GenerateCode(cnBOMPurchaseOrder, "SELECT MAX(Code) FROM MaterialIOParent", 6, "0")
        rstBOMPOParent.Fields("CreatedBy").Value = UserCode
        rstBOMPOParent.Fields("CreatedOn").Value = Now()
        rstBOMPOParent.Fields("Recordstatus").Value = "N"
    Else
        rstBOMPOParent.Fields("ModifiedBy").Value = UserCode
        rstBOMPOParent.Fields("ModifiedOn").Value = Now()
        rstBOMPOParent.Fields("Recordstatus").Value = "M"
    End If
    rstBOMPOParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstBOMPOParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstBOMPOParent.Fields("Supplier").Value = SupplierCode
    rstBOMPOParent.Fields("DeliveryDate").Value = GetDate(MhDateInput3.Text)
    rstBOMPOParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstBOMPOParent.Fields("VAT%").Value = MhRealInput11.Value
    rstBOMPOParent.Fields("VAT").Value = MhRealInput12.Value
    rstBOMPOParent.Fields("Adjustment").Value = MhRealInput14.Value
    rstBOMPOParent.Fields("BillAmount").Value = MhRealInput15.Value
    rstBOMPOParent.Fields("BillNo").Value = (Text8.Text)
    If Not IsDate(MhDateInput2.Text) Then rstBOMPOParent.Fields("BillDate").Value = Null Else rstBOMPOParent.Fields("BillDate").Value = GetDate(MhDateInput2.Text)
    rstBOMPOParent.Fields("PaidAmount").Value = MhRealInput16.Value
    rstBOMPOParent.Fields("AdjustmentRemarks").Value = IIf(MhRealInput14.Value <> 0, TxtAdNar.Text, "")
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBOMPOParent.Fields("BillFeedDate").Value) Then rstBOMPOParent.Fields("BillFeedDate").Value = Now()
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBOMPOParent.Fields("ComputerName").Value) Then rstBOMPOParent.Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    rstBOMPOParent.Fields("FYCode").Value = FYCode
    rstBOMPOParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstBOMPOList.MoveFirst
    rstBOMPOList.Find "[Code] = '" & rstBOMPOParent.Fields("Code").Value & "'"
    If rstBOMPOList.EOF Then rstBOMPOList.AddNew
    rstBOMPOList.Fields("Code").Value = rstBOMPOParent.Fields("Code").Value
    rstBOMPOList.Fields("Name").Value = Pad(rstBOMPOParent.Fields("Name").Value, Space(1), 10, "L")
    rstBOMPOList.Fields("Date").Value = rstBOMPOParent.Fields("Date").Value
    rstSupplierList.MoveFirst
    rstSupplierList.Find "[Code] = '" & rstBOMPOParent.Fields("Supplier").Value & "'"
    rstBOMPOList.Fields("SupplierName").Value = Trim(rstSupplierList.Fields("Col0").Value)
    rstBOMPOList.Fields("BillAmount").Value = rstBOMPOParent.Fields("BillAmount").Value
    rstBOMPOList.Update
    rstBOMPOList.Sort = SortOrder & " Asc"
    rstBOMPOList.Find "[Code] = '" & rstBOMPOParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        DisplayError ("Order No. cannot be blank")
        Text2.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf Not CheckExists(Text3, "Col0", rstSupplierList, SupplierCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnBOMPurchaseOrder, "OutsourceItemPOParent", "Code", "[Name]", Trim(Text2.Text), rstBOMPOParent.Fields("Code").Value, False, FYCode) Then
        Text2.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf Not ChkBOM() Then
        fpSpread2.SetFocus
        CheckMandatoryFields = True: Exit Function
    End If
    'If MhRealInput14.Value <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput16.Value <> 0 Then If MhRealInput16.Value <> MhRealInput15.Value Then MhRealInput14.SetFocus: CheckMandatoryFields = True: Exit Function
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Private Sub LoadBOMList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstBOMPOChild.State = adStateOpen Then rstBOMPOChild.Close
    rstBOMPOChild.Open "SELECT OutsourceItem As BOMCode,M.Name As BOMName,Quantity,Rate,Amount,U.Name As UOMName FROM (OutsourceItemPOChild T INNER JOIN OutsourceItemMaster M ON T.OutsourceItem=M.Code) INNER JOIN GeneralMaster U ON M.UOM=U.Code WHERE T.Code='" & strOrderCode & "' ORDER BY M.Name", cnBOMPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstBOMPOChild.ActiveConnection = Nothing
    If rstBOMPOChild.RecordCount > 0 Then rstBOMPOChild.MoveFirst
    i = 0
    Do While Not rstBOMPOChild.EOF
        i = i + 1
        With fpSpread1
        LoadMasterList
            .SetText 1, i, rstBOMPOChild.Fields("BOMName").Value
            .SetText 2, i, Val(rstBOMPOChild.Fields("Quantity").Value)
            .SetText 3, i, rstBOMPOChild.Fields("UOMName").Value
            .SetText 4, i, Val(rstBOMPOChild.Fields("Rate").Value)
            .SetText 5, i, Val(rstBOMPOChild.Fields("Amount").Value)
            .SetText 6, i, rstBOMPOChild.Fields("BOMCode").Value
        End With
        rstBOMPOChild.MoveNext
    Loop
    If rstBOMPOChild.State = adStateOpen Then rstBOMPOChild.Close
    rstBOMPOChild.Open "SELECT Item As BOMCode,M1.Name As BOMName,Godown As AccountCode,M2.Name As AccountName,Quantity,U.Name As UOMName FROM ((MaterialIOChild T INNER JOIN OutsourceItemMaster M1 ON T.Item=M1.Code) INNER JOIN AccountMaster M2 ON T.Godown=M2.Code) INNER JOIN GeneralMaster U ON M1.UOM=U.Code WHERE T.Code='" & rstBOMPOParent.Fields("IssueOrder").Value & "' ORDER BY M1.Name,M2.Name", cnBOMPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstBOMPOChild.ActiveConnection = Nothing
    If rstBOMPOChild.RecordCount > 0 Then rstBOMPOChild.MoveFirst
    i = 0
    Do While Not rstBOMPOChild.EOF
        i = i + 1
        With fpSpread2
        
            .SetText 1, i, rstBOMPOChild.Fields("BOMName").Value
            .SetText 2, i, rstBOMPOChild.Fields("AccountName").Value
            .SetText 3, i, Val(rstBOMPOChild.Fields("Quantity").Value)
            .SetText 4, i, rstBOMPOChild.Fields("UOMName").Value
            .SetText 5, i, rstBOMPOChild.Fields("AccountCode").Value
            .SetText 6, i, rstBOMPOChild.Fields("BOMCode").Value
        End With
        rstBOMPOChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Bill of Material List")
End Sub
Private Sub CalculateTotal(ByVal strType As String)
    Dim Qty As Variant, Amt As Variant, i As Integer
    If strType = "G" Then   'Calculate GST
        MhRealInput18.Value = 0: MhRealInput19.Value = 0: MhRealInput20.Value = 0
        Qty = 0
        With fpSpread1
            For i = 1 To .DataRowCnt
                .GetText 2, i, Qty: .GetText 5, i, Amt
                MhRealInput18.Value = MhRealInput18.Value + Qty
                MhRealInput19.Value = MhRealInput19.Value + Amt
            Next
        End With
        With fpSpread2
            For i = 1 To .DataRowCnt
                .GetText 3, i, Qty
                MhRealInput20.Value = MhRealInput20.Value + Qty
            Next
        End With
        MhRealInput12.Value = (MhRealInput19.Value + MhRealInput14.Value) * MhRealInput11.Value / 100 'VAT
    Else
        MhRealInput12.Value = (MhRealInput19.Value + MhRealInput14.Value) * MhRealInput11.Value / 100 'VAT
        MhRealInput15.Value = Round(MhRealInput19.Value + MhRealInput12.Value + MhRealInput14.Value, 0)
    End If
End Sub
Private Function UpdateBOMList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 6) As Variant
    On Error GoTo ErrorHandler
    UpdateBOMList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        cnBOMPurchaseOrder.Execute "DELETE FROM OutsourceItemPOChild WHERE Code='" & rstBOMPOParent.Fields("Code").Value & "'"
        cnBOMPurchaseOrder.Execute "DELETE FROM MaterialIOParent WHERE Code='" & rstBOMPOParent.Fields("IssueOrder").Value & "'"
    ElseIf ActionType = "I1" Then
        With fpSpread1
            .GetText 2, .ActiveRow, CellVal(1)  'Quantity
            .GetText 4, .ActiveRow, CellVal(2)  'Rate
            .GetText 5, .ActiveRow, CellVal(3)  'Amount
            .GetText 6, .ActiveRow, CellVal(4)  'Bill Of Material
        End With
        cnBOMPurchaseOrder.Execute "INSERT INTO OutsourceItemPOChild VALUES ('" & rstBOMPOParent.Fields("Code").Value & "','" & CellVal(4) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ")"
    ElseIf ActionType = "I2" Then
        With fpSpread2
            .GetText 3, .ActiveRow, CellVal(1)  'Quantity
            .GetText 5, .ActiveRow, CellVal(2)  'Account
            .GetText 6, .ActiveRow, CellVal(3)  'Bill of Material
        End With
        cnBOMPurchaseOrder.Execute "INSERT INTO MaterialIOChild VALUES ('" & rstBOMPOParent.Fields("IssueOrder").Value & "','1','" & CellVal(3) & "','" & CellVal(2) & "',''," & Val(CellVal(1)) & ")"
    Else
        If DatabaseType = "MS SQL" Then
            cnBOMPurchaseOrder.Execute "INSERT INTO MaterialIOParent VALUES ('" & rstBOMPOParent.Fields("IssueOrder").Value & "','" & Pad(Trim(Text2.Text), Space(1), 10, "L") & "','" & Format(rstBOMPOParent.Fields("Date").Value, "dd-MMM-yyyy") & "','" & SupplierCode & "','1','','" & UserCode & "',GETDATE(),Null,Null,'N','N','" & FYCode & "')"
        Else
            cnBOMPurchaseOrder.Execute "INSERT INTO MaterialIOParent VALUES ('" & rstBOMPOParent.Fields("IssueOrder").Value & "','" & Pad(Trim(Text2.Text), Space(1), 10, "L") & "',#" & Format(rstBOMPOParent.Fields("Date").Value, "mm-dd-yyyy") & "#,'" & SupplierCode & "','1','','" & UserCode & "',NOW(),Null,Null,'N','N','" & FYCode & "')"
        End If
    End If
    Exit Function
ErrorHandler:
    UpdateBOMList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Supplier" Then rstBOMPOList.Filter = "[SupplierName] Like '%" & SrchText & "%'"
End Sub
Public Sub PrintBOMPurchaseOrder(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String)
    Dim rstCompanyMaster As New ADODB.Recordset, rstBOMOrder As New ADODB.Recordset, rstBOMOrderChild As New ADODB.Recordset, Prefix As String
    Dim oOutlookMsg As Outlook.MailItem, FileName As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Prefix = "BM/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/"
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,eMail,GSTIN FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    rstBOMOrder.Open "SELECT '" & Prefix & "'+LTRIM(P.Name) As OrderNo,[Date] As OrderDate,DeliveryDate,LTRIM(M1.PrintName) As SupplierName,VAT,Adjustment,BillAmount,Remarks,LTRIM(M2.PrintName) As BOMName,Quantity,Rate,Amount,[VAT%],LTRIM(G.PrintName)+' ('+LTRIM(G.Value1)+')' As UOM,(SELECT TOP 1 '" & Prefix & "'+LTRIM(P1.Name)+'/'+FORMAT(P1.Date,'dd-MM-yyyy')+'/'+FORMAT([Rate],'0.00') FROM OutsourceItemPOParent P1 INNER JOIN OutsourceItemPOChild C1 ON P1.Code=C1.Code WHERE C1.OutsourceItem=C.OutsourceItem AND P1.Code<P.Code ORDER BY P1.Name DESC) As LastPurchaseRate FROM (((OutsourceItemPOParent P LEFT JOIN OutsourceItemPOChild C ON P.Code = C.Code) LEFT JOIN AccountMaster M1 ON M1.Code=P.Supplier)LEFT JOIN OutsourceItemMaster M2 On C.OutsourceItem=M2.Code)LEFT JOIN GeneralMaster G ON G.Code=M2.UOM WHERE P.Code='" & OrderCode & "' ORDER BY M2.PrintName", cnDatabase, adOpenKeyset, adLockOptimistic
    rstBOMOrderChild.Open "SELECT LTRIM(M1.PrintName) As OutsourceItemName,LTRIM(M3.PrintName) As AccountName,Quantity,M3.Address1 As PrinterAdd1,M3.Address2 As PrinterAdd2,M3.Address3 As PrinterAdd3,M3.Address4 As PrinterAdd4,LTRIM(M3.eMail) As AccountMail,M3.TIN As GSTIN,M3.Mobile,LTRIM(G.PrintName)+' ('+LTRIM(G.Value1)+')' As UOM FROM (((OutsourceItemPOParent P INNER JOIN MaterialIOChild C ON P.IssueOrder=C.Code) INNER JOIN OutsourceItemMaster M1 ON M1.Code=C.Item) INNER JOIN AccountMaster M3 ON C.Godown=M3.Code)LEFT JOIN GeneralMaster G ON G.Code=M1.UOM WHERE P.Code='" & OrderCode & "' ORDER BY M1.PrintName", cnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    rstBOMOrder.ActiveConnection = Nothing: rstBOMOrderChild.ActiveConnection = Nothing
    rptOutsourceItemPurchaseOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptOutsourceItemPurchaseOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptOutsourceItemPurchaseOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value) & Space(1) & "GSTIN : " & Trim(rstCompanyMaster.Fields("GSTIN").Value)
    rptOutsourceItemPurchaseOrder.Text20.SetText "Add : GST @" + Format(rstBOMOrder.Fields("VAT%").Value, "0.00") + "%"
    rptOutsourceItemPurchaseOrder.Text28.SetText " (" & Trim(NumberToWords(rstBOMOrder.Fields("BillAmount").Value, True)) & ")"
    rptOutsourceItemPurchaseOrder.Text27.SetText "for " & Trim(rstBOMOrder.Fields("SupplierName").Value)
    rptOutsourceItemPurchaseOrder.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptOutsourceItemPurchaseOrder.Database.SetDataSource rstBOMOrder, 3, 1
    rptOutsourceItemPurchaseOrder.Subreport1.OpenSubreport.Database.SetDataSource rstBOMOrderChild, 3, 1
    If OutputType = "S" Then
        Set FrmReportViewer.Report = rptOutsourceItemPurchaseOrder
        FrmReportViewer.Show vbModal
    ElseIf OutputType = "P" Then
        rptOutsourceItemPurchaseOrder.PaperSource = crPRBinAuto
        rptOutsourceItemPurchaseOrder.PrintOut False    'Print Report Without Prompt
    Else
        Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
        With oOutlookMsg
            .To = rstBOMOrder.Fields("SupplierMail").Value
            .Subject = "BOM Purchase Order #" & Trim(rstBOMOrder.Fields("OrderNo").Value)
            .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith PO #" & Trim(rstBOMOrder.Fields("OrderNo").Value) & " for doing the needful at your end. An early execution of the order will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of execution of order.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & "Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
            rptOutsourceItemPurchaseOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
            rptOutsourceItemPurchaseOrder.ExportOptions.DestinationType = crEDTDiskFile
            FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
            rptOutsourceItemPurchaseOrder.ExportOptions.DiskFileName = FileName
            rptOutsourceItemPurchaseOrder.Export False
            .Attachments.Add (FileName)
            .Importance = olImportanceHigh
            .ReadReceiptRequested = True
            If CheckEmpty(.To, False) Then .Display Else .Send
        End With
        Set oOutlookMsg = Nothing
    End If
    Set rptOutsourceItemPurchaseOrder = Nothing
    Call CloseRecordset(rstBOMOrder): Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBOMOrderChild)
    On Error GoTo 0
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1: fpSpread1.SetFocus
            CalculateTotal ("G"): CalculateTotal ("N")
        End If
    ElseIf KeyCode = vbKeySpace Then
        Dim BillofMaterial As Variant
        Dim LastPurchaseRate As String
        With fpSpread1
            If .ActiveCol = 1 Then
                .GetText 6, .ActiveRow, BillofMaterial
                On Error Resume Next
                FrmOutsourceItemMaster.SL = True
                FrmOutsourceItemMaster.MasterCode = BillofMaterial
                Load FrmOutsourceItemMaster
                If Err.Number <> 364 Then FrmOutsourceItemMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 5, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then
                    LoadMasterList
                    rstBOMList.MoveFirst: rstBOMList.Find "[Code] ='" & slCode & "'"
                    .SetText 3, .ActiveRow, rstBOMList.Fields("UOMName").Value
                    LastPurchaseRate = GetLastPurchaseRate(slCode)
                    If Not CheckEmpty(GetLastPurchaseRate(slCode), False) Then MsgBox "Last Purchase Rate : " & LastPurchaseRate & " !!!", vbInformation, App.Title
                    .SetFocus
                    Sendkeys "{ENTER}"
                Else
                    .SetActiveCell 1, .ActiveRow
                End If
            End If
        End With
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Qty As Variant, Rate As Variant, BillofMaterial As Variant
    With fpSpread1
        If Col = 1 Or Col = 2 Or Col = 4 Then
            .GetText 1, Row, BillofMaterial
            .GetText 2, Row, Qty
            .GetText 4, Row, Rate
            If BillofMaterial = "" Then .SetText 5, Row, "" Else .SetText 5, Row, Qty * Rate: CalculateTotal ("G"): CalculateTotal ("N")
        End If
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub fpSpread2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread2.DeleteRows fpSpread2.ActiveRow, 1: fpSpread2.SetFocus
            CalculateTotal ("G")
        End If
    ElseIf KeyCode = vbKeySpace Then
        Dim BillofMaterial As Variant, Account As Variant
        With fpSpread2
            .GetText 1, .ActiveRow, BillofMaterial
            If .ActiveCol = 1 Then
                If BillofMaterial = "" Then
                    fpSpread1.GetText 1, fpSpread1.ActiveRow, BillofMaterial
                    .SetText 1, .ActiveRow, BillofMaterial
                    fpSpread1.GetText 6, fpSpread1.ActiveRow, BillofMaterial
                    .SetText 6, .ActiveRow, BillofMaterial
                    If BillofMaterial <> "" Then
                        fpSpread1.GetText 3, fpSpread1.ActiveRow, BillofMaterial
                        .SetText 4, .ActiveRow, BillofMaterial
                        Sendkeys "{ENTER}"
                    End If
                End If
            ElseIf .ActiveCol = 2 Then
                If BillofMaterial <> "" Then
                    .GetText 2, .ActiveRow, Account
                    Text6.Text = FixQuote(Account)
                    If rstAccountList.RecordCount = 0 Then DisplayError ("No Record in Account Master"): .SetActiveCell 1, .ActiveRow: Exit Sub Else rstAccountList.MoveFirst
                    rstAccountList.Find "[Col0] = '" & RTrim(Account) & "'"
                    SelectionType = "S"
                    AccountCode = ""
                    Call LoadSelectionList(rstAccountList, "List of Accounts...", "Name")
                    SearchOrder = 0
                    Call DisplaySelectionList(Text6, AccountCode)
                    Call CloseForm(FrmSelectionList)
                    If AccountCode = "" Then
                        .SetActiveCell 2, .ActiveRow
                    Else
                        .SetText 2, .ActiveRow, Text6.Text
                        .SetText 5, .ActiveRow, AccountCode
                        Sendkeys "{ENTER}"
                    End If
                End If
            End If
        End With
    End If
End Sub
Private Sub fpSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Paper As Variant, Qty As Variant
    With fpSpread2
        If Col = 3 Then CalculateTotal ("G")
    End With
End Sub
Private Sub fpSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Function ChkBOM() As Boolean
    Dim i As Integer, K As Integer, BOM01 As Variant, Qty01 As Variant, BOM02 As Variant, Qty02 As Variant, Price As Variant, Qty As Double
    ChkBOM = True
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.GetText 1, i, BOM01
        fpSpread1.GetText 2, i, Qty01
        fpSpread1.GetText 4, i, Price
        If Val(Price) = 0 Then DisplayError ("Price of Bill of Material at row #" & Trim(Str(i)) & " is zero"): ChkBOM = False: Exit Function
        If fpSpread2.DataRowCnt = 0 Then ChkBOM = True: Exit Function
        Qty = 0
        With fpSpread2
            For K = 1 To .DataRowCnt
                .GetText 1, K, BOM02
                If BOM01 = BOM02 Then .GetText 3, K, Qty02: Qty = Qty + Qty02
            Next
        End With
        If Qty01 <> Qty Then DisplayError ("Purchased vs Issued quantity difference for Bill of Material - " & BOM01): ChkBOM = False: Exit Function
    Next
End Function
Private Sub LockFields(ByVal bVal As Boolean)
    Dim O As Object
    For Each O In Me
        If TypeName(O) = "TextBox" Then
            O.Locked = bVal
        ElseIf TypeName(O) = "TDBNumber" Then
            O.ReadOnly = bVal
        ElseIf TypeName(O) = "fpSpread" Then
            O.Enabled = Not bVal
        End If
    Next
End Sub
Private Sub cmdUpload_Click()
'    cdUpload.DialogTitle = "Select File"
'    cmdUpload.Show
'    CommonDialog1.DialogTitle = "Select File"
'    CommonDialog1.Show
'
'        ofdSpecimen.MultiSelect = False
'        If ofdSpecimen.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
'            FilePath = ofdSpecimen.FileName
'            For i As Integer = 0 To lvwOrderList.Items.Count - 1
'                If lvwOrderList.Items(i).SubItems(0).Text.Trim = System.IO.Path.GetFileName(FilePath) AndAlso lvwOrderList.Items(i).SubItems(1).Text.Trim = cmbSession.Text.Trim AndAlso lvwOrderList.Items(i).SubItems(2).Text.Trim = "S" Then MessageBox.Show("Duplicate Entry !!!", "Specimen", MessageBoxButtons.OK, MessageBoxIcon.Information) : DuplicateEntry = True : Exit For
'            Next
'            If Not DuplicateEntry Then UploadImageOrFile()
'        End If
'        ofdSpecimen.Dispose()
End Sub
Private Sub LoadMasterList()
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "SELECT LTRIM(Name) As Col0,Code FROM AccountMaster ORDER BY Name", cnBOMPurchaseOrder, adOpenKeyset, adLockReadOnly
    If rstBOMList.State = adStateOpen Then rstBOMList.Close
    rstBOMList.Open "SELECT P.Name As Col0,C.Name As UOMName,P.Code FROM OutsourceItemMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code ORDER BY P.Name", cnBOMPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    rstBOMList.ActiveConnection = Nothing
End Sub
Private Function GetLastPurchaseRate(ByVal Item As String) As String
    On Error GoTo ErrorHandler
    If rstLastPurchaseRate.State = adStateOpen Then rstLastPurchaseRate.Close
    rstLastPurchaseRate.Open "SELECT TOP 1 Rate FROM OutsourceItemPOParent P INNER JOIN OutsourceItemPOChild C ON P.Code=C.Code WHERE OutsourceItem='" & Item & "' AND P.Code < '" & IIf(CheckNull(rstBOMPOParent.Fields("Code").Value) = "", "999999", rstBOMPOParent.Fields("Code").Value) & "' ORDER BY P.Name DESC", cnBOMPurchaseOrder, adOpenKeyset, adLockReadOnly
    If rstLastPurchaseRate.RecordCount > 0 Then GetLastPurchaseRate = Trim(rstLastPurchaseRate.Fields("Rate").Value)
    Exit Function
ErrorHandler:
    DisplayError ("Failed to fetch Last Purchase Rate")
End Function
