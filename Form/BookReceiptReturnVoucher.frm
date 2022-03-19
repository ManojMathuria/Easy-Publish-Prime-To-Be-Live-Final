VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmBookReceiptReturnVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Receipt-Return Voucher"
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
      TabIndex        =   11
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
      Picture         =   "BookReceiptReturnVoucher.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8030
         Left            =   120
         TabIndex        =   13
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
         TabPicture(0)   =   "BookReceiptReturnVoucher.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "BookReceiptReturnVoucher.frx":0038
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
            Left            =   605
            MaxLength       =   40
            TabIndex        =   15
            Top             =   7590
            Width           =   12785
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7070
            Left            =   120
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   450
            Width           =   13260
            _ExtentX        =   23389
            _ExtentY        =   12462
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16776960
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
            BeginProperty Column03 
               DataField       =   "PartyName"
               Caption         =   "Party Name"
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
                  Format          =   "dd-MM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
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
                  ColumnWidth     =   2534.74
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   4545.071
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
                  ColumnWidth     =   2429.858
               EndProperty
               BeginProperty Column05 
                  Locked          =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7340
            Left            =   -74880
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   480
            Width           =   13260
            _Version        =   65536
            _ExtentX        =   23389
            _ExtentY        =   12947
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
            Picture         =   "BookReceiptReturnVoucher.frx":0054
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
               Left            =   1560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   3
               Top             =   950
               Width           =   11595
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   24
               Top             =   6110
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
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookReceiptReturnVoucher.frx":0070
               Picture         =   "BookReceiptReturnVoucher.frx":008C
               Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
                  Height          =   285
                  Left            =   9030
                  TabIndex        =   25
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1140
                  _Version        =   65536
                  _ExtentX        =   2011
                  _ExtentY        =   503
                  Calculator      =   "BookReceiptReturnVoucher.frx":00A8
                  Caption         =   "BookReceiptReturnVoucher.frx":00C8
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "BookReceiptReturnVoucher.frx":0134
                  Keys            =   "BookReceiptReturnVoucher.frx":0152
                  Spin            =   "BookReceiptReturnVoucher.frx":019C
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
               Left            =   1320
               MaxLength       =   255
               TabIndex        =   9
               Top             =   6900
               Width           =   1530
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
               Left            =   1560
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
               Left            =   1560
               MaxLength       =   40
               TabIndex        =   4
               Top             =   1265
               Width           =   8865
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
               Left            =   1560
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   2
               Top             =   630
               Width           =   11595
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   18
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Vch No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookReceiptReturnVoucher.frx":01C4
               Picture         =   "BookReceiptReturnVoucher.frx":01E0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   19
               Top             =   630
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
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
               Caption         =   " Party Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookReceiptReturnVoucher.frx":01FC
               Picture         =   "BookReceiptReturnVoucher.frx":0218
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   20
               Top             =   1265
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookReceiptReturnVoucher.frx":0234
               Picture         =   "BookReceiptReturnVoucher.frx":0250
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   10860
               TabIndex        =   21
               Top             =   6900
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Caption         =   " Challan Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookReceiptReturnVoucher.frx":026C
               Picture         =   "BookReceiptReturnVoucher.frx":0288
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Left            =   120
               TabIndex        =   22
               Top             =   6900
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Challan No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookReceiptReturnVoucher.frx":02A4
               Picture         =   "BookReceiptReturnVoucher.frx":02C0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
               Height          =   330
               Left            =   10860
               TabIndex        =   23
               Top             =   6580
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Caption         =   " Cartage Amt"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookReceiptReturnVoucher.frx":02DC
               Picture         =   "BookReceiptReturnVoucher.frx":02F8
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
               Calendar        =   "BookReceiptReturnVoucher.frx":0314
               Caption         =   "BookReceiptReturnVoucher.frx":042C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookReceiptReturnVoucher.frx":0498
               Keys            =   "BookReceiptReturnVoucher.frx":04B6
               Spin            =   "BookReceiptReturnVoucher.frx":0514
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
               Height          =   4335
               Left            =   120
               TabIndex        =   6
               Top             =   1785
               Width           =   13035
               _Version        =   524288
               _ExtentX        =   22992
               _ExtentY        =   7646
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
               MaxCols         =   8
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "BookReceiptReturnVoucher.frx":053C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
               Height          =   330
               Left            =   120
               TabIndex        =   26
               Top             =   6580
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " No. of Box"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookReceiptReturnVoucher.frx":0EDC
               Picture         =   "BookReceiptReturnVoucher.frx":0EF8
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
               Height          =   330
               Left            =   12060
               TabIndex        =   8
               Top             =   6580
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calculator      =   "BookReceiptReturnVoucher.frx":0F14
               Caption         =   "BookReceiptReturnVoucher.frx":0F34
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookReceiptReturnVoucher.frx":0FA0
               Keys            =   "BookReceiptReturnVoucher.frx":0FBE
               Spin            =   "BookReceiptReturnVoucher.frx":1008
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
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   2160
               Width           =   11715
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   10860
               TabIndex        =   28
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Vch Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookReceiptReturnVoucher.frx":1030
               Picture         =   "BookReceiptReturnVoucher.frx":104C
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
               Height          =   330
               Left            =   1320
               TabIndex        =   7
               Top             =   6580
               Width           =   1530
               _Version        =   65536
               _ExtentX        =   2699
               _ExtentY        =   582
               Calculator      =   "BookReceiptReturnVoucher.frx":1068
               Caption         =   "BookReceiptReturnVoucher.frx":1088
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookReceiptReturnVoucher.frx":10F4
               Keys            =   "BookReceiptReturnVoucher.frx":1112
               Spin            =   "BookReceiptReturnVoucher.frx":115C
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
            Begin TDBDate6Ctl.TDBDate MhDateInput2 
               Height          =   330
               Left            =   12060
               TabIndex        =   10
               Top             =   6900
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "BookReceiptReturnVoucher.frx":1184
               Caption         =   "BookReceiptReturnVoucher.frx":129C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookReceiptReturnVoucher.frx":1308
               Keys            =   "BookReceiptReturnVoucher.frx":1326
               Spin            =   "BookReceiptReturnVoucher.frx":1384
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
               TabIndex        =   29
               Top             =   950
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
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
               Caption         =   " Material Centre"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookReceiptReturnVoucher.frx":13AC
               Picture         =   "BookReceiptReturnVoucher.frx":13C8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   10410
               TabIndex        =   30
               Top             =   1265
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
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
               Caption         =   " Bill Type"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookReceiptReturnVoucher.frx":13E4
               Picture         =   "BookReceiptReturnVoucher.frx":1400
            End
            Begin MSForms.ComboBox cmbChallanType 
               Height          =   330
               Left            =   11250
               TabIndex        =   5
               Top             =   1265
               Width           =   1905
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "3360;582"
               MatchEntry      =   0
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Calibri"
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   13240
               Y1              =   6480
               Y2              =   6480
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
               Y1              =   1675
               Y2              =   1675
            End
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
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
            ForeColor       =   &H8000000E&
            Height          =   330
            Left            =   120
            TabIndex        =   16
            Top             =   7590
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   12
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
Attribute VB_Name = "FrmBookReceiptReturnVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AML As String 'A-Add M-Modify L-List
Public VchCode As String  'Vch to Modify
Public VchType As String
Dim cnItemReceiptReturnVoucher As New ADODB.Connection
Dim rstItemRVParent As New ADODB.Recordset, rstItemRVChild As New ADODB.Recordset, rstItemRVList As New ADODB.Recordset
Dim rstCompanyMaster As New ADODB.Recordset, rstPartyList As New ADODB.Recordset, rstMaterialCentreList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstRefList As New ADODB.Recordset, rstOrderList As New ADODB.Recordset
Dim PartyCode As String, ItemCode As String, RefCode As String, MaterialCentreCode As String
Dim SortOrder, PrevStr
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim EditMode As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    Me.Caption = "Item " + IIf(VchType = "R", "Receipt", "Issue") + " Voucher"
    cnItemReceiptReturnVoucher.CursorLocation = adUseClient
    cnItemReceiptReturnVoucher.Open cnDatabase.ConnectionString
    rstItemRVParent.CursorLocation = adUseClient
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,EMail,Website FROM CompanyMaster", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
    rstItemList.Open "SELECT Name As Col0,Price,Code FROM BookMaster ORDER BY Name", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
    rstPartyList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
    rstMaterialCentreList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
    rstItemRVList.Open "SELECT T.Code,T.Name,Date,M2.Name As MaterialCentreName,M1.Name As PartyName,ChallanNo,ChallanDate FROM (BookRVParent T INNER JOIN AccountMaster M1 ON T.Party=M1.Code) INNER JOIN AccountMaster M2 ON T.MaterialCentre=M2.Code WHERE LEFT(T.Type,1)='" & VchType & "'ORDER BY T.Name", cnItemReceiptReturnVoucher, adOpenKeyset, adLockOptimistic
    rstItemRVList.Filter = adFilterNone
    If rstItemRVList.RecordCount > 0 Then
        If CheckEmpty(VchCode, False) Then
            rstItemRVList.MoveLast
        Else
            rstItemRVList.MoveFirst
            rstItemRVList.Find "[Code]='" & VchCode & "'"
        End If
    End If
    Set DataGrid1.DataSource = rstItemRVList
    rstItemRVList.ActiveConnection = Nothing
    rstItemList.ActiveConnection = Nothing
    rstPartyList.ActiveConnection = Nothing
    rstMaterialCentreList.ActiveConnection = Nothing
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstItemRVList.EOF Or rstItemRVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    cmbChallanType.AddItem "FG", 0
    cmbChallanType.AddItem "UFG", 1
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    If MdiMainMenu.MnuBookReceiptVch.Enabled Then AddModifyList
    MdiMainMenu.mnuBookIssueVch.Enabled = False
    MdiMainMenu.MnuBookReceiptVch.Enabled = False
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
    ElseIf Shift = 0 And KeyCode = vbKeyF3 And Toolbar1.Buttons.Item(4).Enabled Then
        If InStr(1, "Text3_Text5", Me.ActiveControl.Name) > 0 Then
            On Error Resume Next
            FrmAccountMaster.AccountType = "01"
            FrmAccountMaster.AL = "A"
            Load FrmAccountMaster
            If Err.Number <> 364 Then FrmAccountMaster.Caption = "Account Master": FrmAccountMaster.Show
        ElseIf Me.ActiveControl.Name = "fpSpread1" Then
            If fpSpread1.ActiveCol = 1 Then
                On Error Resume Next
                FrmBookMaster.BookType = "F"
                FrmBookMaster.AL = "A"
                Load FrmBookMaster
                If Err.Number <> 364 Then FrmBookMaster.Show
            End If
        End If
        KeyCode = 0
    ElseIf Shift = 4 And KeyCode = vbKeyM And Toolbar1.Buttons.Item(4).Enabled Then
        If Me.ActiveControl.Name = "Text3" Then
            If Not CheckEmpty(Text3.Text, False) Then
                FrmAccountMaster.AccountType = "01"
                FrmAccountMaster.AL = "M"
                FrmAccountMaster.MasterCode = PartyCode
                Load FrmAccountMaster
                If Err.Number <> 364 Then FrmAccountMaster.Caption = "Account Master": FrmAccountMaster.Show
            End If
        ElseIf Me.ActiveControl.Name = "Text5" Then
            If Not CheckEmpty(Text5.Text, False) Then
                FrmAccountMaster.AccountType = "01"
                FrmAccountMaster.AL = "M"
                FrmAccountMaster.MasterCode = MaterialCentreCode
                Load FrmAccountMaster
                If Err.Number <> 364 Then FrmAccountMaster.Caption = "Account Master": FrmAccountMaster.Show
            End If
        ElseIf Me.ActiveControl.Name = "fpSpread1" Then
            Dim Item As Variant
            fpSpread1.GetText 1, fpSpread1.ActiveRow, Item
            If Not CheckEmpty(Item, False) Then
                fpSpread1.GetText 7, fpSpread1.ActiveRow, Item
                FrmBookMaster.BookType = "F"
                FrmBookMaster.AL = "M"
                FrmBookMaster.MasterCode = Item
                Load FrmBookMaster
                If Err.Number <> 364 Then FrmBookMaster.Show
            End If
        End If
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        If Toolbar1.Buttons.Item(1).Enabled Then
            SSTab1.Tab = 1: SSTab1.SetFocus
        Else
           If Me.ActiveControl.Name <> "fpSpread1" Then SendKeys "{TAB}"
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
    Call CloseRecordset(rstItemRVList)
    Call CloseRecordset(rstItemRVParent)
    Call CloseRecordset(rstItemRVChild)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstPartyList)
    Call CloseRecordset(rstMaterialCentreList)
    Call CloseRecordset(rstRefList)
    Call CloseRecordset(rstOrderList)
    Call CloseConnection(cnItemReceiptReturnVoucher)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuBookIssueVch.Enabled = True
    MdiMainMenu.MnuBookReceiptVch.Enabled = True
End Sub
Private Sub Text1_Change()
    If rstItemRVList.RecordCount = 0 Then Exit Sub
    rstItemRVList.MoveFirst
    If Text1.Text <> "" Then
        If SortOrder = "Name" Then
           rstItemRVList.Find "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
        Else
           rstItemRVList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        End If
        If rstItemRVList.EOF Then
            rstItemRVList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then rstItemRVList.Bookmark = dblBookMark
            Else
                PrevStr = ""
            End If
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            SendKeys "{End}"
        Else
            PrevStr = Text1.Text
            dblBookMark = DataGrid1.Bookmark
        End If
    Else
        PrevStr = ""
    End If
    If Not (rstItemRVList.EOF Or rstItemRVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    If rstItemRVList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstItemRVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstItemRVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstItemRVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstItemRVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstItemRVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstItemRVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstItemRVList
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
            If Not (rstItemRVList.EOF Or rstItemRVList.BOF) Then
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
        If rstItemRVParent.State = adStateOpen Then rstItemRVParent.Close
        rstItemRVParent.Open "SELECT * FROM BookRVParent WHERE Code=''", cnItemReceiptReturnVoucher, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstItemRVParent) Then
            Text2.Text = GenerateCode(cnItemReceiptReturnVoucher, "SELECT MAX(VAL(Name))  FROM BookRVParent WHERE LEFT(Type,1)='" & VchType & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text3.SetFocus
            blnRecordExist = False
            cnItemReceiptReturnVoucher.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstItemRVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstItemRVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnItemReceiptReturnVoucher.BeginTrans
            With rstItemRVChild
                If .State = adStateOpen Then
                    If .RecordCount > 0 Then .MoveFirst
                    Do While Not .EOF
                        Call UpdateStatus(VchType, .Fields("VchCode").Value, .Fields("ItemCode").Value, .Fields("Quantity").Value, "-")
                        .MoveNext
                    Loop
                End If
            End With
            cnItemReceiptReturnVoucher.Execute "DELETE FROM BookRVParent WHERE Code='" & rstItemRVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstItemRVList.Delete
                rstItemRVList.MoveNext
                If rstItemRVList.RecordCount > 0 And rstItemRVList.EOF Then rstItemRVList.MoveLast
                cnItemReceiptReturnVoucher.CommitTrans
                ShowProgressInStatusBar True
                Timer1.Enabled = True
            Else
                DisplayError (Err.Description)
                cnItemReceiptReturnVoucher.RollbackTrans
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
        If UpdateRecord(rstItemRVParent) Then
            If UpdateItemList("D") Then
                UpdateFlag = 1
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 3, i
                        .GetText 3, i, CellVal01    'Quantity
                        .GetText 6, i, CellVal02    'Ref No.
                        .GetText 7, i, CellVal03    'Item Code
                        If Val(CellVal01) <> 0 And CellVal02 <> "" And CellVal03 <> "" Then
                            If Not UpdateItemList("I") Then UpdateFlag = 0: Exit For
                        End If
                    Next
                End With
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnItemReceiptReturnVoucher.CommitTrans
            If rstItemRVParent.State = adStateOpen Then rstItemRVParent.Close
            rstItemRVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            Call MsgBox("Record updated !!!", vbInformation, App.Title)
            If AML = "A" Then
                AddModifyList
            ElseIf AML = "M" Then
                FrmGetVchNoToModify.Show
                Unload Me
            Else
                SSTab1.Tab = 0
            End If
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstItemRVParent) Then
            cnItemReceiptReturnVoucher.RollbackTrans
            If rstItemRVParent.State = adStateOpen Then rstItemRVParent.Close
            rstItemRVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            If AML = "A" Then
                Unload Me
            ElseIf AML = "M" Then
                FrmGetVchNoToModify.Show
                Unload Me
            Else
                SSTab1.Tab = 0
            End If
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstItemRVList.Filter = adFilterNone
        rstItemRVList.ActiveConnection = cnItemReceiptReturnVoucher
        Do While Not RefreshRecord(rstItemRVList): Loop
        Set DataGrid1.DataSource = rstItemRVList
        rstItemRVList.ActiveConnection = Nothing
        If rstItemRVList.RecordCount > 0 Then rstItemRVList.MoveLast
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Material Centre", 0
            .Combo1.AddItem "Party", 1
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstItemRVList.RecordCount = 0 Then Exit Sub
        Call PrintItemReceiptVch(rstItemRVList.Fields("Code").Value, "P")
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstItemRVList.RecordCount = 0 Then Exit Sub
        Call PrintItemReceiptVch(rstItemRVList.Fields("Code").Value, "S")
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstItemRVList.RecordCount > 0 Then rstItemRVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstItemRVList.RecordCount > 0 Then
            rstItemRVList.MovePrevious
            If rstItemRVList.BOF Then rstItemRVList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstItemRVList.RecordCount > 0 Then
            rstItemRVList.MoveNext
            If rstItemRVList.EOF Then rstItemRVList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstItemRVList.RecordCount > 0 Then rstItemRVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstItemRVList.EOF Or rstItemRVList.BOF) Then
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
    If ColIndex = 0 Or ColIndex = 2 Then
        SortOrder = DataGrid1.Columns(ColIndex).DataField
        rstItemRVList.Sort = "[" + SortOrder & "] Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstItemRVList.EOF Or rstItemRVList.BOF) Then
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
    If rstItemRVList.RecordCount = 0 Then
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
    If rstItemRVParent.EOF Or rstItemRVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnItemReceiptReturnVoucher, "BookRVParent", "Code", "[Name]+LEFT([Type],1)", Trim(Text2.Text) & VchType, rstItemRVParent.Fields("Code").Value, False) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text3.Text)
        If rstPartyList.RecordCount = 0 Then DisplayError ("No Record in Party Master"): Exit Sub Else rstPartyList.MoveFirst
        rstPartyList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": PartyCode = ""
        Call LoadSelectionList(rstPartyList, "List of Party(s)...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, PartyCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(PartyCode) <> "" Then SendKeys "{TAB}" Else Text3.Text = ""
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If CheckEmpty(Text3.Text, False) Then Cancel = True
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text5.Text)
        If rstMaterialCentreList.RecordCount = 0 Then DisplayError ("No Record in Material Centre Master"): Exit Sub Else rstMaterialCentreList.MoveFirst
        rstMaterialCentreList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": MaterialCentreCode = ""
        Call LoadSelectionList(rstMaterialCentreList, "List of Material Centre(s)...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text5, MaterialCentreCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(MaterialCentreCode) <> "" Then SendKeys "{TAB}" Else Text5.Text = ""
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    If CheckEmpty(Text5.Text, False) Then Cancel = True
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput2.Text)) Then Cancel = True
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstItemRVList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstItemRVParent.State = adStateOpen Then rstItemRVParent.Close
    rstItemRVParent.Open "SELECT * FROM BookRVParent WHERE Code='" & FixQuote(rstItemRVList.Fields("Code").Value) & "'", cnItemReceiptReturnVoucher, adOpenKeyset, adLockOptimistic
    If rstItemRVParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    MhDateInput2.Text = "  -  -    "        'Bill Date
    Text2.Text = "" 'Vch No.
    Text3.Text = "" 'Party Name
    Text5.Text = "" 'Material Centre Name
    Text4.Text = "" 'Remarks
    cmbChallanType.ListIndex = 0
    cmbChallanType.Enabled = True
    Text9.Text = "" 'Challan No.
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput2.Text = Format(Date, "dd-MM-yyyy")
    MhRealInput19.Value = 0 'Total Quantity
    MhRealInput1.Value = 0  'No of Box
    MhRealInput2.Value = 0  'Cartage Amount
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
End Sub
Private Sub LoadFields()
    If rstItemRVParent.EOF Or rstItemRVParent.BOF Then Exit Sub
    Text2.Text = rstItemRVParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstItemRVParent.Fields("Date").Value, "dd-MM-yyyy")
    PartyCode = rstItemRVParent.Fields("Party").Value
    If rstPartyList.RecordCount > 0 Then rstPartyList.MoveFirst
    rstPartyList.Find "[Code] = '" & PartyCode & "'"
    If Not rstPartyList.EOF Then Text3.Text = rstPartyList.Fields("Col0").Value
    MaterialCentreCode = rstItemRVParent.Fields("MaterialCentre").Value
    If rstMaterialCentreList.RecordCount > 0 Then rstMaterialCentreList.MoveFirst
    rstMaterialCentreList.Find "[Code] = '" & MaterialCentreCode & "'"
    If Not rstMaterialCentreList.EOF Then Text5.Text = rstMaterialCentreList.Fields("Col0").Value
    Text4.Text = rstItemRVParent.Fields("Remarks").Value
    cmbChallanType.ListIndex = IIf(Right(rstItemRVParent.Fields("Type").Value, 1) = "F", 0, 1)
    MhRealInput1.Value = Val(rstItemRVParent.Fields("Box").Value)
    MhRealInput2.Value = Val(rstItemRVParent.Fields("Cartage").Value)
    Text9.Text = rstItemRVParent.Fields("ChallanNo").Value
    If Not IsNull(rstItemRVParent.Fields("ChallanDate").Value) Then MhDateInput2.Text = Format(rstItemRVParent.Fields("ChallanDate").Value, "dd-MM-yyyy")
    Call LoadItemList(rstItemRVParent.Fields("Code").Value)
    CalculateTotal
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstItemRVParent.RecordCount = 0 Then Exit Sub
    If rstItemRVParent.State = adStateOpen Then rstItemRVParent.Close
    rstItemRVParent.CursorLocation = adUseServer
    rstItemRVParent.Open "SELECT * FROM BookRVParent WHERE Code='" & FixQuote(rstItemRVList.Fields("Code").Value) & "'", cnItemReceiptReturnVoucher, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstItemRVParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    cmbChallanType.Enabled = False
    Text3.SetFocus
    blnRecordExist = True
    cnItemReceiptReturnVoucher.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstItemRVParent.EOF Or rstItemRVParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstItemRVParent.Fields("Code").Value = GenerateCode(cnItemReceiptReturnVoucher, "SELECT MAX(Code) FROM BookRVParent", 6, "0")
        rstItemRVParent.Fields("CreatedBy").Value = UserCode
        rstItemRVParent.Fields("CreatedOn").Value = Now()
        rstItemRVParent.Fields("Recordstatus").Value = "N"
    Else
        rstItemRVParent.Fields("ModifiedBy").Value = UserCode
        rstItemRVParent.Fields("ModifiedOn").Value = Now()
        rstItemRVParent.Fields("Recordstatus").Value = "M"
    End If
    rstItemRVParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstItemRVParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstItemRVParent.Fields("Party").Value = PartyCode
    rstItemRVParent.Fields("MaterialCentre").Value = MaterialCentreCode
    rstItemRVParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstItemRVParent.Fields("Box").Value = MhRealInput1.Value
    rstItemRVParent.Fields("Cartage").Value = MhRealInput2.Value
    rstItemRVParent.Fields("ChallanNo").Value = Trim(Text9.Text)
    If Not IsDate(MhDateInput2.Text) Then rstItemRVParent.Fields("ChallanDate").Value = Null Else rstItemRVParent.Fields("ChallanDate").Value = GetDate(MhDateInput2.Text)
    rstItemRVParent.Fields("Type").Value = VchType & IIf(cmbChallanType.ListIndex = 0, "F", "U")
    rstItemRVParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstItemRVList.MoveFirst
    rstItemRVList.Find "[Code] = '" & rstItemRVParent.Fields("Code").Value & "'"
    If rstItemRVList.EOF Then rstItemRVList.AddNew
    rstItemRVList.Fields("Code").Value = rstItemRVParent.Fields("Code").Value
    rstItemRVList.Fields("Name").Value = Pad(rstItemRVParent.Fields("Name").Value, Space(1), 10, "L")
    rstItemRVList.Fields("Date").Value = rstItemRVParent.Fields("Date").Value
    rstMaterialCentreList.MoveFirst
    rstMaterialCentreList.Find "[Code] = '" & rstItemRVParent.Fields("MaterialCentre").Value & "'"
    rstItemRVList.Fields("MaterialCentreName").Value = Trim(rstMaterialCentreList.Fields("Col0").Value)
    rstPartyList.MoveFirst
    rstPartyList.Find "[Code] = '" & rstItemRVParent.Fields("Party").Value & "'"
    rstItemRVList.Fields("PartyName").Value = Trim(rstPartyList.Fields("Col0").Value)
    rstItemRVList.Fields("ChallanNo").Value = rstItemRVParent.Fields("ChallanNo").Value
    rstItemRVList.Fields("ChallanDate").Value = rstItemRVParent.Fields("ChallanDate").Value
    rstItemRVList.Update
    rstItemRVList.Sort = SortOrder & " Asc"
    rstItemRVList.Find "[Code] = '" & rstItemRVParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        DisplayError ("Voucher No. cannot be blank"):   Text2.SetFocus:         CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text5.Text, False) Then
        Text5.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnItemReceiptReturnVoucher, "BookRVParent", "Code", "[Name]+LEFT([Type],1)", Trim(Text2.Text) & VchType, rstItemRVParent.Fields("Code").Value, False) Then
        Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text9.Text, False) Then
        Text9.SetFocus: CheckMandatoryFields = True: Exit Function
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
Private Sub LoadItemList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstItemRVChild.State = adStateOpen Then rstItemRVChild.Close
    rstItemRVChild.Open "SELECT T.Item As ItemCode,M.Name As ItemName,Ref As VchCode,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/'+IIF(RIGHT(Ref,2)='08','FG',IIF(RIGHT(Ref,2)='05','MF',IIF(RIGHT(Ref,2)='06','SF','CB'))) As VchNo,Quantity,Rate,Amount FROM (BookRVChild T INNER JOIN BookMaster M ON T.Item=M.Code) INNER JOIN BookPOParent P ON LEFT(T.Ref,6)=P.Code WHERE T.Code='" & strOrderCode & "' ORDER BY M.Name", cnItemReceiptReturnVoucher, adOpenKeyset, adLockOptimistic
    rstItemRVChild.ActiveConnection = Nothing
    If rstItemRVChild.RecordCount > 0 Then rstItemRVChild.MoveFirst
    i = 0
    Do While Not rstItemRVChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstItemRVChild.Fields("ItemName").Value
            .SetText 2, i, rstItemRVChild.Fields("VchNo").Value
            .SetText 3, i, Val(rstItemRVChild.Fields("Quantity").Value)
            .SetText 4, i, Val(rstItemRVChild.Fields("Rate").Value)
            .SetText 5, i, Val(rstItemRVChild.Fields("Amount").Value)
            .SetText 6, i, rstItemRVChild.Fields("VchCode").Value
            .SetText 7, i, rstItemRVChild.Fields("ItemCode").Value
            .SetText 8, i, Val(rstItemRVChild.Fields("Quantity").Value)
        End With
        rstItemRVChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Book List")
End Sub
Private Sub CalculateTotal()
    Dim i As Integer, Qty As Variant
    MhRealInput19.Value = 0
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 3, i, Qty
            MhRealInput19.Value = Val(MhRealInput19.Text) + Qty
        Next
    End With
End Sub
Private Function UpdateItemList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 5) As Variant
    On Error GoTo ErrorHandler
    UpdateItemList = True
    If ActionType = "D" Then
        If Not blnRecordExist Then Exit Function
        With rstItemRVChild
            If .State = adStateOpen Then
                If .RecordCount > 0 Then .MoveFirst
                Do While Not .EOF
                    Call UpdateStatus(VchType, .Fields("VchCode").Value, .Fields("ItemCode").Value, .Fields("Quantity").Value, "-")
                    .MoveNext
                Loop
            End If
        End With
        cnItemReceiptReturnVoucher.Execute "DELETE FROM BookRVChild WHERE Code='" & rstItemRVParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 3, .ActiveRow, CellVal(1)  'Quantity
            .GetText 4, .ActiveRow, CellVal(2)  'Rate
            .GetText 5, .ActiveRow, CellVal(3)  'Amount
            .GetText 6, .ActiveRow, CellVal(4)  'Ref No.
            .GetText 7, .ActiveRow, CellVal(5)  'Item Code
        End With
        cnItemReceiptReturnVoucher.Execute "INSERT INTO BookRVChild VALUES ('" & rstItemRVParent.Fields("Code").Value & "','" & CellVal(4) & "','" & CellVal(5) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ")"
        Call UpdateStatus(VchType, CellVal(4), CellVal(5), Val(CellVal(1)), "+")
    End If
    Exit Function
ErrorHandler:
    UpdateItemList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Party" Then
        rstItemRVList.Filter = "[PartyName] Like '%" & SrchText & "%'"
    ElseIf SrchFor = "Material Centre" Then
        rstItemRVList.Filter = "[MaterialCentreName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1: fpSpread1.SetFocus
            CalculateTotal
        End If
    ElseIf KeyCode = vbKeySpace Then
        Dim Item As Variant, Qty As Variant, Ref As Variant
        With fpSpread1
            If .ActiveCol = 1 Then
                .GetText .ActiveCol, .ActiveRow, Item
                Text6.Text = FixQuote(Item)
                If rstItemList.RecordCount = 0 Then DisplayError ("No Record in Item Master"): .SetActiveCell 1, .ActiveRow: .SetFocus: Exit Sub Else rstItemList.MoveFirst
                rstItemList.Find "[Col0] = '" & FixQuote(Trim(Item)) & "'"
                SelectionType = "S"
                ItemCode = ""
                Call LoadSelectionList(rstItemList, "List of Items...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text6, ItemCode)
                Call CloseForm(FrmSelectionList)
                If ItemCode = "" Then
                    .SetActiveCell 1, .ActiveRow
                Else
                    rstItemList.MoveFirst: rstItemList.Find "[Code] ='" & ItemCode & "'"
                    .SetText 1, .ActiveRow, Text6.Text
                    .SetText 4, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                    .SetText 7, .ActiveRow, ItemCode
                    .SetFocus
                    SendKeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 2 Then
                .GetText 7, .ActiveRow, Item   'Item Code
                If Item = "" Then Exit Sub
                If rstRefList.State = adStateOpen Then rstRefList.Close
                .GetText 6, .ActiveRow, Ref    'Ref Code
                .GetText 8, .ActiveRow, Qty    'Old Billing Quantity
                If rstRefList.State = adStateOpen Then rstRefList.Close
                If VchType = "R" Then   'Receipt Vch
                    If cmbChallanType.ListIndex = 0 Then
                        rstRefList.Open "SELECT TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG Bal Qty: '+TRIM(IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'08'='" & Ref & "'," & Val(Qty) & ",0)) As Col0,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'08'='" & Ref & "'," & Val(Qty) & ",0) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG' As Name,P.Code+'08' As Code FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE Book='" + Item + "' AND P.Binder='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'08'='" & Ref & "'," & Val(Qty) & ",0)>0", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
                    Else
                        rstRefList.Open "SELECT TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF Bal Qty: '+TRIM(IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'05'='" & Ref & "'," & Val(Qty) & ",0)) As Col0,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'05'='" & Ref & "'," & Val(Qty) & ",0) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As Name,P.Code+'05' As Code FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE Book='" + Item + "' AND P.BookPrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'05'='" & Ref & "'," & Val(Qty) & ",0)>0 UNION " & _
                                                         "SELECT TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/SF Bal Qty: '+TRIM(IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'06'='" & Ref & "'," & Val(Qty) & ",0)) As Col0,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'06'='" & Ref & "'," & Val(Qty) & ",0) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/SF' As Name,P.Code+'06' As Code FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE Book='" + Item + "' AND P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'06'='" & Ref & "'," & Val(Qty) & ",0)>0 UNION " & _
                                                         "SELECT TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CB Bal Qty: '+TRIM(IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'09'='" & Ref & "'," & Val(Qty) & ",0)) As Col0,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'09'='" & Ref & "'," & Val(Qty) & ",0) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CB' As Name,P.Code+'09' As Code FROM (BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C ON P.Code=C.Code WHERE C.Book='" + Item + "' AND P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)+IIF(P.Code+'09'='" & Ref & "'," & Val(Qty) & ",0)>0", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
                    End If
                Else    'Issue Vch
                    If cmbChallanType.ListIndex = 0 Then
                        rstRefList.Open "SELECT TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG Bal Qty: '+TRIM(IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'08'='" & Ref & "'," & Val(Qty) & ",0)) As Col0,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'08'='" & Ref & "'," & Val(Qty) & ",0) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG' As Name,P.Code+'08' As Code FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE Book='" + Item + "' AND P.Binder='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'08'='" & Ref & "'," & Val(Qty) & ",0)>0", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
                    Else
                        rstRefList.Open "SELECT TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF Bal Qty: '+TRIM(IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'05'='" & Ref & "'," & Val(Qty) & ",0)) As Col0,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'05'='" & Ref & "'," & Val(Qty) & ",0) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As Name,P.Code+'05' As Code FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE Book='" + Item + "' AND P.BookPrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'05'='" & Ref & "'," & Val(Qty) & ",0)>0 UNION " & _
                                                         "SELECT TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/SF Bal Qty: '+TRIM(IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'06'='" & Ref & "'," & Val(Qty) & ",0)) As Col0,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'06'='" & Ref & "'," & Val(Qty) & ",0) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/SF' As Name,P.Code+'06' As Code FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE Book='" + Item + "' AND P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'06'='" & Ref & "'," & Val(Qty) & ",0)>0 UNION " & _
                                                         "SELECT TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CB Bal Qty: '+TRIM(IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'09'='" & Ref & "'," & Val(Qty) & ",0)) As Col0,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'09'='" & Ref & "'," & Val(Qty) & ",0) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CB' As Name,P.Code+'09' As Code FROM (BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C ON P.Code=C.Code WHERE C.Book='" + Item + "' AND P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)+IIF(P.Code+'09'='" & Ref & "'," & Val(Qty) & ",0)>0", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
                    End If
                End If
                .GetText .ActiveCol, .ActiveRow, Item
                Text6.Text = FixQuote(Item)
                If rstRefList.RecordCount = 0 Then DisplayError ("No Pending Order"): .SetActiveCell 2, .ActiveRow: .SetFocus: Exit Sub Else rstRefList.MoveFirst
                rstRefList.Find "[Col0] = '" & FixQuote(RTrim(Item)) & "'"
                SelectionType = "S"
                RefCode = ""
                Call LoadSelectionList(rstRefList, "List of Pending Orders...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text6, RefCode)
                Call CloseForm(FrmSelectionList)
                If RefCode = "" Then
                    .SetActiveCell 2, .ActiveRow
                Else
                    rstRefList.MoveFirst: rstRefList.Find "[Code] ='" & RefCode & "'"
                    .SetText 2, .ActiveRow, Trim(rstRefList.Fields("Name").Value)
                    .GetText 3, .ActiveRow, Qty
                    If Qty = 0 Then .SetText 3, .ActiveRow, Val(rstRefList.Fields("BalQty").Value)
                    .SetText 6, .ActiveRow, RefCode
                    .SetFocus
                    SendKeys "{ENTER}"
                End If
            End If
        End With
    ElseIf KeyCode = vbKeyF11 Then
        If fpSpread1.DataRowCnt = 0 Then LoadOrderList
    End If
    If fpSpread1.DataRowCnt > 0 Then cmbChallanType.Enabled = False Else cmbChallanType.Enabled = True
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Qty As Variant, Rate As Variant
    With fpSpread1
        If Col = 1 Or Col = 4 Then
            .GetText 3, Row, Qty
            .GetText 4, Row, Rate
            If Qty > 0 Then .SetText 5, Row, Qty * Rate
            CalculateTotal
        End If
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub UpdateStatus(ByVal VchType As String, ByVal VchCode As String, ByVal Item As String, ByVal Quantity As Integer, ByVal Operation As String)
    If VchType = "R" Then
        If cmbChallanType.ListIndex = 0 Then   'FG
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code SET C.QuantityReceived=C.QuantityReceived" + Operation + Trim(Quantity) + " WHERE P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '08'='" & Right(VchCode, 2) & "'"
                '
            cnItemReceiptReturnVoucher.Execute "UPDATE (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code SET C.Status='D' WHERE C1.QuantityReceived>0 AND RIGHT(P.Type,1)='P' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '08'='" & Right(VchCode, 2) & "'"  'PO
            cnItemReceiptReturnVoucher.Execute "UPDATE (BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code SET C.Status='D' WHERE C1.QuantityReceived>0 AND RIGHT(P.Type,1)='P' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '08'='" & Right(VchCode, 2) & "'"  'PO
            cnItemReceiptReturnVoucher.Execute "UPDATE (BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code) INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code SET C.Status='D' WHERE C1.QuantityReceived>0 AND RIGHT(P.Type,1)='P' AND P.Code='" + Left(VchCode, 6) + "' AND C.Book='" & Item & "' AND '08'='" & Right(VchCode, 2) & "'"  'PO
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code SET C.Status='D' WHERE C.QuantityReceived>0 AND RIGHT(P.Type,1)='P' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '08'='" & Right(VchCode, 2) & "'"  'PO
        Else
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code SET C.QuantityReceived=C.QuantityReceived" + Operation + Trim(Quantity) + " WHERE P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '05'='" & Right(VchCode, 2) & "'"
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code SET C.QuantityReceived=C.QuantityReceived" + Operation + Trim(Quantity) + " WHERE P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '06'='" & Right(VchCode, 2) & "'"
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code SET C.QuantityReceived=C.QuantityReceived" + Operation + Trim(Quantity) + " WHERE P.Code='" + Left(VchCode, 6) + "' AND C.Book='" & Item & "' AND '09'='" & Right(VchCode, 2) & "'"
            '
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild05  C ON P.Code=C.Code SET Status='D' WHERE C.QuantityReceived>0 AND RIGHT(P.Type,1)='P' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '05'='" & Right(VchCode, 2) & "'"  'PO
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild06  C ON P.Code=C.Code SET Status='D' WHERE C.QuantityReceived>0 AND RIGHT(P.Type,1)='P' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '06'='" & Right(VchCode, 2) & "'"  'PO
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code SET C.Status='D' WHERE C.QuantityReceived>0 AND RIGHT(P.Type,1)='P' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '09'='" & Right(VchCode, 2) & "'"  'PO
        End If
    Else
        If cmbChallanType.ListIndex = 0 Then   'FG
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code SET C.QuantityIssued=C.QuantityIssued" + Operation + Trim(Quantity) + " WHERE P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '08'='" & Right(VchCode, 2) & "'"
            '
            cnItemReceiptReturnVoucher.Execute "UPDATE (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code SET C.Status='D' WHERE C1.QuantityIssued>0 AND RIGHT(P.Type,1)='S' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '08'='" & Right(VchCode, 2) & "'"  'SO
            cnItemReceiptReturnVoucher.Execute "UPDATE (BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code SET C.Status='D' WHERE C1.QuantityIssued>0 AND RIGHT(P.Type,1)='S' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '08'='" & Right(VchCode, 2) & "'"  'SO
            cnItemReceiptReturnVoucher.Execute "UPDATE (BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code) INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code SET C.Status='D' WHERE C1.QuantityIssued>0 AND RIGHT(P.Type,1)='S' AND P.Code='" + Left(VchCode, 6) + "' AND C.Book='" & Item & "' AND '08'='" & Right(VchCode, 2) & "'"  'SO
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code SET C.Status='D' WHERE C.QuantityIssued>0 AND RIGHT(P.Type,1)='S' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '08'='" & Right(VchCode, 2) & "'"  'SO
        Else
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code SET C.QuantityIssued=C.QuantityIssued" + Operation + Trim(Quantity) + " WHERE P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '05'='" & Right(VchCode, 2) & "'"
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code SET C.QuantityIssued=C.QuantityIssued" + Operation + Trim(Quantity) + " WHERE P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '06'='" & Right(VchCode, 2) & "'"
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code SET C.QuantityIssued=C.QuantityIssued" + Operation + Trim(Quantity) + " WHERE P.Code='" + Left(VchCode, 6) + "' AND C.Book='" & Item & "' AND '09'='" & Right(VchCode, 2) & "'"
            '
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild05  C ON P.Code=C.Code SET Status='D' WHERE C.QuantityIssued>0 AND RIGHT(P.Type,1)='S' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '05'='" & Right(VchCode, 2) & "'"  'SO
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild06  C ON P.Code=C.Code SET Status='D' WHERE C.QuantityIssued>0 AND RIGHT(P.Type,1)='S' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '06'='" & Right(VchCode, 2) & "'"  'SO
            cnItemReceiptReturnVoucher.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code SET C.Status='D' WHERE C.QuantityIssued>0 AND RIGHT(P.Type,1)='S' AND P.Code='" + Left(VchCode, 6) + "' AND P.Book='" & Item & "' AND '09'='" & Right(VchCode, 2) & "'"  'SO
        End If
    End If
End Sub
Private Sub LoadOrderList()
    If rstOrderList.State = adStateOpen Then rstOrderList.Close
    If VchType = "R" Then   'Receipt Vch
        If cmbChallanType.ListIndex = 0 Then
            rstOrderList.Open "SELECT P.Code+'08'+P.Book As VchCode,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG' As VchNo,P.Date As VchDate,I.Name As Item,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived) As BalQty,IIF(C.Status='D','Delivered','Undelivered') As Status FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Binder='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)>0 ORDER BY I.Name,P.Date,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG'", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
        Else
            rstOrderList.Open "SELECT P.Code+'05'+P.Book As VchCode,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF'  As VchNo,P.Date As VchDate,I.Name As Item,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived) As BalQty,IIF(C.Status='D','Delivered','Undelivered') As Status FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.BookPrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)>0 UNION " & _
                                                  "SELECT P.Code+'06'+P.Book As VchCode,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/SF' As VchNo,P.Date As VchDate,I.Name As Item,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived) As BalQty,IIF(C.Status='D','Delivered','Undelivered') As Status FROM (BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)>0 UNION " & _
                                                  "SELECT P.Code+'09'+C.Book As VchCode,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CB'  As VchNo,P.Date As VchDate,I.Name As Item,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived) As BalQty,IIF(C.Status='D','Delivered','Undelivered') As Status FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C ON P.Code=C.Code) INNER JOIN BookMaster I ON C.Book=I.Code WHERE P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived)>0 " & _
                                                  "ORDER BY Item,VchDate,VchNo", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
        End If
    Else    'Issue Vch
        If cmbChallanType.ListIndex = 0 Then
            rstOrderList.Open "SELECT P.Code+'08'+P.Book As VchCode,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG' As VchNo,P.Date As VchDate,I.Name As Item,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued) As BalQty,IIF(C.Status='D','Delivered','Undelivered') As Status FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Binder='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)>0 ORDER BY I.Name,P.Date,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG'", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
        Else
            rstOrderList.Open "SELECT P.Code+'05'+P.Book As VchCode,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As VchNo,P.Date As VchDate,I.Name As Item,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued) As BalQty,IIF(C.Status='D','Delivered','Undelivered') As Status FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.BookPrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)>0 UNION " & _
                                                  "SELECT P.Code+'06'+P.Book As VchCode,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/SF' As VchNo,P.Date As VchDate,I.Name As Item,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued) As BalQty,IIF(C.Status='D','Delivered','Undelivered') As Status FROM (BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)>0 UNION " & _
                                                  "SELECT P.Code+'09'+C.Book As VchCode,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CB'  As VchNo,P.Date As VchDate,I.Name As Item,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued) As BalQty,IIF(C.Status='D','Delivered','Undelivered') As Status FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C ON P.Code=C.Code) INNER JOIN BookMaster I ON C.Book=I.Code WHERE P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued)>0 " & _
                                                  "ORDER BY Item,VchDate,VchNo", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
        End If
    End If
    rstOrderList.ActiveConnection = Nothing
    If rstOrderList.RecordCount = 0 Then DisplayError ("No Pending Order Exists"): fpSpread1.SetFocus: Exit Sub
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
            FrmOrderList.fpSpread1.SetText 1, i, .Fields("Item").Value
            FrmOrderList.fpSpread1.SetText 2, i, .Fields("VchNo").Value
            FrmOrderList.fpSpread1.SetText 3, i, Format(.Fields("VchDate").Value, "dd-MM-yy")
            FrmOrderList.fpSpread1.SetText 4, i, Val(.Fields("TotalQty").Value)
            FrmOrderList.fpSpread1.SetText 5, i, Val(.Fields("BalQty").Value)
            FrmOrderList.fpSpread1.SetText 6, i, .Fields("Status").Value
            FrmOrderList.fpSpread1.SetText 7, i, 0
            FrmOrderList.fpSpread1.SetText 8, i, .Fields("VchCode").Value
            .MoveNext
        Loop
        FrmOrderList.fpSpread1.SetActiveCell 7, 1
    End With
    FrmOrderList.Check2 = 0
    FrmOrderList.Show vbModal
    If Not CheckEmpty(FrmOrderList.VchCodeList, False) Then
        If rstOrderList.State = adStateOpen Then rstOrderList.Close
        If VchType = "R" Then   'Receipt Vch
            If cmbChallanType.ListIndex = 0 Then
                rstOrderList.Open "SELECT I.Name As ItemName,I.Code As ItemCode,I.Price,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG' As VchNo,P.Code+'08' As VchCode FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Code+'08'+P.Book IN (" & FrmOrderList.VchCodeList & ") ORDER BY I.Name,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG'", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
            Else
                rstOrderList.Open "SELECT I.Name As ItemName,I.Code As ItemCode,I.Price,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As VchNo,P.Code+'05' As VchCode FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Code+'05'+P.Book IN (" & FrmOrderList.VchCodeList & ") UNION " & _
                                                      "SELECT I.Name As ItemName,I.Code As ItemCode,I.Price,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/SF' As VchNo,P.Code+'06' As VchCode FROM (BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Code+'06'+P.Book IN (" & FrmOrderList.VchCodeList & ") UNION " & _
                                                      "SELECT I.Name As ItemName,I.Code As ItemCode,I.Price,IIF(RIGHT(P.Type,1)='S',C.QuantityIssued-C.QuantityReceived,C.ActualQuantity+C.QuantityIssued-C.QuantityReceived) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CB' As VchNo,P.Code+'09' As VchCode FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C ON P.Code=C.Code) INNER JOIN BookMaster I ON C.Book=I.Code WHERE P.Code+'09'+C.Book IN (" & FrmOrderList.VchCodeList & ") " & _
                                                      "ORDER BY ItemName,VchNo", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
            End If
        Else    'Issue Vch
            If cmbChallanType.ListIndex = 0 Then
                rstOrderList.Open "SELECT I.Name As ItemName,I.Code As ItemCode,I.Price,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG' As VchNo,P.Code+'08' As VchCode FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Code+'08'+P.Book IN (" & FrmOrderList.VchCodeList & ") ORDER BY I.Name,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FG'", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
            Else
                rstOrderList.Open "SELECT I.Name As ItemName,I.Code As ItemCode,I.Price,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As VchNo,P.Code+'05' As VchCode FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Code+'05'+P.Book IN (" & FrmOrderList.VchCodeList & ") UNION " & _
                                                      "SELECT I.Name As ItemName,I.Code As ItemCode,I.Price,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/SF' As VchNo,P.Code+'06' As VchCode FROM (BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Code+'06'+P.Book IN (" & FrmOrderList.VchCodeList & ") UNION " & _
                                                      "SELECT I.Name As ItemName,I.Code As ItemCode,I.Price,IIF(RIGHT(P.Type,1)='S',C.ActualQuantity+C.QuantityReceived-C.QuantityIssued,C.QuantityReceived-C.QuantityIssued) As BalQty,TRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CB' As VchNo,P.Code+'09' As VchCode FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C ON P.Code=C.Code) INNER JOIN BookMaster I ON C.Book=I.Code WHERE P.Code+'09'+C.Book IN (" & FrmOrderList.VchCodeList & ") " & _
                                                      "ORDER BY ItemName,VchNo", cnItemReceiptReturnVoucher, adOpenKeyset, adLockReadOnly
            End If
        End If
        If rstOrderList.RecordCount > 0 Then
            i = 0
            With fpSpread1
                Do While Not rstOrderList.EOF
                    i = i + 1
                    .SetText 1, i, rstOrderList.Fields("ItemName").Value
                    .SetText 2, i, rstOrderList.Fields("VchNo").Value
                    .SetText 3, i, Val(rstOrderList.Fields("BalQty").Value)
                    .SetText 4, i, Val(rstOrderList.Fields("Price").Value)
                    .SetText 5, i, Val(rstOrderList.Fields("BalQty").Value) * Val(rstOrderList.Fields("Price").Value)
                    .SetText 6, i, rstOrderList.Fields("VchCode").Value
                    .SetText 7, i, rstOrderList.Fields("ItemCode").Value
                    rstOrderList.MoveNext
                Loop
                Call CalculateTotal
            End With
        End If
    End If
    CloseForm FrmOrderList
End Sub
Public Sub PrintItemReceiptVch(ByVal OrderCode As String, Optional ByVal OutputType As String)
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptBookReceiptVoucher.Text11.SetText "Item " + IIf(VchType = "I", "Issue", "Receipt") + " Voucher"
    rptBookReceiptVoucher.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptBookReceiptVoucher.Text6.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False)) Then
        rptBookReceiptVoucher.Text8.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False) Then
        rptBookReceiptVoucher.Text8.SetText "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptBookReceiptVoucher.Text8.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value)
    Else
        rptBookReceiptVoucher.Section9.Suppress = True
    End If
    If rstItemRVChild.State = adStateOpen Then rstItemRVChild.Close
    rstItemRVChild.Open "SELECT IIF(RIGHT(C.Ref,2)='08','FG',IIF(RIGHT(C.Ref,2)='05','MF',IIF(RIGHT(C.Ref,2)='06','SF','CB')))+'/'+TRIM(P.Name) As VchNo,P.Date As VchDate,A2.Name As MaterialCentreName,A1.Name As BinderName,P.Remarks,Box,Cartage,ChallanNo,ChallanDate,B.Name As BookName,TRIM(T.Name) As PONo,Quantity,Rate As Price,Amount FROM ((((BookRVParent P INNER JOIN BookRVChild C ON P.Code=C.Code) INNER JOIN BookMaster B ON C.Item=B.Code) INNER JOIN AccountMaster A1 ON P.Party=A1.Code) INNER JOIN BookPOParent T ON LEFT(C.Ref,6)=T.Code) INNER JOIN AccountMaster A2 ON P.MaterialCentre=A2.Code WHERE P.Code='" + OrderCode + "' ORDER BY B.Name", cnItemReceiptReturnVoucher, adOpenKeyset, adLockOptimistic
    rptBookReceiptVoucher.Database.SetDataSource rstItemRVChild, 3, 1
    Screen.MousePointer = vbNormal
    If OutputType = "S" Then
        Set FrmReportViewer.Report = rptBookReceiptVoucher
        FrmReportViewer.Show vbModal
    Else
        rptBookReceiptVoucher.PrintOut
    End If
    Set rptBookReceiptVoucher = Nothing
    On Error GoTo 0
End Sub
Private Sub AddModifyList()
    If AML = "A" Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1): Text3.SetFocus
    ElseIf AML = "M" Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2): Text3.SetFocus
    Else
        Text1.SetFocus
    End If
End Sub
