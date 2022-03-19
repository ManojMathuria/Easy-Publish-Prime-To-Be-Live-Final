VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmItemIssueReceiptVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Issue Receipt Voucher"
   ClientHeight    =   9315
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
   ScaleHeight     =   9315
   ScaleWidth      =   13740
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9300
      Left            =   -105
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   13830
      _Version        =   65536
      _ExtentX        =   24395
      _ExtentY        =   16404
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
      Picture         =   "ItemIssueReceiptVoucher.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   9075
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   120
         Width           =   13485
         _ExtentX        =   23786
         _ExtentY        =   16007
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
         TabPicture(0)   =   "ItemIssueReceiptVoucher.frx":001C
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
         TabPicture(1)   =   "ItemIssueReceiptVoucher.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtNotes"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "btnNotes"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Mh3dFrame2"
         Tab(1).Control(2).Enabled=   0   'False
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
            TabIndex        =   57
            ToolTipText     =   "Open Notes"
            Top             =   8550
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
            TabIndex        =   56
            Top             =   8550
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
            Left            =   720
            MaxLength       =   40
            TabIndex        =   48
            Top             =   8550
            Width           =   8100
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   8025
            Left            =   120
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   450
            Width           =   13260
            _ExtentX        =   23389
            _ExtentY        =   14155
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
               DataField       =   "PartyName"
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
            BeginProperty Column06 
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
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1230.236
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   3899.906
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   2534.74
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
                  ColumnWidth     =   1725.165
               EndProperty
               BeginProperty Column05 
                  Locked          =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1200.189
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7935
            Left            =   -74880
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   480
            Width           =   13260
            _Version        =   65536
            _ExtentX        =   23389
            _ExtentY        =   13996
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
            Picture         =   "ItemIssueReceiptVoucher.frx":0054
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
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   0
               Top             =   120
               Width           =   3090
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
               Left            =   9900
               MaxLength       =   40
               TabIndex        =   5
               Top             =   675
               Width           =   3250
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
               Left            =   5460
               MaxLength       =   255
               TabIndex        =   10
               Top             =   1305
               Width           =   3135
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
               Left            =   5460
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   4
               Top             =   675
               Width           =   3135
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
               Top             =   990
               Width           =   3195
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   26
               Top             =   6690
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
               Picture         =   "ItemIssueReceiptVoucher.frx":0070
               Picture         =   "ItemIssueReceiptVoucher.frx":008C
               Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
                  Height          =   285
                  Left            =   9870
                  TabIndex        =   27
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   930
                  _Version        =   65536
                  _ExtentX        =   1640
                  _ExtentY        =   503
                  Calculator      =   "ItemIssueReceiptVoucher.frx":00A8
                  Caption         =   "ItemIssueReceiptVoucher.frx":00C8
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "ItemIssueReceiptVoucher.frx":0134
                  Keys            =   "ItemIssueReceiptVoucher.frx":0152
                  Spin            =   "ItemIssueReceiptVoucher.frx":019C
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
                  ValueVT         =   223084549
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
                  Height          =   285
                  Left            =   11610
                  TabIndex        =   30
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1185
                  _Version        =   65536
                  _ExtentX        =   2090
                  _ExtentY        =   503
                  Calculator      =   "ItemIssueReceiptVoucher.frx":01C4
                  Caption         =   "ItemIssueReceiptVoucher.frx":01E4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "ItemIssueReceiptVoucher.frx":0250
                  Keys            =   "ItemIssueReceiptVoucher.frx":026E
                  Spin            =   "ItemIssueReceiptVoucher.frx":02B8
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
               Left            =   5460
               MaxLength       =   25
               TabIndex        =   1
               Top             =   105
               Width           =   3015
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
               TabIndex        =   9
               Top             =   1305
               Width           =   3195
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
               Top             =   675
               Width           =   3195
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   24
               Top             =   675
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
               Picture         =   "ItemIssueReceiptVoucher.frx":02E0
               Picture         =   "ItemIssueReceiptVoucher.frx":02FC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   25
               Top             =   1305
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ItemIssueReceiptVoucher.frx":0318
               Picture         =   "ItemIssueReceiptVoucher.frx":0334
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   9900
               TabIndex        =   2
               Top             =   105
               Width           =   3255
               _Version        =   65536
               _ExtentX        =   5741
               _ExtentY        =   582
               Calendar        =   "ItemIssueReceiptVoucher.frx":0350
               Caption         =   "ItemIssueReceiptVoucher.frx":0468
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":04D4
               Keys            =   "ItemIssueReceiptVoucher.frx":04F2
               Spin            =   "ItemIssueReceiptVoucher.frx":0550
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
               TabIndex        =   13
               Top             =   1830
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
               MaxCols         =   12
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "ItemIssueReceiptVoucher.frx":0578
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
               Left            =   840
               MaxLength       =   100
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   3240
               Width           =   11715
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   8580
               TabIndex        =   29
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
               Picture         =   "ItemIssueReceiptVoucher.frx":123E
               Picture         =   "ItemIssueReceiptVoucher.frx":125A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   31
               Top             =   990
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
               Picture         =   "ItemIssueReceiptVoucher.frx":1276
               Picture         =   "ItemIssueReceiptVoucher.frx":1292
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   4380
               TabIndex        =   32
               Top             =   990
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
               Caption         =   " Item Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ItemIssueReceiptVoucher.frx":12AE
               Picture         =   "ItemIssueReceiptVoucher.frx":12CA
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   9945
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   7170
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "ItemIssueReceiptVoucher.frx":12E6
               Caption         =   "ItemIssueReceiptVoucher.frx":1306
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":1372
               Keys            =   "ItemIssueReceiptVoucher.frx":1390
               Spin            =   "ItemIssueReceiptVoucher.frx":13DA
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
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   7170
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   582
               Calculator      =   "ItemIssueReceiptVoucher.frx":1402
               Caption         =   "ItemIssueReceiptVoucher.frx":1422
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":148E
               Keys            =   "ItemIssueReceiptVoucher.frx":14AC
               Spin            =   "ItemIssueReceiptVoucher.frx":14F6
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
               ValueVT         =   1903493125
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel35 
               Height          =   645
               Left            =   10785
               TabIndex        =   35
               Top             =   7170
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
               Picture         =   "ItemIssueReceiptVoucher.frx":151E
               Picture         =   "ItemIssueReceiptVoucher.frx":153A
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   645
               Left            =   11985
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   7170
               Width           =   1170
               _Version        =   65536
               _ExtentX        =   2055
               _ExtentY        =   1147
               Calculator      =   "ItemIssueReceiptVoucher.frx":1556
               Caption         =   "ItemIssueReceiptVoucher.frx":1576
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":15E2
               Keys            =   "ItemIssueReceiptVoucher.frx":1600
               Spin            =   "ItemIssueReceiptVoucher.frx":164A
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
               Height          =   645
               Left            =   1200
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   7170
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1147
               Calculator      =   "ItemIssueReceiptVoucher.frx":1672
               Caption         =   "ItemIssueReceiptVoucher.frx":1692
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":16FE
               Keys            =   "ItemIssueReceiptVoucher.frx":171C
               Spin            =   "ItemIssueReceiptVoucher.frx":1766
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
               Height          =   645
               Left            =   120
               TabIndex        =   38
               Top             =   7170
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
               Picture         =   "ItemIssueReceiptVoucher.frx":178E
               Picture         =   "ItemIssueReceiptVoucher.frx":17AA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
               Height          =   330
               Left            =   8430
               TabIndex        =   39
               Top             =   7485
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
               Picture         =   "ItemIssueReceiptVoucher.frx":17C6
               Picture         =   "ItemIssueReceiptVoucher.frx":17E2
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
               Height          =   330
               Left            =   9945
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   7485
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "ItemIssueReceiptVoucher.frx":17FE
               Caption         =   "ItemIssueReceiptVoucher.frx":181E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":188A
               Keys            =   "ItemIssueReceiptVoucher.frx":18A8
               Spin            =   "ItemIssueReceiptVoucher.frx":18F2
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
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   7485
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   582
               Calculator      =   "ItemIssueReceiptVoucher.frx":191A
               Caption         =   "ItemIssueReceiptVoucher.frx":193A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":19A6
               Keys            =   "ItemIssueReceiptVoucher.frx":19C4
               Spin            =   "ItemIssueReceiptVoucher.frx":1A0E
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
               ValueVT         =   1903493125
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   645
               Left            =   5235
               TabIndex        =   15
               Top             =   7170
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1147
               Calculator      =   "ItemIssueReceiptVoucher.frx":1A36
               Caption         =   "ItemIssueReceiptVoucher.frx":1A56
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":1AC2
               Keys            =   "ItemIssueReceiptVoucher.frx":1AE0
               Spin            =   "ItemIssueReceiptVoucher.frx":1B2A
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
               TabIndex        =   42
               Top             =   7170
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
               Picture         =   "ItemIssueReceiptVoucher.frx":1B52
               Picture         =   "ItemIssueReceiptVoucher.frx":1B6E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   645
               Left            =   2280
               TabIndex        =   43
               Top             =   7170
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
               Picture         =   "ItemIssueReceiptVoucher.frx":1B8A
               Picture         =   "ItemIssueReceiptVoucher.frx":1BA6
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   645
               Left            =   3120
               TabIndex        =   14
               Top             =   7170
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   1147
               Calculator      =   "ItemIssueReceiptVoucher.frx":1BC2
               Caption         =   "ItemIssueReceiptVoucher.frx":1BE2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":1C4E
               Keys            =   "ItemIssueReceiptVoucher.frx":1C6C
               Spin            =   "ItemIssueReceiptVoucher.frx":1CB6
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
               Left            =   8430
               TabIndex        =   44
               Top             =   7170
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
               Picture         =   "ItemIssueReceiptVoucher.frx":1CDE
               Picture         =   "ItemIssueReceiptVoucher.frx":1CFA
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   645
               Left            =   3675
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   7170
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   1147
               Calculator      =   "ItemIssueReceiptVoucher.frx":1D16
               Caption         =   "ItemIssueReceiptVoucher.frx":1D36
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":1DA2
               Keys            =   "ItemIssueReceiptVoucher.frx":1DC0
               Spin            =   "ItemIssueReceiptVoucher.frx":1E0A
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
               TabIndex        =   45
               Top             =   7170
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
               Picture         =   "ItemIssueReceiptVoucher.frx":1E32
               Picture         =   "ItemIssueReceiptVoucher.frx":1E4E
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   645
               Left            =   7395
               TabIndex        =   16
               Top             =   7170
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   1138
               Calculator      =   "ItemIssueReceiptVoucher.frx":1E6A
               Caption         =   "ItemIssueReceiptVoucher.frx":1E8A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":1EF6
               Keys            =   "ItemIssueReceiptVoucher.frx":1F14
               Spin            =   "ItemIssueReceiptVoucher.frx":1F5E
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
               Left            =   4380
               TabIndex        =   46
               Top             =   675
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
               Caption         =   " Mat Centre"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ItemIssueReceiptVoucher.frx":1F86
               Picture         =   "ItemIssueReceiptVoucher.frx":1FA2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
               Height          =   330
               Left            =   11175
               TabIndex        =   47
               Top             =   1305
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
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
               Caption         =   " No. of Box"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ItemIssueReceiptVoucher.frx":1FBE
               Picture         =   "ItemIssueReceiptVoucher.frx":1FDA
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
               Height          =   330
               Left            =   12135
               TabIndex        =   12
               Top             =   1305
               Width           =   1020
               _Version        =   65536
               _ExtentX        =   1799
               _ExtentY        =   582
               Calculator      =   "ItemIssueReceiptVoucher.frx":1FF6
               Caption         =   "ItemIssueReceiptVoucher.frx":2016
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":2082
               Keys            =   "ItemIssueReceiptVoucher.frx":20A0
               Spin            =   "ItemIssueReceiptVoucher.frx":20EA
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   8580
               TabIndex        =   50
               Top             =   1305
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
               Caption         =   " Challan Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ItemIssueReceiptVoucher.frx":2112
               Picture         =   "ItemIssueReceiptVoucher.frx":212E
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput2 
               Height          =   330
               Left            =   9900
               TabIndex        =   11
               Top             =   1305
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   582
               Calendar        =   "ItemIssueReceiptVoucher.frx":214A
               Caption         =   "ItemIssueReceiptVoucher.frx":2262
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "ItemIssueReceiptVoucher.frx":22CE
               Keys            =   "ItemIssueReceiptVoucher.frx":22EC
               Spin            =   "ItemIssueReceiptVoucher.frx":234A
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   330
               Left            =   120
               TabIndex        =   51
               Top             =   120
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
               Picture         =   "ItemIssueReceiptVoucher.frx":2372
               Picture         =   "ItemIssueReceiptVoucher.frx":238E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   4380
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
               Caption         =   " Vch No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ItemIssueReceiptVoucher.frx":23AA
               Picture         =   "ItemIssueReceiptVoucher.frx":23C6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   4380
               TabIndex        =   53
               Top             =   1305
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
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
               Picture         =   "ItemIssueReceiptVoucher.frx":23E2
               Picture         =   "ItemIssueReceiptVoucher.frx":23FE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   8580
               TabIndex        =   54
               Top             =   990
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
               Caption         =   " Receipt Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ItemIssueReceiptVoucher.frx":241A
               Picture         =   "ItemIssueReceiptVoucher.frx":2436
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   8580
               TabIndex        =   55
               Top             =   675
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
               Caption         =   " Consignee"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ItemIssueReceiptVoucher.frx":2452
               Picture         =   "ItemIssueReceiptVoucher.frx":246E
            End
            Begin MSForms.ComboBox cmbChallanType 
               Height          =   330
               Left            =   9900
               TabIndex        =   8
               Top             =   990
               Width           =   3255
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "5741;582"
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
               Y1              =   7065
               Y2              =   7065
            End
            Begin MSForms.ComboBox cmbItemType 
               Height          =   330
               Left            =   5460
               TabIndex        =   7
               Top             =   990
               Width           =   3135
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "5530;582"
               ListWidth       =   2822
               MatchEntry      =   0
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Calibri"
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   13240
               Y1              =   575
               Y2              =   575
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   13240
               Y1              =   1725
               Y2              =   1725
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   8805
            TabIndex        =   49
            Top             =   8550
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
            Picture         =   "ItemIssueReceiptVoucher.frx":248A
            Picture         =   "ItemIssueReceiptVoucher.frx":24A6
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
            TabIndex        =   22
            Top             =   8550
            Width           =   615
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   19
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
Attribute VB_Name = "frmItemIssueReceiptVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Vch Type=NNNNFI/NNNNUI/NNNNFR/NNNNUR (F-Finished U-Unfinished I-Issue R-Receipt) & BOM=NNNNXXXXXXXXXXXXFI/NNNNXXXXXXXXXXXXMF (MF/ME/CF/MO/BN/BM) & 05-Purchase Challan 06-Purchase Return Challan 07-Sales Return Challan 08-Sales Challan
Public VchType As String 'R-Item Receipt I-Item Issue
Dim cnDeliveryChallan As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset, rstDeliveryCVList As New ADODB.Recordset, rstDeliveryCVParent As New ADODB.Recordset, rstDeliveryCVChild As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstTaxList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstNarrationList As New ADODB.Recordset, rstHSNCodeList As New ADODB.Recordset, rstOrderList As New ADODB.Recordset, rstVchSeriesList As New ADODB.Recordset, rstMaterialCentreList As New ADODB.Recordset
Dim PartyCode As String, ConsigneeCode As String, TaxCode As String, ItemCode As String, RefCode As String, NarrationCode As String, HSNCode As String, MaterialCentreCode As String, VchPrefix As String, TranType As String, VchNumbering As String, VchSeriesCode As String, AutoVchNo As String, oVchSeriesCode As String, oVchNo As String
Dim SortOrder, PrevStr, dblBookMark As Double, blnRecordExist As Boolean, EditMode As Boolean
Private Sub btnNotes_Click()
    frmNotes.NotesFlag = 7
    frmNotes.Label1.Caption = "Notes : Voucher No. : " & Text2.Text
    frmNotes.Show (vbModal)
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    Me.Caption = "Item " & IIf(VchType = "I", "Issued to", "Received from") & " Party Voucher"
    Mh3dLabel12.Caption = IIf(VchType = "I", " Mat Centre", " Mat Centre")
    Mh3dLabel14.Caption = IIf(VchType = "I", " Issue", " Receipt") & " Type"
    DataGrid1.Columns(3).Caption = IIf(VchType = "I", "Consignee Name", "Material Centre")
    TranType = IIf(VchType = "I", "08", "05")
    cnDeliveryChallan.CursorLocation = adUseClient
    cnDeliveryChallan.Open cnDatabase.ConnectionString
    LoadMasterList
    rstDeliveryCVList.Open "SELECT T.Code,LTRIM(T.Name) As Name,V.Name As VchSeriesName,Date,T.Type,M1.Name As PartyName,M2.Name As MaterialCentreName,ChallanNo,ChallanDate,Amount FROM (JobworkBVParent T INNER JOIN AccountMaster M1 ON T.Party=M1.Code) INNER JOIN AccountMaster M2 ON " & IIf(VchType = "I", "T.Consignee", "T.MaterialCentre") & "=M2.Code INNER JOIN VchSeriesMaster V ON T.VchSeries=V.Code WHERE LEFT(Type,2) IN (" & IIf(VchType = "R", "'05','07'", "'06','08'") & ") AND RIGHT(Type,1)='" & VchType & "' AND FYCode='" & FYCode & "' ORDER BY T.Name", cnDeliveryChallan, adOpenKeyset, adLockPessimistic
    rstDeliveryCVParent.CursorLocation = adUseClient
    rstDeliveryCVList.Filter = adFilterNone
    If rstDeliveryCVList.RecordCount > 0 Then rstDeliveryCVList.MoveLast
    Set DataGrid1.DataSource = rstDeliveryCVList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstDeliveryCVList.EOF Or rstDeliveryCVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstDeliveryCVList.ActiveConnection = Nothing
    cmbItemType.AddItem "Finished Goods", 0
    cmbItemType.AddItem "Unfinished Goods", 1
    cmbChallanType.AddItem IIf(VchType = "I", "Sale", "Purchase") & " Challan", 0
    cmbChallanType.AddItem IIf(VchType = "I", "Purchase Return", "Sale Return") & " Challan", 1
'    cmbChallanType.AddItem "Challan Reversal", 2
'    cmbChallanType.AddItem "To be Billed", 3
'    cmbChallanType.AddItem "Not to be Billed", 4
    SetButtonsForNoRecord
    fpSpread1.TextTip = TextTipFloating
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    With MdiMainMenu
        .mnuMaterialOutJobWork.Enabled = False: .mnuMaterialInJobWork.Enabled = False
    End With
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        EditMode = False
        If SSTab1.Tab = 0 Then  'List
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If InStr(1, "fpSpread1", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True
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
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF8)) And Toolbar1.Buttons.Item(3).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2)) And Toolbar1.Buttons.Item(4).Enabled Then 'Save
        EditMode = False
        If InStr(1, "fpSpread1", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True
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
    Call CloseRecordset(rstDeliveryCVList)
    Call CloseRecordset(rstDeliveryCVParent)
    Call CloseRecordset(rstDeliveryCVChild)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstTaxList)
    Call CloseRecordset(rstNarrationList)
    Call CloseRecordset(rstHSNCodeList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstOrderList)
    Call CloseConnection(cnDeliveryChallan)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuMaterialOutJobWork.Enabled = True
    MdiMainMenu.mnuMaterialInJobWork.Enabled = True
End Sub
Private Sub Text1_Change()
    On Error Resume Next
    With rstDeliveryCVList
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
    If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstDeliveryCVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstDeliveryCVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstDeliveryCVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstDeliveryCVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstDeliveryCVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstDeliveryCVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstDeliveryCVList
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
            If Not (rstDeliveryCVList.EOF Or rstDeliveryCVList.BOF) Then
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
        Text10.SetFocus
    End If
End Sub
Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text10.Text)
        If rstVchSeriesList.RecordCount = 0 Then DisplayError ("No Record in Voucher Series Master"): Text10.SetFocus: Exit Sub Else rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": VchSeriesCode = ""
        Call LoadSelectionList(rstVchSeriesList, "List of Voucher Series...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text10, VchSeriesCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(VchSeriesCode) <> "" Then Sendkeys "{TAB}" Else Text10.Text = ""
    End If
End Sub
Private Sub Text10_Validate(Cancel As Boolean)
    If CheckEmpty(Text10.Text, False) Then
        Cancel = True
    Else
        rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        VchNumbering = rstVchSeriesList.Fields("VchNumbering").Value
        If VchNumbering = "A" Then Text2.Locked = True Else Text2.Locked = False
        If Not blnRecordExist Then 'Vch-New
            If VchNumbering = "A" Then
                AutoVchNo = GenerateCode(cnDeliveryChallan, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & "F" + VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
            End If
        Else 'Vch-Old
            If VchSeriesCode = oVchSeriesCode Then
                Text2.Text = Text2.Text 'oVchNo
            Else
                If VchNumbering = "A" Then
                    AutoVchNo = GenerateCode(cnDeliveryChallan, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & "F" + VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                    Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                End If
            End If
        End If
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    Dim CellVal As Variant, i As Integer
    If Button.Index = 1 Then
        If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
        rstDeliveryCVParent.Open "SELECT * FROM JobworkBVParent WHERE Code=''", cnDeliveryChallan, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstDeliveryCVParent) Then
            Text2.Text = GenerateCode(cnDeliveryChallan, "SELECT MAX(CONVERT(INT,AutoVchNo)) FROM JobworkBVParent WHERE Right(Type,2)='" & "F" + VchType & "' AND FYCode='" & FYCode & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text10.SetFocus
            blnRecordExist = False
            cnDeliveryChallan.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnDeliveryChallan.BeginTrans
            With rstDeliveryCVChild
                If .State = adStateOpen Then
                    If .RecordCount > 0 Then .MoveFirst
                    Do While Not .EOF
                        If Not CheckEmpty(.Fields("VchCode").Value, False) Then Call UpdateStatus(.Fields("VchCode").Value, .Fields("Quantity").Value, "-")
                        .MoveNext
                    Loop
                End If
            End With
            cnDeliveryChallan.Execute "DELETE FROM JobworkBVParent WHERE Code='" & rstDeliveryCVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstDeliveryCVList.Delete
                rstDeliveryCVList.MoveNext
                If rstDeliveryCVList.RecordCount > 0 And rstDeliveryCVList.EOF Then rstDeliveryCVList.MoveLast
                cnDeliveryChallan.CommitTrans
                ShowProgressInStatusBar True
                Timer1.Enabled = True
                Text1.Text = ""
                rstDeliveryCVList.Filter = adFilterNone
            Else
                DisplayError (Err.Description)
                cnDeliveryChallan.RollbackTrans
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
        If UpdateRecord(rstDeliveryCVParent) Then
            If UpdateItemList("D", 0) Then
                UpdateFlag = 1
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 1, i
                        .GetText 4, i, CellVal
                        If Val(CellVal) <> 0 Then
                            If Not UpdateItemList("I", i) Then UpdateFlag = 0: Exit For
                        End If
                    Next
                End With
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnDeliveryChallan.CommitTrans
            If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
            rstDeliveryCVParent.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstDeliveryCVParent) Then
            cnDeliveryChallan.RollbackTrans
            If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
            rstDeliveryCVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstDeliveryCVList.Filter = adFilterNone
        rstDeliveryCVList.ActiveConnection = cnDeliveryChallan
        Do While Not RefreshRecord(rstDeliveryCVList): Loop
        Set DataGrid1.DataSource = rstDeliveryCVList
        rstDeliveryCVList.ActiveConnection = Nothing
        If rstDeliveryCVList.RecordCount > 0 Then rstDeliveryCVList.MoveLast
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem IIf(VchType = "P", "Material Centre", "Consignee"), 0
            .Combo1.AddItem "Party", 1
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
        DisplayMenu "P"
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
        DisplayMenu "S"
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstDeliveryCVList.RecordCount > 0 Then rstDeliveryCVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstDeliveryCVList.RecordCount > 0 Then
            rstDeliveryCVList.MovePrevious
            If rstDeliveryCVList.BOF Then rstDeliveryCVList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstDeliveryCVList.RecordCount > 0 Then
            rstDeliveryCVList.MoveNext
            If rstDeliveryCVList.EOF Then rstDeliveryCVList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstDeliveryCVList.RecordCount > 0 Then rstDeliveryCVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstDeliveryCVList.EOF Or rstDeliveryCVList.BOF) Then
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
        rstDeliveryCVList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstDeliveryCVList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstDeliveryCVList.EOF Or rstDeliveryCVList.BOF) Then
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
    If rstDeliveryCVList.RecordCount = 0 Then
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
Private Sub Text2_Validate(Cancel As Boolean)
    If rstDeliveryCVParent.EOF Or rstDeliveryCVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnDeliveryChallan, "JobworkBVParent", "Code", "[Name]+RIGHT([Type],1)", Trim(Text2.Text) & VchType, rstDeliveryCVParent.Fields("Code").Value, False, FYCode) Then
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
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = PartyCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        PartyCode = slCode: Text3.Text = slName
        If Not CheckEmpty(PartyCode, False) Then LoadMasterList: Sendkeys "{TAB}"
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
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = IIf(VchType = "I", "*99999", "*99999")
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
    If rstDeliveryCVList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
    rstDeliveryCVParent.Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstDeliveryCVList.Fields("Code").Value) & "'", cnDeliveryChallan, adOpenKeyset, adLockOptimistic
    If rstDeliveryCVParent.RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
End Sub
Private Sub ClearFields()
    Text2.Text = "" 'Vch No.
    Text9.Text = "" 'Challan No
    MhDateInput2.Text = "  -  -    " 'Challan Date
    MhRealInput13.Value = 0 'Box
    Text3.Text = "": PartyCode = "" 'Party Name
    Text7.Text = "": MaterialCentreCode = "" 'Material Centre Name
    Text5.Text = "": TaxCode = "" 'Tax Name
    Text4.Text = "" 'Remarks
    cmbItemType.ListIndex = 0: cmbItemType.Enabled = True
    cmbChallanType.ListIndex = 0: cmbChallanType.Enabled = True: cmbChallanType_Click
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhRealInput1.Value = 0
    MhRealInput2.Value = 0
    MhRealInput3.Value = 0
    MhRealInput4.Value = 0
    MhRealInput5.Value = 0
    MhRealInput6.Value = 0
    MhRealInput12.Value = 0
    MhRealInput7.Value = 0
    MhRealInput8.Value = 0
    MhRealInput9.Value = 0
    MhRealInput10.Value = 0
    MhRealInput11.Value = 0
    Text10.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    PartyCode = "": ConsigneeCode = "": MaterialCentreCode = "": TaxCode = "": VchSeriesCode = "": oVchSeriesCode = "": oVchNo = "": AutoVchNo = ""
End Sub
Private Sub LoadFields()
    With rstDeliveryCVParent
        If .EOF Or .BOF Then Exit Sub
        Text2.Text = .Fields("Name").Value
        MhDateInput1.Text = Format(.Fields("Date").Value, "dd-MM-yyyy")
        Text9.Text = CheckNull(.Fields("ChallanNo").Value)
        If Not IsNull(.Fields("ChallanDate").Value) Then MhDateInput2.Text = Format(.Fields("ChallanDate").Value, "dd-MM-yyyy")
        MhRealInput13.Value = Val(.Fields("Box").Value)
       
       PartyCode = .Fields("Party").Value
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        rstAccountList.Find "[Code] = '" & PartyCode & "'"
        If Not rstAccountList.EOF Then Text3.Text = rstAccountList.Fields("Col0").Value
        
'        If VchType = "R" Then MaterialCentreCode = .Fields("MaterialCentre").Value Else MaterialCentreCode = .Fields("Consignee").Value
'        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
'        rstAccountList.Find "[Code] = '" & MaterialCentreCode & "'"
'        If Not rstAccountList.EOF Then Text7.Text = rstAccountList.Fields("Col0").Value
        
        MaterialCentreCode = .Fields("MaterialCentre").Value
        If rstMaterialCentreList.RecordCount > 0 Then rstMaterialCentreList.MoveFirst
        rstMaterialCentreList.Find "[Code] = '" & MaterialCentreCode & "'"
        If Not rstMaterialCentreList.EOF Then Text7.Text = rstMaterialCentreList.Fields("Col0").Value
        
        ConsigneeCode = .Fields("Consignee").Value
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        rstAccountList.Find "[Code] = '" & ConsigneeCode & "'"
        If Not rstAccountList.EOF Then Text8.Text = rstAccountList.Fields("Col0").Value
        
        TaxCode = .Fields("Tax").Value
        If rstTaxList.RecordCount > 0 Then rstTaxList.MoveFirst
        rstTaxList.Find "[Code] = '" & TaxCode & "'"
        If Not rstTaxList.EOF Then Text5.Text = rstTaxList.Fields("Col0").Value
        VchSeriesCode = .Fields("VchSeries").Value: oVchSeriesCode = VchSeriesCode
        If rstVchSeriesList.RecordCount > 0 Then rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        If Not rstVchSeriesList.EOF Then Text10.Text = rstVchSeriesList.Fields("Col0").Value
        Text4.Text = .Fields("Remarks").Value
        cmbItemType.ListIndex = IIf(Mid(.Fields("Type").Value, 5, 1) = "F", 0, 1)
        cmbChallanType.ListIndex = IIf(VchType = "I", IIf(Left(.Fields("Type").Value, 2) = "08", 0, IIf(Left(.Fields("Type").Value, 2) = "06", 1, IIf(Left(.Fields("Type").Value, 2) = "11", 2, IIf(Left(.Fields("Type").Value, 2) = "13", 3, 4)))), IIf(Left(.Fields("Type").Value, 2) = "05", 0, IIf(Left(.Fields("Type").Value, 2) = "07", 1, IIf(Left(.Fields("Type").Value, 2) = "12", 2, IIf(Left(.Fields("Type").Value, 2) = "14", 3, 4)))))
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
        Call LoadItemList(.Fields("Code").Value)
    End With
    CalculateTotal
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstDeliveryCVParent.RecordCount = 0 Then Exit Sub
    If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
    rstDeliveryCVParent.CursorLocation = adUseServer
    rstDeliveryCVParent.Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstDeliveryCVList.Fields("Code").Value) & "'", cnDeliveryChallan, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstDeliveryCVParent.Fields("RecordStatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    cmbItemType.Enabled = False
    cmbChallanType.Enabled = False
    Text10.SetFocus
    blnRecordExist = True
    cnDeliveryChallan.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    With rstDeliveryCVParent
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnDeliveryChallan, "SELECT MAX(Code) FROM JobworkBVParent", 6, "0")
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
        .Fields("Consignee").Value = ConsigneeCode
        .Fields("MaterialCentre").Value = MaterialCentreCode
        .Fields("Tax").Value = TaxCode
        .Fields("Remarks").Value = Trim(Text4.Text)
        .Fields("Rebate%").Value = MhRealInput4.Value
        .Fields("Rebate").Value = MhRealInput5.Value
        .Fields("Freight").Value = MhRealInput6.Value
        .Fields("Adjustment").Value = MhRealInput12.Value
        .Fields("TaxableAmount").Value = MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput12.Value
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
        .Fields("Type").Value = VchPrefix & IIf(cmbItemType.ListIndex = 0, "F", "U") & VchType
        .Fields("FYCode").Value = FYCode
        .Fields("RecordStatus").Value = "N"
        .Fields("Notes").Value = txtNotes.Text
        .Fields("SalesType").Value = ""
    End With
End Sub
Private Sub AddToList()
    On Error Resume Next
    With rstDeliveryCVList
        .MoveFirst
        .Find "[Code] = '" & rstDeliveryCVParent.Fields("Code").Value & "'"
        If .EOF Then .AddNew
        .Fields("VchSeriesName").Value = Text10.Text
        .Fields("Code").Value = rstDeliveryCVParent.Fields("Code").Value
        .Fields("Name").Value = rstDeliveryCVParent.Fields("Name").Value
        .Fields("Date").Value = rstDeliveryCVParent.Fields("Date").Value
        .Fields("PartyName").Value = Trim(Text3.Text)
        .Fields("MaterialCentreName").Value = Trim(Text7.Text)
        .Fields("Type").Value = rstDeliveryCVParent.Fields("Type").Value
        .Fields("Amount").Value = MhRealInput11.Value
        .Fields("ChallanNo").Value = rstDeliveryCVParent.Fields("ChallanNo").Value
        .Fields("ChallanDate").Value = rstDeliveryCVParent.Fields("ChallanDate").Value
        .Update
        .Sort = SortOrder & " Asc"
        .Find "[Code] = '" & rstDeliveryCVParent.Fields("Code").Value & "'"
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        DisplayError ("Voucher No. cannot be blank"): Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnDeliveryChallan, "JobworkBVParent", "Code", "[Name]+RIGHT(Type,1)", Trim(Text2.Text) & VchType, rstDeliveryCVParent.Fields("Code").Value, False, FYCode) Then
        If Not blnRecordExist Then
            Dim VchNo As String
            VchNo = GenerateCode(cnDeliveryChallan, "SELECT MAX(CONVERT(INT,Name)) FROM JobworkBVParent WHERE RIGHT(Type,1)='" & VchType & "' AND FYCode='" & FYCode & "'", 10, Space(1))
            If Trim(VchNo) <> Trim(Text2.Text) Then DisplayError ("Vch No. changed from " & Trim(Text2.Text) & " to " & Trim(VchNo))
            Text2.Text = VchNo: Exit Function
        Else
            Text2.SetFocus: CheckMandatoryFields = True: Exit Function
        End If
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text7.Text, False) Then
        Text7.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text5.Text, False) Then
        Text5.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf fpSpread1.DataRowCnt = 0 Then
        DisplayError ("Blank Voucher cannot be saved"): fpSpread1.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf fpSpread1.DataRowCnt > 0 Then
        Dim i As Integer, CellVal As Variant
        With fpSpread1
                For i = 1 To .DataRowCnt
                    .GetText 7, i, CellVal
                    If CheckEmpty(CellVal, False) Then DisplayError ("Narration at row #" & Trim(Str(i)) & " is blank"): CheckMandatoryFields = True: .SetFocus: Exit Function
                Next
        End With
    End If
End Function
Private Sub LoadItemList(ByVal VchNo As String)
    Dim i As Integer, SQL As String
    On Error GoTo ErrorHandler
    If rstDeliveryCVChild.State = adStateOpen Then rstDeliveryCVChild.Close
    If cmbChallanType.ListIndex < 2 Then
        If cmbItemType.ListIndex = 0 Then 'Finished with ref
            SQL = "SELECT I.Code As ItemCode,I.Name As ItemName,H.Code As HSNCode,H.Name As HSNName,T.Quantity,T.Quantity+R.EstQty01-R.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Code As NarrationCode,N.Name As NarrationName,SrNo,CASE WHEN Ref IS NULL THEN '' ELSE R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) END As VchCode,CASE WHEN Ref IS NULL THEN '' ELSE LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) END As VchNo FROM (((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,14)=R.Code+'XXXXXXXXXXXXFI') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "'"
        Else 'Unfinished with Ref
'           SQL = "SELECT I.Code As ItemCode,I.Name+'_'+E.Name+'_Printing' As ItemName,H.Code As HSNCode,H.Name As HSNName,T.Quantity,T.Quantity+C.ActualQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild05 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+'XXXXXXMF') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE T.Code='" & VchNo & "'"
            SQL = "SELECT I.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.ActualQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,14)=R.Code+'XXXXXXXXXXXXMF') INNER JOIN BookPOChild05 C ON R.Code=C.Code) INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "'"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.ActualQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild06 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+'XXXXXXME') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code  WHERE T.Code='" & VchNo & "'"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.ActualQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild0901 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Book+'XXXXXXCF') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN BookMaster E ON C.Book=E.Code WHERE T.Code='" & VchNo & "'"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+E.Name+'_'+O.Name As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.Quantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+E.Code+O.Code+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild07 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+C.Operation+'MO') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code WHERE T.Code='" & VchNo & "'"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_Binding' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.ActualQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,14)=R.Code+'XXXXXXXXXXXXBN') INNER JOIN BookPOChild08 C ON R.Code=C.Code) INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "'"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+CASE WHEN C.Category='1' THEN O.Name WHEN C.Category='2' THEN P.Name ELSE U.Name END As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.OrderQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+C.Item+'XXXXX'+C.Category+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild0801 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Item+'XXXXX'+C.Category+'BM') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) LEFT JOIN OutsourceItemMaster O ON C.Category+C.Item='1'+O.Code) LEFT JOIN PaperMaster P ON C.Category+C.Item='2'+P.Code) LEFT JOIN BookMaster U ON C.Category+C.Item='3'+U.Code WHERE T.Code='" & VchNo & "'"
        End If
    Else 'Finished without ref
        SQL = "SELECT I.Code As ItemCode,I.Name As ItemName,H.Code As HSNCode,H.Name As HSNName,T.Quantity,T.Quantity+R.EstQty01-R.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Code As NarrationCode,N.Name As NarrationName,SrNo,CASE WHEN Ref IS NULL THEN '' ELSE R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) END As VchCode,CASE WHEN Ref IS NULL THEN '' ELSE LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) END As VchNo FROM (((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) LEFT JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,12)=R.Code+'XXXXXXXXXXXX') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "' AND RIGHT(T.BOM,2)='FI'"
    End If
    SQL = SQL + " ORDER BY SrNo"
    rstDeliveryCVChild.Open SQL, cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstDeliveryCVChild.ActiveConnection = Nothing
    If rstDeliveryCVChild.RecordCount > 0 Then rstDeliveryCVChild.MoveFirst
    i = 0
    Do Until rstDeliveryCVChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstDeliveryCVChild.Fields("ItemName").Value
            .SetText 2, i, rstDeliveryCVChild.Fields("HSNName").Value
            .SetText 3, i, rstDeliveryCVChild.Fields("VchNo").Value
            .SetText 4, i, Val(rstDeliveryCVChild.Fields("Quantity").Value)
            .SetText 5, i, Val(rstDeliveryCVChild.Fields("Rate").Value)
            .SetText 6, i, Val(rstDeliveryCVChild.Fields("Amount").Value)
            .SetText 7, i, rstDeliveryCVChild.Fields("NarrationName").Value
            .SetText 8, i, rstDeliveryCVChild.Fields("NarrationCode").Value
            .SetText 9, i, rstDeliveryCVChild.Fields("VchCode").Value
            .SetText 10, i, rstDeliveryCVChild.Fields("ItemCode").Value
            .SetText 11, i, rstDeliveryCVChild.Fields("HSNCode").Value
            .SetText 12, i, Val(CheckNull(rstDeliveryCVChild.Fields("PendingQty").Value))
        End With
        rstDeliveryCVChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Item List")
End Sub
Private Function UpdateItemList(ByVal ActionType As String, ByVal SrNo As Integer) As Boolean
    Dim CellVal(1 To 7) As Variant, BOM As String
    On Error GoTo ErrorHandler
    UpdateItemList = True
    If ActionType = "D" Then
        If Not blnRecordExist Then Exit Function
        With rstDeliveryCVChild
            If .State = adStateOpen Then
                If .RecordCount > 0 Then .MoveFirst
                Do While Not .EOF
                    If Not CheckEmpty(.Fields("VchCode").Value, False) Then Call UpdateStatus(.Fields("VchCode").Value, .Fields("Quantity").Value, "-")
                    .MoveNext
                Loop
            End If
        End With
        cnDeliveryChallan.Execute "DELETE FROM JobworkBVChild WHERE Code='" & rstDeliveryCVParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 4, .ActiveRow, CellVal(1) 'Qnty
            .GetText 5, .ActiveRow, CellVal(2)  'Rate
            .GetText 6, .ActiveRow, CellVal(3)  'Amnt
            .GetText 8, .ActiveRow, CellVal(4)  'Narration Code
            .GetText 9, .ActiveRow, CellVal(5)  'VchCode=SOCode+Element+Operation+ItemType for Sales/Purchase Challan & Null for Others
            .GetText 10, .ActiveRow, CellVal(6)  'Item Code
            .GetText 11, .ActiveRow, CellVal(7)  'HSN Code
        End With
        BOM = VchPrefix + IIf(CheckEmpty(CellVal(5), False), "XXXXXXXXXXXXFI", Right(CellVal(5), 14)) 'BOM='0410'+Element+Operation+ItemType
        cnDeliveryChallan.Execute "INSERT INTO JobworkBVChild VALUES ('" & rstDeliveryCVParent.Fields("Code").Value & "','" & Left(CellVal(5), 6) & "','" & BOM & "','" & CellVal(6) & "','" & CellVal(7) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ",'" & CellVal(4) & "'," & SrNo & ",'','','','','',0,'XXXXXX')"
        If Not CheckEmpty(CellVal(5), False) Then Call UpdateStatus(CellVal(5), Val(CellVal(1)), "+")
    End If
    Exit Function
ErrorHandler:
    UpdateItemList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Party" Then
        rstDeliveryCVList.Filter = "[PartyName] Like '%" & SrchText & "%'"
    ElseIf SrchFor = IIf(VchType = "P", "Material Centre", "Consignee") Then
        rstDeliveryCVList.Filter = "[MaterialCentreName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cVal As Variant
    With fpSpread1
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus: CalculateTotal
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then 'Item
                If cmbChallanType.ListIndex < 2 And fpSpread1.ActiveCol = 1 Then Exit Sub
                .GetText 1, .ActiveRow, cVal 'Item
                Text6.Text = FixQuote(cVal)
                If rstItemList.RecordCount = 0 Then DisplayError ("No record in Item Master"): .SetActiveCell 1, .ActiveRow: .SetFocus: Exit Sub Else rstItemList.MoveFirst
                rstItemList.Find "[Col0] = '" & FixQuote(Trim(cVal)) & "'"
                SelectionType = "S": ItemCode = ""
                Call LoadSelectionList(rstItemList, "List of Items...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text6, ItemCode)
                Call CloseForm(FrmSelectionList)
                If CheckEmpty(ItemCode, False) Then
                    .SetActiveCell 1, .ActiveRow
                Else
                    rstItemList.MoveFirst: rstItemList.Find "[Code] ='" & ItemCode & "'"
                    .SetText 1, .ActiveRow, Text6.Text 'Item Name
                    .SetText 10, .ActiveRow, ItemCode
                    .SetText 5, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                    .GetText 11, .ActiveRow, cVal 'HSN
                    If CheckEmpty(cVal, False) Then .SetText 2, .ActiveRow, rstItemList.Fields("HSNName").Value: .SetText 11, .ActiveRow, rstItemList.Fields("HSNCode").Value
                    .SetFocus
                    Sendkeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 2 Then
                .GetText 11, .ActiveRow, cVal 'HSN Code
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "18"
                FrmGeneralMaster.MasterCode = cVal
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 2, .ActiveRow, slName: .SetText 11, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then LoadMasterList: Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 7 Then
                .GetText 8, .ActiveRow, cVal 'Short Narration
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "17"
                FrmGeneralMaster.MasterCode = cVal
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 7, .ActiveRow, slName: .SetText 8, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then LoadMasterList: Sendkeys "{ENTER}"
            End If
        ElseIf KeyCode = vbKeyF11 Then
            If .DataRowCnt = 0 And cmbChallanType.ListIndex < 2 Then LoadOrderList
        End If
        If .DataRowCnt > 0 Then cmbItemType.Enabled = False: cmbChallanType.Enabled = False Else cmbItemType.Enabled = True: cmbChallanType.Enabled = True
    End With
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Item As Variant, Qty As Variant, Rate As Variant
    With fpSpread1
        If Col = 4 Or Col = 5 Then
            .GetText 1, Row, Item
            .GetText 4, Row, Qty
            .GetText 5, Row, Rate
            If Not CheckEmpty(Item, False) Then .SetText 6, Row, Qty * Rate: CalculateTotal Else .SetText 4, Row, "": .SetText 5, Row, "": .SetText 6, Row, ""
        End If
    End With
End Sub
Private Sub fpSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As FPSpreadADO.TextTipFetchMultilineConstants, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim PendingQty As Variant
    fpSpread1.GetText 12, Row, PendingQty
    If Val(PendingQty) = 0 Then Exit Sub
    If Col = 4 Then
        fpSpread1.SetTextTipAppearance "Calibri", 10, False, False, &HC0FFFF, &H80000008
        TipText = "Pending : " & Trim(PendingQty)
        ShowTip = True
    End If
End Sub
Private Sub CalculateTotal()
    Dim i As Integer, Qty As Variant, Amt As Variant
    MhRealInput1.Value = 0: MhRealInput2.Value = 0
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 4, i, Qty
            .GetText 6, i, Amt
            MhRealInput1.Value = MhRealInput1.Value + Val(Qty)
            MhRealInput2.Value = MhRealInput2.Value + Val(Amt)
        Next
        MhRealInput3.Value = MhRealInput2.Value
        MhRealInput5.Value = (MhRealInput3.Value * MhRealInput4.Value) / 100
        MhRealInput8.Value = ((MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput12.Value) * MhRealInput7.Value) / 100
        MhRealInput10.Value = ((MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput12.Value) * MhRealInput9.Value) / 100
        MhRealInput11.Value = Round(MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput8.Value + MhRealInput10.Value + MhRealInput12.Value, 0)
    End With
End Sub
Private Sub DisplayMenu(ByVal OutputTo As String)   'Original/Duplicate/Triplicate
    Dim menusel As String
    If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
    menusel = DisplayPopupMenu(Me.hwnd, 2)
    Call PrintItemIRVch(rstDeliveryCVList.Fields("Code").Value, rstDeliveryCVList.Fields("Type").Value, Choose(menusel, "O", "D", "T"), OutputTo) 'Original/Duplicate/Triplicate
    If Not (rstDeliveryCVList.EOF Or rstDeliveryCVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.SetFocus
End Sub
Private Sub UpdateStatus(ByVal VchCode As String, ByVal Quantity As Long, ByVal Operation As String)
    If InStr(1, "FI", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOParent SET DeliveredQuantityC=DeliveredQuantityC" + Operation + Trim(Quantity) + " WHERE Code+'XXXXXXXXXXXXFI'='" + VchCode + "'"
    End If
    If InStr(1, "FI_MF", Right(VchCode, 2)) > 0 Then
'       cnDeliveryChallan.Execute "UPDATE BookPOChild05 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+Element+'XXXXXXMF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        cnDeliveryChallan.Execute "UPDATE BookPOChild05 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+'XXXXXXXXXXXXMF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
    If InStr(1, "FI_ME", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOChild06 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+Element+'XXXXXXME'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
    If InStr(1, "FI_CF", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOChild0901 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+Book+'XXXXXXCF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
    If InStr(1, "FI_MO", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOChild07 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+Element+Operation+'MO'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
    If InStr(1, "FI_BN", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOChild08 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+'XXXXXXXXXXXXBN'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
    If InStr(1, "FI_BM", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOChild0801 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+Item+'XXXXX'+Category+'BM'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
End Sub
Private Sub LoadOrderList()
    Dim SQL As String
    If rstOrderList.State = adStateOpen Then rstOrderList.Close
    If cmbItemType.ListIndex = 0 Then 'Finished Item
        SQL = "SELECT DISTINCT P.Code+'XXXXXXXXXXXXFI' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FI' As VchNo,P.Date As VchDate,I.Name As Item,P.EstQty01 As OrderedQty,P.DeliveredQuantityC As ChallanQty,P.DeliveredQuantityB As DirectQty FROM (BookPOParent P INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN BookPOChild0801 C ON P.Code=C.Code WHERE (P.BookPrinter='" & PartyCode & "' OR P.TitlePrinter='" & PartyCode & "' OR P.Laminator='" & PartyCode & "' OR P.Binder='" & PartyCode & "' OR C.Vendor='" & PartyCode & "') AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND P.EstQty01-P.DeliveredQuantityB-P.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)<>'FI' AND LEFT(BOM,2)='" & TranType & "') ORDER BY I.Name,P.Date,VchNo"
    Else 'Unfinished Item
'       SQL = "SELECT P.Code+E.Code+'XXXXXXMF' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.BookPrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL "
        SQL = "SELECT P.Code+'XXXXXXXXXXXXMF' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As VchNo,P.Date As VchDate,I.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.BookPrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL " & _
                    "SELECT P.Code+E.Code+'XXXXXXME' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/ME' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL " & _
                    "SELECT P.Code+E.Code+'XXXXXXCF' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CF' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN BookMaster E ON C.Book=E.Code WHERE P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL " & _
                    "SELECT P.Code+E.Code+O.Code+'MO' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MO' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_'+O.Name As Item,C.Quantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM (((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code WHERE P.Laminator='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.Quantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL " & _
                    "SELECT P.Code+'XXXXXXXXXXXXBN' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/BN' As VchNo,P.Date As VchDate,I.Name+'_Binding' As Item,C.ActualQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Binder='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL " & _
                    "SELECT P.Code+C.Item+'XXXXX'+C.Category+'BM' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/BM' As VchNo,P.Date As VchDate,I.Name+'_'+CASE WHEN C.Category='1' THEN O.Name WHEN C.Category='2' THEN R.Name ELSE U.Name END As Item,C.OrderQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((((BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN OutsourceItemMaster O ON C.Category+C.Item='1'+O.Code) LEFT JOIN PaperMaster R ON C.Category+C.Item='2'+R.Code) LEFT JOIN BookMaster U ON C.Category+C.Item='3'+U.Code WHERE C.Vendor='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.OrderQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "')" & _
                    " ORDER BY Item,VchDate,VchNo"
    End If
    rstOrderList.Open SQL, cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstOrderList.ActiveConnection = Nothing
    If rstOrderList.RecordCount = 0 Then DisplayError ("No Pending Order Exists"): fpSpread1.SetFocus: Exit Sub
    With FrmOrderList.fpSpread1
        .Row = SpreadHeader + 1
        .Col = 5: .Text = "Dlvrd-Bill"
        .Col = 6: .Text = "Dlvrd-Challan"
    End With
    Load FrmOrderList
    FrmOrderList.Text2 = Text3.Text
    Dim i As Integer, Delivered As Long, UnitRate As Double
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
            FrmOrderList.fpSpread1.SetText 2, i, .Fields("VchNo").Value: FrmOrderList.fpSpread1.SetText 10, i, .Fields("VchCode").Value
            FrmOrderList.fpSpread1.SetText 3, i, Format(.Fields("VchDate").Value, "dd-MMM-yy")
            FrmOrderList.fpSpread1.SetText 4, i, Val(.Fields("OrderedQty").Value) 'Ordered
            FrmOrderList.fpSpread1.SetText 5, i, Val(.Fields("DirectQty").Value) 'Delivered-Bill
            FrmOrderList.fpSpread1.SetText 6, i, Val(.Fields("ChallanQty").Value) 'Delivered-Challan
            FrmOrderList.fpSpread1.SetText 7, i, Val(.Fields("OrderedQty").Value) - Val(.Fields("DirectQty").Value) - Val(.Fields("ChallanQty").Value) 'Pending
            Delivered = Val(.Fields("ChallanQty").Value) + Val(.Fields("DirectQty").Value)
            FrmOrderList.fpSpread1.SetText 8, i, IIf(Delivered = 0, "Undelivered", IIf(Delivered < Val(.Fields("OrderedQty").Value), "Under Delivery", "Delivered"))
            FrmOrderList.fpSpread1.SetText 9, i, 0
            .MoveNext
        Loop
        FrmOrderList.fpSpread1.SetActiveCell 9, 1
    End With
    FrmOrderList.Check2 = 0
    FrmOrderList.Check1.Visible = False
    CenterForm FrmOrderList
    FrmOrderList.Show vbModal
    If Not CheckEmpty(FrmOrderList.VchCodeList, False) Then
        If rstOrderList.State = adStateOpen Then rstOrderList.Close
        If cmbItemType.ListIndex = 0 Then    'Finished Item
            SQL = "SELECT I.Code As ItemCode,I.Name As ItemName,P.UnitRate,100 As ProfitMargin,H.Code As HSNCode,H.Name As HSNName,P.EstQty01-P.DeliveredQuantityB-P.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FI' As VchNo,P.Code+'XXXXXXXXXXXXFI' As VchCode FROM (BookPOParent P INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+'XXXXXXXXXXXXFI' IN (" & FrmOrderList.VchCodeList & ") ORDER BY I.Name,VchNo"
        Else 'Unfinished Item
'           SQL = "SELECT I.Code As ItemCode,I.Name+'_'+E.Name+'_Printing' As ItemName,ROUND((C.PrintAmount+C.Adjustment+C.PlateAmount+C.PAdjustment+C.PaperAmount+C.RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As VchNo,P.Code+E.Code+'XXXXXXMF' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.Code+E.Code+'XXXXXXMF' IN (" & FrmOrderList.VchCodeList & ") UNION ALL "
            SQL = "SELECT I.Code As ItemCode,I.Name+'_Printing' As ItemName,ROUND((C.PrintAmount1+C.PrintAmount2+C.PrintAmount4+C.Adjustment+C.PlateAmount1+C.PlateAmount2+C.PlateAmount4+C.PAdjustment+C.PaperAmount1+C.PaperAmount2+C.PaperAmount4+C.RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As VchNo,P.Code+'XXXXXXXXXXXXMF' As VchCode FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+'XXXXXXXXXXXXMF' IN (" & FrmOrderList.VchCodeList & ") UNION ALL " & _
                        "SELECT I.Code As ItemCode,I.Name+'_'+E.Name+'_Printing' As ItemName,ROUND((C.PrintAmount+C.Adjustment+C.PlateAmount+C.PAdjustment+C.PaperAmount+C.RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/ME' As VchNo,P.Code+E.Code+'XXXXXXME' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.Code+E.Code+'XXXXXXME' IN (" & FrmOrderList.VchCodeList & ") UNION ALL " & _
                        "SELECT I.Code As ItemCode,I.Name+'_'+E.Name+'_Printing' As ItemName,ROUND(((C1.PrintAmount+C1.Adjustment+C1.PlateAmount+C1.PAdjustment+C1.PaperAmount+C1.RAdjustment)/(SELECT SUM(ActualQuantity) FROM BookPOChild0901 WHERE Code=P.Code))*C.ActualQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CF' As VchNo,P.Code+E.Code+'XXXXXXCF' As VchCode FROM ((((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C ON C.Code=C1.Code) INNER JOIN BookMaster I ON  P.Book=I.Code) INNER JOIN BookMaster E ON C.Book=E.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+'XXXXXXCF' IN (" & FrmOrderList.VchCodeList & ") UNION ALL " & _
                        "SELECT I.Code As ItemCode,I.Name+'_'+E.Name+'_'+O.Name As ItemName,ROUND((C.Amount+C.Adjustment)/C.Quantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.Quantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MO' As VchNo,P.Code+E.Code+O.Code+'MO' As VchCode FROM ((((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+O.Code+'MO' IN (" & FrmOrderList.VchCodeList & ") UNION ALL " & _
                        "SELECT I.Code As ItemCode,I.Name+'_Binding' As ItemName,ROUND((C.BillAmount-C.VAT)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/BN' As VchNo,P.Code+'XXXXXXXXXXXXBN' As VchCode FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+'XXXXXXXXXXXXBN' IN (" & FrmOrderList.VchCodeList & ") UNION ALL " & _
                        "SELECT I.Code As ItemCode,I.Name+'_'+CASE WHEN C.Category='1' THEN O.Name WHEN C.Category='2' THEN R.Name ELSE U.Name END As ItemName,ROUND(C.Amount/C.OrderQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.OrderQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/BM' As VchNo,P.Code+C.Item+'XXXXX'+C.Category+'BM' As VchCode FROM (((((BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN OutsourceItemMaster O ON C.Category+C.Item='1'+O.Code) LEFT JOIN PaperMaster R ON C.Category+C.Item='2'+R.Code) LEFT JOIN BookMaster U ON C.Category+C.Item='3'+U.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+C.Item+'XXXXX'+C.Category+'BM' IN (" & FrmOrderList.VchCodeList & ") " & _
                        "ORDER BY ItemName,VchNo"
        End If
        rstOrderList.Open SQL, cnDeliveryChallan, adOpenKeyset, adLockReadOnly
        If rstOrderList.RecordCount > 0 Then
            i = 0
            With fpSpread1
                Do While Not rstOrderList.EOF
                    i = i + 1
                    .SetText 1, i, rstOrderList.Fields("ItemName").Value
                    .SetText 2, i, rstOrderList.Fields("HSNName").Value: .SetText 11, i, rstOrderList.Fields("HSNCode").Value
                    .SetText 3, i, rstOrderList.Fields("VchNo").Value
                    .SetText 4, i, Val(rstOrderList.Fields("BalQty").Value)
                    UnitRate = Val(rstOrderList.Fields("UnitRate").Value) + (Val(rstOrderList.Fields("UnitRate").Value) * Val(rstOrderList.Fields("ProfitMargin").Value)) / 100
                    .SetText 5, i, Round(UnitRate, 3)
                    .SetText 6, i, Val(rstOrderList.Fields("BalQty").Value) * Round(UnitRate, 3) 'quantity * rate
                    .SetText 7, i, Mid(rstOrderList.Fields("Narration").Value, InStr(1, rstOrderList.Fields("Narration").Value, "-") + 1, 40)
                    .SetText 8, i, Left(rstOrderList.Fields("Narration").Value, InStr(1, rstOrderList.Fields("Narration").Value, "-") - 1)
                    .SetText 9, i, rstOrderList.Fields("VchCode").Value
                    .SetText 10, i, rstOrderList.Fields("ItemCode").Value
                    rstOrderList.MoveNext
                Loop
                Call CalculateTotal
            End With
        End If
    End If
    CloseForm FrmOrderList
End Sub
Private Sub cmbChallanType_Click()
    VchPrefix = IIf(VchType = "I", Choose(cmbChallanType.ListIndex + 1, "08", "06", "12", "14", "16"), Choose(cmbChallanType.ListIndex + 1, "05", "07", "11", "13", "15")) & IIf(cmbChallanType.ListIndex < 2, "10", "00") 'Challan Reversal for To/Not To be billed only
End Sub
Private Sub LoadMasterList()
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    If rstMaterialCentreList.State = adStateOpen Then rstMaterialCentreList.Close
    rstMaterialCentreList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='*99999' ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstMaterialCentreList.ActiveConnection = Nothing
    If rstTaxList.State = adStateOpen Then rstTaxList.Close
    rstTaxList.Open "SELECT Name As Col0,[IGST%],[SGST%],[CGST%],Code FROM TaxMaster ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstTaxList.ActiveConnection = Nothing
    If rstNarrationList.State = adStateOpen Then rstNarrationList.Close
    rstNarrationList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='17' ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstNarrationList.ActiveConnection = Nothing
    If rstHSNCodeList.State = adStateOpen Then rstHSNCodeList.Close
    rstHSNCodeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='18' ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstHSNCodeList.ActiveConnection = Nothing
    If rstItemList.State = adStateOpen Then rstItemList.Close
    rstItemList.Open "SELECT I.Name As Col0,I.Price,I.Code,H.Code As HSNCode,H.Name As HSNName FROM BookMaster I INNER JOIN GeneralMaster H ON I.HSNCode=H.Code ORDER BY I.Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstItemList.ActiveConnection = Nothing
    If rstVchSeriesList.State = adStateOpen Then rstVchSeriesList.Close
    rstVchSeriesList.Open "SELECT Name As Col0,Prefix,Suffix,VchNumbering,Code FROM VchSeriesMaster WHERE VchType='" & IIf("F" & VchType = "FR", "05", IIf("F" & VchType = "FI", "08", IIf("F" & VchType = "FR", "07", "06"))) & "F" & VchType & "' ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstVchSeriesList.ActiveConnection = Nothing
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Public Sub PrintItemIRVch(ByVal VchCode As String, ByVal VchType As String, ByVal BillType As String, Optional ByVal OutputType As String)
Dim ChallanType As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
    rstDeliveryCVParent.Open "SELECT TYPE FROM JobworkBVParent WHERE Code='" + Left(VchCode, 6) + "' ", cnDeliveryChallan, adOpenKeyset, adLockOptimistic
    If rstDeliveryCVParent.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    ChallanType = (rstDeliveryCVParent.Fields("TYPE").Value)
    rstDeliveryCVParent.ActiveConnection = Nothing
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,EMail,Website,GSTIN,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,Prefix,Suffix FROM CompanyMaster P INNER JOIN CompChild C ON P.Code=C.Code WHERE VchType= " & IIf(ChallanType = "0510FR", 5, IIf(ChallanType = "0710FR", 7, IIf(ChallanType = "0610FI", 6, IIf(ChallanType = "0810FI", 8, 0)))), cnDeliveryChallan, adOpenKeyset, adLockOptimistic
    rstDeliveryCVChild.Open "SELECT LTRIM(P1.Name) As BillNo,P1.Date As BillDate,A.PrintName As Party,A.Address1 As PartyAddress1,A.Address2 As PartyAddress2,A.Address3 As PartyAddress3,A.Address4 As PartyAddress4,A.TIN As PartyGSTIN,IIf(P1.Type= '0810FI',C.PrintName,IIf(P1.Type= '0610FI',C.PrintName,IIf(P1.Type= '0510FR',M.PrintName,IIf(P1.Type= '0710FR',M.PrintName,'')))) As Consignee,IIf(P1.Type= '0810FI',C.Address1,IIf(P1.Type= '0610FI',C.Address1,IIf(P1.Type= '0510FR',M.Address1,IIf(P1.Type= '0710FR',M.Address1,'')))) As ConsigneeAddress1,IIf(P1.Type= '0810FI',C.Address2,IIf(P1.Type= '0610FI',C.Address2,IIf(P1.Type= '0510FR',M.Address2,IIf(P1.Type= '0710FR',M.Address2,'')))) As ConsigneeAddress2, " & _
                                                  "IIf(P1.Type= '0810FI',C.Address3,IIf(P1.Type= '0610FI',C.Address3,IIf(P1.Type= '0510FR',M.Address3,IIf(P1.Type= '0710FR',M.Address3,'')))) As ConsigneeAddress3,IIf(P1.Type= '0810FI',C.Address4,IIf(P1.Type= '0610FI',C.Address4,IIf(P1.Type= '0510FR',M.Address4,IIf(P1.Type= '0710FR',M.Address4,'')))) As ConsigneeAddress4,IIf(P1.Type= '0810FI',C.TIN,IIf(P1.Type= '0610FI',C.TIN,IIf(P1.Type= '0510FR',M.TIN,IIf(P1.Type= '0710FR',M.TIN,'')))) As ConsigneeGSTIN,P1.[Rebate%],P1.Rebate,P1.Freight,P1.Adjustment,P1.TaxableAmount,P1.[IGST%],P1.IGST,P1.[SGST%],P1.SGST,P1.[CGST%],P1.CGST,P1.Amount As TotalAmount,P1.Remarks,N.PrintName As Narration,I.PrintName As Item,H.PrintName As HSNCode," & _
                                                  "'Finish Size: '+LTRIM(S.PrintName) As FinishSize," & _
                                                  "C1.Quantity,C1.Rate,C1.Amount,N.Name As SrNo,'' As cmbTitle,LTRIM(C1.Code)+LTRIM(C1.SrNo) As Ref,M.PrintName As MC FROM (((((((JobworkBVParent P1 INNER JOIN JobworkBVChild C1 ON P1.Code=C1.Code)INNER JOIN AccountMaster A ON P1.Party=A.Code)INNER JOIN AccountMaster C ON P1.Consignee=C.Code)INNER JOIN BookMaster I ON C1.Item=I.Code)LEFT JOIN AccountMaster M ON P1.MaterialCentre=M.Code)LEFT JOIN GeneralMaster N ON C1.Narration=N.Code)LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code)LEFT JOIN GeneralMaster S ON I.FinishSize=S.Code WHERE P1.Code='" + Left(VchCode, 6) + "' ORDER BY I.PrintName,N.Name", cnDeliveryChallan, adOpenKeyset, adLockOptimistic
    
    '+', '+LTRIM(Case When C2.Pages1 IS NULL Then I.Pages Else (C2.Pages1+C2.Pages2+C2.Pages4)End)+' pages/'+LTRIM( Case When C2.Pages1 IS NULL Then I.Forms Else (C2.Forms1+C2.Forms2+C2.Forms4)End)+'f ('+ Case When C2.Pages1 IS NULL Then LTRIM(IIF(I.OneColorForms<>0,LTRIM(I.OneColorForms)+'f-1Col','')+' '+IIF(I.TwoColorForms<>0,LTRIM(I.TwoColorForms)+'f-2Col','')+' '+IIF(I.FourColorForms<>0,LTRIM(I.FourColorForms)+'f-4Col','')) Else LTRIM(IIF(C2.Forms1<>0,LTRIM(C2.Forms1)+'f-1Col','')+' '+IIF(C2.Forms2<>0,LTRIM(C2.Forms2)+'f-2Col','')+' '+IIF(C2.Forms4<>0,LTRIM(C2.Forms4)+'f-4Col','')) End+')'
    If rstDeliveryCVChild.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rstDeliveryCVChild.ActiveConnection = Nothing
    rptItemIssueReceiptVoucher.Text1.SetText IIf(ChallanType = "0710FR", "Sales Return", IIf(ChallanType = "0510FR", "Purchase", IIf(ChallanType = "0810FI", "Sales ", IIf(ChallanType = "0610FI", "Purchase Return", "Stock Transfer")))) & " Challan"
    rptItemIssueReceiptVoucher.Text13.SetText IIf(ChallanType = "0710FR", "Buyer :", IIf(ChallanType = "0510FR", "Supplier :", IIf(ChallanType = "0810FI", "Buyer :", IIf(ChallanType = "0610FI", "Supplier :", "From: Material Centre"))))
    rptItemIssueReceiptVoucher.Text7.SetText IIf(ChallanType = "0710FR", "Material Centre :", IIf(ChallanType = "0510FR", "Material Centre :", IIf(ChallanType = "0810FI", "Consignee :", IIf(ChallanType = "0610FI", "Consignee :", "TO: Material Centre"))))
    rptItemIssueReceiptVoucher.Text35.SetText "Printed on " & Format(Now, "dd-MMM-yyyy") & " at " & Format(Now, "hh:mm")
    rptItemIssueReceiptVoucher.Text40.SetText IIf(BillType = "O", "(ORIGINAL FOR RECIPIENT)", IIf(BillType = "D", "(DUPLICATE FOR SUPPLIER)", "(TRIPLICATE FOR SUPPLIER)"))
    rptItemIssueReceiptVoucher.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptItemIssueReceiptVoucher.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False)) Then
        rptItemIssueReceiptVoucher.Text4.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value) & Space(1) & "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptItemIssueReceiptVoucher.Text4.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False) Then
        rptItemIssueReceiptVoucher.Text4.SetText "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    End If
    rptItemIssueReceiptVoucher.Text8.SetText "GSTIN/UIN : " & Trim(rstCompanyMaster.Fields("GSTIN").Value)
    rptItemIssueReceiptVoucher.Text10.SetText "(" & UCase(Trim(NumberToWords(rstDeliveryCVChild.Fields("TotalAmount").Value, False))) & ")"
    rptItemIssueReceiptVoucher.Text11.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptItemIssueReceiptVoucher.Text26.SetText CheckNull(rstCompanyMaster.Fields("Declaration01").Value)
    rptItemIssueReceiptVoucher.Text25.SetText CheckNull(rstCompanyMaster.Fields("Declaration02").Value)
    rptItemIssueReceiptVoucher.Text22.SetText CheckNull(rstCompanyMaster.Fields("Declaration03").Value)
    rptItemIssueReceiptVoucher.Text12.SetText CheckNull(rstCompanyMaster.Fields("Declaration04").Value)
    rptItemIssueReceiptVoucher.Text9.SetText CheckNull(rstCompanyMaster.Fields("Declaration05").Value)
    rptItemIssueReceiptVoucher.Text30.SetText CheckNull(rstCompanyMaster.Fields("Declaration06").Value)
    rptItemIssueReceiptVoucher.Text31.SetText CheckNull(rstCompanyMaster.Fields("Declaration07").Value)
    rptItemIssueReceiptVoucher.Database.SetDataSource rstDeliveryCVChild, 3, 1
    rptItemIssueReceiptVoucher.DiscardSavedData
    Screen.MousePointer = vbNormal
    If OutputType = "S" Then
        Set FrmReportViewer.Report = rptItemIssueReceiptVoucher
        FrmReportViewer.Show vbModal
    Else
        If rstDeliveryCVList.State = adStateClosed Then  'For Print Utility
            rptItemIssueReceiptVoucher.PaperSource = crPRBinAuto
            rptItemIssueReceiptVoucher.PrintOut False
        Else
            rptItemIssueReceiptVoucher.PaperSource = crPRBinAuto
            rptItemIssueReceiptVoucher.PrintOut
        End If
    End If
    Set rptItemIssueReceiptVoucher = Nothing
    If rstDeliveryCVList.State = adStateClosed Then  'For Print Utility
        Call CloseRecordset(rstCompanyMaster)
    End If
    Call CloseRecordset(rstDeliveryCVChild)
    On Error GoTo 0
End Sub
