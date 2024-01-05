VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDebitCreditVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17715
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
   MaxButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   17715
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   8820
      Left            =   -105
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   19830
      _Version        =   65536
      _ExtentX        =   34978
      _ExtentY        =   15557
      _StockProps     =   77
      BackColor       =   16776946
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
      Picture         =   "DebitCreditVoucher.frx":0000
      Begin VB.CommandButton AddPayment 
         Caption         =   "Alt+ F5>> Add Payment"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   17040
         TabIndex        =   28
         Top             =   420
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton AddReceipt 
         Caption         =   " Alt+ F6>> Add Receipt"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   17040
         TabIndex        =   27
         Top             =   780
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton AddJournal 
         Caption         =   " Alt+ F7>> Add Journal"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   17040
         TabIndex        =   26
         Top             =   1140
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton AddDebitNote 
         Caption         =   "  Ctrl+ F6>> Add Debit Note"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   17040
         TabIndex        =   25
         Top             =   1860
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton AddCountra 
         Caption         =   " Ctrl+ F5>> Add Countra"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   17040
         TabIndex        =   24
         Top             =   1500
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton AddCreditNote 
         Caption         =   "  Ctrl+ F7>> Add Credit Note"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   17040
         TabIndex        =   23
         Top             =   2220
         Visible         =   0   'False
         Width           =   2655
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   8595
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   17625
         _ExtentX        =   31089
         _ExtentY        =   15161
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
         TabPicture(0)   =   "DebitCreditVoucher.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Toolbar2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Mh3dLabel1(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "DataGrid1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Text1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "DebitCreditVoucher.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dLabel1(1)"
         Tab(1).Control(1)=   "Mh3dFrame2"
         Tab(1).Control(2)=   "txtNotes"
         Tab(1).Control(3)=   "btnNotes"
         Tab(1).Control(4)=   "txtAccount"
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtAccount 
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
            Left            =   -61560
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   22
            ToolTipText     =   "Open Notes"
            Top             =   7320
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
            Left            =   -61560
            TabIndex        =   21
            Top             =   8070
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
            Height          =   345
            Left            =   -61560
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   7
            ToolTipText     =   "Open Notes"
            Top             =   7680
            Visible         =   0   'False
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
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   4
            Top             =   8180
            Width           =   7380
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "DebitCreditVoucher.frx":0054
            Height          =   7665
            Left            =   120
            TabIndex        =   3
            Top             =   450
            Width           =   17385
            _ExtentX        =   30665
            _ExtentY        =   13520
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   9164542
            HeadLines       =   1
            RowHeight       =   22
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "."
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "Date"
               Caption         =   "Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "VchSeriesName"
               Caption         =   "    Vch Series"
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
               DataField       =   "Name"
               Caption         =   "   Vch/Bill No."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "dd-MM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Account"
               Caption         =   "Account"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Debit"
               Caption         =   "              Debit"
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
            BeginProperty Column05 
               DataField       =   "Credit"
               Caption         =   "             Credit"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "ShortNarration"
               Caption         =   "Short Narration"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0.00"
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
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1349.858
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1604.976
               EndProperty
               BeginProperty Column03 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   4995.213
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
               EndProperty
               BeginProperty Column06 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   4334.74
               EndProperty
            EndProperty
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   8500
            TabIndex        =   6
            Top             =   8175
            Width           =   8295
            _Version        =   65536
            _ExtentX        =   14631
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
            Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D->Delete-Vch F8->Delete-Vch  Ctrl+S->Save F2->Save  F12->Duplicate"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "DebitCreditVoucher.frx":0064
            Picture         =   "DebitCreditVoucher.frx":0080
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7935
            Left            =   -74880
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   480
            Width           =   13260
            _Version        =   65536
            _ExtentX        =   23389
            _ExtentY        =   13996
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
            Picture         =   "DebitCreditVoucher.frx":009C
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   1740
               MaxLength       =   100
               TabIndex        =   14
               Top             =   7210
               Width           =   11415
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   7500
               MaxLength       =   25
               TabIndex        =   13
               Top             =   105
               Width           =   1830
            End
            Begin VB.TextBox Text8 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   9
               Top             =   105
               Width           =   1890
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   525
               Left            =   120
               TabIndex        =   10
               Top             =   6420
               Width           =   13035
               _Version        =   65536
               _ExtentX        =   22992
               _ExtentY        =   926
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
               Picture         =   "DebitCreditVoucher.frx":00B8
               Picture         =   "DebitCreditVoucher.frx":00D4
               Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
                  Height          =   525
                  Left            =   5910
                  TabIndex        =   11
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1470
                  _Version        =   65536
                  _ExtentX        =   2593
                  _ExtentY        =   926
                  Calculator      =   "DebitCreditVoucher.frx":00F0
                  Caption         =   "DebitCreditVoucher.frx":0110
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "DebitCreditVoucher.frx":017C
                  Keys            =   "DebitCreditVoucher.frx":019A
                  Spin            =   "DebitCreditVoucher.frx":01E4
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
               Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
                  Height          =   525
                  Left            =   7365
                  TabIndex        =   12
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1470
                  _Version        =   65536
                  _ExtentX        =   2593
                  _ExtentY        =   926
                  Calculator      =   "DebitCreditVoucher.frx":020C
                  Caption         =   "DebitCreditVoucher.frx":022C
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "DebitCreditVoucher.frx":0298
                  Keys            =   "DebitCreditVoucher.frx":02B6
                  Spin            =   "DebitCreditVoucher.frx":0300
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
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   390
               Left            =   6060
               TabIndex        =   15
               Top             =   105
               Width           =   1470
               _Version        =   65536
               _ExtentX        =   2593
               _ExtentY        =   688
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
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
               Picture         =   "DebitCreditVoucher.frx":0328
               Picture         =   "DebitCreditVoucher.frx":0344
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   570
               Left            =   120
               TabIndex        =   16
               Top             =   7210
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   1005
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Long Narration"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "DebitCreditVoucher.frx":0360
               Picture         =   "DebitCreditVoucher.frx":037C
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   390
               Left            =   11460
               TabIndex        =   17
               Top             =   105
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   688
               Calendar        =   "DebitCreditVoucher.frx":0398
               Caption         =   "DebitCreditVoucher.frx":04B0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "DebitCreditVoucher.frx":051C
               Keys            =   "DebitCreditVoucher.frx":053A
               Spin            =   "DebitCreditVoucher.frx":0598
               AlignHorizontal =   2
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
               Height          =   5715
               Left            =   120
               TabIndex        =   18
               Top             =   750
               Width           =   13035
               _Version        =   524288
               _ExtentX        =   22992
               _ExtentY        =   10081
               _StockProps     =   64
               EditEnterAction =   5
               EditModeReplace =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               GridColor       =   4227327
               MaxCols         =   7
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "DebitCreditVoucher.frx":05C0
               VisibleCols     =   6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   390
               Index           =   0
               Left            =   10260
               TabIndex        =   19
               Top             =   105
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   688
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
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
               Picture         =   "DebitCreditVoucher.frx":0D93
               Picture         =   "DebitCreditVoucher.frx":0DAF
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   390
               Left            =   120
               TabIndex        =   20
               Top             =   105
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   688
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
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
               Picture         =   "DebitCreditVoucher.frx":0DCB
               Picture         =   "DebitCreditVoucher.frx":0DE7
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   13240
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   13240
               Y1              =   7080
               Y2              =   7080
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   810
            Index           =   1
            Left            =   -61485
            TabIndex        =   29
            Top             =   480
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   1429
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
            Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D->Delete-Row F8->Delete-Row  Ctrl+S->Save F2->Save  F12->Duplicate"
            AutoSize        =   -1  'True
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "DebitCreditVoucher.frx":0E03
            Multiline       =   -1  'True
            Picture         =   "DebitCreditVoucher.frx":0E1F
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   630
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   0
            _ExtentX        =   0
            _ExtentY        =   1111
            ButtonWidth     =   3625
            ButtonHeight    =   1005
            AllowCustomize  =   0   'False
            Appearance      =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   23
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "F1 >> Help "
                  Key             =   "F1"
                  Object.ToolTipText     =   "F1 >> Help"
                  Object.Tag             =   "F!"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Alt+F1 >> Add Account "
                  Key             =   "Alt+F1"
                  Object.ToolTipText     =   "Alt+F1 >> Add Account "
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Alt+F2 >>Add Item "
                  Key             =   "Alt+F2"
                  Object.ToolTipText     =   "Alt+F2 >>Add Item "
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Alt+F5 >>Add Payment "
                  Key             =   "Alt+F5"
                  Object.ToolTipText     =   "Alt+F5 >>Add Payment "
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Alt+F6 >>Add Receipt "
                  Key             =   "Alt+F6"
                  Object.ToolTipText     =   "Alt+F6 >>Add Receipt "
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Alt+F7 >>Add Journal "
                  Key             =   "Alt+F7"
                  Object.ToolTipText     =   "Alt+F7 >>Add Journal"
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Ctrl+F5 >>Add Countra "
                  Key             =   "Ctrl+F5"
                  Object.ToolTipText     =   "Ctrl+F5 >>Add Countra "
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Ctrl+F6 >>Add Debit "
                  Key             =   "Ctrl+F6"
                  Object.ToolTipText     =   "Ctrl+F6 >>Add DebitNote"
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Ctrl+F7 >>Add Credit "
                  Key             =   "Ctrl+F7"
                  Object.ToolTipText     =   "Ctrl+F7 >>Add Credit "
               EndProperty
               BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "F10>> Calculator "
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H008BD6FE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Find"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            TabIndex        =   5
            Top             =   8175
            Width           =   975
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17715
      _ExtentX        =   31247
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
Attribute VB_Name = "frmDebitCreditVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchCode As String  'Vch to Modify
Public VchType As String 'PI-Payment Issued, PR-Payment Receipt,JE-Journal Entry, CE-Countra Entry,DN-Debit Note,CN-CreditNote
Dim cnDebitCreditVoucher As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstJEAccountList As New ADODB.Recordset, rstSelfAccountList As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstVchSeriesList As New ADODB.Recordset, rstItemList As New ADODB.Recordset
Dim rstDebitCreditVoucherList As New ADODB.Recordset, rstDebitCreditVoucherParent As New ADODB.Recordset, rstDebitCreditVoucherChild As New ADODB.Recordset
Dim VchPrefix As String, VchNumbering As String, VchSeriesCode As String, oVchSeriesCode As String, oVchNo As String, AutoVchNo As String, AccountCode As String ', MaterialCentreCode As String, TaxCode As String
Dim SortOrder, PrevStr, dblBookMark As Double, blnRecordExist As Boolean, EditMode As Boolean, oDCFlag As Integer
Dim oDebit As Variant, oCredit As Variant, dDCFlag As Long
Private Sub AddPayment_Click()
If EditMode Then
    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
        Me.ActiveControl.SetFocus
    End If
Else
    If VchType = "PI" Then Exit Sub
    CloseForm Me
    frmDebitCreditVoucher.VchType = "PI"
    Load frmDebitCreditVoucher
    If Err.Number <> 364 Then frmDebitCreditVoucher.Show
    frmDebitCreditVoucher.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(1)
    Exit Sub
End If
End Sub
Private Sub AddReceipt_Click()
If EditMode Then
    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
        Me.ActiveControl.SetFocus
    End If
Else
    If VchType = "PR" Then Exit Sub
    CloseForm Me
    frmDebitCreditVoucher.VchType = "PR"
    Load frmDebitCreditVoucher
    If Err.Number <> 364 Then frmDebitCreditVoucher.Show
    frmDebitCreditVoucher.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(1)
    Exit Sub
End If
End Sub
Private Sub AddJournal_Click()
If EditMode Then
    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
        Me.ActiveControl.SetFocus
    End If
Else
    If VchType = "JE" Then Exit Sub
    CloseForm Me
    frmDebitCreditVoucher.VchType = "JE"
    Load frmDebitCreditVoucher
    If Err.Number <> 364 Then frmDebitCreditVoucher.Show
    frmDebitCreditVoucher.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(1)
    Exit Sub
End If
End Sub
Private Sub AddCountra_Click()
If EditMode Then
    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
        Me.ActiveControl.SetFocus
    End If
Else
    If VchType = "CE" Then Exit Sub
    CloseForm Me
    frmDebitCreditVoucher.VchType = "CE"
    Load frmDebitCreditVoucher
    If Err.Number <> 364 Then frmDebitCreditVoucher.Show
    frmDebitCreditVoucher.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(1)
    Exit Sub
End If
End Sub
Private Sub AddDebitNote_Click()
If EditMode Then
    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
        Me.ActiveControl.SetFocus
    End If
Else
    If VchType = "DN" Then Exit Sub
    CloseForm Me
    frmDebitCreditVoucher.VchType = "DN"
    Load frmDebitCreditVoucher
    If Err.Number <> 364 Then frmDebitCreditVoucher.Show
    frmDebitCreditVoucher.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(1)
    Exit Sub
End If
End Sub
Private Sub AddCreditNote_Click()
If EditMode Then
    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
        Me.ActiveControl.SetFocus
    End If
Else
    If VchType = "CN" Then Exit Sub
    CloseForm Me
    frmDebitCreditVoucher.VchType = "CN"
    Load frmDebitCreditVoucher
    If Err.Number <> 364 Then frmDebitCreditVoucher.Show
    frmDebitCreditVoucher.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(1)
    Exit Sub
End If
End Sub
Private Sub btnNotes_Click()
    frmNotes.NotesFlag = 3
    frmNotes.Label1.Caption = "Notes : Voucher No.: " & Text2.Text
    frmNotes.Show (vbModal)
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    Me.Left = 120
    Me.Top = 1200
    WheelHook DataGrid1
    BusySystemIndicator True
        Dim Cols As Long, C As Long
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader
            fpSpread1.UserColAction = UserColActionSort
            Cols = fpSpread1.MaxCols
        For C = 1 To Cols
            fpSpread1.ColUserSortIndicator(C) = ColUserSortIndicatorDescending
            Next
    Me.Caption = IIf(VchType = "PI", "Payment Voucher", IIf(VchType = "PR", "Receipt Voucher", IIf(VchType = "JE", "Journal Voucher", IIf(VchType = "CE", "Contra Voucher", IIf(VchType = "DN", "Debit Note Voucher", "Credit Note Voucher")))))
    VchPrefix = IIf(VchType = "PI", "51", IIf(VchType = "PR", "52", IIf(VchType = "JE", "53", IIf(VchType = "CE", "54", IIf(VchType = "DN", "55", "56"))))) & IIf(VchType = "CE", "10", "01") '10-Contra not affected 01-affected
    DataGrid1.Caption = IIf(VchType = "PI", "List of Payment Vouchers", IIf(VchType = "PR", "List of Receipt Vouchers", IIf(VchType = "53", "List of Journal Vouchers", IIf(VchType = "54", "List of Contra Vouchers", IIf(VchType = "55", "List of Debit Note Vouchers", "List of Credit Note Vouchers")))))
    cnDebitCreditVoucher.CursorLocation = adUseClient
    
    If cnDebitCreditVoucher.State Then cnDebitCreditVoucher.Close
    cnDebitCreditVoucher.Open cnDatabase.ConnectionString
    
    rstDebitCreditVoucherParent.CursorLocation = adUseClient
    LoadMasterList
    
    With rstDebitCreditVoucherList
    If rstDebitCreditVoucherList.State Then rstDebitCreditVoucherList.Close
        .Open "SELECT T.Code As Code,T.Name,V.Name As VchSeriesName,Date,A.Name As Account, C.Debit,C.Credit,ShortNarration,T.Type FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE RIGHT(Type,2)='" & VchType & "' AND T.FYCode='" & FYCode & "' ORDER BY T.CODE", cnDebitCreditVoucher, adOpenKeyset, adLockPessimistic
        .Filter = adFilterNone
        If .RecordCount > 0 Then
            .MoveLast
            If Not CheckEmpty(VchCode, False) Then .MoveFirst: .Find "[Code]='" & VchCode & "'"
        End If

''        Set DataGrid1.DataSource = rstDebitCreditVoucherList

If rstDebitCreditVoucherList.State = adStateOpen And Not rstDebitCreditVoucherList.EOF Then
    Set DataGrid1.DataSource = rstDebitCreditVoucherList
Else
    ' Handle the case where the recordset is not in a valid state.
End If

        BusySystemIndicator False
        SSTab1.Tab = 0
    If FrmAccountLedger.dSortBy = True Then
        SortOrder = "Code"
    Else
        SortOrder = "Name"
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
        .mnuFinanceModuleParent.Enabled = False
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
            Toolbar1_ButtonClick .Item(6)
            KeyCode = 0
        ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2)) And .Item(4).Enabled Then 'Save
            If MhRealInput1.Value <> MhRealInput2.Value Then
            If MsgBox("Variation in Debit (" & MhRealInput1 & ") and Credit (" & MhRealInput2 & ") !!! Change?", vbCritical, "Confirm Change !") = vbYes Then fpSpread1.SetFocus: Exit Sub
            Else
            If Not EditMode Then Toolbar1_ButtonClick .Item(4)
            Toolbar1_ButtonClick .Item(6)
            KeyCode = 0
            End If
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
        ElseIf Shift = vbAltMask And KeyCode = vbKeyF5 Then
            'AddPayment_Click
            Toolbar2_ButtonClick .Item(6)
            KeyCode = 0
        ElseIf Shift = vbAltMask And KeyCode = vbKeyF6 Then
            'AddReceipt_Click
            Toolbar2_ButtonClick .Item(7)
            KeyCode = 0
        ElseIf Shift = vbAltMask And KeyCode = vbKeyF7 Then
            Toolbar2_ButtonClick .Item(8)
            'AddJournal_Click
            KeyCode = 0
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF5 Then
            'AddCountra_Click
            Toolbar2_ButtonClick .Item(9)
            KeyCode = 0
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF6 Then
            Toolbar2_ButtonClick .Item(10)
            'AddDebitNote_Click
            KeyCode = 0
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF7 Then
        Toolbar2_ButtonClick .Item(11)
            'AddCreditNote_Click
            KeyCode = 0
        ElseIf Shift = 0 And KeyCode = vbKeyF10 Then
            Shell "calc.exe", vbNormalFocus
        End If
    End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstDebitCreditVoucherList)
    Call CloseRecordset(rstDebitCreditVoucherParent)
    Call CloseRecordset(rstDebitCreditVoucherChild)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstSelfAccountList)
    Call CloseRecordset(rstJEAccountList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstVchSeriesList)
    Call CloseConnection(cnDebitCreditVoucher)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuFinanceModuleParent.Enabled = True
End Sub
Private Sub Text1_Change()
On Error Resume Next
    With rstDebitCreditVoucherList
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
    With rstDebitCreditVoucherList
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
    On Error Resume Next
    If Toolbar1.Buttons.Item(1).Enabled Then 'Add Button Enabled
        If SSTab1.Tab = 1 Then
            ViewRecord
        Else
            If Not (rstDebitCreditVoucherList.EOF Or rstDebitCreditVoucherList.BOF) Then
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
    Dim HiLiteRecord As Boolean, UpdateFlag As Integer, CellVal01 As Variant, CellVal02 As Variant, CellVal03 As Variant, CellVal04 As Variant, CellVal05 As Variant, CellVal06 As Variant, i As Integer
    With rstDebitCreditVoucherList
        If Button.Index = 1 Then
            If rstDebitCreditVoucherParent.State = adStateOpen Then rstDebitCreditVoucherParent.Close
            rstDebitCreditVoucherParent.Open "SELECT * FROM DebitCreditParent WHERE Code=''", cnDebitCreditVoucher, adOpenKeyset, adLockOptimistic
            ClearFields
            If AddRecord(rstDebitCreditVoucherParent) Then
                Text2.Text = GenerateCode(cnDebitCreditVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "LTRIM(CONVERT(INT,[AutoVchNo])))", "VAL([AutoVchNo]))") & "  FROM  DebitCreditParent WHERE RIGHT(Type,2)='" & VchType & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
                Call SetButtons(False)
                SSTab1.Tab = 1
                Text8.SetFocus
                blnRecordExist = False
                cnDebitCreditVoucher.BeginTrans
            End If
        ElseIf Button.Index = 2 Then
            If .RecordCount = 0 Then Exit Sub
            SSTab1.Tab = 1
            EditRecord
        ElseIf Button.Index = 3 Then
            If .RecordCount = 0 Then Exit Sub
            If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
            SSTab1.Tab = 1
            If chkRef("SELECT RefCode FROM DebitCreditRef WHERE VchCode='" & .Fields("Code").Value & "' AND RefCode IN (SELECT RefCode FROM DebitCreditRef WHERE VchCode<>'" & .Fields("Code").Value & "')") Then
                DisplayError ("Failed to delete the record")
            ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
                On Error Resume Next
                MdiMainMenu.MousePointer = vbHourglass
                cnDebitCreditVoucher.BeginTrans
                cnDebitCreditVoucher.Execute "DELETE FROM DebitCreditRef WHERE VchCode='" & .Fields("Code").Value & "'"
                cnDebitCreditVoucher.Execute "DELETE FROM DebitCreditParent WHERE Code='" & .Fields("Code").Value & "'"
                MdiMainMenu.MousePointer = vbNormal
                If Err.Number = 0 Then
                    .Delete
                    .MoveNext
                    If .RecordCount > 0 And .EOF Then .MoveLast
                    cnDebitCreditVoucher.CommitTrans
                    ShowProgressInStatusBar True
                    Timer1.Enabled = True
                    Text1.Text = ""
                    .Filter = adFilterNone
                Else
                    DisplayError (Err.Description)
                    cnDebitCreditVoucher.RollbackTrans
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
            If oDCFlag = 1 Then oDCFlag = 0: Exit Sub
            If UpdateRecord(rstDebitCreditVoucherParent) Then
                If UpdateItemList("D", 0, "") Then
                    UpdateFlag = 1
                   With fpSpread1
                       For i = 1 To .DataRowCnt
                           .SetActiveCell 1, i
                           .GetText 1, i, CellVal01 'TOA
                           .GetText 2, i, CellVal02 'Account
                           .GetText 3, i, CellVal03 'Debit
                           .GetText 4, i, CellVal04 'Credit
                           .GetText 6, i, CellVal05 'Account Code
                           .GetText 7, i, CellVal06 'Ref Code
                           If (CellVal01) <> "" And Not CheckEmpty(CellVal05, False) And Val(CellVal03) <> 0 Or Val(CellVal04) <> 0 Then If Not UpdateItemList("I", i, CellVal06) Then UpdateFlag = 0: Exit For
                       Next
                   End With
                End If
            End If
            If UpdateFlag Then
                AddToList
                cnDebitCreditVoucher.CommitTrans
                If rstDebitCreditVoucherParent.State = adStateOpen Then rstDebitCreditVoucherParent.Close
                rstDebitCreditVoucherParent.CursorLocation = adUseClient
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
            If CancelRecordUpdate(rstDebitCreditVoucherParent) Then
                cnDebitCreditVoucher.RollbackTrans
                If rstDebitCreditVoucherParent.State = adStateOpen Then rstDebitCreditVoucherParent.Close
                rstDebitCreditVoucherParent.CursorLocation = adUseClient
                Call SetButtons(True)
                SetButtonsForNoRecord
                SSTab1.Tab = 0
            End If
        ElseIf Button.Index = 6 Then
            SSTab1.Tab = 0
            Set DataGrid1.DataSource = Nothing
            .Filter = adFilterNone
            RefreshData rstDebitCreditVoucherList
            Set DataGrid1.DataSource = rstDebitCreditVoucherList
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
            Call PrintDebitCreditVoucher(.Fields("Code").Value, Right(.Fields("Type").Value, 2), "P")
            HiLiteRecord = True
        ElseIf Button.Index = 10 Then
            If .RecordCount = 0 Then Exit Sub
            Call PrintDebitCreditVoucher(.Fields("Code").Value, Right(.Fields("Type").Value, 2), "S")
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
        rstDebitCreditVoucherList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstDebitCreditVoucherList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstDebitCreditVoucherList.EOF Or rstDebitCreditVoucherList.BOF) Then
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
    If rstDebitCreditVoucherList.RecordCount = 0 Then
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
                AutoVchNo = GenerateCode(cnDebitCreditVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  DebitCreditParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
            End If
        Else 'Vch-Old
            If VchSeriesCode = oVchSeriesCode Then
                Text2.Text = oVchNo
            Else
                If VchNumbering = "A" Then
                    AutoVchNo = GenerateCode(cnDebitCreditVoucher, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  DebitCreditParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                    Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                End If
            End If
        End If
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean) 'Vch No.
    With rstDebitCreditVoucherParent
        If .EOF Or .BOF Then Exit Sub
        If CheckEmpty(Text2, True) Then
            Cancel = True
        ElseIf CheckDuplicate(cnDebitCreditVoucher, "DebitCreditParent", "Code", "[Name]+RIGHT(Type,2)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, .Fields("Code").Value, False, FYCode) Then
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
'    If KeyCode = vbKeySpace Then
'        On Error Resume Next
'        FrmAccountMaster.SL = True
'        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = IIf(VchType = "ST", "*99999", "")
'        FrmAccountMaster.MasterCode = PartyCode
'        Load FrmAccountMaster
'        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
'        On Error GoTo 0
'        PartyCode = slCode: Text3.Text = slName
'        If Not CheckEmpty(PartyCode, False) Then LoadMasterList: Sendkeys "{TAB}"
'    ElseIf KeyCode = vbKeyDelete Then
'        PartyCode = "": Text3.Text = ""
'    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
'    If CheckEmpty(Text3.Text, False) Then Cancel = True
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeySpace Then
'        On Error Resume Next
'        FrmAccountMaster.SL = True
'        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = "*99999"
'        FrmAccountMaster.MasterCode = MaterialCentreCode
'        Load FrmAccountMaster
'        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
'        On Error GoTo 0
'        MaterialCentreCode = slCode: Text7.Text = slName
'        If Not CheckEmpty(MaterialCentreCode, False) Then LoadMasterList: Sendkeys "{TAB}"
'    ElseIf KeyCode = vbKeyDelete Then
'        MaterialCentreCode = "": Text7.Text = ""
'    End If
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
   ' If CheckEmpty(Text7.Text, False) Then Cancel = True
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeySpace Then
'        On Error Resume Next
'        FrmTaxMaster.SL = True
'        FrmTaxMaster.MasterCode = TaxCode
'        Load FrmTaxMaster
'        If Err.Number <> 364 Then FrmTaxMaster.Show vbModal
'        On Error GoTo 0
'        TaxCode = slCode: Text5.Text = slName
'        If Not CheckEmpty(TaxCode, False) Then
'            rstTaxList.MoveFirst: rstTaxList.Find "[Code] = '" & TaxCode & "'"
'            If Val(rstTaxList.Fields("SGST%").Value) > 0 Then   'Intra-State GST
'                MhRealInput7.Value = Val(rstTaxList.Fields("CGST%").Value)
'                MhRealInput9.Value = Val(rstTaxList.Fields("SGST%").Value)
'            Else    'Inter-State GST
'                MhRealInput7.Value = Val(rstTaxList.Fields("IGST%").Value)
'                MhRealInput9.Value = 0
'            End If
'            CalculateTotal
'            LoadMasterList
'            Sendkeys "{TAB}"
'        End If
'    ElseIf KeyCode = vbKeyDelete Then
'        TaxCode = "": Text5.Text = ""
'    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
'    If CheckEmpty(Text5.Text, False) Then Cancel = True
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
    If rstDebitCreditVoucherList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    With rstDebitCreditVoucherParent
        If .State = adStateOpen Then .Close
        .Open "SELECT * FROM DebitCreditParent WHERE Code='" & FixQuote(rstDebitCreditVoucherList.Fields("Code").Value) & "'", cnDebitCreditVoucher, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End With
End Sub
Private Sub ClearFields()
    Text8.Text = "" 'Vch Series
    Text2.Text = "" 'Vch No.
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    Text4.Text = "" 'Remarks
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    MhRealInput1.Value = 0
    MhRealInput2.Value = 0
     VchSeriesCode = "": oVchSeriesCode = "": oVchNo = "": AutoVchNo = "": AccountCode = "":
End Sub
Private Sub LoadFields()
    With rstDebitCreditVoucherParent
        If .EOF Or .BOF Then Exit Sub
        VchSeriesCode = .Fields("VchSeries").Value: oVchSeriesCode = VchSeriesCode
        If rstVchSeriesList.RecordCount > 0 Then rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        If Not rstVchSeriesList.EOF Then Text8.Text = rstVchSeriesList.Fields("Col0").Value
        AutoVchNo = Trim(.Fields("AutoVchNo").Value)
        Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value) '.Fields("Name").Value
        oVchNo = Trim(Text2.Text)
        MhDateInput1.Text = Format(.Fields("Date").Value, "dd-MM-yyyy")
        Text4.Text = .Fields("LongNarration").Value
        Call LoadItemList(.Fields("Code").Value)
    txtNotes.Text = .Fields("Notes").Value
    End With
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    With rstDebitCreditVoucherParent
        If .RecordCount = 0 Then Exit Sub
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseServer
        .Open "SELECT * FROM DebitCreditParent WHERE Code='" & FixQuote(rstDebitCreditVoucherList.Fields("Code").Value) & "'", cnDebitCreditVoucher, adOpenKeyset, adLockPessimistic
        MdiMainMenu.MousePointer = vbHourglass
        .Fields("RecordStatus") = "N"
    End With
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text8.SetFocus
    blnRecordExist = True
    cnDebitCreditVoucher.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If MhRealInput1.Value <> MhRealInput2.Value Then
    oDCFlag = 1
    If MsgBox("Variation in Debit Amount ( " & MhRealInput1 & " ) and Credit  Amount ( " & MhRealInput2 & " ) !!! Change?", vbCritical, "Confirm Change !!!") = vbYes Then fpSpread1.SetFocus
    With fpSpread1
    .SetActiveCell 2, 1
    End With
    Exit Sub
    Else
    With rstDebitCreditVoucherParent
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnDebitCreditVoucher, "SELECT MAX(Code) FROM DebitCreditParent", 6, "0")
            .Fields("CreatedBy").Value = UserCode
            .Fields("CreatedOn").Value = Now()
            .Fields("Recordstatus").Value = "N"
        Else
            .Fields("ModifiedBy").Value = UserCode
            .Fields("ModifiedOn").Value = Now()
            .Fields("Recordstatus").Value = "M"
        End If
        .Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 15, "L") 'Pad(Trim(AutoVchNo), Space(1), 10, "L") '
        .Fields("VchSeries").Value = VchSeriesCode
        .Fields("AutoVchNo").Value = Pad(Trim(AutoVchNo), Space(1), 10, "L")
        .Fields("Date").Value = GetDate(MhDateInput1.Text)
        .Fields("LongNarration").Value = Trim(Text4.Text)
        .Fields("Debit").Value = MhRealInput1.Value
        .Fields("Credit").Value = MhRealInput2.Value
        .Fields("Type").Value = VchPrefix & VchType
        .Fields("FYCode").Value = FYCode
        .Fields("RecordStatus").Value = "N"
        .Fields("Notes").Value = txtNotes.Text
    End With
    End If
End Sub
Private Sub AddToList()
    On Error Resume Next
    With rstDebitCreditVoucherList
        .MoveFirst
        .Find "[Code] = '" & rstDebitCreditVoucherParent.Fields("Code").Value & "'"
        If .EOF Then .AddNew
        .Fields("Code").Value = rstDebitCreditVoucherParent.Fields("Code").Value
        .Fields("Name").Value = Pad(rstDebitCreditVoucherParent.Fields("Name").Value, Space(1), 10, "L")
        .Fields("Date").Value = rstDebitCreditVoucherParent.Fields("Date").Value
        .Fields("Debit").Value = MhRealInput1.Value
        .Fields("Credit").Value = MhRealInput2.Value
        .Fields("LongNarration").Value = rstDebitCreditVoucherParent.Fields("LongNarration").Value
        .Fields("Type").Value = rstDebitCreditVoucherParent.Fields("Type").Value
        .Fields("VchSeriesName").Value = Text8.Text
        .Update
        .Sort = SortOrder & " Asc"
        .Find "[Code] = '" & rstDebitCreditVoucherParent.Fields("Code").Value & "'"
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text8.Text, False) Then
        Text8.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text2.Text, False) Then
        DisplayError ("Voucher No. cannot be blank"): Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnDebitCreditVoucher, "DebitCreditParent", "Code", "[Name]+RIGHT(Type,2)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, rstDebitCreditVoucherParent.Fields("Code").Value, False, FYCode) Then
        Text2.SetFocus: CheckMandatoryFields = True: Exit Function
'    ElseIf CheckEmpty(Text3.Text, False) Then 'Party/From Mat Centre
'        Text3.SetFocus:   CheckMandatoryFields = True: Exit Function
'    ElseIf CheckEmpty(Text7.Text, False) Then 'Mat Centre
'        Text7.SetFocus:   CheckMandatoryFields = True: Exit Function
'    ElseIf CheckEmpty(Text5.Text, False) Then 'Tax
'        Text5.SetFocus:   CheckMandatoryFields = True: Exit Function
'    ElseIf VchType = "ST" Then
'        If Text3.Text = Text7.Text Then DisplayError ("Source & Target Material Centres cann't be same"): Text3.SetFocus: CheckMandatoryFields = True: Exit Function
    End If
End Function
Private Sub LoadItemList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    With rstDebitCreditVoucherChild
        If .State = adStateOpen Then .Close
        .Open "SELECT T.Code,TOA,A.Code As AccountCode,A.Name As AccountName,(T.Debit) As Debit,(T.Credit) As Credit,T.ShortNarration,T.RefCode,T.SrNo FROM DebitCreditChild T INNER JOIN AccountMaster A ON T.Account=A.Code WHERE T.Code='" & strOrderCode & "'  ORDER BY SrNo", cnDebitCreditVoucher, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        If .RecordCount > 0 Then .MoveFirst
        i = 0
        Do While Not .EOF
            i = i + 1
            fpSpread1.SetText 1, i, .Fields("TOA").Value
            fpSpread1.SetText 2, i, .Fields("AccountName").Value
            fpSpread1.SetText 3, i, Val(.Fields("Debit").Value)
            fpSpread1.SetText 4, i, Val(.Fields("Credit").Value)
            fpSpread1.SetText 5, i, (.Fields("ShortNarration").Value)
            fpSpread1.SetText 6, i, .Fields("AccountCode").Value
            fpSpread1.SetText 7, i, .Fields("RefCode").Value
            .MoveNext
        Loop
    End With
    CalculateTotal
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Item List")
End Sub
Private Function UpdateItemList(ByVal ActionType As String, ByVal SrNo As Integer, ByVal RefCode As String) As Boolean
    Dim CellVal(1 To 8) As Variant
    On Error GoTo ErrorHandler
    UpdateItemList = True
    If ActionType = "D" Then
        If Not blnRecordExist Then Exit Function
        cnDebitCreditVoucher.Execute "DELETE FROM DebitCreditRef WHERE VchCode='" & rstDebitCreditVoucherParent.Fields("Code").Value & "'"
        cnDebitCreditVoucher.Execute "DELETE FROM DebitCreditChild WHERE Code='" & rstDebitCreditVoucherParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 1, .ActiveRow, CellVal(1)  'Type of Account
            .GetText 3, .ActiveRow, CellVal(2)  'Debit
            .GetText 4, .ActiveRow, CellVal(3)  'Credit
            .GetText 5, .ActiveRow, CellVal(4)  'Narration
            .GetText 6, .ActiveRow, CellVal(5)  'Account Code
        End With
            If VchType = "PI" Or VchType = "PR" Or VchType = "JE" Or VchType = "CE" Or VchType = "DN" Or VchType = "CN" Then
                If CheckEmpty(RefCode, False) Then RefCode = GenerateCode(cnDebitCreditVoucher, "SELECT MAX(RefCode) FROM DebitCreditRef", 6, "0")
                cnDebitCreditVoucher.Execute "INSERT INTO DebitCreditChild VALUES ('" & rstDebitCreditVoucherParent.Fields("Code").Value & "','" & CellVal(1) & "','','" & VchPrefix & VchType & "','" & CellVal(5) & "'," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ",'" & (CellVal(4)) & "'," & SrNo & ",'" & RefCode & "')"
                cnDebitCreditVoucher.Execute "INSERT INTO DebitCreditRef VALUES ('" & RefCode & "',1,'" & VchPrefix & VchType & "','" & rstDebitCreditVoucherParent.Fields("Code").Value & "','" & rstDebitCreditVoucherParent.Fields("Name").Value & "','" & Format(rstDebitCreditVoucherParent.Fields("Date").Value, "dd-MMM-yyyy") & "','" & CellVal(5) & "'," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ",'" & (CellVal(1)) & "')"
            Else
                'CellVal(1) = IIf(VchType = "SO", Val(CellVal(1)), 0 - Val(CellVal(1))) '+ve/+ve for SO & -ve/-ve for PO (child/ref)
               ' If CheckEmpty(RefCode, False) Then RefCode = GenerateCode(cnDebitCreditVoucher, "SELECT MAX(RefCode) FROM DebitCreditRef", 6, "0")
                'cnDebitCreditVoucher.Execute "INSERT INTO DebitCreditRef VALUES ('" & RefCode & "',1,'" & VchPrefix & VchType & "','" & rstDebitCreditVoucherParent.Fields("Code").Value & "','" & rstDebitCreditVoucherParent.Fields("Name").Value & "','" & Format(rstDebitCreditVoucherParent.Fields("Date").Value, "dd-MMM-yyyy") & "','" & CellVal(6) & "'," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ",'" & (CellVal(1)) & "')"
            End If
        'cnDebitCreditVoucher.Execute "INSERT INTO DebitCreditChild VALUES ('" & rstDebitCreditVoucherParent.Fields("Code").Value & "','','" & VchPrefix & "FI" & "','" & CellVal(5) & "','" & CellVal(6) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(4)) & ",Null," & SrNo & ",'','','','',''," & Val(CellVal(3)) & ",'" & RefCode & "')"
    End If
    Exit Function
ErrorHandler:
    UpdateItemList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
'    If SrchFor = "Party" Then
'        rstDebitCreditVoucherList.Filter = "[PartyName] Like '%" & SrchText & "%'"
'    ElseIf SrchFor = "Material Centre" Then
'        rstDebitCreditVoucherList.Filter = "[MaterialCentreName] Like '%" & SrchText & "%'"
'    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Item As Variant, i As Integer, x As Integer, cVal(1 To 8) As Variant
With fpSpread1
    If Shift = 0 And KeyCode = vbKeyF9 Then
            .GetText 8, .ActiveRow, Item  'Ref Code
                If Not CheckEmpty(Item, False) Then
                    If chkRef("SELECT RefCode FROM DebitCreditRef WHERE RefCode='" & Item & "' AND VchCode<>'" & rstDebitCreditVoucherParent.Fields("Code").Value & "'") Then DisplayError ("Failed to delete the record"): .SetFocus
                ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
                    .DeleteRows .ActiveRow, 1: .SetFocus: CalculateTotal
                End If
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
            If .ActiveCol = 3 Or .ActiveCol = 4 Then  'Debit
            .GetText 1, .ActiveRow, cVal(1)  'TOA
            .GetText 3, .ActiveRow, cVal(3)  'Debit
            .GetText 4, .ActiveRow, cVal(4)  'Credit
            If cVal(1) = "D" And cVal(3) = 0 Then Call MsgBox("Debit Amount Can't be zero !!!", vbInformation, App.Title): .SetActiveCell 3, .ActiveRow: .SetFocus:
            If cVal(1) = "C" And cVal(4) = 0 Then Call MsgBox("Credit Amount Can't be zero !!!", vbInformation, App.Title): .SetActiveCell 4, .ActiveRow: .SetFocus:
            End If
    ElseIf KeyCode = vbKeyF3 Then
            If .ActiveCol = 2 Then
                .GetText 6, .ActiveRow, Item 'Ref Code
                If Not CheckEmpty(Item, False) Then If chkRef("SELECT RefCode FROM DebitCreditRef WHERE RefCode='" & Item & "' AND VchCode<>'" & rstDebitCreditVoucherParent.Fields("Code").Value & "'") Then Exit Sub
                .GetText 6, .ActiveRow, Item
                On Error Resume Next
                FrmAccountMaster.SL = True
                FrmAccountMaster.MasterCode = Item
                Load FrmAccountMaster
                If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText 6, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then
                    'rstItemList.MoveFirst: rstItemList.Find "[Code] ='" & slCode & "'"
                    rstJEAccountList.MoveFirst: rstJEAccountList.Find "[Code] ='" & slCode & "'"
                    LoadMasterList
                    .SetFocus
                    Sendkeys "{ENTER}"
                End If
            End If
    ElseIf KeyCode = vbKeySpace Then
        Dim Account As Variant
            CalculateTotal
            If .ActiveCol = 1 Then  'Switch TOA  Value D/C
            .GetText 1, .ActiveRow, cVal(1)  'TOA
            'Switch TOA  Value D/C
                If cVal(1) = "C" Or cVal(1) = "" Then
                .SetText 1, .ActiveRow, "D"
                ElseIf cVal(1) = "D" Or cVal(1) = "" Then
                .SetText 1, .ActiveRow, "C"
                End If
            ElseIf .ActiveCol = 2 Then 'Select Account
                .GetText 1, .ActiveRow, cVal(1)  'TOA
                .GetText .ActiveCol, .ActiveRow, Account
                txtAccount.Text = FixQuote(Account)
                    If cVal(1) = "" And dDCFlag <= 0 Then
                        cVal(1) = "D"
                    ElseIf cVal(1) = "" And dDCFlag > 0 Then
                        cVal(1) = "C": fpSpread1.SetText 1, fpSpread1.ActiveRow, "C"
                    End If
                    
                    If cVal(1) = "D" And VchType = "PI" Or cVal(1) = "C" And VchType = "PR" Then
                        If rstAccountList.RecordCount = 0 Then DisplayError ("No Record in Size Master"): .SetActiveCell 1, .ActiveRow: .SetFocus: Exit Sub Else rstAccountList.MoveFirst
                        rstAccountList.Find "[Col0] = '" & FixQuote(Trim(Account)) & "'"
                        SelectionType = "S": AccountCode = ""
                    Call LoadSelectionList(rstAccountList, "List of Accounts...", "Name")
                    
                    ElseIf cVal(1) = "C" And VchType = "PI" Or cVal(1) = "D" And VchType = "PR" Or cVal(1) = "D" And VchType = "CE" Or cVal(1) = "C" And VchType = "CE" Then
                        If rstSelfAccountList.RecordCount = 0 Then DisplayError ("No Record in Size Master"): .SetActiveCell 1, .ActiveRow: .SetFocus: Exit Sub Else rstSelfAccountList.MoveFirst
                        rstSelfAccountList.Find "[Col0] = '" & FixQuote(Trim(Account)) & "'"
                        SelectionType = "S": AccountCode = ""
                        Call LoadSelectionList(rstSelfAccountList, "List of Accounts...", "Name")
                    
                    ElseIf cVal(1) = "C" And VchType = "JE" Or cVal(1) = "D" And VchType = "JE" Or cVal(1) = "C" And VchType = "DN" Or cVal(1) = "D" And VchType = "DN" Or cVal(1) = "C" And VchType = "CN" Or cVal(1) = "D" And VchType = "CN" Then
                        If rstJEAccountList.RecordCount = 0 Then DisplayError ("No Record in Size Master"): .SetActiveCell 1, .ActiveRow: .SetFocus: Exit Sub Else rstJEAccountList.MoveFirst
                        rstJEAccountList.Find "[Col0] = '" & FixQuote(Trim(Account)) & "'"
                        SelectionType = "S": AccountCode = ""
                        Call LoadSelectionList(rstJEAccountList, "List of Accounts...", "Name")
                    End If
                    
                SearchOrder = 0
                Call DisplaySelectionList(txtAccount, AccountCode)
                Call CloseForm(FrmSelectionList)
                If AccountCode = "" Then
                    .SetText 2, .ActiveRow, txtAccount.Text: .SetActiveCell 2, .ActiveRow
                ElseIf .ActiveCol = 2 Then
                    .SetText 2, .ActiveRow, txtAccount.Text
                    .SetText 6, .ActiveRow, AccountCode
                    .SetFocus
                    Sendkeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 2 Then
                If AccountCode = "" Then .SetText 2, .ActiveRow, txtAccount.Text: .SetActiveCell 2, .ActiveRow
                If AccountCode <> "" Then .SetText 2, .ActiveRow, txtAccount.Text: .SetText 6, .ActiveRow, AccountCode
                    .SetFocus
                    Sendkeys "{ENTER}"
            End If
    End If
End With
End Sub
Private Sub fpSpread1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    Dim TOA As Variant, Account As Variant, Debit As Variant, Credit As Variant, Col As Long
    With fpSpread1
                .GetText 1, Row, TOA
                .GetText 2, Row, Account
                .GetText 3, Row, Debit
                .GetText 4, Row, Credit
        If Col = 1 And Row > 1 Then 'TOA
            If dDCFlag <= 0 And TOA = "" Then
                .SetText 1, Row, "D": .SetActiveCell 2, .ActiveRow: CalculateTotal
            ElseIf dDCFlag > 0 And TOA = "" Then
                .SetText 1, Row, "C": .SetActiveCell 2, .ActiveRow: CalculateTotal
            End If
        End If
    End With
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim TOA As Variant, Account As Variant, Debit As Variant, Credit As Variant, cVal(1 To 8) As Variant
    With fpSpread1
                .GetText 1, Row, TOA
                .GetText 2, Row, Account
                .GetText 3, Row, Debit
                .GetText 4, Row, Credit
        If Col = 1 And Row = 1 Then 'TOA
                If TOA = "" Then .SetText 1, Row, "D": .SetText 4, Row, 0: CalculateTotal
        ElseIf Col = 2 Then  'Accounts
                If TOA = "C" And Not CheckEmpty(Account, False) And Credit = 0 Then .SetActiveCell 4, .ActiveRow: .SetText 4, Row, IIf(dDCFlag < 0, 0, dDCFlag): CalculateTotal
                If TOA = "D" And Not CheckEmpty(Account, False) And Debit = 0 Then .SetActiveCell 3, .ActiveRow: .SetText 3, Row, IIf(dDCFlag < 0, Abs(dDCFlag), 0): CalculateTotal
        ElseIf Col = 3 Then  'Debit
                If TOA = "D" And Debit > 0 Then .SetText 4, Row, 0: CalculateTotal
                If Col = 3 And TOA = "D" And Debit <= 0 And Not CheckEmpty(Account, False) Then .SetActiveCell 3, .ActiveRow: CalculateTotal
        ElseIf Col = 4 Then  'Credit
                If TOA = "C" And Credit > 0 Then .SetText 3, Row, 0: CalculateTotal
                If Col = 4 And TOA = "C" And Credit <= 0 And Not CheckEmpty(Account, False) Then .SetActiveCell 4, .ActiveRow: CalculateTotal:
        ElseIf Col = 5 Then  'TOA
        .GetText 1, .ActiveRow + 1, cVal(1) 'Next Row TOA
            If dDCFlag = 0 Then
            fpSpread1.SetActiveCell .ActiveCol, .ActiveRow: Text4.SetFocus
            ElseIf dDCFlag <= 0 And cVal(1) = "" Then
                .SetText 1, Row + 1, "D": .SetActiveCell 2, .ActiveRow + 1: CalculateTotal
            ElseIf dDCFlag > 0 And cVal(1) = "" Then
                .SetText 1, Row + 1, "C": .SetActiveCell 2, .ActiveRow + 1: CalculateTotal
            End If
        End If
    End With
End Sub
Private Sub CalculateTotal()
    Dim i As Integer, Debit As Variant, Credit As Variant
    MhRealInput1.Value = 0: MhRealInput2.Value = 0
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 3, i, Debit
            .GetText 4, i, Credit
            MhRealInput1.Value = MhRealInput1.Value + Val(Debit)
            MhRealInput2.Value = MhRealInput2.Value + Val(Credit)
        Next
            oDebit = MhRealInput1.Value: oCredit = MhRealInput2.Value: dDCFlag = oDebit - oCredit
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
    End Sub
Private Sub DuplicateRecord()
    Dim Tbl As String
    Tbl = "T" & GetFileNameFromPath(GetTemporaryFileName()): Tbl = Left(Tbl, InStr(1, Tbl, ".", vbTextCompare) - 1)
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    Dim VchCode As String, VchNo As String
    VchCode = GenerateCode(cnDebitCreditVoucher, "SELECT MAX(Code) FROM DebitCreditParent", 6, "0")
    VchNo = GenerateCode(cnDebitCreditVoucher, "SELECT MAX(VAL(Name)) FROM DebitCreditParent WHERE RIGHT(Type,2)='" & VchType & "'", 10, Space(1))
    With cnDebitCreditVoucher
        .BeginTrans
        .Execute "SELECT * INTO " & Tbl & " FROM DebitCreditParent Where Code = '" & rstDebitCreditVoucherList.Fields("Code").Value & "'"
        .Execute "UPDATE " & Tbl & " SET Code='" & VchCode & "',Name='" & Pad(Trim(VchNo), Space(1), 10, "L") & "',[Date]=NOW()"
        .Execute "INSERT INTO DebitCreditParent SELECT * FROM " & Tbl
        .Execute "DROP TABLE " & Tbl
        .Execute "SELECT * INTO " & Tbl & " FROM DebitCreditChild Where Code = '" & rstDebitCreditVoucherList.Fields("Code").Value & "'"
        .Execute "UPDATE " & Tbl & " SET Code='" & VchCode & "'"
        .Execute "UPDATE " & Tbl & " SET RefCode=''"
        .Execute "INSERT INTO DebitCreditChild SELECT * FROM " & Tbl
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
    cnDebitCreditVoucher.RollbackTrans
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
Public Sub PrintDebitCreditVoucher(ByVal VchCode As String, ByVal VchType As String, Optional ByVal OutputType As String)
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstDebitCreditVoucherChild.State = adStateOpen Then rstDebitCreditVoucherChild.Close
    If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,EMail,Website,GSTIN,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,Prefix,Suffix FROM CompanyMaster P INNER JOIN CompChild C ON P.Code=C.Code WHERE VchType= " & IIf(Right(VchType, 2) = "PI", 51, IIf(Right(VchType, 2) = "PR", 52, IIf(Right(VchType, 2) = "JE", 53, IIf(Right(VchType, 2) = "CE", 54, IIf(Right(VchType, 2) = "DN", 55, 56))))), cnDebitCreditVoucher, adOpenKeyset, adLockReadOnly
    rstCompanyMaster.ActiveConnection = Nothing
    
    rstDebitCreditVoucherChild.Open "SELECT '" & LTrim(rstCompanyMaster.Fields("Prefix").Value) & "'+LTRIM(T.Name)+'" & LTrim(rstCompanyMaster.Fields("Suffix").Value) & "' As BillNo,T.Code As Code,T.Name,V.Name As VchSeriesName,Date,A.Name As Account, C.Debit,C.Credit,ShortNarration,LongNarration,T.Type,C.TOA,T.Debit As TotalDebit,T.Credit As TotalCredit " & _
                                "FROM DebitCreditParent T INNER JOIN DebitCreditChild C On C.Code=T.Code INNER JOIN VchSeriesMaster V ON T.VchSeries=V.Code INNER JOIN AccountMaster A On C.Account=A.Code WHERE T.Code='" + Left(VchCode, 6) + "' ORDER BY TOA DESC", cnDebitCreditVoucher, adOpenKeyset, adLockOptimistic
    
    If rstDebitCreditVoucherChild.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rstDebitCreditVoucherChild.ActiveConnection = Nothing
    rptDebitCreditVoucher.Text1.SetText IIf(VchType = "PI", "Payment Voucher", IIf(VchType = "PR", "Receipt Voucher", IIf(VchType = "JE", "Journal Voucher", IIf(VchType = "CE", "Contra Voucher", IIf(VchType = "DN", "Debit Note Voucher", "Credit Note Voucher")))))
    rptDebitCreditVoucher.Text35.SetText "Printed on " & Format(Now, "dd-MMM-yyyy") & " at " & Format(Now, "hh:mm")
    'rptDebitCreditVoucher.Text40.SetText IIf(BillType = "O", "(ORIGINAL FOR RECIPIENT)", IIf(BillType = "D", "(DUPLICATE FOR SUPPLIER)", "(TRIPLICATE FOR SUPPLIER)"))
    If Len(LTrim(rstCompanyMaster.Fields("PrintName").Value)) <> 25 Then rptDebitCreditVoucher.Text2.Font = 48
    rptDebitCreditVoucher.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptDebitCreditVoucher.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False)) Then
        rptDebitCreditVoucher.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value) & Space(1) & "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptDebitCreditVoucher.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False) Then
        rptDebitCreditVoucher.Text24.SetText "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    End If
    rptDebitCreditVoucher.Text18.SetText "(" & UCase(Trim(NumberToWords(rstDebitCreditVoucherChild.Fields("TotalCredit").Value, False))) & ")"
    'rptDebitCreditVoucher.Text2.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptDebitCreditVoucher.Database.SetDataSource rstDebitCreditVoucherChild, 3, 1
    rptDebitCreditVoucher.DiscardSavedData
    Screen.MousePointer = vbNormal
    If OutputType = "S" Then
        Set FrmReportViewer.Report = rptDebitCreditVoucher
        FrmReportViewer.Show vbModal
    Else
        If rstDebitCreditVoucherList.State = adStateClosed Then  'For Print Utility
            rptDebitCreditVoucher.PaperSource = crPRBinAuto
            rptDebitCreditVoucher.PrintOut False
        Else
            rptDebitCreditVoucher.PaperSource = crPRBinAuto
            rptDebitCreditVoucher.PrintOut
        End If
    End If
    Set rptDebitCreditVoucher = Nothing
    If rstDebitCreditVoucherList.State = adStateClosed Then  'For Print Utility
        Call CloseRecordset(rstCompanyMaster)
    End If
    Call CloseRecordset(rstDebitCreditVoucherChild)
    On Error GoTo 0
End Sub
Private Sub LoadMasterList(Optional ByVal LoadSelected As Boolean)
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "SELECT Name As Col0,Code FROM AccountMaster Where [Group] NOT IN ('*26007','*26004') ORDER BY Name", cnDebitCreditVoucher, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    
    If rstSelfAccountList.State = adStateOpen Then rstSelfAccountList.Close
    rstSelfAccountList.Open "SELECT Name As Col0,Code FROM AccountMaster Where [Group] IN ('*26007','*26004') ORDER BY Name", cnDebitCreditVoucher, adOpenKeyset, adLockReadOnly
    rstSelfAccountList.ActiveConnection = Nothing
    
    If rstJEAccountList.State = adStateOpen Then rstJEAccountList.Close
    rstJEAccountList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnDebitCreditVoucher, adOpenKeyset, adLockReadOnly
    rstJEAccountList.ActiveConnection = Nothing
    
    If rstVchSeriesList.State = adStateOpen Then rstVchSeriesList.Close
    rstVchSeriesList.Open "SELECT Name As Col0,Prefix,Suffix,VchNumbering,Code FROM VchSeriesMaster WHERE Left(FYCode,2)='" & Left(FYCode, 2) & "' AND VchType= '" & IIf(VchType = "PI", "51", IIf(VchType = "PR", "52", IIf(VchType = "JE", "53", IIf(VchType = "CE", "54", IIf(VchType = "CN", "55", "56"))))) & VchType & "' ORDER BY Name", cnDebitCreditVoucher, adOpenKeyset, adLockReadOnly
    rstVchSeriesList.ActiveConnection = Nothing
End Sub
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        'MnuHelp_Click (1)
    ElseIf Button.Index >= 6 And Button.Index <= 11 Then
        If Me.Name <> "frmDebitCreditVoucher" Then
            frmDebitCreditVoucher.VchType = Choose(Button.Index - 5, "PI", "PR", "JE", "CE", "DN", "CN")
            Load frmDebitCreditVoucher
            If Err.Number <> 364 Then frmDebitCreditVoucher.Show
            frmDebitCreditVoucher.Toolbar1_ButtonClick frmDebitCreditVoucher.Toolbar1.Buttons.Item(1)
        Else
            Call CloseForm(frmDebitCreditVoucher)
            frmDebitCreditVoucher.VchType = Choose(Button.Index - 5, "PI", "PR", "JE", "CE", "DN", "CN")
            Load frmDebitCreditVoucher
            If Err.Number <> 364 Then frmDebitCreditVoucher.Show
            frmDebitCreditVoucher.Toolbar1_ButtonClick frmDebitCreditVoucher.Toolbar1.Buttons.Item(1)
        End If
    
    End If
End Sub

