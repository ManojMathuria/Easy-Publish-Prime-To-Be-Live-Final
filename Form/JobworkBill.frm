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
Begin VB.Form frmJobworkBill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jobwork Bill"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16500
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16500
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9180
      Left            =   15
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   16470
      _Version        =   65536
      _ExtentX        =   29051
      _ExtentY        =   16192
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
      Picture         =   "JobworkBill.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8955
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Width           =   16245
         _ExtentX        =   28654
         _ExtentY        =   15796
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
         TabPicture(0)   =   "JobworkBill.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Mh3dLabel1(3)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Mh3dLabel1(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "DataGrid1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Text1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "JobworkBill.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dLabel1(1)"
         Tab(1).Control(1)=   "Mh3dFrame2"
         Tab(1).Control(2)=   "Mh3dLabel55"
         Tab(1).ControlCount=   3
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel55 
            Height          =   330
            Left            =   -61970
            TabIndex        =   51
            Top             =   7515
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
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
            Caption         =   "  -> Get Jobwork Paper Cost <-"
            Alignment       =   0
            BevelStyle      =   1
            BorderStyle     =   2
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "JobworkBill.frx":0054
            BevelStyleInside=   1
            Picture         =   "JobworkBill.frx":0070
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
            Left            =   840
            MaxLength       =   40
            TabIndex        =   19
            Top             =   8430
            Width           =   10740
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7905
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   450
            Width           =   16020
            _ExtentX        =   28258
            _ExtentY        =   13944
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
               DataField       =   "VchSeriesName"
               Caption         =   "Vch Series"
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
            BeginProperty Column03 
               DataField       =   "PartyName"
               Caption         =   "Buyer Name"
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
               Caption         =   "Consignee Name"
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
            BeginProperty Column05 
               DataField       =   "Amount"
               Caption         =   "Amount"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "IntegrationStatus"
               Caption         =   "Integration Status"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   "0.00"
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
                  Locked          =   -1  'True
                  ColumnWidth     =   2324.977
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1379.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1934.929
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   3374.929
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   3314.835
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1349.858
               EndProperty
               BeginProperty Column06 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1769.953
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8295
            Left            =   -74880
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   480
            Width           =   16020
            _Version        =   65536
            _ExtentX        =   28257
            _ExtentY        =   14631
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
            Picture         =   "JobworkBill.frx":008C
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
               Left            =   11760
               MaxLength       =   40
               TabIndex        =   3
               Top             =   120
               Width           =   4140
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
               TabIndex        =   0
               Top             =   120
               Width           =   2370
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   9960
               TabIndex        =   24
               Top             =   945
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
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
               Picture         =   "JobworkBill.frx":00A8
               Picture         =   "JobworkBill.frx":00C4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   9960
               TabIndex        =   47
               Top             =   630
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
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
               Picture         =   "JobworkBill.frx":00E0
               Picture         =   "JobworkBill.frx":00FC
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
               Left            =   11760
               MaxLength       =   40
               TabIndex        =   6
               Top             =   630
               Width           =   4140
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
               Left            =   7020
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   5
               Top             =   630
               Width           =   2970
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
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   7
               Top             =   945
               Width           =   3435
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   25
               Top             =   7050
               Width           =   15795
               _Version        =   65536
               _ExtentX        =   27861
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
               Picture         =   "JobworkBill.frx":0118
               Picture         =   "JobworkBill.frx":0134
               Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
                  Height          =   285
                  Left            =   9870
                  TabIndex        =   26
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   930
                  _Version        =   65536
                  _ExtentX        =   1640
                  _ExtentY        =   503
                  Calculator      =   "JobworkBill.frx":0150
                  Caption         =   "JobworkBill.frx":0170
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "JobworkBill.frx":01DC
                  Keys            =   "JobworkBill.frx":01FA
                  Spin            =   "JobworkBill.frx":0244
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
                  Left            =   11610
                  TabIndex        =   29
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1185
                  _Version        =   65536
                  _ExtentX        =   2090
                  _ExtentY        =   503
                  Calculator      =   "JobworkBill.frx":026C
                  Caption         =   "JobworkBill.frx":028C
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "JobworkBill.frx":02F8
                  Keys            =   "JobworkBill.frx":0316
                  Spin            =   "JobworkBill.frx":0360
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
               Left            =   5220
               MaxLength       =   25
               TabIndex        =   1
               Top             =   120
               Width           =   2250
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
               Left            =   11760
               MaxLength       =   40
               TabIndex        =   9
               Top             =   950
               Width           =   4140
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
               TabIndex        =   4
               Top             =   630
               Width           =   3435
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   0
               Left            =   4260
               TabIndex        =   22
               Top             =   120
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
               Caption         =   "  Vch No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":0388
               Picture         =   "JobworkBill.frx":03A4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   23
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
               Caption         =   " Party"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":03C0
               Picture         =   "JobworkBill.frx":03DC
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   8520
               TabIndex        =   2
               Top             =   120
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   582
               Calendar        =   "JobworkBill.frx":03F8
               Caption         =   "JobworkBill.frx":0510
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":057C
               Keys            =   "JobworkBill.frx":059A
               Spin            =   "JobworkBill.frx":05F8
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
               TabIndex        =   10
               Top             =   1500
               Width           =   15795
               _Version        =   524288
               _ExtentX        =   27861
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
               MaxCols         =   17
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "JobworkBill.frx":0620
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
               Index           =   0
               Left            =   7560
               TabIndex        =   28
               Top             =   120
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
               Caption         =   "  Vch Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":1503
               Picture         =   "JobworkBill.frx":151F
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   30
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
               Caption         =   " Tax Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":153B
               Picture         =   "JobworkBill.frx":1557
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   11865
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   7530
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calculator      =   "JobworkBill.frx":1573
               Caption         =   "JobworkBill.frx":1593
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":15FF
               Keys            =   "JobworkBill.frx":161D
               Spin            =   "JobworkBill.frx":1667
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
               ValueVT         =   208535557
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
               Height          =   330
               Left            =   10710
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   7530
               Width           =   1170
               _Version        =   65536
               _ExtentX        =   2064
               _ExtentY        =   582
               Calculator      =   "JobworkBill.frx":168F
               Caption         =   "JobworkBill.frx":16AF
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":171B
               Keys            =   "JobworkBill.frx":1739
               Spin            =   "JobworkBill.frx":1783
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
               ValueVT         =   208535557
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel35 
               Height          =   645
               Left            =   12945
               TabIndex        =   33
               Top             =   7530
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
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
               Caption         =   " Post-Tax Amount"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":17AB
               Picture         =   "JobworkBill.frx":17C7
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   645
               Left            =   14520
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   7530
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   1138
               Calculator      =   "JobworkBill.frx":17E3
               Caption         =   "JobworkBill.frx":1803
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":186F
               Keys            =   "JobworkBill.frx":188D
               Spin            =   "JobworkBill.frx":18D7
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
               ValueVT         =   208535557
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   645
               Left            =   1560
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   7530
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   1138
               Calculator      =   "JobworkBill.frx":18FF
               Caption         =   "JobworkBill.frx":191F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":198B
               Keys            =   "JobworkBill.frx":19A9
               Spin            =   "JobworkBill.frx":19F3
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
               TabIndex        =   36
               Top             =   7530
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
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
               Caption         =   " Pre-Tax Amount"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":1A1B
               Picture         =   "JobworkBill.frx":1A37
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
               Height          =   330
               Left            =   9630
               TabIndex        =   37
               Top             =   7845
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
               Caption         =   " SGST"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":1A53
               Picture         =   "JobworkBill.frx":1A6F
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
               Height          =   330
               Left            =   11865
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   7845
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calculator      =   "JobworkBill.frx":1A8B
               Caption         =   "JobworkBill.frx":1AAB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":1B17
               Keys            =   "JobworkBill.frx":1B35
               Spin            =   "JobworkBill.frx":1B7F
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
               ValueVT         =   208535557
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
               Height          =   330
               Left            =   10710
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   7845
               Width           =   1170
               _Version        =   65536
               _ExtentX        =   2064
               _ExtentY        =   582
               Calculator      =   "JobworkBill.frx":1BA7
               Caption         =   "JobworkBill.frx":1BC7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":1C33
               Keys            =   "JobworkBill.frx":1C51
               Spin            =   "JobworkBill.frx":1C9B
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
               ValueVT         =   208535557
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   645
               Left            =   6195
               TabIndex        =   12
               Top             =   7530
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   1138
               Calculator      =   "JobworkBill.frx":1CC3
               Caption         =   "JobworkBill.frx":1CE3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":1D4F
               Keys            =   "JobworkBill.frx":1D6D
               Spin            =   "JobworkBill.frx":1DB7
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
               ValueVT         =   208535557
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   645
               Left            =   5235
               TabIndex        =   40
               Top             =   7530
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
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
               Picture         =   "JobworkBill.frx":1DDF
               Picture         =   "JobworkBill.frx":1DFB
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   645
               Left            =   2520
               TabIndex        =   41
               Top             =   7530
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
               Picture         =   "JobworkBill.frx":1E17
               Picture         =   "JobworkBill.frx":1E33
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   645
               Left            =   3360
               TabIndex        =   11
               Top             =   7530
               Width           =   930
               _Version        =   65536
               _ExtentX        =   1640
               _ExtentY        =   1138
               Calculator      =   "JobworkBill.frx":1E4F
               Caption         =   "JobworkBill.frx":1E6F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":1EDB
               Keys            =   "JobworkBill.frx":1EF9
               Spin            =   "JobworkBill.frx":1F43
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
               ValueVT         =   208535557
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   9630
               TabIndex        =   42
               Top             =   7530
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
               Caption         =   " IGST/CGST"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":1F6B
               Picture         =   "JobworkBill.frx":1F87
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   645
               Left            =   4275
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   7530
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   1138
               Calculator      =   "JobworkBill.frx":1FA3
               Caption         =   "JobworkBill.frx":1FC3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":202F
               Keys            =   "JobworkBill.frx":204D
               Spin            =   "JobworkBill.frx":2097
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
               ValueVT         =   208601093
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   645
               Left            =   7395
               TabIndex        =   43
               Top             =   7530
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
               Picture         =   "JobworkBill.frx":20BF
               Picture         =   "JobworkBill.frx":20DB
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   645
               Left            =   8475
               TabIndex        =   13
               Top             =   7530
               Width           =   1170
               _Version        =   65536
               _ExtentX        =   2064
               _ExtentY        =   1138
               Calculator      =   "JobworkBill.frx":20F7
               Caption         =   "JobworkBill.frx":2117
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":2183
               Keys            =   "JobworkBill.frx":21A1
               Spin            =   "JobworkBill.frx":21EB
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
               ValueVT         =   208601093
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Left            =   5220
               TabIndex        =   44
               Top             =   630
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
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
               Caption         =   "  Mat Centre"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":2213
               Picture         =   "JobworkBill.frx":222F
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   5220
               TabIndex        =   45
               Top             =   945
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
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
               Caption         =   " Billing Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":224B
               Picture         =   "JobworkBill.frx":2267
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   120
               TabIndex        =   48
               Top             =   120
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
               Picture         =   "JobworkBill.frx":2283
               Picture         =   "JobworkBill.frx":229F
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   9960
               TabIndex        =   49
               Top             =   120
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
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
               Picture         =   "JobworkBill.frx":22BB
               Picture         =   "JobworkBill.frx":22D7
            End
            Begin MSForms.ComboBox cmbBillingType 
               Height          =   330
               Left            =   7020
               TabIndex        =   8
               Top             =   945
               Width           =   2970
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "5239;582"
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
               X2              =   16080
               Y1              =   7425
               Y2              =   7425
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   16080
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   16080
               Y1              =   1365
               Y2              =   1365
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   11565
            TabIndex        =   46
            Top             =   8430
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
            Picture         =   "JobworkBill.frx":22F3
            Picture         =   "JobworkBill.frx":230F
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   3
            Left            =   13800
            TabIndex        =   50
            Top             =   0
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
            Caption         =   " Alt+V-> Preveiw Voucher"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "JobworkBill.frx":232B
            Picture         =   "JobworkBill.frx":2347
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   1
            Left            =   -67440
            TabIndex        =   52
            Top             =   0
            Width           =   8775
            _Version        =   65536
            _ExtentX        =   15478
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
            Caption         =   " F6->Edit Reference Order  F7->Get Revised Unit Rates  Ctrl+E->Edit Voucher  F9->Delete Row Ctrl+S->Save"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "JobworkBill.frx":2363
            Picture         =   "JobworkBill.frx":237F
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
            TabIndex        =   20
            Top             =   8430
            Width           =   735
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   16500
      _ExtentX        =   29104
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
Attribute VB_Name = "frmJobworkBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Vch Type=NNNNSU/NNNNSC/NNNNSJ/NNNNPU/NNNNPC/NNNNPJ (U-Unit Cost C-Jobwork Unit Cost J-Jobwork P-Purchase S-Sale Q-Sales Quotation Z-Purchase Quotation) & BOM=NNNNXXXXXXXXXXXXFI/NNNNXXXXXXXXXXXXMF (MF/ME/CF/MO/BN/BM) & 01-Purchase 04-Sale 23-Purchase Quotation 24-Sales Quotation
Public VchType As String, oVchType As String
Public dSortBy As Boolean, uRate As Variant, uRateMF As Variant, uRateME As Variant, uRateCF As Variant, uRateMO As Variant, uRateBN As Variant, uRateBM As Variant
Public vDate As Variant, vtCode, vtType As Variant
Dim cnJobworkBill As New ADODB.Connection, cnTally As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset, rstJobworkBVList As New ADODB.Recordset, rstJobworkBVParent As New ADODB.Recordset, rstJobworkBVChild As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstTaxList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstNarrationList As New ADODB.Recordset, rstHSNCodeList As New ADODB.Recordset, rstOrderList As New ADODB.Recordset, rstVchSeriesList As New ADODB.Recordset, rstMaterialCentreList As New ADODB.Recordset, rstSalesTypeList As New ADODB.Recordset
Dim BuyerCode As String, PartyStateCode As String, TaxCode As String, ItemCode As String, RefCode As String, NarrationCode As String, HSNCode As String, ConsigneeCode As String, VchPrefix As String, TranType As String, oVchNo As String, oVchDate As Date, oVchSeriesCode As String, AutoVchNo As String, SalesTypeCode As String
Dim SortOrder, PrevStr, dblBookMark As Double, blnRecordExist As Boolean, EditMode As Boolean, VchSeries As String, VchSeriesCode As String, MaterialCentreCode As String, VchNumbering As String
Dim frmDlvChDespatchDetails As New FrmDespatchDetails, frmDlvChConsigneeDetails As New FrmConsigneeDetails
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
'    Me.Left = (MdiMainMenu.ScaleWidth - Me.Width) \ 2
 '   Me.Left = (MdiMainMenu.ScaleHeight - Me.Height) \ 2
    WheelHook DataGrid1
    BusySystemIndicator True
    VchType = Choose(Val(VchType), "SU", "SC", "SJ", "PU", "PC", "PJ", "QU", "QC", "QJ", "ZU", "ZC", "ZJ")
    Me.Caption = Switch(Left(VchType, 1) = "S", "Sales ", Left(VchType, 1) = "P", "Purchase ", Left(VchType, 1) = "Q", "Sales Quotation ", Left(VchType, 1) = "Z", "Purchase Quotation ") + "(" + IIf(Right(VchType, 1) = "U", "Unit Cost", IIf(Right(VchType, 1) = "C", "Job Work Unit Cost", "Job Work")) + ")"
    Mh3dLabel15.Caption = IIf(Left(VchType, 1) = "S", " Sales Type ", " Purchase Type ")
    TranType = IIf(Left(VchType, 1) = "S", "04", IIf(Left(VchType, 1) = "P", "01", IIf(Left(VchType, 1) = "Q", "24", "23")))
    cnJobworkBill.CursorLocation = adUseClient: cnTally.CursorLocation = adUseClient
    cnJobworkBill.Open cnDatabase.ConnectionString
    LoadMasterList
    rstJobworkBVList.Open "SELECT T.Code,T.Name,Date,T.Type,P.Name As PartyName,C.Name As ConsigneeName,Amount,(Select Name From VchSeriesMaster Where Code=VchSeries) As VchSeriesName,(Select Name From AccountMaster Where Code=MaterialCentre) As MaterialCentre,(Select Name From AccountMaster Where Code=SalesType) As SalesType,IntegrationStatus FROM (JobworkBVParent T INNER JOIN AccountMaster P ON T.Party=P.Code) INNER JOIN AccountMaster C ON T.Consignee=C.Code INNER JOIN AccountMaster C1 ON T.MaterialCentre=C1.Code LEFT JOIN AccountMaster C2 ON T.SalesType=C2.Code WHERE RIGHT(Type,2)='" & VchType & "' AND T.FYCode='" & FYCode & "' ORDER BY T.Code", cnJobworkBill, adOpenKeyset, adLockPessimistic
    rstJobworkBVParent.CursorLocation = adUseClient
    rstJobworkBVList.Filter = adFilterNone
    If rstJobworkBVList.RecordCount > 0 Then rstJobworkBVList.MoveLast
    Set DataGrid1.DataSource = rstJobworkBVList
    BusySystemIndicator False
    SSTab1.Tab = 0
'    SortOrder = "Name"
     If FrmStockLedger.dSortBy = True Or FrmAccountLedger.dSortBy = True Then
        SortOrder = "Code"
    Else
        SortOrder = "AutoVchNo"
    End If
    If Not (rstJobworkBVList.EOF Or rstJobworkBVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstJobworkBVList.ActiveConnection = Nothing
    cmbBillingType.AddItem "Direct", 0 'Against Sales/Purchase Order or Direct
    cmbBillingType.AddItem "Against Challan", 1 'Against Sales/Purchase Challan
    SetButtonsForNoRecord
    fpSpread1.TextTip = TextTipFloating
    If InStr(1, "Q_Z", Left(VchType, 1)) > 0 Then
        Mh3dLabel15.Visible = False: Text10.Visible = False
        'Mh3dLabel5.Left = 4620: Mh3dLabel5.Width = 1215: Text2.Left = 5820:
        Text2.Width = 4650: Mh3dLabel1(0).Width = 1815
        Mh3dLabel1(0).Left = 9960: MhDateInput1.Left = 11760: 'MhDateInput1.Width = 1815
        Mh3dLabel6.Visible = False: cmbBillingType.Visible = False: Text5.Width = 9000 '3425 '7455
        cmbBillingType.ListIndex = 0
    End If
    Load frmDlvChDespatchDetails
    Load frmDlvChConsigneeDetails
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    With MdiMainMenu
        .mnuSalesJW.Enabled = False: .mnuPurchaseJW.Enabled = False: .mnuSalesQuotationJW.Enabled = False: .mnuPurchaseQuotationJW.Enabled = False
    End With
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        EditMode = False
        If SSTab1.Tab = 0 Then
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
    'Edit PO/SO
    ElseIf (Shift = 0 And KeyCode = vbKeyF6) And (Left(VchType, 1) = "S" Or Left(VchType, 1) = "P" Or Left(VchType, 1) = "Q" Or Left(VchType, 1) = "Z") Then
            vDate = MhDateInput1.Value: vtType = IIf(Left(VchType, 1) = "S", "FS", "FP"): fpSpread1.GetText 14, fpSpread1.ActiveRow, vtCode: vtCode = Left(vtCode, 6)
            If vDate = "" Then
                Exit Sub
            ElseIf FinancialYearFrom > vDate Or vDate = "" Then
                If MsgBox("You Can't Open Previous Financial Voucher in Current Year,... To Open This Voucher, Please Switch Financial Year ", vbCritical, "   Switch Financial Year !!!") = vbOK Then Exit Sub
            ElseIf vtType = "FP" Or vtType = "FS" Then
            dSortBy = True
                    On Error Resume Next
                    FrmBookPrintOrder.BookPOType = vtType
                    Load FrmBookPrintOrder
                    If Err.Number <> 364 Then FrmBookPrintOrder.Show 'vbModal
                    FrmBookPrintOrder.Text1 = vtCode
                    KeyCode = vbKeyE
                If Shift = 0 And KeyCode = vbKeyReturn Then 'View
                    FrmBookPrintOrder.SSTab1.Tab = 1
                ElseIf Shift = 0 And KeyCode = vbKeyE Then 'Edir
                    FrmBookPrintOrder.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(2)
                ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'Delete
                    FrmBookPrintOrder.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(3)
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    FrmBookPrintOrder.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                End If
            End If
    'Get Revised Unit Rates
    ElseIf (Shift = 0 And KeyCode = vbKeyF7) Then
        Dim Item As Variant, Qty As Variant, Rate As Variant
        With fpSpread1
        .GetText 3, .ActiveRow, Item
        If uRate <> 0 And Right(Item, 2) = "FI" Then fpSpread1.SetText 5, fpSpread1.ActiveRow, uRate
        If uRateMF <> 0 And Right(Item, 2) = "MF" Then fpSpread1.SetText 5, fpSpread1.ActiveRow, uRateMF
        If uRateME <> 0 And Right(Item, 2) = "ME" Then fpSpread1.SetText 5, fpSpread1.ActiveRow, uRateME
        If uRateCF <> 0 And Right(Item, 2) = "CF" Then fpSpread1.SetText 5, fpSpread1.ActiveRow, uRateCF
        If uRateMO <> 0 And Right(Item, 2) = "MO" Then fpSpread1.SetText 5, fpSpread1.ActiveRow, uRateMO
        If uRateBN <> 0 And Right(Item, 2) = "BN" Then fpSpread1.SetText 5, fpSpread1.ActiveRow, uRateBN
        If uRateBM <> 0 And Right(Item, 2) = "BM" Then fpSpread1.SetText 5, fpSpread1.ActiveRow, uRateBM
        fpSpread1.SetActiveCell 5, fpSpread1.ActiveRow
        
            If .ActiveCol = 4 Or .ActiveCol = 5 Then 'Quantity & Rate
                .GetText 1, .ActiveRow, Item
                .GetText 4, .ActiveRow, Qty
                .GetText 5, .ActiveRow, Rate
                If Not CheckEmpty(Item, False) Then .SetText 6, .ActiveRow, Qty * Rate: CalculateTotal Else .SetText 4, .ActiveRow, "": .SetText 5, .ActiveRow, "": .SetText 6, .ActiveRow, ""
            End If
        End With
        dSortBy = False: uRate = 0: uRateMF = 0: uRateME = 0: uRateCF = 0: uRateMO = 0: uRateBN = 0: uRateBM = 0
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
Private Sub Mh3dLabel55_Click()
Call GetJobWorkPaperCost
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    Dim CellVal As Variant, i As Integer
    If Button.Index = 1 Then
        If rstJobworkBVParent.State = adStateOpen Then rstJobworkBVParent.Close
        rstJobworkBVParent.Open "SELECT * FROM JobworkBVParent WHERE Code=''", cnJobworkBill, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstJobworkBVParent) Then
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text9.SetFocus
            blnRecordExist = False
            cnJobworkBill.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstJobworkBVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstJobworkBVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnJobworkBill.BeginTrans
            With rstJobworkBVChild
                If .State = adStateOpen Then
                    If .RecordCount > 0 Then .MoveFirst
                    Do While Not .EOF
                        If Not CheckEmpty(.Fields("VchCode").Value, False) Then Call UpdateStatus(.Fields("VchCode").Value, .Fields("Quantity").Value, "-")
                        .MoveNext
                    Loop
                End If
            End With
            cnJobworkBill.Execute "DELETE FROM JobworkBVParent WHERE Code='" & rstJobworkBVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstJobworkBVList.Delete
                rstJobworkBVList.MoveNext
                If rstJobworkBVList.RecordCount > 0 And rstJobworkBVList.EOF Then rstJobworkBVList.MoveLast
                cnJobworkBill.CommitTrans
                ShowProgressInStatusBar True
                Timer1.Enabled = True
                Text1.Text = ""
                rstJobworkBVList.Filter = adFilterNone
                If Left(VchType, 1) = "S" Then If BusyIntegration Or TallyIntegration Then DelOldVch True
            Else
                DisplayError ("Failed to delete the record")
                cnJobworkBill.RollbackTrans
            End If
            On Error GoTo 0
        End If
        SetButtons (True)
        SetButtonsForNoRecord
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 4 Then
        If CheckMandatoryFields Then Exit Sub
        frmDlvChDespatchDetails.Show vbModal
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstJobworkBVParent) Then
            If UpdateItemList("D", 0) Then
                UpdateFlag = 1
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 1, i
                        .GetText 6, i, CellVal 'Amount
                        If Val(CellVal) <> 0 Then
                            If Not UpdateItemList("I", i) Then UpdateFlag = 0: Exit For
                        End If
                    Next
                End With
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnJobworkBill.CommitTrans
            If rstJobworkBVParent.State = adStateOpen Then rstJobworkBVParent.Close
            rstJobworkBVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            Call MsgBox("Record updated !!!", vbInformation, App.Title)
            If Left(VchType, 1) = "S" Then
                    If MsgBox("Are you Wants to Generate the JSON File For 'e-Invoice' ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Export !") = vbYes Then Get_JSON ("INV")
                If BusyIntegration Then
                    If MsgBox("Are you sure to export the Voucher in Busy?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Export !") = vbYes Then PushVch
                ElseIf TallyIntegration Then
                    If MsgBox("Are you sure to export the Voucher in Tally?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Export !") = vbYes Then PushVch
                End If
            End If
            SSTab1.Tab = 0
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstJobworkBVParent) Then
            cnJobworkBill.RollbackTrans
            If rstJobworkBVParent.State = adStateOpen Then rstJobworkBVParent.Close
            rstJobworkBVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstJobworkBVList.Filter = adFilterNone
        rstJobworkBVList.ActiveConnection = cnJobworkBill
        Do While Not RefreshRecord(rstJobworkBVList): Loop
        Set DataGrid1.DataSource = rstJobworkBVList
        rstJobworkBVList.ActiveConnection = Nothing
        If rstJobworkBVList.RecordCount > 0 Then rstJobworkBVList.MoveLast
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
        If rstJobworkBVList.RecordCount = 0 Then Exit Sub
        DisplayMenu "P"
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstJobworkBVList.RecordCount = 0 Then Exit Sub
        DisplayMenu "S"
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstJobworkBVList.RecordCount > 0 Then rstJobworkBVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstJobworkBVList.RecordCount > 0 Then
            rstJobworkBVList.MovePrevious
            If rstJobworkBVList.BOF Then rstJobworkBVList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstJobworkBVList.RecordCount > 0 Then
            rstJobworkBVList.MoveNext
            If rstJobworkBVList.EOF Then rstJobworkBVList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstJobworkBVList.RecordCount > 0 Then rstJobworkBVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstJobworkBVList.EOF Or rstJobworkBVList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        Text1.SetFocus
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstJobworkBVList)
    Call CloseRecordset(rstJobworkBVParent)
    Call CloseRecordset(rstJobworkBVChild)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstMaterialCentreList)
    Call CloseRecordset(rstSalesTypeList)
    Call CloseRecordset(rstTaxList)
    Call CloseRecordset(rstNarrationList)
    Call CloseRecordset(rstHSNCodeList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstVchSeriesList)
    Call CloseRecordset(rstOrderList)
    Call CloseConnection(cnJobworkBill)
    Call CloseConnection(cnTally)
    Call CloseForm(frmDlvChDespatchDetails)
    Call CloseForm(frmDlvChConsigneeDetails)
    ShowProgressInStatusBar False
    DisableChildMenu
    With MdiMainMenu
        .mnuSalesJW.Enabled = True: .mnuPurchaseJW.Enabled = True: .mnuSalesQuotationJW.Enabled = True: .mnuPurchaseQuotationJW.Enabled = True
    End With
End Sub
Private Sub Text1_Change()
    On Error Resume Next
    With rstJobworkBVList
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
    If rstJobworkBVList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstJobworkBVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstJobworkBVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstJobworkBVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstJobworkBVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstJobworkBVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstJobworkBVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstJobworkBVList
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
            If Not (rstJobworkBVList.EOF Or rstJobworkBVList.BOF) Then
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
        Text9.SetFocus
    End If
End Sub
Private Sub ClearFields()
    Text9.Text = "": VchSeriesCode = ""
    Text2.Text = "" 'Vch No.
    Text3.Text = "": BuyerCode = "" 'Party Name
'    Text7.Text = "": ConsigneeCode = "" 'Consignee Name
    Text7.Text = "" 'Material Centre
    Text8.Text = "" 'Consignee
    Text10.Text = "" 'SalesType
    Text5.Text = "": TaxCode = "" 'Tax Name
    Text4.Text = "" 'Remarks
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
    cmbBillingType.ListIndex = 0: cmbBillingType.Enabled = True: cmbBillingType_Click
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    PartyStateCode = ""
    VchSeriesCode = "": oVchSeriesCode = "": oVchNo = "": AutoVchNo = ""
    frmDlvChDespatchDetails.Text1.Text = "": frmDlvChDespatchDetails.Text2.Text = "": frmDlvChDespatchDetails.Text3.Text = "": frmDlvChDespatchDetails.Text4.Text = "": frmDlvChDespatchDetails.MhDateInput1.Value = Null: frmDlvChDespatchDetails.Text5.Text = "": frmDlvChDespatchDetails.MhDateInput2.Value = Null
    frmDlvChConsigneeDetails.Text1.Text = "": frmDlvChConsigneeDetails.Text2.Text = "": frmDlvChConsigneeDetails.Text3.Text = "": frmDlvChConsigneeDetails.Text4.Text = "": frmDlvChConsigneeDetails.Text5.Text = "": frmDlvChConsigneeDetails.Text6.Text = ""
End Sub
Private Sub LoadFields()
    With rstJobworkBVParent
        If .EOF Or .BOF Then Exit Sub
        VchSeriesCode = .Fields("VchSeries").Value: oVchSeriesCode = VchSeriesCode
        If rstVchSeriesList.RecordCount > 0 Then rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        If Not rstVchSeriesList.EOF Then Text9.Text = rstVchSeriesList.Fields("Col0").Value
        AutoVchNo = Trim(.Fields("AutoVchNo").Value)
        oVchNo = Trim(.Fields("AutoVchNo").Value) 'Trim(Text2.Text)
        Text2.Text = .Fields("Name").Value
        oVchDate = Format(.Fields("Date").Value, "dd-MMM-yyyy")
        MhDateInput1.Text = Format(.Fields("Date").Value, "dd-MM-yyyy")
        BuyerCode = .Fields("Party").Value
        PartyStateCode = .Fields("State").Value
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        rstAccountList.Find "[Code] = '" & BuyerCode & "'"
        If Not rstAccountList.EOF Then Text3.Text = rstAccountList.Fields("Col0").Value
        MaterialCentreCode = .Fields("MaterialCentre").Value
        If rstMaterialCentreList.RecordCount > 0 Then rstMaterialCentreList.MoveFirst
        rstMaterialCentreList.Find "[Code] = '" & MaterialCentreCode & "'"
        If Not rstMaterialCentreList.EOF Then Text7.Text = rstMaterialCentreList.Fields("Col0").Value
        ConsigneeCode = .Fields("Consignee").Value
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        rstAccountList.Find "[Code] = '" & ConsigneeCode & "'"
        If Not rstAccountList.EOF Then Text8.Text = rstAccountList.Fields("Col0").Value
        
        SalesTypeCode = .Fields("SalesType").Value
        If rstSalesTypeList.RecordCount > 0 Then rstSalesTypeList.MoveFirst
        rstSalesTypeList.Find "[Code] = '" & SalesTypeCode & "'"
        If Not rstSalesTypeList.EOF Then Text10.Text = rstSalesTypeList.Fields("Col0").Value
        
        TaxCode = .Fields("Tax").Value
        If rstTaxList.RecordCount > 0 Then rstTaxList.MoveFirst
        rstTaxList.Find "[Code] = '" & TaxCode & "'"
        If Not rstTaxList.EOF Then Text5.Text = rstTaxList.Fields("Col0").Value
        
        Text4.Text = .Fields("Remarks").Value
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
        frmDlvChDespatchDetails.Text1.Text = CheckNull(.Fields("Transport").Value): frmDlvChDespatchDetails.Text2.Text = CheckNull(.Fields("GRNo").Value): frmDlvChDespatchDetails.Text3.Text = CheckNull(.Fields("VehicleNo").Value): frmDlvChDespatchDetails.Text4.Text = CheckNull(.Fields("Station").Value): If Not IsNull(.Fields("GRDate").Value) Then frmDlvChDespatchDetails.MhDateInput1.Value = .Fields("GRDate").Value: frmDlvChDespatchDetails.Text5.Text = CheckNull(.Fields("eWayBill").Value): If Not IsNull(.Fields("eWayBillDate").Value) Then frmDlvChDespatchDetails.MhDateInput2.Value = .Fields("eWayBillDate").Value
        'If Mid(.Fields("Type").Value, 5, 2) <> "ZU" Then cmbBillingType.ListIndex = IIf(Mid(.Fields("Type").Value, 3, 2) = "10", 0, 1)
        If InStr(1, "QU_QC_QJ_ZU_ZC_ZJ", Right(VchType, 2)) = 0 Then cmbBillingType.ListIndex = IIf(Mid(.Fields("Type").Value, 3, 2) = "10", 0, 1)
        Call LoadItemList(.Fields("Code").Value)
    End With
    CalculateTotal
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstJobworkBVParent.RecordCount = 0 Then Exit Sub
    If rstJobworkBVParent.State = adStateOpen Then rstJobworkBVParent.Close
    rstJobworkBVParent.CursorLocation = adUseServer
    rstJobworkBVParent.Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstJobworkBVList.Fields("Code").Value) & "'", cnJobworkBill, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstJobworkBVParent.Fields("RecordStatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    cmbBillingType.Enabled = False
    Text9.SetFocus
    blnRecordExist = True
    cnJobworkBill.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    With rstJobworkBVParent
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnJobworkBill, "SELECT MAX(Code) FROM JobworkBVParent", 6, "0")
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
        .Fields("Party").Value = BuyerCode
        .Fields("MaterialCentre").Value = MaterialCentreCode ' ""
        .Fields("Consignee").Value = ConsigneeCode
        .Fields("Tax").Value = TaxCode
        .Fields("Remarks").Value = Trim(Text4.Text)
        .Fields("Rebate%").Value = MhRealInput4.Value
        .Fields("Rebate").Value = MhRealInput5.Value
        .Fields("Freight").Value = MhRealInput6.Value
        .Fields("Adjustment").Value = MhRealInput12.Value
        .Fields("TaxableAmount").Value = MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput12.Value
        If MhRealInput9.Value > 0 Then 'Intra-State Supply
            .Fields("CGST%").Value = MhRealInput7.Value
            .Fields("CGST").Value = MhRealInput8.Value
            .Fields("SGST%").Value = MhRealInput9.Value
            .Fields("SGST").Value = MhRealInput10.Value
            .Fields("IGST%").Value = 0
            .Fields("IGST").Value = 0
        Else 'Inter-State Supply
            .Fields("CGST%").Value = 0
            .Fields("CGST").Value = 0
            .Fields("SGST%").Value = 0
            .Fields("SGST").Value = 0
            .Fields("IGST%").Value = MhRealInput7.Value
            .Fields("IGST").Value = MhRealInput8.Value
        End If
        .Fields("Amount").Value = MhRealInput11.Value
        .Fields("Type").Value = VchPrefix + VchType
        .Fields("FYCode").Value = FYCode
        .Fields("RecordStatus").Value = "N"
        .Fields("SalesType").Value = SalesTypeCode
        .Fields("Transport").Value = frmDlvChDespatchDetails.Text1.Text
        .Fields("GRNo").Value = frmDlvChDespatchDetails.Text2.Text
        If frmDlvChDespatchDetails.MhDateInput1.ValueIsNull Then .Fields("GRDate").Value = Null Else .Fields("GRDate").Value = GetDate(frmDlvChDespatchDetails.MhDateInput1.Text)
        .Fields("VehicleNo").Value = frmDlvChDespatchDetails.Text3.Text
        .Fields("Station").Value = frmDlvChDespatchDetails.Text4.Text
        .Fields("eWayBill").Value = frmDlvChDespatchDetails.Text5.Text
        If frmDlvChDespatchDetails.MhDateInput2.ValueIsNull Then .Fields("eWayBillDate").Value = Null Else .Fields("eWayBillDate").Value = GetDate(frmDlvChDespatchDetails.MhDateInput2.Text)
    End With
End Sub
Private Sub AddToList()
    On Error Resume Next
    With rstJobworkBVList
        .MoveFirst
        .Find "[Code] = '" & rstJobworkBVParent.Fields("Code").Value & "'"
        If .EOF Then .AddNew
        .Fields("Code").Value = rstJobworkBVParent.Fields("Code").Value
        .Fields("Name").Value = Trim(rstJobworkBVParent.Fields("Name").Value)
        .Fields("VchSeriesName").Value = Text9.Text
        .Fields("Date").Value = rstJobworkBVParent.Fields("Date").Value
        .Fields("PartyName").Value = Trim(Text3.Text)
        .Fields("MaterialCentre").Value = Trim(Text7.Text)
        .Fields("ConsigneeName").Value = Trim(Text8.Text)
        .Fields("SalesType").Value = Trim(Text10.Text)
        .Fields("Amount").Value = MhRealInput11.Value
        .Fields("Type").Value = rstJobworkBVParent.Fields("Type").Value
        .Update
        .Sort = SortOrder & " Asc"
        .Find "[Code] = '" & rstJobworkBVParent.Fields("Code").Value & "'"
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        DisplayError ("Voucher No. cannot be blank"): Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnJobworkBill, "JobworkBVParent", "Code", "[Name]+RIGHT(Type,2)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, rstJobworkBVParent.Fields("Code").Value, False, FYCode) Then
        Dim VchNo As String
        If Not blnRecordExist Then 'Vch-New
            If VchNumbering = "A" Then
                AutoVchNo = GenerateCode(cnJobworkBill, "SELECT MAX(CONVERT(INT,AutoVchNo)) FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 15, Space(1))
                VchNo = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                If Trim(VchNo) <> Trim(Text2.Text) Then DisplayError ("Vch No. changed from " & Trim(Text2.Text) & " to " & Trim(VchNo))
                Text2.Text = VchNo
            End If
        Else 'Vch-Old
            If VchSeriesCode <> oVchSeriesCode Then
                If VchNumbering = "A" Then
                    AutoVchNo = GenerateCode(cnJobworkBill, "SELECT MAX(CONVERT(INT,AutoVchNo)) FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 15, Space(1))
                    VchNo = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                    If Trim(VchNo) <> Trim(Text2.Text) Then DisplayError ("Vch No. changed from " & Trim(Text2.Text) & " to " & Trim(VchNo))
                    Text2.Text = VchNo
                End If
            End If
            Text2.SetFocus: CheckMandatoryFields = True: Exit Function
        End If
    ElseIf CheckEmpty(Text3.Text, False) Then 'Party
        Text3.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text7.Text, False) Then 'Material Centre
        Text7.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text8.Text, False) Then 'Consignee
        Text8.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text10.Text, False) And InStr(1, "S_P", Left(VchType, 1)) > 0 Then 'SalesType
        Text10.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text5.Text, False) Then 'Tax
        Text5.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf fpSpread1.DataRowCnt = 0 Then
        DisplayError ("Blank Voucher cannot be saved"): fpSpread1.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf fpSpread1.DataRowCnt > 0 Then
        Dim i As Integer, CellVal As Variant
        With fpSpread1
                For i = 1 To .DataRowCnt
                    .GetText 2, i, CellVal
                    If CheckEmpty(CellVal, False) Then DisplayError ("HSN Code at row #" & Trim(Str(i)) & " is blank"): CheckMandatoryFields = True: .SetFocus: .SetActiveCell 2, i: Exit Function
                    .GetText 7, i, CellVal
                    If CheckEmpty(CellVal, False) Then DisplayError ("Narration at row #" & Trim(Str(i)) & " is blank"): CheckMandatoryFields = True: .SetFocus: .SetActiveCell 7, i: Exit Function
                Next
        End With
    End If
End Function
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
    If rstJobworkBVList.RecordCount = 0 Then
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
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text9.Text)
        If rstVchSeriesList.RecordCount = 0 Then DisplayError ("No Record in Voucher Series Master"): Text9.SetFocus: Exit Sub Else rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": VchSeriesCode = ""
        Call LoadSelectionList(rstVchSeriesList, "List of Voucher Series...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text9, VchSeriesCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(VchSeriesCode) <> "" Then Sendkeys "{TAB}" Else Text9.Text = ""
    End If
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
    If CheckEmpty(Text9.Text, False) Then
        Cancel = True
    Else
        rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        VchNumbering = rstVchSeriesList.Fields("VchNumbering").Value
        If VchNumbering = "A" Then Text2.Locked = True Else Text2.Locked = False
        If Not blnRecordExist Then 'Vch-New
            If VchNumbering = "A" Then
                AutoVchNo = GenerateCode(cnJobworkBill, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 15, Space(1))
                Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
            End If
        Else 'Vch-Old
            If VchSeriesCode = oVchSeriesCode Then
                Text2.Text = Text2.Text 'oVchNo
            Else
                If VchNumbering = "A" Then
                    AutoVchNo = GenerateCode(cnJobworkBill, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                    Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                End If
            End If
        End If
    End If
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstJobworkBVList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstJobworkBVList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstJobworkBVList.EOF Or rstJobworkBVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstJobworkBVParent.EOF Or rstJobworkBVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnJobworkBill, "JobworkBVParent", "Code", "[Name]+RIGHT(Type,2)", Trim(Text2.Text) & VchType, rstJobworkBVParent.Fields("Code").Value, False, FYCode) Then
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
        FrmAccountMaster.MasterCode = BuyerCode
        FrmAccountMaster.StateCode = PartyStateCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        BuyerCode = slCode: Text3.Text = slName
        If Not IsNull(slStateCode) Then
        PartyStateCode = slStateCode
        End If
        If Not CheckEmpty(BuyerCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
'    If CheckEmpty(Text3.Text, False) Then Cancel = True
'    If CheckEmpty(Text7.Text, False) Then Text7.Text = Text3.Text: ConsigneeCode = BuyerCode
    If CheckEmpty(Text3.Text, False) Then Cancel = True
    If CheckEmpty(Text8.Text, False) Then Text8.Text = Text3.Text: ConsigneeCode = BuyerCode
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
Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = IIf(Left(VchType, 1) = "S", "*26027", "*26025")
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
        LoadMasterList
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
Private Sub MhRealInput4_Validate(Cancel As Boolean)
    CalculateTotal
End Sub
Private Sub MhRealInput6_Validate(Cancel As Boolean)
    CalculateTotal
End Sub
Private Sub MhRealInput12_Validate(Cancel As Boolean)   'Adjustment
    CalculateTotal
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Party" Then rstJobworkBVList.Filter = "[PartyName] Like '%" & SrchText & "%'"
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cVal As Variant
    With fpSpread1
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus: CalculateTotal
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then 'Item
                .GetText 14, .ActiveRow, cVal 'Ref
                If (Not CheckEmpty(cVal, False)) Or cmbBillingType.ListIndex = 1 Or Right(VchType, 1) = "J" Then Exit Sub 'Non-blank Ref OR Against Challan or VchType is SJ/PJ
                .GetText 1, .ActiveRow, cVal 'Item
                Text6.Text = FixQuote(cVal)
                If rstItemList.RecordCount = 0 Then DisplayError ("No record in Item Master"): .SetActiveCell 1, .ActiveRow: .SetFocus: Exit Sub Else rstItemList.MoveFirst
                rstItemList.Find "[Col0] = '" & FixQuote(Trim(cVal)) & "'"
                SelectionType = "S": ItemCode = ""
                Call LoadSelectionList(rstItemList, "List of Items...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text6, ItemCode)
                Call CloseForm(FrmSelectionList)
                If ItemCode = "" Then
                    .SetActiveCell 1, .ActiveRow
                Else
                    rstItemList.MoveFirst: rstItemList.Find "[Code] ='" & ItemCode & "'"
                    .SetText 1, .ActiveRow, Text6.Text
                    .SetText 15, .ActiveRow, ItemCode
                    .GetText 16, .ActiveRow, cVal 'HSN
                    If CheckEmpty(cVal, False) Then .SetText 2, .ActiveRow, rstItemList.Fields("HSNName").Value: .SetText 16, .ActiveRow, rstItemList.Fields("HSNCode").Value
                    .GetText 14, .ActiveRow, cVal 'Ref
                    If CheckEmpty(cVal, False) Then .SetText 5, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                    .SetFocus
                    Sendkeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 2 Then 'HSN
                .GetText 16, .ActiveRow, cVal 'HSN Code
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "18"
                FrmGeneralMaster.MasterCode = cVal
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 2, .ActiveRow, slName: .SetText 16, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then LoadMasterList: Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 7 Then 'Short Nattaion
                .GetText 13, .ActiveRow, cVal 'Short Narration
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "17"
                FrmGeneralMaster.MasterCode = cVal
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 7, .ActiveRow, slName: .SetText 13, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then LoadMasterList: Sendkeys "{ENTER}"
            End If
        ElseIf KeyCode = vbKeyF11 Then
            If .DataRowCnt = 0 Then LoadOrderList
        End If
        If .DataRowCnt > 0 Then cmbBillingType.Enabled = False Else cmbBillingType.Enabled = True
    End With
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Item As Variant, Qty As Variant, Rate As Variant
    With fpSpread1
        If Col = 4 Or Col = 5 Then 'Quantity & Rate
            .GetText 1, Row, Item
            .GetText 4, Row, Qty
            .GetText 5, Row, Rate
            If Not CheckEmpty(Item, False) Then .SetText 6, Row, Qty * Rate: CalculateTotal Else .SetText 4, Row, "": .SetText 5, Row, "": .SetText 6, Row, ""
        End If
    End With
End Sub
Private Sub fpSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As FPSpreadADO.TextTipFetchMultilineConstants, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim PendingQty As Variant
    fpSpread1.GetText 17, Row, PendingQty
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
Private Sub ViewRecord()
    ClearFields
    If rstJobworkBVList.EOF Then Exit Sub
    FindRecord
    LoadMasterList
    LoadFields
End Sub
Private Sub FindRecord()
    If rstJobworkBVParent.State = adStateOpen Then rstJobworkBVParent.Close
    rstJobworkBVParent.Open "SELECT *, (Select State From AccountMaster Where Code=Party) AS State FROM JobworkBVParent WHERE Code='" & FixQuote(rstJobworkBVList.Fields("Code").Value) & "'", cnJobworkBill, adOpenKeyset, adLockOptimistic
    If rstJobworkBVParent.RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
End Sub
Private Sub cmbBillingType_Click()
    VchPrefix = TranType & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "01", Choose(cmbBillingType.ListIndex + 1, "10", "01")) '10-Stock affected 01-Stock not affected
End Sub
Private Sub UpdateStatus(ByVal VchCode As String, ByVal Quantity As Long, ByVal Operation As String)
    If InStr(1, "FI", Right(VchCode, 2)) > 0 Then
        If cmbBillingType.ListIndex = 0 Then 'Direct
            cnJobworkBill.Execute "UPDATE BookPOParent SET BilledAllB=BilledAllB" & Operation & Trim(Quantity) & ",DeliveredQuantityB=DeliveredQuantityB" & Operation & Trim(Quantity) & " WHERE Code+'XXXXXXXXXXXXFI'='" & VchCode & "'"
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            cnJobworkBill.Execute "UPDATE BookPOParent SET BilledAllC=BilledAllC" & Operation & Trim(Quantity) & " WHERE Code+'XXXXXXXXXXXXFI'='" & VchCode & "'"
        End If
    End If
    If InStr(1, "FI_MF", Right(VchCode, 2)) > 0 Then
'        If cmbBillingType.ListIndex = 0 Then 'Direct
'            cnJobworkBill.Execute "UPDATE BookPOChild05 SET BilledMFB=BilledMFB" & Operation & Trim(Quantity) & ",DeliveredQuantityB=DeliveredQuantityB"& Operation & Trim(Quantity) & " WHERE (Code+Element+'XXXXXXMF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
'        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
'            cnJobworkBill.Execute "UPDATE BookPOChild05 SET BilledMFC=BilledMFC" & Operation & Trim(Quantity) & " WHERE (Code+Element+'XXXXXXMF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
'        End If
        If cmbBillingType.ListIndex = 0 Then 'Direct
            cnJobworkBill.Execute "UPDATE BookPOChild05 SET BilledMFB=BilledMFB" & Operation & Trim(Quantity) & ",DeliveredQuantityB=DeliveredQuantityB" & Operation & Trim(Quantity) & " WHERE (Code+Element+'XXXXXXMF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            cnJobworkBill.Execute "UPDATE BookPOChild05 SET BilledMFC=BilledMFC" & Operation & Trim(Quantity) & " WHERE (Code+Element+'XXXXXXMF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        End If
    End If
    If InStr(1, "FI_ME", Right(VchCode, 2)) > 0 Then
        If cmbBillingType.ListIndex = 0 Then 'Direct
            cnJobworkBill.Execute "UPDATE BookPOChild06 SET BilledMEB=BilledMEB" & Operation & Trim(Quantity) & ",DeliveredQuantityB=DeliveredQuantityB" & Operation & Trim(Quantity) & " WHERE (Code+Element+'XXXXXXME'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            cnJobworkBill.Execute "UPDATE BookPOChild06 SET BilledMEC=BilledMEC" & Operation & Trim(Quantity) & " WHERE (Code+Element+'XXXXXXME'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        End If
    End If
    If InStr(1, "FI_CF", Right(VchCode, 2)) > 0 Then
        If cmbBillingType.ListIndex = 0 Then 'Direct
            cnJobworkBill.Execute "UPDATE BookPOChild0901 SET BilledCFB=BilledCFB" & Operation & Trim(Quantity) & ",DeliveredQuantityB=DeliveredQuantityB" & Operation & Trim(Quantity) & " WHERE (Code+Book+'XXXXXXCF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            cnJobworkBill.Execute "UPDATE BookPOChild0901 SET BilledCFC=BilledCFC" & Operation & Trim(Quantity) & " WHERE (Code+Book+'XXXXXXCF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        End If
    End If
    If InStr(1, "FI_MO", Right(VchCode, 2)) > 0 Then
        If cmbBillingType.ListIndex = 0 Then 'Direct
            cnJobworkBill.Execute "UPDATE BookPOChild07 SET BilledMOB=BilledMOB" & Operation & Trim(Quantity) & ",DeliveredQuantityB=DeliveredQuantityB" & Operation & Trim(Quantity) & " WHERE (Code+Element+Operation+'MO'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            cnJobworkBill.Execute "UPDATE BookPOChild07 SET BilledMOC=BilledMOC" & Operation & Trim(Quantity) & " WHERE (Code+Element+Operation+'MO'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        End If
    End If
    If InStr(1, "FI_BN", Right(VchCode, 2)) > 0 Then
        If cmbBillingType.ListIndex = 0 Then 'Direct
            cnJobworkBill.Execute "UPDATE BookPOChild08 SET BilledBNB=BilledBNB" & Operation & Trim(Quantity) & ",DeliveredQuantityB=DeliveredQuantityB" & Operation & Trim(Quantity) & " WHERE (Code+'XXXXXX'+BinderyProcess+'BN'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            cnJobworkBill.Execute "UPDATE BookPOChild08 SET BilledBNC=BilledBNC" & Operation & Trim(Quantity) & " WHERE (Code+'XXXXXX'+BinderyProcess+'BN'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        End If
    End If
    If InStr(1, "FI_BM", Right(VchCode, 2)) > 0 Then
        If cmbBillingType.ListIndex = 0 Then 'Direct
            cnJobworkBill.Execute "UPDATE BookPOChild0801 SET BilledBMB=BilledBMB" & Operation & Trim(Quantity) & ",DeliveredQuantityB=DeliveredQuantityB" & Operation & Trim(Quantity) & " WHERE (Code+Item+'XXXXX'+Category+'BM'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            cnJobworkBill.Execute "UPDATE BookPOChild0801 SET BilledBMC=BilledBMC" & Operation & Trim(Quantity) & " WHERE (Code+Item+'XXXXX'+Category+'BM'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        End If
    End If
End Sub
Private Sub LoadItemList(ByVal VchNo As String)
    Dim i As Integer, SQL As String
    On Error GoTo ErrorHandler
    If rstJobworkBVChild.State = adStateOpen Then rstJobworkBVChild.Close
    If InStr(1, "U_C", Right(VchType, 1)) > 0 Then 'Unit/Jobwork Unit Cost
        SQL = "SELECT I.Name As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "R.EstQty01-R.DeliveredQuantityC-R.BilledAllB", "R.DeliveredQuantityC-R.BilledAllC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,CASE WHEN Ref IS NULL THEN '' ELSE R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) END As VchCode,CASE WHEN Ref IS NULL THEN '' ELSE LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) END As VchNo FROM (((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) LEFT JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,12)=R.Code+'XXXXXXXXXXXX') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "' AND RIGHT(T.BOM,2)='FI'"
    ElseIf Right(VchType, 1) = "J" Then 'Jobwork
        'MF
        SQL = "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.ActualQuantity-C.DeliveredQuantityC-C.BilledMFB", "C.DeliveredQuantityC-C.BilledMFC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild05 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+'XXXXXXMF') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE T.Code='" & VchNo & "'"
'                             SQL = "SELECT I.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.ActualQuantity-C.DeliveredQuantityC-C.BilledMFB", "C.DeliveredQuantityC-C.BilledMFC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,14)=R.Code+'XXXXXXXXXXXXMF') INNER JOIN BookPOChild05 C ON R.Code=C.Code) INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "'"
        SQL = SQL + " UNION ALL " 'ME
        SQL = SQL + "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.ActualQuantity-C.DeliveredQuantityC-C.BilledMEB", "C.DeliveredQuantityC-C.BilledMEC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild06 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+'XXXXXXME') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code  WHERE T.Code='" & VchNo & "'"
        SQL = SQL + " UNION ALL " 'CF
        SQL = SQL + "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.ActualQuantity-C.DeliveredQuantityC-C.BilledCFB", "C.DeliveredQuantityC-C.BilledCFC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild0901 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Book+'XXXXXXCF') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN BookMaster E ON C.Book=E.Code WHERE T.Code='" & VchNo & "'"
        SQL = SQL + " UNION ALL " 'MO
        SQL = SQL + "SELECT I.Name+'_'+E.Name+'_'+O.Name As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.Quantity-C.DeliveredQuantityC-C.BilledMOB", "C.DeliveredQuantityC-C.BilledMOC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+E.Code+O.Code+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild07 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+C.Operation+'MO') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code WHERE T.Code='" & VchNo & "'"
        SQL = SQL + " UNION ALL " 'BN
                        '        SQL = SQL + "SELECT I.Name+'_Binding' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.Quantity-C.DeliveredQuantityC-C.BilledBNB", "C.DeliveredQuantityC-C.BilledBNC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,14)=R.Code+'XXXXXXXXXXXXBN') INNER JOIN BookPOChild08 C ON R.Code=C.Code) INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "'"
          SQL = SQL + "SELECT I.Name+'_Binding'+'_'+O.Name As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.Quantity-C.DeliveredQuantityC-C.BilledBNB", "C.DeliveredQuantityC-C.BilledBNC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+'XXXXXX'+O.Code+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOChild08 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+'XXXXXX'+C.BinderyProcess+'BN') INNER JOIN BookPOParent R ON R.Code=C.Code) INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN GeneralMaster O ON C.BinderyProcess=O.Code WHERE T.Code='" & VchNo & "'"
        SQL = SQL + " UNION ALL " 'BM
        SQL = SQL + "SELECT I.Name+'_'+CASE WHEN C.Category='1' THEN O.Name WHEN C.Category='2' THEN P.Name ELSE U.Name END As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.OrderQuantity-C.DeliveredQuantityC-C.BilledBMB", "C.DeliveredQuantityC-C.BilledBMC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+C.Item+'XXXXX'+C.Category+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo " & _
                                "FROM (((((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild0801 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Item+'XXXXX'+C.Category+'BM') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) LEFT JOIN OutsourceItemMaster O ON C.Category+C.Item='1'+O.Code) LEFT JOIN PaperMaster P ON C.Category+C.Item='2'+P.Code) LEFT JOIN BookMaster U ON C.Category+C.Item='3'+U.Code WHERE T.Code='" & VchNo & "'"
    End If
    SQL = SQL + " ORDER BY SrNo"
    rstJobworkBVChild.Open SQL, cnJobworkBill, adOpenKeyset, adLockReadOnly
    rstJobworkBVChild.ActiveConnection = Nothing
    If rstJobworkBVChild.RecordCount > 0 Then rstJobworkBVChild.MoveFirst
    i = 0
    Do Until rstJobworkBVChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstJobworkBVChild.Fields("ItemName").Value
            .SetText 2, i, rstJobworkBVChild.Fields("HSNName").Value
            .SetText 3, i, rstJobworkBVChild.Fields("VchNo").Value
            .SetText 4, i, Val(rstJobworkBVChild.Fields("Quantity").Value)
            .SetText 5, i, Val(rstJobworkBVChild.Fields("Rate").Value)
            .SetText 6, i, Val(rstJobworkBVChild.Fields("Amount").Value)
            .SetText 7, i, rstJobworkBVChild.Fields("NarrationName").Value
            .SetText 8, i, rstJobworkBVChild.Fields("LongNarration01").Value
            .SetText 9, i, rstJobworkBVChild.Fields("LongNarration02").Value
            .SetText 10, i, rstJobworkBVChild.Fields("LongNarration03").Value
            .SetText 11, i, rstJobworkBVChild.Fields("LongNarration04").Value
            .SetText 12, i, rstJobworkBVChild.Fields("LongNarration05").Value
            .SetText 13, i, rstJobworkBVChild.Fields("NarrationCode").Value
            .SetText 14, i, rstJobworkBVChild.Fields("VchCode").Value
            .SetText 15, i, rstJobworkBVChild.Fields("ItemCode").Value
            .SetText 16, i, rstJobworkBVChild.Fields("HSNCode").Value
            .SetText 17, i, Val(CheckNull(rstJobworkBVChild.Fields("PendingQty").Value))
        End With
        rstJobworkBVChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Item List")
End Sub
Private Function UpdateItemList(ByVal ActionType As String, ByVal SrNo As Integer) As Boolean
    Dim CellVal(1 To 12) As Variant, BOM As String
    On Error GoTo ErrorHandler
    UpdateItemList = True
    If ActionType = "D" Then
        If Not blnRecordExist Then Exit Function
        With rstJobworkBVChild
            If .State = adStateOpen Then
                If .RecordCount > 0 Then .MoveFirst
                Do While Not .EOF
                    If Not CheckEmpty(.Fields("VchCode").Value, False) Then Call UpdateStatus(.Fields("VchCode").Value, .Fields("Quantity").Value, "-")
                    .MoveNext
                Loop
            End If
        End With
        cnJobworkBill.Execute "DELETE FROM JobworkBVChild WHERE Code='" & rstJobworkBVParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 4, .ActiveRow, CellVal(1) 'Qnty
            .GetText 5, .ActiveRow, CellVal(2) 'Rate
            .GetText 6, .ActiveRow, CellVal(3) 'Amnt
            .GetText 8, .ActiveRow, CellVal(4) 'Long Narration I
            .GetText 9, .ActiveRow, CellVal(5) 'Long Narration II
            .GetText 10, .ActiveRow, CellVal(6) 'Long Narration III
            .GetText 11, .ActiveRow, CellVal(7) 'Long Narration IV
            .GetText 12, .ActiveRow, CellVal(8) 'Long Narration V
            .GetText 13, .ActiveRow, CellVal(9) 'Short Narration Code
            .GetText 14, .ActiveRow, CellVal(10) 'VchCode=SOCode+Element+Operation+ItemType & Null for Direct Sales without Sales Order
            .GetText 15, .ActiveRow, CellVal(11) 'Item Code
            .GetText 16, .ActiveRow, CellVal(12) 'HSN Code
        End With
        BOM = VchPrefix + IIf(CheckEmpty(CellVal(10), False), "XXXXXXXXXXXXFI", Right(CellVal(10), 14)) 'BOM='0410'+Element+Operation+ItemType
        cnJobworkBill.Execute "INSERT INTO JobworkBVChild VALUES ('" & rstJobworkBVParent.Fields("Code").Value & "'," & IIf(CheckEmpty(CellVal(10), False), "Null", "'" & Left(CellVal(10), 6) & "'") & ",'" & BOM & "','" & CellVal(11) & "','" & CellVal(12) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ",'" & CellVal(9) & "'," & SrNo & ",'" & CellVal(4) & "','" & CellVal(5) & "','" & CellVal(6) & "','" & CellVal(7) & "','" & CellVal(8) & "','0','XXXXXX')" 'Ref=Null for Direct Sales without Sales Order
        If Not CheckEmpty(CellVal(10), False) Then Call UpdateStatus(CellVal(10), Val(CellVal(1)), "+")
    End If
    Exit Function
ErrorHandler:
    UpdateItemList = False
End Function
Private Sub LoadOrderList()
    Dim SQL As String
    If rstOrderList.State = adStateOpen Then rstOrderList.Close
    If InStr(1, "U_C", Right(VchType, 1)) > 0 Then 'Unit/Jobwork Unit Cost
        SQL = "SELECT DISTINCT P.Code+'XXXXXXXXXXXXFI' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/FI' As VchNo,P.Date As VchDate,I.Name As Item,P.EstQty01 As OrderedQty,P.BilledAllC As BilledQtyC,P.BilledAllB As BilledQtyD,P.DeliveredQuantityC As ChallanQty,P.DeliveredQuantityB As DirectQty,ISNULL(J.Quantity,0) As ClearQty "
        SQL = SQL + "FROM (BookPOParent P INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN BookPOChild0801 C ON P.Code=C.Code LEFT JOIN JobworkBVClear J ON P.Code=J.RefCode WHERE (P.BookPrinter='" & BuyerCode & "' OR P.TitlePrinter='" & BuyerCode & "' OR P.Laminator='" & BuyerCode & "' OR P.Binder='" & BuyerCode & "' OR C.Vendor='" & BuyerCode & "') AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)<>'FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(P.EstQty01-P.DeliveredQuantityC-P.BilledAllB-ISNULL(J.Quantity,0))>0" 'Ordered-Delivered(Challan)-Billed(Direct)=Pending Quantity
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(P.DeliveredQuantityC-P.BilledAllC-ISNULL(J.Quantity,0))>0" 'Delivered(Challan)-Billed(Challan)=Pending Quantity
        End If
        SQL = SQL + " ORDER BY I.Name,P.Date,VchNo"
    ElseIf Right(VchType, 1) = "J" Then 'Jobwork
       SQL = "SELECT P.Code AS oCode,P.Code+E.Code+'XXXXXXMF' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/MF' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.BilledMFC As BilledQtyC,C.BilledMFB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty,0 As ClearQty FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.BookPrinter='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
'      SQL = "SELECT P.Code AS oCode,P.Code+'XXXXXXXXXXXXMF'  As  VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/MF' As VchNo,P.Date As  VchDate,I.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.BilledMFC As BilledQtyC,C.BilledMFB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty,0 As ClearQty FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.BookPrinter='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledMFB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledMFC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT P.Code AS oCode,P.Code+E.Code+'XXXXXXME' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/ME' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.BilledMEC As BilledQtyC,C.BilledMEB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty,0 As ClearQty FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.TitlePrinter='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledMEB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledMEC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT P.Code AS oCode,P.Code+E.Code+'XXXXXXCF' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/CF' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.BilledCFC As BilledQtyC,C.BilledCFB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty,0 As ClearQty FROM ((BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN BookMaster E ON C.Book=E.Code WHERE P.TitlePrinter='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledCFB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledCFC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT P.Code AS oCode,P.Code+E.Code+O.Code+'MO' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/MO' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_'+O.Name As Item,C.Quantity As OrderedQty,C.BilledMOC As BilledQtyC,C.BilledMOB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty,0 As ClearQty FROM (((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code WHERE P.Laminator='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.Quantity-C.DeliveredQuantityC-C.BilledMOB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledMOC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " UNION ALL "
'        SQL = SQL + "SELECT P.Code AS oCode,P.Code+'XXXXXXXXXXXXBN' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/BN' As VchNo,P.Date As VchDate,I.Name+'_Binding' As Item,C.Quantity As OrderedQty,C.BilledBNC As BilledQtyC,C.BilledBNB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty,0 As ClearQty FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Binder='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        SQL = SQL + "SELECT P.Code AS oCode,P.Code+'XXXXXX'+O.Code+'BN' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/BN' As VchNo,P.Date As VchDate,I.Name+'_Binding'+'_'+O.Name As Item,C.Quantity As OrderedQty,C.BilledBNC As BilledQtyC,C.BilledBNB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty,0 As ClearQty FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster O ON C.BinderyProcess=O.Code WHERE P.Binder='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.Quantity-C.DeliveredQuantityC-C.BilledBNB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledBNC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT P.Code AS oCode,P.Code+C.Item+'XXXXX'+C.Category+'BM' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/BM' As VchNo,P.Date As VchDate,I.Name+'_'+CASE WHEN C.Category='1' THEN O.Name WHEN C.Category='2' THEN R.Name ELSE U.Name END As Item,C.OrderQuantity As OrderedQty,C.BilledBMC As BilledQtyC,C.BilledBMB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty,0 As ClearQty FROM ((((BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN OutsourceItemMaster O ON C.Category+C.Item='1'+O.Code) LEFT JOIN PaperMaster R ON C.Category+C.Item='2'+R.Code) LEFT JOIN BookMaster U ON C.Category+C.Item='3'+U.Code WHERE C.Vendor='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND C.Amount<>0 " & _
                                "AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.OrderQuantity-C.DeliveredQuantityC-C.BilledBMB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledBMC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " ORDER BY oCode" ',Item,VchNo,VchDate"
    End If
    rstOrderList.Open SQL, cnJobworkBill, adOpenKeyset, adLockReadOnly
    rstOrderList.ActiveConnection = Nothing
    If rstOrderList.RecordCount = 0 Then DisplayError ("No Pending Order Exists"): fpSpread1.SetFocus: Exit Sub
    Load FrmOrderList
    FrmOrderList.Text2 = Text3.Text
    Dim i As Integer, Delivered As Long, Pending As Long, UnitRate As Double
    With rstOrderList
        For i = 1 To .RecordCount
            With FrmOrderList.fpSpread1
                .MaxRows = .MaxRows + 1
                .InsertRows i, 1
            End With
        Next
        i = 0
        Do Until .EOF
            If cmbBillingType.ListIndex = 0 Then 'Direct
                Pending = Val(.Fields("OrderedQty").Value) - Val(.Fields("ChallanQty").Value) - Val(.Fields("BilledQtyD").Value) - Val(.Fields("ClearQty").Value) 'Pending=Ordered-Delivered(Challan)-Billed(Direct)-ClearQty
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                Pending = Val(.Fields("ChallanQty").Value) - Val(.Fields("BilledQtyC").Value) - Val(.Fields("ClearQty").Value) 'Pending=Delivered(Challan)-Billed(Challan)-ClearQty
            End If
            i = i + 1
            FrmOrderList.fpSpread1.SetText 1, i, .Fields("Item").Value
            FrmOrderList.fpSpread1.SetText 2, i, .Fields("VchNo").Value
            FrmOrderList.fpSpread1.SetText 3, i, Format(.Fields("VchDate").Value, "dd-MM-yy")
            FrmOrderList.fpSpread1.SetText 4, i, Val(.Fields("OrderedQty").Value)
            FrmOrderList.fpSpread1.SetText 5, i, Val(.Fields("BilledQtyC").Value) + Val(.Fields("BilledQtyD").Value) 'Total Billed
            FrmOrderList.fpSpread1.SetText 6, i, Val(.Fields("OrderedQty").Value) - Val(.Fields("BilledQtyC").Value) - Val(.Fields("BilledQtyD").Value) 'Total Unbilled
            FrmOrderList.fpSpread1.SetText 7, i, Pending
            Delivered = Val(.Fields("ChallanQty").Value) + Val(.Fields("DirectQty").Value)
            FrmOrderList.fpSpread1.SetText 8, i, IIf(Delivered = 0, "Undelivered", IIf(Delivered < Val(.Fields("OrderedQty").Value), "Under Delivery", "Delivered"))
            FrmOrderList.fpSpread1.SetText 9, i, 1
            FrmOrderList.fpSpread1.SetText 10, i, .Fields("VchCode").Value
            .MoveNext
        Loop
        FrmOrderList.fpSpread1.SetActiveCell 10, 1
    End With
    FrmOrderList.Check2 = 1 'Select All
    FrmOrderList.Check1.Visible = False 'Show All
    FrmOrderList.Show vbModal
    If Not CheckEmpty(FrmOrderList.VchCodeList, False) Then
        If rstOrderList.State = adStateOpen Then rstOrderList.Close
        If InStr(1, "U_C", Right(VchType, 1)) > 0 Then
            SQL = "SELECT I.Name As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,P.UnitRate,0 As ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/FI' As VchNo,P.Code+'XXXXXXXXXXXXFI' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(P.EstQty01-P.DeliveredQuantityC-P.BilledAllB) As PendingQty" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(P.DeliveredQuantityC-P.BilledAllC) As PendingQty" 'Delivered(Challan)-Billed(Challan)
            End If
            SQL = SQL + " FROM (BookPOParent P INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+'XXXXXXXXXXXXFI' IN (" & FrmOrderList.VchCodeList & ") ORDER BY I.Name,VchNo"
        
        ElseIf Right(VchType, 1) = "J" Then 'Jobwork
           SQL = "SELECT  P.Code AS oCode,I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND((C.PrintAmount+Adjustment+C.PlateAmount+PAdjustment+PaperAmount+RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/MF' As VchNo,P.Code+E.Code+'XXXXXXMF' As VchCode,"
'            SQL = "SELECT P.Code AS oCode,I.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND((C.PrintAmount+Adjustment+C.PlateAmount+PAdjustment+PaperAmount+RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/MF' As VchNo,P.Code+'XXXXXXXXXXXXMF' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledMFB) As PendingQty" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledMFC) As PendingQty" 'Delivered(Challan)-Billed(Challan)
            End If
           SQL = SQL + " FROM (((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+'XXXXXXMF' IN (" & FrmOrderList.VchCodeList & ")"
'            SQL = SQL + " FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+'XXXXXXXXXXXXMF' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT P.Code AS oCode,I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND((C.PrintAmount+C.Adjustment+C.PlateAmount+C.PAdjustment+C.PaperAmount+C.RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/ME' As VchNo,P.Code+E.Code+'XXXXXXME' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledMEB) As PendingQty" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledMEC) As PendingQty" 'Delivered(Challan)-Billed(Challan)
            End If
            SQL = SQL + " FROM (((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+'XXXXXXME' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT P.Code AS oCode,I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND(((C1.PrintAmount+C1.Adjustment+C1.PlateAmount+C1.PAdjustment+C1.PaperAmount+C1.RAdjustment)/(SELECT SUM(ActualQuantity) FROM BookPOChild0901 WHERE Code=P.Code)),3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/CF' As VchNo,P.Code+E.Code+'XXXXXXCF' As VchCode,"
            'SQL = SQL + "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND(((C1.PrintAmount+C1.Adjustment+C1.PlateAmount+C1.PAdjustment+C1.PaperAmount+C1.RAdjustment)/(SELECT SUM(ActualQuantity) FROM BookPOChild0901 WHERE Code=P.Code))*C.ActualQuantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/CF' As VchNo,P.Code+E.Code+'XXXXXXCF' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledCFB) As Pending" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledCFC) As Pending" 'Delivered(Challan)-Billed(Challan)
            End If
            SQL = SQL + "  FROM ((((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C ON C.Code=C1.Code) INNER JOIN BookMaster I ON  P.Book=I.Code) INNER JOIN BookMaster E ON C.Book=E.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+'XXXXXXCF' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT P.Code AS oCode,I.Name+'_'+E.Name+'_'+O.Name As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND((C.Amount+C.Adjustment)/C.Quantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/MO' As VchNo,P.Code+E.Code+O.Code+'MO' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.Quantity-C.DeliveredQuantityC-C.BilledMOB) As PendingQty" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledMOC) As PendingQty" 'Delivered(Challan)-Billed(Challan)
            End If
            SQL = SQL + " FROM ((((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+O.Code+'MO' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT P.Code AS oCode,I.Name+'_Binding'+'_'+O.Name As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND((C.BillAmount-C.GST)/C.Quantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/BN' As VchNo,P.Code+'XXXXXX'+O.Code+'BN' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.Quantity-C.DeliveredQuantityC-C.BilledBNB) As PendingQty" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledBNC) As PendingQty" 'Delivered(Challan)-Billed(Challan)
            End If
            SQL = SQL + " FROM (((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster O ON C.BinderyProcess=O.Code ) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+'XXXXXX'+O.Code+'BN' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT P.Code AS oCode,I.Name+'_'+CASE WHEN C.Category='1' THEN O.Name WHEN C.Category='2' THEN R.Name ELSE U.Name END As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND(C.Amount/C.OrderQuantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/BM' As VchNo,P.Code+C.Item+'XXXXX'+C.Category+'BM' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.OrderQuantity-C.DeliveredQuantityC-C.BilledBMB) As PendingQty" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledBMC) As PendingQty" 'Delivered(Challan)-Billed(Challan)
            End If
            SQL = SQL + " FROM (((((BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN OutsourceItemMaster O ON C.Category+C.Item='1'+O.Code) LEFT JOIN PaperMaster R ON C.Category+C.Item='2'+R.Code) LEFT JOIN BookMaster U ON C.Category+C.Item='3'+U.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+C.Item+'XXXXX'+C.Category+'BM' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " ORDER BY oCode" 'ItemName,VchNo"
        End If
        rstOrderList.Open SQL, cnJobworkBill, adOpenKeyset, adLockReadOnly
        If rstOrderList.RecordCount > 0 Then
            i = 0
            With fpSpread1
                Do While Not rstOrderList.EOF
                    i = i + 1
                    .SetText 1, i, rstOrderList.Fields("ItemName").Value: .SetText 15, i, rstOrderList.Fields("ItemCode").Value
                    .SetText 2, i, rstOrderList.Fields("HSNName").Value: .SetText 16, i, rstOrderList.Fields("HSNCode").Value
                    .SetText 3, i, rstOrderList.Fields("VchNo").Value: .SetText 14, i, rstOrderList.Fields("VchCode").Value
                    .SetText 4, i, Val(rstOrderList.Fields("PendingQty").Value): .SetText 17, i, Val(rstOrderList.Fields("PendingQty").Value)
                    UnitRate = Val(rstOrderList.Fields("UnitRate").Value) + (Val(rstOrderList.Fields("UnitRate").Value) * Val(rstOrderList.Fields("ProfitMargin").Value)) / 100
                    .SetText 5, i, Round(UnitRate, 3)
                    .SetText 6, i, Val(rstOrderList.Fields("PendingQty").Value) * Round(UnitRate, 3) 'quantity * rate
                    .SetText 7, i, Mid(rstOrderList.Fields("Narration").Value, InStr(1, rstOrderList.Fields("Narration").Value, "-") + 1, 40): .SetText 13, i, Left(rstOrderList.Fields("Narration").Value, InStr(1, rstOrderList.Fields("Narration").Value, "-") - 1)
                    rstOrderList.MoveNext
                Loop
'                If VchType <> "SU" Then GetJobWorkPaperCost
                If InStr(1, "SC_PC_QC_ZC_SX_PX_QX_ZX", Right(VchType, 2)) > 0 Then GetJobWorkPaperCost
                Call CalculateTotal
            End With
        End If
    End If
    CloseForm FrmOrderList
End Sub
Private Function GetJobWorkPaperCost()
Dim i As Long, SQLwt As String, Rate As Variant, CellVal(1 To 8) As Variant, n As Long, K As Long
            With fpSpread1
            For i = 1 To .DataRowCnt
            .SetActiveCell 8, i
            .GetText 14, fpSpread1.ActiveRow, vtCode: vtCode = Right(vtCode, 2)
        If VchType = "SU" Then Exit Function
        If InStr(1, "MF_ME_CF_BM_FI", vtCode) > 0 Then
            .GetText 14, fpSpread1.ActiveRow, vtCode: vtCode = Left(vtCode, 6)
                SQLwt = "SELECT ISNUll(C5.PaperConsumptionsheets/(Select Value1 From GeneralMAster Where Code=PM1.UOM)*PM1.[Weight/Unit]/P.EstQty01,0) +"
'                SQLwt = SQLwt + "ISNUll(C5.PaperConsumptionsheets2/(Select Value1 From GeneralMAster Where Code=PM2.UOM)*PM2.[Weight/Unit]/P.EstQty01,0) +"
'                SQLwt = SQLwt + "ISNUll(C5.PaperConsumptionsheets4/(Select Value1 From GeneralMAster Where Code=PM4.UOM)*PM4.[Weight/Unit]/P.EstQty01,0) +"
                SQLwt = SQLwt + "ISNUll(C6.PaperConsumptionKg/P.EstQty01,0) +"
                SQLwt = SQLwt + "ISNULL(C9.PaperConsumptionSheets/(Select Value1 From GeneralMAster Where Code=PM9.UOM)*PM9.[Weight/Unit]/P.EstQty01,0) As Pwt "
                SQLwt = SQLwt + "FROM ((((((BookPOParent P LEFT JOIN BookPOChild05 C5 ON P.Code=C5.Code) LEFT JOIN BookPOChild06 C6 ON P.Code=C6.Code) LEFT JOIN BookPOChild08 C8 ON P.Code=C8.Code) LEFT JOIN BookPOChild09 C9 ON P.Code=C9.Code) LEFT JOIN PaperMaster PM1 ON C5.Paper=PM1.Code) LEFT JOIN PaperMaster PM6 ON C6.Paper=PM6.Code) LEFT JOIN PaperMaster PM9 ON C9.Paper=PM9.Code "
                SQLwt = SQLwt + " WHERE P.Code='" & vtCode & "'"
                If rstOrderList.State = adStateOpen Then rstOrderList.Close
                rstOrderList.Open SQLwt, cnJobworkBill, adOpenKeyset, adLockReadOnly
                .GetText 4, fpSpread1.ActiveRow, CellVal(1) 'Qty
                .GetText 5, fpSpread1.ActiveRow, CellVal(2) 'Jobwork Rate
                .GetText 8, fpSpread1.ActiveRow, CellVal(3) 'LongNarration01
                .GetText 9, fpSpread1.ActiveRow, CellVal(4) 'LongNarration02
                .GetText 10, fpSpread1.ActiveRow, CellVal(5) 'LongNarration03
                .GetText 11, fpSpread1.ActiveRow, CellVal(6) 'LongNarration04
                .GetText 12, fpSpread1.ActiveRow, CellVal(7) 'LongNarration05
                K = 0
                For n = 3 To 7
                    If CellVal(n) <> "" And InStr(1, "Paper Suply By Party Cost Aprox Rs.:", Left(CellVal(n), 36)) = 0 Then
                        If InStr(1, "Jobwork Cost Rs.:", Left(CellVal(n), 17)) = 0 Then K = K + 1
                    End If
                Next
                If K < 4 Then
                    K = 10
                    For n = 3 To 7
                        If CellVal(n) <> "" And InStr(1, "Paper Suply By Party Cost Aprox Rs.:", Left(CellVal(n), 36)) = 0 Then
                            If K < 13 And InStr(1, "Jobwork Cost Rs.:", Left(CellVal(n), 17)) = 0 Then .SetText K, i, CellVal(n): K = K + 1
                        End If
                    Next
                Else
                    If MsgBox("Do You Want's To Replace Old Long Narrations ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") = vbYes Then
                    K = 10
                    For n = 3 To 7
                        If CellVal(n) <> "" And InStr(1, "Paper Suply By Party Cost Aprox Rs.:", Left(CellVal(n), 36)) = 0 Then
                            If K < 13 And InStr(1, "Jobwork Cost Rs.:", Left(CellVal(n), 17)) = 0 Then .SetText K, i, CellVal(n): K = K + 1
                        End If
                    Next
                    Else
                        Exit For
                    End If
                End If
                Rate = InputBox("Paper Rate @..", , Val(Rate), 12490, 6940)
                If Val(Rate) = 0 Then Exit Function
                    .SetText 8, i, "Paper Suply By Party Cost Aprox Rs.: " & Format(Val(rstOrderList.Fields("Pwt")) * CellVal(1) * Val(Rate), "##,##,##,###") & " /-"
                    .SetText 9, i, "Jobwork Cost Rs.: " & Format(CellVal(1) * CellVal(2), "##,##,##,###") & " /-"
        End If
            Next
            End With
End Function
Private Sub LoadMasterList()
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    If rstMaterialCentreList.State = adStateOpen Then rstMaterialCentreList.Close
    rstMaterialCentreList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='*99999' ORDER BY Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
    rstMaterialCentreList.ActiveConnection = Nothing
    If rstSalesTypeList.State = adStateOpen Then rstSalesTypeList.Close
    rstSalesTypeList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='" & IIf(Left(VchType, 1) = "S", "*26027", "*26025") & "' ORDER BY Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
    rstSalesTypeList.ActiveConnection = Nothing
    If rstTaxList.State = adStateOpen Then rstTaxList.Close
    If PartyStateCode = "" Or PartyStateCode = Null Then
    rstTaxList.Open "SELECT Name As Col0,[IGST%],[SGST%],[CGST%],Region,Code FROM TaxMaster ORDER BY Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
    Else
    rstTaxList.Open "SELECT Name As Col0,[IGST%],[SGST%],[CGST%],Region,Code FROM TaxMaster Where Region='" & IIf(CompStateCode = PartyStateCode, "L", "I") & "' ORDER BY Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
    End If
    rstTaxList.ActiveConnection = Nothing
    If rstNarrationList.State = adStateOpen Then rstNarrationList.Close
    rstNarrationList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='17' ORDER BY Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
    rstNarrationList.ActiveConnection = Nothing
    If rstHSNCodeList.State = adStateOpen Then rstHSNCodeList.Close
    rstHSNCodeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='18' ORDER BY Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
    rstHSNCodeList.ActiveConnection = Nothing
    If rstItemList.State = adStateOpen Then rstItemList.Close
    rstItemList.Open "SELECT I.Name As Col0,I.Price,I.Code,H.Code As HSNCode,H.Name As HSNName FROM BookMaster I INNER JOIN GeneralMaster H ON I.HSNCode=H.Code ORDER BY I.Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
    rstItemList.ActiveConnection = Nothing
    If rstVchSeriesList.State = adStateOpen Then rstVchSeriesList.Close
    rstVchSeriesList.Open "SELECT Name As Col0,Prefix,Suffix,VchNumbering,Code FROM VchSeriesMaster WHERE Left(FYCode,2)='" & Left(FYCode, 2) & "' AND VchType ='" & Switch(Left(VchType, 1) = "S", "04", Left(VchType, 1) = "P", "01", Left(VchType, 1) = "Z", "23", Left(VchType, 1) = "Q", "24") & VchType & "' ORDER BY Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
    rstVchSeriesList.ActiveConnection = Nothing
End Sub
Private Sub DisplayMenu(ByVal OutputTo As String)
    'If Left(VchType, 1) = "P" Then Exit Sub
    Dim menusel As String
    If rstJobworkBVList.RecordCount = 0 Then Exit Sub
    menusel = DisplayPopupMenu(Me.hwnd, 2)
    If menusel = 0 Then menusel = 1
    Call PrintJobworkBillVch(rstJobworkBVList.Fields("Code").Value, rstJobworkBVList.Fields("Type").Value, Choose(menusel, "O", "D", "T"), OutputTo) 'Original/Duplicate/Triplicate
    If Not (rstJobworkBVList.EOF Or rstJobworkBVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.SetFocus
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Private Sub PushVch()
    Dim XMLStr, SaleAccount, UOM, i, ChallanTbl
    With rstCompanyMaster
        If .State = adStateOpen Then .Close
        '.Open "SELECT VchSeries,Account,UOM FROM AppConfig WHERE VchType='SF'", cnJobworkBill, adOpenKeyset, adLockReadOnly
        .Open "SELECT Name+'-'+VchNAme As VchSeries,VchNAme As Account,'No.' As UOM FROM VchSeriesMaster WHERE Right(VchType,2)='" & VchType & "'", cnJobworkBill, adOpenKeyset, adLockReadOnly
        VchSeries = .Fields("VchSeries").Value: SaleAccount = .Fields("Account").Value: UOM = .Fields("UOM").Value
    End With
    With rstCompanyMaster
        If .State = adStateOpen Then .Close
        .Open "SELECT * FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnJobworkBill, adOpenKeyset, adLockReadOnly
    End With
    With rstJobworkBVChild
        ChallanTbl = "": XMLStr = ""
        
    If Mid(rstJobworkBVList.Fields("Type").Value, 3, 2) <> "10" Then
        ChallanTbl = ChallanTbl + "SELECT BOM,(Select Sum(STRC1.Quantity) FROM JobworkBVParent STRP1 INNER JOIN JobworkBVChild STRC1 On STRC1.code=STRP1.Code WHERE STRC1.Ref = D.Ref AND Ref <>''AND (RefCode='XXXXXX' OR RefCode='') AND STRP1.Type IN ('0110PU','0110PC','0110PJ','0310PU','0310PC','0310PJ','0510FR','0710FR','0210OU','0210OC','0210OJ','0410TU','0410TC','0410TJ','0610FI','0810FI') AND Ref IN (SELECT Code FROM BookPOParent) AND Ref <>''AND (RefCode='XXXXXX' OR RefCode='')) AS DeliveredQty,"
        ChallanTbl = ChallanTbl + " (Select VchDetails From (SELECT DISTINCT STRC.Ref As TRef,(Select  Ltrim(Name) From BookPOParent Where Code=STRC.Ref) PO,RIGHT((Select TYPE From BookPOParent Where Code=Ref),1)+'O/'+LTRIM((Select Name From BookPOParent Where Code=Ref))+'/JW/'+IIF(FORMAT((Select Date From BookPOParent Where Code=Ref),'MM')<4,Convert(Nvarchar,(Convert(int,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy'))-1)) +'-'+Convert(Nvarchar,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy')),Convert(Nvarchar,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy')) +'-'+ Convert(Nvarchar,Convert(int,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy'))+1)) As VchBillNo,"
        ChallanTbl = ChallanTbl + "SUBSTRING((SELECT ',  '+STRP1.Name AS [text()]  FROM dbo.JobworkBVParent STRP1 INNER JOIN JobworkBVChild STRC1 On STRC1.code=STRP1.Code "
        ChallanTbl = ChallanTbl + "WHERE STRC1.Ref = STRC.Ref AND Ref <>''AND (RefCode='XXXXXX' OR RefCode='') AND STRP1.Type IN ('0110PU','0110PC','0110PJ','0310PU','0310PC','0310PJ','0510FR','0710FR','0210OU','0210OC','0210OJ','0410TU','0410TC','0410TJ','0610FI','0810FI') AND Ref IN (SELECT Code FROM BookPOParent) AND Ref <>''AND (RefCode='XXXXXX' OR RefCode='') "
        ChallanTbl = ChallanTbl + "ORDER BY STRC1.Ref,STRC1.Code "
        ChallanTbl = ChallanTbl + "FOR XML PATH (''), TYPE).value('text()[1]','nvarchar(max)'), 2, 1000) VchDetails "
        ChallanTbl = ChallanTbl + "FROM dbo.JobworkBVParent STRP INNER JOIN JobworkBVChild STRC On STRC.Code=STRP.Code WHERE Ref <>''AND (RefCode='XXXXXX' OR RefCode='') AND STRP.Type IN ('0110PU','0110PC','0110PJ','0310PU','0310PC','0310PJ','0510FR','0710FR','0210OU','0210OC','0210OJ','0410TU','0410TC','0410TJ','0610FI','0810FI') "
        ChallanTbl = ChallanTbl + "AND Ref IN (SELECT Code FROM BookPOParent)) As Challan Where  TRef=D.Ref) As VchDetails,"
    Else
        ChallanTbl = ChallanTbl + "Select '' As DeliveredQty,'' As VchDetails,"
    End If
        If .State = adStateOpen Then .Close
        XMLStr = XMLStr + ChallanTbl
        XMLStr = XMLStr + "G.Name As State,N.Name As cState,'India' As CountryOfResidence,Left((Select V.Prefix From VchSeriesMaster V Where Code=VchSeries),2)+'/'+'" & Right(VchType, 2) & "'+'/'+Ltrim(AutoVchNo)+'/'+ '" & FYFromTo & "'  As BillNo,H.Date As BillDate,M.PrintName As MatCentre," + _
                        "B.PrintName As Buyer,Party As AccountCode,B.Address1 As bAddress1,B.Address2 As bAddress2,B.Address3 As bAddress3,B.Address4 As bAddress4,B.TIN As bGSTIN,(Select Name From GeneralMaster Where Code= B.State) AS State,C.PrintName As Consignee,C.Address1 As cAddress1,C.Address2 As cAddress2,C.Address3 As cAddress3,C.Address4 As cAddress4,C.TIN As cGSTIN," + _
                        "H.TaxableAmount,H.[Rebate%],H.Rebate,H.Freight,H.Adjustment,H.Tax,H.[IGST%],H.IGST,H.[SGST%],H.SGST,H.[CGST%],H.CGST,H.Amount As FinalAmount,H.Remarks," + _
                        "I.Name As ItemName,(Select PrintName From GeneralMaster Where Code=I.IntegrationUnit) As IntegrationUnit,I.ItemIntegrationName As Item,I.BusyCode As ItemAlias,D.Rate,D.[Disc%],ABS(D.Quantity) As Quantity,D.Amount,'' As DeliveryNoteNo,ISNULL(H.GRNo,'') As DispatchDocNo,ISNULL(H.Transport,'') As DishpatchThrough,ISNULL(H.Station,'') As Destination,ISNULL(H.Transport,'') As [CarrierName/Agent],ISNULL(H.GRNo,'') As [BillofLoading/LR-RRNO],ISNULL(H.GRDate,'') As GRDate,ISNULL(H.VehicleNo,'') As MotorVehicleNo,ISNULL(H.eWayBill,'') As eWayBill,ISNULL(H.eWayBillDate,'') As eWayBillDate,IIF((Select Name From AccountMaster Where Code=H.SalesType)='Sales','Bill of Supply','Bill of Supply') As DocType,'Supply' As SubType,'Generated by me' As BILLSTATUS,D.Item As ItemCode,D.LongNarration01,D.LongNarration02,D.LongNarration03,D.LongNarration04,D.LongNarration05,CASE WHEN Ref IS NULL THEN '' ELSE 'Ref. No.: '+LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(D.BOM,2) END As RefOrderNo," & _
                        "(Convert(nvarchar,(ISNULL((Select PaperConsumptionsheets From BookPOChild05 Where Code=R.Code),0)+ISNULL((Select PaperConsumptionsheets From BookPOChild06 Where Code=R.Code),0)+ISNULL((Select PaperConsumptionsheets From BookPOChild09 Where Code=R.Code),0)))+' Sheets') As PaperConsumptionsheets " & _
                        "FROM (((((((JobWorkBVParent H INNER JOIN AccountMaster B ON H.Party=B.Code) INNER JOIN AccountMaster C ON H.Consignee=C.Code) INNER JOIN AccountMaster M ON H.MaterialCentre=M.Code) INNER JOIN JobWorkBVChild D ON H.Code=D.Code) INNER JOIN BookMaster I ON D.Item=I.Code) Left JOIN BookPOParent R ON D.Ref=R.Code) LEFT Join GeneralMaster G ON B.State=G.Code ) INNER Join GeneralMaster N ON C.State=N.Code " + _
                        "WHERE H.Code='" + rstJobworkBVList.Fields("Code").Value + "'"
        .Open XMLStr, cnJobworkBill, adOpenKeyset, adLockReadOnly
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
            XMLStr = XMLStr + "<STATENAME>" + Replace(Trim(.Fields("State").Value), "&", "&amp;") + "</STATENAME>" 'State
            XMLStr = XMLStr + "<COUNTRYOFRESIDENCE>" + Replace(Trim(.Fields("COUNTRYOFRESIDENCE").Value), "&", "&amp;") + "</COUNTRYOFRESIDENCE>" 'State
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
            XMLStr = XMLStr + "<CONSIGNEESTATENAME>" + Replace(Trim(.Fields("cState").Value), "&", "&amp;") + "</CONSIGNEESTATENAME>"
            XMLStr = XMLStr + "<CONSIGNEECOUNTRYNAME>" + Replace(Trim(.Fields("COUNTRYOFRESIDENCE").Value), "&", "&amp;") + "</CONSIGNEECOUNTRYNAME>"
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
            If Not CheckEmpty(Trim(.Fields("VchDetails").Value), False) Then XMLStr = XMLStr + "<BASICSHIPDELIVERYNOTE>" + Replace(Trim(.Fields("VchDetails").Value), "&", "&amp;") + "</BASICSHIPDELIVERYNOTE>"
        If cmbBillingType.ListIndex = 0 Then
            'Delivery Note Date
            If Not CheckEmpty(Trim(.Fields("BillDate").Value), False) Then XMLStr = XMLStr + "<BASICSHIPPINGDATE>" + Replace(Trim(.Fields("BillDate").Value), "&", "&amp;") + "</BASICSHIPPINGDATE>"
            'Despatch Doc No
            If Not CheckEmpty(Trim(.Fields("BillNo").Value), False) Then XMLStr = XMLStr + "<BASICSHIPDOCUMENTNO>" + Replace(Trim(.Fields("BillNo").Value), "&", "&amp;") + "</BASICSHIPDOCUMENTNO>"
        End If
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
            XMLStr = XMLStr + "<CONSIGNEESTATENAME>" + Replace(Trim(.Fields("State").Value), "&", "&amp;") + "</CONSIGNEESTATENAME>"

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
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + "StockItem:" + Replace(Trim(.Fields("ItemName").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'StockItemName
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + Replace(Trim(.Fields("LongNarration01").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'LongNarration01
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + Replace(Trim(.Fields("LongNarration02").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'LongNarration02
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + Replace(Trim(.Fields("LongNarration03").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'LongNarration03
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + Replace(Trim(.Fields("LongNarration04").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'LongNarration04
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + Replace(Trim(.Fields("LongNarration05").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'LongNarration05
        If VchType <> "SU" Then
            If InStr(1, "MF_ME_CF_BM", Right(Trim(.Fields("BOM").Value), 2)) > 0 Then
                If MsgBox("Print Paper Consumption Sheets ?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Quit !") = vbYes Then
                     XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + "Paper Consumption : " + Replace(Trim(.Fields("PaperConsumptionsheets").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'PaperConsumptionsheets
                End If
            End If
        End If
            If MsgBox("Print Order Reference ?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Quit !") = vbYes Then
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + Replace(Trim(.Fields("RefOrderNo").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'RefOrderNo
            End If
    If Trim(.Fields("DeliveredQty").Value) <> "" Then
            If MsgBox("Print Delivered Quantity And Challan Details ?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Quit !") = vbYes Then
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + "Quantity Delivered : " + Replace(Trim(.Fields("DeliveredQty").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'Challan Details
                XMLStr = XMLStr + "<BASICUSERDESCRIPTION>" + "Against Challans : " + Replace(Trim(.Fields("VchDetails").Value), "&", "&amp;") + "</BASICUSERDESCRIPTION>" 'Challan Details
            End If
    End If
                XMLStr = XMLStr + "</BASICUSERDESCRIPTION.LIST>"
                
                XMLStr = XMLStr + "<STOCKITEMNAME>" + Replace(Trim(.Fields("Item").Value), "&", "&amp;") + "</STOCKITEMNAME>"
                
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISAUTONEGATE>No</ISAUTONEGATE>"
                XMLStr = XMLStr + "<ISCUSTOMSCLEARANCE>No</ISCUSTOMSCLEARANCE>"
                XMLStr = XMLStr + "<ISTRACKCOMPONENT>No</ISTRACKCOMPONENT>"
                XMLStr = XMLStr + "<ISTRACKPRODUCTION>No</ISTRACKPRODUCTION>"
                XMLStr = XMLStr + "<ISPRIMARYITEM>No</ISPRIMARYITEM>"
                XMLStr = XMLStr + "<ISSCRAP>No</ISSCRAP>"
                    
                XMLStr = XMLStr + "<RATE>" + Format(Val(.Fields("Rate").Value), "0.00") + "/" + Replace(Trim(.Fields("IntegrationUnit").Value), "&", "&amp;") + "</RATE>"
                XMLStr = XMLStr + "<DISCOUNT>" + Format(Val(.Fields("Disc%").Value), "0.00") + "</DISCOUNT>"
                XMLStr = XMLStr + "<AMOUNT>" + Format(Val(.Fields("Amount").Value), "0.00") + "</AMOUNT>"
                XMLStr = XMLStr + "<ACTUALQTY>" + Format(Val(.Fields("Quantity").Value), "0.00") + " " + Replace(Trim(.Fields("IntegrationUnit").Value), "&", "&amp;") + "</ACTUALQTY>"
                XMLStr = XMLStr + "<BILLEDQTY>" + Format(Val(.Fields("Quantity").Value), "0.00") + " " + Replace(Trim(.Fields("IntegrationUnit").Value), "&", "&amp;") + "</BILLEDQTY>"
                
    'BATCHALLOCATIONS
                XMLStr = XMLStr + "<BATCHALLOCATIONS.LIST>"
                XMLStr = XMLStr + "<GODOWNNAME>" + Replace(Trim(.Fields("MatCentre").Value), "&", "&amp;") + "</GODOWNNAME>"
                XMLStr = XMLStr + "<BATCHNAME>Primary Batch</BATCHNAME>"
                XMLStr = XMLStr + "<DESTINATIONGODOWNNAME>" + Replace(Trim(.Fields("MatCentre").Value), "&", "&amp;") + "</DESTINATIONGODOWNNAME>"
                XMLStr = XMLStr + "<INDENTNO/>"
                XMLStr = XMLStr + "<ORDERNO/>"
                XMLStr = XMLStr + "<TRACKINGNUMBER/>"
                XMLStr = XMLStr + "<DYNAMICCSTISCLEARED>No</DYNAMICCSTISCLEARED>"
                XMLStr = XMLStr + "<AMOUNT>" + Format(Val(.Fields("Amount").Value), "0.00") + "</AMOUNT>"
                XMLStr = XMLStr + "<ACTUALQTY>" + Format(Val(.Fields("Quantity").Value), "0.00") + " " + Replace(Trim(.Fields("IntegrationUnit").Value), "&", "&amp;") + "</ACTUALQTY>"
                XMLStr = XMLStr + "<BILLEDQTY>" + Format(Val(.Fields("Quantity").Value), "0.00") + " " + Replace(Trim(.Fields("IntegrationUnit").Value), "&", "&amp;") + "</BILLEDQTY>"
                XMLStr = XMLStr + "<ADDITIONALDETAILS.LIST>        </ADDITIONALDETAILS.LIST>"
                XMLStr = XMLStr + "<VOUCHERCOMPONENTLIST.LIST>        </VOUCHERCOMPONENTLIST.LIST>"
                XMLStr = XMLStr + "</BATCHALLOCATIONS.LIST>"
                
                 XMLStr = XMLStr + "<ACCOUNTINGALLOCATIONS.LIST>"
                XMLStr = XMLStr + "<OLDAUDITENTRYIDS.LIST TYPE=""Number"">"
                XMLStr = XMLStr + "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"
                XMLStr = XMLStr + "</OLDAUDITENTRYIDS.LIST>"
                
                XMLStr = XMLStr + "<LEDGERNAME>" + Replace(SaleAccount, "&", "&amp;") + "</LEDGERNAME>"
                XMLStr = XMLStr + "<GSTCLASS/>"
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<LEDGERFROMITEM>No</LEDGERFROMITEM>"
                XMLStr = XMLStr + "<REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"
                XMLStr = XMLStr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                XMLStr = XMLStr + "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"
                XMLStr = XMLStr + "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"
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
    If GSTMethod = "1" Then
    'Case-1
                XMLStr = XMLStr + "<LEDGERENTRIES.LIST>"
                XMLStr = XMLStr + "<ROUNDTYPE>Normal Rounding</ROUNDTYPE>"
                XMLStr = XMLStr + "<LEDGERNAME>Output IGST</LEDGERNAME>"
                XMLStr = XMLStr + "<METHODTYPE>GST</METHODTYPE>"
                XMLStr = XMLStr + "<GSTCLASS/>"
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<LEDGERFROMITEM>No</LEDGERFROMITEM>"
                XMLStr = XMLStr + "<REMOVEZEROENTRIES>Yes</REMOVEZEROENTRIES>"
                XMLStr = XMLStr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                XMLStr = XMLStr + "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"
                XMLStr = XMLStr + "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"
                XMLStr = XMLStr + "<AMOUNT>" + Trim(Format(Val(.Fields("IGST").Value), "0.00")) + "</AMOUNT>"
                XMLStr = XMLStr + "<VATEXPAMOUNT>" + Trim(Format(Val(.Fields("IGST").Value), "0.00")) + "</VATEXPAMOUNT>"
                XMLStr = XMLStr + "</LEDGERENTRIES.LIST>"
    ElseIf GSTMethod = "2" Then
    'Case-2
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
    End If
Else
    If GSTMethod = "1" Then
    'Case-1
                XMLStr = XMLStr + "<LEDGERENTRIES.LIST>"
                XMLStr = XMLStr + "<LEDGERNAME>Output CGST</LEDGERNAME>"
                XMLStr = XMLStr + "<GSTCLASS/>"
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<LEDGERFROMITEM>No</LEDGERFROMITEM>"
                XMLStr = XMLStr + "<REMOVEZEROENTRIES>Yes</REMOVEZEROENTRIES>"
                XMLStr = XMLStr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                XMLStr = XMLStr + "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"
                XMLStr = XMLStr + "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"
                XMLStr = XMLStr + "<AMOUNT>" + Trim(Format(Val(.Fields("CGST").Value), "0.00")) + "</AMOUNT>"
                XMLStr = XMLStr + "<VATEXPAMOUNT>" + Trim(Format(Val(.Fields("CGST").Value), "0.00")) + "</VATEXPAMOUNT>"
                XMLStr = XMLStr + "</LEDGERENTRIES.LIST>"

                XMLStr = XMLStr + "<LEDGERENTRIES.LIST>"
                XMLStr = XMLStr + "<ROUNDTYPE>Normal Rounding</ROUNDTYPE>"
                XMLStr = XMLStr + "<LEDGERNAME>Output SGST</LEDGERNAME>"
                XMLStr = XMLStr + "<GSTCLASS/>"
                XMLStr = XMLStr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<LEDGERFROMITEM>No</LEDGERFROMITEM>"
                XMLStr = XMLStr + "<REMOVEZEROENTRIES>Yes</REMOVEZEROENTRIES>"
                XMLStr = XMLStr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                XMLStr = XMLStr + "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"
                XMLStr = XMLStr + "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"
                XMLStr = XMLStr + "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"
                XMLStr = XMLStr + "<AMOUNT>" + Trim(Format(Val(.Fields("SGST").Value), "0.00")) + "</AMOUNT>"
                XMLStr = XMLStr + "<VATEXPAMOUNT>" + Trim(Format(Val(.Fields("SGST").Value), "0.00")) + "</VATEXPAMOUNT>"
                XMLStr = XMLStr + "</LEDGERENTRIES.LIST>"
    ElseIf GSTMethod = "2" Then
    'Case-2
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
Public Sub PrintJobworkBillVch(ByVal VchCode As String, ByVal VchType As String, ByVal BillType As String, Optional ByVal OutputType As String)
    Dim ChallanTbl  As String, SQL As String, oSQL As String, FS01 As String, FS02 As String, FS03 As String, FS04 As String, FS05 As String, FS06 As String, FS07 As String, FS08 As String, FS09 As String, FS10 As String, FS11 As String, FS12 As String, FS13 As String, FS14 As String, FS15 As String, FS16 As String, FS17 As String
    FS01 = "'Size: '+LTRIM(S.PrintName) "
    FS02 = "'Size: '+LTRIM(S.PrintName)+IIf(I.Pages = 0 AND (ISNull(C2.Pages1) = True OR C2.Pages1+C2.Pages2+C2.Pages4 = 0), '', ', '+LTRIM(IIF(ISNULL(C2.Pages1)=True,I.Pages,C2.Pages1+C2.Pages2+C2.Pages4))+' pages/'+LTRIM(IIF(ISNULL(C2.Pages1)=True,I.Forms,C2.Forms1+C2.Forms2+C2.Forms4))+'f ('+IIF(ISNULL(C2.Pages1)=True,LTRIM(IIF(I.OneColorForms<>0,LTRIM(I.OneColorForms)+'f-1Col','')+' '+IIF(I.TwoColorForms<>0,LTRIM(I.TwoColorForms)+'f-2Col','')+' '+IIF(I.FourColorForms<>0,LTRIM(I.FourColorForms)+'f-4Col','')),LTRIM(IIF(C2.Forms1<>0,LTRIM(C2.Forms1)+'f-1Col','')+' '+IIF(C2.Forms2<>0,LTRIM(C2.Forms2)+'f-2Col','')+' '+IIF(C2.Forms4<>0,LTRIM(C2.Forms4)+'f-4Col','')))+')')" 'Billing Against Sale Order
'    FS03 = "'MF-Text Plates: '+LTRIM((C2.[TotalPlates1-¼]+C2.[TotalPlates1-½]+C2.[TotalPlates1-1]+C2.[RevisedPlates1])*1)+ '-plates- 1Col, '+LTRIM((C2.[TotalPlates2-¼]+C2.[TotalPlates2-½]+C2.[TotalPlates2-1]+C2.[RevisedPlates2])*2)+ '-plates- 2Col, '+LTRIM((C2.[TotalPlates4-¼]+C2.[TotalPlates4-½]+C2.[TotalPlates4-1]+C2.[RevisedPlates4])*4)+ '-plates- 4Col  = ' +LTRIM(((C2.[TotalPlates1-¼]+C2.[TotalPlates1-½]+C2.[TotalPlates1-1]+C2.[RevisedPlates1])*1)+((C2.[TotalPlates2-¼]+C2.[TotalPlates2-½]+C2.[TotalPlates2-1]+C2.[RevisedPlates2])*2)+((C2.[TotalPlates4-¼]+C2.[TotalPlates4-½]+C2.[TotalPlates4-1]+C2.[RevisedPlates4])*4))+'  Nos.'"
'    FS04 = "'MF-Text Ptg.  : '+LTRIM(Forms1)+ 'forms- 1Col, '+LTRIM(Forms2)+ 'forms- 2Col, '+LTRIM(Forms4)+ 'forms- 4Col  = ' +LTRIM(Forms1+Forms2+forms4)+'  Nos.'"
'    FS05 = "'ME-Title Plates: '+LTRIM(C3.TotalPlates)+ '-Plates ( '+LTRIM(C3.FrontPrintingType)+' + '+LTRIM(C3.BackPrintingType)+' ) - Color'"
'    FS06 = "'ME-Title Ptg.   : '+'( '+LTRIM(C3.FrontPrintingType)+' + '+LTRIM(C3.BackPrintingType)+' ) - Color Printing'"
'    FS07 = "'CF-Title Plates: '+LTRIM(C5.TotalPlates)+ '-Plates ( '+LTRIM(C5.FrontPrintingColor)+' + '+LTRIM(C5.BackPrintingColor)+' ) - Color'"
'    FS08 = "'CF-Title Ptg. : '+'( '+LTRIM(C5.FrontPrintingColor)+' + '+LTRIM(C5.BackPrintingColor)+' ) - Color Printing'"
'    FS09 = "'Misc. Operations : '+'Text & Title Finishing'"
'    FS10 = "'Binding : '+'( '+LTRIM(C4.BindingType)+')'"
'    FS11 = "'BOM : '+' BOM Items' "
'    FS12 = "'Paper-(1Col) : '+LTRIM(PM1.Name)"
'    FS13 = "'Paper-(2Col) : '+LTRIM(PM2.Name)"
'    FS14 = "'Paper-(4Col) : '+LTRIM(PM4.Name)"
'    FS15 = "'Paper-(ME) : '+LTRIM(PM5.Name)"
'    FS16 = "'Paper-(CF) : '+LTRIM(PM6.Name)"
'    FS17 = "'Paper : '+' Total Paper Value '"
    If Mid(VchType, 3, 2) <> "10" Then
    ChallanTbl = ChallanTbl + "' Delivered Quantity: '+(Select Convert(nvarchar,Convert(float,Sum(STRC1.Quantity))) FROM JobworkBVParent STRP1 INNER JOIN JobworkBVChild STRC1 On STRC1.code=STRP1.Code WHERE STRC1.Ref = C1.Ref AND Ref <>''AND (RefCode='XXXXXX' OR RefCode='') AND STRP1.Type IN ('0110PU','0110PC','0110PJ','0310PU','0310PC','0310PJ','0510FR','0710FR','0210OU','0210OC','0210OJ','0410TU','0410TC','0410TJ','0610FI','0810FI') AND Ref IN (SELECT Code FROM BookPOParent) AND Ref <>''AND (RefCode='XXXXXX' OR RefCode='')) AS DeliveredQty,"
    ChallanTbl = ChallanTbl + "'Against Challan Details: '+(Select VchDetails From (SELECT DISTINCT STRC.Ref As TRef,(Select  Ltrim(Name) From BookPOParent Where Code=STRC.Ref) PO,RIGHT((Select TYPE From BookPOParent Where Code=Ref),1)+'O/'+LTRIM((Select Name From BookPOParent Where Code=Ref))+'/JW/'+IIF(FORMAT((Select Date From BookPOParent Where Code=Ref),'MM')<4,Convert(Nvarchar,(Convert(int,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy'))-1)) +'-'+Convert(Nvarchar,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy')),Convert(Nvarchar,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy')) +'-'+ Convert(Nvarchar,Convert(int,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy'))+1)) As VchBillNo,"
    ChallanTbl = ChallanTbl + "SUBSTRING((SELECT ',  '+STRP1.Name AS [text()]  FROM dbo.JobworkBVParent STRP1 INNER JOIN JobworkBVChild STRC1 On STRC1.code=STRP1.Code "
    ChallanTbl = ChallanTbl + "WHERE STRC1.Ref = STRC.Ref AND Ref <>''AND (RefCode='XXXXXX' OR RefCode='') AND STRP1.Type IN ('0110PU','0110PC','0110PJ','0310PU','0310PC','0310PJ','0510FR','0710FR','0210OU','0210OC','0210OJ','0410TU','0410TC','0410TJ','0610FI','0810FI') AND Ref IN (SELECT Code FROM BookPOParent) AND Ref <>''AND (RefCode='XXXXXX' OR RefCode='') "
    ChallanTbl = ChallanTbl + "ORDER BY STRC1.Ref,STRC1.Code "
    ChallanTbl = ChallanTbl + "FOR XML PATH (''), TYPE).value('text()[1]','nvarchar(max)'), 2, 1000) VchDetails "
    ChallanTbl = ChallanTbl + "FROM dbo.JobworkBVParent STRP INNER JOIN JobworkBVChild STRC On STRC.Code=STRP.Code WHERE Ref <>''AND (RefCode='XXXXXX' OR RefCode='') AND STRP.Type IN ('0110PU','0110PC','0110PJ','0310PU','0310PC','0310PJ','0510FR','0710FR','0210OU','0210OC','0210OJ','0410TU','0410TC','0410TJ','0610FI','0810FI') "
    ChallanTbl = ChallanTbl + "AND Ref IN (SELECT Code FROM BookPOParent)) As Challan Where  TRef=C1.Ref) As VchDetails,"
    End If
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstJobworkBVChild.State = adStateOpen Then rstJobworkBVChild.Close
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,EMail,Website,GSTIN,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,BankName,AccountNo,IFSC,Prefix,Suffix FROM CompanyMaster P INNER JOIN CompChild C ON P.Code=C.Code WHERE VchType=" & IIf(Right(VchType, 2) = "SU", 4, IIf(Right(VchType, 2) = "SC", 4, IIf(Right(VchType, 2) = "SJ", 4, 23))), cnJobworkBill, adOpenKeyset, adLockReadOnly
    rstCompanyMaster.ActiveConnection = Nothing
      SQL = "SELECT Distinct LTRIM(P.Name) +'/' +'" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "' As BillNo,P.Date As BillDate," & _
                "A.PrintName As Party,A.Address1 As PartyAddress1,A.Address2 As PartyAddress2,A.Address3 As PartyAddress3,A.Address4 As PartyAddress4,A.TIN As PartyGSTIN,C.PrintName As Consignee,C.Address1 As ConsigneeAddress1,C.Address2 As ConsigneeAddress2,C.Address3 As ConsigneeAddress3,C.Address4 As ConsigneeAddress4,C.TIN As ConsigneeGSTIN," & _
                "P.[Rebate%],P.Rebate,P.Freight,P.Adjustment,P.TaxableAmount,P.[IGST%],P.IGST,P.[SGST%],P.SGST,P.[CGST%],P.CGST,P.Amount As TotalAmount,P.Remarks," & _
                "N.PrintName As Narration,'StockItem: '+I.PrintName As Item,H.PrintName As HSNCode,C1.Quantity,C1.Rate,C1.Amount,N.Name As Narration,LTRIM(C1.Code)+LTRIM(C1.SrNo) As Ref,"
      oSQL = SQL
    cnJobworkBill.CommandTimeout = 300
    If InStr(1, "SC_PC_QC_ZC", Right(VchType, 2)) > 0 Then
                                                                SQL = SQL + "'' As cmbTitle," & FS01 & " As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,"
                                                                SQL = SQL + "IIF((ISNULL((Select PaperConsumptionsheets From BookPOChild05 Where Code=P2.Code),0)+ISNULL((Select Sum(PaperConsumptionsheets) From BookPOChild06 Where Code=P2.Code),0)+ISNULL((Select PaperConsumptionsheets From BookPOChild09 Where Code=P2.Code),0))=0,'',"
                                                                SQL = SQL + "ISNULL('Paper Consumption : '+(Convert(nvarchar,Round(Format((ISNULL((Select PaperConsumptionsheets From BookPOChild05 Where Code=P2.Code),0)+ISNULL((Select Sum(PaperConsumptionsheets) From BookPOChild06 Where Code=P2.Code),0)+ISNULL((Select PaperConsumptionsheets From BookPOChild09 Where Code=P2.Code),0))/P2.EstQty01*C1.Quantity,'00.00'),.5))+' Sheets'),'')) As PaperConsumptionsheets,"
                                                                SQL = SQL + "IIF(C1.Ref IS NULL,'','Ref. No. :')+(CASE WHEN C1.Ref IS NULL THEN '' ELSE LTRIM(P2.Name)+'/'+IIF(P2.Type='OP','CRN/',RIGHT(P2.Type,1)+'O/')+RIGHT(C1.BOM,2) +' \ dt : '+Convert(nchar,Format(P2.Date,'dd-MMM-yyyy')) END) As RefOrderNo,"
                                                                SQL = SQL + "IIF(Len(I.ISBN)=0,'','ISBN - '+I.ISBN) As ISBN,"
                                                                SQL = SQL + ChallanTbl + "'' AS FS05,'' AS FS06,'' AS FS07,Left((Select V.Prefix From VchSeriesMaster V Where Code=VchSeries),2)+'/'+'" & Right(VchType, 2) & "'+'/'+Ltrim(AutoVchNo) As VchNo,P.ChallanNo,P.ChallanDate,P.Transport,P.GRNo,P.GRDate,P.VehicleNo,P.Station,eWayBill,eWayBillDate "
                                                                SQL = SQL + "FROM((((((((((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) LEFT JOIN BookMaster I ON C1.Item=I.Code) LEFT JOIN GeneralMaster N ON C1.Narration=N.Code) LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code) LEFT JOIN (BookPOParent P2 LEFT JOIN BookPOChild05 C2 ON P2.Code=C2.Code) ON C1.Ref=P2.Code) LEFT JOIN BookPOChild06 C3 ON P2.Code=C3.Code) LEFT JOIN BookPOChild08 C4 ON P2.Code=C4.Code) LEFT JOIN BookPOChild09 C5 ON P2.Code=C5.Code) LEFT JOIN GeneralMaster B ON C4.BindingType=B.Code) LEFT JOIN PaperMaster PM1 ON C2.Paper=PM1.Code) LEFT JOIN PaperMaster PM5 ON C3.Paper=PM5.Code) LEFT JOIN PaperMaster PM6 ON C5.Paper=PM6.Code) LEFT JOIN GeneralMaster S ON I.FinishSize=S.Code) LEFT JOIN AccountMaster A ON P.Party=A.Code) LEFT JOIN AccountMaster C ON P.Consignee=C.Code "
                                                                SQL = SQL + "WHERE P.Code='" + VchCode + "' ORDER BY Item,Ref " ', cnJobworkBill, adOpenKeyset, adLockReadOnly
    ElseIf InStr(1, "SX_PX_QX_ZX", Right(VchType, 2)) > 0 Then
        rstJobworkBVChild.Open SQL & "'' As cmbTitle," & FS02 & " As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM (((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN (BookPOParent P2 INNER JOIN BookPOChild05 C2 ON P2.Code=C2.Code) ON C1.Ref=P2.Code) INNER JOIN GeneralMaster S ON I.FinishSize=S.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code WHERE P.Code='" + VchCode + "' AND BOM='TP' UNION " & _
                                                                SQL & "'' As cmbTitle,'Finish Size: '+LTRIM(S.PrintName)+', '+LTRIM(C2.FrontPrintingType)+'+'+LTRIM(C2.BackPrintingType)+'Col' As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM (((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN (BookPOParent P2 INNER JOIN BookPOChild06 C2 ON P2.Code=C2.Code) ON C1.Ref=P2.Code) INNER JOIN GeneralMaster S ON I.FinishSize=S.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code WHERE P.Code='" + VchCode + "' AND BOM='CP' UNION " & _
                                                                SQL & "I2.PrintName+', Finish Size: '+LTRIM(S.PrintName)+', '+LTRIM(C3.FrontPrintingColor)+'+'+LTRIM(C3.BackPrintingColor)+'Col, '+LTRIM(C3.[Ups/Plate])+'Ups, Qty: '+LTRIM(C3.ActualQuantity) As cmbItem,'' As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM ((((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN (BookPOParent P2 INNER JOIN (BookPOChild09 C2 INNER JOIN BookPOChild0901 C3 ON C2.Code=C3.Code) ON P2.Code=C2.Code) ON C1.Ref=P2.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code) INNER JOIN BookMaster I2 ON C3.Book=I2.Code) INNER JOIN GeneralMaster S ON I2.FinishSize=S.Code WHERE P.Code='" + VchCode + "' AND BOM='JP' UNION " & _
                                                                SQL & "'' As cmbTitle,'Finish Size: '+LTRIM(S.PrintName) As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM ((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN GeneralMaster S ON I.FinishSize=S.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code WHERE P.Code='" + VchCode + "' AND BOM='MO' UNION " & _
                                                                SQL & "'' As cmbTitle,'Finish Size: '+LTRIM(S.PrintName)+', '+LTRIM(IIF(ISNULL(C3.Pages1)=True,I.Pages,C3.Pages1+C3.Pages2+C3.Pages4))+' pages/'+LTRIM(C2.BindingForms+C2.ExtraForms)+'f, '+LTRIM(B.PrintName) As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM (((((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN (BookPOParent P2 INNER JOIN BookPOChild08 C2 ON P2.Code=C2.Code) ON C1.Ref=P2.Code) INNER JOIN GeneralMaster S ON I.FinishSize=S.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code) INNER JOIN GeneralMaster B ON C2.BindingType=B.Code) LEFT JOIN BookPOChild05  C3 ON P2.Code=C3.Code " & _
                                                                "WHERE P.Code='" + VchCode + "' AND BOM='BD' ORDER BY Item,SrNo,cmbTitle", cnJobworkBill, adOpenKeyset, adLockOptimistic
    ElseIf InStr(1, "SU_PU_QU_ZU", Right(VchType, 2)) > 0 Then
                                                            SQL = SQL + "'' As cmbTitle," & FS01 & " As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,"
                                                If MsgBox("Print Paper Details ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") = vbYes Then
                                                            SQL = SQL + "IIF((ISNULL((Select PaperConsumptionsheets From BookPOChild05 Where Code=P2.Code),0)+ISNULL((Select Sum(PaperConsumptionsheets) From BookPOChild06 Where Code=P2.Code),0)+ISNULL((Select PaperConsumptionsheets From BookPOChild09 Where Code=P2.Code),0))=0,'',"
                                                            SQL = SQL + "ISNULL('Paper Consumption : '+(Convert(nvarchar,Round(Format((ISNULL((Select PaperConsumptionsheets From BookPOChild05 Where Code=P2.Code),0)+ISNULL((Select SUM(PaperConsumptionsheets) From BookPOChild06 Where Code=P2.Code),0)+ISNULL((Select PaperConsumptionsheets From BookPOChild09 Where Code=P2.Code),0))/P2.EstQty01*C1.Quantity,'00.00'),.5))+' Sheets'),'')) As PaperConsumptionsheets,"
                                                Else
                                                            SQL = SQL + "'' As PaperConsumptionsheets,"
                                                End If
                                                            SQL = SQL + "IIF(C1.Ref IS NULL,'','Ref. No. :')+(CASE WHEN Ref IS NULL THEN '' ELSE LTRIM(P2.Name)+'/'+IIF(P2.Type='OP','CRN/',RIGHT(P2.Type,1)+'O/')+RIGHT(C1.BOM,2)+' \ dt : '+Convert(nchar,Format(P2.Date,'dd-MMM-yyyy')) END) As RefOrderNo,"
                                                            SQL = SQL + "IIF(Len(I.ISBN)=0,'','ISBN - '+I.ISBN) As ISBN,"
                                                            SQL = SQL + ChallanTbl + "'' AS FS05,'' AS FS06,'' AS FS07,Left((Select V.Prefix From VchSeriesMaster V Where Code=VchSeries),2)+'/'+'" & Right(VchType, 2) & "'+'/'+Ltrim(AutoVchNo) As VchNo,P.ChallanNo,P.ChallanDate,P.Transport,P.GRNo,P.GRDate,P.VehicleNo,P.Station,eWayBill,eWayBillDate "
                                                            SQL = SQL + "FROM (((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN GeneralMaster S ON I.FinishSize=S.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code) LEFT JOIN BookPOParent P2 ON C1.Ref=P2.Code WHERE P.Code='" + VchCode + "'  ORDER BY Item,Ref "    ', cnJobworkBill, adOpenKeyset, adLockOptimistic 'AND BOM='WS'
    ElseIf InStr(1, "SJ_PJ_QJ_ZJ", Right(VchType, 2)) > 0 Then
                                                            'MF
                                                            SQL = SQL + "'' As cmbTitle," & FS01 & " As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,"
                                                            SQL = SQL + " IIF(ElementPrintName<>'',ElementPrintName,(Select Name From ElementMaster Where Code=Element)) +' Plates: '+"
                                                            SQL = SQL + "IIF(((C2.[TotalPlates-¼]+C2.[TotalPlates-½]+C2.[TotalPlates-1-F&B]+C2.[TotalPlates-1-W&T]+C2.[RevisedPlates]))=0,'',LTRIM((C2.[TotalPlates-¼]+C2.[TotalPlates-½]+C2.[TotalPlates-1-F&B]+C2.[TotalPlates-1-W&T]+C2.[RevisedPlates]))+ ' - plates') AS Plates,"
                                                            SQL = SQL + "'Forms: '+IIF(Forms=0,'',LTRIM(Forms)+' forms ( '+(Select Convert(NvarChar,Name) From GeneralMaster Where code=Color)+')')    As ptgSets, "
                                                            SQL = SQL + "IIF(ISNULL(PaperConsumptionsheets,0)=0,'',"
                                                            SQL = SQL + "ISNULL('Paper Consumption : '+(Convert(nvarchar,Round(Format((ISNULL(PaperConsumptionsheets,0)/P2.EstQty01*C1.Quantity),'00.00'),.5))+' Sheets'),'')) As PaperConsumptionsheets,"
                                                            SQL = SQL + "IIF(C1.Ref IS NULL,'','Ref. No. :')+(CASE WHEN C1.Ref IS NULL THEN '' ELSE LTRIM(P2.Name)+'/'+IIF(P2.Type='OP','CRN/',RIGHT(P2.Type,1)+'O/')+RIGHT(C1.BOM,2) +' \ dt : '+Convert(nchar,Format(P2.Date,'dd-MMM-yyyy')) END) As RefOrderNo,"
                                                            SQL = SQL + "IIF(Len(I.ISBN)=0,'','ISBN - '+I.ISBN) As ISBN,'' AS FS04,'' AS FS05,'' AS FS06,'' AS FS07,Left((Select V.Prefix From VchSeriesMaster V Where Code=VchSeries),2)+'/'+'" & Right(VchType, 2) & "'+'/'+Ltrim(AutoVchNo) As VchNo,P.ChallanNo,P.ChallanDate,P.Transport,P.GRNo,P.GRDate,P.VehicleNo,P.Station,eWayBill,eWayBillDate "
                                                            SQL = SQL + "FROM((((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) LEFT JOIN BookMaster I ON C1.Item=I.Code) LEFT JOIN GeneralMaster N ON C1.Narration=N.Code) LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code) Left JOIN GeneralMaster S ON I.FinishSize=S.Code)LEFT JOIN BookPOParent P2 ON C1.Ref=P2.Code)LEFT JOIN BookPOChild05 C2 ON C2.Code=P2.Code And C2.Element=SUBSTRING(C1.BOM,5,6))LEFT JOIN AccountMaster A ON P.Party=A.Code) LEFT JOIN AccountMaster C ON P.Consignee=C.Code "
                                                            SQL = SQL + "WHERE Right(C1.BOM,2)='MF' AND P.Code='" + VchCode + "'"
                                                            SQL = SQL + " UNION All "
                                                            'ME
                                                            SQL = SQL + oSQL + "'' As cmbTitle," & FS01 & " As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,"
                                                            SQL = SQL + "(Select Name From ElementMaster Where Code=SubString(BOM,5,6) And Right(C1.BOM,2)='ME')+': '+LTRIM(C3.TotalPlates)+ '- plates  , Color: ( '+(Select Name From GeneralMaster Where Code=C3.FrontPrintingType)+' + '+(Select Name From GeneralMaster Where Code=C3.BackPrintingType)+' )' AS Plates,"
                                                            SQL = SQL + "'Sets: '+Convert(nvarchar,Sets)+' - Nos ( '+(Select Name From GeneralMaster Where Code=C3.FrontPrintingType)+' + '+(Select Name From GeneralMaster Where Code=C3.BackPrintingType)+' ) - Color Printing' As ptgSets,"
                                                            SQL = SQL + "IIF((ISNULL((Select PaperConsumptionsheets From BookPOChild06 Where Code=P2.Code AND Code=SubString(BOM,5,6) And Right(C1.BOM,2)='ME'),0))=0,'',"
                                                            SQL = SQL + "ISNULL('Paper Consumption : '+(Convert(nvarchar,Round(Format((ISNULL((Select PaperConsumptionsheets From BookPOChild06 Where Code=P2.Code AND Code=SubString(BOM,5,6) And Right(C1.BOM,2)='ME'),0))/P2.EstQty01*C1.Quantity,'00.00'),.5))+' Sheets'),'')) As PaperConsumptionsheets,"
                                                            SQL = SQL + "IIF(C1.Ref IS NULL,'','Ref. No. :')+(CASE WHEN C1.Ref IS NULL THEN '' ELSE LTRIM(P2.Name)+'/'+IIF(P2.Type='OP','CRN/',RIGHT(P2.Type,1)+'O/')+RIGHT(C1.BOM,2) +' \ dt : '+Convert(nchar,Format(P2.Date,'dd-MMM-yyyy')) END) As RefOrderNo,"
                                                            SQL = SQL + "IIF(Len(I.ISBN)=0,'','ISBN - '+I.ISBN) As ISBN,'' AS FS04,'' AS FS05,'' AS FS06,'' AS FS07,Left((Select V.Prefix From VchSeriesMaster V Where Code=VchSeries),2)+'/'+'" & Right(VchType, 2) & "'+'/'+Ltrim(AutoVchNo) As VchNo,P.ChallanNo,P.ChallanDate,P.Transport,P.GRNo,P.GRDate,P.VehicleNo,P.Station,eWayBill,eWayBillDate "
                                                            SQL = SQL + "FROM(((((((((JobworkBVParent P LEFT JOIN JobworkBVChild C1 ON P.Code=C1.Code) LEFT JOIN BookMaster I ON C1.Item=I.Code) LEFT JOIN GeneralMaster N ON C1.Narration=N.Code) LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code) Left JOIN GeneralMaster S ON I.FinishSize=S.Code)LEFT JOIN BookPOParent P2 ON C1.Ref=P2.Code)LEFT JOIN BookPOChild06 C3 ON C3.Code=P2.Code And C3.Element=SUBSTRING(C1.BOM,5,6)) Left JOIN ElementMaster E ON E.Code=C3.Element And Right(C1.BOM,2)='ME')LEFT JOIN AccountMaster A ON P.Party=A.Code) LEFT JOIN AccountMaster C ON P.Consignee=C.Code "
                                                            SQL = SQL + "WHERE Right(C1.BOM,2)='ME' AND P.Code='" + VchCode + "'"
                                                            SQL = SQL + " UNION All "
                                                            'CF
                                                            SQL = SQL + oSQL + "'' As cmbTitle," & FS01 & " As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,"
                                                            SQL = SQL + "'CF- Plates: '+LTRIM(C5.TotalPlates)+ '-Plates ( '+LTRIM(C5.FrontPrintingColor)+' + '+LTRIM(C5.BackPrintingColor)+' ) Color'+'  Nos.' AS Plates,"
                                                            SQL = SQL + "'Sets: '+IIF(Calculation='S','1',Convert(nvarchar,(Select Count(Code) From BookPOChild0901 Where Code=C5.code)))+' - Nos ( '+LTRIM(C5.FrontPrintingColor)+' + '+LTRIM(C5.BackPrintingColor)+' ) - Color Printing' As ptgSets,"
                                                            SQL = SQL + "IIF((ISNULL((Select PaperConsumptionsheets From BookPOChild09 Where Code=P2.Code),0))=0,'',"
                                                            SQL = SQL + "ISNULL('Paper Consumption : '+(Convert(nvarchar,Round(Format((ISNULL((Select PaperConsumptionsheets From BookPOChild09 Where Code=P2.Code),0))/P2.EstQty01*C1.Quantity,'00.00'),.5))+' Sheets'),'')) As PaperConsumptionsheets,"
                                                            SQL = SQL + "IIF(C1.Ref IS NULL,'','Ref. No. :')+(CASE WHEN C1.Ref IS NULL THEN '' ELSE LTRIM(P2.Name)+'/'+IIF(P2.Type='OP','CRN/',RIGHT(P2.Type,1)+'O/')+RIGHT(C1.BOM,2) +' \ dt : '+Convert(nchar,Format(P2.Date,'dd-MMM-yyyy')) END) As RefOrderNo,"
                                                            SQL = SQL + "IIF(Len(I.ISBN)=0,'','ISBN - '+I.ISBN) As ISBN,'' AS FS04,'' AS FS05,'' AS FS06,'' AS FS07,Left((Select V.Prefix From VchSeriesMaster V Where Code=VchSeries),2)+'/'+'" & Right(VchType, 2) & "'+'/'+Ltrim(AutoVchNo) As VchNo,P.ChallanNo,P.ChallanDate,P.Transport,P.GRNo,P.GRDate,P.VehicleNo,P.Station,eWayBill,eWayBillDate "
                                                            SQL = SQL + "FROM((((((((JobworkBVParent P LEFT JOIN JobworkBVChild C1 ON P.Code=C1.Code) LEFT JOIN BookMaster I ON C1.Item=I.Code) LEFT JOIN GeneralMaster N ON C1.Narration=N.Code) LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code) Left JOIN GeneralMaster S ON I.FinishSize=S.Code)LEFT JOIN BookPOParent P2 ON C1.Ref=P2.Code)LEFT JOIN BookPOChild09 C5 ON C5.Code=P2.Code)LEFT JOIN AccountMaster A ON P.Party=A.Code) LEFT JOIN AccountMaster C ON P.Consignee=C.Code "
                                                            SQL = SQL + "WHERE Right(C1.BOM,2)='CF' AND P.Code='" + VchCode + "'"
                                                            SQL = SQL + " UNION All "
                                                            'MO
                                                            SQL = SQL + oSQL + "'' As cmbTitle," & FS01 & " As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,"
                                                            SQL = SQL + "(Select Name From ElementMaster Where Code=SubString(BOM,5,6) And Right(C1.BOM,2)='MO')+' : '+(Select Name From GeneralMaster Where Code=SubString(BOM,11,6) And Right(C1.BOM,2)='MO') AS Plates,Convert(nvarchar,Format(Number,'0'))+' : '+OperationCountName As ptgSets,'' As PaperConsumptionsheets,"
                                                            SQL = SQL + "IIF(C1.Ref IS NULL,'','Ref. No. :')+(CASE WHEN C1.Ref IS NULL THEN '' ELSE LTRIM(P2.Name)+'/'+IIF(P2.Type='OP','CRN/',RIGHT(P2.Type,1)+'O/')+RIGHT(C1.BOM,2) +' \ dt : '+Convert(nchar,Format(P2.Date,'dd-MMM-yyyy')) END) As RefOrderNo,"
                                                            SQL = SQL + "IIF(Len(I.ISBN)=0,'','ISBN - '+I.ISBN) As ISBN,'' AS FS04,'' AS FS05,'' AS FS06,'' AS FS07,Left((Select V.Prefix From VchSeriesMaster V Where Code=VchSeries),2)+'/'+'" & Right(VchType, 2) & "'+'/'+Ltrim(AutoVchNo) As VchNo,P.ChallanNo,P.ChallanDate,P.Transport,P.GRNo,P.GRDate,P.VehicleNo,P.Station,eWayBill,eWayBillDate "
                                                            SQL = SQL + "FROM((((((((JobworkBVParent P LEFT JOIN JobworkBVChild C1 ON P.Code=C1.Code) LEFT JOIN BookMaster I ON C1.Item=I.Code) LEFT JOIN GeneralMaster N ON C1.Narration=N.Code) LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code) Left JOIN GeneralMaster S ON I.FinishSize=S.Code)LEFT JOIN BookPOParent P2 ON C1.Ref=P2.Code)LEFT JOIN BookPOChild07 C6 ON C6.Code=P2.Code And C6.Element=SUBSTRING(C1.BOM,5,6))LEFT JOIN AccountMaster A ON P.Party=A.Code) LEFT JOIN AccountMaster C ON P.Consignee=C.Code "
                                                            SQL = SQL + "WHERE Right(C1.BOM,2)='MO' AND P.Code='" + VchCode + "'"
                                                            SQL = SQL + " UNION All "
                                                            'BN
                                                            SQL = SQL + oSQL + "'' As cmbTitle," & FS01 & " As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,"
                                                            SQL = SQL + "(Select Name From GeneralMaster Where Code=C4.BinderyProcess)+' : '+IIF(IIF((Select Value1 From GeneralMaster Where Code=BinderyProcess)>0,1,C4.Number)=0,'',Convert(nvarchar,IIF((Select Value1 From GeneralMaster Where Code=BinderyProcess)>0,1,C4.Number))+ ' Nos ') AS Plates,"
                                                            SQL = SQL + "'Amount:=' + IIF((((C4.Rate*P2.ProfitMargin/100)+C4.Rate)+(C4.Adjustment/C4.Quantity*CalcValue))=0,'',' (' +(Select Name From GeneralMaster Where Code=BinderyProcess)+': '+Convert(nvarchar,Format(((C1.Quantity*IIF((Select Value1 From GeneralMaster Where Code=BinderyProcess)>0,1,C4.Number))),'######0.00'))+' X @ '+Convert(nvarchar,Format((((C4.Rate*P2.ProfitMargin/100)+C4.Rate)+(C4.Adjustment/IIF((Select Value1 From GeneralMaster Where Code=BinderyProcess)>0,C4.Quantity*CalcValue,C4.Quantity*CalcValue/C4.Number))),'######0.00'))+' Per '+Convert(nvarchar,Format(C4.CalcValue,'######0.00'))+' = Rs.'+Convert(nvarchar,Format(C1.Quantity*IIF((Select Value1 From GeneralMaster Where Code=BinderyProcess)>0,1,C4.Number)*(((C4.Rate*P2.ProfitMargin/100)+C4.Rate)+(C4.Adjustment/IIF((Select Value1 From GeneralMaster Where Code=BinderyProcess)>0,C4.Quantity*CalcValue,C4.Quantity*CalcValue/C4.Number)))/C4.CalcValue,'######0.00'))+' ) ')   As ptgSets,"
                                                            SQL = SQL + "'' As PaperConsumptionsheets,"
                                                            SQL = SQL + "IIF(C1.Ref IS NULL,'','Ref. No. :')+(CASE WHEN C1.Ref IS NULL THEN '' ELSE LTRIM(P2.Name)+'/'+IIF(P2.Type='OP','CRN/',RIGHT(P2.Type,1)+'O/')+RIGHT(C1.BOM,2) +' \ dt : '+Convert(nchar,Format(P2.Date,'dd-MMM-yyyy')) END) As RefOrderNo,"
                                                            SQL = SQL + "IIF(Len(I.ISBN)=0,'','ISBN - '+I.ISBN) As ISBN,'' AS FS04,'' AS FS05,'' AS FS06,'' AS FS07,Left((Select V.Prefix From VchSeriesMaster V Where Code=VchSeries),2)+'/'+'" & Right(VchType, 2) & "'+'/'+Ltrim(AutoVchNo) As VchNo,P.ChallanNo,P.ChallanDate,P.Transport,P.GRNo,P.GRDate,P.VehicleNo,P.Station,eWayBill,eWayBillDate "
                                                            SQL = SQL + "FROM((((((((JobworkBVParent P LEFT JOIN JobworkBVChild C1 ON P.Code=C1.Code) LEFT JOIN BookMaster I ON C1.Item=I.Code) LEFT JOIN GeneralMaster N ON C1.Narration=N.Code) LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code) Left JOIN GeneralMaster S ON I.FinishSize=S.Code)LEFT JOIN BookPOParent P2 ON C1.Ref=P2.Code)LEFT JOIN BookPOChild08 C4 ON C4.Code=P2.Code And C4.BinderyProcess+'BN'=SUBSTRING(C1.BOM,11,8))LEFT JOIN AccountMaster A ON P.Party=A.Code) LEFT JOIN AccountMaster C ON P.Consignee=C.Code "
                                                            SQL = SQL + "WHERE Right(C1.BOM,2)='BN' AND P.Code='" + VchCode + "'"
                                                            SQL = SQL + " UNION All "
                                                            'BM
                                                            SQL = SQL + oSQL + "'' As cmbTitle," & FS01 & " As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,"
                                                            SQL = SQL + "IIF(C7.Category='1','OutSource: '+(SELECT Name FROM OutsourceItemMaster WHERE Code=C7.Item),IIF(C7.Category='2','Paper: '+(SELECT P.Name+' (UOM : '+LTRIM(U.Name)+')' As Name FROM PaperMaster P INNER JOIN GeneralMaster U ON P.UOM=U.Code WHERE P.Code=C7.Item),'BOM: '+(SELECT Name FROM BookMaster WHERE Code=C7.Item)))+' ( '+Convert(nvarchar,C7.OrderQuantity)+' Qty. X '+Convert(nvarchar,C7.[Consumption/Item])+' Nos. = '+Convert(nvarchar,C7.TotalConsumption)+' Nos.) ' AS Plates,"
                                                            SQL = SQL + "'Amount:= ( '+Convert(nvarchar,C7.TotalConsumption)+' Qty. X '+Convert(nvarchar,C7.[Rate])+' @. = Rs. '+Convert(nvarchar,C7.Amount)+' ) ' As ptgSets,'' As PaperConsumptionsheets,"
                                                            SQL = SQL + "IIF(C1.Ref IS NULL,'','Ref. No. :')+(CASE WHEN C1.Ref IS NULL THEN '' ELSE LTRIM(P2.Name)+'/'+IIF(P2.Type='OP','CRN/',RIGHT(P2.Type,1)+'O/')+RIGHT(C1.BOM,2) +' \ dt : '+Convert(nchar,Format(P2.Date,'dd-MMM-yyyy')) END) As RefOrderNo,"
                                                            SQL = SQL + "IIF(Len(I.ISBN)=0,'','ISBN - '+I.ISBN) As ISBN,'' AS FS04,'' AS FS05,'' AS FS06,'' AS FS07,Left((Select V.Prefix From VchSeriesMaster V Where Code=VchSeries),2)+'/'+'" & Right(VchType, 2) & "'+'/'+Ltrim(AutoVchNo) As VchNo,P.ChallanNo,P.ChallanDate,P.Transport,P.GRNo,P.GRDate,P.VehicleNo,P.Station,eWayBill,eWayBillDate "
                                                            SQL = SQL + "FROM((((((((JobworkBVParent P LEFT JOIN JobworkBVChild C1 ON P.Code=C1.Code) LEFT JOIN BookMaster I ON C1.Item=I.Code) LEFT JOIN GeneralMaster N ON C1.Narration=N.Code) LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code) Left JOIN GeneralMaster S ON I.FinishSize=S.Code)LEFT JOIN BookPOParent P2 ON C1.Ref=P2.Code)LEFT JOIN BookPOChild0801 C7 ON C7.Code=P2.Code AND C7.Item+'BN'=SUBSTRING(C1.BOM,5,6))LEFT JOIN AccountMaster A ON P.Party=A.Code) LEFT JOIN AccountMaster C ON P.Consignee=C.Code "
                                                            SQL = SQL + "WHERE Right(C1.BOM,2)='BM' AND P.Code='" + VchCode + "' Order By RefOrderNo"
    End If
                                                            rstJobworkBVChild.Open SQL, cnJobworkBill, adOpenKeyset, adLockOptimistic
    If rstJobworkBVChild.RecordCount = 0 Then On Error GoTo 0: Screen.MousePointer = vbNormal: Exit Sub
    If MsgBox("Print Item Detail?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Quit !") = vbNo Then rptJobworkBill.Section25.Suppress = True
    rstJobworkBVChild.ActiveConnection = Nothing
        With rptJobworkBill
                .Field17.DecimalPlaces = 3
                .Field25.DecimalPlaces = 0
                .Text35.Font.Size = 8
                .Field25.ThousandsSeparators = True
            If FYFromToFlag = "True" Then
                .Field58.Suppress = True
                .Text45.Suppress = False
                .Text45.SetText rstJobworkBVChild.Fields("VchNo").Value + "/" + FYFromTo
            Else
                .Field58.Suppress = False
                .Text45.Suppress = True
            End If
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

    rptJobworkBill.Text1.SetText IIf(Mid(VchType, 5, 1) = "Q", "Sales Quotation", IIf(Mid(VchType, 5, 1) = "Z", "Purchase Quotation", "Tax Invoice"))
    rptJobworkBill.Text35.SetText "Printed on " & Format(Now, "dd-MMM-yyyy") & " at " & Format(Now, "hh:mm")
    rptJobworkBill.Text40.SetText IIf(BillType = "O", "(ORIGINAL FOR RECIPIENT)", IIf(BillType = "D", "(DUPLICATE FOR SUPPLIER)", "(TRIPLICATE FOR SUPPLIER)"))
    rptJobworkBill.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptJobworkBill.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False)) Then
        rptJobworkBill.Text4.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value) & Space(1) & "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptJobworkBill.Text4.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False) Then
        rptJobworkBill.Text4.SetText "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    End If
    rptJobworkBill.Text8.SetText "GSTIN/UIN : " & Trim(rstCompanyMaster.Fields("GSTIN").Value)
'    If Trim(rstJobworkBVChild.Fields("ChallanNo").Value) <> "" Then .Text37.SetText Trim(rstJobworkBVChild.Fields("ChallanNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("ChallanDate").Value, "dd-MM-yy") Else .Text33.SetText ""
'    If Trim(rstJobworkBVChild.Fields("GRNo").Value) = "" And rstJobworkBVChild.Fields("VehicleNo").Value <> "" Then .Text38.SetText Trim(rstJobworkBVChild.Fields("VehicleNo").Value): .Text34.SetText "Vehicle NO.  :"
'    If Trim(rstJobworkBVChild.Fields("GRNo").Value) = "" And rstJobworkBVChild.Fields("VehicleNo").Value = "" Then .Text34.SetText ""
'    If Trim(rstJobworkBVChild.Fields("GRNo").Value) = "Null" Or rstJobworkBVChild.Fields("VehicleNo").Value = "Null" Then .Text34.SetText ""
'    If Trim(rstJobworkBVChild.Fields("GRNo").Value) <> "" Then .Text38.SetText Trim(rstJobworkBVChild.Fields("GRNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("GRDate").Value, "dd-MM-yy")
'    If rstJobworkBVChild.Fields("Transport").Value Then .Text39.SetText Trim(rstJobworkBVChild.Fields("Transport").Value) Else .Text36.SetText ""
'    If Trim(rstJobworkBVChild.Fields("Station").Value) <> "" Then .Text41.SetText Trim(rstJobworkBVChild.Fields("Station").Value) Else .Text42.SetText ""
'    If Trim(rstJobworkBVChild.Fields("eWayBill").Value) <> "" Then .Text32.SetText Trim(rstJobworkBVChild.Fields("eWayBill").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("eWayBillDate").Value, "dd-MM-yy") Else .Text43.SetText ""
            .Text33.SetText "": .Text36.SetText "": .Text34.SetText "": .Text42.SetText "": .Text43.SetText ""
'Challan No.
        If .Text33.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("ChallanNo").Value) <> "" Then .Text33.SetText "Challan No.     :": .Text37.SetText Trim(rstJobworkBVChild.Fields("ChallanNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("ChallanDate").Value, "dd-MM-yy") Else .Text33.SetText ""
        ElseIf .Text36.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("ChallanNo").Value) <> "" Then .Text36.SetText "Challan No.     :": .Text39.SetText Trim(rstJobworkBVChild.Fields("ChallanNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("ChallanDate").Value, "dd-MM-yy") Else .Text36.SetText ""
        ElseIf .Text34.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("ChallanNo").Value) <> "" Then .Text34.SetText "Challan No.     :": .Text38.SetText Trim(rstJobworkBVChild.Fields("ChallanNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("ChallanDate").Value, "dd-MM-yy") Else .Text34.SetText ""
        ElseIf .Text42.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("ChallanNo").Value) <> "" Then .Text42.SetText "Challan No.     :": .Text41.SetText Trim(rstJobworkBVChild.Fields("ChallanNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("ChallanDate").Value, "dd-MM-yy") Else .Text42.SetText ""
        ElseIf .Text43.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("ChallanNo").Value) <> "" Then .Text43.SetText "Challan No.     :": .Text32.SetText Trim(rstJobworkBVChild.Fields("ChallanNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("ChallanDate").Value, "dd-MM-yy") Else .Text43.SetText ""
        End If
'Transport
        If .Text33.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("Transport").Value) <> "" Then .Text33.SetText "Transport        :": .Text37.SetText Trim(rstJobworkBVChild.Fields("Transport").Value) Else .Text33.SetText ""
        ElseIf .Text36.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("Transport").Value) <> "" Then .Text36.SetText "Transport        :": .Text39.SetText Trim(rstJobworkBVChild.Fields("Transport").Value) Else .Text36.SetText ""
        ElseIf .Text34.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("Transport").Value) <> "" Then .Text34.SetText "Transport        :": .Text38.SetText Trim(rstJobworkBVChild.Fields("Transport").Value) Else .Text34.SetText ""
        ElseIf .Text42.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("Transport").Value) <> "" Then .Text42.SetText "Transport        :": .Text41.SetText Trim(rstJobworkBVChild.Fields("Transport").Value) Else .Text42.SetText ""
        ElseIf .Text43.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("Transport").Value) <> "" Then .Text43.SetText "Transport        :": .Text32.SetText Trim(rstJobworkBVChild.Fields("Transport").Value) Else .Text43.SetText ""
        End If
'Gr/RR No.
        If .Text33.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("GRNo").Value) <> "" Then .Text33.SetText "Gr/RR No.       :": .Text37.SetText Trim(rstJobworkBVChild.Fields("GRNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("GRDate").Value, "dd-MM-yy") Else .Text33.SetText ""
        ElseIf .Text36.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("GRNo").Value) <> "" Then .Text36.SetText "Gr/RR No.       :": .Text39.SetText Trim(rstJobworkBVChild.Fields("GRNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("GRDate").Value, "dd-MM-yy") Else .Text36.SetText ""
        ElseIf .Text34.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("GRNo").Value) <> "" Then .Text34.SetText "Gr/RR No.       :": .Text38.SetText Trim(rstJobworkBVChild.Fields("GRNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("GRDate").Value, "dd-MM-yy") Else .Text34.SetText ""
        ElseIf .Text42.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("GRNo").Value) <> "" Then .Text42.SetText "Gr/RR No.       :": .Text41.SetText Trim(rstJobworkBVChild.Fields("GRNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("GRDate").Value, "dd-MM-yy") Else .Text42.SetText ""
        ElseIf .Text43.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("GRNo").Value) <> "" Then .Text43.SetText "Gr/RR No.       :": .Text32.SetText Trim(rstJobworkBVChild.Fields("GRNo").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("GRDate").Value, "dd-MM-yy") Else .Text43.SetText ""
        End If
'Station
        If .Text33.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("Station").Value) <> "" Then .Text33.SetText "Station            :": .Text37.SetText Trim(rstJobworkBVChild.Fields("Station").Value) Else .Text33.SetText ""
        ElseIf .Text36.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("Station").Value) <> "" Then .Text36.SetText "Station            :": .Text39.SetText Trim(rstJobworkBVChild.Fields("Station").Value) Else .Text36.SetText ""
        ElseIf .Text34.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("Station").Value) <> "" Then .Text34.SetText "Station            :": .Text38.SetText Trim(rstJobworkBVChild.Fields("Station").Value) Else .Text34.SetText ""
        ElseIf .Text42.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("Station").Value) <> "" Then .Text42.SetText "Station            :": .Text41.SetText Trim(rstJobworkBVChild.Fields("Station").Value) Else .Text42.SetText ""
        ElseIf .Text43.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("Station").Value) <> "" Then .Text43.SetText "Station            :": .Text32.SetText Trim(rstJobworkBVChild.Fields("Station").Value) Else .Text43.SetText ""
        End If
'e-way Bill No.
        If .Text33.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("eWayBill").Value) <> "" Then .Text33.SetText "e-way Bill#   :": .Text37.SetText Trim(rstJobworkBVChild.Fields("eWayBill").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("eWayBillDate").Value, "dd-MM-yy") Else .Text33.SetText ""
        ElseIf .Text36.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("eWayBill").Value) <> "" Then .Text36.SetText "e-way Bill#   :": .Text39.SetText Trim(rstJobworkBVChild.Fields("eWayBill").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("eWayBillDate").Value, "dd-MM-yy") Else .Text36.SetText ""
        ElseIf .Text34.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("eWayBill").Value) <> "" Then .Text34.SetText "e-way Bill#   :": .Text38.SetText Trim(rstJobworkBVChild.Fields("eWayBill").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("eWayBillDate").Value, "dd-MM-yy") Else .Text34.SetText ""
        ElseIf .Text42.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("eWayBill").Value) <> "" Then .Text42.SetText "e-way Bill#   :": .Text41.SetText Trim(rstJobworkBVChild.Fields("eWayBill").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("eWayBillDate").Value, "dd-MM-yy") Else .Text42.SetText ""
        ElseIf .Text43.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("eWayBill").Value) <> "" Then .Text43.SetText "e-way Bill#   :": .Text32.SetText Trim(rstJobworkBVChild.Fields("eWayBill").Value) + " Dt : " & Format(rstJobworkBVChild.Fields("eWayBillDate").Value, "dd-MM-yy") Else .Text43.SetText ""
        End If
'VehicleNo
        If .Text33.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("VehicleNo").Value) <> "" Then .Text33.SetText "Vehicle No.     :": .Text37.SetText Trim(rstJobworkBVChild.Fields("VehicleNo").Value) Else .Text33.SetText ""
        ElseIf .Text36.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("VehicleNo").Value) <> "" Then .Text36.SetText "Vehicle No.     :": .Text39.SetText Trim(rstJobworkBVChild.Fields("VehicleNo").Value) Else .Text36.SetText ""
        ElseIf .Text34.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("VehicleNo").Value) <> "" Then .Text34.SetText "Vehicle No.     :": .Text38.SetText Trim(rstJobworkBVChild.Fields("VehicleNo").Value) Else .Text34.SetText ""
        ElseIf .Text42.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("VehicleNo").Value) <> "" Then .Text42.SetText "Vehicle No.     :": .Text41.SetText Trim(rstJobworkBVChild.Fields("VehicleNo").Value) Else .Text42.SetText ""
        ElseIf .Text43.Text = "" Then
            If Trim(rstJobworkBVChild.Fields("VehicleNo").Value) <> "" Then .Text43.SetText "Vehicle No.     :": .Text32.SetText Trim(rstJobworkBVChild.Fields("VehicleNo").Value) Else .Text43.SetText ""
        End If
    rptJobworkBill.Text10.SetText "(" & UCase(Trim(NumberToWords(rstJobworkBVChild.Fields("TotalAmount").Value, False))) & ")"
    rptJobworkBill.Text11.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptJobworkBill.Text26.SetText CheckNull(rstCompanyMaster.Fields("Declaration01").Value)
    rptJobworkBill.Text25.SetText CheckNull(rstCompanyMaster.Fields("Declaration02").Value)
    rptJobworkBill.Text22.SetText CheckNull(rstCompanyMaster.Fields("Declaration03").Value)
    rptJobworkBill.Text12.SetText CheckNull(rstCompanyMaster.Fields("Declaration04").Value)
    rptJobworkBill.Text9.SetText CheckNull(rstCompanyMaster.Fields("Declaration05").Value)
    rptJobworkBill.Text30.SetText CheckNull(rstCompanyMaster.Fields("Declaration06").Value)
    rptJobworkBill.Text31.SetText CheckNull(rstCompanyMaster.Fields("Declaration07").Value)
    rptJobworkBill.Text330.SetText "Bank Name             : " & CheckNull(rstCompanyMaster.Fields("BankName").Value)
    rptJobworkBill.Text340.SetText "A/c No.                    : " & CheckNull(rstCompanyMaster.Fields("AccountNo").Value)
    rptJobworkBill.Text360.SetText "Branch & IFS Code : " & CheckNull(rstCompanyMaster.Fields("IFSC").Value)

    rptJobworkBill.Database.SetDataSource rstJobworkBVChild, 3, 1
    'rptJobworkBill.DiscardSavedData
    Screen.MousePointer = vbNormal
    If OutputType = "S" Then
        Set FrmReportViewer.Report = rptJobworkBill
        FrmReportViewer.Show vbModal
    Else
        If rstJobworkBVList.State = adStateClosed Then  'For Print Utility
            rptJobworkBill.PaperSource = crPRBinAuto
            rptJobworkBill.PrintOut False
        Else
            rptJobworkBill.PaperSource = crPRBinAuto
            rptJobworkBill.PrintOut
        End If
    End If
End With
    Set rptJobworkBill = Nothing
    If rstJobworkBVList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster) 'For Print Utility
    Call CloseRecordset(rstJobworkBVChild)
    On Error GoTo 0
End Sub
Private Sub UpdateIntegration()
cnJobworkBill.Execute "Update JobworkBVParent Set IntegrationStatus='True' WHERE Code='" & rstJobworkBVList.Fields("Code").Value & "' AND RIGHT(Type,2)='" & VchType & "' AND FYCode='" & FYCode & "'"
End Sub
Private Function Get_JSON(Doc As String)
    Dim xmlHttp As New MSXML2.XMLHTTP60
    Dim URL As String
    Dim jsonName, jsonData As String
    Dim Response As String
    Dim jsonFile As Integer
    jsonFile = FreeFile
    Dim SQL, SaleAccount, UOM, i, ChallanTbl
    Dim SNo As Integer
    With rstCompanyMaster
        If .State = adStateOpen Then .Close
        .Open "SELECT Name+'-'+VchNAme As VchSeries,VchNAme As Account,'No.' As UOM FROM VchSeriesMaster WHERE Right(VchType,2)='" & VchType & "'", cnJobworkBill, adOpenKeyset, adLockReadOnly
        VchSeries = .Fields("VchSeries").Value: SaleAccount = .Fields("Account").Value: UOM = .Fields("UOM").Value
    End With
    With rstCompanyMaster
        If .State = adStateOpen Then .Close
        .Open "SELECT GSTIN AS GSTIN ,C.Name AS SupName,Address1+' '+Address2 AS SupAdd1,Address3 AS SupAdd2,Convert(nchar,G.Name) AS SupStateName,Convert(nchar,G.PrintName) AS SupStateCode,C.Address4 AS SupPincode,C.eMail,C.Phone FROM CompanyMaster C Inner Join GeneralMaster G ON G.Code=C.State  WHERE FYCode='" & FYCode & "'", cnJobworkBill, adOpenKeyset, adLockReadOnly
    End With
    With rstJobworkBVChild
        ChallanTbl = "": SQL = ""
        ChallanTbl = ChallanTbl + "Select '' As DeliveredQty,'' As VchDetails,"
        If .State = adStateOpen Then .Close
        SQL = SQL + ChallanTbl
        SQL = SQL + "b.TIN As BGSTIN,B.Name As Buyer,G.Name As BStateName,b.Address1+' '+b.Address2 As bAddress1,b.Address3 As bAddress2,b.Address4 As bPinCode,G.PrintName As BStateCode,B.Phone as bPhone,B.eMail As beMail,N.Name As cState,'India' As CountryOfResidence,Left((Select V.Prefix From VchSeriesMaster V Where Code=VchSeries),3)+'/'+'" & Right(VchType, 2) & "'+'/'+Ltrim(AutoVchNo)+'/'+ '" & FYFromTo & "'  As BillNo,Format(H.Date,'dd/MM/yyyy') As BillDate,M.PrintName As MatCentre," & _
                        "H.TaxableAmount,H.[Rebate%],H.Rebate,H.Freight,H.Adjustment,H.Tax,H.[IGST%],H.IGST,H.[SGST%],H.SGST,H.[CGST%],H.CGST,H.Amount As FinalAmount,H.Remarks,(Select PrintName From GeneralMaster Where Code=I.HSNCode) As HSNCode, " & _
                        "I.Name As ItemName,(Select PrintName From GeneralMaster Where Code=I.IntegrationUnit) As IntegrationUnit,I.ItemIntegrationName As Item,I.BusyCode As ItemAlias,D.Rate,D.[Disc%],ABS(D.Quantity) As Quantity,D.Amount,IIF(H.[Rebate%]<>0,Format(D.Amount/H.[Rebate%],'#######0.00'),'0') As Discount,'' As DeliveryNoteNo,ISNULL(H.GRNo,'') As DispatchDocNo,ISNULL(H.Transport,'') As DishpatchThrough,ISNULL(H.Station,'') As Destination,ISNULL(H.Transport,'') As [CarrierName/Agent],ISNULL(H.GRNo,'') As [BillofLoading/LR-RRNO],ISNULL(H.GRDate,'') As GRDate,ISNULL(H.VehicleNo,'') As MotorVehicleNo," & _
                        "ISNULL(H.eWayBill,'') As eWayBill,ISNULL(H.eWayBillDate,'') As eWayBillDate,IIF((Select Name From AccountMaster Where Code=H.SalesType)='Sales','Bill of Supply','Bill of Supply') As DocType,'Supply' As SubType,'Generated by me' As BILLSTATUS,D.Item As ItemCode,D.LongNarration01,D.LongNarration02,D.LongNarration03,D.LongNarration04,D.LongNarration05,CASE WHEN Ref IS NULL THEN '' ELSE 'Ref. No.: '+LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(D.BOM,2) END As RefOrderNo,(Convert(nvarchar,(ISNULL((Select PaperConsumptionsheets From BookPOChild05 Where Code=R.Code),0)+ISNULL((Select PaperConsumptionsheets From BookPOChild06 Where Code=R.Code),0)+ISNULL((Select PaperConsumptionsheets From BookPOChild09 Where Code=R.Code),0)))+' Sheets') As PaperConsumptionsheets,Ltrim(AutoVchNo) As AutoVchNo " & _
                        "FROM (((((((JobWorkBVParent H INNER JOIN AccountMaster B ON H.Party=B.Code) INNER JOIN AccountMaster C ON H.Consignee=C.Code) INNER JOIN AccountMaster M ON H.MaterialCentre=M.Code) INNER JOIN JobWorkBVChild D ON H.Code=D.Code) INNER JOIN BookMaster I ON D.Item=I.Code) Left JOIN BookPOParent R ON D.Ref=R.Code) LEFT Join GeneralMaster G ON B.State=G.Code ) INNER Join GeneralMaster N ON C.State=N.Code " & _
                        "WHERE H.Code='" + rstJobworkBVList.Fields("Code").Value + "'"
        .Open SQL, cnJobworkBill, adOpenKeyset, adLockReadOnly
        'Print E-Invoice.json
        
        If Dir(App.Path & "\JSON\EINV_" & FYCode & .Fields("AutoVchNo").Value & "_EISI" & ".json", vbDirectory) <> "" Then
        If MsgBox("Are you Wants to Over-write the existing JSON File For 'e-Invoice' ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Export !") = vbYes Then Exit Function
        End If
        jsonName = App.Path & "\JSON\EINV_" & FYCode & .Fields("AutoVchNo").Value & "_EISI" & ".json"
       Close jsonFile
        Open jsonName For Output As jsonFile
        jsonData = jsonData + "["
        jsonData = jsonData + "{"
        jsonData = jsonData + """" + "Version" + """" + " : " + """" + "1.1" + """" + ","
        
        jsonData = jsonData + """" + "TranDtls" + """" + " : " + "{"
        
            jsonData = jsonData + """" + "TaxSch" + """" + " : "
            jsonData = jsonData + """" + "GST" + """" + ","
            jsonData = jsonData + """" + "SupTyp" + """" + " : "
            User_Input ("2")
            jsonData = jsonData + """" + FrmDialog.uInput + """" + ","
            jsonData = jsonData + """" + "IgstOnIntra" + """" + " : "
            User_Input ("3")
            jsonData = jsonData + """" + FrmDialog.uInput + """" + ","
            jsonData = jsonData + """" + "RegRev" + """" + " : "
            User_Input ("4")
            jsonData = jsonData + """" + FrmDialog.uInput + """" + ","
            jsonData = jsonData + """" + "EcmGstin" + """" + " : "
            User_Input ("5")
            jsonData = jsonData + """" + FrmDialog.uInput + """" + "},"

        jsonData = jsonData + """" + "DocDtls" + """" + " : " + "{"
            jsonData = jsonData + """" + "Typ" + """" + " : "
            jsonData = jsonData + """" + "INV" + """" + "," 'Invoice,Credit Note,Debit Note
            jsonData = jsonData + """" + "No" + """" + " : "
            jsonData = jsonData + """" + .Fields("BillNo").Value + """" + ","
            jsonData = jsonData + """" + "Dt" + """" + " : "
            jsonData = jsonData + """" + .Fields("BillDate").Value + """" + "},"
        
        jsonData = jsonData + """" + "SellerDtls" + """" + " : " + "{"
            jsonData = jsonData + """" + "Gstin" + """" + " : "
            jsonData = jsonData + """" + Trim(rstCompanyMaster.Fields("GSTIN").Value) + """" + ","
            jsonData = jsonData + """" + "LglNm" + """" + " : "
            jsonData = jsonData + """" + Trim(rstCompanyMaster.Fields("SupName").Value) + """" + ","
            jsonData = jsonData + """" + "TrdNm" + """" + " : "
            jsonData = jsonData + """" + Trim(rstCompanyMaster.Fields("SupName").Value) + """" + ","
            jsonData = jsonData + """" + "Addr1" + """" + " : "
            jsonData = jsonData + """" + Trim(rstCompanyMaster.Fields("SupAdd1").Value) + """" + ","
            jsonData = jsonData + """" + "Addr2" + """" + " : "
            jsonData = jsonData + """" + Trim(rstCompanyMaster.Fields("SupAdd2").Value) + """" + ","
            jsonData = jsonData + """" + "Loc" + """" + " : "
            jsonData = jsonData + """" + Trim(rstCompanyMaster.Fields("SupStateName").Value) + """" + ","
            jsonData = jsonData + """" + "Pin" + """" + " : "
            jsonData = jsonData + """" + Trim(rstCompanyMaster.Fields("SupPinCode").Value) + """" + ","
            jsonData = jsonData + """" + "Stcd" + """" + " : "
            jsonData = jsonData + """" + Trim(rstCompanyMaster.Fields("SupStateCode").Value) + """" + ","
            jsonData = jsonData + """" + "Ph" + """" + " : "
            jsonData = jsonData + """" + Trim(rstCompanyMaster.Fields("Phone").Value) + """" + ","
            jsonData = jsonData + """" + "Em" + """" + " : "
            jsonData = jsonData + """" + Trim(rstCompanyMaster.Fields("email").Value) + """" + "},"
            
        jsonData = jsonData + """" + "BuyerDtls" + """" + " : " + "{"
            jsonData = jsonData + """" + "Gstin" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("BGSTIN").Value) + """" + ","
            jsonData = jsonData + """" + "LglNm" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("Buyer").Value) + """" + ","
            jsonData = jsonData + """" + "TrdNm" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("Buyer").Value) + """" + ","
            jsonData = jsonData + """" + "Pos" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("BStateName").Value) + """" + ","
            jsonData = jsonData + """" + "Addr1" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("bAddress1").Value) + """" + ","
            jsonData = jsonData + """" + "Addr2" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("bAddress2").Value) + """" + ","
            jsonData = jsonData + """" + "Loc" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("BStateName").Value) + """" + ","
            jsonData = jsonData + """" + "Pin" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("BPinCode").Value) + """" + ","
            jsonData = jsonData + """" + "Stcd" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("BStateCode").Value) + """" + ","
            jsonData = jsonData + """" + "Ph" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("BPhone").Value) + """" + ","
            jsonData = jsonData + """" + "Em" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("Bemail").Value) + """" + "},"
        
        jsonData = jsonData + """" + "DispDtls" + """" + " : "
        jsonData = jsonData + """" + "Null" + """" + ","
        jsonData = jsonData + """" + "ShipDtls" + """" + " : "
        jsonData = jsonData + """" + "Null" + """" + ","
        
        jsonData = jsonData + """" + "ValDtls" + """" + " : " + "{"
            jsonData = jsonData + """" + "AssVal" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("TaxableAmount").Value) + """" + ","
            jsonData = jsonData + """" + "IgstVal" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("IGST").Value) + """" + ","
            jsonData = jsonData + """" + "CgstVal" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("CGST").Value) + """" + ","
            jsonData = jsonData + """" + "SgstVal" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("SGST").Value) + """" + ","
            jsonData = jsonData + """" + "CesVal" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + ","
            jsonData = jsonData + """" + "StCesVal" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + ","
            jsonData = jsonData + """" + "Discount" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("Rebate").Value) + """" + ","
            jsonData = jsonData + """" + "OthChrg" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("Freight").Value + .Fields("Adjustment").Value) + """" + ","
            jsonData = jsonData + """" + "RndOffAmt" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + ","
            jsonData = jsonData + """" + "TotInvVal" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("FinalAmount").Value) + """" + "},"
        
        jsonData = jsonData + """" + "ExpDtls" + """" + " : " + "{"
            jsonData = jsonData + """" + "ShipBNo" + """" + " : "
            jsonData = jsonData + """" + "Null" + """" + ","
            jsonData = jsonData + """" + "ShipBDt" + """" + " : "
            jsonData = jsonData + """" + "Null" + """" + ","
            jsonData = jsonData + """" + "Port" + """" + " : "
            jsonData = jsonData + """" + "Null" + """" + ","
            jsonData = jsonData + """" + "RefClm" + """" + " : "
            jsonData = jsonData + """" + "Null" + """" + ","
            jsonData = jsonData + """" + "ForCur" + """" + " : "
            jsonData = jsonData + """" + "Null" + """" + ","
            jsonData = jsonData + """" + "CntCode" + """" + " : "
            jsonData = jsonData + """" + "Null" + """" + ","
            jsonData = jsonData + """" + "ExpDuty" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + "},"
        jsonData = jsonData + """" + "EwbDtls" + """" + " : "
        jsonData = jsonData + """" + "Null" + """" + ","

        jsonData = jsonData + """" + "ItemList" + """" + " : " + "["
        .MoveFirst
Do While Not .EOF
            SNo = SNo + 1
            jsonData = jsonData + "{"
            jsonData = jsonData + """" + "SlNo" + """" + " : "
            jsonData = jsonData + """" + CStr(SNo) + """" + ","
            jsonData = jsonData + """" + "PrdDesc" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("ItemName").Value) + """" + ","
            jsonData = jsonData + """" + "IsServc" + """" + " : " 'IS Service Y-Yes/N-No
            jsonData = jsonData + """" + "Y" + """" + ","
            jsonData = jsonData + """" + "HsnCd" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("HSNCode").Value) + """" + ","
            jsonData = jsonData + """" + "Barcde" + """" + " : "
            jsonData = jsonData + """" + "Null" + """" + ","
            jsonData = jsonData + """" + "Qty" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("Quantity").Value) + """" + ","
            jsonData = jsonData + """" + "Unit" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("IntegrationUnit").Value) + """" + ","
            jsonData = jsonData + """" + "UnitPrice" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("Rate").Value) + """" + ","
            jsonData = jsonData + """" + "TotAmt" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("Amount").Value) + """" + ","
            jsonData = jsonData + """" + "Discount" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("Discount").Value) + """" + ","
            jsonData = jsonData + """" + "PreTaxVal" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + ","
            jsonData = jsonData + """" + "AssAmt" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("Amount").Value - .Fields("Discount").Value) + """" + ","
            jsonData = jsonData + """" + "GstRt" + """" + " : "
            jsonData = jsonData + """" + CStr(Trim(.Fields("IGST%").Value + .Fields("SGST%").Value + .Fields("CGST%").Value)) + """" + ","
            jsonData = jsonData + """" + "IgstAmt" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("IGST").Value) + """" + ","
            jsonData = jsonData + """" + "CgstAmt" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("CGST").Value) + """" + ","
            jsonData = jsonData + """" + "SgstAmt" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("SGST").Value) + """" + ","
            jsonData = jsonData + """" + "CesRt" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + ","
            jsonData = jsonData + """" + "CesAmt" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + ","
            jsonData = jsonData + """" + "CesNonAdvlAmt" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + ","
            jsonData = jsonData + """" + "StateCesRt" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + ","
            jsonData = jsonData + """" + "StateCesAmt" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + ","
            jsonData = jsonData + """" + "StateCesNonAdvlAmt" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + ","
            jsonData = jsonData + """" + "OthChrg" + """" + " : "
            jsonData = jsonData + """" + "0" + """" + ","
            jsonData = jsonData + """" + "TotItemVal" + """" + " : "
            jsonData = jsonData + """" + Trim(.Fields("Amount").Value - .Fields("Discount").Value + .Fields("IGST").Value + .Fields("CGST").Value + .Fields("SGST").Value) + """" + ","
            jsonData = jsonData + """" + "BchDtls" + """" + " : "
            jsonData = jsonData + """" + "Null" + """"

            If .RecordCount = SNo Then
                jsonData = jsonData + "}"
            ElseIf .RecordCount > SNo Then
                jsonData = jsonData + "},"
            End If
            
    .MoveNext
    Loop
            jsonData = jsonData + "]}]"
        Print #jsonFile, jsonData
        Close jsonFile
'        'Print ewayBill.json
'        Open App.Path & "\ewaybill.json" For Output As jsonFile
    End With
End Function
Public Function User_Input(Find As Variant)
    FrmDialog.uInput = ""
    FrmDialog.Flag = Find
    Load FrmDialog
    Screen.MousePointer = vbNormal
        If Find = 2 Then
            FrmDialog.Frame1.Caption = "Select Supply Type"
        ElseIf Find = 3 Then
            FrmDialog.Frame1.Caption = "Select IGST on Intra"
        ElseIf Find = 4 Then
            FrmDialog.Frame1.Caption = "Select Reverse Charge Mechanism"
        ElseIf Find = 5 Then
            FrmDialog.Frame1.Caption = "E-Commerce GST"
        End If
    FrmDialog.Show vbModal
End Function

'New To Be Live

