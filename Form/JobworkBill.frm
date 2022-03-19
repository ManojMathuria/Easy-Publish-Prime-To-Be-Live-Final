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
   ClientHeight    =   8595
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
   ScaleHeight     =   8595
   ScaleWidth      =   13740
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   8580
      Left            =   15
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   13715
      _Version        =   65536
      _ExtentX        =   24192
      _ExtentY        =   15134
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
         Height          =   8360
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Width           =   13485
         _ExtentX        =   23786
         _ExtentY        =   14737
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
         Tab(0).Control(1)=   "Mh3dLabel1(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "DataGrid1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Text1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "JobworkBill.frx":0038
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
            Left            =   600
            MaxLength       =   40
            TabIndex        =   19
            Top             =   7830
            Width           =   8220
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7305
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   450
            Width           =   13260
            _ExtentX        =   23389
            _ExtentY        =   12885
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
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "Amount"
               Caption         =   "                 Amount"
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
                  Locked          =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   4529.764
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
                  ColumnWidth     =   4545.071
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  Locked          =   -1  'True
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7575
            Left            =   -74880
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   480
            Width           =   13260
            _Version        =   65536
            _ExtentX        =   23389
            _ExtentY        =   13361
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
            Picture         =   "JobworkBill.frx":0054
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
               Left            =   11280
               MaxLength       =   40
               TabIndex        =   3
               Top             =   120
               Width           =   1860
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
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   0
               Top             =   120
               Width           =   2130
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   8640
               TabIndex        =   24
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
               Picture         =   "JobworkBill.frx":0070
               Picture         =   "JobworkBill.frx":008C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   8640
               TabIndex        =   47
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
               Caption         =   " Consignee"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":00A8
               Picture         =   "JobworkBill.frx":00C4
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
               Left            =   9960
               MaxLength       =   40
               TabIndex        =   6
               Top             =   630
               Width           =   3180
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
               Width           =   2850
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
               TabIndex        =   7
               Top             =   945
               Width           =   3435
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   25
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
               Picture         =   "JobworkBill.frx":00E0
               Picture         =   "JobworkBill.frx":00FC
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
                  Calculator      =   "JobworkBill.frx":0118
                  Caption         =   "JobworkBill.frx":0138
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "JobworkBill.frx":01A4
                  Keys            =   "JobworkBill.frx":01C2
                  Spin            =   "JobworkBill.frx":020C
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
                  Calculator      =   "JobworkBill.frx":0234
                  Caption         =   "JobworkBill.frx":0254
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "JobworkBill.frx":02C0
                  Keys            =   "JobworkBill.frx":02DE
                  Spin            =   "JobworkBill.frx":0328
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
               Left            =   4380
               MaxLength       =   25
               TabIndex        =   1
               Top             =   120
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
               Left            =   9960
               MaxLength       =   40
               TabIndex        =   9
               Top             =   950
               Width           =   3180
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
               Left            =   3660
               TabIndex        =   22
               Top             =   120
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
               Picture         =   "JobworkBill.frx":0350
               Picture         =   "JobworkBill.frx":036C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   23
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
               Caption         =   " Party"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":0388
               Picture         =   "JobworkBill.frx":03A4
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   8280
               TabIndex        =   2
               Top             =   120
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   582
               Calendar        =   "JobworkBill.frx":03C0
               Caption         =   "JobworkBill.frx":04D8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":0544
               Keys            =   "JobworkBill.frx":0562
               Spin            =   "JobworkBill.frx":05C0
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
               TabIndex        =   10
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
               SpreadDesigner  =   "JobworkBill.frx":05E8
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
               Left            =   7440
               TabIndex        =   28
               Top             =   120
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
               Caption         =   " Vch Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":14A7
               Picture         =   "JobworkBill.frx":14C3
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   30
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
               Picture         =   "JobworkBill.frx":14DF
               Picture         =   "JobworkBill.frx":14FB
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   9940
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   6810
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "JobworkBill.frx":1517
               Caption         =   "JobworkBill.frx":1537
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":15A3
               Keys            =   "JobworkBill.frx":15C1
               Spin            =   "JobworkBill.frx":160B
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
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   6810
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   582
               Calculator      =   "JobworkBill.frx":1633
               Caption         =   "JobworkBill.frx":1653
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":16BF
               Keys            =   "JobworkBill.frx":16DD
               Spin            =   "JobworkBill.frx":1727
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
               TabIndex        =   33
               Top             =   6810
               Width           =   1225
               _Version        =   65536
               _ExtentX        =   2161
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
               Picture         =   "JobworkBill.frx":174F
               Picture         =   "JobworkBill.frx":176B
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   650
               Left            =   12000
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   6810
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   1147
               Calculator      =   "JobworkBill.frx":1787
               Caption         =   "JobworkBill.frx":17A7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":1813
               Keys            =   "JobworkBill.frx":1831
               Spin            =   "JobworkBill.frx":187B
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
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   6810
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1147
               Calculator      =   "JobworkBill.frx":18A3
               Caption         =   "JobworkBill.frx":18C3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":192F
               Keys            =   "JobworkBill.frx":194D
               Spin            =   "JobworkBill.frx":1997
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
               Top             =   6815
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
               Caption         =   " Pre-Tax Amt"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":19BF
               Picture         =   "JobworkBill.frx":19DB
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
               Picture         =   "JobworkBill.frx":19F7
               Picture         =   "JobworkBill.frx":1A13
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
               Height          =   330
               Left            =   9940
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   7130
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "JobworkBill.frx":1A2F
               Caption         =   "JobworkBill.frx":1A4F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":1ABB
               Keys            =   "JobworkBill.frx":1AD9
               Spin            =   "JobworkBill.frx":1B23
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
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   7130
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   582
               Calculator      =   "JobworkBill.frx":1B4B
               Caption         =   "JobworkBill.frx":1B6B
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":1BD7
               Keys            =   "JobworkBill.frx":1BF5
               Spin            =   "JobworkBill.frx":1C3F
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
               TabIndex        =   12
               Top             =   6810
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   1147
               Calculator      =   "JobworkBill.frx":1C67
               Caption         =   "JobworkBill.frx":1C87
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":1CF3
               Keys            =   "JobworkBill.frx":1D11
               Spin            =   "JobworkBill.frx":1D5B
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
               TabIndex        =   40
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
               Picture         =   "JobworkBill.frx":1D83
               Picture         =   "JobworkBill.frx":1D9F
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   650
               Left            =   2280
               TabIndex        =   41
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
               Picture         =   "JobworkBill.frx":1DBB
               Picture         =   "JobworkBill.frx":1DD7
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   650
               Left            =   3120
               TabIndex        =   11
               Top             =   6810
               Width           =   570
               _Version        =   65536
               _ExtentX        =   1005
               _ExtentY        =   1147
               Calculator      =   "JobworkBill.frx":1DF3
               Caption         =   "JobworkBill.frx":1E13
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":1E7F
               Keys            =   "JobworkBill.frx":1E9D
               Spin            =   "JobworkBill.frx":1EE7
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
               TabIndex        =   42
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
               Picture         =   "JobworkBill.frx":1F0F
               Picture         =   "JobworkBill.frx":1F2B
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
               Calculator      =   "JobworkBill.frx":1F47
               Caption         =   "JobworkBill.frx":1F67
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":1FD3
               Keys            =   "JobworkBill.frx":1FF1
               Spin            =   "JobworkBill.frx":203B
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
               TabIndex        =   43
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
               Picture         =   "JobworkBill.frx":2063
               Picture         =   "JobworkBill.frx":207F
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   645
               Left            =   7395
               TabIndex        =   13
               Top             =   6810
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   1138
               Calculator      =   "JobworkBill.frx":209B
               Caption         =   "JobworkBill.frx":20BB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "JobworkBill.frx":2127
               Keys            =   "JobworkBill.frx":2145
               Spin            =   "JobworkBill.frx":218F
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
               TabIndex        =   44
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
               Caption         =   "  Mat Centre"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":21B7
               Picture         =   "JobworkBill.frx":21D3
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   4620
               TabIndex        =   45
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
               Caption         =   " Billing Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":21EF
               Picture         =   "JobworkBill.frx":220B
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   120
               TabIndex        =   48
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
               Picture         =   "JobworkBill.frx":2227
               Picture         =   "JobworkBill.frx":2243
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   9960
               TabIndex        =   49
               Top             =   120
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
               Caption         =   " Purchase Type"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "JobworkBill.frx":225F
               Picture         =   "JobworkBill.frx":227B
            End
            Begin MSForms.ComboBox cmbBillingType 
               Height          =   330
               Left            =   5820
               TabIndex        =   8
               Top             =   945
               Width           =   2850
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "5027;582"
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
            Top             =   7830
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
            Picture         =   "JobworkBill.frx":2297
            Picture         =   "JobworkBill.frx":22B3
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
            Top             =   7830
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   16
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
Attribute VB_Name = "frmJobworkBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Vch Type=NNNNSU/NNNNSC/NNNNSJ/NNNNPU/NNNNPC/NNNNPJ (U-Unit Cost C-Jobwork Unit Cost J-Jobwork P-Purchase S-Sale Q-Sales Quotation Z-Purchase Quotation) & BOM=NNNNXXXXXXXXXXXXFI/NNNNXXXXXXXXXXXXMF (MF/ME/CF/MO/BN/BM) & 01-Purchase 04-Sale 23-Purchase Quotation 24-Sales Quotation
Public VchType As String
Dim cnJobworkBill As New ADODB.Connection, cnTally As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset, rstJobworkBVList As New ADODB.Recordset, rstJobworkBVParent As New ADODB.Recordset, rstJobworkBVChild As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstTaxList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstNarrationList As New ADODB.Recordset, rstHSNCodeList As New ADODB.Recordset, rstOrderList As New ADODB.Recordset, rstVchSeriesList As New ADODB.Recordset, rstMaterialCentreList As New ADODB.Recordset, rstSalesTypeList As New ADODB.Recordset
Dim BuyerCode As String, TaxCode As String, ItemCode As String, RefCode As String, NarrationCode As String, HSNCode As String, ConsigneeCode As String, VchPrefix As String, TranType As String, oVchNo As String, oVchDate As Date, oVchSeriesCode As String, AutoVchNo As String, SalesTypeCode As String
Dim SortOrder, PrevStr, dblBookMark As Double, blnRecordExist As Boolean, EditMode As Boolean, VchSeries As String, VchSeriesCode As String, MaterialCentreCode As String, VchNumbering As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    VchType = Choose(Val(VchType), "SU", "SC", "SJ", "PU", "PC", "PJ", "QU", "QC", "QJ", "ZU", "ZC", "ZJ")
    Me.Caption = Switch(Left(VchType, 1) = "S", "Sales ", Left(VchType, 1) = "P", "Purchase ", Left(VchType, 1) = "Q", "Sales Quotation ", Left(VchType, 1) = "Z", "Purchase Quotation ") + "(" + IIf(Right(VchType, 1) = "U", "Unit Cost", IIf(Right(VchType, 1) = "C", "Job Work Unit Cost", "Job Work")) + ")"
    Mh3dLabel15.Caption = IIf(Left(VchType, 1) = "S", " Sales Type ", " Purchase Type ")
    TranType = IIf(Left(VchType, 1) = "S", "04", IIf(Left(VchType, 1) = "P", "01", IIf(Left(VchType, 1) = "Q", "24", "23")))
    cnJobworkBill.CursorLocation = adUseClient: cnTally.CursorLocation = adUseClient
    cnJobworkBill.Open cnDatabase.ConnectionString
    LoadMasterList
    rstJobworkBVList.Open "SELECT T.Code,T.Name,Date,T.Type,P.Name As PartyName,C.Name As ConsigneeName,Amount,(Select Name From VchSeriesMaster Where Code=VchSeries) As VchSeriesName,(Select Name From AccountMaster Where Code=MaterialCentre) As MaterialCentre,(Select Name From AccountMaster Where Code=SalesType) As SalesType FROM (JobworkBVParent T INNER JOIN AccountMaster P ON T.Party=P.Code) INNER JOIN AccountMaster C ON T.Consignee=C.Code INNER JOIN AccountMaster C1 ON T.MaterialCentre=C1.Code LEFT JOIN AccountMaster C2 ON T.SalesType=C2.Code WHERE RIGHT(Type,2)='" & VchType & "' AND FYCode='" & FYCode & "' ORDER BY T.Name", cnJobworkBill, adOpenKeyset, adLockPessimistic
    rstJobworkBVParent.CursorLocation = adUseClient
    rstJobworkBVList.Filter = adFilterNone
    If rstJobworkBVList.RecordCount > 0 Then rstJobworkBVList.MoveLast
    Set DataGrid1.DataSource = rstJobworkBVList
    BusySystemIndicator False
    SSTab1.Tab = 0
'    SortOrder = "Name"
     If FrmStockLedger.dSortBy = True Then
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
        Mh3dLabel5.Left = 4620: Mh3dLabel5.Width = 1215: Text2.Left = 5820: Text2.Width = 2850
        Mh3dLabel1(0).Left = 10960: MhDateInput1.Left = 11805
        Mh3dLabel6.Visible = False: cmbBillingType.Visible = False: Text5.Width = 7455
        cmbBillingType.ListIndex = 0
    End If
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
                If BusyIntegration Or TallyIntegration Then
                    If MsgBox("Are you sure to export the Voucher?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Export !") = vbYes Then PushVch
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
    VchSeriesCode = "": oVchSeriesCode = "": oVchNo = "": AutoVchNo = ""
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
        cmbBillingType.ListIndex = IIf(Mid(.Fields("Type").Value, 3, 2) = "10", 0, 1)
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
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        BuyerCode = slCode: Text3.Text = slName
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
        MhRealInput11.Value = MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput8.Value + MhRealInput10.Value + MhRealInput12.Value
    End With
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstJobworkBVList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstJobworkBVParent.State = adStateOpen Then rstJobworkBVParent.Close
    rstJobworkBVParent.Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstJobworkBVList.Fields("Code").Value) & "'", cnJobworkBill, adOpenKeyset, adLockOptimistic
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
            cnJobworkBill.Execute "UPDATE BookPOChild05 SET BilledMFB=BilledMFB" & Operation & Trim(Quantity) & ",DeliveredQuantityB=DeliveredQuantityB" & Operation & Trim(Quantity) & " WHERE (Code+'XXXXXXXXXXXXMF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            cnJobworkBill.Execute "UPDATE BookPOChild05 SET BilledMFC=BilledMFC" & Operation & Trim(Quantity) & " WHERE (Code+'XXXXXXXXXXXXMF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
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
            cnJobworkBill.Execute "UPDATE BookPOChild08 SET BilledBNB=BilledBNB" & Operation & Trim(Quantity) & ",DeliveredQuantityB=DeliveredQuantityB" & Operation & Trim(Quantity) & " WHERE (Code+'XXXXXXXXXXXXBN'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            cnJobworkBill.Execute "UPDATE BookPOChild08 SET BilledBNC=BilledBNC" & Operation & Trim(Quantity) & " WHERE (Code+'XXXXXXXXXXXXBN'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
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
'        SQL = "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.ActualQuantity-C.DeliveredQuantityC-C.BilledMFB", "C.DeliveredQuantityC-C.BilledMFC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild05 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+'XXXXXXMF') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE T.Code='" & VchNo & "'"
        SQL = "SELECT I.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.ActualQuantity-C.DeliveredQuantityC-C.BilledMFB", "C.DeliveredQuantityC-C.BilledMFC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,14)=R.Code+'XXXXXXXXXXXXMF') INNER JOIN BookPOChild05 C ON R.Code=C.Code) INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "'"
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.ActualQuantity-C.DeliveredQuantityC-C.BilledMEB", "C.DeliveredQuantityC-C.BilledMEC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild06 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+'XXXXXXME') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code  WHERE T.Code='" & VchNo & "'"
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.ActualQuantity-C.DeliveredQuantityC-C.BilledCFB", "C.DeliveredQuantityC-C.BilledCFC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild0901 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Book+'XXXXXXCF') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN BookMaster E ON C.Book=E.Code WHERE T.Code='" & VchNo & "'"
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT I.Name+'_'+E.Name+'_'+O.Name As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.Quantity-C.DeliveredQuantityC-C.BilledMOB", "C.DeliveredQuantityC-C.BilledMOC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+E.Code+O.Code+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild07 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+C.Operation+'MO') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code WHERE T.Code='" & VchNo & "'"
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT I.Name+'_Binding' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+" & IIf(cmbBillingType.ListIndex = 0, "C.ActualQuantity-C.DeliveredQuantityC-C.BilledBNB", "C.DeliveredQuantityC-C.BilledBNC") & " As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,14)=R.Code+'XXXXXXXXXXXXBN') INNER JOIN BookPOChild08 C ON R.Code=C.Code) INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "'"
        SQL = SQL + " UNION ALL "
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
        SQL = "SELECT DISTINCT P.Code+'XXXXXXXXXXXXFI' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/FI' As VchNo,P.Date As VchDate,I.Name As Item,P.EstQty01 As OrderedQty,P.BilledAllC As BilledQtyC,P.BilledAllB As BilledQtyD,P.DeliveredQuantityC As ChallanQty,P.DeliveredQuantityB As DirectQty FROM (BookPOParent P INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN BookPOChild0801 C ON P.Code=C.Code WHERE (P.BookPrinter='" & BuyerCode & "' OR P.TitlePrinter='" & BuyerCode & "' OR P.Laminator='" & BuyerCode & "' OR P.Binder='" & BuyerCode & "' OR C.Vendor='" & BuyerCode & "') AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)<>'FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(P.EstQty01-P.DeliveredQuantityC-P.BilledAllB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)=Pending Quantity
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(P.DeliveredQuantityC-P.BilledAllC)>0" 'Delivered(Challan)-Billed(Challan)=Pending Quantity
        End If
        SQL = SQL + " ORDER BY I.Name,P.Date,VchNo"
    ElseIf Right(VchType, 1) = "J" Then 'Jobwork
'       SQL = "SELECT P.Code+E.Code+'XXXXXXMF' As VchCode,LTRIM(P.Name)+'/'+" & iif(instr(1,"Q_Z",LEFT(VchType,1))>0,"LEFT(P.Type,1)","RIGHT(P.Type,1)") & "+'O/MF' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.BilledMFC As BilledQtyC,C.BilledMFB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.BookPrinter='" & BuyerCode & "' AND " & iif(instr(1,"Q_Z",LEFT(VchType,1))>0,"LEFT(P.Type,1)","RIGHT(P.Type,1)") & "='" & iif(instr(1,"Q_Z",LEFT(VchType,1))>0,"O",Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        SQL = "SELECT P.Code+'XXXXXXXXXXXXMF' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/MF' As VchNo,P.Date As VchDate,I.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.BilledMFC As BilledQtyC,C.BilledMFB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.BookPrinter='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledMFB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledMFC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT P.Code+E.Code+'XXXXXXME' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/ME' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.BilledMEC As BilledQtyC,C.BilledMEB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.TitlePrinter='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledMEB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledMEC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT P.Code+E.Code+'XXXXXXCF' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/CF' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.BilledCFC As BilledQtyC,C.BilledCFB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN BookMaster E ON C.Book=E.Code WHERE P.TitlePrinter='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledCFB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledCFC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT P.Code+E.Code+O.Code+'MO' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/MO' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_'+O.Name As Item,C.Quantity As OrderedQty,C.BilledMOC As BilledQtyC,C.BilledMOB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM (((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code WHERE P.Laminator='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.Quantity-C.DeliveredQuantityC-C.BilledMOB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledMOC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT P.Code+'XXXXXXXXXXXXBN' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/BN' As VchNo,P.Date As VchDate,I.Name+'_Binding' As Item,C.ActualQuantity As OrderedQty,C.BilledBNC As BilledQtyC,C.BilledBNB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Binder='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledBNB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledBNC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " UNION ALL "
        SQL = SQL + "SELECT P.Code+C.Item+'XXXXX'+C.Category+'BM' As VchCode,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/BM' As VchNo,P.Date As VchDate,I.Name+'_'+CASE WHEN C.Category='1' THEN O.Name WHEN C.Category='2' THEN R.Name ELSE U.Name END As Item,C.OrderQuantity As OrderedQty,C.BilledBMC As BilledQtyC,C.BilledBMB As BilledQtyD,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((((BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN OutsourceItemMaster O ON C.Category+C.Item='1'+O.Code) LEFT JOIN PaperMaster R ON C.Category+C.Item='2'+R.Code) LEFT JOIN BookMaster U ON C.Category+C.Item='3'+U.Code WHERE C.Vendor='" & BuyerCode & "' AND " & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "='" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "O", Left(VchType, 1)) & "' AND C.Amount<>0 " & _
                                "AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') AND "
        If cmbBillingType.ListIndex = 0 Then 'Direct
            SQL = SQL + "(C.OrderQuantity-C.DeliveredQuantityC-C.BilledBMB)>0" 'Ordered-Delivered(Challan)-Billed(Direct)
        ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
            SQL = SQL + "(C.DeliveredQuantityC-C.BilledBMC)>0" 'Delivered(Challan)-Billed(Challan)
        End If
        SQL = SQL + " ORDER BY Item,VchNo,VchDate"
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
                Pending = Val(.Fields("OrderedQty").Value) - Val(.Fields("ChallanQty").Value) - Val(.Fields("BilledQtyD").Value) 'Pending=Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                Pending = Val(.Fields("ChallanQty").Value) - Val(.Fields("BilledQtyC").Value) 'Pending=Delivered(Challan)-Billed(Challan)
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
        ElseIf Right(VchType, 1) = "J" Then
'           SQL = "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND((C.PrintAmount+Adjustment+C.PlateAmount+PAdjustment+PaperAmount+RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & iif(instr(1,"Q_Z",LEFT(VchType,1))>0,"LEFT(P.Type,1)","RIGHT(P.Type,1)") & "+'O/MF' As VchNo,P.Code+E.Code+'XXXXXXMF' As VchCode,"
            SQL = "SELECT I.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND((C.PrintAmount1+C.PrintAmount2+C.PrintAmount4+Adjustment+C.PlateAmount1+C.PlateAmount2+C.PlateAmount4+PAdjustment+PaperAmount1+PaperAmount2+PaperAmount4+RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/MF' As VchNo,P.Code+'XXXXXXXXXXXXMF' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledMFB) As PendingQty" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledMFC) As PendingQty" 'Delivered(Challan)-Billed(Challan)
            End If
'           SQL = SQL + " FROM (((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+'XXXXXXMF' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+'XXXXXXXXXXXXMF' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND((C.PrintAmount+C.Adjustment+C.PlateAmount+C.PAdjustment+C.PaperAmount+C.RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/ME' As VchNo,P.Code+E.Code+'XXXXXXME' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledMEB) As PendingQty" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledMEC) As PendingQty" 'Delivered(Challan)-Billed(Challan)
            End If
            SQL = SQL + " FROM (((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+'XXXXXXME' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND(((C1.PrintAmount+C1.Adjustment+C1.PlateAmount+C1.PAdjustment+C1.PaperAmount+C1.RAdjustment)/(SELECT SUM(ActualQuantity) FROM BookPOChild0901 WHERE Code=P.Code))*C.ActualQuantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/CF' As VchNo,P.Code+E.Code+'XXXXXXCF' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledCFB) As Pending" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledCFC) As Pending" 'Delivered(Challan)-Billed(Challan)
            End If
            SQL = SQL + "  FROM ((((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C ON C.Code=C1.Code) INNER JOIN BookMaster I ON  P.Book=I.Code) INNER JOIN BookMaster E ON C.Book=E.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+'XXXXXXCF' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+E.Name+'_'+O.Name As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND((C.Amount+C.Adjustment)/C.Quantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/MO' As VchNo,P.Code+E.Code+O.Code+'MO' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.Quantity-C.DeliveredQuantityC-C.BilledMOB) As PendingQty" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledMOC) As PendingQty" 'Delivered(Challan)-Billed(Challan)
            End If
            SQL = SQL + " FROM ((((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+O.Code+'MO' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_Binding' As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND((C.BillAmount-C.VAT)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/BN' As VchNo,P.Code+'XXXXXXXXXXXXBN' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.ActualQuantity-C.DeliveredQuantityC-C.BilledBNB) As PendingQty" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledBNC) As PendingQty" 'Delivered(Challan)-Billed(Challan)
            End If
            SQL = SQL + " FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+'XXXXXXXXXXXXBN' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+CASE WHEN C.Category='1' THEN O.Name WHEN C.Category='2' THEN R.Name ELSE U.Name END As ItemName,I.Code As ItemCode,H.Code As HSNCode,H.Name As HSNName,ROUND(C.Amount/C.OrderQuantity,3) As UnitRate,P.ProfitMargin,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+" & IIf(InStr(1, "Q_Z", Left(VchType, 1)) > 0, "LEFT(P.Type,1)", "RIGHT(P.Type,1)") & "+'O/BM' As VchNo,P.Code+C.Item+'XXXXX'+C.Category+'BM' As VchCode,"
            If cmbBillingType.ListIndex = 0 Then 'Direct
                SQL = SQL + "(C.OrderQuantity-C.DeliveredQuantityC-C.BilledBMB) As PendingQty" 'Ordered-Delivered(Challan)-Billed(Direct)
            ElseIf cmbBillingType.ListIndex = 1 Then 'Against Challan
                SQL = SQL + "(C.DeliveredQuantityC-C.BilledBMC) As PendingQty" 'Delivered(Challan)-Billed(Challan)
            End If
            SQL = SQL + " FROM (((((BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN OutsourceItemMaster O ON C.Category+C.Item='1'+O.Code) LEFT JOIN PaperMaster R ON C.Category+C.Item='2'+R.Code) LEFT JOIN BookMaster U ON C.Category+C.Item='3'+U.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+C.Item+'XXXXX'+C.Category+'BM' IN (" & FrmOrderList.VchCodeList & ")"
            SQL = SQL + " ORDER BY ItemName,VchNo"
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
                Call CalculateTotal
            End With
        End If
    End If
    CloseForm FrmOrderList
End Sub
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
    rstTaxList.Open "SELECT Name As Col0,[IGST%],[SGST%],[CGST%],Region,Code FROM TaxMaster ORDER BY Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
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
    rstVchSeriesList.Open "SELECT Name As Col0,Prefix,Suffix,VchNumbering,Code FROM VchSeriesMaster WHERE VchType='" & Switch(Left(VchType, 1) = "S", "04", Left(VchType, 1) = "P", "01", Left(VchType, 1) = "Z", "23", Left(VchType, 1) = "Q", "24") & VchType & "' ORDER BY Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
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
    Dim xmlstr, SaleAccount, UOM, i
    With rstCompanyMaster
        If .State = adStateOpen Then .Close
        .Open "SELECT VchSeries,Account,UOM FROM AppConfig WHERE VchType='" & VchType & "'", cnJobworkBill, adOpenKeyset, adLockReadOnly
        VchSeries = .Fields("VchSeries").Value: SaleAccount = .Fields("Account").Value: UOM = .Fields("UOM").Value
    End With
    With rstJobworkBVChild
        If .State = adStateOpen Then .Close
        xmlstr = "SELECT LTRIM(H.Name) As BillNo,H.Date As BillDate,ISNULL(M.PrintName,'Main Location') As MatCentre," + _
                        "B.PrintName As Buyer,B.Address1 As bAddress1,B.Address2 As bAddress2,B.Address3 As bAddress3,B.Address4 As bAddress4,B.TIN As bGSTIN,C.PrintName As Consignee,C.Address1 As cAddress1,C.Address2 As cAddress2,C.Address3 As cAddress3,C.Address4 As cAddress4,C.TIN As cGSTIN," + _
                        "H.TaxableAmount,H.[Rebate%],H.Rebate,H.Freight,H.Adjustment,H.Tax,H.[IGST%],H.IGST,H.[SGST%],H.SGST,H.[CGST%],H.CGST,H.Amount As FinalAmount,H.Remarks," + _
                        "I.BusyCode As Item,D.Rate,D.[Disc%],D.Quantity,D.Amount " & _
                        "FROM ((((JobWorkBVParent H INNER JOIN AccountMaster B ON H.Party=B.Code) INNER JOIN AccountMaster C ON H.Consignee=C.Code) LEFT JOIN AccountMaster M ON H.MaterialCentre=M.Code) INNER JOIN JobWorkBVChild D ON H.Code=D.Code) INNER JOIN BookMaster I ON D.Item=I.Code " + _
                        "WHERE H.Code='" + rstJobworkBVList.Fields("Code").Value + "'"
        .Open xmlstr, cnJobworkBill, adOpenKeyset, adLockReadOnly
        xmlstr = ""
        If TallyIntegration Then
            Dim Dom As Object
            Set Dom = CreateObject("MSXML2.DomDocument")
            Dom.async = False
            xmlstr = xmlstr + "<ENVELOPE>"
            xmlstr = xmlstr + "<HEADER>"
            xmlstr = xmlstr + "<TALLYREQUEST>Import Data</TALLYREQUEST>"
            xmlstr = xmlstr + "</HEADER>"
            xmlstr = xmlstr + "<BODY>"
            xmlstr = xmlstr + "<IMPORTDATA>"
            xmlstr = xmlstr + "<REQUESTDESC>"
            xmlstr = xmlstr + "<REPORTNAME>Vouchers</REPORTNAME>"
            xmlstr = xmlstr + "<STATICVARIABLES>"
            xmlstr = xmlstr + "<SVCURRENTCOMPANY>##SVCURRENTCOMPANY</SVCURRENTCOMPANY>" '##SVCURRENTCOMPANY-Current Open Company
            xmlstr = xmlstr + "</STATICVARIABLES>"
            xmlstr = xmlstr + "</REQUESTDESC>"
            xmlstr = xmlstr + "<REQUESTDATA>"
            xmlstr = xmlstr + "<TALLYMESSAGE xmlns:UDF=""TallyUDF"">"
            xmlstr = xmlstr + "<VOUCHER ACTION=""Create"">"
            xmlstr = xmlstr + "<VOUCHERTYPENAME>" + Replace(Trim(VchSeries), "&", "&amp;") + "</VOUCHERTYPENAME>"
            xmlstr = xmlstr + "<VOUCHERNUMBER>" + Replace(Trim(.Fields("BillNo").Value), "&", "&amp;") + "</VOUCHERNUMBER>" 'Vch No.
            xmlstr = xmlstr + "<DATE>" + Format(.Fields("BillDate").Value, "yyyyMMdd") + "</DATE>" 'Vch Date
            If Not CheckEmpty(Trim(.Fields("Remarks").Value), False) Then xmlstr = xmlstr + "<NARRATION>" + Replace(Trim(.Fields("Remarks").Value), "&", "&amp;") + "</NARRATION>" 'Narration
            xmlstr = xmlstr + "<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>"
            xmlstr = xmlstr + "<ISINVOICE>Yes</ISINVOICE>"
            xmlstr = xmlstr + "<HASDISCOUNTS>Yes</HASDISCOUNTS>"
            xmlstr = xmlstr + "<PARTYNAME>" + Replace(Trim(.Fields("Buyer").Value), "&", "&amp;") + "</PARTYNAME>" 'Buyer Info
            If Not CheckEmpty(Trim(.Fields("bAddress1").Value) + Trim(.Fields("bAddress2").Value) + Trim(.Fields("bAddress3").Value) + Trim(.Fields("bAddress4").Value), False) Then
                xmlstr = xmlstr + "<ADDRESS.LIST TYPE=""String"">"
                If Not CheckEmpty(Trim(.Fields("bAddress1").Value), False) Then xmlstr = xmlstr + "<ADDRESS>" + Replace(Trim(.Fields("bAddress1").Value), "&", "&amp;") + "</ADDRESS>"
                If Not CheckEmpty(Trim(.Fields("bAddress2").Value), False) Then xmlstr = xmlstr + "<ADDRESS>" + Replace(Trim(.Fields("bAddress2").Value), "&", "&amp;") + "</ADDRESS>"
                If Not CheckEmpty(Trim(.Fields("bAddress3").Value), False) Then xmlstr = xmlstr + "<ADDRESS>" + Replace(Trim(.Fields("bAddress3").Value), "&", "&amp;") + "</ADDRESS>"
                If Not CheckEmpty(Trim(.Fields("bAddress4").Value), False) Then xmlstr = xmlstr + "<ADDRESS>" + Replace(Trim(.Fields("bAddress4").Value), "&", "&amp;") + "</ADDRESS>"
                xmlstr = xmlstr + "</ADDRESS.LIST>"
            End If
            If Not CheckEmpty(Trim(.Fields("bGSTIN").Value), False) Then xmlstr = xmlstr + "<PARTYGSTIN>" + Replace(Trim(.Fields("bGSTIN").Value), "&", "&amp;") + "</PARTYGSTIN>"
            xmlstr = xmlstr + "<BASICBUYERNAME>" + Replace(Trim(.Fields("Consignee").Value), "&", "&amp;") + "</BASICBUYERNAME>" 'Consignee Info
            If Not CheckEmpty(Trim(.Fields("cAddress1").Value) + Trim(.Fields("cAddress2").Value) + Trim(.Fields("cAddress3").Value) + Trim(.Fields("cAddress4").Value), False) Then
                xmlstr = xmlstr + "<BASICBUYERADDRESS.LIST TYPE=""String"">"
                If Not CheckEmpty(Trim(.Fields("cAddress1").Value), False) Then xmlstr = xmlstr + "<BASICBUYERADDRESS>" + Replace(Trim(.Fields("cAddress1").Value), "&", "&amp;") + "</BASICBUYERADDRESS>"
                If Not CheckEmpty(Trim(.Fields("cAddress2").Value), False) Then xmlstr = xmlstr + "<BASICBUYERADDRESS>" + Replace(Trim(.Fields("cAddress2").Value), "&", "&amp;") + "</BASICBUYERADDRESS>"
                If Not CheckEmpty(Trim(.Fields("cAddress3").Value), False) Then xmlstr = xmlstr + "<BASICBUYERADDRESS>" + Replace(Trim(.Fields("cAddress3").Value), "&", "&amp;") + "</BASICBUYERADDRESS>"
                If Not CheckEmpty(Trim(.Fields("cAddress4").Value), False) Then xmlstr = xmlstr + "<BASICBUYERADDRESS>" + Replace(Trim(.Fields("cAddress4").Value), "&", "&amp;") + "</BASICBUYERADDRESS>"
                xmlstr = xmlstr + "</BASICBUYERADDRESS.LIST>"
            End If
            If Not CheckEmpty(Trim(.Fields("cGSTIN").Value), False) Then xmlstr = xmlstr + "<CONSIGNEEGSTIN>" + Replace(Trim(.Fields("cGSTIN").Value), "&", "&amp;") + "</CONSIGNEEGSTIN>"
            .MoveFirst
            Do Until .EOF
                xmlstr = xmlstr + "<INVENTORYENTRIES.LIST>"
                xmlstr = xmlstr + "<STOCKITEMNAME>" + Replace(Trim(.Fields("Item").Value), "&", "&amp;") + "</STOCKITEMNAME>"
                xmlstr = xmlstr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                xmlstr = xmlstr + "<RATE>" + Format(Val(.Fields("Rate").Value), "0.00") + "/" + Replace(UOM, "&", "&amp;") + "</RATE>"
                xmlstr = xmlstr + "<DISCOUNT>" + Format(Val(.Fields("Disc%").Value), "0.00") + "</DISCOUNT>"
                xmlstr = xmlstr + "<AMOUNT>" + Format(Val(.Fields("Amount").Value), "0.00") + "</AMOUNT>"
                xmlstr = xmlstr + "<ACTUALQTY>" + Format(Val(.Fields("Quantity").Value), "0.00") + " " + Replace(UOM, "&", "&amp;") + "</ACTUALQTY>"
                xmlstr = xmlstr + "<BILLEDQTY>" + Format(Val(.Fields("Quantity").Value), "0.00") + " " + Replace(UOM, "&", "&amp;") + "</BILLEDQTY>"
                xmlstr = xmlstr + "<BATCHALLOCATIONS.LIST>"
                xmlstr = xmlstr + "<GODOWNNAME>" + Replace(Trim(.Fields("MatCentre").Value), "&", "&amp;") + "</GODOWNNAME>"
                xmlstr = xmlstr + "<DESTINATIONGODOWNNAME>" + Replace(Trim(.Fields("MatCentre").Value), "&", "&amp;") + "</DESTINATIONGODOWNNAME>"
                xmlstr = xmlstr + "</BATCHALLOCATIONS.LIST>"
                xmlstr = xmlstr + "<ACCOUNTINGALLOCATIONS.LIST>"
                xmlstr = xmlstr + "<LEDGERNAME>" + Replace(SaleAccount, "&", "&amp;") + "</LEDGERNAME>"
                xmlstr = xmlstr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                xmlstr = xmlstr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                xmlstr = xmlstr + "<AMOUNT>" + Format(Val(.Fields("Amount").Value), "0.00") + "</AMOUNT>"
                xmlstr = xmlstr + "</ACCOUNTINGALLOCATIONS.LIST>"
                xmlstr = xmlstr + "</INVENTORYENTRIES.LIST>"
                .MoveNext
            Loop
            .MoveFirst
            xmlstr = xmlstr + "<LEDGERENTRIES.LIST>"
            xmlstr = xmlstr + "<LEDGERNAME>" + Replace(Trim(.Fields("Buyer").Value), "&", "&amp;") + "</LEDGERNAME>" 'Buyer Name
            xmlstr = xmlstr + "<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>"
            xmlstr = xmlstr + "<ISPARTYLEDGER>Yes</ISPARTYLEDGER>"
            xmlstr = xmlstr + "<AMOUNT>" + Trim(0 - Val(.Fields("FinalAmount").Value)) + "</AMOUNT>" 'Vch Amount
            xmlstr = xmlstr + "<BILLALLOCATIONS.LIST>"
            xmlstr = xmlstr + "<NAME>" + Replace(Trim(.Fields("BillNo").Value), "&", "&amp;") + "</NAME>" 'Vch No.
            xmlstr = xmlstr + "<BILLTYPE></BILLTYPE>"
            xmlstr = xmlstr + "<AMOUNT>" + Trim(0 - Val(.Fields("FinalAmount").Value)) + "</AMOUNT>" 'Vch Amount
            xmlstr = xmlstr + "</BILLALLOCATIONS.LIST>"
            xmlstr = xmlstr + "</LEDGERENTRIES.LIST>"
            If Val(.Fields("Rebate").Value) > 0 Then
                xmlstr = xmlstr + "<LEDGERENTRIES.LIST>"
                xmlstr = xmlstr + "<BASICRATEOFINVOICETAX.LIST TYPE=""Number"">"
                xmlstr = xmlstr + "<BASICRATEOFINVOICETAX> " + Trim(Format(0 - Val(.Fields("Rebate%").Value), "0.00")) + "</BASICRATEOFINVOICETAX>"
                xmlstr = xmlstr + "</BASICRATEOFINVOICETAX.LIST>"
                xmlstr = xmlstr + "<LEDGERNAME>Discount</LEDGERNAME>"
                xmlstr = xmlstr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                xmlstr = xmlstr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                xmlstr = xmlstr + "<AMOUNT>" + Trim(Format(0 - Val(.Fields("Rebate").Value), "0.00")) + "</AMOUNT>"
                xmlstr = xmlstr + "<VATEXPAMOUNT>" + Trim(0 - Format(Val(.Fields("Rebate").Value), "0.00")) + "</VATEXPAMOUNT>"
                xmlstr = xmlstr + "</LEDGERENTRIES.LIST>"
            End If
            If Val(.Fields("Freight").Value) > 0 Then
                xmlstr = xmlstr + "<LEDGERENTRIES.LIST>"
                xmlstr = xmlstr + "<LEDGERNAME>" + Replace("Packing & Forwarding Charges", "&", "&amp;") + "</LEDGERNAME>"
                xmlstr = xmlstr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                xmlstr = xmlstr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                xmlstr = xmlstr + "<AMOUNT>" + Format(Val(.Fields("Freight").Value), "0.00") + "</AMOUNT>"
                xmlstr = xmlstr + "<VATEXPAMOUNT>" + Format(Val(.Fields("Freight").Value), "0.00") + "</VATEXPAMOUNT>"
                xmlstr = xmlstr + "</LEDGERENTRIES.LIST>"
            End If
            rstTaxList.MoveFirst
            rstTaxList.Find "[Code]='" & .Fields("Tax").Value & "'"
            If rstTaxList.Fields("Region").Value = "I" Then
                xmlstr = xmlstr + "<LEDGERENTRIES.LIST>"
                xmlstr = xmlstr + "<BASICRATEOFINVOICETAX.LIST TYPE=""Number"">"
                xmlstr = xmlstr + "<BASICRATEOFINVOICETAX> " + Trim(Format(Val(.Fields("IGST%").Value), "0.00")) + "</BASICRATEOFINVOICETAX>"
                xmlstr = xmlstr + "</BASICRATEOFINVOICETAX.LIST>"
                xmlstr = xmlstr + "<LEDGERNAME>IGST-" + IIf(Val(.Fields("IGST%").Value) = 0, "Exempted", Trim(Format(Val(.Fields("IGST%").Value), "0.00")) + "%") + "</LEDGERNAME>"
                xmlstr = xmlstr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                xmlstr = xmlstr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                xmlstr = xmlstr + "<AMOUNT>" + Trim(Format(Val(.Fields("IGST").Value), "0.00")) + "</AMOUNT>"
                xmlstr = xmlstr + "<VATEXPAMOUNT>" + Trim(Format(Val(.Fields("IGST").Value), "0.00")) + "</VATEXPAMOUNT>"
                xmlstr = xmlstr + "</LEDGERENTRIES.LIST>"
            Else
                xmlstr = xmlstr + "<LEDGERENTRIES.LIST>"
                xmlstr = xmlstr + "<BASICRATEOFINVOICETAX.LIST TYPE=""Number"">"
                xmlstr = xmlstr + "<BASICRATEOFINVOICETAX> " + Trim(Format(Val(.Fields("CGST%").Value), "0.00")) + "</BASICRATEOFINVOICETAX>"
                xmlstr = xmlstr + "</BASICRATEOFINVOICETAX.LIST>"
                xmlstr = xmlstr + "<LEDGERNAME>CGST-" + IIf(Val(.Fields("CGST%").Value) = 0, "Exempted", Trim(Format(Val(.Fields("CGST%").Value), "0.00")) + "%") + "</LEDGERNAME>"
                xmlstr = xmlstr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                xmlstr = xmlstr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                xmlstr = xmlstr + "<AMOUNT>" + Trim(Format(Val(.Fields("CGST").Value), "0.00")) + "</AMOUNT>"
                xmlstr = xmlstr + "<VATEXPAMOUNT>" + Trim(Format(Val(.Fields("CGST").Value), "0.00")) + "</VATEXPAMOUNT>"
                xmlstr = xmlstr + "</LEDGERENTRIES.LIST>"
                xmlstr = xmlstr + "<LEDGERENTRIES.LIST>"
                xmlstr = xmlstr + "<BASICRATEOFINVOICETAX.LIST TYPE=""Number"">"
                xmlstr = xmlstr + "<BASICRATEOFINVOICETAX> " + Trim(Format(Val(.Fields("SGST%").Value), "0.00")) + "</BASICRATEOFINVOICETAX>"
                xmlstr = xmlstr + "</BASICRATEOFINVOICETAX.LIST>"
                xmlstr = xmlstr + "<LEDGERNAME>SGST-" + IIf(Val(.Fields("SGST%").Value) = 0, "Exempted", Trim(Format(Val(.Fields("SGST%").Value), "0.00")) + "%") + "</LEDGERNAME>"
                xmlstr = xmlstr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                xmlstr = xmlstr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                xmlstr = xmlstr + "<AMOUNT>" + Trim(Format(Val(.Fields("SGST").Value), "0.00")) + "</AMOUNT>"
                xmlstr = xmlstr + "<VATEXPAMOUNT>" + Trim(Format(Val(.Fields("SGST").Value), "0.00")) + "</VATEXPAMOUNT>"
                xmlstr = xmlstr + "</LEDGERENTRIES.LIST>"
            End If
            If Val(.Fields("Adjustment").Value) <> 0 Then
                xmlstr = xmlstr + "<LEDGERENTRIES.LIST>"
                xmlstr = xmlstr + "<LEDGERNAME>Round Off</LEDGERNAME>"
                xmlstr = xmlstr + "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"
                xmlstr = xmlstr + "<ISPARTYLEDGER>No</ISPARTYLEDGER>"
                xmlstr = xmlstr + "<AMOUNT>" + Format(Val(.Fields("Adjustment").Value), "0.00") + "</AMOUNT>"
                xmlstr = xmlstr + "<VATEXPAMOUNT>" + Format(Val(.Fields("Adjustment").Value), "0.00") + "</VATEXPAMOUNT>"
                xmlstr = xmlstr + "</LEDGERENTRIES.LIST>"
            End If
            xmlstr = xmlstr + "</VOUCHER>"
            xmlstr = xmlstr + "</TALLYMESSAGE>"
            xmlstr = xmlstr + "</REQUESTDATA>"
            xmlstr = xmlstr + "</IMPORTDATA>"
            xmlstr = xmlstr + "</BODY>"
            xmlstr = xmlstr + "</ENVELOPE>"
            Dim WinHttpReq As Object
            Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
            With WinHttpReq
                .Open "POST", "http://localhost:" + ReadFromFile("Tally Port"), False
                Do While True
                    On Error Resume Next
                    DelOldVch False
                    .Send xmlstr
                    If Err.Number = -2147012867 Then
                        If MsgBox(Err.Description + "Tally is not open. Would you like to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Proceed !") = vbNo Then Exit Do
                    Else
                        .WaitForResponse 4000
                        Dom.Loadxml .responseText
                        If Dom.SelectSingleNode("//CREATED").Text = "1" Then
                            MsgBox "Voucher Exported to Tally !!!", vbInformation, App.Title: Exit Do
                        Else
                            If MsgBox(Dom.SelectSingleNode("//LINEERROR").Text + " Would you like to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Proceed !") = vbNo Then Exit Do
                        End If
                    End If
                Loop
            End With
        End If
    End With
End Sub
Private Sub DelOldVch(ByVal dspMsg As Boolean)
    Dim xmlstr
    If TallyIntegration Then
        Dim Dom As Object
        Set Dom = CreateObject("MSXML2.DomDocument")
        Dom.async = False
        xmlstr = xmlstr + "<ENVELOPE>"
        xmlstr = xmlstr + "<HEADER>"
        xmlstr = xmlstr + "<TALLYREQUEST>Import Data</TALLYREQUEST>"
        xmlstr = xmlstr + "</HEADER>"
        xmlstr = xmlstr + "<BODY>"
        xmlstr = xmlstr + "<IMPORTDATA>"
        xmlstr = xmlstr + "<REQUESTDESC>"
        xmlstr = xmlstr + "<REPORTNAME>Vouchers</REPORTNAME>"
        xmlstr = xmlstr + "<STATICVARIABLES>"
        xmlstr = xmlstr + "<SVCURRENTCOMPANY>##SVCURRENTCOMPANY</SVCURRENTCOMPANY>" '##SVCURRENTCOMPANY-Current Open Company
        xmlstr = xmlstr + "</STATICVARIABLES>"
        xmlstr = xmlstr + "</REQUESTDESC>"
        xmlstr = xmlstr + "<REQUESTDATA>"
        xmlstr = xmlstr + "<TALLYMESSAGE xmlns:UDF=""TallyUDF"">"
        xmlstr = xmlstr + "<VOUCHER DATE='" & Format(oVchDate, "yyyyMMdd") & "' TAGNAME = ""Voucher Number"" TAGVALUE='" & oVchNo & "' ACTION=""Delete"">"
        xmlstr = xmlstr + "<VOUCHERTYPENAME>" + VchSeries + "</VOUCHERTYPENAME>"
        xmlstr = xmlstr + "</VOUCHER>"
        xmlstr = xmlstr + "</TALLYMESSAGE>"
        xmlstr = xmlstr + "</REQUESTDATA>"
        xmlstr = xmlstr + "</IMPORTDATA>"
        xmlstr = xmlstr + "</BODY>"
        xmlstr = xmlstr + "</ENVELOPE>"
        Dim WinHttpReq As Object
        Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
        With WinHttpReq
            .Open "POST", "http://localhost:" + ReadFromFile("Tally Port"), False
            Do While True
                On Error Resume Next
                .Send xmlstr
                If Err.Number = -2147012867 Then
                    If MsgBox(Err.Description + "Tally is not open. Would you like to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Proceed !") = vbNo Then Exit Do
                Else
                    .WaitForResponse 4000
                    Dom.Loadxml .responseText
                    If Dom.SelectSingleNode("//DELETED").Text = "1" Then
                        If dspMsg Then MsgBox "Voucher from Tally Deleted !!!", vbInformation, App.Title: Exit Do
                    Else
                        If Dom.SelectSingleNode("//LINEERROR").Text = "Voucher does not exist!" Then
                            Exit Do
                        Else
                            If MsgBox(Dom.SelectSingleNode("//LINEERROR").Text + " Would you like to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Proceed !") = vbNo Then Exit Do
                        End If
                    End If
                End If
            Loop
        End With
    End If
End Sub
Public Sub PrintJobworkBillVch(ByVal VchCode As String, ByVal VchType As String, ByVal BillType As String, Optional ByVal OutputType As String)
    Dim SQL As String, FS01 As String, FS02 As String, FS03 As String, FS04 As String, FS05 As String, FS06 As String, FS07 As String, FS08 As String, FS09 As String, FS10 As String, FS11 As String, FS12 As String, FS13 As String, FS14 As String, FS15 As String, FS16 As String, FS17 As String
    FS01 = "'Finish Size: '+LTRIM(S.PrintName)+IIf(I.Pages = 0, '', ', '+LTRIM(I.Pages)+' pages/'+LTRIM(I.Forms)+'f ('+LTRIM(IIF(OneColorForms<>0,LTRIM(OneColorForms)+'f-1Col','')+' '+IIF(TwoColorForms<>0,LTRIM(TwoColorForms)+'f-2Col','')+' '+IIF(FourColorForms<>0,LTRIM(FourColorForms)+'f-4Col',''))+')')" 'Direct Billing
    FS02 = "'Finish Size: '+LTRIM(S.PrintName)+IIf(I.Pages = 0 AND (ISNull(C2.Pages1) = True OR C2.Pages1+C2.Pages2+C2.Pages4 = 0), '', ', '+LTRIM(IIF(ISNULL(C2.Pages1)=True,I.Pages,C2.Pages1+C2.Pages2+C2.Pages4))+' pages/'+LTRIM(IIF(ISNULL(C2.Pages1)=True,I.Forms,C2.Forms1+C2.Forms2+C2.Forms4))+'f ('+IIF(ISNULL(C2.Pages1)=True,LTRIM(IIF(I.OneColorForms<>0,LTRIM(I.OneColorForms)+'f-1Col','')+' '+IIF(I.TwoColorForms<>0,LTRIM(I.TwoColorForms)+'f-2Col','')+' '+IIF(I.FourColorForms<>0,LTRIM(I.FourColorForms)+'f-4Col','')),LTRIM(IIF(C2.Forms1<>0,LTRIM(C2.Forms1)+'f-1Col','')+' '+IIF(C2.Forms2<>0,LTRIM(C2.Forms2)+'f-2Col','')+' '+IIF(C2.Forms4<>0,LTRIM(C2.Forms4)+'f-4Col','')))+')')" 'Billing Against Sale Order
    FS03 = "'MF-Text Plates: '+LTRIM((C2.[TotalPlates1-¼]+C2.[TotalPlates1-½]+C2.[TotalPlates1-1]+C2.[RevisedPlates1])*1)+ '-plates- 1Col, '+LTRIM((C2.[TotalPlates2-¼]+C2.[TotalPlates2-½]+C2.[TotalPlates2-1]+C2.[RevisedPlates2])*2)+ '-plates- 2Col, '+LTRIM((C2.[TotalPlates4-¼]+C2.[TotalPlates4-½]+C2.[TotalPlates4-1]+C2.[RevisedPlates4])*4)+ '-plates- 4Col  = ' +LTRIM(((C2.[TotalPlates1-¼]+C2.[TotalPlates1-½]+C2.[TotalPlates1-1]+C2.[RevisedPlates1])*1)+((C2.[TotalPlates2-¼]+C2.[TotalPlates2-½]+C2.[TotalPlates2-1]+C2.[RevisedPlates2])*2)+((C2.[TotalPlates4-¼]+C2.[TotalPlates4-½]+C2.[TotalPlates4-1]+C2.[RevisedPlates4])*4))+'  Nos.'"
    FS04 = "'MF-Text Ptg.  : '+LTRIM(Forms1)+ 'forms- 1Col, '+LTRIM(Forms2)+ 'forms- 2Col, '+LTRIM(Forms4)+ 'forms- 4Col  = ' +LTRIM(Forms1+Forms2+forms4)+'  Nos.'"
    FS05 = "'ME-Title Plates: '+LTRIM(C3.TotalPlates)+ '-Plates ( '+LTRIM(FrontPrintingType)+' + '+LTRIM(BackPrintingType)+' ) - Color'"
    FS06 = "'ME-Title Ptg.   : '+'( '+LTRIM(FrontPrintingType)+' + '+LTRIM(BackPrintingType)+' ) - Color Printing'"
    FS07 = "'CF-Title Plates: '+LTRIM(C5.TotalPlates)+ '-Plates ( '+LTRIM(FrontPrintingColor)+' + '+LTRIM(BackPrintingColor)+' ) - Color'"
    FS08 = "'CF-Title Ptg. : '+'( '+LTRIM(FrontPrintingColor)+' + '+LTRIM(BackPrintingColor)+' ) - Color Printing'"
    FS09 = "'Misc. Operations : '+'Text & Title Finishing'"
    FS10 = "'Binding : '+'( '+LTRIM(C4.BindingType)+')'"
    FS11 = "'BOM : '+' BOM Items' "
    FS12 = "'Paper-(1Col) : '+LTRIM(PM1.Name)"
    FS13 = "'Paper-(2Col) : '+LTRIM(PM2.Name)"
    FS14 = "'Paper-(4Col) : '+LTRIM(PM4.Name)"
    FS15 = "'Paper-(ME) : '+LTRIM(PM5.Name)"
    FS16 = "'Paper-(CF) : '+LTRIM(PM6.Name)"
    FS17 = "'Paper : '+' Total Paper Value '"
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    If rstJobworkBVChild.State = adStateOpen Then rstJobworkBVChild.Close
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,EMail,Website,GSTIN,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,BankName,AccountNo,IFSC,Prefix,Suffix FROM CompanyMaster P INNER JOIN CompChild C ON P.Code=C.Code WHERE VchType=" & IIf(Right(VchType, 2) = "SU", 8, IIf(Right(VchType, 2) = "SJ", 9, 23)), cnJobworkBill, adOpenKeyset, adLockReadOnly
    rstCompanyMaster.ActiveConnection = Nothing
    SQL = "SELECT LTRIM(P.Name) +'/' +'" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "' As BillNo,P.Date As BillDate," & _
                "A.PrintName As Party,A.Address1 As PartyAddress1,A.Address2 As PartyAddress2,A.Address3 As PartyAddress3,A.Address4 As PartyAddress4,A.TIN As PartyGSTIN,C.PrintName As Consignee,C.Address1 As ConsigneeAddress1,C.Address2 As ConsigneeAddress2,C.Address3 As ConsigneeAddress3,C.Address4 As ConsigneeAddress4,C.TIN As ConsigneeGSTIN," & _
                "P.[Rebate%],P.Rebate,P.Freight,P.Adjustment,P.TaxableAmount,P.[IGST%],P.IGST,P.[SGST%],P.SGST,P.[CGST%],P.CGST,P.Amount As TotalAmount,P.Remarks," & _
                "N.PrintName As Narration,I.PrintName As Item,H.PrintName As HSNCode,C1.Quantity,C1.Rate,C1.Amount,N.Name As SrNo,LTRIM(C1.Code)+LTRIM(C1.SrNo) As Ref, "
        cnJobworkBill.CommandTimeout = 300
If Right(VchType, 2) = "SC" Or Right(VchType, 2) = "SJ" Or Right(VchType, 2) = "PC" Or Right(VchType, 2) = "PJ" Or Right(VchType, 2) = "QC" Or Right(VchType, 2) = "QJ" Or Right(VchType, 2) = "ZC" Or Right(VchType, 2) = "ZJ" Then
        rstJobworkBVChild.Open SQL & "'' As cmbTitle," & FS01 & ",LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05," & _
                "IIF(Right(C1.BOM,2)='MF'," & FS03 & ",IIF(Right(C1.BOM,2)='ME'," & FS05 & ",IIF(Right(C1.BOM,2)='CF'," & FS07 & ",IIF(Right(C1.BOM,2)='MO'," & FS09 & ",IIF(Right(C1.BOM,2)='BP'," & FS10 & ",IIF(Right(C1.BOM,2)='BM'," & FS11 & ",''))))))," & _
                "IIF(Right(C1.BOM,2)='MF'," & FS04 & ",IIF(Right(C1.BOM,2)='ME'," & FS06 & ",IIF(Right(C1.BOM,2)='CF'," & FS08 & ",'')))," & _
                "IIF(Right(C1.BOM,2)='MF'," & FS12 & ",IIF(Right(C1.BOM,2)='ME'," & FS15 & ",IIF(Right(C1.BOM,2)='CF'," & FS16 & ",'')))," & _
                "IIF(Right(C1.BOM,2)='MF'," & FS13 & ",'')," & _
                "IIF(Right(C1.BOM,2)='MF'," & FS14 & ",'')," & _
                "IIF(Right(C1.BOM,2)='MF',((C2.PBillAmount-C2.Pvat)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='ME',((C3.PBillAmount-C3.Pvat)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='CF',((C5.PlateAmount-C5.PGST)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='MO',0,IIF(Right(C1.BOM,2)='BP',0,IIF(Right(C1.BOM,2)='BM',0,0))))))," & _
                "IIF(Right(C1.BOM,2)='MF',((C2.BillAmount-C2.vat)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='ME',((C3.BillAmount-C3.vat)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='CF',((C5.PrintAmount-C5.GST)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='MO',0,IIF(Right(C1.BOM,2)='BP',0,IIF(Right(C1.BOM,2)='BM',0,0))))))," & _
                "IIF(Right(C1.BOM,2)='MF',((C2.RBillAmount-C2.vat)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='ME',((C3.RBillAmount-C3.vat)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='CF',((C5.PaperAmount-C5.GST)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='MO',0,IIF(Right(C1.BOM,2)='BP',0,IIF(Right(C1.BOM,2)='BM',0,0))))))," & _
                "C2.BillAmount,C2.PBillAmount,IIF(Right(C1.BOM,2)='MF',((C2.RBillAmount-C2.vat)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='ME',((C3.RBillAmount-C3.vat)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='CF',((C5.PaperAmount-C5.GST)/P2.EstQty01)*C1.Quantity,IIF(Right(C1.BOM,2)='MO',0,IIF(Right(C1.BOM,2)='BP',0,IIF(Right(C1.BOM,2)='BM',0,0)))))),C2.RBillAmount,P2.EstQty01,C2.Vat,C2.Pvat,C2.Rvat,P2.Name As OrderNo," & _
                "(C2.[TotalPlates1-¼]+C2.[TotalPlates1-½]+C2.[TotalPlates1-1]+C2.[RevisedPlates1])*1 As TotalPlates1,((C2.[TotalPlates1-¼]+C2.[TotalPlates1-½]+C2.[TotalPlates1-1]+C2.[RevisedPlates1])*1)*PlateRate1 As PlateAmount1,(C2.[TotalPlates2-¼]+C2.[TotalPlates2-½]+C2.[TotalPlates2-1]+C2.[RevisedPlates2])*2 As TotalPlates2,((C2.[TotalPlates2-¼]+C2.[TotalPlates2-½]+C2.[TotalPlates2-1]+C2.[RevisedPlates2])*2)*PlateRate2 As PlateAmount2,(C2.[TotalPlates4-¼]+C2.[TotalPlates4-½]+C2.[TotalPlates4-1]+C2.[RevisedPlates4])*4 As TotalPlates4,((C2.[TotalPlates4-¼]+C2.[TotalPlates4-½]+C2.[TotalPlates4-1]+C2.[RevisedPlates4])*4)*PlateRate4 As PlateAmount4,P.ChallanNo,P.ChallanDate,P.Transport,P.GrNo,P.GrDate,P.VehicleNo,P.Station " & _
                "FROM((((((((((((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) LEFT JOIN BookMaster I ON C1.Item=I.Code) LEFT JOIN GeneralMaster N ON C1.Narration=N.Code) LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code) LEFT JOIN (BookPOParent P2 LEFT JOIN BookPOChild05 C2 ON P2.Code=C2.Code) ON C1.Ref=P2.Code) LEFT JOIN BookPOChild06 C3 ON P2.Code=C3.Code) LEFT JOIN BookPOChild08 C4 ON P2.Code=C4.Code) LEFT JOIN BookPOChild09 C5 ON P2.Code=C5.Code) LEFT JOIN GeneralMaster B ON C4.BindingType=B.Code) LEFT JOIN PaperMaster PM1 ON C2.Paper1=PM1.Code) LEFT JOIN PaperMaster PM2 ON C2.Paper1=PM2.Code) LEFT JOIN PaperMaster PM4 ON C2.Paper1=PM4.Code) LEFT JOIN PaperMaster PM5 ON C3.Paper=PM5.Code) LEFT JOIN PaperMaster PM6 ON C5.Paper=PM6.Code) LEFT JOIN GeneralMaster S ON I.FinishSize=S.Code) LEFT JOIN AccountMaster A ON P.Party=A.Code) LEFT JOIN AccountMaster C ON P.Consignee=C.Code " & _
                "WHERE P.Code='" + VchCode + "' ORDER BY I.PrintName,N.Name", cnJobworkBill, adOpenKeyset, adLockReadOnly
ElseIf Right(VchType, 2) = "SX" Or Right(VchType, 2) = "PX" Or Right(VchType, 2) = "QX" Or Right(VchType, 2) = "ZX" Then
        rstJobworkBVChild.Open SQL & "'' As cmbTitle," & FS02 & " As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM (((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN (BookPOParent P2 INNER JOIN BookPOChild05 C2 ON P2.Code=C2.Code) ON C1.Ref=P2.Code) INNER JOIN GeneralMaster S ON I.FinishSize=S.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code WHERE P.Code='" + VchCode + "' AND BOM='TP' UNION " & _
                                                             SQL & "'' As cmbTitle,'Finish Size: '+LTRIM(S.PrintName)+', '+LTRIM(C2.FrontPrintingType)+'+'+LTRIM(C2.BackPrintingType)+'Col' As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM (((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN (BookPOParent P2 INNER JOIN BookPOChild06 C2 ON P2.Code=C2.Code) ON C1.Ref=P2.Code) INNER JOIN GeneralMaster S ON I.FinishSize=S.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code WHERE P.Code='" + VchCode + "' AND BOM='CP' UNION " & _
                                                             SQL & "I2.PrintName+', Finish Size: '+LTRIM(S.PrintName)+', '+LTRIM(C3.FrontPrintingColor)+'+'+LTRIM(C3.BackPrintingColor)+'Col, '+LTRIM(C3.[Ups/Plate])+'Ups, Qty: '+LTRIM(C3.ActualQuantity) As cmbItem,'' As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM ((((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN (BookPOParent P2 INNER JOIN (BookPOChild09 C2 INNER JOIN BookPOChild0901 C3 ON C2.Code=C3.Code) ON P2.Code=C2.Code) ON C1.Ref=P2.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code) INNER JOIN BookMaster I2 ON C3.Book=I2.Code) INNER JOIN GeneralMaster S ON I2.FinishSize=S.Code WHERE P.Code='" + VchCode + "' AND BOM='JP' UNION " & _
                                                             SQL & "'' As cmbTitle,'Finish Size: '+LTRIM(S.PrintName) As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05 FROM ((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN GeneralMaster S ON I.FinishSize=S.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code WHERE P.Code='" + VchCode + "' AND BOM='MO' UNION " & _
                                                             SQL & "'' As cmbTitle,'Finish Size: '+LTRIM(S.PrintName)+', '+LTRIM(IIF(ISNULL(C3.Pages1)=True,I.Pages,C3.Pages1+C3.Pages2+C3.Pages4))+' pages/'+LTRIM(C2.BindingForms+C2.ExtraForms)+'f, '+LTRIM(B.PrintName) As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,P.ChallanNo,P.ChallanDate,P.Transport,P.GrNo,P.GrDate,P.VehicleNo,P.Station  FROM (((((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN (BookPOParent P2 INNER JOIN BookPOChild08 C2 ON P2.Code=C2.Code) ON C1.Ref=P2.Code) INNER JOIN GeneralMaster S ON I.FinishSize=S.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code) INNER JOIN GeneralMaster B ON C2.BindingType=B.Code) LEFT JOIN BookPOChild05  C3 ON P2.Code=C3.Code " & _
                                                                          "WHERE P.Code='" + VchCode + "' AND BOM='BD' ORDER BY Item,SrNo,cmbTitle", cnJobworkBill, adOpenKeyset, adLockOptimistic
    ElseIf Right(VchType, 2) = "SU" Or Right(VchType, 2) = "PU" Or Right(VchType, 2) = "QU" Or Right(VchType, 2) = "ZU" Then  '"SG"
        rstJobworkBVChild.Open SQL & "'' As cmbTitle," & FS01 & " As FinishSize,LongNarration01,LongNarration02,LongNarration03,LongNarration04,LongNarration05,P.ChallanNo,P.ChallanDate,P.Transport,P.GrNo,P.GrDate,P.VehicleNo,P.Station  FROM ((((((JobworkBVParent P INNER JOIN JobworkBVChild C1 ON P.Code=C1.Code) INNER JOIN BookMaster I ON C1.Item=I.Code) INNER JOIN GeneralMaster N ON C1.Narration=N.Code) INNER JOIN GeneralMaster H ON C1.HSNCode=H.Code) INNER JOIN GeneralMaster S ON I.FinishSize=S.Code) INNER JOIN AccountMaster A ON P.Party=A.Code) INNER JOIN AccountMaster C ON P.Consignee=C.Code WHERE P.Code='" + VchCode + "'  ORDER BY Item,SrNo", cnJobworkBill, adOpenKeyset, adLockOptimistic 'AND BOM='WS'
    End If
    If rstJobworkBVChild.RecordCount = 0 Then On Error GoTo 0: Screen.MousePointer = vbNormal: Exit Sub
    If MsgBox("Print Item Detail?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Quit !") = vbNo Then rptJobworkBill.Section25.Suppress = True
    rstJobworkBVChild.ActiveConnection = Nothing
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
    
    If Not CheckEmpty(Trim(rstJobworkBVChild.Fields("ChallanNo").Value), False) Then rptItemIssueReceiptVoucher.Text37.SetText "Challan No: " & Trim(rstJobworkBVChild.Fields("ChallanNo").Value) + " dt: " & Trim(rstJobworkBVChild.Fields("ChallanDate").Value)
    If Not CheckEmpty(Trim(rstJobworkBVChild.Fields("GRNo").Value), False) Then rptItemIssueReceiptVoucher.Text38.SetText "GR No: " & Trim(rstJobworkBVChild.Fields("GR No").Value) + " GR dt: " & Trim(rstJobworkBVChild.Fields("GR Date").Value)
    If Not CheckEmpty(Trim(rstJobworkBVChild.Fields("VehicleNo").Value), False) Then rptItemIssueReceiptVoucher.Text39.SetText "Vehicle No: " & Trim(rstJobworkBVChild.Fields("VehicleNo").Value)
    If Not CheckEmpty(Trim(rstJobworkBVChild.Fields("Station").Value), False) Then rptItemIssueReceiptVoucher.Text41.SetText "Station: " & Trim(rstJobworkBVChild.Fields("Station").Value)
    rptJobworkBill.Text8.SetText "GSTIN/UIN : " & Trim(rstCompanyMaster.Fields("GSTIN").Value)
    rptJobworkBill.Text10.SetText "(" & UCase(Trim(NumberToWords(rstJobworkBVChild.Fields("TotalAmount").Value, False))) & ")"
    rptJobworkBill.Text11.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptJobworkBill.Text26.SetText CheckNull(rstCompanyMaster.Fields("Declaration01").Value)
    rptJobworkBill.Text25.SetText CheckNull(rstCompanyMaster.Fields("Declaration02").Value)
    rptJobworkBill.Text22.SetText CheckNull(rstCompanyMaster.Fields("Declaration03").Value)
    rptJobworkBill.Text12.SetText CheckNull(rstCompanyMaster.Fields("Declaration04").Value)
    rptJobworkBill.Text9.SetText CheckNull(rstCompanyMaster.Fields("Declaration05").Value)
    rptJobworkBill.Text30.SetText CheckNull(rstCompanyMaster.Fields("Declaration06").Value)
    rptJobworkBill.Text31.SetText CheckNull(rstCompanyMaster.Fields("Declaration07").Value)
    rptJobworkBill.Text33.SetText "Bank Name             : " & CheckNull(rstCompanyMaster.Fields("BankName").Value)
    rptJobworkBill.Text34.SetText "A/c No.                    : " & CheckNull(rstCompanyMaster.Fields("AccountNo").Value)
    rptJobworkBill.Text36.SetText "Branch & IFS Code : " & CheckNull(rstCompanyMaster.Fields("IFSC").Value)
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
    Set rptJobworkBill = Nothing
    If rstJobworkBVList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster) 'For Print Utility
    Call CloseRecordset(rstJobworkBVChild)
    On Error GoTo 0
End Sub
