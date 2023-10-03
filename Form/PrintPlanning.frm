VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPrintPlanning 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Planning"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
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
   ScaleHeight     =   4875
   ScaleWidth      =   8715
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   4870
      Left            =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   8700
      _Version        =   65536
      _ExtentX        =   15346
      _ExtentY        =   8590
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
      Picture         =   "PrintPlanning.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   4630
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   8176
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
         TabPicture(0)   =   "PrintPlanning.frx":001C
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
         TabPicture(1)   =   "PrintPlanning.frx":0038
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
            TabIndex        =   13
            Top             =   4160
            Width           =   7760
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3645
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   450
            Width           =   8235
            _ExtentX        =   14526
            _ExtentY        =   6429
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "Voucher No."
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
               Caption         =   "Voucher Date"
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
               DataField       =   "Particulars"
               Caption         =   "Particulars"
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
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   5534.929
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   3865
            Left            =   -74880
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   480
            Width           =   8235
            _Version        =   65536
            _ExtentX        =   14526
            _ExtentY        =   6817
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
            Picture         =   "PrintPlanning.frx":0054
            Begin VB.TextBox MhRealInput4 
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
               Left            =   6870
               MaxLength       =   13
               TabIndex        =   8
               Text            =   "0.000"
               Top             =   1400
               Visible         =   0   'False
               Width           =   1030
            End
            Begin VB.TextBox MhRealInput3 
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
               Left            =   5820
               MaxLength       =   13
               TabIndex        =   7
               Text            =   "0.00"
               Top             =   1400
               Visible         =   0   'False
               Width           =   1060
            End
            Begin VB.TextBox MhRealInput2 
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
               Left            =   5220
               MaxLength       =   13
               TabIndex        =   6
               Text            =   "0.00"
               Top             =   1400
               Visible         =   0   'False
               Width           =   610
            End
            Begin VB.TextBox MhRealInput1 
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
               Left            =   4380
               MaxLength       =   13
               TabIndex        =   5
               Text            =   "0"
               Top             =   1400
               Visible         =   0   'False
               Width           =   850
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
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   3430
               Width           =   6690
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
               Left            =   435
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   4
               Top             =   1400
               Visible         =   0   'False
               Width           =   3960
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
               TabIndex        =   2
               Top             =   630
               Width           =   6570
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   16
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
               Caption         =   " Voucher No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PrintPlanning.frx":0070
               Picture         =   "PrintPlanning.frx":008C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   5475
               TabIndex        =   17
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
               Caption         =   " Voucher Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PrintPlanning.frx":00A8
               Picture         =   "PrintPlanning.frx":00C4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   18
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
               Picture         =   "PrintPlanning.frx":00E0
               Picture         =   "PrintPlanning.frx":00FC
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2090
               Left            =   120
               TabIndex        =   3
               Top             =   1150
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   3678
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   9164542
               HeadLines       =   1
               RowHeight       =   20
               TabAction       =   2
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
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "BookName"
                  Caption         =   "Book Name"
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
               BeginProperty Column01 
                  DataField       =   "Quantity"
                  Caption         =   " Quantity"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "Forms"
                  Caption         =   " Forms"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "PaperWastage%"
                  Caption         =   "Wastage (%)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "PaperConsumption"
                  Caption         =   "Consumption"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.000"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
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
                     ColumnWidth     =   3945.26
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     ColumnAllowSizing=   -1  'True
                     Locked          =   -1  'True
                     ColumnWidth     =   840.189
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   599.811
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1049.953
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1019.906
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   120
               TabIndex        =   20
               Top             =   3430
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
               Caption         =   " Size"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PrintPlanning.frx":0118
               Picture         =   "PrintPlanning.frx":0134
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   7040
               TabIndex        =   1
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "PrintPlanning.frx":0150
               Caption         =   "PrintPlanning.frx":0268
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PrintPlanning.frx":02D4
               Keys            =   "PrintPlanning.frx":02F2
               Spin            =   "PrintPlanning.frx":0350
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
            Begin VB.Line Line3 
               X1              =   0
               X2              =   8280
               Y1              =   3330
               Y2              =   3330
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   8280
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   8280
               Y1              =   1050
               Y2              =   1050
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   3900
            TabIndex        =   21
            Top             =   0
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
            Picture         =   "PrintPlanning.frx":0378
            Picture         =   "PrintPlanning.frx":0394
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
            Top             =   4160
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
      Width           =   8715
      _ExtentX        =   15372
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
Attribute VB_Name = "FrmPrintPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnPrintPlanning As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstPrintPVList As New ADODB.Recordset
Dim rstPrintPVParent As New ADODB.Recordset
Dim WithEvents rstPrintPVChild As ADODB.Recordset
Attribute rstPrintPVChild.VB_VarHelpID = -1
Dim rstBookList As New ADODB.Recordset
Dim rstCheckRef As New ADODB.Recordset
Dim BookCode As String
Dim AddTitleEntry As String
Dim TitleEntryCode As String
Dim TitleEntryName As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim SortOrder, OutputTo As String
Public PlanningType As String
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstPrintPVList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstPrintPVList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstPrintPVList.EOF Or rstPrintPVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    If PlanningType = "1" Then
        DataGrid2.Columns(0).Caption = "Item Name"
        Me.Caption = "Print Planning [Multi Form Format]"
    Else
        DataGrid2.Columns(0).Caption = "Item Name"
        Me.Caption = "Print Planning [Spread Format]"
    End If
    cnPrintPlanning.CursorLocation = adUseClient
    cnPrintPlanning.Open cnDatabase.ConnectionString
    rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnPrintPlanning, adOpenKeyset, adLockReadOnly
    rstPrintPVList.Open "Select PrintPVParent.Code, PrintPVParent.Name, Date, Particulars From PrintPVParent Where PlanningType = '" & PlanningType & "' AND FYCode='" & FYCode & "' Order By PrintPVParent.Name", cnPrintPlanning, adOpenKeyset, adLockOptimistic
    rstPrintPVParent.CursorLocation = adUseClient
    Set rstPrintPVChild = New ADODB.Recordset
    If rstPrintPVList.RecordCount > 0 Then rstPrintPVList.MoveLast
    Set DataGrid1.DataSource = rstPrintPVList
    BusySystemIndicator False
    SSTab1.Tab = 0
    If Not (rstPrintPVList.EOF Or rstPrintPVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstPrintPVList.ActiveConnection = Nothing
    LoadMasterList
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True
    MdiMainMenu.mnuPrintPlanningModule.Enabled = False
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
                If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "MhRealInput3" And Me.ActiveControl.Name <> "MhRealInput4" Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
            If Not Me.ActiveControl Is Nothing Then
                If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "MhRealInput3" And Me.ActiveControl.Name <> "MhRealInput4" Then KeyCode = 0
            Else    'if Form Unloaded in case of Add
                KeyCode = 0
            End If
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
        If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "MhRealInput3" And Me.ActiveControl.Name <> "MhRealInput4" Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
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
            SSTab1.Tab = 1
            SSTab1.SetFocus
        Else
            If Me.ActiveControl.Name <> "MhRealInput4" Then
                Sendkeys "{TAB}"
            End If
        End If
        If Me.ActiveControl.Name <> "MhRealInput4" Then
            KeyCode = 0
        End If
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstPrintPVList)
    Call CloseRecordset(rstPrintPVParent)
    Call CloseRecordset(rstPrintPVChild)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstCheckRef)
    Call CloseConnection(cnPrintPlanning)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuPrintPlanningModule.Enabled = True
End Sub
Private Sub Text1_Change()
    If rstPrintPVList.RecordCount = 0 Then Exit Sub
    rstPrintPVList.MoveFirst
    If Text1.Text <> "" Then
        rstPrintPVList.Find "[Name] Like '%" & FixQuote(Text1.Text) & "%'"
        If rstPrintPVList.EOF Then
            rstPrintPVList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstPrintPVList.Bookmark = dblBookMark
                End If
            Else
                PrevStr = ""
            End If
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            Sendkeys "{End}"
        Else
            PrevStr = Text1.Text
            dblBookMark = DataGrid1.Bookmark
        End If
    Else
        PrevStr = ""
    End If
    If Not (rstPrintPVList.EOF Or rstPrintPVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstPrintPVList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstPrintPVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstPrintPVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstPrintPVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstPrintPVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstPrintPVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstPrintPVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstPrintPVList
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
            If Not (rstPrintPVList.EOF Or rstPrintPVList.BOF) Then
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
        Text2.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    
    If Button.Index = 1 Then
        If rstPrintPVParent.State = adStateOpen Then
           rstPrintPVParent.Close
        End If
        rstPrintPVParent.Open "Select * From PrintPVParent Where Code = ''", cnPrintPlanning, adOpenKeyset, adLockOptimistic
        ClearFields ("P")
        ClearFields ("C")
        Call LoadBookList("")
        If rstPrintPVChild.State = adStateClosed Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        If AddRecord(rstPrintPVParent) Then
            Text2.Text = GenerateCode(cnPrintPlanning, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,Name))", "VAL(Name))") & "  FROM PrintPVParent WHERE PlanningType = '" & PlanningType & "' AND FYCode='" & FYCode & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            cnPrintPlanning.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstPrintPVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstPrintPVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If CheckRef Then
            DisplayError ("Failed to delete the record")
        ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnPrintPlanning.Execute "Delete From PrintPVParent Where Code = '" & rstPrintPVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstPrintPVList.Delete
                rstPrintPVList.MoveNext
                If rstPrintPVList.RecordCount > 0 And rstPrintPVList.EOF Then
                    rstPrintPVList.MoveLast
                End If
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
        MakeTextBoxInvisible (False)
        If blnRecordExist And AllowTransactionsModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Voucher")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        AddTitleEntry = vbNo
        If PlanningType = "1" Then
            If CheckEmpty(rstPrintPVParent.Fields("Code").Value, False) Then
                AddTitleEntry = MsgBox("Do you wish to add entry for Title also ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Title Entry !")
            End If
        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstPrintPVParent) Then
            UpdateFlag = 1
            If AddTitleEntry = vbYes Then
                If Not TitleEntry Then
                    UpdateFlag = 0
                End If
            End If
            If UpdateFlag Then
                UpdateFlag = 0
                If UpdateBookList("D") Then
                     UpdateFlag = 1
                     If rstPrintPVChild.RecordCount <> 0 Then
                          rstPrintPVChild.MoveFirst
                          Do While Not rstPrintPVChild.EOF
                              If Val(rstPrintPVChild.Fields("Quantity").Value) <> 0 Then
                                   If Not UpdateBookList("U") Then
                                        UpdateFlag = 0
                                        Exit Do
                                    End If
                              End If
                              rstPrintPVChild.MoveNext
                          Loop
                     End If
                End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnPrintPlanning.CommitTrans
            If rstPrintPVParent.State = adStateOpen Then rstPrintPVParent.Close
            rstPrintPVParent.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstPrintPVParent) Then
            cnPrintPlanning.RollbackTrans
            If rstPrintPVParent.State = adStateOpen Then rstPrintPVParent.Close
            rstPrintPVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstPrintPVList.ActiveConnection = cnPrintPlanning
        Do While Not RefreshRecord(rstPrintPVList)
        Loop
        Set DataGrid1.DataSource = rstPrintPVList
        rstPrintPVList.ActiveConnection = Nothing
        If rstPrintPVList.RecordCount > 0 Then rstPrintPVList.MoveLast
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstPrintPVList.RecordCount = 0 Then Exit Sub
        OutputTo = "P"
        PrintPrintPlanning
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstPrintPVList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        PrintPrintPlanning
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstPrintPVList.RecordCount > 0 Then rstPrintPVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstPrintPVList.RecordCount > 0 Then
            rstPrintPVList.MovePrevious
            If rstPrintPVList.BOF Then
                rstPrintPVList.MoveNext
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstPrintPVList.RecordCount > 0 Then
            rstPrintPVList.MoveNext
            If rstPrintPVList.EOF Then
                rstPrintPVList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstPrintPVList.RecordCount > 0 Then rstPrintPVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstPrintPVList.EOF Or rstPrintPVList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        Text1.SetFocus
    End If
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
    End If
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
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstPrintPVList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(9).Enabled = False
        Toolbar1.Buttons.Item(10).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstPrintPVParent.EOF Or rstPrintPVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnPrintPlanning, "PrintPVParent", "Code", "[Name]+PlanningType", Trim(Text2.Text) & PlanningType, rstPrintPVParent.Fields("Code").Value, False, FYCode) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not ValidateDate(Me.ActiveControl) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If rstPrintPVChild.RecordCount = 0 Then
        Sendkeys "^"
        Call AddRecord(rstPrintPVChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
    End If
End Sub
Private Sub ViewRecord()
    ClearFields ("P")
    ClearFields ("C")
    If rstPrintPVList.EOF Then
        If rstPrintPVChild.State = adStateOpen Then rstPrintPVChild.Close
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstPrintPVParent.State = adStateOpen Then
       rstPrintPVParent.Close
    End If
    rstPrintPVParent.Open "Select * From PrintPVParent Where Code = '" & FixQuote(rstPrintPVList.Fields("Code").Value) & "'", cnPrintPlanning, adOpenKeyset, adLockOptimistic
    If rstPrintPVParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields(ByVal strType As String)
    If strType = "P" Then
        Text2.Text = ""
        Text4.Text = ""
        MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    ElseIf strType = "C" Then
        Text5.Text = ""
        Text6.Text = ""
        MhRealInput1.Text = "0"
        MhRealInput2.Text = "0.00"
        If PlanningType = "1" Then MhRealInput3.Text = "2.00" Else MhRealInput3.Text = "4.00"
        MhRealInput4.Text = "0.000"
        BookCode = ""
    End If
End Sub
Private Sub LoadFields()
    If rstPrintPVParent.EOF Or rstPrintPVParent.BOF Then Exit Sub
    Text2.Text = rstPrintPVParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstPrintPVParent.Fields("Date").Value, "dd-MM-yyyy")
    Text4.Text = rstPrintPVParent.Fields("Remarks").Value
    Call LoadBookList(rstPrintPVParent.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstPrintPVParent.RecordCount = 0 Then Exit Sub
    If rstPrintPVChild.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    If rstPrintPVParent.State = adStateOpen Then
       rstPrintPVParent.Close
    End If
    rstPrintPVParent.CursorLocation = adUseServer
    rstPrintPVParent.Open "Select * From PrintPVParent Where Code = '" & FixQuote(rstPrintPVList.Fields("Code").Value) & "'", cnPrintPlanning, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstPrintPVParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    cnPrintPlanning.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstPrintPVParent.EOF Or rstPrintPVParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstPrintPVParent.Fields("Code").Value = GenerateCode(cnPrintPlanning, "Select Max(Code) From PrintPVParent", 6, "0")
        rstPrintPVParent.Fields("CreatedBy").Value = UserCode
        rstPrintPVParent.Fields("CreatedOn").Value = Now()
        rstPrintPVParent.Fields("Recordstatus").Value = "N"
    Else
        rstPrintPVParent.Fields("ModifiedBy").Value = UserCode
        rstPrintPVParent.Fields("ModifiedOn").Value = Now()
        rstPrintPVParent.Fields("Recordstatus").Value = "M"
    End If
    rstPrintPVParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstPrintPVParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstPrintPVParent.Fields("PlanningType").Value = PlanningType
    rstPrintPVParent.Fields("Particulars").Value = "Planned " & Format(rstPrintPVChild.RecordCount, 0) & IIf(PlanningType = "1", " Item(s)", " Spread Sheet(s)") & " For Printing"
    rstPrintPVParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstPrintPVParent.Fields("FYCode").Value = FYCode
    rstPrintPVParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstPrintPVList.MoveFirst
    rstPrintPVList.Find "[Code] = '" & rstPrintPVParent.Fields("Code").Value & "'"
    If rstPrintPVList.EOF Then
       rstPrintPVList.AddNew
       rstPrintPVList.Fields("Code").Value = rstPrintPVParent.Fields("Code").Value
    End If
    rstPrintPVList.Fields("Name").Value = Pad(rstPrintPVParent.Fields("Name").Value, Space(1), 10, "L")
    rstPrintPVList.Fields("Date").Value = rstPrintPVParent.Fields("Date").Value
    rstPrintPVList.Fields("Particulars").Value = Trim(rstPrintPVParent.Fields("Particulars").Value)
    rstPrintPVList.Update
    rstPrintPVList.Sort = "Name Asc"
    rstPrintPVList.Find "[Code] = '" & rstPrintPVParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Voucher No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckDuplicate(cnPrintPlanning, "PrintPVParent", "Code", "[Name]+PlanningType", Trim(Text2.Text) & PlanningType, rstPrintPVParent.Fields("Code").Value, False, FYCode) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    End If
End Function
Private Function CheckRef() As Boolean
    On Error GoTo ErrorHandler
    
    If rstCheckRef.State = adStateOpen Then
         rstCheckRef.Close
    End If
    rstCheckRef.Open "Select Ref From " & IIf(PlanningType = "1", "BookPOChild05", "BookPOChild06") & " Where Ref = '" & rstPrintPVList.Fields("Code").Value & "'", cnPrintPlanning, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then
        CheckRef = True
    End If
    Exit Function
ErrorHandler:
    CheckRef = True
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Private Sub LoadBookList(ByVal strVoucherCode As String)
    On Error GoTo ErrorHandler
    If rstPrintPVChild.State = adStateOpen Then rstPrintPVChild.Close
    rstPrintPVChild.Open "Select Book, M1.Name As BookName, M2.Name As SizeName, Quantity, T.Forms, [PaperWastage%], PaperConsumption From BookMaster M1, GeneralMaster M2, PrintPVChild T Where T.Book = M1.Code And M1.[FinishSize] = M2.Code And T.Code = '" & strVoucherCode & "'", cnPrintPlanning, adOpenKeyset, adLockOptimistic
    rstPrintPVChild.ActiveConnection = Nothing
    Set DataGrid2.DataSource = rstPrintPVChild
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load " & IIf(PlanningType = "1", "Item", "Spread Form") & " List")
End Sub
Private Sub DataGrid2_DblClick()
    Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
End Sub
Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyE Then
        If rstPrintPVChild.RecordCount = 0 Then
            KeyCode = 0
            Exit Sub
        End If
        If Val(CheckNull(rstPrintPVChild.Fields("Quantity").Value)) <> 0 Then
            BookCode = rstPrintPVChild.Fields("Book").Value
            Text5.Text = rstPrintPVChild.Fields("BookName").Value
            MhRealInput1.Text = Format(Val(rstPrintPVChild.Fields("Quantity").Value), "0")
            MhRealInput2.Text = Format(Val(rstPrintPVChild.Fields("Forms").Value), "0.00")
            MhRealInput3.Text = Format(Val(rstPrintPVChild.Fields("PaperWastage%").Value), "0.00")
            MhRealInput4.Text = Format(Val(rstPrintPVChild.Fields("PaperConsumption").Value), "0.000")
        End If
        With DataGrid2
            Text5.Visible = True
            Text5.Move .Left + .Columns(0).Left, .Top + .RowTop(.Row), .Columns(0).Width + 10, .RowHeight + 30
            MhRealInput1.Visible = True
            MhRealInput1.Move .Left + .Columns(1).Left, .Top + .RowTop(.Row), .Columns(1).Width + 10, .RowHeight + 30
            MhRealInput2.Visible = True
            MhRealInput2.Move .Left + .Columns(2).Left, .Top + .RowTop(.Row), .Columns(2).Width + 10, .RowHeight + 30
            MhRealInput3.Visible = True
            MhRealInput3.Move .Left + .Columns(3).Left, .Top + .RowTop(.Row), .Columns(3).Width + 10, .RowHeight + 30
            MhRealInput4.Visible = True
            MhRealInput4.Move .Left + .Columns(4).Left, .Top + .RowTop(.Row), .Columns(4).Width + 10, .RowHeight + 30
        End With
        DataGrid2.Enabled = False
        Text5.SetFocus
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Sendkeys "^"
        Call AddRecord(rstPrintPVChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If rstPrintPVChild.RecordCount = 0 Then Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            Set DataGrid2.DataSource = Nothing
            rstPrintPVChild.Delete
            rstPrintPVChild.MoveNext
            Set DataGrid2.DataSource = rstPrintPVChild
            DataGrid2.SetFocus
        End If
        If rstPrintPVChild.RecordCount = 0 Then
            Call ClearFields("C")
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
    ElseIf Shift = vbShiftMask And KeyCode = vbKeyTab Then
       Text4.SetFocus
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        Text2.SetFocus
        KeyCode = 0
    End If
End Sub
Private Sub DataGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim menusel As String
    
    If Button = vbRightButton Then
       menusel = DisplayPopupMenu(Me.hwnd)
        Select Case menusel
            Case 1
                Call DataGrid2_KeyDown(vbKeyA, vbCtrlMask)
            Case 2
                Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
            Case 3
                Call DataGrid2_KeyDown(vbKeyD, vbCtrlMask)
            Case Else
        End Select
    End If
End Sub
Private Sub rstPrintPVChild_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    If Not (rstPrintPVChild.EOF Or rstPrintPVChild.BOF) Then
        If Not IsNull(rstPrintPVChild.Fields("SizeName").Value) Then Text6.Text = rstPrintPVChild.Fields("SizeName").Value
    End If
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    ElseIf KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmBookMaster.SL = True
        FrmBookMaster.ItemType = "F"
        FrmBookMaster.MasterCode = BookCode
        Load FrmBookMaster
        If Err.Number <> 364 Then FrmBookMaster.Show vbModal
        On Error GoTo 0
        BookCode = slCode: Text5.Text = slName
        If Not CheckEmpty(BookCode, False) Then
            LoadMasterList
            rstBookList.MoveFirst
            rstBookList.Find "[Code] = '" & BookCode & "'"
            Text6.Text = rstBookList.Fields("SizeName").Value
            If Val(MhRealInput2.Text) = 0 Then MhRealInput2.Text = Format(Val(rstBookList.Fields("Forms").Value), "0.00")
            Sendkeys "{TAB}"
        End If
    ElseIf KeyCode = vbKeyDelete Then
        Text5.Text = "": BookCode = ""
    End If
End Sub
Private Sub MhRealInput1_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput1_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput1, KeyAscii, 0
End Sub
Private Sub MhRealInput1_Change()
    MhRealInput4.Text = Format(CalculateConsumption(PlanningType, Val(MhRealInput1.Text), Val(MhRealInput2.Text), Val(MhRealInput3.Text)), "0.000")
End Sub
Private Sub MhRealInput1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    If Not ValidateNumber(Me.ActiveControl, 0) Then
        Cancel = True
    ElseIf Val(MhRealInput1.Text) <= 0 Then
        Cancel = True
        MhRealInput1.SetFocus
        FocusSelect Me.ActiveControl
    End If
End Sub
Private Sub MhRealInput2_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput2_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput2, KeyAscii, 2
End Sub
Private Sub MhRealInput2_Change()
    MhRealInput4.Text = Format(CalculateConsumption(PlanningType, Val(MhRealInput1.Text), Val(MhRealInput2.Text), Val(MhRealInput3.Text)), "0.000")
End Sub
Private Sub MhRealInput2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    If Not ValidateNumber(Me.ActiveControl, 2) Then
        Cancel = True
    ElseIf Val(MhRealInput2.Text) <= 0 Then
        Cancel = True
        MhRealInput2.SetFocus
        FocusSelect Me.ActiveControl
    End If
End Sub
Private Sub MhRealInput3_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput3_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput3, KeyAscii, 2
End Sub
Private Sub MhRealInput3_Change()
    MhRealInput4.Text = Format(CalculateConsumption(PlanningType, Val(MhRealInput1.Text), Val(MhRealInput2.Text), Val(MhRealInput3.Text)), "0.000")
End Sub
Private Sub MhRealInput3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput3_Validate(Cancel As Boolean)
    If Not ValidateNumber(Me.ActiveControl, 2) Then
        Cancel = True
    ElseIf Val(MhRealInput3.Text) < 0 Or Val(MhRealInput3.Text) > 99.99 Then
        Cancel = True
        MhRealInput3.SetFocus
        FocusSelect Me.ActiveControl
    End If
End Sub
Private Sub MhRealInput4_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput4_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput4, KeyAscii, 3
End Sub
Private Sub MhRealInput4_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Not ValidateNumber(Me.ActiveControl, 3) Then Exit Sub
        If Val(MhRealInput4.Text) > 0 Then
            rstPrintPVChild.Fields("Book").Value = BookCode
            rstPrintPVChild.Fields("BookName").Value = Trim(Text5.Text)
            rstPrintPVChild.Fields("SizeName").Value = Trim(Text6.Text)
            rstPrintPVChild.Fields("Quantity").Value = Format(Val(MhRealInput1.Text), "0")
            rstPrintPVChild.Fields("Forms").Value = Format(Val(MhRealInput2.Text), "0.00")
            rstPrintPVChild.Fields("PaperWastage%").Value = Format(Val(MhRealInput3.Text), "0.00")
            rstPrintPVChild.Fields("PaperConsumption").Value = Format(Val(MhRealInput4.Text), "0.000")
            rstPrintPVChild.Update
            MakeTextBoxInvisible (False)
            If rstPrintPVChild.AbsolutePosition = rstPrintPVChild.RecordCount Then
                Call DataGrid2_KeyDown(vbKeyA, vbCtrlMask)
            End If
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
       MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)
    Cancel = True
End Sub
Private Sub MakeTextBoxInvisible(ByVal KeyEscPressed As Boolean)
    If KeyEscPressed Then
        If Not (rstPrintPVChild.EOF Or rstPrintPVChild.BOF) Then
            If Val(CheckNull(rstPrintPVChild.Fields("Quantity").Value)) = 0 Then
                rstPrintPVChild.Delete
                rstPrintPVChild.MoveNext
                If rstPrintPVChild.RecordCount > 0 Then rstPrintPVChild.MoveFirst
            End If
        End If
    End If
    Text5.Visible = False
    MhRealInput1.Visible = False
    MhRealInput2.Visible = False
    MhRealInput3.Visible = False
    MhRealInput4.Visible = False
    DataGrid2.Enabled = True
    DataGrid2.SetFocus
End Sub
Private Function CheckDuplicateBook() As Boolean
    Dim dblBookMark As Double
    
    If rstPrintPVChild.RecordCount = 0 Then Exit Function
    If Not (rstPrintPVChild.EOF Or rstPrintPVChild.BOF) Then
       dblBookMark = rstPrintPVChild.Bookmark
    End If
    rstPrintPVChild.MoveFirst
    Do While Not rstPrintPVChild.EOF
          If rstPrintPVChild.Fields("BookName").Value = Trim(Text5.Text) Then
             CheckDuplicateBook = True
             Exit Do
          End If
          rstPrintPVChild.MoveNext
    Loop
    If dblBookMark <> 0 Then
       rstPrintPVChild.Bookmark = dblBookMark
    Else
       rstPrintPVChild.MoveLast
    End If
End Function
Private Function UpdateBookList(ByVal strOption As String) As Boolean
    On Error GoTo ErrorHandler
    UpdateBookList = True
    If strOption = "D" Then
        cnPrintPlanning.Execute "Delete From PrintPVChild Where Code = '" & rstPrintPVParent.Fields("Code").Value & "'"
    Else
        cnPrintPlanning.Execute "INSERT INTO PrintPVChild Values ('" & rstPrintPVParent.Fields("Code").Value & "','" & rstPrintPVChild.Fields("Book").Value & "'," & Val(rstPrintPVChild.Fields("Quantity").Value) & "," & Val(rstPrintPVChild.Fields("Forms").Value) & "," & Val(rstPrintPVChild.Fields("PaperWastage%").Value) & "," & Val(rstPrintPVChild.Fields("PaperConsumption").Value) & ")"
        If AddTitleEntry = vbYes Then
            cnPrintPlanning.Execute "INSERT INTO PrintPVChild Values ('" & TitleEntryCode & "','" & rstPrintPVChild.Fields("Book").Value & "'," & Val(rstPrintPVChild.Fields("Quantity").Value) & "," & Val(rstPrintPVChild.Fields("Forms").Value) & ",4," & CalculateConsumption("2", Val(rstPrintPVChild.Fields("Quantity").Value), Val(rstPrintPVChild.Fields("Forms").Value), 4) & ")"
        End If
    End If
    Exit Function
ErrorHandler:
    UpdateBookList = False
End Function
Private Function TitleEntry() As Boolean
    On Error GoTo ErrorHandler
    TitleEntry = True
    TitleEntryCode = GenerateCode(cnPrintPlanning, "Select Max(Code) From PrintPVParent", 6, "0")
    TitleEntryName = GenerateCode(cnPrintPlanning, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,Name))", "VAL(Name))") & "  FROM  PrintPVParent Where PlanningType = '2'", 10, Space(1))
    If DatabaseType = "MS SQL" Then
        cnPrintPlanning.Execute "INSERT INTO PrintPVParent (Code, Name, [Date], PlanningType, Particulars, Remarks, CreatedBy,CreatedOn,PrintStatus,RecordStatus) Values('" & TitleEntryCode & "','" & TitleEntryName & "','" & GetDate(MhDateInput1.Text) & "','2','" & "Planned " & Format(rstPrintPVChild.RecordCount, 0) & " Title(s) For Printing','','" & UserCode & "',GETDATE(),'N','N')"
    Else
        cnPrintPlanning.Execute "INSERT INTO PrintPVParent (Code, Name, [Date], PlanningType, Particulars, Remarks, CreatedBy) Values('" & TitleEntryCode & "','" & TitleEntryName & "',#" & GetDate(MhDateInput1.Text) & "#,'2','" & "Planned " & Format(rstPrintPVChild.RecordCount, 0) & " Title(s) For Printing','','" & UserCode & "')"
    End If
    Exit Function
ErrorHandler:
    TitleEntry = False
End Function
Private Sub PrintPrintPlanning()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptPrintPlanning.Text1.SetText IIf(PlanningType = "1", "Item", "Spread Form") & " Print Planning"
    rptPrintPlanning.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPrintPlanning.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False)) Then
        rptPrintPlanning.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False) Then
        rptPrintPlanning.Text24.SetText "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptPrintPlanning.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value)
    Else
        rptPrintPlanning.Section5.Suppress = True
    End If
    If rstPrintPVChild.State = adStateOpen Then
        rstPrintPVChild.Close
    End If
    If PlanningType = "1" Then
        rstPrintPVChild.Open "SELECT LTRIM(P.Name) As VchNo,[Date] As VchDate,LTRIM(PrintName) As BookName,'' As BoardName,(SELECT LTRIM(PrintName) FROM GeneralMaster WHERE Code = M.[FinishSize]) As SizeName,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=M.BookPrinter) As BookPrinter,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=M.TitlePrinter) As TitlePrinter,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=M.Laminator) As Laminator,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=M.BinderFresh) As Binder,Quantity,C.Forms,PaperConsumption,P.Remarks FROM (PrintPVParent P INNER JOIN PrintPVChild C ON P.Code=C.Code) INNER JOIN BookMaster M ON C.Book=M.Code WHERE  P.Code = '" & rstPrintPVList.Fields("Code").Value & "' ORDER BY M.PrintName", cnPrintPlanning, adOpenKeyset, adLockOptimistic
    Else
        rstPrintPVChild.Open "SELECT LTRIM(P.Name) As VchNo,[Date] As VchDate,LTRIM(PrintName) As BookName,'' As BoardName,(SELECT LTRIM(PrintName) FROM GeneralMaster WHERE Code = M.[FinishSize]) As SizeName,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=M.TitlePrinter) As BookPrinter,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=M.BookPrinter) As TitlePrinter,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=M.Laminator) As Laminator,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=M.BinderFresh) As Binder,Quantity,C.Forms,PaperConsumption,P.Remarks FROM (PrintPVParent P INNER JOIN PrintPVChild C ON P.Code=C.Code) INNER JOIN BookMaster M ON C.Book=M.Code WHERE  P.Code = '" & rstPrintPVList.Fields("Code").Value & "' ORDER BY M.PrintName", cnPrintPlanning, adOpenKeyset, adLockOptimistic
        rptPrintPlanning.Text11.SetText "SF Party": rptPrintPlanning.Text15.SetText "MF PARTY"
    End If
    rptPrintPlanning.Database.SetDataSource rstPrintPVChild, 3, 1
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptPrintPlanning
        FrmReportViewer.Show vbModal
    Else
        rptPrintPlanning.PaperSource = crPRBinAuto
        rptPrintPlanning.PrintOut
    End If
    Set rptPrintPlanning = Nothing
    On Error GoTo 0
End Sub
Private Sub LoadMasterList()
    If rstBookList.State = adStateOpen Then rstBookList.Close
    rstBookList.Open "Select M1.Name As Col0, M2.Name As SizeName, '' Forms, M1.Code From BookMaster M1 INNER JOIN GeneralMaster M2 ON M1.[FinishSize] = M2.Code ORDER BY M1.Name", cnPrintPlanning, adOpenKeyset, adLockReadOnly
    rstBookList.ActiveConnection = Nothing
End Sub
