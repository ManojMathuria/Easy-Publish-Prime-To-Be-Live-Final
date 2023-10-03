VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRMConsumed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raw Materials Consumption Voucher"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18555
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
   ScaleHeight     =   11113.21
   ScaleMode       =   0  'User
   ScaleWidth      =   18555
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   8955
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   18540
      _Version        =   65536
      _ExtentX        =   32702
      _ExtentY        =   15796
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
      Picture         =   "RMConsumed.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8715
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   18300
         _ExtentX        =   32279
         _ExtentY        =   15372
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
         TabPicture(0)   =   "RMConsumed.frx":001C
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
         TabPicture(1)   =   "RMConsumed.frx":0038
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
            Left            =   1080
            TabIndex        =   12
            Top             =   8250
            Width           =   12555
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7710
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   450
            Width           =   18075
            _ExtentX        =   31882
            _ExtentY        =   13600
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
               Caption         =   "   Voucher No."
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
               Caption         =   "  Voucher Date"
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
               DataField       =   "AccountName"
               Caption         =   "   Godown Name"
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
                  ColumnWidth     =   1395.213
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1379.906
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   14699.91
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8115
            Left            =   -74880
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   480
            Width           =   18075
            _Version        =   65536
            _ExtentX        =   31882
            _ExtentY        =   14314
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
            Picture         =   "RMConsumed.frx":0054
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
               Left            =   1680
               MaxLength       =   60
               TabIndex        =   4
               Top             =   1260
               Width           =   13170
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   6195
               Left            =   120
               TabIndex        =   6
               Top             =   1800
               Width           =   17850
               _Version        =   524288
               _ExtentX        =   31485
               _ExtentY        =   10927
               _StockProps     =   64
               ButtonDrawMode  =   1
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
               MaxCols         =   22
               SpreadDesigner  =   "RMConsumed.frx":0070
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   10
               TabIndex        =   0
               Top             =   120
               Width           =   1650
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
               TabIndex        =   3
               Top             =   950
               Width           =   16290
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
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   2
               Top             =   630
               Width           =   16290
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   15
               Top             =   120
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
               Caption         =   " Vch. No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "RMConsumed.frx":0E1E
               Picture         =   "RMConsumed.frx":0E3A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   14835
               TabIndex        =   16
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
               Caption         =   " Vch. Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "RMConsumed.frx":0E56
               Picture         =   "RMConsumed.frx":0E72
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   17
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
               Caption         =   " Godown Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "RMConsumed.frx":0E8E
               Picture         =   "RMConsumed.frx":0EAA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   18
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
               Picture         =   "RMConsumed.frx":0EC6
               Picture         =   "RMConsumed.frx":0EE2
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   16395
               TabIndex        =   1
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "RMConsumed.frx":0EFE
               Caption         =   "RMConsumed.frx":1016
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "RMConsumed.frx":1082
               Keys            =   "RMConsumed.frx":10A0
               Spin            =   "RMConsumed.frx":10FE
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
            Begin VB.TextBox Text9 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3840
               TabIndex        =   19
               Top             =   2520
               Width           =   2535
            End
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   2955
               Left            =   120
               TabIndex        =   7
               Top             =   4920
               Width           =   17850
               _Version        =   524288
               _ExtentX        =   31485
               _ExtentY        =   5212
               _StockProps     =   64
               EditEnterAction =   5
               EditModePermanent=   -1  'True
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
               MaxCols         =   5
               MaxRows         =   100
               OperationMode   =   2
               SpreadDesigner  =   "RMConsumed.frx":1126
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   20
               Top             =   1260
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
               Caption         =   " Approved By"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "RMConsumed.frx":16ED
               Picture         =   "RMConsumed.frx":1709
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   14835
               TabIndex        =   21
               Top             =   1260
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
               Caption         =   " Approval Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "RMConsumed.frx":1725
               Picture         =   "RMConsumed.frx":1741
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput2 
               Height          =   330
               Left            =   16395
               TabIndex        =   5
               Top             =   1260
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "RMConsumed.frx":175D
               Caption         =   "RMConsumed.frx":1875
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "RMConsumed.frx":18E1
               Keys            =   "RMConsumed.frx":18FF
               Spin            =   "RMConsumed.frx":195D
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
               X2              =   18070
               Y1              =   4830
               Y2              =   4830
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   18070
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   18070
               Y1              =   1680
               Y2              =   1680
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   13620
            TabIndex        =   22
            Top             =   8250
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
            Picture         =   "RMConsumed.frx":1985
            Picture         =   "RMConsumed.frx":19A1
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
            TabIndex        =   13
            Top             =   8250
            Width           =   975
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   18555
      _ExtentX        =   32729
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
Attribute VB_Name = "FrmRMConsumed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VchCode As String
Dim cnMaterialStockAdjustment As New ADODB.Connection
Dim rstMaterialSVList As New ADODB.Recordset, rstMaterialSVParent As New ADODB.Recordset, rstMaterialSVChild As ADODB.Recordset
Dim rstCompanyMaster As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstUserList As New ADODB.Recordset, rstOutsourceItemList As New ADODB.Recordset, rstPaperList As New ADODB.Recordset, rstFreshBookList As New ADODB.Recordset, rstTitleList As New ADODB.Recordset, rstRepairBookList As New ADODB.Recordset, rstRMStock As New ADODB.Recordset
Dim AccountCode As String, ApproverCode As String
', BOM As String, Paper As String, FG As String, UFG As String, Title As String
Dim EditMode As Boolean, SortOrder As String, PrevStr As String, dblBookMark As Double, blnRecordExist As Boolean, OutputTo As String
Dim rstElementList As New ADODB.Recordset, Element As String
Dim BOMCode As Variant, PaperCode As Variant, FGCode As Variant, UFGCode As Variant, TitleCode As Variant, MachineCode As Variant, ElementCode As String, FinishedItemCode As Variant
Dim GSM, ReelSize, Cut_off As Variant
Dim sDate As Date, eDate As Date
Dim PlanWeight, PlanSheet, ActualWeight, Category As Variant, moveFlag As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
    BusySystemIndicator True
    cnMaterialStockAdjustment.CursorLocation = adUseClient
    cnMaterialStockAdjustment.Open cnDatabase.ConnectionString
    
    rstCompanyMaster.Open "SELECT PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website,FinancialYearFrom FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnMaterialStockAdjustment, adOpenKeyset, adLockReadOnly
    sDate = Format(rstCompanyMaster.Fields("FinancialYearFrom").Value, "dd-m-yyyy")
    sDate = Format(sDate, "dd-m-yyyy")
    rstOutsourceItemList.Open "Select M.Name,'1'+M.Code As NCode,U.Name As UOMName From OutsourceItemMaster M INNER JOIN GeneralMaster U ON M.UOM=U.Code Order By M.Name", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstPaperList.Open "SELECT LTRIM(M.Name)+' (UOM : '+LTRIM(C.Name)+')' As Name,'2'+M.Code As NCode,LTRIM(C.Name) AS UOMName FROM PaperMaster M INNER JOIN GeneralMaster C ON M.UOM=C.Code ORDER BY M.Name", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstFreshBookList.Open "Select Name,[Group],'3'+Code As NCode,'Piece' AS UOMName From BookMaster Where Type='F' Order By Name", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstRepairBookList.Open "Select Name,'4'+Code As NCode,'Piece' AS UOMName From BookMaster Where Type='R' Order By Name", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstTitleList.Open "Select Name,[Group],'5'+Code As NCode,'Piece' AS UOMName From BookMaster Where Type='F' Order By Name", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstElementList.Open "SELECT Name,'6'+Code FROM ElementMaster ORDER BY Name", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    
    rstMaterialSVList.Open "SELECT T.Code,T.Name,T.Date,M.Name As AccountName FROM MaterialSVParent T INNER JOIN AccountMaster M ON T.Account = M.Code WHERE Type='R' AND FYCode='" & FYCode & "' ORDER BY T.Name", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstMaterialSVParent.CursorLocation = adUseClient
    Set rstMaterialSVChild = New ADODB.Recordset
    rstMaterialSVList.Filter = adFilterNone
    If rstMaterialSVList.RecordCount > 0 Then rstMaterialSVList.MoveLast
    Set DataGrid1.DataSource = rstMaterialSVList
    BusySystemIndicator False
    SSTab1.Tab = 0
    If FrmStockLedger.dSortBy = True Then
    SortOrder = "Code"
    Else
    SortOrder = "Name"
    End If
    If Not (rstMaterialSVList.EOF Or rstMaterialSVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstMaterialSVList.ActiveConnection = Nothing
    rstAccountList.ActiveConnection = Nothing
    rstUserList.ActiveConnection = Nothing
    rstOutsourceItemList.ActiveConnection = Nothing
    rstPaperList.ActiveConnection = Nothing
    rstRMStock.ActiveConnection = Nothing
    rstFreshBookList.ActiveConnection = Nothing
    rstRepairBookList.ActiveConnection = Nothing
    rstTitleList.ActiveConnection = Nothing
    rstElementList.ActiveConnection = Nothing
'    Call RefreshDropDownList("A")
'    fpSpread1.Col = 4
'    fpSpread1.ColHidden = True
'    fpSpread2.Col = 4
'    fpSpread2.ColHidden = True

Dim i, C As Long, cVal As Variant
With fpSpread1
    .ColHeaderRows = 2:
        For C = 1 To .MaxCols
        .GetText C, 0, cVal
        .Col = C: .Row = SpreadHeader + 1: .Text = cVal:
        Next
        .Col = 1: .Row = SpreadHeader: .Text = " Production Planning": .FontSize = 12:
        .Col = 11: .Row = SpreadHeader: .Text = " Production & Job Performance": .FontSize = 12:
        .AddCellSpan 1, SpreadHeader, 10, 1
        .AddCellSpan 11, SpreadHeader, 17, 1

        .Col = 16: .Row = SpreadHeader + 1: .Text = " Sheets Weight"
        i = 0
        i = i + 1: .ColWidth(i) = 8.25
        i = i + 1: .ColWidth(i) = 34.125
        i = i + 1: .ColWidth(i) = 37
        i = i + 1: .ColWidth(i) = 11.125
        i = i + 1: .ColWidth(i) = 7.125
        i = i + 1: .ColWidth(i) = 10.875
        i = i + 1: .ColWidth(i) = 10.625
        i = i + 1: .ColWidth(i) = 10.25
        i = i + 1: .ColWidth(i) = 10
        i = i + 1: .ColWidth(i) = 10.125
        i = i + 1: .ColWidth(i) = 30
        i = i + 1: .ColWidth(i) = 8
        i = i + 1: .ColWidth(i) = 8
        i = i + 1: .ColWidth(i) = 8
        i = i + 1: .ColWidth(i) = 10
        i = i + 1: .ColWidth(i) = 11
        i = i + 1: .ColWidth(i) = 10
End With
    If VchApprovalRights Then Text5.Enabled = True: MhDateInput2.Enabled = True Else Text5.Enabled = False: MhDateInput2.Enabled = False
    
    'If eDate = "00:00:00" Then eDate = Format(Now, "dd-mm-yyyy") Else eDate = GetDate(MhDateInput1.Text)
    LoadMasterList
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True
    MdiMainMenu.mnuStockJournalRawMaterial.Enabled = False
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
           If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then
              Sendkeys "{TAB}"
           End If
        End If
        If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then
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
    Call CloseRecordset(rstMaterialSVList)
    Call CloseRecordset(rstMaterialSVParent)
    Call CloseRecordset(rstMaterialSVChild)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstUserList)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstRMStock)
    Call CloseRecordset(rstFreshBookList)
    Call CloseRecordset(rstTitleList)
    Call CloseRecordset(rstRepairBookList)
    Call CloseRecordset(rstElementList)
    Call CloseConnection(cnMaterialStockAdjustment)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuStockJournalRawMaterial.Enabled = True
End Sub
Private Sub Text1_Change()
    On Error Resume Next
    With rstMaterialSVList
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
    If rstMaterialSVList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstMaterialSVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstMaterialSVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstMaterialSVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstMaterialSVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstMaterialSVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstMaterialSVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstMaterialSVList
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
            If Not (rstMaterialSVList.EOF Or rstMaterialSVList.BOF) Then
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
    Dim UpdateFlag As Integer, i As Integer
    Dim CellVal As Variant
    
    If Button.Index = 1 Then
        If rstMaterialSVParent.State = adStateOpen Then rstMaterialSVParent.Close
        rstMaterialSVParent.Open "Select * From MaterialSVParent Where Code = ''", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstMaterialSVParent) Then
            Text2.Text = GenerateCode(cnMaterialStockAdjustment, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,Name))", "VAL(Name))") & "  FROM MaterialSVParent WHERE Type='R' AND  FYCode='" & FYCode & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            cnMaterialStockAdjustment.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstMaterialSVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstMaterialSVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnMaterialStockAdjustment.Execute "Delete From MaterialSVParent Where Code = '" & rstMaterialSVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstMaterialSVList.Delete
                rstMaterialSVList.MoveNext
                If rstMaterialSVList.RecordCount > 0 And rstMaterialSVList.EOF Then
                    rstMaterialSVList.MoveLast
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
        If blnRecordExist And AllowTransactionsModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Voucher")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstMaterialSVParent) Then
            If UpdateMaterialList("D", i) Then
                UpdateFlag = 1
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.SetActiveCell 9, i
                    fpSpread1.GetText 9, i, CellVal
                    If Val(CellVal) <> 0 Then
                        If Not UpdateMaterialList("I1", i) Then
                            UpdateFlag = 0
                            Exit For
                        End If
                    End If
                Next
                If UpdateFlag = 1 Then
                    For i = 1 To fpSpread2.DataRowCnt
                        fpSpread2.SetActiveCell 3, i
                        fpSpread2.GetText 3, i, CellVal
                        If Val(CellVal) <> 0 Then
                            If Not UpdateMaterialList("I2", i) Then
                                UpdateFlag = 0
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnMaterialStockAdjustment.CommitTrans
            If rstMaterialSVParent.State = adStateOpen Then
                rstMaterialSVParent.Close
            End If
            rstMaterialSVParent.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstMaterialSVParent) Then
            cnMaterialStockAdjustment.RollbackTrans
            If rstMaterialSVParent.State = adStateOpen Then
                rstMaterialSVParent.Close
            End If
            rstMaterialSVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstMaterialSVList.ActiveConnection = cnMaterialStockAdjustment
        Do While Not RefreshRecord(rstMaterialSVList)
        Loop
        Set DataGrid1.DataSource = rstMaterialSVList
        rstMaterialSVList.ActiveConnection = Nothing
        If rstMaterialSVList.RecordCount > 0 Then rstMaterialSVList.MoveLast
        rstAccountList.ActiveConnection = cnMaterialStockAdjustment
        Do While Not RefreshRecord(rstAccountList)
        Loop
        rstAccountList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Source", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstMaterialSVList.RecordCount = 0 Then Exit Sub
        OutputTo = "P"
        PrintMaterialStockAdjustment
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstMaterialSVList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        PrintMaterialStockAdjustment
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstMaterialSVList.RecordCount > 0 Then rstMaterialSVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstMaterialSVList.RecordCount > 0 Then
            rstMaterialSVList.MovePrevious
            If rstMaterialSVList.BOF Then
                rstMaterialSVList.MoveNext
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstMaterialSVList.RecordCount > 0 Then
            rstMaterialSVList.MoveNext
            If rstMaterialSVList.EOF Then
                rstMaterialSVList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstMaterialSVList.RecordCount > 0 Then rstMaterialSVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstMaterialSVList.EOF Or rstMaterialSVList.BOF) Then
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
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstMaterialSVList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstMaterialSVList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstMaterialSVList.EOF Or rstMaterialSVList.BOF) Then
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
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstMaterialSVList.RecordCount = 0 Then
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
    If rstMaterialSVParent.EOF Or rstMaterialSVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnMaterialStockAdjustment, "MaterialSVParent", "Code", "[Name]+[Type]", Trim(Text2.Text), rstMaterialSVParent.Fields("Code").Value, False, FYCode) Then
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
        FrmAccountMaster.MasterCode = AccountCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        AccountCode = slCode: Text3.Text = slName
        If Not CheckEmpty(AccountCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If CheckEmpty(Text3.Text, False) Then Cancel = True
End Sub
Private Sub Text5_Change()
    If Text5.Text = " " Then
        Text5.Text = "?": Sendkeys "{TAB}"
    ElseIf CheckEmpty(Text5.Text, False) Then
        ApproverCode = "": MhDateInput2.Value = Null
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    Dim SearchString As String
    If CheckEmpty(Text5.Text, False) Then MhDateInput2.Value = Null: Exit Sub
    SearchString = FixQuote(Text5.Text)
    If rstUserList.RecordCount = 0 Then
        DisplayError ("No Record in User Master")
        Cancel = True
        Exit Sub
    Else
        rstUserList.MoveFirst
    End If
    rstUserList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstUserList.EOF Then
        SelectionType = "S"
        ApproverCode = ""
        Call LoadSelectionList(rstUserList, "List of Approver(s)...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text5, ApproverCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text5.Text, False) Then Text5.Text = "?"
        If RTrim(ApproverCode) <> "" Then Sendkeys "{TAB}"
        If MhDateInput2.ValueIsNull Then MhDateInput2.Value = Format(Date, "dd-MM-yyyy")
        Cancel = True
    Else
        ApproverCode = rstUserList.Fields("Code").Value
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstMaterialSVList.EOF Then
        If rstMaterialSVChild.State = adStateOpen Then rstMaterialSVChild.Close
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstMaterialSVParent.State = adStateOpen Then rstMaterialSVParent.Close
    rstMaterialSVParent.Open "SELECT T.*,A.Name As ApproverName FROM MaterialSVParent T LEFT JOIN  UserMaster A ON T.ApprovedBy=A.Code WHERE T.Code = '" & FixQuote(rstMaterialSVList.Fields("Code").Value) & "'", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    If rstMaterialSVParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    fpSpread2.ClearRange 1, 1, fpSpread2.MaxCols, fpSpread2.MaxRows, True
    Text5.Text = ""
    MhDateInput2.Value = Null
End Sub
Private Sub LoadFields()
    If rstMaterialSVParent.EOF Or rstMaterialSVParent.BOF Then Exit Sub
    Text2.Text = rstMaterialSVParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstMaterialSVParent.Fields("Date").Value, "dd-MM-yyyy")
    AccountCode = rstMaterialSVParent.Fields("Account").Value
    If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & AccountCode & "'"
    If Not rstAccountList.EOF Then Text3.Text = rstAccountList.Fields("Col0").Value
    If Not CheckEmpty(rstMaterialSVParent.Fields("ApprovedBy").Value, False) Then ApproverCode = rstMaterialSVParent.Fields("ApprovedBy").Value
    If Not CheckEmpty(rstMaterialSVParent.Fields("ApproverName").Value, False) Then Text5.Text = rstMaterialSVParent.Fields("ApproverName").Value
    Text4.Text = rstMaterialSVParent.Fields("Remarks").Value
    If Not CheckEmpty(rstMaterialSVParent.Fields("ApprovalDate").Value, False) Then MhDateInput2.Text = Format(rstMaterialSVParent.Fields("ApprovalDate").Value, "dd-MM-yyyy")
    Call LoadMaterialList(rstMaterialSVParent.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstMaterialSVParent.RecordCount = 0 Then Exit Sub
    If rstMaterialSVChild.State = adStateClosed Then SSTab1.Tab = 0: Exit Sub
    If rstMaterialSVParent.State = adStateOpen Then rstMaterialSVParent.Close
    rstMaterialSVParent.CursorLocation = adUseServer
    rstMaterialSVParent.Open "SELECT T.*,A.Name As ApproverName FROM MaterialSVParent T LEFT JOIN  UserMaster A ON T.ApprovedBy=A.Code WHERE T.Code = '" & FixQuote(rstMaterialSVList.Fields("Code").Value) & "'", cnMaterialStockAdjustment, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstMaterialSVParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    cnMaterialStockAdjustment.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstMaterialSVParent.EOF Or rstMaterialSVParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstMaterialSVParent.Fields("Code").Value = GenerateCode(cnMaterialStockAdjustment, "Select Max(Code) From MaterialSVParent", 6, "0")
        rstMaterialSVParent.Fields("CreatedBy").Value = UserCode
        rstMaterialSVParent.Fields("CreatedOn").Value = Now()
        rstMaterialSVParent.Fields("Recordstatus").Value = "N"
    Else
        rstMaterialSVParent.Fields("ModifiedBy").Value = UserCode
        rstMaterialSVParent.Fields("ModifiedOn").Value = Now()
        rstMaterialSVParent.Fields("Recordstatus").Value = "M"
    End If
    rstMaterialSVParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstMaterialSVParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstMaterialSVParent.Fields("Account").Value = AccountCode
    rstMaterialSVParent.Fields("ApprovedBy").Value = ApproverCode
    If Not MhDateInput2.ValueIsNull Then rstMaterialSVParent.Fields("ApprovalDate").Value = GetDate(MhDateInput2.Text) Else rstMaterialSVParent.Fields("ApprovalDate").Value = Null
    rstMaterialSVParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstMaterialSVParent.Fields("Type").Value = "R"
    rstMaterialSVParent.Fields("FYCode").Value = FYCode
    rstMaterialSVParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstMaterialSVList.MoveFirst
    rstMaterialSVList.Find "[Code] = '" & rstMaterialSVParent.Fields("Code").Value & "'"
    If rstMaterialSVList.EOF Then
       rstMaterialSVList.AddNew
       rstMaterialSVList.Fields("Code").Value = rstMaterialSVParent.Fields("Code").Value
    End If
    rstMaterialSVList.Fields("Name").Value = Pad(rstMaterialSVParent.Fields("Name").Value, Space(1), 10, "L")
    rstMaterialSVList.Fields("Date").Value = rstMaterialSVParent.Fields("Date").Value
    rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & rstMaterialSVParent.Fields("Account").Value & "'"
    rstMaterialSVList.Fields("AccountName").Value = Trim(rstAccountList.Fields("Col0").Value)
    rstMaterialSVList.Update
    rstMaterialSVList.Sort = SortOrder & " Asc"
    rstMaterialSVList.Find "[Code] = '" & rstMaterialSVParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Order No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
       Text3.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstAccountList, AccountCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(cnMaterialStockAdjustment, "MaterialSVParent", "Code", "[Name]+[Type]", Trim(Text2.Text), rstMaterialSVParent.Fields("Code").Value, False, FYCode) Then
        Text2.SetFocus
        CheckMandatoryFields = True
'    ElseIf CheckItem("1") Then
'       fpSpread1.SetFocus
'        CheckMandatoryFields = True
'    ElseIf CheckItem("2") Then
'       fpSpread2.SetFocus
'        CheckMandatoryFields = True
    ElseIf CheckItem("20") Then
       fpSpread2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckItem("18") Then
       fpSpread2.SetFocus
        CheckMandatoryFields = True
     ElseIf CheckItem("19") Then
       fpSpread2.SetFocus
         CheckMandatoryFields = True
     ElseIf CheckItem("20") Then
       fpSpread2.SetFocus
         CheckMandatoryFields = True
     ElseIf CheckItem("22") Then
        fpSpread2.SetFocus
         CheckMandatoryFields = True
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
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Source" Then rstMaterialSVList.Filter = "[AccountName] Like '%" & SrchText & "%'"
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CurVal As Variant, i As Long
    
    With fpSpread1
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            Dim Production As Variant
            .GetText 15, .ActiveRow, Production
            If Production > 0 Then MsgBox "You Can't Delete This Record Due to Production Done ?", vbExclamation: fpSpread1.SetFocus: Exit Sub
            If MsgBox("Are you sure to delete the record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        
        ElseIf KeyCode = vbKeyReturn Then
                moveFlag = True
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 11 Then
            .GetText 15, .ActiveRow, Production
            If Production > 0 Then MsgBox "You Can't Edit This Record Due to Production Done ?", vbExclamation: fpSpread1.SetFocus: Exit Sub
                .GetText 20, .ActiveRow, MachineCode
                On Error Resume Next
                FrmMachineMaster.SL = True
                FrmMachineMaster.MasterCode = Right(MachineCode, 6)
                Load FrmMachineMaster
                If Err.Number <> 364 Then FrmMachineMaster.Show vbModal
                On Error GoTo 0
                If slCode = "" Then Exit Sub
                MachineCode = slCode
                .SetText 11, .ActiveRow, slName: .SetText 20, .ActiveRow, MachineCode
                .SetActiveCell 14, .ActiveRow
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 3 Then
            .GetText 15, .ActiveRow, Production
            If Production > 0 Then MsgBox "You Can't Edit This Record Due to Production Done ?", vbExclamation: fpSpread1.SetFocus: Exit Sub
                .GetText 19, .ActiveRow, PaperCode
                On Error Resume Next
                FrmPaperMaster.SL = True
                FrmPaperMaster.slAccount = AccountCode
                FrmPaperMaster.slvtCode = rstMaterialSVParent.Fields("Code").Value
                FrmPaperMaster.slDate = rstMaterialSVParent.Fields("Date").Value
                FrmPaperMaster.MasterCode = Right(PaperCode, 6)
                Load FrmPaperMaster
                If Err.Number <> 364 Then FrmPaperMaster.Show vbModal
                On Error GoTo 0
                If slCode = "" Then Exit Sub
                PaperCode = "2" + slCode
                .SetText 3, .ActiveRow, slName: .SetText 19, .ActiveRow, PaperCode
                LoadRMStock (True)
                On Error Resume Next
                .Col = 21: .ColHidden = False
                .Col = 22: .ColHidden = False
                If rstRMStock.RecordCount > 0 Then
                .SetText 4, .ActiveRow, Val(rstRMStock.Fields("RMSTOCK").Value)
                .SetText 5, .ActiveRow, rstRMStock.Fields("Unit").Value
                .SetText 6, .ActiveRow, Val(rstRMStock.Fields("RMSTOCKKG").Value)
                .SetText 21, .ActiveRow, Val(rstRMStock.Fields("RMUOMVALUE").Value)
                .SetText 22, .ActiveRow, rstRMStock.Fields("UOMCODE").Value
                End If
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 2 Then
                .GetText 1, .ActiveRow, CurVal
                If CheckEmpty(CurVal, False) Then Exit Sub
                If CurVal = "BOM Item" Then
                    .GetText 18, .ActiveRow, BOMCode
                    On Error Resume Next
                    FrmOutsourceItemMaster.SL = True
                    FrmOutsourceItemMaster.MasterCode = Right(BOMCode, 6)
                    Load FrmOutsourceItemMaster
                    If Err.Number <> 364 Then FrmOutsourceItemMaster.Show vbModal
                    On Error GoTo 0
                    If slCode = "" Then Exit Sub
                    .SetText 2, .ActiveRow, slName: .SetText 18, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                ElseIf CurVal = "Paper" Then
                    .GetText 18, .ActiveRow, PaperCode
                    On Error Resume Next
                    FrmPaperMaster.SL = True
                    FrmPaperMaster.MasterCode = Right(PaperCode, 6)
                    Load FrmPaperMaster
                    If Err.Number <> 364 Then FrmPaperMaster.Show vbModal
                    On Error GoTo 0
                    If slCode = "" Then Exit Sub
                    .SetText 2, .ActiveRow, slName: .SetText 18, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                ElseIf CurVal = "FG Item" Then
                    .GetText 18, .ActiveRow, FGCode
                    On Error Resume Next
                    FrmBookMaster.SL = True
                    FrmBookMaster.ItemType = "F"
                    FrmBookMaster.MasterCode = Right(FGCode, 6)
                    Load FrmBookMaster
                    If Err.Number <> 364 Then FrmBookMaster.Show vbModal
                    On Error GoTo 0
                    If slCode = "" Then Exit Sub
                    .SetText 2, .ActiveRow, slName: .SetText 18, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                ElseIf CurVal = "UFG Item" Then
                  If .ActiveCol = 2 Then
                    .GetText 15, .ActiveRow, Production
                    If Production > 0 Then MsgBox "You Can't Edit This Record Due to Production Done ?", vbExclamation: fpSpread1.SetFocus: Exit Sub
                    .GetText 18, .ActiveRow, UFGCode
                    On Error Resume Next
                    FrmBookMaster.SL = True
                    FrmBookMaster.ItemType = "R"
                    FrmBookMaster.MasterCode = Right(UFGCode, 6)
                    Load FrmBookMaster
                    If Err.Number <> 364 Then FrmBookMaster.Show vbModal
                    On Error GoTo 0
                    If slCode = "" Then Exit Sub
                    UFGCode = "4" + slCode
                    .SetText 2, .ActiveRow, slName: .SetText 18, .ActiveRow, UFGCode
                    LoadMasterList
                    On Error Resume Next
                    If rstMaterialSVChild.RecordCount <> 0 And slCode = rstMaterialSVChild.Fields("UFGCODE1").Value Then
                        .SetText 3, .ActiveRow, rstMaterialSVChild.Fields("RM").Value: .SetText 19, .ActiveRow, "2" + rstMaterialSVChild.Fields("RMCODE").Value
                        .SetText 4, .ActiveRow, Val(rstMaterialSVChild.Fields("RMSTOCK").Value)
                        .SetText 5, .ActiveRow, rstMaterialSVChild.Fields("RMUOM1").Value
                        .SetText 6, .ActiveRow, Val(rstMaterialSVChild.Fields("RMSTOCKKG").Value)
                        .SetText 7, .ActiveRow, Val(rstMaterialSVChild.Fields("UFGSTOCK").Value)
                        .SetText 8, .ActiveRow, Val(rstMaterialSVChild.Fields("Pending").Value)
                        GSM = Val(rstMaterialSVChild.Fields("GSM").Value)
                        ReelSize = Val(rstMaterialSVChild.Fields("Reel_Width").Value)
                        Cut_off = Val(rstMaterialSVChild.Fields("Cut_off").Value)
                        .GetText 15, .ActiveRow, CurVal
                        If CurVal = 0 Then
                            .SetText 9, .ActiveRow, IIf((Val(rstMaterialSVChild.Fields("PendingWIP").Value) - Val(rstMaterialSVChild.Fields("UFGSTOCK").Value)) > 0, (Val(rstMaterialSVChild.Fields("PendingWIP").Value) - Val(rstMaterialSVChild.Fields("UFGSTOCK").Value)), 0)
                        .GetText 9, .ActiveRow, CurVal
                        .SetText 10, .ActiveRow, ((CurVal / 20000 / 500) * ReelSize * GSM * Cut_off)
                        End If
                        .SetText 12, .ActiveRow, Val(Format(rstMaterialSVChild.Fields("CUT_OFF").Value, "###0.000"))
                        .SetText 13, .ActiveRow, Val(Format(rstMaterialSVChild.Fields("Reel_Width").Value, "###0.000"))
                        .SetText 14, .ActiveRow, Val(rstMaterialSVChild.Fields("GSM").Value)
                        .SetText 21, .ActiveRow, Val(rstMaterialSVChild.Fields("RMUOMVALUE").Value)
                        .SetText 22, .ActiveRow, rstMaterialSVChild.Fields("UOMCode").Value
                        .GetText 9, .ActiveRow, CurVal
                        If CurVal = 0 Then
                            .SetActiveCell 8, .ActiveRow
                         Else
                             .SetActiveCell 10, .ActiveRow
                         End If
                    End If
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                  End If
                ElseIf CurVal = "Title" Then
                    .GetText 18, .ActiveRow, TitleCode
                    On Error Resume Next
                    FrmBookMaster.SL = True
                    FrmBookMaster.ItemType = "F"
                    FrmBookMaster.MasterCode = Right(TitleCode, 6)
                    Load FrmBookMaster
                    If Err.Number <> 364 Then FrmBookMaster.Show vbModal
                    On Error GoTo 0
                    If slCode = "" Then Exit Sub
                    .SetText 2, .ActiveRow, slName: .SetText 18, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                ElseIf CurVal = "Element" Then
                    If rstElementList.RecordCount = 0 Then DisplayError ("No Record in Element Master"): .SetActiveCell 2, .ActiveRow: Exit Sub Else rstElementList.MoveFirst
                    .GetText 18, .ActiveRow, FinishedItemCode
                    Text9.Text = ""
                    If Not CheckEmpty(FinishedItemCode, False) Then
                        ElementCode = Left(FinishedItemCode, 6): FinishedItemCode = Right(FinishedItemCode, 6)
                        rstElementList.Find "[Code] = '" & RTrim(ElementCode) & "'"
                        Text9.Text = rstElementList.Fields("Col0").Value
                    End If
                    SelectionType = "S": ElementCode = ""
                    Call LoadSelectionList(rstElementList, "List of Elements...", "Name")
                    SearchOrder = 0
                    Call DisplaySelectionList(Text9, ElementCode)
                    Call CloseForm(FrmSelectionList)
                    If Not CheckEmpty(ElementCode, False) Then
                        On Error Resume Next
                        FrmBookMaster.SL = True
                        FrmBookMaster.ItemType = "F"
                        FrmBookMaster.MasterCode = FinishedItemCode
                        Load FrmBookMaster
                        If Err.Number <> 364 Then FrmBookMaster.Show vbModal
                        On Error GoTo 0
                        If slCode = "" Then Exit Sub
                        If Not CheckEmpty(slCode, False) Then .SetText 2, .ActiveRow, "[" + Text9.Text + "]-" + slName: .SetText 18, .ActiveRow, ElementCode + "-" + slCode: Sendkeys "{ENTER}"
                    End If
                    If CheckEmpty(ElementCode, False) Or CheckEmpty(slCode, False) Then .SetActiveCell 2, .ActiveRow: .SetText 2, .ActiveRow, "": .SetText 18, .ActiveRow, ""
                End If
            End If
        ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
            'Call RefreshDropDownList("R")
        ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyD) Then
        
        End If
    End With
End Sub
'Private Sub fpSpread2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Shift = 0 And KeyCode = vbKeyF9 Then
'        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
'            fpSpread2.DeleteRows fpSpread2.ActiveRow, 1
'            fpSpread2.SetFocus
'        End If
'    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
'        Call RefreshDropDownList("R")
'    End If
'End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal, Category As Variant
With fpSpread1
    .GetText 18, Row, Category
    .GetText Col, Row, ActiveCellVal
If Col = 1 Or Col = 9 Or Col = 10 Or Col = 15 Then
    If ActiveCellVal = "" Then
   '     Cancel = True
        Exit Sub
    End If
    .GetText 1, Row, Category
    If Col = 1 Then
        .Col = 2
    ElseIf Col = 9 Then
        .GetText Col, Row, ActiveCellVal
        If ActiveCellVal >= 0 Then
                Call Get_Deficient_Weight(True, False)
        .GetText 20, Row, ActiveCellVal
            If ActiveCellVal <> "" And moveFlag = True Then
                .GetText 15, Row, ActiveCellVal
                If ActiveCellVal <= 0 And PlanWeight >= 0 Then .SetActiveCell 15, .ActiveRow
            Else
                If moveFlag = True Then .SetActiveCell 11, .ActiveRow: moveFlag = False
            End If
        End If
    ElseIf Col = 10 Then
        .GetText Col, Row, ActiveCellVal
        If ActiveCellVal >= 0 Then
                Call Get_Deficient_Weight(False, True)
        .GetText 20, Row, ActiveCellVal
            If ActiveCellVal <> "" And moveFlag = True Then
                .GetText 15, Row, ActiveCellVal
                If ActiveCellVal <= 0 And PlanWeight >= 0 Then .SetActiveCell 15, .ActiveRow
            Else
                If moveFlag = True Then .SetActiveCell 11, .ActiveRow: moveFlag = False
            End If
        End If
    ElseIf Col = 15 Then
        Call Get_Deficient_Weight(True, True)
        .GetText Col, Row, ActiveCellVal
        If ActiveCellVal >= 0 Or ActualWeight <= 0 Then
            If ActiveCellVal = 0 And moveFlag = True Then .SetActiveCell 17, .ActiveRow
            moveFlag = False
        End If
    End If
End If
End With
End Sub
Sub Get_Deficient_Weight(ByVal CalcQtyWt As Boolean, CalcQtySheet As Boolean)
Dim ActiveCellVal As Variant
With fpSpread1
            .GetText 9, .ActiveRow, PlanSheet
            .GetText 10, .ActiveRow, PlanWeight
            .GetText 12, .ActiveRow, Cut_off
            .GetText 13, .ActiveRow, ReelSize
            .GetText 14, .ActiveRow, GSM
        If CalcQtyWt Then
            PlanWeight = ((PlanSheet / 20000 / 500) * ReelSize * GSM * Cut_off)
            .SetText 10, .ActiveRow, PlanWeight
        End If
        If CalcQtySheet Then
            On Error Resume Next
            PlanSheet = ((PlanWeight / ReelSize / GSM / Cut_off) * 20000 * 500)
            .SetText 9, .ActiveRow, PlanSheet
        End If
            .GetText 15, .ActiveRow, ActiveCellVal
            ActualWeight = ((ActiveCellVal / 20000 / 500) * ReelSize * GSM * Cut_off)
            .SetText 16, .ActiveRow, ActualWeight
            
        If PlanWeight = 0 And ActualWeight > 0 Then
            .GetText 15, .ActiveRow, PlanSheet
            .SetText 9, .ActiveRow, PlanSheet: .SetText 10, .ActiveRow, ActualWeight: .SetText 17, .ActiveRow, 0
        ElseIf ActualWeight <> 0 And PlanWeight <> 0 Then
            .SetText 17, .ActiveRow, ActualWeight - PlanWeight
        Else
            .SetText 17, .ActiveRow, ""
        End If
End With
End Sub

'Private Sub fpSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Dim ActiveCellVal As Variant, Category As Variant, Item As Variant
'
'    fpSpread2.GetText Col, Row, ActiveCellVal
'    If ActiveCellVal = "" Then
'        Cancel = True
'        Exit Sub
'    End If
'    fpSpread2.GetText 1, Row, Category
'    If Col = 1 Then
'        fpSpread2.Col = 2
'        fpSpread2.TypeComboBoxList = IIf(Category = "BOM", BOM, IIf(Category = "Paper", Paper, IIf(Category = "UFG", UFG, IIf(Category = "FG", FG, Title))))
'    ElseIf Col = 2 Then
'        If Category = "BOM" Then
'           If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
'           rstOutsourceItemList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
'           If Not rstOutsourceItemList.EOF Then
'                fpSpread2.SetText 4, Row, rstOutsourceItemList.Fields("NCode").Value
'                fpSpread2.SetText 6, Row, rstOutsourceItemList.Fields("UOMName").Value
'           End If
'        ElseIf Category = "Paper" Then
'           If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
'           rstPaperList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
'           If Not rstPaperList.EOF Then
'                fpSpread2.SetText 4, Row, rstPaperList.Fields("NCode").Value
'                fpSpread2.SetText 6, Row, rstPaperList.Fields("UOMName").Value
'           End If
'        ElseIf Category = "FG" Then
'           If rstFreshBookList.RecordCount > 0 Then rstFreshBookList.MoveFirst
'           rstFreshBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
'           If Not rstFreshBookList.EOF Then
'                fpSpread2.SetText 4, Row, rstFreshBookList.Fields("NCode").Value
'                fpSpread2.SetText 6, Row, rstFreshBookList.Fields("UOMName").Value
'            End If
'        ElseIf Category = "UFG" Then
'           If rstRepairBookList.RecordCount > 0 Then rstRepairBookList.MoveFirst
'           rstRepairBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
'           If Not rstRepairBookList.EOF Then
'                fpSpread2.SetText 4, Row, rstRepairBookList.Fields("NCode").Value
'                fpSpread2.SetText 6, Row, rstRepairBookList.Fields("UOMName").Value
'           End If
'        Else
'           If rstTitleList.RecordCount > 0 Then rstTitleList.MoveFirst
'           rstTitleList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
'           If Not rstTitleList.EOF Then
'                fpSpread2.SetText 4, Row, rstTitleList.Fields("NCode").Value
'                fpSpread2.SetText 6, Row, rstTitleList.Fields("UOMName").Value
'           End If
'        End If
'    End If
'End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
'Private Sub fpSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'    EditMode = IIf(Mode = 1, True, False)
'End Sub
Private Function CheckItem(ByVal xNumber As String) As Boolean
    Dim i As Integer, K As Integer, Item01 As Variant, Category01 As Variant, Item02 As Variant, Category02 As Variant
    CheckItem = False
    If xNumber = "1" Then
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.SetActiveCell 1, i
            fpSpread1.GetText 18, i, Item01
            fpSpread1.GetText 1, i, Category01
            If Category01 = "BOM Item" Then
                If Left(Item01, 1) <> "1" Then CheckItem = True
            ElseIf Category01 = "Paper" Then
                If Left(Item01, 1) <> "2" Then CheckItem = True
            ElseIf Category01 = "FG Item" Then
                If Left(Item01, 1) <> "3" Then CheckItem = True
            ElseIf Category01 = "UFG Item" Then
                If Left(Item01, 1) <> "4" Then CheckItem = True
            ElseIf Category01 = "Element" Then
                If Left(Item01, 1) <> "6" Then CheckItem = True
            Else
                If Left(Item01, 1) <> "5" Then CheckItem = True
            End If
            If CheckItem Then DisplayError "Data mismatch in row #" & Trim(Str(i)): Exit For
        Next
    ElseIf xNumber = "11" Then
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.SetActiveCell 1, i
            fpSpread1.GetText 20, i, Item01
            If Item01 = "" Then CheckItem = True
            If CheckItem Then DisplayError "Machine Can't Blank in row #" & Trim(Str(i)): fpSpread1.SetActiveCell 11, i: Exit For
        Next
    ElseIf xNumber = "18" Then
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.SetActiveCell 1, i
            fpSpread1.GetText xNumber, i, Item01
            If Item01 = "" Then CheckItem = True
            If CheckItem Then DisplayError "UFG Can't Blank in row #" & Trim(Str(i)): fpSpread1.SetActiveCell 2, i: Exit For
        Next
    ElseIf xNumber = "19" Then
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.SetActiveCell 1, i
            fpSpread1.GetText xNumber, i, Item01
            If Item01 = "" Then CheckItem = True
            If CheckItem Then DisplayError "RM Can't Blank in row #" & Trim(Str(i)): fpSpread1.SetActiveCell 3, i: Exit For
        Next
    ElseIf xNumber = "20" Then
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.SetActiveCell 1, i
            fpSpread1.GetText xNumber, i, Item01
            If Item01 = "" Then CheckItem = True
            If CheckItem Then DisplayError "Machine Can't Blank in row #" & Trim(Str(i)): fpSpread1.SetActiveCell 11, i: Exit For
        Next
    ElseIf xNumber = "22" Then
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.SetActiveCell 1, i
            fpSpread1.GetText xNumber, i, Item01
            If Item01 = "" Then CheckItem = True
            If CheckItem Then DisplayError "Machine Can't Blank in row #" & Trim(Str(i)): fpSpread1.SetActiveCell 5, i: Exit For
        Next
    Else
        For i = 1 To fpSpread2.DataRowCnt
            fpSpread2.SetActiveCell 1, i
            fpSpread2.GetText 18, i, Item01
            fpSpread2.GetText 1, i, Category01
            If Category01 = "BOM Item" Then
                If Left(Item01, 1) <> "1" Then CheckItem = True
            ElseIf Category01 = "Paper" Then
                If Left(Item01, 1) <> "2" Then CheckItem = True
            ElseIf Category01 = "FG Item" Then
                If Left(Item01, 1) <> "3" Then CheckItem = True
            ElseIf Category01 = "UFG Item" Then
                If Left(Item01, 1) <> "4" Then CheckItem = True
            ElseIf Category01 = "Element" Then
                If Left(Item01, 1) <> "6" Then CheckItem = True
            Else
                If Left(Item01, 1) <> "5" Then CheckItem = True
            End If
            If CheckItem Then DisplayError "Data mismatch in row #" & Trim(Str(i)): Exit For
        Next
    End If
    If Not CheckItem Then
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.SetActiveCell 1, i
            fpSpread1.GetText 1, i, Category01
            fpSpread1.GetText 18, i, Item01
            For K = 1 To fpSpread2.DataRowCnt
                fpSpread2.SetActiveCell 1, K
                fpSpread2.GetText 1, K, Category02
                fpSpread2.GetText 4, K, Item02
                If Category02 = Category01 And Item02 = Item01 Then
                    CheckItem = True
                    Exit For
                End If
            Next
            If CheckItem Then
                DisplayError "Same item cann't be generated (row #" & Trim(Str(i)) & ") and consumed (row #" & Trim(Str(K)) & ") simultaneously"
                Exit For
            End If
        Next
    End If
End Function
Private Sub LoadMaterialList(ByVal strOrderCode As String)
    Dim i As Integer, SQL As String
    On Error GoTo ErrorHandler
    eDate = GetDate(MhDateInput1.Text)
    If eDate = "00:00:00" Then eDate = Format(Now, "dd-mm-yyyy")
    If rstMaterialSVChild.State = adStateOpen Then rstMaterialSVChild.Close
    SQL = SQL + "Select *,(RMSTOCK*UOM)/500*((ReelSize*Cutoff*GSM)/20000) AS RMSTOCKKG,(Planning/500*((ReelSize*Cutoff*GSM)/20000)) AS PlanWeight,(QTY/500*((ReelSize*Cutoff*GSM)/20000)) AS SheetWeight,(Planning/500*((ReelSize*Cutoff*GSM)/20000))-(QTY/500*((ReelSize*Cutoff*GSM)/20000)) As Diff "
    SQL = SQL + "From( "
    SQL = SQL + "Select *,(Select Name From PaperMaster Where Code=Right(RMCode,6)) As RMName,(SELECT Convert(Numeric(12,3),dbo.ufnGetPaperStock('" & AccountCode & "',Right(RMCode,6),'JN','" & rstMaterialSVList.Fields("Code").Value & "','" & eDate & "')) As Col1 FROM PaperMaster P Where P.Code=Right(RMCode,6)) As RMSTOCK,(SELECT Name From GeneralMaster Where Code= (SElect UOM From PaperMaster Where Code=Right(RMCode,6))) As UNIT,(SELECT Value1 From GeneralMaster Where Code= (SElect UOM From PaperMaster Where Code=Right(RMCode,6))) As UOM,(Select UOM From PaperMaster Where Code=Right(RMCode,6)) AS UOMCODE, "
    SQL = SQL + "(Select Sum(Stock-Pending)From (SELECT *,(SELECT (dbo.ufnGetItemStock('*00002',FGList,'JN','" & rstMaterialSVList.Fields("Code").Value & "','" & eDate & "'))) AS Stock,(SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & sDate & "' AND '" & eDate & "' AND P.MaterialCentre IN ('*00002') AND C.Item =FGList) As Pending From (Select M.Code As FGList,M2.Item FROM BookMaster M Left JOIN BookChild01 M2 ON M.Code=M2.Code Where  M.Type='F' AND M2.Category='4' AND M.Code IN (Select Item From JobWorkBVRef) AND M2.Item=Right(ItemCode,6)) TBL) As TBL Group By Item) AS FinalPending "
    SQL = SQL + "From( "
    SQL = SQL + "SELECT C.Category,C.Category+C.Item As ItemCode,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=C.Item),IIF(Category='2',(SELECT LTRIM(M1.Name)+' (UOM : '+LTRIM(M2.Name)+')' As Name FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item),(SELECT Name FROM BookMaster WHERE Code=C.Item))) As ItemName,ABS(C.Quantity) As Qty,"
    SQL = SQL + "IIF(Category='1',(SELECT U.Name FROM OutsourceItemMaster M INNER JOIN GeneralMaster U ON M.UOM=U.Code WHERE M.Code=C.Item),IIF(Category='2',(SELECT LTRIM(M2.Name) As UOMName FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item),(SELECT LTRIM(M2.Name) As UOMName FROM BookMaster M1 INNER JOIN GeneralMaster M2 ON M1.IntegrationUnit=M2.Code WHERE M1.Code=C.Item))) As UOMName,Planning,Machine,(Select Name From MachineMaster Where Code=C.Machine) AS MachineName,CUTOFF,ReelSize,GSM,SNO,(SELECT Convert(Numeric(12,0),ISNULL(dbo.ufnGetUFGStock('4',Right(C.Item,6),'" & AccountCode & "','JN','" & rstMaterialSVList.Fields("Code").Value & "','" & eDate & "'),0))As Col1 FROM BookMaster I Where I.Code=Right(C.Item,6))As UFGStock,(Select '2'+Item From MaterialSVChild R Where Code='" & rstMaterialSVList.Fields("Code").Value & "' AND Category=2 AND C.SNO=R.SNO) AS RMCode  "
    SQL = SQL + "From MaterialSVChild C Where C.Code = '" & rstMaterialSVList.Fields("Code").Value & "' And Quantity >= 0 AND Category='4')AS TBL)AS TBL Order By SNO "
    rstMaterialSVChild.Open SQL, cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic

'        rstMaterialSVChild.Open "Select *,(RMSTOCK*UOM)/500*((ReelSize*Cutoff*GSM)/20000) AS RMSTOCKKG,(Planning/500*((ReelSize*Cutoff*GSM)/20000)) AS PlanWeight,(QTY/500*((ReelSize*Cutoff*GSM)/20000)) AS SheetWeight,(Planning/500*((ReelSize*Cutoff*GSM)/20000))-(QTY/500*((ReelSize*Cutoff*GSM)/20000)) As Diff " & _
'                                                 "From( " & _
'                                                 "Select *,(SElect Name From PaperMaster Where Code=Right(RMCode,6)) As RMName,(SELECT Convert(Numeric(12,3),dbo.ufnGetPaperStock('" & AccountCode & "',Right(RMCode,6),'JN','" & rstMaterialSVList.Fields("Code").Value & "','" & eDate & "')) As Col1 FROM PaperMaster P Where P.Code=Right(RMCode,6)) As RMSTOCK,(SELECT Name From GeneralMaster Where Code= (SElect UOM From PaperMaster Where Code=Right(RMCode,6))) As UNIT,(SELECT Value1 From GeneralMaster Where Code= (SElect UOM From PaperMaster Where Code=Right(RMCode,6))) As UOM " & _
'                                                 "From( " & _
'                                                 "SELECT C.Category,C.Category+C.Item As ItemCode,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=C.Item),IIF(Category='2',(SELECT LTRIM(M1.Name)+' (UOM : '+LTRIM(M2.Name)+')' As Name FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item),(SELECT Name FROM BookMaster WHERE Code=C.Item))) As ItemName,ABS(C.Quantity) As Qty,IIF(Category='1',(SELECT U.Name FROM OutsourceItemMaster M INNER JOIN GeneralMaster U ON M.UOM=U.Code WHERE M.Code=C.Item),IIF(Category='2',(SELECT LTRIM(M2.Name) As UOMName " & _
'                                                 "FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item),(SELECT LTRIM(M2.Name) As UOMName FROM BookMaster M1 INNER JOIN GeneralMaster M2 ON M1.IntegrationUnit=M2.Code WHERE M1.Code=C.Item))) As UOMName,Planning,Machine,(Select Name From MachineMaster Where Code=C.Machine) AS MachineName,CUTOFF,ReelSize,GSM,SNO,(SELECT Convert(Numeric(12,0),ISNULL(dbo.ufnGetUFGStock('4',Right(C.Item,6),'" & AccountCode & "','JN','" & rstMaterialSVList.Fields("Code").Value & "','" & eDate & "'),0))As Col1 FROM BookMaster I Where I.Code=Right(C.Item,6))As UFGStock," & _

'                                                 "ISNULL((Select ISNULL(ABS(ISNULL(Stock,0)-ISNULL(PendingSO,0)),0) as DeficientSalesOrder From(Select *,(SELECT (dbo.ufnGetItemStock('*00002',ItemCode,'R','" & rstMaterialSVList.Fields("Code").Value & "','" & eDate & "')) As Col1 " & _
'                                                 "FROM BookMaster I WHERE I.Type='F' AND I.Code=ItemCode) As Stock,ISNULL((SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & sDate & "' AND '" & eDate & "' AND P.MaterialCentre IN ('*00002') AND C.Item=ItemCode),0) As PendingSO From (SELECT I.Code As ItemCode,(Select Name From BookMaster Where Code=I1.Item) As UFG,I1.Category,(Select Code From BookMaster Where Code=I1.Item) As UFGCode FROM BookMaster I Left JOIN BookChild01 I1 ON I.Code=I1.Code WHERE I.Type='F' AND I1.Category='4' AND ISNULL((SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode " & _
'                                                 "WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & sDate & "' AND '" & eDate & "' AND P.MaterialCentre IN ('*00002') AND C.Item=I.Code),0)<>0 AND (Select Code From BookMaster Where Code=I1.Item) =C.Item) AS TBL) AS TBL) ,0) As DeficientSalesOrder,


'(Select '2'+Item From MaterialSVChild R Where Code='" & rstMaterialSVList.Fields("Code").Value & "' AND Category=2 AND C.SNO=R.SNO) AS RMCode " & _
'                                                 "From MaterialSVChild C Where C.Code = '" & rstMaterialSVList.Fields("Code").Value & "' And Quantity >= 0 AND Category='4')AS TBL)AS TBL Order By SNO ", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
       
    If rstMaterialSVChild.RecordCount > 0 Then rstMaterialSVChild.MoveFirst
    i = 0
    Do While Not rstMaterialSVChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, IIf(rstMaterialSVChild.Fields("Category").Value = "1", "BOM Item", IIf(rstMaterialSVChild.Fields("Category").Value = "2", "Paper", IIf(rstMaterialSVChild.Fields("Category").Value = "3", "FG Item", IIf(rstMaterialSVChild.Fields("Category").Value = "4", "UFG Item", IIf(rstMaterialSVChild.Fields("Category").Value = "6", "Element", "Title")))))
            .SetText 2, i, rstMaterialSVChild.Fields("ItemName").Value
            .SetText 3, i, rstMaterialSVChild.Fields("RMName").Value
            .SetText 4, i, Val(rstMaterialSVChild.Fields("RMSTOCK").Value)
            .SetText 5, i, rstMaterialSVChild.Fields("UNIT").Value
            .SetText 6, i, Val(rstMaterialSVChild.Fields("RMSTOCKKG").Value)
            .SetText 7, i, Val(rstMaterialSVChild.Fields("UFGSTOCK").Value)
             If Not IsNull(rstMaterialSVChild.Fields("FinalPending").Value) Then .SetText 8, i, Abs(Val(rstMaterialSVChild.Fields("FinalPending").Value))
            .SetText 9, i, Val(rstMaterialSVChild.Fields("Planning").Value)
            .SetText 10, i, Val(rstMaterialSVChild.Fields("PlanWeight").Value)
            .SetText 11, i, rstMaterialSVChild.Fields("MachineName").Value
            .SetText 12, i, Val(rstMaterialSVChild.Fields("Cutoff").Value)
            .SetText 13, i, Val(rstMaterialSVChild.Fields("ReelSize").Value)
            .SetText 14, i, Val(rstMaterialSVChild.Fields("GSM").Value)
            .SetText 15, i, Val(rstMaterialSVChild.Fields("QTY").Value)
            .SetText 16, i, Val(rstMaterialSVChild.Fields("SheetWeight").Value)
            .SetText 17, i, Val(rstMaterialSVChild.Fields("Diff").Value)
            .SetText 18, i, rstMaterialSVChild.Fields("ItemCODE").Value
            .SetText 19, i, rstMaterialSVChild.Fields("RMCODE").Value
            .SetText 20, i, rstMaterialSVChild.Fields("Machine").Value
            .SetText 21, i, Val(rstMaterialSVChild.Fields("UOM").Value)
            .SetText 22, i, rstMaterialSVChild.Fields("UOMCode").Value
            
        End With
        rstMaterialSVChild.MoveNext
    Loop
'    If rstMaterialSVChild.State = adStateOpen Then rstMaterialSVChild.Close
'    If DatabaseType = "MS SQL" Then
'        rstMaterialSVChild.Open "SELECT C.Category,C.Category+C.Item As ItemCode,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=C.Item),IIF(Category='2',(SELECT LTRIM(M1.Name)+' (UOM : '+LTRIM(M2.Name)+')' As Name FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item),(SELECT Name FROM BookMaster WHERE Code=C.Item))) As ItemName,ABS(C.Quantity) As Qty,IIF(Category='1',(SELECT U.Name FROM OutsourceItemMaster M INNER JOIN GeneralMaster U ON M.UOM=U.Code WHERE M.Code=C.Item),IIF(Category='2',(SELECT LTRIM(M2.Name) As UOMName FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item),'Piece')) As UOMName From MaterialSVChild C Where C.Code = '" & strOrderCode & "' And Quantity < 0 Order By Category", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
'    Else
'        rstMaterialSVChild.Open "SELECT C.Category,C.Category+C.Item As ItemCode,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=C.Item),IIF(Category='2',(SELECT TRIM(M1.Name)+' (UOM : '+LTRIM(M2.Name)+')' As Name FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item),(SELECT Name FROM BookMaster WHERE Code=C.Item))) As ItemName,ABS(C.Quantity) As Qty,IIF(Category='1',(SELECT U.Name FROM OutsourceItemMaster M INNER JOIN GeneralMaster U ON M.UOM=U.Code WHERE M.Code=C.Item),IIF(Category='2',(SELECT TRIM(M2.Name) As UOMName FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item),'Piece')) As UOMName From MaterialSVChild C Where C.Code = '" & strOrderCode & "' And Quantity < 0 Order By Category", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
'    End If
'    rstMaterialSVChild.ActiveConnection = Nothing
'    If rstMaterialSVChild.RecordCount > 0 Then rstMaterialSVChild.MoveFirst
'    i = 0
'    Do While Not rstMaterialSVChild.EOF
'        i = i + 1
'        With fpSpread2
'            .SetText 1, i, IIf(rstMaterialSVChild.Fields("Category").Value = "1", "BOM", IIf(rstMaterialSVChild.Fields("Category").Value = "2", "Paper", IIf(rstMaterialSVChild.Fields("Category").Value = "3", "FG", IIf(rstMaterialSVChild.Fields("Category").Value = "4", "UFG", "Title"))))
'            .Col = 2
'            .TypeComboBoxList = IIf(rstMaterialSVChild.Fields("Category").Value = "1", BOM, IIf(rstMaterialSVChild.Fields("Category").Value = "2", Paper, IIf(rstMaterialSVChild.Fields("Category").Value = "4", UFG, IIf(rstMaterialSVChild.Fields("Category").Value = "3", FG, Title))))
'            .SetText 2, i, rstMaterialSVChild.Fields("ItemName").Value
'            .SetText 3, i, Val(rstMaterialSVChild.Fields("Qty").Value)
'            .SetText 18, i, rstMaterialSVChild.Fields("ItemCode").Value
'            .SetText 6, i, rstMaterialSVChild.Fields("UOMName").Value
'        End With
'        rstMaterialSVChild.MoveNext
'    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Material List")
End Sub
Private Function UpdateMaterialList(ByVal ActionType As String, ByVal SrNo As Integer) As Boolean
    Dim CellVal(1 To 11) As Variant
    On Error GoTo ErrorHandler

    UpdateMaterialList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        cnMaterialStockAdjustment.Execute "Delete From MaterialSVChild WHERE Code = '" & rstMaterialSVParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I1" Then
        With fpSpread1
            .GetText 18, .ActiveRow, CellVal(1) 'UFGCODE
            .GetText 15, .ActiveRow, CellVal(2) 'Qty
            .GetText 9, .ActiveRow, CellVal(3) ' Plan
            .GetText 20, .ActiveRow, CellVal(4) 'MCODE
            .GetText 12, .ActiveRow, CellVal(5) 'Cutoff
            .GetText 13, .ActiveRow, CellVal(6) 'ReelSize
            .GetText 14, .ActiveRow, CellVal(7) 'GSM
            .GetText 22, .ActiveRow, CellVal(8) 'UOM
            .GetText 19, .ActiveRow, CellVal(9) 'RMCODE
            .GetText 21, .ActiveRow, CellVal(10) 'UOM
            CellVal(11) = (CellVal(3) / CellVal(10))
            CellVal(11) = Format((CellVal(3) / CellVal(10) - Int(CellVal(3) / CellVal(10))) * CellVal(10) / 1000 + Int(CellVal(3) / CellVal(10)), "##000.000")
        End With
        cnMaterialStockAdjustment.Execute "Insert Into MaterialSVChild Values ('" & rstMaterialSVParent.Fields("Code").Value & "','" & Left(CellVal(1), 1) & "','" & Right(CellVal(1), 6) & "'," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ",'" & CellVal(4) & "'," & Val(CellVal(5)) & "," & Val(CellVal(6)) & "," & Val(CellVal(7)) & ",'" & CellVal(8) & "'," & SrNo & ")"
        cnMaterialStockAdjustment.Execute "Insert Into MaterialSVChild Values ('" & rstMaterialSVParent.Fields("Code").Value & "','" & Left(CellVal(9), 1) & "','" & Right(CellVal(9), 6) & "'," & 0 - Val(CellVal(11)) & "," & Val(CellVal(3)) & ",'" & CellVal(4) & "'," & Val(CellVal(5)) & "," & Val(CellVal(6)) & "," & Val(CellVal(7)) & ",'" & CellVal(8) & "'," & SrNo & ")"
'    Else
'        With fpSpread2
'            .GetText 1, .ActiveRow, CellVal(1)
'            .GetText 3, .ActiveRow, CellVal(2)
'            .GetText 4, .ActiveRow, CellVal(3)
'        End With
'        cnMaterialStockAdjustment.Execute "Insert Into MaterialSVChild Values ('" & rstMaterialSVParent.Fields("Code").Value & "','" & IIf(CellVal(1) = "BOM", "1", IIf(CellVal(1) = "Paper", "2", IIf(CellVal(1) = "FG", "3", IIf(CellVal(1) = "UFG", "4", "5")))) & "','" & Right(CellVal(3), 6) & "'," & 0 - Val(CellVal(2)) & ")"
    End If
    Exit Function
ErrorHandler:
    UpdateMaterialList = False
End Function
'Private Sub RefreshDropDownList(ByVal xType As String)
'    If xType = "R" Then
'        rstOutsourceItemList.ActiveConnection = cnMaterialStockAdjustment
'        Do While Not RefreshRecord(rstOutsourceItemList)
'        Loop
'        rstOutsourceItemList.ActiveConnection = Nothing
'        rstPaperList.ActiveConnection = cnMaterialStockAdjustment
'        Do While Not RefreshRecord(rstPaperList)
'        Loop
'        rstPaperList.ActiveConnection = Nothing
'        rstFreshBookList.ActiveConnection = cnMaterialStockAdjustment
'        Do While Not RefreshRecord(rstFreshBookList)
'        Loop
'        rstFreshBookList.ActiveConnection = Nothing
'        rstTitleList.ActiveConnection = cnMaterialStockAdjustment
'        Do While Not RefreshRecord(rstTitleList)
'        Loop
'        rstTitleList.ActiveConnection = Nothing
'        rstRepairBookList.ActiveConnection = cnMaterialStockAdjustment
'        Do While Not RefreshRecord(rstRepairBookList)
'        Loop
'        rstRepairBookList.ActiveConnection = Nothing
'        rstElementList.ActiveConnection = cnMaterialStockAdjustment
'        Do While Not RefreshRecord(rstElementList)
'        Loop
'        rstElementList.ActiveConnection = Nothing
'        BOM = "": Paper = "": FG = "": UFG = "": Title = "": Element = ""
'    End If
'    Do While Not rstOutsourceItemList.EOF
'        BOM = IIf(CheckEmpty(BOM, False), "", BOM + Chr$(9)) + rstOutsourceItemList.Fields("Name").Value
'        rstOutsourceItemList.MoveNext
'    Loop
'    Do While Not rstPaperList.EOF
'        Paper = IIf(CheckEmpty(Paper, False), "", Paper + Chr$(9)) + rstPaperList.Fields("Name").Value
'        rstPaperList.MoveNext
'    Loop
'    Do While Not rstFreshBookList.EOF
'        FG = IIf(CheckEmpty(FG, False), "", FG + Chr$(9)) + rstFreshBookList.Fields("Name").Value
'        rstFreshBookList.MoveNext
'    Loop
'    Do While Not rstTitleList.EOF
'        Title = IIf(CheckEmpty(Title, False), "", Title + Chr$(9)) + rstTitleList.Fields("Name").Value
'        rstTitleList.MoveNext
'    Loop
'    rstTitleList.Filter = adFilterNone
'    Do While Not rstRepairBookList.EOF
'        UFG = IIf(CheckEmpty(UFG, False), "", UFG + Chr$(9)) + rstRepairBookList.Fields("Name").Value
'        rstRepairBookList.MoveNext
'    Loop
'    rstElementList.Filter = adFilterNone
'    Do While Not rstElementList.EOF
'        Element = IIf(CheckEmpty(Element, False), "", Element + Chr$(9)) + rstElementList.Fields("Name").Value
'        rstElementList.MoveNext
'    Loop
'End Sub
Private Sub PrintMaterialStockAdjustment()
    On Error Resume Next
If MsgBox("Do You want's to Print Grid Sheet Format?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") = vbYes Then
        Call Print_Sheet
Else
    Screen.MousePointer = vbHourglass
    rptStockJournal.Text1.SetText "Stock Journal Voucher"
    rptStockJournal.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptStockJournal.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False)) Then
        rptStockJournal.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False) Then
        rptStockJournal.Text24.SetText "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptStockJournal.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value)
    Else
        rptStockJournal.Section5.Suppress = True
    End If
    If rstMaterialSVChild.State = adStateOpen Then rstMaterialSVChild.Close
    If DatabaseType = "MS SQL" Then
        rstMaterialSVChild.Open "SELECT LTRIM(Name) As VchNo,[Date] As VchDate,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.Account) As Godown,Category,CASE WHEN Category='1' THEN (SELECT LTRIM(PrintName) FROM OutsourceItemMaster WHERE Code=C.Item) WHEN Category='2' THEN (SELECT LTRIM(M1.PrintName)+' (UOM : '+LTRIM(M2.PrintName)+'='+LTRIM(M2.Value1)+')' FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item) ELSE (SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=C.Item) END As ItemName,CASE WHEN Quantity>=0 THEN 'Items Generated' ELSE 'Items Consumed' END As ItemType,Quantity,Remarks,IIF(Category='1',(SELECT U.Name FROM OutsourceItemMaster M INNER JOIN GeneralMaster U ON M.UOM=U.Code WHERE M.Code=C.Item),IIF(Category='2',(SELECT LTRIM(M2.Name) As UOMName FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item),'Piece')) As UOMName, " & _
        "IIF(Category='1','BOM',IIF(Category='2','Paper',IIF(Category='3','FG',IIF(Category='4','UFG',('Title'))))) As CategoryName FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE P.Code='" & rstMaterialSVList.Fields("Code").Value & "' ", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    Else
        rstMaterialSVChild.Open "SELECT LTRIM(Name) As VchNo,[Date] As VchDate,(SELECT LTRIM(PrintName) FROM AccountMaster WHERE Code=P.Account) As Godown,Category,IIF(Category='1',(SELECT LTRIM(PrintName) FROM OutsourceItemMaster WHERE Code=C.Item),IIF(Category='2',(SELECT LTRIM(M1.PrintName)+' (UOM : '+LTRIM(M2.PrintName)+'='+LTRIM(M2.Value1)+')' FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item),(SELECT LTRIM(PrintName) FROM BookMaster WHERE Code=C.Item))) As ItemName,IIF(Quantity>=0,'Items Generated','Items Consumed') As ItemType,Quantity,Remarks,IIF(Category='1',(SELECT U.Name FROM OutsourceItemMaster M INNER JOIN GeneralMaster U ON M.UOM=U.Code WHERE M.Code=C.Item),IIF(Category='2',(SELECT LTRIM(M2.Name) As UOMName FROM PaperMaster M1 INNER JOIN GeneralMaster M2 ON M1.UOM=M2.Code WHERE M1.Code=C.Item),'Piece')) As UOMName, " & _
        "IIF(Category='1','BOM',IIF(Category='2','Paper',IIF(Category='3','FG',IIF(Category='4','UFG',('Title'))))) As CategoryName FROM MaterialSVParent P Left Join MaterialSVChild C On (P.Code=C.Code And P.Code='" & rstMaterialSVList.Fields("Code").Value & "' )", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    End If
    rptStockJournal.Text27.SetText "for " & Trim(rstMaterialSVChild.Fields("Godown").Value)
    rptStockJournal.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptStockJournal.Database.SetDataSource rstMaterialSVChild, 3, 1
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptStockJournal
        FrmReportViewer.Show vbModal
    Else
        rptStockJournal.PaperSource = crPRBinAuto
        rptStockJournal.PrintOut
    End If
    Set rptStockJournal = Nothing
    On Error GoTo 0

End If
End Sub
Private Sub LoadMasterList()
    On Error Resume Next
    If rstMaterialSVList.EOF Then VchCode = "XXXXXX" Else VchCode = rstMaterialSVList.Fields("Code").Value
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "SELECT LTRIM(Name) As Col0,Code FROM AccountMaster ORDER BY Name", cnMaterialStockAdjustment, adOpenKeyset, adLockReadOnly
    If rstUserList.State = adStateOpen Then rstUserList.Close
    rstUserList.Open "SELECT Name As Col0,Code FROM UserMaster WHERE " & IIf(UserLevel = 1, "1=1", "Code='" & UserCode & "'") & " AND VchApprovalRights=" & IIf(DatabaseType = "MS SQL", 1, True) & " ORDER BY Name", cnMaterialStockAdjustment, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    rstUserList.ActiveConnection = Nothing
    If rstMaterialSVChild.State = adStateOpen Then rstMaterialSVChild.Close
'    rstMaterialSVChild.Open "Select UFG,UFGCODE1,RM,RMCODE,RMUOMVALUE,UOMCODE,RM_UNIT_Weight,RMFORM,RMSTOCK AS RMSTOCK,RMUOM AS RMUOM1,(Convert(numeric,PARSENAME(RMSTOCK,2))*RM_UNIT_Weight)+((RMSTOCK-Convert(numeric,PARSENAME(RMSTOCK,2)))/RMUOMVALUE*1000*RM_UNIT_Weight) AS RMSTOCKKG,IIF(RMFORM='R','KG',RMUOM) AS RMUOM2,UFGSTOCK,PendingSO As Pending,IIF((PendingSO-UFGSTOCK)>0,(PendingSO-UFGSTOCK),0)Planning,'' As Production,'' As Machine,Cut_off,Reel_Width,GSM " & _
'                                            "FROM(" & _
'                                            "Select *,(SELECT Convert(Numeric(12,0),ISNULL(dbo.ufnGetUFGStock('4',UFGCODE1,'" & AccountCode & "','JN','" & VchCode & "','" & eDate & "'),0))As Col1 FROM BookMaster I Where I.Code=UFGCODE1)As UFGSTOCK,(SELECT Convert(Numeric(12,3),dbo.ufnGetPaperStock('" & AccountCode & "',RMCODE,'JN','" & VchCode & "','" & eDate & "')) As Col1 FROM PaperMaster P Where P.Code=RMCODE) As RMSTOCK," & _
'                                            "ISNULL((Select ABS(SUM(Stock-PendingSO)) as DeficientSalesOrder " & _
'                                            "From(" & _
'                                            "SELECT Distinct *,(SELECT (dbo.ufnGetItemStock('*00002',ItemCode,'JN','" & VchCode & "','" & eDate & "')) As Col1 FROM BookMaster I WHERE I.Type='F' AND I.Code=ItemCode) As Stock," & _
'                                            "ISNULL((SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & sDate & "' AND '" & eDate & "' AND P.MaterialCentre IN ('*00002') AND C.Item=ItemCode),0)+ISNULL((SELECT SUM(ABS(R.Quantity)) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & sDate & "' AND '" & eDate & "' AND P.MaterialCentre IN ('*00002') AND C.Item=ItemCode AND R.RefCode=C.RefCode AND VchCode<>C.Code),0) AS SalesOrder," & _
'                                            "ISNULL((SELECT SUM(ABS(R.Quantity)) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & sDate & "' AND '" & eDate & "' AND P.MaterialCentre IN ('*00002') AND C.Item=ItemCode AND R.RefCode=C.RefCode AND VchCode<>C.Code),0) As Dispatched," & _
'                                            "ISNULL((SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & sDate & "' AND '" & eDate & "' AND P.MaterialCentre IN ('*00002') AND C.Item=ItemCode),0) As PendingSO " & _
'                                            "FROM(" & _
'                                            "SELECT I.Code As ItemCode,(Select Name From BookMaster Where Code=I1.Item) As UFG,I1.Category,(Select Code From BookMaster Where Code=I1.Item) As UFGCode2 FROM BookMaster I Left JOIN BookChild01 I1 ON I.Code=I1.Code WHERE I.Type='F' AND I1.Category='4' AND ISNULL((SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & sDate & "' AND '" & eDate & "' AND P.MaterialCentre IN ('*00002') AND C.Item=I.Code),0)<>0)AS TBL) As FG_UFG where UFGCode2=UFGCODE1),0) AS PendingSO " & _
'                                            "FROM(" & _
'                                            "Select I.Name AS UFG,I.Code AS UFGCODE1,(Select NAme From PaperMaster Where Code=C.Item) As RM,(SELECT Value1 From GeneralMaster Where Code=(Select UOM From PaperMaster Where Code=C.Item)) RMUOMVALUE,(Select UOM From PaperMaster Where Code=C.Item) As UOMCODE,(Select [Weight/Unit] From PaperMaster Where Code=C.Item) AS RM_UNIT_Weight,(SELECT NAME From GeneralMaster Where Code=(Select UOM From PaperMaster Where Code=C.Item)) As RMUOM,(Select cmLength From PaperMaster Where Code=C.Item) AS Cut_off,(Select cmWidth From PaperMaster Where Code=C.Item) AS Reel_Width,(Select GSM From PaperMaster Where Code=C.Item) AS GSM,(Select Form From PaperMaster Where Code=C.Item) AS RMFORM,(Select Code From PaperMaster Where Code=C.Item) As RMCODE " & _
'                                            "From BookMaster I Left JOIN BookChild01 C ON I.Code=C.Code Where Type='R')TBL)TBL Where UFGCODE1='" & Right(UFGCode, 6) & "'  ", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstMaterialSVChild.Open "Select UFG,UFGCODE1,RM,RMCODE,RMUOMVALUE,UOMCODE,RM_UNIT_Weight,RMFORM,RMSTOCK AS RMSTOCK,RMUOM AS RMUOM1,(Convert(numeric,PARSENAME(RMSTOCK,2))*RM_UNIT_Weight)+((RMSTOCK-Convert(numeric,PARSENAME(RMSTOCK,2)))/RMUOMVALUE*1000*RM_UNIT_Weight) AS RMSTOCKKG,IIF(RMFORM='R','KG',RMUOM) AS RMUOM2,UFGSTOCK,PendingSO As Pending,PendingWIP,IIF((PendingSO-UFGSTOCK)>0,(PendingSO-UFGSTOCK),0)Planning,'' As Production,'' As Machine,Cut_off,Reel_Width,GSM " & _
                                            "FROM(" & _
                                            "Select *,(SELECT Convert(Numeric(12,0),ISNULL(dbo.ufnGetUFGStock('4',UFGCODE1,'" & AccountCode & "','JN','" & VchCode & "','" & eDate & "'),0))As Col1 FROM BookMaster I Where I.Code=UFGCODE1)As UFGSTOCK,(SELECT Convert(Numeric(12,3),dbo.ufnGetPaperStock('" & AccountCode & "',RMCODE,'JN','" & VchCode & "','" & eDate & "')) As Col1 FROM PaperMaster P Where P.Code=RMCODE) As RMSTOCK," & _
                                            "(Select SUM(ABS(Pending)) From (Select *,(Stock-Pending) As PendingSO From (SELECT *,(Select Quantity From BookChild01 Where Code=FGList AND Item=UFGCODE1) AS WIPPU,(SELECT (dbo.ufnGetItemStock('*00002',FGList,'JN','" & VchCode & "','" & eDate & "'))) AS Stock,(SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & sDate & "' AND '" & eDate & "' AND P.MaterialCentre IN ('*00002') AND C.Item =FGList) As Pending From (Select M.Code As FGList,M2.Item FROM BookMaster M Left JOIN BookChild01 M2 ON M.Code=M2.Code Where  M.Type='F' AND M2.Category='4' AND M.Code IN (Select Item From JobWorkBVRef) AND M2.Item=Right(UFGCODE1,6))TBL) As TBL) As TBL Group By Item) AS PendingSO, " & _
                                            "(Select SUM(ABS(pWIP)) From (Select *,(Stock-Pending)*WIPPU As pWIP From (SELECT *,(Select Quantity From BookChild01 Where Code=FGList AND Item=UFGCODE1) AS WIPPU,(SELECT (dbo.ufnGetItemStock('*00002',FGList,'JN','" & VchCode & "','" & eDate & "'))) AS Stock,(SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & sDate & "' AND '" & eDate & "' AND P.MaterialCentre IN ('*00002') AND C.Item =FGList) As Pending From (Select M.Code As FGList,M2.Item FROM BookMaster M Left JOIN BookChild01 M2 ON M.Code=M2.Code Where  M.Type='F' AND M2.Category='4' AND M.Code IN (Select Item From JobWorkBVRef) AND M2.Item=Right(UFGCODE1,6))TBL) As TBL) As TBL Group By Item) AS PendingWIP " & _
                                            "FROM(" & _
                                            "Select I.Name AS UFG,I.Code AS UFGCODE1,(Select NAme From PaperMaster Where Code=C.Item) As RM,(SELECT Value1 From GeneralMaster Where Code=(Select UOM From PaperMaster Where Code=C.Item)) RMUOMVALUE,(Select UOM From PaperMaster Where Code=C.Item) As UOMCODE,(Select [Weight/Unit] From PaperMaster Where Code=C.Item) AS RM_UNIT_Weight,(SELECT NAME From GeneralMaster Where Code=(Select UOM From PaperMaster Where Code=C.Item)) As RMUOM,(Select cmLength From PaperMaster Where Code=C.Item) AS Cut_off,(Select cmWidth From PaperMaster Where Code=C.Item) AS Reel_Width,(Select GSM From PaperMaster Where Code=C.Item) AS GSM,(Select Form From PaperMaster Where Code=C.Item) AS RMFORM,(Select Code From PaperMaster Where Code=C.Item) As RMCODE " & _
                                            "From BookMaster I Left JOIN BookChild01 C ON I.Code=C.Code Where Type='R')TBL)TBL Where UFGCODE1='" & Right(UFGCode, 6) & "'  ", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    
    rstMaterialSVChild.ActiveConnection = Nothing
    If rstMaterialSVChild.RecordCount = 0 Then
    End If
End Sub
Private Sub LoadRMStock(Optional ByVal LoadSelected As Boolean)
If rstMaterialSVList.EOF Then VchCode = "XXXXXX" Else VchCode = rstMaterialSVList.Fields("Code").Value
If rstRMStock.State = adStateOpen Then rstRMStock.Close
On Error Resume Next
    rstRMStock.Open "Select *,(Convert(numeric,PARSENAME(RMSTOCK,2))*RM_UNIT_Weight)+((RMSTOCK-Convert(numeric,PARSENAME(RMSTOCK,2)))/RMUOMVALUE*1000*RM_UNIT_Weight) AS RMSTOCKKG From(SELECT Convert(Numeric(12,3),dbo.ufnGetPaperStock('" & AccountCode & "','" & Right(PaperCode, 6) & "','JN','" & VchCode & "','" & GetDate(MhDateInput1.Text) & "')) As RMSTOCK,(Select Name From GeneralMAster Where Code=UOM) As Unit,[Weight/Unit] AS RM_UNIT_Weight,(Select Value1 From GeneralMAster Where Code=UOM) As RMUOMVALUE,UOM AS UOMCODE FROM PaperMaster P Where P.Code='" & Right(PaperCode, 6) & "')TBL ", cnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
rstRMStock.ActiveConnection = Nothing
End Sub
Sub Print_Sheet()
        Dim C As Long, PrintHeader As String, CurVal As Variant
With fpSpread1

    If fpSpread1.EditMode = True Then
        fpSpread1.EditMode = False
        DoEvents
    End If
    Set spreadpreview.frm = Me
    Set pagesetup.frmPageSetup = Me
    Set PrintDlg.frmPrintDlg = Me
    Set headerfooter.frmHeaderFooter = Me
    spreadpreview.Show

'
'        .Col = 1: .ColHidden = True
'        .ColHeaderRows = 3:
'        For C = 1 To .MaxCols
'        .GetText C, -999, CurVal
'        .Col = C: .Row = SpreadHeader + 2: .Text = CurVal:
'        Next
'        .Col = 1: .Row = SpreadHeader + 1: .Text = " Production Planning": .FontSize = 12:
'        .Col = 11: .Row = SpreadHeader + 1: .Text = " Production & Job Performance": .FontSize = 12:
'        .AddCellSpan 1, SpreadHeader + 1, 10, 1
'        .AddCellSpan 11, SpreadHeader + 1, 17, 1
'        .Col = 1: .Row = SpreadHeader: .Text = "[" + rstCompanyMaster.Fields("PrintName").Value + "]" + "  Date : " + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy"): .FontSize = 20
'        .AddCellSpan 1, SpreadHeader, 17, 1
'                PrintHeader = Me.Caption
'                ' These are 8.5" X 11" paper dimensions in TWIPS
'                Printer.PaperSize = vbPRPSA4
'                ' Set printing options for sheet
''                .PrintAbortMsg = "Printing - Click Cancel to .Quit"
'                .PrintJobName = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") '& ".pdf"
'                .PrintFooter = "        Export Data Company : " & rstCompanyMaster.Fields("PrintName").Value & " _(" & CompCode & "_" & PrintHeader & ")" & "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]" & " /rPage # ./p " & " Print Date : ( " & Format(Date, "dd-MMM-yyyy") & " )         "
'                .PrintBorder = True
'                .PrintColHeaders = True
'                .PrintColor = True
'                .PrintGrid = True
'                .PrintMarginTop = 500 '1440
'                .PrintMarginBottom = 500 '1440
'                .PrintMarginLeft = 50 '720
'                .PrintMarginRight = 50 '720
'                '.PrintType = SPRD_PRINT_ALL
'                .PrintRowHeaders = True
'                .PrintShadows = True
'                .PrintUseDataMax = True
'                ' Center vertically
'                .PrintCenterOnPageV = False
'                ' Center horizontally
'                .PrintCenterOnPageH = True
'                ' Perform the printing action
'                ' Set the sheet to print
'                .Sheet = 1
'                ' Set scaling method
'                .PrintScalingMethod = PrintScalingMethodZoom
'                ' Set zoom factor
'                .PrintZoomFactor = 0
'                ' Print
'                '.PrintSheet 0
'                .PrintOrientation = PrintOrientationLandscape
'                .PrintSheet
'        .Col = 1: .ColHidden = False
'        .ColHeaderRows = 2
'        .AddCellSpan 1, SpreadHeader + 1, 1, 1
'        .AddCellSpan 11, SpreadHeader + 1, 1, 1
'        .Col = 1: .Row = SpreadHeader: .Text = " Production Planning": .FontSize = 12:
'        .Col = 11: .Row = SpreadHeader: .Text = " Production & Job Performance": .FontSize = 12:
'        .Col = 1: .Row = SpreadHeader + 1: .Text = "Category": .FontSize = 9.75
'        .Col = 11: .Row = SpreadHeader + 1: .Text = "Machine": .FontSize = 9.75
'        .AddCellSpan 1, SpreadHeader, 10, 1
'        .AddCellSpan 11, SpreadHeader, 17, 1
End With
End Sub
