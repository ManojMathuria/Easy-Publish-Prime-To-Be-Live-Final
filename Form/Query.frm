VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Planning Register"
   ClientHeight    =   9570
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   16980
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   16980
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9180
      Left            =   120
      TabIndex        =   8
      Top             =   345
      Width           =   16890
      _Version        =   65536
      _ExtentX        =   29792
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
      Picture         =   "Query.frx":0000
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Query.frx":001C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Query.frx":0560
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Query.frx":0674
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Query.frx":0786
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8400
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   1140
         TabIndex        =   14
         Top             =   8100
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
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
         Caption         =   " Data Count"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "Query.frx":0898
         Picture         =   "Query.frx":08B4
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   435
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   8100
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   582
         Calculator      =   "Query.frx":08D0
         Caption         =   "Query.frx":08F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Query.frx":095C
         Keys            =   "Query.frx":097A
         Spin            =   "Query.frx":09C4
         AlignHorizontal =   1
         AlignVertical   =   2
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
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   13320
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   " Quantity Total"
         Top             =   8100
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   582
         Calculator      =   "Query.frx":09EC
         Caption         =   "Query.frx":0A0C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Query.frx":0A78
         Keys            =   "Query.frx":0A96
         Spin            =   "Query.frx":0AE0
         AlignHorizontal =   1
         AlignVertical   =   2
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
         MinValue        =   -9999999999
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   14475
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   " Amount Total"
         Top             =   8100
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   582
         Calculator      =   "Query.frx":0B08
         Caption         =   "Query.frx":0B28
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Query.frx":0B94
         Keys            =   "Query.frx":0BB2
         Spin            =   "Query.frx":0BFC
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
         MinValue        =   -9999999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1638405
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   12120
         TabIndex        =   13
         Top             =   8100
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
         Caption         =   "  Totals"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "Query.frx":0C24
         Picture         =   "Query.frx":0C40
         Begin VB.CheckBox Check1 
            Height          =   225
            Left            =   720
            TabIndex        =   20
            Top             =   40
            Width           =   225
         End
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
         Left            =   1200
         TabIndex        =   0
         Top             =   8700
         Width           =   10815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   7395
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   17145
         _ExtentX        =   30242
         _ExtentY        =   13044
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   9164542
         ForeColor       =   0
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column02 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column04 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column07 
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column08 
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column09 
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column10 
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column11 
            DataField       =   ""
            Caption         =   ""
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            Locked          =   -1  'True
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1709.858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   3525.166
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   3569.953
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1500.095
            EndProperty
         EndProperty
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   12480
         TabIndex        =   9
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " &From"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "Query.frx":0C5C
         Picture         =   "Query.frx":0C78
      End
      Begin MSMask.MaskEdBox MhDateInput1 
         Height          =   330
         Left            =   13320
         TabIndex        =   3
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   14760
         TabIndex        =   10
         Top             =   120
         Width           =   765
         _Version        =   65536
         _ExtentX        =   1349
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
         Caption         =   " &To"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
      End
      Begin MSMask.MaskEdBox MhDateInput2 
         Height          =   330
         Left            =   15510
         TabIndex        =   4
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   120
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
         Caption         =   " Accounts Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   7920
         TabIndex        =   18
         Top             =   120
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
         Caption         =   " Voucher Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
      End
      Begin Mh3dlblLib.Mh3dLabel CmdExport 
         Height          =   330
         Left            =   13150
         TabIndex        =   6
         Top             =   8700
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
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
         Caption         =   "Export List"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "Query.frx":0C94
         Picture         =   "Query.frx":0CB0
      End
      Begin Mh3dlblLib.Mh3dLabel CmdPrint 
         Height          =   330
         Left            =   12120
         TabIndex        =   5
         Top             =   8700
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
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
         Caption         =   " Print List"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "Query.frx":0CCC
         Picture         =   "Query.frx":0CE8
      End
      Begin Mh3dlblLib.Mh3dLabel CmdLabel 
         Height          =   330
         Left            =   14200
         TabIndex        =   7
         Top             =   8700
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
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
         Caption         =   " Print Vch."
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "Query.frx":0D04
         Picture         =   "Query.frx":0D20
      End
      Begin MSForms.ComboBox ComboBox3 
         Height          =   330
         Left            =   15240
         TabIndex        =   21
         Top             =   8700
         Width           =   1575
         VariousPropertyBits=   612390939
         DisplayStyle    =   3
         Size            =   "2778;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Comic Sans MS"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox ComboBox2 
         Height          =   330
         Left            =   9360
         TabIndex        =   2
         Top             =   120
         Width           =   3015
         VariousPropertyBits=   612390939
         DisplayStyle    =   3
         Size            =   "5318;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Comic Sans MS"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   16900
         Y1              =   8610
         Y2              =   8610
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
         ForeColor       =   &H80000009&
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   8700
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   16920
         Y1              =   600
         Y2              =   600
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   330
         Left            =   1560
         TabIndex        =   1
         Top             =   120
         Width           =   3735
         VariousPropertyBits=   1686132763
         DisplayStyle    =   3
         Size            =   "6588;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Comic Sans MS"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Preview [Alt+V]"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print [Alt+P] "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mail [Alt+M]"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit [Escape]"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Export [Alt+E]"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh [F5]"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancel [Escape]"
         EndProperty
      EndProperty
      Begin VB.CommandButton cmdRefresh 
         Height          =   375
         Left            =   1680
         Picture         =   "Query.frx":0D3C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Refresh [F5]"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Refresh 
         Height          =   375
         Left            =   1680
         Picture         =   "Query.frx":0E86
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Refresh [F5]"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   375
         Left            =   2040
         Picture         =   "Query.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Cancel  [Escape]"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Export 
         Height          =   375
         Left            =   1320
         Picture         =   "Query.frx":10D2
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Export [Alt+E]"
         Top             =   0
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PtgType As String 'IIf(PtgType = 1, "Print Selection Voucher Format-1", IIf(PtgType = 2, "Print Selection Voucher Format-2", IIf(PtgType = 3, "Print All Voucher Of Data Grid", "Print Data Grid")))
Dim cnQuery As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset, rstQueryList As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstQueryDetails As New ADODB.Recordset
Dim OutputTo As String, AutoMode As Boolean
Dim PrevStr As String, SortCol As String, SortOrder As String, HiLiteRecord  As Boolean, i As Double, ATotal As Double, QTotal As Double, VchCode As String
Dim StartColumn As String, StartRow As String, EndColumn As String, EndRow As String, PrintFlag As Boolean
Dim VchType, vDate, vtType, vtCode, vtNo As String
Public dSortBy As Boolean
Private Sub cmdCancel_Click()
    Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
End Sub
Private Sub cmdRefresh_Click()
    Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
        WheelHook DataGrid1
    BusySystemIndicator True
        Me.Caption = "Sales Query Ledger"
    cnQuery.CursorLocation = adUseClient: cnQuery.Open cnDatabase.ConnectionString:
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    If Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd") Then
        MhDateInput2.Text = Format(FinancialYearTo, "dd-mm-yyyy")
    Else
        MhDateInput2.Text = Format(Date, "dd-mm-yyyy")
    End If
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
        rstAccountList.Open "Select PrintName+' ['+Code+']' AS PrintName,PrintName AS Name,Code FROM AccountMaster Order By PrintName", cnQuery, adOpenKeyset, adLockPessimistic
    rstAccountList.ActiveConnection = Nothing
'ComboBox1 Accounts
    ComboBox1.Clear
    i = 0
    ComboBox1.AddItem "All Accounts [XXXXXX]", i
    rstAccountList.MoveFirst
    Do While Not rstAccountList.EOF
    i = i + 1
    ComboBox1.AddItem rstAccountList.Fields("PrintName").Value, i
    rstAccountList.MoveNext
    Loop

'ComboBox2 Voucher Type
    ComboBox2.Clear
    ComboBox2.FontSize = 9
    ComboBox2.AddItem "Purchase", 0 '1
    ComboBox2.AddItem "Purchase Return", 1 '2
    ComboBox2.AddItem "Sale Return", 2 '3
    ComboBox2.AddItem "Sales", 3 '4
    ComboBox2.AddItem "Purchase Challan IN", 4 '5
    ComboBox2.AddItem "Purchase Challan Out", 5 '6
    ComboBox2.AddItem "Sale Challan IN", 6 '7
    ComboBox2.AddItem "Sale Challan Out", 7 '8
    
    ComboBox2.AddItem "Purchase Order", 8 '17
    ComboBox2.AddItem "Sale Order", 9 '18
    ComboBox2.AddItem "Stock Tranfer", 10 '19
    ComboBox2.AddItem "Stock Genral", 11 '20
    ComboBox2.AddItem "Promotional Sale Challan Out", 12 '21
    ComboBox2.AddItem "Promotional Purchase Challan Out", 13 '22
    ComboBox2.AddItem "Purchase Quotation", 14 '23
    ComboBox2.AddItem "Sales Quotation", 15 '24
    
    ComboBox2.AddItem "Payments", 16 '51
    ComboBox2.AddItem "Receipts", 17 '52
    ComboBox2.AddItem "Journal", 18 '53
    ComboBox2.AddItem "Countra", 19 '54
    ComboBox2.AddItem "Debit Note", 20 '55
    ComboBox2.AddItem "Credit Note", 21 '56
    
    ComboBox3.Clear
    ComboBox3.AddItem "Print Selection", 0
    ComboBox3.AddItem "Print All", 1
    ComboBox3.ListIndex = 0
    
    ComboBox1.ListIndex = 0
    ComboBox2.ListIndex = 3
    ComboBox2_Validate (True)
    LoadMasterList

'DataGrid1_Width
    DataGrid1.Columns(0).Width = 645
    DataGrid1.Columns(1).Width = 900
    DataGrid1.Columns(2).Width = 930
    DataGrid1.Columns(3).Width = 1200
    DataGrid1.Columns(4).Width = 1710
    DataGrid1.Columns(5).Width = 1200
    DataGrid1.Columns(6).Width = 2775
    DataGrid1.Columns(7).Width = 3525
    DataGrid1.Columns(8).Width = 1155
    DataGrid1.Columns(9).Width = 1155
    DataGrid1.Columns(10).Width = 3535
    DataGrid1.Columns(11).Width = 1500
'DataGrid1_Caption
    DataGrid1.Columns(0).Caption = "S. No."
    DataGrid1.Columns(1).Caption = "Vch Code"
    DataGrid1.Columns(2).Caption = "Vch. Type"
    DataGrid1.Columns(3).Caption = "Vch. Series"
    DataGrid1.Columns(4).Caption = "Vch. No."
    DataGrid1.Columns(5).Caption = "Vch. Date"
    DataGrid1.Columns(6).Caption = "Vch. Name"
    DataGrid1.Columns(7).Caption = "Account Name"
    DataGrid1.Columns(8).Caption = "Quantity"
    DataGrid1.Columns(9).Caption = "Amount"
    DataGrid1.Columns(10).Caption = "Remark"
    DataGrid1.Columns(11).Caption = "Account Code"
'DataGrid1_Data_Field
    DataGrid1.Columns(0).DataField = "RowNo"
    DataGrid1.Columns(1).DataField = "VchCode"
    DataGrid1.Columns(2).DataField = "VchType"
    DataGrid1.Columns(3).DataField = "VchSeries"
    DataGrid1.Columns(4).DataField = "VchNo"
    DataGrid1.Columns(5).DataField = "VchDate"
    DataGrid1.Columns(6).DataField = "VchName"
    DataGrid1.Columns(7).DataField = "PartyName"
    DataGrid1.Columns(8).DataField = "Quantity"
    DataGrid1.Columns(9).DataField = "Amount"
    DataGrid1.Columns(10).DataField = "Remark"
    DataGrid1.Columns(11).DataField = "Account"
'DataGrid1_Data_Alinement
    DataGrid1.Columns(0).Alignment = dbgCenter
    DataGrid1.Columns(1).Alignment = dbgCenter
    DataGrid1.Columns(2).Alignment = dbgCenter
    DataGrid1.Columns(3).Alignment = dbgCenter
    DataGrid1.Columns(8).Alignment = dbgRight
    DataGrid1.Columns(9).Alignment = dbgRight
    DataGrid1.Columns(10).Alignment = dbgRight
    DataGrid1.Height = 7400
    MhRealInput1 = GetTotal()
    rstQueryList.Filter = adFilterNone
    If rstQueryList.RecordCount > 0 Then
        rstQueryList.MoveFirst
        If Not CheckEmpty(vtCode, False) Then rstQueryList.Find "[Code]='" & vtCode & "'"
    End If
    Set DataGrid1.DataSource = rstQueryList
    SortCol = "PartyName"
    BusySystemIndicator False
    If Not (rstQueryList.EOF Or rstQueryList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstQueryList.ActiveConnection = Nothing
    AutoMode = True
Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Check1_Click()
    If Check1.Value Then MhRealInput1 = GetTotal()
    Check1.Value = 0
End Sub
Private Sub ComboBox1_Click()
    Text1.Text = ""
    FocusSelect Me.ActiveControl
End Sub
Private Sub ComboBox2_Change()
If Not AutoMode Then Exit Sub
Me.Caption = ComboBox2.Text + " Query Ledger"
    Text1.Text = ""
    ComboBox2_Validate (True)
    Text1.SetFocus
End Sub
Private Function GetTotal() As Double
Dim TotalRows As Integer

On Error GoTo M
    ATotal = 0: QTotal = 0: TotalRows = 0
If rstQueryList.RecordCount > 0 Then
rstQueryList.MoveFirst
Do While Not rstQueryList.EOF
        QTotal = QTotal + Val(rstQueryList.Fields("Quantity").Value)
        ATotal = ATotal + Val(rstQueryList.Fields("Amount").Value)
rstQueryList.MoveNext
Loop
End If
    GetTotal = ATotal
    MhRealInput1 = ATotal
    MhRealInput2 = rstQueryList.RecordCount
    MhRealInput3 = QTotal
Exit Function
M:
'Exit Function
    MsgBox Err.Description
    GetTotal = ATotal
    MhRealInput1 = ATotal
    MhRealInput2 = rstQueryList.RecordCount
    MhRealInput3 = QTotal
End Function
Private Sub ComboBox1_Validate(Cancel As Boolean)
    If CheckEmpty(ComboBox2, True) Then
        Cancel = True
    Else
        ComboBox2_Validate (True)
    End If
End Sub
Private Sub ComboBox2_Validate(Cancel As Boolean)
    If CheckEmpty(ComboBox2, True) Then
        Cancel = True
    Else
        VchTypeUpdate
    End If
End Sub
Private Sub VchTypeUpdate() 'ComboBox2
If ComboBox2.ListIndex >= 0 And ComboBox2.ListIndex <= 7 Then
    VchCode = ComboBox2.ListIndex + 1
ElseIf ComboBox2.ListIndex >= 8 And ComboBox2.ListIndex <= 15 Then
    VchCode = ComboBox2.ListIndex + 9
ElseIf ComboBox2.ListIndex >= 16 And ComboBox2.ListIndex <= 21 Then
    VchCode = ComboBox2.ListIndex + 35
End If
    VchCode = Format(VchCode, "00")
        If VchCode >= 51 And VchCode <= 56 Then
            DataGrid1.Columns(8).Caption = "Debit"
            DataGrid1.Columns(9).Caption = "Credit"
            DataGrid1.Columns(8).NumberFormat = "#########0.00"
            MhRealInput3.Format = "#########0.00"
            MhRealInput3.DisplayFormat = "#########0.00"
        Else
            DataGrid1.Columns(8).Caption = "Quantity"
            DataGrid1.Columns(9).Caption = "Amount"
            DataGrid1.Columns(8).NumberFormat = "#########0"
            MhRealInput3.Format = "#########0"
            MhRealInput3.DisplayFormat = "#########0"
        End If
    If AutoMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
End Sub
Private Sub LoadMasterList()
Dim AC As String

    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
        rstCompanyMaster.Open "Select * FROM CompanyMaster Where FYCode='" & FYCode & "'", cnQuery, adOpenKeyset, adLockReadOnly
    rstCompanyMaster.ActiveConnection = Nothing
    
    AC = ComboBox1.Text: AC = "'" + Mid(Right(AC, 8), 2, 6) + "'"
    If AC = "''" Then
        AC = ""
    ElseIf AC = "'XXXXXX'" Then
        AC = ""
    ElseIf Len(AC) = 8 Then
        AC = " AND P.Party = " + AC
    Else
        AC = ""
    End If
    
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
        rstAccountList.Open "Select PrintName+' ['+Code+']' AS PrintName,PrintName AS Name,Code FROM AccountMaster Order By PrintName", cnQuery, adOpenKeyset, adLockPessimistic
    rstAccountList.ActiveConnection = Nothing
'End If
    
If (VchCode >= 1 And VchCode <= 8) Or (VchCode >= 17 And VchCode <= 24) Then
    If rstQueryList.State = adStateOpen Then rstQueryList.Close
        rstQueryList.Open "SELECT ROW_NUMBER() OVER (ORDER BY P.Date) AS RowNo,P.Code VchCode,P.Type VchType,V.Name As VchSeries,P.Name As VchNo,Format(Date,'dd-MM-yyyy') VchDate,V.VchName,(Select PrintName From AccountMaster Where Code=Party) AS PartyName,(Select ISNULL(SUM(Quantity),0) From JobworkBVChild C Where C.Code=P.Code) AS Quantity, ISNULL(Amount,0) AS Amount,P.Remarks AS Remark,Party AS Account From JobworkBVParent P LEFT JOIN VchSeriesMaster V ON V.Code=P.VchSeries Where Left(Type,2)= '" & VchCode & "' AND P.FYCode='" & FYCode & "' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "'  " & AC & "  Order By RowNo ", cnQuery, adOpenKeyset, adLockPessimistic
    rstQueryList.ActiveConnection = Nothing
ElseIf VchCode >= 51 And VchCode <= 56 Then
    If rstQueryList.State = adStateOpen Then rstQueryList.Close
        rstQueryList.Open "SELECT ROW_NUMBER() OVER (ORDER BY P.Date) AS RowNo,P.Code VchCode,P.Type VchType,V.Name As VchSeries,P.Name As VchNo,Format(Date,'dd-MM-yyyy') VchDate,V.VchName,(Select PrintName From AccountMaster Where Code=C.Account) AS PartyName,C.Debit AS Quantity,C.Credit AS Amount, P.LongNarration AS Remark,C.Account From DebitCreditParent P INNER JOIN DebitCreditChild C ON P.Code=C.code LEFT JOIN VchSeriesMaster V ON V.Code=P.VchSeries Where Left(Type,2)= '" & VchCode & "' AND P.FYCode='" & FYCode & "' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "'  " & AC & "  Order By Date ", cnQuery, adOpenKeyset, adLockPessimistic
        rstQueryList.ActiveConnection = Nothing
End If
If AutoMode Then MhRealInput1 = GetTotal(): Text1.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
VchType = "": vDate = "": vtType = "": vtCode = "": vtNo = "":
    If Shift = 0 And KeyCode = vbKeyEscape Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyE Then
        Export_Click
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
     ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then 'Refresh
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF And Toolbar1.Buttons.Item(1).Enabled Then 'First
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(13)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then 'Previous
        'Toolbar1_ButtonClick Toolbar1.Buttons.Item(14)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyN And Toolbar1.Buttons.Item(1).Enabled Then 'Next
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(15)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL And Toolbar1.Buttons.Item(1).Enabled Then 'Last
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(16)
        KeyCode = 0

'Open Transection
    ElseIf ((Shift = 0 And KeyCode = vbKeyReturn) Or (Shift = 0 And KeyCode = vbKeyF8) Or (Shift = 0 And KeyCode = vbKeyF12)) Then      'Open Transection
'Get vtCode,vtType,vtNo,vDate
If rstQueryList.EOF Then Exit Sub
            vDate = FixQuote(rstQueryList.Fields("VchDate").Value): vDate = Format(vDate, "dd-MMM-yyyy")
            vtCode = FixQuote(rstQueryList.Fields("VchCode").Value)
            vtType = FixQuote(rstQueryList.Fields("VchType").Value): vtType = Right(vtType, 2)
            vtNo = FixQuote(rstQueryList.Fields("VchNo").Value)
'ChecK vch FY
            If vDate = "" Then Exit Sub
            If FinancialYearFrom > vDate Or vDate = "" Then
                If MsgBox("You Can't Open Previous Financial Voucher in Current Year,... To Open This Voucher, Please Switch Financial Year ", vbCritical, "   Switch Financial Year !!!") = vbOK Then Exit Sub
'Order FG AND Jobwork
            ElseIf vtType = "FP" Or vtType = "FS" Then
            dSortBy = True
                    On Error Resume Next
                    FrmBookPrintOrder.BookPOType = vtType
                    If Err.Number <> 364 Then FrmBookPrintOrder.Show
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
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                End If
'Purchase Order,Sale Order,Stock Transfer
            ElseIf vtType = "PO" Or vtType = "SO" Or vtType = "ST" Then
            dSortBy = True
                    On Error Resume Next
                    frmSalesOrderVoucher.VchType = vtType
                    If Err.Number <> 364 Then frmSalesOrderVoucher.Show
                    frmSalesOrderVoucher.Text1 = vtCode
                If Shift = 0 And KeyCode = vbKeyReturn Then 'View
                    frmSalesOrderVoucher.SSTab1.Tab = 1
                ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'Delete
                    frmSalesOrderVoucher.Toolbar1_ButtonClick frmSalesOrderVoucher.Toolbar1.Buttons.Item(3)
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmSalesOrderVoucher.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
'                    Call cmdRefresh_Click
                End If
'Stock Journal Voucher
            ElseIf vtType = "JR" Then
            dSortBy = True
                    On Error Resume Next
                    frmStockJournalVoucher.VchType = vtType
                    If Err.Number <> 364 Then frmStockJournalVoucher.Show
                    frmStockJournalVoucher.Text1 = vtCode
                If Shift = 0 And KeyCode = vbKeyReturn Then 'View
                    frmStockJournalVoucher.SSTab1.Tab = 1
                ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'Delete
                    frmStockJournalVoucher.Toolbar1_ButtonClick frmStockJournalVoucher.Toolbar1.Buttons.Item(3)
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmStockJournalVoucher.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                End If
'Sale Voucher
            ElseIf vtType = "SF" Or vtType = "PF" Or vtType = "TF" Or vtType = "OF" Then
            dSortBy = True
                    On Error Resume Next
                    frmSalesVoucher.VchType = vtType
                    If Err.Number <> 364 Then frmSalesVoucher.Show
                    frmSalesVoucher.Text1 = vtCode
                If Shift = 0 And KeyCode = vbKeyReturn Then 'View
                    frmSalesVoucher.SSTab1.Tab = 1
                ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'Delete
                    frmSalesVoucher.Toolbar1_ButtonClick frmSalesVoucher.Toolbar1.Buttons.Item(3)
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmSalesVoucher.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                End If
'Sale Challan Voucher
            ElseIf vtType = "RF" Or vtType = "IF" Then
            dSortBy = True
                    On Error Resume Next
                    frmSalesChallanVoucher.VchType = vtType
                    If Err.Number <> 364 Then frmSalesChallanVoucher.Show
                    frmSalesChallanVoucher.Text1 = vtCode
                If Shift = 0 And KeyCode = vbKeyReturn Then 'View
                    frmSalesChallanVoucher.SSTab1.Tab = 1
                ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'Delete
                    frmSalesChallanVoucher.Toolbar1_ButtonClick frmSalesChallanVoucher.Toolbar1.Buttons.Item(3)
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmSalesChallanVoucher.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                End If
'Jobwork Sale Challan Voucher
            ElseIf vtType = "FR" Or vtType = "FI" Then
            vtType = IIf(vtType = "FR", "R", "I")
            dSortBy = True
                    On Error Resume Next
                    frmItemIssueReceiptVoucher.VchType = vtType
                    If Err.Number <> 364 Then frmItemIssueReceiptVoucher.Show
                    frmItemIssueReceiptVoucher.Text1 = vtCode
                If Shift = 0 And KeyCode = vbKeyReturn Then 'View
                    frmItemIssueReceiptVoucher.SSTab1.Tab = 1
                ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'Delete
                    frmItemIssueReceiptVoucher.Toolbar1_ButtonClick frmItemIssueReceiptVoucher.Toolbar1.Buttons.Item(3)
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmItemIssueReceiptVoucher.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                End If
'Jobwork Sale Voucher
            ElseIf vtType = "SU" Or vtType = "SC" Or vtType = "SJ" Or vtType = "PU" Or vtType = "PC" Or vtType = "PJ" Then
                vtType = IIf(vtType = "SU", 1, IIf(vtType = "SC", 2, IIf(vtType = "SJ", 3, IIf(vtType = "PU", 4, IIf(vtType = "PC", 5, IIf(vtType = "PJ", 6, ""))))))
                dSortBy = True
                    On Error Resume Next
                    frmJobworkBill.VchType = vtType
                    If Err.Number <> 364 Then frmJobworkBill.Show
                    frmJobworkBill.Text1 = vtCode
                If Shift = 0 And KeyCode = vbKeyReturn Then 'View
                    frmJobworkBill.SSTab1.Tab = 1
                ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'Delete
                    frmJobworkBill.Toolbar1_ButtonClick frmJobworkBill.Toolbar1.Buttons.Item(3)
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmJobworkBill.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                            Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
                End If
            End If
        KeyCode = 0
    End If
End Sub
Private Sub DataGrid1_DblClick()
On Error Resume Next
    If Toolbar1.Buttons.Item(17).Enabled Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(17)
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    SortCol = DataGrid1.Columns(ColIndex).DataField
    SortOrder = IIf(SortOrder = "Asc", "Desc", "Asc")
    rstQueryList.Sort = "[" + SortCol & "] " & SortOrder
    DataGrid1.ClearSelCols
    If Not (rstQueryList.EOF Or rstQueryList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub Text1_Change()
On Error Resume Next
    If rstQueryList.RecordCount = 0 Then Exit Sub
    rstQueryList.MoveFirst
    If Len(Text1.Text) > 0 Then
        rstQueryList.Filter = "[" & SortCol & "] Like '%" & FixQuote(Text1.Text) & "%'"
        If rstQueryList.EOF Then  'if Spelling mistake
            rstQueryList.Filter = adFilterNone
            rstQueryList.MoveFirst
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            Sendkeys "{End}"
        Else    'if Spelling alright
            PrevStr = Text1.Text
        End If
    Else
        rstQueryList.Filter = adFilterNone
        rstQueryList.MoveFirst
        MhRealInput1 = GetTotal()
        Set DataGrid1.DataSource = rstQueryList
        PrevStr = ""
    End If
    If Not (rstQueryList.EOF Or rstQueryList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    MhRealInput1 = 0
    MhRealInput3 = 0
    If Check1.Value Then MhRealInput1 = GetTotal()
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim KeyProcessed As Boolean
    If rstQueryList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstQueryList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstQueryList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstQueryList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstQueryList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstQueryList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstQueryList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstQueryList
            .MoveLast
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        With rstQueryList
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
Private Sub Form_Unload(Cancel As Integer)
    AutoMode = False
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstQueryList)
    Call CloseConnection(cnQuery)
End Sub
Private Sub MhDateInput1_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not ValidateDate(Me.ActiveControl) Then Cancel = True
End Sub
Private Sub MhDateInput2_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Not ValidateDate(Me.ActiveControl) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
        FocusSelect Me.ActiveControl
        Cancel = True
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        Load FrmDialog: Screen.MousePointer = vbNormal: FrmDialog.Flag = 6: FrmDialog.Caption = "Format":  FrmDialog.Command1.Caption = "Print Selection Voucher Format-1": FrmDialog.Command2.Caption = "Print Selection Voucher Format-2": FrmDialog.Command3.Caption = "Print Data Grid All Vouchers [Format-1] ": FrmDialog.Command4.Caption = "Print Data Grid All Vouchers [Format-2] ": FrmDialog.Command5.Visible = False: FrmDialog.Show vbModal
        If PtgType = 1 Then ComboBox3.ListIndex = 0: PrintFlag = False: Call PrintLabel
        If PtgType = 2 Then ComboBox3.ListIndex = 0: PrintFlag = False: Call PrintLabel
        If PtgType = 3 Then ComboBox3.ListIndex = 1: PrintFlag = False: Call PrintLabel: ComboBox3.ListIndex = 0
        If PtgType = 4 Then ComboBox3.ListIndex = 1: PrintFlag = False: Call PrintLabel: ComboBox3.ListIndex = 0
    ElseIf Button.Index = 2 Then
        Load FrmDialog: Screen.MousePointer = vbNormal: FrmDialog.Flag = 6: FrmDialog.Caption = "Format":  FrmDialog.Command1.Caption = "Print Selection Voucher Format-1": FrmDialog.Command2.Caption = "Print Selection Voucher Format-2": FrmDialog.Command3.Caption = "Print Data Grid All Vouchers [Format-1] ": FrmDialog.Command4.Caption = "Print Data Grid All Vouchers [Format-2] ": FrmDialog.Command5.Visible = False: FrmDialog.Show vbModal
        If PtgType = 1 Then ComboBox3.ListIndex = 0: PrintFlag = True: Call PrintLabel
        If PtgType = 2 Then ComboBox3.ListIndex = 0: PrintFlag = True: Call PrintLabel
        If PtgType = 3 Then ComboBox3.ListIndex = 1: PrintFlag = True: Call PrintLabel: ComboBox3.ListIndex = 0
        If PtgType = 4 Then ComboBox3.ListIndex = 1: PrintFlag = True: Call PrintLabel: ComboBox3.ListIndex = 0
    ElseIf Button.Index = 3 Then
        Unload Me
    ElseIf Button.Index = 6 Then 'Refresh Button
        Set DataGrid1.DataSource = Nothing
        LoadMasterList
        Set DataGrid1.DataSource = rstQueryList
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then 'Filter Button
        With FrmFilter
            .Combo1.AddItem "Name", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then 'First Record Button
        If rstQueryList.RecordCount > 0 Then rstQueryList.MoveFirst
        HiLiteRecord = True
    ElseIf Button.Index = 14 Then 'Previous Record Button
        If rstQueryList.RecordCount > 0 Then
           rstQueryList.MovePrevious
           If rstQueryList.BOF Then rstQueryList.MoveNext
        End If
        HiLiteRecord = True
    ElseIf Button.Index = 15 Then 'Next Record Button
        If rstQueryList.RecordCount > 0 Then
           rstQueryList.MoveNext
           If rstQueryList.EOF Then
              rstQueryList.MovePrevious
           End If
        End If
        HiLiteRecord = True
    ElseIf Button.Index = 16 Then 'Last Record Button
        If rstQueryList.RecordCount > 0 Then rstQueryList.MoveLast
        HiLiteRecord = True
    ElseIf Button.Index = 17 Then 'Open Record
    
    End If
'If AutoMode Then If Check1.Value Then MhRealInput1 = GetTotal(): Text1.SetFocus
End Sub
Private Sub CmdPrint_Click()
PrintFlag = True
On Error GoTo errHandler_print
Export_Click
    On Error GoTo 0
Exit Sub
errHandler_print:
  On Error GoTo 0
  Exit Sub
End Sub
Private Sub Export_Click()
Screen.MousePointer = vbHourglass
On Error Resume Next
Dim oExcel As Object
Dim oPdf As Object
Dim oBook As Object
Dim oSheet As Object
Dim j As Integer, i As Integer, Cnt As Long
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   Set oSheet = oBook.Worksheets(1)
   On Error GoTo errcode
   With oBook.Worksheets("sheet1").Rows(1)
        .Font.Bold = True
        .Font.Size = 16
        oBook.Worksheets("sheet1").Cells(1, j + 1).Value = Me.Caption
        .Range("A1:L1").Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        oBook.Worksheets("sheet1").Rows(2).Font.Bold = True
        For j = 0 To DataGrid1.Columns.Count - 1
            oBook.Worksheets("sheet1").Cells(2, j + 1).Value = DataGrid1.Columns(j).Caption
           MdiMainMenu.StatusBar1.Panels(2).Text = "Processed record #" & Trim(Str(Cnt)) & " of " & Trim(Str(rstQueryList.RecordCount)) & " !!!"
        Cnt = Cnt + 1
        Next j
    'Next i
   End With
 'Save Data
   oSheet.Range("A3").CopyFromRecordset rstQueryList
     
   With oExcel
            StartColumn = "A"
            StartRow = 1
            EndColumn = "L"
            EndRow = rstQueryList.RecordCount + 1
            oBook.Activate
            If PrintFlag = False Then oExcel.Visible = True
        For i = 0 To DataGrid1.Columns.Count - 1
                    oBook.Worksheets("sheet1").Cells(2, i + 1).Value = DataGrid1.Columns(i).Caption
        Next
            
            .Columns("A:Z").EntireColumn.AutoFit
            .ActiveSheet.pagesetup.Orientation = xlLandscape
            .ActiveSheet.pagesetup.LeftMargin = .InchesToPoints(0.36)
            .ActiveSheet.pagesetup.RightMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.TopMargin = .InchesToPoints(0.5)
            .ActiveSheet.pagesetup.BottomMargin = .InchesToPoints(0.5)
            .ActiveSheet.pagesetup.HeaderMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.FooterMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.PrintArea = StartColumn & StartRow & ":" & EndColumn & EndRow + 1
            .ActiveSheet.pagesetup.Zoom = False
            .ActiveSheet.pagesetup.FitToPagesTall = False
            .ActiveSheet.pagesetup.FitToPagesWide = 1
            .ActiveSheet.pagesetup.PrintGridlines = True
            .ActiveSheet.pagesetup.PrintTitleRows = "$1:$2"
            
            If PrintFlag Then
            .ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF
           Screen.MousePointer = vbDefault
                With CommonDialog1
                    .Copies = 1
                    .flags = &H0&
                    .ShowPrinter
                    oExcel.ActiveSheet.PrintOut
                End With
            End If
   End With
   PrintFlag = False
   Set oBook = Nothing
   Set oSheet = Nothing
   Set oExcel = Nothing
   Screen.MousePointer = vbDefault
   Exit Sub
errcode:
   MsgBox Err.Description, , Err.Source
   PrintFlag = False
   Set oBook = Nothing
   Set oSheet = Nothing
   Set oExcel = Nothing
   Screen.MousePointer = vbDefault
End Sub
Private Sub CmdLabel_Click()
On Error Resume Next
    PrintFlag = True
    PtgType = 2
    Call PrintLabel
End Sub
Public Sub PrintLabel()
   On Error GoTo errcode
    Screen.MousePointer = vbHourglass
    Dim oExcel As Object
    Dim oPdf As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim j, R As Integer, i As Integer, Cnt As Long
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    
    With oBook.Worksheets("sheet1").Columns(1)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With oBook.Worksheets("sheet1").Columns(4)
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    With oBook.Worksheets("sheet1").Columns(6)
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    With oBook.Worksheets("sheet1").Columns(7)
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
   With oExcel
        oBook.Activate
    If PrintFlag = False Then oExcel.Visible = True
    End With
    If ComboBox3.ListIndex = 1 Then rstQueryList.MoveFirst
    Do While Not rstQueryList.EOF
    With rstQueryDetails
        If .State = adStateOpen Then .Close
        .Open "SELECT P.Code VchCode,P.Type VchType,V.Name As VchSeries,P.Name As VchNo,Format(Date,'dd-MM-yyyy') VchDate,V.VchName,A.PrintName AS AccountName,A.Name AS AccountAlias,A.Code As AccountCode,A.Address1+' '+A.Address2+' '+A.Address3+' '+A.City+IIF(A.Address4<>'',' -'+A.Address4,'')+' '+IIF(State='*56000','',IIF(State='','',',State :'+(Select Name From GeneralMaster Where Code=State)))+' '+IIF(Station<>'',',Station :'+Station,'')+IIF(Mobile<>'',' Mobile : '+Mobile,'') AS Address," & _
                    "I.Name AS ItemCode,I.PrintName AS ItemName,(Select PrintName From GeneralMaster Where Code=I.IntegrationUnit) As Unit,Rate,ABS(C.Quantity) AS Quantity,ISNULL(C.Amount,0) AS Amount,P.Remarks AS Remark,P.Transport,C1.*  From JobworkBVParent P INNER JOIN JobworkBVChild C ON C.Code=P.Code Left Join AccountMaster A On A.Code=P.Party Left JOIN BookMaster I ON C.Item=I.Code Left JOIN VchSeriesMaster V ON V.Code=P.VchSeries LEFT JOIN CompChild C1 ON C1.VchType=Left(P.Type,2) WHERE P.Code='" & FixQuote(rstQueryList.Fields("VchCode").Value) & "'  Order By P.Code ", cnQuery, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End With
    
'Loop Start
'Header
    i = i + 1
    With oBook.Worksheets("sheet1").Rows(i)
        oBook.Worksheets("sheet1").Cells(i, 1).Value = rstQueryDetails.Fields("VchName").Value & " Voucher"
    End With
    With oSheet.Range("A" & i & ":H" & i)
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 16
        .Font.Underline = xlUnderlineStyleDouble ' Use double underline
    End With

'Vch. Info
If PtgType = 1 Or PtgType = 3 Then
   i = i + 1: j = 0
   With oBook.Worksheets("sheet1").Rows(i)
        j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Vch. No."
        j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = rstQueryDetails.Fields("VchNo").Value
        j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Party : [" + rstQueryDetails.Fields("AccountAlias").Value + "]" + rstQueryDetails.Fields("AccountName").Value
        j = j + 3: oBook.Worksheets("sheet1").Cells(i, j).Value = "Date:"
        j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = rstQueryDetails.Fields("VchDate").Value
   End With
    With oSheet.Range("G" & i & ":H" & i)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "dd-MM-yyyy"
    End With
   
ElseIf PtgType = 2 Or PtgType = 4 Then
   i = i + 1: j = 0
   With oBook.Worksheets("sheet1").Rows(i)
         j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Party :" + " [" + rstQueryDetails.Fields("AccountAlias").Value + "]" + rstQueryDetails.Fields("AccountName").Value
         j = j + 1: 'oBook.Worksheets("sheet1").Cells(i, j).Value = ""
         j = j + 3: oBook.Worksheets("sheet1").Cells(i, j).Value = "Date :" + rstQueryDetails.Fields("VchDate").Value
         j = j + 1: 'oBook.Worksheets("sheet1").Cells(i, j).Value = ""
   End With
    With oSheet.Range("A" & i & ":C" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    With oSheet.Range("E" & i & ":G" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
   i = i + 1: j = 0
   With oBook.Worksheets("sheet1").Rows(i)
         j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Address :" + rstQueryDetails.Fields("Address").Value
         j = j + 1: ' oBook.Worksheets("sheet1").Cells(i, j).Value = ""
         j = j + 3: oBook.Worksheets("sheet1").Cells(i, j).Value = "Vch. No. :" + rstQueryDetails.Fields("VchNo").Value
         j = j + 1: 'oBook.Worksheets("sheet1").Cells(i, j).Value = ""
   End With
    With oSheet.Range("A" & i & ":C" & i + 1)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    With oSheet.Range("E" & i & ":H" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
   
   i = i + 1: j = 0
   With oBook.Worksheets("sheet1").Rows(i)
         j = j + 1: 'oBook.Worksheets("sheet1").Cells(i, j).Value = ""
         j = j + 1: 'oBook.Worksheets("sheet1").Cells(i, j).Value = ""
         j = j + 3: oBook.Worksheets("sheet1").Cells(i, j).Value = "Trp.: " + rstQueryDetails.Fields("Transport").Value
         j = j + 1: 'oBook.Worksheets("sheet1").Cells(i, j).Value = ""
   End With
'    With oSheet.Range("A" & i & ":C" & i)
'        .Merge
'        .WrapText = True
'        .HorizontalAlignment = xlLeft
'        .VerticalAlignment = xlCenter
'        .NumberFormat = "General"
'    End With
    With oSheet.Range("E" & i & ":H" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
End If

'Details Header
    
    i = i + 1: j = 0
   With oBook.Worksheets("sheet1").Rows(i)
         j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "S. No."
         j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Item Alias"
         j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Item Description"
         j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Qty."
         j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Unit."
         j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Rate."
         j = j + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Amount"
   End With
' Details Header Formatting
    With oSheet.Range("A" & i & ":H" & i).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oSheet.Range("A" & i & ":H" & i).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
' Details
      With oExcel
        Dim QSum As Double, ASum As Double
        Cnt = 0:
        rstQueryDetails.MoveFirst
    Do While Not rstQueryDetails.EOF
            i = i + 1: Cnt = Cnt + 1
             oBook.Worksheets("sheet1").Cells(i, 1).Value = Cnt
            If rstQueryDetails.Fields("ItemCode").Value <> "" Then oBook.Worksheets("sheet1").Cells(i, 2).Value = rstQueryDetails.Fields("ItemCode").Value
            If rstQueryDetails.Fields("ItemName").Value <> "" Then oBook.Worksheets("sheet1").Cells(i, 3).Value = rstQueryDetails.Fields("ItemName").Value
            If rstQueryDetails.Fields("Quantity").Value <> "" Then oBook.Worksheets("sheet1").Cells(i, 4).Value = Val(Format(rstQueryDetails.Fields("Quantity").Value, "###0.00")): QSum = QSum + Val(rstQueryDetails.Fields("Quantity").Value)
            If rstQueryDetails.Fields("Unit").Value <> "" Then oBook.Worksheets("sheet1").Cells(i, 5).Value = rstQueryDetails.Fields("Unit").Value
            If rstQueryDetails.Fields("Rate").Value <> "" Then oBook.Worksheets("sheet1").Cells(i, 6).Value = Val(Format(rstQueryDetails.Fields("Rate").Value, "###0.00"))
            If rstQueryDetails.Fields("Amount").Value <> "" Then oBook.Worksheets("sheet1").Cells(i, 7).Value = Val(Format(rstQueryDetails.Fields("Amount").Value, "###0.00")): ASum = ASum + Val(rstQueryDetails.Fields("Amount").Value)
        rstQueryDetails.MoveNext
    Loop
            i = i + 1
            oBook.Worksheets("sheet1").Cells(i, 2).Value = "Total :"
            oBook.Worksheets("sheet1").Cells(i, 4).Value = Format(QSum, "###0.00")
            oBook.Worksheets("sheet1").Cells(i, 7).Value = Format(ASum, "###0.00")
    End With
   With oSheet.Range("A" & i & ":H" & i).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oSheet.Range("A" & i & ":H" & i).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    rstQueryDetails.MoveFirst
    
    
    i = i + 1:
    With oBook.Worksheets("sheet1").Rows(i)
        oBook.Worksheets("sheet1").Cells(i, 1).Value = rstQueryDetails.Fields("Declaration01").Value
        oBook.Worksheets("sheet1").Cells(i, 5).Value = "Bank :" + rstCompanyMaster.Fields("BankName").Value
    End With
    With oSheet.Range("A" & i & ":C" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    With oSheet.Range("E" & i & ":H" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    
    i = i + 1:
    With oBook.Worksheets("sheet1").Rows(i)
        oBook.Worksheets("sheet1").Cells(i, 1).Value = rstQueryDetails.Fields("Declaration02").Value
        oBook.Worksheets("sheet1").Cells(i, 5).Value = "Account No :" + rstCompanyMaster.Fields("AccountNo").Value
    End With
    With oSheet.Range("A" & i & ":C" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    With oSheet.Range("E" & i & ":H" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    
    i = i + 1:
    With oBook.Worksheets("sheet1").Rows(i)
        oBook.Worksheets("sheet1").Cells(i, 1).Value = rstQueryDetails.Fields("Declaration03").Value
        oBook.Worksheets("sheet1").Cells(i, 5).Value = "IFSC CODE :" + rstCompanyMaster.Fields("IFSC").Value
    End With
    With oSheet.Range("A" & i & ":C" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    With oSheet.Range("E" & i & ":H" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    
    i = i + 1:
    With oBook.Worksheets("sheet1").Rows(i)
        oBook.Worksheets("sheet1").Cells(i, 1).Value = rstQueryDetails.Fields("Declaration04").Value
    End With
    With oSheet.Range("A" & i & ":C" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    i = i + 1:
    With oBook.Worksheets("sheet1").Rows(i)
        oBook.Worksheets("sheet1").Cells(i, 1).Value = rstQueryDetails.Fields("Declaration05").Value
    End With
    With oSheet.Range("A" & i & ":C" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    i = i + 1:
    With oBook.Worksheets("sheet1").Rows(i)
        oBook.Worksheets("sheet1").Cells(i, 1).Value = rstQueryDetails.Fields("Declaration06").Value
    End With
    With oSheet.Range("A" & i & ":C" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    i = i + 1:
    With oBook.Worksheets("sheet1").Rows(i)
        oBook.Worksheets("sheet1").Cells(i, 1).Value = rstQueryDetails.Fields("Declaration07").Value
    End With
    With oSheet.Range("A" & i & ":C" & i)
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "General"
    End With
    
    i = i + 1
If ComboBox3.ListIndex = 0 Then Exit Do
rstQueryList.MoveNext
Loop

   'Loop End
'*****************
   With oExcel
            StartColumn = "A"
            StartRow = 1
            EndColumn = "H"
            EndRow = rstQueryList.RecordCount + 4
    '        oBook.Activate
    '        oExcel.Visible = True
            .Columns("A:Z").EntireColumn.AutoFit
            .ActiveSheet.pagesetup.Orientation = xlPortrait
            .ActiveSheet.pagesetup.LeftMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.RightMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.TopMargin = .InchesToPoints(0.5)
            .ActiveSheet.pagesetup.BottomMargin = .InchesToPoints(0)
            .ActiveSheet.pagesetup.HeaderMargin = .InchesToPoints(0)
            .ActiveSheet.pagesetup.FooterMargin = .InchesToPoints(0)
            StartColumn = "A"
            StartRow = 1
            EndColumn = "H"
            If i <> 0 Then EndRow = i
            .ActiveSheet.pagesetup.PrintArea = StartColumn & StartRow & ":" & EndColumn & EndRow
            .ActiveSheet.pagesetup.Zoom = False
            .ActiveSheet.pagesetup.FitToPagesTall = 1000
            .ActiveSheet.pagesetup.FitToPagesWide = 1
            .ActiveSheet.pagesetup.PrintGridlines = False
            .ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF
            Screen.MousePointer = vbDefault
            If PrintFlag Then
                With CommonDialog1
                    .Copies = 1
                    .flags = &H0&
                    .ShowPrinter
                    oExcel.ActiveSheet.PrintOut
                End With
            End If
   End With
   PtgType = 0
   PrintFlag = False
   Set oBook = Nothing
   Set oSheet = Nothing
   Set oExcel = Nothing
   Screen.MousePointer = vbDefault
   Exit Sub
errcode:
   MsgBox Err.Description, , Err.Source
   PtgType = 0
   PrintFlag = False
   Set oBook = Nothing
   Set oSheet = Nothing
   Set oExcel = Nothing
   Screen.MousePointer = vbDefault
    Call CloseRecordset(rstQueryDetails)
    On Error GoTo 0
End Sub
