VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAccountLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Account Ledger"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19245
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
   LinkTopic       =   "FrmLogin"
   MaxButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   19245
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   9255
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   19245
      _Version        =   65536
      _ExtentX        =   33946
      _ExtentY        =   16325
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
      Picture         =   "AccountLedger.frx":0000
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   13800
         Top             =   1560
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
               Picture         =   "AccountLedger.frx":001C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AccountLedger.frx":0560
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AccountLedger.frx":0674
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AccountLedger.frx":0786
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Preview 
         Caption         =   "&Print Preview"
         Height          =   330
         Left            =   15600
         TabIndex        =   20
         Top             =   8840
         Width           =   1215
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   8370
         Width           =   8535
         _Version        =   65536
         _ExtentX        =   15055
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
         Caption         =   "Ctrl+F->Search  F8->Delete  F9->Hide  Escap->Un-Hide  F12->Duplicate  F5->Refresh"
         FillColor       =   8421504
         TextColor       =   16777215
         Picture         =   "AccountLedger.frx":0898
         Picture         =   "AccountLedger.frx":08B4
      End
      Begin VB.CommandButton Command2 
         Height          =   320
         Left            =   6000
         Picture         =   "AccountLedger.frx":08D0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Search"
         Top             =   8840
         Width           =   375
      End
      Begin VB.CommandButton cmdFilter 
         Height          =   320
         Left            =   5520
         Picture         =   "AccountLedger.frx":0C12
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Filter"
         Top             =   8840
         Width           =   375
      End
      Begin VB.TextBox Text1 
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
         Left            =   3240
         MaxLength       =   40
         TabIndex        =   12
         ToolTipText     =   "Find And Search"
         Top             =   8840
         Width           =   2190
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   7305
         Left            =   120
         TabIndex        =   0
         Top             =   960
         Width           =   19050
         _Version        =   524288
         _ExtentX        =   33602
         _ExtentY        =   12885
         _StockProps     =   64
         ColsFrozen      =   3
         EditEnterAction =   2
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
         MaxCols         =   13
         MaxRows         =   2000
         SelectBlockOptions=   4
         SpreadDesigner  =   "AccountLedger.frx":0F54
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   620
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
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
         Caption         =   " &From"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountLedger.frx":1AB7
         Picture         =   "AccountLedger.frx":1AD3
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   1800
         TabIndex        =   6
         Top             =   620
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
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
         Caption         =   " &To"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountLedger.frx":1AEF
         Picture         =   "AccountLedger.frx":1B0B
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   2190
         TabIndex        =   2
         Top             =   620
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "AccountLedger.frx":1B27
         Caption         =   "AccountLedger.frx":1C3F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountLedger.frx":1CAB
         Keys            =   "AccountLedger.frx":1CC9
         Spin            =   "AccountLedger.frx":1D27
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
         ForeColor       =   255
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
         Left            =   720
         TabIndex        =   1
         Top             =   620
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "AccountLedger.frx":1D4F
         Caption         =   "AccountLedger.frx":1E67
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountLedger.frx":1ED3
         Keys            =   "AccountLedger.frx":1EF1
         Spin            =   "AccountLedger.frx":1F4F
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
         ForeColor       =   255
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   13230
         TabIndex        =   7
         Top             =   105
         Visible         =   0   'False
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
         Caption         =   " &Sort && Filter"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountLedger.frx":1F77
         Picture         =   "AccountLedger.frx":1F93
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   330
         Left            =   1200
         TabIndex        =   8
         Top             =   8840
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountLedger.frx":1FAF
         Caption         =   "AccountLedger.frx":1FCF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountLedger.frx":2033
         Keys            =   "AccountLedger.frx":2051
         Spin            =   "AccountLedger.frx":209B
         AlignHorizontal =   2
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   255
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   332922885
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   8840
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
         Caption         =   " Data Count"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountLedger.frx":20C3
         Picture         =   "AccountLedger.frx":20DF
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   18090
         TabIndex        =   10
         Top             =   8840
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
         Caption         =   " Print Data"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountLedger.frx":20FB
         Picture         =   "AccountLedger.frx":2117
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   16920
         TabIndex        =   11
         Top             =   8840
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
         Caption         =   " Export Data"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountLedger.frx":2133
         Picture         =   "AccountLedger.frx":214F
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   2520
         TabIndex        =   13
         Top             =   8840
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
         Caption         =   " Find"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountLedger.frx":216B
         Picture         =   "AccountLedger.frx":2187
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   14295
         TabIndex        =   16
         Top             =   8840
         Visible         =   0   'False
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
         Caption         =   "Import Data"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountLedger.frx":21A3
         Picture         =   "AccountLedger.frx":21BF
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   13320
         TabIndex        =   19
         Top             =   600
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial MT Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   "Opening Bal. = Rs. 0.00"
         Alignment       =   1
         BorderStyle     =   0
         TextColor       =   0
         Picture         =   "AccountLedger.frx":21DB
         Picture         =   "AccountLedger.frx":21F7
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   13320
         TabIndex        =   21
         Top             =   8370
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial MT Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   "Opening Bal. = Rs. 0.00"
         Alignment       =   1
         BorderStyle     =   0
         TextColor       =   0
         Picture         =   "AccountLedger.frx":2213
         Picture         =   "AccountLedger.frx":222F
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   7695
         _Version        =   65536
         _ExtentX        =   13573
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial MT Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   "Accounts :"
         Alignment       =   0
         BorderStyle     =   0
         TextColor       =   0
         Picture         =   "AccountLedger.frx":224B
         Picture         =   "AccountLedger.frx":2267
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   16800
         TabIndex        =   23
         Top             =   120
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Print Preview"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Print"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Mail"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Exit"
               ImageIndex      =   4
            EndProperty
         EndProperty
         Begin VB.CommandButton Command1 
            Height          =   375
            Left            =   1440
            Picture         =   "AccountLedger.frx":2283
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Refresh"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdRefresh 
            Height          =   375
            Left            =   1440
            Picture         =   "AccountLedger.frx":23CD
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Refresh"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdCancel 
            Height          =   375
            Left            =   1920
            Picture         =   "AccountLedger.frx":2517
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Cancel"
            Top             =   0
            Width           =   375
         End
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   6735
         TabIndex        =   27
         Top             =   600
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial MT Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   "Report Header"
         BorderStyle     =   0
         TextColor       =   0
         Picture         =   "AccountLedger.frx":2619
         Picture         =   "AccountLedger.frx":2635
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   19300
         Y1              =   8760
         Y2              =   8760
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   6480
         TabIndex        =   15
         Top             =   8840
         Width           =   2205
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3889;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Left            =   14430
         TabIndex        =   3
         Top             =   105
         Visible         =   0   'False
         Width           =   2325
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4101;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   19300
         Y1              =   540
         Y2              =   540
      End
   End
End
Attribute VB_Name = "FrmAccountLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dSortBy As Boolean
Public sDate As String, eDate As String, AccountGroupList As String, AccountList As String, VchType As String, Header1 As String, vDate, SCode As Variant, LR As Integer, R As Long, OutputTo As String
Dim rstAccountLedger As New ADODB.Recordset, rstAccountOpening As New ADODB.Recordset, rstCompanyMaster As New ADODB.Recordset, Reset As Long
Dim Debit As Double, Credit As Double, Bal As Variant, DebitTotal As Double, CreditTotal As Double, BalTotal As Double, Code As Variant, TotalFlag As Boolean, HideFlag As Boolean, ExitFlag As Boolean
Dim Opening As Double
Dim oOutlook As New Outlook.Application
Dim EMailID As String, Attachment As String, Message As String
Private Sub Form_Load()
Reset = 0:
    On Error GoTo ErrorHandler
    CenterForm Me
    Me.Top = 1200
    BusySystemIndicator True
    Dim Cols As Long, C As Long
        fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader
        fpSpread1.UserColAction = UserColActionSort
        Cols = fpSpread1.MaxCols
        For C = 1 To Cols
        fpSpread1.ColUserSortIndicator(C) = ColUserSortIndicatorDescending
    Next
    If VchType >= 0 Then
        Combo1.AddItem "Date", 0
        Combo1.AddItem "Type", 1
        Combo1.AddItem "Vch/Bill No.", 2
        Combo1.AddItem "Accounts", 3
        Combo1.AddItem "Debit", 4
        Combo1.AddItem "Credit", 5
        Combo1.AddItem "Balance", 6
        Combo1.AddItem "Short Narration", 7
        Combo1.AddItem "Long Narration", 8
        Combo1.ListIndex = 0

        Combo2.AddItem "Date ", 0
        Combo2.AddItem "Type ", 1
        Combo2.AddItem "Vch/Bill No", 2
        Combo2.AddItem "Accounts", 3
        Combo2.AddItem "Debit", 4
        Combo2.AddItem "Credit", 5
        Combo2.AddItem "Balance", 6
        Combo2.AddItem "Short Narration", 7
        Combo2.AddItem "Long Narration", 8
        Combo2.ListIndex = 0
    End If
    Reset = 1
    If VchType = 2 Then Me.Caption = " Accounts Ledger"
    MhDateInput1.Value = Format(sDate, "dd-MM-yyyy")
    MhDateInput2.Value = Format(eDate, "dd-MM-yyyy")
    cmdRefresh_Click
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TypeCode As Variant
SCode = "": TypeCode = ""
    With fpSpread1
        .GetText 11, .ActiveRow, TypeCode: TypeCode = Right(TypeCode, 2)
    End With
    'Sale
   If (Shift = 0 And KeyCode = vbKeyReturn Or Shift = 0 And KeyCode = vbKeyF8 Or Shift = 0 And KeyCode = vbKeyF12) And (TypeCode = "SF" Or TypeCode = "TF" Or TypeCode = "PF" Or TypeCode = "OF") Then
                With fpSpread1
                    .GetText 12, .ActiveRow, SCode
                End With
        If SCode = "" Then Exit Sub
        'FY Check
        fpSpread1.GetText 1, fpSpread1.ActiveRow, vDate: vDate = Format(vDate, "dd-MMM-yyyy"):
        If vDate = "" Then
            Exit Sub
        ElseIf FinancialYearFrom > vDate Or vDate = "" Then
            If MsgBox("You Can't Open Previous Financial Voucher in Current Year,... To Open This Voucher, Please Switch Financial Year ", vbCritical, "   Switch Financial Year !!!") = vbOK Then Exit Sub
        Else
            On Error Resume Next
            dSortBy = True
            frmSalesVoucher.VchType = TypeCode
            If Err.Number <> 364 Then frmSalesVoucher.Show
            frmSalesVoucher.Text1 = SCode
        End If
        If Shift = 0 And KeyCode = vbKeyReturn Then 'View
            frmSalesVoucher.SSTab1.Tab = 1
        ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'Delete
            frmSalesVoucher.Toolbar1_ButtonClick frmSalesVoucher.Toolbar1.Buttons.Item(3)
            Call cmdRefresh_Click
        ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
            frmSalesVoucher.SSTab1.Tab = 1
            If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
            Call cmdRefresh_Click
        End If
        KeyCode = 0
   'Refresh
   ElseIf Shift = 0 And KeyCode = vbKeyF5 Then 'Refresh
        Call cmdRefresh_Click
   'Jobwork Sale
   ElseIf (Shift = 0 And KeyCode = vbKeyReturn Or Shift = 0 And KeyCode = vbKeyF8 Or Shift = 0 And KeyCode = vbKeyF12) And (TypeCode = "SU" Or TypeCode = "SJ" Or TypeCode = "SC" Or TypeCode = "PU" Or TypeCode = "PJ" Or TypeCode = "PC") Then
                With fpSpread1
                    .GetText 12, .ActiveRow, SCode
                End With
        If SCode = "" Then Exit Sub
        'FY Check
        fpSpread1.GetText 1, fpSpread1.ActiveRow, vDate: vDate = Format(vDate, "dd-MMM-yyyy"):
        If vDate = "" Then
            Exit Sub
        ElseIf FinancialYearFrom > vDate Or vDate = "" Then
            If MsgBox("You Can't Open Previous Financial Voucher in Current Year,... To Open This Voucher, Please Switch Financial Year ", vbCritical, "   Switch Financial Year !!!") = vbOK Then Exit Sub
        Else
            On Error Resume Next
            TypeCode = IIf(TypeCode = "SU", 1, IIf(TypeCode = "SC", 2, IIf(TypeCode = "SJ", 3, IIf(TypeCode = "PU", 4, IIf(TypeCode = "PC", 5, IIf(TypeCode = "PJ", 6, ""))))))
            frmJobworkBill.VchType = TypeCode
            dSortBy = True
            If Err.Number <> 364 Then frmJobworkBill.Show
            frmJobworkBill.Text1 = SCode
        End If
        If Shift = 0 And KeyCode = vbKeyReturn Then 'View
            frmJobworkBill.SSTab1.Tab = 1
        ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'Delete
            frmJobworkBill.Toolbar1_ButtonClick frmJobworkBill.Toolbar1.Buttons.Item(3)
            Call cmdRefresh_Click
        ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate Record
            If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
            Call cmdRefresh_Click
        End If
            KeyCode = 0
    'Debit / Credit
   ElseIf (Shift = 0 And KeyCode = vbKeyReturn Or Shift = 0 And KeyCode = vbKeyF8 Or Shift = 0 And KeyCode = vbKeyF12) And (TypeCode = "PI" Or TypeCode = "PR" Or TypeCode = "JE" Or TypeCode = "CE" Or TypeCode = "DN" Or TypeCode = "CN") Then
                With fpSpread1
                    .GetText 3, .ActiveRow, SCode
                End With
        'FY Check
        fpSpread1.GetText 1, fpSpread1.ActiveRow, vDate: vDate = Format(vDate, "dd-MMM-yyyy"):
        If vDate = "" Then
            Exit Sub
        ElseIf FinancialYearFrom > vDate Or vDate = "" Then
            If MsgBox("You Can't Open Previous Financial Voucher in Current Year,... To Open This Voucher, Please Switch Financial Year ", vbCritical, "   Switch Financial Year !!!") = vbOK Then Exit Sub
        Else
            On Error Resume Next
            frmDebitCreditVoucher.VchType = TypeCode
            If Err.Number <> 364 Then frmDebitCreditVoucher.Show
            frmDebitCreditVoucher.Text1 = SCode
        End If
        If Shift = 0 And KeyCode = vbKeyReturn Then 'View
            frmDebitCreditVoucher.SSTab1.Tab = 1
        ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'Delete
            frmDebitCreditVoucher.Toolbar1_ButtonClick frmDebitCreditVoucher.Toolbar1.Buttons.Item(3)
            Call cmdRefresh_Click
        ElseIf Shift = 0 And KeyCode = vbKeyF12 Then ' Duplicate
            If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
            Call cmdRefresh_Click
            KeyCode = 0
        End If
   ElseIf Shift = 0 And KeyCode = vbKeyReturn Then 'Open
        If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}": KeyCode = 0
   ElseIf Shift = 0 And KeyCode = vbKeyEscape Then ' Close/Hide Row/Unhide Row
     With fpSpread1
        If HideFlag = True Then
            For R = 1 To .DataRowCnt 'Unhide All
                If HideFlag = True Then .Row = R: .RowHidden = False: .SetText 13, R, "":
            Next
            Total_Click
            .SetActiveCell .ActiveCol, 1
            HideFlag = False
        ElseIf HideFlag = False And ExitFlag = False Then
            Call cmdCancel_Click: ExitFlag = False
        End If
            KeyCode = 0
     End With
   ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
            If Text1.Text = "" Then
                MsgBox "Please Provide Search Input", vbInformation
                Text1.SetFocus
            ElseIf Text1.Text <> "" Then
            Call Command2_Click
            End If
        KeyCode = 0
   ElseIf Shift = 0 And KeyCode = vbKeyF9 Then
        With fpSpread1
            R = IIf(.ActiveRow + 1 <> LR, .ActiveRow + 1, 1)
            LR = R
             .SetText 13, .ActiveRow, "True": .Row = .ActiveRow: .RowHidden = True: LR = .Row
            TotalFlag = True: HideFlag = True: Total_Click
            TotalFlag = False
            .SetActiveCell .ActiveCol, R
        End With
        KeyCode = 0
   End If
End Sub
Private Sub cmdRefresh_Click()
    On Error GoTo ErrHandler
    Dim SQL As String, OpSQL As String, mSQL, sSQL, dSQL As String  '[SQL Query,Opening SQL Query,Month SQL Query,Summary SQL Query,Details SQL Query]
    Dim i As Long, R As Long, C As Long, n As Integer
    Debit = 0: Credit = 0: Bal = 0
    CreditTotal = 0: DebitTotal = 0:
    If VchType >= 0 And VchType <= 29 Then 'Account Ledger
        OpSQL = "SELECT (SELECT ISNULL(Sum(C.Credit-C.Debit),0) As Opening FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date < '" & GetDate(MhDateInput1.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account IN (" & AccountList & ")) + " & _
                       "(SELECT ISNULL(Sum(IIF(LEFT(Type,2)='01',(T.Amount),IIF(LEFT(Type,2)='02',(0-T.Amount),IIF(LEFT(Type,2)='03',(T.Amount),IIF(LEFT(Type,2)='04',(0-T.Amount),0))))),0) As Opening FROM JobworkBVParent T LEFT JOIN AccountMaster A On T.Party=A.Code Where T.Date < '" & GetDate(MhDateInput1.Text) & "' And T.Party IN (" & AccountList & ")) +  " & _
                       "(Select ISNULL(Sum(Opening),0) From AccountMaster Where Code IN (" & AccountList & ")) As Opening,(SELECT PRINTNAME FROM AccountMaster Where Code IN (" & AccountList & ")) AS AccountNAME  "
        Screen.MousePointer = vbHourglass
        If rstAccountOpening.State = adStateOpen Then rstAccountOpening.Close
        rstAccountOpening.Open OpSQL, cnDatabase, adOpenKeyset, adLockReadOnly
        If rstAccountOpening.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
        
        mSQL = "WITH Months AS (SELECT TOP 12 CASE WHEN ROW_NUMBER()OVER (ORDER BY (SELECT NULL )) <= 3 THEN ROW_NUMBER() OVER (ORDER BY (SELECT NULL )) +12 Else ROW_NUMBER()OVER (ORDER BY (SELECT NULL )) END AS mCode FROM master.dbo.spt_values) "
        mSQL = mSQL + "SELECT '' Date,'' VchType,'' VchBillNo,FORMAT(DATEADD(month, m.mCode - 4, '" & FinancialYearFrom & "'),'MMMM') MonthYear,ISNULL(SUM(TBL.Debit), 0) AS Debit,ISNULL(SUM(TBL.Credit), 0) AS Credit,'' ShortNarration,'' LongNarration,''Type,''Code,ISNULL(TBL.AccountName,'') As AccountName,CASE WHEN m.mCode <= 12 THEN FORMAT(DATEADD(month, m.mCode - 4, DATEADD(YEAR,-1,'" & FinancialYearFrom & "')), 'dd-MMM-yyyy') Else FORMAT(DATEADD(month, m.mCode - 4, '" & FinancialYearFrom & "'), 'dd-MMM-yyyy') END AS FromDate,CASE WHEN m.mCode <= 12 THEN   FORMAT(DATEADD(Day, -1, DATEADD(month, m.mCode - 3, DATEADD(YEAR,-1,'" & FinancialYearFrom & "'))), 'dd-MMM-yyyy') Else FORMAT(DATEADD(Day, -1, DATEADD(month, m.mCode - 3, '" & FinancialYearFrom & "')), 'dd-MMM-yyyy') END AS ToDate,'' Account,CASE WHEN m.mCode <= 3 THEN m.mCode + 12 ELSE m.mCode END AS mCode FROM Months m LEFT JOIN ( "
        sSQL = "Select IIF(FORMAT(Date, 'MM') > 3, FORMAT(Date, 'MM'), FORMAT(Date, 'MM') + 12) AS mCode,Debit AS Debit,Credit AS Credit,AccountName FROM ( "
        If VchType = 2 Then
             OpSQL = ""
             OpSQL = "Select '" & GetDate(MhDateInput1.Text) & "' As Date,'' VchType,'' As VchBillNo,'Opening' As Account,ABS(IIF(Opening<0,Opening,0)) AS Debit,ABS(IIF(Opening>0,Opening,0)) AS Credit,'' As ShortNarration,'' AS LongNarration,'' AS Type,'' as  Code,(Select PrintName From AccountMaster Where Code IN (" & AccountList & ")) AS AccountName " & _
                            "From (SELECT(SELECT ISNULL(Sum(C.Credit-C.Debit),0) FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date < '" & GetDate(MhDateInput1.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account IN (" & AccountList & ")) " & _
                            "+ (SELECT ISNULL(Sum(IIF(LEFT(Type,2)='01',(T.Amount),IIF(LEFT(Type,2)='02',(0-T.Amount),IIF(LEFT(Type,2)='03',(T.Amount),IIF(LEFT(Type,2)='04',(0-T.Amount),0))))),0) FROM JobworkBVParent T LEFT JOIN AccountMaster A On T.Party=A.Code Where T.Date < '" & GetDate(MhDateInput1.Text) & "' And T.Party IN (" & AccountList & ")) " & _
                            "+ (Select ISNULL(Sum(Opening),0) From AccountMaster Where Code IN (" & AccountList & ")) As Opening ) TBL Union "
        End If
                dSQL = "SELECT Date As Date,'Pymt' As VchTYPE,LTrim(T.Name) As VchBillNo,A.Name As Account,(Select IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 Where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,(Select IIF(C2.TOA='C',Credit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") AND Right(BOM,2)='PI' " & _
                            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
                            "SELECT Date As Date,'Rcpt' As VchTYPE,LTrim(T.Name) As VchBillNo,A.Name As Account,(SELECT IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 Where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,(Select IIF(C2.TOA='C',Credit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") AND Right(BOM,2)='PR' " & _
                            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
                            "SELECT Date As Date,'Jrnl' As VchTYPE,LTrim(T.Name) As VchBillNo,A.Name As Account,(SELECT IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,IIF(TOA='D',C.Debit,0) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") And Right(BOM,2)='JE' " & _
                            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
                            "SELECT Date As Date,'Cntr' As VchTYPE,LTrim(T.Name) As VchBillNo,A.Name As Account,(SELECT IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,(Select IIF(C2.TOA='C',Credit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") AND Right(BOM,2)='CE' " & _
                            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
                            "SELECT Date As Date,'DrNt' As VchTYPE,LTrim(T.Name) As VchBillNo,A.Name As Account,(SELECT IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,(Select IIF(C2.TOA='C',Credit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") AND Right(BOM,2)='DN' " & _
                            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
                            "SELECT Date As Date,'CrNt' As VchTYPE,LTrim(T.Name) As VchBillNo,A.Name As Account,(SELECT IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,(Select IIF(C2.TOA='C',Credit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") AND Right(BOM,2)='CN' " & _
                            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
                            "SELECT Date As Date,'Pur' As VchTYPE,LTrim(T.Name) As VchBillNo,'Purchase' As Account, '0' As Debit,T.Amount As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T LEFT JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where Left(Type,2)='01' And Right(Type,2)='PF' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") Union " & _
                            "SELECT Date As Date,'Pur' As VchTYPE,LTrim(T.Name) As VchBillNo,'Job Work Purchase ' As Account, '0' As Debit,T.Amount As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T LEFT JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where Left(Type,2)='01' And Right(Type,2)<>'PF' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") Union " & _
                            "SELECT Date As Date,'PurRtr' As VchTYPE,LTrim(T.Name) As VchBillNo,'Purchase Return' As Account, T.Amount As Debit,'0' As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T LEFT JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where Left(Type,2)='02' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") Union " & _
                            "SELECT Date As Date,'SaleRtr' As VchTYPE,LTrim(T.Name) As VchBillNo,'Sale Return' As Account, '0' As Debit,T.Amount As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T LEFT JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where Left(Type,2)='03' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") Union " & _
                            "SELECT Date As Date,'Sale' As VchType,LTRIM(T.Name) As VchBillNo,S.Name As Account,(T.Amount+T.Rebate) As Debit,'0' As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) As Type,T.Code As Code,A.Name As AccountName FROM (((JobworkBVParent T LEFT JOIN JobworkBVChild C ON C.Code=T.Code) LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code) LEFT JOIN AccountMaster A ON T.Party=A.Code) LEFT JOIN AccountMaster S ON T.SalesType=S.Code WHERE LEFT(Type,2)='04' AND RIGHT(Type,2)='SF' AND T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND T.Party IN (" & AccountList & ") Union " & _
                            "SELECT Date As Date,'Sale' As VchTYPE,LTRIM(T.Name) As VchBillNo,'Rebate' As Account, '0' As Debit,T.Rebate As Credit,'Rebate' As ShortNarration,'' As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T LEFT JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where T.Rebate>0 And Left(Type,2)='04' And Right(Type,2)='SF' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") Union " & _
                            "SELECT Date As Date,'Sale' As VchTYPE,LTrim(T.Name) As VchBillNo,'Job Work Sale ' As Account, T.Amount As Debit,'0' As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T INNER JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where Left(Type,2)='04' And Right(Type,2)<>'SF' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") "
        
    End If
    If VchType = 1 Then SQL = mSQL + sSQL + dSQL + " ) As T) AS TBL ON m.mCode = TBL.mCode GROUP BY m.mCode,TBL.AccountName ORDER BY m.mCode; "
    If VchType = 2 Then SQL = OpSQL + dSQL + " ORDER BY Date ASC"
    
            Screen.MousePointer = vbHourglass
            If rstAccountLedger.State = adStateOpen Then rstAccountLedger.Close
            rstAccountLedger.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
            If rstAccountLedger.RecordCount = 0 And rstAccountOpening.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub

If OutputTo = "S" Then
    PrintLedger (OutputTo)
ElseIf OutputTo = "" Then
    If rstAccountLedger.RecordCount <> 0 Then
        Mh3dLabel9.Caption = "Account : " + rstAccountLedger.Fields("AccountName").Value: Mh3dLabel9.FontSize = 14
    ElseIf rstAccountLedger.RecordCount = 0 And rstAccountOpening.RecordCount <> 0 Then
            If rstAccountLedger.State = adStateOpen Then rstAccountLedger.Close
            rstAccountLedger.Open "Select Name AS AccountName From AccountMaster Where code IN (" & AccountList & ")", cnDatabase, adOpenKeyset, adLockReadOnly
            If rstAccountLedger.RecordCount <> 0 Then rstAccountLedger.MoveFirst
            If rstAccountLedger.RecordCount <> 0 Then Mh3dLabel9.Caption = "Account : " + rstAccountLedger.Fields("AccountName").Value: Mh3dLabel9.FontSize = 13
            rstAccountOpening.MoveFirst
            Opening = Format(Val(rstAccountOpening.Fields("Opening").Value), "##,##,##,##0.00")
            Mh3dLabel10.Caption = "Opening Balance: = Rs. " & Format(Opening, "##,##,##,##0.00") & IIf(Opening <= 0, " Dr.", " Cr.")
            Screen.MousePointer = vbNormal: Exit Sub
    End If


    rstAccountOpening.MoveFirst
    Opening = Format(Val(rstAccountOpening.Fields("Opening").Value), "##,##,##,##0.00")
    Mh3dLabel10.Caption = ""
    Mh3dLabel10.Caption = "Opening Balance: = Rs. " & Format(Abs(Opening), "##,##,##,##0.00") & IIf(Opening <= 0, " Dr.", " Cr.")
    
'    Bal = Opening
    With fpSpread1
            If .DataRowCnt = 0 Then
            Else
            n = .DataRowCnt:
            fpSpread1.RowHeight(n) = 12.75
           End If
            .ClearRange 1, 1, .MaxCols, .MaxRows, False
            Dim K As Integer
            K = rstAccountLedger.RecordCount
            ' Set number of columns and rows
            fpSpread1.MaxCols = 13
            fpSpread1.MaxRows = IIf(K < 27, 27, K + 1)
            For C = 1 To .MaxCols
            If VchType <= 24 And VchType >= 0 Then fpSpread1.ColHeaderRows = 1: fpSpread1.Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12:
            Next
        rstAccountLedger.MoveFirst
        Do While Not rstAccountLedger.EOF
            i = i + 1
                .SetText 1, i, rstAccountLedger.Fields("Date").Value
                .SetText 2, i, rstAccountLedger.Fields("VchType").Value
                .SetText 3, i, rstAccountLedger.Fields("VchBillNo").Value
            If VchType = 1 Then
                .SetText 4, i, rstAccountLedger.Fields("MonthYear").Value
            ElseIf VchType = 2 Then
                .SetText 4, i, rstAccountLedger.Fields("Account").Value
            End If
                        Debit = Val(rstAccountLedger.Fields("Debit").Value)
                .SetText 5, i, Val(rstAccountLedger.Fields("Debit").Value)
                        Credit = Val(rstAccountLedger.Fields("Credit").Value)
                .SetText 6, i, Val(rstAccountLedger.Fields("Credit").Value)
                        Bal = Bal + Credit - Debit
                .SetText 7, i, Bal
                .SetText 8, i, IIf(Bal <= 0, "Dr.", "Cr.")
                .SetText 9, i, rstAccountLedger.Fields("ShortNarration").Value
                .SetText 10, i, rstAccountLedger.Fields("LongNarration").Value
                .SetText 11, i, rstAccountLedger.Fields("Type").Value
                .SetText 12, i, rstAccountLedger.Fields("Code").Value
             rstAccountLedger.MoveNext
        Loop
                R = i + 1
            For C = 1 To .MaxCols
            fpSpread1.Col = C: fpSpread1.Row = R:  fpSpread1.FontSize = 12: fpSpread1.FontUnderline = True: fpSpread1.ForeColor = vbBlue: 'fpSpread1.FontBold = True:
                       '    .Col = C: .Row = R: .FontSize = 12: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue: '.FontBold = True:
            Next
                    .LockBackColor = RGB(255, 250, 255): Combo1.BackColor = RGB(255, 250, 255): Combo2.BackColor = RGB(255, 250, 255): MhDateInput1.BackColor = RGB(255, 250, 255): MhDateInput2.BackColor = RGB(255, 250, 255):  TDBNumber2.BackColor = RGB(255, 250, 255): Text1.BackColor = RGB(255, 250, 255):
                     .SelectBlockOptions = SelectBlockOptionsAll
                    .AllowMultiBlocks = True:
                    If TDBNumber2 <> 0 Then fpSpread1.SetFocus: fpSpread1.SetActiveCell 1, i + 1
    Mh3dLabel11.Caption = "Closing Balance : = Rs. " & Format(Abs(Bal), "##,##,##,##0.00") & IIf(Bal <= 0, " Dr.", " Cr.")
    Mh3dLabel12.Caption = "Accounts Ledger"
    End With
    TDBNumber2 = i
    fpSpread1.MaxRows = IIf(i < 27, 27, i + 1)
    Call cmdFilter_Click
    Screen.MousePointer = vbNormal
    Exit Sub
ElseIf OutputTo = "V" Or OutputTo = "P" Or OutputTo = "M" Then
    PrintLedger (OutputTo)
End If
Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub Mh3dLabel8_Click()
Dim C As Long, R As Long
Dim JQty As Variant
Dim FileName As String

If Dir(App.Path & "\Export", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Export"

 FileName = App.Path & "\Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"

' Save to xls file type

' Load an Excel-formatted file

fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, False

'MsgBox
    MsgBox "Import Processing....", vbInformation, "Easy Publish...Import !!! "

fpSpread1.ImportExcelBook FileName, ""        '& "\EasyPublish.xls", ""

cmdRefresh.Visible = False
    
    With fpSpread1
    
        fpSpread1.MaxCols = 13
            For C = 1 To fpSpread1.MaxCols
                fpSpread1.GetText C, 1, JQty
               fpSpread1.Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.Text = JQty
               
               If VchType <= 2 Then
                    .LockBackColor = RGB(255, 255, 240): Combo1.BackColor = RGB(255, 255, 240): Combo2.BackColor = RGB(255, 255, 240): MhDateInput1.BackColor = RGB(255, 255, 240): MhDateInput2.BackColor = RGB(255, 255, 240): TDBNumber2.BackColor = RGB(255, 255, 240): Text1.BackColor = RGB(255, 255, 240):
                ElseIf VchType >= 3 And VchType <= 6 Then
                    .LockBackColor = RGB(245, 255, 230): Combo1.BackColor = RGB(245, 255, 230): Combo2.BackColor = RGB(245, 255, 230): MhDateInput1.BackColor = RGB(245, 255, 230): MhDateInput2.BackColor = RGB(245, 255, 230): TDBNumber2.BackColor = RGB(245, 255, 230): Text1.BackColor = RGB(245, 255, 230):
                ElseIf VchType >= 7 And VchType <= 10 Then
                    .LockBackColor = RGB(245, 250, 250): Combo1.BackColor = RGB(245, 250, 250): Combo2.BackColor = RGB(245, 250, 250): MhDateInput1.BackColor = RGB(245, 250, 250): MhDateInput2.BackColor = RGB(245, 250, 250): TDBNumber2.BackColor = RGB(245, 250, 250): Text1.BackColor = RGB(245, 250, 250):
                ElseIf VchType >= 21 And VchType <= 24 Then
                    .LockBackColor = RGB(255, 250, 255): Combo1.BackColor = RGB(255, 250, 255): Combo2.BackColor = RGB(255, 250, 255): MhDateInput1.BackColor = RGB(255, 250, 255): MhDateInput2.BackColor = RGB(255, 250, 255): TDBNumber2.BackColor = RGB(255, 250, 255): Text1.BackColor = RGB(255, 250, 255):
                ElseIf VchType >= 25 And VchType <= 28 Then
                    .LockBackColor = RGB(240, 255, 255): Combo1.BackColor = RGB(240, 255, 255): Combo2.BackColor = RGB(240, 255, 255): MhDateInput1.BackColor = RGB(240, 255, 255): MhDateInput2.BackColor = RGB(240, 255, 255): TDBNumber2.BackColor = RGB(240, 255, 255): Text1.BackColor = RGB(240, 255, 255):
                End If
                
                .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 11:
            
            Next
                If VchType = 0 Then .ColWidth(1) = 49.25: .ColWidth(2) = 15: .ColWidth(3) = 15: .ColWidth(31) = 24: .ColWidth(32) = 22.75: .Col = 31: .ColHidden = False: .Col = 32: .ColHidden = False
                If VchType <= 10 And VchType >= 7 Or VchType <= 28 And VchType >= 25 Then fpSpread1.DeleteRows 1, 2 Else: fpSpread1.DeleteRows 1, 1
                    For R = 1 To .DataRowCnt - 1
                    .Col = 31: .Row = R: .Lock = False
                    Next
                    
    fpSpread1.DeleteRows .DataRowCnt, 1
    
    Call Total_Click
    fpSpread1.Col = 4: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignRight
    fpSpread1.Col = 5: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignRight
    fpSpread1.Col = 31: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignRight
    fpSpread1.Col = 32: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignRight
    End With
End Sub
Private Sub Mh3dLabel6_Click()
Dim x As Boolean, FileName As String, SheetName As String, LogFileName As String
Dim R As Long, C As Long
Dim JQty As Variant

'"Export Data" &
    With fpSpread1
    If VchType >= 0 And VchType <= 30 Then fpSpread1.InsertRows 1, 2
                    .SetText 4, 1, Mh3dLabel9.Caption
                    .SetText 7, 1, Format(Opening, "##,##,##,##0.00")
                    .SetText 8, 1, IIf(Opening <= 0, " Dr.", " Cr.")
                    .SetText 9, 1, " Rs. *** Opening Bal."
                    R = 1
                For C = 1 To .MaxCols
                    .Col = C: .Row = R: .FontBold = True: .FontSize = 12: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbRed: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter:
                Next
                    R = 2
                For C = 1 To .MaxCols
                    .Col = C: .Row = R: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter:
                    .GetText C, 0, JQty
                    .SetText C, 2, JQty
                Next
                                    
                    .ColHeadersShow = True: .PrintColHeaders = True: .PrintRowHeaders = True: .ColHeadersShow = True: .RowHeadersShow = True: .GridShowHoriz = True: .GridShowVert = True
                If VchType >= 0 And VchType <= 30 Then .Col = 4: .Row = 1: .FontBold = True: .FontSize = 14: .FontUnderline = True: .ForeColor = vbRed:
    
    End With

    If Dir(App.Path & "\Export", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Export"
    
    '
    ' Export Excel file and set result to x
     FileName = App.Path & "\Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"
    SheetName = "Sheet1" '"(" & Me.Caption & ")"
    LogFileName = "Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".txt"
    x = fpSpread1.ExportToExcelEx(FileName, SheetName, LogFileName, ExcelSaveFlagNoFormulas)
    ' Display result to user based on T/F value of x
    If x = True Then
    
    MsgBox "Export complete.", vbInformation, "Easy Publish...Export !!! "
        
        Dim oExcel As Object
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open (FileName)
        oExcel.Visible = True
        oExcel.Sheets("Sheet1").Select
        oExcel.Sheets("Sheet1").Unprotect
         Set oExcel = Nothing
    Else
    MsgBox "Export did not succeed.", vbInformation, "Easy Publish...Export !!!"
    End If
    '
    With fpSpread1
    'Delete Header Row
    If VchType >= 0 And VchType <= 30 Then fpSpread1.DeleteRows 1, 2
    End With
End Sub
Private Sub Mh3dLabel5_Click()
With fpSpread1
Dim PrintHeader As String
Dim R As Long, C As Long
Dim JQty As Variant

If VchType >= 0 And VchType <= 30 Then fpSpread1.InsertRows 1, 1
                .SetText 4, 1, Mh3dLabel9.Caption
                .SetText 7, 1, Format(Opening, "##,##,##,##0.00")
                .SetText 8, 1, IIf(Opening <= 0, " Dr.", " Cr.")
                .SetText 9, 1, " Rs. *** Opening Bal."
                R = 1
            For C = 1 To .MaxCols
                .Col = C: .Row = R: .FontBold = True: .FontSize = 12: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbRed: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter:
            Next
                If VchType >= 0 And VchType <= 30 Then .Col = 4: .Row = 1: .FontBold = True: .FontSize = 14: .FontUnderline = True: .ForeColor = vbRed:
PrintHeader = Me.Caption
.LockBackColor = vbWhite
' These are 8.5" X 11" paper dimensions in TWIPS
Const PaperWidth = 12240
Const PaperHeight = 15840
Printer.PaperSize = vbPRPSA4
' Set printing options for sheet
.PrintAbortMsg = "Printing - Click Cancel to .Quit"
.PrintJobName = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") '& ".pdf"
'.PrintHeader = "/cPrint Header/rPage # ./p/n2nd Line"
.PrintFooter = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") & " /rPage # ./p ": .FontSize = 16 '& ".pdf" ' "/cPrint Footer/rPage # ./p/n2nd Line"
.PrintBorder = True
.PrintColHeaders = True
.PrintColor = True
.PrintGrid = True
.PrintMarginTop = 750 '1440
.PrintMarginBottom = 500 '1440
.PrintMarginLeft = 100 '720
.PrintMarginRight = 100 '720
'.PrintType = SPRD_PRINT_ALL
.PrintRowHeaders = True
.PrintShadows = True
.PrintUseDataMax = True
' Center vertically
.PrintCenterOnPageV = False
' Center horizontally
.PrintCenterOnPageH = True
' Perform the printing action
' Set the sheet to print
.Sheet = 1
' Set scaling method
.PrintScalingMethod = PrintScalingMethodZoom
' Set zoom factor
.PrintZoomFactor = 0.75
' Print
'.PrintSheet 0
.PrintOrientation = PrintOrientationLandscape
.PrintSheet
.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
    'Delete Header Row
    If VchType >= 0 And VchType <= 30 Then .DeleteRows 1, 1
 End With
End Sub
Private Sub cmdFilter_Click()
        Call Total_Click
End Sub
Private Sub Command2_Click() ' Search Command
  Dim i As Integer, cVal As Variant, R As Long
    With fpSpread1
    If Text1.Text = "" Then Exit Sub
            If .DataRowCnt = 0 Then Exit Sub
                For i = 1 To .DataRowCnt 'Unhide All
                .Row = i: .RowHidden = False
            Next
        .MaxCols = 13
        
            R = IIf(.ActiveRow + 1 <> LR, .ActiveRow + 1, 1)
            LR = R
            For i = R To .DataRowCnt
            If Combo2.ListIndex >= 0 Then .GetText Combo2.ListIndex + 1, i, cVal
                        If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
                        ''''
                        ElseIf Combo2.ListIndex >= 0 Then
                        .SetActiveCell Combo2.ListIndex + 1, i: Exit Sub
                        Else
                        Exit Sub
                        End If
            Next
    End With
End Sub
Private Sub fpSpread1_BeforeUserSort(ByVal Col As Long, ByVal State As FPSpreadADO.BeforeUserSortStateConstants, DefaultAction As FPSpreadADO.BeforeUserSortDefaultActionConstants)
    Dim n As Integer
    With fpSpread1
        If .DataRowCnt = 0 Then Exit Sub
        n = .DataRowCnt:
        .RowHeight(n) = 12.75
        .DeleteRows n, 1
    End With
    End Sub
Private Sub fpSpread1_AfterUserSort(ByVal Col As Long)
With fpSpread1
    If .DataRowCnt = 0 Then Exit Sub
End With
    Call Total_Click
End Sub
Private Sub Total_Click()
    Dim i As Integer, cVal As Variant, n As Integer, R As Long, C As Long, Cols As Long, Flag As Variant
    Dim DebitVal As Variant, DebitTotal As Variant
    Dim CreditVal As Variant, CreditTotal As Variant
    With fpSpread1
    If .DataRowCnt = 0 Then Exit Sub
            n = .DataRowCnt: DebitVal = 0: CreditVal = 0: Bal = 0 'Bal = Opening
        For i = 1 To .DataRowCnt 'Unhide All
        .GetText 3, i, cVal
            If TotalFlag = False Then .Row = i: .RowHidden = False
            If cVal = "Grand Total" Then fpSpread1.DeleteRows i, 1
        Next
    fpSpread1.MaxCols = 13
         
    For i = 1 To .DataRowCnt
        
    If Combo2.ListIndex >= 0 Then .GetText Combo2.ListIndex + 1, i, cVal
                .GetText 5, i, DebitVal
                .GetText 6, i, CreditVal
                .GetText 13, i, Flag
                .GetText 7, i, cVal
        If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
                .Row = i: .RowHidden = True: n = n - 1: .SetText 13, .ActiveRow, "True": 'Hide Filter
        Else
            .Row = i
        If Not .RowHidden Then
                DebitTotal = DebitTotal + DebitVal '5
                CreditTotal = CreditTotal + CreditVal '6
                Bal = Bal - DebitVal + CreditVal
                .SetText 7, i, Bal
                .SetText 8, i, IIf(Bal <= 0, "Dr.", "Cr.")
        End If
        End If
                TDBNumber2 = n 'Data Count
        Next
                .SetText 3, i, "Grand Total"
                .SetText 5, i, DebitTotal
                .SetText 6, i, CreditTotal
                .SetText 7, i, Bal
                .SetText 8, i, IIf(Bal <= 0, "Dr.", "Cr.")
                .SetText 9, i, "<< Closing Balance. "
    End With
        Call Fomatting_Click
    fpSpread1.MaxRows = IIf(TDBNumber2.Value < 27, i + (27 - TDBNumber2.Value), i + 1)
End Sub
Private Sub Fomatting_Click()
Dim R As Long, C As Long, Cols As Long, Rows As Long
        With fpSpread1
       fpSpread1.MaxCols = 13
            Cols = .MaxCols
            R = .DataRowCnt
            For C = 1 To Cols
            fpSpread1.Col = C: fpSpread1.Row = R:  fpSpread1.FontSize = 12: fpSpread1.FontUnderline = True: fpSpread1.ForeColor = vbBlue: 'fpSpread1.FontBold = True:
        Next
'Formatting
            .SelectBlockOptions = SelectBlockOptionsAll
            If VchType <> 0 Then .SetActiveCell .ActiveCol, R
        End With
End Sub
Private Sub Preview_Click()
Dim PrintHeader As String
'*********************************************************
With fpSpread1
.ColsFrozen = 0
PrintHeader = Me.Caption
.LockBackColor = vbWhite
' These are 8.5" X 11" paper dimensions in TWIPS  12240  15840
Const PaperWidth = 12240
Const PaperHeight = 15840
' Set printing options for sheet
.PrintAbortMsg = "Printing - Click Cancel to .Quit"
.PrintJobName = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") '& ".pdf"
.PrintFooter = "        Export Data Company : " & " " & " _(" & CompCode & "_" & PrintHeader & ")" & "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]" & " /rPage # ./p " & " Print Date : ( " & Format(Date, "dd-MMM-yyyy") & " )         ": .FontSize = 16 '& ".pdf" ' "/cPrint Footer/rPage # ./p/n2nd Line"
.PrintBorder = True
.PrintColHeaders = True
.PrintColor = True
.PrintGrid = True
.PrintMarginTop = 200 '750 '1440
.PrintMarginBottom = 200 '500 '1440
.PrintMarginLeft = 100 '720
.PrintMarginRight = 100 '720
'.PrintType = SPRD_PRINT_ALL
.PrintRowHeaders = True
.PrintShadows = True
.PrintUseDataMax = True
' Center vertically
.PrintCenterOnPageV = False
' Center horizontally
.PrintCenterOnPageH = True
' Perform the printing action
' Set the sheet to print
.Sheet = 1
' Set scaling method
.PrintScalingMethod = PrintScalingMethodZoom
' Set zoom factor
.PrintZoomFactor = 0.75
' Print
'.PrintSheet 0
.PrintOrientation = PrintOrientationLandscape
'.PrintSheet
.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
   
   'If a cell is currently active, turn off edit mode
    If .EditMode = True Then
        .EditMode = False
        DoEvents
    End If
    Set spreadpreview.frm = Me
    Set pagesetup.frmPageSetup = Me
    Set PrintDlg.frmPrintDlg = Me
    Set headerfooter.frmHeaderFooter = Me
    spreadpreview.Show
 End With
End Sub
Private Sub Combo1_Change()
    If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Command1_Click()
With fpSpread1
    fpSpread1.DeleteRows .DataRowCnt, 1
    cmdRefresh_Click
End With
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbEnter And Shift = vbCtrlMask Then Call cmdFilter_Click
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstAccountLedger)
    Call CloseRecordset(rstAccountOpening)
    Call CloseRecordset(rstCompanyMaster)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
OutputTo = ""
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        PrintLedger ("S")
    ElseIf Button.Index = 2 Then
        PrintLedger ("P")
    ElseIf Button.Index = 3 Then
        PrintLedger ("E")
    ElseIf Button.Index = 4 Then
        cmdCancel_Click
    End If
End Sub
Public Sub PrintLedger(ByVal OutputType As String)
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT * FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    rstCompanyMaster.ActiveConnection = Nothing
        If rstCompanyMaster.RecordCount = 0 Then On Error GoTo 0: Exit Sub
        If rstAccountLedger.RecordCount = 0 Then On Error GoTo 0: Exit Sub
        rstAccountLedger.ActiveConnection = Nothing
        With rstAccountLedger
        'ArialMT (Western)
        'Bookman Old Style (Western)
        End With
        rstAccountLedger.MoveFirst
        rptAccountsLedger.Database.SetDataSource rstAccountLedger, 3, 1
        rptAccountsLedger.Database.SetDataSource rstAccountOpening, 3, 2
        rptAccountsLedger.DiscardSavedData
    With rptAccountsLedger
    'Section
    If VchType = "1" Then
        .Section00.Suppress = True
        .Section10.Suppress = True
        .Section20.Suppress = True
        .Section4.Suppress = False
        .Section5.Suppress = False
        .Section6.Suppress = False
        .Section05.Suppress = True
        .Section06.Suppress = True
        .Section09.Suppress = False
    ElseIf VchType = "2" Then
        .Section00.Suppress = False
        .Section10.Suppress = False
        .Section20.Suppress = False
        .Section4.Suppress = True
        .Section5.Suppress = True
        .Section6.Suppress = True
    End If
            .Text.SetText "Printed on " & Format(Now, "dd-MMM-yyyy") & " at " & Format(Now, "hh:mm"): .Text.Font.Size = 7: .Text.Font.Bold = False
            .Text1.SetText rstCompanyMaster.Fields("PrintName").Value: .Text1.Font.Size = 16: .Text2.Font.Bold = True
            .Text2.SetText rstCompanyMaster.Fields("Address1").Value & " " & rstCompanyMaster.Fields("Address2").Value & " " & rstCompanyMaster.Fields("Address3").Value & " " & rstCompanyMaster.Fields("Address4").Value: .Text2.Font.Size = 12: .Text2.Font.Bold = False
            If VchType = "1" Then .Text3.SetText " MONTHLY SUMMARY ": .Text3.Font.Size = 12: .Text3.Font.Bold = True:
            If VchType = "2" Then .Text3.SetText " Account Ledger ": .Text3.Font.Size = 12: .Text3.Font.Bold = True:
            If VchType = "2" Then .Text4.SetText "(" & Format(MhDateInput1.Value, "dd-MM-yyyy") & " to " & Format(MhDateInput2.Value, "dd-MM-yyyy") & ")": .Text4.Font.Size = 11: .Text4.Font.Bold = False
            .Text5.SetText "Accounts : " & rstAccountOpening.Fields("AccountName").Value: .Text5.Font.Size = 12: .Text5.Font.Bold = True
    If VchType = "2" Then
            .Text6.SetText rstCompanyMaster.Fields("PrintName").Value: .Text6.Font.Size = 10: .Text6.Font.Bold = False: .Text6.HorAlignment = crLeftAlign
            If VchType = "2" Then .Text7.SetText "Account Ledger : " & "(" & Format(MhDateInput1.Value, "dd-MM-yyyy") & " to " & Format(MhDateInput2.Value, "dd-MM-yyyy") & ")": .Text7.Font.Size = 10: .Text7.Font.Bold = False: .Text7.HorAlignment = crLeftAlign
            .Text8.SetText "Accounts : " & rstAccountOpening.Fields("AccountName").Value: .Text8.Font.Size = 10: .Text8.Font.Bold = False: .Text8.HorAlignment = crLeftAlign
    End If
            If VchType = "1" Then .Text8.SetText "Accounts : " & rstAccountOpening.Fields("AccountName").Value: .Text8.Font.Size = 12: .Text8.Font.Bold = False: .Text8.HorAlignment = crLeftAlign
    End With
        If OutputType = "S" Then
            Screen.MousePointer = vbNormal
            Set FrmReportViewer.Report = rptAccountsLedger: FrmReportViewer.Show vbModal
        ElseIf OutputType = "P" Then
            rptAccountsLedger.PaperSource = crPRBinAuto
            rptAccountsLedger.PrintOut
        Else
                Dim oOutlookMsg As Outlook.MailItem, FileName As String
                Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                If rstAccountOpening.State = adStateOpen Then rstAccountOpening.Close
                rstAccountOpening.Open "Select EMail From AccountMaster Where Code IN (" & AccountList & ") ", cnDatabase, adOpenKeyset, adLockReadOnly
                If rstAccountOpening.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
                With oOutlookMsg
                    .To = rstAccountOpening.Fields("EMail").Value
                    .Subject = "Account Ledger"
                    .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith " & "Account Ledger from " + Format(GetDate(MhDateInput1.Text), "dd-MMM-yyyy") + " to " + Format(GetDate(MhDateInput2.Text), "dd-MMM-yyyy") & " for doing the needful at your end.<Br><b>Kindly do acknowledge the receipt of the mail</b>.<Br><Br>Thanks & Regards<Br>Accounts Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
                    rptAccountsLedger.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                    rptAccountsLedger.ExportOptions.DestinationType = crEDTDiskFile
                    FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
                    rptAccountsLedger.ExportOptions.DiskFileName = FileName
                    rptAccountsLedger.Export False
                    .Attachments.Add (FileName)
                    .Importance = olImportanceHigh
                    .ReadReceiptRequested = True
                    If CheckEmpty(.To, False) Then .Display Else .Send
                End With
                Set oOutlookMsg = Nothing
        End If
        Set rptAccountsLedger = Nothing
        Screen.MousePointer = vbNormal
'        Call CloseForm(Me)
        Exit Sub
        On Error GoTo 0
        Screen.MousePointer = vbNormal
End Sub
