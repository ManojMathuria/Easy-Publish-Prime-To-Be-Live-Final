VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmBookPOChild07 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Miscellaneous Operations Order Details"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18810
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   18810
   StartUpPosition =   2  'CenterScreen
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   7470
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   105
      Width           =   18615
      _Version        =   65536
      _ExtentX        =   32835
      _ExtentY        =   13176
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
      Picture         =   "BookPOChild07.frx":0000
      Begin VB.CommandButton cmdProceed 
         BackColor       =   &H008BD6FE&
         Height          =   330
         Left            =   17760
         Picture         =   "BookPOChild07.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Save"
         Top             =   105
         Width           =   375
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H008BD6FE&
         Height          =   330
         Left            =   18120
         Picture         =   "BookPOChild07.frx":011E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cancel"
         Top             =   105
         Width           =   375
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
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   9
         Top             =   7015
         Width           =   13590
      End
      Begin VB.TextBox Text5 
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   0
         TabStop         =   0   'False
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   645
         Width           =   6135
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
         MaxLength       =   10
         TabIndex        =   6
         Top             =   6510
         Width           =   1575
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
         Left            =   9360
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   640
         Width           =   9135
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   7800
         TabIndex        =   11
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
         Picture         =   "BookPOChild07.frx":0220
         Picture         =   "BookPOChild07.frx":023C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   7800
         TabIndex        =   12
         Top             =   645
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
         Caption         =   " Item Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild07.frx":0258
         Picture         =   "BookPOChild07.frx":0274
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   13
         Top             =   6510
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
         Picture         =   "BookPOChild07.frx":0290
         Picture         =   "BookPOChild07.frx":02AC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   15510
         TabIndex        =   14
         Top             =   6510
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
         Picture         =   "BookPOChild07.frx":02C8
         Picture         =   "BookPOChild07.frx":02E4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   7800
         TabIndex        =   15
         Top             =   6510
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
         Picture         =   "BookPOChild07.frx":0300
         Picture         =   "BookPOChild07.frx":031C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   13320
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
         Caption         =   " Target Date"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild07.frx":0338
         Picture         =   "BookPOChild07.frx":0354
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   645
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
         Caption         =   " Party Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild07.frx":0370
         Picture         =   "BookPOChild07.frx":038C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   18
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
         Picture         =   "BookPOChild07.frx":03A8
         Picture         =   "BookPOChild07.frx":03C4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   19
         Top             =   7015
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
         Picture         =   "BookPOChild07.frx":03E0
         Picture         =   "BookPOChild07.frx":03FC
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   9360
         TabIndex        =   1
         Top             =   105
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild07.frx":0418
         Caption         =   "BookPOChild07.frx":0530
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":059C
         Keys            =   "BookPOChild07.frx":05BA
         Spin            =   "BookPOChild07.frx":0618
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
      Begin TDBDate6Ctl.TDBDate MhDateInput3 
         Height          =   330
         Left            =   14880
         TabIndex        =   2
         Top             =   105
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild07.frx":0640
         Caption         =   "BookPOChild07.frx":0758
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":07C4
         Keys            =   "BookPOChild07.frx":07E2
         Spin            =   "BookPOChild07.frx":0840
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
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   9360
         TabIndex        =   7
         Top             =   6510
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild07.frx":0868
         Caption         =   "BookPOChild07.frx":0980
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":09EC
         Keys            =   "BookPOChild07.frx":0A0A
         Spin            =   "BookPOChild07.frx":0A68
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
         Left            =   17070
         TabIndex        =   8
         Top             =   6510
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":0A90
         Caption         =   "BookPOChild07.frx":0AB0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":0B1C
         Keys            =   "BookPOChild07.frx":0B3A
         Spin            =   "BookPOChild07.frx":0B84
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
         ValueVT         =   1703936005
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   4875
         Left            =   120
         TabIndex        =   5
         Top             =   1170
         Width           =   18390
         _Version        =   524288
         _ExtentX        =   32438
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
         MaxCols         =   20
         MaxRows         =   100
         OperationMode   =   2
         ScrollBars      =   2
         SpreadDesigner  =   "BookPOChild07.frx":0BAC
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
         Height          =   330
         Left            =   17070
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":187A
         Caption         =   "BookPOChild07.frx":189A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":1906
         Keys            =   "BookPOChild07.frx":1924
         Spin            =   "BookPOChild07.frx":196E
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
         ValueVT         =   1703936005
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   6030
         Width           =   18390
         _Version        =   65536
         _ExtentX        =   32438
         _ExtentY        =   529
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
         Picture         =   "BookPOChild07.frx":1996
         Picture         =   "BookPOChild07.frx":19B2
         Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
            Height          =   330
            Left            =   14220
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1896
            _ExtentY        =   582
            Calculator      =   "BookPOChild07.frx":19CE
            Caption         =   "BookPOChild07.frx":19EE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild07.frx":1A5A
            Keys            =   "BookPOChild07.frx":1A78
            Spin            =   "BookPOChild07.frx":1AC2
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
            ValueVT         =   1922891781
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   2
         Left            =   15360
         TabIndex        =   22
         Top             =   7020
         Width           =   3120
         _Version        =   65536
         _ExtentX        =   5503
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
         Caption         =   " Ctrl+D->Delete  Ctrl+S->Save"
         AutoSize        =   -1  'True
         FillColor       =   8421504
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":1AEA
         Multiline       =   -1  'True
         GlobalMem       =   -1  'True
         Picture         =   "BookPOChild07.frx":1B06
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   18600
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   18600
         Y1              =   6415
         Y2              =   6415
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   18600
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   18600
         Y1              =   6920
         Y2              =   6920
      End
   End
End
Attribute VB_Name = "FrmBookPOChild07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstBookPOChild07 As New ADODB.Recordset
Public PartyCode As String, titleQty As Long
Dim rstElementList As New ADODB.Recordset, rstOperationList As New ADODB.Recordset, rstSizeList As New ADODB.Recordset, rstCalcModeList As New ADODB.Recordset, rstOrderList As New ADODB.Recordset, rstFetchOperationRate As New ADODB.Recordset
Dim Element As String, Operation As String, Size As String, CalcMode As String, ItemCode As String, CalcVal As Variant, i6 As Integer
Dim EditMode As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    ItemCode = FrmBookPrintOrder.rstBookList.Fields("Code").Value
    DisableCloseButton Me
    rstElementList.Open "SELECT Name,Code FROM ElementMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly 'WHERE Type='19'
    rstOperationList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type='7' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstSizeList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type IN ('1','11') ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstCalcModeList.Open "SELECT Name,Value1,Code FROM GeneralMaster WHERE Type='20' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstElementList.ActiveConnection = Nothing: rstOperationList.ActiveConnection = Nothing: rstSizeList.ActiveConnection = Nothing: rstCalcModeList.ActiveConnection = Nothing
    Call RefreshDropDownList("A")
    With fpSpread1
        .Col = 1: .TypeComboBoxList = Element
        .Col = 2: .TypeComboBoxList = Operation
        .Col = 5: .TypeComboBoxList = Size '4
        .Col = 7: .TypeComboBoxList = CalcMode '6
    End With
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text) 'Order No.
    Text4.Text = Trim(FrmBookPrintOrder.Text7.Text) 'Party Name
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text) 'Item Name
    ClearFields
    If rstBookPOChild07.RecordCount > 0 Then rstBookPOChild07.MoveFirst
    If Not CheckEmpty(CheckNull(rstBookPOChild07.Fields("Code").Value), False) Then LoadFields Else InsertOperation
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}": KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        If Not EditMode Then cmdProceed_Click: KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        If Not EditMode Then cmdCancel_Click: KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then CloseForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstElementList)
    Call CloseRecordset(rstOperationList)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstCalcModeList)
    Call CloseRecordset(rstOrderList)
    Call CloseRecordset(rstFetchOperationRate)
End Sub
Private Sub ClearFields()
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")  'Order Date
    MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")    'Target Date
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    MhRealInput18.Value = 0: MhRealInput19.Value = 0
    Text8.Text = "" 'Bill No.
    MhDateInput2.Text = "  -  -    "    'Bill Date
    MhRealInput16.Text = "0.00" 'Paid Amount
    Text6.Text = "" 'Remarks
End Sub
Private Sub LoadFields()
    If rstBookPOChild07.RecordCount = 0 Then Exit Sub
    Dim i As Integer
    MhDateInput1.Text = Format(rstBookPOChild07.Fields("OrderDate").Value, "dd-MM-yyyy")
    MhDateInput3.Text = Format(rstBookPOChild07.Fields("TargetDate").Value, "dd-MM-yyyy")
    With fpSpread1
        If rstOrderList.State = adStateOpen Then rstOrderList.Close
        rstOrderList.Open "SELECT OrderDate,TargetDate,E.Code As ECode,E.Name As EName,O.Code As OCode,O.Name As OName,T.OperationCountName As OCName,[Number],S.Code As SCode,S.Name As SName,Quantity,C.Code As CCode,C.Name As CName,T.CalcValue As CValue,Rate,Amount,Adjustment,[GST%],GST,BillAmount,Status,Narration FROM (((BookPOChild07 T INNER JOIN ElementMaster E ON T.Element=E.Code) INNER JOIN GeneralMaster O ON T.Operation=O.Code) INNER JOIN GeneralMaster C ON T.CalcMode=C.Code) LEFT JOIN GeneralMaster S ON T.[Size]=S.Code WHERE T.Code='" & CheckNull(rstBookPOChild07.Fields("Code").Value) & "' ORDER BY E.Name,O.Name", cnDatabase, adOpenKeyset, adLockReadOnly
        rstOrderList.ActiveConnection = Nothing
        If rstOrderList.RecordCount > 0 Then
            rstOrderList.MoveFirst
            i = 0
            Do While Not rstOrderList.EOF
                i = i + 1
                .SetText 1, i, rstOrderList.Fields("EName").Value
                .SetText 2, i, rstOrderList.Fields("OName").Value
                .SetText 3, i, Val(rstOrderList.Fields("Number").Value)
                .SetText 4, i, CheckNull(rstOrderList.Fields("OCName").Value)
                .SetText 5, i, CheckNull(rstOrderList.Fields("SName").Value)
                .SetText 6, i, Val(rstOrderList.Fields("Quantity").Value)
                .SetText 7, i, CheckNull(rstOrderList.Fields("CName").Value)
                .SetText 8, i, Val(rstOrderList.Fields("CValue").Value)
                .SetText 9, i, Val(rstOrderList.Fields("Rate").Value)
                .SetText 10, i, Val(rstOrderList.Fields("Amount").Value)
                .SetText 11, i, Val(rstOrderList.Fields("Adjustment").Value)
                .SetText 12, i, Val(rstOrderList.Fields("GST%").Value)
                .SetText 13, i, Val(rstOrderList.Fields("BillAmount").Value)
                .SetText 14, i, rstOrderList.Fields("ECode").Value
                .SetText 15, i, rstOrderList.Fields("OCode").Value
                .SetText 16, i, rstOrderList.Fields("SCode").Value
                .SetText 17, i, rstOrderList.Fields("CCode").Value
                .SetText 18, i, Val(rstOrderList.Fields("GST").Value)
                .SetText 19, i, Val(rstOrderList.Fields("Status").Value)
                .SetText 20, i, rstOrderList.Fields("Narration").Value
                rstOrderList.MoveNext
            Loop
        End If
    End With
    Call CalculateAmount
    Text8.Text = rstBookPOChild07.Fields("BillNo").Value
    If Not IsNull(rstBookPOChild07.Fields("BillDate").Value) Then MhDateInput2.Text = Format(rstBookPOChild07.Fields("BillDate").Value, "dd-MM-yyyy")
    MhRealInput16.Text = Format(Val(rstBookPOChild07.Fields("PaidAmount").Value), "0.00")
    Text6.Text = rstBookPOChild07.Fields("Remarks").Value
End Sub
Private Sub SaveFields()
    Dim i As Integer, Number As Variant, Quantity As Variant, Rate As Variant, Amount As Variant, Adjustment As Variant, GST As Variant, BillAmount As Variant, Element As Variant, Operation As Variant, Size As Variant, CalcMode As Variant, CalcValue As Variant, GSTAmt As Variant, Status As Variant, Narration As Variant, OperationCountName As Variant
    With rstBookPOChild07
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            .Delete adAffectCurrent
            .MoveNext
        Loop
    End With
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 3, i, Number
            .GetText 4, i, OperationCountName
            .GetText 6, i, Quantity
            .GetText 8, i, CalcValue
            .GetText 9, i, Rate
            .GetText 10, i, Amount
            .GetText 11, i, Adjustment
            .GetText 12, i, GST
            .GetText 13, i, BillAmount
            .GetText 14, i, Element
            .GetText 15, i, Operation
            .GetText 16, i, Size
            .GetText 17, i, CalcMode
            .GetText 18, i, GSTAmt
            .GetText 19, i, Status
            .GetText 20, i, Narration
            With rstBookPOChild07
                .AddNew
                .Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
                .Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
                .Fields("Element").Value = Element
                .Fields("Operation").Value = Operation
                .Fields("Number").Value = Val(Number)
                .Fields("OperationCountName").Value = OperationCountName
                .Fields("Size").Value = Size
                .Fields("Quantity").Value = Val(Quantity)
                .Fields("CalcMode").Value = CalcMode
                .Fields("CalcValue").Value = Val(CalcValue)
                .Fields("Rate").Value = Val(Rate)
                .Fields("Amount").Value = Val(Amount)
                .Fields("Adjustment").Value = Val(Adjustment)
                .Fields("GST%").Value = Val(GST)
                .Fields("GST").Value = Val(GSTAmt)
                .Fields("BillAmount").Value = Val(BillAmount)
                .Fields("Remarks").Value = Text6.Text
                .Fields("BillNo").Value = Text8.Text
                If Not IsDate(MhDateInput2.Text) Then .Fields("BillDate").Value = Null Else .Fields("BillDate").Value = GetDate(MhDateInput2.Text)
                .Fields("PaidAmount").Value = Val(MhRealInput16.Text)
                .Fields("Status").Value = Status
                .Fields("Narration").Value = Narration
                .Update
            End With
        Next
    End With
End Sub
Private Sub fpSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
  If fpSpread1.ActiveCol = 1 Or fpSpread1.ActiveCol = 2 Or fpSpread1.ActiveCol = 5 Or fpSpread1.ActiveCol = 7 Then fpSpread1.SetText 8, fpSpread1.ActiveRow, "": fpSpread1.SetText 9, fpSpread1.ActiveRow, ""
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf CheckNull(rstBookPOChild07.Fields("Code").Value) = "" Then
        MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
    End If
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If MhDateInput2.ValueIsNull Then Exit Sub
End Sub
Private Sub MhDateInput3_Validate(Cancel As Boolean)
    If Format(GetDate(MhDateInput3.Text), "yyyymmdd") <= Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then DisplayError ("Target Date cann't be prior to Order Date"): MhDateInput3.SetFocus: Cancel = True
End Sub
Private Sub CalculateAmount()   'Calculate Amount
    Dim BTAmount As Variant, BTAmountTotal As Variant, Amount As Variant, TotalAmount As Double, i As Integer
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 10, i, BTAmount
            .GetText 13, i, Amount
            BTAmountTotal = BTAmountTotal + BTAmount
            TotalAmount = TotalAmount + Amount
        Next
    End With
    MhRealInput18.Value = BTAmountTotal
    MhRealInput19.Value = TotalAmount
End Sub
Private Sub cmdProceed_Click()
    SaveFields
    FrmBookPrintOrder.Command3.Enabled = False
    Call CloseForm(Me)
End Sub
Private Sub cmdCancel_Click()
    rstBookPOChild07.CancelUpdate
    Call CloseForm(Me)
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If UserLevel = "3" Then Call DisplayError("You don't have the rights to delete BOM Item"): Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1
            fpSpread1.SetFocus
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        Call RefreshDropDownList("R")
    ElseIf Shift = 0 And KeyCode = vbKeyDelete Then
        If fpSpread1.ActiveCol = 16 Then
            fpSpread1.SetText fpSpread1.ActiveCol, fpSpread1.ActiveRow, ""  'Size Name
            fpSpread1.SetText 16, fpSpread1.ActiveRow, ""   'Size Code
        End If
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    On Error GoTo ErrorHandler
    Dim ActiveCellVal As Variant, CalcType As String, Number  As Variant, Size As Variant, Qty As Variant, CalcMode As Variant, Rate As Variant, xPos As Integer, Amount As Double, Adjustment As Variant, GST As Variant, Operation As Variant, Binding As Variant, AreaRange As Variant, xCalcVal As Variant
    With fpSpread1
        If .ActiveCol <> 7 Then i6 = 0
        .GetText Col, Row, ActiveCellVal
        If Col = 1 Then 'Element
            If Not CheckEmpty(ActiveCellVal, False) Then
                If rstElementList.RecordCount > 0 Then rstElementList.MoveFirst
                rstElementList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
                If Not rstElementList.EOF Then .SetText Col + 13, Row, rstElementList.Fields("Code").Value
            End If
        ElseIf Col = 2 Then 'Operation
            If Not CheckEmpty(ActiveCellVal, False) Then
                If rstOperationList.RecordCount > 0 Then rstOperationList.MoveFirst
                rstOperationList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
                If Not rstOperationList.EOF Then
                    .SetText Col + 13, Row, rstOperationList.Fields("Code").Value
                    .GetText 6, Row, Qty 'Quantity
                    If Val(Qty) = 0 Then .SetText 6, Row, FrmBookPrintOrder.MhRealInput3.Value
                End If
            End If
        ElseIf Col = 5 Then 'Size
            If Not CheckEmpty(ActiveCellVal, False) Then
                If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
                rstSizeList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
                If Not rstSizeList.EOF Then .SetText Col + 11, Row, rstSizeList.Fields("Code").Value
            End If
        ElseIf Col = 7 Then 'Calc Mode
            If Not CheckEmpty(ActiveCellVal, False) Then
                If rstCalcModeList.RecordCount > 0 Then rstCalcModeList.MoveFirst
                rstCalcModeList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
                If Not rstCalcModeList.EOF Then
                    .SetText Col + 10, Row, rstCalcModeList.Fields("Code").Value
                    .GetText 8, .ActiveRow, Qty
                    'CalcVal = IIf(Qty = 0, Val(rstCalcModeList.Fields("Value1").Value), CalcVal)
                    CalcVal = IIf(Val(xCalcVal) = 0, Val(rstCalcModeList.Fields("Value1").Value), Val(CalcVal))
                End If
                i6 = i6 + 1
                If i6 = 1 Then
                    If (rstCalcModeList.Fields("Code").Value) = "*20006" Then
                        CalcVal = InputBox("Please Confirm Input Value for" & Chr(13) & "Quantity Per Packet", "Easy Publish Prime", Val(CalcVal))
                        .SetText 8, .ActiveRow, CalcVal
                    ElseIf (rstCalcModeList.Fields("Code").Value) = "*20009" Then
                        CalcVal = InputBox("Please Confirm Input Value for" & Chr(13) & "Quantity Per Box", "Easy Publish Prime", Val(CalcVal))
                        .SetText 8, .ActiveRow, CalcVal
                    ElseIf (rstCalcModeList.Fields("Code").Value) = "*20010" Then
                        CalcVal = InputBox("Please Confirm Input Value for" & Chr(13) & "Quantity Per Bundle", "Easy Publish Prime", Val(CalcVal))
                        .SetText 8, .ActiveRow, CalcVal
                    ElseIf (rstCalcModeList.Fields("Code").Value) = "*20008" Then
                        CalcVal = InputBox("Please Confirm Input Value for" & Chr(13) & "Per Paisa Inch²", "Easy Publish Prime", Val(CalcVal))
                        .SetText 8, .ActiveRow, CalcVal: .SetText 9, .ActiveRow, 0.01
                    End If
                End If
            End If
        End If
        .GetText 17, Row, CalcMode
        If Not CheckEmpty(CalcMode, False) Then
            If rstCalcModeList.RecordCount > 0 Then rstCalcModeList.MoveFirst
            rstCalcModeList.Find "[Code]='" & FixQuote(CalcMode) & "'"
            If Not rstCalcModeList.EOF Then CalcType = IIf(InStr(1, rstCalcModeList.Fields("Name").Value, "Inch") > 0, "S", "O")
            .GetText 8, Row, CalcVal
            If CalcVal = "" Then CalcVal = Val(rstCalcModeList.Fields("Value1").Value): .SetText 8, .ActiveRow, CalcVal
        End If
        'Fetch Rate
        .GetText 9, Row, Rate
        If Val(Rate) = 0 Then
            .GetText 6, Row, Qty 'Quantity
            .GetText 3, Row, Number
            .GetText 15, Row, Operation
            .GetText 5, Row, AreaRange: If AreaRange <> "" Then AreaRange = Left(AreaRange, 5) * Mid(AreaRange, 7, 5)
            .GetText 16, Row, Size
            If Not (CheckEmpty(Operation, False) And CheckEmpty(CalcMode, False)) And Val(Qty) > 0 Then .SetText 9, Row, FetchOperationRate(Operation, CalcMode, IIf(CalcType = "O", Size, ""), Val(AreaRange), Val(Number), Val(Qty))
        End If
        If Col >= 3 And Col <= 12 Then
            .GetText 3, Row, Number
            .GetText 5, Row, Size
            .GetText 6, Row, Qty
            .GetText 8, Row, CalcVal
            .GetText 11, Row, Adjustment
            .GetText 12, Row, GST
            If CalcType = "S" And (Not CheckEmpty(Size, False)) Then xPos = InStr(1, LCase(Size), "x"): Size = Val(Left(Size, xPos - 1)) * Val(Mid(Size, xPos + 1, 5)) Else Size = 1
            If CalcVal = 0 Then CalcVal = 1
            Amount = Round((Number * Size * Val(Qty) * Val(Rate)) / CalcVal, 2)
            .SetText 10, Row, Amount 'Amount
            .SetText 18, Row, ((Amount + Val(Adjustment)) * Val(GST)) / 100 'GST
            .SetText 13, Row, Round(Amount + Val(Adjustment) + (((Amount + Val(Adjustment)) * Val(GST)) / 100), 0) 'BillAmount
            CalculateAmount
        End If
ErrorHandler:
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub RefreshDropDownList(ByVal xType As String)
    If xType = "R" Then
        rstElementList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstElementList): Loop
        rstElementList.ActiveConnection = Nothing
        rstOperationList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstOperationList): Loop
        rstOperationList.ActiveConnection = Nothing
        rstSizeList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstSizeList): Loop
        rstSizeList.ActiveConnection = Nothing
        rstCalcModeList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstCalcModeList): Loop
        rstCalcModeList.ActiveConnection = Nothing
        Element = "": Operation = "": Size = "": CalcMode = ""
    End If
    With rstElementList
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If Element = "" Then Element = .Fields("Name").Value Else Element = Element + Chr$(9) + .Fields("Name").Value
            .MoveNext
        Loop
    End With
    With rstOperationList
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If Operation = "" Then Operation = .Fields("Name").Value Else Operation = Operation + Chr$(9) + .Fields("Name").Value
            .MoveNext
        Loop
    End With
    With rstSizeList
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If Size = "" Then Size = .Fields("Name").Value Else Size = Size + Chr$(9) + .Fields("Name").Value
            .MoveNext
        Loop
    End With
    With rstCalcModeList
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If CalcMode = "" Then CalcMode = .Fields("Name").Value Else CalcMode = CalcMode + Chr$(9) + .Fields("Name").Value
            .MoveNext
        Loop
    End With
End Sub
Private Function FetchOperationRate(ByVal xOperation As String, ByVal xCalcMode As String, ByVal xSize As String, xAreaRange As Double, xSectionRange As Double, xQtyRange As Double) As Double
    On Error GoTo ErrorHandler
    If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
    rstFetchOperationRate.Open "SELECT TOP 1 Rate FROM AccountChild07 WHERE Code = '" & PartyCode & "' AND BinderyProcess='" & xOperation & "' AND CalcMode='" & xCalcMode & "' AND " & IIf(CheckEmpty(xSize, False), "1=1", "[Size]='" & xSize & "'") & " AND AreaRange>=" & xAreaRange & " ORDER BY Range", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstFetchOperationRate.RecordCount = 0 Then
        If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
        rstFetchOperationRate.Open "SELECT TOP 1 Rate FROM AccountMaster P INNER JOIN AccountChild07 C ON P.Code=C.Code WHERE [Name] Like '%Rate%'  AND Operation='" & xOperation & "' AND CalcMode='" & xCalcMode & "' AND " & IIf(CheckEmpty(xSize, False), "1=1", "[Size]='" & xSize & "'") & " AND AreaRange>=" & xAreaRange & " ORDER BY Range", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstFetchOperationRate.RecordCount > 0 Then FetchOperationRate = Val(rstFetchOperationRate.Fields("Rate").Value)
    Exit Function
ErrorHandler:
    DisplayError (Err.Description)
End Function
Private Sub InsertOperation()
    With rstOrderList
        If .State = adStateOpen Then .Close
        .Open "SELECT E.Code As ECode,E.Name As EName,O.Code As OCode,O.Name As OName,T.OperationCountName As OCName,[Number],S.Code As SCode,S.Name As SName,C.Code As CCode,C.Name As CName,T.CalcValue As CalcVal FROM (((BookChild07 T INNER JOIN GeneralMaster E ON T.Element=E.Code) INNER JOIN GeneralMaster O ON T.Operation=O.Code) INNER JOIN GeneralMaster C ON T.CalcMode=C.Code) LEFT JOIN GeneralMaster S ON T.[Size]=S.Code WHERE T.Code='" & ItemCode & "' ORDER BY E.Name,O.Name", cnDatabase, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        If .RecordCount = 0 Then
            Dim Operation, CalcMode, Size, CalcVal
            With fpSpread1
                .SetText 3, 1, 1: .SetText 6, 1, titleQty 'Set Number,Quantity
                'Element
                If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
                rstFetchOperationRate.Open "SELECT Code,Name FROM ElementMaster WHERE Code='*00027'", cnDatabase, adOpenKeyset, adLockReadOnly
                If rstFetchOperationRate.RecordCount > 0 Then .SetText 1, 1, rstFetchOperationRate.Fields("Name").Value: .SetText 14, 1, rstFetchOperationRate.Fields("Code").Value
                'Operation,Size
                If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
                rstFetchOperationRate.Open "SELECT O.Code As OperationCode,O.Name As OperationName,S.Code As SizeCode,S.Name As SizeName FROM (BookMaster I LEFT JOIN GeneralMaster O ON I.LaminationType=O.Code) LEFT JOIN GeneralMaster S ON I.FinishSize=S.Code WHERE I.Code='" & ItemCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
                If rstFetchOperationRate.RecordCount > 0 Then .SetText 2, 1, rstFetchOperationRate.Fields("OperationName").Value: .SetText 15, 1, rstFetchOperationRate.Fields("OperationCode").Value: .SetText 5, 1, rstFetchOperationRate.Fields("SizeName").Value: .SetText 16, 1, rstFetchOperationRate.Fields("SizeCode").Value: Operation = rstFetchOperationRate.Fields("OperationCode").Value: Size = rstFetchOperationRate.Fields("SizeCode").Value
                If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
                'CalcMode
                rstFetchOperationRate.Open "SELECT Code,Name FROM GeneralMaster WHERE Code='*20001'", cnDatabase, adOpenKeyset, adLockReadOnly
                If rstFetchOperationRate.RecordCount > 0 Then .SetText 7, 1, rstFetchOperationRate.Fields("Name").Value: .SetText 17, 1, rstFetchOperationRate.Fields("Code").Value: CalcMode = rstFetchOperationRate.Fields("Code").Value
                If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
                rstFetchOperationRate.Open "SELECT Rate FROM AccountChild07 WHERE Code = '" & PartyCode & "' AND BinderyProcess='" & Operation & "' AND CalcMode='" & CalcMode & "' AND [Size]='" & Size & "'", cnDatabase, adOpenKeyset, adLockReadOnly
                'Rate
                If rstFetchOperationRate.RecordCount > 0 Then .SetText 9, 1, Val(rstFetchOperationRate.Fields("Rate").Value): .SetText 10, 1, titleQty * Val(rstFetchOperationRate.Fields("Rate").Value): .SetText 13, 1, titleQty * Val(rstFetchOperationRate.Fields("Rate").Value)
            End With
        Else
            Dim i As Integer
            Do While Not .EOF
                i = i + 1
                fpSpread1.SetText 1, i, rstOrderList.Fields("EName").Value
                fpSpread1.SetText 2, i, rstOrderList.Fields("OName").Value
                fpSpread1.SetText 3, i, Val(rstOrderList.Fields("Number").Value)
                fpSpread1.SetText 4, i, CheckNull(rstOrderList.Fields("OCName").Value)
                fpSpread1.SetText 5, i, CheckNull(rstOrderList.Fields("SName").Value)
                fpSpread1.SetText 7, i, CheckNull(rstOrderList.Fields("CName").Value)
                fpSpread1.SetText 8, i, Val(rstOrderList.Fields("CalcValue").Value)
                fpSpread1.SetText 14, i, rstOrderList.Fields("ECode").Value
                fpSpread1.SetText 15, i, rstOrderList.Fields("OCode").Value
                fpSpread1.SetText 16, i, rstOrderList.Fields("SCode").Value
                fpSpread1.SetText 17, i, rstOrderList.Fields("CCode").Value
                .MoveNext
            Loop
        End If
    End With
End Sub

