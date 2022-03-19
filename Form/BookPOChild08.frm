VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmBookPOChild08 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bindery & Finishing Process Order Details"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20010
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
   ScaleWidth      =   20010
   StartUpPosition =   2  'CenterScreen
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   7470
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   105
      Width           =   19815
      _Version        =   65536
      _ExtentX        =   34951
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
      Picture         =   "BookPOChild08.frx":0000
      Begin VB.CommandButton cmdProceed 
         BackColor       =   &H008BD6FE&
         Height          =   330
         Left            =   18360
         Picture         =   "BookPOChild08.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Save"
         Top             =   105
         Width           =   375
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H008BD6FE&
         Height          =   330
         Left            =   18720
         Picture         =   "BookPOChild08.frx":011E
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
         Width           =   12750
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
         Width           =   10335
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
         Picture         =   "BookPOChild08.frx":0220
         Picture         =   "BookPOChild08.frx":023C
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
         Picture         =   "BookPOChild08.frx":0258
         Picture         =   "BookPOChild08.frx":0274
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
         Picture         =   "BookPOChild08.frx":0290
         Picture         =   "BookPOChild08.frx":02AC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   16230
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
         Picture         =   "BookPOChild08.frx":02C8
         Picture         =   "BookPOChild08.frx":02E4
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
         Picture         =   "BookPOChild08.frx":0300
         Picture         =   "BookPOChild08.frx":031C
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
         Picture         =   "BookPOChild08.frx":0338
         Picture         =   "BookPOChild08.frx":0354
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
         Picture         =   "BookPOChild08.frx":0370
         Picture         =   "BookPOChild08.frx":038C
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
         Picture         =   "BookPOChild08.frx":03A8
         Picture         =   "BookPOChild08.frx":03C4
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
         Picture         =   "BookPOChild08.frx":03E0
         Picture         =   "BookPOChild08.frx":03FC
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
         Calendar        =   "BookPOChild08.frx":0418
         Caption         =   "BookPOChild08.frx":0530
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":059C
         Keys            =   "BookPOChild08.frx":05BA
         Spin            =   "BookPOChild08.frx":0618
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
         Calendar        =   "BookPOChild08.frx":0640
         Caption         =   "BookPOChild08.frx":0758
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":07C4
         Keys            =   "BookPOChild08.frx":07E2
         Spin            =   "BookPOChild08.frx":0840
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
         Calendar        =   "BookPOChild08.frx":0868
         Caption         =   "BookPOChild08.frx":0980
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":09EC
         Keys            =   "BookPOChild08.frx":0A0A
         Spin            =   "BookPOChild08.frx":0A68
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
         Left            =   17790
         TabIndex        =   8
         Top             =   6510
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":0A90
         Caption         =   "BookPOChild08.frx":0AB0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":0B1C
         Keys            =   "BookPOChild08.frx":0B3A
         Spin            =   "BookPOChild08.frx":0B84
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
         ValueVT         =   1638405
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   4875
         Left            =   120
         TabIndex        =   5
         Top             =   1170
         Width           =   19590
         _Version        =   524288
         _ExtentX        =   34555
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
         MaxCols         =   26
         MaxRows         =   100
         OperationMode   =   2
         ScrollBars      =   2
         SpreadDesigner  =   "BookPOChild08.frx":0BAC
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
         Height          =   330
         Left            =   18525
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1A38
         Caption         =   "BookPOChild08.frx":1A58
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1AC4
         Keys            =   "BookPOChild08.frx":1AE2
         Spin            =   "BookPOChild08.frx":1B2C
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
         ValueVT         =   1638405
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   6030
         Width           =   19590
         _Version        =   65536
         _ExtentX        =   34555
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
         Picture         =   "BookPOChild08.frx":1B54
         Picture         =   "BookPOChild08.frx":1B70
         Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
            Height          =   330
            Left            =   15405
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1896
            _ExtentY        =   582
            Calculator      =   "BookPOChild08.frx":1B8C
            Caption         =   "BookPOChild08.frx":1BAC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild08.frx":1C18
            Keys            =   "BookPOChild08.frx":1C36
            Spin            =   "BookPOChild08.frx":1C80
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   2
         Left            =   14520
         TabIndex        =   22
         Top             =   7020
         Width           =   4440
         _Version        =   65536
         _ExtentX        =   7832
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
         Caption         =   "  F2->Delete Row  Ctrl+D->Delete Row  Ctrl+S->Save"
         AutoSize        =   -1  'True
         FillColor       =   8421504
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":1CA8
         Multiline       =   -1  'True
         GlobalMem       =   -1  'True
         Picture         =   "BookPOChild08.frx":1CC4
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   19300
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   19300
         Y1              =   6415
         Y2              =   6415
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   19300
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   19300
         Y1              =   6920
         Y2              =   6920
      End
   End
End
Attribute VB_Name = "FrmBookPOChild08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstBookPOChild08 As New ADODB.Recordset
Public PartyCode As String, OrderQty As Long
Dim rstSubItemList As New ADODB.Recordset, rstElementList As New ADODB.Recordset, rstOperationList As New ADODB.Recordset, rstBindingList As New ADODB.Recordset, rstSizeList As New ADODB.Recordset, rstCalcModeList As New ADODB.Recordset, rstOrderList As New ADODB.Recordset, rstFetchOperationRate As New ADODB.Recordset, rstBindingNBinderyProcess As New ADODB.Recordset, rstItemDetails As New ADODB.Recordset
Dim SubItem As String, Element As String, Operation As String, Binding As String, Size As String, CalcMode As String, ItemCode As String, CalcVal As Variant, i6 As Integer, i7 As Integer, Fraction  As Variant
Dim EditMode As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    ItemCode = FrmBookPrintOrder.rstBookList.Fields("Code").Value
    DisableCloseButton Me
    rstSubItemList.Open "SELECT (Select Name From BookMaster Where I1.Item=Code) As Name,I1.Item As Code FROM BookMaster I INNER JOIN BookChild01 I1 ON  I1.Code=I.Code Where I1.Category=4 AND I.Code='" & ItemCode & "' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstSubItemList.RecordCount = 0 Then rstSubItemList.Close: rstSubItemList.Open "SELECT Name As Name,Code As Code FROM BookMaster Where Code='" & ItemCode & "' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstElementList.Open "SELECT Name,Code FROM ElementMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly 'WHERE Type='19'  Where Name Like '%FG%'
    rstOperationList.Open "SELECT Name,Code,Value1 As oValue1 FROM GeneralMaster WHERE Type='7' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstBindingList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type='6' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstSizeList.Open "SELECT Name,Code,Type AS sType FROM GeneralMaster WHERE Type IN ('1','11') ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstCalcModeList.Open "SELECT Name,Value1,Code FROM GeneralMaster WHERE Type='20' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstOperationList.ActiveConnection = Nothing: rstBindingList.ActiveConnection = Nothing: rstSizeList.ActiveConnection = Nothing: rstCalcModeList.ActiveConnection = Nothing: rstSubItemList.ActiveConnection = Nothing: rstElementList.ActiveConnection = Nothing
    Call RefreshDropDownList("A")
    With fpSpread1
        .Col = 1: .TypeComboBoxList = SubItem
        .Col = 2: .TypeComboBoxList = Binding
        .Col = 3: .TypeComboBoxList = Element
        .Col = 4: .TypeComboBoxList = Operation '3
        .Col = 7: .TypeComboBoxList = Size '6'4
        .Col = 10: .TypeComboBoxList = CalcMode '8'6
    End With
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text) 'Order No.
    Text4.Text = Trim(FrmBookPrintOrder.Text7.Text) 'Party Name
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text) 'Item Name
    ClearFields
    If rstBookPOChild08.RecordCount > 0 Then rstBookPOChild08.MoveFirst
    If Not CheckEmpty(CheckNull(rstBookPOChild08.Fields("Code").Value), False) Then LoadFields Else InsertOperation
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
    Call CloseRecordset(rstSubItemList)
    Call CloseRecordset(rstOperationList)
    Call CloseRecordset(rstBindingList)
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
    If rstBookPOChild08.RecordCount = 0 Then Exit Sub
    Dim i As Integer, ICode As String, ECode As String, BCode As String
    MhDateInput1.Text = Format(rstBookPOChild08.Fields("OrderDate").Value, "dd-MM-yyyy")
    MhDateInput3.Text = Format(rstBookPOChild08.Fields("TargetDate").Value, "dd-MM-yyyy")
    With fpSpread1
        If rstOrderList.State = adStateOpen Then rstOrderList.Close
        rstOrderList.Open "SELECT OrderDate,TargetDate,I.Code As ICode,I.Name As IName,E.Code As ECode,E.Name As EName,O.Code As OCode,O.Name As OName,O.Value1 As oValue1,B.Code As BCode,B.Name As BName,T.OperationCountName As OCName,[Number],S.Code As SCode,S.Name As SName,S.Type As SType,Quantity,C.Code As CCode,C.Name As CName,T.CalcValue As CValue,Rate,Amount,Adjustment,[GST%],GST,BillAmount,Status,T.Narration FROM (((((BookPOChild08 T INNER JOIN BookMaster I ON T.SUBITEM=I.Code)INNER JOIN ElementMaster E ON T.Element=E.Code) INNER JOIN GeneralMaster O ON T.BinderyProcess=O.Code) INNER JOIN GeneralMaster B ON T.BindingType=B.Code)INNER JOIN GeneralMaster C ON T.CalcMode=C.Code) LEFT JOIN GeneralMaster S ON T.[Size]=S.Code WHERE T.Code='" & CheckNull(rstBookPOChild08.Fields("Code").Value) & "' ORDER BY E.Name,BinderyProcess,O.Name", cnDatabase, adOpenKeyset, adLockReadOnly
        rstOrderList.ActiveConnection = Nothing
        If rstOrderList.RecordCount > 0 Then
            rstOrderList.MoveFirst
            .MaxCols = 27
            i = 0
            Do While Not rstOrderList.EOF
                i = i + 1
                If ICode = "" Or ICode <> rstOrderList.Fields("ICode").Value Then .SetText 1, i, rstOrderList.Fields("IName").Value: ICode = rstOrderList.Fields("ICode").Value
                If BCode = "" Or (BCode + ICode) <> (rstOrderList.Fields("BCode").Value + rstOrderList.Fields("ICode").Value) Then .SetText 2, i, rstOrderList.Fields("BName").Value: ICode = rstOrderList.Fields("ICode").Value: BCode = rstOrderList.Fields("BCode").Value
                If ECode = "" Or ECode <> rstOrderList.Fields("ECode").Value Then .SetText 3, i, rstOrderList.Fields("EName").Value: ECode = rstOrderList.Fields("ECode").Value
                .SetText 4, i, rstOrderList.Fields("OName").Value 'Operation
                .SetText 5, i, Val(rstOrderList.Fields("Number").Value) 'Number
                .SetText 6, i, CheckNull(rstOrderList.Fields("OCName").Value) 'Operation Count Name
                .SetText 7, i, CheckNull(rstOrderList.Fields("SName").Value) 'Size Name
                .SetText 8, i, IIf(Val(rstOrderList.Fields("STYPE").Value) = 11, 1, 0)
                .SetText 9, i, Val(rstOrderList.Fields("Quantity").Value)
                .SetText 10, i, CheckNull(rstOrderList.Fields("CName").Value) 'Calc Mode
                .SetText 11, i, Val(rstOrderList.Fields("CValue").Value) 'Calc Value
                .SetText 12, i, Val(rstOrderList.Fields("Rate").Value)
                .SetText 13, i, Val(rstOrderList.Fields("Amount").Value)
                .SetText 14, i, Val(rstOrderList.Fields("Adjustment").Value)
                .SetText 15, i, Val(rstOrderList.Fields("GST%").Value)
                .SetText 16, i, Val(rstOrderList.Fields("GST").Value)
                .SetText 17, i, Val(rstOrderList.Fields("BillAmount").Value)
                .SetText 18, i, rstOrderList.Fields("ECode").Value: ECode = rstOrderList.Fields("ECode").Value
                .SetText 19, i, rstOrderList.Fields("OCode").Value
                .SetText 20, i, rstOrderList.Fields("SCode").Value
                .SetText 21, i, rstOrderList.Fields("CCode").Value
                .SetText 22, i, Val(rstOrderList.Fields("Status").Value)
                .SetText 23, i, rstOrderList.Fields("Narration").Value
                .SetText 24, i, rstOrderList.Fields("BCode").Value: BCode = rstOrderList.Fields("BCode").Value
                .SetText 25, i, Val(rstOrderList.Fields("oValue1").Value)
                .SetText 26, i, rstOrderList.Fields("ICode").Value: ICode = rstOrderList.Fields("ICode").Value
                .SetText 27, i, rstOrderList.Fields("STYPE").Value
                rstOrderList.MoveNext
            Loop
        End If
    End With
    Call CalculateAmount
    Text8.Text = rstBookPOChild08.Fields("BillNo").Value
    If Not IsNull(rstBookPOChild08.Fields("BillDate").Value) Then MhDateInput2.Text = Format(rstBookPOChild08.Fields("BillDate").Value, "dd-MM-yyyy")
    MhRealInput16.Text = Format(Val(rstBookPOChild08.Fields("PaidAmount").Value), "0.00")
    Text6.Text = rstBookPOChild08.Fields("Remarks").Value
End Sub
Private Sub SaveFields()
    Dim SubItem As Variant, i As Integer, Number As Variant, Quantity As Variant, Rate As Variant, Amount As Variant, Adjustment As Variant, GST As Variant, BillAmount As Variant, Element As Variant, Operation As Variant, Binding As Variant, Size As Variant, CalcMode As Variant, CalcValue As Variant, GSTAmt As Variant, Status As Variant, Narration As Variant, OperationCountName As Variant, Fraction As Variant
    With rstBookPOChild08
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            .Delete adAffectCurrent
            .MoveNext
        Loop
    End With
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 4, i, Number
            .GetText 5, i, OperationCountName
            .GetText 7, i, Quantity
            .GetText 9, i, CalcValue
            .GetText 8, i, Fraction
            .GetText 10, i, Rate
            .GetText 11, i, Amount
            .GetText 12, i, Adjustment
            .GetText 13, i, GST
            .GetText 14, i, GSTAmt
            .GetText 15, i, BillAmount
            .GetText 16, i, Element
            .GetText 17, i, Operation
            .GetText 18, i, Size
            .GetText 19, i, CalcMode
            .GetText 20, i, Status
            .GetText 21, i, Narration
            .GetText 22, i, Binding
            .GetText 26, i, SubItem
            With rstBookPOChild08
                .AddNew
                .Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
                .Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
                .Fields("SubItem").Value = SubItem
                .Fields("BindingType").Value = Binding
                .Fields("Element").Value = Element
                .Fields("BinderyProcess").Value = Operation
                .Fields("Number").Value = Val(Number)
                .Fields("OperationCountName").Value = OperationCountName
                .Fields("Size").Value = Size
                .Fields("Fraction").Value = Fraction
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
Dim cVal As Variant
fpSpread1.GetText 19, Row, cVal
    If cVal <> "" Then
        If Col = 1 Or Col = 2 Then fpSpread1.SetText Col, Row, ""
    ElseIf cVal = "" Then
        If Col = 2 Then
            fpSpread1.GetText Col, Row, cVal
            If Not CheckEmpty(cVal, False) Then
                If rstBindingList.RecordCount > 0 Then rstBindingList.MoveFirst
                rstBindingList.Find "[Name]='" & FixQuote(cVal) & "'"
                If Not rstBindingList.EOF Then
                    fpSpread1.SetText Col + 20, Row, rstBindingList.Fields("Code").Value
                End If
            End If
            If rstItemDetails.State = adStateOpen Then rstItemDetails.Close
            rstItemDetails.Open "Select BindingForms01,BindingForms02,(Select Name From GeneralMaster Where Code= FinishSize)  As FSize,FinishSize FROM BookMaster WHERE Code='" & ItemCode & "' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
            If rstItemDetails.RecordCount = 0 Then Exit Sub
        Dim i As Integer, OCN As Variant, NO As Variant
        With rstBindingNBinderyProcess
        fpSpread1.GetText 22, Row, cVal
            i = fpSpread1.DataRowCnt
            If .State = adStateOpen Then .Close
            .Open "SELECT B.Code AS BCode,B.Name As BName,(Select Name From GeneralMaster Where Code=IIf(C.BinderyProcess = '*07037', '*20005', IIf(C.BinderyProcess = '*07039', '*20005', IIf(C.BinderyProcess = '*07036', '*20001', IIf(C.BinderyProcess = '*07038', '*20005', IIf(C.BinderyProcess = '*07041', '*20009', IIf(C.BinderyProcess = '*07051', '*20005', '*20006'))))))) As CalcName,(Select Value1 From GeneralMaster Where Code=IIf(C.BinderyProcess = '*07037', '*20005', IIf(C.BinderyProcess = '*07039', '*20005', IIf(C.BinderyProcess = '*07036', '*20001', IIf(C.BinderyProcess = '*07038', '*20005', IIf(C.BinderyProcess = '*07041', '*20009', IIf(C.BinderyProcess = '*07051', '*20005', '*20006'))))))) As CVal FROM BindingTypeChild C INNER JOIN GeneralMaster B ON C.BinderyProcess=B.Code WHERE C.Code='" & cVal & "' ORDER BY B.Name", cnDatabase, adOpenKeyset, adLockReadOnly
            If .RecordCount = 0 Then
            Exit Sub
            ElseIf .RecordCount > 0 Then
                If MsgBox("Want to load all bindery process for this binding type?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Load !") = vbYes Then
                    Do Until .EOF
                        fpSpread1.SetText 3, i, .Fields("BName").Value: fpSpread1.SetText 17, i, .Fields("BCode").Value 'Bindery Process
                        fpSpread1.GetText 17, i, OCN
                                                                    NO = Val(rstItemDetails.Fields("BindingForms01").Value) + Val(rstItemDetails.Fields("BindingForms02").Value) 'Number
                        fpSpread1.SetText 4, i, IIf(OCN = "*07037", NO, IIf(OCN = "*07039", NO, IIf(OCN = "*07051", NO, 1))) 'Number
                        fpSpread1.SetText 5, i, IIf(OCN = "*07037", "Sections", IIf(OCN = "*07039", "Forms", IIf(OCN = "*07051", "Sections", "Nos"))) 'Number
                        fpSpread1.SetText 6, i, rstItemDetails.Fields("FSIZE").Value: fpSpread1.SetText 18, i, rstItemDetails.Fields("FinishSize").Value 'Size
                        fpSpread1.SetText 7, i, FrmBookPrintOrder.MhRealInput3.Value 'Quantity
                        fpSpread1.SetText 8, i, .Fields("CalcName").Value
                        fpSpread1.SetText 9, i, IIf(Val(.Fields("CVal").Value) <> 0, Val(.Fields("CVal").Value), "")
                        fpSpread1.SetText 19, i, IIf(OCN = "*07037", "*20005", IIf(OCN = "*07039", "*20005", IIf(OCN = "*07036", "*20001", IIf(OCN = "*07038", "*20005", IIf(OCN = "*07041", "*20009", IIf(OCN = "*07051", "*20005", "*20006")))))) 'CalcMode
                        fpSpread1.SetText 22, i, cVal
                
                        i = i + 1
                        .MoveNext
                    Loop
                End If
            End If
        End With
    End If
    End If
    If fpSpread1.ActiveCol = 10 Then fpSpread1.SetText 11, fpSpread1.ActiveRow, "": fpSpread1.SetText 12, fpSpread1.ActiveRow, ""
    If fpSpread1.ActiveCol = 4 Or fpSpread1.ActiveCol = 7 Then fpSpread1.SetText 12, fpSpread1.ActiveRow, ""
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf CheckNull(rstBookPOChild08.Fields("Code").Value) = "" Then
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
            .GetText 13, i, BTAmount
            .GetText 17, i, Amount
            TotalAmount = TotalAmount + Amount
            BTAmountTotal = BTAmountTotal + BTAmount
        Next
    End With
    MhRealInput18.Value = BTAmountTotal
    MhRealInput19.Value = TotalAmount
End Sub
Private Sub cmdProceed_Click()
    SaveFields
    FrmBookPrintOrder.Command4.Enabled = False
    Call CloseForm(Me)
End Sub
Private Sub cmdCancel_Click()
    rstBookPOChild08.CancelUpdate
    Call CloseForm(Me)
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF9) Then
        If UserLevel = "3" Then Call DisplayError("You don't have the rights to delete Bindery Process of Item"): Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1
            fpSpread1.SetFocus
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        Call RefreshDropDownList("R")
    ElseIf Shift = 0 And KeyCode = vbKeyDelete Then
        If fpSpread1.ActiveCol = 7 Then
            fpSpread1.SetText fpSpread1.ActiveCol, fpSpread1.ActiveRow, ""  'Size Name
            fpSpread1.SetText 18, fpSpread1.ActiveRow, ""   'Size Code
        End If
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    On Error GoTo ErrorHandler
    Dim ActiveCellVal As Variant, CalcType As String, Number As Variant, Size As Variant, Qty As Variant, CalcMode As Variant, Rate As Variant, xPos As Integer, Amount As Double, Adjustment As Variant, GST As Variant, Operation As Variant, Binding As Variant, AreaRange As Variant, xCalcVal As Variant
    With fpSpread1
        If .ActiveCol <> 7 Then i7 = 0
        If .ActiveCol <> 10 Then i6 = 0
        .GetText Col, Row, ActiveCellVal
        If Col = 1 Then 'SubItem
            If Not CheckEmpty(ActiveCellVal, False) Then
                If rstElementList.RecordCount > 0 Then rstElementList.MoveFirst
                rstElementList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
                If Not rstElementList.EOF Then .SetText Col + 25, Row, rstElementList.Fields("Code").Value
            End If
        ElseIf Col = 2 Then  'Binding
            If Not CheckEmpty(ActiveCellVal, False) Then
                If rstBindingList.RecordCount > 0 Then rstBindingList.MoveFirst
                rstBindingList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
                If Not rstBindingList.EOF Then
                    .SetText Col + 22, Row, rstBindingList.Fields("Code").Value
                End If
            End If
        ElseIf Col = 3 Then 'Element
            If Not CheckEmpty(ActiveCellVal, False) Then
                If rstElementList.RecordCount > 0 Then rstElementList.MoveFirst
                rstElementList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
                If Not rstElementList.EOF Then .SetText Col + 15, Row, rstElementList.Fields("Code").Value
            End If
        ElseIf Col = 4 Then 'Operation
            If Not CheckEmpty(ActiveCellVal, False) Then
                If rstOperationList.RecordCount > 0 Then rstOperationList.MoveFirst
                rstOperationList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
                If Not rstOperationList.EOF Then
                    .SetText Col + 15, Row, rstOperationList.Fields("Code").Value 'OperationCode
                    .SetText Col + 21, Row, Val(rstOperationList.Fields("oValue1").Value)
                    .GetText 9, Row, Qty 'Quantity
                    If Val(Qty) = 0 Then .SetText 9, Row, FrmBookPrintOrder.MhRealInput3.Value 'Quantity
                End If
            End If
        ElseIf Col = 7 Then 'Size
            If Not CheckEmpty(ActiveCellVal, False) Then
                If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
                rstSizeList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
                If Not rstSizeList.EOF Then .SetText Col + 13, Row, rstSizeList.Fields("Code").Value
                i7 = i7 + 1
                If i7 = 1 Then
                    If Val(rstSizeList.Fields("STYPE").Value) = 11 Then
                        .SetText Col + 1, Row, 1
                    Else
                        Fraction = InputBox("Please Confirm Input Value for" & Chr(13) & "Pages Per Form", "Easy Publish Prime", Val(Fraction))
                        .SetText Col + 1, Row, Fraction
                    End If
                End If
            End If
        ElseIf Col = 10 Then 'Calc Mode
            If Not CheckEmpty(ActiveCellVal, False) Then
                If rstCalcModeList.RecordCount > 0 Then rstCalcModeList.MoveFirst
                rstCalcModeList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
                If Not rstCalcModeList.EOF Then
                    .SetText Col + 11, Row, rstCalcModeList.Fields("Code").Value
                    .GetText 11, .ActiveRow, xCalcVal
                    CalcVal = IIf(Val(xCalcVal) = 0, Val(rstCalcModeList.Fields("Value1").Value), Val(xCalcVal))
                End If
                i6 = i6 + 1
                If i6 = 1 Then
                    If (rstCalcModeList.Fields("Code").Value) = "*20006" Then
                        CalcVal = InputBox("Please Confirm Input Value for" & Chr(13) & "Quantity Per Packet", "Easy Publish Prime", Val(CalcVal))
                        .SetText 11, .ActiveRow, Val(CalcVal)
                    ElseIf (rstCalcModeList.Fields("Code").Value) = "*20009" Then
                        CalcVal = InputBox("Please Confirm Input Value for" & Chr(13) & "Quantity Per Box", "Easy Publish Prime", Val(CalcVal))
                        .SetText 11, .ActiveRow, Val(CalcVal)
                    ElseIf (rstCalcModeList.Fields("Code").Value) = "*20010" Then
                        CalcVal = InputBox("Please Confirm Input Value for" & Chr(13) & "Quantity Per Bundle", "Easy Publish Prime", Val(CalcVal))
                        .SetText 11, .ActiveRow, Val(CalcVal)
                    ElseIf (rstCalcModeList.Fields("Code").Value) = "*20008" Then
                        CalcVal = InputBox("Please Confirm Input Value for" & Chr(13) & "Per Paisa Inch²", "Easy Publish Prime", Val(CalcVal))
                        .SetText 11, .ActiveRow, CalcVal: .SetText 10, .ActiveRow, 0.01
                    End If
                End If
            End If
        End If
        .GetText 21, Row, CalcMode
        If Not CheckEmpty(CalcMode, False) Then
            If rstCalcModeList.RecordCount > 0 Then rstCalcModeList.MoveFirst
            rstCalcModeList.Find "[Code]='" & FixQuote(CalcMode) & "'"
            If Not rstCalcModeList.EOF Then CalcType = IIf(InStr(1, rstCalcModeList.Fields("Name").Value, "Inch") > 0, "S", "O")
            .GetText 11, Row, CalcVal
            If CalcVal = "" Or Val(CalcVal) = 0 Then CalcVal = Val(rstCalcModeList.Fields("Value1").Value): .SetText 11, .ActiveRow, Val(CalcVal)
        End If
        'Fetch Rate
        .GetText 12, Row, Rate
        If Val(Rate) = 0 Then
            .GetText 9, Row, Qty 'Quantity
            .GetText 5, Row, Number
            .GetText 19, Row, Operation
            .GetText 7, Row, AreaRange: If AreaRange <> "" Then AreaRange = Left(AreaRange, 5) * Mid(AreaRange, 7, 5)
            .GetText 20, Row, Size
            If Not (CheckEmpty(Operation, False) And CheckEmpty(CalcMode, False)) And Val(Qty) > 0 Then .SetText 12, Row, FetchOperationRate(Operation, CalcMode, IIf(CalcType = "O", Size, ""), Val(AreaRange), Val(Number), Val(Qty))
        End If
        If Col >= 4 And Col <= 15 Then
            .GetText 5, Row, Number
            .GetText 7, Row, Size
            .GetText 9, Row, Qty
            .GetText 11, Row, CalcVal
            .GetText 14, Row, Adjustment
            .GetText 15, Row, GST
            If CalcType = "S" And (Not CheckEmpty(Size, False)) Then xPos = InStr(1, LCase(Size), "x"): Size = Val(Left(Size, xPos - 1)) * Val(Mid(Size, xPos + 1, 5)) Else Size = 1
            .GetText 11, Row, CalcVal
        If CalcVal = 0 Then CalcVal = 1
            .GetText 25, Row, ActiveCellVal
        If ActiveCellVal = 1 Then
                    Amount = Round((Size * Val(Qty) * (Val(Rate)) / CalcVal), 2)
        ElseIf ActiveCellVal = 0 Then
                    Amount = Round((Number * Size * Val(Qty) * Val(Rate)) / CalcVal, 2)
        End If
'            If Operation = "*07037" Or Operation = "*07039" Or Operation = "*07051" Or Operation = "*07053" Then
'                    Amount = Round((Number * Size * Val(Qty) * Val(Rate)) / CalcVal, 2)
'            ElseIf Operation = "*07038" Then
'                    Amount = Round((Size * Val(Qty) * Val(Rate)) / CalcVal, 2)
'            Else
'                    Amount = Round((Size * Val(Qty) * Val(Rate)) / CalcVal, 2)
'            End If
            .SetText 13, Row, Amount 'Amount
            .SetText 16, Row, ((Amount + Val(Adjustment)) * Val(GST)) / 100 'GST
            .SetText 17, Row, Round(Amount + Val(Adjustment) + (((Amount + Val(Adjustment)) * Val(GST)) / 100), 0) 'BillAmount
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
        rstSubItemList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstSubItemList): Loop
        rstSubItemList.ActiveConnection = Nothing
        rstElementList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstElementList): Loop
        rstElementList.ActiveConnection = Nothing
        rstBindingList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstBindingList): Loop
        rstBindingList.ActiveConnection = Nothing
        rstOperationList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstOperationList): Loop
        rstOperationList.ActiveConnection = Nothing
        rstSizeList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstSizeList): Loop
        rstSizeList.ActiveConnection = Nothing
        rstCalcModeList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstCalcModeList): Loop
        rstCalcModeList.ActiveConnection = Nothing
        Element = "": Operation = "": Binding = "": Size = "": CalcMode = "": SubItem = ""
    End If
    With rstSubItemList
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If SubItem = "" Then SubItem = .Fields("Name").Value Else SubItem = SubItem + Chr$(9) + .Fields("Name").Value
            .MoveNext
        Loop
    End With
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
    With rstBindingList
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If Binding = "" Then Binding = .Fields("Name").Value Else Binding = Binding + Chr$(9) + .Fields("Name").Value
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
Private Function FetchOperationRate(ByVal xBinderyProcess As String, ByVal xCalcMode As String, ByVal xSize As String, xAreaRange As Double, xSectionRange As Double, xQtyRange As Double) As Double
    On Error GoTo ErrorHandler
    If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
        rstFetchOperationRate.Open "IF EXISTS (SELECT *FROM AccountMaster P INNER JOIN AccountChild07 C ON P.Code=C.Code WHERE C.Code = '" & PartyCode & "' AND BinderyProcess='" & xBinderyProcess & "' AND CalcMode='" & xCalcMode & "' AND " & IIf(CheckEmpty(xSize, False), "1=1", "[Size]='" & xSize & "'") & " AND AreaRange>=" & xAreaRange & "  AND SectionRange>=" & xSectionRange & " AND QtyRange>=" & xQtyRange & ") " & _
                                                                        "SELECT TOP 1 Rate FROM AccountChild07 WHERE Code = '" & PartyCode & "' AND BinderyProcess='" & xBinderyProcess & "' AND CalcMode='" & xCalcMode & "' AND " & IIf(CheckEmpty(xSize, False), "1=1", "[Size]='" & xSize & "'") & " AND AreaRange>=" & xAreaRange & "  AND SectionRange>=" & xSectionRange & " AND QtyRange>=" & xQtyRange & " ORDER BY AreaRange " & _
                                                        "Else " & _
                                                                        "SELECT TOP 1 Rate FROM AccountChild07 WHERE Code = '" & PartyCode & "' AND BinderyProcess='" & xBinderyProcess & "' AND CalcMode='" & xCalcMode & "' AND AreaRange>=" & xAreaRange & "  AND SectionRange>=" & xSectionRange & " AND QtyRange>=" & xQtyRange & " ORDER BY AreaRange ", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstFetchOperationRate.RecordCount = 0 Then
        If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
        rstFetchOperationRate.Open "SELECT TOP 1 Rate FROM AccountChild07 WHERE Code = '" & PartyCode & "' AND BinderyProcess='" & xBinderyProcess & "' AND CalcMode='" & xCalcMode & "' AND AreaRange>=" & xAreaRange & "  AND SectionRange>=" & xSectionRange & " AND QtyRange>=" & xQtyRange & " ORDER BY AreaRange ", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstFetchOperationRate.RecordCount = 0 Then
        If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
        
        rstFetchOperationRate.Open "IF EXISTS (SELECT *FROM AccountMaster P INNER JOIN AccountChild07 C ON P.Code=C.Code WHERE [Name] Like '%Rate%'  AND BinderyProcess='" & xBinderyProcess & "' AND CalcMode='" & xCalcMode & "' AND " & IIf(CheckEmpty(xSize, False), "1=1", "[Size]='" & xSize & "'") & " AND AreaRange>=" & xAreaRange & " AND SectionRange>=" & xSectionRange & " AND QtyRange>=" & xQtyRange & ")" & _
                                                                        "SELECT TOP 1 Rate FROM AccountMaster P INNER JOIN AccountChild07 C ON P.Code=C.Code WHERE [Name] Like '%Rate%'  AND BinderyProcess='" & xBinderyProcess & "' AND CalcMode='" & xCalcMode & "' AND " & IIf(CheckEmpty(xSize, False), "1=1", "[Size]='" & xSize & "'") & " AND AreaRange>=" & xAreaRange & " AND SectionRange>=" & xSectionRange & " AND QtyRange>=" & xQtyRange & " ORDER BY AreaRange " & _
                                                        "Else " & _
                                                                        "SELECT TOP 1 Rate FROM AccountMaster P INNER JOIN AccountChild07 C ON P.Code=C.Code WHERE [Name] Like '%Rate%'  AND BinderyProcess='" & xBinderyProcess & "' AND CalcMode='" & xCalcMode & "' AND AreaRange>=" & xAreaRange & " AND SectionRange>=" & xSectionRange & " AND QtyRange>=" & xQtyRange & " ORDER BY AreaRange", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstFetchOperationRate.RecordCount = 0 Then
        If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
        rstFetchOperationRate.Open "SELECT TOP 1 Rate FROM AccountMaster P INNER JOIN AccountChild07 C ON P.Code=C.Code WHERE [Name] Like '%Rate%'  AND BinderyProcess='" & xBinderyProcess & "' AND CalcMode='" & xCalcMode & "' AND AreaRange>=" & xAreaRange & " AND SectionRange>=" & xSectionRange & " AND QtyRange>=" & xQtyRange & " ORDER BY AreaRange", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    
    If rstFetchOperationRate.RecordCount > 0 Then FetchOperationRate = Val(rstFetchOperationRate.Fields("Rate").Value)
    Exit Function
ErrorHandler:
    DisplayError (Err.Description)
End Function
Private Sub InsertOperation()
    With rstOrderList
        If .State = adStateOpen Then .Close
        .Open "SELECT E.Code As ECode,E.Name As EName,B.Code As BCode,B.Name As BName,O.Code As OCode,O.Name As OName,O.Value1 As oValue1,T.OperationCountName As OCName,[Number],S.Code As SCode,S.Name As SName,C.Code As CCode,C.Name As CName,T.CalcValue As CalcVal FROM ((((BookChild08 T INNER JOIN ElementMaster E ON T.ElementGroup=E.Code) INNER JOIN GeneralMaster O ON T.BinderyProcess=O.Code) INNER JOIN GeneralMaster B ON T.BindingType=B.Code) INNER JOIN GeneralMaster C ON T.CalcMode=C.Code) LEFT JOIN GeneralMaster S ON T.[Size]=S.Code WHERE T.Code='" & ItemCode & "' ORDER BY E.Name,O.Name", cnDatabase, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        If .RecordCount = 0 Then
            Dim Operation, Binding, CalcMode, Size, CalcVal
            With fpSpread1
                .SetText 5, 1, 1 'Set Number
                .SetText 9, 1, OrderQty 'Set Quantity
                'Element
                If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
                rstFetchOperationRate.Open "SELECT Code,Name FROM ElementMaster WHERE Code='*00016'", cnDatabase, adOpenKeyset, adLockReadOnly
                If rstFetchOperationRate.RecordCount > 0 Then .SetText 3, 1, rstFetchOperationRate.Fields("Name").Value: .SetText 18, 1, rstFetchOperationRate.Fields("Code").Value
                'Operation,Size
                If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
                rstFetchOperationRate.Open "SELECT O.Code As OperationCode,O.Name As OperationName,S.Code As SizeCode,S.Name As SizeName FROM (BookMaster I LEFT JOIN GeneralMaster O ON I.LaminationType=O.Code) LEFT JOIN GeneralMaster S ON I.FinishSize=S.Code WHERE I.Code='" & ItemCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
                If rstFetchOperationRate.RecordCount > 0 Then .SetText 4, 1, rstFetchOperationRate.Fields("OperationName").Value: .SetText 19, 1, rstFetchOperationRate.Fields("OperationCode").Value: .SetText 25, 1, rstFetchOperationRate.Fields("oValue1").Value: .SetText 7, 1, rstFetchOperationRate.Fields("SizeName").Value: .SetText 20, 1, rstFetchOperationRate.Fields("SizeCode").Value: Operation = rstFetchOperationRate.Fields("OperationCode").Value: Size = rstFetchOperationRate.Fields("SizeCode").Value
                If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
                'CalcMode
                rstFetchOperationRate.Open "SELECT Code,Name FROM GeneralMaster WHERE Code='*20005'", cnDatabase, adOpenKeyset, adLockReadOnly
                If rstFetchOperationRate.RecordCount > 0 Then .SetText 10, 1, rstFetchOperationRate.Fields("Name").Value: .SetText 20, 1, rstFetchOperationRate.Fields("Code").Value: CalcMode = rstFetchOperationRate.Fields("Code").Value
                If rstFetchOperationRate.State = adStateOpen Then rstFetchOperationRate.Close
                rstFetchOperationRate.Open "SELECT Rate FROM AccountChild07 WHERE Code = '" & PartyCode & "' AND BinderyProcess='" & Operation & "' AND CalcMode='" & CalcMode & "' AND [Size]='" & Size & "'", cnDatabase, adOpenKeyset, adLockReadOnly
                'Rate
                If rstFetchOperationRate.RecordCount > 0 Then .SetText 12, 1, Val(rstFetchOperationRate.Fields("Rate").Value): .SetText 13, 1, OrderQty * Val(rstFetchOperationRate.Fields("Rate").Value): .SetText 17, 1, OrderQty * Val(rstFetchOperationRate.Fields("Rate").Value)
            End With
        Else
            Dim i As Integer, ECode As String, BCode As String
            Do While Not .EOF
                i = i + 1
                If ECode = "" Or ECode <> rstOrderList.Fields("ECode").Value Then fpSpread1.SetText 1, i, rstOrderList.Fields("EName").Value
                If BCode = "" Or (BCode + ECode) <> (rstOrderList.Fields("BCode").Value + rstOrderList.Fields("ECode").Value) Then fpSpread1.SetText 2, i, rstOrderList.Fields("BName").Value
                fpSpread1.SetText 4, i, rstOrderList.Fields("OName").Value
                fpSpread1.SetText 5, i, Val(rstOrderList.Fields("Number").Value)
                fpSpread1.SetText 6, i, CheckNull(rstOrderList.Fields("OCName").Value)
                fpSpread1.SetText 7, i, CheckNull(rstOrderList.Fields("SName").Value)
                fpSpread1.SetText 9, i, FrmBookPrintOrder.MhRealInput3.Value 'Quantity
                fpSpread1.SetText 10, i, CheckNull(rstOrderList.Fields("CName").Value)
                fpSpread1.SetText 11, i, Val(rstOrderList.Fields("CalcVal").Value)
                fpSpread1.SetText 18, i, rstOrderList.Fields("ECode").Value: ECode = rstOrderList.Fields("ECode").Value
                fpSpread1.SetText 19, i, rstOrderList.Fields("OCode").Value
                fpSpread1.SetText 20, i, rstOrderList.Fields("SCode").Value
                fpSpread1.SetText 21, i, rstOrderList.Fields("CCode").Value
                fpSpread1.SetText 24, i, rstOrderList.Fields("BCode").Value: BCode = rstOrderList.Fields("BCode").Value
                fpSpread1.SetText 25, i, Val(rstOrderList.Fields("oValue1").Value)
                fpSpread1.SetText 26, i, rstOrderList.Fields("ICode").Value
                .MoveNext
            Loop
        End If
    End With
End Sub
