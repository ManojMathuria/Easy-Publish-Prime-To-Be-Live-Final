VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmPaperLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "w2"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20085
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   20085
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   9270
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   19890
      _Version        =   65536
      _ExtentX        =   35084
      _ExtentY        =   16351
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
      Picture         =   "PaperLedger.frx":0000
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   3600
         TabIndex        =   23
         Top             =   8880
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
         Caption         =   " Find"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "PaperLedger.frx":001C
         Picture         =   "PaperLedger.frx":0038
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   2400
         TabIndex        =   29
         Top             =   8880
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   ""
         Alignment       =   0
         BorderColor     =   16777215
         TextColor       =   0
         Picture         =   "PaperLedger.frx":0054
         Picture         =   "PaperLedger.frx":0070
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check1"
         Height          =   255
         Left            =   14400
         TabIndex        =   28
         Top             =   158
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   8400
         TabIndex        =   27
         Top             =   158
         Visible         =   0   'False
         Width           =   255
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   8175
         Left            =   120
         TabIndex        =   4
         Top             =   645
         Width           =   19695
         _cx             =   34740
         _cy             =   14420
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   128
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   14
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   0
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   4
            Left            =   0
            Top             =   0
         End
      End
      Begin VB.CommandButton Preview 
         Caption         =   "&Print Preview"
         Height          =   330
         Left            =   15960
         TabIndex        =   26
         Top             =   8880
         Width           =   1215
      End
      Begin VB.CommandButton Search 
         Height          =   320
         Left            =   11040
         Picture         =   "PaperLedger.frx":008C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Search"
         Top             =   8880
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
         Left            =   4200
         MaxLength       =   100
         TabIndex        =   20
         ToolTipText     =   "Find And Search"
         Top             =   8880
         Width           =   6270
      End
      Begin VB.CommandButton cmdFilter 
         Height          =   320
         Left            =   10560
         Picture         =   "PaperLedger.frx":03CE
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Filter"
         Top             =   8880
         Width           =   375
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   14000
         TabIndex        =   15
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
         Caption         =   " Sort && Filter"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "PaperLedger.frx":0710
         Picture         =   "PaperLedger.frx":072C
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Billing Details"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12000
         TabIndex        =   13
         Top             =   158
         Width           =   1950
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   375
         Left            =   19380
         Picture         =   "PaperLedger.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancel"
         Top             =   90
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   375
         Left            =   19000
         Picture         =   "PaperLedger.frx":084A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Refresh"
         Top             =   90
         Width           =   375
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Show Paper Receipt Against 'PO' Only"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6450
         TabIndex        =   3
         Top             =   158
         Width           =   3375
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   8145
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   19650
         _Version        =   524288
         _ExtentX        =   34660
         _ExtentY        =   14367
         _StockProps     =   64
         ColsFrozen      =   3
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
         MaxCols         =   29
         MaxRows         =   1000
         SelectBlockOptions=   11
         ShadowColor     =   16775408
         SpreadDesigner  =   "PaperLedger.frx":0994
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   105
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
         Picture         =   "PaperLedger.frx":1A22
         Picture         =   "PaperLedger.frx":1A3E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   1800
         TabIndex        =   8
         Top             =   105
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
         Picture         =   "PaperLedger.frx":1A5A
         Picture         =   "PaperLedger.frx":1A76
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   2190
         TabIndex        =   1
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "PaperLedger.frx":1A92
         Caption         =   "PaperLedger.frx":1BAA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperLedger.frx":1C16
         Keys            =   "PaperLedger.frx":1C34
         Spin            =   "PaperLedger.frx":1C92
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
         ReadOnly        =   -1
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
         TabIndex        =   0
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "PaperLedger.frx":1CBA
         Caption         =   "PaperLedger.frx":1DD2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperLedger.frx":1E3E
         Keys            =   "PaperLedger.frx":1E5C
         Spin            =   "PaperLedger.frx":1EBA
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
         ReadOnly        =   -1
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
         Left            =   3270
         TabIndex        =   9
         Top             =   105
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2222
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
         Caption         =   " &Show Voucher"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "PaperLedger.frx":1EE2
         Picture         =   "PaperLedger.frx":1EFE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   18570
         TabIndex        =   17
         Top             =   8860
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
         Caption         =   " Print Data"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "PaperLedger.frx":1F1A
         Picture         =   "PaperLedger.frx":1F36
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   17280
         TabIndex        =   18
         Top             =   8860
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
         Caption         =   " Export Data"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "PaperLedger.frx":1F52
         Picture         =   "PaperLedger.frx":1F6E
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   330
         Left            =   1200
         TabIndex        =   21
         Top             =   8880
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "PaperLedger.frx":1F8A
         Caption         =   "PaperLedger.frx":1FAA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperLedger.frx":200E
         Keys            =   "PaperLedger.frx":202C
         Spin            =   "PaperLedger.frx":2076
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   8880
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
         Picture         =   "PaperLedger.frx":209E
         Picture         =   "PaperLedger.frx":20BA
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Show Challan Details"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9900
         TabIndex        =   10
         Top             =   158
         Width           =   2055
      End
      Begin MSForms.ComboBox Combo4 
         Height          =   330
         Left            =   11520
         TabIndex        =   24
         Top             =   8880
         Width           =   4125
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "7276;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   330
         Left            =   16650
         TabIndex        =   16
         Top             =   120
         Width           =   2300
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4039;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   15080
         TabIndex        =   14
         Top             =   120
         Width           =   1605
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2831;582"
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
         Left            =   4520
         TabIndex        =   2
         Top             =   105
         Width           =   1845
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3254;582"
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
         X2              =   19920
         Y1              =   540
         Y2              =   540
      End
   End
End
Attribute VB_Name = "FrmPaperLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sSQL, sDate As String, eDate As String, ItemList As String, PaperList As String, SupplierList As String, AccountList As String, VchType As String, LR As Integer
Dim rstPaperLedger As New ADODB.Recordset, Reset As Long
Dim nSort, VSFlexFlag, FrozenFlag As Boolean
Dim rstCompanyMaster As New ADODB.Recordset
Dim i, C As Integer
Private Sub Form_Load()
    If VchType = 19 Then VSFlexFlag = True
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    Reset = 0:
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    If VchType = 19 Then Me.Caption = "Paper Stock Ladger"
If VchType <> 19 Then
        With fpSpread1
        .Col = 1: .Row = SpreadHeader
        .UserColAction = UserColActionSort
        For C = 1 To 3
                    .ColUserSortIndicator(C) = ColUserSortIndicatorDisabled
        Next
        For C = 4 To 30
                    .ColUserSortIndicator(C) = ColUserSortIndicatorDescending
        Next
        End With
        If VchType = 11 Then
            Combo1.Clear
            Combo1.AddItem "Receipt ", 0
            Combo1.AddItem "Purchase ", 1
            Combo1.AddItem "Both", 2
            Combo1.ListIndex = 2
            
            Combo2.Clear
            Combo2.AddItem "Show IN Units ", 0
            Combo2.AddItem "Show IN Sheets ", 1
            Combo2.AddItem "Show IN Kg. ", 2
            Combo2.AddItem "All ", 3
            Combo2.ListIndex = 0
            
            Combo3.Clear
            Combo3.AddItem "Sort By Supplier Name", 0
            Combo3.AddItem "Sort By Party Name", 1
            Combo3.AddItem "Sort By Paper Name", 2
            Combo3.AddItem "Sort By Voucher No", 3
            Combo3.AddItem "Sort By All", 4
            Combo3.ListIndex = 4
        ElseIf VchType = 12 Or VchType = 13 Then
            Combo1.Clear
            Combo1.AddItem "Against Receipt ", 0
            Combo1.AddItem "Against Purchase ", 1
            Combo1.AddItem "Against Both", 2
            Combo1.ListIndex = 2
            
            Combo2.Clear
            Combo2.AddItem "Show IN Units ", 0
            Combo2.AddItem "Show IN Sheets ", 1
            Combo2.AddItem "Show IN Kg. ", 2
            Combo2.AddItem "All ", 3
            Combo2.ListIndex = 0
            
            Combo3.Clear
            Combo3.AddItem "Sort By Supplier Name", 0
            Combo3.AddItem "Sort By Party Name", 1
            Combo3.AddItem "Sort By Paper Name", 2
            Combo3.AddItem "Sort By Voucher No", 3
            Combo3.AddItem "Sort By All", 4
            Combo3.ListIndex = 4
        ElseIf VchType = 14 Or VchType = 15 Or VchType = 16 Then
            Combo1.Clear
            Combo1.AddItem "Issue ", 0
            Combo1.AddItem "Sale ", 1
            Combo1.AddItem "Both", 2
            Combo1.ListIndex = 2
        ElseIf VchType = 17 Then
            Combo1.Clear
            Combo1.AddItem "Inward ", 0
            Combo1.AddItem "Outward ", 1
            Combo1.AddItem "Both", 2
            Combo1.ListIndex = 2
            
            Combo2.Clear
            Combo2.AddItem "Show IN Units ", 0
            Combo2.AddItem "Show IN Sheets ", 1
            Combo2.AddItem "Show IN Kg. ", 2
            Combo2.AddItem "All ", 3
            Combo2.ListIndex = 0
            
            Combo3.Clear
            Combo3.AddItem "Sort By From Account Name", 0
            Combo3.AddItem "Sort By To Account Name", 1
            Combo3.AddItem "Sort By Paper Name", 2
            Combo3.AddItem "Sort By Voucher No", 3
            Combo3.AddItem "Sort By All", 4
            Combo3.ListIndex = 4
        ElseIf VchType = 18 Then
            Combo1.Clear
            Combo1.AddItem "Purchase Order ", 0
            Combo1.AddItem "Pending Purchase Order ", 1
            Combo1.AddItem "Both", 2
            Combo1.ListIndex = 2
            
            Combo2.Clear
            Combo2.AddItem "Show IN Units ", 0
            Combo2.AddItem "Show IN Sheets ", 1
            Combo2.AddItem "Show IN Kg. ", 2
            Combo2.AddItem "All ", 3
            Combo2.ListIndex = 0
            
            Combo3.Clear
            Combo3.AddItem "Sort By Supplier Name", 0
            Combo3.AddItem "Sort By Paper Name", 1
            Combo3.AddItem "Sort By Voucher No", 2
            Combo3.AddItem "Sort By All", 3
            Combo3.ListIndex = 3
    Else
            Combo4.Clear
            Combo4.AddItem "Party (Supplied From)", 0
            Combo4.AddItem "Party (Supplied To)", 1
            Combo4.AddItem "Challan No.", 2
            Combo4.AddItem "Challan Date", 3
            Combo4.AddItem "Bill No.", 4
            Combo4.AddItem "Bill Date", 5
            Combo4.AddItem "Paper Name", 6
            Combo4.AddItem "Vch. No.", 7
            Combo4.ListIndex = 0
        End If
    ElseIf VchType = 19 Then
            Combo1.Clear
            Combo1.AddItem " Sheets", 0
            Combo1.AddItem " UOM", 1
            Combo1.AddItem " KGs", 2
            Combo1.AddItem " UOM . Sheet", 3
            Combo1.ListIndex = 1
End If
    
        If VchType = 11 Then Check0.Caption = "Show Paper Receipt Against 'PO' Only": Me.Caption = "Paper Receipt Party-Wise" '11-11
        If VchType = 12 Then Combo1.Width = 2000: Check0.Value = 1: Check0.Visible = False: Me.Caption = "Paper Receipt Order-Wise" '13-12
        If VchType = 13 Then Combo1.Width = 2000: Check0.Value = 1: Check0.Visible = False: Me.Caption = "Paper Receipt Without-Order" '15-13
        If VchType = 14 Then Check0.Caption = "Show Paper Issue Against 'SO' Only": Check2.Visible = False: Me.Caption = "Paper Issue Party-Wise" '12-14
        If VchType = 15 Then Check0.Caption = "Show Paper Issue Against 'SO' Only": Check2.Visible = False: Me.Caption = "Paper Issue Order-Wise" '14-15
        If VchType = 16 Then Check0.Caption = "Show Paper Issue Against 'SO' Only": Check2.Visible = False: Me.Caption = "Paper Issue Without-Order" '16-16
        If VchType = 17 Then Check0.Caption = "Show Paper Issue Against 'SO' Only": Check2.Visible = False: Me.Caption = "Paper Transfer Party-Wise" '17
        If VchType = 18 Then Combo1.Width = 2250: Check0.Left = 7000: Check0.Caption = "Show Paper Pending 'PO' Only": Check3.Visible = True: Check2.Visible = True: Me.Caption = "Paper Pending Order Supplier-Wise" '18
        If VchType = 19 Then Combo1.Width = 1500: Check0.Value = 1: Check2.Value = 1: Check0.Left = 7000: Check0.Width = 2300: Check0.Caption = "Show Total Party-Wise": Check2.Left = 9300: Check2.Width = 2300: Check2.Visible = True: Check2.Caption = "Show Total Size-Wise": Check3.Left = 11600: Check3.Width = 2300: Check3.Visible = True: Check3.Caption = "Show Total UOM-Wise": Check1.Left = 13900: Check1.Width = 2300: Check1.Visible = True: Check1.Caption = "Show Total GSM-Wise": Check4.Left = 16200: Check4.Width = 2300: Check4.Visible = True: Check4.Caption = "Show Total Paper-Wise": Me.Caption = "Paper Stock Ledger": Combo2.Visible = False: Combo3.Visible = False: Combo4.Visible = True: Mh3dLabel4.Visible = False: Mh3dLabel1.Caption = "  A/C Statement Unit": Mh3dLabel1.Width = 1840: Combo1.Width = 1570: Combo1.Left = 5120
        Reset = 1
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
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}": KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        cmdCancel_Click
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        Call cmdRefresh_Click
        KeyCode = 0
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
            If Text1.Text = "" Then
                MsgBox "Please Provide Search Input", vbInformation
                Text1.SetFocus
            ElseIf Text1.Text <> "" Then
            Call Search_Click
            End If
        KeyCode = 0
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyC) Or (Shift = 0 And KeyCode = vbKeyF12) Then
        Call CopyToClipboard
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyV) Or (Shift = 0 And KeyCode = vbKeyF12) Then
        Call PasteFromClipboard
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstPaperLedger)
    Call CloseRecordset(rstCompanyMaster)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Sub Combo1_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Combo2_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Combo3_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub cmdRefresh_Click()
    On Error GoTo ErrHandler
    Dim SQL As String, i As Long, Stock As Long, EffectiveStock As Long
    If VchType = 11 Then 'Receipt Party-wise
                        SQL = "SELECT P.Code,IIF(P.BillNO IS Null,'',P.BillNO) AS BillNO,IIF(P.BillDate IS Null,'',P.BillDate) AS BillDate,P.Name As VchNo,(Date),(IIF(P.OrderType='R','Rpt','Pur')+'-'+Ltrim(P.Name)) As VchRef,IIF('Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ) IS Not Null,'Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ),IIF(OrderType='P','PO-'+LTRIM(P.Name),'Rpt-WPO')) As Ref,M1.Name As PartyAccountName,M2.Name As FromAccountName,IIF(BiltyNo IS NULL,'',BiltyNo) As ChallanNo,IIF(BiltyDate IS Null,'',BiltyDate) As ChallanDate,(Select Name From PaperMaster where Paper=code ) As Paper,(Select G.Name From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS UOM, " & _
                                 "C.[Weight/Unit] as wtUOM,IIF(OrderType='P'And C.Ref IS NULL,C.Quantity,0) AS Quantity,IIF(OrderType='P'And C.Ref IS NULL,C.QuantitySheets,0) As QuantitySheets,IIF(OrderType='P'And C.Ref IS NULL,C.QuantityKg,0) AS QuantityKg,C.[Units/Bundle],C.TotalBundles,IIF(C.Ref IS NOT NULL,C.Quantity,0) AS QuantityIssue,IIF(C.Ref IS NOT NULL,C.QuantitySheets,0) As IssueQtySheets,IIF(C.Ref IS NOT NULL,C.QuantityKg,0) AS IssueQtyKg FROM ((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN AccountMaster M2 ON P.Supplier=M2.Code WHERE OrderType in ('P','R') And C.Ref is not Null " & _
                                 "AND Date<=#" & GetDate(MhDateInput2.Text) & "# AND Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Supplier IN (" & SupplierList & ") AND C.Account IN (" & AccountList & ") AND C.Paper IN (" & PaperList & ") AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.OrderType,1)='R'", "RIGHT(P.OrderType,1)='P'", "1=1") & "" & _
                                 "Union " & _
                                 "SELECT P.Code,IIF(P.BillNO IS Null,'',P.BillNO) AS BillNO,IIF(P.BillDate IS Null,'',P.BillDate) AS BillDate,P.Name As VchNo,(Date),(IIF(P.OrderType='R','Rpt','Pur')+'-'+Ltrim(P.Name)) As VchRef,IIF('Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ) IS Not Null,'Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ),IIF(OrderType='P','PO-'+LTRIM(P.Name),'Rpt-WPO')) As Ref,M1.Name As PartyAccountName,M2.Name As FromAccountName,IIF(BiltyNo IS NULL,'',BiltyNo) As ChallanNo,IIF(BiltyDate IS Null,'',BiltyDate) As ChallanDate,(Select Name From PaperMaster where Paper=code ) As Paper,(Select G.Name From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS UOM, " & _
                                 "C.[Weight/Unit] as wtUOM,IIF(OrderType='P'And C.Ref IS NULL,C.Quantity,0) As Quantity,IIF(OrderType='P'And C.Ref IS NULL,C.QuantitySheets,0) AS QuantitySheets,IIF(OrderType='P'And C.Ref IS NULL,C.QuantityKg,0) ASQuantityKg,C.[Units/Bundle],C.TotalBundles,IIF(OrderType<>'P'And C.Ref IS NULL,C.Quantity,0) As QuantityIssue,IIF(OrderType<>'P'And C.Ref IS NULL,C.QuantitySheets,0) AS IssueQtySheets,IIF(OrderType<>'P'And C.Ref IS NULL,C.QuantityKg,0) AS IssueQtyKg FROM ((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN AccountMaster M2 ON P.Supplier=M2.Code WHERE OrderType in ('P','R') And C.Ref is Null " & _
                                 "AND Date<=#" & GetDate(MhDateInput2.Text) & "# AND Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Supplier IN (" & SupplierList & ") AND C.Account IN (" & AccountList & ") AND C.Paper IN (" & PaperList & ") AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.OrderType,1)='R'", "RIGHT(P.OrderType,1)='P'", "1=1") & "" & _
                                 "ORDER BY " & Choose(Combo3.ListIndex + 1, "FromAccountName,PartyAccountName,Paper,Vchref", "PartyAccountName,FromAccountName,Paper,VchNO", "Paper,FromAccountName,PartyAccountName,Vchref", "VchNO,FromAccountName,PartyAccountName,Paper", "FromAccountName,PartyAccountName,Paper,Vchref") & ""
    ElseIf VchType = 12 Then 'Receipt Order-wise
                        SQL = "SELECT P.Code,IIF(P.BillNO IS Null,'',P.BillNO) AS BillNO,IIF(P.BillDate IS Null,'',P.BillDate) AS BillDate,LTRIM(P.Name) As VchNo,(Date),(IIF(P.OrderType='R','Rpt','Pur')+'-'+Ltrim(P.Name)) As VchRef,IIF('Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ) IS Not Null,'Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ),IIF(OrderType='P','PO','Rpt-WPO')) As Ref,M1.Name As PartyAccountName,M2.Name As FromAccountName,IIF(BiltyNo IS NULL,'',BiltyNo) As ChallanNo,IIF(BiltyDate IS Null,'',BiltyDate) As ChallanDate,(Select Name From PaperMaster where Paper=code ) As Paper,(Select G.Name From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS UOM, " & _
                                 "C.[Weight/Unit] as wtUOM,IIF(OrderType='P'And C.Ref IS NULL,C.Quantity,0) AS Quantity,IIF(OrderType='P'And C.Ref IS NULL,C.QuantitySheets,0) As QuantitySheets,IIF(OrderType='P'And C.Ref IS NULL,C.QuantityKg,0) AS QuantityKg,C.[Units/Bundle],C.TotalBundles,IIF(C.Ref IS NOT NULL,C.Quantity,0) AS QuantityIssue,IIF(C.Ref IS NOT NULL,C.QuantitySheets,0) As IssueQtySheets,IIF(C.Ref IS NOT NULL,C.QuantityKg,0) AS IssueQtyKg  FROM ((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN AccountMaster M2 ON P.Supplier=M2.Code WHERE OrderType in ('P','R') And C.Ref is not Null " & _
                                 "And C.Ref is not Null AND Date<=#" & GetDate(MhDateInput2.Text) & "# AND Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Supplier IN (" & SupplierList & ") AND C.Account IN (" & AccountList & ") AND Paper IN (" & PaperList & ") AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.OrderType,1)='R'", "RIGHT(P.OrderType,1)='P'", "1=1") & "" & _
                                 "Union " & _
                                 "SELECT P.Code,IIF(P.BillNO IS Null,'',P.BillNO) AS BillNO,IIF(P.BillDate IS Null,'',P.BillDate) AS BillDate,LTRIM(P.Name) As VchNo,(Date),(IIF(P.OrderType='R','Rpt','Pur')+'-'+Ltrim(P.Name)) As VchRef,IIF('Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ) IS Not Null,'Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ),IIF(OrderType='P','PO','Rpt-WPO')) As Ref,M1.Name As PartyAccountName,M2.Name As FromAccountName,IIF(BiltyNo IS NULL,'',BiltyNo) As ChallanNo,IIF(BiltyDate IS Null,'',BiltyDate) As ChallanDate,(Select Name From PaperMaster where Paper=code ) As Paper,(Select G.Name From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS UOM, " & _
                                 "C.[Weight/Unit] as wtUOM,IIF(OrderType='P'And C.Ref IS NULL,C.Quantity,0) As Quantity,IIF(OrderType='P'And C.Ref IS NULL,C.QuantitySheets,0) AS QuantitySheets,IIF(OrderType='P'And C.Ref IS NULL,C.QuantityKg,0) ASQuantityKg,C.[Units/Bundle],C.TotalBundles,IIF(OrderType<>'P'And C.Ref IS NULL,C.Quantity,0) As QuantityIssue,IIF(OrderType<>'P'And C.Ref IS NULL,C.QuantitySheets,0) AS IssueQtySheets,IIF(OrderType<>'P'And C.Ref IS NULL,C.QuantityKg,0) AS IssueQtyKg FROM ((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN AccountMaster M2 ON P.Supplier=M2.Code WHERE OrderType in ('P','R') And C.Ref is Null " & _
                                 "And C.Ref is not Null AND Date<=#" & GetDate(MhDateInput2.Text) & "# AND Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Supplier IN (" & SupplierList & ") AND C.Account IN (" & AccountList & ") AND Paper IN (" & PaperList & ") AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.OrderType,1)='R'", "RIGHT(P.OrderType,1)='P'", "1=1") & "" & _
                                 "ORDER BY " & Choose(Combo3.ListIndex + 1, "FromAccountName,PartyAccountName,Paper,Ref", "PartyAccountName,FromAccountName,Paper,Ref", "Paper,FromAccountName,PartyAccountName,Ref", "Ref,FromAccountName,PartyAccountName,Paper", "Ref,FromAccountName,PartyAccountName,Paper") & "" 'ORDER-Wise
    ElseIf VchType = 13 Then 'Receipt Without-Order
                        SQL = "SELECT P.Code,IIF(P.BillNO IS Null,'',P.BillNO) AS BillNO,IIF(P.BillDate IS Null,'',P.BillDate) AS BillDate,P.Name As VchNo,(Date),(IIF(P.OrderType='R','Rpt','Pur')+'-'+Ltrim(P.Name)) As VchRef,IIF('Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ) IS Not Null,'Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ),IIF(OrderType='P','PO','Rpt-WPO')) As Ref,M1.Name As PartyAccountName,M2.Name As FromAccountName,IIF(BiltyNo IS NULL,'',BiltyNo) As ChallanNo,IIF(BiltyDate IS Null,'',BiltyDate) As ChallanDate,(Select Name From PaperMaster where Paper=code ) As Paper,(Select G.Name From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS UOM, " & _
                                 "C.[Weight/Unit] as wtUOM,IIF(OrderType='P'And C.Ref IS NULL,C.Quantity,0) AS Quantity,IIF(OrderType='P'And C.Ref IS NULL,C.QuantitySheets,0) As QuantitySheets,IIF(OrderType='P'And C.Ref IS NULL,C.QuantityKg,0) AS QuantityKg,C.[Units/Bundle],C.TotalBundles,IIF(C.Ref IS NOT NULL,C.Quantity,0) AS QuantityIssue,IIF(C.Ref IS NOT NULL,C.QuantitySheets,0) As IssueQtySheets,IIF(C.Ref IS NOT NULL,C.QuantityKg,0) AS IssueQtyKg  FROM ((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN AccountMaster M2 ON P.Supplier=M2.Code WHERE OrderType in ('R') And C.Ref is not Null " & _
                                 "And C.Ref is Null AND Date<=#" & GetDate(MhDateInput2.Text) & "# AND Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Supplier IN (" & SupplierList & ") AND C.Account IN (" & AccountList & ") AND Paper IN (" & PaperList & ") AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.OrderType,1)='R'", "RIGHT(P.OrderType,1)='P'", "1=1") & "" & _
                                 "Union " & _
                                 "SELECT P.Code,IIF(P.BillNO IS Null,'',P.BillNO) AS BillNO,IIF(P.BillDate IS Null,'',P.BillDate) AS BillDate,P.Name As VchNo,(Date),(IIF(P.OrderType='R','Rpt','Pur')+'-'+Ltrim(P.Name)) As VchRef,IIF('Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ) IS Not Null,'Pur-'+(Select Ltrim(Name) From PaperPOParent where c.Ref=code ),IIF(OrderType='P','PO','Rpt-WPO')) As Ref,M1.Name As PartyAccountName,M2.Name As FromAccountName,IIF(BiltyNo IS NULL,'',BiltyNo) As ChallanNo,IIF(BiltyDate IS Null,'',BiltyDate) As ChallanDate,(Select Name From PaperMaster where Paper=code ) As Paper,(Select G.Name From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS UOM, " & _
                                 "C.[Weight/Unit] as wtUOM,IIF(OrderType='P'And C.Ref IS NULL,C.Quantity,0) As Quantity,IIF(OrderType='P'And C.Ref IS NULL,C.QuantitySheets,0) AS QuantitySheets,IIF(OrderType='P'And C.Ref IS NULL,C.QuantityKg,0) ASQuantityKg,C.[Units/Bundle],C.TotalBundles,IIF(OrderType<>'P'And C.Ref IS NULL,C.Quantity,0) As QuantityIssue,IIF(OrderType<>'P'And C.Ref IS NULL,C.QuantitySheets,0) AS IssueQtySheets,IIF(OrderType<>'P'And C.Ref IS NULL,C.QuantityKg,0) AS IssueQtyKg FROM ((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN AccountMaster M2 ON P.Supplier=M2.Code WHERE OrderType in ('R') And C.Ref is Null " & _
                                 "And C.Ref is Null AND Date<=#" & GetDate(MhDateInput2.Text) & "# AND Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Supplier IN (" & SupplierList & ") AND C.Account IN (" & AccountList & ") AND Paper IN (" & PaperList & ") AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.OrderType,1)='R'", "RIGHT(P.OrderType,1)='P'", "1=1") & "" & _
                                 "ORDER BY " & Choose(Combo3.ListIndex + 1, "FromAccountName,PartyAccountName,Paper,VchNo", "PartyAccountName,FromAccountName,Paper,VchNo", "Paper,FromAccountName,PartyAccountName,Ref", "VchNo,FromAccountName,PartyAccountName,Paper", "VchNo,FromAccountName,PartyAccountName,Paper") & ""
    ElseIf VchType = 14 Or VchType = 15 Or VchType = 16 Then
                        SQL = "SELECT P.Code,IIF(P.BiltyNo IS Null,'',P.BiltyNo) AS BillNO,IIF(P.BiltyDate IS Null,'',P.BiltyDate) AS BillDate,P.Name As VchNo,(Date),(IIF(P.Type='I','Isuue','Sale')+'-'+Ltrim(P.Name)) As VchRef,IIF('Outward-'+(Select Ltrim(Name) From PaperMVParent where C.Code=code ) IS Not Null,'Outward-'+(Select Ltrim(Name) From PaperMVParent where C.Code=code ),IIF(Type='S','SO','Ise-WSO')) As Ref,M1.Name As PartyAccountName,M2.Name As FromAccountName,IIF(BiltyNo IS NULL,'',BiltyNo) As ChallanNo,IIF(BiltyDate IS Null,'',BiltyDate) As ChallanDate,(Select Name From PaperMaster where Paper=code ) As Paper,(Select G.Name From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS UOM, " & _
                                 "C.[Weight/Unit] as wtUOM,C.Quantity,C.QuantitySheets,C.QuantityKg,C.[Units/Bundle],C.TotalBundles  FROM ((PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.AccountTO=M1.Code) INNER JOIN AccountMaster M2 ON P.AccountFrom=M2.Code WHERE Type in ('S','I') And C.Code is not Null " & _
                                 "AND Date<=#" & GetDate(MhDateInput2.Text) & "# AND Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.AccountFrom IN (" & SupplierList & ") AND P.AccountTo IN (" & AccountList & ") AND Paper IN (" & PaperList & ") AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='I'", "RIGHT(P.Type,1)='S'", "1=1") & "" & _
                                 "Union " & _
                                 "SELECT P.Code,IIF(P.BiltyNo IS Null,'',P.BiltyNo) AS BillNO,IIF(P.BiltyDate IS Null,'',P.BiltyDate) AS BillDate,P.Name As VchNo,(Date),(IIF(P.Type='I','Isuue','Sale')+'-'+Ltrim(P.Name)) As VchRef,IIF('Outward-'+(Select Ltrim(Name) From PaperMVParent where C.Code=code ) IS Not Null,'Outward-'+(Select Ltrim(Name) From PaperMVParent where C.Code=code ),IIF(Type='S','SO','Ise-WSO')) As Ref,M1.Name As PartyAccountName,M2.Name As FromAccountName,IIF(BiltyNo IS NULL,'',BiltyNo) As ChallanNo,IIF(BiltyDate IS Null,'',BiltyDate) As ChallanDate,(Select Name From PaperMaster where Paper=code ) As Paper,(Select G.Name From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS UOM, " & _
                                 "C.[Weight/Unit] as wtUOM,C.Quantity,C.QuantitySheets,C.QuantityKg,C.[Units/Bundle],C.TotalBundles  FROM ((PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.AccountTo=M1.Code) INNER JOIN AccountMaster M2 ON P.AccountFrom=M2.Code WHERE Type in ('S','I') And C.Code is Null " & _
                                 "AND Date<=#" & GetDate(MhDateInput2.Text) & "# AND Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.AccountFrom IN (" & SupplierList & ") AND P.AccountTo IN (" & AccountList & ") AND Paper IN (" & PaperList & ") AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='I'", "RIGHT(P.Type,1)='S'", "1=1") & "" & _
                                 "ORDER BY M1.Name,Paper,Vchref"
    ElseIf VchType = 17 Then
                        SQL = "SELECT P.Code,IIF(P.BiltyNo IS Null,'',P.BiltyNo) AS BillNO,IIF(P.BiltyDate IS Null,'',P.BiltyDate) AS BillDate,P.Name As VchNo,(Date),(IIF(P.Type='T','Transfer','Issue')+'-'+Ltrim(P.Name)) As VchRef,IIF('Outward-'+(Select Ltrim(Name) From PaperMVParent where C.Code=code ) IS Not Null,'Outward-'+(Select LTRIM(Name) From PaperMVParent where C.Code=code ),IIF(Type='T','Tranfer','Issue-WSO')) As Ref,M1.Name As PartyAccountName,M2.Name As FromAccountName,IIF(BiltyNo IS NULL,'',BiltyNo) As ChallanNo,IIF(BiltyDate IS Null,'',BiltyDate) As ChallanDate,(Select Name From PaperMaster where Paper=code ) As Paper,(Select G.Name From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS UOM, " & _
                                 "C.[Weight/Unit] as wtUOM,C.Quantity,C.QuantitySheets,C.QuantityKg,C.[Units/Bundle],C.TotalBundles  FROM ((PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.AccountTO=M1.Code) INNER JOIN AccountMaster M2 ON P.AccountFrom=M2.Code WHERE Type in ('T') And C.Code is not Null " & _
                                 "AND Date<=#" & GetDate(MhDateInput2.Text) & "# AND Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.AccountFrom IN (" & SupplierList & ") AND P.AccountTo IN (" & AccountList & ") AND Paper IN (" & PaperList & ") AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='T'", "RIGHT(P.Type,1)='S'", "1=1") & "" & _
                                 "Union " & _
                                 "SELECT P.Code,IIF(P.BiltyNo IS Null,'',P.BiltyNo) AS BillNO,IIF(P.BiltyDate IS Null,'',P.BiltyDate) AS BillDate,P.Name As VchNo,(Date),(IIF(P.Type='T','Transfer','Issue')+'-'+Ltrim(P.Name)) As VchRef,IIF('Outward-'+(Select Ltrim(Name) From PaperMVParent where C.Code=code ) IS Not Null,'Outward-'+(Select LTRIM(Name) From PaperMVParent where C.Code=code ),IIF(Type='T','Transfer','Issue-WSO')) As Ref,M1.Name As PartyAccountName,M2.Name As FromAccountName,IIF(BiltyNo IS NULL,'',BiltyNo) As ChallanNo,IIF(BiltyDate IS Null,'',BiltyDate) As ChallanDate,(Select Name From PaperMaster where Paper=code ) As Paper,(Select G.Name From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS UOM, " & _
                                 "C.[Weight/Unit] as wtUOM,C.Quantity,C.QuantitySheets,C.QuantityKg,C.[Units/Bundle],C.TotalBundles  FROM ((PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.AccountTo=M1.Code) INNER JOIN AccountMaster M2 ON P.AccountFrom=M2.Code WHERE Type in ('T') And C.Code is Null " & _
                                 "AND Date<=#" & GetDate(MhDateInput2.Text) & "# AND Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.AccountFrom IN (" & SupplierList & ") AND P.AccountTo IN (" & AccountList & ") AND Paper IN (" & PaperList & ") AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='T'", "RIGHT(P.Type,1)='S'", "1=1") & "" & _
                                 "ORDER BY " & Choose(Combo3.ListIndex + 1, "FromAccountName,PartyAccountName,Paper,VchNo", "PartyAccountName,FromAccountName,Paper,VchNo", "Paper,FromAccountName,PartyAccountName,Ref", "VchNo,FromAccountName,PartyAccountName,Paper", "VchNo,FromAccountName,PartyAccountName,Paper") & ""
    ElseIf VchType = 18 Then
                        SQL = "SELECT P.Code,IIF(P.BillNO IS Null,'',P.BillNO) AS BillNO,IIF(P.BillDate IS Null,'',P.BillDate) AS BillDate,P.Name As VchNo,LTRIM(Date) As Date,(IIF(P.OrderType='R','Rpt','Pur')+'-'+Ltrim(P.Name)) As VchRef,'' As Ref,'' As PartyAccountName,(Select Name From AccountMaster Where P.Supplier=Code )AS FromAccountName,IIF(BiltyNo IS NULL,'',BiltyNo) As ChallanNo,IIF(BiltyDate IS Null,'',BiltyDate) As ChallanDate,(Select Name From PaperMaster where Paper=code ) As Paper,(Select G.Name From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS UOM," & _
                                 "C.[Weight/Unit] as wtUOM,SUM(ABS(C.QuantitySheets))/(Select G.Value1 From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS Quantity,SUM(ABS(C.QuantitySheets)) AS QuantitySheets,SUM(ABS(C.QuantityKg))As QuantityKg,C.[Units/Bundle],C.TotalBundles,C.[Rate/Kg],C.[Rate/Unit],C.[Amount],(Select G.Value1 From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) AS Units, " & _
                                 "(Select ISNULL(sum(QuantitySheets),0) From PaperIOChild where Ref+Paper=C.Code+C.Paper)/(Select G.Value1 From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) as QuantityIssue,(Select ISNULL(sum(QuantitySheets),0) From PaperIOChild where Ref+Paper=C.Code+C.Paper) AS IssueQtySheets,(Select ISNULL(sum(QuantityKg),0) From PaperIOChild where Ref+Paper=C.Code+C.Paper) AS IssueQtyKg, " & _
                                 "(SUM(ABS(C.QuantitySheets))-(Select ISNULL(sum(QuantitySheets),0) From PaperIOChild where Ref+Paper=C.Code+C.Paper))/(Select G.Value1 From GeneralMaster G Inner Join PaperMaster M ON M.UOM=G.Code Where C.Paper=M.Code ) As PendingQty,SUM(ABS(C.QuantitySheets))-(Select ISNULL(sum(QuantitySheets),0) From PaperIOChild where Ref+Paper=C.Code+C.Paper) As PendingQtySheets,SUM(ABS(C.QuantityKg))-(Select ISNULL(sum(QuantityKg),0) From PaperIOChild where Ref+Paper=C.Code+C.Paper) As PendingQtyKG " & _
                                 "FROM (PaperPOParent P INNER JOIN PaperPOChild C ON P.Code=C.Code)INNER JOIN AccountMaster M2 ON P.Supplier=M2.Code " & _
                                 "AND Date<=#" & GetDate(MhDateInput2.Text) & "# AND Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Supplier IN (" & SupplierList & ") AND Paper IN (" & PaperList & ") AND " & Choose(Combo1.ListIndex + 1, "(Select ISNULL(sum(QuantitySheets),0) From PaperIOChild where Ref+Paper=C.Code+C.Paper)>=0 ", "(Select ISNULL(sum(QuantitySheets),0) From PaperIOChild where Ref+Paper=C.Code+C.Paper)=0 ", "1=1 ") & "" & _
                                 "Group By C.Code,C.Paper,P.Code,OrderType,[Weight/Unit],P.Supplier,P.Name,BillNO,P.BillDate,BiltyNo,BiltyDate,[Units/Bundle],TotalBundles,(Date),C.[Rate/Kg],C.[Rate/Unit],C.[Amount] " & _
                                 "ORDER BY " & Choose(Combo3.ListIndex + 1, "FromAccountName", "Paper", "VchNO", "FromAccountName,Paper,VchNO") & ""
    ElseIf VchType = 19 Then
                            SQL = sSQL 'Paper Stock Ledger from Paper Stock Register
    End If
    
    If DatabaseType = "MS SQL" Then SQL = Replace(SQL, "#", "'")
    Screen.MousePointer = vbHourglass   'vbNormal
    If rstPaperLedger.State = adStateOpen Then rstPaperLedger.Close
    rstPaperLedger.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
    If rstPaperLedger.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
    
    If VchType = 19 Then VSFlexFlag = True: GoTo NXT:
    With fpSpread1
        If VchType = 11 Or VchType = 12 Or VchType = 13 Then
            .ColWidth(2) = 28: .ColWidth(8) = 52.75: .ColWidth(11) = 7: .ColWidth(14) = 8.25: .ColWidth(13) = 8
            .Col = 4: .ColHidden = True
            .Col = 5: .ColHidden = True
            .Col = 6: .ColHidden = True
            .Col = 7: .ColHidden = True
            .Col = 12: .ColHidden = True
            .Col = 13: .ColHidden = True
            .Col = 14: .ColHidden = True
            .Col = 15: .ColHidden = True
            .Col = 16: .ColHidden = True
            .Col = 17: .ColHidden = True
            .Col = 18: .ColHidden = True
            .Col = 19: .ColHidden = True 'Qty[IN]
            .Col = 20: .ColHidden = True 'UOM
            .Col = 21: .ColHidden = True 'QtySheets[IN]
            .Col = 22: .ColHidden = True 'QtyKg[IN]
            .Col = 23: .ColHidden = True 'Qty[Pending]
            .Col = 24: .ColHidden = True 'UOM
            .Col = 25: .ColHidden = True 'QtySheets[Pending]
            .Col = 26: .ColHidden = True 'QtyKg[Pending]
            .Col = 27: .ColHidden = False 'Rate/Kg
            .Col = 28: .ColHidden = False 'Rate/Unit
            .Col = 29: .ColHidden = False 'Amount
            If Combo2.ListIndex = 0 Then .ColWidth(2) = 25: .ColWidth(8) = 55.25: .ColWidth(24) = 10: .Col = 13: .ColHidden = False: .Col = 19: .ColHidden = False: .Col = 20: .ColHidden = False
            If Combo2.ListIndex = 1 Then .ColWidth(2) = 29: .ColWidth(8) = 55: .ColWidth(24) = 10: .Col = 15: .ColHidden = False: .Col = 21: .ColHidden = False: .Col = 25: .ColHidden = True
            If Combo2.ListIndex = 2 Then .ColWidth(2) = 29.75: .ColWidth(8) = 54: .ColWidth(24) = 10: .Col = 16: .ColHidden = False: .Col = 22: .ColHidden = False: .Col = 26: .ColHidden = True
            If Combo2.ListIndex = 3 Then .ColWidth(2) = 29.5: .ColWidth(8) = 50.5: .ColWidth(24) = 10: .ColWidth(29) = 13.25: .Col = 13: .ColHidden = False: .Col = 19: .ColHidden = False: .Col = 14: .ColHidden = False: .Col = 15: .ColHidden = False: .Col = 21: .ColHidden = False: .Col = 25: .ColHidden = True: .Col = 16: .ColHidden = False: .Col = 22: .ColHidden = False: .Col = 26: .ColHidden = True
            If VchType = 13 Then
            If Combo2.ListIndex = 0 Then .ColWidth(2) = 25: .ColWidth(8) = 58.75: .ColWidth(24) = 10: .Col = 13: .ColHidden = True: .Col = 19: .ColHidden = False: .Col = 20: .ColHidden = False
            If Combo2.ListIndex = 1 Then .ColWidth(2) = 28: .ColWidth(3) = 28: .ColWidth(8) = 61.5: .ColWidth(24) = 10: .Col = 15: .ColHidden = True: .Col = 21: .ColHidden = False: .Col = 25: .ColHidden = True
            If Combo2.ListIndex = 2 Then .ColWidth(2) = 30: .ColWidth(8) = 59.75: .ColWidth(24) = 10: .Col = 16: .ColHidden = True: .Col = 22: .ColHidden = False: .Col = 26: .ColHidden = True
            If Combo2.ListIndex = 3 Then .ColWidth(2) = 25: .ColWidth(8) = 58.75: .Col = 13: .ColHidden = True: .Col = 14: .ColHidden = True: .Col = 15: .ColHidden = True: .Col = 16: .ColHidden = True: .Col = 19: .ColHidden = False: .Col = 20: .ColHidden = False: .Col = 21: .ColHidden = False: .Col = 25: .ColHidden = False: .Col = 22: .ColHidden = False: .Col = 26: .ColHidden = True
            End If
        ElseIf VchType = 14 Or VchType = 15 Or VchType = 16 Or VchType = 17 Then
            .Col = 4: .ColHidden = True 'ChallanNo
            .Col = 5: .ColHidden = True 'ChallanDate
            .Col = 6: .ColHidden = True 'BillNO
            .Col = 7: .ColHidden = True 'BillDate
            .Col = 12: .ColHidden = True
            .Col = 13: .ColHidden = True
            .Col = 14: .ColHidden = True
            .Col = 15: .ColHidden = True
            .Col = 16: .ColHidden = True
            .Col = 17: .ColHidden = True
            .Col = 18: .ColHidden = True
            .Col = 19: .ColHidden = True 'Qty[IN]
            .Col = 20: .ColHidden = True 'UOM
            .Col = 21: .ColHidden = True 'QtySheets[IN]
            .Col = 22: .ColHidden = True 'QtyKg[IN]
            .Col = 23: .ColHidden = True 'Qty[Pending]
            .Col = 24: .ColHidden = True 'UOM
            .Col = 25: .ColHidden = True 'QtySheets[Pending]
            .Col = 26: .ColHidden = True 'QtyKg[Pending]
            .Col = 27: .ColHidden = False 'Rate/Kg
            .Col = 28: .ColHidden = False 'Rate/Unit
            .Col = 29: .ColHidden = False 'Amount
            If VchType = 17 Then
            If Combo2.ListIndex = 0 Then .ColWidth(2) = 25: .ColWidth(8) = 58.25: .ColWidth(9) = 10: .ColWidth(24) = 10: .Col = 13: .ColHidden = False: .Col = 14: .ColHidden = False:
            If Combo2.ListIndex = 1 Then .ColWidth(2) = 28: .ColWidth(3) = 28: .ColWidth(8) = 61.5: .ColWidth(24) = 10: .Col = 15: .ColHidden = True: .Col = 21: .ColHidden = False: .Col = 25: .ColHidden = True
            If Combo2.ListIndex = 2 Then .ColWidth(2) = 30: .ColWidth(8) = 59.75: .ColWidth(24) = 10: .Col = 16: .ColHidden = True: .Col = 22: .ColHidden = False: .Col = 26: .ColHidden = True
            If Combo2.ListIndex = 3 Then .ColWidth(2) = 25: .ColWidth(8) = 58.75: .Col = 13: .ColHidden = True: .Col = 14: .ColHidden = True: .Col = 15: .ColHidden = True: .Col = 16: .ColHidden = True: .Col = 19: .ColHidden = False: .Col = 20: .ColHidden = False: .Col = 21: .ColHidden = False: .Col = 25: .ColHidden = False: .Col = 22: .ColHidden = False: .Col = 26: .ColHidden = True
            End If
        ElseIf VchType = 18 Then
            .Col = 3: .ColHidden = True 'Supplier
            .Col = 4: .ColHidden = True 'ChallanNo
            .Col = 5: .ColHidden = True 'ChallanDate
            .Col = 6: .ColHidden = True 'BillNO
            .Col = 7: .ColHidden = True 'BillDate
            .Col = 10: .ColHidden = True 'Date
            .Col = 11: .ColHidden = True 'Ref
            .Col = 13: .ColHidden = True 'Qty[PO]
            .Col = 14: .ColHidden = True 'UOM
            .Col = 15: .ColHidden = True 'QtySheets[PO]
            .Col = 16: .ColHidden = True 'QtyKg[PO]
            .Col = 17: .ColHidden = True 'Units/Bundle
            .Col = 18: .ColHidden = True 'TotalBundles
            .Col = 19: .ColHidden = True 'Qty[IN]
            .Col = 20: .ColHidden = True 'UOM
            .Col = 21: .ColHidden = True 'QtySheets[IN]
            .Col = 22: .ColHidden = True 'QtyKg[IN]
            .Col = 23: .ColHidden = True 'Qty[Pending]
            .Col = 24: .ColHidden = True 'UOM
            .Col = 25: .ColHidden = True 'QtySheets[Pending]
            .Col = 26: .ColHidden = True 'QtyKg[Pending]
            .Col = 27: .ColHidden = False 'Rate/Kg
            .Col = 28: .ColHidden = False 'Rate/Unit
            .Col = 29: .ColHidden = False 'Amount
            If Combo2.ListIndex = 0 Then .ColWidth(2) = 30: .ColWidth(8) = 56.75: .ColWidth(24) = 10: .Col = 13: .ColHidden = False: .Col = 14: .ColHidden = True: .Col = 19: .ColHidden = False: .Col = 20: .ColHidden = True: .Col = 23: .ColHidden = False: .Col = 24: .ColHidden = False
            If Combo2.ListIndex = 1 Then .ColWidth(2) = 29: .ColWidth(8) = 57.25: .ColWidth(24) = 10: .Col = 15: .ColHidden = False: .Col = 21: .ColHidden = False: .ColHidden = False: .Col = 25: .ColHidden = False
            If Combo2.ListIndex = 2 Then .ColWidth(2) = 30.25: .ColWidth(8) = 58.5: .ColWidth(24) = 10: .Col = 16: .ColHidden = False: .Col = 22: .ColHidden = False: .ColHidden = False: .Col = 26: .ColHidden = False
            If Combo2.ListIndex = 3 Then .ColWidth(2) = 29.5: .ColWidth(8) = 56.75: .ColWidth(24) = 10: .ColWidth(29) = 13.25:  .Col = 13: .ColHidden = False: .Col = 14: .ColHidden = False: .Col = 19: .ColHidden = False: .Col = 20: .ColHidden = True: .ColHidden = False: .Col = 24: .ColHidden = True: .Col = 26: .ColHidden = False: .Col = 15: .ColHidden = False: .Col = 21: .ColHidden = False: .ColHidden = False: .Col = 25: .ColHidden = False: .Col = 16: .ColHidden = False: .Col = 25: .ColHidden = False: .ColHidden = False: .Col = 26: .ColHidden = False
        ElseIf VchType = 4 Then
        ElseIf VchType = 5 Then
        ElseIf VchType = 6 Then
        ElseIf VchType = 7 Then
        ElseIf VchType = 8 Then
        ElseIf VchType = 9 Then
        ElseIf VchType = 10 Then
        End If
        
        If VchType = 11 Or VchType = 12 Or VchType = 18 Then
        Call Check0_Click
        Call Check2_Click
        Call Check3_Click
        End If
        
        .ClearRange 1, 1, .MaxCols, .MaxRows, True
        rstPaperLedger.MoveFirst
        Do While Not rstPaperLedger.EOF
                i = i + 1
                .SetText 1, i, i 'S.No.
                .SetText 2, i, rstPaperLedger.Fields("FromAccountName").Value: If Len(rstPaperLedger.Fields("FromAccountName").Value) > 25 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: fpSpread1.TypeHAlign = TypeHAlignRight
                .SetText 3, i, rstPaperLedger.Fields("PartyAccountName").Value: If Len(rstPaperLedger.Fields("PartyAccountName").Value) > 25 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: fpSpread1.TypeHAlign = TypeHAlignRight
                .SetText 4, i, rstPaperLedger.Fields("ChallanNo").Value
                .SetText 5, i, rstPaperLedger.Fields("ChallanDate").Value
                .SetText 6, i, rstPaperLedger.Fields("BillNO").Value
                .SetText 7, i, rstPaperLedger.Fields("BillDate").Value
                .SetText 8, i, rstPaperLedger.Fields("Paper").Value: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstPaperLedger.Fields("Paper").Value) > 75 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignRight
                .SetText 9, i, rstPaperLedger.Fields("VchRef").Value
                .SetText 10, i, rstPaperLedger.Fields("Date").Value
                .SetText 11, i, rstPaperLedger.Fields("Ref").Value
                .SetText 12, i, val(rstPaperLedger.Fields("wtUOM").Value)
                .SetText 13, i, IIf(val(rstPaperLedger.Fields("Quantity").Value) = 0, "", val(rstPaperLedger.Fields("Quantity").Value))
                .SetText 14, i, rstPaperLedger.Fields("UOM").Value
                .SetText 15, i, IIf(val(rstPaperLedger.Fields("QuantitySheets").Value) = 0, "", val(rstPaperLedger.Fields("QuantitySheets").Value))
                .SetText 16, i, IIf(val(rstPaperLedger.Fields("QuantityKg").Value) = 0, "", val(rstPaperLedger.Fields("QuantityKg").Value))
                .SetText 17, i, val(rstPaperLedger.Fields("Units/Bundle").Value)
                .SetText 18, i, val(rstPaperLedger.Fields("TotalBundles").Value)
            If VchType = 11 Or VchType = 12 Or VchType = 13 Or VchType = 18 Or VchType = 19 Then
                .SetText 19, i, IIf(val(rstPaperLedger.Fields("QuantityIssue").Value) = 0, "", val(rstPaperLedger.Fields("QuantityIssue").Value))
                .SetText 20, i, rstPaperLedger.Fields("UOM").Value
                .SetText 21, i, IIf(val(rstPaperLedger.Fields("IssueQtySheets").Value) = 0, "", val(rstPaperLedger.Fields("IssueQtySheets").Value))
                .SetText 22, i, IIf(val(rstPaperLedger.Fields("IssueQtyKg").Value) = 0, "", val(rstPaperLedger.Fields("IssueQtyKg").Value))
            End If
            If VchType = 18 Or VchType = 19 Then
                .SetText 23, i, IIf(val(rstPaperLedger.Fields("PendingQty").Value) = 0, "", val(rstPaperLedger.Fields("PendingQty").Value))
                .SetText 24, i, rstPaperLedger.Fields("UOM").Value
                .SetText 25, i, IIf(val(rstPaperLedger.Fields("PendingQtySheets").Value) = 0, "", val(rstPaperLedger.Fields("PendingQtySheets").Value))
                .SetText 26, i, IIf(val(rstPaperLedger.Fields("PendingQtyKG").Value) = 0, "", val(rstPaperLedger.Fields("PendingQtyKG").Value))
                .SetText 27, i, val(rstPaperLedger.Fields("Rate/Kg").Value)
                .SetText 28, i, val(rstPaperLedger.Fields("Rate/Unit").Value)
                .SetText 29, i, val(rstPaperLedger.Fields("Amount").Value)
            End If
            rstPaperLedger.MoveNext
        Loop
    .LockBackColor = RGB(245, 250, 250): Combo1.BackColor = RGB(245, 250, 250): Combo2.BackColor = RGB(245, 250, 250): Combo3.BackColor = RGB(245, 250, 250): MhDateInput1.BackColor = RGB(245, 250, 250): MhDateInput2.BackColor = RGB(245, 250, 250): 'TDBNumber1.BackColor = RGB(245, 250, 250): TDBNumber2.BackColor = RGB(245, 250, 250): Text1.BackColor = RGB(245, 250, 250):
    End With
TDBNumber2 = i
NXT:
Call Print_Grid
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
Private Sub Check0_Click() 'Show Paper PO
    Dim PO As Variant, i As Long
If VSFlexFlag Then
If Reset = 1 Then Call VSFlexGrid1_AfterDataRefresh
Else
If VchType = 18 Or VchType = 19 Then
        With fpSpread1
        For i = 1 To .DataRowCnt
            If Check0.Value Then
                .GetText 26, i, PO
                If PO = 0 Then .Row = i: .RowHidden = True Else .Row = i: .RowHidden = False
            ElseIf Check0.Value Then
                .GetText 26, i, PO
                If PO > 0 Then .Row = i: .RowHidden = False Else .Row = i: .RowHidden = True
            Else
                .Row = i: .RowHidden = False
            End If
        Next
    End With
Else
    With fpSpread1
        For i = 1 To .DataRowCnt
            If Check0.Value Then
                .GetText 11, i, PO
                If Left(PO, 3) = "Pur" Then .Row = i: .RowHidden = False Else .Row = i: .RowHidden = True
            ElseIf Check0.Value Then
                .GetText 11, i, PO
                If Left(PO, 4) = "Sale" Then .Row = i: .RowHidden = False Else .Row = i: .RowHidden = True
            Else
                .Row = i: .RowHidden = False
            End If
        Next
    End With
End If
End If
End Sub
Private Sub Check2_Click() 'Show Bill Details
If VSFlexFlag Then
If Reset = 1 Then Call VSFlexGrid1_AfterDataRefresh
Else
    With fpSpread1
            If Check2.Value Then
                .Col = 6: .ColHidden = False
                .Col = 7: .ColHidden = False
            Else
                .Col = 6: .ColHidden = True
                .Col = 7: .ColHidden = True
            End If
    End With
End If
End Sub
Private Sub Check3_Click() 'Show Challan Details
If VSFlexFlag Then
If Reset = 1 Then Call VSFlexGrid1_AfterDataRefresh
Else
    With fpSpread1
            If Check3.Value Then
                .Col = 4: .ColHidden = False
                .Col = 5: .ColHidden = False
            Else
                .Col = 4: .ColHidden = True
                .Col = 5: .ColHidden = True
            End If
    End With
End If
End Sub
Private Sub Check1_Click() 'Show GSM TOTAL
If VSFlexFlag Then
    If Reset = 1 Then Call VSFlexGrid1_AfterDataRefresh
End If
End Sub
Private Sub Check4_Click() 'Show Paper TOTAL
If VSFlexFlag Then
    If Reset = 1 Then Call VSFlexGrid1_AfterDataRefresh
End If
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub Mh3dLabel5_Click()
Dim PrintHeader As String
    On Error GoTo ErrHandler
Screen.MousePointer = vbHourglass

' These are 8.25" X 11.75" paper dimensions in TWIPS
Const PaperWidth = 12240
Const PaperHeight = 15840
Printer.PaperSize = vbPRPSA4

PrintHeader = "Export Data Company : " & rstCompanyMaster.Fields("PrintName").Value & " _(" & CompCode & "_" & PrintHeader & ")" & "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]" & " /rPage # ./p " & " Print Date : ( " & Format(Date, "dd-MMM-yyyy") & " )         "
If VSFlexFlag = True Then
    With Me.VSFlexGrid1
    
    .PrintGrid PrintHeader, True, PrintOrientationPortrait, 50, 50
    
    End With
Else
With fpSpread1
    .LockBackColor = vbWhite
    ' Set printing options for sheet
    .PrintAbortMsg = "Printing - Click Cancel to .Quit"
    .PrintJobName = "Export Data" & "(" & CompCode & "_Vch-" & VchType & ")" & Format(Date, "dd-MMM-yyyy")
    .PrintHeader = "" ' "/cPrint Header/rPage # ./p/n2nd Line"
    .PrintFooter = "" ' "/cPrint Footer/rPage # ./p/n2nd Line"
    .PrintBorder = True
    .PrintColHeaders = True
    .PrintColor = True
    .PrintGrid = True
    .PrintMarginTop = 1000 '1440
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
 End With
End If
    Screen.MousePointer = vbNormal
Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
Private Sub Mh3dLabel6_Click()
Dim x As Boolean, FileName As String, SheetName As String, LogFileName As String
Dim R As Long, C As Long
On Error Resume Next
Screen.MousePointer = vbHourglass
If VSFlexFlag = True Then
    With Me.VSFlexGrid1
        If Dir(App.Path & "\Export", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Export"
        FileName = App.Path & "\Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"
        SheetName = "Sheet1" '"(" & Me.Caption & ")"
        .SaveGrid FileName, flexFileExcel, SaveExcelSettings.flexXLSaveFixedCells
        If Dir(FileName, vbDirectory) <> "" Then x = True
        If x = True Then
        MsgBox "Export complete.", vbInformation, "Easy Publish...Export !!! "
        Shell "C:\WINDOWS\explorer.exe """ & FileName & "", vbNormalFocus
        Else
        MsgBox "Export did not succeed.", vbInformation, "Easy Publish...Export !!!"
        End If
    End With
ElseIf VSFlexFlag = False Then
    With fpSpread1
        If VchType <= 10 And VchType >= 7 Or VchType <= 28 And VchType >= 25 Then fpSpread1.InsertRows 1, 2 Else fpSpread1.InsertRows 1, 1
            R = 1
        For C = 1 To .MaxCols
            .Col = C: .Row = R: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: '.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
        Next
            .SetText 1, 1, "S.NO.": .SetText 2, 1, "Party (Supplied From)": .SetText 3, 1, "Party (Supplied To)": .SetText 4, 1, "Challan No.": .SetText 5, 1, "Challan Date": .SetText 6, 1, "Bill No.": .SetText 7, 1, "Bill Date": .SetText 8, 1, "Paper Name": .SetText 9, 1, "Vch. No.": .SetText 10, 1, "Vch. Date": .SetText 11, 1, "Vch. Ref.": .SetText 12, 1, "Weight/Unit": .SetText 13, 1, "Quantity": .SetText 14, 1, "Unit": .SetText 15, 1, "Quantity (Sheets)": .SetText 16, 1, "Quantity (Kgs.)": .SetText 17, 1, "Bundles": .SetText 18, 1, "Total Bundles": .SetText 19, 1, "Quantity (IN)": .SetText 20, 1, " Unit": .SetText 21, 1, "Quantity IN (Sheets)": .SetText 22, 1, "Quantity IN (Kgs.)": .SetText 23, 1, "Pending Qty.": .SetText 24, 1, " Unit": .SetText 25, 1, "Pending Qty. (Sheets)": .SetText 26, 1, "Pending Qty. (Kgs.)": .SetText 27, 1, "Rate/Kg.": .SetText 28, 1, "Rate/Unit": .SetText 29, 1, "Amount": .ColHeadersShow = True
            .PrintColHeaders = True: .PrintRowHeaders = True: .ColHeadersShow = True: .RowHeadersShow = True: .GridShowHoriz = True: .GridShowVert = True
    End With
        '
        ' Export Excel file and set result to x
        FileName = "Export Data" & "(" & CompCode & "_Vch-" & VchType & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"
        SheetName = "Export Data" & "(" & CompCode & "_Vch-" & VchType & ")"
        LogFileName = "Export Data" & "(" & CompCode & "_Vch-" & VchType & ")" & Format(Date, "dd-MMM-yyyy") & ".txt"
        x = fpSpread1.ExportToExcelEx(FileName, SheetName, LogFileName, ExcelSaveFlagNoFormulas)
        ' Display result to user based on T/F value of x
    If x = True Then
        MsgBox "Export complete.", vbInformation, "Easy Publish...Export !!! "
            Dim oExcel As Object
            Set oExcel = CreateObject("Excel.Application")
            oExcel.Workbooks.Open (App.Path & "\" & FileName)
            oExcel.Visible = True
            Set oExcel = Nothing
    Else
        MsgBox "Export did not succeed.", vbInformation, "Easy Publish...Export !!!"
    End If
'
With fpSpread1
    'Delete Header Row
    If VchType <= 10 And VchType >= 7 Or VchType <= 28 And VchType >= 25 Then fpSpread1.DeleteRows 1, 2 Else: fpSpread1.DeleteRows 1, 1
End With
End If
Screen.MousePointer = vbNormal
Exit Sub

ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
Private Sub cmdFilter_Click()
    Dim i, n, C, R As Integer, Cval As Variant
    Dim StockVal As Variant, StockTotal As Variant
    Dim PVal As Variant, PTotal As Variant
    Dim PRVal As Variant, PRTotal As Variant
    Dim PCVal As Variant, PCTotal As Variant
    Dim PRCVal As Variant, PRCTotal As Variant
    Dim SVal As Variant, STotal As Variant
    Dim SRVal As Variant, SRTotal As Variant
    Dim SCVal As Variant, SCTotal As Variant
    Dim SRCVal As Variant, SRCTotal As Variant
    Dim SJIVal As Variant, SJITotal As Variant
    Dim SJOVal As Variant, SJOTotal As Variant
    Dim POVal As Variant, POTotal As Variant
    Dim SOVal As Variant, SOTotal As Variant
    Dim EStockVal As Variant, EStockTotal As Variant
    Dim AVal As Variant, ATotal As Variant
    Dim NPVal As Variant, NPValTotal As Variant
    Dim NSVal As Variant, NSValTotal As Variant
    Dim PAVal As Variant, PAValTotal As Variant
    Dim SAVal As Variant, SAValTotal As Variant
    Dim PRAVal As Variant, PRAValTotal As Variant
    Dim SRAVal As Variant, SRAValTotal As Variant
    Dim NPAVal As Variant, NPAValTotal As Variant
    Dim NSAVal As Variant, NSAValTotal As Variant

If VSFlexFlag = True Then
    With Me.VSFlexGrid1
        If VSFlexGrid1.Rows - 1 = 1 Then Exit Sub
                If VSFlexGrid1.Rows - 1 = 0 Then Exit Sub
                For i = 5 To VSFlexGrid1.Rows - 1 'Unhide All
                    VSFlexGrid1.RowHidden(i) = False
                Next
                If Text1.Text = "" Then Exit Sub
                If Combo4.Value = "" Then Combo4.ListIndex = 2
                    R = IIf(VSFlexGrid1.Row + 1 <> LR, VSFlexGrid1.Row + 1, 1)
                    LR = R
                For i = 0 To VSFlexGrid1.RightCol  'Match Col Header
                C = C + 1
                Cval = VSFlexGrid1.TextMatrix(0, C)
                If Combo4.Value = Cval Then Exit For
                Next
                
                    For i = 5 To VSFlexGrid1.Rows - 1
                    Cval = VSFlexGrid1.TextMatrix(i, C)
                                If InStr(StrConv(Cval, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
                                VSFlexGrid1.RowHidden(i) = True
                                Else
                                VSFlexGrid1.RowHidden(i) = False
                                End If
                    Next
    
    'StockVal
    .Subtotal flexSTSum, -1, 4, "(#,##0.00)", vbBlue, vbWhite, True, "Grand Total", , True
    .Subtotal flexSTSum, -1, 5, "(#,##0)", vbWhite, vbBlue, True, , , True
    .Subtotal flexSTSum, -1, 6, "(#,##0)", vbWhite, vbBlue, True, , , True
    .Subtotal flexSTSum, -1, 7, "(#,##0)", vbWhite, vbBlue, True, , , True
    .Subtotal flexSTSum, -1, 8, "(#,##0)", vbWhite, vbBlue, True, , , True
    VSFlexGrid1.RowHidden(.Rows - 1) = False
    'VSFlexGrid1.TextMatrix(.Rows - 1, 1) = "Grand Total"
    StockVal = 0
            For C = 4 To 8
                For i = 7 To VSFlexGrid1.Rows - 2 'Match Col Header
                If VSFlexGrid1.RowHidden(i) = False Then
                    StockVal = StockVal + val(VSFlexGrid1.TextMatrix(i, C)) 'Total
                End If
                Next
                VSFlexGrid1.TextMatrix(VSFlexGrid1.Rows - 1, C) = StockVal
                StockVal = 0
            Next
    End With
Else
    With fpSpread1
    n = .DataRowCnt: StockVal = 0
        For i = 1 To .DataRowCnt 'Unhide All
            .Row = i: .RowHidden = False
        Next
        If CheckEmpty(Text1.Text, False) Then TDBNumber2 = n - 1: Exit Sub
        C = Combo4.ListIndex + 2
        .SetActiveCell C, 1
        For i = 1 To .DataRowCnt
'        If Combo4.ListIndex = 0 Then .SetActiveCell C, i: .GetText C, i, cVal Else .SetActiveCell C, 1 ': .GetText 3, i, cVal
        .SetActiveCell C, i: .GetText C, i, Cval  'Else .SetActiveCell C, 1 ': .GetText 3, i, cVal
                .GetText 4, i, StockVal
                .GetText 6, i, PVal
                .GetText 7, i, PRVal
                .GetText 8, i, PCVal
                .GetText 9, i, PRCVal
                .GetText 10, i, SVal
                .GetText 11, i, SRVal
                .GetText 12, i, SCVal
                .GetText 13, i, SRCVal
                .GetText 14, i, SJIVal
                .GetText 15, i, SJOVal
                .GetText 16, i, POVal
                .GetText 17, i, SOVal
                .GetText 18, i, EStockVal
                .GetText 20, i, AVal
                .GetText 21, i, NPVal
                .GetText 22, i, NSVal
                .GetText 24, i, PAVal
                .GetText 25, i, SAVal
                .GetText 26, i, PRAVal
                .GetText 27, i, SRAVal
                .GetText 28, i, NPAVal
                .GetText 29, i, NSAVal
        If InStr(StrConv(Cval, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
        .Row = i: .RowHidden = True: n = n - 1
            Else
        .SetActiveCell C, i '1
        StockTotal = StockTotal + StockVal '4
        PTotal = PTotal + PVal '6
        PRTotal = PRTotal + PRVal '7
        PCTotal = PCTotal + PCVal '8
        PRCTotal = PRCTotal + PRCVal '9
        STotal = STotal + SVal '10
        SRTotal = SRTotal + SRVal '11
        SCTotal = SCTotal + SCVal '12
        SRCTotal = SRCTotal + SRCVal '13
        SJITotal = SJITotal + SJIVal '14
        SJOTotal = SJOTotal + SJOVal '15
        POTotal = POTotal + POVal '16
        SOTotal = SOTotal + SOVal '17
        EStockTotal = EStockTotal + EStockVal '18
        ATotal = ATotal + AVal '20
        NPValTotal = NPValTotal + NPVal '21
        NSValTotal = NSValTotal + NSVal '22
        PAValTotal = PAValTotal + PAVal '24
        SAValTotal = SAValTotal + SAVal '25
        PRAValTotal = PRAValTotal + PRAVal '26
        SRAValTotal = SRAValTotal + SRAVal '27
        NPAValTotal = NPAValTotal + NPAVal '28
        NSAValTotal = NSAValTotal + NSAVal '29
        End If
            TDBNumber2 = n
        Next
        .SetText 4, i - 1, StockTotal
        .SetText 6, i - 1, PTotal
        .SetText 7, i - 1, PRTotal
        .SetText 8, i - 1, PCTotal
        .SetText 9, i - 1, PRCTotal
        .SetText 10, i - 1, STotal
        .SetText 11, i - 1, SRTotal
        .SetText 12, i - 1, SCTotal
        .SetText 13, i - 1, SRCTotal
        .SetText 14, i - 1, SJITotal
        .SetText 15, i - 1, SJOTotal
        .SetText 16, i - 1, POTotal
        .SetText 17, i - 1, SOTotal
        .SetText 18, i - 1, EStockTotal
        .SetText 20, i - 1, ATotal
        .SetText 21, i - 1, NPValTotal
        .SetText 22, i - 1, NSValTotal
        .SetText 24, i - 1, PAValTotal
        .SetText 25, i - 1, SAValTotal
        .SetText 26, i - 1, PRAValTotal
        .SetText 27, i - 1, SRAValTotal
        .SetText 28, i - 1, NPAValTotal
        .SetText 29, i - 1, NSAValTotal
        If VchType >= 7 Then .Row = 1: .RowHidden = False:
        '.Row = i - 1: .RowHidden = False: .SelectBlockOptions = SelectBlockOptionsAll
        .SetActiveCell 1, 1
        .SetActiveCell 1, i - 1
    End With
End If
Screen.MousePointer = vbNormal
Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
Private Sub Preview_Click()
Dim PrintHeader As String, x As Boolean
Dim SheetName, FileName As String
Dim Cval As Variant, i, C As Integer
On Error Resume Next
If VSFlexFlag = True Then
    With Me.VSFlexGrid1
        If Dir(App.Path & "\Export", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Export"
        FileName = App.Path & "\Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"
        SheetName = "Sheet1" '"(" & Me.Caption & ")"
        .SaveGrid FileName, flexFileExcel, SaveExcelSettings.flexXLSaveFixedCells
        If Dir(FileName, vbDirectory) <> "" Then x = True
        If x = True Then
        'MsgBox "Export complete.", vbInformation, "Easy Publish...Export !!! "
        'Shell "C:\WINDOWS\explorer.exe """ & FileName & "", vbNormalFocus
        Else
        MsgBox "File Not Exist", vbInformation, "Easy Publish...Import !!!"
        End If
    End With
If Dir(App.Path & "\Export", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Export"

FileName = App.Path & "\Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"
' Load an Excel-formatted file
fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, False

'MsgBox
    'MsgBox "Import Processing....", vbInformation, "Easy Publish...Import !!! "
fpSpread1.ImportExcelBook FileName, ""        '& "\EasyPublish.xls", ""
End If
'*********************************************************
With fpSpread1
    .ColsFrozen = 0
    PrintHeader = Me.Caption
    .LockBackColor = vbWhite
    ' These are 8.5" X 11" paper dimensions in TWIPS  12240  15840
    Const PaperWidth = 12240
    Const PaperHeight = 15840
    Printer.PaperSize = vbPRPSA4
    ' Set printing options for sheet
    .PrintAbortMsg = "Printing - Click Cancel to .Quit"
    .PrintJobName = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") '& ".pdf"
    '.PrintHeader = "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy"): .PrintHeader=: .Font = 20 '"/cPrint Header/rPage # ./p/n2nd Line"
    .PrintFooter = "        Export Data Company : " & rstCompanyMaster.Fields("PrintName").Value & " _(" & CompCode & "_" & PrintHeader & ")" & "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]" & " /rPage # ./p " & " Print Date : ( " & Format(Date, "dd-MMM-yyyy") & " )         ": '.FontSize = 16 '& ".pdf" ' "/cPrint Footer/rPage # ./p/n2nd Line"
    .PrintBorder = True
    If VchType <> 19 Then .PrintColHeaders = True Else .PrintColHeaders = False
    .PrintColor = True
    .PrintGrid = True
    .PrintMarginTop = 200 '750 '1440
    .PrintMarginBottom = 200 '500 '1440
    .PrintMarginLeft = 100 '720
    .PrintMarginRight = 100 '720
    '.PrintType = SPRD_PRINT_ALL
    If VchType <> 19 Then .PrintRowHeaders = True Else .PrintRowHeaders = False
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
    If VchType = 19 Then
        .PrintZoomFactor = 0.95
        C = 0
        For i = 2 To 9
            .GetText i, 1, Cval: .SetText i, 0, Cval
        Next
            .Col = 5: .CellType = CellTypeNumber: .TypeNumberDecPlaces = 2
            .Col = 6: .CellType = CellTypeNumber: .TypeNumberDecPlaces = 0
            .Col = 7: .CellType = CellTypeNumber: .TypeNumberDecPlaces = 0
            .Col = 8: .CellType = CellTypeNumber: .TypeNumberDecPlaces = 0
            .Col = 9: .CellType = CellTypeNumber: .TypeNumberDecPlaces = 0
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintGrid = False
        .DeleteRows 1, 1
    Else
        .PrintZoomFactor = 0.75
    End If
    ' Print
    '.PrintSheet 0
    If VchType = 19 Then
        .PrintOrientation = PrintOrientationPortrait
    Else
        .PrintOrientation = PrintOrientationLandscape
    End If
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
Private Sub Search_Click()
Dim i, C As Integer, Cval As Variant, R As Long
  If VSFlexGrid1.Visible Then
              C = 0
            If VSFlexGrid1.BottomRow = 1 Then Exit Sub
            If Text1.Text = "" Then Exit Sub
                If VSFlexGrid1.BottomRow = 0 Then Exit Sub
                For i = 1 To VSFlexGrid1.Rows - 1 'Unhide All
                    VSFlexGrid1.RowHidden(i) = False
                Next
                For i = 1 To VSFlexGrid1.RightCol  'Match Col Header
                C = C + 1
                Cval = VSFlexGrid1.TextMatrix(0, C)
                If Combo4.Value = Cval Then Exit For
                Next
                
                R = IIf(VSFlexGrid1.Row + 1 <> LR, VSFlexGrid1.Row + 1, 1)
                LR = R
            For C = 1 To VSFlexGrid1.RightCol
                For i = R To VSFlexGrid1.Rows - 1
                Cval = VSFlexGrid1.TextMatrix(i, C)
                            If InStr(StrConv(Cval, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
                            VSFlexGrid1.Row = i: VSFlexGrid1.Col = C: VSFlexGrid1.RowPosition(VSFlexGrid1.Row) = i: VSFlexGrid1.ColPosition(VSFlexGrid1.Col) = C
                            Else
                            VSFlexGrid1.Row = i: VSFlexGrid1.Col = C: VSFlexGrid1.TopRow = i: VSFlexGrid1.LeftCol = C: Exit Sub
                            End If
                Next
           Next
  Else
        With fpSpread1
                    If Text1.Text = "" Then Exit Sub
                    If .DataRowCnt = 0 Then Exit Sub
                
                    For i = 1 To .DataRowCnt 'Unhide All
                        .Row = i: .RowHidden = False
                    Next
            
                    R = IIf(.ActiveRow + 1 <> LR, .ActiveRow + 1, 1)
                    LR = R
                    For i = R To .DataRowCnt
                    If Combo4.ListIndex = Combo4.ListIndex Then .GetText Combo4.ListIndex + 2, i, Cval                                                                 'Else .SetActiveCell 3, 1: .GetText 3, i, cVal
                                If InStr(StrConv(Cval, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
                                
                                Else
                                        .SetActiveCell Combo4.ListIndex + 2, i: Exit Sub
                                End If
                    Next
        End With
End If
End Sub
Public Function Print_Grid()
Dim aSNO, pSNO, dPrint As Long
Dim GodownNameH, SizeNameH, PaperNameH, UOMH, GSMH As String
Dim INWardGTF, OUTWardGTF, BalGTF
    On Error GoTo ErrHandler
Dim i, C As Long
dPrint = 0: pSNO = 0: aSNO = 0: GodownNameH = "": SizeNameH = "": PaperNameH = "": UOMH = "": GSMH = "": INWardGTF = 0: OUTWardGTF = 0: BalGTF = 0
    MdiMainMenu.MousePointer = vbHourglass: ShowProgressInStatusBar True: Timer1.Enabled = True
    If VchType = 19 Then
           With VSFlexGrid1
                    .Clear: .Cols = 16: .Rows = rstPaperLedger.RecordCount + 10
                    rstPaperLedger.MoveFirst
                    If VchType = 19 Then i = 6
    Do While Not rstPaperLedger.EOF
        C = 1: i = i + 1
        If GodownNameH <> rstPaperLedger.Fields("GodownName").Value Then
        'Party Header
                aSNO = aSNO + 1
                .TextMatrix(i, 0) = "A/C-" & aSNO
                .TextMatrix(i, C) = rstPaperLedger.Fields("GodownName").Value: .RowHeight(i) = 400: .Cell(flexcpFontSize, i, C) = 12: .Cell(flexcpFontBold, i, C) = True: .Cell(flexcpBackColor, i, C) = vbWhite: .Cell(flexcpFontUnderline, i, C) = True: .Cell(flexcpForeColor, i, C) = RGB(128, 0, 64)
                 GodownNameH = rstPaperLedger.Fields("GodownName").Value
                 
                 .TextMatrix(i, 9) = rstPaperLedger.Fields("GodownName").Value
                 .TextMatrix(i, 10) = rstPaperLedger.Fields("SizeName").Value
                 .TextMatrix(i, 11) = rstPaperLedger.Fields("UOM").Value
                 .TextMatrix(i, 12) = rstPaperLedger.Fields("GSM").Value
                 .TextMatrix(i, 13) = rstPaperLedger.Fields("PaperName").Value
                    i = i + 1
                 .Rows = .Rows + 1
         End If
        If SizeNameH <> rstPaperLedger.Fields("SizeName").Value Then
        'Size Header
                .TextMatrix(i, C) = rstPaperLedger.Fields("SizeName").Value: .Cell(flexcpFontSize, i, C) = 10: .Cell(flexcpFontBold, i, C) = True: .Cell(flexcpBackColor, i, C) = vbWhite: .Cell(flexcpFontUnderline, i, C) = False: .Cell(flexcpForeColor, i, C) = RGB(113, 0, 113)
                 SizeNameH = rstPaperLedger.Fields("SizeName").Value
                 
                 .TextMatrix(i, 9) = rstPaperLedger.Fields("GodownName").Value
                 .TextMatrix(i, 10) = rstPaperLedger.Fields("SizeName").Value
                 .TextMatrix(i, 11) = rstPaperLedger.Fields("UOM").Value
                 .TextMatrix(i, 12) = rstPaperLedger.Fields("GSM").Value
                 .TextMatrix(i, 13) = rstPaperLedger.Fields("PaperName").Value
                 
                 i = i + 1: .Rows = .Rows + 1
         End If
        If UOMH <> rstPaperLedger.Fields("UOM").Value Then
        'UOM Header
                .TextMatrix(i, C) = rstPaperLedger.Fields("UOM").Value: .Cell(flexcpFontSize, i, C) = 10: .Cell(flexcpFontBold, i, C) = True: .Cell(flexcpBackColor, i, C) = vbWhite: .Cell(flexcpFontItalic, i, C) = True: .Cell(flexcpForeColor, i, C) = RGB(80, 80, 160)
                 UOMH = rstPaperLedger.Fields("UOM").Value
                 
                 .TextMatrix(i, 9) = rstPaperLedger.Fields("GodownName").Value
                 .TextMatrix(i, 10) = rstPaperLedger.Fields("SizeName").Value
                 .TextMatrix(i, 11) = rstPaperLedger.Fields("UOM").Value
                 .TextMatrix(i, 12) = rstPaperLedger.Fields("GSM").Value
                 .TextMatrix(i, 13) = rstPaperLedger.Fields("PaperName").Value
                 
                 i = i + 1: .Rows = .Rows + 1
         End If
        If GSMH <> rstPaperLedger.Fields("GSM").Value Then
        'GSM Header
                .TextMatrix(i, C) = "GSM : " & rstPaperLedger.Fields("GSM").Value: .Cell(flexcpFontSize, i, C) = 10: .Cell(flexcpFontBold, i, C) = True: .Cell(flexcpBackColor, i, C) = vbWhite: .Cell(flexcpFontUnderline, i, C) = False: .Cell(flexcpForeColor, i, C) = vbRed
                 GSMH = rstPaperLedger.Fields("GSM").Value
                 
                 .TextMatrix(i, 9) = rstPaperLedger.Fields("GodownName").Value
                 .TextMatrix(i, 10) = rstPaperLedger.Fields("SizeName").Value
                 .TextMatrix(i, 11) = rstPaperLedger.Fields("UOM").Value
                 .TextMatrix(i, 12) = rstPaperLedger.Fields("GSM").Value
                 .TextMatrix(i, 13) = rstPaperLedger.Fields("PaperName").Value
                 
                 i = i + 1: .Rows = .Rows + 1
         End If
        If PaperNameH <> rstPaperLedger.Fields("PaperName").Value Then
        'Paper Name Header
                .TextMatrix(i, C) = rstPaperLedger.Fields("PaperName").Value: .Cell(flexcpFontSize, i, C) = 9: .Cell(flexcpFontBold, i, C) = True: .Cell(flexcpBackColor, i, C) = vbWhite: .Cell(flexcpFontUnderline, i, C) = False: .Cell(flexcpForeColor, i, C) = RGB(0, 106, 106)
                 PaperNameH = rstPaperLedger.Fields("PaperName").Value
                 
                 .TextMatrix(i, 9) = rstPaperLedger.Fields("GodownName").Value
                 .TextMatrix(i, 10) = rstPaperLedger.Fields("SizeName").Value
                 .TextMatrix(i, 11) = rstPaperLedger.Fields("UOM").Value
                 .TextMatrix(i, 12) = rstPaperLedger.Fields("GSM").Value
                 .TextMatrix(i, 13) = rstPaperLedger.Fields("PaperName").Value
                 
                 i = i + 1: .Rows = .Rows + 1
         End If
                C = 0
                C = C + 1: .TextMatrix(i, C) = rstPaperLedger.Fields("VchType").Value + IIf(rstPaperLedger.Fields("VchNo").Value <> "", "-" + rstPaperLedger.Fields("VchNo").Value, "")
                C = C + 1: .TextMatrix(i, C) = Format(rstPaperLedger.Fields("VchDate").Value, "dd-mm-yy")
                C = C + 1: .TextMatrix(i, C) = rstPaperLedger.Fields("Particulars").Value: .WordWrap = True: .AutoSizeMode = flexAutoSizeRowHeight
                C = C + 1: .TextMatrix(i, C) = Format(val(rstPaperLedger.Fields("Forms").Value), "###0.00")
                C = C + 1: .TextMatrix(i, C) = val(rstPaperLedger.Fields("BookQuantity").Value)
'********************************IN
            If InStr(1, "PI_SI_MI_CN_OB", StrConv(rstPaperLedger.Fields("VchType").Value, vbUpperCase)) > 0 And val(rstPaperLedger.Fields("Quantity").Value) > 0 Then
                    If Combo1.ListIndex = 0 Then
                        C = C + 1: .TextMatrix(i, C) = val(rstPaperLedger.Fields("Quantity").Value)
                    ElseIf Combo1.ListIndex = 1 Then
                        C = C + 1: .TextMatrix(i, C) = Format(val(rstPaperLedger.Fields("Quantity").Value) / val(rstPaperLedger.Fields("SPU").Value), "###0.000")
                    ElseIf Combo1.ListIndex = 2 Then
                        C = C + 1: .TextMatrix(i, C) = Format(val(rstPaperLedger.Fields("Quantity").Value) / val(rstPaperLedger.Fields("SPU").Value) * val(rstPaperLedger.Fields("Weight/Unit").Value), "###0.000") 'Weight/Unit
                    ElseIf Combo1.ListIndex = 3 Then
                        C = C + 1: .TextMatrix(i, C) = Format(val(rstPaperLedger.Fields("Quantity").Value) / val(rstPaperLedger.Fields("SPU").Value), "###0.000")
                        'C = C + 1: .TextMatrix(i, C) = Format(Int(Val(Abs(rstPaperLedger.Fields("Quantity").Value)) / Val(rstPaperLedger.Fields("SPU").Value)) + ((Val(Abs(rstPaperLedger.Fields("Quantity").Value)) / Val(rstPaperLedger.Fields("SPU").Value)) - Int(Val(Abs(rstPaperLedger.Fields("Quantity").Value)) / Val(rstPaperLedger.Fields("SPU").Value))) * Val(rstPaperLedger.Fields("SPU").Value) / 1000, "####0.000")
                    End If
                C = C + 1
'********************************OUT
            Else
                C = C + 1
                If Combo1.ListIndex = 0 Then
                    C = C + 1: .TextMatrix(i, C) = val(Abs(rstPaperLedger.Fields("Quantity").Value))
                ElseIf Combo1.ListIndex = 1 Then
                    C = C + 1: .TextMatrix(i, C) = Format(val(Abs(rstPaperLedger.Fields("Quantity").Value)) / val(rstPaperLedger.Fields("SPU").Value), "###0.000")
                ElseIf Combo1.ListIndex = 2 Then
                    C = C + 1: .TextMatrix(i, C) = Format(val(Abs(rstPaperLedger.Fields("Quantity").Value)) / val(rstPaperLedger.Fields("SPU").Value) * val(rstPaperLedger.Fields("Weight/Unit").Value), "###0.000") 'Weight/Unit
                ElseIf Combo1.ListIndex = 3 Then
                    C = C + 1: .TextMatrix(i, C) = Format(val(Abs(rstPaperLedger.Fields("Quantity").Value)) / val(rstPaperLedger.Fields("SPU").Value), "###0.000")
                    'C = C + 1: .TextMatrix(i, C) = Format(Int(Val(Abs(rstPaperLedger.Fields("Quantity").Value)) / Val(rstPaperLedger.Fields("SPU").Value)) + ((Val(Abs(rstPaperLedger.Fields("Quantity").Value)) / Val(rstPaperLedger.Fields("SPU").Value)) - Int(Val(Abs(rstPaperLedger.Fields("Quantity").Value)) / Val(rstPaperLedger.Fields("SPU").Value))) * Val(rstPaperLedger.Fields("SPU").Value) / 1000, "####0.000")
                End If
            End If
'********************************Bal
                C = C + 1:  .TextMatrix(i, C) = Format(val(val(.TextMatrix(i, C - 2)) - val(.TextMatrix(i, C - 1))), "###0.000"): .Cell(flexcpForeColor, i, C) = vbWhite: pSNO = pSNO + 1
'********************************
                 .TextMatrix(i, 9) = rstPaperLedger.Fields("GodownName").Value
                 .TextMatrix(i, 10) = rstPaperLedger.Fields("SizeName").Value
                 .TextMatrix(i, 11) = rstPaperLedger.Fields("UOM").Value
                 .TextMatrix(i, 12) = rstPaperLedger.Fields("GSM").Value
                 .TextMatrix(i, 13) = rstPaperLedger.Fields("PaperName").Value
                 .TextMatrix(i, 14) = rstPaperLedger.Fields("SPU").Value
            dPrint = dPrint + 1
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstPaperLedger.RecordCount & " !!!"
                rstPaperLedger.MoveNext
            If MdiMainMenu.ProgressBar1.Value + Round((100 / rstPaperLedger.RecordCount), 2) <= 100 Then
                MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + Round((100 / rstPaperLedger.RecordCount), 2)
            End If
            Loop
          .Rows = i + 1
            End With
    Else
         Set VSFlexGrid1.DataSource = rstPaperLedger
    End If
    TDBNumber2.Value = pSNO

Call VSFlexGrid_Format_Headers
Call VSFlexGrid1_AfterDataRefresh
    Timer1.Enabled = False
    ShowProgressInStatusBar False
    MdiMainMenu.MousePointer = vbNormal
    Screen.MousePointer = vbNormal
    Exit Function
ErrHandler:
    Timer1.Enabled = False
    ShowProgressInStatusBar False
    MdiMainMenu.MousePointer = vbNormal
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Function
Private Function VSFlexGrid_Format_Headers()
On Error GoTo ErrHandler
Dim i As Long
Dim C As Long
        With VSFlexGrid1
        .ColWidth(0) = 5
            If VchType = 19 Then
            C = 0
                    C = C + 1: .TextMatrix(i, C) = "Vch No": .ColWidth(1) = 1020
                    C = C + 1: .TextMatrix(i, C) = "Date": .ColWidth(2) = 850
                    C = C + 1: .TextMatrix(i, C) = "Particulars": .ColWidth(3) = 5120
                    C = C + 1: .TextMatrix(i, C) = "Forms": .ColWidth(4) = 600
                    C = C + 1: .TextMatrix(i, C) = "Quantity": .ColWidth(5) = 700
                    C = C + 1: .TextMatrix(i, C) = "IN": .ColWidth(6) = 990
                    C = C + 1: .TextMatrix(i, C) = "OUT": .ColWidth(7) = 990
                    C = C + 1: .TextMatrix(i, C) = "Balance": .ColWidth(8) = 990
            End If
        End With
Screen.MousePointer = vbNormal
Exit Function
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Function
Private Sub VSFlexGrid1_AfterDataRefresh()
    On Error GoTo ErrHandler
Dim C As Variant
Dim T As Long
Dim GroupOn As Long
Dim i As Integer
    With VSFlexGrid1
    If VchType = 19 Then
        GroupOn = 9
        .FrozenRows = 3
    End If
    'Subtotal
nSort = False
    .SubtotalPosition = flexSTBelow
    .MultiTotals = True
    .Subtotal flexSTClear
    For i = 5 To VSFlexGrid1.Rows - 1 'Unhide All
        VSFlexGrid1.RowHidden(i) = False
    Next
If Combo1.ListIndex = 0 Then
    .Subtotal flexSTSum, -1, 4, "(#,##0.00)", vbBlue, vbWhite, True, "Grand Total", , True: .TextMatrix(.Rows - 1, 1) = "Grand Total"
    .Subtotal flexSTSum, -1, 5, "(#,##0)", vbWhite, vbBlue, True, , , True
    .Subtotal flexSTSum, -1, 6, "(#,##0)", vbWhite, vbBlue, True, , , True
    .Subtotal flexSTSum, -1, 7, "(#,##0)", vbWhite, vbBlue, True, , , True
    .Subtotal flexSTSum, -1, 8, "(#,##0)", vbWhite, vbBlue, True, , , True
Else
    .Subtotal flexSTSum, -1, 4, "(#,##0.00)", vbBlue, vbWhite, True, "Grand Total", , True: .TextMatrix(.Rows - 1, 1) = "Grand Total"
    .Subtotal flexSTSum, -1, 5, "(#,##0.000)", vbWhite, vbBlue, True, , , True
    .Subtotal flexSTSum, -1, 6, "(#,##0.000)", vbWhite, vbBlue, True, , , True
    .Subtotal flexSTSum, -1, 7, "(#,##0.000)", vbWhite, vbBlue, True, , , True
    .Subtotal flexSTSum, -1, 8, "(#,##0.000)", vbWhite, vbBlue, True, , , True
End If
'Party Total
If Check0.Value Then
    If Combo1.ListIndex = 0 Then
        .Subtotal flexSTSum, 9, 4, "(#,##0.00)", &H8000000F, RGB(128, 0, 64), True, , 9, True
        .Subtotal flexSTSum, 9, 5, "(#,##0)", &H8000000F, RGB(128, 0, 64), True, , 9, True
        .Subtotal flexSTSum, 9, 6, "(#,##0)", &H8000000F, RGB(128, 0, 64), True, , 9, True
        .Subtotal flexSTSum, 9, 7, "(#,##0)", &H8000000F, RGB(128, 0, 64), True, , 9, True
        .Subtotal flexSTSum, 9, 8, "(#,##0)", &H8000000F, RGB(128, 0, 64), True, , 9, True
        .RowHidden(7) = True
    Else
        .Subtotal flexSTSum, 9, 4, "(#,##0.00)", &H8000000F, RGB(128, 0, 64), True, , 9, True
        .Subtotal flexSTSum, 9, 5, "(#,##0)", &H8000000F, RGB(128, 0, 64), True, , 9, True
        .Subtotal flexSTSum, 9, 6, "(#,##0.000)", &H8000000F, RGB(128, 0, 64), True, , 9, True
        .Subtotal flexSTSum, 9, 7, "(#,##0.000)", &H8000000F, RGB(128, 0, 64), True, , 9, True
        .Subtotal flexSTSum, 9, 8, "(#,##0.000)", &H8000000F, RGB(128, 0, 64), True, , 9, True
        .RowHidden(7) = True
    End If
End If
 'Size Total
If Check2.Value Then
    'If FrmPaperStockRegister.Check6.Value Then
    If Combo1.ListIndex = 0 Then
        .Subtotal flexSTSum, 10, 4, "(#,##0.00)", &H8000000F, RGB(113, 0, 113), True, , 10, True
        .Subtotal flexSTSum, 10, 5, "(#,##0)", &H8000000F, RGB(113, 0, 113), True, , 10, True
        .Subtotal flexSTSum, 10, 6, "(#,##0)", &H8000000F, RGB(113, 0, 113), True, , 10, True
        .Subtotal flexSTSum, 10, 7, "(#,##0)", &H8000000F, RGB(113, 0, 113), True, , 10, True
        .Subtotal flexSTSum, 10, 8, "(#,##0)", &H8000000F, RGB(113, 0, 113), True, , 10, True
        .RowHidden(7) = True
    Else
        .Subtotal flexSTSum, 10, 4, "(#,##0.00)", &H8000000F, RGB(113, 0, 113), True, , 10, True
        .Subtotal flexSTSum, 10, 5, "(#,##0)", &H8000000F, RGB(113, 0, 113), True, , 10, True
        .Subtotal flexSTSum, 10, 6, "(#,##0.000)", &H8000000F, RGB(113, 0, 113), True, , 10, True
        .Subtotal flexSTSum, 10, 7, "(#,##0.000)", &H8000000F, RGB(113, 0, 113), True, , 10, True
        .Subtotal flexSTSum, 10, 8, "(#,##0.000)", &H8000000F, RGB(113, 0, 113), True, , 10, True
        .RowHidden(7) = True
    End If
End If
 'UOM Total
'If FrmPaperStockRegister.Check7.Value Then
If Check3.Value Then
    If Combo1.ListIndex = 0 Then
        .Subtotal flexSTSum, 11, 4, "(#,##0.00)", &H8000000F, RGB(80, 80, 160), True, , 11, True
        .Subtotal flexSTSum, 11, 5, "(#,##0)", &H8000000F, RGB(80, 80, 160), True, , 11, True
        .Subtotal flexSTSum, 11, 6, "(#,##0)", &H8000000F, RGB(80, 80, 160), True, , 11, True
        .Subtotal flexSTSum, 11, 7, "(#,##0)", &H8000000F, RGB(80, 80, 160), True, , 11, True
        .Subtotal flexSTSum, 11, 8, "(#,##0)", &H8000000F, RGB(80, 80, 160), True, , 11, True
        .RowHidden(7) = True
    Else
        .Subtotal flexSTSum, 11, 4, "(#,##0.00)", &H8000000F, RGB(80, 80, 160), True, , 11, True
        .Subtotal flexSTSum, 11, 5, "(#,##0)", &H8000000F, RGB(80, 80, 160), True, , 11, True
        .Subtotal flexSTSum, 11, 6, "(#,##0.000)", &H8000000F, RGB(80, 80, 160), True, , 11, True
        .Subtotal flexSTSum, 11, 7, "(#,##0.000)", &H8000000F, RGB(80, 80, 160), True, , 11, True
        .Subtotal flexSTSum, 11, 8, "(#,##0.000)", &H8000000F, RGB(80, 80, 160), True, , 11, True
        .RowHidden(7) = True
    End If
End If
 'GSM Total
    'If FrmPaperStockRegister.Check5.Value Then
If Check1.Value Then
    If Combo1.ListIndex = 0 Then
        .Subtotal flexSTSum, 12, 4, "(#,##0.00)", &H8000000F, vbRed, True, , 12, True
        .Subtotal flexSTSum, 12, 5, "(#,##0)", &H8000000F, vbRed, True, , 12, True
        .Subtotal flexSTSum, 12, 6, "(#,##0)", &H8000000F, vbRed, True, , 12, True
        .Subtotal flexSTSum, 12, 7, "(#,##0)", &H8000000F, vbRed, True, , 12, True
        .Subtotal flexSTSum, 12, 8, "(#,##0)", &H8000000F, vbRed, True, , 12, True
        .RowHidden(7) = True
    Else
         .Subtotal flexSTSum, 12, 4, "(#,##0.00)", &H8000000F, vbRed, True, , 12, True
        .Subtotal flexSTSum, 12, 5, "(#,##0)", &H8000000F, vbRed, True, , 12, True
        .Subtotal flexSTSum, 12, 6, "(#,##0.000)", &H8000000F, vbRed, True, , 12, True
        .Subtotal flexSTSum, 12, 7, "(#,##0.000)", &H8000000F, vbRed, True, , 12, True
        .Subtotal flexSTSum, 12, 8, "(#,##0.000)", &H8000000F, vbRed, True, , 12, True
        .RowHidden(7) = True
    End If
End If
 'Paper Total
If Check4.Value Then
        'If FrmPaperStockRegister.Check4.Value Then
    If Combo1.ListIndex = 0 Then
        .Subtotal flexSTSum, 13, 4, "(#,##0.00)", &H8000000F, RGB(0, 106, 106), True, , 13, True
        .Subtotal flexSTSum, 13, 5, "(#,##0)", &H8000000F, RGB(0, 106, 106), True, , 13, True
        .Subtotal flexSTSum, 13, 6, "(#,##0)", &H8000000F, RGB(0, 106, 106), True, , 13, True
        .Subtotal flexSTSum, 13, 7, "(#,##0)", &H8000000F, RGB(0, 106, 106), True, , 13, True
        .Subtotal flexSTSum, 13, 8, "(#,##0)", &H8000000F, RGB(0, 106, 106), True, , 13, True
        .RowHidden(7) = True
    Else
        .Subtotal flexSTSum, 13, 4, "(#,##0.00)", &H8000000F, RGB(0, 106, 106), True, , 13, True
        .Subtotal flexSTSum, 13, 5, "(#,##0)", &H8000000F, RGB(0, 106, 106), True, , 13, True
        .Subtotal flexSTSum, 13, 6, "(#,##0.000)", &H8000000F, RGB(0, 106, 106), True, , 13, True
        .Subtotal flexSTSum, 13, 7, "(#,##0.000)", &H8000000F, RGB(0, 106, 106), True, , 13, True
        .Subtotal flexSTSum, 13, 8, "(#,##0.000)", &H8000000F, RGB(0, 106, 106), True, , 13, True
        .RowHidden(7) = True
    End If
End If

    For C = 2 To (.Cols - 1)
        .ExplorerBar = flexExSort
        .ColSort(C) = flexSortCustom
        .AllowUserResizing = flexResizeBoth
    Next
    
    For C = 9 To (.Cols - 1)
    .ColHidden(C) = True
    Next
     
     .AutoSizeMode = flexAutoSizeRowHeight
    C = .Rows - 1
    .TextMatrix(C, 0) = ""
    
    For C = 1 To 3
            .MergeCells = flexMergeSpill
            .MergeCol(C) = False
            .MergeRow(C) = True
            .WordWrap = True
            .RowHidden(C + 2) = True
    Next

    C = 3
    i = 1
    .TextMatrix(i, C) = "Paper Stock Register (" & IIf(FrmPaperStockRegister.Check2.Value, "              Summarised", "Detailed") & ")": .Cell(flexcpFontSize, i, C) = 12: .Cell(flexcpFontBold, i, C) = True: .Cell(flexcpBackColor, i, C) = vbWhite: .Cell(flexcpFontUnderline, i, C) = True: .Cell(flexcpForeColor, i, C) = RGB(128, 0, 0): .Cell(flexcpAlignment, i, C) = flexAlignRightCenter 'CenterCenter
    i = 2
    .TextMatrix(i, C) = Trim(rstCompanyMaster.Fields("PrintName").Value): .RowHeight(i) = 500: .Cell(flexcpFontSize, i, C) = 18: .Cell(flexcpFontBold, i, C) = True: .Cell(flexcpBackColor, i, C) = vbWhite: .Cell(flexcpFontUnderline, i, C) = False: .Cell(flexcpForeColor, i, C) = RGB(145, 0, 72): .Cell(flexcpAlignment, i, C) = flexAlignCenterCenter 'RGB(49, 120, 61)
    i = 3
    .TextMatrix(i, C) = "From [" + Format(GetDate(FrmPaperStockRegister.MhDateInput1.Text), "dd-mm-yyyy") + "] To [" + Format(GetDate(FrmPaperStockRegister.MhDateInput2.Text), "dd-mm-yyyy") + "] [" & IIf(FrmPaperStockRegister.Option1.Value, "Including In-Transit", IIf(FrmPaperStockRegister.Option2.Value, "Excluding In-Transit", "In-Transit Only")) & "]": .RowHeight(i) = 400: .Cell(flexcpFontSize, i, C) = 14: .Cell(flexcpFontBold, i, C) = True: .Cell(flexcpBackColor, i, C) = vbWhite: .Cell(flexcpFontUnderline, i, C) = False: .Cell(flexcpForeColor, i, C) = RGB(80, 80, 160): .Cell(flexcpAlignment, i, C) = flexAlignCenterCenter
    i = 6: C = 1
    .TextMatrix(i, C) = "Paper A/C IN :" + Combo1.Value: .Cell(flexcpFontSize, i, C) = 12: .Cell(flexcpFontBold, i, C) = True: .Cell(flexcpBackColor, i, C) = vbWhite: .Cell(flexcpFontUnderline, i, C) = True: .Cell(flexcpForeColor, i, C) = vbRed: .Cell(flexcpFontItalic, i, C) = True 'RGB(128, 0, 0)
    .AutoSize 1, 3
    
    For C = 1 To 8
    'Combo4.AddItem
    Combo4.AddItem .TextMatrix(0, C), C - 1
    Next
'    Combo4.ListIndex = 0
    If Combo1.ListIndex = 3 Then Find_SubTotal_Row
    End With
Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
Sub Find_SubTotal_Row()
Dim Cval, SPU, Col As Variant

    For Col = 6 To 8
        For i = 4 To VSFlexGrid1.Rows - 1
            If val(Format(VSFlexGrid1.TextMatrix(i, Col), "###0.000")) <> 0 Then
            If val(Format(VSFlexGrid1.TextMatrix(i, Col), "###0.000")) < 0 Then C = -1 Else C = 1
                    Cval = val(Format(VSFlexGrid1.TextMatrix(i, Col), "###0.000"))
                    SPU = IIf(val(VSFlexGrid1.TextMatrix(i, 14)) <> 0, val(VSFlexGrid1.TextMatrix(i, 14)), val(VSFlexGrid1.TextMatrix(i - 1, 14))): VSFlexGrid1.TextMatrix(i, 14) = SPU
                    VSFlexGrid1.TextMatrix(i, Col) = Format(Format(Int(Abs(Cval)) + ((Abs(Cval) - Int(Abs(Cval))) * SPU / 1000), "####0.000") * C, "####0.000")
            End If
        Next
    Next
End Sub
Private Sub CopyToClipboard()
    Dim selectedData As String
    Dim i As Integer
    ' Get the selected data from the grid
    For i = VSFlexGrid1.RowSel To VSFlexGrid1.Row
        selectedData = selectedData & VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColSel) & vbCrLf
    Next i
    ' Copy the selected data to the clipboard
    Clipboard.SetText selectedData
End Sub
Private Sub PasteFromClipboard()
    Dim clipboardData As String
    Dim dataRows() As String
    Dim i As Integer
    ' Get the data from the clipboard
    clipboardData = Clipboard.GetText
    ' Split the clipboard data into individual rows
    dataRows = Split(clipboardData, vbCrLf)
    ' Paste the data into the grid
    For i = 0 To UBound(dataRows)
        'VSFlexGrid1.TextMatrix(VSFlexGrid1.Row + i, VSFlexGrid1.Col) = dataRows(i)
    Next i
End Sub
Private Sub VSFlexGrid1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
Mh3dLabel9.Alignment = mhAlignmentCenter
Mh3dLabel9.Caption = "Ln" & NewRow & " Col " & NewCol
End Sub
