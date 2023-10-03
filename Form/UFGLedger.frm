VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmUFGLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " UFG Ledger"
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
      TabIndex        =   5
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
      Picture         =   "UFGLedger.frx":0000
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   5280
         TabIndex        =   26
         Top             =   120
         Width           =   6660
         _Version        =   65536
         _ExtentX        =   11747
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
         Picture         =   "UFGLedger.frx":001C
         Picture         =   "UFGLedger.frx":0038
      End
      Begin VB.CommandButton Preview 
         Caption         =   "&Print Preview"
         Height          =   330
         Left            =   15960
         TabIndex        =   25
         Top             =   8880
         Width           =   1215
      End
      Begin VB.CommandButton Search 
         Height          =   320
         Left            =   11040
         Picture         =   "UFGLedger.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   3240
         MaxLength       =   100
         TabIndex        =   19
         ToolTipText     =   "Find And Search"
         Top             =   8880
         Width           =   7230
      End
      Begin VB.CommandButton cmdFilter 
         Height          =   320
         Left            =   10560
         Picture         =   "UFGLedger.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Filter"
         Top             =   8880
         Width           =   375
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   14000
         TabIndex        =   14
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
         Picture         =   "UFGLedger.frx":06D8
         Picture         =   "UFGLedger.frx":06F4
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Summrize"
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
         TabIndex        =   12
         Top             =   158
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   375
         Left            =   19380
         Picture         =   "UFGLedger.frx":0710
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel"
         Top             =   90
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   375
         Left            =   19000
         Picture         =   "UFGLedger.frx":0812
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Refresh"
         Top             =   90
         Width           =   375
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
         TabIndex        =   9
         Top             =   158
         Width           =   2055
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
         TabIndex        =   4
         Top             =   600
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
         MaxCols         =   12
         MaxRows         =   1000
         SelectBlockOptions=   11
         ShadowColor     =   16775408
         SpreadDesigner  =   "UFGLedger.frx":095C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   6
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
         Picture         =   "UFGLedger.frx":140A
         Picture         =   "UFGLedger.frx":1426
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   1800
         TabIndex        =   7
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
         Picture         =   "UFGLedger.frx":1442
         Picture         =   "UFGLedger.frx":145E
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
         Calendar        =   "UFGLedger.frx":147A
         Caption         =   "UFGLedger.frx":1592
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "UFGLedger.frx":15FE
         Keys            =   "UFGLedger.frx":161C
         Spin            =   "UFGLedger.frx":167A
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
         Calendar        =   "UFGLedger.frx":16A2
         Caption         =   "UFGLedger.frx":17BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "UFGLedger.frx":1826
         Keys            =   "UFGLedger.frx":1844
         Spin            =   "UFGLedger.frx":18A2
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
         TabIndex        =   8
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
         Picture         =   "UFGLedger.frx":18CA
         Picture         =   "UFGLedger.frx":18E6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   18570
         TabIndex        =   16
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
         Picture         =   "UFGLedger.frx":1902
         Picture         =   "UFGLedger.frx":191E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   17280
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
         Caption         =   " Export Data"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "UFGLedger.frx":193A
         Picture         =   "UFGLedger.frx":1956
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   330
         Left            =   1200
         TabIndex        =   20
         Top             =   8880
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "UFGLedger.frx":1972
         Caption         =   "UFGLedger.frx":1992
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "UFGLedger.frx":19F6
         Keys            =   "UFGLedger.frx":1A14
         Spin            =   "UFGLedger.frx":1A5E
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
         TabIndex        =   21
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
         Picture         =   "UFGLedger.frx":1A86
         Picture         =   "UFGLedger.frx":1AA2
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   2520
         TabIndex        =   22
         Top             =   8880
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
         Picture         =   "UFGLedger.frx":1ABE
         Picture         =   "UFGLedger.frx":1ADA
      End
      Begin MSForms.ComboBox Combo4 
         Height          =   330
         Left            =   11520
         TabIndex        =   23
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
         TabIndex        =   15
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
         TabIndex        =   13
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
Attribute VB_Name = "FrmUFGLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sDate As String, eDate As String, ItemList As String, PaperList As String, SupplierList As String, AccountList As String, VchType As String, LR As Integer
Dim rstUFGLedger As New ADODB.Recordset, Reset As Long
Dim rstCompanyMaster As New ADODB.Recordset
Dim oVchType As String, Header1 As String, VchCode As String, PartyH As String, ItemH As String, OrderH As String, OrderF As Double, INWardF As Double, OUTWardF As Double, AmountF As Double, SNo As Long, aSNO As Long, pSNO As Long
Dim Opening As Double, Debit As Double, Credit As Double, Bal As Variant
Dim OrderGTF As Double, INWardGTF As Double, OUTWardGTF As Double, AmountGTF As Double
Private Sub Combo1_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Combo2_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Combo3_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub SortAction()
fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader
fpSpread1.UserColAction = UserColActionSort
fpSpread1.ColUserSortIndicator(1) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(2) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(3) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(4) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(5) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(6) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(7) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(8) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(9) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(10) = ColUserSortIndicatorDescending
End Sub
Private Sub Form_Load()
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    Me.Caption = "WIP Ledger"
    Reset = 0:
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    Call SortAction
    If VchType = 1 Then
'        Combo1.AddItem "Receipt ", 0
'        Combo1.AddItem "Purchase ", 1
'        Combo1.AddItem "Both", 2
'        Combo1.ListIndex = 2
'        Combo2.AddItem "Show IN Units ", 0
'        Combo2.AddItem "Show IN Sheets ", 1
'        Combo2.AddItem "Show IN Kg. ", 2
'        Combo2.AddItem "All ", 3
'        Combo2.ListIndex = 0
'        Combo3.AddItem "Sort By Supplier Name", 0
'        Combo3.AddItem "Sort By Party Name", 1
'        Combo3.AddItem "Sort By Paper Name", 2
'        Combo3.AddItem "Sort By Voucher No", 3
'        Combo3.AddItem "Sort By All", 4
'        Combo3.ListIndex = 4
    End If
'        Combo4.AddItem "Party (Supplied From)", 0
'        Combo4.AddItem "Party (Supplied To)", 1
'        Combo4.AddItem "Challan No.", 2
'        Combo4.AddItem "Challan Date", 3
'        Combo4.AddItem "Bill No.", 4
'        Combo4.AddItem "Bill Date", 5
'        Combo4.AddItem "Paper Name", 6
'        Combo4.AddItem "Vch. No.", 7
'        Combo4.ListIndex = 0
    Reset = 1
    If VchType = 2 Then Check2.Visible = True
: Me.Caption = " "     '11-11
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
        cmdRefresh_Click
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstUFGLedger)
    Call CloseRecordset(rstCompanyMaster)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Sub Format_Grid()
    With fpSpread1
If VchType = 1 Or VchType = 2 Then
        MhDateInput1.Visible = False
        MhDateInput2.Visible = False
        Mh3dLabel2.Visible = False
        Mh3dLabel3.Visible = False
        Mh3dLabel1.Visible = False: Combo1.Visible = False: Mh3dLabel4.Visible = False: Combo2.Visible = False: Combo3.Visible = False
        Check0.Visible = False: Check2.Visible = False: Check3.Visible = False
            .ColWidth(1) = 5
            If VchType = 1 Then .ColWidth(2) = 38.125 + 13.625 Else .ColWidth(2) = 38.125
            .ColWidth(3) = 12.5
            .ColWidth(4) = 8
            .Col = 4: .ColHidden = True
            .ColWidth(5) = 49.25
            .ColWidth(6) = 13.625
            If VchType = 1 Then .Col = 6: .ColHidden = True
            .ColWidth(7) = 8
            .Col = 7: .ColHidden = True
            .ColWidth(8) = 8
            .Col = 8: .ColHidden = True
            .ColWidth(9) = 17.875
            .Col = 10: .ColHidden = True
            .ColWidth(12) = 9.25
End If
        .ClearRange 1, 1, .MaxCols, .MaxRows, True
        .SetText 1, 0, "S.No."
    If VchType = 1 Then
    Mh3dLabel9.Caption = "UFG-LEDGER": Mh3dLabel9.Alignment = mhAlignmentCenter
        .SetText 2, 0, "FG Name"
        .SetText 3, 0, "FG Stock"
        .SetText 4, 0, "FG Code"
        .SetText 5, 0, "UFG Name"
        .SetText 6, 0, "Stock in Kg"
        .SetText 7, 0, "UFG Code"
        .SetText 8, 0, "Category"
        .SetText 9, 0, "UFG Quantity/Unit"
        .SetText 10, 0, " "
        .SetText 11, 0, "UFG Stock"
        .SetText 12, 0, "Unit"
    ElseIf VchType = 2 Then
    Mh3dLabel9.Caption = "Raw-Material-LEDGER": Mh3dLabel9.Alignment = mhAlignmentCenter: Mh3dLabel9.FontBold = True
        .SetText 2, 0, "UFG Name"
        .SetText 3, 0, "UFG Stock"
        .SetText 4, 0, "UFG Code"
        .SetText 5, 0, "RM Name"
        .SetText 6, 0, "RM Stock in Kg"
        .SetText 7, 0, "RM Code"
        .SetText 8, 0, "Category"
        .SetText 9, 0, "RM Quantity/Unit"
        .SetText 10, 0, " "
        .SetText 11, 0, "RM Stock"
        .SetText 12, 0, "Unit"
    End If
End With
End Sub
Private Sub cmdRefresh_Click()
    On Error GoTo ErrHandler
    Dim SQL As String, i As Long, C As Long, Flag As Boolean
   
    If VchType = 1 Then 'Order-wise
            '---Category='1'---BOM
                                    SQL = "Select *,0 AS RMStockKG,'' AS UOM,(SELECT Convert(Numeric(12,0),ISNULL(dbo.ufnGetItemStock('*00002',FGCODE,'XX','XXXXXX',GETDATE()),0))As Col1)As FGStock,(SELECT Convert(Numeric(12,0),ISNULL(dbo.ufnGetUFGStock('4',Right(UFGCode,6),'000000','XX','XXXXXX',GETDATE()),0))As Col1) As RMStock From(Select (Select Name From BookMaster Where C.Code=Code) As FGName,C.Code As FGCODE,C.Category,(Select Name From OutsourceItemMaster Where Code=C.Item) As UFG,(Select Code From OutsourceItemMaster Where Code=C.Item) As UFGCode,C.quantity As [UFGReq/UNIT],Format(C.quantity*(Select SUM(Quantity) From JobworkBVParent P Inner Join JobworkBVChild C ON P.Code=C.Code  Where Left(Type,2)=18),'###00.000') As [UFGReqSheets],'FS' As pvtType FROM OutsourceItemMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='1' "
    End If
    If VchType = 2 Then Flag = True
    If VchType <> 1 Then
    If Flag = False Then SQL = SQL + " UNION "
            '---Category='2'---SUB_UFG_Paper
            If Check2.Value Then SQL = SQL + "Select  Distinct ''FGName,''FGCODE,''Category,UFG,UOM,UOMValue,[Weight/Unit],UFGCode,[UFGReq/UNIT],UFGReqSheets,pvtType,FGStock,RMStock,RMStockKG From("
                                    SQL = SQL + "Select *,((PARSENAME(RMStock,2)*1)*UOMValue+(PARSENAME(RMStock,1)*1)*UOMValue)/UOMValue*[Weight/Unit] AS RMStockKG "
                                    SQL = SQL + "From ( Select *,(SELECT Convert(Numeric(12,0),ISNULL(dbo.ufnGetUFGStock('4',Right(FGCODE,6),'000000','XX','XXXXXX',GETDATE()),0))As Col1)As FGStock,(SELECT Convert(Numeric(12,3),dbo.ufnGetPaperStock('000000',Right(UFGCode,6),'XX','XXXXXX',GETDATE()))) As RMStock "
                                    SQL = SQL + "From (Select (Select Name From BookMaster Where C.Code=Code) As FGName,C.Code As FGCODE,C.Category,(Select Name From PaperMaster Where Code=C.Item) As UFG,(Select Name From GeneralMaster Where Code= I.UOM) AS UOM,(Select Value1 From GeneralMaster Where Code= I.UOM) AS UOMValue,I.[Weight/Unit],(Select Code From PaperMaster Where Code=C.Item) As UFGCode,C.quantity As [UFGReq/UNIT],Format(C.quantity*(Select SUM(Quantity) From JobworkBVParent P Inner Join JobworkBVChild C ON P.Code=C.Code  Where Left(Type,2)=18),'###00.000') As [UFGReqSheets],'FS' As pvtType FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2') AS TBL ) AS TBL "
                                    
                                    'SQL = SQL + "Select (Select Name From BookMaster Where C.Code=Code) As FGName,C.Code As FGCODE,C.Category,(Select Name From PaperMaster Where Code=C.Item) As UFG,(Select Code From PaperMaster Where Code=C.Item) As UFGCode,C.quantity As [UFGReq/UNIT],Format(C.quantity*(Select SUM(Quantity) From JobworkBVParent P Inner Join JobworkBVChild C ON P.Code=C.Code  Where Left(Type,2)=18),'###00.000') As [UFGReqSheets],'FS' As pvtType FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2'"
    End If
    If VchType = 1 Then
        SQL = SQL + " UNION "
             '---Category='4'---UFG
                                    SQL = SQL + "Select I.Name As FGNAME,I.Code As FGCODE,C.Category,(Select Name From BookMaster Where Code=C.Item) As UFG,(Select Code From BookMaster Where Code=C.Item) As UFGCode,C.quantity As [UFGReq/UNIT],Format(C.quantity*(Select SUM(Quantity) From JobworkBVParent P Inner Join JobworkBVChild C ON P.Code=C.Code  Where Left(Type,2)=18),'###00.000') As [UFGReqSheets],'FS' As pvtType FROM BookMaster I Left JOIN BookChild01 C ON I.Code=C.Code WHERE C.Category='4' ) AS TBL" & _
            " "
    End If
    If VchType = 1 Then
        SQL = SQL + "Order By FGName,Category,UFG ASC"
    ElseIf VchType = 2 Then
        If Check2.Value Then SQL = SQL + ") AS TBL Order By UFG " Else SQL = SQL + "Order By UFG "
    ElseIf VchType = 3 Then
        SQL = SQL + "Order By SUBUFG ASC"
    End If
    Format_Grid
    Screen.MousePointer = vbHourglass   'vbNormal
    If rstUFGLedger.State = adStateOpen Then rstUFGLedger.Close
    rstUFGLedger.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
    If rstUFGLedger.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
    With fpSpread1
        i = 0
        rstUFGLedger.MoveFirst
        Do While Not rstUFGLedger.EOF
            i = i + 1
            If ItemH <> rstUFGLedger.Fields("FGName").Value Then
                .SetText 2, i, rstUFGLedger.Fields("FGNAME").Value
                .SetText 3, i, Val(rstUFGLedger.Fields("FGStock").Value)
                ItemH = rstUFGLedger.Fields("FGNAME").Value
                aSNO = aSNO + 1
                C = C + 1
                .SetText 1, i, C
            End If
            .SetText 4, i, rstUFGLedger.Fields("FGCODE").Value
            .SetText 5, i, rstUFGLedger.Fields("UFG").Value
            .SetText 6, i, Val(rstUFGLedger.Fields("RMStockKG").Value)
            .SetText 7, i, rstUFGLedger.Fields("UFGCode").Value
            .SetText 8, i, Val(rstUFGLedger.Fields("Category").Value)
            .SetText 9, i, Val(rstUFGLedger.Fields("UFGREQ/UNIT").Value)
            .SetText 10, i, ""
            .SetText 11, i, Val(rstUFGLedger.Fields("RMStock").Value)
            If VchType = 1 Then .SetText 12, i, "Sheet" Else .SetText 12, i, rstUFGLedger.Fields("UOM").Value

            rstUFGLedger.MoveNext
        Loop
    .LockBackColor = RGB(245, 250, 250): Combo1.BackColor = RGB(245, 250, 250): Combo2.BackColor = RGB(245, 250, 250): Combo3.BackColor = RGB(245, 250, 250): MhDateInput1.BackColor = RGB(245, 250, 250): MhDateInput2.BackColor = RGB(245, 250, 250): 'TDBNumber1.BackColor = RGB(245, 250, 250): TDBNumber2.BackColor = RGB(245, 250, 250): Text1.BackColor = RGB(245, 250, 250):
    End With

TDBNumber2 = i
       If VchType <> 1 And VchType <> 2 Then cmdFilter_Click
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
    
End Sub
Private Sub Check2_Click() 'Show Bill Details
    With fpSpread1
            If Check2.Value Then
                .Col = 6: .ColHidden = False
                .Col = 7: .ColHidden = False
            Else
                .Col = 6: .ColHidden = True
                .Col = 7: .ColHidden = True
            End If
    End With
End Sub
Private Sub Check3_Click() 'Show Challan Details
    With fpSpread1
            If Check3.Value Then
                .Col = 4: .ColHidden = False
                .Col = 5: .ColHidden = False
            Else
                .Col = 4: .ColHidden = True
                .Col = 5: .ColHidden = True
            End If
    End With
End Sub
Private Sub Check0_Click() 'Show Paper PO
    Dim PO As Variant, i As Long
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
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub Mh3dLabel5_Click()
With fpSpread1
.LockBackColor = vbWhite
' These are 8.5" X 11" paper dimensions in TWIPS
Const PaperWidth = 12240
Const PaperHeight = 15840
Printer.PaperSize = vbPRPSA4
' Set printing options for sheet
fpSpread1.PrintAbortMsg = "Printing - Click Cancel to .Quit"
fpSpread1.PrintJobName = "Export Data" & "(" & CompCode & "_Vch-" & VchType & ")" & Format(Date, "dd-MMM-yyyy")
fpSpread1.PrintHeader = "" ' "/cPrint Header/rPage # ./p/n2nd Line"
fpSpread1.PrintFooter = "" ' "/cPrint Footer/rPage # ./p/n2nd Line"
fpSpread1.PrintBorder = True
fpSpread1.PrintColHeaders = True
fpSpread1.PrintColor = True
fpSpread1.PrintGrid = True
fpSpread1.PrintMarginTop = 1000 '1440
fpSpread1.PrintMarginBottom = 500 '1440
fpSpread1.PrintMarginLeft = 100 '720
fpSpread1.PrintMarginRight = 100 '720
'fpSpread1.PrintType = SPRD_PRINT_ALL
fpSpread1.PrintRowHeaders = True
fpSpread1.PrintShadows = True
fpSpread1.PrintUseDataMax = True
' Center vertically
fpSpread1.PrintCenterOnPageV = False
' Center horizontally
fpSpread1.PrintCenterOnPageH = True
' Perform the printing action
' Set the sheet to print
fpSpread1.Sheet = 1
' Set scaling method
fpSpread1.PrintScalingMethod = PrintScalingMethodZoom
' Set zoom factor
fpSpread1.PrintZoomFactor = 0.75
' Print
'fpSpread1.PrintSheet 0
fpSpread1.PrintOrientation = PrintOrientationLandscape
fpSpread1.PrintSheet
.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
 End With
End Sub
Private Sub Mh3dLabel6_Click()
Dim x As Boolean, FileName As String, SheetName As String, LogFileName As String
Dim R As Long, C As Long
With fpSpread1
If VchType <= 10 And VchType >= 7 Or VchType <= 28 And VchType >= 25 Then fpSpread1.InsertRows 1, 2 Else fpSpread1.InsertRows 1, 1
                R = 1
            For C = 1 To .MaxCols
                .Col = C: .Row = R: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: '.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
            Next
                .SetText 1, 1, "S.NO.": .SetText 2, 1, "Party (Supplied From)": .SetText 3, 1, "Party (Supplied To)": .SetText 4, 1, "Challan No.": .SetText 5, 1, "Challan Date": .SetText 6, 1, "Bill No.": .SetText 7, 1, "Bill Date": .SetText 8, 1, "Paper Name": .SetText 9, 1, "Vch. No.": .SetText 10, 1, "Vch. Date": .SetText 11, 1, "Vch. Ref.": .SetText 12, 1, "Weight/Unit": .SetText 13, 1, "Quantity": .SetText 14, 1, "Unit": .SetText 15, 1, "Quantity (Sheets)": .SetText 16, 1, "Quantity (Kgs.)": .SetText 17, 1, "Bundles": .SetText 18, 1, "Total Bundles": .SetText 19, 1, "Quantity (IN)": .SetText 20, 1, " Unit": .SetText 21, 1, "Quantity IN (Sheets)": .SetText 22, 1, "Quantity IN (Kgs.)": .SetText 23, 1, "Pending Qty.": .SetText 24, 1, " Unit": .SetText 25, 1, "Pending Qty. (Sheets)": .SetText 26, 1, "Pending Qty. (Kgs.)": .SetText 27, 1, "Rate/Kg.": .SetText 28, 1, "Rate/Unit": .SetText 29, 1, "Amount": .ColHeadersShow = True
                .PrintColHeaders = True: .PrintRowHeaders = True: .ColHeadersShow = True: .RowHeadersShow = True: .GridShowHoriz = True: .GridShowVert = True
'If VchType <= 10 And VchType >= 7 Or VchType <= 28 And VchType >= 25 Then .SetText 1, 2, Header1: .Col = 1: .Row = 2: .FontBold = True: .FontSize = 14: .FontUnderline = True: .ForeColor = vbRed:
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
End Sub
Private Sub cmdFilter_Click()
     Dim i As Integer, cVal As Variant, n As Integer, C As Integer
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
    With fpSpread1
    n = .DataRowCnt: StockVal = 0
        For i = 1 To .DataRowCnt 'Unhide All
            .Row = i: .RowHidden = False
        Next
        'If CheckEmpty(Text1.Text, False) Then TDBNumber2 = n - 1: Exit Sub
        'C = Combo4.ListIndex + 2
        '.SetActiveCell C, 1
        C = 8
        For i = 1 To .DataRowCnt
'        If Combo4.ListIndex = 0 Then .SetActiveCell C, i: .GetText C, i, cVal Else .SetActiveCell C, 1 ': .GetText 3, i, cVal
        .SetActiveCell C, i: .GetText C, i, cVal  'Else .SetActiveCell C, 1 ': .GetText 3, i, cVal
'                .GetText 4, i, StockVal
'                .GetText 6, i, PVal
'                .GetText 7, i, PRVal
                .GetText 8, i, PCVal
                .GetText 9, i, PRCVal
'                .GetText 10, i, SVal
'                .GetText 11, i, SRVal
'                .GetText 12, i, SCVal
'                .GetText 13, i, SRCVal
'                .GetText 14, i, SJIVal
'                .GetText 15, i, SJOVal
'                .GetText 16, i, POVal
'                .GetText 17, i, SOVal
'                .GetText 18, i, EStockVal
'                .GetText 20, i, AVal
'                .GetText 21, i, NPVal
'                .GetText 22, i, NSVal
'                .GetText 24, i, PAVal
'                .GetText 25, i, SAVal
'                .GetText 26, i, PRAVal
'                .GetText 27, i, SRAVal
'                .GetText 28, i, NPAVal
'                .GetText 29, i, NSAVal
        If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
        .Row = i: .RowHidden = True: n = n - 1
            Else
        .SetActiveCell C, i '1
        StockTotal = StockTotal + StockVal '4
'        PTotal = PTotal + PVal '6
'        PRTotal = PRTotal + PRVal '7
        PCTotal = PCTotal + PCVal '8
        PRCTotal = PRCTotal + PRCVal '9
'        STotal = STotal + SVal '10
'        SRTotal = SRTotal + SRVal '11
'        SCTotal = SCTotal + SCVal '12
'        SRCTotal = SRCTotal + SRCVal '13
'        SJITotal = SJITotal + SJIVal '14
'        SJOTotal = SJOTotal + SJOVal '15
'        POTotal = POTotal + POVal '16
'        SOTotal = SOTotal + SOVal '17
'        EStockTotal = EStockTotal + EStockVal '18
'        ATotal = ATotal + AVal '20
'        NPValTotal = NPValTotal + NPVal '21
'        NSValTotal = NSValTotal + NSVal '22
'        PAValTotal = PAValTotal + PAVal '24
'        SAValTotal = SAValTotal + SAVal '25
'        PRAValTotal = PRAValTotal + PRAVal '26
'        SRAValTotal = SRAValTotal + SRAVal '27
'        NPAValTotal = NPAValTotal + NPAVal '28
'        NSAValTotal = NSAValTotal + NSAVal '29
        End If
            TDBNumber2 = n
        Next
        .SetText 2, i, "Grand Total"
'        .SetText 4, i - 1, StockTotal
'        .SetText 6, i - 1, PTotal
'        .SetText 7, i - 1, PRTotal
        .SetText 8, i, PCTotal
        .SetText 9, i, PRCTotal
'        .SetText 10, i - 1, STotal
'        .SetText 11, i - 1, SRTotal
'        .SetText 12, i - 1, SCTotal
'        .SetText 13, i - 1, SRCTotal
'        .SetText 14, i - 1, SJITotal
'        .SetText 15, i - 1, SJOTotal
'        .SetText 16, i - 1, POTotal
'        .SetText 17, i - 1, SOTotal
'        .SetText 18, i - 1, EStockTotal
'        .SetText 20, i - 1, ATotal
'        .SetText 21, i - 1, NPValTotal
'        .SetText 22, i - 1, NSValTotal
'        .SetText 24, i - 1, PAValTotal
'        .SetText 25, i - 1, SAValTotal
'        .SetText 26, i - 1, PRAValTotal
'        .SetText 27, i - 1, SRAValTotal
'        .SetText 28, i - 1, NPAValTotal
'        .SetText 29, i - 1, NSAValTotal
'        If VchType >= 7 Then .Row = 1: .RowHidden = False:
        '.Row = i - 1: .RowHidden = False: .SelectBlockOptions = SelectBlockOptionsAll
        .SetActiveCell 1, 1
        .SetActiveCell 1, i
        For C = 1 To .MaxCols
        '.Col = C: .Row = i: .FontBold = True: .FontSize = 12: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
        .Col = C: .Row = i: .FontBold = True: .FontSize = 12: .BackColor = &H8000000F:  .ForeColor = vbBlue: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
        Next
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
'Printer.PaperSize = vbPRPSA4
' Set printing options for sheet
fpSpread1.PrintAbortMsg = "Printing - Click Cancel to .Quit"
fpSpread1.PrintJobName = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") '& ".pdf"
'fpSpread1.PrintHeader = "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy"): fpSpread1.PrintHeader=: .Font = 20 '"/cPrint Header/rPage # ./p/n2nd Line"
fpSpread1.PrintFooter = "        Export Data Company : " & rstCompanyMaster.Fields("PrintName").Value & " _(" & CompCode & "_" & PrintHeader & ")" & "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]" & " /rPage # ./p " & " Print Date : ( " & Format(Date, "dd-MMM-yyyy") & " )         ": .FontSize = 16 '& ".pdf" ' "/cPrint Footer/rPage # ./p/n2nd Line"
fpSpread1.PrintBorder = True
fpSpread1.PrintColHeaders = True
fpSpread1.PrintColor = True
fpSpread1.PrintGrid = True
fpSpread1.PrintMarginTop = 200 '750 '1440
fpSpread1.PrintMarginBottom = 200 '500 '1440
fpSpread1.PrintMarginLeft = 100 '720
fpSpread1.PrintMarginRight = 100 '720
'fpSpread1.PrintType = SPRD_PRINT_ALL
fpSpread1.PrintRowHeaders = True
fpSpread1.PrintShadows = True
fpSpread1.PrintUseDataMax = True
' Center vertically
fpSpread1.PrintCenterOnPageV = False
' Center horizontally
fpSpread1.PrintCenterOnPageH = True
' Perform the printing action
' Set the sheet to print
fpSpread1.Sheet = 1
' Set scaling method
fpSpread1.PrintScalingMethod = PrintScalingMethodZoom
' Set zoom factor
fpSpread1.PrintZoomFactor = 0.75
' Print
'fpSpread1.PrintSheet 0
fpSpread1.PrintOrientation = PrintOrientationLandscape
'fpSpread1.PrintSheet
.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
   
   'If a cell is currently active, turn off edit mode
    If fpSpread1.EditMode = True Then
        fpSpread1.EditMode = False
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
              Dim i As Integer, cVal As Variant, R As Long
        With fpSpread1
                If Text1.Text = "" Then Exit Sub
                If .DataRowCnt = 0 Then Exit Sub
                
                For i = 1 To .DataRowCnt 'Unhide All
                .Row = i: .RowHidden = False
                    Next
        
                    
                    R = IIf(.ActiveRow + 1 <> LR, .ActiveRow + 1, 1)
                    LR = R
                    For i = R To .DataRowCnt
                    If Combo4.ListIndex = Combo4.ListIndex Then .GetText Combo4.ListIndex + 2, i, cVal                                                                 'Else .SetActiveCell 3, 1: .GetText 3, i, cVal
                                If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
                                
                                Else
                                .SetActiveCell Combo4.ListIndex + 2, i: Exit Sub
                                End If
                    Next

        End With
End Sub
Private Function UFG_Publish()
Dim i As Long, dPrint As Long
'OrderPGTF = 0: INWardPGTF = 0: OUTWardPGTF = 0: AmountPGTF = 0
'OrderGTF = 0: INWardGTF = 0: OUTWardGTF = 0: AmountGTF = 0
PartyH = "": OrderH = "": ItemH = "": INWardF = 0: OUTWardF = 0: SNo = 0: aSNO = 0: pSNO = 0: OrderF = 0: Bal = 0: AmountF = 0
    With fpSpread1
    .RowHeadersAutoText = DispBlank
        rstUFGLedger.MoveFirst
        Do While Not rstUFGLedger.EOF
        If VchType = 1 Then
            i = i + 1
            If PartyH <> rstUFGLedger.Fields("Name").Value Then
                aSNO = aSNO + 1
                .SetText 0, i, "Item-" & aSNO
'Party Header
                .SetText 2, i, "Item : " + rstUFGLedger.Fields("Name").Value: .Col = 2: .Row = i: .FontBold = True: .FontSize = 12: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbRed: pSNO = 0
                fpSpread1.Col = 2: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstUFGLedger.Fields("Name").Value) > 33 Then fpSpread1.RowHeight(i) = 36: fpSpread1.TypeHAlign = TypeHAlignRight
                PartyH = rstUFGLedger.Fields("Name").Value
                If i > 2 Then i = i + 1
            End If
            If OrderH <> rstUFGLedger.Fields("UFG").Value And (VchType = 1) And rstUFGLedger.Fields("UFG").Value <> "" Then
                If i > 2 Then
                .SetText 0, i, " "
'SUBTOTAL Footer
                .SetText 5, i, "SUBTOTAL": .SetText 6, i, OrderF: .SetText 8, i, INWardF: .SetText 23, i, OUTWardF: .SetText 24, i, Bal: .SetText 25, i, "Units": .SetText 27, i, AmountF: INWardGTF = INWardGTF + INWardF: INWardF = 0: OUTWardGTF = OUTWardGTF + OUTWardF: OUTWardF = 0: OrderGTF = OrderGTF + OrderF: OrderF = 0: AmountGTF = AmountGTF + AmountF: AmountF = 0: SNo = 0
                .Col = 5: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True: .TypeHAlign = TypeHAlignRight:
                .Col = 6: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                .Col = 8: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                .Col = 23: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                .Col = 24: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                .Col = 25: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                .Col = 27: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                End If
                Bal = 0: Credit = 0: Debit = 0: i = i + 1
                pSNO = pSNO + 1
'Order No Header
                If VchType = 1 And rstUFGLedger.Fields("pvtType").Value = "FP" Then VchCode = "S": Bal = Val(rstUFGLedger.Fields("Ordered").Value) * IIf(VchCode = "S", -1, 1)
                .SetText 0, i, IIf(VchCode = "S", "P", "S") & "-" & pSNO
                .SetText 2, i, rstUFGLedger.Fields("UFG").Value: .Col = 5: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue:
                .SetText 6, i, Val(rstUFGLedger.Fields("Ordered").Value): .Col = 6: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue:
                .SetText 32, i, Trim(rstUFGLedger.Fields("pvtCode").Value)
                .SetText 35, i, Trim(rstUFGLedger.Fields("pvtType").Value)
                OrderF = Val(rstUFGLedger.Fields("Ordered").Value)
                OrderH = rstUFGLedger.Fields("VchBillNo").Value
                If VchType <> 45 Then i = i + 1
            End If
        ElseIf VchType = 36 Or VchType = 38 Then
                    i = i + 1
            If PartyH <> rstUFGLedger.Fields("AccountName").Value Then
                If i > 2 Then
                .SetText 0, i, " "
                    .SetText 5, i, "SUBTOTAL": .SetText 6, i, OrderF: .SetText 8, i, INWardF: .SetText 23, i, OUTWardF:: If VchType = 36 Or VchType = 38 Then .SetText 24, i, (IIf(VchCode = "S", -1, 1) * OrderF) - OUTWardF + INWardF: .SetText 25, i, "Units": .SetText 27, i, AmountF: INWardGTF = INWardGTF + INWardF: INWardF = 0: OUTWardGTF = OUTWardGTF + OUTWardF: OUTWardF = 0: OrderGTF = OrderGTF + OrderF: OrderF = 0: AmountGTF = AmountGTF + AmountF: AmountF = 0: SNo = 0
                    .Col = 5: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True: .TypeHAlign = TypeHAlignRight
                    .Col = 6: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 8: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 23: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 24: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 25: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 27: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    i = i + 1
                End If
                aSNO = aSNO + 1
                .SetText 0, i, "A/C-" & aSNO
                .SetText 5, i, "Party : " + rstUFGLedger.Fields("AccountName").Value: .Col = 5: .Row = i: .FontBold = True: .FontSize = 14: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbRed:
                fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstUFGLedger.Fields("AccountName").Value) > 33 Then fpSpread1.RowHeight(i) = 36: fpSpread1.TypeHAlign = TypeHAlignRight
                PartyH = rstUFGLedger.Fields("AccountName").Value
                i = i + 1
            End If
        ElseIf VchType = 40 Or VchType = 43 Then
                    i = i + 1
            If PartyH <> rstUFGLedger.Fields("AccountName").Value Then
                If i > 2 Then
                .SetText 0, i, " "
                    .SetText 5, i, "SUBTOTAL": .SetText 6, i, OrderF: .SetText 8, i, INWardF: .SetText 23, i, OUTWardF: If VchType = 40 Or VchType = 43 Then .SetText 24, i, Bal: .SetText 25, i, "Units": .SetText 27, i, AmountF: INWardGTF = INWardGTF + INWardF: INWardF = 0: OUTWardGTF = OUTWardGTF + OUTWardF: OUTWardF = 0: OrderGTF = OrderGTF + OrderF: OrderF = 0: AmountGTF = AmountGTF + AmountF: AmountF = 0: SNo = 0: Bal = 0
                    .Col = 5: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True: .TypeHAlign = TypeHAlignRight
                    .Col = 6: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 8: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 23: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 24: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 25: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 27: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    i = i + 1
                End If
                aSNO = aSNO + 1
                .SetText 0, i, "A/C-" & aSNO
                .SetText 5, i, "Party : " + rstUFGLedger.Fields("AccountName").Value: .Col = 5: .Row = i: .FontBold = True: .FontSize = 14: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbRed:
                fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstUFGLedger.Fields("AccountName").Value) > 33 Then fpSpread1.RowHeight(i) = 36: fpSpread1.TypeHAlign = TypeHAlignRight
                PartyH = rstUFGLedger.Fields("AccountName").Value
                i = i + 1
            End If
        ElseIf VchType = 41 Or VchType = 44 Then
                    i = i + 1
            If ItemH <> rstUFGLedger.Fields("ItemName").Value Then
                If i > 2 Then
                .SetText 0, i, " "
                    .SetText 5, i, "SUBTOTAL": .SetText 6, i, OrderF: .SetText 8, i, INWardF: .SetText 23, i, OUTWardF: If VchType = 41 Or VchType = 44 Then .SetText 24, i, Bal: .SetText 25, i, "Units": .SetText 27, i, AmountF: INWardGTF = INWardGTF + INWardF: INWardF = 0: OUTWardGTF = OUTWardGTF + OUTWardF: OUTWardF = 0: OrderGTF = OrderGTF + OrderF: OrderF = 0: AmountGTF = AmountGTF + AmountF: AmountF = 0: SNo = 0: Bal = 0
                    .Col = 5: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True: .TypeHAlign = TypeHAlignRight
                    .Col = 6: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 8: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 23: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 24: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 25: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    .Col = 27: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                    i = i + 1
                End If
                aSNO = aSNO + 1
                .SetText 0, i, "I-" & aSNO
                .SetText 5, i, "Item : " + rstUFGLedger.Fields("ItemName").Value: .Col = 5: .Row = i: .FontBold = True: .FontSize = 14: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbRed:
                fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstUFGLedger.Fields("ItemName").Value) > 33 Then fpSpread1.RowHeight(i) = 36: fpSpread1.TypeHAlign = TypeHAlignRight
                ItemH = rstUFGLedger.Fields("ItemName").Value
                i = i + 1
            End If
        End If
'Pending Order
        If VchType = 34 Or VchType = 35 Or VchType = 37 Or VchType = 45 Then
        If VchType = 34 Or VchType = 35 Or VchType = 37 Or VchType = 45 And rstUFGLedger.Fields("VchBillNo").Value = "" Then
            SNo = SNo + 1
            .SetText 0, i, SNo
            .SetText 1, i, rstUFGLedger.Fields("vtDate").Value
            .SetText 2, i, Trim(rstUFGLedger.Fields("vtNo").Value)
            .SetText 3, i, rstUFGLedger.Fields("TypeRef").Value: fpSpread1.Col = 3: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstUFGLedger.Fields("TypeRef").Value) > 10 Then fpSpread1.RowHeight(i) = 25.5:
            .SetText 5, i, rstUFGLedger.Fields("MaterialCentre").Value & IIf(rstUFGLedger.Fields("Remarks") <> "" Or rstUFGLedger.Fields("ChallanNo") <> "", " ->> ", "") & IIf(rstUFGLedger.Fields("Remarks") <> "", " RemarK : " & rstUFGLedger.Fields("Remarks"), "") & IIf(rstUFGLedger.Fields("ChallanNo") <> "", " (Ch.No." + rstUFGLedger.Fields("ChallanNo") & "_ Ch. dt." & rstUFGLedger.Fields("ChallanDate") & ")", ""): fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstUFGLedger.Fields("MaterialCentre").Value & IIf(rstUFGLedger.Fields("RemarkS") <> "", " -> RemarK : " & rstUFGLedger.Fields("Remarks"), "") & IIf(rstUFGLedger.Fields("ChallanNo") <> "", " (Ch.No." + rstUFGLedger.Fields("ChallanNo") & "_ Ch. dt." & rstUFGLedger.Fields("ChallanDate") & ")", "")) > 75 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignRight
                Credit = Val(rstUFGLedger.Fields("INward").Value)
                INWardF = INWardF + Credit
            .SetText 8, i, Val(rstUFGLedger.Fields("INward").Value)
                Debit = Val(rstUFGLedger.Fields("OutWard").Value)
                OUTWardF = OUTWardF + Debit
            .SetText 23, i, Val(rstUFGLedger.Fields("OutWard").Value)
                Bal = Bal + Credit - Debit
            .SetText 24, i, Bal
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstUFGLedger.Fields("Rate").Value)
                AmountF = AmountF + Val(rstUFGLedger.Fields("Amount").Value)
            .SetText 27, i, Val(rstUFGLedger.Fields("Amount").Value)
            .SetText 32, i, Trim(rstUFGLedger.Fields("vtCode").Value)
            .SetText 35, i, rstUFGLedger.Fields("vtType").Value
            dPrint = dPrint + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstUFGLedger.RecordCount & " !!!"
        End If
        ElseIf VchType = 36 Or VchType = 38 Then
            SNo = SNo + 1
            .SetText 0, i, SNo
            .SetText 1, i, rstUFGLedger.Fields("VchDate").Value
            .SetText 2, i, rstUFGLedger.Fields("VchBillNo").Value
            .SetText 3, i, IIf(VchCode = "S", "Purchase Order", "Sales Order")
            .SetText 5, i, rstUFGLedger.Fields("ItemName").Value: fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstUFGLedger.Fields("ItemName").Value) > 48 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignRight
                OrderF = OrderF + Val(rstUFGLedger.Fields("Ordered").Value)
            .SetText 6, i, Val(rstUFGLedger.Fields("Ordered").Value)
                Credit = Val(rstUFGLedger.Fields("INward").Value)
                INWardF = INWardF + Credit
            .SetText 8, i, Val(rstUFGLedger.Fields("INward").Value)
                Debit = Val(rstUFGLedger.Fields("OutWard").Value)
                OUTWardF = OUTWardF + Debit
            .SetText 23, i, Val(rstUFGLedger.Fields("OutWard").Value)
            .SetText 24, i, (IIf(VchCode = "S", -1, 1) * Val(rstUFGLedger.Fields("Ordered").Value)) - Val(rstUFGLedger.Fields("OutWard").Value) + Val(rstUFGLedger.Fields("INward").Value)
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstUFGLedger.Fields("Rate").Value)
                AmountF = AmountF + Val(rstUFGLedger.Fields("Amount").Value)
            .SetText 27, i, Val(rstUFGLedger.Fields("Amount").Value)
            .SetText 32, i, rstUFGLedger.Fields("vtCode").Value
            .SetText 35, i, rstUFGLedger.Fields("vtType").Value
            dPrint = dPrint + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstUFGLedger.RecordCount & " !!!"
        ElseIf VchType = 39 Or VchType = 42 Then
            SNo = SNo + 1
            i = i + 1
            .SetText 0, i, SNo
            .SetText 1, i, rstUFGLedger.Fields("VchDate").Value
            .SetText 2, i, rstUFGLedger.Fields("VchBillNo").Value
            .SetText 3, i, IIf(VchCode = "S", "Purchase Order", "Sales Order")
            .SetText 5, i, rstUFGLedger.Fields("ItemName").Value: fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstUFGLedger.Fields("ItemName").Value) > 75 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignLeft
                OrderF = OrderF + Val(rstUFGLedger.Fields("Ordered").Value)
            .SetText 6, i, Val(rstUFGLedger.Fields("Ordered").Value)
                Credit = Val(rstUFGLedger.Fields("Dispatched").Value)
                INWardF = INWardF + Credit
            .SetText 8, i, Val(rstUFGLedger.Fields("Dispatched").Value)
                Debit = Val(rstUFGLedger.Fields("Dispatched").Value)
                OUTWardF = OUTWardF + Debit
            .SetText 23, i, Val(rstUFGLedger.Fields("Dispatched").Value)
            .SetText 24, i, Val(rstUFGLedger.Fields("Balance").Value)
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstUFGLedger.Fields("Rate").Value)
                AmountF = AmountF + Val(rstUFGLedger.Fields("Amount").Value)
            .SetText 27, i, Val(rstUFGLedger.Fields("Amount").Value)
            .SetText 32, i, rstUFGLedger.Fields("vtCode").Value
            .SetText 35, i, rstUFGLedger.Fields("iCode").Value
            dPrint = dPrint + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstUFGLedger.RecordCount & " !!!"
        ElseIf VchType = 40 Or VchType = 43 Then
            SNo = SNo + 1
            .SetText 0, i, SNo
            .SetText 1, i, rstUFGLedger.Fields("VchDate").Value
            .SetText 2, i, rstUFGLedger.Fields("VchBillNo").Value
            .SetText 3, i, IIf(VchCode = "S", "Purchase Order", "Sales Order")
            .SetText 5, i, rstUFGLedger.Fields("ItemName").Value: fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstUFGLedger.Fields("ItemName").Value) > 48 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignRight
                OrderF = OrderF + Val(rstUFGLedger.Fields("Ordered").Value)
            .SetText 6, i, Val(rstUFGLedger.Fields("Ordered").Value)
                Credit = Val(rstUFGLedger.Fields("Dispatched").Value)
                INWardF = INWardF + Credit
            .SetText 8, i, Val(rstUFGLedger.Fields("Dispatched").Value)
                Debit = Val(rstUFGLedger.Fields("Dispatched").Value)
                OUTWardF = OUTWardF + Debit
            .SetText 23, i, Val(rstUFGLedger.Fields("Dispatched").Value)
                Bal = Bal + Val(rstUFGLedger.Fields("Balance").Value)
            .SetText 24, i, Val(rstUFGLedger.Fields("Balance").Value)
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstUFGLedger.Fields("Rate").Value)
                AmountF = AmountF + Val(rstUFGLedger.Fields("Amount").Value)
            .SetText 27, i, Val(rstUFGLedger.Fields("Amount").Value)
            .SetText 32, i, rstUFGLedger.Fields("vtCode").Value
            dPrint = dPrint + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstUFGLedger.RecordCount & " !!!"
        ElseIf VchType = 41 Or VchType = 44 Then
            SNo = SNo + 1
            .SetText 0, i, SNo
            .SetText 1, i, rstUFGLedger.Fields("VchDate").Value
            .SetText 2, i, rstUFGLedger.Fields("VchBillNo").Value
            .SetText 3, i, IIf(VchCode = "S", "Purchase Order", "Sales Order")
            .SetText 5, i, rstUFGLedger.Fields("AccountName").Value: fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstUFGLedger.Fields("ItemName").Value) > 48 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignRight
                OrderF = OrderF + Val(rstUFGLedger.Fields("Ordered").Value)
            .SetText 6, i, Val(rstUFGLedger.Fields("Ordered").Value)
                Credit = Val(rstUFGLedger.Fields("Dispatched").Value)
                INWardF = INWardF + Credit
            .SetText 8, i, Val(rstUFGLedger.Fields("Dispatched").Value)
                Debit = Val(rstUFGLedger.Fields("Dispatched").Value)
                OUTWardF = OUTWardF + Debit
            .SetText 23, i, Val(rstUFGLedger.Fields("Dispatched").Value)
                Bal = Bal + Val(rstUFGLedger.Fields("Balance").Value)
            .SetText 24, i, Val(rstUFGLedger.Fields("Balance").Value)
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstUFGLedger.Fields("Rate").Value)
                AmountF = AmountF + Val(rstUFGLedger.Fields("Amount").Value)
            .SetText 27, i, Val(rstUFGLedger.Fields("Amount").Value)
            .SetText 32, i, rstUFGLedger.Fields("vtCode").Value
            dPrint = dPrint + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstUFGLedger.RecordCount & " !!!"
        End If
NXT:
            rstUFGLedger.MoveNext
            If MdiMainMenu.ProgressBar1.Value + Round((100 / rstUFGLedger.RecordCount), 2) <= 100 Then
                MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + Round((100 / rstUFGLedger.RecordCount), 2)
            End If
        Loop
        If i > 2 Then
            i = i + 1: .SetText 0, i, " ": .SetText 0, i + 1, " ": .SetText 0, i + 2, " "
            If VchType = 39 Or VchType = 42 Then .SetText 5, i, "TOTAL" Else .SetText 5, i, "SUBTOTAL"
            .SetText 6, i, OrderF: .SetText 8, i, INWardF: .SetText 23, i, OUTWardF: .SetText 24, i, Bal: .SetText 25, i, "Units": .SetText 27, i, AmountF: If VchType = 36 Or VchType = 38 Then .SetText 24, i, (IIf(VchCode = "S", -1, 1) * OrderF) - OUTWardF + INWardF
            .Col = 5: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True: .TypeHAlign = TypeHAlignRight
            .Col = 6: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 8: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 23: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 24: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 25: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 27: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
       End If
        INWardGTF = INWardGTF + INWardF: INWardF = 0: OUTWardGTF = OUTWardGTF + OUTWardF: OUTWardF = 0: OrderGTF = OrderGTF + OrderF: OrderF = 0: AmountGTF = AmountGTF + AmountF: AmountF = 0:
         .SetText 5, i + 1, "Grand TOTAL": .SetText 6, i + 1, OrderGTF: .SetText 8, i + 1, INWardGTF: .SetText 23, i + 1, OUTWardGTF: .SetText 24, i + 1, (IIf(VchCode = "S", -1, 1) * OrderGTF) - OUTWardGTF + INWardGTF: .SetText 25, i + 1, "Units": .SetText 27, i + 1, AmountGTF: If VchType = 36 Or VchType = 38 Then .SetText 24, i, (IIf(VchCode = "S", -1, 1) * OrderGTF) - OUTWardGTF + INWardGTF
            .Col = 5: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True: .TypeHAlign = TypeHAlignRight
            .Col = 6: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 8: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 23: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 24: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 25: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 27: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
End With
End Function

