VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmPaperLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Paper Ledger"
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
      Picture         =   "PaperLedger.frx":0000
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
         Picture         =   "PaperLedger.frx":001C
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
         Picture         =   "PaperLedger.frx":035E
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
         Picture         =   "PaperLedger.frx":06A0
         Picture         =   "PaperLedger.frx":06BC
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
         TabIndex        =   12
         Top             =   158
         Width           =   1950
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   375
         Left            =   19380
         Picture         =   "PaperLedger.frx":06D8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel"
         Top             =   90
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   375
         Left            =   19000
         Picture         =   "PaperLedger.frx":07DA
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
         SpreadDesigner  =   "PaperLedger.frx":0924
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
         Picture         =   "PaperLedger.frx":19B2
         Picture         =   "PaperLedger.frx":19CE
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
         Picture         =   "PaperLedger.frx":19EA
         Picture         =   "PaperLedger.frx":1A06
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
         Calendar        =   "PaperLedger.frx":1A22
         Caption         =   "PaperLedger.frx":1B3A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperLedger.frx":1BA6
         Keys            =   "PaperLedger.frx":1BC4
         Spin            =   "PaperLedger.frx":1C22
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
         Calendar        =   "PaperLedger.frx":1C4A
         Caption         =   "PaperLedger.frx":1D62
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperLedger.frx":1DCE
         Keys            =   "PaperLedger.frx":1DEC
         Spin            =   "PaperLedger.frx":1E4A
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
         Picture         =   "PaperLedger.frx":1E72
         Picture         =   "PaperLedger.frx":1E8E
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
         Picture         =   "PaperLedger.frx":1EAA
         Picture         =   "PaperLedger.frx":1EC6
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
         Picture         =   "PaperLedger.frx":1EE2
         Picture         =   "PaperLedger.frx":1EFE
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
         Calculator      =   "PaperLedger.frx":1F1A
         Caption         =   "PaperLedger.frx":1F3A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperLedger.frx":1F9E
         Keys            =   "PaperLedger.frx":1FBC
         Spin            =   "PaperLedger.frx":2006
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
         Picture         =   "PaperLedger.frx":202E
         Picture         =   "PaperLedger.frx":204A
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
         Picture         =   "PaperLedger.frx":2066
         Picture         =   "PaperLedger.frx":2082
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
Attribute VB_Name = "FrmPaperLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sDate As String, eDate As String, ItemList As String, PaperList As String, SupplierList As String, AccountList As String, VchType As String, LR As Integer
Dim rstPaperLedger As New ADODB.Recordset, Reset As Long
Dim rstCompanyMaster As New ADODB.Recordset
Private Sub Combo1_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Combo2_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Combo3_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Form_Load()
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    Reset = 0:
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    
fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader
fpSpread1.UserColAction = UserColActionSort
fpSpread1.ColUserSortIndicator(1) = ColUserSortIndicatorDisabled
fpSpread1.ColUserSortIndicator(2) = ColUserSortIndicatorDisabled
fpSpread1.ColUserSortIndicator(3) = ColUserSortIndicatorDisabled
fpSpread1.ColUserSortIndicator(4) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(5) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(6) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(7) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(8) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(9) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(10) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(11) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(12) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(13) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(14) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(15) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(16) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(17) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(18) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(19) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(20) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(21) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(22) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(23) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(24) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(25) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(26) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(27) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(28) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(29) = ColUserSortIndicatorDescending
fpSpread1.ColUserSortIndicator(30) = ColUserSortIndicatorDescending
    
    If VchType = 11 Then
    Combo1.AddItem "Receipt ", 0
    Combo1.AddItem "Purchase ", 1
    Combo1.AddItem "Both", 2
    Combo1.ListIndex = 2
    Combo2.AddItem "Show IN Units ", 0
    Combo2.AddItem "Show IN Sheets ", 1
    Combo2.AddItem "Show IN Kg. ", 2
    Combo2.AddItem "All ", 3
    Combo2.ListIndex = 0
    Combo3.AddItem "Sort By Supplier Name", 0
    Combo3.AddItem "Sort By Party Name", 1
    Combo3.AddItem "Sort By Paper Name", 2
    Combo3.AddItem "Sort By Voucher No", 3
    Combo3.AddItem "Sort By All", 4
    Combo3.ListIndex = 4
    ElseIf VchType = 12 Or VchType = 13 Then
    Combo1.AddItem "Against Receipt ", 0
    Combo1.AddItem "Against Purchase ", 1
    Combo1.AddItem "Against Both", 2
    Combo1.ListIndex = 2
    Combo2.AddItem "Show IN Units ", 0
    Combo2.AddItem "Show IN Sheets ", 1
    Combo2.AddItem "Show IN Kg. ", 2
    Combo2.AddItem "All ", 3
    Combo2.ListIndex = 0
    Combo3.AddItem "Sort By Supplier Name", 0
    Combo3.AddItem "Sort By Party Name", 1
    Combo3.AddItem "Sort By Paper Name", 2
    Combo3.AddItem "Sort By Voucher No", 3
    Combo3.AddItem "Sort By All", 4
    Combo3.ListIndex = 4
    ElseIf VchType = 14 Or VchType = 15 Or VchType = 16 Then
    Combo1.AddItem "Issue ", 0
    Combo1.AddItem "Sale ", 1
    Combo1.AddItem "Both", 2
    Combo1.ListIndex = 2
    ElseIf VchType = 17 Then
    Combo1.AddItem "Inward ", 0
    Combo1.AddItem "Outward ", 1
    Combo1.AddItem "Both", 2
    Combo1.ListIndex = 2
    Combo2.AddItem "Show IN Units ", 0
    Combo2.AddItem "Show IN Sheets ", 1
    Combo2.AddItem "Show IN Kg. ", 2
    Combo2.AddItem "All ", 3
    Combo2.ListIndex = 0
    Combo3.AddItem "Sort By From Account Name", 0
    Combo3.AddItem "Sort By To Account Name", 1
    Combo3.AddItem "Sort By Paper Name", 2
    Combo3.AddItem "Sort By Voucher No", 3
    Combo3.AddItem "Sort By All", 4
    Combo3.ListIndex = 4
    ElseIf VchType = 18 Then
    Combo1.AddItem "Purchase Order ", 0
    Combo1.AddItem "Pending Purchase Order ", 1
    Combo1.AddItem "Both", 2
    Combo1.ListIndex = 2
    Combo2.AddItem "Show IN Units ", 0
    Combo2.AddItem "Show IN Sheets ", 1
    Combo2.AddItem "Show IN Kg. ", 2
    Combo2.AddItem "All ", 3
    Combo2.ListIndex = 0
    Combo3.AddItem "Sort By Supplier Name", 0
    Combo3.AddItem "Sort By Paper Name", 1
    Combo3.AddItem "Sort By Voucher No", 2
    Combo3.AddItem "Sort By All", 3
    Combo3.ListIndex = 3
    End If
    Combo4.AddItem "Party (Supplied From)", 0
    Combo4.AddItem "Party (Supplied To)", 1
    Combo4.AddItem "Challan No.", 2
    Combo4.AddItem "Challan Date", 3
    Combo4.AddItem "Bill No.", 4
    Combo4.AddItem "Bill Date", 5
    Combo4.AddItem "Paper Name", 6
    Combo4.AddItem "Vch. No.", 7
    Combo4.ListIndex = 0
    
    Reset = 1
    If VchType = 11 Then Check0.Caption = "Show Paper Receipt Against 'PO' Only": Me.Caption = "Paper Receipt Party-Wise" '11-11
    If VchType = 12 Then Combo1.Width = 2000: Check0.value = 1: Check0.Visible = False: Me.Caption = "Paper Receipt Order-Wise" '13-12
    If VchType = 13 Then Combo1.Width = 2000: Check0.value = 1: Check0.Visible = False: Me.Caption = "Paper Receipt Without-Order" '15-13
    If VchType = 14 Then Check0.Caption = "Show Paper Issue Against 'SO' Only": Check2.Visible = False: Me.Caption = "Paper Issue Party-Wise" '12-14
    If VchType = 15 Then Check0.Caption = "Show Paper Issue Against 'SO' Only": Check2.Visible = False: Me.Caption = "Paper Issue Order-Wise" '14-15
    If VchType = 16 Then Check0.Caption = "Show Paper Issue Against 'SO' Only": Check2.Visible = False: Me.Caption = "Paper Issue Without-Order" '16-16
    If VchType = 17 Then Check0.Caption = "Show Paper Issue Against 'SO' Only": Check2.Visible = False: Me.Caption = "Paper Transfer Party-Wise" '17
    If VchType = 18 Then Combo1.Width = 2250: Check0.Left = 7000: Check0.Caption = "Show Paper Pending 'PO' Only": Check3.Visible = True: Check2.Visible = True: Me.Caption = "Paper Pending Order Supplier-Wise" '18
    MhDateInput1.value = Format(sDate, "dd-MM-yyyy")
    MhDateInput2.value = Format(eDate, "dd-MM-yyyy")
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
    End If
    If DatabaseType = "MS SQL" Then SQL = Replace(SQL, "#", "'")
    Screen.MousePointer = vbHourglass   'vbNormal
    If rstPaperLedger.State = adStateOpen Then rstPaperLedger.Close
    rstPaperLedger.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
    If rstPaperLedger.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
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
        ElseIf VchType = 18 Or VchType = 19 Then
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
            .SetText 2, i, rstPaperLedger.Fields("FromAccountName").value: If Len(rstPaperLedger.Fields("FromAccountName").value) > 25 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: fpSpread1.TypeHAlign = TypeHAlignRight
            .SetText 3, i, rstPaperLedger.Fields("PartyAccountName").value: If Len(rstPaperLedger.Fields("PartyAccountName").value) > 25 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: fpSpread1.TypeHAlign = TypeHAlignRight
            .SetText 4, i, rstPaperLedger.Fields("ChallanNo").value
            .SetText 5, i, rstPaperLedger.Fields("ChallanDate").value
            .SetText 6, i, rstPaperLedger.Fields("BillNO").value
            .SetText 7, i, rstPaperLedger.Fields("BillDate").value
            .SetText 8, i, rstPaperLedger.Fields("Paper").value: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstPaperLedger.Fields("Paper").value) > 75 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignRight
            .SetText 9, i, rstPaperLedger.Fields("VchRef").value
            .SetText 10, i, rstPaperLedger.Fields("Date").value
            .SetText 11, i, rstPaperLedger.Fields("Ref").value
            .SetText 12, i, Val(rstPaperLedger.Fields("wtUOM").value)
            .SetText 13, i, IIf(Val(rstPaperLedger.Fields("Quantity").value) = 0, "", Val(rstPaperLedger.Fields("Quantity").value))
            .SetText 14, i, rstPaperLedger.Fields("UOM").value
            .SetText 15, i, IIf(Val(rstPaperLedger.Fields("QuantitySheets").value) = 0, "", Val(rstPaperLedger.Fields("QuantitySheets").value))
            .SetText 16, i, IIf(Val(rstPaperLedger.Fields("QuantityKg").value) = 0, "", Val(rstPaperLedger.Fields("QuantityKg").value))
            .SetText 17, i, Val(rstPaperLedger.Fields("Units/Bundle").value)
            .SetText 18, i, Val(rstPaperLedger.Fields("TotalBundles").value)
            If VchType = 11 Or VchType = 12 Or VchType = 13 Or VchType = 18 Or VchType = 19 Then
            .SetText 19, i, IIf(Val(rstPaperLedger.Fields("QuantityIssue").value) = 0, "", Val(rstPaperLedger.Fields("QuantityIssue").value))
            .SetText 20, i, rstPaperLedger.Fields("UOM").value
            .SetText 21, i, IIf(Val(rstPaperLedger.Fields("IssueQtySheets").value) = 0, "", Val(rstPaperLedger.Fields("IssueQtySheets").value))
            .SetText 22, i, IIf(Val(rstPaperLedger.Fields("IssueQtyKg").value) = 0, "", Val(rstPaperLedger.Fields("IssueQtyKg").value))
            End If
            If VchType = 18 Or VchType = 19 Then
            .SetText 23, i, IIf(Val(rstPaperLedger.Fields("PendingQty").value) = 0, "", Val(rstPaperLedger.Fields("PendingQty").value))
            .SetText 24, i, rstPaperLedger.Fields("UOM").value
            .SetText 25, i, IIf(Val(rstPaperLedger.Fields("PendingQtySheets").value) = 0, "", Val(rstPaperLedger.Fields("PendingQtySheets").value))
            .SetText 26, i, IIf(Val(rstPaperLedger.Fields("PendingQtyKG").value) = 0, "", Val(rstPaperLedger.Fields("PendingQtyKG").value))
            .SetText 27, i, Val(rstPaperLedger.Fields("Rate/Kg").value)
            .SetText 28, i, Val(rstPaperLedger.Fields("Rate/Unit").value)
            .SetText 29, i, Val(rstPaperLedger.Fields("Amount").value)
            End If
            rstPaperLedger.MoveNext
        Loop
    .LockBackColor = RGB(245, 250, 250): Combo1.BackColor = RGB(245, 250, 250): Combo2.BackColor = RGB(245, 250, 250): Combo3.BackColor = RGB(245, 250, 250): MhDateInput1.BackColor = RGB(245, 250, 250): MhDateInput2.BackColor = RGB(245, 250, 250): 'TDBNumber1.BackColor = RGB(245, 250, 250): TDBNumber2.BackColor = RGB(245, 250, 250): Text1.BackColor = RGB(245, 250, 250):
    End With
TDBNumber2 = i
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
    
End Sub
Private Sub Check2_Click() 'Show Bill Details
    With fpSpread1
            If Check2.value Then
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
            If Check3.value Then
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
            If Check0.value Then
                .GetText 26, i, PO
                If PO = 0 Then .Row = i: .RowHidden = True Else .Row = i: .RowHidden = False
            ElseIf Check0.value Then
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
            If Check0.value Then
                .GetText 11, i, PO
                If Left(PO, 3) = "Pur" Then .Row = i: .RowHidden = False Else .Row = i: .RowHidden = True
            ElseIf Check0.value Then
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
Dim x As Boolean, fileName As String, SheetName As String, LogFileName As String
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
fileName = "Export Data" & "(" & CompCode & "_Vch-" & VchType & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"
SheetName = "Export Data" & "(" & CompCode & "_Vch-" & VchType & ")"
LogFileName = "Export Data" & "(" & CompCode & "_Vch-" & VchType & ")" & Format(Date, "dd-MMM-yyyy") & ".txt"
x = fpSpread1.ExportToExcelEx(fileName, SheetName, LogFileName, ExcelSaveFlagNoFormulas)
' Display result to user based on T/F value of x
If x = True Then
MsgBox "Export complete.", vbInformation, "Easy Publish...Export !!! "
    Dim oExcel As Object
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\" & fileName)
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
        If CheckEmpty(Text1.Text, False) Then TDBNumber2 = n - 1: Exit Sub
        C = Combo4.ListIndex + 2
        .SetActiveCell C, 1
        For i = 1 To .DataRowCnt
'        If Combo4.ListIndex = 0 Then .SetActiveCell C, i: .GetText C, i, cVal Else .SetActiveCell C, 1 ': .GetText 3, i, cVal
        .SetActiveCell C, i: .GetText C, i, cVal  'Else .SetActiveCell C, 1 ': .GetText 3, i, cVal
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
        If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
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
fpSpread1.PrintFooter = "        Export Data Company : " & rstCompanyMaster.Fields("PrintName").value & " _(" & CompCode & "_" & PrintHeader & ")" & "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]" & " /rPage # ./p " & " Print Date : ( " & Format(Date, "dd-MMM-yyyy") & " )         ": .FontSize = 16 '& ".pdf" ' "/cPrint Footer/rPage # ./p/n2nd Line"
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
