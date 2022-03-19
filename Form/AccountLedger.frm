VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmAccountLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Account Ledger"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19485
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
   ScaleHeight     =   9255
   ScaleWidth      =   19485
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   18600
      Picture         =   "AccountLedger.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Refresh"
      Top             =   210
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   9030
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   19290
      _Version        =   65536
      _ExtentX        =   34025
      _ExtentY        =   15928
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
      Picture         =   "AccountLedger.frx":014A
      Begin VB.CommandButton Preview 
         Caption         =   "&Print Preview"
         Height          =   330
         Left            =   15600
         TabIndex        =   24
         Top             =   8640
         Width           =   1215
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   2
         Left            =   8760
         TabIndex        =   22
         Top             =   8625
         Width           =   6735
         _Version        =   65536
         _ExtentX        =   11880
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
         Picture         =   "AccountLedger.frx":0166
         Picture         =   "AccountLedger.frx":0182
      End
      Begin VB.CommandButton Command2 
         Height          =   320
         Left            =   6000
         Picture         =   "AccountLedger.frx":019E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Search"
         Top             =   8620
         Width           =   375
      End
      Begin VB.CommandButton cmdFilter 
         Height          =   320
         Left            =   5520
         Picture         =   "AccountLedger.frx":04E0
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Filter"
         Top             =   8620
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
         TabIndex        =   14
         ToolTipText     =   "Find And Search"
         Top             =   8620
         Width           =   2190
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   375
         Left            =   18840
         Picture         =   "AccountLedger.frx":0822
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancel"
         Top             =   90
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   375
         Left            =   18480
         Picture         =   "AccountLedger.frx":0924
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Refresh"
         Top             =   90
         Width           =   375
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   7905
         Left            =   120
         TabIndex        =   0
         Top             =   645
         Width           =   19050
         _Version        =   524288
         _ExtentX        =   33602
         _ExtentY        =   13944
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
         SpreadDesigner  =   "AccountLedger.frx":0A6E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   15240
         TabIndex        =   5
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
         Picture         =   "AccountLedger.frx":15D1
         Picture         =   "AccountLedger.frx":15ED
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   16920
         TabIndex        =   6
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
         Picture         =   "AccountLedger.frx":1609
         Picture         =   "AccountLedger.frx":1625
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   17310
         TabIndex        =   2
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "AccountLedger.frx":1641
         Caption         =   "AccountLedger.frx":1759
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountLedger.frx":17C5
         Keys            =   "AccountLedger.frx":17E3
         Spin            =   "AccountLedger.frx":1841
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
         Left            =   15840
         TabIndex        =   1
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "AccountLedger.frx":1869
         Caption         =   "AccountLedger.frx":1981
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountLedger.frx":19ED
         Keys            =   "AccountLedger.frx":1A0B
         Spin            =   "AccountLedger.frx":1A69
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
         Index           =   0
         Left            =   5550
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
         Picture         =   "AccountLedger.frx":1A91
         Picture         =   "AccountLedger.frx":1AAD
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   330
         Left            =   1200
         TabIndex        =   10
         Top             =   8620
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountLedger.frx":1AC9
         Caption         =   "AccountLedger.frx":1AE9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountLedger.frx":1B4D
         Keys            =   "AccountLedger.frx":1B6B
         Spin            =   "AccountLedger.frx":1BB5
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   8620
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
         Picture         =   "AccountLedger.frx":1BDD
         Picture         =   "AccountLedger.frx":1BF9
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   18090
         TabIndex        =   12
         Top             =   8625
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
         Picture         =   "AccountLedger.frx":1C15
         Picture         =   "AccountLedger.frx":1C31
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   16920
         TabIndex        =   13
         Top             =   8625
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
         Picture         =   "AccountLedger.frx":1C4D
         Picture         =   "AccountLedger.frx":1C69
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   2520
         TabIndex        =   15
         Top             =   8625
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
         Picture         =   "AccountLedger.frx":1C85
         Picture         =   "AccountLedger.frx":1CA1
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   17295
         TabIndex        =   18
         Top             =   8505
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
         Picture         =   "AccountLedger.frx":1CBD
         Picture         =   "AccountLedger.frx":1CD9
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
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
         Picture         =   "AccountLedger.frx":1CF5
         Picture         =   "AccountLedger.frx":1D11
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   7920
         TabIndex        =   23
         Top             =   120
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
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
         Picture         =   "AccountLedger.frx":1D2D
         Picture         =   "AccountLedger.frx":1D49
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   6480
         TabIndex        =   17
         Top             =   8625
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
         Left            =   5310
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
Public sDate As String, eDate As String, AccountGroupList As String, AccountList As String, VchType As String, Header1 As String, SCode As Variant, LR As Integer, R As Long
Dim rstAccountLedger As New ADODB.Recordset, rstAccountOpening As New ADODB.Recordset, Reset As Long
Dim Debit As Double, Credit As Double, Bal As Variant, DebitTotal As Double, CreditTotal As Double, BalTotal As Double, Code As Variant, TotalFlag As Boolean, HideFlag As Boolean, ExitFlag As Boolean
Dim Opening As Double
Private Sub Combo1_Change()
    If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Command1_Click()
With fpSpread1
    fpSpread1.DeleteRows .DataRowCnt, 1
    cmdRefresh_Click
End With
End Sub
Private Sub Form_Load()
Reset = 0:
    On Error GoTo ErrorHandler
    CenterForm Me
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
    If VchType = 23 Then Me.Caption = " Accounts Ledger"
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
Dim TypeCode As Variant
SCode = "": TypeCode = ""
    With fpSpread1
        .GetText 11, .ActiveRow, TypeCode: TypeCode = Right(TypeCode, 2)
    End With
    'Sale
   If (Shift = 0 And KeyCode = vbKeyReturn Or Shift = 0 And KeyCode = vbKeyF8 Or Shift = 0 And KeyCode = vbKeyF12) And (TypeCode = "SF" Or TypeCode = "TF" Or TypeCode = "PF" Or TypeCode = "OF") Then
                With fpSpread1
                    .GetText 3, .ActiveRow, SCode
                End With
        If SCode = "" Then Exit Sub
        On Error Resume Next
        frmSalesVoucher.VchType = TypeCode
        If Err.Number <> 364 Then frmSalesVoucher.Show
        frmSalesVoucher.Text1 = SCode
        
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
                    .GetText 3, .ActiveRow, SCode
                End With
        If SCode = "" Then Exit Sub
        On Error Resume Next
        TypeCode = IIf(TypeCode = "SU", 1, IIf(TypeCode = "SC", 2, IIf(TypeCode = "SJ", 3, IIf(TypeCode = "PU", 4, IIf(TypeCode = "PC", 5, IIf(TypeCode = "PJ", 6, ""))))))
        frmJobworkBill.VchType = TypeCode
        If Err.Number <> 364 Then frmJobworkBill.Show
        frmJobworkBill.Text1 = SCode
        
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
        If SCode = "" Then Exit Sub
        On Error Resume Next
        frmDebitCreditVoucher.VchType = TypeCode
        If Err.Number <> 364 Then frmDebitCreditVoucher.Show
        frmDebitCreditVoucher.Text1 = SCode
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
            Call cmdRefresh_Click: ExitFlag = True
        ElseIf ExitFlag = True Then
            Call cmdCancel_Click: ExitFlag = False
        End If
            KeyCode = 0
     End With
   ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
        MsgBox "Please Provide Search Input", vbInformation
        Text1.SetFocus
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
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbEnter And Shift = vbCtrlMask Then Call cmdFilter_Click
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstAccountLedger)
    Call CloseRecordset(rstAccountOpening)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Sub cmdRefresh_Click()
    On Error GoTo ErrHandler
    Dim SQL As String, OpSQL As String, i As Long, R As Long, C As Long
    Debit = 0: Credit = 0: Bal = 0
    CreditTotal = 0: DebitTotal = 0:
     If VchType >= 0 And VchType <= 29 Then 'Account Ledger
  OpSQL = "SELECT(SELECT ISNULL(Sum(C.Credit-C.Debit),0) As Opening FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date < '" & GetDate(MhDateInput1.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account IN (" & AccountList & ")) + " & _
                             "(SELECT ISNULL(Sum(IIF(LEFT(Type,2)='01',(T.Amount),IIF(LEFT(Type,2)='02',(0-T.Amount),IIF(LEFT(Type,2)='03',(T.Amount),IIF(LEFT(Type,2)='04',(0-T.Amount),0))))),0) As Opening FROM JobworkBVParent T LEFT JOIN AccountMaster A On T.Party=A.Code Where T.Date < '" & GetDate(MhDateInput1.Text) & "' And T.Party IN (" & AccountList & ")) +  " & _
                             "(Select ISNULL(Sum(Opening),0) From AccountMaster Where Code IN (" & AccountList & ")) As Opening "
            Screen.MousePointer = vbHourglass
            If rstAccountOpening.State = adStateOpen Then rstAccountOpening.Close
            rstAccountOpening.Open OpSQL, cnDatabase, adOpenKeyset, adLockReadOnly
            If rstAccountOpening.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
            
  SQL = "SELECT Date As Date,'Pymt' As VchTYPE,T.Name As VchBillNo,A.Name As Account,(Select IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 Where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,(Select IIF(C2.TOA='C',Credit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") AND Right(BOM,2)='PI' " & _
            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
            "SELECT Date As Date,'Rcpt' As VchTYPE,T.Name As VchBillNo,A.Name As Account,(SELECT IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 Where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,(Select IIF(C2.TOA='C',Credit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") AND Right(BOM,2)='PR' " & _
            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
            "SELECT Date As Date,'Jrnl' As VchTYPE,T.Name As VchBillNo,A.Name As Account,(SELECT IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,IIF(TOA='D',C.Debit,0) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") And Right(BOM,2)='JE' " & _
            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
            "SELECT Date As Date,'Cntr' As VchTYPE,T.Name As VchBillNo,A.Name As Account,(SELECT IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,(Select IIF(C2.TOA='C',Credit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") AND Right(BOM,2)='CE' " & _
            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
            "SELECT Date As Date,'DrNt' As VchTYPE,T.Name As VchBillNo,A.Name As Account,(SELECT IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,(Select IIF(C2.TOA='C',Credit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") AND Right(BOM,2)='DN' " & _
            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
            "SELECT Date As Date,'CrNt' As VchTYPE,T.Name As VchBillNo,A.Name As Account,(SELECT IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,(Select IIF(C2.TOA='C',Credit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") AND Right(BOM,2)='CN' " & _
            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _
            "SELECT Date As Date,'Pur' As VchTYPE,T.Name As VchBillNo,'Purchase' As Account, '0' As Debit,T.Amount As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T LEFT JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where Left(Type,2)='01' And Right(Type,2)='PF' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") Union " & _
            "SELECT Date As Date,'Pur' As VchTYPE,T.Name As VchBillNo,'Job Work Purchase ' As Account, '0' As Debit,T.Amount As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T LEFT JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where Left(Type,2)='01' And Right(Type,2)<>'PF' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") Union " & _
            "SELECT Date As Date,'PurRtr' As VchTYPE,T.Name As VchBillNo,'Purchase Return' As Account, T.Amount As Debit,'0' As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T LEFT JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where Left(Type,2)='02' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") Union " & _
            "SELECT Date As Date,'SaleRtr' As VchTYPE,T.Name As VchBillNo,'Sale Return' As Account, '0' As Debit,T.Amount As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T LEFT JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where Left(Type,2)='03' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") Union " & _
            "SELECT Date As Date,'Sale' As VchType,LTRIM(T.Name) As VchBillNo,S.Name As Account,(T.Amount+T.Rebate) As Debit,'0' As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) As Type,T.Code As Code,A.Name As AccountName FROM (((JobworkBVParent T LEFT JOIN JobworkBVChild C ON C.Code=T.Code) LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code) LEFT JOIN AccountMaster A ON T.Party=A.Code) LEFT JOIN AccountMaster S ON T.SalesType=S.Code WHERE LEFT(Type,2)='04' AND RIGHT(Type,2)='SF' AND T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND T.Party IN (" & AccountList & ") Union " & _
            "SELECT Date As Date,'Sale' As VchTYPE,T.Name As VchBillNo,'Rebate' As Account, '0' As Debit,T.Rebate As Credit,'Rebate' As ShortNarration,'' As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T LEFT JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where T.Rebate>0 And Left(Type,2)='04' And Right(Type,2)='SF' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") Union " & _
            "SELECT Date As Date,'Sale' As VchTYPE,T.Name As VchBillNo,'Job Work Sale ' As Account, T.Amount As Debit,'0' As Credit,'' As ShortNarration,T.Remarks As LongNarration,LTRIM(T.Type) AS Type,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM JobworkBVParent T INNER JOIN JobworkBVChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On T.Party=A.Code Where Left(Type,2)='04' And Right(Type,2)<>'SF' And T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' And T.Party IN (" & AccountList & ") ORDER BY Date ASC"
    End If
'            "SELECT Date As Date,'Jrnl' As VchTYPE,T.Name As VchBillNo,A.Name As Account,(SELECT IIF(C2.TOA='D',Debit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Debit,(Select IIF(C2.TOA='C',Credit,0) From DebitCreditChild C2 where C2.Account IN (" & AccountList & ") AND T.Code=C2.Code) As Credit,C.ShortNarration As ShortNarration,T.LongNarration As LongNarration,LTRIM(T.Type) AS Type ,T.Code As Code,(Select Name From AccountMaster Master Where Code=" & AccountList & ") As AccountName FROM DebitCreditParent T LEFT JOIN DebitCreditChild C On C.Code=T.Code LEFT JOIN VchSeriesMaster V ON T.VchSeries=V.Code LEFT JOIN AccountMaster A On C.Account=A.Code WHERE T.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND C.Code IN (Select Code From DebitCreditChild C1 Where C1.Account IN (" & AccountList & ")) AND C.Account NOT IN (" & AccountList & ") And Right(BOM,2)='JE' " & _
'            "AND TOA= IIF((Select TOA From DebitCreditChild C1 Where T.Code=C1.Code AND C1.Account IN (" & AccountList & "))='D','C','D') Union " & _

            
            Screen.MousePointer = vbHourglass
            If rstAccountLedger.State = adStateOpen Then rstAccountLedger.Close
            rstAccountLedger.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
            If rstAccountLedger.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
            
            Dim n As Integer
    Mh3dLabel9.Caption = "Account : " + rstAccountLedger.Fields("AccountName").value
    rstAccountOpening.MoveFirst
    Opening = Format(Val(rstAccountOpening.Fields("Opening").value), "##,##,##,##0.00")
    Mh3dLabel10.Caption = "Opening Bal. = Rs. " & Format(Opening, "##,##,##,##0.00") & IIf(Opening <= 0, " Dr.", " Cr.")
    
    Bal = Opening
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
            If VchType <= 24 And VchType >= 21 Then fpSpread1.ColHeaderRows = 1: fpSpread1.Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12:
            Next
        rstAccountLedger.MoveFirst
        Do While Not rstAccountLedger.EOF
            i = i + 1
                .SetText 1, i, rstAccountLedger.Fields("Date").value
                .SetText 2, i, rstAccountLedger.Fields("VchType").value
                .SetText 3, i, rstAccountLedger.Fields("VchBillNo").value
                .SetText 4, i, rstAccountLedger.Fields("Account").value
                        Debit = Val(rstAccountLedger.Fields("Debit").value)
                .SetText 5, i, Val(rstAccountLedger.Fields("Debit").value)
                        Credit = Val(rstAccountLedger.Fields("Credit").value)
                .SetText 6, i, Val(rstAccountLedger.Fields("Credit").value)
                        Bal = Bal + Credit - Debit
                .SetText 7, i, Bal
                .SetText 8, i, IIf(Bal <= 0, "Dr.", "Cr.")
                .SetText 9, i, rstAccountLedger.Fields("ShortNarration").value
                .SetText 10, i, rstAccountLedger.Fields("LongNarration").value
                .SetText 11, i, rstAccountLedger.Fields("Type").value
                .SetText 12, i, rstAccountLedger.Fields("Code").value

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
    End With
    TDBNumber2 = i
    fpSpread1.MaxRows = IIf(i < 27, 27, i + 1)
    Call cmdFilter_Click
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
'Private Sub PendingCheck_Click()
'If TDBNumber1.Value <= 0 And PendingCheck.Value Then ZeroStock.Value = 0
'    Call cmdRefresh_Click
'End Sub
'Private Sub ZeroStock_Click()
'If ZeroStock.Value Then NegativeStock.Value = 0
'If TDBNumber1.Value <= 0 And ZeroStock.Value Then PendingCheck.Value = 0
'    Call cmdRefresh_Click
'End Sub
'Private Sub NegativeStock_Click()
'   If NegativeStock.Value And PendingCheck.Value = 0 Then ZeroStock.Value = 0
'   If NegativeStock.Value And PendingCheck.Value And TDBNumber1 > 0 Then ZeroStock.Value = 0: TDBNumber1.Value = 0
'Call cmdRefresh_Click
'End Sub
'Private Sub TDBNumber1_Change()
'Dim n As Long, i As Long
'If TDBNumber1 > 0 Then ZeroStock.Value = 1
'    With fpSpread1
'    If .DataRowCnt = 0 Then Exit Sub
'            n = .DataRowCnt:
'        For i = 1 To .DataRowCnt 'Unhide All
'            .Row = i: .RowHidden = False
'    Next
'End With
'Call cmdRefresh_Click
'End Sub
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
Dim fileName As String

If Dir(App.Path & "\Export", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Export"

 fileName = App.Path & "\Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"

' Save to xls file type

' Load an Excel-formatted file

fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, False

'MsgBox
    MsgBox "Import Processing....", vbInformation, "Easy Publish...Import !!! "

fpSpread1.ImportExcelBook fileName, ""        '& "\EasyPublish.xls", ""

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
Dim x As Boolean, fileName As String, SheetName As String, LogFileName As String
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
     fileName = App.Path & "\Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"
    SheetName = "Sheet1" '"(" & Me.Caption & ")"
    LogFileName = "Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".txt"
    x = fpSpread1.ExportToExcelEx(fileName, SheetName, LogFileName, ExcelSaveFlagNoFormulas)
    ' Display result to user based on T/F value of x
    If x = True Then
    
    MsgBox "Export complete.", vbInformation, "Easy Publish...Export !!! "
        
        Dim oExcel As Object
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open (fileName)
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
fpSpread1.PrintAbortMsg = "Printing - Click Cancel to .Quit"
fpSpread1.PrintJobName = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") '& ".pdf"
'fpSpread1.PrintHeader = "/cPrint Header/rPage # ./p/n2nd Line"
fpSpread1.PrintFooter = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") & " /rPage # ./p ": .FontSize = 16 '& ".pdf" ' "/cPrint Footer/rPage # ./p/n2nd Line"
fpSpread1.PrintBorder = True
fpSpread1.PrintColHeaders = True
fpSpread1.PrintColor = True
fpSpread1.PrintGrid = True
fpSpread1.PrintMarginTop = 750 '1440
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
    'Delete Header Row
    If VchType >= 0 And VchType <= 30 Then fpSpread1.DeleteRows 1, 1
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
        fpSpread1.MaxCols = 13
        
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
fpSpread1.RowHeight(n) = 12.75
    End With
 fpSpread1.DeleteRows n, 1
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
            n = .DataRowCnt: DebitVal = 0: CreditVal = 0: Bal = Opening
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
                .GetText 3, i, cVal
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
    fpSpread1.MaxRows = IIf(TDBNumber2.value < 27, i + (27 - TDBNumber2.value), i + 1)
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
fpSpread1.PrintAbortMsg = "Printing - Click Cancel to .Quit"
fpSpread1.PrintJobName = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") '& ".pdf"
fpSpread1.PrintFooter = "        Export Data Company : " & " " & " _(" & CompCode & "_" & PrintHeader & ")" & "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]" & " /rPage # ./p " & " Print Date : ( " & Format(Date, "dd-MMM-yyyy") & " )         ": .FontSize = 16 '& ".pdf" ' "/cPrint Footer/rPage # ./p/n2nd Line"
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
