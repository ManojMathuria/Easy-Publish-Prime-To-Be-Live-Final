VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Object = "{96548BD2-D0BF-46B1-B519-8F2268D49306}#1.0#0"; "vsvport8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStockLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Stock Status"
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
   ScaleHeight     =   9255
   ScaleWidth      =   19485
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   18600
      Picture         =   "StockLedger.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Refresh"
      Top             =   210
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   9030
      Left            =   120
      TabIndex        =   5
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
      Picture         =   "StockLedger.frx":014A
      Begin VB.CheckBox Check2 
         Caption         =   "Show Subtotal"
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
         Left            =   3840
         TabIndex        =   37
         Top             =   158
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComctlLib.Slider Zoom 
         Height          =   75
         Left            =   17400
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Zoom"
         Top             =   8400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   132
         _Version        =   393216
         Min             =   -5
         Max             =   5
         TickStyle       =   2
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   32
         Top             =   645
         Width           =   19095
         _cx             =   33681
         _cy             =   13996
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         BackColorSel    =   -2147483635
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
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   38
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         ExplorerBar     =   5
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
         Begin VSViewPort8LibCtl.VSViewPort VSViewPort1 
            Height          =   7875
            Left            =   570
            TabIndex        =   33
            Top             =   -30
            Width           =   19050
            _cx             =   33602
            _cy             =   13891
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            MousePointer    =   0
            BackColor       =   -2147483633
            AutoScroll      =   -1  'True
            VirtualWidth    =   1000
            VirtualHeight   =   1000
            LargeChangeHorz =   300
            LargeChangeVert =   300
            SmallChangeHorz =   30
            SmallChangeVert =   30
            Track           =   0   'False
            MouseScroll     =   0   'False
            ProportionalBars=   -1  'True
            FocusTrack      =   0   'False
            FocusMarginLeft =   0
            FocusMarginTop  =   0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
         Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
            Height          =   7935
            Left            =   600
            TabIndex        =   34
            Top             =   0
            Width           =   19095
            _cx             =   33681
            _cy             =   13996
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            MousePointer    =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoRTF         =   -1  'True
            Preview         =   -1  'True
            DefaultDevice   =   0   'False
            PhysicalPage    =   -1  'True
            AbortWindow     =   -1  'True
            AbortWindowPos  =   0
            AbortCaption    =   "Printing..."
            AbortTextButton =   "Cancel"
            AbortTextDevice =   "on the %s on %s"
            AbortTextPage   =   "Now printing Page %d of"
            FileName        =   ""
            MarginLeft      =   1440
            MarginTop       =   1440
            MarginRight     =   1440
            MarginBottom    =   1440
            MarginHeader    =   0
            MarginFooter    =   0
            IndentLeft      =   0
            IndentRight     =   0
            IndentFirst     =   0
            IndentTab       =   720
            SpaceBefore     =   0
            SpaceAfter      =   0
            LineSpacing     =   100
            Columns         =   1
            ColumnSpacing   =   180
            ShowGuides      =   2
            LargeChangeHorz =   300
            LargeChangeVert =   300
            SmallChangeHorz =   30
            SmallChangeVert =   30
            Track           =   0   'False
            ProportionalBars=   -1  'True
            Zoom            =   42.2974176313446
            ZoomMode        =   3
            ZoomMax         =   400
            ZoomMin         =   10
            ZoomStep        =   25
            EmptyColor      =   -2147483636
            TextColor       =   0
            HdrColor        =   0
            BrushColor      =   0
            BrushStyle      =   0
            PenColor        =   0
            PenStyle        =   0
            PenWidth        =   0
            PageBorder      =   0
            Header          =   ""
            Footer          =   ""
            TableSep        =   "|;"
            TableBorder     =   7
            TablePen        =   0
            TablePenLR      =   0
            TablePenTB      =   0
            NavBar          =   3
            NavBarColor     =   -2147483633
            ExportFormat    =   0
            URL             =   ""
            Navigation      =   3
            NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
            AutoLinkNavigate=   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
         Begin VSReport8LibCtl.VSReport VSReport1 
            Left            =   9240
            Top             =   1560
            _rv             =   800
            ReportName      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OnOpen          =   ""
            OnClose         =   ""
            OnNoData        =   ""
            OnPage          =   ""
            OnError         =   ""
            MaxPages        =   0
            DoEvents        =   -1  'True
            BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
               Width           =   0
               MarginLeft      =   1440
               MarginTop       =   1440
               MarginRight     =   1440
               MarginBottom    =   1440
               Columns         =   1
               ColumnLayout    =   0
               Orientation     =   0
               PageHeader      =   0
               PageFooter      =   0
               PictureAlign    =   7
               PictureShow     =   1
               PaperSize       =   0
            EndProperty
            BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
               ConnectionString=   ""
               RecordSource    =   ""
               Filter          =   ""
               MaxRecords      =   0
            EndProperty
            GroupCount      =   0
            SectionCount    =   5
            BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
               Name            =   "Detail"
               Visible         =   0   'False
               Height          =   0
               CanGrow         =   -1  'True
               CanShrink       =   0   'False
               KeepTogether    =   -1  'True
               ForcePageBreak  =   0
               BackColor       =   16777215
               Repeat          =   0   'False
               OnFormat        =   ""
               OnPrint         =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
               Name            =   "Header"
               Visible         =   0   'False
               Height          =   0
               CanGrow         =   -1  'True
               CanShrink       =   0   'False
               KeepTogether    =   -1  'True
               ForcePageBreak  =   0
               BackColor       =   16777215
               Repeat          =   0   'False
               OnFormat        =   ""
               OnPrint         =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
               Name            =   "Footer"
               Visible         =   0   'False
               Height          =   0
               CanGrow         =   -1  'True
               CanShrink       =   0   'False
               KeepTogether    =   -1  'True
               ForcePageBreak  =   0
               BackColor       =   16777215
               Repeat          =   0   'False
               OnFormat        =   ""
               OnPrint         =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
               Name            =   "Page Header"
               Visible         =   0   'False
               Height          =   0
               CanGrow         =   -1  'True
               CanShrink       =   0   'False
               KeepTogether    =   -1  'True
               ForcePageBreak  =   0
               BackColor       =   16777215
               Repeat          =   0   'False
               OnFormat        =   ""
               OnPrint         =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
               Name            =   "Page Footer"
               Visible         =   0   'False
               Height          =   0
               CanGrow         =   -1  'True
               CanShrink       =   0   'False
               KeepTogether    =   -1  'True
               ForcePageBreak  =   0
               BackColor       =   16777215
               Repeat          =   0   'False
               OnFormat        =   ""
               OnPrint         =   ""
               Object.Tag             =   ""
            EndProperty
            FieldCount      =   0
         End
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   18880
         Picture         =   "StockLedger.frx":0166
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Zoom"
         Top             =   8625
         Width           =   375
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   4
         Left            =   1920
         Top             =   1200
      End
      Begin VB.CommandButton Preview 
         Caption         =   "&Print Preview"
         Height          =   330
         Left            =   15360
         TabIndex        =   30
         Top             =   8640
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Height          =   320
         Left            =   5880
         Picture         =   "StockLedger.frx":04D8
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Search"
         Top             =   8640
         Width           =   375
      End
      Begin VB.CommandButton cmdFilter 
         Height          =   320
         Left            =   5400
         Picture         =   "StockLedger.frx":081A
         Style           =   1  'Graphical
         TabIndex        =   21
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
         TabIndex        =   19
         ToolTipText     =   "Find And Search"
         Top             =   8620
         Width           =   2070
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
         Height          =   330
         Left            =   13680
         TabIndex        =   14
         Top             =   105
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "StockLedger.frx":0B5C
         Caption         =   "StockLedger.frx":0B7C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StockLedger.frx":0BE0
         Keys            =   "StockLedger.frx":0BFE
         Spin            =   "StockLedger.frx":0C48
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
      Begin VB.CheckBox Check1 
         Caption         =   "Show Amount "
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
         Left            =   7440
         TabIndex        =   13
         Top             =   158
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox NegativeStock 
         Caption         =   "Show Negative Stock Items"
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
         Left            =   4920
         TabIndex        =   12
         Top             =   158
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   375
         Left            =   18840
         Picture         =   "StockLedger.frx":0C70
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel"
         Top             =   90
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   375
         Left            =   18480
         Picture         =   "StockLedger.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Refresh"
         Top             =   90
         Width           =   375
      End
      Begin VB.CheckBox ZeroStock 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Purchase Return Greater Than Equal >>>>"
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
         Left            =   9360
         TabIndex        =   9
         Top             =   158
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.CheckBox PendingCheck 
         Caption         =   "Show Pending"
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
         Left            =   3360
         TabIndex        =   4
         Top             =   158
         Visible         =   0   'False
         Width           =   1455
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   7905
         Left            =   120
         TabIndex        =   0
         Top             =   660
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
         MaxCols         =   35
         MaxRows         =   2000
         SelectBlockOptions=   4
         SpreadDesigner  =   "StockLedger.frx":0EBC
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
         Picture         =   "StockLedger.frx":2061
         Picture         =   "StockLedger.frx":207D
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
         Picture         =   "StockLedger.frx":2099
         Picture         =   "StockLedger.frx":20B5
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   2190
         TabIndex        =   2
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "StockLedger.frx":20D1
         Caption         =   "StockLedger.frx":21E9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StockLedger.frx":2255
         Keys            =   "StockLedger.frx":2273
         Spin            =   "StockLedger.frx":22D1
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
         TabIndex        =   1
         Top             =   120
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "StockLedger.frx":22F9
         Caption         =   "StockLedger.frx":2411
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StockLedger.frx":247D
         Keys            =   "StockLedger.frx":249B
         Spin            =   "StockLedger.frx":24F9
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
         Left            =   14760
         TabIndex        =   8
         Top             =   120
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
         Picture         =   "StockLedger.frx":2521
         Picture         =   "StockLedger.frx":253D
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   330
         Left            =   1200
         TabIndex        =   15
         Top             =   8620
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "StockLedger.frx":2559
         Caption         =   "StockLedger.frx":2579
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StockLedger.frx":25DD
         Keys            =   "StockLedger.frx":25FB
         Spin            =   "StockLedger.frx":2645
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
         TabIndex        =   16
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
         Picture         =   "StockLedger.frx":266D
         Picture         =   "StockLedger.frx":2689
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   17880
         TabIndex        =   17
         Top             =   8625
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
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
         Picture         =   "StockLedger.frx":26A5
         Picture         =   "StockLedger.frx":26C1
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   16680
         TabIndex        =   18
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
         Picture         =   "StockLedger.frx":26DD
         Picture         =   "StockLedger.frx":26F9
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   2520
         TabIndex        =   20
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
         Picture         =   "StockLedger.frx":2715
         Picture         =   "StockLedger.frx":2731
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   15380
         TabIndex        =   23
         Top             =   8620
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
         Picture         =   "StockLedger.frx":274D
         Picture         =   "StockLedger.frx":2769
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel 
         Height          =   330
         Left            =   8520
         TabIndex        =   27
         Top             =   8625
         Visible         =   0   'False
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
         Picture         =   "StockLedger.frx":2785
         Picture         =   "StockLedger.frx":27A1
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   13680
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
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
         Picture         =   "StockLedger.frx":27BD
         Picture         =   "StockLedger.frx":27D9
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   3360
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
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
         Picture         =   "StockLedger.frx":27F5
         Picture         =   "StockLedger.frx":2811
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   12720
         TabIndex        =   24
         Top             =   8625
         Visible         =   0   'False
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
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
         Caption         =   "Create Stock Journal Voucher"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "StockLedger.frx":282D
         Picture         =   "StockLedger.frx":2849
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   9360
         TabIndex        =   31
         Top             =   105
         Visible         =   0   'False
         Width           =   5415
         _Version        =   65536
         _ExtentX        =   9551
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
         Picture         =   "StockLedger.frx":2865
         Picture         =   "StockLedger.frx":2881
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   6360
         TabIndex        =   22
         Top             =   8625
         Width           =   2085
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3678;582"
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
         Left            =   16110
         TabIndex        =   3
         Top             =   105
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
   Begin VSReport8LibCtl.VSReport VSReport4 
      Left            =   9600
      Top             =   4440
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
   Begin VSReport8LibCtl.VSReport VSReport3 
      Left            =   9600
      Top             =   4440
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
   Begin VSReport8LibCtl.VSReport VSReport2 
      Left            =   9600
      Top             =   4440
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
End
Attribute VB_Name = "FrmStockLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nSort As Boolean, VSFlexFlag As Boolean, FontFlag As Boolean
Public dSortBy As Boolean
Public sDate As String, eDate As String, ItemList As String, oItemName As String, ItemGroupList As String, MatCList As String, AccountList As String, VchType As String
Dim rstStockLedger As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstItemOpening As New ADODB.Recordset, rstCompanyMaster As New ADODB.Recordset
Dim Reset As Long, sysStock As Variant, phyStock As Variant, LR As Integer, R As Long, TotalFlag As Boolean, HideFlag As Boolean, ExitFlag As Boolean
Dim Opening As Double, Debit As Double, Credit As Double, Bal As Variant
Dim oMcCode As Variant, oPartyCode As Variant
Public sMcCode As Variant, SCode As Variant, oSCode As Variant, vTypeCode As Variant, vtCode As Variant, vtType As Variant, vtNo As Variant, vDate As Variant
Dim oVchType As String, Header1 As String, VchCode As String, PartyH As String, ItemH As String, OrderH As String, OrderF As Double, INWardF As Double, OUTWardF As Double, AmountF As Double, SNo As Long, aSNO As Long, pSNO As Long
Dim OrderGTF As Double, INWardGTF As Double, OUTWardGTF As Double, AmountGTF As Double
Dim OrderPGTF As Double, INWardPGTF As Double, OUTWardPGTF As Double, AmountPGTF As Double, ClearFlag As Boolean, unClearFlag As Boolean
Private Sub Form_Load()
If VchType <> 34 And VchType <> 45 And VchType <> 30 Then VchCode = ""
If VchType = 35 Or VchType = 36 Or VchType = 39 Or VchType = 40 Or VchType = 41 Then VchCode = "S"
If VchType = 37 Or VchType = 38 Or VchType = 42 Or VchType = 43 Or VchType = 44 Then VchCode = "P"
If VchType = 34 Or VchType = 45 Then VchCode = VchCode
If VchType = 49 Then VSFlexFlag = False
Reset = 0:
If SCode <> "" Then SCode = SCode Else SCode = ""
If VchType = 31 Or VchType = 49 Then If SCode = "" Then SCode = ItemList
    On Error GoTo ErrorHandler
    CenterForm Me
    Me.Top = (MdiMainMenu.ScaleHeight - Me.Height) \ 2 + 1000
    BusySystemIndicator True
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    If VchType <= 2 Or VchType = 33 Then
        Combo1.Clear
        Combo1.AddItem "Item Ascending", 0
        Combo1.AddItem "Item Descending", 1
        Combo1.AddItem "Item Group Ascending", 2
        Combo1.AddItem "Item Group Descending", 3
        Combo1.AddItem "MRP Ascending", 4
        Combo1.AddItem "MRP Descending", 5
        Combo1.AddItem "All", 6
        Combo1.ListIndex = 0
    ElseIf (VchType >= 3 And VchType <= 33) Or (VchType >= 53 And VchType <= 69) Then
        Combo1.Clear
        Combo1.AddItem "Item Ascending", 0
        Combo1.AddItem "Item Descending", 1
        Combo1.AddItem "Item Group Ascending", 2
        Combo1.AddItem "Item Group Descending", 3
        Combo1.AddItem "All", 4
        Combo1.ListIndex = 0
    ElseIf VchType = 45 Or VchType >= 34 And VchType <= 38 Then
        Combo1.Clear
        Combo1.AddItem " Date", 0
        Combo1.AddItem " Vch/Bill No", 1
        Combo1.AddItem " Particulars", 2
        Combo1.ListIndex = 1
    ElseIf VchType >= 39 And VchType <= 44 Then
        Combo1.Clear
        Combo1.AddItem " Date", 0
        Combo1.AddItem " Vch/Bill No", 1
        Combo1.AddItem " Account-Wise", 2
        Combo1.AddItem " Item-Wise", 3
            If VchType = 39 Or VchType = 42 Then
                Combo1.ListIndex = 1
            ElseIf VchType = 40 Or VchType = 43 Then
                Combo1.ListIndex = 2
            ElseIf VchType = 41 Or VchType = 44 Then
                Combo1.ListIndex = 3
            End If
    ElseIf VchType >= 46 And VchType <= 48 Then
        Combo1.Clear
        Combo1.AddItem " Sales Direct", 0
        Combo1.AddItem " Sales Against Challan", 1
        Combo1.ListIndex = 1
    ElseIf VchType = 103 Then
        Combo1.Visible = False
    ElseIf VchType >= 101 Then
        Combo1.Clear
        Combo1.AddItem "Item Ascending", 0
        Combo1.AddItem "Item Descending", 1
        Combo1.AddItem "WIP Ascending", 2
        Combo1.AddItem "WIP Descending", 3
        Combo1.AddItem "RM Ascending", 4
        Combo1.AddItem "RM Descending", 5
        Combo1.ListIndex = 0
    End If
    If VchType >= 34 And VchType <= 45 Then
        Combo2.Clear
        Combo2.AddItem " Date", 0
        Combo2.AddItem " Vch/Bill No", 1
        Combo2.AddItem " Particulars", 2
        Combo2.AddItem " Buyers Name", 3
        Combo2.ListIndex = 3
    ElseIf VchType = 46 Then
        Combo2.Clear
        Combo2.AddItem " Date", 0
        Combo2.AddItem " Vch/Bill No", 1
        Combo2.AddItem " Particulars", 2
        Combo2.AddItem " Unit Rate", 3
        Combo2.AddItem " Buyer Name", 4
        Combo2.ListIndex = 0
    ElseIf VchType = 47 Then
        Combo2.Clear
        Combo2.AddItem " Buyer Name", 0
        Combo2.ListIndex = 0
    ElseIf VchType >= 101 Then
        Combo2.Clear
        Combo2.AddItem "Item", 0
        Combo2.AddItem "WIP", 1
        Combo2.AddItem "RM", 2
        Combo2.ListIndex = 0
    Else
        Combo2.Clear
        Combo2.AddItem " Item Name", 0
        Combo2.AddItem " Item Group", 1
        Combo2.ListIndex = 0
    End If
    Reset = 1
    Combo1.Visible = True: Mh3dLabel1(0).Visible = True: MhDateInput1.ReadOnly = True: MhDateInput2.ReadOnly = True: Combo2.Visible = True: Command1.Visible = True: cmdFilter.Visible = True: Mh3dLabel7.Visible = True: Text1.Visible = True: Mh3dLabel.Visible = True: Command2.Visible = True: Mh3dLabel10.Visible = True
    If VchType <> 0 Then Command1 = False
    If VchType > 0 Then Mh3dLabel.Visible = True:
    If VchType = 0 Then Me.Caption = "Physical Stock Audit Ledger Item-Wise": Mh3dLabel8.Visible = True: Mh3dLabel9.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    '
    If VchType = 1 Then PendingCheck.Visible = True:  ZeroStock.Visible = True:  NegativeStock.Visible = True: Me.Caption = "Inventory Movement Ledger Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False:: Mh3dLabel11.Visible = False:
    If VchType = 2 Then PendingCheck.Visible = True:  ZeroStock.Visible = True:  NegativeStock.Visible = True: Me.Caption = "Stock Status Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:: Mh3dLabel10.Visible = False:: Mh3dLabel11.Visible = False:
    '
    If VchType = 3 Then Check1.Visible = True: ZeroStock.Visible = True:   ZeroStock.Caption = "Show Sales Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = "Sales Analysis Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 4 Then Check1.Visible = True: ZeroStock.Visible = True:   ZeroStock.Caption = "Show Sales Return Greater Than Equal >>>>": ZeroStock.Width = 4000: ZeroStock.Left = 9540: Check1.Left = 3500: Me.Caption = "Sales Return Analysis Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 5 Then Check1.Visible = True: ZeroStock.Visible = True:   ZeroStock.Caption = "Show Qty. Greater Than Equal >>>>": ZeroStock.Width = 3200: ZeroStock.Left = 10300: Check1.Left = 3500: Me.Caption = "Sales And Sales Return Analysis Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 6 Then Check1.Visible = True: ZeroStock.Visible = True:    ZeroStock.Caption = "Show Sales Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = "Net Sales Analysis Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    '
    If VchType = 7 Then Check1.Visible = True:  ZeroStock.Visible = True: ZeroStock.Caption = "Show Sales Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = "Sales Analysis One Party Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 8 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Left = 10000: ZeroStock.Width = 3800: ZeroStock.Caption = "Show Sales Return Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = "Sales Return Analysis One Party Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 9 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Qty. Greater Than Equal >>>>": ZeroStock.Width = 3200: ZeroStock.Left = 10300: Check1.Left = 3500: Me.Caption = "Sales And Sales Return Analysis One Party Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 10 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Sales Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = "Net Sales Analysis One Party Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    '
    If VchType = 21 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Sales Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = "Sales Analysis Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 22 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Left = 10000: ZeroStock.Width = 3800: ZeroStock.Caption = "Show Sales Return Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = "Sales Return Analysis Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 23 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Qty. Greater Than Equal >>>>": ZeroStock.Width = 3200: ZeroStock.Left = 10300: Check1.Left = 3500: Me.Caption = "Sales And Sales Return Analysis Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 24 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Sales Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = "Net Sales Analysis Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
     '
    If VchType = 25 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Sales Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = "Sales Analysis One-Item Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 26 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Left = 10000: ZeroStock.Width = 3800: ZeroStock.Caption = "Show Sales Return Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = "Sales Return Analysis One-Item Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 27 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Qty. Greater Than Equal >>>>": ZeroStock.Width = 3200: ZeroStock.Left = 10300: Check1.Left = 3500: Me.Caption = "Sales And Sales Return Analysis One-Item Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 28 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Sales Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = "Net Sales Analysis One-Item Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    '
    If (VchType >= 29 And VchType <= 30) Or VchType = 69 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False: ZeroStock.Visible = False: TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = False: Combo1.Visible = False: PendingCheck.Visible = False: ZeroStock.Caption = IIf(vtCode = 18, "Show Sales", "Sales Purchase") + "Greater Than Equal >>>>": Check1.Left = 3500: Me.Caption = IIf(vtCode = 18, "Pending Purchase Order", "Pending Sales Order")
    If (VchType = 31 Or Right(VchType, 2) = 48) Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: Mh3dLabel10.Visible = True: Mh3dLabel11.Visible = True: ZeroStock.Visible = False: TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = False: Combo1.Visible = False: PendingCheck.Visible = False: ZeroStock.Caption = "Show Sales Greater Than Equal >>>>": Check1.Visible = False: Me.Caption = "Item Ledger"
    If Right(VchType, 2) = 48 And Left(VchType, 2) = "04" Then Me.Caption = "Sales Voucher-Wise"
    If Right(VchType, 2) = 48 And Left(VchType, 2) = "01" Then Me.Caption = "Purchase Voucher-Wise"
    If VchType = 32 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: Mh3dLabel10.Visible = True: Mh3dLabel11.Visible = True: ZeroStock.Visible = False: TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = False: Combo1.Visible = False: PendingCheck.Visible = False: ZeroStock.Caption = "Show Sales Greater Than Equal >>>>": Check1.Visible = False: Me.Caption = "Item Ledger Material Centre-Wise"
    If VchType = 33 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = True:  ZeroStock.Visible = False: NegativeStock.Visible = True: Me.Caption = "Short-Item Analysis Item-Wise": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False:
    If VchType = 34 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:  ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Orders Status Voucher-Wise ": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 35 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:   ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Purchase Orders-Party-Wise-Detailed": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 36 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:   ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Purchase Orders-Party-Wise-Summarized": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 37 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:  ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Purchase Orders-Party-Wise-Detailed": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 38 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:  ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Sales Orders-Party-Wise-Summarized ": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 39 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:  ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Purchase Orders Order-Wise ": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 40 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:  ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Purchase Orders Party-Wise ": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 41 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:  ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Purchase Orders Item-Wise ": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 42 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:  ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Sales Orders Order-Wise ": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 43 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:  ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Sales Orders Party-Wise ": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 44 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:  ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Sales Orders Item-Wise ": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 45 Then MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: PendingCheck.Visible = False:  ZeroStock.Visible = False: NegativeStock.Visible = False: Me.Caption = "Orders Status Voucher-Wise ": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = True: Combo1.Visible = True: Mh3dLabel10.Visible = False: Mh3dLabel11.Visible = False
    If VchType = 46 Then Check1.Visible = True: Check1.Caption = "Select All": Me.Caption = "Pending Sales Order-Wise ": Mh3dLabel1(0).Caption = "   Pending ": Mh3dLabel12.Visible = True: Mh3dLabel12.Caption = "F9->Clear Order Quantity F10-> Retrieve Order Quantity": Mh3dLabel.Caption = "Ctrl+F->Search  F8->Delete  Escap->Un-Hide  F12->Duplicate  F5->Refresh"
    If VchType = 47 Then Check1.Visible = True: Check1.Caption = "Select All": Me.Caption = "Pending Sales Party-Wise ": Mh3dLabel1(0).Caption = "   Pending "
    '
    If VchType = 53 Then Check1.Visible = True: ZeroStock.Visible = True:   ZeroStock.Caption = "Show Purchase Greater Than Equal >>>>": ZeroStock.Width = 3700: ZeroStock.Left = 9800: Check1.Left = 3500: Me.Caption = "Purchase Analysis Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 54 Then Check1.Visible = True: ZeroStock.Visible = True:   ZeroStock.Caption = "Show Purchase Return Greater Than Equal >>>>": ZeroStock.Width = 4215: ZeroStock.Left = 9300: Check1.Left = 3500: Me.Caption = "Purchase Return Analysis Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 55 Then Check1.Visible = True: ZeroStock.Visible = True:   ZeroStock.Caption = "Show Qty. Greater Than Equal >>>>": ZeroStock.Width = 3200: ZeroStock.Left = 10300: Check1.Left = 3500: Me.Caption = "Purchase And Purchase Return Analysis Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 56 Then Check1.Visible = True: ZeroStock.Visible = True:    ZeroStock.Caption = "Show Purchase Greater Than Equal >>>>": ZeroStock.Width = 3700: ZeroStock.Left = 9800: Check1.Left = 3500: Me.Caption = "Net Purchase Analysis Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    '
    If VchType = 57 Then Check1.Visible = True:  ZeroStock.Visible = True: ZeroStock.Caption = "Show Purchase Greater Than Equal >>>>": ZeroStock.Width = 3700: ZeroStock.Left = 9800: Check1.Left = 3500: Me.Caption = "Purchase Analysis One Party Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 58 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Purchase Return Greater Than Equal >>>>": ZeroStock.Width = 4215: ZeroStock.Left = 9300: Check1.Left = 3500: Me.Caption = "Purchase Return Analysis One Party Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 59 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Qty. Greater Than Equal >>>>": ZeroStock.Width = 3200: ZeroStock.Left = 10300: Check1.Left = 3500: Me.Caption = "Purchase And Purchase Return Analysis One Party Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 60 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Purchase Greater Than Equal >>>>": ZeroStock.Width = 3700: ZeroStock.Left = 9800: Check1.Left = 3500: Me.Caption = "Net Purchase Analysis One Party Item-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    '
    If VchType = 61 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Purchase Greater Than Equal >>>>": ZeroStock.Width = 3700: ZeroStock.Left = 9800: Check1.Left = 3500: Me.Caption = "Purchase Analysis Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 62 Then Check1.Visible = True: ZeroStock.Visible = True:  ZeroStock.Caption = "Show Purchase Return Greater Than Equal >>>>": ZeroStock.Width = 4215: ZeroStock.Left = 9300: Check1.Left = 3500: Me.Caption = "Purchase Return Analysis Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 63 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Qty. Greater Than Equal >>>>": ZeroStock.Width = 3200: ZeroStock.Left = 10300: Check1.Left = 3500: Me.Caption = "Purchase And Purchase Return Analysis Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 64 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Purchase Greater Than Equal >>>>": ZeroStock.Width = 3700: ZeroStock.Left = 9800: Check1.Left = 3500: Me.Caption = "Net Purchase Analysis Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
     '
    If VchType = 65 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Purchase Greater Than Equal >>>>": ZeroStock.Width = 3700: ZeroStock.Left = 9800: Check1.Left = 3500: Me.Caption = "Purchase Analysis One-Item Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 66 Then Check1.Visible = True: ZeroStock.Visible = True:  ZeroStock.Caption = "Show Purchase Return Greater Than Equal >>>>": ZeroStock.Width = 4215: ZeroStock.Left = 9300: Check1.Left = 3500: Me.Caption = "Purchase Return Analysis One-Item Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 67 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Qty. Greater Than Equal >>>>": ZeroStock.Width = 3200: ZeroStock.Left = 10300: Check1.Left = 3500: Me.Caption = "Purchase And Purchase Return Analysis One-Item Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    If VchType = 68 Then Check1.Visible = True: ZeroStock.Visible = True: ZeroStock.Caption = "Show Purchase Greater Than Equal >>>>": ZeroStock.Width = 3700: ZeroStock.Left = 9800: Check1.Left = 3500: Me.Caption = "Net Purchase Analysis One-Item Party-Wise": TDBNumber1.Visible = True: Mh3dLabel1(0).Visible = True: Combo1.Visible = True:
    
    If VchType = 103 Then Check2.Visible = True: Check1.Visible = False: ZeroStock.Visible = False: Check1.Left = 3500: Me.Caption = "WIP Pending Item-Wise Ledger": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = False: Combo1.Visible = False
    If VchType = 104 Then Check2.Visible = True: Check1.Visible = False: ZeroStock.Visible = False: Check1.Left = 3500: Me.Caption = "RM Pending Item-Wise": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = False: Combo1.Visible = False
    If VchType = 105 Then Check2.Visible = True: Check1.Visible = False: ZeroStock.Visible = False: Check1.Left = 3500: Me.Caption = "RM & WIP Item-Wise [PSS] ": TDBNumber1.Visible = False: Mh3dLabel1(0).Visible = False: Combo1.Visible = False
    If VchType = 49 Then Combo1.Visible = False: Mh3dLabel1(0).Visible = False: MhDateInput1.ReadOnly = True: MhDateInput2.ReadOnly = True: Combo2.Visible = False: Command1.Visible = False: cmdFilter.Visible = False: Mh3dLabel7.Visible = False: Text1.Visible = False: Mh3dLabel.Visible = False: Command2.Visible = False: Mh3dLabel10.Visible = False: Me.Caption = " Inventory - Montly Summary":
    MhDateInput1.Value = Format(sDate, "dd-MM-yyyy")
    MhDateInput2.Value = Format(eDate, "dd-MM-yyyy")
    cmdRefresh_Click
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub cmdRefresh_Click()
    On Error GoTo ErrHandler
    Dim SQL As String, i As Long, R As Long, C As Long, Stock As Long, EffectiveStock As Long, StockTotal As Long, POTotal As Long, SOTotal As Long, EStockTotal As Long, AmountTotal As Double, PurchaseTotal As Long, PurchaseReturnTotal As Long, PurchaseChallanTotal As Long, PurchaseReturnChallanTotal As Long, SalesTotal As Long, SalesReturnTotal As Long, SalesChallanTotal As Long, SalesReturnChallanTotal As Long, StockJournalINTotal As Long, StockJournalOUTTotal As Long, NetPurchaseTotal As Long, NetSalesTotal As Long, PurchaseAmountTotal As Double, SalesAmountTotal As Double, PurchaseReturnAmountTotal As Double, SalesReturnAmountTotal As Double, NetPurchaseAmountTotal As Double, NetSalesAmountTotal As Double, ICode As Variant
    Dim OpSQL As String, dPrint As Long
    Debit = 0: Credit = 0: Bal = 0
    If VchType = 31 Or VchType = 32 Or VchType = 49 Then ' Item Ledger
    oMcCode = IIf(sMcCode <> "", "P.MaterialCentre", "P.Party")
    OpSQL = "Select ISNULL(Sum(INWard),0) As INWard,ISNULL(Sum(OutWard),0) As OutWard, ISNULL(Sum(INWard),0)-ISNULL(Sum(OutWard),0)+ISNULL((SELECT OPBAL From BookChild I Where MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ") And Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")),0) As Opening,(SELECT ISNULL(Sum(OPBAL),0) From BookChild I Where MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ") And Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")) As OPBAL From (" & _
                "SELECT ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') UNION ALL " & _
                "SELECT '0' As INWard,ISNULL(ABS(Quantity),0) As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT '0' As INWard,ISNULL(ABS(Quantity),0) As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")  AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') UNION ALL " & _
                "SELECT ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='05' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT '0' As INWard,ISNULL(ABS(Quantity),0) As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='06' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT '0' As INWard,ISNULL(ABS(Quantity),0) As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='08' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='07' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity<0  UNION ALL " & _
                "SELECT IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity>0 UNION ALL " & _
                "SELECT IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity<0 UNION ALL " & _
                "SELECT IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard  FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date < '" & GetDate(MhDateInput1.Text) & "'  AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity>0 ) As Tbl "
                    Screen.MousePointer = vbHourglass
                    If rstItemOpening.State = adStateOpen Then rstItemOpening.Close
                        rstItemOpening.Open OpSQL, cnDatabase, adOpenKeyset, adLockReadOnly
                    If rstItemOpening.RecordCount = 0 Then Screen.MousePointer = vbNormal: 'Exit Sub
                        If rstItemOpening.RecordCount = 0 Then
                        OpSQL = "Select 0 As INWard,0 As OutWard, (SELECT ISNULL(Sum(OPBAL),0) From BookChild I Where I.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ") AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")) As Opening,Name As Item,Code As ItemCode,(SELECT ISNULL(Sum(OPBAL),0) From BookChild I Where I.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ") AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")) As OPBAL From BookMaster Where Code IN (" & IIf(SCode <> "", SCode, ItemList) & ")"
                        If rstItemOpening.State = adStateOpen Then rstItemOpening.Close
                        rstItemOpening.Open OpSQL, cnDatabase, adOpenKeyset, adLockReadOnly
                        If rstItemOpening.RecordCount = 0 Then Screen.MousePointer = vbNormal: 'Exit Sub
                    End If
    End If
    If VchType <= 2 Or VchType = 33 Then 'Stock Ledger
    SQL = "SELECT * FROM (" & _
                "SELECT " & IIf(VchType <= 10 And VchType >= 7, "(select name from AccountMaster where code='" & AccountList & "')", "''") & " as OneParty,I.Name As Item,I.Price  As MRP,G.Name As ItemGroup," & _
                "ISNULL((SELECT SUM(0-R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='17' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code),0)+ISNULL((SELECT SUM(EstQty01-(DeliveredQuantityC+DeliveredQuantityB)) FROM BookPOParent WHERE LEFT(Type,1)<>'O' AND RIGHT(Type,1)='P' AND Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND MaterialCentre IN (" & AccountList & ") AND Book=I.CODE),0) As PendingPO," & _
                "ISNULL((SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code),0)+ISNULL((SELECT SUM(EstQty01-(DeliveredQuantityC+DeliveredQuantityB)) FROM BookPOParent WHERE LEFT(Type,1)<>'O' AND RIGHT(Type,1)='S' AND Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND MaterialCentre IN (" & AccountList & ") AND Book=I.CODE),0) As PendingSO," & _
                "ISNULL((SELECT SUM(OPBAL) FROM BookChild C WHERE C.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code),0)  As OPBAL," & _
                "ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='05' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code),0)  As PurchaseChallan," & _
                "ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='06' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code),0)  As PurchaseReturnChallan," & _
                "ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='07' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code),0)  As SalesChallan," & _
                "ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='08' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code),0)  As SalesReturnChallan," & _
                "ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code And SubString(P.Type,3,2)='10'),0)  As Purchase," & _
                "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code And SubString(P.Type,3,2)='10'),0)  As PurchaseAmount," & _
                "ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code),0)  As PurchaseReturn," & _
                "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code),0)  As PurchaseReturnAmount," & _
                "ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code And SubString(P.Type,3,2)='10'),0)  As Sales," & _
                "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code And SubString(P.Type,3,2)='10'),0)  As SalesAmount," & _
                "ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code),0)  As SalesReturn," & _
                "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code),0)  As SalesReturnAmount," & _
                "ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code AND C.Quantity>0),0) As StockTransferIN," & _
                "ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND Party IN (" & AccountList & ") AND C.Item=I.Code AND C.Quantity<0),0) As StockTransferOUT," & _
                "ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND Party IN (" & AccountList & ") AND C.Item=I.Code AND C.Quantity>0),0) As StockJournalIN," & _
                "ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND MaterialCentre IN (" & AccountList & ") AND C.Item=I.Code AND C.Quantity<0),0) As StockJournalOUT," & _
                "ISNULL((SELECT ABS(SUM(Quantity)) From JobworkBVRef Where RefCode IN (Select RefCode From JobworkBVRef C inner join JobworkBVParent P on P.code=C.vchcode WHERE LEFT(C.VchType,2)='23' AND C.VchDate BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item = I.Code)),0) As PQ , " & _
                "ISNULL((SELECT ABS(SUM(Quantity)) From JobworkBVRef Where RefCode IN (Select RefCode From JobworkBVRef C inner join JobworkBVParent P on P.code=C.vchcode WHERE LEFT(C.VchType,2)='24' AND C.VchDate BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & AccountList & ") AND C.Item = I.Code)),0) As SQ " & _
                " ,I.Code As code,I.HSNCode FROM BookMaster I INNER JOIN GeneralMaster G ON I.[Group]=G.Code WHERE I.Code IN (" & ItemList & ")" & _
                ") As Tbl ORDER BY " & Choose(Combo1.ListIndex + 1, "Item ASC,MRP,ItemGroup", "Item DESC,MRP,ItemGroup", "ItemGroup ASC,Item,MRP", "ItemGroup DESC,Item,MRP", "MRP ASC,Item,ItemGroup", "MRP DESC,Item,ItemGroup", "Item ASC,MRP,ItemGroup") & ""
    ElseIf (VchType >= 3 And VchType <= 10) Or (VchType >= 53 And VchType <= 60) Then 'Item-Wise'( Sale And Purchase Ledger)
    SQL = "SELECT * FROM (" & _
                "SELECT " & IIf(((VchType >= 7 And VchType <= 10) Or (VchType >= 57 And VchType <= 60)), "(select name from AccountMaster where code=" & IIf(sMcCode <> "", sMcCode, AccountList) & ")", "''") & " as OneParty,I.Name As Item,I.Price  As MRP,G.Name As ItemGroup,I.Code As code,'' As HSNCode," & _
                SQL & _
                "ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=I.Code),0)  As Purchase," & _
                "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=I.Code),0)  As PurchaseAmount," & _
                "ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=I.Code),0)  As PurchaseReturn," & _
                "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=I.Code),0)  As PurchaseReturnAmount," & _
                "ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=I.Code),0)  As Sales," & _
                "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=I.Code),0)  As SalesAmount," & _
                "ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=I.Code),0)  As SalesReturn," & _
                "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=I.Code),0)  As SalesReturnAmount " & _
                "FROM BookMaster I INNER JOIN GeneralMaster G ON I.[Group]=G.Code WHERE I.Code IN (" & ItemList & ")" & _
                ") As Tbl ORDER BY " & Choose(Combo1.ListIndex + 1, "Item ASC,MRP,ItemGroup", "Item DESC,MRP,ItemGroup", "ItemGroup ASC,Item,MRP", "ItemGroup DESC,Item,MRP", "MRP ASC,Item,ItemGroup", "MRP DESC,Item,ItemGroup", "Item ASC,MRP,ItemGroup") & ""
    ElseIf (VchType >= 21 And VchType <= 28) Or (VchType >= 61 And VchType <= 68) Then 'Party-Wise'( Sale And Purchase Ledger)
    SQL = "SELECT * FROM (" & _
              "SELECT " & IIf(((VchType >= 25 And VchType <= 28) Or (VchType >= 65 And VchType <= 68)), "(select name from BookMaster where code=" & IIf(SCode <> "", SCode, ItemList) & ")", "''") & " as OneItem,I.Name As Item,'' As MRP,G.Name As ItemGroup,I.Code As Code,'' As HSNCode," & _
              "ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND P.Party=I.Code),0)  As Purchase," & _
              "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND P.Party=I.Code),0)  As PurchaseAmount," & _
              "ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND P.Party=I.Code),0)  As PurchaseReturn," & _
              "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND P.Party=I.Code),0)  As PurchaseReturnAmount," & _
              "ISNULL((SELECT SUM(ABS(C.Quantity)) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND P.Party=I.Code),0)  As Sales," & _
              "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND P.Party=I.Code),0)  As SalesAmount," & _
              "ISNULL((SELECT SUM(C.Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND P.Party=I.Code),0)  As SalesReturn," & _
              "ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND P.Party=I.Code),0)  As SalesReturnAmount " & _
              "FROM AccountMaster I INNER JOIN GeneralMaster G ON I.[Group]=G.Code WHERE I.Code IN (" & AccountList & ")" & _
              ") As Tbl ORDER BY " & Choose(Combo1.ListIndex + 1, "Item ASC,ItemGroup", "Item DESC,ItemGroup", "ItemGroup ASC,Item", "ItemGroup DESC,Item", "Item ASC,ItemGroup") & ""
    ElseIf VchType >= 29 And VchType <= 30 Then 'Pending Order
    SQL = "Select Code As VchCode,Date As Date,RIGHT(Type,1)+'O/'+LTRIM(Name)+'/JW/'+IIF(FORMAT(Date,'MM')<4,Convert(Nvarchar,(Convert(int,FORMAT(Date,'yy'))-1)) +'-'+Convert(Nvarchar,FORMAT(Date,'yy')),Convert(Nvarchar,FORMAT(Date,'yy')) +'-'+ Convert(Nvarchar,Convert(int,FORMAT(Date,'yy'))+1)) As VchBillNo,(Select Name From BookMAster Where Code=Book) As Item,(Select Name From AccountMaster Where Code=IIF(BookPrinter IS NOT NULL, BookPrinter,IIF(TitlePrinter IS NOT NULL,TitlePrinter,IIF(Laminator IS NOT NULL,Laminator,IIF(Binder IS NOT NULL,Binder,''))))) As Details," & _
             "EstQty01 As Ordered,0 As Dispatched,(EstQty01-(DeliveredQuantityB+DeliveredQuantityC)) As Pending,'No.' As Unit,UnitRate As Rate,(UnitRate*(EstQty01-(DeliveredQuantityB+DeliveredQuantityC))) As Amount,Type AS VchType,IIF(Right(Type,2)='FP','PO','SO') AS Type From BookPOParent Where Type='" & IIf(vTypeCode = "18", "FP", "FS") & "' AND Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND Book IN (" & IIf(SCode <> "", SCode, ItemList) & ")   AND (DeliveredQuantityC) = 0 AND (DeliveredQuantityB) = 0 " & _
             "Union " & _
             "Select P1.Code As VchCode,P1.Date As Date,RIGHT(P1.Type,1)+'O/'+LTRIM(P1.Name)+'/JW/'+IIF(FORMAT(P1.Date,'MM')<4,Convert(Nvarchar,(Convert(int,FORMAT(P1.Date,'yy'))-1)) +'-'+Convert(Nvarchar,FORMAT(P1.Date,'yy')),Convert(Nvarchar,FORMAT(P1.Date,'yy')) +'-'+ Convert(Nvarchar,Convert(int,FORMAT(P1.Date,'yy'))+1)) As VchBillNo,(Select Name From BookMaster A Where A.Code=C.Item) As Item,(Select Name From AccountMaster A Where A.Code=P.Party) As Details," & _
             "AVG(P1.EstQty01) As Ordered,Sum((IIF(LEFT(BOM,18) IN ('0310XXXXXXXXXXXXFI','0710XXXXXXXXXXXXFI','0110XXXXXXXXXXXXFI','0510XXXXXXXXXXXXFI','0000'),C.Quantity,0))+ABS(IIF(LEFT(BOM,18) IN ('0410XXXXXXXXXXXXFI','0810XXXXXXXXXXXXFI','0210XXXXXXXXXXXXFI','0610XXXXXXXXXXXXFI','0000'),C.Quantity,0))) As Dispatched,AVG(P1.EstQty01)-Sum((IIF(LEFT(BOM,18) IN ('0310XXXXXXXXXXXXFI','0710XXXXXXXXXXXXFI','0110XXXXXXXXXXXXFI','0510XXXXXXXXXXXXFI','0000'),C.Quantity,0))+ABS(IIF(LEFT(BOM,18) IN ('0410XXXXXXXXXXXXFI','0810XXXXXXXXXXXXFI','0210XXXXXXXXXXXXFI','0610XXXXXXXXXXXXFI','0000'),C.Quantity,0))) As Pending,'No.' As Unit,AVG(C.Rate) As Rate," & _
             "AVG(C.Rate)*(AVG(P1.EstQty01)-Sum((IIF(LEFT(BOM,18) IN ('0310XXXXXXXXXXXXFI','0710XXXXXXXXXXXXFI','0110XXXXXXXXXXXXFI','0510XXXXXXXXXXXXFI','0000'),C.Quantity,0))+ABS(IIF(LEFT(BOM,18) IN ('0410XXXXXXXXXXXXFI','0810XXXXXXXXXXXXFI','0210XXXXXXXXXXXXFI','0610XXXXXXXXXXXXFI','0000'),C.Quantity,0)))) As Amount,P1.Type As VchType,IIF(Right(P1.Type,2)='FP','PO','SO') AS Type " & _
             "From JobworkBVChild C INNER JOIN JobworkBVParent P ON P.Code=C.Code INNER JOIN BookPOParent P1 ON P1.CODE=C.Ref Where (RefCode =''  Or  RefCode='XXXXXX') AND LEFT(IIF(C.BOM IS NOT Null,C.BOM,'0000FI'),18) IN " & IIf(Left(VchCode, 1) = "S", "('0110XXXXXXXXXXXXFI','0210XXXXXXXXXXXXFI','0510XXXXXXXXXXXXFI','0610XXXXXXXXXXXXFI','0000FI')", "('0310XXXXXXXXXXXXFI','0410XXXXXXXXXXXXFI','0710XXXXXXXXXXXXFI','0810XXXXXXXXXXXXFI','0000FI')") & _
             " AND C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") " & _
             "Group By P1.Code,P1.Date,RIGHT(P1.Type,1)+'O/'+LTRIM(P1.Name)+'/JW/'+IIF(FORMAT(P1.Date,'MM')<4,Convert(Nvarchar,(Convert(int,FORMAT(P1.Date,'yy'))-1)) +'-'+Convert(Nvarchar,FORMAT(P1.Date,'yy')),Convert(Nvarchar,FORMAT(P1.Date,'yy')) +'-'+ Convert(Nvarchar,Convert(int,FORMAT(P1.Date,'yy'))+1)),Item,Party,P1.Type " & _
             "Union " & _
             "SELECT DISTINCT C.RefCode As VchCode,P.Date AS Date,LTRIM(P.Name) AS VchBillNo,(Select PrintName From BookMaster Where Code=C.Item ) AS Item,(Select PrintName From AccountMaster Where Code=P.Party) AS Details,IIF(LEFT(BOM,4)+Right(BOM,2)IN ('1701FI','1801FI'),ABS(C.Quantity),0) AS Ordered,ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0) As Dispatched,(IIF(LEFT(BOM,4)+Right(BOM,2)IN ('1701FI','1801FI'),ABS(C.[Quantity]),0)-ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0)) As Bal,'Units' As Unit,C.Rate As Rate,(C.Rate*(IIF(LEFT(BOM,4)+Right(BOM,2)IN ('1701FI','1801FI'),ABS(C.[Quantity]),0)-ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0))) As Amount,P.TYPE As VchType,Right(Type,2) AS Type  " & _
             "FROM JobworkBVParent P INNER JOIN JobworkBVChild C ON P.Code=C.Code Left Join JobworkBVRef R ON R.VchCode=C.Code " & _
             "WHERE LEFT((C.BOM),6) IN ('" & IIf(VchCode = "P", "1801FI", "1701FI") & "') AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "' AND LEFT(P.Code,1)<>'*' AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & _
             IIf(FrmItemSelectionList.Option3.Value, "ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0)<ABS(C.Quantity)", IIf(FrmItemSelectionList.Option1.Value, "ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0)>=ABS(C.Quantity)", IIf(FrmItemSelectionList.Option2.Value, "IIf(Right(P.Type,1)='P',ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0),ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0))>=0", 1))) & "  " & _
             "AND C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") " & _
             "ORDER BY VchBillNo "
    ElseIf VchType = 31 Then ' Item Ledger Date-wise
    oMcCode = IIf(sMcCode <> "", "P.MaterialCentre", "P.Party")
    SQL = "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code=  " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') UNION ALL " & _
                "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,'0' As INWard,ISNULL(ABS(Quantity),0) As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code= " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,'0' As INWard,ISNULL(ABS(Quantity),0) As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code= " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")  AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') UNION ALL " & _
                "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code= " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code= " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='05' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,'0' As INWard,ISNULL(ABS(Quantity),0) As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code= " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='06' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,'0' As INWard,ISNULL(ABS(Quantity),0) As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code= " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='08' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code= " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='07' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code= " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity<0  UNION ALL " & _
                "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code= " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity>0 UNION ALL " & _
                "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code= " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity<0 UNION ALL " & _
                "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard,'Units' As Unit,Rate,(Rate*ISNULL(ABS(Quantity),0)) As Amount,BOM,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) AS Type,(Select Name From BookMaster Where Code=C.Item) As Item,(Select Name From AccountMaster Where Code= P.Party) As Party,(Select Name From AccountMaster Where Code= " & oMcCode & " ) As MaterialCentre,P.Type VchType FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity>0" & _
                "Order By P.Date ASC "
    ElseIf VchType = 49 Then ' Item Ledger Date-wise
    oMcCode = IIf(sMcCode <> "", "P.MaterialCentre", "P.Party")
    SQL = "Select mCode,MonthYear,Sum(INWard) As INWard,Sum(OutWard) As OutWard,Item As Item,Format(DATEADD(month, mCode-4, '" & FinancialYearFrom & "'),'dd-MMM-yyyy') AS FromDate,Format(DATEADD(Day,-01,DATEADD(month, mCode-3, '" & FinancialYearFrom & "')),'dd-MMM-yyyy') As ToDate From(" & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') UNION ALL " & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, '0' As INWard,ISNULL(ABS(Quantity),0) As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, '0' As INWard,ISNULL(ABS(Quantity),0) As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")  AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') UNION ALL " & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='05' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, '0' As INWard,ISNULL(ABS(Quantity),0) As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='06' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, '0' As INWard,ISNULL(ABS(Quantity),0) As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='08' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, ISNULL(ABS(Quantity),0) As INWard,'0' As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='07' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity<0  UNION ALL " & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity>0 UNION ALL " & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity<0 UNION ALL " & _
                "SELECT IIF(FORMAT(P.Date, 'MM')>3,FORMAT(P.Date, 'MM'),FORMAT(P.Date, 'MM')+12) AS mCode,FORMAT(P.Date, 'MMMM') AS MonthYear, IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As INWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As OutWard,(Select Name From BookMaster Where Code=C.Item) As Item FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & oMcCode & " IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity>0" & _
                ") As TBL Group By mCode,MonthYear,Item Order By mCode Asc "
      ElseIf VchType = 32 Then 'One Item Ledger Material Centre-wise
      SQL = "Select ISNULL(Sum(oINWard),0) As oINWard,ISNULL(Sum(oOutWard),0) As oOutWard, ISNULL(Sum(oINWard),0)-ISNULL(Sum(oOutWard),0)+(SELECT ISNULL(Sum(OPBAL),0) From BookChild I Where I.MaterialCentre = TBL.MaterialCentre AND Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")) As Opening,ISNULL(Sum(cINWard),0) As cINWard,ISNULL(Sum(cOutWard),0) As cOutWard,ISNULL(Sum(cINWard),0)-ISNULL(Sum(cOutWard),0)+ISNULL(Sum(oINWard),0)-ISNULL(Sum(oOutWard),0)+(SELECT ISNULL(Sum(OPBAL),0) From BookChild I Where I.MaterialCentre = TBL.MaterialCentre AND Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")) As Closing,(Select Name From AccountMaster Where Code= MaterialCentre) As MaterialCentreName,MaterialCentre,(SELECT ISNULL(Sum(OPBAL),0) From BookChild I Where I.MaterialCentre = TBL.MaterialCentre AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")) As OPBAL From ( " & _
                "Select ISNULL(ABS(Quantity),0) As oINWard,'0' As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') UNION ALL " & _
                "Select  '0' As oINWard,ISNULL(ABS(Quantity),0) As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  '0' As oINWard,ISNULL(ABS(Quantity),0) As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') UNION ALL " & _
                "Select  ISNULL(ABS(Quantity),0) As oINWard,'0' As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  ISNULL(ABS(Quantity),0) As oINWard,'0' As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='05' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  '0' As oINWard,ISNULL(ABS(Quantity),0) As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='06' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  '0' As oINWard,ISNULL(ABS(Quantity),0) As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='08' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  ISNULL(ABS(Quantity),0) As oINWard,'0' As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='07' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As oINWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity<0 UNION ALL " & _
                "Select  IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As oINWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity<0 UNION ALL " & _
                "Select  IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As oINWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity>0 UNION ALL " & _
                "Select  IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As oINWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As oOutWard,'0' As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date < '" & GetDate(MhDateInput1.Text) & "' And P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity>0 UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,ISNULL(ABS(Quantity),0) As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='01' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,'0' As cINWard,ISNULL(ABS(Quantity),0) As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='02' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,'0' As cINWard,ISNULL(ABS(Quantity),0) As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='04' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,ISNULL(ABS(Quantity),0) As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='03' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,ISNULL(ABS(Quantity),0) As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='05' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,'0' As cINWard,ISNULL(ABS(Quantity),0) As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='06' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,'0' As cINWard,ISNULL(ABS(Quantity),0) As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='08' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,ISNULL(ABS(Quantity),0) As cINWard,'0' As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='07' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As cINWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As cOutWard,IIF(Quantity<0,PArty,MaterialCentre) As MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity>0 UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As cINWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity>0 UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As cINWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As cOutWard,IIF(Quantity<0,PArty,MaterialCentre) As MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='19' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity<0 UNION ALL " & _
                "Select  '0' As oINWard,'0' As oOutWard,IIF((Quantity)<0,'0',ISNULL(ABS(Quantity),0)) As cINWard,IIF((Quantity)<0,ISNULL(ABS(Quantity),0),'0') As cOutWard,MaterialCentre,C.Code FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE LEFT(P.Type,2)='20' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  Item IN (" & IIf(SCode <> "", SCode, ItemList) & ") AND C.Quantity<0) As Tbl Group By MaterialCentre "
      ElseIf VchType = 34 Then ' Order Status
        SQL = "Select P.Date As vtDate,P.Name As VchBillNo,IIF(LEFT(P.Type,4)='0110','Purchase',IIF(LEFT(P.Type,4)='0210','Purchase Return',IIF(LEFT(P.Type,4)='0510','Pur Challan IN',IIF(LEFT(P.Type,4)='0610','Pur Challan Out',IIF(LEFT(P.Type,4)='0310','Sales Return',IIF(LEFT(P.Type,4)='0410','Sales',IIF(LEFT(P.Type,4)='0710','Sales Challan IN',IIF(LEFT(P.Type,4)='0810','Sales Challan Out','Order Status')))))))) As TypeRef,(Select Name From BookMaster A Where A.Code=C.Item) As ItemName,(Select Name From AccountMaster A Where A.Code=P.Party) As AccountName,(Select Name From AccountMaster A Where A.Code=P.MaterialCentre) As MaterialCentre,P.Remarks,P.ChallanDate,P.ChallanNo, " & _
                  "P1.EstQty01 As Ordered,IIF(LEFT(BOM,18) IN ('0310XXXXXXXXXXXXFI','0710XXXXXXXXXXXXFI','0110XXXXXXXXXXXXFI','0510XXXXXXXXXXXXFI','0000'),C.Quantity,0) As INward,ABS(IIF(LEFT(BOM,18) IN ('0410XXXXXXXXXXXXFI','0810XXXXXXXXXXXXFI','0210XXXXXXXXXXXXFI','0610XXXXXXXXXXXXFI','0000'),C.Quantity,0)) As OutWard,(IIF(P1.Type = 'FP', -1, 1) * P1.EstQty01)-ABS(IIF(LEFT(BOM,18) IN ('0410XXXXXXXXXXXXFI','0810XXXXXXXXXXXXFI','0210XXXXXXXXXXXXFI','0610XXXXXXXXXXXXFI','0000'),C.Quantity,0))+IIF(LEFT(BOM,18) IN ('0310XXXXXXXXXXXXFI','0710XXXXXXXXXXXXFI','0110XXXXXXXXXXXXFI','0510XXXXXXXXXXXXFI','0000'),C.Quantity,0) As Pending,C.Rate,C.Amount,ISNULL(P.Name,'') AS vtNO,ISNULL(P.Type,'') AS vtType,ISNULL(P.Code,'') AS vtCode,P1.Code As pvtCode," & _
                  "RIGHT(P1.Type,1)+'O/'+LTRIM(P1.Name)+'/JW/'+IIF(FORMAT(P1.Date,'MM')<4,Convert(Nvarchar,(Convert(int,FORMAT(P1.Date,'yy'))-1)) +'-'+Convert(Nvarchar,FORMAT(P1.Date,'yy')),Convert(Nvarchar,FORMAT(P1.Date,'yy')) +'-'+ Convert(Nvarchar,Convert(int,FORMAT(P1.Date,'yy'))+1)) As VchBillNo, P1.Date As pvtDate,P1.Type As pvtType " & _
                  "From JobworkBVChild C INNER JOIN JobworkBVParent P ON P.Code=C.Code INNER JOIN BookPOParent P1 ON P1.CODE=C.Ref Where Ref=" & SCode & " AND (RefCode =''  Or  RefCode='XXXXXX') AND LEFT(IIF(C.BOM IS NOT Null,C.BOM,'0000FI'),18) IN " & IIf(Left(VchCode, 1) = "S", "('0110XXXXXXXXXXXXFI','0210XXXXXXXXXXXXFI','0510XXXXXXXXXXXXFI','0610XXXXXXXXXXXXFI','0000FI')", "('0310XXXXXXXXXXXXFI','0410XXXXXXXXXXXXFI','0710XXXXXXXXXXXXFI','0810XXXXXXXXXXXXFI','0000FI')") & "  ORDER BY P.Date "
      ElseIf VchType = 35 Or VchType = 37 Then 'Sale & Purchase Order Status Detailed
        SQL = "SELECT (Select Name From AccountMaster Where Code=Party) As AccountName,RIGHT((Select TYPE From BookPOParent Where Code=Ref),1)+'O/'+LTRIM((Select Name From BookPOParent Where Code=Ref))+'/JW/'+IIF(FORMAT((Select Date From BookPOParent Where Code=Ref),'MM')<4,Convert(Nvarchar,(Convert(int,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy'))-1)) +'-'+Convert(Nvarchar,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy')),Convert(Nvarchar,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy')) +'-'+ Convert(Nvarchar,Convert(int,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy'))+1)) As VchBillNo,(Select Date From BookPOParent Where Code=Ref) AS pvtDate,(Select Name From BookMaster Where Code=Item) As ItemName,ISNULL(ChallanNo,'') As ChallanNo,ISNULL(ChallanDate,'') As ChallanDate,(Select EstQty01 From BookPOParent Where Code=Ref) As Ordered,(Quantity) As INward,0 as Outward,Party As BCode,P.Name As GRNNo,Date As vtDate," & _
                  "(Select Name From AccountMaster Where Code=MaterialCentre)As MaterialCentre,(Select Name From AccountMaster Where Code=Party)As Party,Remarks,(Select LTrim(Name) From BookPOParent Where Code=Ref) As PO,(Select Name From BookMaster Where Code=Item) As Book,C.Quantity As Qty,ISNULL(C.Rate,(Select UnitRate From BookPOParent Where Code=Ref)) AS Rate,ISNULL(C.Amount,((Select EstQty01 From BookPOParent Where Code=Ref)*(Select UnitRate From BookPOParent Where Code=Ref))) As Amount,P.BOX,P.Freight,ISNULL(P.TYPE,'') AS TYPE,RIGHT(P.Type,2)+'-'+LTRIM(P.Name) As MRNNo,(Select Type From BookPOParent Where Code=Ref) As pvtType,Item As Code,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) As TypeRef,IIF(C.BOM IS NOT Null,C.BOM,'0000FI') As BOM,(Select Code From BookPOParent Where Code=Ref) As Code,(Select Code From BookPOParent Where Code=Ref) As pvtCode,ISNULL(P.Name,'') AS vtNO,ISNULL(P.Type,'') AS vtType,ISNULL(P.Code,'') AS vtCode " & _
                  "FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE P.Type IN ('0110PU','0110PC','0110PJ','0310PU','0310PC','0310PJ','0510FR','0710FR') AND Ref IN (SELECT Code FROM BookPOParent) AND LEFT((Select Type From BookPOParent Where Code=Ref),1)<>'O' AND RIGHT((Select Type From BookPOParent Where Code=Ref),1)<>'" & Left(VchCode, 1) & "' AND LEFT((Select Code From BookPOParent Where Code=Ref),1)<>'*' AND (Select Date From BookPOParent Where Code=Ref)>='" & GetDate(MhDateInput1.Text) & "' AND (Select Date From BookPOParent Where Code=Ref)<='" & GetDate(MhDateInput2.Text) & "' AND " & _
                   IIf(FrmItemSelectionList.Option3.Value, "(Select DeliveredQuantityC+DeliveredQuantityB From BookPOParent Where Code=Ref)<(Select EstQty01 From BookPOParent Where Code=Ref)", IIf(FrmItemSelectionList.Option1.Value, "(Select DeliveredQuantityC+DeliveredQuantityB From BookPOParent Where Code=Ref)>=(Select EstQty01 From BookPOParent Where Code=Ref)", "1=1")) & " AND P.MaterialCentre IN (" & MatCList & ") AND P.Party IN (" & AccountList & ") AND C.Item IN (" & ItemList & ") " & _
                  "UNION ALL " & _
                  "SELECT (Select Name From AccountMaster Where Code=Party) As AccountName,RIGHT((Select TYPE From BookPOParent Where Code=Ref),1)+'O/'+LTRIM((Select Name From BookPOParent Where Code=Ref))+'/JW/'+IIF(FORMAT((Select Date From BookPOParent Where Code=Ref),'MM')<4,Convert(Nvarchar,(Convert(int,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy'))-1)) +'-'+Convert(Nvarchar,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy')),Convert(Nvarchar,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy')) +'-'+ Convert(Nvarchar,Convert(int,FORMAT((Select Date From BookPOParent Where Code=Ref),'yy'))+1)) As VchBillNo,(Select Date From BookPOParent Where Code=Ref) AS pvtDate,(Select Name From BookMaster Where Code=Item) As ItemName,ISNULL(ChallanNo,'') As ChallanNo,ISNULL(ChallanDate,'') As ChallanDate,(Select EstQty01 From BookPOParent Where Code=Ref) As Ordered,0 As INward,(Quantity) as Outward,Party As BCode,P.Name As GRNNo,Date As vtDate," & _
                  "(Select Name From AccountMaster Where Code=MaterialCentre)As MaterialCentre,(Select Name From AccountMaster Where Code=Party)As Party,Remarks,(Select LTrim(Name) From BookPOParent Where Code=Ref) As PO,(Select Name From BookMaster Where Code=Item) As Book,C.Quantity As Qty,ISNULL(C.Rate,(Select UnitRate From BookPOParent Where Code=Ref)) AS Rate,ISNULL(C.Amount,((Select EstQty01 From BookPOParent Where Code=Ref)*(Select UnitRate From BookPOParent Where Code=Ref))) As Amount,P.BOX,P.Freight,ISNULL(P.TYPE,'') AS TYPE,RIGHT(P.Type,2)+'-'+LTRIM(P.Name) As MRNNo,(Select Type From BookPOParent Where Code=Ref) As pvtType,Item As Code,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) As TypeRef,IIF(C.BOM IS NOT Null,C.BOM,'0000FI') As BOM,(Select Code From BookPOParent Where Code=Ref) As Code,(Select Code From BookPOParent Where Code=Ref) As pvtCode,ISNULL(P.Name,'') AS vtNO,ISNULL(P.Type,'') AS vtType,ISNULL(P.Code,'') AS vtCode " & _
                  "FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE P.Type IN ('0210OU','0210OC','0210OJ','0410TU','0410TC','0410TJ','0610FI','0810FI') AND Ref IN (SELECT Code FROM BookPOParent) AND LEFT((Select Type From BookPOParent Where Code=Ref),1)<>'O' AND RIGHT((Select Type From BookPOParent Where Code=Ref),1)<>'" & Left(VchCode, 1) & "' AND LEFT((Select Code From BookPOParent Where Code=Ref),1)<>'*' AND (Select Date From BookPOParent Where Code=Ref)>='" & GetDate(MhDateInput1.Text) & "' AND (Select Date From BookPOParent Where Code=Ref)<='" & GetDate(MhDateInput2.Text) & "' AND " & _
                   IIf(FrmItemSelectionList.Option3.Value, "(Select DeliveredQuantityC+DeliveredQuantityB From BookPOParent Where Code=Ref)<(Select EstQty01 From BookPOParent Where Code=Ref)", IIf(FrmItemSelectionList.Option1.Value, "(Select DeliveredQuantityC+DeliveredQuantityB From BookPOParent Where Code=Ref)>=(Select EstQty01 From BookPOParent Where Code=Ref)", "1=1")) & " AND P.MaterialCentre IN (" & MatCList & ") AND P.Party IN (" & AccountList & ") AND C.Item IN (" & ItemList & ") " & _
                  "UNION ALL " & _
                 "SELECT (Select Name From AccountMaster Where Code=IIF(Binder IS NOT NULL AND Binder<>'' ,Binder,IIF(BookPrinter IS NOT NULL AND BookPrinter<>'',BookPrinter,IIF(TitlePrinter IS NOT NULL AND TitlePrinter<>'',TitlePrinter,Laminator)))) As AccountName,RIGHT(TYPE,1)+'O/'+LTRIM(Name)+'/JW/'+IIF(FORMAT(Date,'MM')<4,Convert(Nvarchar,(Convert(int,FORMAT(Date,'yy'))-1)) +'-'+Convert(Nvarchar,FORMAT(Date,'yy')),Convert(Nvarchar,FORMAT(Date,'yy')) +'-'+ Convert(Nvarchar,Convert(int,FORMAT(Date,'yy'))+1)) As VchBillNo,Date AS pvtDate," & _
                 "(Select Name From BookMaster Where Code=Book) As ItemName,'' As ChallanNo,'' As ChallanDate,EstQty01 As Ordered,0 As INward,0 as Outward,'' As BCode,'' As GRNNo,Date As vtDate,(Select Name From AccountMaster Where Code=MaterialCentre)As MaterialCentre,'' As Party,'' As Remarks,LTrim(Name) As PO,(Select Name From BookMaster Where Code=Book) As Book,0 As Qty,UnitRate AS Rate,(EstQty01*UnitRate) As Amount,0 As BOX,0 As Freight,'' As Type,'' As MRNNo,Type As pvtType,Book As Code,'' As TypeRef,'0000FI' As BOM,Code,Code As pvtCode,'' AS vtNO,'' AS vtType,'' AS vtCode " & _
                 "From BookPOParent Where Code NOT IN (Select Distinct C.Ref FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE P.Type IN ('0110PU','0110PC','0110PJ','0310PU','0310PC','0310PJ','0510FR','0710FR','0210OU','0210OC','0210OJ','0410TU','0410TC','0410TJ','0610FI','0810FI')) AND LEFT(Type,1)<>'O' AND RIGHT(Type ,1)<>'" & Left(VchCode, 1) & "' AND LEFT(Code ,1)<>'*' AND Date >='" & GetDate(MhDateInput1.Text) & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND " & _
                 IIf(FrmItemSelectionList.Option3.Value, "(DeliveredQuantityC+DeliveredQuantityB)<(EstQty01)", IIf(FrmItemSelectionList.Option1.Value, "(DeliveredQuantityC+DeliveredQuantityB)>=(EstQty01)", "1=1")) & " AND MaterialCentre IN (" & MatCList & ") AND IIF(Binder IS NOT NULL AND Binder<>'' ,Binder,IIF(BookPrinter IS NOT NULL AND BookPrinter<>'',BookPrinter,IIF(TitlePrinter IS NOT NULL AND TitlePrinter<>'',TitlePrinter,Laminator))) IN (" & AccountList & ") AND Book IN (" & ItemList & ") " & _
                 "ORDER BY AccountName," & Choose(Combo1.ListIndex + 1, "P.Date,pvtCode,vtCode", "pvtCode,vtCode", "ItemName,pvtCode,vtCode,P.Date") & ""
      ElseIf VchType = 36 Or VchType = 38 Then 'Sale & Purchase Order Status Summarized
            SQL = "SELECT IIF(A1.PrintName IS NOT NULL,LTRIM(A1.PrintName),IIF(A2.PrintName IS NOT NULL,A2.PrintName,IIF(A3.PrintName IS NOT NULL,LTRIM(A3.PrintName),LTRIM(A4.PrintName)))) As AccountName,RIGHT(T.Type,1)+'O/'+LTRIM(T.Name)+'/JW/'+IIF(FORMAT(T.Date,'MM')<4,Convert(Nvarchar,(Convert(int,FORMAT(T.Date,'yy'))-1)) +'-'+Convert(Nvarchar,FORMAT(T.Date,'yy')),Convert(Nvarchar,FORMAT(T.Date,'yy')) +'-'+ Convert(Nvarchar,Convert(int,FORMAT(T.Date,'yy'))+1)) As VchBillNo,T.Date As VchDate,I.PrintName As ItemName,T.EstQty01 As Ordered,ISNULL((SELECT SUM(Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE P.MaterialCentre IN (" & MatCList & ") AND P.Type IN ('0110PU','0110PC','0110PJ','0310PU','0310PC','0310PJ','0510FR','0710FR') AND Ref IN (T.Code)),0) As INward,ISNULL((SELECT SUM(Quantity) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE P.MaterialCentre IN (" & MatCList & ") AND " & _
                  "P.Type IN ('0210OU','0210OC','0210OJ','0410TU','0410TC','0410TJ','0610FI','0810FI') AND Ref IN (T.Code)),0) As Outward," & _
                  "ISNULL((SELECT Avg(C.Rate) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE P.MaterialCentre IN (" & MatCList & ") AND P.Type IN ('0110PU','0110PC','0110PJ','0310PU','0310PC','0310PJ','0510FR','0710FR','0210OU','0210OC','0210OJ','0410TU','0410TC','0410TJ','0610FI','0810FI') AND Ref IN (T.Code)),Avg(T.UnitRate)) AS Rate,ISNULL((SELECT SUM(C.Amount) FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code WHERE P.MaterialCentre IN (" & MatCList & ") AND P.Type IN ('0110PU','0110PC','0110PJ','0310PU','0310PC','0310PJ','0510FR','0710FR','0210OU','0210OC','0210OJ','0410TU','0410TC','0410TJ','0610FI','0810FI') AND Ref IN (T.Code)),SUM(T.EstQty01*T.UnitRate)) AS Amount," & _
                  "I.Code As ItemCode,IIF(A1.Code IS NOT NULL,LTRIM(A1.Code),IIF(A2.Code IS NOT NULL,LTRIM(A2.Code),IIF(A3.Code IS NOT NULL,LTRIM(A3.Code),LTRIM(A4.Code)))) As AccountCode,T.Code As vtCode,T.Name As vtNo,T.Type As vtType " & _
                  "FROM ((((BookPOParent T INNER JOIN BookMaster As I ON T.Book=I.Code) LEFT JOIN AccountMaster As A1 ON T.Binder=A1.Code) LEFT JOIN AccountMaster As A2 ON T.BookPrinter=A2.Code) LEFT JOIN AccountMaster As A3 ON T.TitlePrinter=A3.Code) LEFT JOIN AccountMaster As A4 ON T.Laminator=A4.Code WHERE LEFT(T.Type,1)<>'O' AND RIGHT(T.Type,1)<>'" & Left(VchCode, 1) & "' AND LEFT(T.Code,1)<>'*' AND T.Date>='" & GetDate(MhDateInput1.Text) & "' AND T.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & _
                   IIf(FrmItemSelectionList.Option3.Value, "T.DeliveredQuantityC+T.DeliveredQuantityB<T.EstQty01", IIf(FrmItemSelectionList.Option1.Value, "T.DeliveredQuantityC+T.DeliveredQuantityB>=T.EstQty01", "1=1")) & " AND IIF(A1.Code IS NOT NULL,LTRIM(A1.Code),IIF(A2.Code IS NOT NULL,LTRIM(A2.Code),IIF(A3.Code IS NOT NULL,LTRIM(A3.Code),LTRIM(A4.Code)))) IN (" & AccountList & ") AND I.Code IN (" & ItemList & ") " & _
                  "Group By T.TYPE,IIF(A1.PrintName IS NOT NULL,LTRIM(A1.PrintName),IIF(A2.PrintName IS NOT NULL,A2.PrintName,IIF(A3.PrintName IS NOT NULL,LTRIM(A3.PrintName),LTRIM(A4.PrintName)))),RIGHT(T.Type,1)+'O/'+LTRIM(T.Name)+'/JW/'+IIF(FORMAT(T.Date,'MM')<4,Convert(Nvarchar,(Convert(int,FORMAT(T.Date,'yy'))-1)) +'-'+Convert(Nvarchar,FORMAT(T.Date,'yy')),Convert(Nvarchar,FORMAT(T.Date,'yy')) +'-'+ Convert(Nvarchar,Convert(int,FORMAT(T.Date,'yy'))+1)),IIF(A1.Code IS NOT NULL,LTRIM(A1.Code),IIF(A2.Code IS NOT NULL,LTRIM(A2.Code),IIF(A3.Code IS NOT NULL,LTRIM(A3.Code),LTRIM(A4.Code)))),T.Date,I.PrintName,T.EstQty01,LTRIM(T.Name),T.Code,I.Code,T.Name " & _
                  "ORDER BY IIF(A1.PrintName IS NOT NULL,LTRIM(A1.PrintName),IIF(A2.PrintName IS NOT NULL,A2.PrintName,IIF(A3.PrintName IS NOT NULL,LTRIM(A3.PrintName),LTRIM(A4.PrintName))))," & Choose(Combo1.ListIndex + 1, "T.Date,LTRIM(T.Name)", "T.Code,LTRIM(T.Name)", "I.PrintName,LTRIM(T.Name),T.Code") & ""
      ElseIf VchType = 39 Or VchType = 40 Or VchType = 41 Or VchType = 42 Or VchType = 43 Or VchType = 44 Then 'Sale & Purchased Order Status Summrized
        SQL = "SELECT DISTINCT (Select PrintName From AccountMaster Where Code=P.Party) AS AccountName,LTRIM(P.Name) AS VchBillNo,P.Date AS VchDate,(Select PrintName From BookMaster Where Code=C.Item ) AS ItemName,P.ChallanNo,(P.ChallanDate),IIF(LEFT(BOM,4)+Right(BOM,2)IN ('1701FI','1801FI'),ABS(C.Quantity),0) AS Ordered,ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0) As Dispatched,(IIF(LEFT(BOM,4)+Right(BOM,2)IN ('1701FI','1801FI'),ABS(C.[Quantity]),0)-ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0)) As Balance," & _
                  "(P.Party) AS BCode,LTRIM(P.Name) AS GRNNo,P.Date,(Select Name From AccountMaster Where Code=P.Consignee) AS Consignee,(Select Name From AccountMaster Where Code=P.Party) AS Party,P.Remarks,LTRIM(P.Name) AS PO,(Select PrintName From BookMaster Where Code=C.Item ) AS Book,C.Quantity As Qty,C.Rate,C.Amount,P.BOX,P.Freight,P.TYPE, Right(P.Type,2)+'-'+LTRIM(P.Name) As MRNNo,Right(P.Type,2) As vtType,C.Item as iCode,(Select VchName From VchSeriesMaster Where Code=P.vchSeries) As TypeRef,C.BOM  As BOM,P.Code,C.RefCode As vtCode,P.Remarks AS RemarkC " & _
                  "FROM JobworkBVParent P INNER JOIN JobworkBVChild C ON P.Code=C.Code Left Join JobworkBVRef R ON R.VchCode=C.Code " & _
                  "WHERE LEFT((C.BOM),6) IN ('" & IIf(VchCode = "P", "1801FI", "1701FI") & "') AND Right(P.Type,1)<>'" & Left(VchCode, 1) & "' AND LEFT(P.Code,1)<>'*' AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & _
                  IIf(FrmItemSelectionList.Option3.Value, "ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0)<ABS(C.Quantity)", IIf(FrmItemSelectionList.Option1.Value, "ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0)>=ABS(C.Quantity)", IIf(FrmItemSelectionList.Option2.Value, "IIf(Right(P.Type,1)='P',ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0),ISNULL((Select ABS(Sum (Quantity)) From JobworkBVRef Where RefCode=C.RefCode AND VchCode<>C.Code),0))>=0", 1))) & "  " & _
                  "AND (P.Party) IN (" & AccountList & ") AND C.Item IN (" & ItemList & ") " & _
                  "ORDER BY " & Choose(Combo1.ListIndex + 1, "P.Date", "P.Code", "AccountName", "ItemName") & ""
      ElseIf VchType = 45 Then 'Sale & Purchase Order Status Detailed
         SQL = "Select VchDate As vtDate,IIF(Right(VchType,2)='" & IIf(VchCode = "P", "SO", "PO") & "',LTRIM(VchNo),'') As VchBillNo,VchDate As pvtDate,VchCode As pvtCode,(Select Name From BookMaster A Where A.Code=Item) As ItemName,(Select VchName From VchSeriesMaster Where Code=(Select VchSeries From JobworkBVParent P Where P.Code=VchCode)) As TypeRef,(Select ISNULL(ChallanDate,'') From JobworkBVParent Where Code=JobworkBVRef.VchCode) As ChallanDate,(Select ISNULL(ChallanNo,'') From JobworkBVParent Where Code=JobworkBVRef.VchCode) As ChallanNo,(Select Name From AccountMaster A Where A.Code=Party) As AccountName,(Select Name From AccountMaster A Where A.Code=(Select MaterialCentre From JobworkBVParent P Where P.Code=VchCode)) As MaterialCentre,(Select ISNULL(Remarks,'') From JobworkBVParent Where Code=JobworkBVRef.VchCode) As Remarks,(Select ABS(Quantity) From JobworkBVRef Where RefCode  =" & SCode & " And (Left(VchType,2) ='17' OR Left(VchType,2) ='18')) As Ordered," & _
                   "ISNULL(IIF(Quantity>0,Quantity,0),0)  As INward,ISNULL(IIF(Quantity<0,ABS(Quantity),0),0)  As OutWard,'0' As Pending,Rate,(ISNULL(IIF(Quantity>0,Quantity,0),0)-ISNULL(IIF(Quantity<0,ABS(Quantity),0),0))*Rate As Amount,ISNULL(VchNo,'') AS vtNO,ISNULL(VchCode,'') AS vtCode,VchType As vtType,VchType As pvtType FROM JobworkBVRef " & _
                   "WHERE RefCode=" & SCode & " AND Left(VchType,2) NOT IN ('','') ORDER BY " & Choose(Combo1.ListIndex + 1, "VchDate", "VchCode", "AccountName", "ItemName") & ""
      ElseIf VchType >= 46 And VchType <= 47 Then 'Pending Sale & Purchase Order Status Detailed
            If VchType = 46 Then
                        SQL = SQL + "Select VchType,ItemCode,UnitRate,VchCode,VchNo,VchDate,Item,BuyerCode,BuyerName,OrderedQty,Pending,PendingAmount,BilledQtyC,BilledQtyD,ChallanQty,DirectQty,ClearQty,CreatedBy,CreatedOn,Remarks From ("
            ElseIf VchType = 47 Then
                        SQL = SQL + "Select BuyerCode,BuyerName,Sum(OrderedQty) AS OrderedQty,SUM(Pending) AS Pending,Sum(PendingAmount) As PendingAmount,Sum(BilledQtyC) As BilledQtyC,SUM(BilledQtyD) As BilledQtyD,SUM(ChallanQty) AS ChallanQty,SUM(DirectQty) AS DirectQty,Sum(ClearQty) As ClearQty,CreatedBy,CreatedOn,Remarks From ("
            End If
                            SQL = SQL + "SELECT DISTINCT IIF(P.BookPrinter<>'',P.BookPrinter,IIF(P.TitlePrinter<>'',P.TitlePrinter,IIF(P.Laminator<>'',P.Laminator,IIF(P.Binder<>'',P.Binder,IIF(C.Vendor<>'',C.Vendor,'000000'))))) As BuyerCode,(Select PrintName From AccountMaster Where Code= IIF(P.BookPrinter<>'',P.BookPrinter,IIF(P.TitlePrinter<>'',P.TitlePrinter,IIF(P.Laminator<>'',P.Laminator,IIF(P.Binder<>'',P.Binder,IIF(C.Vendor<>'',C.Vendor,'000000')))))) As BuyerName,"
                        
                        If Combo1.ListIndex = 0 Then 'Direct 'Ordered-Delivered(Challan)-Billed(Direct)=Pending Quantity
                            SQL = SQL + "(P.EstQty01-P.DeliveredQuantityC-P.BilledAllB-ISNULL(J.Quantity,0)) As Pending,(P.EstQty01-P.DeliveredQuantityC-P.BilledAllB-ISNULL(J.Quantity,0))*P.UnitRate As PendingAmount,"
                        ElseIf Combo1.ListIndex = 1 Then 'Against Challan 'Delivered(Challan)-Billed(Challan)=Pending Quantity
                            SQL = SQL + "(P.DeliveredQuantityC-P.BilledAllC-ISNULL(J.Quantity,0)) As Pending,(P.DeliveredQuantityC-P.BilledAllC-ISNULL(J.Quantity,0))*P.UnitRate As PendingAmount,"
                        End If
                            
                            SQL = SQL + "'0000'+P.Type As VchType,P.Book As ItemCode,P.UnitRate,P.Code As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/JW' As VchNo,P.Date As VchDate,I.Name As Item,P.EstQty01 As OrderedQty,P.BilledAllC As BilledQtyC,P.BilledAllB As BilledQtyD,P.DeliveredQuantityC As ChallanQty,P.DeliveredQuantityB As DirectQty,ISNULL(J.Quantity,0) As ClearQty,J.CreatedBy,J.CreatedOn,J.Remarks "
                            SQL = SQL + "FROM ((BookPOParent P INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN BookPOChild0801 C ON P.Code=C.Code) Left Join JobworkBVClear J On P.code=J.RefCode "
                            SQL = SQL + "WHERE RIGHT(P.Type,1)=" & IIf(VchType = 46 Or VchType = 47, "'S'", "'P'") & "  AND "
                        
                        If Combo1.ListIndex = 0 Then 'Direct 'Ordered-Delivered(Challan)-Billed(Direct)=Pending Quantity
                            SQL = SQL + IIf(FrmItemSelectionList.Option3.Value, "(P.EstQty01-P.DeliveredQuantityC-P.BilledAllB-ISNULL(J.Quantity,0))>0", IIf(FrmItemSelectionList.Option1.Value, "(P.EstQty01-P.DeliveredQuantityC-P.BilledAllB-ISNULL(J.Quantity,0))<=0", "1=1")) & ""
                        ElseIf Combo1.ListIndex = 1 Then 'Against Challan 'Delivered(Challan)-Billed(Challan)=Pending Quantity
                            SQL = SQL + IIf(FrmItemSelectionList.Option3.Value, "(P.DeliveredQuantityC-P.BilledAllC-ISNULL(J.Quantity,0))>0", IIf(FrmItemSelectionList.Option1.Value, "(P.DeliveredQuantityC-P.BilledAllC-ISNULL(J.Quantity,0))<=0 AND (P.DeliveredQuantityC+P.BilledAllB+ISNULL(J.Quantity,0))>=P.EstQty01", "1=1")) & ""
                        End If
                            
                            SQL = SQL + " AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND IIF(P.BookPrinter<>'',P.BookPrinter,IIF(P.TitlePrinter<>'',P.TitlePrinter,IIF(P.Laminator<>'',P.Laminator,IIF(P.Binder<>'',P.Binder,IIF(C.Vendor<>'',C.Vendor,'000000'))))) IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  P.Book IN (" & IIf(SCode <> "", SCode, ItemList) & ") "
            If VchType = 46 Then
                            SQL = SQL + " ) AS TBL ORDER BY BuyerName,Item,VchDate,VchNo"
            ElseIf VchType = 47 Then
                            SQL = SQL + " ) AS TBL Group BY BuyerCode,BuyerName,CreatedBy,CreatedOn,Remarks "
                            SQL = SQL + " ORDER BY BuyerName"
            End If
      ElseIf Right(VchType, 2) = 48 And Left(VchType, 2) = "04" Then ' Sales Ledger
            SQL = "Select P.Date As VchDate,P.Name As VchBillNo,V.VchName AS Type,V.Name AS VchSeries,(Select Name From BookMaster Where Code=Item) As Item,(Select Name From AccountMaster Where Code=Party) AS Party,ISNULL(ABS(Quantity),0) As INWard,0 As OutWard,'Unit' As Unit,C.Rate,C.Amount,(Select Name From AccountMaster Where Code=MaterialCentre) As MaterialCentre,Type As VchType, BOM FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code  INNER JOIN VchSeriesMaster V ON V.Code=P.vchSeries WHERE LEFT(P.Type,2) = '" & Left(VchType, 2) & "' AND '" & Left(VchType, 2) & "'='04' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & AccountList & ") AND  C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")  AND P.VchSeries IN (" & ItemGroupList & ") AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') "
            
      ElseIf Right(VchType, 2) = 48 Then ' Sales Ledger
            SQL = "SELECT C.Code As VchCode,P.Date As VchDate,P.Name As VchBillNo,0 As INWard,ISNULL(SUM(ABS(Quantity)),0) As OutWard,'Units' As Unit,0 As Rate,Sum(C.Amount) As Amount,BOM,V.VchName AS Type,V.Name AS VchSeries,'' As Item,A.Name As Party,A.Name As MaterialCentre,P.Type VchType,C.Code As VchCode FROM JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code INNER JOIN VchSeriesMaster V ON V.Code=P.vchSeries INNER JOIN BookMaster I ON I.Code=C.Item INNER JOIN AccountMaster A ON A.Code=P.Party " & _
                  "WHERE LEFT(P.Type,2)='" & Left(VchType, 2) & "' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.Party IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND  C.Item IN (" & IIf(SCode <> "", SCode, ItemList) & ")  AND P.VchSeries IN (" & ItemGroupList & ")  AND SubString(P.Type,3,2)='10' And Right(BOM,2) NOT IN ('MO','ME','MF','BM') " & _
                  "Group By C.Code,P.Date,P.Name,BOM,V.VchName,V.Name,A.Name,P.Type " & _
                  "Order By V.Name,C.Code ASC "
'''''      ElseIf VchType = 103 Or VchType = 104 Or VchType = 105 Then ' WIP & RM Ledger
'''''            SQL = ""
'''''                If FrmItemSelectionList.Check1.Value And VchType = 103 Then SQL = SQL + "Select UFG,'' Category,UFGCode,'' [UFGReq/UNIT],'' UFGPages,'' Color,SUM(Stock) AS Stock,SUM(SalesOrder) AS SalesOrder,SUM(Dispatched) As Dispatched,SUM(PendingSO) AS PendingSO,SUM(DeficientSalesOrder) AS DeficientSalesOrder,SUM(UFGRequired) AS UFGRequired,(SELECT Convert(Numeric,ISNULL(dbo.ufnGetUFGStock('4',UFGCode,'000000','XX','XXXXXX','" & GetDate(MhDateInput2.Text) & "'),0)))As UFGStock,SUM(UFGRequired)+(SELECT Convert(Numeric,ISNULL(dbo.ufnGetUFGStock('4',UFGCode,'000000','XX','XXXXXX','" & GetDate(MhDateInput2.Text) & "'),0))) As FinalUFGRequired From( "
'''''                If VchType = 103 Or VchType = 104 Then SQL = SQL + "SELECT "
'''''                If VchType = 103 Then SQL = SQL + "UFG,Category,UFGCode,[UFGReq/UNIT], UFGPages,Color,FG,SUM(Stock) AS Stock,SUM(SalesOrder) AS SalesOrder,SUM(Dispatched) As Dispatched,SUM(PendingSO) AS PendingSO,SUM(DeficientSalesOrder) AS DeficientSalesOrder,SUM(UFGRequired) AS UFGRequired,SUM(UFGStock) As UFGStock,SUM(FinalUFGRequired) As FinalUFGRequired "
'''''                If VchType = 104 Then SQL = SQL + "SubUFG,SubUFG_Make,SubUFG_GSM,SubUFG_CUTOFF,SubUFGCode,SUBUFGCategory,[SubUFGReq/UNIT],[Weight/Unit],UOM_Name,UOM,SUM(SUBUFGReqSheets) AS SUBUFGReqSheets,SUM(SUBUFGReqKg) AS SUBUFGReqKg,SUM(SubUFGStkUOM) AS SubUFGStkUOM,SUM(SubUFGStockKg) As SubUFGStockKg,SUM(FinalSUBUFGReqKg) AS FinalSUBUFGReqKg "
'''''                If VchType = 103 Or VchType = 104 Then SQL = SQL + "From  ("
'''''                SQL = SQL + "Select*,Convert(Numeric(12,3),(PARSENAME(SubUFGStkUOM,2)+(SubUFGStkUOM-PARSENAME(SubUFGStkUOM,2))*2)*[Weight/Unit]) AS SubUFGStockKg,IIF(SUBUFGReqKg + Convert(Numeric(12, 3), (PARSENAME(SubUFGStkUOM, 2) + (SubUFGStkUOM - PARSENAME(SubUFGStkUOM, 2)) * 2) * [Weight/Unit]) < 0, SUBUFGReqKg + Convert(Numeric(12, 3), (PARSENAME(SubUFGStkUOM, 2) + (SubUFGStkUOM - PARSENAME(SubUFGStkUOM, 2)) * 2) * [Weight/Unit]), 0) As FinalSUBUFGReqKg " & _
'''''                    "From(Select*,Convert(Numeric(12,3),IIF(FinalUFGRequired<0,FinalUFGRequired*[SubUFGReq/UNIT],0)) AS [SUBUFGReqSheets],Convert(Numeric(12,3),IIF(FinalUFGRequired<0,FinalUFGRequired*[SubUFGReq/UNIT],0)/UOM*[Weight/Unit])  As [SUBUFGReqKg],(SELECT Convert(Numeric(12,3),dbo.ufnGetPaperStock('000000',SubUFGCode,'XX','XXXXXX','" & GetDate(MhDateInput2.Text) & "')) As Col1 FROM PaperMaster P Where P.Code=SubUFGCode)As SubUFGStkUOM " & _
'''''                    "From (Select*,(Select(Select Name From PaperMaster Where Code=C.Item) As UFG FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2' AND C.code =UFGCode) AS SubUFG,(Select(Select Make+'-'+SubMake From PaperMaster Where Code=C.Item) As UFG FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2' AND C.code =UFGCode) AS SubUFG_Make,(Select(Select GSM From PaperMaster Where Code=C.Item) As UFG FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2'  AND C.code =UFGCode) AS SubUFG_GSM,(Select(Select cmWidth From PaperMaster Where Code=C.Item) As UFG FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2'  AND C.code =UFGCode) AS SubUFG_CUTOFF," & _
'''''                    "(Select(Select Code From PaperMaster Where Code=C.Item) As UFG FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2' AND C.code =UFGCode) AS SubUFGCode,(Select C.Category FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2' AND C.code =UFGCode) As [SUBUFGCategory],(Select C.Quantity FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2' AND C.code =UFGCode) AS [SubUFGReq/UNIT],(Select Distinct I.[Weight/Unit] FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2'   AND C.code =UFGCode) AS [Weight/Unit],(Select (Select Name From GeneralMaster Where Code=I.UOM) As UFG FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2'   AND C.code =UFGCode) AS [UOM_Name],(Select (Select Value1 From GeneralMaster Where Code=I.UOM) As UFG FROM PaperMaster I Left JOIN BookChild01 C ON I.Code=C.Item WHERE C.Category='2' AND C.code =UFGCode) AS [UOM] " & _
'''''                    "From (Select *,UFGStock+((Stock-PendingSO)*[UFGReq/UNIT]) As [FinalUFGRequired] " & _
'''''                    "From (Select *,(Stock-PendingSO) as DeficientSalesOrder,((Stock-PendingSO)*[UFGReq/UNIT]) As [UFGRequired],(SELECT Convert(Numeric,ISNULL(dbo.ufnGetUFGStock('4',UFGCode,'000000','XX','XXXXXX','" & GetDate(MhDateInput2.Text) & "'),0)) As Col1 FROM BookMaster I Where I.Code=UFGCode)As UFGStock " & _
'''''                    "From (SELECT Distinct *,(Select Name From BookMaster Where Code=ItemCode) FG,[UFGReq/UNIT]*ISNULL((Select AVG(Pages*Ups/Sets) From BookChild06 Where Code=ItemCode),0) As UFGPages,ISNULL(Convert(nvarchar,ISNULL((Select Left(Value1,1) From GeneralMaster Where Code= (Select Top 1 FrontPrintingType From BookChild06 Where Code=UFGCode)),0))+'+'+Convert(nvarchar,ISNULL((Select Left(Value1,1) From GeneralMaster Where Code= (Select Top 1 BackPrintingType From BookChild06 Where Code=UFGCode)),0)),'') As Color," & _
'''''                    "(SELECT dbo.ufnGetItemStock(" & (AccountList) & ",ItemCode,'XX','XXXXXX','" & GetDate(MhDateInput2.Text) & "') As Col1 FROM BookMaster I WHERE I.Type='F' AND I.Code=ItemCode) As Stock,ISNULL((SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=ItemCode),0)+ISNULL((SELECT SUM(ABS(R.Quantity)) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND " & _
'''''                    "C.Item=ItemCode AND R.RefCode=C.RefCode AND VchCode<>C.Code),0) AS SalesOrder," & _
'''''                    "ISNULL((SELECT SUM(ABS(R.Quantity)) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=ItemCode AND R.RefCode=C.RefCode AND VchCode<>C.Code),0) As Dispatched,ISNULL((SELECT SUM(R.Quantity) " & _
'''''                    "FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=ItemCode),0) As PendingSO,ISNULL((SELECT CONVERT(DECIMAL(12,2),Avg(R.Rate)) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=ItemCode),0) As PendingSORate,ISNULL((SELECT SUM(R.Quantity*R.Rate) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND " & _
'''''                    "P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND " & _
'''''                    "P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=ItemCode),0) As PendingSOAmount " & _
'''''                    "From(SELECT I.Name As Item,(Select Name From GeneralMaster Where Code= I.BindingType) AS Binding,(Select Name AS Size From GeneralMaster Where Code= I.FinishSize) AS FinishSize,ISNULL((Select Sum(Pages) From BookChild06 Where Code=I.Code),0) As Pages,I.Code As ItemCode,I.Price,(Select Name From BookMaster Where Code=I1.Item) As UFG,I1.Category,(Select Code From BookMaster Where Code=I1.Item) As UFGCode,I1.quantity As [UFGReq/UNIT] " & _
'''''                    "From BookMaster I Left JOIN BookChild01 I1 ON I.Code=I1.Code WHERE I.Type='F' AND I1.Category='4' "
'''''                If FrmItemSelectionList.Option1.Value Then    'Close
'''''                    SQL = SQL + "AND ISNULL((SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=I.Code),0)=0 "
'''''                ElseIf FrmItemSelectionList.Option3.Value Then    'Pending
'''''                    SQL = SQL + "AND ISNULL((SELECT SUM(R.Quantity) FROM (JobWorkBVParent P INNER JOIN JobWorkBVChild C ON P.Code=C.Code) INNER JOIN JobWorkBVRef R ON C.RefCode=R.RefCode WHERE LEFT(P.Type,2)='18' AND P.Date BETWEEN '" & GetDate(MhDateInput1.Text) & "' AND '" & GetDate(MhDateInput2.Text) & "' AND P.MaterialCentre IN (" & IIf(sMcCode <> "", sMcCode, AccountList) & ")  AND C.Item=I.Code),0)<>0 "
'''''                End If
'''''            SQL = SQL + "AND I.Code IN (" & IIf(SCode <> "", SCode, ItemList) & ")) As FG_UFG) AS PendingSO) AS SUBUFG) AS TBL) AS TBL) AS TBL "
'''''                If VchType = 105 Then SQL = SQL + "ORDER BY " & Choose(Combo1.ListIndex + 1, "Item,UFG,SubUFG ASC", "Item,UFG,SubUFG DESC", "UFG,Item,SubUFG ASC", "UFG,Item,SubUFG DESC", "SubUFG,UFG,Item ASC", "SubUFG,UFG,Item DESC") & " "
'''''                If VchType = 103 Or VchType = 104 Then SQL = SQL + ") AS TBL "
'''''                If VchType = 103 Then
'''''                    If FrmItemSelectionList.Check1.Value Then
'''''                        SQL = SQL + " Group By UFG,Category,UFGCode,[UFGReq/UNIT],UFGPages,Color,FG ) AS TBL Group By UFG,UFGCode ORDER BY UFG ASC "
'''''                    Else
'''''                        SQL = SQL + " Group By UFG,Category,UFGCode,[UFGReq/UNIT],UFGPages,Color,FG ORDER BY UFG ASC "
'''''                    End If
'''''                End If
'''''                If VchType = 104 Then SQL = SQL + " Group By SubUFGCode,SubUFG,SubUFG_Make, SubUFG_GSM, SubUFG_CUTOFF,SUBUFGCategory,[SubUFGReq/UNIT],[Weight/Unit],UOM_Name,UOM ORDER BY SubUFG ASC "
      End If
        Screen.MousePointer = vbHourglass
        MdiMainMenu.StatusBar1.Panels(2).Text = "Wait For Data Spooling !!!"
        If rstStockLedger.State = adStateOpen Then rstStockLedger.Close
        rstStockLedger.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
            
        MdiMainMenu.MousePointer = vbHourglass
        ShowProgressInStatusBar True
        Timer1.Enabled = True
        If rstStockLedger.RecordCount = 0 And oVchType <> "" Then VchType = oVchType
        VSViewPort1.Visible = False
        VSPrinter1.Visible = False
        VSFlexGrid1.Visible = False
        If rstStockLedger.RecordCount = 0 Then
            With fpSpread1
                .MaxCols = 19: .MaxRows = 28
                For C = 1 To .MaxCols
                    fpSpread1.Col = C: fpSpread1.Row = SpreadHeader: .Text = " "
                Next
                For R = 1 To .MaxRows
                    fpSpread1.Col = 0: fpSpread1.Row = R: .Text = " "
                Next
                .ClearRange -1, 1, .MaxCols, .MaxRows, False
                Mh3dLabel11.Caption = "": Mh3dLabel10.Caption = ""
                MsgBox "No Records Found....", vbInformation, "Easy Publish...Reports !!! "
            End With
                MdiMainMenu.MousePointer = vbNormal
                ShowProgressInStatusBar False
                Timer1.Enabled = False
                Screen.MousePointer = vbNormal: Exit Sub
        ElseIf rstStockLedger.RecordCount > 5000 Or VchType = 103 Or VchType = 104 Or VchType = 105 Then
                VSFlexGrid1.Visible = True
                Call PublishGrid
                VSFlexFlag = True
        End If
If VSFlexFlag = False Then
    Dim n As Integer
    With fpSpread1
            If .DataRowCnt = 0 Then
            Else
            n = .DataRowCnt:
                fpSpread1.RowHeight(n) = 12.75
           End If
            .ClearRange -1, 1, .MaxCols, .MaxRows, False
            Dim K As Integer
            If VchType >= 35 And VchType <> 48 Then K = 9999 Else K = rstStockLedger.RecordCount
            ' Set number of columns and rows
                
                If VchType = 46 Then fpSpread1.MaxCols = 38 Else fpSpread1.MaxCols = 35
                fpSpread1.MaxRows = IIf(K < 27, 27, K + 1)
                Call FormatCol
                Call Check1_Click
                Call FormatHeader

    
    If (VchType >= 34 And VchType <= 45) Or VchType = 49 Then
            Call Print_fpSpread
    Else
        rstStockLedger.MoveFirst
        Do While Not rstStockLedger.EOF
            If VchType = 0 Or VchType = 1 Or VchType = 2 Or VchType = 33 Then
                    If PendingCheck.Value Then
                    If Val(rstStockLedger.Fields("PendingPO").Value) = 0 And Val(rstStockLedger.Fields("PendingSO").Value) = 0 Then GoTo NXT
                    End If
                        Stock = Val(rstStockLedger.Fields("PurchaseChallan").Value) - Val(rstStockLedger.Fields("PurchaseReturnChallan").Value) - Val(rstStockLedger.Fields("SalesChallan").Value) + Val(rstStockLedger.Fields("SalesReturnChallan").Value) + Val(rstStockLedger.Fields("Purchase").Value) - Val(rstStockLedger.Fields("PurchaseReturn").Value) - Val(rstStockLedger.Fields("Sales").Value) + Val(rstStockLedger.Fields("SalesReturn").Value) + Val(rstStockLedger.Fields("StockJournalIN").Value) - Val(rstStockLedger.Fields("StockJournalOUT").Value) + Val(rstStockLedger.Fields("StockTransferIN").Value) - Val(rstStockLedger.Fields("StockTransferOUT").Value)
                    If VchType <= 2 Then EffectiveStock = Stock + Val(rstStockLedger.Fields("PendingPO").Value) - Val(rstStockLedger.Fields("PendingSO").Value)
                    If VchType = 33 Then EffectiveStock = Stock + Val(rstStockLedger.Fields("PendingPO").Value) - Val(rstStockLedger.Fields("PendingSO").Value) - Val(rstStockLedger.Fields("SQ").Value)
                    If NegativeStock.Value Then
                        If EffectiveStock >= 0 Then GoTo NXT
                    End If
                    If ZeroStock.Value Then
                    If TDBNumber1.Value = 0 Then
                        If EffectiveStock <> TDBNumber1.Value Then GoTo NXT
                    Else
                        If EffectiveStock >= TDBNumber1.Value Then GoTo NXT
                    End If
                    End If
            ElseIf VchType = 3 Or VchType = 7 Or VchType = 21 Or VchType = 25 Then
                If TDBNumber1.Value = 0 Then
                    If Val(rstStockLedger.Fields("Sales").Value) = TDBNumber1.Value Then GoTo NXT
                Else
                    If Val(rstStockLedger.Fields("Sales").Value) < TDBNumber1.Value Then GoTo NXT
                End If
            ElseIf VchType = 4 Or VchType = 8 Or VchType = 22 Or VchType = 26 Then
                If TDBNumber1.Value = 0 Then
                    If Val(rstStockLedger.Fields("SalesReturn").Value) = TDBNumber1.Value Then GoTo NXT
                Else
                    If Val(rstStockLedger.Fields("SalesReturn").Value) < TDBNumber1.Value Then GoTo NXT
                End If
            ElseIf VchType = 5 Or VchType = 9 Or VchType = 23 Or VchType = 27 Then
                If TDBNumber1.Value = 0 Then
                    If Val(rstStockLedger.Fields("Sales").Value) = TDBNumber1.Value And Val(rstStockLedger.Fields("SalesReturn").Value) = TDBNumber1.Value Then GoTo NXT
                Else
                    If Val(rstStockLedger.Fields("Sales").Value) < TDBNumber1.Value And Val(rstStockLedger.Fields("SalesReturn").Value) < TDBNumber1.Value Then GoTo NXT
                End If
            ElseIf VchType = 6 Or VchType = 10 Or VchType = 24 Or VchType = 28 Then
                If TDBNumber1.Value = 0 Then
                    If Val(rstStockLedger.Fields("Sales").Value) = TDBNumber1.Value Then GoTo NXT
                Else
                    If Val(rstStockLedger.Fields("Sales").Value) < TDBNumber1.Value Then GoTo NXT
                End If
            ElseIf VchType >= 29 And VchType <= 30 Then
                If TDBNumber1.Value = 0 Then
                    If Val(rstStockLedger.Fields("Pending").Value) = TDBNumber1.Value Then GoTo NXT
                Else
                    If Val(rstStockLedger.Fields("Pending").Value) < TDBNumber1.Value Then GoTo NXT
                End If
            End If
        i = i + 1
'Pending Order
        If VchType >= 29 And VchType <= 30 Then
            .SetText 1, i, rstStockLedger.Fields("Date").Value
            .SetText 2, i, rstStockLedger.Fields("VchBillNo").Value
            .SetText 3, i, rstStockLedger.Fields("Type").Value: .Col = 3: .Row = i: .TypeHAlign = TypeHAlignCenter
            .SetText 5, i, rstStockLedger.Fields("Details").Value
            .SetText 6, i, Val(rstStockLedger.Fields("Ordered").Value)
            .SetText 24, i, Val(rstStockLedger.Fields("Pending").Value)
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstStockLedger.Fields("Rate").Value)
            .SetText 27, i, Val(rstStockLedger.Fields("Amount").Value)
            .SetText 32, i, rstStockLedger.Fields("VchCode").Value
            .SetText 35, i, rstStockLedger.Fields("VchType").Value
            dPrint = dPrint + 1
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
'Item Ledger
        ElseIf VchType = 31 Or Right(VchType, 2) = 48 Then
            .SetText 1, i, rstStockLedger.Fields("VchDate").Value
            .SetText 2, i, rstStockLedger.Fields("VchBillNo").Value
            .SetText 3, i, rstStockLedger.Fields("Type").Value ': .Col = 3: .Row = i: .TypeHAlign = TypeHAlignCenter
            If Right(VchType, 2) = 48 Then .Col = 4: .Row = i: .CellType = CellTypeEdit: .SetText 4, i, rstStockLedger.Fields("VchSeries").Value
            .SetText 5, i, rstStockLedger.Fields("Party").Value:
                Credit = Val(rstStockLedger.Fields("INward").Value)
            .SetText 6, i, Val(rstStockLedger.Fields("INward").Value)
                Debit = Val(rstStockLedger.Fields("OutWard").Value)
            .SetText 23, i, Val(rstStockLedger.Fields("OutWard").Value)
                Bal = Bal + Credit - Debit
            .SetText 24, i, Bal
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstStockLedger.Fields("Rate").Value)
            If VchType = 31 Then .SetText 27, i, Bal * Val(rstStockLedger.Fields("Rate").Value) Else .SetText 27, i, Val(rstStockLedger.Fields("Amount").Value)
            .SetText 32, i, rstStockLedger.Fields("VchCode").Value
            .SetText 35, i, rstStockLedger.Fields("VchType").Value
            dPrint = dPrint + 1
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
'Pending Quotations and Short Item Analysis
        ElseIf VchType = 33 Then
            .SetText 3, i, rstStockLedger.Fields("Item").Value: .Col = 3: .Row = i: .TypeHAlign = TypeHAlignLeft
            .SetText 4, i, Val(rstStockLedger.Fields("MRP").Value)
            .SetText 5, i, rstStockLedger.Fields("ItemGroup").Value
            .SetText 6, i, Stock + Val(rstStockLedger.Fields("OPBAL").Value)
            .SetText 7, i, "Units"
            .SetText 8, i, Val(rstStockLedger.Fields("SQ").Value)
            .SetText 18, i, Val(rstStockLedger.Fields("PendingPO").Value)
            .SetText 19, i, Val(rstStockLedger.Fields("PendingSO").Value)
            .SetText 20, i, EffectiveStock
            .SetText 22, i, EffectiveStock * Val(rstStockLedger.Fields("MRP").Value)
            .SetText 32, i, (rstStockLedger.Fields("Code").Value)
            .SetText 35, i, rstStockLedger.Fields("HSNCode").Value
            dPrint = dPrint + 1
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
'Voucher Details
        ElseIf VchType = 32 Then
            .SetText 3, i, rstStockLedger.Fields("MaterialCentreName").Value
            .SetText 6, i, Val(rstStockLedger.Fields("Opening").Value)
            .SetText 20, i, Val(rstStockLedger.Fields("cINWard").Value)
            .SetText 23, i, Val(rstStockLedger.Fields("cOutWard").Value)
            .SetText 24, i, Val(rstStockLedger.Fields("Closing").Value)
            .SetText 25, i, "Units"
            .SetText 35, i, rstStockLedger.Fields("MaterialCentre").Value
            dPrint = dPrint + 1
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
'Sales,Sales Return,Purchase,Purchase Return
        ElseIf VchType >= 46 And VchType <= 48 Then
                If VchType = 46 Then
                    .SetText 1, i, rstStockLedger.Fields("VchDate").Value
                    .SetText 2, i, rstStockLedger.Fields("VchNo").Value
                    .SetText 3, i, rstStockLedger.Fields("Item").Value ': .Col = 3: .Row = i: .TypeHAlign = TypeHAlignLeft
                    .SetText 4, i, Val(rstStockLedger.Fields("UnitRate").Value)
                End If
                    .SetText 5, i, "  " + rstStockLedger.Fields("BuyerName").Value
                    .SetText 6, i, Val(rstStockLedger.Fields("OrderedQty").Value)
                    .SetText 7, i, "Units"
                If Combo1.ListIndex = 0 Then
                    If Val(rstStockLedger.Fields("OrderedQty").Value) - Val(rstStockLedger.Fields("ClearQty").Value) - Val(rstStockLedger.Fields("ChallanQty").Value) - Val(rstStockLedger.Fields("BilledQtyD").Value) < 0 Then
                        .SetText 8, i, Val("0")
                    Else
                        .SetText 8, i, Val(rstStockLedger.Fields("OrderedQty").Value) - Val(rstStockLedger.Fields("ClearQty").Value) - Val(rstStockLedger.Fields("ChallanQty").Value) - Val(rstStockLedger.Fields("BilledQtyD").Value)
                    End If
                If VchType = 46 Then
                    If Val(rstStockLedger.Fields("OrderedQty").Value) - Val(rstStockLedger.Fields("ClearQty").Value) - Val(rstStockLedger.Fields("ChallanQty").Value) - Val(rstStockLedger.Fields("BilledQtyD").Value) < 0 Then
                        .SetText 9, i, Val("0")
                    Else
                        .SetText 9, i, (Val(rstStockLedger.Fields("OrderedQty").Value) - Val(rstStockLedger.Fields("ClearQty").Value) - Val(rstStockLedger.Fields("ChallanQty").Value) - Val(rstStockLedger.Fields("BilledQtyD").Value)) * Val(rstStockLedger.Fields("UnitRate").Value)
                    End If
                ElseIf VchType = 47 Then
                    .SetText 9, i, Val(rstStockLedger.Fields("PendingAmount").Value)
                End If
                ElseIf Combo1.ListIndex = 1 Then
                    If Val(rstStockLedger.Fields("ChallanQty").Value) - Val(rstStockLedger.Fields("ClearQty").Value) - Val(rstStockLedger.Fields("BilledQtyC").Value) > 0 Then .SetText 8, i, Val(rstStockLedger.Fields("ChallanQty").Value) - Val(rstStockLedger.Fields("ClearQty").Value) - Val(rstStockLedger.Fields("BilledQtyC").Value) Else .SetText 8, i, Val(0)
                    If Val(rstStockLedger.Fields("OrderedQty").Value) <= Val(rstStockLedger.Fields("ChallanQty").Value) - Val(rstStockLedger.Fields("BilledQtyC").Value) Then .Col = 8: .Row = i: .FontBold = True: .FontSize = 10:  .ForeColor = vbRed: .Col = 9: .Row = i: .FontBold = True: .FontSize = 10: .ForeColor = vbRed: .Col = 11: .Row = i: .FontBold = True: .FontSize = 10: .ForeColor = vbRed: .Col = 12: .Row = i: .FontBold = True: .FontSize = 10: .ForeColor = vbRed
                    If Val(rstStockLedger.Fields("OrderedQty").Value) <= Val(rstStockLedger.Fields("ChallanQty").Value) + Val(rstStockLedger.Fields("BilledQtyD").Value) Then .Col = 8: .Row = i: .FontBold = True: .FontSize = 10:  .ForeColor = vbRed: .Col = 9: .Row = i: .FontBold = True: .FontSize = 10: .ForeColor = vbRed: .Col = 11: .Row = i: .FontBold = True: .FontSize = 10: .ForeColor = vbRed: .Col = 12: .Row = i: .FontBold = True: .FontSize = 10: .ForeColor = vbRed
                If VchType = 46 Then
                    If Val(rstStockLedger.Fields("ChallanQty").Value) - Val(rstStockLedger.Fields("ClearQty").Value) - Val(rstStockLedger.Fields("BilledQtyC").Value) > 0 Then .SetText 9, i, (Val(rstStockLedger.Fields("ChallanQty").Value) - Val(rstStockLedger.Fields("ClearQty").Value) - Val(rstStockLedger.Fields("BilledQtyC").Value)) * Val(rstStockLedger.Fields("UnitRate").Value) Else .SetText 9, i, Val(0)
                ElseIf VchType = 47 Then
                    .SetText 9, i, Val(rstStockLedger.Fields("PendingAmount").Value)
                End If
                    End If
                    .SetText 10, i, Val(rstStockLedger.Fields("BilledQtyC").Value)
                    .SetText 11, i, Val(rstStockLedger.Fields("BilledQtyD").Value)
                    .SetText 12, i, Val(rstStockLedger.Fields("ChallanQty").Value)
                    .SetText 13, i, Val(rstStockLedger.Fields("DirectQty").Value)
                    .SetText 14, i, Val(rstStockLedger.Fields("ClearQty").Value)
                If VchType = 46 Then
                    .SetText 25, i, rstStockLedger.Fields("VchType").Value
                    .SetText 26, i, Val(rstStockLedger.Fields("UnitRate").Value)
                End If
                    .SetText 32, i, rstStockLedger.Fields("BuyerCode").Value
                If VchType = 46 Then
                    .SetText 34, i, rstStockLedger.Fields("ItemCode").Value
                    .SetText 35, i, rstStockLedger.Fields("VchCode").Value
                    .SetText 36, i, rstStockLedger.Fields("CreatedBy").Value
                    .SetText 37, i, rstStockLedger.Fields("CreatedOn").Value
                    .SetText 38, i, rstStockLedger.Fields("Remarks").Value
                    .Col = 15: .Row = i: .CellType = CellTypeCheckBox: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                End If
                    dPrint = dPrint + 1
                    MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
        ElseIf (VchType >= 3 And VchType <= 10) Or (VchType >= 21 And VchType <= 28) Or (VchType >= 53 And VchType <= 60) Or (VchType >= 61 And VchType <= 68) Then
            .SetText 3, i, rstStockLedger.Fields("Item").Value: .Col = 3: .Row = i: .TypeHAlign = TypeHAlignLeft
            .SetText 4, i, Val(rstStockLedger.Fields("MRP").Value)
            .SetText 5, i, rstStockLedger.Fields("ItemGroup").Value
            .SetText 8, i, Val(rstStockLedger.Fields("Purchase").Value)
                    PurchaseTotal = PurchaseTotal + Val(rstStockLedger.Fields("Purchase").Value)
            .SetText 9, i, Val(rstStockLedger.Fields("PurchaseReturn").Value)
                    PurchaseReturnTotal = PurchaseReturnTotal + Val(rstStockLedger.Fields("PurchaseReturn").Value)
            .SetText 12, i, Val(rstStockLedger.Fields("Sales").Value)
                    SalesTotal = SalesTotal + Val(rstStockLedger.Fields("Sales").Value)
            .SetText 13, i, Val(rstStockLedger.Fields("SalesReturn").Value)
                    SalesReturnTotal = SalesReturnTotal + Val(rstStockLedger.Fields("SalesReturn").Value)
            .SetText 23, i, Val(rstStockLedger.Fields("Purchase").Value) - Val(rstStockLedger.Fields("PurchaseReturn").Value) 'NetPurchase
                    NetPurchaseTotal = NetPurchaseTotal + Val(rstStockLedger.Fields("Purchase").Value) - Val(rstStockLedger.Fields("PurchaseReturn").Value) 'NetPurchase
            .SetText 24, i, Val(rstStockLedger.Fields("Sales").Value) - Val(rstStockLedger.Fields("SalesReturn").Value) 'NetSales
                    NetSalesTotal = NetSalesTotal + Val(rstStockLedger.Fields("Sales").Value) - Val(rstStockLedger.Fields("SalesReturn").Value) 'NetSales
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstStockLedger.Fields("PurchaseAmount").Value)
                    PurchaseAmountTotal = PurchaseAmountTotal + Val(rstStockLedger.Fields("PurchaseAmount").Value)
            .SetText 27, i, Val(rstStockLedger.Fields("SalesAmount").Value)
                    SalesAmountTotal = SalesAmountTotal + Val(rstStockLedger.Fields("SalesAmount").Value)
            .SetText 28, i, Val(rstStockLedger.Fields("PurchaseReturnAmount").Value)
                    PurchaseReturnAmountTotal = PurchaseReturnAmountTotal + Val(rstStockLedger.Fields("PurchaseReturnAmount").Value)
            .SetText 29, i, Val(rstStockLedger.Fields("SalesReturnAmount").Value)
                    SalesReturnAmountTotal = SalesReturnAmountTotal + Val(rstStockLedger.Fields("SalesReturnAmount").Value)
            .SetText 30, i, Val(rstStockLedger.Fields("PurchaseAmount").Value) - Val(rstStockLedger.Fields("PurchaseReturnAmount").Value) 'NetPurchaseAmount
                    NetPurchaseAmountTotal = NetPurchaseAmountTotal + Val(rstStockLedger.Fields("PurchaseAmount").Value) - Val(rstStockLedger.Fields("PurchaseReturnAmount").Value) 'NetPurchaseAmount
            .SetText 31, i, Val(rstStockLedger.Fields("SalesAmount").Value) - Val(rstStockLedger.Fields("SalesReturnAmount").Value) 'NetSalesAmount
                    NetSalesAmountTotal = NetSalesAmountTotal + Val(rstStockLedger.Fields("SalesAmount").Value) - Val(rstStockLedger.Fields("SalesReturnAmount").Value) 'NetSalesAmount
            .SetText 32, i, (rstStockLedger.Fields("Code").Value)
            .SetText 33, i, 0
            .SetText 34, i, 0
            .SetText 35, i, rstStockLedger.Fields("HSNCode").Value
            dPrint = dPrint + 1
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
'Stock , Sale, Purchase
        ElseIf VchType <= 2 Then
            .SetText 3, i, rstStockLedger.Fields("Item").Value: .Col = 3: .Row = i: .TypeHAlign = TypeHAlignLeft
            .SetText 4, i, Val(rstStockLedger.Fields("MRP").Value)
            .SetText 5, i, rstStockLedger.Fields("ItemGroup").Value
            .SetText 6, i, Stock + Val(rstStockLedger.Fields("OPBAL").Value)
            .SetText 7, i, "Units"
            .SetText 8, i, Val(rstStockLedger.Fields("Purchase").Value)
                    PurchaseTotal = PurchaseTotal + Val(rstStockLedger.Fields("Purchase").Value)
            .SetText 9, i, Val(rstStockLedger.Fields("PurchaseReturn").Value)
                    PurchaseReturnTotal = PurchaseReturnTotal + Val(rstStockLedger.Fields("PurchaseReturn").Value)
            .SetText 10, i, Val(rstStockLedger.Fields("PurchaseChallan").Value)
                    PurchaseChallanTotal = PurchaseChallanTotal + Val(rstStockLedger.Fields("PurchaseChallan").Value)
            .SetText 11, i, Val(rstStockLedger.Fields("PurchaseReturnChallan").Value)
                    PurchaseReturnChallanTotal = PurchaseReturnChallanTotal + Val(rstStockLedger.Fields("PurchaseReturnChallan").Value)
            .SetText 12, i, Val(rstStockLedger.Fields("Sales").Value)
                    SalesTotal = SalesTotal + Val(rstStockLedger.Fields("Sales").Value)
            .SetText 13, i, Val(rstStockLedger.Fields("SalesReturn").Value)
                    SalesReturnTotal = SalesReturnTotal + Val(rstStockLedger.Fields("SalesReturn").Value)
            .SetText 14, i, Val(rstStockLedger.Fields("SalesChallan").Value)
                    SalesChallanTotal = SalesChallanTotal + Val(rstStockLedger.Fields("SalesChallan").Value)
            .SetText 15, i, Val(rstStockLedger.Fields("SalesReturnChallan").Value)
                    SalesReturnChallanTotal = SalesReturnChallanTotal + Val(rstStockLedger.Fields("SalesReturnChallan").Value)
            .SetText 16, i, Val(rstStockLedger.Fields("StockJournalIN").Value)
                    StockJournalINTotal = StockJournalINTotal + Val(rstStockLedger.Fields("StockJournalIN").Value)
            .SetText 17, i, Val(rstStockLedger.Fields("StockJournalOUT").Value)
                    StockJournalOUTTotal = StockJournalOUTTotal + Val(rstStockLedger.Fields("StockJournalOUT").Value)
            .SetText 18, i, Val(rstStockLedger.Fields("PendingPO").Value)
                    POTotal = POTotal + Val(rstStockLedger.Fields("PendingPO").Value)
            .SetText 19, i, Val(rstStockLedger.Fields("PendingSO").Value)
                    SOTotal = SOTotal + Val(rstStockLedger.Fields("PendingSO").Value)
            .SetText 20, i, EffectiveStock
            .SetText 21, i, Val(rstStockLedger.Fields("MRP").Value)
            .SetText 22, i, EffectiveStock * Val(rstStockLedger.Fields("MRP").Value)
                    AmountTotal = AmountTotal + EffectiveStock * Val(rstStockLedger.Fields("MRP").Value)
            .SetText 23, i, Val(rstStockLedger.Fields("Purchase").Value) - Val(rstStockLedger.Fields("PurchaseReturn").Value)
                    NetPurchaseTotal = NetPurchaseTotal + Val(rstStockLedger.Fields("Purchase").Value) - Val(rstStockLedger.Fields("PurchaseReturn").Value)
            .SetText 24, i, Val(rstStockLedger.Fields("Sales").Value) - Val(rstStockLedger.Fields("SalesReturn").Value)
                    NetSalesTotal = NetSalesTotal + Val(rstStockLedger.Fields("Sales").Value) - Val(rstStockLedger.Fields("SalesReturn").Value)
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstStockLedger.Fields("PurchaseAmount").Value)
                    PurchaseAmountTotal = PurchaseAmountTotal + Val(rstStockLedger.Fields("PurchaseAmount").Value)
            .SetText 27, i, Val(rstStockLedger.Fields("SalesAmount").Value)
                    SalesAmountTotal = SalesAmountTotal + Val(rstStockLedger.Fields("SalesAmount").Value)
            .SetText 28, i, Val(rstStockLedger.Fields("PurchaseReturnAmount").Value)
                    PurchaseReturnAmountTotal = PurchaseReturnAmountTotal + Val(rstStockLedger.Fields("PurchaseReturnAmount").Value)
            .SetText 29, i, Val(rstStockLedger.Fields("SalesReturnAmount").Value)
                    SalesReturnAmountTotal = SalesReturnAmountTotal + Val(rstStockLedger.Fields("SalesReturnAmount").Value)
            .SetText 30, i, Val(rstStockLedger.Fields("PurchaseAmount").Value) - Val(rstStockLedger.Fields("PurchaseReturnAmount").Value)
                    NetPurchaseAmountTotal = NetPurchaseAmountTotal + Val(rstStockLedger.Fields("PurchaseAmount").Value) - Val(rstStockLedger.Fields("PurchaseReturnAmount").Value)
            .SetText 31, i, Val(rstStockLedger.Fields("SalesAmount").Value) - Val(rstStockLedger.Fields("SalesReturnAmount").Value)
            .SetText 32, i, (rstStockLedger.Fields("Code").Value)
            .SetText 33, i, 0
            .SetText 34, i, 0
            .SetText 35, i, rstStockLedger.Fields("HSNCode").Value
            NetSalesAmountTotal = NetSalesAmountTotal + Val(rstStockLedger.Fields("SalesAmount").Value) - Val(rstStockLedger.Fields("SalesReturnAmount").Value)
            StockTotal = StockTotal + Stock
            EStockTotal = EStockTotal + EffectiveStock
            dPrint = dPrint + 1
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
        End If
NXT:
            rstStockLedger.MoveNext
            If MdiMainMenu.ProgressBar1.Value + Round((100 / rstStockLedger.RecordCount), 2) <= 100 Then
                MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + Round((100 / rstStockLedger.RecordCount), 2)
            End If
        Loop
    End If
End With
End If
    

With fpSpread1
If VSFlexFlag = False Then
    If VchType < 34 Or Right(VchType, 2) = 48 Then
        R = i + 1
        For C = 1 To .MaxCols
            .Col = C: .Row = R: .FontBold = True: .FontSize = 12.5: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue:
        Next
    End If
    End If
        If VchType <= 2 Or VchType = 33 Then
            .LockBackColor = RGB(255, 255, 240): Combo1.BackColor = RGB(255, 255, 240): Combo2.BackColor = RGB(255, 255, 240): MhDateInput1.BackColor = RGB(255, 255, 240): MhDateInput2.BackColor = RGB(255, 255, 240): TDBNumber1.BackColor = RGB(255, 255, 240): TDBNumber2.BackColor = RGB(255, 255, 240): Text1.BackColor = RGB(255, 255, 240):
        ElseIf (VchType >= 3 And VchType <= 6) Or (VchType >= 53 And VchType <= 56) Then
            .LockBackColor = RGB(245, 255, 230): Combo1.BackColor = RGB(245, 255, 230): Combo2.BackColor = RGB(245, 255, 230): MhDateInput1.BackColor = RGB(245, 255, 230): MhDateInput2.BackColor = RGB(245, 255, 230): TDBNumber1.BackColor = RGB(245, 255, 230): TDBNumber2.BackColor = RGB(245, 255, 230): Text1.BackColor = RGB(245, 255, 230):
        ElseIf (VchType >= 7 And VchType <= 10) Or (VchType >= 57 And VchType <= 60) Then
            .LockBackColor = RGB(245, 250, 250): Combo1.BackColor = RGB(245, 250, 250): Combo2.BackColor = RGB(245, 250, 250): MhDateInput1.BackColor = RGB(245, 250, 250): MhDateInput2.BackColor = RGB(245, 250, 250): TDBNumber1.BackColor = RGB(245, 250, 250): TDBNumber2.BackColor = RGB(245, 250, 250): Text1.BackColor = RGB(245, 250, 250):
        ElseIf (VchType >= 21 And VchType <= 24) Or (VchType >= 61 And VchType <= 64) Then
            .LockBackColor = RGB(255, 250, 255): Combo1.BackColor = RGB(255, 250, 255): Combo2.BackColor = RGB(255, 250, 255): MhDateInput1.BackColor = RGB(255, 250, 255): MhDateInput2.BackColor = RGB(255, 250, 255): TDBNumber1.BackColor = RGB(255, 250, 255): TDBNumber2.BackColor = RGB(255, 250, 255): Text1.BackColor = RGB(255, 250, 255):
        ElseIf (VchType >= 25 And VchType <= 30) Or (VchType >= 65 And VchType <= 68) Then
            .LockBackColor = RGB(240, 255, 255): Combo1.BackColor = RGB(240, 255, 255): Combo2.BackColor = RGB(240, 255, 255): MhDateInput1.BackColor = RGB(240, 255, 255): MhDateInput2.BackColor = RGB(240, 255, 255): TDBNumber1.BackColor = RGB(240, 255, 255): TDBNumber2.BackColor = RGB(240, 255, 255): Text1.BackColor = RGB(240, 255, 255):
        End If
If VSFlexFlag = False Then
        .SelectBlockOptions = SelectBlockOptionsAll: .AllowMultiBlocks = True: If TDBNumber2 <> 0 Then fpSpread1.SetFocus: fpSpread1.SetActiveCell 3, LR 'i + 1
        If VchType < 34 Or Right(VchType, 2) = 48 Then TDBNumber2 = i: fpSpread1.MaxRows = IIf(i < 27, 27, i + 1): Call cmdFilter_Click Else TDBNumber2 = fpSpread1.DataRowCnt: fpSpread1.MaxRows = IIf(fpSpread1.DataRowCnt < 27, 27, fpSpread1.DataRowCnt + 1)
        If VchType >= 53 And VchType <= 68 Then TDBNumber2 = i: fpSpread1.MaxRows = IIf(i < 27, 27, i + 1): Call cmdFilter_Click
        If VchType = 46 Or VchType = 47 Then Call cmdFilter_Click
        'Item Ledger
        If VchType = 31 Then fpSpread1.SetText 24, i + 1, Bal: fpSpread1.GetText 26, i + 1, Bal: fpSpread1.SetText 26, i + 1, Bal / i: fpSpread1.GetText 27, i, Bal: fpSpread1.SetText 27, i + 1, Bal:
        If VchType = 32 Then
                Mh3dLabel11.Caption = "Material Centre : " + "All"
            If rstItemOpening.RecordCount <> 0 Then rstItemOpening.MoveFirst
            fpSpread1.GetText 6, i + 1, Bal: Bal = Format(Bal, "##,##,##,##0.00"): Mh3dLabel10.Caption = "Opening Balance :  " & Format(Bal, "##,##,##,##0.00") & IIf(Opening <= 0, " Units", " Units")
        End If
End If
End With
    
    Timer1.Enabled = False
    ShowProgressInStatusBar False
    MdiMainMenu.MousePointer = vbNormal
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHandler:
    Timer1.Enabled = False
    ShowProgressInStatusBar False
    MdiMainMenu.MousePointer = vbNormal
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
Private Function Print_fpSpread()
Dim i As Long, dPrint As Long
OrderPGTF = 0: INWardPGTF = 0: OUTWardPGTF = 0: AmountPGTF = 0
OrderGTF = 0: INWardGTF = 0: OUTWardGTF = 0: AmountGTF = 0
PartyH = "": OrderH = "": ItemH = "": INWardF = 0: OUTWardF = 0: SNo = 0: aSNO = 0: pSNO = 0: OrderF = 0: Bal = 0: AmountF = 0
    With fpSpread1
    .RowHeadersAutoText = DispBlank
        rstStockLedger.MoveFirst
        Do While Not rstStockLedger.EOF
        If VchType = 34 Or VchType = 35 Or VchType = 37 Or VchType = 45 Then
            i = i + 1
            If PartyH <> rstStockLedger.Fields("AccountName").Value Then
                aSNO = aSNO + 1
                .SetText 0, i, "A/C-" & aSNO
'Party Header
                .SetText 5, i, "Party : " + rstStockLedger.Fields("AccountName").Value: .Col = 5: .Row = i: .FontBold = True: .FontSize = 12: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbRed: pSNO = 0
                fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstStockLedger.Fields("AccountName").Value) > 33 Then fpSpread1.RowHeight(i) = 36: fpSpread1.TypeHAlign = TypeHAlignRight
                PartyH = rstStockLedger.Fields("AccountName").Value
                If i > 2 Then i = i + 1
            End If
            If OrderH <> rstStockLedger.Fields("VchBillNo").Value And (VchType = 34 Or VchType = 35 Or VchType = 37 Or VchType = 45) And rstStockLedger.Fields("VchBillNo").Value <> "" Then
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
                If VchType = 34 And rstStockLedger.Fields("pvtType").Value = "FP" Then VchCode = "S": Bal = Val(rstStockLedger.Fields("Ordered").Value) * IIf(VchCode = "S", -1, 1)
                If VchType = 34 And rstStockLedger.Fields("pvtType").Value = "FS" Then VchCode = "P": Bal = Val(rstStockLedger.Fields("Ordered").Value) * IIf(VchCode = "S", -1, 1)
                If VchType = 45 And Right(rstStockLedger.Fields("pvtType").Value, 2) = "PO" Then VchCode = "S": Bal = Val(rstStockLedger.Fields("Ordered").Value) * IIf(VchCode = "S", -1, 1)
                If VchType = 45 And Right(rstStockLedger.Fields("pvtType").Value, 2) = "SO" Then VchCode = "P": Bal = Val(rstStockLedger.Fields("Ordered").Value) * IIf(VchCode = "S", -1, 1)
                If VchType = 35 And rstStockLedger.Fields("pvtType").Value = "FP" Then VchCode = "S": Bal = Val(rstStockLedger.Fields("Ordered").Value) * IIf(VchCode = "S", -1, 1)
                If VchType = 37 And rstStockLedger.Fields("pvtType").Value = "FS" Then VchCode = "P": Bal = Val(rstStockLedger.Fields("Ordered").Value) * IIf(VchCode = "S", -1, 1)
                .SetText 0, i, IIf(VchCode = "S", "P", "S") & "-" & pSNO
                .SetText 1, i, rstStockLedger.Fields("pvtDate").Value: .Col = 1: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue:
                .SetText 2, i, rstStockLedger.Fields("VchBillNo").Value: .Col = 2: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue:
                .SetText 5, i, rstStockLedger.Fields("ItemName").Value: .Col = 5: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue:
                .SetText 6, i, Val(rstStockLedger.Fields("Ordered").Value): .Col = 6: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue:
                .SetText 32, i, Trim(rstStockLedger.Fields("pvtCode").Value)
                .SetText 35, i, Trim(rstStockLedger.Fields("pvtType").Value)
                OrderF = Val(rstStockLedger.Fields("Ordered").Value)
                OrderH = rstStockLedger.Fields("VchBillNo").Value
                If VchType <> 45 Then i = i + 1
            End If
        ElseIf VchType = 36 Or VchType = 38 Then
                    i = i + 1
            If PartyH <> rstStockLedger.Fields("AccountName").Value Then
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
                .SetText 5, i, "Party : " + rstStockLedger.Fields("AccountName").Value: .Col = 5: .Row = i: .FontBold = True: .FontSize = 14: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbRed:
                fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstStockLedger.Fields("AccountName").Value) > 33 Then fpSpread1.RowHeight(i) = 36: fpSpread1.TypeHAlign = TypeHAlignRight
                PartyH = rstStockLedger.Fields("AccountName").Value
                i = i + 1
            End If
        ElseIf VchType = 40 Or VchType = 43 Then
                    i = i + 1
            If PartyH <> rstStockLedger.Fields("AccountName").Value Then
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
                .SetText 5, i, "Party : " + rstStockLedger.Fields("AccountName").Value: .Col = 5: .Row = i: .FontBold = True: .FontSize = 14: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbRed:
                fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstStockLedger.Fields("AccountName").Value) > 33 Then fpSpread1.RowHeight(i) = 36: fpSpread1.TypeHAlign = TypeHAlignRight
                PartyH = rstStockLedger.Fields("AccountName").Value
                i = i + 1
            End If
        ElseIf VchType = 41 Or VchType = 44 Then
                    i = i + 1
            If ItemH <> rstStockLedger.Fields("ItemName").Value Then
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
                .SetText 5, i, "Item : " + rstStockLedger.Fields("ItemName").Value: .Col = 5: .Row = i: .FontBold = True: .FontSize = 14: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbRed:
                fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstStockLedger.Fields("ItemName").Value) > 33 Then fpSpread1.RowHeight(i) = 36: fpSpread1.TypeHAlign = TypeHAlignRight
                ItemH = rstStockLedger.Fields("ItemName").Value
                i = i + 1
            End If
        End If
        
'Pending Order
        If VchType = 34 Or VchType = 35 Or VchType = 37 Or VchType = 45 Then
        If VchType = 34 Or VchType = 35 Or VchType = 37 Or VchType = 45 And rstStockLedger.Fields("VchBillNo").Value = "" Then
            SNo = SNo + 1
            .SetText 0, i, SNo
            .SetText 1, i, rstStockLedger.Fields("vtDate").Value
            .SetText 2, i, Trim(rstStockLedger.Fields("vtNo").Value)
            .SetText 3, i, rstStockLedger.Fields("TypeRef").Value: fpSpread1.Col = 3: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstStockLedger.Fields("TypeRef").Value) > 10 Then fpSpread1.RowHeight(i) = 25.5:
            .SetText 5, i, rstStockLedger.Fields("MaterialCentre").Value & IIf(rstStockLedger.Fields("Remarks") <> "" Or rstStockLedger.Fields("ChallanNo") <> "", " ->> ", "") & IIf(rstStockLedger.Fields("Remarks") <> "", " RemarK : " & rstStockLedger.Fields("Remarks"), "") & IIf(rstStockLedger.Fields("ChallanNo") <> "", " (Ch.No." + rstStockLedger.Fields("ChallanNo") & "_ Ch. dt." & rstStockLedger.Fields("ChallanDate") & ")", ""): fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstStockLedger.Fields("MaterialCentre").Value & IIf(rstStockLedger.Fields("RemarkS") <> "", " -> RemarK : " & rstStockLedger.Fields("Remarks"), "") & IIf(rstStockLedger.Fields("ChallanNo") <> "", " (Ch.No." + rstStockLedger.Fields("ChallanNo") & "_ Ch. dt." & rstStockLedger.Fields("ChallanDate") & ")", "")) > 75 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignRight
                Credit = Val(rstStockLedger.Fields("INward").Value)
                INWardF = INWardF + Credit
            .SetText 8, i, Val(rstStockLedger.Fields("INward").Value)
                Debit = Val(rstStockLedger.Fields("OutWard").Value)
                OUTWardF = OUTWardF + Debit
            .SetText 23, i, Val(rstStockLedger.Fields("OutWard").Value)
                Bal = Bal + Credit - Debit
            .SetText 24, i, Bal
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstStockLedger.Fields("Rate").Value)
                AmountF = AmountF + Val(rstStockLedger.Fields("Amount").Value)
            .SetText 27, i, Val(rstStockLedger.Fields("Amount").Value)
            .SetText 32, i, Trim(rstStockLedger.Fields("vtCode").Value)
            .SetText 35, i, rstStockLedger.Fields("vtType").Value
            dPrint = dPrint + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
        End If
        ElseIf VchType = 36 Or VchType = 38 Then
            SNo = SNo + 1
            .SetText 0, i, SNo
            .SetText 1, i, rstStockLedger.Fields("VchDate").Value
            .SetText 2, i, rstStockLedger.Fields("VchBillNo").Value
            .SetText 3, i, IIf(VchCode = "S", "Purchase Order", "Sales Order")
            .SetText 5, i, rstStockLedger.Fields("ItemName").Value: fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstStockLedger.Fields("ItemName").Value) > 48 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignRight
                OrderF = OrderF + Val(rstStockLedger.Fields("Ordered").Value)
            .SetText 6, i, Val(rstStockLedger.Fields("Ordered").Value)
                Credit = Val(rstStockLedger.Fields("INward").Value)
                INWardF = INWardF + Credit
            .SetText 8, i, Val(rstStockLedger.Fields("INward").Value)
                Debit = Val(rstStockLedger.Fields("OutWard").Value)
                OUTWardF = OUTWardF + Debit
            .SetText 23, i, Val(rstStockLedger.Fields("OutWard").Value)
            .SetText 24, i, (IIf(VchCode = "S", -1, 1) * Val(rstStockLedger.Fields("Ordered").Value)) - Val(rstStockLedger.Fields("OutWard").Value) + Val(rstStockLedger.Fields("INward").Value)
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstStockLedger.Fields("Rate").Value)
                AmountF = AmountF + Val(rstStockLedger.Fields("Amount").Value)
            .SetText 27, i, Val(rstStockLedger.Fields("Amount").Value)
            .SetText 32, i, rstStockLedger.Fields("vtCode").Value
            .SetText 35, i, rstStockLedger.Fields("vtType").Value
            dPrint = dPrint + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
        ElseIf VchType = 39 Or VchType = 42 Then
            SNo = SNo + 1
            i = i + 1
            .SetText 0, i, SNo
            .SetText 1, i, rstStockLedger.Fields("VchDate").Value
            .SetText 2, i, rstStockLedger.Fields("VchBillNo").Value
            .SetText 3, i, IIf(VchCode = "S", "Purchase Order", "Sales Order")
            .SetText 5, i, rstStockLedger.Fields("ItemName").Value: fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstStockLedger.Fields("ItemName").Value) > 75 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignLeft
                OrderF = OrderF + Val(rstStockLedger.Fields("Ordered").Value)
            .SetText 6, i, Val(rstStockLedger.Fields("Ordered").Value)
                Credit = Val(rstStockLedger.Fields("Dispatched").Value)
                INWardF = INWardF + Credit
            .SetText 8, i, Val(rstStockLedger.Fields("Dispatched").Value)
                Debit = Val(rstStockLedger.Fields("Dispatched").Value)
                OUTWardF = OUTWardF + Debit
            .SetText 23, i, Val(rstStockLedger.Fields("Dispatched").Value)
            .SetText 24, i, Val(rstStockLedger.Fields("Balance").Value)
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstStockLedger.Fields("Rate").Value)
                AmountF = AmountF + Val(rstStockLedger.Fields("Amount").Value)
            .SetText 27, i, Val(rstStockLedger.Fields("Amount").Value)
            .SetText 32, i, rstStockLedger.Fields("vtCode").Value
            .SetText 35, i, rstStockLedger.Fields("iCode").Value
            dPrint = dPrint + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
        ElseIf VchType = 40 Or VchType = 43 Then
            SNo = SNo + 1
            .SetText 0, i, SNo
            .SetText 1, i, rstStockLedger.Fields("VchDate").Value
            .SetText 2, i, rstStockLedger.Fields("VchBillNo").Value
            .SetText 3, i, IIf(VchCode = "S", "Purchase Order", "Sales Order")
            .SetText 5, i, rstStockLedger.Fields("ItemName").Value: fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstStockLedger.Fields("ItemName").Value) > 48 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignRight
                OrderF = OrderF + Val(rstStockLedger.Fields("Ordered").Value)
            .SetText 6, i, Val(rstStockLedger.Fields("Ordered").Value)
                Credit = Val(rstStockLedger.Fields("Dispatched").Value)
                INWardF = INWardF + Credit
            .SetText 8, i, Val(rstStockLedger.Fields("Dispatched").Value)
                Debit = Val(rstStockLedger.Fields("Dispatched").Value)
                OUTWardF = OUTWardF + Debit
            .SetText 23, i, Val(rstStockLedger.Fields("Dispatched").Value)
                Bal = Bal + Val(rstStockLedger.Fields("Balance").Value)
            .SetText 24, i, Val(rstStockLedger.Fields("Balance").Value)
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstStockLedger.Fields("Rate").Value)
                AmountF = AmountF + Val(rstStockLedger.Fields("Amount").Value)
            .SetText 27, i, Val(rstStockLedger.Fields("Amount").Value)
            .SetText 32, i, rstStockLedger.Fields("vtCode").Value
            dPrint = dPrint + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
        ElseIf VchType = 41 Or VchType = 44 Then
            SNo = SNo + 1
            .SetText 0, i, SNo
            .SetText 1, i, rstStockLedger.Fields("VchDate").Value
            .SetText 2, i, rstStockLedger.Fields("VchBillNo").Value
            .SetText 3, i, IIf(VchCode = "S", "Purchase Order", "Sales Order")
            .SetText 5, i, rstStockLedger.Fields("AccountName").Value: fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstStockLedger.Fields("ItemName").Value) > 48 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignRight
                OrderF = OrderF + Val(rstStockLedger.Fields("Ordered").Value)
            .SetText 6, i, Val(rstStockLedger.Fields("Ordered").Value)
                Credit = Val(rstStockLedger.Fields("Dispatched").Value)
                INWardF = INWardF + Credit
            .SetText 8, i, Val(rstStockLedger.Fields("Dispatched").Value)
                Debit = Val(rstStockLedger.Fields("Dispatched").Value)
                OUTWardF = OUTWardF + Debit
            .SetText 23, i, Val(rstStockLedger.Fields("Dispatched").Value)
                Bal = Bal + Val(rstStockLedger.Fields("Balance").Value)
            .SetText 24, i, Val(rstStockLedger.Fields("Balance").Value)
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstStockLedger.Fields("Rate").Value)
                AmountF = AmountF + Val(rstStockLedger.Fields("Amount").Value)
            .SetText 27, i, Val(rstStockLedger.Fields("Amount").Value)
            .SetText 32, i, rstStockLedger.Fields("vtCode").Value
            dPrint = dPrint + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
'Item Ledger Summarize
        ElseIf VchType = 49 Then
        i = i + 1
            .SetText 3, i, rstStockLedger.Fields("MonthYear").Value 'Month
        If i = 1 Then
            .SetText 6, i, Val(rstItemOpening.Fields("Opening").Value)
            Opening = Val(rstItemOpening.Fields("Opening").Value)
        Else
            .SetText 6, i, Opening + (INWardF - OUTWardF)
        End If
            .SetText 20, i, Val(rstStockLedger.Fields("INWard").Value)
                Credit = Val(rstStockLedger.Fields("INWard").Value)
                INWardF = INWardF + Credit
            .SetText 23, i, Val(rstStockLedger.Fields("OutWard").Value)
                Debit = Val(rstStockLedger.Fields("OutWard").Value)
                OUTWardF = OUTWardF + Debit
                Bal = Opening + (INWardF - OUTWardF)
            .SetText 24, i, Bal
            .SetText 25, i, "Units"
            .SetText 32, i, rstStockLedger.Fields("FromDate").Value
            .SetText 35, i, rstStockLedger.Fields("ToDate").Value
            dPrint = dPrint + 1
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
        ElseIf VchType = 101 Or VchType = 102 Or VchType = 103 Or VchType = 104 Or VchType = 105 Then
            SNo = SNo + 1
            i = i + 1
            .SetText 0, i, SNo
            .SetText 1, i, rstStockLedger.Fields("Item").Value
            .SetText 2, i, rstStockLedger.Fields("VchBillNo").Value
            .SetText 3, i, IIf(VchCode = "S", "Purchase Order", "Sales Order")
            .SetText 5, i, rstStockLedger.Fields("AccountName").Value: fpSpread1.Col = 5: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstStockLedger.Fields("ItemName").Value) > 48 Then fpSpread1.RowHeight(i) = 25.5: fpSpread1.TypeHAlign = TypeHAlignRight
                OrderF = OrderF + Val(rstStockLedger.Fields("Ordered").Value)
            .SetText 6, i, Val(rstStockLedger.Fields("Ordered").Value)
                Credit = Val(rstStockLedger.Fields("Dispatched").Value)
                INWardF = INWardF + Credit
            .SetText 8, i, Val(rstStockLedger.Fields("Dispatched").Value)
                Debit = Val(rstStockLedger.Fields("Dispatched").Value)
                OUTWardF = OUTWardF + Debit
            .SetText 23, i, Val(rstStockLedger.Fields("Dispatched").Value)
                Bal = Bal + Val(rstStockLedger.Fields("Balance").Value)
            .SetText 24, i, Val(rstStockLedger.Fields("Balance").Value)
            .SetText 25, i, "Units"
            .SetText 26, i, Val(rstStockLedger.Fields("Rate").Value)
                AmountF = AmountF + Val(rstStockLedger.Fields("Amount").Value)
            .SetText 27, i, Val(rstStockLedger.Fields("Amount").Value)
            .SetText 32, i, rstStockLedger.Fields("vtCode").Value
            dPrint = dPrint + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
        End If
NXT:
            rstStockLedger.MoveNext
            If MdiMainMenu.ProgressBar1.Value + Round((100 / rstStockLedger.RecordCount), 2) <= 100 Then
                MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + Round((100 / rstStockLedger.RecordCount), 2)
            End If
        Loop
        If i > 2 Or VchType = 49 Then
            i = i + 1: .SetText 0, i, " ": .SetText 0, i + 1, " ": .SetText 0, i + 2, " "
            If VchType = 39 Or VchType = 42 Then .SetText 5, i, "TOTAL" Else .SetText 5, i, "SUBTOTAL"
            If VchType = 49 Then
                .SetText 3, i, ""
                 .SetText 20, i, INWardF: .SetText 23, i, OUTWardF: .SetText 24, i, Bal: .SetText 25, i, "Units"
                 Mh3dLabel10.Caption = "Closing Balance = " & Bal & " Units ": Mh3dLabel10.Visible = True
            Else
                .SetText 6, i, OrderF: .SetText 8, i, INWardF: .SetText 23, i, OUTWardF: .SetText 24, i, Bal: .SetText 25, i, "Units": .SetText 27, i, AmountF: If VchType = 36 Or VchType = 38 Then .SetText 24, i, (IIf(VchCode = "S", -1, 1) * OrderF) - OUTWardF + INWardF
            End If

                .Col = 5: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True: .TypeHAlign = TypeHAlignRight
                .Col = 6: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                .Col = 8: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                .Col = 20: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                .Col = 23: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                .Col = 24: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                .Col = 25: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
                .Col = 27: .Row = i: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F:  .ForeColor = RGB(128, 0, 64): .TypeVAlign = TypeVAlignTop: .FontUnderline = True
       End If
If VchType <> 49 Then
        INWardGTF = INWardGTF + INWardF: INWardF = 0: OUTWardGTF = OUTWardGTF + OUTWardF: OUTWardF = 0: OrderGTF = OrderGTF + OrderF: OrderF = 0: AmountGTF = AmountGTF + AmountF: AmountF = 0:
         .SetText 5, i + 1, "Grand TOTAL": .SetText 6, i + 1, OrderGTF: .SetText 8, i + 1, INWardGTF: .SetText 23, i + 1, OUTWardGTF: .SetText 24, i + 1, (IIf(VchCode = "S", -1, 1) * OrderGTF) - OUTWardGTF + INWardGTF: .SetText 25, i + 1, "Units": .SetText 27, i + 1, AmountGTF: If VchType = 36 Or VchType = 38 Then .SetText 24, i, (IIf(VchCode = "S", -1, 1) * OrderGTF) - OUTWardGTF + INWardGTF
            .Col = 5: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True: .TypeHAlign = TypeHAlignRight
            .Col = 6: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 8: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 23: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 24: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 25: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
            .Col = 27: .Row = i + 1: .FontBold = True: .FontSize = 11: .BackColor = &H8000000F: .ForeColor = &H808000: .TypeVAlign = TypeVAlignTop: .FontUnderline = True
End If
End With
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Check VchCode Next & Previous
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Right(oVchType, 2) <> Format(VchType, "00") Then
            oVchType = oVchType + Format(VchType, "00")
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        If oVchType <> "" And oVchType <> Format(VchType, "00") Then
            oVchType = Get_oVchType(oVchType): VchType = Right(oVchType, 2)
        ElseIf oVchType <> "" And oVchType = Format(VchType, "00") Then
            oVchType = Get_oVchType(oVchType): VchType = Right(oVchType, 2)
        End If
        If oVchType = "" And Format(VchType, "00") = "" And HideFlag = False And ExitFlag = False Then
            Call cmdCancel_Click: ExitFlag = False: KeyCode = 0: Exit Sub
        End If
    End If

With fpSpread1
        If Shift = 0 And KeyCode = vbKeyReturn And (VchType = 34 Or VchType = 35 Or VchType = 37 Or VchType = 36 Or VchType = 38 Or VchType = 39 Or VchType = 40 Or VchType = 41 Or VchType = 42 Or VchType = 43 Or VchType = 44 Or VchType = 45 Or Right(VchType, 2) = 48) Then .GetText 32, .ActiveRow, SCode: SCode = "'" & SCode & "'": If SCode = "''" Then Exit Sub
        If Shift = 0 And KeyCode = vbKeyReturn And VchType = 46 Then .GetText 35, .ActiveRow, SCode: SCode = "'" & SCode & "'": If SCode = "''" Then Exit Sub
        
        If (Shift = vbCtrlMask And KeyCode <> vbKeyEscape) And ((VchType >= 3 And VchType <= 10) Or (VchType >= 53 And VchType <= 60)) Then .GetText 32, .ActiveRow, SCode: SCode = "'" & SCode & "'": If SCode = "''" Then Exit Sub
        If (Shift = 0 And KeyCode <> vbKeyEscape) And ((VchType >= 3 And VchType <= 10) Or (VchType >= 53 And VchType <= 60)) Then .GetText 32, .ActiveRow, SCode: SCode = "'" & SCode & "'": If SCode = "''" Then Exit Sub
        If (Shift = 0 And KeyCode <> vbKeyEscape) And ((VchType >= 21 And VchType <= 28) Or (VchType >= 61 And VchType <= 68)) Then .GetText 32, .ActiveRow, sMcCode: sMcCode = "'" & sMcCode & "'": If sMcCode = "''" Then Exit Sub

'Specific VchType = 46
        If (Shift = 0 And KeyCode = vbKeyF9) And VchType = 46 Then
            If Check1.Value <> 1 Then .GetText 8, .ActiveRow, SCode
            If SCode = 0 Then ClearFlag = False: MsgBox "You Can't Clear This Order  !!!", vbCritical, "   Order Quantity Clear  !!!": SCode = "": Exit Sub
            If SCode <> 0 Then ClearFlag = True: ClearQty (True): MsgBox " ( " & SCode & " ) Order Quantity Clear  !!!", vbCritical, "   Order Quantity Clear  !!!": SCode = "": ClearFlag = False: Form_Load: Exit Sub
        ElseIf (Shift = 0 And KeyCode = vbKeyF10) And VchType = 46 Then
            .GetText 14, .ActiveRow, SCode
            If SCode = 0 Then unClearFlag = False: MsgBox "Order Quantity Can't Retrieve  !!!", vbCritical, "   Retrieve Pending Order !!!": SCode = "": Exit Sub
            If SCode <> 0 Then unClearFlag = True: ClearQty (True): MsgBox " ( " & SCode & " )  Order Quantity Retrieve !!!", vbCritical, "   Retrieve Pending Order !!!": SCode = "": unClearFlag = False:: Form_Load: Exit Sub
        End If
    
    
    If (Shift = 0 And KeyCode = vbKeyReturn) And ((VchType >= 3 And VchType <= 6) Or (VchType >= 53 And VchType <= 56)) Then
        If VchType = 3 Then oVchType = VchType: VchType = 25 'One Item-Party-wise 'Sales Ok
        If VchType = 4 Then oVchType = VchType: VchType = 26 'One Item-Party-wise 'Sales Returns
        If VchType = 5 Then oVchType = VchType: VchType = 27 'One Item-Party-wise'Sales And Sales Returns
        If VchType = 6 Then oVchType = VchType: VchType = 28 'One Item-Party-wise'Net Sales
        If SCode = "" Then Exit Sub
            Form_Load
            KeyCode = 0
    ElseIf (Shift = 0 And KeyCode = vbKeyReturn) And (VchType = 36 Or VchType = 38 Or VchType = 39 Or VchType = 40 Or VchType = 41 Or VchType = 42 Or VchType = 43 Or VchType = 44) Then
        If SCode = "" Then Exit Sub
        If VchType = 36 Then oVchType = VchType: VchType = 34 'One Item JobWork Voucher-wise
        If VchType = 38 Then oVchType = VchType: VchType = 34 'One Item JobWork Voucher-wise
        
        If VchType = 39 Then oVchType = VchType: VchType = 45 'One Item Voucher-wise
        If VchType = 40 Then oVchType = VchType: VchType = 45 'One Item Voucher-wise
        If VchType = 41 Then oVchType = VchType: VchType = 45 'One Item Voucher-wise
        
        If VchType = 42 Then oVchType = VchType: VchType = 45 'One Item Voucher-wise
        If VchType = 43 Then oVchType = VchType: VchType = 45 'One Item Voucher-wise
        If VchType = 44 Then oVchType = VchType: VchType = 45 'One Item Voucher-wise
            Form_Load
            KeyCode = 0
    ElseIf (Shift = 0 And KeyCode = vbKeyReturn) And ((VchType >= 7 And VchType <= 10) Or (VchType >= 57 And VchType <= 60)) Then
        If SCode = "" Then Exit Sub
        If VchType = 7 Then VchType = 25 'One Party-Item-wise 'Sales Ok
        If VchType = 8 Then VchType = 26 'One Party-Item-wise 'Sales Returns
        If VchType = 9 Then VchType = 27 'One Party-Item-wise 'Sales And 'Sales Returns
        If VchType = 10 Then VchType = 28 'One Party-Item-wise 'Net Sales
            sMcCode = "'"
            Form_Load
            KeyCode = 0
    ElseIf (Shift = 0 And KeyCode = vbKeyReturn) And ((VchType >= 21 And VchType <= 24) Or (VchType >= 61 And VchType <= 64)) Then
        If VchType = 21 Then VchType = 7 'One Item-Party-wise 'Sales
        If VchType = 22 Then VchType = 8 'One Item-Party-wise 'Sales Return
        If VchType = 23 Then VchType = 9 'One Item-Party-wise 'Sales And Sales Return
        If VchType = 24 Then VchType = 10 'One Item-Party-wise 'Net Sales
        If SCode = "" Then Exit Sub
            Form_Load
            KeyCode = 0
    ElseIf (Shift = 0 And KeyCode = vbKeyReturn) And ((VchType >= 25 And VchType <= 28) Or (VchType >= 65 And VchType <= 68)) Then
        If sMcCode = "" Then Exit Sub
        If VchType = 25 Then VchType = 7 'One Item-Party-wise 'Sales Ok
        If VchType = 26 Then VchType = 8 'One Item-Party-wise 'Sales Return
        If VchType = 27 Then VchType = 9 'One Item-Party-wise 'Sales And Sales Return
        If VchType = 28 Then VchType = 10 'One Item-Party-wise 'Net Sales
            SCode = ""
            Form_Load
            KeyCode = 0


'vbKeyEscape
    ElseIf (Shift = 0 And KeyCode = vbKeyEscape) And (VchType = 34 Or VchType = 45) And SCode <> "" Then
        If VchType = 34 Then VchType = oVchType: If oVchType = 30 Then SCode = oSCode 'Party-wise Order Status Sumarized
        If VchType = 45 Then VchType = oVchType: SCode = oSCode
        If SCode = "" Then Exit Sub
        If oVchType <> 30 Then sMcCode = "'": SCode = ""
           oVchType = ""
            Form_Load
            KeyCode = 0
    ElseIf (Shift = 0 And KeyCode = vbKeyEscape) And ((VchType >= 3 And VchType <= 6) Or (VchType >= 53 And VchType <= 56)) And SCode <> "" Then
            Call cmdCancel_Click: ExitFlag = False
    ElseIf (Shift = 0 And KeyCode = vbKeyEscape) And ((VchType >= 7 And VchType <= 10) Or (VchType >= 57 And VchType <= 60)) And SCode <> "" Then
        If VchType = 7 Then VchType = 21 'Party-wise 'Sales OK
        If VchType = 8 Then VchType = 22 'Party-wise 'Sales Return
        If VchType = 9 Then VchType = 23 'Party-wise'Sales And Sales Return
        If VchType = 10 Then VchType = 24 'Party-wise 'Net Sales
        If SCode = "" Then Exit Sub
            sMcCode = "'": SCode = ""
            Form_Load
            KeyCode = 0
    ElseIf (Shift = 0 And KeyCode = vbKeyEscape) And ((VchType >= 25 And VchType <= 28) Or (VchType >= 65 And VchType <= 68)) And sMcCode <> "" Then
        If VchType = 25 Then VchType = 3 'Item-Wise 'Sales
        If VchType = 25 Then VchType = 7 'One Party-Item-wise 'Sales
        If VchType = 26 Then VchType = 4 'Item-wise 'Sales Return
        If VchType = 26 Then VchType = 8 'Item-wise 'Sales Return
        If VchType = 27 Then VchType = 5 'Item-wise 'Sales Return
        If VchType = 27 Then VchType = 9 'One Item-Party-wise 'Sales
        If VchType = 28 Then VchType = 6 'One Item-Party-wise 'Sales
        SCode = ""
        sMcCode = ""
        Form_Load
        KeyCode = 0
    
'vbKeyReturn
    ElseIf (Shift = 0 And KeyCode = vbKeyReturn) And (VchType = 2 Or VchType = 1 Or VchType = 30 Or VchType = 32 Or VchType = 33 Or VchType = 49) Then   'One-Item Pending
        
        LR = fpSpread1.ActiveRow
        If (VchType = 1 Or VchType = 2) And fpSpread1.ActiveCol = 6 Or fpSpread1.ActiveCol = 18 Or fpSpread1.ActiveCol = 19 Or fpSpread1.ActiveCol = 24 Then
            If VchType = 30 Then fpSpread1.GetText 35, fpSpread1.ActiveRow, vtType: vtType = Right(vtType, 2): 'fpSpread1.GetText 32, fpSpread1.ActiveRow, SCode: SCode = "'" & SCode & "'": If SCode = "''" Then Exit Sub
            If fpSpread1.ActiveCol = 18 Then vTypeCode = "18": VchCode = "S"
            If fpSpread1.ActiveCol = 19 Then vTypeCode = "19": VchCode = "P"
            If (VchType = 1 Or VchType = 2) Then fpSpread1.GetText 32, fpSpread1.ActiveRow, SCode: SCode = "'" & SCode & "'": If SCode = "''" Then Exit Sub
            If VchType = 2 And fpSpread1.ActiveCol = 6 Then oVchType = VchType: VchType = 32: Text1.Text = "" 'Item Ledger Material Center-Wise
            If VchType = 1 And (fpSpread1.ActiveCol = 18 Or fpSpread1.ActiveCol = 19) Then VchType = 29 'Pending Order
            If (VchType = 2 Or VchType = 33) And (fpSpread1.ActiveCol = 18 Or fpSpread1.ActiveCol = 19) Then oVchType = VchType: VchType = 30 'Pending Order
            If VchType = 30 And (fpSpread1.ActiveCol = 24) And (vtType = "FP" Or vtType = "FS") Then oVchType = VchType: oSCode = SCode: VchType = 34: VchCode = Right(vtType, 1): VchCode = IIf(VchCode = "P", "S", "P"): fpSpread1.GetText 32, fpSpread1.ActiveRow, SCode: SCode = "'" & SCode & "'": If SCode = "''" Then Exit Sub 'Pending Order
            If VchType = 30 And (fpSpread1.ActiveCol = 24) And (vtType = "PO" Or vtType = "SO") Then oVchType = VchType: oSCode = SCode: VchType = 45: fpSpread1.GetText 32, fpSpread1.ActiveRow, SCode: SCode = "'" & SCode & "'": If SCode = "''" Then Exit Sub 'Pending Order
            Form_Load
        ElseIf VchType = 32 Then
            If VchType = 32 Then oVchType = oVchType + Format(VchType, "00"): VchType = 31 'Item Ledger Material Centre-Wise
                oSCode = SCode
                fpSpread1.GetText 35, fpSpread1.ActiveRow, sMcCode: sMcCode = "'" & sMcCode & "'": If sMcCode = "''" Then Exit Sub
                Form_Load
        ElseIf VchType = 31 Then
                VchType = 31 'One-Item-Ledger
                oVchType = oVchType + Format(VchType, "00")
                oSCode = SCode
                fpSpread1.GetText 32, fpSpread1.ActiveRow, vDate: sDate = Format(vDate, "dd-MM-yyyy")
                fpSpread1.GetText 35, fpSpread1.ActiveRow, vDate: eDate = Format(vDate, "dd-MM-yyyy")
                Form_Load
        ElseIf VchType = 49 Then
                VchType = 31 'One-Item-Ledger
                oVchType = oVchType + Format(VchType, "00")
                oSCode = SCode
                fpSpread1.GetText 32, fpSpread1.ActiveRow, vDate: sDate = Format(vDate, "dd-MM-yyyy")
                fpSpread1.GetText 35, fpSpread1.ActiveRow, vDate: eDate = Format(vDate, "dd-MM-yyyy")
                Form_Load
        ElseIf SCode = "" Then
                Exit Sub
        End If
            KeyCode = 0
'Open Transection
    ElseIf ((Shift = 0 And KeyCode = vbKeyReturn) Or (Shift = 0 And KeyCode = vbKeyF8) Or (Shift = 0 And KeyCode = vbKeyF12)) And (VchType = 29 Or VchType = 30 Or Right(VchType, 2) = 48 Or VchType = 31 Or VchType = 32 Or VchType = 34 Or VchType = 35 Or VchType = 36 Or VchType = 37 Or VchType = 38 Or VchType = 45 Or VchType = 46) And SCode <> "" Then     'Open Transection
            fpSpread1.GetText 1, fpSpread1.ActiveRow, vDate: vDate = Format(vDate, "dd-MMM-yyyy"):
            
            If VchType = 46 Then
                SCode = ""
            Else
                If oSCode = "" Then oSCode = SCode
            End If
            
            If VchType = 46 Then
                fpSpread1.GetText 2, fpSpread1.ActiveRow, vtNo: fpSpread1.GetText 35, fpSpread1.ActiveRow, vtCode: fpSpread1.GetText 25, fpSpread1.ActiveRow, vtType: vtType = Right(vtType, 2)
                Else
                fpSpread1.GetText 2, fpSpread1.ActiveRow, vtNo: fpSpread1.GetText 32, fpSpread1.ActiveRow, vtCode: fpSpread1.GetText 35, fpSpread1.ActiveRow, vtType: vtType = Right(vtType, 2)
            End If
            
            If VchType = 34 Or VchType = 35 Or VchType = 37 Or VchType = 36 Or VchType = 38 Or VchType = 45 Then fpSpread1.GetText 32, fpSpread1.ActiveRow, vtCode: fpSpread1.GetText 35, fpSpread1.ActiveRow, vtType: vtType = Right(vtType, 2)
            
            If vDate = "" Then Exit Sub
            If FinancialYearFrom > vDate Or vDate = "" Then
                If MsgBox("You Can't Open Previous Financial Voucher in Current Year,... To Open This Voucher, Please Switch Financial Year ", vbCritical, "   Switch Financial Year !!!") = vbOK Then Exit Sub
                
'Order FG AND Jobwork
            ElseIf vtType = "FP" Or vtType = "FS" Then
            dSortBy = True
                If VchType = 46 Then SCode = ""
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
                    Call cmdRefresh_Click
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
                    Call cmdRefresh_Click
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmSalesOrderVoucher.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                    Call cmdRefresh_Click
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
                    Call cmdRefresh_Click
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmStockJournalVoucher.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                    Call cmdRefresh_Click
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
                    Call cmdRefresh_Click
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmSalesVoucher.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                    Call cmdRefresh_Click
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
                    Call cmdRefresh_Click
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmSalesChallanVoucher.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                    Call cmdRefresh_Click
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
                    Call cmdRefresh_Click
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmItemIssueReceiptVoucher.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                    Call cmdRefresh_Click
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
                    Call cmdRefresh_Click
                ElseIf Shift = 0 And KeyCode = vbKeyF12 Then 'Duplicate
                    frmJobworkBill.SSTab1.Tab = 1
                    If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then Exit Sub
                    Call cmdRefresh_Click
                End If
                
            End If
        KeyCode = 0
        

'Escape
    ElseIf (Shift = 0 And KeyCode = vbKeyEscape) And (VchType = 29 Or VchType = 30 Or VchType = 31 Or VchType = 32 Or VchType = 49) And SCode <> "" Then   'Stock Ledger Closing
        'VchType = Right(oVchType, 2):
        sDate = FrmItemSelectionList.MhDateInput1: eDate = FrmItemSelectionList.MhDateInput2
        
'        If VchType = 29 Then VchType = 1 'Inventry Movement Ledger Closing
'        If VchType = 30 Then VchType = 2 'Stock Ledger Closing
'        If VchType = 32 Then VchType = 2 'Stock Ledger Closing
'        If VchType = 31 Then VchType = Right(oVchType, 2): oVchType = GetOvchType(oVchType): sDate = FrmItemSelectionList.MhDateInput1: eDate = FrmItemSelectionList.MhDateInput2   'VchType = 32 'Stock Ledger Closing
        If VchType = 32 Then SCode = oSCode
        If oVchType <> 49 Then sMcCode = ""
        Form_Load
        'fpSpread1.SetActiveCell vTypeCode, LR
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'Delete
        If VSFlexGrid1.Visible Then
        On Error Resume Next
                R = IIf(VSFlexGrid1.Row + 1 <> LR, VSFlexGrid1.Row + 1, 1)
                LR = R
                VSFlexGrid1.RemoveItem (VSFlexGrid1.Row): LR = VSFlexGrid1.Row
                VSFlexGrid1.Row = R
                    Call VSFlexGrid1_AfterDataRefresh
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyF9 Then
    If VSFlexGrid1.Visible Then
    On Error Resume Next
            R = IIf(VSFlexGrid1.Row + 1 <> LR, VSFlexGrid1.Row + 1, 1)
            LR = R
            VSFlexGrid1.RowHidden(VSFlexGrid1.Row) = True: LR = VSFlexGrid1.Row
            VSFlexGrid1.Row = R
'            Call VSFlexGrid1_AfterDataRefresh
    Else
            R = IIf(.ActiveRow + 1 <> LR, .ActiveRow + 1, 1)
            LR = R
             .Row = .ActiveRow: .RowHidden = True: LR = .Row
            TotalFlag = True: HideFlag = True: If VchType < 35 Then Total_Click
            TotalFlag = False
            .SetActiveCell .ActiveCol, R
    End If
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        If VSFlexFlag = False Then
            If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}": KeyCode = 0
        ElseIf VSFlexFlag = True Then
            If Me.ActiveControl.Name <> "VSFlexGrid1" Then Sendkeys "{TAB}": KeyCode = 0
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then ' Close/Hide Row/Unhide Row
        If HideFlag = True Then
            For R = 1 To .DataRowCnt 'Unhide All
                If HideFlag = True Then .Row = R: .RowHidden = False: .SetText 13, R, "":
            Next
            If VchType < 35 Then Total_Click
            .SetActiveCell .ActiveCol, 1
            HideFlag = False
        ElseIf HideFlag = False And ExitFlag = False Then
            Call cmdCancel_Click: ExitFlag = False
        End If
            KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        Call cmdRefresh_Click
        KeyCode = 0
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
            If Text1.Text = "" Then
                MsgBox "Please Provide Search Input", vbInformation
                Text1.SetFocus
            ElseIf Text1.Text <> "" Then
            Call Command2_Click
            End If
        KeyCode = 0
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyC) Or (Shift = 0 And KeyCode = vbKeyF12) Then
        Call CopyToClipboard
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyV) Or (Shift = 0 And KeyCode = vbKeyF12) Then
        Call PasteFromClipboard
    End If
End With
End Sub
Private Sub Combo1_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Command1_Click()
With fpSpread1
    fpSpread1.DeleteRows .DataRowCnt, 1
    Call cmdRefresh_Click
    fpSpread1.Col = 6: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignRight ' Stock Qty.
    fpSpread1.Col = 7: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignCenter 'Units
    fpSpread1.Col = 33: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignRight 'Physical Stock Quantity
    fpSpread1.Col = 34: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignRight 'Stock Impact
End With
End Sub
Private Sub Check1_Click()
Dim C As Long
    With fpSpread1
If VchType = 46 Then
    Dim i As Integer, CellVal As Variant
    With fpSpread1
        For i = 1 To .DataRowCnt - 1
            .GetText 8, i, CellVal 'Pending
            If Val(CellVal) > 0 Then .SetText 15, i, Check1.Value
        Next
    End With
Else
            If Check1.Value Then
                If VchType = 1 Then  'Stock Ledger
                    .Col = 32: .ColHidden = True
                    .ColWidth(16) = 13.5
                    .ColWidth(17) = 13.5
                    .ColWidth(20) = 13.5
                ElseIf VchType = 3 Or VchType = 7 Or VchType = 21 Or VchType = 25 Or VchType = 53 Or VchType = 57 Or VchType = 61 Or VchType = 65 Then 'Sales,Purchase
                    If VchType = 53 Or VchType = 57 Or VchType = 61 Or VchType = 65 Then
                    .ColWidth(8) = 40
                    .Col = 26: .ColHidden = False 'SalesAmount
                    .ColWidth(26) = 19.875
                    ElseIf VchType = 3 Or VchType = 7 Or VchType = 21 Or VchType = 25 Then
                    .Col = 27: .ColHidden = False 'SalesAmount
                    .ColWidth(27) = 15
                    .ColWidth(12) = 45
                    End If
                    .ColWidth(3) = 50
                    .ColWidth(5) = 30.75
                    .ColWidth(25) = 11
                ElseIf VchType = 4 Or VchType = 8 Or VchType = 22 Or VchType = 26 Or VchType = 54 Or VchType = 58 Or VchType = 62 Or VchType = 66 Then 'Sales & Purchase Return
                    .ColWidth(3) = 50
                    .ColWidth(5) = 30.75
                    If VchType = 4 Or VchType = 8 Or VchType = 22 Or VchType = 26 Then
                    .Col = 29: .ColHidden = False 'SalesReturnAmount
                    .ColWidth(13) = 45
                    .ColWidth(29) = 15
                    ElseIf VchType = 54 Or VchType = 58 Or VchType = 62 Or VchType = 66 Then
                    .Col = 28: .ColHidden = False 'Purchase ReturnAmount
                    .ColWidth(9) = 35.5
                    .ColWidth(28) = 24.5
                    End If
                    .ColWidth(25) = 11
                ElseIf VchType = 5 Or VchType = 9 Or VchType = 23 Or VchType = 27 Or VchType = 55 Or VchType = 59 Or VchType = 63 Or VchType = 67 Then 'Sales And Sales Return
                    .ColWidth(3) = 50
                    If VchType = 55 Or VchType = 59 Or VchType = 63 Or VchType = 67 Then .ColWidth(5) = 19.375 Else .ColWidth(5) = 30.5
                    .ColWidth(8) = 15
                    .ColWidth(9) = 15
                    .ColWidth(12) = 15
                    .ColWidth(13) = 15
                    .ColWidth(25) = 11
                    If VchType = 5 Or VchType = 9 Or VchType = 23 Or VchType = 27 Then
                    .Col = 27: .ColHidden = False: .Col = 29: .ColHidden = False 'SalesAmount 'SalesReturnAmount
                    .ColWidth(27) = 15
                    .ColWidth(29) = 15
                    ElseIf VchType = 55 Or VchType = 59 Or VchType = 63 Or VchType = 67 Then
                    .Col = 26: .ColHidden = False: .Col = 28: .ColHidden = False 'SalesAmount 'SalesReturnAmount
                    .ColWidth(26) = 19.375
                    .ColWidth(28) = 21.75
                    End If
                ElseIf VchType = 6 Or VchType = 10 Or VchType = 24 Or VchType = 28 Or VchType = 56 Or VchType = 60 Or VchType = 64 Or VchType = 68 Then 'Net Sale
                    .ColWidth(3) = 50
                    .ColWidth(5) = 30.75
                    .ColWidth(7) = 11
                    If VchType = 6 Or VchType = 10 Or VchType = 24 Or VchType = 28 Then
                    .Col = 30: .ColHidden = True: .Col = 31: .ColHidden = False 'Net Sales Amount
                    .ColWidth(24) = 45
                    .ColWidth(31) = 15
                    ElseIf VchType = 56 Or VchType = 60 Or VchType = 64 Or VchType = 68 Then
                    .Col = 30: .ColHidden = False: .Col = 31: .ColHidden = True 'Net Purchase Amount
                    .ColWidth(23) = 39.625
                    .ColWidth(30) = 20.375
                    End If
                    .ColWidth(25) = 11
                    
                ElseIf VchType >= 29 And VchType <= 30 Then  'Pending Purchase Order Voucher-wise
                    For C = 1 To 3
                    .Col = C: .ColHidden = False
                    Next
                    .ColWidth(1) = 15
                    .ColWidth(2) = 15
                    .ColWidth(3) = 50
                    For C = 4 To 5
                    .Col = C: .ColHidden = True
                    Next
                    .Col = 6: .ColHidden = False
                    .ColWidth(6) = 15
                    For C = 7 To 23
                    .Col = C: .ColHidden = True
                    Next
                    For C = 24 To 27
                    .Col = C: .ColHidden = False
                    Next
                    .ColWidth(24) = 15
                    .ColWidth(25) = 11.5
                    .ColWidth(26) = 15
                    .ColWidth(27) = 15
                    For C = 28 To 35
                    .Col = C: .ColHidden = True
                    Next
                ElseIf VchType = 33 Then 'Closing Stock
                    .Col = 1: .ColHidden = True: .Col = 2: .ColHidden = True
                    For C = 3 To 8
                    .Col = C: .ColHidden = False
                    Next
                    .ColWidth(3) = 40.25 'Item
                    .ColWidth(4) = 7.75 'MRP
                    .ColWidth(5) = 13 'Item Group
                    .ColWidth(6) = 12 'Stock Qty.
                    .ColWidth(7) = 11 'Unit
                    .ColWidth(8) = 16 'SQ
                    For C = 9 To 17
                    .Col = C: .ColHidden = True
                    Next
                    For C = 18 To 20
                    .Col = C: .ColHidden = False
                    Next
                    .ColWidth(18) = 12 'Pending P/O
                    .ColWidth(19) = 12 'Pending S/O
                    .ColWidth(20) = 14 'Effective Stock
                    .Col = 21: .ColHidden = True
                    .Col = 22: .ColHidden = False
                    .ColWidth(22) = 13.25 ' Amount
                    For C = 23 To 35
                    .Col = C: .ColHidden = True
                    Next
                End If
            Else
                If VchType = 3 Or VchType = 7 Or VchType = 21 Or VchType = 25 Or VchType = 53 Or VchType = 57 Or VchType = 61 Or VchType = 65 Then 'Sales
                    .Col = 26: .ColHidden = True 'PurchaseAmount
                    .Col = 27: .ColHidden = True 'SalesAmount
                    .ColWidth(3) = 50
                    .ColWidth(5) = 30.75
                    .ColWidth(8) = 60
                    .ColWidth(12) = 60
                    .ColWidth(25) = 11
                ElseIf VchType = 4 Or VchType = 8 Or VchType = 22 Or VchType = 26 Or VchType = 54 Or VchType = 58 Or VchType = 62 Or VchType = 66 Then 'Sales & Purchase Return
                    .Col = 28: .ColHidden = True 'Purchase Return Amount
                    .Col = 29: .ColHidden = True 'Sales Return Amount
                    .ColWidth(3) = 50
                    .ColWidth(5) = 30.75
                    .ColWidth(9) = 60
                    .ColWidth(13) = 60
                    .ColWidth(25) = 11
                ElseIf VchType = 5 Or VchType = 9 Or VchType = 23 Or VchType = 27 Or VchType = 55 Or VchType = 59 Or VchType = 63 Or VchType = 67 Then 'Sales And Sales Return
                    .Col = 26: .ColHidden = True: .Col = 28: .ColHidden = True: .Col = 27: .ColHidden = True: .Col = 29: .ColHidden = True 'SalesAmount'Sales Return Amount
                    .ColWidth(3) = 50
                    .ColWidth(5) = 30.75
                    .ColWidth(8) = 30
                    .ColWidth(9) = 30
                    .ColWidth(12) = 30
                    .ColWidth(13) = 30
                    .ColWidth(25) = 11
                ElseIf VchType = 6 Or VchType = 10 Or VchType = 24 Or VchType = 28 Or VchType = 56 Or VchType = 60 Or VchType = 64 Or VchType = 68 Then 'Net Sale
                    .Col = 30: .ColHidden = True: .Col = 31: .ColHidden = True 'Net Sales Amount
                    .ColWidth(3) = 50
                    .ColWidth(5) = 30.75
                    .ColWidth(7) = 11
                    .ColWidth(23) = 60
                    .ColWidth(24) = 60
                    .ColWidth(25) = 11
                ElseIf VchType = 33 Then 'Closing Stock
                    .Col = 1: .ColHidden = True: .Col = 2: .ColHidden = True
                    For C = 3 To 8
                    .Col = C: .ColHidden = False
                    Next
                    .ColWidth(3) = 46.5 'Item
                    .ColWidth(4) = 7.75 'MRP
                    .ColWidth(5) = 13 'Item Group
                    .ColWidth(6) = 14 'Stock Qty.
                    .ColWidth(7) = 11 'Unit
                    .ColWidth(8) = 16 'SQ
                    For C = 9 To 17
                    .Col = C: .ColHidden = True
                    Next
                    For C = 18 To 20
                    .Col = C: .ColHidden = False
                    Next
                    .ColWidth(18) = 14 'Pending P/O
                    .ColWidth(19) = 14 'Pending S/O
                    .ColWidth(20) = 15 'Effective Stock
                    For C = 21 To 35
                    .Col = C: .ColHidden = True
                    Next
                    End If
            End If
End If
    End With
End Sub
Private Sub Check2_Click()
If Check2.Value Then
    VSFlexGrid1.Subtotal flexSTClear
    VSFlexGrid1_AfterDataRefresh
Else
    VSFlexGrid1.Subtotal flexSTClear
End If
End Sub
Private Sub PendingCheck_Click()
If TDBNumber1.Value <= 0 And PendingCheck.Value Then ZeroStock.Value = 0
If VSFlexFlag = True Then VSFlexGrid1.Subtotal flexSTClear
    Call cmdRefresh_Click
End Sub
Private Sub ZeroStock_Click()
If ZeroStock.Value Then NegativeStock.Value = 0
If TDBNumber1.Value <= 0 And ZeroStock.Value Then PendingCheck.Value = 0
    Call cmdRefresh_Click
End Sub
Private Sub NegativeStock_Click()
   If NegativeStock.Value And PendingCheck.Value = 0 Then ZeroStock.Value = 0
   If NegativeStock.Value And PendingCheck.Value And TDBNumber1 > 0 Then ZeroStock.Value = 0: TDBNumber1.Value = 0
Call cmdRefresh_Click
End Sub
Private Sub TDBNumber1_Change()
Dim n As Long, i As Long
If TDBNumber1 > 0 Then ZeroStock.Value = 1
    With fpSpread1
    If .DataRowCnt = 0 Then Exit Sub
            n = .DataRowCnt:
        For i = 1 To .DataRowCnt 'Unhide All
            .Row = i: .RowHidden = False
    Next
End With
Call cmdRefresh_Click
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
    
        fpSpread1.MaxCols = 35
            For C = 1 To .MaxCols
                fpSpread1.GetText C, 1, JQty
                fpSpread1.Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.Text = JQty: .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 11:
            Next
'                .Col = C: .Row = 1: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter:
                
            If VchType <= 2 Then
                .LockBackColor = RGB(255, 255, 240): Combo1.BackColor = RGB(255, 255, 240): Combo2.BackColor = RGB(255, 255, 240): MhDateInput1.BackColor = RGB(255, 255, 240): MhDateInput2.BackColor = RGB(255, 255, 240): TDBNumber1.BackColor = RGB(255, 255, 240): TDBNumber2.BackColor = RGB(255, 255, 240): Text1.BackColor = RGB(255, 255, 240):
            ElseIf (VchType >= 3 And VchType <= 6) Or (VchType >= 53 And VchType <= 56) Then
                .LockBackColor = RGB(245, 255, 230): Combo1.BackColor = RGB(245, 255, 230): Combo2.BackColor = RGB(245, 255, 230): MhDateInput1.BackColor = RGB(245, 255, 230): MhDateInput2.BackColor = RGB(245, 255, 230): TDBNumber1.BackColor = RGB(245, 255, 230): TDBNumber2.BackColor = RGB(245, 255, 230): Text1.BackColor = RGB(245, 255, 230):
            ElseIf (VchType >= 7 And VchType <= 10) Or (VchType >= 57 And VchType <= 60) Then
                .LockBackColor = RGB(245, 250, 250): Combo1.BackColor = RGB(245, 250, 250): Combo2.BackColor = RGB(245, 250, 250): MhDateInput1.BackColor = RGB(245, 250, 250): MhDateInput2.BackColor = RGB(245, 250, 250): TDBNumber1.BackColor = RGB(245, 250, 250): TDBNumber2.BackColor = RGB(245, 250, 250): Text1.BackColor = RGB(245, 250, 250):
            ElseIf (VchType >= 21 And VchType <= 24) Or (VchType >= 61 And VchType <= 64) Then
                .LockBackColor = RGB(255, 250, 255): Combo1.BackColor = RGB(255, 250, 255): Combo2.BackColor = RGB(255, 250, 255): MhDateInput1.BackColor = RGB(255, 250, 255): MhDateInput2.BackColor = RGB(255, 250, 255): TDBNumber1.BackColor = RGB(255, 250, 255): TDBNumber2.BackColor = RGB(255, 250, 255): Text1.BackColor = RGB(255, 250, 255):
            ElseIf (VchType >= 25 And VchType <= 28) Or (VchType >= 65 And VchType <= 68) Then
                .LockBackColor = RGB(240, 255, 255): Combo1.BackColor = RGB(240, 255, 255): Combo2.BackColor = RGB(240, 255, 255): MhDateInput1.BackColor = RGB(240, 255, 255): MhDateInput2.BackColor = RGB(240, 255, 255): TDBNumber1.BackColor = RGB(240, 255, 255): TDBNumber2.BackColor = RGB(240, 255, 255): Text1.BackColor = RGB(240, 255, 255):
            End If
            
            If VchType = 0 Then .ColWidth(3) = 49.25: .ColWidth(4) = 15: .ColWidth(5) = 15: .ColWidth(33) = 24: .ColWidth(34) = 22.75: .Col = 33: .ColHidden = False: .Col = 34: .ColHidden = False
            If (VchType <= 10 And VchType >= 7) Or (VchType <= 28 And VchType >= 25) Or (VchType >= 57 And VchType <= 60) Then fpSpread1.DeleteRows 1, 2 Else: fpSpread1.DeleteRows 1, 1
            For R = 1 To .DataRowCnt - 1
            .Col = 33: .Row = R: .Lock = False
            Next
                    
            fpSpread1.DeleteRows .DataRowCnt, 1
    
    Call Total_Click
    fpSpread1.Col = 6: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignRight
    fpSpread1.Col = 7: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignCenter
    fpSpread1.Col = 33: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignRight
    fpSpread1.Col = 34: fpSpread1.Row = .DataRowCnt: .TypeHAlign = TypeHAlignRight
    End With
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbEnter And Shift = vbCtrlMask Then Call cmdFilter_Click
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstStockLedger)
    Call CloseRecordset(rstItemOpening)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstItemList)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
    sMcCode = "": SCode = "": oSCode = "":  vtType = "": vDate = "": vTypeCode = "": vtNo = "":
End Sub
Private Sub cmdFilter_Click()
 Dim i As Integer, cVal As Variant, n As Integer, R As Long, C As Long, Cols As Long
 
C = 0
    If VSFlexGrid1.Visible Then
    If VSFlexGrid1.BottomRow = 1 Then Exit Sub
    If Text1.Text = "" Then Exit Sub
            If VSFlexGrid1.BottomRow = 0 Then Exit Sub
              VSFlexGrid1.Subtotal flexSTClear
              n = VSFlexGrid1.BottomRow
            For i = 1 To VSFlexGrid1.Rows  'Unhide All
                VSFlexGrid1.RowHidden(i) = False
            Next

            For i = 0 To VSFlexGrid1.RightCol  'Match Col Header
            C = C + 1
            If C > VSFlexGrid1.RightCol Then Exit Sub
            cVal = StrConv(VSFlexGrid1.TextMatrix(0, C), vbUpperCase)
            If StrConv(Combo2.Value, vbUpperCase) = cVal Then Exit For
            
            Next
            
    For i = 1 To VSFlexGrid1.BottomRow
                    If VSFlexGrid1.BottomRow < i Or n = 0 Then Exit For
                    If Combo2.ListIndex >= 0 And n <> 0 Then cVal = VSFlexGrid1.TextMatrix(i, C)
                    
            If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) <> 0 Then
                '****'
            Else
                If Not VSFlexGrid1.RowHidden(i) Then
                    VSFlexGrid1.RemoveItem (i): LR = i: n = n - 1 'Hide Filter
                    i = i - 1
                End If
            End If
                    TDBNumber2 = n 'Data Count
    Next
    
    Call VSFlexGrid1_AfterDataRefresh
  Else
    Call Total_Click
  End If
End Sub
Private Sub Command2_Click()
  Dim i As Integer, cVal As Variant, R As Long, C As Long
  If VSFlexGrid1.Visible Then
    If VSFlexGrid1.BottomRow = 1 Then Exit Sub
  If Text1.Text = "" Then Exit Sub
            If VSFlexGrid1.BottomRow = 0 Then Exit Sub
            For i = 1 To VSFlexGrid1.Rows  'Unhide All
                VSFlexGrid1.RowHidden(i) = False
            Next

            
            For i = 1 To VSFlexGrid1.RightCol  'Match Col Header
            C = C + 1
            cVal = VSFlexGrid1.TextMatrix(0, C)
            If Combo2.Value = cVal Then Exit For
            Next
            
            R = IIf(VSFlexGrid1.Row + 1 <> LR, VSFlexGrid1.Row + 1, 1)
            LR = R
            
            For i = R To VSFlexGrid1.BottomRow
            If Combo2.ListIndex >= 0 Then cVal = VSFlexGrid1.TextMatrix(i, C)
                        If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
                        ElseIf Combo2.ListIndex >= 0 Then
                        VSFlexGrid1.Row = i: VSFlexGrid1.Col = C:  Exit Sub
                        End If
            Next
  Else
    With fpSpread1
    If Text1.Text = "" Then Exit Sub
            If .DataRowCnt = 0 Then Exit Sub
                For i = 1 To .DataRowCnt 'Unhide All
                .Row = i: .RowHidden = False
            Next
        fpSpread1.MaxCols = 35
        If VchType < 28 And Combo2.ListIndex = 0 Then C = Combo2.ListIndex + 3
        If VchType < 28 And Combo2.ListIndex = 1 Then C = Combo2.ListIndex + 4
        If VchType >= 29 And Combo2.ListIndex = 0 Then C = Combo2.ListIndex + 3
        If VchType >= 29 And Combo2.ListIndex = 1 Then C = Combo2.ListIndex + 4
        
        If VchType >= 34 And VchType <= 45 Then C = Combo2.ListIndex + 1
        If VchType >= 34 And VchType <= 45 And Combo2.ListIndex = 3 Then C = Combo2.ListIndex + 2
        If VchType >= 53 And VchType <= 68 And Combo2.ListIndex = 0 Then C = Combo2.ListIndex + 3
        If VchType >= 53 And VchType <= 68 And Combo2.ListIndex = 1 Then C = Combo2.ListIndex + 4
        If VchType = 46 Then C = Combo2.ListIndex + 1
        If VchType = 47 Then C = Combo2.ListIndex + 5
            R = IIf(.ActiveRow + 1 <> LR, .ActiveRow + 1, 1)
            LR = R
            
            For i = R To .DataRowCnt
            If Combo2.ListIndex >= 0 Then .GetText C, i, cVal
                        If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
                        ElseIf Combo2.ListIndex >= 0 Then
                        .SetActiveCell C, i: Exit Sub
                        End If
            Next
    End With
    End If
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
    Dim i As Integer, cVal As Variant, n As Integer, R As Long, C As Long, Cols As Long
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
    Dim PStockVal As Variant, PStockTotal As Variant
    Dim JStockVal As Variant, JStockTotal As Variant
    With fpSpread1
        If .DataRowCnt = 0 Then Exit Sub
        n = .DataRowCnt: StockVal = 0
            For i = 1 To .DataRowCnt 'Unhide All
            .GetText 3, i, cVal
                If TotalFlag = False Then .Row = i: .RowHidden = False
                If cVal = "Grand Total" Then fpSpread1.DeleteRows i, 1
            Next
        If VchType = 46 Then fpSpread1.MaxCols = 38 Else fpSpread1.MaxCols = 35

        If VchType <= 28 And Combo2.ListIndex = 0 Then C = Combo2.ListIndex + 3
        If VchType <= 28 And Combo2.ListIndex = 1 Then C = Combo2.ListIndex + 4
        If VchType >= 29 And Combo2.ListIndex = 0 Then C = Combo2.ListIndex + 3
        If VchType >= 29 And Combo2.ListIndex = 1 Then C = Combo2.ListIndex + 4
        If VchType <= 53 And Combo2.ListIndex = 0 Then C = Combo2.ListIndex + 3
        If VchType <= 53 And Combo2.ListIndex = 1 Then C = Combo2.ListIndex + 4
        If VchType = 46 Then C = Combo2.ListIndex + 1
        If VchType = 47 Then C = Combo2.ListIndex + 5
        If Right(VchType, 2) = 48 And Combo2.ListIndex = 0 Then C = Combo2.ListIndex + 2
    
    
    For i = 1 To .DataRowCnt
                If Combo2.ListIndex >= 0 Then .GetText C, i, cVal
                .GetText 6, i, StockVal
                .GetText 8, i, PVal
                .GetText 9, i, PRVal
                .GetText 10, i, PCVal
                .GetText 11, i, PRCVal
                .GetText 12, i, SVal
                .GetText 13, i, SRVal
                .GetText 14, i, SCVal
                .GetText 15, i, SRCVal
                .GetText 16, i, SJIVal
                .GetText 17, i, SJOVal
                .GetText 18, i, POVal
                .GetText 19, i, SOVal
                .GetText 20, i, EStockVal
                .GetText 22, i, AVal
                .GetText 23, i, NPVal
                .GetText 24, i, NSVal
                .GetText 26, i, PAVal
                .GetText 27, i, SAVal
                .GetText 28, i, PRAVal
                .GetText 29, i, SRAVal
                .GetText 30, i, NPAVal
                .GetText 31, i, NSAVal
                .GetText 33, i, PStockVal
                .GetText 34, i, JStockVal
                 .GetText 3, i, cVal
                If cVal = "Grand Total" Then fpSpread1.DeleteRows .DataRowCnt, 1
                 .GetText 4, i, cVal
                If cVal = "Grand Total" Then fpSpread1.DeleteRows .DataRowCnt, 1
                 .GetText 5, i, cVal
                If cVal = "Grand Total" Then fpSpread1.DeleteRows .DataRowCnt, 1
                .GetText C, i, cVal
            If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
                    .Row = i: .RowHidden = True: n = n - 1 'Hide Filter
            Else
                  .Row = i
                If Not .RowHidden Then
                        StockTotal = Val(StockTotal) + Val(StockVal) '6
                        PTotal = Val(PTotal) + Val(PVal) '8
                        PRTotal = Val(PRTotal) + Val(PRVal) '9
                        PCTotal = Val(PCTotal) + Val(PCVal) '10
                        PRCTotal = Val(PRCTotal) + Val(PRCVal) '11
                        STotal = Val(STotal) + Val(SVal) '12
                        SRTotal = Val(SRTotal) + Val(SRVal) '13
                        SCTotal = Val(SCTotal) + Val(SCVal) '14
                        SRCTotal = Val(SRCTotal) + Val(SRCVal) '15
                        SJITotal = Val(SJITotal) + Val(SJIVal) '16
                        SJOTotal = Val(SJOTotal) + Val(SJOVal) '17
                        POTotal = Val(POTotal) + Val(POVal) '18
                        SOTotal = Val(SOTotal) + Val(SOVal) '19
                        EStockTotal = Val(EStockTotal) + Val(EStockVal) '20
                        ATotal = Val(ATotal) + Val(AVal) '22
                        NPValTotal = Val(NPValTotal) + Val(NPVal) '23
                        NSValTotal = Val(NSValTotal) + Val(NSVal) '24
                        PAValTotal = Val(PAValTotal) + Val(PAVal) '26
                        SAValTotal = Val(SAValTotal) + Val(SAVal) '27
                        PRAValTotal = Val(PRAValTotal) + Val(PRAVal) '28
                        SRAValTotal = Val(SRAValTotal) + Val(SRAVal) '29
                        NPAValTotal = Val(NPAValTotal) + Val(NPAVal) '30
                        NSAValTotal = Val(NSAValTotal) + Val(NSAVal) '31
                        PStockTotal = Val(PStockTotal) + Val(PStockVal) '33
                        JStockTotal = Val(JStockTotal) + Val(JStockVal) '34
                    End If
            End If
                    TDBNumber2 = n 'Data Count
        Next
    
                For C = 3 To 5
                          .Col = C
                If Not .ColHidden Then
                .SetText C, i, "Grand Total": Exit For
                End If
                Next
                .SetText 6, i, StockTotal
                .SetText 7, i, "Units"
                .SetText 8, i, PTotal
                .SetText 9, i, PRTotal
                .SetText 10, i, PCTotal
                .SetText 11, i, PRCTotal
                .SetText 12, i, STotal
                .SetText 13, i, SRTotal
                .SetText 14, i, SCTotal
                .SetText 15, i, SRCTotal
                .SetText 16, i, SJITotal
                .SetText 17, i, SJOTotal
                .SetText 18, i, POTotal
                .SetText 19, i, SOTotal
                .SetText 20, i, EStockTotal
                .SetText 22, i, ATotal
                .SetText 23, i, NPValTotal
                .SetText 24, i, NSValTotal
                .SetText 25, i, "Units"
                .SetText 26, i, PAValTotal
                .SetText 27, i, SAValTotal
                .SetText 28, i, PRAValTotal
                .SetText 29, i, SRAValTotal
                .SetText 30, i, NPAValTotal
                .SetText 31, i, NSAValTotal
                .SetText 33, i, PStockTotal
                .SetText 34, i, JStockTotal
    End With
                Call Fomatting_Click
    fpSpread1.MaxRows = IIf(TDBNumber2.Value < 27, i + (27 - TDBNumber2.Value), i + 1)
End Sub
Private Sub Fomatting_Click()
Dim R As Long, C As Long, Cols As Long, Rows As Long
        With fpSpread1
       If VchType = 46 Then fpSpread1.MaxCols = 38 Else fpSpread1.MaxCols = 35
            Cols = .MaxCols
            R = .DataRowCnt
            For C = 1 To Cols
            fpSpread1.Col = C: fpSpread1.Row = R: fpSpread1.FontBold = True: fpSpread1.FontSize = 12.5: fpSpread1.FontUnderline = True: fpSpread1.ForeColor = vbBlue:
        Next
            'Formatting
            If VchType = 0 Then
                C = 33
            For R = 1 To (.DataRowCnt - 1)
                .Lock = False
            Next

            Else
                .SelectBlockOptions = SelectBlockOptionsAll
            End If
                If VchType <> 0 Then .SetActiveCell 3, LR
        End With
End Sub
Private Function FormatCol()
Dim i As Long, R As Long, C As Long
With fpSpread1
        If VchType <= 28 Then .Col = 1: .ColHidden = True: .Col = 2: .ColHidden = True:
        If VchType >= 53 And VchType <= 68 Then .Col = 1: .ColHidden = True: .Col = 2: .ColHidden = True:
        If VchType = 0 Then 'Physical Stock Audit Ledger
            For C = 3 To 7
            .Col = C: .ColHidden = False
            Next
            .ColWidth(3) = 49.25 'Item
            .ColWidth(4) = 15 'MRP
            .ColWidth(5) = 15 'Item Group
            .ColWidth(6) = 14  'Stock Qty.
            .ColWidth(7) = 11.5 'Unit
            For C = 8 To 32
            .Col = C: .ColHidden = True
            Next
            For C = 33 To 34
            .Col = C: .ColHidden = False
            Next
            .ColWidth(33) = 24 'Phical Stock Qty.
            .ColWidth(34) = 22.75 'Stock Impact
            .Col = 35: .ColHidden = True
        ElseIf VchType = 1 Then 'Inventory Ledger
            For C = 3 To 20
            .Col = C: .ColHidden = False
            Next
            .ColWidth(3) = 50
            .ColWidth(4) = 12
            .ColWidth(5) = 25
            .ColWidth(6) = 12
            .ColWidth(7) = 12
            .ColWidth(8) = 20
            .ColWidth(9) = 20
            .ColWidth(10) = 17.375
            .ColWidth(11) = 20
            .ColWidth(12) = 24.125
            .ColWidth(13) = 20
            .ColWidth(14) = 20
            .ColWidth(15) = 20
            .ColWidth(16) = 17.5
            .ColWidth(17) = 24.125
            .ColWidth(18) = 19.75
            .ColWidth(19) = 22.25
            .ColWidth(20) = 17.875
            .Col = 21: .ColHidden = True
            For C = 22 To 24
            .Col = C: .ColHidden = False
            Next
            .ColWidth(22) = 17.375
            .ColWidth(23) = 24.125
            .ColWidth(24) = 19.875
            .Col = 25: .ColHidden = True
            For C = 26 To 31
            .Col = C: .ColHidden = False
            Next
            .ColWidth(26) = 22.125
            .ColWidth(27) = 35.625
            .ColWidth(28) = 24
            .ColWidth(29) = 20
            .ColWidth(30) = 22
            .ColWidth(31) = 35.75
            For C = 32 To 35
            .Col = C: .ColHidden = True
            Next
        ElseIf VchType = 2 Then 'Closing Stock
            For C = 3 To 7
            .Col = C: .ColHidden = False
            Next
            .ColWidth(3) = 49.25 'Item
            .ColWidth(4) = 8 'MRP
            .ColWidth(5) = 15 'Item Group
            .ColWidth(6) = 14 'Stock Qty.
            .ColWidth(7) = 11 'Unit
            For C = 8 To 17
            .Col = C: .ColHidden = True
            Next
            For C = 18 To 20
            .Col = C: .ColHidden = False
            Next
            .ColWidth(18) = 12 'Pending P/O
            .ColWidth(19) = 12 'Pending S/O
            .ColWidth(20) = 14 'Effective Stock
            .Col = 21: .ColHidden = True
            .Col = 22: .ColHidden = False
            .ColWidth(22) = 16 ' Amount
            For C = 23 To 35
            .Col = C: .ColHidden = True
            Next
        ElseIf VchType = 3 Or VchType = 7 Or VchType = 21 Or VchType = 25 Or VchType = 53 Or VchType = 57 Or VchType = 61 Or VchType = 65 Then 'Sales
            .Col = 3: .ColHidden = False
            .ColWidth(3) = 50 'Item
            .Col = 4: .ColHidden = True:
            .Col = 5: .ColHidden = False
            .ColWidth(5) = 30.75 'Item Group
            .Col = 6: .ColHidden = True:
            For C = 7 To 11
            .Col = C: .ColHidden = True
            Next
        If VchType = 53 Or VchType = 57 Or VchType = 61 Or VchType = 65 Then
            .Col = 8: .ColHidden = False
            .ColWidth(8) = 60 'Sales
            .Col = 12: .ColHidden = True
        ElseIf VchType = 3 Or VchType = 7 Or VchType = 21 Or VchType = 25 Then
            .Col = 8: .ColHidden = True
            .Col = 12: .ColHidden = False
            .ColWidth(12) = 60 'Purchase
        End If
            For C = 13 To 24
            .Col = C: .ColHidden = True
            Next
            .Col = 25: .ColHidden = False
            .ColWidth(25) = 11 'Unit
            For C = 26 To 35
            .Col = C: .ColHidden = True
            Next
        ElseIf VchType = 4 Or VchType = 8 Or VchType = 22 Or VchType = 26 Or VchType = 54 Or VchType = 58 Or VchType = 62 Or VchType = 66 Then 'Sales & Purchase Return
            .Col = 3: .ColHidden = False
            .ColWidth(3) = 50 'Item
            .Col = 4: .ColHidden = True:
            .Col = 5: .ColHidden = False
            .ColWidth(5) = 30.75 'Item Group
            .Col = 6: .ColHidden = True:
            For C = 7 To 12
            .Col = C: .ColHidden = True
            Next
        If VchType = 54 Or VchType = 58 Or VchType = 62 Or VchType = 66 Then
            .Col = 9: .ColHidden = False
            .ColWidth(9) = 60 'Sales
            .Col = 13: .ColHidden = True
        ElseIf VchType = 4 Or VchType = 8 Or VchType = 22 Or VchType = 26 Then
            .Col = 9: .ColHidden = True
            .Col = 13: .ColHidden = False
            .ColWidth(13) = 60 'Purchase
        End If
            For C = 14 To 24
            .Col = C: .ColHidden = True
            Next
            .Col = 25: .ColHidden = False
            .ColWidth(25) = 11 'Unit
            For C = 26 To 35
            .Col = C: .ColHidden = True
            Next
        ElseIf VchType = 5 Or VchType = 9 Or VchType = 23 Or VchType = 27 Or VchType = 55 Or VchType = 59 Or VchType = 63 Or VchType = 67 Then 'Sales And Sales Return
            .Col = 3: .ColHidden = False
            .ColWidth(3) = 50 'Item
            .Col = 4: .ColHidden = True:
            .Col = 5: .ColHidden = False
            .ColWidth(5) = 30.75 'Item Group
            .Col = 6: .ColHidden = True:
            For C = 7 To 11
            .Col = C: .ColHidden = True
            Next
        If VchType = 55 Or VchType = 59 Or VchType = 63 Or VchType = 67 Then
            .Col = 8: .ColHidden = False
            .ColWidth(8) = 30 'Sales
            .Col = 9: .ColHidden = False
            .ColWidth(9) = 30 'Sales Return
            .Col = 12: .ColHidden = True
            .Col = 13: .ColHidden = True
        ElseIf VchType = 5 Or VchType = 9 Or VchType = 23 Or VchType = 27 Then
            .Col = 8: .ColHidden = True
            .Col = 9: .ColHidden = True
            .Col = 12: .ColHidden = False
            .ColWidth(12) = 30 'Sales
            .Col = 13: .ColHidden = False
            .ColWidth(13) = 30 'Sales Return
        End If
            For C = 14 To 24
            .Col = C: .ColHidden = True
            Next
            .Col = 25: .ColHidden = False
            .ColWidth(25) = 11 'Unit
            For C = 26 To 35
            .Col = C: .ColHidden = True
            Next
        ElseIf VchType = 6 Or VchType = 10 Or VchType = 24 Or VchType = 28 Or VchType = 56 Or VchType = 60 Or VchType = 64 Or VchType = 68 Then 'Net Sale
            .Col = 3: .ColHidden = False
            .ColWidth(3) = 50 'Item
            .Col = 4: .ColHidden = True:
            .Col = 5: .ColHidden = False
            .ColWidth(5) = 30 'Item Group
            For C = 6 To 23
            .Col = C: .ColHidden = True
            Next
        If VchType = 56 Or VchType = 60 Or VchType = 64 Or VchType = 68 Then
            .Col = 23: .ColHidden = False
            .Col = 24: .ColHidden = True
            .ColWidth(23) = 60 'Sales
            .Col = 25: .ColHidden = False
            .ColWidth(25) = 11 'Unit
        ElseIf VchType = 6 Or VchType = 10 Or VchType = 24 Or VchType = 28 Then
            .Col = 23: .ColHidden = True
            .Col = 24: .ColHidden = False
            .ColWidth(24) = 60 'Sales
            .Col = 25: .ColHidden = False
            .ColWidth(25) = 11 'Unit
        End If
            For C = 26 To 35
            .Col = C: .ColHidden = True
            Next
        ElseIf VchType >= 29 And VchType <= 30 Then  'Pending Purchase Order Voucher-wise
            For C = 1 To 3
            .Col = C: .ColHidden = False
            Next
            .ColWidth(1) = 10
            .ColWidth(2) = 15
            .ColWidth(3) = 15
            .Col = 4: .ColHidden = True
            .Col = 5: .ColHidden = False
            .ColWidth(5) = 40
            .Col = 6: .ColHidden = False
            .ColWidth(6) = 15
            For C = 7 To 23
            .Col = C: .ColHidden = True
            Next
            For C = 24 To 27
            .Col = C: .ColHidden = False
            Next
            .ColWidth(24) = 15
            .ColWidth(25) = 11.5
            .ColWidth(26) = 12.25
            .ColWidth(27) = 17.5
            For C = 28 To 35
            .Col = C: .ColHidden = True
            Next
        ElseIf VchType = 31 Or Right(VchType, 2) = 48 Then 'Item Ledger
            For C = 1 To 3
            .Col = C: .ColHidden = False
            Next
            .ColWidth(1) = 10 'Date
            .ColWidth(2) = 13 'Vch/BillNo
            .ColWidth(3) = 23 'Vch Type
            If Right(VchType, 2) = 48 Then
                .Col = 4: .ColHidden = False
                .ColWidth(4) = 23 'Voucher Series
            Else
                .Col = 4: .ColHidden = True
            End If
            .Col = 5: .ColHidden = False
            .ColWidth(5) = 30 'Particulars
            .Col = 6: .ColHidden = False
            .ColWidth(6) = 13 'Qty1
            For C = 7 To 22
            .Col = C: .ColHidden = True
            Next
            For C = 23 To 27
            .Col = C: .ColHidden = False
            Next
            .ColWidth(23) = 13 'Qty2
            If Right(VchType, 2) = 48 Then .ColWidth(24) = 0 Else .ColWidth(24) = 13 'Qty2
            .ColWidth(25) = 11.5 'Unit
            If Right(VchType, 2) = 48 Then .ColWidth(26) = 0 Else .ColWidth(26) = 10.25 'Rate
            .ColWidth(27) = 14.5 'Amount
            For C = 28 To 38
            .Col = C: .ColHidden = True
            Next
        ElseIf VchType = 49 Or VchType = 32 Then  'Item Ledger Material Centre-wise
            For C = 1 To 2
            .Col = C: .ColHidden = True
            Next
            .Col = 3: .ColHidden = False
            .ColWidth(3) = 38 'Material Centre
            .Col = 4: .ColHidden = True
            .Col = 5: .ColHidden = True
            .Col = 6: .ColHidden = False
            .ColWidth(6) = 25.5 'Opening
            For C = 7 To 22
            .Col = C: .ColHidden = True
            Next
            .Col = 20: .ColHidden = False
            .ColWidth(20) = 25.5 'cINWard
            For C = 23 To 25
            .Col = C: .ColHidden = False
            Next
            .ColWidth(23) = 25.5 'cOutWard
            .ColWidth(24) = 25.5 'Closing
            .ColWidth(25) = 11.5 'Unit
            For C = 26 To 35
            .Col = C: .ColHidden = True
            Next
        ElseIf VchType = 33 Then 'Closing Stock
            .Col = 1: .ColHidden = True: .Col = 2: .ColHidden = True
            For C = 3 To 8
            .Col = C: .ColHidden = False
            Next
            .ColWidth(3) = 40.25 'Item
            .ColWidth(4) = 7.75 'MRP
            .ColWidth(5) = 13 'Item Group
            .ColWidth(6) = 12 'Stock Qty.
            .ColWidth(7) = 11 'Unit
            .ColWidth(8) = 16 'SQ
            For C = 9 To 17
            .Col = C: .ColHidden = True
            Next
            For C = 18 To 20
            .Col = C: .ColHidden = False
            Next
            .ColWidth(18) = 12 'Pending P/O
            .ColWidth(19) = 12 'Pending S/O
            .ColWidth(20) = 14 'Effective Stock
            .Col = 21: .ColHidden = True
            .Col = 22: .ColHidden = False
            .ColWidth(22) = 13.25 ' Amount
            For C = 23 To 35
            .Col = C: .ColHidden = True
            Next
        ElseIf VchType >= 34 And VchType <= 45 Then  'Pending Purchase Order Voucher-wise
            .ColWidth(0) = 6
            For C = 1 To 3
            .Col = C: .ColHidden = False
            Next
            .ColWidth(1) = 10.5
            .ColWidth(2) = 13
            .ColWidth(3) = 14.75
            .Col = 4: .ColHidden = True
            .Col = 5: .ColHidden = False
            If VchType = 45 Then .Col = 8: .ColHidden = False: .Col = 23: .ColHidden = False
            If VchType >= 39 And VchType <= 44 Then .ColWidth(5) = 46.75 Else .ColWidth(5) = 38.875 '55.75 '
            .Col = 6: .ColHidden = False
            .ColWidth(6) = 12
            .Col = 7: .ColHidden = True
            .Col = 8: .ColHidden = False
            .ColWidth(8) = 12
            For C = 9 To 22
            .Col = C: .ColHidden = True
            Next
            For C = 23 To 27
            .Col = C: .ColHidden = False
            Next
            .ColWidth(23) = 13
            .ColWidth(24) = 11
            .ColWidth(25) = 6.625 '7.375
            If VchType = 41 Then .ColWidth(26) = 10.625 Else .ColWidth(26) = 6.375  '11
            If VchType >= 39 And VchType <= 44 Then .ColWidth(27) = 16.25 Else .ColWidth(27) = 11
            For C = 28 To 35
            .Col = C: .ColHidden = True
            Next
            If VchType >= 39 And VchType <= 41 Then .Col = 23: .ColHidden = True
            If VchType >= 42 And VchType <= 44 Then .Col = 8: .ColHidden = True
        ElseIf VchType >= 46 And VchType <= 48 Then  'Pending Sale AND Purchase
            For C = 1 To 14
            .Col = C: .ColHidden = False
            Next
        If VchType = 46 Then
            .ColWidth(1) = 10 'Date
            .ColWidth(2) = 13 'Vch/BillNo
            .ColWidth(3) = 42.125 'Particulars
            .ColWidth(4) = 8.25 'Unit Rate
            .ColWidth(5) = 30.625 'Buyer Name
            .ColWidth(6) = 11.5 'Qty
            .ColWidth(7) = 8.625 'Unit
            .ColWidth(8) = 13 'Pending Qty
            .ColWidth(9) = 14.25 'Pending Amount
            .ColWidth(10) = 19.125 'BilledQtyC
            .ColWidth(11) = 19.5 'BilledQtyD
            .ColWidth(12) = 20.25 'ChallanQty
            .ColWidth(13) = 27.25 'DirectQty
            .ColWidth(14) = 13 'Clear Qty
            .ColWidth(15) = 3 'Check Box
        ElseIf VchType = 47 Then
            .ColWidth(1) = 10 'Date
            .ColWidth(2) = 13 'Vch/BillNo
            .ColWidth(3) = 42.125 'Particulars
            .ColWidth(4) = 8.25 'Unit Rate
            .ColWidth(5) = 28.625 'Buyer Name
            .ColWidth(6) = 11 'Qty
            .ColWidth(7) = 5.825 'Unit
            .ColWidth(8) = 11.5 'Pending Qty
            .ColWidth(9) = 13.875 'Pending Amount
            .ColWidth(10) = 18 'BilledQtyC
            .ColWidth(11) = 11.625 'BilledQtyD
            .ColWidth(12) = 19.25 'ChallanQty
            .ColWidth(13) = 18.5 'DirectQty
            .ColWidth(14) = 13 'Clear Qty
            For C = 1 To 4
                .Col = C: .ColHidden = True
            Next
        End If
            For C = 16 To 35
            .Col = C: .ColHidden = True
            Next
            .ColWidth(36) = 10 'CreatedBY
            .ColWidth(37) = 15 'CreatedOn
            .ColWidth(38) = 40 'Remarks
        End If
End With
End Function
Private Function FormatHeader()
Dim i As Long, R As Long, C As Long
With fpSpread1
            For C = 1 To .MaxCols
                fpSpread1.Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12:
            Next
        If (VchType >= 0 And VchType <= 6) Or (VchType >= 53 And VchType <= 56) Then
            fpSpread1.ColHeaderRows = 1:
            'Header1
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Item Name": fpSpread1.FontBold = False 'fpSpread1.Col = 2: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Vch/Bill No.":
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Item Name": fpSpread1.Col = 4: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "MRP": fpSpread1.Col = 5: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Parent Group": fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Stock Qty.": fpSpread1.Col = 7: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Units": fpSpread1.Col = 8: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Purchases Qty.":
            fpSpread1.Col = 9: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Purchases Return Qty.": fpSpread1.Col = 10: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Purchases Challan": fpSpread1.Col = 11: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Purchases Return Challan": fpSpread1.Col = 12: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Qty.":
            fpSpread1.Col = 13: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Return Qty.": fpSpread1.Col = 14: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Challan": fpSpread1.Col = 15: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Return Challan": fpSpread1.Col = 16: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Stock Journal IN":
            fpSpread1.Col = 17: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Stock Journal OUT": fpSpread1.Col = 18: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Pending P/O": fpSpread1.Col = 19: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Pending S/O": fpSpread1.Col = 20: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Effective Stock": fpSpread1.Col = 21: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Price":
            fpSpread1.Col = 22: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = " Amount": fpSpread1.Col = 23: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Net Purchases": fpSpread1.Col = 24: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Net Sales": fpSpread1.Col = 25: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Units": fpSpread1.Col = 26: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = " Purchases Amount":
            fpSpread1.Col = 27: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Amount": fpSpread1.Col = 28: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Purchases Return Amount": fpSpread1.Col = 29: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Return Amt.": fpSpread1.Col = 30: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Net Purchases Amount":
            fpSpread1.Col = 31: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Net Sales Amount":
        ElseIf (VchType >= 7 And VchType <= 10) Or (VchType >= 57 And VchType <= 60) Then
                fpSpread1.ColHeaderRows = 2:
            For C = 4 To .MaxCols
                .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "": .Col = C: fpSpread1.Row = SpreadHeader + 1: fpSpread1.FontSize = 12:
            Next
        'Header1
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Party : " + rstStockLedger.Fields("OneParty").Value: fpSpread1.FontSize = 12:
            Header1 = "Party : " + rstStockLedger.Fields("OneParty").Value: fpSpread1.Row = SpreadHeader + 1: fpSpread1.FontSize = 12:
        'Header2
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Date": fpSpread1.Col = 2: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Vch/Bill No.":
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Item Name": fpSpread1.Col = 4: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "MRP": fpSpread1.Col = 5: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Parent Group": fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Stock Qty.": fpSpread1.Col = 7: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Units": fpSpread1.Col = 8: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Purchases Qty.":
            fpSpread1.Col = 9: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Purchases Return Qty.": fpSpread1.Col = 10: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Purchases Challan": fpSpread1.Col = 11: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Purchases Return Challan": fpSpread1.Col = 12: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Qty.":
            fpSpread1.Col = 13: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Return Qty.": fpSpread1.Col = 14: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Challan": fpSpread1.Col = 15: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Return Challan": fpSpread1.Col = 16: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Stock Journal IN":
            fpSpread1.Col = 17: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Stock Journal OUT": fpSpread1.Col = 18: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Pending P/O": fpSpread1.Col = 19: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Pending S/O": fpSpread1.Col = 20: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Effective Stock": fpSpread1.Col = 21: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Price":
            fpSpread1.Col = 22: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = " Amount": fpSpread1.Col = 23: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Net Purchases": fpSpread1.Col = 24: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Net Sales": fpSpread1.Col = 25: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Units": fpSpread1.Col = 26: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = " Purchases Amount":
            fpSpread1.Col = 27: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Amount": fpSpread1.Col = 28: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Purchases Return Amount": fpSpread1.Col = 29: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Return Amt.": fpSpread1.Col = 30: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Net Purchases Amount":
            fpSpread1.Col = 31: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Net Sales Amount":
        ElseIf VchType <= 20 And VchType >= 11 Then
            'VchType Used For Paper Ledger
        ElseIf (VchType >= 21 And VchType <= 24) Or (VchType >= 61 And VchType <= 64) Then
        'Header1
            fpSpread1.ColHeaderRows = 1:
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Date": fpSpread1.Col = 2: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Vch/Bill No.":
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Party Name": fpSpread1.Col = 4: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "MRP": fpSpread1.Col = 5: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Parent Group": fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Stock Qty.": fpSpread1.Col = 7: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Units": fpSpread1.Col = 8: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Purchases Qty.":
            fpSpread1.Col = 9: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Purchases Return Qty.": fpSpread1.Col = 10: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Purchases Challan": fpSpread1.Col = 11: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Purchases Return Challan": fpSpread1.Col = 12: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Qty.":
            fpSpread1.Col = 13: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Return Qty.": fpSpread1.Col = 14: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Challan": fpSpread1.Col = 15: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Return Challan": fpSpread1.Col = 16: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Stock Journal IN":
            fpSpread1.Col = 17: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Stock Journal OUT": fpSpread1.Col = 18: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Pending P/O": fpSpread1.Col = 19: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Pending S/O": fpSpread1.Col = 20: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Effective Stock": fpSpread1.Col = 21: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Price":
            fpSpread1.Col = 22: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = " Amount": fpSpread1.Col = 23: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Net Purchases": fpSpread1.Col = 24: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Net Sales": fpSpread1.Col = 25: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Units": fpSpread1.Col = 26: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = " Purchases Amount":
            fpSpread1.Col = 27: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Amount": fpSpread1.Col = 28: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Purchases Return Amount": fpSpread1.Col = 29: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Sales Return Amt.": fpSpread1.Col = 30: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Net Purchases Amount":
            fpSpread1.Col = 31: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Net Sales Amount":
        ElseIf (VchType >= 25 And VchType <= 28) Or (VchType >= 65 And VchType <= 68) Then
                fpSpread1.ColHeaderRows = 2:
                For C = 1 To .MaxCols
                    .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "": .Col = C: fpSpread1.Row = SpreadHeader + 1: fpSpread1.FontSize = 12:
                Next
        'Header1
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Item : " + rstStockLedger.Fields("OneItem").Value: fpSpread1.FontSize = 12:
            Header1 = "Item : " + rstStockLedger.Fields("OneItem").Value: fpSpread1.Row = SpreadHeader + 1: fpSpread1.FontSize = 12:
        'Header2
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Date": fpSpread1.Col = 2: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Vch/Bill No.":
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Party Name": fpSpread1.Col = 4: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "MRP": fpSpread1.Col = 5: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Parent Group": fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Stock Qty.": fpSpread1.Col = 7: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Units": fpSpread1.Col = 8: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Purchases Qty.":
            fpSpread1.Col = 9: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Purchases Return Qty.": fpSpread1.Col = 10: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Purchases Challan": fpSpread1.Col = 11: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Purchases Return Challan": fpSpread1.Col = 12: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Qty.":
            fpSpread1.Col = 13: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Return Qty.": fpSpread1.Col = 14: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Challan": fpSpread1.Col = 15: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Return Challan": fpSpread1.Col = 16: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Stock Journal IN":
            fpSpread1.Col = 17: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Stock Journal OUT": fpSpread1.Col = 18: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Pending P/O": fpSpread1.Col = 19: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Pending S/O": fpSpread1.Col = 20: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Effective Stock": fpSpread1.Col = 21: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Price":
            fpSpread1.Col = 22: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = " Amount": fpSpread1.Col = 23: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Net Purchases": fpSpread1.Col = 24: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Net Sales": fpSpread1.Col = 25: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Units": fpSpread1.Col = 26: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = " Purchases Amount":
            fpSpread1.Col = 27: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Amount": fpSpread1.Col = 28: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Purchases Return Amount": fpSpread1.Col = 29: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Sales Return Amt.": fpSpread1.Col = 30: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Net Purchases Amount":
            fpSpread1.Col = 31: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Net Sales Amount":
        ElseIf VchType >= 29 And VchType <= 30 Then
            fpSpread1.ColHeaderRows = 2:
        For C = 1 To .MaxCols
            .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "": .Col = C: fpSpread1.Row = SpreadHeader + 1: fpSpread1.FontSize = 12:
        Next
    'Header1
            fpSpread1.AddCellSpan 1, SpreadHeader, 3, 1
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader: fpSpread1.Text = " Item : " + rstStockLedger.Fields("Item").Value: fpSpread1.FontSize = 12: fpSpread1.FontBold = True
            fpSpread1.TypeHAlign = TypeHAlignCenter
            Header1 = " Item : " + rstStockLedger.Fields("Item").Value:
    'Header2
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Date": fpSpread1.Col = 2: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Vch/Bill No.":
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Vch Type":
            fpSpread1.Col = 5: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Particulars": fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Ordered Qty.":
            fpSpread1.Col = 24: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Pending": fpSpread1.Col = 25: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Units":
            fpSpread1.Col = 26: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Rate": fpSpread1.Col = 27: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Amount":
        ElseIf VchType = 31 Then
            fpSpread1.ColHeaderRows = 2:
        For C = 1 To .MaxCols
            .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "": .Col = C: fpSpread1.Row = SpreadHeader + 1: fpSpread1.FontSize = 12:
        Next
    'Header1
            fpSpread1.AddCellSpan 1, SpreadHeader, 5, 1
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader: fpSpread1.Text = " Item : " + rstStockLedger.Fields("Item").Value: fpSpread1.FontSize = 12: fpSpread1.FontBold = True
            fpSpread1.TypeHAlign = TypeHAlignCenter
            Header1 = " Item : " + rstStockLedger.Fields("Item").Value:
    'Header2
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Date":
            fpSpread1.Col = 2: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Vch/Bill No.":
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Vch Type":
            fpSpread1.Col = 5: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Particulars":
            fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "INward Qty.":
            fpSpread1.Col = 23: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Out Ward Qty.":
            fpSpread1.Col = 24: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Daily Bal.":
            fpSpread1.Col = 25: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Units":
            fpSpread1.Col = 26: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Rate":
            fpSpread1.Col = 27: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Amount":
        ElseIf VchType = 32 Or VchType = 49 Then
            fpSpread1.ColHeaderRows = 2:
            For C = 1 To .MaxCols
            .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "": .Col = C: fpSpread1.Row = SpreadHeader + 1: fpSpread1.FontSize = 12:
            Next
    'Header1
            If rstItemList.State = adStateOpen Then rstItemList.Close
            rstItemList.Open "SELECT PrintName As Item FROM BookMaster WHERE Code=" & ItemList & "", cnDatabase, adOpenKeyset, adLockReadOnly
            rstItemList.MoveFirst
                    fpSpread1.AddCellSpan 1, SpreadHeader, 5, 1
                    fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader: fpSpread1.Text = " Item : " + rstItemList.Fields("Item").Value: fpSpread1.FontSize = 12: fpSpread1.FontBold = True
                    fpSpread1.TypeHAlign = TypeHAlignCenter
                    Header1 = " Item : " + rstItemList.Fields("Item").Value:
    'Header2 rstItemOpening
            If VchType = 32 Then fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = " Material Centre ": fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Opening Qty.": fpSpread1.Col = 20: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "IN-Ward Qty.": fpSpread1.Col = 23: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Out-Ward Qty": fpSpread1.Col = 24: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Closing Qty": fpSpread1.Col = 25: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Units":
            If VchType = 49 Then fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = " Month ": fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Opening Qty.": fpSpread1.Col = 20: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "IN-Ward Qty.": fpSpread1.Col = 23: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Out-Ward Qty": fpSpread1.Col = 24: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Closing Qty": fpSpread1.Col = 25: fpSpread1.Row = SpreadHeader + 1: fpSpread1.Text = "Units":
        ElseIf VchType = 33 Then
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Item Name": fpSpread1.Col = 4: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "MRP": fpSpread1.Col = 5: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Parent Group": fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Stock Qty.": fpSpread1.Col = 7: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Units": fpSpread1.Col = 8: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Pending Quotation Qty.":
            fpSpread1.Col = 18: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Pending P/O": fpSpread1.Col = 19: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Pending S/O": fpSpread1.Col = 20: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Effective Stock": fpSpread1.Col = 21: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = "Price": fpSpread1.Col = 22: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12: fpSpread1.Text = " Amount":
            'Header1
        ElseIf VchType >= 34 And VchType <= 44 Then
            fpSpread1.ColHeaderRows = 1:
        For C = 1 To .MaxCols
            .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "": .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12:
        Next
    'Header1
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Date":
            fpSpread1.Col = 2: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Vch/Bill No.":
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Voucher Type":
            fpSpread1.Col = 5: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Particulars":
            fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Ordered Qty.":
            fpSpread1.Col = 8: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "INward Qty.":
            fpSpread1.Col = 23: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Outward Qty.":
            fpSpread1.Col = 24: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Pending Qty":
            fpSpread1.Col = 25: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Units":
            fpSpread1.Col = 26: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Rate":
            fpSpread1.Col = 27: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Amount":
        ElseIf Right(VchType, 2) = 48 Then
            fpSpread1.ColHeaderRows = 1:
        For C = 1 To .MaxCols
            .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "": .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12:
        Next
    'Header1
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Date":
            fpSpread1.Col = 2: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Vch/Bill No.":
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Vch Type":
            fpSpread1.Col = 4: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Vch Series":
            fpSpread1.Col = 5: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Particulars":
            fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "INward Qty.":
            fpSpread1.Col = 23: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Out Ward Qty.":
            fpSpread1.Col = 24: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Daily Bal.":
            fpSpread1.Col = 25: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Units":
            fpSpread1.Col = 26: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Rate":
            fpSpread1.Col = 27: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Amount":
            Mh3dLabel11.Caption = ""
            Mh3dLabel10.Caption = ""
        ElseIf VchType >= 46 And VchType <= 47 Then
            fpSpread1.ColHeaderRows = 1:
        For C = 1 To .MaxCols
            .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "": .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 11:
        Next
    'Header1
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Date"
            fpSpread1.Col = 2: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Vch/Bill No."
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Particulars"
            fpSpread1.Col = 4: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Unit Rate"
            fpSpread1.Col = 5: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Buyers Name"
            fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Ordered Qty."
            fpSpread1.Col = 7: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Unit"
            fpSpread1.Col = 8: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Pending Qty"
            fpSpread1.Col = 9: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Pending Amount"
            fpSpread1.Col = 10: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Billed Against Challan"
            fpSpread1.Col = 11: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Billed Direct"
            fpSpread1.Col = 12: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Supply Against Challan"
            fpSpread1.Col = 13: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Supply Against Billing"
            fpSpread1.Col = 14: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Clear Quantity"
            fpSpread1.Col = 15: fpSpread1.Row = SpreadHeader: fpSpread1.Text = " "
            fpSpread1.Col = 36: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Created By"
            fpSpread1.Col = 37: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Created On"
            fpSpread1.Col = 38: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Remarks"
        End If
        If VchType = 31 Then
            If Len(sMcCode) > 10 Then
                Mh3dLabel11.Caption = "Material Centre :  All"
            Else
                Mh3dLabel11.Caption = "Material Centre : " + rstStockLedger.Fields("MaterialCentre").Value
            End If
            If rstItemOpening.RecordCount <> 0 Then rstItemOpening.MoveFirst
            If rstItemOpening.RecordCount <> 0 Then Opening = Format(Val(rstItemOpening.Fields("Opening").Value), "##,##,##,##0.00")
            Mh3dLabel10.Caption = "Opening Balance :  " & Format(Opening, "##,##,##,##0.00") & IIf(Opening <= 0, " Units", " Units")
            Bal = Bal + Opening
        End If
        If VchType = 45 Then
            fpSpread1.ColHeaderRows = 1:
        For C = 1 To .MaxCols
            .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "": .Col = C: fpSpread1.Row = SpreadHeader: fpSpread1.FontSize = 12:
        Next
    'Header1
            fpSpread1.Col = 1: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Date": fpSpread1.Col = 2: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Vch/Bill No.":
            fpSpread1.Col = 3: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Vch Type":
            fpSpread1.Col = 5: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Particulars": fpSpread1.Col = 6: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Ordered Qty.":
            fpSpread1.Col = 8: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "INward Qty.":
            fpSpread1.Col = 23: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Out Ward Qty.":
            fpSpread1.Col = 24: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Pending": fpSpread1.Col = 25: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Units":
            fpSpread1.Col = 26: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Rate": fpSpread1.Col = 27: fpSpread1.Row = SpreadHeader: fpSpread1.Text = "Amount":
        End If
End With
End Function
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = 0 And KeyCode = vbKeyReturn) And VchType = 0 Then 'Enter Physical Stock
            With fpSpread1
            If .ActiveCol = 33 And fpSpread1.ActiveRow < fpSpread1.DataRowCnt Then
                fpSpread1.GetText 6, fpSpread1.ActiveRow, sysStock
                fpSpread1.GetText 33, fpSpread1.ActiveRow, phyStock
                fpSpread1.SetText 34, fpSpread1.ActiveRow, Val(phyStock) - Val(sysStock)
            End If
            End With
    End If
End Sub
Private Sub Preview_Click()
Dim PrintHeader As String
'Dim R As Long, C As Long, i As Long
'*********************************************************
If VSFlexGrid1.Visible = True Then Preview.Visible = False: Exit Sub
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
Function ClearQty(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 17) As Variant, i As Long
    Dim lpBuff As String * 1024
    On Error GoTo ErrorHandler
        With fpSpread1
If Check1.Value Then
For i = 1 To .DataRowCnt - 1
            .SetActiveCell 8, i
            .GetText 15, .ActiveRow, CellVal(17) 'CheckBox
If CellVal(17) = 1 Then
            .GetText 35, .ActiveRow, CellVal(1) 'RefCode
            .GetText 25, .ActiveRow, CellVal(2) 'VchType
            .GetText 2, .ActiveRow, CellVal(3) 'VchNo
            .GetText 1, .ActiveRow, CellVal(4) 'VchDate
            .GetText 32, .ActiveRow, CellVal(5) 'Buyer
            .GetText 34, .ActiveRow, CellVal(6): CellVal(6) = Pad(Trim(CellVal(6)), 0, 6, "L") 'Item
            .GetText 8, .ActiveRow, CellVal(7) 'Qty
            .GetText 4, .ActiveRow, CellVal(8) 'Rate
            .GetText 14, .ActiveRow, CellVal(9) 'ClearQty
            .GetText 36, .ActiveRow, CellVal(14) 'CreatedBy
            .GetText 37, .ActiveRow, CellVal(12) 'CreatedOn
            .GetText 38, .ActiveRow, CellVal(13) 'Remarks
            CellVal(12) = Format(CellVal(12), "YYYY-MM-DD hh:mm:ss")
            If ClearFlag Then CellVal(11) = "Confirm Clear Quantity !!!" Else CellVal(11) = "Confirm Retrieve Quantity !!!"
            If CellVal(12) <> "" Then CellVal(12) = CellVal(12): CellVal(16) = Format(Now(), "YYYY-MM-DD hh:mm:ss")
            If CellVal(12) = "" Then CellVal(12) = Format(Now(), "YYYY-MM-DD hh:mm:ss"): CellVal(16) = Null
            If CellVal(14) <> "" Then CellVal(14) = CellVal(14): CellVal(15) = UserCode
            If CellVal(14) = "" Then CellVal(14) = UserCode: CellVal(15) = Null
            
            If ClearFlag Then
                CellVal(10) = InputBox(CellVal(11), , CellVal(7))
                If Val(CellVal(10)) > CellVal(7) Then SCode = 0: Exit Function
            Else
                CellVal(10) = InputBox(CellVal(11), , CellVal(9))
                If Val(CellVal(10)) > CellVal(9) Then SCode = 0: Exit Function
            End If
                CellVal(13) = InputBox("Confirm Remarks !!!", , CellVal(13))
            If ClearFlag Then
                CellVal(10) = CellVal(10) + CellVal(9)
                SCode = Val(SCode) + Val(CellVal(10))
            Else
                SCode = Val(SCode) + Val(CellVal(10))
                CellVal(10) = Abs(CellVal(9) - Val(CellVal(10)))
            End If
            
            GetComputerName lpBuff, Len(lpBuff)
        cnDatabase.Execute "IF EXISTS (SELECT ORDINAL_POSITION FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'JobworkBVClear' AND COLUMN_Name ='RefCode') Print 'Col_Exist' ELSE CREATE TABLE JobworkBVClear([RefCode] [nvarchar](6) NOT NULL,[VchType] [nvarchar](6) NOT NULL,[VchNo] [nvarchar](25) NOT NULL,[VchDate] [datetime] NOT NULL,[Party] [nvarchar](6) NOT NULL,[Item] [nvarchar](6) NOT NULL,[Quantity] [decimal](12, 0) NOT NULL,[Rate] [decimal](12, 2) NOT NULL,[Remarks] [nvarchar](100) NULL,[CreatedBy] [nvarchar](6) NOT NULL ,[CreatedOn] [datetime] NOT NULL ,[ModifiedBy] [nvarchar](6) NULL,[ModifiedOn] [datetime] NULL,[ComputerName] [nvarchar](40) NULL) ON [PRIMARY]"
        cnDatabase.Execute "IF EXISTS (SELECT ORDINAL_POSITION FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'JobworkBVClear' AND COLUMN_Name ='Remarks') Print 'Col_Exist' ELSE ALTER TABLE JobworkBVClear ADD  [Remarks] [nvarchar](100) NULL,[CreatedBy] [nvarchar](6) NOT NULL Default('000001'),[CreatedOn] [datetime] NOT NULL Default('" & CellVal(12) & "' ) ,[ModifiedBy] [nvarchar](6) NULL,[ModifiedOn] [datetime] NULL,[ComputerName] [nvarchar](40) NULL"
    If ClearFlag = True Then
        cnDatabase.Execute "INSERT INTO JobworkBVClear VALUES ('" & CellVal(1) & "','" & CellVal(2) & "','" & CellVal(3) & "','" & Format(CellVal(4), "dd-MMM-yyyy") & "','" & CellVal(5) & "','" & CellVal(6) & "'," & Val(CellVal(10)) & "," & Val(CellVal(8)) & ",'" & CellVal(13) & "' ,'" & CellVal(14) & "' ,'" & CellVal(12) & "','" & CellVal(15) & "' ,'" & CellVal(16) & "','" & Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1) & "' )"
        cnDatabase.Execute "DELETE FROM JobworkBVClear WHERE RefCode='" & CellVal(1) & "' AND VchType='" & CellVal(2) & "' AND Quantity='" & CellVal(9) & "'"
        .SetText 8, .ActiveRow, Abs(CellVal(7) + CellVal(9) - CellVal(10)) ' Qty
        .SetText 14, .ActiveRow, CellVal(10) 'Clear Qty
        .SetText 38, .ActiveRow, CellVal(13) 'Remarks
    End If
    If unClearFlag = True Then
        cnDatabase.Execute "DELETE FROM JobworkBVClear WHERE RefCode='" & CellVal(1) & "' AND VchType='" & CellVal(2) & "' AND Quantity='" & CellVal(9) & "'"
            .SetText 8, .ActiveRow, CellVal(7) + CellVal(9) - CellVal(10) ' Qty
            .SetText 14, .ActiveRow, CellVal(10) 'Clear Qty
            .SetText 38, .ActiveRow, CellVal(13) 'Remarks
        cnDatabase.Execute "INSERT INTO JobworkBVClear VALUES ('" & CellVal(1) & "','" & CellVal(2) & "','" & CellVal(3) & "','" & Format(CellVal(4), "dd-MMM-yyyy") & "','" & CellVal(5) & "','" & CellVal(6) & "'," & Val(CellVal(10)) & "," & Val(CellVal(8)) & ",'" & CellVal(13) & "' ,'" & CellVal(14) & "' ,'" & CellVal(12) & "','" & CellVal(15) & "' ,'" & CellVal(16) & "','" & Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1) & "' )"
    End If
   End If
            Next
            ClearFlag = False
            unClearFlag = False
Else
            .GetText 35, .ActiveRow, CellVal(1) 'RefCode
            .GetText 25, .ActiveRow, CellVal(2) 'VchType
            .GetText 2, .ActiveRow, CellVal(3) 'VchNo
            .GetText 1, .ActiveRow, CellVal(4) 'VchDate
            .GetText 32, .ActiveRow, CellVal(5) 'Buyer
            .GetText 34, .ActiveRow, CellVal(6): CellVal(6) = Pad(Trim(CellVal(6)), 0, 6, "L") 'Item
            .GetText 8, .ActiveRow, CellVal(7) 'Qty
            .GetText 4, .ActiveRow, CellVal(8) 'Rate
            .GetText 14, .ActiveRow, CellVal(9) 'ClearQty
            .GetText 36, .ActiveRow, CellVal(14) 'CreatedBy
            .GetText 37, .ActiveRow, CellVal(12) 'CreatedOn
            .GetText 38, .ActiveRow, CellVal(13) 'Remarks
            CellVal(12) = Format(CellVal(12), "YYYY-MM-DD hh:mm:ss")
            If ClearFlag Then CellVal(11) = "Confirm Clear Quantity !!!" Else CellVal(11) = "Confirm Retrieve Quantity !!!"
            If CellVal(12) <> "" Then CellVal(12) = CellVal(12): CellVal(16) = Format(Now(), "YYYY-MM-DD hh:mm:ss")
            If CellVal(12) = "" Then CellVal(12) = Format(Now(), "YYYY-MM-DD hh:mm:ss"): CellVal(16) = Null
            If CellVal(14) <> "" Then CellVal(14) = CellVal(14): CellVal(15) = UserCode
            If CellVal(14) = "" Then CellVal(14) = UserCode: CellVal(15) = Null
            
            If ClearFlag Then
                CellVal(10) = InputBox(CellVal(11), , CellVal(7))
                If Val(CellVal(10)) > CellVal(7) Then SCode = 0: Exit Function
            Else
                CellVal(10) = InputBox(CellVal(11), , CellVal(9))
                If Val(CellVal(10)) > CellVal(9) Then SCode = 0: Exit Function
            End If
                CellVal(13) = InputBox("Confirm Remarks !!!", , CellVal(13))
            If ClearFlag Then
                CellVal(10) = CellVal(10) + CellVal(9)
                SCode = CellVal(10)
            Else
                SCode = CellVal(10)
                CellVal(10) = Abs(CellVal(9) - Val(CellVal(10)))
            End If
            
            GetComputerName lpBuff, Len(lpBuff)
        cnDatabase.Execute "IF EXISTS (SELECT ORDINAL_POSITION FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'JobworkBVClear' AND COLUMN_Name ='RefCode') Print 'Col_Exist' ELSE CREATE TABLE JobworkBVClear([RefCode] [nvarchar](6) NOT NULL,[VchType] [nvarchar](6) NOT NULL,[VchNo] [nvarchar](25) NOT NULL,[VchDate] [datetime] NOT NULL,[Party] [nvarchar](6) NOT NULL,[Item] [nvarchar](6) NOT NULL,[Quantity] [decimal](12, 0) NOT NULL,[Rate] [decimal](12, 2) NOT NULL,[Remarks] [nvarchar](100) NULL,[CreatedBy] [nvarchar](6) NOT NULL ,[CreatedOn] [datetime] NOT NULL ,[ModifiedBy] [nvarchar](6) NULL,[ModifiedOn] [datetime] NULL,[ComputerName] [nvarchar](40) NULL) ON [PRIMARY]"
        cnDatabase.Execute "IF EXISTS (SELECT ORDINAL_POSITION FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'JobworkBVClear' AND COLUMN_Name ='Remarks') Print 'Col_Exist' ELSE ALTER TABLE JobworkBVClear ADD  [Remarks] [nvarchar](100) NULL,[CreatedBy] [nvarchar](6) NOT NULL Default('000001'),[CreatedOn] [datetime] NOT NULL Default('" & CellVal(12) & "' ) ,[ModifiedBy] [nvarchar](6) NULL,[ModifiedOn] [datetime] NULL,[ComputerName] [nvarchar](40) NULL"
    If ClearFlag = True Then
        cnDatabase.Execute "INSERT INTO JobworkBVClear VALUES ('" & CellVal(1) & "','" & CellVal(2) & "','" & CellVal(3) & "','" & Format(CellVal(4), "dd-MMM-yyyy") & "','" & CellVal(5) & "','" & CellVal(6) & "'," & Val(CellVal(10)) & "," & Val(CellVal(8)) & ",'" & CellVal(13) & "' ,'" & CellVal(14) & "' ,'" & CellVal(12) & "','" & CellVal(15) & "' ,'" & CellVal(16) & "','" & Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1) & "' )"
        cnDatabase.Execute "DELETE FROM JobworkBVClear WHERE RefCode='" & CellVal(1) & "' AND VchType='" & CellVal(2) & "' AND Quantity='" & CellVal(9) & "'"
        .SetText 8, .ActiveRow, Abs(CellVal(7) + CellVal(9) - CellVal(10)) ' Qty
        .SetText 14, .ActiveRow, CellVal(10) 'Clear Qty
        .SetText 38, .ActiveRow, CellVal(13) 'Remarks
        ClearFlag = False
    End If
    If unClearFlag = True Then
        cnDatabase.Execute "DELETE FROM JobworkBVClear WHERE RefCode='" & CellVal(1) & "' AND VchType='" & CellVal(2) & "' AND Quantity='" & CellVal(9) & "'"
            .SetText 8, .ActiveRow, CellVal(7) + CellVal(9) - CellVal(10) ' Qty
            .SetText 14, .ActiveRow, CellVal(10) 'Clear Qty
            .SetText 38, .ActiveRow, CellVal(13) 'Remarks
        cnDatabase.Execute "INSERT INTO JobworkBVClear VALUES ('" & CellVal(1) & "','" & CellVal(2) & "','" & CellVal(3) & "','" & Format(CellVal(4), "dd-MMM-yyyy") & "','" & CellVal(5) & "','" & CellVal(6) & "'," & Val(CellVal(10)) & "," & Val(CellVal(8)) & ",'" & CellVal(13) & "' ,'" & CellVal(14) & "' ,'" & CellVal(12) & "','" & CellVal(15) & "' ,'" & CellVal(16) & "','" & Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1) & "' )"
            unClearFlag = False
    End If

End If
            .SetActiveCell 8, .ActiveRow
            Check1.Value = 0
        End With
    Exit Function
ErrorHandler:
    ClearFlag = False
    unClearFlag = False
    ClearQty = False
End Function
Private Sub Mh3dLabel5_Click()
Dim PrintHeader As String
Dim R As Long, C As Long
Dim JQty As Variant
    On Error GoTo ErrHandler
Screen.MousePointer = vbHourglass

Const PaperWidth = 12240
Const PaperHeight = 15840

PrintHeader = "Export Data Company : " & rstCompanyMaster.Fields("PrintName").Value & " _(" & CompCode & "_" & PrintHeader & ")" & "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]" & " /rPage # ./p " & " Print Date : ( " & Format(Date, "dd-MMM-yyyy") & " )         "
If VSFlexFlag = True Then
    With Me.VSFlexGrid1
    .PrintGrid PrintHeader, True, PrintOrientationLandscape, 50, 300
 
    End With
Else
With fpSpread1
.MaxRows = .MaxRows + 2
    If VchType >= 0 Then fpSpread1.InsertRows 1, 2
    .SetText 5, 1, rstCompanyMaster.Fields("PrintName").Value: .Col = 5: .Row = 1: .FontBold = True: .FontSize = 20: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = RGB(1, 106, 106): .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: '.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
    .SetText 5, 2, "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]": .Col = 5: .Row = 2: .FontBold = True: .FontSize = 16: .BackColor = &H8000000F:  .ForeColor = RGB(20, 106, 106): .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter:  '.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
    R = 1
For C = 1 To .MaxCols
'    .Col = C: .Row = R: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: '.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
 '   .GetText C, 0, JQty
    '.SetText C, 1, JQty
Next


PrintHeader = Me.Caption
.LockBackColor = vbWhite
' These are 8.5" X 11" paper dimensions in TWIPS
Printer.PaperSize = vbPRPSA4
' Set printing options for sheet
fpSpread1.PrintAbortMsg = "Printing - Click Cancel to .Quit"
fpSpread1.PrintJobName = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") '& ".pdf"
'fpSpread1.PrintHeader = "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy"): fpSpread1.PrintHeader=: .Font = 20 '"/cPrint Header/rPage # ./p/n2nd Line"
fpSpread1.PrintFooter = "        Export Data Company : " & rstCompanyMaster.Fields("PrintName").Value & " _(" & CompCode & "_" & PrintHeader & ")" & "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]" & " /rPage # ./p " & " Print Date : ( " & Format(Date, "dd-MMM-yyyy") & " )         ": .FontSize = 16 '& ".pdf" ' "/cPrint Footer/rPage # ./p/n2nd Line"
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
    'Delete Row
    If VchType >= 0 Then fpSpread1.DeleteRows 1, 2
    .MaxRows = .MaxRows - 2

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
Dim JQty As Variant
    On Error GoTo ErrHandler
'''''With Me.VSFlexGrid1
'''''    Set .DataSource = Nothing
'''''    .LoadGrid App.Path & "\Customers.xls", flexFileExcel
'''''End With

'"Export Data" &
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
    If (VchType <= 10 And VchType >= 7) Or (VchType <= 28 And VchType >= 25) Or (VchType >= 57 And VchType <= 60) Then fpSpread1.InsertRows 1, 2 Else fpSpread1.InsertRows 1, 1
                    R = 1
                For C = 1 To .MaxCols
                    .Col = C: .Row = R: .FontBold = True: .FontSize = 10: .BackColor = &H8000000F: .FontUnderline = True: .ForeColor = vbBlue: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: '.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
                    .GetText C, 0, JQty
                    .SetText C, 1, JQty
                Next
'                    .SetText 1, 1, "Details": .SetText 2, 1, "MRP": .SetText 3, 1, "Parent Group": .SetText 4, 1, "Stock Qty.": .SetText 5, 1, "Units": .SetText 6, 1, "Purchases Qty.": .SetText 7, 1, "Purchases Return Qty.": .SetText 8, 1, "Purchases Challan": .SetText 9, 1, "Purchases Return Challan": .SetText 10, 1, "Sales Qty.": .SetText 11, 1, "Sales Return Qty.": .SetText 12, 1, "Sales Challan": .SetText 13, 1, "Sales Return Challan": .SetText 14, 1, "Stock Journal IN": .SetText 15, 1, "Stock Journal OUT": .SetText 16, 1, "Pending P/O": .SetText 17, 1, "Pending S/O": .SetText 18, 1, "Effective Stock": .SetText 19, 1, "Price": .SetText 20, 1, " Amount": .SetText 21, 1, "Net Purchases": .SetText 22, 1, "Net Sales": .SetText 23, 1, "Units": .SetText 24, 1, " Purchases Amount": .SetText 25, 1, "Sales Amount": .SetText 26, 1, "Purchases Return Amount": .SetText 27, 1, "Sales Return Amt.": .SetText 28, 1, "Net Purchases Amount": .SetText 29, 1, "Net Sales Amount": .SetText 30, 1, "ICODE":
'                    If VchType = 0 Then: .SetText 31, 1, "Physical Stock Quantity": .SetText 32, 1, "Stock Impact":
                .ColHeadersShow = True: .PrintColHeaders = True: .PrintRowHeaders = True: .ColHeadersShow = True: .RowHeadersShow = True: .GridShowHoriz = True: .GridShowVert = True
                If (VchType <= 10 And VchType >= 7) Or (VchType <= 28 And VchType >= 25) Or (VchType >= 57 And VchType <= 60) Then .SetText 1, 2, Header1: .Col = 1: .Row = 2: .FontBold = True: .FontSize = 14: .FontUnderline = True: .ForeColor = vbRed:
    
    End With
    If Dir(App.Path & "\Export", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Export"
    '
     FileName = App.Path & "\Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"
    SheetName = "Sheet1" '"(" & Me.Caption & ")"
    LogFileName = "Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".txt"
    x = fpSpread1.ExportToExcelEx(FileName, SheetName, LogFileName, ExcelSaveFlagNoFormulas)
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
    If (VchType <= 10 And VchType >= 7) Or (VchType <= 28 And VchType >= 25) Or (VchType >= 57 And VchType <= 60) Then fpSpread1.DeleteRows 1, 2 Else: fpSpread1.DeleteRows 1, 1
    End With
End If
Screen.MousePointer = vbNormal
Exit Sub

ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
Private Function VSFlexGrid_Format_Headers()
Dim i As Long
Dim C As Long
    On Error GoTo ErrHandler
    With VSFlexGrid1
        .ColWidth(0) = 250
    If VchType = 0 Then 'Physical Stock Audit Ledger
        .TextMatrix(i, 1) = ""
        .ColHidden(1) = True
        .TextMatrix(i, 2) = ""
        .ColHidden(2) = True
        .TextMatrix(i, 3) = "Item Name"
        .TextMatrix(i, 4) = "MRP"
        .TextMatrix(i, 5) = "Parent Group"
        .TextMatrix(i, 6) = "Stock Qty."
        .TextMatrix(i, 7) = "Units"
        For C = 3 To 7
        .ColHidden(C) = False
        Next
        .ColWidth(3) = 49.25 'Item
        .ColWidth(4) = 15 'MRP
        .ColWidth(5) = 15 'Item Group
        .ColWidth(6) = 14  'Stock Qty.
        .ColWidth(7) = 11.5 'Unit
        For C = 8 To 32
        .ColHidden(C) = True
        Next
        For C = 33 To 34
        .ColHidden(C) = False
        Next
        .ColWidth(33) = 24 'Phical Stock Qty.
        .ColWidth(34) = 22.75 'Stock Impact
        .ColHidden(35) = True
        .TextMatrix(i, 33) = "Physical Stock Quantity"
        .TextMatrix(i, 34) = "Stock Impact"
        .TextMatrix(i, 35) = "HSNCODE"
    ElseIf VchType <= 2 Or VchType = 33 Then
'        .ColWidth(0) = 500
        .TextMatrix(i, 1) = ""
        .ColHidden(1) = True
        .TextMatrix(i, 2) = ""
        .ColHidden(2) = True
        .TextMatrix(i, 3) = "Item Name"
        .TextMatrix(i, 4) = "MRP"
        .TextMatrix(i, 5) = "Parent Group"
        .TextMatrix(i, 6) = "Stock Qty."
        .TextMatrix(i, 7) = "Units"
         If VchType <= 2 And VchType <> 1 Then
            For C = 8 To 17
                    .ColHidden(C) = True
             Next
             ElseIf VchType = 33 And VchType <> 1 Then
             For C = 9 To 17
                    .ColHidden(C) = True
             Next
                    .TextMatrix(i, 8) = "Pending Quotation"
         End If
         If VchType = 1 Then .TextMatrix(i, 8) = "Purchases Qty."
        .TextMatrix(i, 9) = "Purchases Return Qty."
        .TextMatrix(i, 10) = "Purchases Challan"
        .TextMatrix(i, 11) = "Purchases Return Challan"
        .TextMatrix(i, 12) = "Sales Qty."
        .TextMatrix(i, 13) = "Sales Return Qty."
        .TextMatrix(i, 14) = "Sales Challan"
        .TextMatrix(i, 15) = "Sales Return Challan"
        .TextMatrix(i, 16) = "Stock Journal IN"
        .TextMatrix(i, 17) = "Stock Journal OUT"
        .TextMatrix(i, 18) = "Pending P/O"
        .TextMatrix(i, 19) = "Pending S/O"
        .TextMatrix(i, 20) = "Effective Stock"
        .TextMatrix(i, 21) = "Price"
        .ColHidden(21) = True
        .TextMatrix(i, 22) = " Amount"
        .TextMatrix(i, 23) = "Net Purchases"
        .TextMatrix(i, 24) = "Net Sales"
        .TextMatrix(i, 25) = "Units"
        .ColHidden(25) = True
        .TextMatrix(i, 26) = " Purchases Amount"
        .TextMatrix(i, 27) = "Sales Amount"
        .TextMatrix(i, 28) = "Purchases Return Amount"
        .TextMatrix(i, 29) = "Sales Return Amt."
        .TextMatrix(i, 30) = "Net Purchases Amount"
        .TextMatrix(i, 31) = "Net Sales Amount"
        .TextMatrix(i, 32) = "Code"
        .TextMatrix(i, 33) = "Physical Stock Quantity"
        .TextMatrix(i, 34) = "Stock Impact"
        .TextMatrix(i, 35) = "HSNCODE"
        If VchType <> 1 Then
            For C = 23 To 35
                    .ColHidden(C) = True
             Next
        ElseIf VchType = 1 Then
            For C = 32 To 35
                    .ColHidden(C) = True
             Next
        End If
    ElseIf VchType > 101 Then
            If (Combo1.ListIndex = 0 Or Combo1.ListIndex = 1) And VchType = 105 Then
'                                 .ColWidth(0) = 500
                C = C + 1: .TextMatrix(i, C) = "Item"
                C = C + 1: .TextMatrix(i, C) = "Binding"
                C = C + 1: .TextMatrix(i, C) = "FinishSize"
                C = C + 1: .TextMatrix(i, C) = "Pages"
                .ColHidden(C) = True
                C = C + 1: .TextMatrix(i, C) = "Price"
                C = C + 1: .TextMatrix(i, C) = "WIP"
            ElseIf (Combo1.ListIndex = 2 Or Combo1.ListIndex = 3) And VchType = 105 Then
                C = C + 1: .TextMatrix(i, C) = "WIP"
                C = C + 1: .TextMatrix(i, C) = "Item"
                C = C + 1: .TextMatrix(i, C) = "FinishSize"
                C = C + 1: .TextMatrix(i, C) = "Pages"
                C = C + 1: .TextMatrix(i, C) = "Price"
            ElseIf (Combo1.ListIndex = 4 Or Combo1.ListIndex = 5) And VchType = 105 Then
                C = C + 1: .TextMatrix(i, C) = "RM"
                C = C + 1: .TextMatrix(i, C) = "WIP"
                C = C + 1: .TextMatrix(i, C) = "Item"
                C = C + 1: .TextMatrix(i, C) = "FinishSize"
                C = C + 1: .TextMatrix(i, C) = "Pages"
                C = C + 1: .TextMatrix(i, C) = "Price"
            End If

    If VchType = 103 Then C = C + 1: .TextMatrix(i, C) = "WIP"
                If VchType = 103 Or VchType = 105 Then
                        C = C + 1: .TextMatrix(i, C) = "WIP/UNIT"
                        If FrmItemSelectionList.Check1.Value Then .ColHidden(C) = True
                        C = C + 1: .TextMatrix(i, C) = "WIP Pages"
                        If FrmItemSelectionList.Check1.Value Then .ColHidden(C) = True
                        C = C + 1: .TextMatrix(i, C) = "Color"
                        If FrmItemSelectionList.Check1.Value Then .ColHidden(C) = True
                        If FrmItemSelectionList.Check1.Value = False And VchType <> 105 Then C = C + 1: .TextMatrix(i, C) = "FG Name(Actual BarCode)"
                        C = C + 1: .TextMatrix(i, C) = "Stock"
                        C = C + 1: .TextMatrix(i, C) = "Sales Order"
                        C = C + 1: .TextMatrix(i, C) = "Dispatched"
                        C = C + 1: .TextMatrix(i, C) = "Pending SO"
                    If VchType = 105 Then
                        C = C + 1: .TextMatrix(i, C) = "Rate"
                        C = C + 1: .TextMatrix(i, C) = "Amount"
                    End If
                        C = C + 1: .TextMatrix(i, C) = "Deficient Order"
                        C = C + 1: .TextMatrix(i, C) = "WIP Required"
                        C = C + 1: .TextMatrix(i, C) = "WIP Stock"
                        C = C + 1: .TextMatrix(i, C) = "Final WIP Required"
            End If
                    If VchType = 104 Or VchType = 105 Then
                        C = C + 1: .TextMatrix(i, C) = "RM"
                    End If
                If VchType = 104 Or VchType = 105 Then
                        C = C + 1: .TextMatrix(i, C) = "RM Make"
                        C = C + 1: .TextMatrix(i, C) = "RM GSM"
                        C = C + 1: .TextMatrix(i, C) = "RM CUT-OFF"
                        C = C + 1: .TextMatrix(i, C) = "RM/UNIT"
                        C = C + 1: .TextMatrix(i, C) = "Weight/Unit"
                        C = C + 1: .TextMatrix(i, C) = "Unit Name"
                        C = C + 1: .TextMatrix(i, C) = "Qty/Unit"
                        C = C + 1: .TextMatrix(i, C) = "RM Req IN Sheets"
                        C = C + 1: .TextMatrix(i, C) = "RM Req IN Kgs"
                        C = C + 1: .TextMatrix(i, C) = "RM Stock UOM"
                        C = C + 1: .TextMatrix(i, C) = "RM Stock IN Kgs"
                        C = C + 1: .TextMatrix(i, C) = "Final RM Req Kgs"
                End If
    End If
    End With
        Mh3dLabel11.Caption = ""
        Mh3dLabel10.Caption = ""

Screen.MousePointer = vbNormal
Exit Function
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Function
Private Function VSFlexGrid_Format_Cols_Headers()
Dim C As Long
With VSFlexGrid1
End With
End Function
Private Function PublishGrid()
Dim i, Stock, StockTotal, PurchaseTotal, PurchaseReturnTotal, PurchaseChallanTotal, PurchaseReturnChallanTotal, SalesTotal, SalesReturnTotal, SalesChallanTotal, SalesReturnChallanTotal, StockJournalINTotal, StockJournalOUTTotal, POTotal, SOTotal, EffectiveStock As Long, NetPurchaseTotal, NetSalesTotal, EStockTotal As Long
Dim AmountTotal, PurchaseAmountTotal, SalesAmountTotal, PurchaseReturnAmountTotal, SalesReturnAmountTotal, NetPurchaseAmountTotal, NetSalesAmountTotal As Double
Dim dPrint As Long
Dim C As Long
On Error GoTo ErrHandler

With VSFlexGrid1
    .Clear
    Zoom.Visible = True
    If VchType <= 2 Or VchType = 33 Then
        .Cols = 36
        .Rows = rstStockLedger.RecordCount + 1
        rstStockLedger.MoveFirst
        Do While Not rstStockLedger.EOF
                If PendingCheck.Value Then
                    If Val(rstStockLedger.Fields("PendingPO").Value) = 0 And Val(rstStockLedger.Fields("PendingSO").Value) = 0 Then GoTo NXT
                End If
                    Stock = Val(rstStockLedger.Fields("PurchaseChallan").Value) - Val(rstStockLedger.Fields("PurchaseReturnChallan").Value) - Val(rstStockLedger.Fields("SalesChallan").Value) + Val(rstStockLedger.Fields("SalesReturnChallan").Value) + Val(rstStockLedger.Fields("Purchase").Value) - Val(rstStockLedger.Fields("PurchaseReturn").Value) - Val(rstStockLedger.Fields("Sales").Value) + Val(rstStockLedger.Fields("SalesReturn").Value) + Val(rstStockLedger.Fields("StockJournalIN").Value) - Val(rstStockLedger.Fields("StockJournalOUT").Value) + Val(rstStockLedger.Fields("StockTransferIN").Value) - Val(rstStockLedger.Fields("StockTransferOUT").Value)
                If VchType <= 2 Then EffectiveStock = Stock + Val(rstStockLedger.Fields("PendingPO").Value) - Val(rstStockLedger.Fields("PendingSO").Value)
                If VchType = 33 Then EffectiveStock = Stock + Val(rstStockLedger.Fields("PendingPO").Value) - Val(rstStockLedger.Fields("PendingSO").Value) - Val(rstStockLedger.Fields("SQ").Value)
                If NegativeStock.Value Then
                    If EffectiveStock >= 0 Then GoTo NXT
                End If
                If ZeroStock.Value Then
                    If TDBNumber1.Value = 0 Then
                        If EffectiveStock <> TDBNumber1.Value Then GoTo NXT
                    Else
                        If EffectiveStock >= TDBNumber1.Value Then GoTo NXT
                    End If
                End If
        
        
        i = i + 1
                .TextMatrix(i, 0) = i
            If VchType <= 2 Then
                .TextMatrix(i, 1) = ""
                .TextMatrix(i, 3) = rstStockLedger.Fields("Item").Value
                .TextMatrix(i, 4) = Format(Val(rstStockLedger.Fields("MRP").Value), "###0.00")
                .TextMatrix(i, 5) = rstStockLedger.Fields("ItemGroup").Value
                .TextMatrix(i, 6) = Stock + Val(rstStockLedger.Fields("OPBAL").Value): If .TextMatrix(i, 6) < 0 Then .Cell(flexcpForeColor, i, 6) = vbRed Else .Cell(flexcpForeColor, i, 6) = vbBlack
                .TextMatrix(i, 7) = "Units"
                .TextMatrix(i, 8) = Val(rstStockLedger.Fields("Purchase").Value): If .TextMatrix(i, 8) < 0 Then .Cell(flexcpForeColor, i, 8) = vbRed Else .Cell(flexcpForeColor, i, 8) = vbBlack
                        PurchaseTotal = PurchaseTotal + Val(rstStockLedger.Fields("Purchase").Value)
                .TextMatrix(i, 9) = Val(rstStockLedger.Fields("PurchaseReturn").Value): If .TextMatrix(i, 9) < 0 Then .Cell(flexcpForeColor, i, 9) = vbRed Else .Cell(flexcpForeColor, i, 9) = vbBlack
                        PurchaseReturnTotal = PurchaseReturnTotal + Val(rstStockLedger.Fields("PurchaseReturn").Value)
                .TextMatrix(i, 10) = Val(rstStockLedger.Fields("PurchaseChallan").Value): If .TextMatrix(i, 10) < 0 Then .Cell(flexcpForeColor, i, 10) = vbRed Else .Cell(flexcpForeColor, i, 10) = vbBlack
                        PurchaseChallanTotal = PurchaseChallanTotal + Val(rstStockLedger.Fields("PurchaseChallan").Value)
                .TextMatrix(i, 11) = Val(rstStockLedger.Fields("PurchaseReturnChallan").Value): If .TextMatrix(i, 11) < 0 Then .Cell(flexcpForeColor, i, 11) = vbRed Else .Cell(flexcpForeColor, i, 11) = vbBlack
                        PurchaseReturnChallanTotal = PurchaseReturnChallanTotal + Val(rstStockLedger.Fields("PurchaseReturnChallan").Value)
                .TextMatrix(i, 12) = Val(rstStockLedger.Fields("Sales").Value): If .TextMatrix(i, 12) < 0 Then .Cell(flexcpForeColor, i, 12) = vbRed Else .Cell(flexcpForeColor, i, 12) = vbBlack
                        SalesTotal = SalesTotal + Val(rstStockLedger.Fields("Sales").Value)
                .TextMatrix(i, 13) = Val(rstStockLedger.Fields("SalesReturn").Value): If .TextMatrix(i, 13) < 0 Then .Cell(flexcpForeColor, i, 13) = vbRed Else .Cell(flexcpForeColor, i, 13) = vbBlack
                        SalesReturnTotal = SalesReturnTotal + Val(rstStockLedger.Fields("SalesReturn").Value)
                .TextMatrix(i, 14) = Val(rstStockLedger.Fields("SalesChallan").Value): If .TextMatrix(i, 14) < 0 Then .Cell(flexcpForeColor, i, 14) = vbRed Else .Cell(flexcpForeColor, i, 14) = vbBlack
                        SalesChallanTotal = SalesChallanTotal + Val(rstStockLedger.Fields("SalesChallan").Value)
                .TextMatrix(i, 15) = Val(rstStockLedger.Fields("SalesReturnChallan").Value): If .TextMatrix(i, 15) < 0 Then .Cell(flexcpForeColor, i, 15) = vbRed Else .Cell(flexcpForeColor, i, 15) = vbBlack
                        SalesReturnChallanTotal = SalesReturnChallanTotal + Val(rstStockLedger.Fields("SalesReturnChallan").Value)
                .TextMatrix(i, 16) = Val(rstStockLedger.Fields("StockJournalIN").Value): If .TextMatrix(i, 16) < 0 Then .Cell(flexcpForeColor, i, 16) = vbRed Else .Cell(flexcpForeColor, i, 16) = vbBlack
                        StockJournalINTotal = StockJournalINTotal + Val(rstStockLedger.Fields("StockJournalIN").Value)
                .TextMatrix(i, 17) = Val(rstStockLedger.Fields("StockJournalOUT").Value): If .TextMatrix(i, 17) < 0 Then .Cell(flexcpForeColor, i, 17) = vbRed Else .Cell(flexcpForeColor, i, 17) = vbBlack
                        StockJournalOUTTotal = StockJournalOUTTotal + Val(rstStockLedger.Fields("StockJournalOUT").Value)
                .TextMatrix(i, 18) = Val(rstStockLedger.Fields("PendingPO").Value): If .TextMatrix(i, 18) < 0 Then .Cell(flexcpForeColor, i, 18) = vbRed Else .Cell(flexcpForeColor, i, 18) = vbBlack
                        POTotal = POTotal + Val(rstStockLedger.Fields("PendingPO").Value)
                .TextMatrix(i, 19) = Val(rstStockLedger.Fields("PendingSO").Value): If .TextMatrix(i, 19) < 0 Then .Cell(flexcpForeColor, i, 19) = vbRed Else .Cell(flexcpForeColor, i, 19) = vbBlack
                        SOTotal = SOTotal + Val(rstStockLedger.Fields("PendingSO").Value)
                .TextMatrix(i, 20) = EffectiveStock: If .TextMatrix(i, 20) < 0 Then .Cell(flexcpForeColor, i, 20) = vbRed Else .Cell(flexcpForeColor, i, 20) = vbBlack
                .TextMatrix(i, 21) = Val(rstStockLedger.Fields("MRP").Value)
                .TextMatrix(i, 22) = EffectiveStock * Val(rstStockLedger.Fields("MRP").Value): If .TextMatrix(i, 22) < 0 Then .Cell(flexcpForeColor, i, 22) = vbRed Else .Cell(flexcpForeColor, i, 22) = vbBlack
                        AmountTotal = AmountTotal + EffectiveStock * Val(rstStockLedger.Fields("MRP").Value)
                .TextMatrix(i, 23) = Val(rstStockLedger.Fields("Purchase").Value) - Val(rstStockLedger.Fields("PurchaseReturn").Value)
                        NetPurchaseTotal = NetPurchaseTotal + Val(rstStockLedger.Fields("Purchase").Value) - Val(rstStockLedger.Fields("PurchaseReturn").Value)
                .TextMatrix(i, 24) = Val(rstStockLedger.Fields("Sales").Value) - Val(rstStockLedger.Fields("SalesReturn").Value)
                        NetSalesTotal = NetSalesTotal + Val(rstStockLedger.Fields("Sales").Value) - Val(rstStockLedger.Fields("SalesReturn").Value)
                .TextMatrix(i, 25) = "Units"
                .TextMatrix(i, 26) = Val(rstStockLedger.Fields("PurchaseAmount").Value)
                        PurchaseAmountTotal = PurchaseAmountTotal + Val(rstStockLedger.Fields("PurchaseAmount").Value)
                .TextMatrix(i, 27) = Val(rstStockLedger.Fields("SalesAmount").Value)
                        SalesAmountTotal = SalesAmountTotal + Val(rstStockLedger.Fields("SalesAmount").Value)
                .TextMatrix(i, 28) = Val(rstStockLedger.Fields("PurchaseReturnAmount").Value)
                        PurchaseReturnAmountTotal = PurchaseReturnAmountTotal + Val(rstStockLedger.Fields("PurchaseReturnAmount").Value)
                .TextMatrix(i, 29) = Val(rstStockLedger.Fields("SalesReturnAmount").Value)
                        SalesReturnAmountTotal = SalesReturnAmountTotal + Val(rstStockLedger.Fields("SalesReturnAmount").Value)
                .TextMatrix(i, 30) = Val(rstStockLedger.Fields("PurchaseAmount").Value) - Val(rstStockLedger.Fields("PurchaseReturnAmount").Value)
                        NetPurchaseAmountTotal = NetPurchaseAmountTotal + Val(rstStockLedger.Fields("PurchaseAmount").Value) - Val(rstStockLedger.Fields("PurchaseReturnAmount").Value)
                .TextMatrix(i, 31) = Val(rstStockLedger.Fields("SalesAmount").Value) - Val(rstStockLedger.Fields("SalesReturnAmount").Value)
                .TextMatrix(i, 32) = (rstStockLedger.Fields("Code").Value)
'                .TextMatrix(i, 33) = ""
'                .TextMatrix(i, 34) = ""
                .TextMatrix(i, 35) = rstStockLedger.Fields("HSNCode").Value
                NetSalesAmountTotal = NetSalesAmountTotal + Val(rstStockLedger.Fields("SalesAmount").Value) - Val(rstStockLedger.Fields("SalesReturnAmount").Value)
                StockTotal = StockTotal + Stock
                EStockTotal = EStockTotal + EffectiveStock
                dPrint = dPrint + 1
                MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
'Pending Quotations and Short Item Analysis
        ElseIf VchType = 33 Then
            .TextMatrix(i, 3) = rstStockLedger.Fields("Item").Value
            .TextMatrix(i, 4) = Val(rstStockLedger.Fields("MRP").Value)
            .TextMatrix(i, 5) = rstStockLedger.Fields("ItemGroup").Value
            .TextMatrix(i, 6) = Stock + Val(rstStockLedger.Fields("OPBAL").Value): If .TextMatrix(i, 6) < 0 Then .Cell(flexcpForeColor, i, 6) = vbRed Else .Cell(flexcpForeColor, i, 6) = vbBlack
            .TextMatrix(i, 7) = "Units"
            .TextMatrix(i, 8) = Val(rstStockLedger.Fields("SQ").Value): If .TextMatrix(i, 8) < 0 Then .Cell(flexcpForeColor, i, 8) = vbRed Else .Cell(flexcpForeColor, i, 8) = vbBlack
            .TextMatrix(i, 18) = Val(rstStockLedger.Fields("PendingPO").Value): If .TextMatrix(i, 18) < 0 Then .Cell(flexcpForeColor, i, 18) = vbRed Else .Cell(flexcpForeColor, i, 18) = vbBlack
            .TextMatrix(i, 19) = Val(rstStockLedger.Fields("PendingSO").Value): If .TextMatrix(i, 19) < 0 Then .Cell(flexcpForeColor, i, 19) = vbRed Else .Cell(flexcpForeColor, i, 19) = vbBlack
            .TextMatrix(i, 20) = EffectiveStock: If .TextMatrix(i, 20) < 0 Then .Cell(flexcpForeColor, i, 20) = vbRed Else .Cell(flexcpForeColor, i, 20) = vbBlack
            .TextMatrix(i, 22) = EffectiveStock * Val(rstStockLedger.Fields("MRP").Value): If .TextMatrix(i, 22) < 0 Then .Cell(flexcpForeColor, i, 22) = vbRed Else .Cell(flexcpForeColor, i, 22) = vbBlack
            .TextMatrix(i, 32) = (rstStockLedger.Fields("Code").Value)
            .TextMatrix(i, 35) = rstStockLedger.Fields("HSNCode").Value
            dPrint = dPrint + 1
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
        End If
            TDBNumber2 = dPrint
NXT:
                rstStockLedger.MoveNext
            If MdiMainMenu.ProgressBar1.Value + Round((100 / rstStockLedger.RecordCount), 2) <= 100 Then
                MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + Round((100 / rstStockLedger.RecordCount), 2)
            End If
            Loop
            .Rows = i + 1
'Pending Quotations and Short Item Analysis
    ElseIf VchType >= 101 Then
        If VchType = 103 Then 'WIP
        .Cols = 14
        ElseIf VchType = 104 Then 'RM
        .Cols = 14
        ElseIf VchType = 105 Then
        .Cols = 33
        End If
        .Rows = rstStockLedger.RecordCount + 1
        rstStockLedger.MoveFirst
        Do While Not rstStockLedger.EOF
            i = i + 1
            C = 0
            If (Combo1.ListIndex = 0 Or Combo1.ListIndex = 1) And VchType = 105 Then
'                C = C + 0: .TextMatrix(i, C) = "#" & i
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("Item").Value
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("Binding").Value
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("FinishSize").Value
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("Pages").Value)
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("Price").Value)
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("UFG").Value
            ElseIf (Combo1.ListIndex = 2 Or Combo1.ListIndex = 3) And VchType = 105 Then
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("UFG").Value
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("Item").Value
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("Binding").Value
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("FinishSize").Value
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("Pages").Value)
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("Price").Value)
            ElseIf (Combo1.ListIndex = 4 Or Combo1.ListIndex = 5) And VchType = 105 Then
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("SUBUFG").Value
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("UFG").Value
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("Item").Value
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("Binding").Value
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("FinishSize").Value
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("Pages").Value)
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("Price").Value
            End If
            If VchType = 103 Or VchType = 105 Then
                If VchType = 103 Then C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("UFG").Value
                C = C + 1: .TextMatrix(i, C) = Format(rstStockLedger.Fields("UFGReq/UNIT").Value, "###0.000")
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("UFGPAGES").Value)
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("Color").Value
                If FrmItemSelectionList.Check1.Value = False And VchType = 103 Then C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("FG").Value
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("Stock").Value)
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("SalesOrder").Value)
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("Dispatched").Value)
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("PendingSO").Value)
            If VchType = 105 Then
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("PendingSORate").Value)
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("PendingSOAmount").Value)
            End If
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("DeficientSalesOrder").Value)
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("UFGRequired").Value)
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("UFGStock").Value
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("FinalUFGRequired").Value)
            End If
            If VchType <> 103 And (Combo1.ListIndex <> 4 And Combo1.ListIndex <> 5 Or VchType = 104) Then
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("SUBUFG").Value
            End If
            If VchType = 104 Or VchType = 105 Then
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("SubUFG_Make").Value
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("SubUFG_GSM").Value)
                C = C + 1: .TextMatrix(i, C) = Format(rstStockLedger.Fields("SubUFG_CUTOFF").Value, "###0.00")
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("SubUFGReq/UNIT").Value)
                C = C + 1: .TextMatrix(i, C) = Format(rstStockLedger.Fields("Weight/Unit").Value, "###0.000")
                C = C + 1: .TextMatrix(i, C) = rstStockLedger.Fields("UOM_Name").Value
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("UOM").Value)
                C = C + 1: .TextMatrix(i, C) = Val(rstStockLedger.Fields("SUBUFGReqSheets").Value)
                C = C + 1: .TextMatrix(i, C) = Format(rstStockLedger.Fields("SUBUFGReqKg").Value, "###0.000")
                C = C + 1: .TextMatrix(i, C) = Format(rstStockLedger.Fields("SubUFGStkUOM").Value, "###0.000")
                C = C + 1: .TextMatrix(i, C) = Format(rstStockLedger.Fields("SubUFGStockKg").Value, "###0.000")
                C = C + 1: .TextMatrix(i, C) = Format(rstStockLedger.Fields("FinalSUBUFGReqKg").Value, "###0.000")
            End If
            dPrint = dPrint + 1
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstStockLedger.RecordCount & " !!!"
                
                rstStockLedger.MoveNext
            If MdiMainMenu.ProgressBar1.Value + Round((100 / rstStockLedger.RecordCount), 2) <= 100 Then
                MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + Round((100 / rstStockLedger.RecordCount), 2)
            End If
            Loop
          .Rows = i + 1
          TDBNumber2.Value = i
    Else
    
        Set VSFlexGrid1.DataSource = rstStockLedger
        
    End If

End With

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
Private Sub VSFlexGrid1_AfterDataRefresh()
Dim C As Variant
Dim T As Long
Dim GroupOn As Long
nSort = False
If VchType <= 2 Or VchType = 33 Then
GroupOn = 5
VSFlexGrid1.FrozenCols = 7
ElseIf VchType > 101 Then
GroupOn = 3
VSFlexGrid1.FrozenCols = 3
End If
'Subtotal
With VSFlexGrid1
.SubtotalPosition = flexSTBelow

If nSort = True Then
    .MultiTotals = True
    .Subtotal flexSTClear
        For C = 1 To .Cols - 1
        Err.Number = 0
        On Error Resume Next
        T = .TextMatrix(1, C)
        If Err.Number = 0 Then
            If InStr(1, "#_ITEM_MRP_PRICE_FINISH SIZE_PAGES_RM GSM_RM CUT-OFF_RATE_BAL_WIP_WIP/UNIT_RM_RM/UNIT_WEIGHT/UNIT", StrConv(.TextMatrix(0, C), vbUpperCase)) > 0 Then
                .Subtotal flexSTAverage, GroupOn, C, "(#,##0)", RGB(240, 230, 247), RGB(128, 0, 64), True, "Sub Total", GroupOn, True
                For T = 1 To .Rows - 1
                If .TextMatrix(T, 3) = "" Then .TextMatrix(T, 3) = "Sub Total"
                Next
            Else
                .Subtotal flexSTSum, GroupOn, C, "(#,##0)", RGB(240, 230, 247), RGB(128, 0, 64), True, "Sub Total", GroupOn, True
                .Subtotal flexSTSum, GroupOn, C, "(#,##0)", RGB(241, 248, 248), vbRed, True, , 1, True: .TextMatrix(.Rows - 1, 1) = "Grand Total"
            End If
        End If
        Next
    nSort = False
ElseIf nSort = False Then
    .MultiTotals = True
    .Subtotal flexSTClear
        For C = 1 To .Cols - 1
        Err.Number = 0
        On Error Resume Next
        T = .TextMatrix(1, C)
        If Err.Number = 0 Then
            If InStr(1, "#_ITEM_MRP_PRICE_FINISH SIZE_PAGES_RM GSM_RM CUT-OFF_RATE_BAL_WIP_WIP/UNIT_RM_RM/UNIT_WEIGHT/UNIT", StrConv(.TextMatrix(0, C), vbUpperCase)) > 0 Then
            
            Else
            
                If FrmItemSelectionList.Check1.Value = False Then .Subtotal flexSTSum, 1, C, "(#,##0)", RGB(240, 230, 247), vbBlue, True, "Sub-Total", 1, True
                .Subtotal flexSTSum, 0, C, "(#,##0)", RGB(240, 230, 247), vbRed, True, , 0, True: .TextMatrix(.Rows - 1, 1) = "Grand Total"
            End If
        End If
        Next
    nSort = True
End If
    For C = 1 To (.Cols - 1)
        .AutoSize C
        .ExplorerBar = flexExSort
        .ColSort(C) = flexSortCustom
        .AllowUserResizing = flexResizeBoth
    Next
    .Col = 33
C = .Rows - 1

.TextMatrix(C, 0) = ""
End With
End Sub
Private Sub VSFlexGrid1_AfterSort(ByVal Col As Long, Order As Integer)
Dim C As Variant
Dim T As Long
VSFlexGrid1.SubtotalPosition = flexSTBelow
'VSFlexGrid1.AutoResize = True
With VSFlexGrid1

If nSort = True Then
    .MultiTotals = True
    .Subtotal flexSTClear
        
        For C = 1 To .Cols - 1
        Err.Number = 0
        On Error Resume Next
        T = .TextMatrix(1, C)
        If Err.Number = 0 Then
            If InStr(1, "#_ITEM_MRP_PRICE_FINISH SIZE_PAGES_RM GSM_RM CUT-OFF_RATE_BAL_WIP_WIP/UNIT_RM_RM/UNIT_WEIGHT/UNIT", StrConv(.TextMatrix(0, C), vbUpperCase)) > 0 Then
            
            Else
                .Subtotal flexSTSum, .Col, C, "(#,##0)", RGB(240, 230, 247), RGB(128, 0, 64), True, "Sub Total", .Col, True
                .Subtotal flexSTSum, 1, C, "(#,##0)", RGB(240, 230, 247), &H808000, True, , 1, True
                .Subtotal flexSTSum, 0, C, "(#,##0)", RGB(240, 230, 247), vbRed, True, , 0, True: .TextMatrix(.Rows - 1, 1) = "Grand Total"
            End If
        End If
        Next
        
    nSort = False

ElseIf nSort = False Then
    .MultiTotals = True
    .Subtotal flexSTClear
        For C = 1 To .Cols - 1
        Err.Number = 0
        On Error Resume Next
        T = .TextMatrix(1, C)
        If Err.Number = 0 Then
            If InStr(1, "#_ITEM_MRP_PRICE_FINISH SIZE_PAGES_RM GSM_RM CUT-OFF_RATE_BAL_WIP_WIP/UNIT_RM_RM/UNIT_WEIGHT/UNIT", StrConv(.TextMatrix(0, C), vbUpperCase)) > 0 Then
    
            Else
                .Subtotal flexSTSum, 1, C, "(#,##0)", RGB(240, 230, 247), &H808000, True, "Subtotal", 1, True
                .Subtotal flexSTSum, 0, C, "(#,##0)", RGB(240, 230, 247), vbRed, True, , 0, True: .TextMatrix(.Rows - 1, 1) = "Grand Total"
            End If
        End If
        Next
    nSort = True
End If
    
    For C = 1 To (.Cols - 1)
        .AutoSize C
        .ExplorerBar = flexExSort
        .ColSort(C) = flexSortCustom
        .AllowUserResizing = flexResizeBoth
    Next
    
    C = .Rows - 1
: .TextMatrix(.Rows - 1, 0) = ""
End With
End Sub
Private Sub Zoom_Click()
Dim C As Long

With VSFlexGrid1
C = Zoom.Value
If .Font.Size >= 8.25 And FontFlag = False Then
    .Font.Size = .Font.Size + 1
    If .Font.Size >= 13.25 Then FontFlag = True
ElseIf .Font.Size > 8.25 And FontFlag = True Then
    .Font.Size = .Font.Size - 1
    If .Font.Size <= 8.25 Then FontFlag = False
End If
.Font.Size = .Font.Size + C

    .AutoResize = True
    For C = 1 To (.Cols - 1)
        .AutoSize C
        .ExplorerBar = flexExSort
        .ColSort(C) = flexSortCustom
        .AllowUserResizing = flexResizeBoth
        .Cell(flexcpFontSize, .Rows - 1, C) = .Font.Size
    Next
If FontFlag = False Then
    Command3.ToolTipText = "Zoom IN " & (.Font.Size - 8.25) * 20 & "%"
ElseIf FontFlag = True Then
    Command3.ToolTipText = "Zoom Out " & (.Font.Size - 8.25) * 20 & "%"
End If

End With
End Sub
Private Sub VSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    With VSFlexGrid1
If VchType = 0 Then
    .Col = 33
    
        If .Col = 33 Then .Editable = flexEDKbdMouse
        
        If (Shift = 0 And KeyCode = vbKeyReturn) And VchType = 0 And .Col = 33 Then 'Enter Physical Stock
    
                    If .Col = 33 And .Row < .Rows Then
                            sysStock = .TextMatrix(.Row, 6)
                            phyStock = .TextMatrix(.Row, 33)
                        If phyStock = "" Then
                            '.TextMatrix(.Row, 34) = Val(sysStock)
                            .Subtotal flexSTSum, 1, 33, "(#,##0)", RGB(240, 230, 247), vbBlue, True, "Grand Total", 1, True
                            .Subtotal flexSTSum, 1, 34, "(#,##0)", RGB(240, 230, 247), vbBlue, True, "Grand Total", 1, True
                        ElseIf phyStock <> "" Then
                                .TextMatrix(.Row, 34) = Val(phyStock) - Val(sysStock)
                                If .TextMatrix(.Row, 34) < 0 Then
                                    .Cell(flexcpForeColor, .Row, 34) = vbRed
                                ElseIf .TextMatrix(.Row, 34) > 0 Then
                                    .Cell(flexcpForeColor, .Row, 34) = vbBlack
                                End If
                        End If
                    End If
                                
                                .Subtotal flexSTSum, 1, 33, "(#,##0)", RGB(240, 230, 247), vbBlue, True, "Grand Total", 1, True
                                .Subtotal flexSTSum, 1, 34, "(#,##0)", RGB(240, 230, 247), vbBlue, True, "Grand Total", 1, True
                                .Editable = flexEDNone

        End If
End If
        If (Shift = 0 And KeyCode = vbKeyReturn) And .Row + 1 < .Rows Then
            .Row = .Row + 1
            KeyCode = 0
        End If
    End With
End Sub
Private Sub Mh3dLabel9_Click()
    Dim i As Long, j As Long, K As Long, C As Long
    Dim JItem As Variant, JQty As Variant, JRate As Variant, JICode As Variant, JHSNCode As Variant
    On Error Resume Next
    frmStockJournalVoucher.VchType = "JR"
    Load frmStockJournalVoucher
    If Err.Number <> 364 Then frmStockJournalVoucher.Show
    frmStockJournalVoucher.Toolbar1_ButtonClick frmStockJournalVoucher.Toolbar1.Buttons.Item(1)
                    j = 1: K = 1
If VSFlexFlag = False Then
                For i = 1 To fpSpread1.DataRowCnt - 1
                fpSpread1.GetText 34, i, JQty
                fpSpread1.GetText 3, i, JItem
    
    'Get Stock Journal Qty
    If JItem <> "Grand Total" And JQty > 0 Or JQty < 0 Then
                    C = C + 1
                    fpSpread1.GetText 3, i, JItem
                    fpSpread1.GetText 34, i, JQty
                    fpSpread1.GetText 4, i, JRate
                    fpSpread1.GetText 32, i, JICode
                    fpSpread1.GetText 35, i, JHSNCode
    
    'Set Stock Journal Qty
                If JQty > 0 Then
                    
                        frmStockJournalVoucher.fpSpread1.SetText 1, j, JItem
                        frmStockJournalVoucher.fpSpread1.SetText 2, j, JQty
                        frmStockJournalVoucher.fpSpread1.SetText 3, j, JRate
                        frmStockJournalVoucher.fpSpread1.SetText 4, j, JQty * JRate
                        frmStockJournalVoucher.fpSpread1.SetText 5, j, JICode
                        frmStockJournalVoucher.fpSpread1.SetText 6, j, JHSNCode
        'Active Row Generated
                            j = j + 1
                    
                ElseIf JQty < 0 Then
                    
                        frmStockJournalVoucher.fpSpread2.SetText 1, K, JItem
                        frmStockJournalVoucher.fpSpread2.SetText 2, K, JQty * -1
                        frmStockJournalVoucher.fpSpread2.SetText 3, K, JRate
                        frmStockJournalVoucher.fpSpread2.SetText 4, K, JQty * JRate * -1
                        frmStockJournalVoucher.fpSpread2.SetText 5, K, JICode
                        frmStockJournalVoucher.fpSpread2.SetText 6, K, JHSNCode
        'Active Row Consumed
                            K = K + 1
                End If
        End If
            Next
            
ElseIf VSFlexFlag = True Then
                For i = 1 To VSFlexGrid1.Rows - 2
                JQty = VSFlexGrid1.TextMatrix(i, 34)
                JItem = VSFlexGrid1.TextMatrix(i, 3)
    'Get Stock Journal Qty
    If JItem <> "Grand Total" And (JQty > 0 Or JQty < 0) And JQty <> "" Then
                    C = C + 1
                    JItem = VSFlexGrid1.TextMatrix(i, 3)
                    JQty = VSFlexGrid1.TextMatrix(i, 34)
                    JRate = VSFlexGrid1.TextMatrix(i, 4)
                    JICode = VSFlexGrid1.TextMatrix(i, 32)
                    JHSNCode = VSFlexGrid1.TextMatrix(i, 35)
    
    'Set Stock Journal Qty
                If JQty > 0 Then
                    
                        frmStockJournalVoucher.fpSpread1.SetText 1, j, JItem
                        frmStockJournalVoucher.fpSpread1.SetText 2, j, JQty
                        frmStockJournalVoucher.fpSpread1.SetText 3, j, JRate
                        frmStockJournalVoucher.fpSpread1.SetText 4, j, JQty * JRate
                        frmStockJournalVoucher.fpSpread1.SetText 5, j, JICode
                        frmStockJournalVoucher.fpSpread1.SetText 6, j, JHSNCode
        'Active Row Generated
                            j = j + 1
                    
                ElseIf JQty < 0 Then
                    
                        frmStockJournalVoucher.fpSpread2.SetText 1, K, JItem
                        frmStockJournalVoucher.fpSpread2.SetText 2, K, JQty * -1
                        frmStockJournalVoucher.fpSpread2.SetText 3, K, JRate
                        frmStockJournalVoucher.fpSpread2.SetText 4, K, JQty * JRate * -1
                        frmStockJournalVoucher.fpSpread2.SetText 5, K, JICode
                        frmStockJournalVoucher.fpSpread2.SetText 6, K, JHSNCode
        'Active Row Consumed
                            K = K + 1
                End If
        End If
            Next
End If
        If C = 0 Then
                frmStockJournalVoucher.Toolbar1_ButtonClick frmStockJournalVoucher.Toolbar1.Buttons.Item(5)
                MsgBox ("There is Zero Item to Create Stock Journal Voucher"), vbCritical
                Call CloseForm(frmStockJournalVoucher)
        Else
                Mh3dLabel6_Click
                Call CloseForm(FrmStockLedger)
        End If
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
