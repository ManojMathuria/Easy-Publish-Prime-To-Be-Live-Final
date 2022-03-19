VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProductionScheduling 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production Scheduling"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19845
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
   ScaleHeight     =   9375
   ScaleWidth      =   19845
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   19260
      Picture         =   "ProductionScheduling.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   200
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   18780
      Picture         =   "ProductionScheduling.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save"
      Top             =   200
      Width           =   375
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   375
      Left            =   18300
      Picture         =   "ProductionScheduling.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Refresh"
      Top             =   200
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   9150
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   19650
      _Version        =   65536
      _ExtentX        =   34660
      _ExtentY        =   16140
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
      Picture         =   "ProductionScheduling.frx":034E
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3000
         Top             =   2400
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
               Picture         =   "ProductionScheduling.frx":036A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductionScheduling.frx":08AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductionScheduling.frx":09C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductionScheduling.frx":0AD4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   7725
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   19410
         _Version        =   524288
         _ExtentX        =   34237
         _ExtentY        =   13626
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
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
         MaxCols         =   34
         MaxRows         =   1000
         SelectBlockOptions=   11
         SpreadDesigner  =   "ProductionScheduling.frx":0BE6
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   18150
         TabIndex        =   42
         Top             =   530
         Width           =   1380
         _ExtentX        =   2434
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
      End
      Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
         Height          =   330
         Left            =   6150
         TabIndex        =   43
         Top             =   105
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   582
         _StockProps     =   77
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
         Picture         =   "ProductionScheduling.frx":2033
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
            Height          =   330
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
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
            Caption         =   " &Show"
            Alignment       =   0
            FillColor       =   9164542
            TextColor       =   0
            Picture         =   "ProductionScheduling.frx":204F
            Picture         =   "ProductionScheduling.frx":206B
         End
         Begin MSForms.OptionButton Option1 
            Height          =   300
            Left            =   720
            TabIndex        =   46
            Top             =   10
            Width           =   615
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "1085;529"
            Value           =   "0"
            Caption         =   "All "
            FontName        =   "Calibri"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.OptionButton Option2 
            Height          =   300
            Left            =   1440
            TabIndex        =   44
            Top             =   10
            Width           =   1095
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "1931;529"
            Value           =   "1"
            Caption         =   "Pending"
            FontName        =   "Calibri"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.OptionButton Option3 
            Height          =   300
            Left            =   2640
            TabIndex        =   45
            Top             =   10
            Width           =   855
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "1508;529"
            Value           =   "0"
            Caption         =   "Close"
            FontName        =   "Calibri"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set &Machine Time"
         Height          =   330
         Left            =   9840
         TabIndex        =   41
         Top             =   480
         Width           =   3650
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Pending For Schedule"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6150
         TabIndex        =   27
         Top             =   480
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Freeze Panes"
         Height          =   330
         Left            =   9840
         TabIndex        =   40
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Preview 
         Caption         =   "&Print Preview"
         Height          =   330
         Left            =   15720
         TabIndex        =   39
         Top             =   8760
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Select All To Deschedule"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   15720
         TabIndex        =   29
         Top             =   500
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Select All To  Schedule"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   13560
         TabIndex        =   28
         Top             =   500
         Width           =   2055
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
         TabIndex        =   15
         ToolTipText     =   "Find And Search"
         Top             =   8760
         Width           =   3390
      End
      Begin VB.CommandButton cmdFilter 
         Height          =   320
         Left            =   6720
         Picture         =   "ProductionScheduling.frx":2087
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Filter"
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   320
         Left            =   7200
         Picture         =   "ProductionScheduling.frx":23C9
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Search"
         Top             =   8760
         Width           =   375
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   8
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
         Picture         =   "ProductionScheduling.frx":270B
         Picture         =   "ProductionScheduling.frx":2727
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   1800
         TabIndex        =   9
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
         Picture         =   "ProductionScheduling.frx":2743
         Picture         =   "ProductionScheduling.frx":275F
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
         Calendar        =   "ProductionScheduling.frx":277B
         Caption         =   "ProductionScheduling.frx":2893
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ProductionScheduling.frx":28FF
         Keys            =   "ProductionScheduling.frx":291D
         Spin            =   "ProductionScheduling.frx":297B
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   720
         TabIndex        =   0
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "ProductionScheduling.frx":29A3
         Caption         =   "ProductionScheduling.frx":2ABB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ProductionScheduling.frx":2B27
         Keys            =   "ProductionScheduling.frx":2B45
         Spin            =   "ProductionScheduling.frx":2BA3
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   3270
         TabIndex        =   10
         Top             =   105
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
         Caption         =   " &Show Order"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ProductionScheduling.frx":2BCB
         Picture         =   "ProductionScheduling.frx":2BE7
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   9840
         TabIndex        =   11
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " &Machine"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ProductionScheduling.frx":2C03
         Picture         =   "ProductionScheduling.frx":2C1F
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   330
         Left            =   1200
         TabIndex        =   16
         Top             =   8760
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "ProductionScheduling.frx":2C3B
         Caption         =   "ProductionScheduling.frx":2C5B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ProductionScheduling.frx":2CBF
         Keys            =   "ProductionScheduling.frx":2CDD
         Spin            =   "ProductionScheduling.frx":2D27
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
         ValueVT         =   176685061
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   8760
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
         Picture         =   "ProductionScheduling.frx":2D4F
         Picture         =   "ProductionScheduling.frx":2D6B
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   18330
         TabIndex        =   18
         Top             =   8760
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
         Picture         =   "ProductionScheduling.frx":2D87
         Picture         =   "ProductionScheduling.frx":2DA3
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   17040
         TabIndex        =   19
         Top             =   8760
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
         Picture         =   "ProductionScheduling.frx":2DBF
         Picture         =   "ProductionScheduling.frx":2DDB
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   2520
         TabIndex        =   20
         Top             =   8760
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
         Picture         =   "ProductionScheduling.frx":2DF7
         Picture         =   "ProductionScheduling.frx":2E13
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel 
         Height          =   330
         Left            =   9840
         TabIndex        =   21
         Top             =   8760
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
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
         Caption         =   "Ctrl+F->Search  F8->Delete  F9->Hide  Escap->Un-Hide  F5->Refresh"
         FillColor       =   8421504
         TextColor       =   16777215
         Picture         =   "ProductionScheduling.frx":2E2F
         Picture         =   "ProductionScheduling.frx":2E4B
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   15735
         TabIndex        =   23
         Top             =   8760
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
         Picture         =   "ProductionScheduling.frx":2E67
         Picture         =   "ProductionScheduling.frx":2E83
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   13560
         TabIndex        =   25
         Top             =   120
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
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
         Caption         =   "Click To Schedule M/C"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ProductionScheduling.frx":2E9F
         Picture         =   "ProductionScheduling.frx":2EBB
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   15720
         TabIndex        =   26
         Top             =   120
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   "Click To Deschedule M/C"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ProductionScheduling.frx":2ED7
         Picture         =   "ProductionScheduling.frx":2EF3
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
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
         Caption         =   " &M/C Lunch Time From"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ProductionScheduling.frx":2F0F
         Picture         =   "ProductionScheduling.frx":2F2B
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   2760
         TabIndex        =   33
         Top             =   480
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
         Picture         =   "ProductionScheduling.frx":2F47
         Picture         =   "ProductionScheduling.frx":2F63
      End
      Begin TDBTime6Ctl.TDBTime TDBTime4 
         Height          =   330
         Left            =   2040
         TabIndex        =   36
         Top             =   480
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Caption         =   "ProductionScheduling.frx":2F7F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "ProductionScheduling.frx":2FE3
         Spin            =   "ProductionScheduling.frx":3033
         AlignHorizontal =   2
         AlignVertical   =   2
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn:ss"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn:ss"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "13:00:00"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0.541666666666667
      End
      Begin TDBTime6Ctl.TDBTime TDBTime5 
         Height          =   330
         Left            =   3120
         TabIndex        =   37
         Top             =   480
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Caption         =   "ProductionScheduling.frx":305B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "ProductionScheduling.frx":30BF
         Spin            =   "ProductionScheduling.frx":310F
         AlignHorizontal =   2
         AlignVertical   =   2
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn:ss"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn:ss"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "13:30:00"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0.5625
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         Left            =   4560
         MaxLength       =   40
         TabIndex        =   38
         ToolTipText     =   "Find And Search"
         Top             =   840
         Visible         =   0   'False
         Width           =   870
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   3720
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
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
         Caption         =   " Shift Hrs."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ProductionScheduling.frx":3137
         Picture         =   "ProductionScheduling.frx":3153
      End
      Begin TDBTime6Ctl.TDBTime TDBTime1 
         Height          =   330
         Left            =   1920
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Caption         =   "ProductionScheduling.frx":316F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "ProductionScheduling.frx":31D3
         Spin            =   "ProductionScheduling.frx":3223
         AlignHorizontal =   2
         AlignVertical   =   2
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn:ss"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn:ss"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "09:00:00"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0.375
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
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
         Caption         =   " &M/C Start Time From"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ProductionScheduling.frx":324B
         Picture         =   "ProductionScheduling.frx":3267
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   2640
         TabIndex        =   31
         Top             =   840
         Visible         =   0   'False
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
         Picture         =   "ProductionScheduling.frx":3283
         Picture         =   "ProductionScheduling.frx":329F
      End
      Begin TDBTime6Ctl.TDBTime TDBTime2 
         Height          =   330
         Left            =   3000
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Caption         =   "ProductionScheduling.frx":32BB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "ProductionScheduling.frx":331F
         Spin            =   "ProductionScheduling.frx":336F
         AlignHorizontal =   2
         AlignVertical   =   2
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn:ss"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn:ss"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.999988425925926
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "09:00:00"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0.375
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Left            =   3840
         TabIndex        =   48
         Top             =   480
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
         Caption         =   " &Show Plates"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ProductionScheduling.frx":3397
         Picture         =   "ProductionScheduling.frx":33B3
      End
      Begin MSForms.ComboBox Combo4 
         Height          =   330
         Left            =   4920
         TabIndex        =   49
         Top             =   480
         Width           =   1200
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2117;582"
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
         Left            =   7680
         TabIndex        =   22
         Top             =   8760
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
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   10680
         TabIndex        =   12
         Top             =   120
         Width           =   2805
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4948;582"
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
         Left            =   4350
         TabIndex        =   2
         Top             =   105
         Width           =   1785
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3149;582"
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
         X2              =   19810
         Y1              =   900
         Y2              =   900
      End
   End
End
Attribute VB_Name = "FrmProductionScheduling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchType As String, OutputTo As String
Dim oOutlook As New Outlook.Application
Dim cnProductionScheduling As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstMachineMaster As New ADODB.Recordset
Dim rstProductionScheduling As New ADODB.Recordset
Public rstBookPOChild0501 As New ADODB.Recordset
Dim EditMode As Boolean, MachineCode As Variant
Dim Reset As Long, TotalFlag As Boolean, LR As Integer, SaveFlag As Boolean
Dim sDate As Date, eDate As Date, mStart As Date, mLstart As Date, mLstop As Date, mStop As Date, mHrs As Date, sStopHrs As Long, nShiftHrs As Double, rShiftHrs As Double, jobHrs As Double, hh As Long, mm As Long
Dim mSHH As Long, mSMM As Long, mCSHH As Long, mCMM As Long, mLsHH As Long, mLsMM As Long, mLeHH As Long, mLeMM As Long, mLTMM As Date
Dim timeFlag As Boolean, freezeFlag As Boolean
Private Sub Command1_Click()
    On Error Resume Next
    FrmMachineMaster.SL = False
    Load FrmMachineMaster
    If Err.Number <> 364 Then FrmMachineMaster.Caption = "Machine Master": FrmMachineMaster.Show
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT * FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    cnProductionScheduling.CursorLocation = adUseClient
    cnProductionScheduling.Open cnDatabase.ConnectionString
    rstBookPOChild0501.Open "SELECT Code,Color,Machine,[Plan],formsPrinted,platesIssued,paperIssued,SNo FROM BookPOChild0501 ORDER BY Code", cnProductionScheduling, adOpenKeyset, adLockOptimistic
    'Combo1.AddItem
    Combo1.Clear
    Combo1.AddItem "Sales Order", 0
    Combo1.AddItem "Purchase Order", 1
    Combo1.AddItem "Both", 2
    Combo1.ListIndex = 0
    Reset = 0:
    LoadMasterList
    Dim i As Long
    rstMachineMaster.MoveFirst
    Do While Not rstMachineMaster.EOF
    'Combo2.AddItem
    Combo2.AddItem rstMachineMaster.Fields("Col0") + "  M/C_Code <<<>>> " & Trim(rstMachineMaster.Fields("Code").Value), i
    i = i + 1
    rstMachineMaster.MoveNext
    Loop
    Combo2.AddItem "All", i
    Combo2.ListIndex = 0
    
    'Combo3.AddItem
    Combo3.Clear
    Combo3.AddItem " Party", 0
    Combo3.AddItem " Ref No", 1
    Combo3.AddItem " Ref Date", 2
    Combo3.AddItem " Item", 3
    Combo3.AddItem " Size", 4
    Combo3.AddItem " Col", 5
    Combo3.ListIndex = 1
    
    'Combo4
    Combo4.Clear
    Combo4.AddItem " New", 0
    Combo4.AddItem " Old", 1
    Combo4.AddItem " Both", 2
    Combo4.ListIndex = 2
    
    If VchType <> 2 Then Combo4.Visible = False: Mh3dLabel18.Visible = False
    Reset = 1
    MhDateInput1.Value = Format(FinancialYearFrom, "dd-MM-yyyy")
    If Format(Date, "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then MhDateInput1.Value = Format(FinancialYearTo, "dd-MM-yyyy") Else MhDateInput2.Value = Format(Date, "dd-MM-yyyy")
LoadMasterList
    cmdRefresh_Click
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub cmdFilter_Click()
        Call Total_Click
End Sub
Private Sub Combo2_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Combo4_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Check2_Click()
Dim MC As Variant
MC = Right(Combo2.Value, 6)
If MC = "*21046" Then: Check2 = 1: Exit Sub
Call cmdRefresh_Click
End Sub
Private Sub Check3_Click()
Dim i As Integer, cVal As Variant, cVa(1 To 3) As Variant, MC As Variant
If Check3.Value = 0 Then Exit Sub
MC = Right(Combo2.Value, 6)
If MC = "All" Or MC = "*21046" Then MsgBox " Please Select Machine !!! ", vbCritical: Check3.Value = 0: Combo2.SetFocus: Exit Sub
If Check3.Value And Check4.Value Then Check4.Value = 0
With fpSpread1
For i = 1 To .DataRowCnt
    .GetText 26, i, cVal
.Row = i
If Not .RowHidden Then
    If (cVal = 0 Or Val(cVal) > 0) And Check3.Value = 0 Then
        .SetText 9, i, 0
    ElseIf Val(cVal) > 0 Then
        .SetText 9, i, 0
    ElseIf cVal = 0 Then
        .SetText 9, i, 1
    End If
        .GetText 9, i, cVa(2)
    If MC <> "All" Then cVa(3) = Val(cVa(3)) + Val(cVa(2))
End If
Next i
End With
MsgBox "( " & Val(cVa(3)) & " )  Jobs Selected To For schedule"
Check3.Value = 0
End Sub
Private Sub Check4_Click()
Dim i As Integer, cVal As Variant, cVa(1 To 3) As Variant, MC As Variant
If Check4.Value = 0 Then Exit Sub
MC = Right(Combo2.Value, 6)
If MC = "All" Or MC = "*21046" Then MsgBox " Please Select Machine !!! ", vbCritical: Check4.Value = 0: Combo2.SetFocus: Exit Sub
If Check3.Value And Check4.Value Then Check3.Value = 0
With fpSpread1
For i = 1 To .DataRowCnt
.GetText 26, i, cVal
.Row = i
If Not .RowHidden Then
If (cVal = 0 Or Val(cVal) > 0) And Check4.Value = 0 Then
    If Not .RowHidden Then .SetText 9, i, 0
ElseIf Val(cVal) > 0 Then
    If Not .RowHidden Then .SetText 9, i, 1
ElseIf cVal = 0 Then
    If Not .RowHidden Then .SetText 9, i, 0
End If
.GetText 9, i, cVa(2)
If MC <> "All" Then cVa(3) = Val(cVa(3)) + Val(cVa(2))
End If
Next i
End With
MsgBox "( " & Val(cVa(3)) & " )  Jobs Selected To For schedule"
Check4.Value = 0
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}": KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        If Not EditMode Then cmdProceed_Click: KeyCode = 0
    ElseIf KeyCode = vbKeyF5 Then
        Call cmdRefresh_Click
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        If Not EditMode Then cmdCancel_Click: KeyCode = 0
    
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
        If Text1.Text = "" Then
            MsgBox "Please Provide Search Input", vbInformation
            Text1.SetFocus
        ElseIf Text1.Text <> "" Then
        Call Command2_Click
        End If
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstMachineMaster)
    Call CloseRecordset(rstBookPOChild0501)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(cnProductionScheduling)
    End Sub
Private Sub cmdProceed_Click()
    Mh3dLabel11_Click
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Sub Mh3dLabel11_Click()
Dim i As Integer, cVal(1 To 4) As Variant, MC As Variant, rVal As Variant
    cVal(3) = 0
    MC = Right(Combo2.Value, 6)
    If MC = "All" Or MC = "*21046" Then MsgBox " Please Select Machine !!! ", vbCritical: Combo2.SetFocus: Exit Sub
    With fpSpread1
    For i = 1 To .DataRowCnt
            .GetText 25, i, cVal(1): .GetText 2, i, rVal: .GetText 9, i, cVal(2):
        If (cVal(1) = "" Or cVal(1) = "*21046") And rVal <> "" And cVal(2) = 1 And MC <> "All" And MC <> "*21046" Then
            .SetText 25, i, MC
        ElseIf cVal(1) = "" And rVal <> "" And cVal(2) = "" Then
            .SetText 25, i, "*21046"
        End If
            If MC <> "*21046" Then cVal(3) = Val(cVal(3)) + Val(cVal(2))
        Next
    End With
        SaveFlag = True
        SaveFields
    If Combo2.Value <> "All" Then MsgBox "( " & cVal(3) & " ) Jobs Add To schedule For " & Chr(13) & Chr(13) & " <<<>>> " & Left(Combo2.Value, Int(Len(Combo2.Value)) - 6) & "", vbInformation, " Easy Publish !!! "
End Sub
Private Sub Mh3dLabel12_Click()
Dim i As Integer, cVal(1 To 4) As Variant, MC As Variant, rVal As Variant, oPlan As Variant
    MC = Right(Combo2.Value, 6)
    If MC = "All" Or MC = "*21046" Then MsgBox " Please Select Machine !!! ", vbCritical: Combo2.SetFocus: Exit Sub
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 25, i, cVal(1): .GetText 2, i, rVal: .GetText 9, i, cVal(2): .GetText 10, i, cVal(4): .GetText 34, i, oPlan
        If (cVal(1) = MC Or cVal(1) <> "*21046") And rVal <> "" And cVal(2) = 1 And MC <> "All" And cVal(4) = 0 Then
            .SetText 25, i, "*21046"
        ElseIf cVal(1) = "" And rVal <> "" And cVal(2) = "" Then
            .SetText 25, i, "*21046"
        ElseIf (cVal(1) = MC Or cVal(1) <> "*21046") And rVal <> "" And cVal(2) = 1 And MC <> "All" And cVal(4) <> 0 Then
            .SetText 8, i, cVal(4): .SetText 11, i, 0: fpSpread1.SetText 32, i, Val(oPlan) - cVal(4)
        End If
            If rVal <> "" Then .SetText 9, i, 0
            If cVal(4) = 0 Then cVal(3) = Val(cVal(3)) + Val(cVal(2))
        Next
    End With
        SaveFlag = True
        SaveFields
    MsgBox "( " & cVal(3) & " ) Jobs Removed To Reschedule From " & Chr(13) & Chr(13) & " <<<>>> " & Left(Combo2.Value, Int(Len(Combo2.Value)) - 6) & "", vbInformation, " Easy Publish !!! "
End Sub
Private Sub SaveFields()
    Dim i As Integer, Code As Variant, Color As Variant, Machine As Variant, formsPrinted As Variant, platesIssued As Variant, paperIssued As Variant, SrNo As Variant, n As Integer, SNo As Variant, CellVal(1 To 7) As Variant, ActiveCellVal As Variant, Plan As Variant
         i = 0: SrNo = "000000"
    With rstProductionScheduling
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
        cnProductionScheduling.Execute "DELETE FROM BookPOChild0501 WHERE Code='" & rstProductionScheduling.Fields("RefCode").Value & "'"
        .MoveNext
        Loop
    End With
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 2, i, ActiveCellVal
            .GetText 9, i, CellVal(1)
            .GetText 25, i, Machine
        If ActiveCellVal <> "" And Trim(Machine) <> "*21046" Then
            .SetText 26, i, i - 1
                    .GetText 6, i, Color
                    .GetText 25, i, Machine
                    .GetText 8, i, Plan
                    .GetText 10, i, formsPrinted
                    .GetText 24, i, Code
                    .GetText 17, i, platesIssued
                    .GetText 22, i, paperIssued
                    .GetText 26, i, SNo
                With rstBookPOChild0501
                    .AddNew
                    .Fields("Code").Value = Code
                    .Fields("Color").Value = Color
                    .Fields("Machine").Value = Trim(Machine)
                    .Fields("Plan").Value = Plan
                    .Fields("formsPrinted").Value = formsPrinted
                    .Fields("platesIssued").Value = platesIssued
                    .Fields("PaperIssued").Value = paperIssued
                    .Fields("SNO").Value = SNo
                    .Update
                End With
        End If
        Next
    End With
        cmdRefresh_Click
        SaveFlag = False
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    With fpSpread1
        If .ActiveCol = 31 Then
            If KeyCode = vbKeySpace Then
                .GetText 25, .ActiveRow, MachineCode
                On Error Resume Next
                FrmMachineMaster.SL = True
                FrmMachineMaster.MasterCode = MachineCode
                Load FrmMachineMaster
                If Err.Number <> 364 Then FrmMachineMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText 25, .ActiveRow, slCode
                If fpSpread1.ActiveCol = 31 Then fpSpread1.SetActiveCell 31, fpSpread1.ActiveRow
                If Not CheckEmpty(slCode, False) Then LoadMasterList: Sendkeys "{ENTER}"
            ElseIf KeyCode = vbKeyDelete Then
                .SetText .ActiveCol, .ActiveRow, "": .SetText 25, .ActiveRow, ""
            End If
        End If
    End With
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim TotalQty As Variant, IssuedQty As Variant, oIssuedQty As Variant, RefNo As Variant, oRefNo As Variant, Plan As Variant, PendingPlan As Variant, SNo As Variant, oPlan As Variant, i As Long
        fpSpread1.GetText 2, Row, RefNo
    If RefNo <> "" Then
        If Col = 8 Or Col = 10 Or Col = 13 Or Col = 17 Or Col = 22 Then 'Plan  & Issued
        'TotalQty & Plan
            If Col = 10 Then
                fpSpread1.GetText Col - 2, Row, TotalQty
            ElseIf Col = 8 Then
                fpSpread1.GetText 7, Row, TotalQty
                fpSpread1.GetText Col, Row, Plan
            Else
                fpSpread1.GetText Col - 1, Row, TotalQty
            End If
            
        'Issued
            If Col = 10 Or Col = 13 Or Col = 17 Or Col = 22 Then
                fpSpread1.GetText Col, Row, IssuedQty
            ElseIf Col = 17 Then
                IssuedQty = 0
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.GetText 2, i, oRefNo
                    If RefNo = oRefNo Then
                        fpSpread1.GetText Col, Row, IssuedQty
                        IssuedQty = IssuedQty + oIssuedQty
                    End If
                Next
            ElseIf Col = 8 Then
                fpSpread1.GetText Col + 2, Row, IssuedQty
            End If
            
            
        'Balance
            If Col = 10 Or Col = 13 Or Col = 17 Or Col = 22 Then
                If Val(IssuedQty) > Val(TotalQty) Then
                Cancel = True
            Else
                fpSpread1.SetText Col + 1, Row, Val(TotalQty) - Val(IssuedQty)
            End If
        
        ElseIf Col = 8 Then
            fpSpread1.SetText Col + 3, Row, Val(Plan) - Val(IssuedQty)
        End If
            
            fpSpread1.GetText 26, Row, SNo
            fpSpread1.GetText 32, Row, PendingPlan
            If Col = 10 Or Col = 17 Or Col = 22 Then
                If Val(IssuedQty) > Val(TotalQty) Then Cancel = True
            ElseIf Col = 8 And SNo > 0 And Val(Plan) > Val(PendingPlan) Then
                fpSpread1.GetText 34, Row, oPlan
                'fpSpread1.SetText 8, Row, Val(oPlan)
                If Val(Plan) < Val(oPlan) Then fpSpread1.SetText 32, Row, Val(oPlan) - Val(Plan)
                If Val(Plan) > Val(oPlan) Then fpSpread1.SetText 8, Row, Val(oPlan): fpSpread1.SetText 11, Row, Val(oPlan) - IssuedQty
                fpSpread1.SetText 9, Row, 1
            ElseIf Col = 8 And Val(Plan) > Val(TotalQty) Then
                fpSpread1.SetText 8, Row, Val(TotalQty): Cancel = True: MsgBox "You Can't Plan Greater Than " & Chr(13) & Chr(13) & " Actual >>>" & (TotalQty) & Chr(13) & Chr(13) & " Plan >>>" & (Plan), vbInformation, "Easy Publish...Alert !!! "
            ElseIf Col = 8 And Val(Plan) < Val(IssuedQty) Then
                fpSpread1.SetText 8, Row, Val(IssuedQty): Cancel = True: MsgBox "You Can't Plan Less Than " & Chr(13) & Chr(13) & " Printed >>>" & (IssuedQty) & Chr(13) & Chr(13) & " Plan >>>" & (Plan), vbInformation, "Easy Publish...Alert !!! "
            ElseIf Col = 8 And SNo = 0 And Val(Plan) > Val(PendingPlan) Then
                fpSpread1.SetText 8, Row, Val(PendingPlan): Cancel = True: MsgBox "You Can't Plan Greater Than " & Chr(13) & Chr(13) & " Pending >>>" & (PendingPlan) & Chr(13) & Chr(13) & " Plan >>>" & (Plan), vbInformation, "Easy Publish...Alert !!! "
            End If
        End If
    End If
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub FormatCols()
Dim C As Long
    With fpSpread1
        .ColWidth(1) = 22.75
        .ColWidth(2) = 10.75
        .ColWidth(3) = 8.75
        .ColWidth(4) = 29
        .ColWidth(5) = 9.75
        .ColWidth(6) = 4.75
        .ColWidth(7) = 8
        .ColWidth(9) = 3.25
        .ColWidth(10) = 8
        .ColWidth(11) = 8
        .ColWidth(18) = 8
'        .ColWidth(19) = 53.75
        .ColWidth(20) = 8.75
        .ColWidth(21) = 8.75
        .ColWidth(22) = 8.75
        .ColWidth(23) = 8.75
        .ColWidth(24) = 10
        .ColWidth(25) = 10
        .ColWidth(26) = 8
        .ColWidth(27) = 7.75
        .ColWidth(28) = 7.75
        .ColWidth(29) = 16.25
        .ColWidth(30) = 16.25
        .ColWidth(31) = 17.5
        .ColWidth(32) = 10.25
        .ColWidth(33) = 10.25
    'Col Span
    .ColHeaderRows = 4
        .AddCellSpan 1, SpreadHeader, 34, 1
        .Col = 1: .Row = SpreadHeader: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .FontBold = True: .FontSize = 20: .FontUnderline = True: .ForeColor = RGB(1, 106, 106):
        .SetText 1, SpreadHeader, rstCompanyMaster.Fields("PrintName").Value:
        .AddCellSpan 1, SpreadHeader + 1, 34, 1
        .Col = 1: .Row = SpreadHeader + 1: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter:  .FontBold = True: .FontSize = 16: .ForeColor = RGB(20, 106, 106):
        .SetText 1, SpreadHeader + 1, Me.Caption + "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]":
    'Header Text
        .Row = SpreadHeader + 2
        .Col = 7: .Text = "Forms"
        .Col = 12: .Text = "Order"
        .Col = 16: .Text = "Plate"
        .Col = 19: .Text = "Paper Name"
        .Col = 20: .Text = "Paper"
        .Col = 27: .Text = "Actual"
        .Row = SpreadHeader + 3
        'Forms
        .Col = 7: .Text = " Actual"
        .Col = 8: .Text = "    Plan"
        .Col = 9: .Text = " "
        .Col = 10: .Text = "Printed"
        .Col = 11: .Text = "       Bal"
        'Ordered
        .Col = 12: .Text = "Quantity"
        .Col = 13: .Text = "     Out"
        .Col = 14: .Text = "     Bal"
        'Plate
        .Col = 15: .Text = "Plate Type"
        'Plate
        .Col = 16: .Text = "Total"
        .Col = 17: .Text = "Issued"
        .Col = 18: .Text = "    Bal"
        'Paper
        .Col = 20: .Text = "Wastage"
        .Col = 21: .Text = "      Total"
        .Col = 22: .Text = "    Issued"
        .Col = 23: .Text = "         Bal"
        'Time
        .Col = 27: .Text = "IMP"
        .Col = 28: .Text = "HH:MM"
        'Row Span
        .AddCellSpan 0, SpreadHeader + 2, 1, 2
        .AddCellSpan 1, SpreadHeader + 2, 1, 2
        .AddCellSpan 2, SpreadHeader + 2, 1, 2
        .AddCellSpan 3, SpreadHeader + 2, 1, 2
        .AddCellSpan 4, SpreadHeader + 2, 1, 2
        .AddCellSpan 5, SpreadHeader + 2, 1, 2
        .AddCellSpan 6, SpreadHeader + 2, 1, 2
        .AddCellSpan 7, SpreadHeader + 2, 1, 2
        .AddCellSpan 7, SpreadHeader + 2, 5, 1
        .AddCellSpan 11, SpreadHeader + 2, 1, 2
        .AddCellSpan 12, SpreadHeader + 2, 1, 2
        .AddCellSpan 12, SpreadHeader + 2, 3, 1
        .AddCellSpan 14, SpreadHeader + 2, 1, 2
        .AddCellSpan 15, SpreadHeader + 2, 1, 2
        .AddCellSpan 16, SpreadHeader + 2, 3, 1
        .AddCellSpan 18, SpreadHeader + 2, 1, 2
        .AddCellSpan 19, SpreadHeader + 2, 1, 2
        .AddCellSpan 20, SpreadHeader + 2, 4, 1
        .AddCellSpan 23, SpreadHeader + 2, 1, 2
        .AddCellSpan 24, SpreadHeader + 2, 1, 2
        .AddCellSpan 25, SpreadHeader + 2, 1, 2
        .AddCellSpan 26, SpreadHeader + 2, 1, 2
        .AddCellSpan 27, SpreadHeader + 2, 2, 1
        .AddCellSpan 29, SpreadHeader + 2, 1, 2
        .AddCellSpan 30, SpreadHeader + 2, 1, 2
        .AddCellSpan 31, SpreadHeader + 2, 1, 2
        .AddCellSpan 32, SpreadHeader + 2, 1, 2
        .AddCellSpan 33, SpreadHeader + 2, 1, 2
If VchType = 1 Then
        For C = 13 To 13
        .Col = C: .ColHidden = True
        Next
        For C = 14 To 20
        .Col = C: .ColHidden = True
        Next
        For C = 21 To 26
        .Col = C: .ColHidden = True
        Next
        .Col = 31: .ColHidden = True
        .Col = 32: .ColHidden = True
        .Col = 33: .ColHidden = True
        .Col = 34: .ColHidden = True
ElseIf VchType = 2 Then
        .Row = SpreadHeader + 1
        .Col = 31: .Text = "Ptg. Machine"
        .AddCellSpan 31, SpreadHeader + 2, 1, 2
        .ColWidth(4) = 30
        .ColWidth(15) = 17.5
        .ColWidth(16) = 8
        .ColWidth(17) = 8
        .ColWidth(18) = 8
        For C = 9 To 14
        .Col = C: .ColHidden = True
        Next
        .Col = 16: .ColHidden = False
        .Col = 17: .ColHidden = False
        .Col = 18: .ColHidden = False
        For C = 19 To 34
        .Col = C: .ColHidden = True
        Next
        .Col = 31: .ColHidden = False
ElseIf VchType = 3 Then
        .Row = SpreadHeader + 1
        .Col = 31: .Text = "Ptg. Machine"
        .AddCellSpan 31, SpreadHeader + 2, 1, 2
        .ColWidth(1) = 21
        .ColWidth(2) = 10.75
        .ColWidth(4) = 25
        .ColWidth(5) = 9
        .ColWidth(7) = 7
        .ColWidth(8) = 7
        .ColWidth(19) = 49
        .ColWidth(20) = 8
        .Col = 6: .ColHidden = True
        For C = 9 To 18
        .Col = C: .ColHidden = True
        Next
        For C = 19 To 23
        .Col = C: .ColHidden = False
        Next
        For C = 24 To 34
        .Col = C: .ColHidden = True
        Next
End If
'        For C = 1 To 34
'        .Col = C: .ColHidden = False
'        Next
    End With
End Sub
Private Sub cmdRefresh_Click()
Dim mcN As String, MC As String, C As Long, cVal(1 To 7) As Variant, Col As Variant, K As Long

sDate = #12:00:00 AM#: eDate = #12:00:00 AM#: mStart = #12:00:00 AM#: mLstart = #12:00:00 AM#: mLstop = #12:00:00 AM#: mStop = #12:00:00 AM#: mHrs = #12:00:00 AM#
sStopHrs = 0: nShiftHrs = 0: jobHrs = 0: hh = 0: mm = 0: rShiftHrs = 0
mSHH = 0: mSMM = 0: mCSHH = 0: mCMM = 0: mLsHH = 0: mLsMM = 0: mLeHH = 0: mLeMM = 0: mLTMM = #12:00:00 AM#
timeFlag = False
Text1.Text = "": TDBNumber2 = 0: Col = 0
fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.MaxCols = 34
fpSpread1.RowHeadersAutoText = DispBlank
fpSpread1.ColHeadersAutoText = DispBlank
    FormatCols
    MC = Right(Combo2.Value, 6)
    If MC = "*21046" And Check2.Value = 0 Then Check2.Value = 1:
    If MC = "*21046" And Option2 <> True Then Option2 = True
    On Error GoTo ErrHandler
    Dim SQL As String, i As Long
    'MF1
  SQL = "SELECT (Select Name From ElementMaster Where Code='*00011') As Element,RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-MF','UFG-MF')+'/'+LTRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],'1' As Col,A.Name As Party,C.Forms1 As TotalForms,ISNULL(C1.[Plan],C.Forms1) AS [Plan],ISNULL(C1.formsPrinted,0) AS formsPrinted,C.ActualQuantity As TotalQty,(P.DeliveredQuantityC+P.DeliveredQuantityB) As QtyIssued,IIF(PlateType1='1','Deepatch',IIF(PlateType1='2','PS',IIF(PlateType1='3','Wipeon','CTP'))) As Plate,IIF(Processing='R',[RevisedPlates1],[TotalPlates1-¼]+[TotalPlates1-½]+[TotalPlates1-1]+[RevisedPlates1]) As TotalPlates,ISNULL(C1.platesIssued,0) AS platesIssued, LTRIM(R.Name) As Paper,PARSENAME(PaperWastageFinal1,2)*U.Value1+(PARSENAME(PaperWastageFinal1,1)) As Wastage,PaperConsumptionsheets1 As TotalPaper,ISNULL(C1.paperIssued,0) As paperIssued,P.Code+'MF1*00011' As RefCode,M.Name AS MAC,M.Code AS MCode,MakeReadyTime,Efficiency,ISNULL(C1.SNo,0) AS SNo, " & _
"M.StartTime,M.EndTime,C.Forms1-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF1*00011'),0) AS PendingPlan,C.Processing FROM ((((((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.Size1=S.Code) INNER JOIN PaperMaster R ON C.Paper1=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'MF1*00011'=C1.Code) LEFT JOIN AccountMaster A ON P.BookPrinter=A.Code LEFT JOIN MachineMaster M ON M.Code=ISNULL(C1.Machine,'*21046') WHERE Forms1<>0 AND " & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>=1") & " AND LEFT(P.Type,1)<>'O' AND P.Date>='" & GetDate(MhDateInput1.Text) & "'  AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Option1.Value, "1=1", IIf(Option2.Value, "ISNULL(C1.[Plan],C.Forms1)-ISNULL(C1.FormsPrinted,0)<>0", IIf(VchType = 1 And Option3.Value, "ISNULL(C1.[Plan],C.Forms1)-ISNULL(C1.FormsPrinted,0)=0", "1=1"))) & _
            " AND " & IIf(Combo4.ListIndex = 0, "C.Processing='N'", IIf(Combo4.ListIndex = 1, "C.Processing='O'", "1=1")) & " AND " & IIf(Option1.Value, "1=1", IIf(Option2.Value, "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)>0", "1=1")) & " AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='S'", "RIGHT(P.Type,1)='P'", "1=1") & " AND " & IIf(Check2.Value = 0, "C1.Machine IS NOT NULL", IIf(MC = "All", "1=1", "(C1.Machine IN ('" & MC & "','*21046') OR " & IIf(Check2.Value, "(C1.Machine IS NULL)", "C1.Machine IN ('')") & ")")) & " AND " & IIf(Check2.Value = 0, "M.Code<>'*21046'", "1=1") & " AND " & _
            IIf(VchType = 2 And Option1.Value, "1=1", IIf(VchType = 2 And Option2.Value, "IIF(Processing='R',[RevisedPlates1],[TotalPlates1-¼]+[TotalPlates1-½]+[TotalPlates1-1]+[RevisedPlates1])-ISNULL(C1.platesIssued,0)<>0", IIf(VchType = 2 And Option3.Value, "((IIF(Processing='R',[RevisedPlates1],[TotalPlates1-¼]+[TotalPlates1-½]+[TotalPlates1-1]+[RevisedPlates1])-ISNULL(C1.platesIssued,0)=0) OR (P.DeliveredQuantityC+P.DeliveredQuantityB)>0)", "1=1"))) & " AND " & IIf(Check2.Value, "1=1", IIf(MC = "All", "1=1", "'" & MC & "'=C1.Machine"))
    'MF2
    SQL = SQL + " UNION " & _
           "SELECT (Select Name From ElementMaster Where Code='*00012') As Element,RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-MF','UFG-MF')+'/'+LTRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],'2' As Col,A.Name As Party,C.Forms2 As TotalForms,ISNULL(C1.[Plan],C.Forms2)AS [Plan],ISNULL(C1.formsPrinted,0)  AS formsPrinted,C.ActualQuantity As TotalQty,(P.DeliveredQuantityC+P.DeliveredQuantityB) As QtyIssued,IIF(PlateType2='1','Deepatch',IIF(PlateType2='2','PS',IIF(PlateType2='3','Wipeon','CTP'))) As Plate,IIF(Processing='R',[RevisedPlates2],[TotalPlates2-¼]+[TotalPlates2-½]+[TotalPlates2-1]+[RevisedPlates2])*2 As TotalPlates,ISNULL(C1.platesIssued,0)  AS platesIssued,LTRIM(R.Name) As Paper,PARSENAME(PaperWastageFinal2,2)*U.Value1+(PARSENAME(PaperWastageFinal2,1)) As Wastage,PaperConsumptionsheets2 As TotalPaper,ISNULL(C1.paperIssued,0) As paperIssued,P.Code+'MF2*00012' As RefCode,M.Name AS MAC,M.Code AS MCode,MakeReadyTime,Efficiency,ISNULL(C1.SNo,0) AS SNo," & _
"M.StartTime,M.EndTime,C.Forms2-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF2*00012'),0) AS PendingPlan,C.Processing FROM ((((((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.Size2=S.Code) INNER JOIN PaperMaster R ON C.Paper2=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'MF2*00012'=C1.Code) LEFT JOIN AccountMaster A ON P.BookPrinter=A.Code LEFT JOIN MachineMaster M ON M.Code=ISNULL(C1.Machine,'*21046') WHERE Forms2<>0 AND " & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>=2") & " AND LEFT(P.Type,1)<>'O' AND P.Date>='" & GetDate(MhDateInput1.Text) & "'  AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Option1.Value, "1=1", IIf(Option2.Value, "ISNULL(C1.[Plan],C.Forms2)-ISNULL(C1.FormsPrinted,0)<>0", IIf(VchType = 1 And Option3.Value, "ISNULL(C1.[Plan],C.Forms2)-ISNULL(C1.FormsPrinted,0)=0", "1=1"))) & _
           " AND " & IIf(Combo4.ListIndex = 0, "C.Processing='N'", IIf(Combo4.ListIndex = 1, "C.Processing='O'", "1=1")) & " AND " & IIf(Option1.Value, "1=1", IIf(Option2.Value, "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)>0", "1=1")) & " AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='S'", "RIGHT(P.Type,1)='P'", "1=1") & " AND " & IIf(Check2.Value = 0, "C1.Machine IS NOT NULL", IIf(MC = "All", "1=1", "(C1.Machine IN ('" & MC & "','*21046') OR " & IIf(Check2.Value, "(C1.Machine IS NULL)", "C1.Machine IN ('')") & ")")) & " AND " & IIf(Check2.Value = 0, "M.Code<>'*21046'", "1=1") & " AND " & _
           IIf(VchType = 2 And Option1.Value, "1=1", IIf(VchType = 2 And Option2.Value, "IIF(Processing='R',[RevisedPlates2],[TotalPlates2-¼]+[TotalPlates2-½]+[TotalPlates2-1]+[RevisedPlates2])*2-ISNULL(C1.platesIssued,0)<>0", IIf(VchType = 2 And Option3.Value, "((IIF(Processing='R',[RevisedPlates2],[TotalPlates2-¼]+[TotalPlates2-½]+[TotalPlates2-1]+[RevisedPlates2])*2-ISNULL(C1.platesIssued,0)=0) OR (P.DeliveredQuantityC+P.DeliveredQuantityB)>0)", "1=1"))) & " AND " & IIf(Check2.Value, "1=1", IIf(MC = "All", "1=1", "'" & MC & "'=C1.Machine"))
    'MF4
    SQL = SQL + " UNION " & _
          "SELECT (Select Name From ElementMaster Where Code='*00013') As Element,RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-MF','UFG-MF')+'/'+LTRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],'4' As Col,A.Name As Party,C.Forms4 As TotalForms,ISNULL(C1.[Plan],C.Forms4)AS [Plan],ISNULL(C1.formsPrinted,0)  AS formsPrinted,C.ActualQuantity As TotalQty,(P.DeliveredQuantityC+P.DeliveredQuantityB) As QtyIssued,IIF(PlateType4='1','Deepatch',IIF(PlateType4='2','PS',IIF(PlateType4='3','Wipeon','CTP'))) As Plate,IIF(Processing='R',[RevisedPlates4],[TotalPlates4-¼]+[TotalPlates4-½]+[TotalPlates4-1]+[RevisedPlates4])*4 As TotalPlates,ISNULL(C1.platesIssued,0)  AS platesIssued,LTRIM(R.Name) As Paper,PARSENAME(PaperWastageFinal4,2)*U.Value1+(PARSENAME(PaperWastageFinal4,1)) As Wastage,PaperConsumptionsheets4 As TotalPaper,ISNULL(C1.paperIssued,0) AS paperIssued,P.Code+'MF4*00013' As RefCode,M.Name AS MAC,M.Code AS MCode,MakeReadyTime,Efficiency,ISNULL(C1.SNo,0) AS SNo, " & _
"M.StartTime,M.EndTime,C.Forms4-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF4*00013'),0) AS PendingPlan,C.Processing FROM ((((((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.Size4=S.Code) INNER JOIN PaperMaster R ON C.Paper4=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'MF4*00013'=C1.Code) LEFT JOIN AccountMaster A ON P.BookPrinter=A.Code LEFT JOIN MachineMaster M ON M.Code=ISNULL(C1.Machine,'*21046') WHERE Forms4<>0 AND " & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>=4") & " AND LEFT(P.Type,1)<>'O' AND P.Date>='" & GetDate(MhDateInput1.Text) & "'  AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Option1.Value, "1=1", IIf(Option2.Value, "ISNULL(C1.[Plan],C.Forms4)-ISNULL(C1.FormsPrinted,0)<>0", IIf(VchType = 1 And Option3.Value, "ISNULL(C1.[Plan],C.Forms4)-ISNULL(C1.FormsPrinted,0)=0", "1=1"))) & _
          " AND " & IIf(Combo4.ListIndex = 0, "C.Processing='N'", IIf(Combo4.ListIndex = 1, "C.Processing='O'", "1=1")) & " AND " & IIf(Option1.Value, "1=1", IIf(Option2.Value, "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)>0", "1=1")) & " AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='S'", "RIGHT(P.Type,1)='P'", "1=1") & " AND " & IIf(Check2.Value = 0, "C1.Machine IS NOT NULL", IIf(MC = "All", "1=1", "(C1.Machine IN ('" & MC & "','*21046') OR " & IIf(Check2.Value, "(C1.Machine IS NULL)", "C1.Machine IN ('')") & ")")) & " AND " & IIf(Check2.Value = 0, "M.Code<>'*21046'", "1=1") & " AND " & _
          IIf(VchType = 2 And Option1.Value, "1=1", IIf(VchType = 2 And Option2.Value, "IIF(Processing='R',[RevisedPlates4],[TotalPlates4-¼]+[TotalPlates4-½]+[TotalPlates4-1]+[RevisedPlates4])*4-ISNULL(C1.platesIssued,0)<>0", IIf(VchType = 2 And Option3.Value, "((IIF(Processing='R',[RevisedPlates4],[TotalPlates4-¼]+[TotalPlates4-½]+[TotalPlates4-1]+[RevisedPlates4])*4-ISNULL(C1.platesIssued,0)=0) OR (P.DeliveredQuantityC+P.DeliveredQuantityB)>0)", "1=1"))) & " AND " & IIf(Check2.Value, "1=1", IIf(MC = "All", "1=1", "'" & MC & "'=C1.Machine"))

    'ME
    SQL = SQL + " UNION " & _
              "SELECT (Select Name From ElementMaster Where Code=C.Element) As Element," & _
              "RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-ME','UFG-ME')+'/'+LTRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],IIF(LTRIM(BackPrintingType)=0,LTRIM(FrontPrintingType),LTRIM(FrontPrintingType)+' + '+LTRIM(BackPrintingType)) As Col,A.Name As Party,Sets As TotalForms,ISNULL(C1.[Plan],Sets) AS [Plan],ISNULL(C1.formsPrinted,0) AS formsPrinted,C.ActualQuantity As TotalQty,(P.DeliveredQuantityC+P.DeliveredQuantityB) As QtyIssued,IIF(PlateType='1','Deepatch',IIF(PlateType='2','PS',IIF(PlateType='3','Wipeon','CTP'))) As Plate,(FrontPrintingType+BackPrintingType) As TotalPlates,ISNULL(C1.platesIssued,0) AS platesIssued, LTRIM(R.Name) As Paper,PARSENAME(PaperWastageFinal,2)*U.Value1+(PARSENAME(PaperWastageFinal,1)) As Wastage,PaperConsumptionsheets As TotalPaper,ISNULL(C1.paperIssued,0) AS paperIssued,P.Code+'ME1'+C.Element As RefCode,M.Name AS MAC,M.Code AS MCode,MakeReadyTime,Efficiency,ISNULL(C1.SNo,0) AS SNo, " & _
            "M.StartTime,M.EndTime,Sets-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'ME1*00011'),0) AS PendingPlan,C.Processing FROM ((((((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.[Size]=S.Code) INNER JOIN PaperMaster R ON C.Paper=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'ME1'+C.Element=C1.Code) LEFT JOIN AccountMaster A ON P.TitlePrinter=A.Code LEFT JOIN MachineMaster M ON M.Code=ISNULL(C1.Machine,'*21046') WHERE Sets<>0 AND " & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>=LTRIM(FrontPrintingType)") & " AND LEFT(P.Type,1)<>'O' AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND  " & _
            IIf(Option1.Value, "1=1", IIf(Option2.Value, "ISNULL(C1.[Plan],Sets)-ISNULL(C1.FormsPrinted,0)<>0", IIf(VchType = 1 And Option3.Value, "ISNULL(C1.[Plan],Sets)-ISNULL(C1.FormsPrinted,0)=0", "1=1"))) & _
              " AND " & IIf(Combo4.ListIndex = 0, "C.Processing='N'", IIf(Combo4.ListIndex = 1, "C.Processing='O'", "1=1")) & " AND " & IIf(Option1.Value, "1=1", IIf(Option2.Value, "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)>0", "1=1")) & " AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='S'", "RIGHT(P.Type,1)='P'", "1=1") & " AND " & IIf(Check2.Value = 0, "C1.Machine IS NOT NULL", IIf(MC = "All", "1=1", "(C1.Machine IN ('" & MC & "','*21046') OR " & IIf(Check2.Value, "(C1.Machine IS NULL)", "C1.Machine IN ('')") & ")")) & " AND " & IIf(Check2.Value = 0, "M.Code<>'*21046'", "1=1") & " AND " & _
              IIf(VchType = 2 And Option1.Value, "1=1", IIf(VchType = 2 And Option2.Value, "(FrontPrintingType+BackPrintingType) -ISNULL(C1.platesIssued,0)<>0", IIf(VchType = 2 And Option3.Value, "(((FrontPrintingType+BackPrintingType)-ISNULL(C1.platesIssued,0)=0) OR (P.DeliveredQuantityC+P.DeliveredQuantityB)>0)", "1=1"))) & " AND " & IIf(Check2.Value, "1=1", IIf(MC = "All", "1=1", "'" & MC & "'=C1.Machine"))
    'CF
    SQL = SQL + " UNION " & _
              "SELECT (Select Name From ElementMaster Where Code='*00015') As Element," & _
              "RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-CF','UFG-CF')+'/'+LTRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],Convert(nvarchar,C.FrontPrintingColor)+Convert(nvarchar,C.BackPrintingColor) As Col,A.Name As Party,1 As TotalForms,ISNULL(C1.[Plan],1) AS [Plan],ISNULL(C1.formsPrinted,0) AS formsPrinted,C.ActualQuantity As TotalQty,(P.DeliveredQuantityC+P.DeliveredQuantityB) As QtyIssued,IIF(PlateType='1','Deepatch',IIF(PlateType='2','PS',IIF(PlateType='3','Wipeon','CTP'))) As Plate,C.TotalPlates As TotalPlates,ISNULL(C1.PlatesIssued,0) AS PlatesIssued,LTRIM(R.Name) As Paper, PARSENAME(PaperWastageFinal,2)*U.Value1+(PARSENAME(PaperWastageFinal,1)) As Wastage,PaperConsumptionsheets As TotalPaper,ISNULL(C1.paperIssued,0) AS paperIssued,P.Code+'CF1*00015' As RefCode,M.Name AS MAC,M.Code AS MCode,MakeReadyTime,Efficiency,ISNULL(C1.SNo,0) AS SNo, " & _
              "M.StartTime,M.EndTime,1-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'CF1*00015'),0) AS PendingPlan,C.Plate As Processing FROM (((((((BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code)INNER JOIN BookPOChild0901 C9 ON C.Code=C9.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.[Size]=S.Code) INNER JOIN PaperMaster R ON C.Paper=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'CF1*00015'=C1.Code) LEFT JOIN AccountMaster A ON P.TitlePrinter=A.Code LEFT JOIN MachineMaster M ON M.Code=ISNULL(C1.Machine,'*21046') WHERE (" & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>= LTRIM(C.FrontPrintingColor)") & "  OR " & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>= LTRIM(C.BackPrintingColor)") & " ) AND LEFT(P.Type,1)<>'O' AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & _
              "' AND " & IIf(Combo4.ListIndex = 0, "C.Plate='N'", IIf(Combo4.ListIndex = 1, "C.Plate='O'", "1=1")) & " AND " & IIf(Option1.Value, "1=1", IIf(Option2.Value, "1-ISNULL(C1.FormsPrinted,0)<>0", "1-ISNULL(C1.FormsPrinted,0)=0")) & " AND " & IIf(Option1.Value, "1=1", IIf(Option2.Value, "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)>0", "1=1")) & " AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='S'", "RIGHT(P.Type,1)='P'", "1=1") & " AND " & IIf(Check2.Value = 0, "C1.Machine IS NOT NULL", IIf(MC = "All", "1=1", "(C1.Machine IN ('" & MC & "','*21046') OR " & IIf(Check2.Value, "(C1.Machine IS NULL)", "C1.Machine IN ('')") & ")")) & " AND " & IIf(Check2.Value = 0, "M.Code<>'*21046'", "1=1") & " AND " & _
              IIf(VchType = 2 And Option1.Value, "1=1", IIf(VchType = 2 And Option2.Value, "((C.TotalPlates-ISNULL(C1.platesIssued,0)<>0) OR (P.DeliveredQuantityC+P.DeliveredQuantityB)>0)", IIf(VchType = 2 And Option3.Value, "C.TotalPlates-ISNULL(C1.platesIssued,0)=0", "1=1"))) & _
              " AND " & IIf(Check2.Value, "1=1", IIf(MC = "All", "1=1", "'" & MC & "'=C1.Machine"))
    'Pending
    'MF1
    SQL = SQL + " UNION " & _
              "SELECT Distinct (Select Name From ElementMaster Where Code='*00011') As Element," & _
              "RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-MF','UFG-MF')+'/'+LTRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],'1' As Col,A.Name As Party,C.Forms1 As TotalForms,C.Forms1-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF1*00011'),0) AS [Plan],0 AS formsPrinted,C.ActualQuantity As TotalQty,(P.DeliveredQuantityC+P.DeliveredQuantityB) As QtyIssued,IIF(PlateType1='1','Deepatch',IIF(PlateType1='2','PS',IIF(PlateType1='3','Wipeon','CTP'))) As Plate,IIF(Processing='R',[RevisedPlates1],[TotalPlates1-¼]+[TotalPlates1-½]+[TotalPlates1-1]+[RevisedPlates1]) As TotalPlates,0 AS platesIssued,LTRIM(R.Name) As Paper,PARSENAME(PaperWastageFinal1,2)*U.Value1+(PARSENAME(PaperWastageFinal1,1)) As Wastage,PaperConsumptionsheets1 As TotalPaper,ISNULL(C1.paperIssued,0) As paperIssued,P.Code+'MF1*00011' As RefCode,M.Name AS MAC,M.Code AS MCode,MakeReadyTime,Efficiency,0 AS SNo, " & _
              "M.StartTime,M.EndTime,C.Forms1-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF1*00011'),0) AS PendingPlan,C.Processing FROM ((((((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.Size1=S.Code) INNER JOIN PaperMaster R ON C.Paper1=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'MF1*00011'=C1.Code) LEFT JOIN AccountMaster A ON P.BookPrinter=A.Code LEFT JOIN MachineMaster M ON M.Code=('*21046') WHERE Forms1<>0 AND " & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>=1") & " AND LEFT(P.Type,1)<>'O' AND P.Date>='" & GetDate(MhDateInput1.Text) & "'  AND P.Date<='" & GetDate(MhDateInput2.Text) & _
              "' AND " & IIf(Option2.Value, "C.Forms1-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF1*00011'),0)<>0", "C.Forms1-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF1*00011'),0)<>0") & _
              " AND " & IIf(Combo4.ListIndex = 0, "C.Processing='N'", IIf(Combo4.ListIndex = 1, "C.Processing='O'", "1=1")) & " AND " & IIf(Option2.Value, "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)>0", "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)<=0") & "AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='S'", "RIGHT(P.Type,1)='P'", "1=1") & " AND " & IIf(Check2.Value = 0, "C1.Machine IS NOT NULL", IIf(MC = "All", "1=1", "(ISNULL(C1.Machine,'*21046') IN ('" & MC & "','*21046') OR " & IIf(Check2.Value, "(C1.Machine IS NULL)", "C1.Machine IN ('')") & ")")) & " AND " & IIf(Check2.Value = 0, "M.Code<>'*21046'", "1=1") & " AND " & _
              IIf(VchType = 2 And Option1.Value, "1=1", IIf(VchType = 2 And Option2.Value, "IIF(Processing='R',[RevisedPlates1],[TotalPlates1-¼]+[TotalPlates1-½]+[TotalPlates1-1]+[RevisedPlates1])-ISNULL(C1.platesIssued,0)<>0", IIf(VchType = 2 And Option3.Value, "IIF(Processing='R',[RevisedPlates1],[TotalPlates1-¼]+[TotalPlates1-½]+[TotalPlates1-1]+[RevisedPlates1])-ISNULL(C1.platesIssued,0)=0", "1=1"))) & " AND " & IIf(Check2.Value, "1=1", IIf(MC = "All", "1=1", "M.Code <>'*21046'"))
    'MF2
    SQL = SQL + " UNION " & _
              "SELECT Distinct (Select Name From ElementMaster Where Code='*00012') As Element," & _
              "RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-MF','UFG-MF')+'/'+LTRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],'2' As Col,A.Name As Party,C.Forms2 As TotalForms,C.Forms2-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF2*00012'),0) AS [Plan],0  AS formsPrinted,C.ActualQuantity As TotalQty,(P.DeliveredQuantityC+P.DeliveredQuantityB) As QtyIssued,IIF(PlateType2='1','Deepatch',IIF(PlateType2='2','PS',IIF(PlateType2='3','Wipeon','CTP'))) As Plate,IIF(Processing='R',[RevisedPlates2],[TotalPlates2-¼]+[TotalPlates2-½]+[TotalPlates2-1]+[RevisedPlates2])*2 As TotalPlates,0 AS platesIssued,LTRIM(R.Name) As Paper,PARSENAME(PaperWastageFinal2,2)*U.Value1+(PARSENAME(PaperWastageFinal2,1)) As Wastage,PaperConsumptionsheets2 As TotalPaper,ISNULL(C1.paperIssued,0) As paperIssued,P.Code+'MF2*00012' As RefCode,M.Name AS MAC,M.Code AS MCode,MakeReadyTime,Efficiency,0 AS SNo, " & _
              "M.StartTime,M.EndTime,C.Forms2-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF2*00012'),0) AS PendingPlan,C.Processing FROM ((((((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.Size2=S.Code) INNER JOIN PaperMaster R ON C.Paper2=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'MF2*00012'=C1.Code) LEFT JOIN AccountMaster A ON P.BookPrinter=A.Code LEFT JOIN MachineMaster M ON M.Code=('*21046') WHERE Forms2<>0 AND " & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>=2") & " AND LEFT(P.Type,1)<>'O' AND P.Date>='" & GetDate(MhDateInput1.Text) & "'  AND P.Date<='" & GetDate(MhDateInput2.Text) & _
              "' AND " & IIf(Option2.Value, "C.Forms2-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF2*00012'),0)<>0", "C.Forms2-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF2*00012'),0)<>0") & _
              " AND " & IIf(Combo4.ListIndex = 0, "C.Processing='N'", IIf(Combo4.ListIndex = 1, "C.Processing='O'", "1=1")) & " AND " & IIf(Option2.Value, "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)>0", "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)<=0") & " AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='S'", "RIGHT(P.Type,1)='P'", "1=1") & " AND " & IIf(Check2.Value = 0, "C1.Machine IS NOT NULL", IIf(MC = "All", "1=1", "(ISNULL(C1.Machine,'*21046') IN ('" & MC & "','*21046') OR " & IIf(Check2.Value, "(C1.Machine IS NULL)", "C1.Machine IN ('')") & ")")) & " AND " & IIf(Check2.Value = 0, "M.Code<>'*21046'", "1=1") & " AND " & _
              IIf(VchType = 2 And Option1.Value, "1=1", IIf(VchType = 2 And Option2.Value, "IIF(Processing='R',[RevisedPlates2],[TotalPlates2-¼]+[TotalPlates2-½]+[TotalPlates2-1]+[RevisedPlates2])*2-ISNULL(C1.platesIssued,0)<>0", IIf(VchType = 2 And Option3.Value, "IIF(Processing='R',[RevisedPlates2],[TotalPlates2-¼]+[TotalPlates2-½]+[TotalPlates2-1]+[RevisedPlates2])*2-ISNULL(C1.platesIssued,0)=0", "1=1"))) & " AND " & IIf(Check2.Value, "1=1", IIf(MC = "All", "1=1", "M.Code <>'*21046'"))
    'MF4
    SQL = SQL + " UNION " & _
              "SELECT Distinct (Select Name From ElementMaster Where Code='*00013') As Element," & _
              "RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-MF','UFG-MF')+'/'+LTRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],'4' As Col,A.Name As Party,C.Forms4 As TotalForms,C.Forms4-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF4*00013'),0) AS [Plan],0  AS formsPrinted,C.ActualQuantity As TotalQty,(P.DeliveredQuantityC+P.DeliveredQuantityB) As QtyIssued,IIF(PlateType4='1','Deepatch',IIF(PlateType4='2','PS',IIF(PlateType4='3','Wipeon','CTP'))) As Plate,IIF(Processing='R',[RevisedPlates4],[TotalPlates4-¼]+[TotalPlates4-½]+[TotalPlates4-1]+[RevisedPlates4])*4 As TotalPlates,0 AS platesIssued,LTRIM(R.Name) As Paper,PARSENAME(PaperWastageFinal4,2)*U.Value1+(PARSENAME(PaperWastageFinal4,1)) As Wastage,PaperConsumptionsheets4 As TotalPaper,ISNULL(C1.paperIssued,0) AS paperIssued,P.Code+'MF4*00013' As RefCode,M.Name AS MAC,M.Code AS MCode,MakeReadyTime,Efficiency,0 AS SNo, " & _
              "M.StartTime,M.EndTime,C.Forms4-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF4*00013'),0) AS PendingPlan,C.Processing FROM ((((((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.Size4=S.Code) INNER JOIN PaperMaster R ON C.Paper4=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'MF4*00013'=C1.Code) LEFT JOIN AccountMaster A ON P.BookPrinter=A.Code LEFT JOIN MachineMaster M ON M.Code=('*21046') WHERE Forms4<>0 AND " & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>=4") & " AND LEFT(P.Type,1)<>'O' AND P.Date>='" & GetDate(MhDateInput1.Text) & "'  AND P.Date<='" & GetDate(MhDateInput2.Text) & _
              "' AND " & IIf(Option2.Value, "C.Forms4-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF4*00013'),0)<>0", "C.Forms4-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'MF4*00013'),0)<>0") & _
              " AND " & IIf(Combo4.ListIndex = 0, "C.Processing='N'", IIf(Combo4.ListIndex = 1, "C.Processing='O'", "1=1")) & " AND " & IIf(Option2.Value, "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)>0", "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)<=0") & " AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='S'", "RIGHT(P.Type,1)='P'", "1=1") & " AND " & IIf(Check2.Value = 0, "C1.Machine IS NOT NULL", IIf(MC = "All", "1=1", "(ISNULL(C1.Machine,'*21046') IN ('" & MC & "','*21046') OR " & IIf(Check2.Value, "(C1.Machine IS NULL)", "C1.Machine IN ('')") & ")")) & " AND " & IIf(Check2.Value = 0, "M.Code<>'*21046'", "1=1") & " AND " & _
              IIf(VchType = 2 And Option1.Value, "1=1", IIf(VchType = 2 And Option2.Value, "IIF(Processing='R',[RevisedPlates4],[TotalPlates4-¼]+[TotalPlates4-½]+[TotalPlates4-1]+[RevisedPlates4])*4-ISNULL(C1.platesIssued,0)<>0", IIf(VchType = 2 And Option3.Value, "IIF(Processing='R',[RevisedPlates4],[TotalPlates4-¼]+[TotalPlates4-½]+[TotalPlates4-1]+[RevisedPlates4])*4-ISNULL(C1.platesIssued,0)=0", "1=1"))) & " AND " & IIf(Check2.Value, "1=1", IIf(MC = "All", "1=1", "M.Code <>'*21046'"))
    'ME
    SQL = SQL + " UNION " & _
              "SELECT Distinct (Select Name From ElementMaster Where Code=C.Element) As Element," & _
              "RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-ME','UFG-ME')+'/'+LTRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],IIF(LTRIM(BackPrintingType)=0,LTRIM(FrontPrintingType),LTRIM(FrontPrintingType)+' + '+LTRIM(BackPrintingType)) As Col,A.Name As Party,Sets As TotalForms,C.Sets-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'ME1'+C.Element),0) AS [Plan],0 AS formsPrinted,C.ActualQuantity As TotalQty,(P.DeliveredQuantityC+P.DeliveredQuantityB) As QtyIssued,IIF(PlateType='1','Deepatch',IIF(PlateType='2','PS',IIF(PlateType='3','Wipeon','CTP'))) As Plate,(FrontPrintingType+BackPrintingType) As TotalPlates,0 AS platesIssued, LTRIM(R.Name) As Paper,PARSENAME(PaperWastageFinal,2)*U.Value1+(PARSENAME(PaperWastageFinal,1)) As Wastage,PaperConsumptionsheets As TotalPaper,ISNULL(C1.paperIssued,0) AS paperIssued,P.Code+'ME1'+C.Element As RefCode," & _
              "M.Name AS MAC,M.Code AS MCode,MakeReadyTime,Efficiency,0 AS SNo, " & _
              "M.StartTime,M.EndTime,Sets-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'ME1'+C.Element),0) AS PendingPlan,C.Processing FROM ((((((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.[Size]=S.Code) INNER JOIN PaperMaster R ON C.Paper=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'ME1'+C.Element=C1.Code) LEFT JOIN AccountMaster A ON P.TitlePrinter=A.Code LEFT JOIN MachineMaster M ON M.Code=ISNULL(C1.Machine,'*21046') WHERE Sets<>0 AND " & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>=LTRIM(FrontPrintingType)") & " AND LEFT(P.Type,1)<>'O' AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND " & _
              "P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & IIf(Option2.Value, "C.Sets-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'ME1'+C.Element),0)<>0", "C.Sets-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'ME1'+C.Element),0)<>0") & " AND " & IIf(Option2.Value, "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)>0", "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)<=0") & _
              " AND " & IIf(Combo4.ListIndex = 0, "C.Processing='N'", IIf(Combo4.ListIndex = 1, "C.Processing='O'", "1=1")) & " AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='S'", "RIGHT(P.Type,1)='P'", "1=1") & " AND " & IIf(Check2.Value = 0, "C1.Machine IS NOT NULL", IIf(MC = "All", "1=1", "(ISNULL(C1.Machine,'*21046') IN ('" & MC & "','*21046') OR " & IIf(Check2.Value, "(C1.Machine IS NULL)", "C1.Machine IN ('')") & ")")) & " AND " & IIf(Check2.Value = 0, "M.Code<>'*21046'", "1=1") & " AND " & _
              IIf(VchType = 2 And Option1.Value, "1=1", IIf(VchType = 2 And Option2.Value, "(FrontPrintingType+BackPrintingType)-ISNULL(C1.platesIssued,0)<>0", IIf(VchType = 2 And Option3.Value, "(FrontPrintingType+BackPrintingType)-ISNULL(C1.platesIssued,0)=0", "1=1"))) & " AND " & IIf(Check2.Value, "1=1", IIf(MC = "All", "1=1", "M.Code <>'*21046'"))
    'CF
    SQL = SQL + " UNION " & _
             "SELECT Distinct (Select Name From ElementMaster Where Code='*00015') As Element," & _
             "RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-CF','UFG-CF')+'/'+LTRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],Convert(nvarchar,C.FrontPrintingColor)+Convert(nvarchar,C.BackPrintingColor) As Col,A.Name As Party,1 As TotalForms,1-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'CF1*00015'),0) AS [Plan],ISNULL(0,0) AS formsPrinted,C.ActualQuantity As TotalQty,(P.DeliveredQuantityC+P.DeliveredQuantityB) As QtyIssued,IIF(PlateType='1','Deepatch',IIF(PlateType='2','PS',IIF(PlateType='3','Wipeon','CTP'))) As Plate,C.TotalPlates As TotalPlates,0 AS PlatesIssued,LTRIM(R.Name) As Paper, PARSENAME(PaperWastageFinal,2)*U.Value1+(PARSENAME(PaperWastageFinal,1)) As Wastage,PaperConsumptionsheets As TotalPaper,ISNULL(C1.paperIssued,0) AS paperIssued,P.Code+'CF1*00015' As RefCode,M.Name AS MAC,M.Code AS MCode,MakeReadyTime,Efficiency,0 AS SNo," & _
             "M.StartTime,M.EndTime,1-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'CF1*00015'),0) AS PendingPlan,C.Plate As Processing FROM (((((((BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code)INNER JOIN BookPOChild0901 C9 ON C.Code=C9.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.[Size]=S.Code) INNER JOIN PaperMaster R ON C.Paper=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'CF1*00015'=C1.Code) LEFT JOIN AccountMaster A ON P.TitlePrinter=A.Code LEFT JOIN MachineMaster M ON M.Code=('*21046') WHERE (" & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>= LTRIM(C.FrontPrintingColor)") & "  OR " & IIf(MC = "All", "1=1", "(Select Units From MachineMaster Where Code='" & MC & "')>= LTRIM(C.BackPrintingColor)") & " ) AND LEFT(P.Type,1)<>'O' AND P.Date>='" & GetDate(MhDateInput1.Text) & "' AND P.Date<='" & GetDate(MhDateInput2.Text) & "' AND " & _
              IIf(Option2.Value, "1-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'CF1*00015'),0)<>0", "1-ISNULL((Select Sum ([Plan]) From BookPOChild0501 Where Code=P.Code+'CF1*00015'),0)<>0") & " AND " & IIf(Combo4.ListIndex = 0, "C.Plate='N'", IIf(Combo4.ListIndex = 1, "C.Plate='O'", "1=1")) & " AND " & IIf(Option1.Value, "1=1", IIf(Option2.Value, "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)>0", "(C.ActualQuantity-P.DeliveredQuantityC+P.DeliveredQuantityB)<=0")) & " AND " & Choose(Combo1.ListIndex + 1, "RIGHT(P.Type,1)='S'", "RIGHT(P.Type,1)='P'", "1=1") & " AND " & IIf(Check2.Value = 0, "C1.Machine IS NOT NULL", IIf(MC = "All", "1=1", "(C1.Machine IN ('" & MC & "','*21046') OR " & IIf(Check2.Value, "(C1.Machine IS NULL)", "C1.Machine IN ('')") & ")")) & " AND " & IIf(Check2.Value = 0, "M.Code<>'*21046'", "1=1") & _
              " AND " & IIf(VchType = 2 And Option1.Value, "1=1", IIf(VchType = 2 And Option2.Value, "C.TotalPlates-ISNULL(C1.platesIssued,0)<>0", IIf(VchType = 2 And Option3.Value, "C.TotalPlates-ISNULL(C1.platesIssued,0)=0", "1=1"))) & " AND " & IIf(Check2.Value, "1=1", "M.Code <>'*21046'") & _
              " ORDER BY MAC,SNo,TotalForms,RefDate"
    Screen.MousePointer = vbHourglass
    If rstProductionScheduling.State = adStateOpen Then rstProductionScheduling.Close
    rstProductionScheduling.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
    If rstProductionScheduling.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
    With fpSpread1
       .ClearRange -1, 1, .MaxCols, .MaxRows, False
       .MaxRows = (rstProductionScheduling.RecordCount + rstMachineMaster.RecordCount)
       .ColsFrozen = 0
       rstProductionScheduling.MoveFirst
       Do While Not rstProductionScheduling.EOF
            i = i + 1
            If mcN = rstProductionScheduling.Fields("MAC").Value Then TDBNumber2 = TDBNumber2 + 1 'Data Count
            If mcN <> rstProductionScheduling.Fields("MAC").Value Or i = 1 Then
            .Col = 1: .Row = i: .FontBold = True: .FontSize = 13: .BackColor = &H80C0FF: .FontUnderline = True: .ForeColor = vbRed: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignLeft: .TypeVAlign = TypeVAlignCenter
                If Not IsNull(rstProductionScheduling.Fields("MAC").Value) Then
                mcN = rstProductionScheduling.Fields("MAC").Value: TDBTime1 = rstProductionScheduling.Fields("StartTime"): TDBTime2 = rstProductionScheduling.Fields("EndTime"): timeFlag = False
                    For C = 1 To .MaxCols
                    .Col = C: .Row = i:  .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignLeft: FontUnderline = True
                    .SetText C, i, ""
                    Next
                    'Col Span
                    '.AddCellSpan 1, i, 3, 1
                    .SetText 1, i, " " & rstProductionScheduling.Fields("MAC").Value: K = 0: .SetText .RowHeaderCols - 1, i, " "
                    'Col Span
                    '.AddCellSpan 4, i, 30, 1
                    .Col = 4: .Row = i: .FontBold = True: .FontSize = 11: .BackColor = &H80C0FF: .FontUnderline = True: .ForeColor = vbRed: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignLeft:
                    .SetText 4, i, " Shift Time From: " & Format(TDBTime1.Value, "HH:MM") & " TO :" & Format(TDBTime2.Value, "HH:MM") & " "
                    .Col = 9: .Row = i: .CellType = CellTypeEdit: i = i + 1
                Else
                     .SetText 1, i, "Z-Machine To Be Decide": i = i + 1: mcN = "Z-Machine To Be Decide": K = 0: .SetText .RowHeaderCols - 1, i, " "
                End If
            End If
                K = K + 1
                .SetText .RowHeaderCols - 1, i, K
                .SetText 1, i, rstProductionScheduling.Fields("Party").Value: fpSpread1.Col = 1: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstProductionScheduling.Fields("Party").Value) > 24 Then fpSpread1.RowHeight(i) = 24: fpSpread1.TypeHAlign = TypeHAlignLeft: fpSpread1.TypeVAlign = TypeVAlignCenter
                .SetText 2, i, rstProductionScheduling.Fields("RefNo").Value
                .SetText 3, i, Format(rstProductionScheduling.Fields("RefDate").Value, "dd-MMM-yy")
                .SetText 4, i, rstProductionScheduling.Fields("Item").Value & " >>>" & rstProductionScheduling.Fields("Element").Value & "<<<": fpSpread1.Col = 1: fpSpread1.Row = i: fpSpread1.CellType = CellTypeStaticText: fpSpread1.TypeTextWordWrap = True: If Len(rstProductionScheduling.Fields("Item").Value & " >>>" & rstProductionScheduling.Fields("Element").Value & "<<<") > 40 Then fpSpread1.RowHeight(i) = 24: fpSpread1.TypeHAlign = TypeHAlignLeft: fpSpread1.TypeVAlign = TypeVAlignCenter
                .SetText 5, i, Left(rstProductionScheduling.Fields("Size").Value, 11)
                .SetText 6, i, rstProductionScheduling.Fields("Col").Value
                .SetText 7, i, Val(rstProductionScheduling.Fields("TotalForms").Value)
            If Val(rstProductionScheduling.Fields("QtyIssued").Value) >= Val(rstProductionScheduling.Fields("TotalQty").Value) And IsNull(rstProductionScheduling.Fields("Plan").Value) Then
                .SetText 8, i, Val(rstProductionScheduling.Fields("TotalForms").Value)
                .SetText 34, i, Val(rstProductionScheduling.Fields("TotalForms").Value)
            ElseIf Not IsNull(rstProductionScheduling.Fields("Plan").Value) Then
                If Val(rstProductionScheduling.Fields("Plan").Value) = 0 Then .SetText 8, i, Val(rstProductionScheduling.Fields("TotalForms").Value): .SetText 34, i, Val(rstProductionScheduling.Fields("TotalForms").Value) Else .SetText 8, i, Val(rstProductionScheduling.Fields("Plan").Value): .SetText 34, i, Val(rstProductionScheduling.Fields("Plan").Value)
            Else
                .SetText 8, i, Val(rstProductionScheduling.Fields("TotalForms").Value)
                .SetText 34, i, Val(rstProductionScheduling.Fields("TotalForms").Value)
            End If
                .Col = 9: .Row = i: .CellType = CellTypeCheckBox: .TypeHAlign = TypeHAlignCenter
                .GetText 8, i, cVal(6) 'Plan
            If Val(rstProductionScheduling.Fields("QtyIssued").Value) >= Val(rstProductionScheduling.Fields("TotalQty").Value) Then
                .SetText 10, i, Val(cVal(6))
                .SetText 11, i, 0
            ElseIf Not IsNull(rstProductionScheduling.Fields("formsPrinted").Value) Then
                .SetText 10, i, Val(rstProductionScheduling.Fields("formsPrinted").Value)
                .SetText 11, i, Val(cVal(6)) - Val(rstProductionScheduling.Fields("formsPrinted").Value)
            Else
                .SetText 10, i, 0 'Printed
                .SetText 11, i, Val(cVal(6)) 'Bal
            End If
            .SetText 12, i, Val(rstProductionScheduling.Fields("TotalQty").Value)
            .SetText 13, i, Val(rstProductionScheduling.Fields("QtyIssued").Value)
            .SetText 14, i, Val(rstProductionScheduling.Fields("TotalQty").Value) - Val(rstProductionScheduling.Fields("QtyIssued").Value)
            .SetText 15, i, rstProductionScheduling.Fields("Plate").Value & " <<< " & IIf(rstProductionScheduling.Fields("Processing").Value = "O", "Old Plate", "New Plate") & " >>> "
            .SetText 16, i, Val(rstProductionScheduling.Fields("TotalPlates").Value)
            If Val(rstProductionScheduling.Fields("QtyIssued").Value) >= Val(rstProductionScheduling.Fields("TotalQty").Value) Then
                .SetText 17, i, Val(rstProductionScheduling.Fields("TotalPlates").Value)
                .SetText 18, i, 0
            ElseIf Not IsNull(rstProductionScheduling.Fields("platesIssued").Value) Then
                .SetText 17, i, Val(rstProductionScheduling.Fields("platesIssued").Value)
                .SetText 18, i, Val(rstProductionScheduling.Fields("TotalPlates").Value) - Val(rstProductionScheduling.Fields("platesIssued").Value)
            Else
                .SetText 17, i, 0
                .SetText 18, i, Val(rstProductionScheduling.Fields("TotalPlates").Value)
            End If
            .SetText 19, i, rstProductionScheduling.Fields("Paper").Value
                .SetText 20, i, Val(rstProductionScheduling.Fields("Wastage").Value)
                .SetText 21, i, Val(rstProductionScheduling.Fields("TotalPaper").Value)
            If Val(rstProductionScheduling.Fields("QtyIssued").Value) > 0 Then
                .SetText 22, i, Val(rstProductionScheduling.Fields("TotalPaper").Value)
                .SetText 23, i, 0
            ElseIf Not IsNull(rstProductionScheduling.Fields("paperIssued").Value) Then
                .SetText 22, i, Val(rstProductionScheduling.Fields("paperIssued").Value)
                .SetText 23, i, Val(rstProductionScheduling.Fields("TotalPaper").Value) - Val(rstProductionScheduling.Fields("paperIssued").Value)
            Else
                .SetText 22, i, 0
                .SetText 23, i, Val(rstProductionScheduling.Fields("TotalPaper").Value)
            End If
            .SetText 24, i, rstProductionScheduling.Fields("RefCode").Value
            If Not IsNull(Trim((rstProductionScheduling.Fields("MCode").Value))) Then
            .SetText 25, i, Trim(rstProductionScheduling.Fields("MCode").Value)
            Else
            .SetText 25, i, "*21046"
            End If
            If Not IsNull(rstProductionScheduling.Fields("SNO").Value) Then
            .SetText 26, i, rstProductionScheduling.Fields("SNO").Value
            Else
            .SetText 26, i, 0
            End If
            .GetText 8, i, cVal(1): .GetText 10, i, cVal(2): .GetText 11, i, cVal(3): .GetText 12, i, cVal(4): .GetText 2, i - 1, cVal(7)
            .SetText 27, i, (cVal(4) * (cVal(1) - cVal(2)))
    If Not IsNull(rstProductionScheduling.Fields("MakeReadyTime").Value) Then
            .SetText 28, i, Int((rstProductionScheduling.Fields("MakeReadyTime").Value / 60) * (cVal(1) - cVal(2)) + (cVal(4) * (cVal(1) - cVal(2))) / rstProductionScheduling.Fields("Efficiency").Value) & ":" & Format(Int(Round((rstProductionScheduling.Fields("MakeReadyTime").Value / 60) * (cVal(1) - cVal(2)) + (cVal(4) * (cVal(1) - cVal(2))) / rstProductionScheduling.Fields("Efficiency").Value - Int((rstProductionScheduling.Fields("MakeReadyTime").Value / 60) * (cVal(1) - cVal(2)) + (cVal(4) * (cVal(1) - cVal(2))) / rstProductionScheduling.Fields("Efficiency").Value), 2) * 60), "00")
                        'M/R Time
                    cVal(5) = Int((rstProductionScheduling.Fields("MakeReadyTime").Value / 60) * (cVal(1) - cVal(2)) + (cVal(4) * (cVal(1) - cVal(2))) / rstProductionScheduling.Fields("Efficiency").Value) & ":" & Format(Int(Round((rstProductionScheduling.Fields("MakeReadyTime").Value / 60) * (cVal(1) - cVal(2)) + (cVal(4) * (cVal(1) - cVal(2))) / rstProductionScheduling.Fields("Efficiency").Value - Int((rstProductionScheduling.Fields("MakeReadyTime").Value / 60) * (cVal(1) - cVal(2)) + (cVal(4) * (cVal(1) - cVal(2))) / rstProductionScheduling.Fields("Efficiency").Value), 2) * 60), "00")
        hh = Left(cVal(5), Len(cVal(5)) - 3): mm = Right(cVal(5), 2)
                        cVal(5) = (Val(Round((((rstProductionScheduling.Fields("MakeReadyTime").Value / 60) * (cVal(1) - cVal(2))) + ((cVal(4) * (cVal(1) - cVal(2))) / rstProductionScheduling.Fields("Efficiency").Value)), 2)))
        
        If cVal(7) = "" And cVal(5) <> 0 Then  'M/R Time i = 2 And
            .SetText 29, i, Format((MhDateInput2.Value + (TDBTime1.Value)), "dd-MM-yyyy hh:mm AM/PM")
        ElseIf cVal(5) <> 0 And sDate < MhDateInput2.Value Then
            .SetText 29, i, Format((MhDateInput2.Value + (TDBTime1.Value)), "dd-MM-yyyy hh:mm AM/PM")
        ElseIf cVal(5) <> 0 Then
            .SetText 29, i, Format(sDate, "dd-MM-yyyy hh:mm AM/PM")
        Else
            .SetText 29, i, ""
        End If
        If cVal(5) <> 0 Then Call FindDate
        If i = 2 And cVal(5) <> 0 Then
            .SetText 30, i, Format(eDate, "dd-MM-yyyy hh:mm AM/PM")
                      'sDate = Format(MhDateInput2.Value + (9 / TDBTime5.Value) + (Round((((rstProductionScheduling.Fields("MakeReadyTime").Value / 60) * (cVal(1) - cVal(2))) + ((cVal(4) * (cVal(1) - cVal(2))) / rstProductionScheduling.Fields("Efficiency").Value)), 2) / TDBTime5.Value), "dd-MM-yyyy hh:mm AM/PM")
        ElseIf cVal(5) <> 0 Then
            .SetText 30, i, Format(eDate, "dd-MM-yyyy hh:mm AM/PM")
        Else
            .SetText 30, i, ""
        End If
        End If
            .SetText 31, i, rstProductionScheduling.Fields("MAC").Value
            .SetText 32, i, Val(rstProductionScheduling.Fields("PendingPlan").Value)
            .SetText 33, i, Val(rstProductionScheduling.Fields("TotalQty").Value)
            rstProductionScheduling.MoveNext
        Loop
        .ColsFrozen = 0
    End With
    Call Total_Click
Screen.MousePointer = vbNormal
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
Private Sub Total_Click()
    Dim i As Integer, j As Integer, cVal As Variant, n As Integer, R As Long, C As Long, Cols As Long, nVal As Variant
    Dim cValue(1 To 33) As Variant, cValTotal(1 To 33) As Variant
    With fpSpread1
    n = 0
        If .DataRowCnt = 0 Then Exit Sub
        n = .DataRowCnt:
    For i = 1 To .DataRowCnt 'Unhide All
        .GetText 4, i, cVal
        If TotalFlag = False Then .Row = i: .RowHidden = False
        If cVal = "Grand Total" Then fpSpread1.DeleteRows i, 1: n = n - 1
    Next

    .MaxRows = n + 1
    C = Combo3.ListIndex + 1
    'Get Value
For i = 1 To .DataRowCnt
        If Combo3.ListIndex >= 0 Then .GetText C, i, cVal
        For j = 1 To 33
            .GetText j, i, cValue(j)
            .GetText 7, i, cValue(7)
        Next j
                 
            .GetText C, i, cVal
            .GetText 2, i, nVal
            If nVal = "" Then n = n - 1
            .GetText 4, i, nVal
            If nVal = "Grand Total" Then fpSpread1.DeleteRows .DataRowCnt, 1:
'Filter
            .GetText 2, i, nVal
    If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
            If cVal <> "" And nVal <> "" Then .Row = i: .RowHidden = True: n = n - 1 'Hide Filter
    Else
            .Row = i
            If Not .RowHidden Then
                For j = 7 To 27
                    cValTotal(j) = Val(cValTotal(j)) + Val(cValue(j)) '(7 To 27)
                Next j
                If cValue(28) <> "" Then cValTotal(28) = cValTotal(28) + Left(cValue(28), (Len(cValue(28)) - 3)) + Right(cValue(28), 2) / 60 '28
                For j = 32 To 33
                    cValTotal(j) = Val(cValTotal(j)) + Val(cValue(j)) '(32 To 33)
                Next j
            End If
    End If
Next
    TDBNumber2 = n 'Data Count


        'Set Value
                .SetText 4, i, "Grand Total"
                .SetText 7, i, cValTotal(7)
                .SetText 8, i, cValTotal(8)
                .SetText 10, i, cValTotal(10)
                .SetText 11, i, cValTotal(11)
                .SetText 12, i, cValTotal(12)
                .SetText 13, i, cValTotal(13)
                .SetText 14, i, cValTotal(14)
                .SetText 16, i, cValTotal(16)
                .SetText 17, i, cValTotal(17)
                .SetText 18, i, cValTotal(18)
                .SetText 20, i, cValTotal(20)
                .SetText 21, i, cValTotal(21)
                .SetText 22, i, cValTotal(22)
                .SetText 23, i, cValTotal(23)
                .SetText 27, i, cValTotal(27)
                .SetText 28, i, Int(cValTotal(28)) & ":" & Format(((cValTotal(28) - Int(cValTotal(28))) * 60), "00")
                If nShiftHrs = 0 Then nShiftHrs = 24
                .SetText 29, i, Int((Int(cValTotal(28)) / (nShiftHrs / 60))) & " Days " & Int(Right(Round(cValTotal(28) / 24, 2), 2) / 100 * 24) & " Hrs."
                .SetText 30, i, Int((Int(cValTotal(28)) / (nShiftHrs / 60))) & " Days " & Int(Right(Round(cValTotal(28) / 24, 2), 2) / 100 * 24) & " Hrs."
                .SetText 31, i, rstProductionScheduling.Fields("MAC")
                .SetText 32, i, cValTotal(32)
                .SetText 33, i, cValTotal(33)
    End With
    Call Fomatting_Click
    fpSpread1.MaxRows = IIf(TDBNumber2.Value < 27, i + (27 - TDBNumber2.Value), i + 1)
End Sub
Private Sub Fomatting_Click()
Dim R As Long, C As Long, Cols As Long, Rows As Long
        With fpSpread1
            Cols = .MaxCols
            R = .DataRowCnt
            For C = 1 To Cols
            fpSpread1.Col = C: fpSpread1.Row = R: fpSpread1.FontBold = True: fpSpread1.FontSize = 12.5: fpSpread1.FontUnderline = True: fpSpread1.ForeColor = vbBlue:
        Next
                .SelectBlockOptions = SelectBlockOptionsAll
                .SetActiveCell 3, LR
        End With
End Sub
Private Sub Command2_Click()
  Dim i As Integer, cVal As Variant, R As Long, C As Long
    With fpSpread1
    If Text1.Text = "" Then Exit Sub
            If .DataRowCnt = 0 Then Exit Sub
                For i = 1 To .DataRowCnt 'Unhide All
                .Row = i: .RowHidden = False
            Next
            R = IIf(.ActiveRow + 1 <> LR, .ActiveRow + 1, 1)
            LR = R
If VchType = 1 Then C = Combo3.ListIndex + 1
            For i = R To .DataRowCnt
            If Combo3.ListIndex >= 0 Then .GetText C, i, cVal
                        If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
                        ElseIf Combo3.ListIndex >= 0 Then
                        .SetActiveCell C, i: Exit Sub
                        End If
            Next
    End With
End Sub
Private Sub Mh3dLabel7_Click()
Dim x As Boolean, FileName As String, SheetName As String, LogFileName As String
Dim R As Long, C As Long, Header1 As String
Dim JQty As Variant
    '"Export Data" &
    With fpSpread1
        .ColsFrozen = 0
        fpSpread1.InsertRows 1, 5
              For R = 1 To 5
                    For C = 1 To .MaxCols
                    .Col = C: .Row = R: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .SetText C, R, ""
                Next C
            Next R
'                'Header-1
                .AddCellSpan 1, 1, 34, 1:
                R = 1
                .Col = 1: .Row = R: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .FontBold = True: .FontSize = 20: .ForeColor = RGB(1, 106, 106): .FontUnderline = True:
                .GetText 1, SpreadHeader, JQty
                .SetText 1, R, JQty
                'Header-2
                .AddCellSpan 1, 2, 34, 1:
                R = 2
                .Col = 1: .Row = R: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .FontBold = True: .FontSize = 16: .ForeColor = RGB(1, 106, 106): .FontUnderline = True:
                .GetText 1, SpreadHeader + 1, JQty
                .SetText 1, R, JQty
               .AddCellSpan 1, 3, 34, 1
               .AddCellSpan 1, 4, 34, 1
                'Header-3
                R = 5
                For C = 1 To .MaxCols
                .Col = C: .Row = R: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .FontBold = True: .FontSize = 10: ForeColor = RGB(1, 106, 106): .FontUnderline = True:
                .GetText C, SpreadHeader + 2, JQty
                .SetText C, R, JQty
                Next
    End With
    If Dir(App.Path & "\Export", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Export"
    '
    ' Export Excel file and set result to x
     FileName = App.Path & "\Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"
    SheetName = "Sheet1" '"(" & Me.Caption & ")"
    LogFileName = "Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".txt"
    x = fpSpread1.ExportToExcelEx(FileName, SheetName, LogFileName, ExcelSaveFlagNone)
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
    fpSpread1.DeleteRows 1, 5
    End With
End Sub
Private Sub Option1_Click()
Call cmdRefresh_Click
End Sub
Private Sub Option2_Click()
Call cmdRefresh_Click
End Sub
Private Sub Option3_Click()
Call cmdRefresh_Click
End Sub
Private Sub Preview_Click()
Dim PrintHeader As String
Dim R As Long, C As Long, i As Long
Dim JQty As Variant
Dim answer As Integer
'*********************************************************
With fpSpread1
.ColsFrozen = 0
Command3.Caption = "Freeze Panes"
freezeFlag = False
PrintHeader = Me.Caption
.LockBackColor = vbWhite
'' These are 8.5" X 11" paper dimensions in TWIPS  12240  15840
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
'fpSpread1.PrintShadows = True
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
'.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
   
   'If a cell is currently active, turn off edit mode
    If fpSpread1.EditMode = True Then
        fpSpread1.EditMode = False
        DoEvents
    End If
    Set spreadpreview.frm = Me
    Set pagesetup.frmPageSetup = Me
    Set PrintDlg.frmPrintDlg = Me
    Set headerfooter.frmHeaderFooter = Me
MsgBox "Processing Complete"
          spreadpreview.Show
 End With
End Sub
Private Sub Mh3dLabel6_Click()
Dim i As Long
With fpSpread1
Dim PrintHeader As String
Dim R As Long, C As Long
Dim JQty As Variant
Dim answer As Integer
    .ColsFrozen = 0
    .MaxRows = .MaxRows + 2
    .InsertRows 1, 2
    .SetText .RowHeaderCols - 1, 1, " ": .SetText .RowHeaderCols - 1, 2, " "
    .Col = 9: fpSpread1.ColHidden = True
    For i = .DataRowCnt To .MaxRows
    .SetText .RowHeaderCols - 1, i, " "
    Next i
    For R = 1 To 2
        For C = 1 To .MaxCols
            .Col = C: .Row = R:  .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: '.BackColor = &H8000000F: .FontBold = True: .FontSize = 20: .FontUnderline = True: .ForeColor = vbRed:
            .SetText C, R, ""
        Next C
    Next R
    'Col Span
    .AddCellSpan 1, 1, 30, 1
    .AddCellSpan 1, 2, 30, 1
    
     .Col = 1: .Row = 1: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .FontBold = True: .FontSize = 20: .FontUnderline = True: .ForeColor = RGB(1, 106, 106): .SetText 1, 1, rstCompanyMaster.Fields("PrintName").Value:
     .Col = 1: .Row = 2: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter:  .FontBold = True: .FontSize = 16: .ForeColor = RGB(20, 106, 106): .SetText 1, 2, Me.Caption + "  From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]":
    R = 1
PrintHeader = Me.Caption
.LockBackColor = vbWhite
' These are 8.5" X 11" paper dimensions in TWIPS  12240  15840
Const PaperWidth = 12240
Const PaperHeight = 15840
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
MsgBox "Processing Complete"
fpSpread1.PrintOrientation = PrintOrientationLandscape
fpSpread1.PrintSheet
    'Delete Row
    .DeleteRows 1, 2
    .MaxRows = .MaxRows - 2
    .Col = 9: fpSpread1.ColHidden = False
 End With
End Sub
Private Sub Command3_Click()
With fpSpread1
If freezeFlag = False Then Command3.Caption = "UnFreeze Panes" Else Command3.Caption = "Freeze Panes"
If freezeFlag = False Then
        .ColsFrozen = 5
Else
        .ColsFrozen = 0
        .RowsFrozen = 0
End If
If freezeFlag = True Then freezeFlag = False Else freezeFlag = True
End With
End Sub
Private Sub LoadMasterList()
    If rstMachineMaster.State = adStateOpen Then rstMachineMaster.Close
    rstMachineMaster.Open "SELECT LTRIM(Name) As Col0,Code,Units,MakeReadyTime,Efficiency,MinSizeWidth,MinSizeLength,MaxSizeWidth,MaxSizeLength,Category FROM MachineMaster ", cnDatabase, adOpenKeyset, adLockReadOnly
    rstMachineMaster.ActiveConnection = Nothing
End Sub
Private Function FindDate()
    If timeFlag = False Then
    mSHH = Left(TDBTime1, 2): mSMM = Mid(TDBTime1, 4, 2): 'Machine Start Time
    If TDBTime2 <= TDBTime1 Then mCSHH = 24 + Left(TDBTime2, 2): mCMM = Mid(TDBTime2, 4, 2) Else mCSHH = Left(TDBTime2, 2): mCMM = Mid(TDBTime2, 4, 2) 'Machine Stop Time
    mLsHH = Left(TDBTime4, 2): mLsMM = Mid(TDBTime4, 4, 2) 'Machine Lunch Start Time
    mLeHH = Left(TDBTime5, 2): mLeMM = Mid(TDBTime5, 4, 2) 'Machine Lunch Stop Time
    mLTMM = TDBTime4 - TDBTime5
    'Mid(mLTMM, 4, 2)
    mStart = DateAdd("n", (mSHH * 60 + mSMM), Date)
    'MsgBox mStart, , "Machine Shift Start Time"
    
    mLstart = DateAdd("n", (mLsHH * 60 + mLsMM), Date)
    'MsgBox mLstart, , "Machine Lunch Start Time"
    
    mLstop = DateAdd("n", (mLeHH * 60 + mLeMM), Date)
    'MsgBox mLstop, , "Machine Lunch Stop Time"
    
    mStop = DateAdd("n", (mCSHH * 60 + mCMM), Date)
    'MsgBox mStop, , "Machine Shift Close Time"
    
    mHrs = mStop - mStart
    mHrs = mHrs + mLTMM
    nShiftHrs = Int(Left(mHrs, 2)) * 60
    nShiftHrs = nShiftHrs + Int(Mid(mHrs, 4, 2)) / 60
    nShiftHrs = nShiftHrs + Int(Right(mHrs, 2)) / 60 / 60
    If TDBTime2 <= TDBTime1 Then sStopHrs = TDBTime2 - TDBTime2 Else sStopHrs = (24 * 60) - nShiftHrs - Mid(mLTMM, 4, 2)
    sDate = mStart: eDate = mStop
    timeFlag = True
    End If
    
    jobHrs = ((hh * 60 + mm))
    
    Do While jobHrs > 0
        If rShiftHrs <= 0 Then rShiftHrs = nShiftHrs
'Case-1 (Before Lunch)
        If DateAdd("n", (jobHrs), sDate) <= mLstart Then
            eDate = DateAdd("n", (jobHrs), sDate)
            rShiftHrs = rShiftHrs - jobHrs
            jobHrs = 0
            sDate = eDate
            If eDate = mLstart Then sDate = DateAdd("n", (Mid(mLTMM, 4, 2)), eDate)
'Case-2 (After Lunch)
        ElseIf DateAdd("n", (jobHrs), sDate) > mLstart And jobHrs <= rShiftHrs Then
            If sDate <= mLstart Then eDate = DateAdd("n", (jobHrs + Mid(mLTMM, 4, 2)), sDate) Else eDate = DateAdd("n", (jobHrs), sDate)
            rShiftHrs = rShiftHrs - jobHrs
            If rShiftHrs = 0 Then mStop = DateAdd("h", (24), mStop)
            If rShiftHrs = 0 Then mLstart = DateAdd("h", (24), mLstart)
            jobHrs = 0
            sDate = eDate
'Case-3 (After Shift)
        ElseIf DateAdd("n", (jobHrs + IIf(sDate <= mLstart, Mid(mLTMM, 4, 2), 0)), sDate) >= mStop And jobHrs >= rShiftHrs Then
            If sDate <= mLstart And DateAdd("n", (rShiftHrs), sDate) >= mLstart Then eDate = DateAdd("n", (rShiftHrs + Mid(mLTMM, 4, 2)), sDate) Else eDate = DateAdd("n", (rShiftHrs), sDate)
            mLstart = DateAdd("h", (24), mLstart)
            jobHrs = jobHrs - rShiftHrs
            rShiftHrs = rShiftHrs - rShiftHrs
            sDate = eDate
            If eDate >= mStop Then sDate = DateAdd("n", (sStopHrs), eDate)
            If rShiftHrs = 0 Then mStop = DateAdd("h", (24), mStop)
'Case-4 (Under Shift)
        ElseIf DateAdd("n", (jobHrs + IIf(sDate <= mLstart, Mid(mLTMM, 4, 2), 0)), sDate) <= mStop And jobHrs >= rShiftHrs Then
            If sDate <= mLstart Then eDate = DateAdd("n", (sStopHrs + rShiftHrs + Mid(mLTMM, 4, 2)), sDate) Else eDate = DateAdd("n", (sStopHrs + rShiftHrs), sDate)
            mLstart = DateAdd("h", (24), mLstart)
            jobHrs = jobHrs - rShiftHrs
            rShiftHrs = 0
            If rShiftHrs = 0 Then mStop = DateAdd("h", (24), mStop)
            sDate = eDate
        Else
            MsgBox eDate & " >>>Failed to Calculate Days<<< ", , "Failed to Calculate Days "
            Exit Do
        End If
    Loop
End Function
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 4 Then CloseForm Me: Exit Sub
    OutputTo = Choose(Button.Index, "S", "P", "M")
    PrintProductionSchedule
End Sub
Private Sub PrintProductionSchedule()
    On Error Resume Next
    Dim iCount As Integer
    rptProductionSchedule.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptProductionSchedule.Text9.SetText IIf(Option1.Value, "All Schedule ", IIf(Option2.Value, "Pending Schedule ", "Close Schedule")) & " From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]"
    rstProductionScheduling.ActiveConnection = Nothing
    Screen.MousePointer = vbNormal
    If rstProductionScheduling.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rstProductionScheduling.MoveFirst
    rptProductionSchedule.Database.SetDataSource rstProductionScheduling, 3, 1
    rptProductionSchedule.DiscardSavedData
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptProductionSchedule: FrmReportViewer.Show vbModal
    ElseIf OutputTo = "P" Then
        rptProductionSchedule.PaperSource = crPRBinAuto
        rptProductionSchedule.PrintOut
    Else
        If iCount >= 0 Then
            Dim oOutlookMsg As Outlook.MailItem, FileName As String
            Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
            With oOutlookMsg
                .To = Trim(rstCompanyMaster.Fields("eMail").Value)
                .Subject = IIf(Option1.Value, "All Schedule ", IIf(Option2.Value, "Pending Schedule ", "Close Schedule")) & "Printing Floor"
                .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith " & IIf(Option1.Value, "All Schedule ", IIf(Option2.Value, "Pending Schedule ", "Close Schedule")) & "Printing Floor from " + Format(GetDate(MhDateInput1.Text), "dd-MMM-yyyy") + " to " + Format(GetDate(MhDateInput2.Text), "dd-MMM-yyyy") & " for doing the needful at your end.<Br><b>Kindly do acknowledge the receipt of the mail</b>.<Br><Br>Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
                rptProductionSchedule.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptProductionSchedule.ExportOptions.DestinationType = crEDTDiskFile
                FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
                rptProductionSchedule.ExportOptions.DiskFileName = FileName
                rptProductionSchedule.Export False
                .Attachments.Add (FileName)
                .Importance = olImportanceHigh
                .ReadReceiptRequested = True
                If CheckEmpty(.To, False) Then .Display Else .Send
            End With
            Set oOutlookMsg = Nothing
        End If
    End If
    Set rptProductionSchedule = Nothing
    On Error GoTo 0
End Sub
