VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPaperSearchList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Papers..."
   ClientHeight    =   8145
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   16335
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   16335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFilter 
      Height          =   320
      Left            =   15920
      Picture         =   "PaperSearchList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Filter"
      Top             =   7375
      Width           =   375
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   7035
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   345
      Width           =   16250
      _Version        =   524288
      _ExtentX        =   28663
      _ExtentY        =   12409
      _StockProps     =   64
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
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
      GridColor       =   4227327
      MaxCols         =   12
      MaxRows         =   1000
      ScrollBars      =   2
      SpreadDesigner  =   "PaperSearchList.frx":0532
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
      Left            =   530
      TabIndex        =   0
      ToolTipText     =   "Find"
      Top             =   7760
      Width           =   13380
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Proceed"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PaperSearchList.frx":104C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PaperSearchList.frx":115E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   0
      Left            =   45
      TabIndex        =   11
      Top             =   7760
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   9164542
      BackColor       =   0
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
      Picture         =   "PaperSearchList.frx":1270
      Picture         =   "PaperSearchList.frx":128C
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   2
      Left            =   13890
      TabIndex        =   12
      Top             =   7755
      Width           =   2405
      _Version        =   65536
      _ExtentX        =   4242
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
      Caption         =   " Enter->Proceed  Esc->Cancel"
      Alignment       =   0
      FillColor       =   8421504
      TextColor       =   16777215
      Picture         =   "PaperSearchList.frx":12A8
      Picture         =   "PaperSearchList.frx":12C4
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   4
      Left            =   10180
      TabIndex        =   13
      Top             =   7365
      Width           =   480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   9164542
      BackColor       =   0
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
      Caption         =   " GSM"
      Alignment       =   0
      FillColor       =   9164542
      Picture         =   "PaperSearchList.frx":12E0
      Picture         =   "PaperSearchList.frx":12FC
   End
   Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
      Height          =   330
      Left            =   10645
      TabIndex        =   3
      ToolTipText     =   "GSM"
      Top             =   7365
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   582
      Calculator      =   "PaperSearchList.frx":1318
      Caption         =   "PaperSearchList.frx":1338
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "PaperSearchList.frx":13A4
      Keys            =   "PaperSearchList.frx":13C2
      Spin            =   "PaperSearchList.frx":140C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   9999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ""
      ShowContextMenu =   1
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   5
      Left            =   11125
      TabIndex        =   14
      Top             =   7365
      Width           =   195
      _Version        =   65536
      _ExtentX        =   344
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   9164542
      BackColor       =   0
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
      Caption         =   "±"
      FillColor       =   9164542
      Picture         =   "PaperSearchList.frx":1434
      Picture         =   "PaperSearchList.frx":1450
   End
   Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
      Height          =   330
      Left            =   11305
      TabIndex        =   4
      ToolTipText     =   "GSM Tolerance"
      Top             =   7365
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   582
      Calculator      =   "PaperSearchList.frx":146C
      Caption         =   "PaperSearchList.frx":148C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "PaperSearchList.frx":14F8
      Keys            =   "PaperSearchList.frx":1516
      Spin            =   "PaperSearchList.frx":1560
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   9999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ""
      ShowContextMenu =   1
      ValueVT         =   1325268997
      Value           =   1
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   1
      Left            =   11785
      TabIndex        =   15
      Top             =   7365
      Width           =   630
      _Version        =   65536
      _ExtentX        =   1111
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   9164542
      BackColor       =   0
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
      Caption         =   " Width"
      Alignment       =   0
      FillColor       =   9164542
      Picture         =   "PaperSearchList.frx":1588
      Picture         =   "PaperSearchList.frx":15A4
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   3
      Left            =   13030
      TabIndex        =   16
      Top             =   7365
      Width           =   195
      _Version        =   65536
      _ExtentX        =   344
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   9164542
      BackColor       =   0
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
      Caption         =   "±"
      FillColor       =   9164542
      Picture         =   "PaperSearchList.frx":15C0
      Picture         =   "PaperSearchList.frx":15DC
   End
   Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
      Height          =   330
      Left            =   12400
      TabIndex        =   5
      ToolTipText     =   "Width"
      Top             =   7365
      Width           =   645
      _Version        =   65536
      _ExtentX        =   1147
      _ExtentY        =   582
      Calculator      =   "PaperSearchList.frx":15F8
      Caption         =   "PaperSearchList.frx":1618
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "PaperSearchList.frx":1684
      Keys            =   "PaperSearchList.frx":16A2
      Spin            =   "PaperSearchList.frx":16EC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###########0.00"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###########0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999.99
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ""
      ShowContextMenu =   1
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
      Height          =   330
      Left            =   13210
      TabIndex        =   6
      ToolTipText     =   "Width Tolerance"
      Top             =   7365
      Width           =   645
      _Version        =   65536
      _ExtentX        =   1147
      _ExtentY        =   582
      Calculator      =   "PaperSearchList.frx":1714
      Caption         =   "PaperSearchList.frx":1734
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "PaperSearchList.frx":17A0
      Keys            =   "PaperSearchList.frx":17BE
      Spin            =   "PaperSearchList.frx":1808
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###########0.00"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###########0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999.99
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ""
      ShowContextMenu =   1
      ValueVT         =   5
      Value           =   1
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   6
      Left            =   13840
      TabIndex        =   17
      Top             =   7365
      Width           =   630
      _Version        =   65536
      _ExtentX        =   1111
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   9164542
      BackColor       =   0
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
      Caption         =   " Length"
      Alignment       =   0
      FillColor       =   9164542
      Picture         =   "PaperSearchList.frx":1830
      Picture         =   "PaperSearchList.frx":184C
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   7
      Left            =   15085
      TabIndex        =   18
      Top             =   7365
      Width           =   195
      _Version        =   65536
      _ExtentX        =   344
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   9164542
      BackColor       =   0
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
      Caption         =   "±"
      FillColor       =   9164542
      Picture         =   "PaperSearchList.frx":1868
      Picture         =   "PaperSearchList.frx":1884
   End
   Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
      Height          =   330
      Left            =   14455
      TabIndex        =   7
      ToolTipText     =   "Length"
      Top             =   7365
      Width           =   645
      _Version        =   65536
      _ExtentX        =   1138
      _ExtentY        =   582
      Calculator      =   "PaperSearchList.frx":18A0
      Caption         =   "PaperSearchList.frx":18C0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "PaperSearchList.frx":192C
      Keys            =   "PaperSearchList.frx":194A
      Spin            =   "PaperSearchList.frx":1994
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###########0.00"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###########0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999.99
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ""
      ShowContextMenu =   1
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
      Height          =   330
      Left            =   15275
      TabIndex        =   8
      ToolTipText     =   "Length Tolerance"
      Top             =   7365
      Width           =   645
      _Version        =   65536
      _ExtentX        =   1138
      _ExtentY        =   582
      Calculator      =   "PaperSearchList.frx":19BC
      Caption         =   "PaperSearchList.frx":19DC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "PaperSearchList.frx":1A48
      Keys            =   "PaperSearchList.frx":1A66
      Spin            =   "PaperSearchList.frx":1AB0
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###########0.00"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###########0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999.99
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ""
      ShowContextMenu =   1
      ValueVT         =   1325268997
      Value           =   1
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   8
      Left            =   7990
      TabIndex        =   20
      Top             =   7365
      Width           =   720
      _Version        =   65536
      _ExtentX        =   1270
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   9164542
      BackColor       =   0
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
      Caption         =   " Quality"
      Alignment       =   0
      FillColor       =   9164542
      Picture         =   "PaperSearchList.frx":1AD8
      Picture         =   "PaperSearchList.frx":1AF4
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   9
      Left            =   5920
      TabIndex        =   21
      Top             =   7365
      Width           =   600
      _Version        =   65536
      _ExtentX        =   1058
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   9164542
      BackColor       =   0
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
      Caption         =   " Grade"
      Alignment       =   0
      FillColor       =   9164542
      Picture         =   "PaperSearchList.frx":1B10
      Picture         =   "PaperSearchList.frx":1B2C
   End
   Begin MSForms.ComboBox cmbGrade 
      Height          =   330
      Left            =   6505
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7365
      Width           =   1500
      VariousPropertyBits=   545282075
      BackColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "2646;582"
      MatchEntry      =   0
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Calibri"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbQuality 
      Height          =   330
      Left            =   8695
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7365
      Width           =   1500
      VariousPropertyBits=   545282075
      BackColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "2646;582"
      MatchEntry      =   0
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Calibri"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FrmPaperSearchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstPaperSearchList As New ADODB.Recordset, PaperCode As String, PaperName As String
Dim PrevStr As String
Private Sub Form_Load()
    Dim rstPaperQualityList As New ADODB.Recordset, i As Integer
    cmbGrade.AddItem "A", 0
    cmbGrade.AddItem "B", 1
    cmbGrade.AddItem "C", 2
    cmbGrade.AddItem "D", 3
    With rstPaperQualityList
        .Open "SELECT Name FROM GeneralMaster WHERE Type=16 ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
        Do Until .EOF
            cmbQuality.AddItem .Fields("Name").Value, i: i = i + 1
            .MoveNext
        Loop
    End With
    cmbQuality.ListIndex = -1
    cmbGrade.ListIndex = -1
    CloseRecordset rstPaperQualityList
    rstPaperSearchList.Filter = adFilterNone
    Set fpSpread1.DataSource = rstPaperSearchList
End Sub
Private Sub Form_Activate()
    cmdFilter_Click
    If Not CheckEmpty(Text1.Text, False) And rstPaperSearchList.RecordCount > 0 Then rstPaperSearchList.MoveFirst: rstPaperSearchList.Find "[Col0]='" & Text1.Text & "'"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0: Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        KeyCode = 0: Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub
Private Sub Text1_Change()
    With rstPaperSearchList
        If .RecordCount = 0 Then Exit Sub
        If Len(Text1.Text) > 0 Then
            .MoveFirst
            .Filter = "[Col0] Like '%" & Text1.Text & "%'"
            If Not .EOF Then
                PrevStr = Text1.Text
            Else
                .Filter = adFilterNone: .MoveFirst
                Text1.Text = PrevStr
                Sendkeys "{End}"
            End If
        Else
            .Filter = adFilterNone: .MoveFirst
            PrevStr = ""
        End If
    End With
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    With rstPaperSearchList
        If .RecordCount = 0 Then Exit Sub
        If Shift = 0 And KeyCode = vbKeyBack Then
            .MoveFirst
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyUp Then
            .MovePrevious
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyDown Then
            .MoveNext
            If .EOF Then .MoveLast
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
            .Move -20
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
            .Move 20
            If .EOF Then .MoveLast
            KeyProcessed = True
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyHome Then
            .MoveFirst
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyEnd Then
            .MoveLast
            If .EOF Then .MoveLast
            KeyProcessed = True
        End If
    End With
    If KeyProcessed Then KeyCode = 0
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim Paper As Variant
    With fpSpread1
        .GetText 7, .ActiveRow, Paper
        PaperCode = Paper
        .GetText 1, .ActiveRow, Paper
        PaperName = Paper
    End With
    If Button.Index = 2 Then PaperCode = ""
    Set rstPaperSearchList = Nothing
    Me.Hide
End Sub
Private Sub cmdFilter_Click()
    If MhRealInput3.Value = 0 Or MhRealInput5.Value = 0 Then DisplayError ("Both Width & Length are mandatory"): Exit Sub
    Dim x As String, GSM As String, Quality As String, Grade As String
    GSM = IIf(MhRealInput1.Value = 0, "[GSM]<>0", "[GSM]>=" + Trim(MhRealInput1.Value - MhRealInput2.Value) + " AND [GSM]<=" + Trim(MhRealInput1.Value + MhRealInput2.Value))
    Quality = IIf(cmbQuality.ListIndex = -1, "[Quality]<>''", "[Quality]='" + Trim(cmbQuality.Text) & "'")
    Grade = IIf(cmbGrade.ListIndex = -1, "[Grade]<>''", "[Grade]='" + Trim(cmbGrade.Text) & "'")
    x = x + "((" + GSM + ") AND (" + Quality + ") AND (" + Grade + ") AND (([inWidth]>=" + Trim(MhRealInput3.Value - MhRealInput4.Value) + " AND [inWidth]<=" + Trim(MhRealInput3.Value + MhRealInput4.Value) + ") AND ([inLength]>=" + Trim(MhRealInput5.Value - MhRealInput6.Value) + " AND [inLength]<=" + Trim(MhRealInput5.Value + MhRealInput6.Value) + "))) OR ((" + GSM + ") AND (" + Quality + ") AND (" + Grade + ") AND (([inWidth]>=" + Trim(MhRealInput5.Value - MhRealInput4.Value) + " AND [inWidth]<=" + Trim(MhRealInput5.Value + MhRealInput4.Value) + ") AND ([inLength]>=" + Trim(MhRealInput3.Value * 2 - MhRealInput6.Value) + " AND [inLength]<=" + Trim(MhRealInput3.Value * 2 + MhRealInput6.Value) + ")))"
    With rstPaperSearchList
        .Filter = adFilterNone
        .Filter = x
        If .RecordCount = 0 Then .Filter = adFilterNone
    End With
End Sub
Private Sub cmbGrade_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmbGrade.ListIndex = -1
End Sub
Private Sub cmbQuality_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmbQuality.ListIndex = -1
End Sub
