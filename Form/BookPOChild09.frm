VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmBookPOChild09 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Combo Items Sheet Printing Order Details"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16260
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
   ScaleHeight     =   7845
   ScaleWidth      =   16260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProceed 
      BackColor       =   &H008BD6FE&
      Height          =   375
      Left            =   15293
      Picture         =   "BookPOChild09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008BD6FE&
      Height          =   375
      Left            =   15293
      Picture         =   "BookPOChild09.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   98
      TabStop         =   0   'False
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   7635
      Left            =   120
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   120
      Width           =   14595
      _Version        =   65536
      _ExtentX        =   25744
      _ExtentY        =   13467
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
      Picture         =   "BookPOChild09.frx":0204
      Begin VB.TextBox Text7 
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
         MaxLength       =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   645
         Width           =   4455
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
         TabIndex        =   5
         Top             =   960
         Width           =   4455
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput23 
         Height          =   330
         Left            =   7680
         TabIndex        =   35
         ToolTipText     =   "Minimum Sheets"
         Top             =   5535
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0220
         Caption         =   "BookPOChild09.frx":0240
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":02AC
         Keys            =   "BookPOChild09.frx":02CA
         Spin            =   "BookPOChild09.frx":0314
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#####0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Left            =   9120
         TabIndex        =   96
         Top             =   5535
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   582
         _StockProps     =   77
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
         Caption         =   " Final Wastage"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":033C
         Picture         =   "BookPOChild09.frx":0358
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput29 
         Height          =   330
         Left            =   8385
         TabIndex        =   36
         ToolTipText     =   "Maximum Sheets"
         Top             =   5535
         Width           =   750
         _Version        =   65536
         _ExtentX        =   1323
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0374
         Caption         =   "BookPOChild09.frx":0394
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0400
         Keys            =   "BookPOChild09.frx":041E
         Spin            =   "BookPOChild09.frx":0468
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#####0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   9120
         TabIndex        =   82
         Top             =   5850
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
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
         Caption         =   " GST Paper"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":0490
         Picture         =   "BookPOChild09.frx":04AC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel39 
         Height          =   330
         Left            =   12000
         TabIndex        =   90
         Top             =   5220
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
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
         Caption         =   " Consumption"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":04C8
         Picture         =   "BookPOChild09.frx":04E4
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput33 
         Height          =   330
         Left            =   7680
         TabIndex        =   40
         Top             =   5850
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0500
         Caption         =   "BookPOChild09.frx":0520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":058C
         Keys            =   "BookPOChild09.frx":05AA
         Spin            =   "BookPOChild09.frx":05F4
         AlignHorizontal =   1
         AlignVertical   =   2
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
         MinValue        =   -9999999999.99
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
      Begin VB.TextBox Text10 
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
         TabIndex        =   47
         Top             =   6705
         Width           =   1815
      End
      Begin VB.TextBox Text9 
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
         TabIndex        =   15
         Top             =   3905
         Width           =   12810
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
         TabIndex        =   50
         Top             =   7215
         Width           =   12810
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   105
         Width           =   1575
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
         TabIndex        =   44
         Top             =   6390
         Width           =   1815
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   30
         Top             =   5225
         Width           =   4450
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   6120
         TabIndex        =   53
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
         Picture         =   "BookPOChild09.frx":061C
         Picture         =   "BookPOChild09.frx":0638
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   54
         Top             =   960
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
         Caption         =   " Printing Size"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":0654
         Picture         =   "BookPOChild09.frx":0670
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   55
         Top             =   4725
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
         Caption         =   " Total Plates"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":068C
         Picture         =   "BookPOChild09.frx":06A8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   3480
         TabIndex        =   56
         Top             =   4725
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
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
         Caption         =   " Plate Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":06C4
         Picture         =   "BookPOChild09.frx":06E0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   3480
         TabIndex        =   57
         Top             =   4410
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
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
         Caption         =   " Print Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":06FC
         Picture         =   "BookPOChild09.frx":0718
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   6120
         TabIndex        =   58
         Top             =   4410
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
         Caption         =   " Adjustment Ptg."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":0734
         Picture         =   "BookPOChild09.frx":0750
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   6120
         TabIndex        =   59
         Top             =   3585
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
         Caption         =   " Plate Type"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":076C
         Picture         =   "BookPOChild09.frx":0788
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   60
         Top             =   4410
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
         Caption         =   " Total Forms"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":07A4
         Picture         =   "BookPOChild09.frx":07C0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Left            =   6120
         TabIndex        =   61
         Top             =   5220
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
         Caption         =   " Forms/Sheet"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":07DC
         Picture         =   "BookPOChild09.frx":07F8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
         Height          =   330
         Left            =   12000
         TabIndex        =   62
         Top             =   5535
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   582
         _StockProps     =   77
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
         Caption         =   " Total Consumption"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":0814
         Picture         =   "BookPOChild09.frx":0830
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   63
         Top             =   6390
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
         Caption         =   " Party Bill No."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":084C
         Picture         =   "BookPOChild09.frx":0868
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   6120
         TabIndex        =   64
         Top             =   6390
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
         Picture         =   "BookPOChild09.frx":0884
         Picture         =   "BookPOChild09.frx":08A0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   3480
         TabIndex        =   65
         Top             =   6390
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
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
         Picture         =   "BookPOChild09.frx":08BC
         Picture         =   "BookPOChild09.frx":08D8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   11355
         TabIndex        =   66
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
         Picture         =   "BookPOChild09.frx":08F4
         Picture         =   "BookPOChild09.frx":0910
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   67
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
         Picture         =   "BookPOChild09.frx":092C
         Picture         =   "BookPOChild09.frx":0948
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   68
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
         Picture         =   "BookPOChild09.frx":0964
         Picture         =   "BookPOChild09.frx":0980
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   69
         Top             =   7215
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
         Picture         =   "BookPOChild09.frx":099C
         Picture         =   "BookPOChild09.frx":09B8
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   1680
         TabIndex        =   16
         ToolTipText     =   "Front"
         Top             =   4410
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":09D4
         Caption         =   "BookPOChild09.frx":09F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0A60
         Keys            =   "BookPOChild09.frx":0A7E
         Spin            =   "BookPOChild09.frx":0AC8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   255
         Format          =   "###########0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   4440
         TabIndex        =   18
         ToolTipText     =   "Front"
         Top             =   4410
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0AF0
         Caption         =   "BookPOChild09.frx":0B10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0B7C
         Keys            =   "BookPOChild09.frx":0B9A
         Spin            =   "BookPOChild09.frx":0BE4
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   4440
         TabIndex        =   25
         Top             =   4725
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0C0C
         Caption         =   "BookPOChild09.frx":0C2C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0C98
         Keys            =   "BookPOChild09.frx":0CB6
         Spin            =   "BookPOChild09.frx":0D00
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   330
         Left            =   7680
         TabIndex        =   20
         Top             =   4410
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0D28
         Caption         =   "BookPOChild09.frx":0D48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0DB4
         Keys            =   "BookPOChild09.frx":0DD2
         Spin            =   "BookPOChild09.frx":0E1C
         AlignHorizontal =   1
         AlignVertical   =   2
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
         MinValue        =   -9999999999.99
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
         Height          =   330
         Left            =   7680
         TabIndex        =   46
         Top             =   6390
         Width           =   6810
         _Version        =   65536
         _ExtentX        =   12012
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0E44
         Caption         =   "BookPOChild09.frx":0E64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0ED0
         Keys            =   "BookPOChild09.frx":0EEE
         Spin            =   "BookPOChild09.frx":0F38
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   1680
         TabIndex        =   24
         Top             =   4725
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0F60
         Caption         =   "BookPOChild09.frx":0F80
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0FEC
         Keys            =   "BookPOChild09.frx":100A
         Spin            =   "BookPOChild09.frx":1054
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   255
         Format          =   "###########0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   4440
         TabIndex        =   34
         ToolTipText     =   "%"
         Top             =   5535
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":107C
         Caption         =   "BookPOChild09.frx":109C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1108
         Keys            =   "BookPOChild09.frx":1126
         Spin            =   "BookPOChild09.frx":1170
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99.99
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
         Value           =   4
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   11040
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4410
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":1198
         Caption         =   "BookPOChild09.frx":11B8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1224
         Keys            =   "BookPOChild09.frx":1242
         Spin            =   "BookPOChild09.frx":128C
         AlignHorizontal =   1
         AlignVertical   =   2
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
         Height          =   330
         Left            =   10200
         TabIndex        =   21
         Top             =   4410
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":12B4
         Caption         =   "BookPOChild09.frx":12D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1340
         Keys            =   "BookPOChild09.frx":135E
         Spin            =   "BookPOChild09.frx":13A8
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99.99
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   13395
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   5535
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":13D0
         Caption         =   "BookPOChild09.frx":13F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":145C
         Keys            =   "BookPOChild09.frx":147A
         Spin            =   "BookPOChild09.frx":14C4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "########0.000"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   255
         Format          =   "########0.000"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.999
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
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   7650
         TabIndex        =   1
         Top             =   105
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild09.frx":14EC
         Caption         =   "BookPOChild09.frx":1604
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1670
         Keys            =   "BookPOChild09.frx":168E
         Spin            =   "BookPOChild09.frx":16EC
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
         Left            =   12915
         TabIndex        =   2
         Top             =   105
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild09.frx":1714
         Caption         =   "BookPOChild09.frx":182C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1898
         Keys            =   "BookPOChild09.frx":18B6
         Spin            =   "BookPOChild09.frx":1914
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
         Left            =   4440
         TabIndex        =   45
         Top             =   6390
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calendar        =   "BookPOChild09.frx":193C
         Caption         =   "BookPOChild09.frx":1A54
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1AC0
         Keys            =   "BookPOChild09.frx":1ADE
         Spin            =   "BookPOChild09.frx":1B3C
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel29 
         Height          =   330
         Left            =   120
         TabIndex        =   70
         Top             =   3580
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
         Caption         =   " Plate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":1B64
         Picture         =   "BookPOChild09.frx":1B80
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel30 
         Height          =   330
         Left            =   6120
         TabIndex        =   71
         Top             =   960
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
         Caption         =   " Imposition"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":1B9C
         Picture         =   "BookPOChild09.frx":1BB8
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
         Height          =   330
         Left            =   11040
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   4725
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":1BD4
         Caption         =   "BookPOChild09.frx":1BF4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1C60
         Keys            =   "BookPOChild09.frx":1C7E
         Spin            =   "BookPOChild09.frx":1CC8
         AlignHorizontal =   1
         AlignVertical   =   2
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
         Height          =   330
         Left            =   10200
         TabIndex        =   27
         Top             =   4725
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":1CF0
         Caption         =   "BookPOChild09.frx":1D10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1D7C
         Keys            =   "BookPOChild09.frx":1D9A
         Spin            =   "BookPOChild09.frx":1DE4
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99.99
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel33 
         Height          =   330
         Left            =   120
         TabIndex        =   72
         Top             =   3905
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
         Caption         =   " Plate Party"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":1E0C
         Picture         =   "BookPOChild09.frx":1E28
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel34 
         Height          =   330
         Left            =   120
         TabIndex        =   73
         Top             =   6705
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
         Caption         =   " P/Party Bill No."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":1E44
         Picture         =   "BookPOChild09.frx":1E60
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel35 
         Height          =   330
         Left            =   6120
         TabIndex        =   74
         Top             =   6705
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
         Picture         =   "BookPOChild09.frx":1E7C
         Picture         =   "BookPOChild09.frx":1E98
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel36 
         Height          =   330
         Left            =   3480
         TabIndex        =   75
         Top             =   6705
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
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
         Picture         =   "BookPOChild09.frx":1EB4
         Picture         =   "BookPOChild09.frx":1ED0
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
         Height          =   330
         Left            =   7680
         TabIndex        =   49
         Top             =   6705
         Width           =   6810
         _Version        =   65536
         _ExtentX        =   12012
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":1EEC
         Caption         =   "BookPOChild09.frx":1F0C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1F78
         Keys            =   "BookPOChild09.frx":1F96
         Spin            =   "BookPOChild09.frx":1FE0
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
      Begin TDBDate6Ctl.TDBDate MhDateInput4 
         Height          =   330
         Left            =   4440
         TabIndex        =   48
         Top             =   6705
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calendar        =   "BookPOChild09.frx":2008
         Caption         =   "BookPOChild09.frx":2120
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":218C
         Keys            =   "BookPOChild09.frx":21AA
         Spin            =   "BookPOChild09.frx":2208
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
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   1635
         Left            =   120
         TabIndex        =   7
         Top             =   1485
         Width           =   14370
         _Version        =   524288
         _ExtentX        =   25347
         _ExtentY        =   2884
         _StockProps     =   64
         ButtonDrawMode  =   8
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
         MaxCols         =   8
         MaxRows         =   1000
         ScrollBars      =   2
         SpreadDesigner  =   "BookPOChild09.frx":2230
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel31 
         Height          =   330
         Left            =   12000
         TabIndex        =   76
         Top             =   4410
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
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
         Caption         =   " Print Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":2B02
         Picture         =   "BookPOChild09.frx":2B1E
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
         Height          =   330
         Left            =   13395
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4410
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":2B3A
         Caption         =   "BookPOChild09.frx":2B5A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":2BC6
         Keys            =   "BookPOChild09.frx":2BE4
         Spin            =   "BookPOChild09.frx":2C2E
         AlignHorizontal =   1
         AlignVertical   =   2
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel37 
         Height          =   330
         Left            =   12000
         TabIndex        =   77
         Top             =   4725
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
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
         Caption         =   " Plate Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":2C56
         Picture         =   "BookPOChild09.frx":2C72
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
         Height          =   330
         Left            =   13395
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4725
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":2C8E
         Caption         =   "BookPOChild09.frx":2CAE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":2D1A
         Keys            =   "BookPOChild09.frx":2D38
         Spin            =   "BookPOChild09.frx":2D82
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
         Height          =   330
         Left            =   6120
         TabIndex        =   78
         Top             =   4725
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
         Caption         =   " Adjustment Plate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":2DAA
         Picture         =   "BookPOChild09.frx":2DC6
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput28 
         Height          =   330
         Left            =   13395
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   5220
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":2DE2
         Caption         =   "BookPOChild09.frx":2E02
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":2E6E
         Keys            =   "BookPOChild09.frx":2E8C
         Spin            =   "BookPOChild09.frx":2ED6
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "########0.000"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   255
         Format          =   "########0.000"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.999
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput30 
         Height          =   330
         Left            =   10200
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Final Wastage"
         Top             =   5535
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":2EFE
         Caption         =   "BookPOChild09.frx":2F1E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":2F8A
         Keys            =   "BookPOChild09.frx":2FA8
         Spin            =   "BookPOChild09.frx":2FF2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "########0.000"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   255
         Format          =   "########0.000"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.999
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
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   79
         Top             =   2280
         Width           =   11370
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   2580
         TabIndex        =   17
         ToolTipText     =   "Back"
         Top             =   4410
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":301A
         Caption         =   "BookPOChild09.frx":303A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":30A6
         Keys            =   "BookPOChild09.frx":30C4
         Spin            =   "BookPOChild09.frx":310E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   255
         Format          =   "###########0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
         Height          =   330
         Left            =   5280
         TabIndex        =   19
         ToolTipText     =   "Back"
         Top             =   4410
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":3136
         Caption         =   "BookPOChild09.frx":3156
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":31C2
         Keys            =   "BookPOChild09.frx":31E0
         Spin            =   "BookPOChild09.frx":322A
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   80
         Top             =   5850
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
         Caption         =   " Paper Rate/Kg."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":3252
         Picture         =   "BookPOChild09.frx":326E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   12000
         TabIndex        =   81
         Top             =   5850
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
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
         Caption         =   " Paper Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":328A
         Picture         =   "BookPOChild09.frx":32A6
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput31 
         Height          =   330
         Left            =   1680
         TabIndex        =   39
         ToolTipText     =   "Front"
         Top             =   5850
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":32C2
         Caption         =   "BookPOChild09.frx":32E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":334E
         Keys            =   "BookPOChild09.frx":336C
         Spin            =   "BookPOChild09.frx":33B6
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput32 
         Height          =   330
         Left            =   13395
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   5850
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":33DE
         Caption         =   "BookPOChild09.frx":33FE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":346A
         Keys            =   "BookPOChild09.frx":3488
         Spin            =   "BookPOChild09.frx":34D2
         AlignHorizontal =   1
         AlignVertical   =   2
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
         Height          =   330
         Left            =   7680
         TabIndex        =   31
         Top             =   5220
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":34FA
         Caption         =   "BookPOChild09.frx":351A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":3586
         Keys            =   "BookPOChild09.frx":35A4
         Spin            =   "BookPOChild09.frx":35EE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   261292037
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput35 
         Height          =   330
         Left            =   11040
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   5850
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":3616
         Caption         =   "BookPOChild09.frx":3636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":36A2
         Keys            =   "BookPOChild09.frx":36C0
         Spin            =   "BookPOChild09.frx":370A
         AlignHorizontal =   1
         AlignVertical   =   2
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput34 
         Height          =   330
         Left            =   10200
         TabIndex        =   41
         Top             =   5850
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":3732
         Caption         =   "BookPOChild09.frx":3752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":37BE
         Keys            =   "BookPOChild09.frx":37DC
         Spin            =   "BookPOChild09.frx":3826
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99.99
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   6120
         TabIndex        =   83
         Top             =   5850
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
         Caption         =   " Adjustment Paper"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":384E
         Picture         =   "BookPOChild09.frx":386A
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput38 
         Height          =   330
         Left            =   9330
         TabIndex        =   86
         TabStop         =   0   'False
         ToolTipText     =   "Paper Amount BT"
         Top             =   7215
         Visible         =   0   'False
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":3886
         Caption         =   "BookPOChild09.frx":38A6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":3912
         Keys            =   "BookPOChild09.frx":3930
         Spin            =   "BookPOChild09.frx":397A
         AlignHorizontal =   1
         AlignVertical   =   2
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
         ValueVT         =   147062789
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput37 
         Height          =   330
         Left            =   10725
         TabIndex        =   85
         TabStop         =   0   'False
         ToolTipText     =   "Plate Amount BT"
         Top             =   7215
         Visible         =   0   'False
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":39A2
         Caption         =   "BookPOChild09.frx":39C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":3A2E
         Keys            =   "BookPOChild09.frx":3A4C
         Spin            =   "BookPOChild09.frx":3A96
         AlignHorizontal =   1
         AlignVertical   =   2
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput36 
         Height          =   330
         Left            =   12120
         TabIndex        =   84
         TabStop         =   0   'False
         ToolTipText     =   "Printing Amount BT"
         Top             =   7215
         Visible         =   0   'False
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":3ABE
         Caption         =   "BookPOChild09.frx":3ADE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":3B4A
         Keys            =   "BookPOChild09.frx":3B68
         Spin            =   "BookPOChild09.frx":3BB2
         AlignHorizontal =   1
         AlignVertical   =   2
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
         ValueVT         =   147062789
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
         Height          =   330
         Left            =   9120
         TabIndex        =   87
         Top             =   4410
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
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
         Caption         =   " GST Ptg."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":3BDA
         Picture         =   "BookPOChild09.frx":3BF6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         DataField       =   "C"
         Height          =   330
         Index           =   0
         Left            =   9120
         TabIndex        =   88
         Top             =   4725
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
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
         Caption         =   " GST Plate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":3C12
         Picture         =   "BookPOChild09.frx":3C2E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   300
         Left            =   120
         TabIndex        =   89
         Top             =   3105
         Width           =   14370
         _Version        =   65536
         _ExtentX        =   25347
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
         Picture         =   "BookPOChild09.frx":3C4A
         Picture         =   "BookPOChild09.frx":3C66
         Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
            Height          =   300
            Left            =   7995
            TabIndex        =   8
            Top             =   0
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1759
            _ExtentY        =   529
            Calculator      =   "BookPOChild09.frx":3C82
            Caption         =   "BookPOChild09.frx":3CA2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild09.frx":3D0E
            Keys            =   "BookPOChild09.frx":3D2C
            Spin            =   "BookPOChild09.frx":3D76
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###########0"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   255
            Format          =   "###########0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999
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
         Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
            Height          =   300
            Left            =   9945
            TabIndex        =   9
            Top             =   0
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   529
            Calculator      =   "BookPOChild09.frx":3D9E
            Caption         =   "BookPOChild09.frx":3DBE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild09.frx":3E2A
            Keys            =   "BookPOChild09.frx":3E48
            Spin            =   "BookPOChild09.frx":3E92
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###########0"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   255
            Format          =   "###########0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999
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
         Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
            Height          =   300
            Left            =   11160
            TabIndex        =   10
            Top             =   0
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1759
            _ExtentY        =   529
            Calculator      =   "BookPOChild09.frx":3EBA
            Caption         =   "BookPOChild09.frx":3EDA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild09.frx":3F46
            Keys            =   "BookPOChild09.frx":3F64
            Spin            =   "BookPOChild09.frx":3FAE
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###########0"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   255
            Format          =   "###########0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999
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
         Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
            Height          =   300
            Left            =   12105
            TabIndex        =   11
            Top             =   0
            Width           =   1020
            _Version        =   65536
            _ExtentX        =   1799
            _ExtentY        =   529
            Calculator      =   "BookPOChild09.frx":3FD6
            Caption         =   "BookPOChild09.frx":3FF6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild09.frx":4062
            Keys            =   "BookPOChild09.frx":4080
            Spin            =   "BookPOChild09.frx":40CA
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###########0"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   255
            Format          =   "###########0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999
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
         Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
            Height          =   300
            Left            =   13095
            TabIndex        =   12
            Top             =   0
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   529
            Calculator      =   "BookPOChild09.frx":40F2
            Caption         =   "BookPOChild09.frx":4112
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild09.frx":417E
            Keys            =   "BookPOChild09.frx":419C
            Spin            =   "BookPOChild09.frx":41E6
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###########0"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   255
            Format          =   "###########0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999
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
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
         Height          =   330
         Left            =   7680
         TabIndex        =   26
         Top             =   4725
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":420E
         Caption         =   "BookPOChild09.frx":422E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":429A
         Keys            =   "BookPOChild09.frx":42B8
         Spin            =   "BookPOChild09.frx":4302
         AlignHorizontal =   1
         AlignVertical   =   2
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
         MinValue        =   -9999999999.99
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Left            =   3480
         TabIndex        =   91
         Top             =   5535
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
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
         Caption         =   " Wastage %"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":432A
         Picture         =   "BookPOChild09.frx":4346
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   120
         TabIndex        =   92
         Top             =   5225
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
         Caption         =   " Paper Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":4362
         Picture         =   "BookPOChild09.frx":437E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel50 
         Height          =   330
         Left            =   120
         TabIndex        =   93
         Top             =   5535
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
         Caption         =   " Paper By Party"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":439A
         Picture         =   "BookPOChild09.frx":43B6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   1680
         TabIndex        =   94
         Top             =   5535
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
         Caption         =   ""
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":43D2
         Picture         =   "BookPOChild09.frx":43EE
         Begin VB.CheckBox chkPaper 
            Caption         =   "Check2"
            Height          =   210
            Left            =   780
            TabIndex        =   33
            Top             =   60
            UseMaskColor    =   -1  'True
            Value           =   1  'Checked
            Width           =   210
         End
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Left            =   6120
         TabIndex        =   95
         Top             =   5535
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         _StockProps     =   77
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
         Caption         =   " Wastage(Min.,Max.)"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":440A
         Picture         =   "BookPOChild09.frx":4426
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel40 
         Height          =   330
         Left            =   6120
         TabIndex        =   97
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
         Caption         =   " Calculation"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":4442
         Picture         =   "BookPOChild09.frx":445E
      End
      Begin MSForms.ComboBox Combo4 
         Height          =   330
         Left            =   7650
         TabIndex        =   4
         Top             =   645
         Width           =   6840
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "12065;582"
         ListRows        =   3
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Left            =   1680
         TabIndex        =   13
         Top             =   3580
         Width           =   4455
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "7858;582"
         ListRows        =   4
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   14600
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   14600
         Y1              =   4315
         Y2              =   4315
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   14600
         Y1              =   1375
         Y2              =   1375
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   330
         Left            =   7650
         TabIndex        =   6
         Top             =   960
         Width           =   6840
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "12065;582"
         ListRows        =   3
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   14600
         Y1              =   7125
         Y2              =   7125
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   14600
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   14600
         Y1              =   6280
         Y2              =   6280
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   7650
         TabIndex        =   14
         Top             =   3585
         Width           =   6840
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "12065;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   14600
         Y1              =   5135
         Y2              =   5135
      End
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   570
      Index           =   2
      Left            =   14760
      TabIndex        =   99
      Top             =   1080
      Width           =   1440
      _Version        =   65536
      _ExtentX        =   2540
      _ExtentY        =   1005
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
      Picture         =   "BookPOChild09.frx":447A
      Multiline       =   -1  'True
      GlobalMem       =   -1  'True
      Picture         =   "BookPOChild09.frx":4496
   End
End
Attribute VB_Name = "FrmBookPOChild09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstBookPOChild09 As New ADODB.Recordset, rstBookPOChild0901 As New ADODB.Recordset
Public SizeCode As String, PartyCode As String, RoundOffQty As Boolean
Dim rstPlateMakerList As New ADODB.Recordset, rstSizeList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstPaperList As New ADODB.Recordset, rstFetchRate As New ADODB.Recordset
Dim PlateMakerCode As String, ItemCode As String, PaperCode As String, fPlateCode As String, bPlateCode As String, fColor As Integer, bColor As Integer, fPlate As Integer
Dim PaperBalance As Long, EditMode As Boolean, CutOffSize As Integer
Dim SPU As Long, Wt As Double, ActualQty As Long
Dim TotalPlate As Double, TotalFormFront As Double, TotalFormBack As Double, MaxPrintingQuantity As Double, ActualQuantity As Double, BillingQuantity As Double, FrontPrintingColor As Double, BackPrintingColor As Double, TotalFormsFront As Double, TotalFormsBack As Double, FrontColor As Double, BackColors As Double
Private Sub chkPaper_Validate(Cancel As Boolean)
    If MhRealInput31.Value = 0 Then
       MhRealInput31.Value = 0         'Paper Rate Must Be Zero
    ElseIf MhRealInput31.Value <> 0 And chkPaper.Value = 1 Then
       If MsgBox("Paper By Party Selected  So That Paper Rate Rs.(" & MhRealInput31.Text & ") Must Be Zero  ! Change Rate ?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Change !") = vbYes Then MhRealInput31.Value = 0
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    Me.Left = (MdiMainMenu.ScaleWidth - Me.Width) \ 2
    BusySystemIndicator True
    DisableCloseButton Me
    AbortPO = False
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text)     'Order No.
    Text7.Text = Trim(FrmBookPrintOrder.Text9.Text)    'Vendor Name
    Combo1.AddItem "Old", 0: Combo1.AddItem "New", 1: Combo1.AddItem "Revised", 2: Combo1.AddItem "Cancelled", 3
    Combo2.AddItem "Deep-etch", 0: Combo2.AddItem "PS", 1: Combo2.AddItem "Wipe-on", 2: Combo2.AddItem "CTP", 3
    Combo3.AddItem "F&B", 0: Combo3.AddItem "W&T", 1
    Combo4.AddItem "Single Set Calculation", 0: Combo4.AddItem "Individual Set calculation", 1
    ClearFields
    rstPlateMakerList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstSizeList.Open "SELECT Name As Col0,Code From GeneralMaster WHERE Type='1' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstItemList.Open "SELECT Name As Col0,'' as TitleFrontColor,'' as TitleBackColor,Code From BookMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstPlateMakerList.ActiveConnection = Nothing
    rstSizeList.ActiveConnection = Nothing
    rstItemList.ActiveConnection = Nothing
    Dim Quantity As Long
    With rstBookPOChild0901
        If .State = adStateOpen And .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Quantity = Quantity + .Fields("ActualQuantity").Value
                .MoveNext
            Loop
            .MoveFirst
        End If
    End With
    LoadMasterList
    If Val(CheckNull(Quantity)) = 0 Then  'New Record
        Combo1.ListIndex = 0    'Plate
        Combo2.ListIndex = 3    'Plate Type
        Combo3.ListIndex = 0    'Imposition
        Combo4.ListIndex = 0    'Calculation
        MhDateInput1.Value = FrmBookPrintOrder.MhDateInput1.Value   'Order Date
        MhDateInput3.Value = DateAdd("d", 2, MhDateInput1.Value)    'Target Date
        PlateMakerCode = PartyCode
        If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
        rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
        If Not rstPlateMakerList.EOF Then Text9.Text = rstPlateMakerList.Fields("Col0").Value
        If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
        rstSizeList.Find "[Code] = '" & SizeCode & "'"
        If Not rstSizeList.EOF Then Text4.Text = rstSizeList.Fields("Col0").Value
    Else
        LoadFields
    End If
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}": KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        cmdProceed_Click
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        cmdCancel_Click
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstPlateMakerList)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstFetchRate)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)    'Order Date
    If MhDateInput1.ValueIsNull Then
        Cancel = True
    ElseIf Format(MhDateInput1.Value, "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(MhDateInput1.Value, "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf MhDateInput3.ValueIsNull Then
        MhDateInput3.Value = DateAdd("d", 2, MhDateInput1.Value)
    End If
End Sub
Private Sub MhDateInput3_Validate(Cancel As Boolean)    'Target Date
    If MhDateInput3.ValueIsNull Then
        Cancel = True
    ElseIf Format(MhDateInput3.Value, "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(MhDateInput3.Value, "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        DisplayError ("Target Date cann't be prior to Order Date"): Cancel = True
    End If
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)  'Plate Type
    Dim ItemCode As Variant, i As Integer, BookCode As String
    If Combo2.ListIndex = 1 Or Combo2.ListIndex = 3 Then    'PS/CTP Plate Details
        On Error Resume Next
        With fpSpread1
            For i = 1 To .DataRowCnt
                .GetText 8, i, ItemCode
                BookCode = BookCode + "'" + ItemCode + "',"
            Next
        End With
        BookCode = Left(BookCode, Len(BookCode) - 1)
        If BookCode <> "" Then
            FrmPSPlateRegister.ItemCode = ItemCode
            FrmPSPlateRegister.OrderCode = IIf(CheckNull(rstBookPOChild09.Fields("Code").Value) = "", "999999", rstBookPOChild09.Fields("Code").Value)
            FrmPSPlateRegister.OrderDate = GetDate(MhDateInput1.Text)
            FrmPSPlateRegister.OrderType = "09"
            FrmPSPlateRegister.PlateType = ""
            Load FrmPSPlateRegister
            If Err.Number <> 364 Then FrmPSPlateRegister.Show vbModal
        End If
        On Error GoTo 0
    End If
    Call GetPartyRates("L", "F")   'Plate Rate
End Sub
Private Sub MhRealInput1_GotFocus()       'ActualQuantity
    If MhRealInput1.Value = 0 Then
        MhRealInput1.Value = ActualQuantity
    ElseIf MhRealInput1.Value <> ActualQuantity Then
       If MsgBox("Variation in Calculated Actual Quantity [" & Trim(ActualQuantity) & "]" & vbNewLine & "                           and" & vbNewLine & " Existing Actual Quantity [" & Trim(MhRealInput1.Value) & "] " & vbNewLine & "                           ! Change?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Change !") = vbYes Then MhRealInput1.Value = ActualQuantity
        End If
End Sub
Private Sub MhRealInput2_GotFocus()     'MaxPrintingQuantity
    If MhRealInput2.Value = 0 Then
        MhRealInput2.Value = MaxPrintingQuantity
    ElseIf MhRealInput2.Value <> MaxPrintingQuantity Then
       If MsgBox("Variation in Calculated Total Printing Quantity [" & Trim(MaxPrintingQuantity) & "]" & vbNewLine & "                           and" & vbNewLine & " Existing Total Printing Quantity [" & Trim(MhRealInput2.Value) & "] " & vbNewLine & "                           ! Change?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Change !") = vbYes Then MhRealInput2.Value = MaxPrintingQuantity
        End If
End Sub
Private Sub MhRealInput11_GotFocus()
    If MhDateInput1.ReadOnly Then Exit Sub
    Call GetPartyRates("W", "F")
End Sub
Private Sub MhRealInput23_GotFocus()
    If MhDateInput1.ReadOnly Then Exit Sub
    Call GetPartyRates("M", "F")
End Sub
Private Sub MhRealInput23_Validate(Cancel As Boolean)   'Wastage Min - Front
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateConsumption
    MhRealInput12_Validate False
End Sub
Private Sub MhRealInput29_GotFocus()
    If MhDateInput1.ReadOnly Then Exit Sub
    Call GetPartyRates("X", "F")
End Sub
Private Sub MhRealInput29_Validate(Cancel As Boolean)   'Wastage Max - Front
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateConsumption
    MhRealInput12_Validate False
End Sub
Private Sub MhRealInput36_GotFocus()
    If MhDateInput1.ReadOnly Then Exit Sub
    Call GetPartyRates("M", "B")
End Sub
Private Sub MhRealInput36_Validate(Cancel As Boolean)   'Wastage Min - Back
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateConsumption
End Sub
Private Sub MhRealInput15_GotFocus()    'BillingQuantity
    If MhRealInput15.Value = 0 Then
        MhRealInput15.Value = BillingQuantity
    ElseIf MhRealInput15.Value <> BillingQuantity Then
       If MsgBox("Variation in Calculated Billing Quantity [" & Trim(BillingQuantity) & "]" & vbNewLine & "                           and" & vbNewLine & " Existing Billing Quantity [" & Trim(MhRealInput15.Value) & "] " & vbNewLine & "                           ! Change?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Change !") = vbYes Then MhRealInput15.Value = BillingQuantity
        End If
End Sub
Private Sub MhRealInput19_GotFocus()    'FrontPrintingColor
    If MhRealInput19.Value = 0 Then
        MhRealInput19.Value = FrontPrintingColor
    ElseIf MhRealInput19.Value <> FrontPrintingColor Then
       If MsgBox("Variation in Calculated Front Printing Color [" & Trim(FrontPrintingColor) & "]" & vbNewLine & "                           and" & vbNewLine & " Existing Total Front Printing Color [" & Trim(MhRealInput19.Value) & "] " & vbNewLine & "                           ! Change?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Change !") = vbYes Then MhRealInput19.Value = FrontPrintingColor
        End If
End Sub
Private Sub MhRealInput20_GotFocus()    'BackPrintingColor
    If MhRealInput20.Value = 0 Then
        MhRealInput20.Value = BackPrintingColor
    ElseIf MhRealInput20.Value <> BackPrintingColor Then
       If MsgBox("Variation in Calculated Back Printing Color [" & Trim(BackPrintingColor) & "]" & vbNewLine & "                           and" & vbNewLine & " Existing Back Printing Color [" & Trim(MhRealInput20.Value) & "] " & vbNewLine & "                           ! Change?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Change !") = vbYes Then MhRealInput20.Value = BackPrintingColor
        End If
End Sub
Private Sub MhRealInput3_GotFocus()     'Total Plates
    If MhRealInput3.Value = 0 Then
        MhRealInput3.Value = IIf(Combo3.ListIndex = 0, MhRealInput19.Value + MhRealInput20.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value))
    ElseIf MhRealInput3.Value <> IIf(Combo3.ListIndex = 0, MhRealInput19.Value + MhRealInput20.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value)) Then
       If MsgBox("Variation in Calculated Total Plates [" & Trim(IIf(Combo3.ListIndex = 0, MhRealInput19.Value + MhRealInput20.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value))) & "]" & vbNewLine & "                           and" & vbNewLine & " Existing Total Plates [" & Trim(MhRealInput3.Value) & "] " & vbNewLine & "                           ! Change?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Change !") = vbYes Then MhRealInput3.Value = IIf(Combo3.ListIndex = 0, MhRealInput19.Value + MhRealInput20.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value))
    End If
        Call GetPartyRates("L", "F")   'Plate Rate
    End Sub
   Private Sub MhRealInput6_GotFocus()      'TotalFormFront
    If MhRealInput6.Value = 0 Then
        MhRealInput6.Value = IIf(Combo4.ListIndex = 1, TotalFormsFront, IIf(Combo3.ListIndex = 0, MhRealInput15.Value * MhRealInput19.Value, IIf((MhRealInput15.Value * MhRealInput19.Value) < (MhRealInput15.Value * MhRealInput20.Value), (MhRealInput15.Value * MhRealInput20.Value), (MhRealInput15.Value * MhRealInput19.Value)))) 'TotalFormsFront
    ElseIf MhRealInput6.Value <> IIf(Combo4.ListIndex = 1, TotalFormsFront, IIf(Combo3.ListIndex = 0, MhRealInput15.Value * MhRealInput19.Value, IIf((MhRealInput15.Value * MhRealInput19.Value) < (MhRealInput15.Value * MhRealInput20.Value), (MhRealInput15.Value * MhRealInput20.Value), (MhRealInput15.Value * MhRealInput19.Value)))) Then
       If MsgBox("Variation in Calculated Total Forms Front [" & Trim(IIf(Combo4.ListIndex = 1, TotalFormsFront, IIf(Combo3.ListIndex = 0, MhRealInput15.Value * MhRealInput19.Value, IIf((MhRealInput15.Value * MhRealInput19.Value) < (MhRealInput15.Value * MhRealInput20.Value), (MhRealInput15.Value * MhRealInput20.Value), (MhRealInput15.Value * MhRealInput19.Value))))) & "]" & vbNewLine & "                           and" & vbNewLine & " Existing Total Forms Front [" & Trim(MhRealInput6.Value) & "] " & vbNewLine & "                           ! Change?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Change !") = vbYes Then MhRealInput6.Value = IIf(Combo4.ListIndex = 1, TotalFormsFront, IIf(Combo3.ListIndex = 0, MhRealInput15.Value * MhRealInput19.Value, IIf((MhRealInput15.Value * MhRealInput19.Value) < (MhRealInput15.Value * MhRealInput20.Value), (MhRealInput15.Value * MhRealInput20.Value), (MhRealInput15.Value * MhRealInput19.Value)))) 'Total Forms [Front]MhRealInput6.Value 'TotalFormsFront
    End If
End Sub
Private Sub MhRealInput7_GotFocus()         'TotalFormBack
    If MhRealInput7.Value = 0 Then
        MhRealInput7.Value = IIf(Combo4.ListIndex = 1, TotalFormsBack, IIf(Combo3.ListIndex = 0, MhRealInput15.Value * MhRealInput20.Value, 0)) 'TotalFormsBack
    ElseIf MhRealInput7.Value <> IIf(Combo4.ListIndex = 1, TotalFormsBack, IIf(Combo3.ListIndex = 0, MhRealInput15.Value * MhRealInput20.Value, 0)) Then
       If MsgBox("Variation in Calculated Total Forms Back [" & Trim(IIf(Combo4.ListIndex = 1, TotalFormsBack, IIf(Combo3.ListIndex = 0, MhRealInput15.Value * MhRealInput20.Value, 0))) & "]" & vbNewLine & "                           and" & vbNewLine & " Existing Total Back Front[" & Trim(MhRealInput7.Value) & "]" & vbNewLine & "                           ! Change?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Change !") = vbYes Then MhRealInput7.Value = IIf(Combo4.ListIndex = 1, TotalFormsBack, IIf(Combo3.ListIndex = 0, MhRealInput15.Value * MhRealInput20.Value, 0)) 'TotalFormsBack
        End If
    End Sub
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text9.Text)
        If rstPlateMakerList.RecordCount = 0 Then DisplayError ("No Record in Account Master"): Text9.SetFocus: Exit Sub Else rstPlateMakerList.MoveFirst
        rstPlateMakerList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": PlateMakerCode = ""
        Call LoadSelectionList(rstPlateMakerList, "List of Accounts...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text9, PlateMakerCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(PlateMakerCode) <> "" Then Sendkeys "{TAB}" Else Text9.Text = ""
    End If
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
    If CheckEmpty(Text9.Text, False) Then Cancel = True
End Sub
Private Sub Combo3_Validate(Cancel As Boolean)  'Imposition
'    MhRealInput3.Value = IIf(Combo3.ListIndex = 0, MhRealInput19.Value + MhRealInput20.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value)): Call GetPrinterRates("W"): Call CalculatePlateAmount
If MhRealInput3.Value = 0 Then
        MhRealInput3.Value = IIf(Combo3.ListIndex = 0, MhRealInput19.Value + MhRealInput20.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value))
    ElseIf MhRealInput3.Value <> IIf(Combo3.ListIndex = 0, MhRealInput19.Value + MhRealInput20.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value)) Then
       If MsgBox("Variation in Calculated Total Plates [" & Trim(IIf(Combo3.ListIndex = 0, MhRealInput19.Value + MhRealInput20.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value))) & "]" & vbNewLine & "                           and" & vbNewLine & " Existing Total Plates[" & Trim(MhRealInput3.Value) & "] " & vbNewLine & "                           ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput3.Value = IIf(Combo3.ListIndex = 0, MhRealInput19.Value + MhRealInput20.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value))
: Call GetPartyRates("W"): Call CalculatePlateAmount
        End If
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text4.Text)
        If rstSizeList.RecordCount = 0 Then DisplayError ("No Record in Size Master"): Text4.SetFocus: Exit Sub Else rstSizeList.MoveFirst
        rstSizeList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": SizeCode = ""
        Call LoadSelectionList(rstSizeList, "List of Sizes...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text4, SizeCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(SizeCode) <> "" Then GetPartyRates ("A"): Sendkeys "{TAB}" Else Text4.Text = ""
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    Dim Ups As Integer
    If CheckEmpty(Text4.Text, False) Then Cancel = True
    Ups = CalUps()
    If Ups = 0 Then Cancel = True: Exit Sub
    If Ups > 0 Then
        If Ups <> MhRealInput12.Value And MhRealInput12.Value <> 0 Then
            If MsgBox("Calculated [Ups/Sheet] are different from existing Ups ! Change Ups?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput12.Value = Ups: MhRealInput12_Validate False
        Else
            MhRealInput12.Value = Ups: MhRealInput12_Validate False
        End If
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Color As Variant
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1: fpSpread1.SetFocus: CalculateTotal
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If Not EditMode Then fpSpread1.SetText fpSpread1.ActiveCol, fpSpread1.ActiveRow, ""
    ElseIf KeyCode = vbKeySpace Then
        Dim Item As Variant
        With fpSpread1
            If .ActiveCol = 1 Then
                .GetText .ActiveCol, .ActiveRow, Item
                Text2.Text = FixQuote(Item)
                If rstItemList.RecordCount = 0 Then DisplayError ("No Record in Item Master"): .SetActiveCell 1, .ActiveRow: Exit Sub Else rstItemList.MoveFirst
                rstItemList.Find "[Col0] = '" & RTrim(Item) & "'"
                SelectionType = "S"
                ItemCode = ""
                Call LoadSelectionList(rstItemList, "List of Items...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text2, ItemCode)
                Call CloseForm(FrmSelectionList)
                If CheckEmpty(ItemCode, False) Then
                    .SetActiveCell 1, .ActiveRow
                Else
                    .SetText 1, .ActiveRow, Text2.Text
                    .SetText 8, .ActiveRow, ItemCode
                    .GetText 6, .ActiveRow, Color
                    If Val(Color) = 0 Then
                        .SetText 6, .ActiveRow, Trim(rstItemList.Fields("TitleFrontColor").Value)
                    ElseIf Val(Color) <> rstItemList.Fields("TitleFrontColor").Value Then
                        If MsgBox("Variation in Current (" & Trim(Color) & ") and Master (" & Trim(rstItemList.Fields("TitleFrontColor").Value) & ") Front Color !!! Change Color?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 6, .ActiveRow, Trim(rstItemList.Fields("TitleFrontColor").Value)
                    End If
                    .GetText 7, .ActiveRow, Color
                    If Val(Color) = 0 Then
                        .SetText 7, .ActiveRow, Trim(rstItemList.Fields("TitleBackColor").Value)
                    ElseIf Val(Color) <> rstItemList.Fields("TitleBackColor").Value Then
                        If MsgBox("Variation in Current (" & Trim(Color) & ") and Master (" & Trim(rstItemList.Fields("TitleBackColor").Value) & ") Back Color !!! Change Color?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 7, .ActiveRow, Trim(rstItemList.Fields("TitleBackColor").Value)
                    End If
                    Sendkeys "{ENTER}"
                End If
            End If
        End With
    ElseIf KeyCode = vbKeyF3 Then
        With fpSpread1
            If .ActiveCol = 1 Then
                .GetText 8, .ActiveRow, Item
                On Error Resume Next
                FrmBookMaster.SL = True
                FrmBookMaster.ItemType = "F"
                FrmBookMaster.MasterCode = Item
                Load FrmBookMaster
                If Err.Number <> 364 Then FrmBookMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText 8, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then LoadMasterList: Sendkeys "{ENTER}"
            End If
        End With
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Qty As Variant, Ups As Variant, BillingQty As Long
    With fpSpread1
        If Col = 2 Or Col = 3 Then
            .GetText 2, Row, Qty    'Actual Quantity
            .GetText 3, Row, Ups    'Ups/Plate
            If Val(Qty) > 0 Then
                If Val(Ups) > 0 Then
                    BillingQty = Val(Qty) / Val(Ups)
                        .SetText 4, Row, BillingQty   'Printing Quantity
                    If RoundOffQty Then 'Billing Quantity = Round off Printing Quantity
                        If BillingQty <= 1000 Then BillingQty = 1000 Else BillingQty = IIf(Int(BillingQty / 1000) = 0, 1000, Int(BillingQty / 1000) * 1000) + IIf(BillingQty Mod 1000 <= IIf(BillingQty <= 10000, 299, 599), 0, 1000)
                    End If
                    .GetText 5, Row, Qty    'Billing Quantity
                    If Val(Qty) = 0 Then
                        .SetText 5, Row, BillingQty
                    ElseIf Val(Qty) <> BillingQty Then
                        fpSpread1.SetActiveCell Col + 2, Row: If MsgBox("Variation in Current (" & Trim(Qty) & ") and Calculated (" & Trim(BillingQty) & ") Billing Quantities !!! Change Quantity?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 5, Row, BillingQty
                    End If
                End If
                CalculateTotal
            End If
        ElseIf Col >= 5 Then
            CalculateTotal
        End If
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub MhRealInput6_Validate(Cancel As Boolean)    'Total Forms [Front]
    Call GetPartyRates("P", "F")
End Sub
Private Sub MhRealInput7_Validate(Cancel As Boolean)    'Total Forms [Front]
    Call GetPartyRates("P", "B")
End Sub
Private Sub MhRealInput5_Validate(Cancel As Boolean)    'Print Rate [Front]
    CalculatePrintAmount
End Sub
Private Sub MhRealInput8_Validate(Cancel As Boolean)    'Print Rate [Back]
    MhRealInput5_Validate False
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean)    'Adjustment [Printing]
    MhRealInput5_Validate False
End Sub
Private Sub MhRealInput18_Validate(Cancel As Boolean)   'GST% [Printing]
    MhRealInput5_Validate False
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)    'Plate Rate
    CalculatePlateAmount
End Sub
Private Sub MhRealInput26_Validate(Cancel As Boolean)   'Adjustment [Plate]
    MhRealInput4_Validate False
End Sub
Private Sub MhRealInput21_Validate(Cancel As Boolean)   'GST% [Plate]
    MhRealInput4_Validate False
End Sub
Private Sub MhRealInput33_Validate(Cancel As Boolean)   'Adjustment [Paper]
    MhRealInput31_Validate False
End Sub
Private Sub MhRealInput34_Validate(Cancel As Boolean)   'GST% [Paper]
    MhRealInput31_Validate False
End Sub
Private Sub MhRealInput31_Validate(Cancel As Boolean)    'Paper Rate
    CalculatePaperAmount
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        LoadMasterList True
        With FrmPaperSearchList
            Set .rstPaperSearchList = rstPaperList
            FrmPaperSearchList.MhRealInput3.Value = Val(Left(Text4.Text, 5))
            FrmPaperSearchList.MhRealInput5.Value = Val(Mid(Text4.Text, 7, 5))
            FrmPaperSearchList.Text1.Text = Text1.Text: Sendkeys "{End}"
            Load FrmPaperSearchList
            .Show vbModal
            If Not CheckEmpty(.PaperCode, False) Then PaperCode = .PaperCode: Text1.Text = .PaperName: Sendkeys "{TAB}"
        End With
        Call CloseForm(FrmPaperSearchList)
    ElseIf KeyCode = vbKeyF3 Then
        On Error Resume Next
        FrmPaperMaster.SL = True
        FrmPaperMaster.MasterCode = PaperCode
        Load FrmPaperMaster
        If Err.Number <> 364 Then FrmPaperMaster.Show vbModal
        On Error GoTo 0
        PaperCode = slCode: Text1.Text = slName
        If Not CheckEmpty(PaperCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        PaperCode = "": Text1.Text = ""
    End If
End Sub
Private Sub Text1_Validate(Cancel As Boolean)   'Paper
    If CheckEmpty(Text1.Text, False) Then Cancel = True
    If Not CheckEmpty(PaperCode, False) Then
        rstPaperList.MoveFirst
        rstPaperList.Find "[Code]='" & PaperCode & "'"
        Text1.Text = rstPaperList.Fields("Col0").Value: SPU = Val(rstPaperList.Fields("SPU").Value)
        If rstPaperList.Fields("Form").Value = "R" Then
            Do While True
                CutOffSize = InputBox("Reel Cut Off (mm)", "Easy Publish", Val(CutOffSize))
                If Val(CutOffSize) = 0 Then DisplayError ("Reel Cut off Size cann't be zero"): Cancel = True Else Exit Do
            Loop
        Else
            CutOffSize = 0
        End If
    End If
    Dim Ups As Integer
    Ups = CalUps()
    If Ups = 0 Then Cancel = True: Exit Sub
    If Ups > 0 Then
        If Ups <> MhRealInput12.Value And MhRealInput12.Value <> 0 Then
            If MsgBox("Calculated [Ups/Sheet] are different from existing Ups ! Change Ups?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput12.Value = Ups: MhRealInput12_Validate False
        Else
            MhRealInput12.Value = Ups: MhRealInput12_Validate False
        End If
    End If
End Sub
Private Sub MhRealInput12_Validate(Cancel As Boolean)
    CalculateConsumption
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)
    MhRealInput12_Validate False
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateConsumption
End Sub
Private Sub ClearFields()
    MhDateInput1.Value = Date
    MhDateInput3.Value = DateAdd("d", 2, MhDateInput1.Value)
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Text9.Text = ""
    Combo3.ListIndex = 0
    Combo4.ListIndex = 0
    Text4.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    Text2.Text = ""
    MhRealInput1.Value = 0
    MhRealInput2.Value = 0
    MhRealInput15.Value = 0
    MhRealInput19.Value = 0
    MhRealInput20.Value = 0
    MhRealInput6.Value = 0
    MhRealInput7.Value = 0
    chkPaper.Value = 1
    MhRealInput5.Value = 0
    MhRealInput8.Value = 0
    MhRealInput9.Value = 0
    MhRealInput18.Value = 0
    MhRealInput17.Value = 0
    MhRealInput25.Value = 0
    MhRealInput3.Value = 0
    MhRealInput4.Value = 0
    MhRealInput26.Value = 0
    MhRealInput33.Value = 0
    MhRealInput21.Value = 0
    MhRealInput34.Value = 0
    MhRealInput22.Value = 0
    MhRealInput35.Value = 0
    MhRealInput27.Value = 0
    Text1.Text = "": PaperCode = ""
    MhRealInput12.Value = 0
    MhRealInput28.Value = 0
    MhRealInput11.Value = 0
    MhRealInput23.Value = 0
    MhRealInput29.Value = 0
    MhRealInput30.Value = 0
    MhRealInput13.Value = 0
    Text8.Text = ""
    MhDateInput2.Value = ""
    MhRealInput16.Value = 0
    Text10.Text = ""
    MhDateInput4.Value = ""
    MhRealInput24.Value = 0
    MhRealInput31.Value = 0
    MhRealInput32.Value = 0
    MhRealInput36.Value = 0
    MhRealInput37.Value = 0
    MhRealInput38.Value = 0
    Text6.Text = ""
    CutOffSize = 0
End Sub
Private Sub LoadFields()
    With rstBookPOChild09
        MhDateInput1.Value = .Fields("OrderDate").Value
        MhRealInput1.Value = .Fields("ActualQuantity").Value
        MhRealInput2.Value = .Fields("MaxPrintingQuantity").Value
        MhRealInput15.Value = .Fields("BillingQuantity").Value
        MhRealInput19.Value = .Fields("FrontPrintingColor").Value
        MhRealInput20.Value = .Fields("BackPrintingColor").Value
        MhDateInput3.Value = .Fields("TargetDate").Value
        Combo1.ListIndex = IIf(.Fields("Plate").Value = "O", 0, IIf(.Fields("Plate").Value = "N", 1, IIf(.Fields("Plate").Value = "R", 2, 3))) 'O:Old N:New R:Revised C:Cancelled
        Combo2.ListIndex = Val(.Fields("PlateType").Value) - 1
        PlateMakerCode = .Fields("PlateMaker").Value
        If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
        rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
        If Not rstPlateMakerList.EOF Then Text9.Text = Trim(rstPlateMakerList.Fields("Col0").Value)
        Combo3.ListIndex = IIf(.Fields("Imposition").Value = "F", 0, 1)
        Combo4.ListIndex = IIf(.Fields("Calculation").Value = "S", 0, 1)
        SizeCode = .Fields("Size").Value
        If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
        rstSizeList.Find "[Code] = '" & SizeCode & "'"
        If Not rstSizeList.EOF Then Text4.Text = rstSizeList.Fields("Col0").Value
        Call LoadItemList
        MhRealInput6.Value = .Fields("TotalFormsFront").Value
        MhRealInput7.Value = .Fields("TotalFormsBack").Value
        chkPaper.Value = IIf(.Fields("PaperByParty").Value, 1, 0)
        MhRealInput5.Value = .Fields("PrintRateFront").Value
        MhRealInput8.Value = .Fields("PrintRateBack").Value
        MhRealInput36.Value = .Fields("PrintAmountBT").Value
        MhRealInput9.Value = .Fields("Adjustment").Value
        MhRealInput18.Value = .Fields("GST%").Value
        MhRealInput17.Value = .Fields("GST").Value
        MhRealInput25.Value = .Fields("PrintAmount").Value
        MhRealInput3.Value = .Fields("TotalPlates").Value
        MhRealInput4.Value = .Fields("PlateRate").Value
        MhRealInput37.Value = .Fields("PlateAmountBT").Value
        MhRealInput26.Value = .Fields("PAdjustment").Value
        MhRealInput21.Value = .Fields("PGST%").Value
        MhRealInput22.Value = .Fields("PGST").Value
        MhRealInput27.Value = .Fields("PlateAmount").Value
        MhRealInput31.Value = .Fields("PaperRate").Value
        MhRealInput38.Value = .Fields("PaperAmountBT").Value
        MhRealInput33.Value = .Fields("RAdjustment").Value
        MhRealInput34.Value = .Fields("RGST%").Value
        MhRealInput35.Value = .Fields("RGST").Value
        MhRealInput32.Value = .Fields("PaperAmount").Value
        PaperCode = .Fields("Paper").Value
        If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
        rstPaperList.Find "[Code] = '" & PaperCode & "'"
        If Not rstPaperList.EOF Then Text1.Text = rstPaperList.Fields("Col0").Value: SPU = Val(rstPaperList.Fields("SPU").Value): Wt = Val(rstPaperList.Fields("Wt").Value)
        MhRealInput12.Value = .Fields("Ups/Sheet").Value
        MhRealInput28.Value = .Fields("PaperConsumption").Value
        MhRealInput11.Value = .Fields("PaperWastage%").Value
        MhRealInput23.Value = .Fields("PaperWastageMin").Value
        MhRealInput29.Value = .Fields("PaperWastageMax").Value
        MhRealInput30.Value = .Fields("PaperWastageFinal").Value
        MhRealInput13.Value = .Fields("PaperConsumptionOther").Value
        Text8.Text = .Fields("BillNo").Value
        If Not IsNull(.Fields("BillDate").Value) Then MhDateInput2.Value = .Fields("BillDate").Value
        MhRealInput16.Value = .Fields("PaidAmount").Value
        Text10.Text = .Fields("PBillNo").Value
        If Not IsNull(.Fields("PBillDate").Value) Then MhDateInput4.Value = .Fields("PBillDate").Value
        MhRealInput24.Value = .Fields("PPaidAmount").Value
        Text6.Text = .Fields("Remarks").Value
        CutOffSize = .Fields("CutOffSize").Value
    End With
End Sub
Private Sub SaveFields()
    With rstBookPOChild09
        .Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
        .Fields("ActualQuantity").Value = MhRealInput1.Value
        .Fields("MaxPrintingQuantity").Value = MhRealInput2.Value
        .Fields("BillingQuantity").Value = MhRealInput15.Value
        .Fields("FrontPrintingColor").Value = MhRealInput19.Value
        .Fields("BackPrintingColor").Value = MhRealInput20.Value
        .Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
        .Fields("Plate").Value = Choose(Combo1.ListIndex + 1, "O", "N", "R", "C")
        .Fields("PlateType").Value = Trim(Str(Combo2.ListIndex + 1))
        .Fields("PlateMaker").Value = PlateMakerCode
        .Fields("Imposition").Value = Choose(Combo3.ListIndex + 1, "F", "W")
        .Fields("Calculation").Value = Choose(Combo4.ListIndex + 1, "S", "M")
        .Fields("Size").Value = SizeCode
        UpdateItemList
        .Fields("ActualQuantity").Value = ActualQty
        .Fields("TotalFormsFront").Value = MhRealInput6.Value
        .Fields("TotalFormsBack").Value = MhRealInput7.Value
        .Fields("PrintRateFront").Value = MhRealInput5.Value
        .Fields("PrintRateBack").Value = MhRealInput8.Value
        .Fields("PrintAmountBT").Value = MhRealInput36.Value
        .Fields("Adjustment").Value = MhRealInput9.Value
        .Fields("GST%").Value = MhRealInput18.Value
        .Fields("GST").Value = MhRealInput17.Value
        .Fields("PrintAmount").Value = MhRealInput25.Value
        .Fields("TotalPlates").Value = MhRealInput3.Value
        .Fields("PlateRate").Value = MhRealInput4.Value
        .Fields("PlateAmountBT").Value = MhRealInput37.Value
        .Fields("PAdjustment").Value = MhRealInput26.Value
        .Fields("RAdjustment").Value = MhRealInput33.Value
        .Fields("PGST%").Value = MhRealInput21.Value
        .Fields("PGST").Value = MhRealInput22.Value
        .Fields("PlateAmount").Value = MhRealInput27.Value
        .Fields("PaperByParty").Value = chkPaper.Value
        .Fields("PaperRate").Value = MhRealInput31.Value
        .Fields("PaperAmountBT").Value = MhRealInput38.Value
        .Fields("RGST%").Value = MhRealInput34.Value
        .Fields("RGST").Value = MhRealInput35.Value
        .Fields("PaperAmount").Value = MhRealInput32.Value
        .Fields("Paper").Value = PaperCode
        .Fields("Ups/Sheet").Value = MhRealInput12.Value
        .Fields("PaperConsumption").Value = MhRealInput28.Value
        .Fields("PaperWastage%").Value = MhRealInput11.Value
        .Fields("PaperWastageMin").Value = MhRealInput23.Value
        .Fields("PaperWastageMax").Value = MhRealInput29.Value
        .Fields("PaperWastageFinal").Value = MhRealInput30.Value
        .Fields("PaperConsumptionOther").Value = MhRealInput13.Value
        .Fields("PaperConsumptionSheets").Value = CLng(Int(MhRealInput13.Value) * SPU) + ((MhRealInput13.Value - Int(MhRealInput13.Value)) * 1000)
        .Fields("BillNo").Value = Text8.Text
        If MhDateInput2.ValueIsNull Then .Fields("BillDate").Value = Null Else .Fields("BillDate").Value = GetDate(MhDateInput2.Text)
        .Fields("PaidAmount").Value = MhRealInput16.Value
        .Fields("PBillNo").Value = Text10.Text
        If MhDateInput4.ValueIsNull Then .Fields("PBillDate").Value = Null Else .Fields("PBillDate").Value = GetDate(MhDateInput4.Text)
        .Fields("PPaidAmount").Value = MhRealInput24.Value
        .Fields("Remarks").Value = Text6.Text
        If Not CheckEmpty(Text8.Text, False) Then If IsNull(.Fields("BillFeedDate").Value) Then .Fields("BillFeedDate").Value = Now()
        .Fields("CutOffSize").Value = CutOffSize
        Dim lpBuff As String * 1024
        GetComputerName lpBuff, Len(lpBuff)
        If Not CheckEmpty(Text8.Text, False) Then If IsNull(.Fields("ComputerName").Value) Then .Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    End With
End Sub
Private Sub LoadItemList()
    On Error GoTo ErrHandler
    Dim i As Integer
    With rstBookPOChild0901
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            i = i + 1
            fpSpread1.SetText 1, i, .Fields("ItemName").Value
            fpSpread1.SetText 2, i, Val(.Fields("ActualQuantity").Value)
            fpSpread1.SetText 3, i, Val(.Fields("Ups/Plate").Value)
            fpSpread1.SetText 4, i, Val(.Fields("PrintingQuantity").Value)
            fpSpread1.SetText 5, i, Val(.Fields("BillingQuantity").Value)
            fpSpread1.SetText 6, i, Trim(.Fields("FrontPrintingColor").Value)
            fpSpread1.SetText 7, i, Trim(.Fields("BackPrintingColor").Value)
            fpSpread1.SetText 8, i, .Fields("ItemCode").Value
            .MoveNext
        Loop
    End With
    CalculateTotal
    Exit Sub
ErrHandler:
    DisplayError ("Failed to Load Item List")
End Sub
Private Sub UpdateItemList()
    Dim i As Integer, Qty As Variant
    With rstBookPOChild0901
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            .Delete: .MoveNext
        Loop
    End With
    ActualQty = 0
    For i = 1 To fpSpread1.DataRowCnt
        Call AddRecord(rstBookPOChild0901)
        With rstBookPOChild0901
            fpSpread1.GetText 1, i, Qty: .Fields("ItemName").Value = Qty
            fpSpread1.GetText 2, i, Qty: .Fields("ActualQuantity").Value = Val(Qty): ActualQty = ActualQty + Val(Qty)
            fpSpread1.GetText 3, i, Qty: .Fields("Ups/Plate").Value = Val(Qty)
            fpSpread1.GetText 4, i, Qty: .Fields("PrintingQuantity").Value = Val(Qty)
            fpSpread1.GetText 5, i, Qty: .Fields("BillingQuantity").Value = Val(Qty)
            fpSpread1.GetText 6, i, Qty: .Fields("FrontPrintingColor").Value = Val(Qty)
            fpSpread1.GetText 7, i, Qty: .Fields("BackPrintingColor").Value = Val(Qty)
            fpSpread1.GetText 8, i, Qty: .Fields("ItemCode").Value = Qty
        End With
        rstBookPOChild0901.Update
    Next
End Sub
Private Sub CalculateTotal()
    Dim i As Integer, Qty As Variant
    With fpSpread1
        ActualQuantity = 0: MaxPrintingQuantity = 0: BillingQuantity = 0: FrontPrintingColor = 0: BackPrintingColor = 0: TotalFormsFront = 0: TotalFormsBack = 0: FrontColor = 0: BackColors = 0
        For i = 1 To .DataRowCnt
            .GetText 2, i, Qty: ActualQuantity = ActualQuantity + Val(Qty)                                                'MhRealInput1.Value = MhRealInput1.Value + Val(Qty)
            If Combo4.ListIndex = 1 Then 'Indiviusal Calculations
            .GetText 4, i, Qty: MaxPrintingQuantity = MaxPrintingQuantity + Val(Qty)                                   'MhRealInput2.Value Then MhRealInput2.Value = Val(Qty)
            Else
            .GetText 4, i, Qty: If Val(Qty) > MaxPrintingQuantity Then MaxPrintingQuantity = Val(Qty)
            End If
            If Combo4.ListIndex = 1 Then 'Indiviusal Calculations
            .GetText 5, i, Qty: BillingQuantity = BillingQuantity + Val(Qty)                                                   'MhRealInput15.Value Then MhRealInput15.Value = Val(Qty)
            Else
            .GetText 5, i, Qty: If Val(Qty) > BillingQuantity Then BillingQuantity = Val(Qty)
            End If
            If Combo4.ListIndex = 1 Then 'Indiviusal Calculations
            .GetText 6, i, Qty: FrontPrintingColor = FrontPrintingColor + Val(Qty)                                          'MhRealInput19.Value Then MhRealInput19.Value = Val(Qty)
            Else
            .GetText 6, i, Qty: If Val(Qty) > FrontPrintingColor Then FrontPrintingColor = Val(Qty)
            End If
            If Combo4.ListIndex = 1 Then 'Indiviusal Calculations
            .GetText 7, i, Qty: BackPrintingColor = BackPrintingColor + Val(Qty)                                           'MhRealInput20.Value Then MhRealInput20.Value = Val(Qty)
            Else
            .GetText 7, i, Qty: If Val(Qty) > BackPrintingColor Then BackPrintingColor = Val(Qty)
            End If
            If Combo4.ListIndex = 1 Then 'FrontColor
            .GetText 6, i, Qty: FrontColor = Qty
            Else
            .GetText 6, i, Qty: If Val(Qty) > FrontColor Then FrontColor = Val(Qty)
            End If
            If Combo4.ListIndex = 1 Then 'BackColors
            .GetText 7, i, Qty: BackColors = Qty
            Else
            .GetText 7, i, Qty: If Val(Qty) > BackColors Then BackColors = Val(Qty)
            End If
            If Combo3.ListIndex = 0 And Combo4.ListIndex = 1 Then 'Indiviusal Calculations Total Forms Front
            .GetText 5, i, Qty: TotalFormsFront = TotalFormsFront + (Val(Qty) * FrontColor)
            ElseIf Combo3.ListIndex = 1 And Combo4.ListIndex = 1 Then 'Indiviusal Calculations Total Forms Front
            .GetText 5, i, Qty: TotalFormsFront = TotalFormsFront + ((Val(Qty) * IIf(FrontColor > BackColors, FrontColor, BackColors)))
            Else
            .GetText 5, i, Qty: If (Val(Qty) * FrontColor) > TotalFormsFront Then TotalFormsFront = (Val(Qty) * FrontColor)
            End If
            If Combo3.ListIndex = 0 And Combo4.ListIndex = 1 Then 'Indiviusal Calculations Total Forms Back
            .GetText 5, i, Qty: TotalFormsBack = TotalFormsBack + (Val(Qty) * BackColors)
            ElseIf Combo3.ListIndex = 1 And Combo4.ListIndex = 1 Then 'Indiviusal Calculations Total Forms Back
            .GetText 5, i, Qty: TotalFormsBack = TotalFormsBack + (Val(Qty) * 0)
            Else
            .GetText 5, i, Qty: If (Val(Qty) * BackColors) > TotalFormsBack Then TotalFormsBack = (Val(Qty) * BackColors)
            End If
        Next
         If MhRealInput2.Value = 0 Then
            MhRealInput2.Value = MaxPrintingQuantity
        ElseIf MhRealInput2.Value <> MaxPrintingQuantity Then
        End If
'      MaxPrintingQuantity = IIf(Val(Qty) > MaxPrintingQuantity, MaxPrintingQuantity, Val(Qty))
        TotalPlate = IIf(Combo3.ListIndex = 0, MhRealInput19.Value + MhRealInput20.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value))     'MhRealInput3.Value
'        MhRealInput3.Value = TotalPlate
        TotalFormFront = IIf(Combo4.ListIndex = 1, TotalFormsFront, IIf(Combo3.ListIndex = 0, MhRealInput15.Value * MhRealInput19.Value, IIf((MhRealInput15.Value * MhRealInput19.Value) < (MhRealInput15.Value * MhRealInput20.Value), (MhRealInput15.Value * MhRealInput20.Value), (MhRealInput15.Value * MhRealInput19.Value)))) 'Total Forms [Front]MhRealInput6.Value
'        MhRealInput6.Value = TotalFormFront
        TotalFormBack = IIf(Combo4.ListIndex = 1, TotalFormsBack, IIf(Combo3.ListIndex = 0, MhRealInput15.Value * MhRealInput20.Value, 0)) 'Total Forms [Back]MhRealInput7.Value
'        MhRealInput7.Value = TotalFormBack
    End With
End Sub
Private Sub GetPartyRates(ByVal RateType As String, Optional ByVal Position As String)
    If (MhRealInput6.Value + MhRealInput7.Value) = 0 Or CheckEmpty(SizeCode, False) Or (MhRealInput19.Value + MhRealInput20.Value) = 0 Then Exit Sub
    Dim frontPlateRate As Double, backPlateRate As Double, frontPrintRate As Double, backPrintRate As Double, frontPaperWastageRate As Double, backPaperWastageRate As Double, frontPaperWastageMin As Long, backPaperWastageMin As Long, frontPaperWastageMax As Long, backPaperWastageMax As Long
    On Error GoTo ErrorHandler
    fColor = FrontPrintingColor: bColor = BackPrintingColor:
    'Fetching Front Rates
    With rstFetchRate
        If (MhRealInput19.Value + MhRealInput20.Value) <> 0 Then
            If RateType = "L" Then  'Plate Rate
            fPlateCode = " Like '%" & Choose(Combo2.ListIndex + 1, "Deep-etch", "PS", "Wipe-on", "CTP") & "%'"
            bPlateCode = " Like '%" & Choose(Combo2.ListIndex + 1, "Deep-etch", "PS", "Wipe-on", "CTP") & "%'"
                If .State = adStateOpen Then .Close
                .Open "SELECT TOP 1 P.* FROM AccountChild06 P INNER JOIN SizeGroupChild C ON P.[SizeGroup]=C.Code WHERE P.Code='" & PartyCode & "' AND C.[Size]='" & SizeCode & "' AND (Select Name From GeneralMaster Where Code=P.Plate)" & fPlateCode & " AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
                If .RecordCount = 0 Then
                    If .State = adStateOpen Then .Close
                    .Open "SELECT TOP 1 C1.* FROM (AccountMaster P INNER JOIN AccountChild06 C1 ON P.Code=C1.Code) INNER JOIN SizeGroupChild C2 ON C1.[SizeGroup]=C2.Code WHERE [Name] LIKE '%Rate%' AND C2.[Size]='" & SizeCode & "' AND (Select Name From GeneralMaster Where Code=C1.Plate)" & fPlateCode & " AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
                End If
                If .RecordCount > 0 Then frontPlateRate = Val(.Fields("Rate").Value)
            Else 'Ptg Rate
                If .State = adStateOpen Then .Close
                .Open "SELECT TOP 1 P.* FROM AccountChild05 P INNER JOIN GeneralMaster G ON P.Color=G.Code INNER JOIN SizeGroupChild C ON P.[SizeGroup]=C.Code WHERE P.Code='" & PartyCode & "' AND C.[Size]='" & SizeCode & "' AND Value1='" & fColor & "' AND [Range]>=" & MhRealInput6.Value & " AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY [Range],wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
                If .RecordCount = 0 Then
                    If .State = adStateOpen Then .Close
                    .Open "SELECT TOP 1 C1.* FROM (AccountMaster P INNER JOIN AccountChild05 C1 ON P.Code=C1.Code) INNER JOIN GeneralMaster G ON C1.Color=G.Code INNER JOIN SizeGroupChild C2 ON C1.[SizeGroup]=C2.Code WHERE P.Name LIKE '%Rate%' AND C2.[Size]='" & SizeCode & "' AND G.Value1=" & fColor & " AND [Range]>=" & MhRealInput6.Value & " AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY [Range],wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
                End If
                If .RecordCount > 0 Then
                    If RateType = "P" Then  'Print Rate
                        frontPrintRate = Val(.Fields("PrintingRate").Value)
                    ElseIf RateType = "W" Then  'Paper Wastage (Percentage)
                        frontPaperWastageRate = Val(.Fields("PaperWastageRate").Value)
                    ElseIf RateType = "M" Then  'Paper Wastage (Minimum Sheets)
                        frontPaperWastageMin = Val(.Fields("PaperWastageMin").Value)
                    ElseIf RateType = "X" Then  'Paper Wastage (Max Sheets)
                        frontPaperWastageMax = Val(.Fields("PaperWastageMax").Value)
                    End If
                End If
            End If
        End If
        'Fetching Back Rates
        If (MhRealInput19.Value + MhRealInput20.Value) <> 0 Then
            If RateType = "L" Then  'Plate Rate
'                If .State = adStateOpen Then .Close
'                .Open "SELECT TOP 1 P.* FROM AccountChild06 P INNER JOIN SizeGroupChild C ON P.[SizeGroup]=C.Code WHERE P.Code='" & PartyCode & "' AND C.[Size]='" & SizeCode & "' AND [Plate]='" & bPlateCode & "' AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
'                If .RecordCount = 0 Then
'                    If .State = adStateOpen Then .Close
'                    .Open "SELECT TOP 1 C1.* FROM (AccountMaster P INNER JOIN AccountChild06 C1 ON P.Code=C1.Code) INNER JOIN SizeGroupChild C2 ON C1.[SizeGroup]=C2.Code WHERE [Name] LIKE '%Rate%' AND C2.[Size]='" & SizeCode & "' AND [Plate]='" & bPlateCode & "' AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
'                End If
'                If .RecordCount > 0 Then backPlateRate = Val(.Fields("Rate").Value)
            Else
                If .State = adStateOpen Then .Close
                .Open "SELECT TOP 1 P.* FROM AccountChild05 P INNER JOIN GeneralMaster G ON P.Color=G.Code INNER JOIN SizeGroupChild C ON P.[SizeGroup]=C.Code WHERE P.Code='" & PartyCode & "' AND C.[Size]='" & SizeCode & "' AND [Value1]='" & bColor & "' AND [Range]>=" & MhRealInput6.Value & " AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY [Range],wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
                If .RecordCount = 0 Then
                    If .State = adStateOpen Then .Close
                    .Open "SELECT TOP 1 C1.* FROM (AccountMaster P INNER JOIN AccountChild05 C1 ON P.Code=C1.Code) INNER JOIN GeneralMaster G ON C1.Color=G.Code INNER JOIN SizeGroupChild C2 ON C1.[SizeGroup]=C2.Code WHERE P.Name LIKE '%Rate%' AND C2.[Size]='" & SizeCode & "' AND G.Value1=" & bColor & " AND [Range]>=" & MhRealInput6.Value & " AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY [Range],wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
                End If
                If .RecordCount > 0 Then
                    If RateType = "P" Then  'Print Rate
                        backPrintRate = Val(.Fields("PrintingRate").Value)
                    ElseIf RateType = "W" Then  'Paper Wastage (Percentage)
                        backPaperWastageRate = Val(.Fields("PaperWastageRate").Value)
                    ElseIf RateType = "M" Then  'Paper Wastage (Minimum Sheets)
                        backPaperWastageMin = Val(.Fields("PaperWastageMin").Value)
                    ElseIf RateType = "X" Then  'Paper Wastage (Max Sheets)
                        backPaperWastageMax = Val(.Fields("PaperWastageMax").Value)
                    End If
                End If
            End If
        End If
    End With
    'Value Posting
    If RateType = "L" Then
        If Position = "F" Then
            If MhRealInput3.Value > 0 Then 'total front plates
                If Combo1.ListIndex > 0 Then 'not old
                    If frontPlateRate > 0 Then
                        If MhRealInput4.Value = 0 Then
                            MhRealInput4.Value = frontPlateRate
                        ElseIf MhRealInput4.Value <> frontPlateRate Then
                            If MsgBox("Front Plate Rate [" & Trim(MhRealInput4.Value) & "] is different from that in Master [" & Trim(Format(frontPlateRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput4.Value = frontPlateRate
                        End If
                    End If
                Else
                    If fPlate Then MhRealInput4.Value = 0 'old 3 times usable plates
                End If
            Else
                MhRealInput4.Value = 0
            End If
        End If
        If Position = "B" Then
            If MhRealInput37.Value > 0 Then 'total back plates
                If Combo1.ListIndex > 0 Then 'not old
                    If backPlateRate > 0 Then
                        If MhRealInput4.Value = 0 Then
                            MhRealInput4.Value = backPlateRate
                        ElseIf MhRealInput4.Value <> backPlateRate Then
                            If MsgBox("Back Plate Rate [" & Trim(MhRealInput4.Value) & "] is different from that in Master [" & Trim(Format(backPlateRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput4.Value = backPlateRate
                        End If
                    End If
                Else
                    If fPlate Then MhRealInput4.Value = 0
                End If
            Else
                MhRealInput4.Value = 0
            End If
        End If
    ElseIf RateType = "P" Then
        If Position = "F" Then
            If MhRealInput6.Value > 0 Then 'Total Forms
                If frontPrintRate > 0 Then
                    If MhRealInput5.Value = 0 Then
                        MhRealInput5.Value = frontPrintRate
                    ElseIf MhRealInput5.Value <> frontPrintRate Then
                        If MsgBox("Front Print Rate [" & Trim(MhRealInput5.Value) & "] is different from that in Master [" & Trim(Format(frontPrintRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput5.Value = frontPrintRate
                    End If
                End If
            Else
                MhRealInput5.Value = 0
            End If
        End If
        If Position = "B" Then
            If MhRealInput7.Value > 0 Then 'Total back plates
                If backPrintRate > 0 Then
                    If MhRealInput8.Value = 0 Then
                        MhRealInput8.Value = backPrintRate
                    ElseIf MhRealInput8.Value <> backPrintRate Then
                        If MsgBox("Back Print Rate [" & Trim(MhRealInput8.Value) & "] is different from that in Master [" & Trim(Format(backPrintRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput8.Value = backPrintRate
                    End If
                End If
            Else
                MhRealInput8.Value = 0
            End If
        End If
    ElseIf RateType = "W" Then
        If Position = "F" Then
            If fColor > 0 Then
                If frontPaperWastageRate > 0 Then
                    If MhRealInput11.Value = 0 Then
                        MhRealInput11.Value = frontPaperWastageRate
                    ElseIf MhRealInput11.Value <> frontPaperWastageRate Then
                        If MsgBox("Front Paper Wastage Rate [" & Trim(MhRealInput11.Value) & "] is different from that in Master [" & Trim(Format(frontPaperWastageRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput11.Value = frontPaperWastageRate
                    End If
                End If
            Else
                MhRealInput11.Value = 0
            End If
        End If
        If Position = "B" Then
            If IIf(Combo3.ListIndex = 0, bColor > 0, False) Then
                If backPaperWastageRate > 0 Then
                    If MhRealInput35.Value = 0 Then
                        MhRealInput35.Value = backPaperWastageRate
                    ElseIf MhRealInput35.Value <> backPaperWastageRate Then
                        If MsgBox("Back Paper Wastage Rate [" & Trim(MhRealInput35.Value) & "] is different from that in Master [" & Trim(Format(backPaperWastageRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput35.Value = backPaperWastageRate
                    End If
                End If
            Else
                MhRealInput35.Value = 0
            End If
        End If
    ElseIf RateType = "M" Then
        If Position = "F" Then
            If fColor > 0 Then
                If frontPaperWastageMin > 0 Then
                    If MhRealInput23.Value = 0 Then
                        MhRealInput23.Value = frontPaperWastageMin
                    ElseIf MhRealInput23.Value <> frontPaperWastageMin Then
                        If MsgBox("Front Paper Wastage Min [" & Trim(MhRealInput23.Value) & "] is different from that in Master [" & Trim(Format(frontPaperWastageMin, "#0")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput23.Value = frontPaperWastageMin
                    End If
                End If
            Else
                MhRealInput23.Value = 0
            End If
        End If
    ElseIf RateType = "X" Then
        If Position = "F" Then
            If fColor > 0 Then
                If frontPaperWastageMax > 0 Then
                    If MhRealInput29.Value = 0 Then
                        MhRealInput29.Value = frontPaperWastageMax
                    ElseIf MhRealInput29.Value <> frontPaperWastageMax Then
                        If MsgBox("Front Paper Wastage Max [" & Trim(MhRealInput29.Value) & "] is different from that in Master [" & Trim(Format(frontPaperWastageMax, "#0")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput29.Value = frontPaperWastageMax
                    End If
                End If
            Else
                MhRealInput29.Value = 0
            End If
        End If
        If Position = "B" Then
            If IIf(Combo3.ListIndex = 0, bColor > 0, False) Then
                If backPaperWastageMin > 0 Then
                    If MhRealInput36.Value = 0 Then
                        MhRealInput36.Value = backPaperWastageMin
                    ElseIf MhRealInput36.Value <> backPaperWastageMin Then
                        If MsgBox("Back Paper Wastage Min [" & Trim(MhRealInput36.Value) & "] is different from that in Master [" & Trim(Format(backPaperWastageMin, "#0")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput36.Value = backPaperWastageMin
                    End If
                End If
            Else
                MhRealInput36.Value = 0
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub
'Private Sub GetPrinterRates(ByVal RateFor As String)
'    Dim PlateRate As Double, PrintRateFront As Double, PrintRateBack As Double, PaperWastageRate As Double, PaperWastageMin As Long
'    If MhRealInput2.Value = 0 Then Exit Sub
'    On Error GoTo ErrorHandler
'    With rstPrinterRates
'        If .State = adStateOpen Then .Close
'        .Open "SELECT TOP 1 P.* FROM AccountChild05 P INNER JOIN SizeGroupChild C ON P.[SizeGroup]=C.Code WHERE C.[Size]='" & SizeCode & "' AND P.Code='" & PrinterCode & "'  AND Range>=" & MhRealInput2.Value & " ORDER BY Range", cnDatabase, adOpenKeyset, adLockReadOnly
'        If .RecordCount = 0 Then
'            If .State = adStateOpen Then .Close
'            .Open "SELECT TOP 1 C1.* FROM (AccountMaster P INNER JOIN AccountChild05 C1 ON P.Code=C1.Code) INNER JOIN SizeGroupChild C2 ON C1.[SizeGroup]=C2.Code WHERE Name LIKE '%Rate%' AND C2.[Size]='" & SizeCode & "' AND Range>= " & MhRealInput2.Value & " ORDER BY Range", cnDatabase, adOpenKeyset, adLockReadOnly
'        End If
'        If .RecordCount > 0 Then
'            If RateFor = "L" Then   'Plate
'                PlateRate = Val(.Fields(Choose(Combo2.ListIndex + 1, "DeepatchPlateRate1", "PSPlateRate1", "WipeonPlateRate1", "CTPPlateRate1")).Value)
'            ElseIf RateFor = "W" Then   'Wastage
'                If Combo3.ListIndex = 0 Then    'F&B
'                    If MhRealInput19.Value > 0 Then PaperWastageRate = PaperWastageRate + Val(.Fields("PaperWastageRate" & Trim(MhRealInput19.Text)).Value): PaperWastageMin = PaperWastageMin + Val(.Fields("PaperWastageMin" & Trim(MhRealInput19.Text)).Value)
'                    If MhRealInput20.Value > 0 Then PaperWastageRate = PaperWastageRate + Val(.Fields("PaperWastageRate" & Trim(MhRealInput20.Text)).Value): PaperWastageMin = PaperWastageMin + Val(.Fields("PaperWastageMin" & Trim(MhRealInput20.Text)).Value)
'                Else    'W&T
'                    If MhRealInput19.Value > 0 And MhRealInput20.Value > 0 Then
'                        If MhRealInput19.Value > MhRealInput20.Value Then PaperWastageRate = Val(.Fields("PaperWastageRate" & Trim(MhRealInput19.Text)).Value): PaperWastageMin = PaperWastageMin + Val(.Fields("PaperWastageMin" & Trim(MhRealInput19.Text)).Value)
'                        If MhRealInput20.Value > MhRealInput19.Value Then PaperWastageRate = Val(.Fields("PaperWastageRate" & Trim(MhRealInput20.Text)).Value): PaperWastageMin = PaperWastageMin + Val(.Fields("PaperWastageMin" & Trim(MhRealInput20.Text)).Value)
'                    End If
'                End If
'            Else    'All
'                PlateRate = Val(.Fields(Choose(Combo2.ListIndex + 1, "DeepatchPlateRate1", "PSPlateRate1", "WipeonPlateRate1", "CTPPlateRate1")).Value)
'                If MhRealInput19.Value > 0 Then PrintRateFront = Val(.Fields("PrintRate" & Trim(MhRealInput19.Text)).Value)
'                If MhRealInput20.Value > 0 Then PrintRateBack = Val(.Fields("PrintRate" & Trim(MhRealInput20.Text)).Value)
'                If Combo3.ListIndex = 0 Then
'                    If MhRealInput19.Value > 0 Then PaperWastageRate = PaperWastageRate + Val(.Fields("PaperWastageRate" & Trim(MhRealInput19.Text)).Value): PaperWastageMin = PaperWastageMin + Val(.Fields("PaperWastageMin" & Trim(MhRealInput19.Text)).Value)
'                    If MhRealInput20.Value > 0 Then PaperWastageRate = PaperWastageRate + Val(.Fields("PaperWastageRate" & Trim(MhRealInput20.Text)).Value): PaperWastageMin = PaperWastageMin + Val(.Fields("PaperWastageMin" & Trim(MhRealInput20.Text)).Value)
'                Else
'                    If MhRealInput19.Value > 0 And MhRealInput20.Value > 0 Then
'                        If MhRealInput19.Value > MhRealInput20.Value Then PaperWastageRate = Val(.Fields("PaperWastageRate" & Trim(MhRealInput19.Text)).Value): PaperWastageMin = PaperWastageMin + Val(.Fields("PaperWastageMin" & Trim(MhRealInput19.Text)).Value)
'                        If MhRealInput20.Value > MhRealInput19.Value Then PaperWastageRate = Val(.Fields("PaperWastageRate" & Trim(MhRealInput20.Text)).Value): PaperWastageMin = PaperWastageMin + Val(.Fields("PaperWastageMin" & Trim(MhRealInput20.Text)).Value)
'                    End If
'                End If
'            End If
'        End If
'        If RateFor = "L" Then
'            If MhRealInput4.Value <> PlateRate Then
'                If MhRealInput4.Value = 0 Then
'                    MhRealInput4.Value = PlateRate
'                ElseIf MhRealInput4.Value <> PlateRate Then
'                    If MsgBox("Variation in Current (" & Trim(MhRealInput4.Value) & ") and Master (" & Trim(PlateRate) & ") Plate Rate !!! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput4.Value = PlateRate
'                End If
'            End If
'        ElseIf RateFor = "W" Then
'            If MhRealInput11.Value = 0 Then
'                MhRealInput11.Value = PaperWastageRate
'            ElseIf MhRealInput11.Value <> PaperWastageRate Then
'                If MsgBox("Variation in Current (" & Trim(MhRealInput11.Value) & ") and Master (" & Trim(PaperWastageRate) & ") Paper Wastage Rate !!! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput11.Value = PaperWastageRate
'            End If
'            If MhRealInput23.Value = 0 Then
'                MhRealInput23.Value = PaperWastageMin
'            ElseIf MhRealInput23.Value <> PaperWastageMin Then
'                If MsgBox("Variation in Current (" & Trim(MhRealInput23.Value) & ") and Master (" & Trim(PaperWastageMin) & ") Paper Wastage Minimum !!! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput23.Value = PaperWastageMin
'            End If
'        Else
'            If MhRealInput5.Value = 0 Then
'                MhRealInput5.Value = PrintRateFront
'            ElseIf MhRealInput5.Value <> PrintRateFront Then
'                If MsgBox("Variation in Current (" & Trim(MhRealInput5.Value) & ") and Master (" & Trim(PrintRateFront) & ") Print Rate [Front] !!! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput5.Value = PrintRateFront
'            End If
'            If MhRealInput8.Value = 0 Then
'                MhRealInput8.Value = PrintRateBack
'            ElseIf MhRealInput8.Value <> PrintRateBack Then
'                If MsgBox("Variation in Current (" & Trim(MhRealInput8.Value) & ") and Master (" & Trim(PrintRateBack) & ") Print Rate [Back] !!! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput8.Value = PrintRateBack
'            End If
'            If MhRealInput4.Value <> PlateRate Then
'                If MhRealInput4.Value = 0 Then
'                    MhRealInput4.Value = PlateRate
'                ElseIf MhRealInput4.Value <> PlateRate Then
'                    If MsgBox("Variation in Current (" & Trim(MhRealInput4.Value) & ") and Master (" & Trim(PlateRate) & ") Plate Rate !!! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput4.Value = PlateRate
'                End If
'            End If
'            If MhRealInput11.Value = 0 Then
'                MhRealInput11.Value = PaperWastageRate
'            ElseIf MhRealInput11.Value <> PaperWastageRate Then
'                If MsgBox("Variation in Current (" & Trim(MhRealInput11.Value) & ") and Master (" & Trim(PaperWastageRate) & ") Paper Wastage Rate !!! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput11.Value = PaperWastageRate
'            End If
'            If MhRealInput23.Value = 0 Then
'                MhRealInput23.Value = PaperWastageMin
'            ElseIf MhRealInput23.Value <> PaperWastageMin Then
'                If MsgBox("Variation in Current (" & Trim(MhRealInput23.Value) & ") and Master (" & Trim(PaperWastageMin) & ") Paper Wastage Minimum !!! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput23.Value = PaperWastageMin
'            End If
'        End If
'        If RateFor = "L" Then
'            Call CalculatePlateAmount
'        ElseIf RateFor = "W" Then
'            Call CalculateConsumption
'        Else
'            Call CalculatePrintAmount: Call CalculatePlateAmount: Call CalculateConsumption
'        End If
'    End With
'Exit Sub
'ErrorHandler:
'    DisplayError ("Failed to Fetch Printer Rates")
'End Sub
Private Sub CalculatePrintAmount()
    Dim TaxableAmt As Double
    TaxableAmt = IIf(MhRealInput6.Value < 1000, 1, MhRealInput6.Value / 1000) * MhRealInput5.Value + IIf(MhRealInput7.Value < 1000 And MhRealInput7.Value > 0, 1, MhRealInput7.Value / 1000) * MhRealInput8.Value
    MhRealInput36.Value = TaxableAmt
    TaxableAmt = TaxableAmt + MhRealInput9.Value
    MhRealInput17.Value = TaxableAmt * MhRealInput18.Value / 100
    MhRealInput25.Value = Round(TaxableAmt + MhRealInput17.Value, 0)
End Sub
Private Sub CalculatePlateAmount()
    Dim TaxableAmt As Double
    TaxableAmt = MhRealInput3.Value * MhRealInput4.Value
    MhRealInput37.Value = TaxableAmt
    TaxableAmt = TaxableAmt + MhRealInput26.Value
    MhRealInput22.Value = TaxableAmt * MhRealInput21.Value / 100
    MhRealInput27.Value = Round(TaxableAmt + MhRealInput22.Value, 0)
End Sub
Private Sub CalculatePaperAmount()
    Dim TaxableAmt As Double
    If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
    rstPaperList.Find "[Code] = '" & PaperCode & "'"
    If Not rstPaperList.EOF Then
        If Val(SPU) = 0 Then SPU = 500
        TaxableAmt = (MhRealInput31.Value * Wt * ((Int(MhRealInput13.Value) * SPU + (MhRealInput13.Value - Int(MhRealInput13.Value)) * 1000) / SPU))    'Paper Amount without Adjustment before Tax
        MhRealInput38.Value = TaxableAmt
        TaxableAmt = TaxableAmt + MhRealInput33.Value   'Paper Amount with Adjustment before Tax
        MhRealInput35.Value = TaxableAmt * MhRealInput34.Value / 100    'Tax
        MhRealInput32.Value = Round(TaxableAmt + MhRealInput35.Value, 0)    'Paper Amount after Tax
    End If
End Sub
Private Sub CalculateConsumption()
    Dim Consumption As Long, MinWastage As Long, MaxWastage As Long, i As Integer, FinalWastage As Long, Qty As Variant, Wastage As Long
    With fpSpread1
    MaxPrintingQuantity = 0: Qty = 0: FinalWastage = 0: MinWastage = 0: MaxWastage = 0
        
        If MhRealInput12.Value > 0 Then 'Printable Forms/Paper Sheet
            If Val(SPU) = 0 Then SPU = 500
            Consumption = MhRealInput2.Value    'Consumption (in Printing Size Sheets)
        For i = 1 To .DataRowCnt
            If Combo4.ListIndex = 1 Then 'Indiviusal Calculations Wastage
            '.GetText 4, i, Qty: MinWastage = ((Val(Qty) * MhRealInput11) / 100): If ((Val(Qty) * MhRealInput11) / 100) < MhRealInput23.Value Then MinWastage = MhRealInput23.Value: If MhRealInput23.Value = 0 Then MinWastage = ((Val(Qty) * MhRealInput11) / 100)
            '.GetText 4, i, Qty: MaxWastage = ((Val(Qty) * MhRealInput11) / 100): If ((Val(Qty) * MhRealInput11) / 100) > MhRealInput29.Value Then MaxWastage = MhRealInput29.Value: If MhRealInput29.Value = 0 Then MaxWastage = ((Val(Qty) * MhRealInput11) / 100)
            '.GetText 4, i, Qty: FinalWastage = FinalWastage + MinWastage 'IIf(MhRealInput23.Value > ((Val(Qty) * MhRealInput11) / 100), MhRealInput23.Value, IIf(MhRealInput23.Value < ((Val(Qty) * MhRealInput11) / 100), MinWastage, IIf(MhRealInput29.Value > ((Val(Qty) * MhRealInput11) / 100), MaxWastage, MhRealInput29.Value)))
            .GetText 4, i, Qty: Wastage = 0: If ((Val(Qty) * IIf(MhRealInput11 = 0, 0.01, MhRealInput11)) / 100) > Wastage Then Wastage = ((Val(Qty) * MhRealInput11) / 100): If ((Val(Qty) * IIf(MhRealInput11 = 0, 0, MhRealInput11)) / 100) < MhRealInput23.Value Then Wastage = MhRealInput23.Value: If MhRealInput23.Value = 0 Then Wastage = ((Val(Qty) * MhRealInput11) / 100)
            .GetText 4, i, Qty: If Wastage > MhRealInput29.Value Then Wastage = MhRealInput29.Value: If Wastage < MhRealInput29.Value Then Wastage = Wastage: If MhRealInput29.Value = 0 Then Wastage = Wastage
            .GetText 4, i, Qty: FinalWastage = FinalWastage + Wastage
            Else                                   'Single Calculations Wastage
            .GetText 4, i, Qty: If ((Val(Qty) * IIf(MhRealInput11 = 0, 0.01, MhRealInput11)) / 100) > FinalWastage Then FinalWastage = ((Val(Qty) * MhRealInput11) / 100): If ((Val(Qty) * IIf(MhRealInput11 = 0, 0, MhRealInput11)) / 100) < MhRealInput23.Value Then FinalWastage = MhRealInput23.Value: If MhRealInput23.Value = 0 Then FinalWastage = ((Val(Qty) * MhRealInput11) / 100)
            .GetText 4, i, Qty: If FinalWastage > MhRealInput29.Value Then FinalWastage = MhRealInput29.Value: If FinalWastage < MhRealInput29.Value Then FinalWastage = FinalWastage: If MhRealInput29.Value = 0 Then FinalWastage = FinalWastage
            End If
        Next
            Consumption = Consumption / MhRealInput12.Value 'Consumption (in Paper Size Sheets)
            MhRealInput28.Value = Int(Val(Consumption) / SPU) + ((Val(Consumption) Mod SPU) / 1000)
            MhRealInput30.Value = Int(Val(FinalWastage) / SPU) + ((Val(FinalWastage) Mod SPU) / 1000)
            Consumption = Consumption + FinalWastage  'Consumption With FinalWastage (in Printing Size Sheets)
            MhRealInput13.Value = Int(Val(Consumption) / SPU) + ((Val(Consumption) Mod SPU) / 1000)
            
'            Wastage = (Consumption * MhRealInput11.Value) / 100     'Wastage (in Printing Size Sheets)
'            If MhRealInput23.Value > 0 And Wastage < MhRealInput23.Value Then FinalWastage = MhRealInput23.Value 'Comparison with Minimum Wastage
'            If MhRealInput29.Value > 0 And Wastage > MhRealInput29.Value Then FinalWastage = MhRealInput29.Value 'Comparison with Maximum Wastage
'    With rstBookPOChild0901
'        If .RecordCount > 0 Then .MoveFirst
'        Do While Not .EOF
'            i = i + 1
'            If Combo4.ListIndex = 1 And MhRealInput23.Value > 0 And Wastage / .RecordCount < MhRealInput23.Value And Wastage / .RecordCount > MhRealInput29.Value Then
'            FinalWastage = (Consumption * MhRealInput11.Value) / 100 / MhRealInput12.Value  'Wastage (in Paper Size Sheets)
'            ElseIf Combo4.ListIndex = 1 Then 'Wastage
'            FinalWastage = Wastage * .RecordCount / MhRealInput12.Value 'Wastage (in Paper Size Sheets)
'            Else
'            FinalWastage = Wastage * 1 / MhRealInput12.Value 'Wastage (in Paper Size Sheets)
'            End If
'            .MoveNext
'        Loop
'    End With
        
        End If
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    If Combo1.ListIndex < 0 Then Combo1.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo2.ListIndex < 0 Then Combo2.SetFocus: CheckMandatoryFields = True: Exit Function
    If CheckEmpty(Text9.Text, False) Then Text9.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo3.ListIndex < 0 Then Combo3.SetFocus: CheckMandatoryFields = True: Exit Function
    If CheckEmpty(Text4.Text, False) Then Text4.SetFocus: CheckMandatoryFields = True: Exit Function
    'If Val(MhRealInput31.Text) = 0 And chkPaper.Value = 0 Then MhRealInput31.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput12.Value <= 0 Then MhRealInput12.SetFocus: CheckMandatoryFields = True: Exit Function
End Function
Private Sub cmdProceed_Click()
    Dim Stock As Double, VchDate As Date
    VchDate = FrmBookPOChild09.MhDateInput1.Value
    If CheckMandatoryFields Then Exit Sub
    If Left(FrmBookPrintOrder.BookPOType, 1) <> "O" Then
        Stock = CalculatePaperBalance(IIf(chkPaper.Value, PartyCode, "000000"), PaperCode, CheckNull(rstBookPOChild09.Fields("Code").Value), "BPOT", VchDate): Stock = Fix(Val(Stock)) * Val(SPU) + Round(Val(Stock) - Fix(Val(Stock)), 3) * 1000
        If Val(SPU) = 0 Then SPU = 500
        PaperBalance = Stock - (CLng(Int(MhRealInput13.Value) * SPU) + (MhRealInput13.Value - Int(MhRealInput13.Value)) * 1000)
                If PaperBalance < 0 Then
                        If UserLevel <= 2 Then
                            If MsgBox("Stock (" & Trim(Format(CLng(Fix(0 - Abs(PaperBalance) / Val(SPU))) + ((0 - Abs(PaperBalance) Mod Val(SPU)) / 1000), "0.000")) & ") (" & Trim(Format((PaperBalance / SPU) * Wt, "0.000")) & " Kg) of the Paper - " & Trim(Text1.Text) & " )" & vbCrLf & " is going negative ! Would you like to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbNo Then Exit Sub
                        Else
                            Call DisplayError("Cann't Save !! Stock (" & Trim(Format(CLng(Fix(0 - Abs(PaperBalance) / Val(SPU))) + ((0 - Abs(PaperBalance) Mod Val(SPU)) / 1000), "0.000")) & ") (" & Trim(Format((PaperBalance / SPU) * Wt, "0.000")) & " Kg) of the Paper - " & Trim(Text1.Text) & " )" & " is going negative "): AbortPO = True: Exit Sub
                        End If
                End If
    End If
    SaveFields
    rstBookPOChild09.Update
    Call CloseForm(Me)
End Sub
Private Sub cmdCancel_Click()
    rstBookPOChild09.CancelUpdate
    Call CloseForm(Me)
End Sub
Private Function CalUps() As Integer
        If CheckEmpty(PaperCode, False) Or CheckEmpty(SizeCode, False) Then CalUps = -1: Exit Function
        Dim FL As Double, FR As Double, PL As Double, PW As Double
        rstPaperList.MoveFirst
        rstPaperList.Find "[Code]='" & PaperCode & "'"
        FL = Val(Left(Text4.Text, InStr(1, Text4.Text, "X") - 1)): FR = Val(Mid(Text4.Text, InStr(1, Text4.Text, "X") + 1)) 'Printing Size Left & Right
        PL = IIf(rstPaperList.Fields("Form").Value = "R", Val(CutOffSize) / 25.4, Val(rstPaperList.Fields("inLength").Value)): PW = Val(rstPaperList.Fields("inWidth").Value) 'Paper Area Length & Width
        If Abs(FL - PL) <= 1 Then PL = FL
        If Abs(FR - PL) <= 1 Then PL = FR
        If Abs(FL - PW) <= 1 Then PW = FL
        If Abs(FR - PW) <= 1 Then PW = FR
        CalUps = CalcUps(PL * PW, FL * FR)
End Function
Private Sub LoadMasterList(Optional ByVal LoadSelected As Boolean)
    If rstPaperList.State = adStateOpen Then rstPaperList.Close
    If LoadSelected Then
        rstPaperList.Open "SELECT * FROM (SELECT LTRIM(P.Name)+' (UOM : '+LTRIM(C.Name)+'='+LTRIM(C.Value1)+')' As Col0,FORMAT(dbo.ufnGetPaperStock('" & IIf(chkPaper.Value, PartyCode, "000000") & "',P.Code,'PO','" & CheckNull(rstBookPOChild09.Fields("Code").Value) & "','" & GetDate(MhDateInput1.Text) & "'),'#0.000') As Col1,C.Name As UOM,GSM,inWidth,inLength,P.Code,C.Value1 As SPU,[Form],[Weight/Unit] As Wt,LTRIM(Q.Name) As Quality,Grade FROM (PaperMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code) INNER JOIN GeneralMaster Q ON P.Quality=Q.Code) As Tbl WHERE CONVERT(DECIMAL(12,3),Col1)<>0 ORDER BY Col0", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstPaperList.Open "SELECT LTRIM(P.Name)+' (UOM : '+LTRIM(C.Name)+'='+LTRIM(C.Value1)+')' As Col0,FORMAT(0,'#0.000') As Col1,C.Name As UOM,GSM,inWidth,inLength,P.Code,C.Value1 As SPU,[Form],[Weight/Unit] As Wt,LTRIM(Q.Name) As Quality,Grade FROM (PaperMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code) INNER JOIN GeneralMaster Q ON P.Quality=Q.Code ORDER BY Col0", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    rstPaperList.ActiveConnection = Nothing
End Sub
