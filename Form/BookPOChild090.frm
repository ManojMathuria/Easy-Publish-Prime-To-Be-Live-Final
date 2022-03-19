VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmBookPOChild09 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Combo Items Sheet Printing Order Details"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14835
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
   ScaleHeight     =   8265
   ScaleWidth      =   14835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   14360
      Picture         =   "BookPOChild09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   14360
      Picture         =   "BookPOChild09.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   8045
      Left            =   120
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   105
      Width           =   14135
      _Version        =   65536
      _ExtentX        =   24933
      _ExtentY        =   14190
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
      Begin VB.TextBox TxtAdNar 
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
         TabIndex        =   51
         Top             =   7590
         Width           =   12330
      End
      Begin VB.TextBox Text3 
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
         Left            =   7680
         MaxLength       =   40
         TabIndex        =   4
         Top             =   645
         Width           =   6330
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput33 
         Height          =   330
         Left            =   7680
         TabIndex        =   40
         Top             =   5915
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
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
         Top             =   6755
         Width           =   1570
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
         TabIndex        =   12
         Top             =   3915
         Width           =   12330
      End
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
         Top             =   7275
         Width           =   12330
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
         Top             =   6430
         Width           =   1570
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
         TabIndex        =   28
         Top             =   5275
         Width           =   4450
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   6120
         TabIndex        =   55
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
         Picture         =   "BookPOChild09.frx":033C
         Picture         =   "BookPOChild09.frx":0358
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   56
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
         Picture         =   "BookPOChild09.frx":0374
         Picture         =   "BookPOChild09.frx":0390
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   57
         Top             =   4755
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
         Picture         =   "BookPOChild09.frx":03AC
         Picture         =   "BookPOChild09.frx":03C8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   3240
         TabIndex        =   58
         Top             =   4755
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
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
         Caption         =   " Plate Rate-F&&B"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":03E4
         Picture         =   "BookPOChild09.frx":0400
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   3240
         TabIndex        =   59
         Top             =   4440
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
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
         Caption         =   " Print Rate-F&&B"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":041C
         Picture         =   "BookPOChild09.frx":0438
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   6120
         TabIndex        =   60
         Top             =   4435
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
         Caption         =   " Adjustment"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":0454
         Picture         =   "BookPOChild09.frx":0470
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   6120
         TabIndex        =   61
         Top             =   3600
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
         Caption         =   " Plate Type-F&&B"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":048C
         Picture         =   "BookPOChild09.frx":04A8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   62
         Top             =   4435
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
         Caption         =   " Impressions"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":04C4
         Picture         =   "BookPOChild09.frx":04E0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Left            =   8640
         TabIndex        =   63
         Top             =   5275
         Width           =   990
         _Version        =   65536
         _ExtentX        =   1746
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
         Caption         =   " Ups/Sheet"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":04FC
         Picture         =   "BookPOChild09.frx":0518
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
         Height          =   330
         Left            =   11400
         TabIndex        =   64
         Top             =   5595
         Width           =   1530
         _Version        =   65536
         _ExtentX        =   2699
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
         Caption         =   " Consumption-Kgs"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":0534
         Picture         =   "BookPOChild09.frx":0550
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   65
         Top             =   6430
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
         Picture         =   "BookPOChild09.frx":056C
         Picture         =   "BookPOChild09.frx":0588
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   6120
         TabIndex        =   66
         Top             =   6430
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
         Picture         =   "BookPOChild09.frx":05A4
         Picture         =   "BookPOChild09.frx":05C0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   3240
         TabIndex        =   67
         Top             =   6430
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
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
         Picture         =   "BookPOChild09.frx":05DC
         Picture         =   "BookPOChild09.frx":05F8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   10875
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
         Caption         =   " Target Date"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":0614
         Picture         =   "BookPOChild09.frx":0630
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   69
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
         Picture         =   "BookPOChild09.frx":064C
         Picture         =   "BookPOChild09.frx":0668
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   70
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
         Picture         =   "BookPOChild09.frx":0684
         Picture         =   "BookPOChild09.frx":06A0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   71
         Top             =   7275
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
         Picture         =   "BookPOChild09.frx":06BC
         Picture         =   "BookPOChild09.frx":06D8
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   1680
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Front"
         Top             =   4435
         Width           =   1570
         _Version        =   65536
         _ExtentX        =   2769
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":06F4
         Caption         =   "BookPOChild09.frx":0714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0780
         Keys            =   "BookPOChild09.frx":079E
         Spin            =   "BookPOChild09.frx":07E8
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   4560
         TabIndex        =   14
         ToolTipText     =   "Front"
         Top             =   4435
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0810
         Caption         =   "BookPOChild09.frx":0830
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":089C
         Keys            =   "BookPOChild09.frx":08BA
         Spin            =   "BookPOChild09.frx":0904
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   4560
         TabIndex        =   22
         Top             =   4755
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":092C
         Caption         =   "BookPOChild09.frx":094C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":09B8
         Keys            =   "BookPOChild09.frx":09D6
         Spin            =   "BookPOChild09.frx":0A20
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   330
         Left            =   7680
         TabIndex        =   16
         Top             =   4435
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0A48
         Caption         =   "BookPOChild09.frx":0A68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0AD4
         Keys            =   "BookPOChild09.frx":0AF2
         Spin            =   "BookPOChild09.frx":0B3C
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
         Height          =   330
         Left            =   7680
         TabIndex        =   46
         Top             =   6430
         Width           =   6330
         _Version        =   65536
         _ExtentX        =   11165
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0B64
         Caption         =   "BookPOChild09.frx":0B84
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0BF0
         Keys            =   "BookPOChild09.frx":0C0E
         Spin            =   "BookPOChild09.frx":0C58
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   1680
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4755
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0C80
         Caption         =   "BookPOChild09.frx":0CA0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0D0C
         Keys            =   "BookPOChild09.frx":0D2A
         Spin            =   "BookPOChild09.frx":0D74
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   1680
         TabIndex        =   32
         ToolTipText     =   "%"
         Top             =   5595
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0D9C
         Caption         =   "BookPOChild09.frx":0DBC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0E28
         Keys            =   "BookPOChild09.frx":0E46
         Spin            =   "BookPOChild09.frx":0E90
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
         ValueVT         =   2088828933
         Value           =   4
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   10440
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4435
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0EB8
         Caption         =   "BookPOChild09.frx":0ED8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":0F44
         Keys            =   "BookPOChild09.frx":0F62
         Spin            =   "BookPOChild09.frx":0FAC
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
         Height          =   330
         Left            =   9620
         TabIndex        =   17
         Top             =   4435
         Width           =   835
         _Version        =   65536
         _ExtentX        =   1473
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":0FD4
         Caption         =   "BookPOChild09.frx":0FF4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1060
         Keys            =   "BookPOChild09.frx":107E
         Spin            =   "BookPOChild09.frx":10C8
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   12915
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   5595
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":10F0
         Caption         =   "BookPOChild09.frx":1110
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":117C
         Keys            =   "BookPOChild09.frx":119A
         Spin            =   "BookPOChild09.frx":11E4
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   7680
         TabIndex        =   1
         Top             =   105
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild09.frx":120C
         Caption         =   "BookPOChild09.frx":1324
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1390
         Keys            =   "BookPOChild09.frx":13AE
         Spin            =   "BookPOChild09.frx":140C
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
         Left            =   12435
         TabIndex        =   2
         Top             =   105
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild09.frx":1434
         Caption         =   "BookPOChild09.frx":154C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":15B8
         Keys            =   "BookPOChild09.frx":15D6
         Spin            =   "BookPOChild09.frx":1634
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
         Left            =   4560
         TabIndex        =   45
         Top             =   6430
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild09.frx":165C
         Caption         =   "BookPOChild09.frx":1774
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":17E0
         Keys            =   "BookPOChild09.frx":17FE
         Spin            =   "BookPOChild09.frx":185C
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
         TabIndex        =   72
         Top             =   3600
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
         Caption         =   " Plate-F&&B"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":1884
         Picture         =   "BookPOChild09.frx":18A0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel30 
         Height          =   330
         Left            =   6120
         TabIndex        =   73
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
         Picture         =   "BookPOChild09.frx":18BC
         Picture         =   "BookPOChild09.frx":18D8
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
         Height          =   330
         Left            =   10440
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4755
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":18F4
         Caption         =   "BookPOChild09.frx":1914
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1980
         Keys            =   "BookPOChild09.frx":199E
         Spin            =   "BookPOChild09.frx":19E8
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
         Height          =   330
         Left            =   9620
         TabIndex        =   25
         Top             =   4755
         Width           =   835
         _Version        =   65536
         _ExtentX        =   1473
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":1A10
         Caption         =   "BookPOChild09.frx":1A30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1A9C
         Keys            =   "BookPOChild09.frx":1ABA
         Spin            =   "BookPOChild09.frx":1B04
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput23 
         Height          =   330
         Left            =   4800
         TabIndex        =   34
         ToolTipText     =   "Minimum"
         Top             =   5595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":1B2C
         Caption         =   "BookPOChild09.frx":1B4C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1BB8
         Keys            =   "BookPOChild09.frx":1BD6
         Spin            =   "BookPOChild09.frx":1C20
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel33 
         Height          =   330
         Left            =   120
         TabIndex        =   74
         Top             =   3915
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
         Picture         =   "BookPOChild09.frx":1C48
         Picture         =   "BookPOChild09.frx":1C64
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel34 
         Height          =   330
         Left            =   120
         TabIndex        =   75
         Top             =   6755
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
         Picture         =   "BookPOChild09.frx":1C80
         Picture         =   "BookPOChild09.frx":1C9C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel35 
         Height          =   330
         Left            =   6120
         TabIndex        =   76
         Top             =   6755
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
         Picture         =   "BookPOChild09.frx":1CB8
         Picture         =   "BookPOChild09.frx":1CD4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel36 
         Height          =   330
         Left            =   3240
         TabIndex        =   77
         Top             =   6755
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
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
         Picture         =   "BookPOChild09.frx":1CF0
         Picture         =   "BookPOChild09.frx":1D0C
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
         Height          =   330
         Left            =   7680
         TabIndex        =   49
         Top             =   6755
         Width           =   6330
         _Version        =   65536
         _ExtentX        =   11165
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":1D28
         Caption         =   "BookPOChild09.frx":1D48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1DB4
         Keys            =   "BookPOChild09.frx":1DD2
         Spin            =   "BookPOChild09.frx":1E1C
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput4 
         Height          =   330
         Left            =   4560
         TabIndex        =   48
         Top             =   6755
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild09.frx":1E44
         Caption         =   "BookPOChild09.frx":1F5C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":1FC8
         Keys            =   "BookPOChild09.frx":1FE6
         Spin            =   "BookPOChild09.frx":2044
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
         Width           =   13890
         _Version        =   524288
         _ExtentX        =   24500
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
         SpreadDesigner  =   "BookPOChild09.frx":206C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel31 
         Height          =   330
         Left            =   11400
         TabIndex        =   78
         Top             =   4435
         Width           =   1530
         _Version        =   65536
         _ExtentX        =   2699
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
         Picture         =   "BookPOChild09.frx":2946
         Picture         =   "BookPOChild09.frx":2962
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
         Height          =   330
         Left            =   12915
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   4435
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":297E
         Caption         =   "BookPOChild09.frx":299E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":2A0A
         Keys            =   "BookPOChild09.frx":2A28
         Spin            =   "BookPOChild09.frx":2A72
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel37 
         Height          =   330
         Left            =   11400
         TabIndex        =   79
         Top             =   4755
         Width           =   1530
         _Version        =   65536
         _ExtentX        =   2699
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
         Picture         =   "BookPOChild09.frx":2A9A
         Picture         =   "BookPOChild09.frx":2AB6
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
         Height          =   330
         Left            =   12915
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   4755
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":2AD2
         Caption         =   "BookPOChild09.frx":2AF2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":2B5E
         Keys            =   "BookPOChild09.frx":2B7C
         Spin            =   "BookPOChild09.frx":2BC6
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
         Height          =   330
         Left            =   6120
         TabIndex        =   80
         Top             =   4755
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
         Caption         =   " Adjustment"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":2BEE
         Picture         =   "BookPOChild09.frx":2C0A
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput28 
         Height          =   330
         Left            =   10320
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   5595
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":2C26
         Caption         =   "BookPOChild09.frx":2C46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":2CB2
         Keys            =   "BookPOChild09.frx":2CD0
         Spin            =   "BookPOChild09.frx":2D1A
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput29 
         Height          =   330
         Left            =   5520
         TabIndex        =   35
         ToolTipText     =   "Maximum"
         Top             =   5595
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":2D42
         Caption         =   "BookPOChild09.frx":2D62
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":2DCE
         Keys            =   "BookPOChild09.frx":2DEC
         Spin            =   "BookPOChild09.frx":2E36
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput30 
         Height          =   330
         Left            =   7680
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Final"
         Top             =   5595
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":2E5E
         Caption         =   "BookPOChild09.frx":2E7E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":2EEA
         Keys            =   "BookPOChild09.frx":2F08
         Spin            =   "BookPOChild09.frx":2F52
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   2460
         TabIndex        =   21
         ToolTipText     =   "Back"
         Top             =   4755
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":2F7A
         Caption         =   "BookPOChild09.frx":2F9A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":3006
         Keys            =   "BookPOChild09.frx":3024
         Spin            =   "BookPOChild09.frx":306E
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
         Height          =   330
         Left            =   5280
         TabIndex        =   15
         ToolTipText     =   "Back"
         Top             =   4435
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":3096
         Caption         =   "BookPOChild09.frx":30B6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":3122
         Keys            =   "BookPOChild09.frx":3140
         Spin            =   "BookPOChild09.frx":318A
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   81
         Top             =   5915
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
         Caption         =   " Paper Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":31B2
         Picture         =   "BookPOChild09.frx":31CE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   11400
         TabIndex        =   82
         Top             =   5915
         Width           =   1530
         _Version        =   65536
         _ExtentX        =   2699
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
         Picture         =   "BookPOChild09.frx":31EA
         Picture         =   "BookPOChild09.frx":3206
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput31 
         Height          =   330
         Left            =   1680
         TabIndex        =   39
         ToolTipText     =   "Front"
         Top             =   5915
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":3222
         Caption         =   "BookPOChild09.frx":3242
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":32AE
         Keys            =   "BookPOChild09.frx":32CC
         Spin            =   "BookPOChild09.frx":3316
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput32 
         Height          =   330
         Left            =   12915
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   5915
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":333E
         Caption         =   "BookPOChild09.frx":335E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":33CA
         Keys            =   "BookPOChild09.frx":33E8
         Spin            =   "BookPOChild09.frx":3432
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
         Height          =   330
         Left            =   9600
         TabIndex        =   30
         Top             =   5280
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":345A
         Caption         =   "BookPOChild09.frx":347A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":34E6
         Keys            =   "BookPOChild09.frx":3504
         Spin            =   "BookPOChild09.frx":354E
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   8640
         TabIndex        =   83
         Top             =   5915
         Width           =   990
         _Version        =   65536
         _ExtentX        =   1746
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
         Caption         =   " GST"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":3576
         Picture         =   "BookPOChild09.frx":3592
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput35 
         Height          =   330
         Left            =   10320
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   5910
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":35AE
         Caption         =   "BookPOChild09.frx":35CE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":363A
         Keys            =   "BookPOChild09.frx":3658
         Spin            =   "BookPOChild09.frx":36A2
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput34 
         Height          =   330
         Left            =   9600
         TabIndex        =   41
         Top             =   5910
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":36CA
         Caption         =   "BookPOChild09.frx":36EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":3756
         Keys            =   "BookPOChild09.frx":3774
         Spin            =   "BookPOChild09.frx":37BE
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   6120
         TabIndex        =   84
         Top             =   5915
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
         Caption         =   " Adjustment"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":37E6
         Picture         =   "BookPOChild09.frx":3802
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
         Height          =   330
         Left            =   8640
         TabIndex        =   85
         Top             =   4435
         Width           =   990
         _Version        =   65536
         _ExtentX        =   1746
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
         Caption         =   " GST"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":381E
         Picture         =   "BookPOChild09.frx":383A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         DataField       =   "C"
         Height          =   330
         Left            =   8640
         TabIndex        =   86
         Top             =   4755
         Width           =   990
         _Version        =   65536
         _ExtentX        =   1746
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
         Caption         =   " GST"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":3856
         Picture         =   "BookPOChild09.frx":3872
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   300
         Left            =   120
         TabIndex        =   87
         Top             =   3100
         Width           =   13890
         _Version        =   65536
         _ExtentX        =   24500
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
         Picture         =   "BookPOChild09.frx":388E
         Picture         =   "BookPOChild09.frx":38AA
         Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
            Height          =   300
            Left            =   7540
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   0
            Width           =   997
            _Version        =   65536
            _ExtentX        =   1759
            _ExtentY        =   529
            Calculator      =   "BookPOChild09.frx":38C6
            Caption         =   "BookPOChild09.frx":38E6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild09.frx":3952
            Keys            =   "BookPOChild09.frx":3970
            Spin            =   "BookPOChild09.frx":39BA
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
            ReadOnly        =   1
            Separator       =   ""
            ShowContextMenu =   1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
            Height          =   300
            Left            =   9490
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   0
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   529
            Calculator      =   "BookPOChild09.frx":39E2
            Caption         =   "BookPOChild09.frx":3A02
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild09.frx":3A6E
            Keys            =   "BookPOChild09.frx":3A8C
            Spin            =   "BookPOChild09.frx":3AD6
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
            ReadOnly        =   1
            Separator       =   ""
            ShowContextMenu =   1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
            Height          =   300
            Left            =   10710
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   0
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1759
            _ExtentY        =   529
            Calculator      =   "BookPOChild09.frx":3AFE
            Caption         =   "BookPOChild09.frx":3B1E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild09.frx":3B8A
            Keys            =   "BookPOChild09.frx":3BA8
            Spin            =   "BookPOChild09.frx":3BF2
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
            ReadOnly        =   1
            Separator       =   ""
            ShowContextMenu =   1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
            Height          =   300
            Left            =   11685
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   0
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1746
            _ExtentY        =   529
            Calculator      =   "BookPOChild09.frx":3C1A
            Caption         =   "BookPOChild09.frx":3C3A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild09.frx":3CA6
            Keys            =   "BookPOChild09.frx":3CC4
            Spin            =   "BookPOChild09.frx":3D0E
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
            ReadOnly        =   1
            Separator       =   ""
            ShowContextMenu =   1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
            Height          =   300
            Left            =   12660
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   0
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1746
            _ExtentY        =   529
            Calculator      =   "BookPOChild09.frx":3D36
            Caption         =   "BookPOChild09.frx":3D56
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "BookPOChild09.frx":3DC2
            Keys            =   "BookPOChild09.frx":3DE0
            Spin            =   "BookPOChild09.frx":3E2A
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
            ReadOnly        =   1
            Separator       =   ""
            ShowContextMenu =   1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
         Height          =   330
         Left            =   7680
         TabIndex        =   24
         Top             =   4755
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":3E52
         Caption         =   "BookPOChild09.frx":3E72
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":3EDE
         Keys            =   "BookPOChild09.frx":3EFC
         Spin            =   "BookPOChild09.frx":3F46
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   120
         TabIndex        =   93
         Top             =   5275
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
         Picture         =   "BookPOChild09.frx":3F6E
         Picture         =   "BookPOChild09.frx":3F8A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel50 
         Height          =   330
         Left            =   11400
         TabIndex        =   94
         Top             =   5280
         Width           =   1530
         _Version        =   65536
         _ExtentX        =   2699
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
         Picture         =   "BookPOChild09.frx":3FA6
         Picture         =   "BookPOChild09.frx":3FC2
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Left            =   6120
         TabIndex        =   95
         Top             =   5595
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
         Caption         =   " Paper Wastage"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":3FDE
         Picture         =   "BookPOChild09.frx":3FFA
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel40 
         Height          =   330
         Left            =   6120
         TabIndex        =   96
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
         Caption         =   " Ref No."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":4016
         Picture         =   "BookPOChild09.frx":4032
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel57 
         Height          =   330
         Left            =   6120
         TabIndex        =   97
         Top             =   5275
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
         Caption         =   " Cut Off Size"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":404E
         Picture         =   "BookPOChild09.frx":406A
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput52 
         Height          =   330
         Left            =   7680
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   5275
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":4086
         Caption         =   "BookPOChild09.frx":40A6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":4112
         Keys            =   "BookPOChild09.frx":4130
         Spin            =   "BookPOChild09.frx":417A
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
         ForeColor       =   255
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
         Height          =   330
         Left            =   12915
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   5280
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         _StockProps     =   77
         TintColor       =   16711935
         Alignment       =   0
         AutoSize        =   0   'False
         BevelSize       =   0
         BevelStyle      =   0
         BorderColor     =   -2147483642
         BorderStyle     =   1
         FillColor       =   16777215
         FontStyle       =   0
         FontTransparent =   0   'False
         LightColor      =   -2147483643
         ShadowColor     =   -2147483632
         TextColor       =   -2147483640
         WallPaper       =   0
         NoPrefix        =   0   'False
         FormatString    =   ""
         Caption         =   ""
         Picture         =   "BookPOChild09.frx":41A2
         Begin VB.CheckBox chkPaper 
            Caption         =   "Check1"
            Height          =   210
            Left            =   450
            TabIndex        =   31
            Top             =   80
            Width           =   210
         End
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Left            =   120
         TabIndex        =   99
         Top             =   5595
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
         Caption         =   " Wastage %-F&&B"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":41BE
         Picture         =   "BookPOChild09.frx":41DA
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
         Height          =   330
         Left            =   2460
         TabIndex        =   33
         ToolTipText     =   "%"
         Top             =   5595
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":41F6
         Caption         =   "BookPOChild09.frx":4216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":4282
         Keys            =   "BookPOChild09.frx":42A0
         Spin            =   "BookPOChild09.frx":42EA
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel49 
         Height          =   330
         Left            =   3240
         TabIndex        =   100
         Top             =   5595
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
         Caption         =   " Wastage Sht-F&&B"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":4312
         Picture         =   "BookPOChild09.frx":432E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   8640
         TabIndex        =   101
         Top             =   5595
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Caption         =   " Consumption-UOM"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":434A
         Picture         =   "BookPOChild09.frx":4366
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput36 
         Height          =   330
         Left            =   5280
         TabIndex        =   23
         ToolTipText     =   "Back"
         Top             =   4755
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild09.frx":4382
         Caption         =   "BookPOChild09.frx":43A2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild09.frx":440E
         Keys            =   "BookPOChild09.frx":442C
         Spin            =   "BookPOChild09.frx":4476
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Left            =   120
         TabIndex        =   102
         Top             =   7590
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
         Caption         =   " Adj.Remarks"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild09.frx":449E
         Picture         =   "BookPOChild09.frx":44BA
      End
      Begin MSForms.ComboBox Combo22 
         Height          =   330
         Left            =   10760
         TabIndex        =   11
         Top             =   3600
         Width           =   3250
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "5733;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox Combo11 
         Height          =   330
         Left            =   3960
         TabIndex        =   9
         Top             =   3600
         Width           =   2175
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3836;582"
         ListRows        =   4
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
         TabIndex        =   8
         Top             =   3600
         Width           =   2295
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4048;582"
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
         X2              =   14090
         Y1              =   3490
         Y2              =   3490
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   14090
         Y1              =   4335
         Y2              =   4335
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   14090
         Y1              =   1375
         Y2              =   1375
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   330
         Left            =   7680
         TabIndex        =   6
         Top             =   960
         Width           =   6330
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "11165;582"
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
         X2              =   14090
         Y1              =   7175
         Y2              =   7175
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   14090
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   14090
         Y1              =   6330
         Y2              =   6330
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   7650
         TabIndex        =   10
         Top             =   3600
         Width           =   3120
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "5503;582"
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
         X2              =   14090
         Y1              =   5175
         Y2              =   5175
      End
   End
End
Attribute VB_Name = "FrmBookPOChild09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchCode As String, VchType As String, PartyCode As String, RoundOffQty As Boolean, rstBookPOChild09 As New ADODB.Recordset, rstBookPOChild0901 As New ADODB.Recordset
Dim rstPlateMakerList As New ADODB.Recordset, rstSizeList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstPaperList As New ADODB.Recordset, rstFetchRate As New ADODB.Recordset, rstOrderList As New ADODB.Recordset
Dim PlateMakerCode As String, PaperCode As String, SizeCode As String
Dim PaperBalance As Long, SPU As Long, Wt As Double, inWidth As Double, GSM As Double, EditMode As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    DisableCloseButton Me
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text) 'Order No.
    Text7.Text = Trim(FrmBookPrintOrder.Text6.Text) 'Party Name
    Combo1.AddItem "Old", 0: Combo1.AddItem "New", 1: Combo1.AddItem "Revised", 2
    Combo11.AddItem "Old", 0: Combo11.AddItem "New", 1: Combo11.AddItem "Revised", 2
    Combo2.AddItem "Deepatch", 0: Combo2.AddItem "PS", 1: Combo2.AddItem "Wipeon", 2: Combo2.AddItem "CTP", 3
    Combo22.AddItem "Deepatch", 0: Combo22.AddItem "PS", 1: Combo22.AddItem "Wipeon", 2: Combo22.AddItem "CTP", 3
    Combo3.AddItem "F&B", 0: Combo3.AddItem "W&T", 1
    LoadMasterList
    ClearFields
    LoadFields
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name <> "fpSpread1" Then SendKeys "{TAB}": KeyCode = 0
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
    Call CloseRecordset(rstPlateMakerList)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstFetchRate)
    Call CloseRecordset(rstOrderList)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)    'Order Date
    If Format(MhDateInput1.Value, "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(MhDateInput1.Value, "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf CheckEmpty(VchCode, False) Then
        MhDateInput3.Value = DateAdd("d", 2, MhDateInput1.Value)
    End If
End Sub
Private Sub MhDateInput3_Validate(Cancel As Boolean)    'Target Date
    If Format(GetDate(MhDateInput3.Text), "yyyymmdd") <= Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then DisplayError ("Target Date cann't be prior to Order Date"): MhDateInput3.SetFocus: Cancel = True
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1.ListIndex = -1 Then Cancel = True
    If Combo1.ListIndex = 0 Then If InStr(1, "1_3", Trim(Combo2.ListIndex)) > 0 Then MhRealInput4.Value = 0
End Sub
Private Sub Combo11_Validate(Cancel As Boolean)
    If Combo11.ListIndex = -1 Then Cancel = True
    If Combo11.ListIndex = 0 Then If InStr(1, "1_3", Trim(Combo22.ListIndex)) > 0 Then MhRealInput36.Value = 0
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)  'Front Plate Type
    If Left(VchType, 1) = "O" Then Exit Sub
    If InStr(1, "1_3", Trim(Combo2.ListIndex)) > 0 Then 'PS/CTP Plate Details
        Dim ItemCode As Variant, i As Integer, BookCode As Variant
        On Error Resume Next
        With fpSpread1
            For i = 1 To .DataRowCnt
                .GetText 8, i, BookCode
                ItemCode = ItemCode + "'" + BookCode + "',"
            Next
        End With
        ItemCode = Left(ItemCode, Len(ItemCode) - 1)
        If Not CheckEmpty(ItemCode, False) Then
            FrmPSPlateRegister.ItemCode = ItemCode
            FrmPSPlateRegister.OrderCode = IIf(CheckEmpty(VchCode, False), "999999", VchCode)
            FrmPSPlateRegister.OrderDate = GetDate(MhDateInput1.Text)
            FrmPSPlateRegister.TblSuffix = "09"
            FrmPSPlateRegister.OrderType = VchType
            FrmPSPlateRegister.PlateType = "F"
            Load FrmPSPlateRegister
            If Err.Number <> 364 Then FrmPSPlateRegister.Show vbModal
        End If
        On Error GoTo 0
    End If
End Sub
Private Sub Combo22_Validate(Cancel As Boolean)  'Back Plate Type
    If Left(VchType, 1) = "O" Then Exit Sub
    If InStr(1, "1_3", Trim(Combo22.ListIndex)) > 0 Then 'PS/CTP Plate Details
        Dim ItemCode As Variant, i As Integer, BookCode As Variant
        On Error Resume Next
        With fpSpread1
            For i = 1 To .DataRowCnt
                .GetText 8, i, BookCode
                ItemCode = ItemCode + "'" + BookCode + "',"
            Next
        End With
        ItemCode = Left(ItemCode, Len(ItemCode) - 1)
        If Not CheckEmpty(ItemCode, False) Then
            FrmPSPlateRegister.ItemCode = ItemCode
            FrmPSPlateRegister.OrderCode = IIf(CheckEmpty(VchCode, False), "999999", VchCode)
            FrmPSPlateRegister.OrderDate = GetDate(MhDateInput1.Text)
            FrmPSPlateRegister.TblSuffix = "09"
            FrmPSPlateRegister.OrderType = VchType
            FrmPSPlateRegister.PlateType = "B"
            Load FrmPSPlateRegister
            If Err.Number <> 364 Then FrmPSPlateRegister.Show vbModal
        End If
        On Error GoTo 0
    End If
End Sub
Private Sub Combo3_Validate(Cancel As Boolean)  'Imposition
    'Plates
    MhRealInput3.Value = IIf(Combo3.ListIndex = 0, MhRealInput19.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, 0))
    MhRealInput7.Value = IIf(Combo3.ListIndex = 0, MhRealInput20.Value, IIf(MhRealInput20.Value > MhRealInput19.Value, MhRealInput20.Value, 0))
    'Plate Rate
    MhRealInput4.Value = IIf(Combo3.ListIndex = 0, MhRealInput4.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput4.Value, 0))
    MhRealInput36.Value = IIf(Combo3.ListIndex = 0, MhRealInput36.Value, IIf(MhRealInput20.Value > MhRealInput19.Value, MhRealInput36.Value, 0))
    'Print Rate
    MhRealInput5.Value = IIf(Combo3.ListIndex = 0, MhRealInput5.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput5.Value, 0))
    MhRealInput8.Value = IIf(Combo3.ListIndex = 0, MhRealInput8.Value, IIf(MhRealInput20.Value > MhRealInput19.Value, MhRealInput8.Value, 0))
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "1"
        FrmGeneralMaster.MasterCode = SizeCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        SizeCode = slCode: Text4.Text = slName
        If Not CheckEmpty(SizeCode, False) Then LoadMasterList: SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        SizeCode = "": Text4.Text = ""
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)   'Size
    If CheckEmpty(Text4.Text, False) Then Cancel = True
End Sub
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01"
        FrmAccountMaster.MasterCode = PlateMakerCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        PlateMakerCode = slCode: Text9.Text = slName
        If Not CheckEmpty(PlateMakerCode, False) Then LoadMasterList: SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        PlateMakerCode = "": Text9.Text = ""
    End If
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
    If CheckEmpty(Text9.Text, False) Then Cancel = True
End Sub
Private Sub MhRealInput5_GotFocus()
    Call GetPrinterRates("P", "F")
End Sub
Private Sub MhRealInput5_Validate(Cancel As Boolean)    'Front Print Rate
    CalculatePrintAmount
End Sub
Private Sub MhRealInput8_GotFocus()
    Call GetPrinterRates("P", "B")
End Sub
Private Sub MhRealInput8_Validate(Cancel As Boolean)    'Back Print Rate
    CalculatePrintAmount
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean)    'Adjustment
    CalculatePrintAmount
End Sub
Private Sub MhRealInput18_Validate(Cancel As Boolean)   'GST%
    CalculatePrintAmount
End Sub
Private Sub MhRealInput4_GotFocus()
    Call GetPrinterRates("L", "F")
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)    'Front Plate Rate
    CalculatePlateAmount
End Sub
Private Sub MhRealInput36_GotFocus()
    Call GetPrinterRates("L", "B")
End Sub
Private Sub MhRealInput36_Validate(Cancel As Boolean)    'Back Plate Rate
    CalculatePlateAmount
End Sub
Private Sub MhRealInput26_Validate(Cancel As Boolean)   'Plate Adjustment
    CalculatePlateAmount
End Sub
Private Sub MhRealInput21_Validate(Cancel As Boolean)   'PGST%
    CalculatePlateAmount
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmPaperMaster.SL = True
        FrmPaperMaster.MasterCode = PaperCode
        Load FrmPaperMaster
        If Err.Number <> 364 Then FrmPaperMaster.Show vbModal
        On Error GoTo 0
        PaperCode = slCode: Text1.Text = slName
        If Not CheckEmpty(PaperCode, False) Then LoadMasterList: SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        PaperCode = "": Text1.Text = ""
    End If
End Sub
Private Sub Text1_Validate(Cancel As Boolean)   'Paper
    If CheckEmpty(Text1.Text, False) Then
        Cancel = True
    Else
        rstPaperList.MoveFirst
        rstPaperList.Find "[Code]='" & PaperCode & "'"
        Text1.Text = rstPaperList.Fields("Col0").Value: SPU = Val(rstPaperList.Fields("SPU").Value): Wt = Val(rstPaperList.Fields("Wt").Value): inWidth = Val(rstPaperList.Fields("inWidth").Value): GSM = Val(rstPaperList.Fields("GSM").Value)
        If rstPaperList.Fields("Form").Value = "S" Then MhRealInput52.Value = 0 Else MhRealInput52.Value = Val(rstPaperList.Fields("inLength").Value)
        CalculateConsumption
        If CheckEmpty(SizeCode, False) Then Exit Sub
        Dim FL As Double, FR As Double, PL As Double, PR As Double
        PL = Val(Left(Text4.Text, InStr(1, Text4.Text, "X") - 1)): PR = Val(Mid(Text4.Text, InStr(1, Text4.Text, "X") + 1, 5))
        FL = Val(rstPaperList.Fields("inWidth").Value): FR = Val(rstPaperList.Fields("inLength").Value)
        Call CalcUps(FL * FR, PL * PR)
        If Left(VchType, 1) <> "O" Then
            PaperBalance = CalculatePaperBalance(IIf(chkPaper.Value, PartyCode, "*00000"), PaperCode, VchCode, "BPO") 'Sheets
            MsgBox "Stock Available : (" & Trim(Format(CLng(Int(PaperBalance / SPU)) + ((PaperBalance Mod SPU) / 1000), "0.000")) & ") (" & Trim(Format((PaperBalance / SPU) * Wt, "0.000")) & " Kg)"
        End If
    End If
End Sub
Private Sub MhRealInput12_GotFocus()
    Dim FL As Double, FR As Double, PL As Double, PR As Double, Ups01 As Integer, Ups02 As Integer, Ups03 As Integer, Ups As Integer
    If CheckEmpty(PaperCode, False) Or CheckEmpty(SizeCode, False) Then Exit Sub
    rstPaperList.MoveFirst
    rstPaperList.Find "[Code]='" & PaperCode & "'"
    PL = Val(Left(Text4.Text, InStr(1, Text4.Text, "X") - 1)): PR = Val(Mid(Text4.Text, InStr(1, Text4.Text, "X") + 1, 5))
    FL = Val(rstPaperList.Fields("inWidth").Value): FR = Val(rstPaperList.Fields("inLength").Value)
    Ups01 = Int(IIf(FL > FR, FL, FR) / IIf(PL > PR, PL, PR)) * Int(IIf(FL < FR, FL, FR) / IIf(PL < PR, PL, PR)): Ups02 = Int(IIf(FL > FR, FL, FR) / IIf(PL < PR, PL, PR)) * Int(IIf(FL < FR, FL, FR) / IIf(PL > PR, PL, PR)): Ups03 = Int((FL * FR) / (PL * PR))
    Ups = IIf(Ups03 > IIf(Ups01 > Ups02, Ups01, Ups02), Ups03, IIf(Ups01 > Ups02, Ups01, Ups02))
    If Ups > 0 Then
        If MhRealInput12.Value = 0 Then
            MhRealInput12.Value = Ups
        ElseIf Ups <> MhRealInput12.Value Then
            If MsgBox("Variation in Calculated [" & Trim(Ups) & "] and Existing [" & Trim(MhRealInput12.Value) & "] Ups/Sheet ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput12.Value = Ups
        End If
    End If
End Sub
Private Sub MhRealInput12_Validate(Cancel As Boolean)   'Titles/Sheet For Calculating Paper Consumption
    CalculateConsumption
End Sub
Private Sub MhRealInput11_GotFocus()
    Call GetPrinterRates("W", "F")
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)   'Wastage Percentage - Front
    CalculateConsumption
End Sub
Private Sub MhRealInput14_GotFocus()
    Call GetPrinterRates("W", "B")
End Sub
Private Sub MhRealInput14_Validate(Cancel As Boolean)   'Wastage Percentage - Back
    CalculateConsumption
End Sub
Private Sub MhRealInput23_GotFocus()
    Call GetPrinterRates("M", "F")
End Sub
Private Sub MhRealInput23_Validate(Cancel As Boolean)   'Wastage Min - Front
    CalculateConsumption
End Sub
Private Sub MhRealInput29_GotFocus()
    Call GetPrinterRates("M", "B")
End Sub
Private Sub MhRealInput29_Validate(Cancel As Boolean)   'Wastage Min - Back
    CalculateConsumption
End Sub
Private Sub MhRealInput31_Validate(Cancel As Boolean)   'Paper Rate
    MhRealInput32.Value = MhRealInput31.Value * MhRealInput13.Value
    CalculatePaperAmount
End Sub
Private Sub MhRealInput33_Validate(Cancel As Boolean)   'Paper Adjustment
    CalculatePaperAmount
End Sub
Private Sub MhRealInput34_Validate(Cancel As Boolean)   'RGST%
    CalculatePaperAmount
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    With fpSpread1
        If Shift = vbCtrlMask And KeyCode = vbKeyD Then
            If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
                .DeleteRows .ActiveRow, 1: .SetFocus: CalculateTotal
            End If
        ElseIf Shift = 0 And KeyCode = vbKeyDelete Then
            If Not EditMode Then .SetText .ActiveCol, .ActiveRow, ""
        ElseIf KeyCode = vbKeySpace Then
            Dim Item As Variant
            If .ActiveCol = 1 Then
                .GetText 8, .ActiveRow, Item
                On Error Resume Next
                FrmBookMaster.SL = True
                FrmBookMaster.BookType = IIf(Left(VchType, 1) = "O", "F", Left(VchType, 1))
                FrmBookMaster.MasterCode = Item
                Load FrmBookMaster
                If Err.Number <> 364 Then FrmBookMaster.Show vbModal
                On Error GoTo 0
                .SetText 1, .ActiveRow, slName
                .SetText 8, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then
                    LoadMasterList
                    SendKeys "{ENTER}"
                Else
                    .SetActiveCell 1, .ActiveRow
                End If
            End If
        End If
    End With
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Qty As Variant, Ups As Variant, TotalForms As Double
    With fpSpread1
        If Col = 2 Or Col = 3 Then
            .GetText 2, Row, Qty    'Actual Quantity
            .GetText 3, Row, Ups    'Ups/Plate
            If Val(Qty) > 0 Then
                If Val(Ups) > 0 Then
                    TotalForms = Val(Qty) / Val(Ups)
                    TotalForms = TotalForms + IIf(TotalForms - Int(TotalForms) > 0, 1, 0)
                    If TotalForms > 0 Then
                        .SetText 4, Row, TotalForms    'Printing Quantity
                        If RoundOffQty Then
                            If TotalForms < 1000 Then TotalForms = 1000
                            TotalForms = IIf(Int(TotalForms / 1000) = 0, 1000, Int(TotalForms / 1000) * 1000) + IIf(TotalForms Mod 1000 <= IIf(TotalForms <= 20000, 299, 599), 0, 1000)
                        End If
                    End If
                    .GetText 5, Row, Qty    'Billing Quantity
                    If Val(Qty) = 0 Then
                        .SetText 5, Row, TotalForms
                    ElseIf Val(Qty) <> TotalForms Then
                        fpSpread1.SetActiveCell Col + 2, Row: If MsgBox("Variation in Calculated [" & Trim(TotalForms) & "] and Existing [" & Trim(Qty) & "] Impressions ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 5, Row, TotalForms
                    End If
                Else
                    .SetText 4, Row, ""
                    .SetText 5, Row, ""
                End If
                CalculateTotal
            Else
                .SetText 4, Row, ""
                .SetText 5, Row, ""
            End If
        ElseIf Col >= 5 Then
            CalculateTotal
        End If
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub cmdProceed_Click()
    Dim i As Integer, Item As Variant
    With fpSpread1
        For i = 1 To .DataRowCnt
            .SetActiveCell i, 1
            .GetText 1, i, Item
            If Not CheckEmpty(Item, False) Then If CheckDuplicateItem(i, Item) Then .SetFocus: Exit Sub
        Next
    End With
    If CheckMandatoryFields Then Exit Sub
    If Left(VchType, 1) <> "O" Then
        Dim Stock As Long
        PaperBalance = CalculatePaperBalance(IIf(chkPaper.Value, PartyCode, "*00000"), PaperCode, VchCode, "BPO") 'Sheets
        Stock = PaperBalance - (CLng(Int(MhRealInput13.Value) * SPU) + ((MhRealInput13.Value - Int(MhRealInput13.Value)) * 1000)) 'Sheets
        If Stock < 0 Then If MsgBox("Stock (" & Trim(Format(CLng(Fix(Stock / SPU)) + ((Stock Mod SPU) / 1000), "0.000")) & ") (" & Trim(Format((Stock / SPU) * Wt, "0.000")) & " Kg) of the Paper - " & Trim(Text1.Text) & vbCrLf & " is going negative ! Would you like to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbNo Then Exit Sub
    End If
    SaveFields
    FrmBookPrintOrder.Command2.Enabled = False
    Call CloseForm(Me)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Sub ClearFields()
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy") 'Order Date
    MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy") 'Target Date
    Text3.Text = ""
    Combo1.ListIndex = 0
    Combo11.ListIndex = 0
    Combo2.ListIndex = 3
    Combo22.ListIndex = 3
    Combo3.ListIndex = 0
    Text4.Text = "": SizeCode = ""
    Text9.Text = "": PlateMakerCode = PartyCode
    If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
    rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
    If Not rstPlateMakerList.EOF Then Text9.Text = rstPlateMakerList.Fields("Col0").Value
    MhRealInput6.Value = 0
    MhRealInput5.Value = 0
    MhRealInput8.Value = 0
    MhRealInput9.Value = 0
    MhRealInput18.Value = 0
    MhRealInput17.Value = 0
    MhRealInput25.Value = 0
    MhRealInput3.Value = 0
    MhRealInput7.Value = 0
    MhRealInput4.Value = 0
    MhRealInput36.Value = 0
    MhRealInput26.Value = 0
    MhRealInput21.Value = 0
    MhRealInput22.Value = 0
    MhRealInput27.Value = 0
    chkPaper.Value = 1
    Text1.Text = "": PaperCode = "": SPU = 0: Wt = 0: inWidth = 0: GSM = 0
    MhRealInput52.Value = 0
    MhRealInput12.Value = 0
    MhRealInput11.Value = 0
    MhRealInput14.Value = 0
    MhRealInput23.Value = 0
    MhRealInput29.Value = 0
    MhRealInput30.Value = 0
    MhRealInput28.Value = 0
    MhRealInput13.Value = 0
    MhRealInput31.Value = 0
    MhRealInput33.Value = 0
    MhRealInput34.Value = 0
    MhRealInput35.Value = 0
    MhRealInput32.Value = 0
    Text6.Text = ""
    Text8.Text = ""
    MhDateInput2.Value = ""
    Text10.Text = ""
    MhDateInput4.Value = ""
    MhRealInput16.Value = 0
    MhRealInput24.Value = 0
    TxtAdNar.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    MhRealInput1.Value = 0
    MhRealInput2.Value = 0
    MhRealInput15.Value = 0
    MhRealInput19.Value = 0
    MhRealInput20.Value = 0
End Sub
Private Sub LoadFields()
    With rstBookPOChild09
        If .RecordCount = 0 Then Exit Sub
        MhDateInput1.Text = Format(.Fields("OrderDate").Value, "dd-MM-yyyy")
        MhDateInput3.Text = Format(.Fields("TargetDate").Value, "dd-MM-yyyy")
        Text3.Text = .Fields("Ref").Value
        Combo1.ListIndex = IIf(.Fields("Plate").Value = "O", 0, IIf(.Fields("Plate").Value = "N", 1, 2))
        Combo11.ListIndex = IIf(.Fields("Plate").Value = "O", 0, IIf(.Fields("Plate").Value = "N", 1, 2))
        Combo2.ListIndex = Val(.Fields("PlateType").Value) - 1
        Combo22.ListIndex = Val(.Fields("PlateType").Value) - 1
        Combo3.ListIndex = IIf(.Fields("Imposition").Value = "F", 0, 1)
        SizeCode = .Fields("Size").Value
        If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
        rstSizeList.Find "[Code] = '" & SizeCode & "'"
        If Not rstSizeList.EOF Then Text4.Text = rstSizeList.Fields("Col0").Value
        PlateMakerCode = .Fields("PlateMaker").Value
        If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
        rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
        If Not rstPlateMakerList.EOF Then Text9.Text = Trim(rstPlateMakerList.Fields("Col0").Value)
        MhRealInput6.Value = .Fields("TotalForms").Value
        MhRealInput5.Value = Val(.Fields("PrintRateFront").Value)
        MhRealInput8.Value = Val(.Fields("PrintRateBack").Value)
        MhRealInput9.Value = Val(.Fields("Adjustment").Value)
        MhRealInput18.Value = Val(.Fields("GST%").Value)
        MhRealInput17.Value = Val(.Fields("GST").Value)
        MhRealInput25.Value = Val(.Fields("PrintAmount").Value)
        MhRealInput3.Value = Val(.Fields("TotalPlates").Value)
        MhRealInput7.Value = Val(.Fields("TotalPlatesBack").Value)
        MhRealInput4.Value = Val(.Fields("PlateRate").Value)
        MhRealInput36.Value = Val(.Fields("PlateRateBack").Value)
        MhRealInput26.Value = Val(.Fields("PAdjustment").Value)
        MhRealInput21.Value = Val(.Fields("PGST%").Value)
        MhRealInput22.Value = Val(.Fields("PGST").Value)
        MhRealInput27.Value = Val(.Fields("PlateAmount").Value)
        chkPaper.Value = IIf(.Fields("PaperByParty").Value, 1, 0)
        PaperCode = .Fields("Paper").Value
        If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
        rstPaperList.Find "[Code] = '" & PaperCode & "'"
        If Not rstPaperList.EOF Then Text1.Text = rstPaperList.Fields("Col0").Value: SPU = Val(rstPaperList.Fields("SPU").Value): Wt = Val(rstPaperList.Fields("Wt").Value): inWidth = Val(rstPaperList.Fields("inWidth").Value): GSM = Val(rstPaperList.Fields("GSM").Value)
        MhRealInput52.Value = Val(.Fields("CutOffSize").Value)
        MhRealInput12.Value = Val(.Fields("Ups/Sheet").Value)
        MhRealInput11.Value = Val(.Fields("PaperWastage%").Value)
        MhRealInput14.Value = Val(.Fields("PaperWastage%Back").Value)
        MhRealInput23.Value = Val(.Fields("PaperWastageMin").Value)
        MhRealInput29.Value = Val(.Fields("PaperWastageMinBack").Value)
        MhRealInput30.Value = Val(.Fields("PaperWastageFinal").Value)
        MhRealInput28.Value = Val(.Fields("PaperConsumptionOther").Value)
        MhRealInput13.Value = Val(.Fields("PaperConsumptionKg").Value)
        MhRealInput31.Value = Val(.Fields("PaperRate").Value)
        MhRealInput33.Value = Val(.Fields("RAdjustment").Value)
        MhRealInput34.Value = Val(.Fields("RGST%").Value)
        MhRealInput35.Value = Val(.Fields("RGST").Value)
        MhRealInput32.Value = Val(.Fields("PaperAmount").Value)
        Text6.Text = .Fields("Remarks").Value
        Text8.Text = .Fields("BillNo").Value
        If Not IsNull(.Fields("BillDate").Value) Then MhDateInput2.Value = .Fields("BillDate").Value
        Text10.Text = .Fields("PBillNo").Value
        If Not IsNull(.Fields("PBillDate").Value) Then MhDateInput4.Value = .Fields("PBillDate").Value
        MhRealInput16.Value = Val(.Fields("PaidAmount").Value)
        MhRealInput24.Value = Val(.Fields("PPaidAmount").Value)
        TxtAdNar.Text = .Fields("AdjustmentRemarks").Value
        Call LoadItemList
    End With
End Sub
Private Sub SaveFields()
    With rstBookPOChild09
        If .RecordCount = 0 Then .AddNew
        .Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
        .Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
        .Fields("Ref").Value = Text3.Text
        .Fields("Plate").Value = Choose(Combo1.ListIndex + 1, "O", "N", "R")
        .Fields("PlateBack").Value = Choose(Combo11.ListIndex + 1, "O", "N", "R")
        .Fields("PlateType").Value = Trim(Str(Combo2.ListIndex + 1))
        .Fields("PlateTypeBack").Value = Trim(Str(Combo22.ListIndex + 1))
        .Fields("Imposition").Value = Choose(Combo3.ListIndex + 1, "F", "W")
        .Fields("Size").Value = SizeCode
        .Fields("PlateMaker").Value = PlateMakerCode
        .Fields("TotalForms").Value = MhRealInput6.Value
        .Fields("PrintRateFront").Value = MhRealInput5.Value
        .Fields("PrintRateBack").Value = MhRealInput8.Value
        .Fields("Adjustment").Value = MhRealInput9.Value
        .Fields("GST%").Value = MhRealInput18.Value
        .Fields("GST").Value = MhRealInput17.Value
        .Fields("PrintAmount").Value = MhRealInput25.Value
        .Fields("TotalPlates").Value = MhRealInput3.Value
        .Fields("TotalPlatesBack").Value = MhRealInput7.Value
        .Fields("PlateRate").Value = MhRealInput4.Value
        .Fields("PlateRateBack").Value = MhRealInput36.Value
        .Fields("PAdjustment").Value = MhRealInput26.Value
        .Fields("PGST%").Value = MhRealInput21.Value
        .Fields("PGST").Value = MhRealInput22.Value
        .Fields("PlateAmount").Value = MhRealInput27.Value
        .Fields("PaperByParty").Value = chkPaper.Value
        .Fields("Paper").Value = PaperCode
        .Fields("CutOffSize").Value = MhRealInput52.Value
        .Fields("Ups/Sheet").Value = MhRealInput12.Value
        .Fields("PaperWastage%").Value = MhRealInput11.Value
        .Fields("PaperWastage%Back").Value = MhRealInput14.Value
        .Fields("PaperWastageMin").Value = MhRealInput23.Value
        .Fields("PaperWastageMinBack").Value = MhRealInput29.Value
        .Fields("PaperWastageFinal").Value = MhRealInput30.Value
        .Fields("PaperConsumptionSheets").Value = CLng(Int(MhRealInput13.Value) * SPU) + ((MhRealInput13.Value - Int(MhRealInput13.Value)) * 1000)
        .Fields("PaperConsumptionOther").Value = MhRealInput28.Value
        .Fields("PaperConsumptionKg").Value = MhRealInput13.Value
        .Fields("PaperRate").Value = MhRealInput31.Value
        .Fields("RAdjustment").Value = MhRealInput33.Value
        .Fields("RGST%").Value = MhRealInput34.Value
        .Fields("RGST").Value = MhRealInput35.Value
        .Fields("PaperAmount").Value = MhRealInput32.Value
        .Fields("Remarks").Value = Text6.Text
        .Fields("BillNo").Value = Text8.Text
        If Not IsDate(MhDateInput2.Text) Then .Fields("BillDate").Value = Null Else .Fields("BillDate").Value = GetDate(MhDateInput2.Text)
        .Fields("PBillNo").Value = Text10.Text
        If Not IsDate(MhDateInput4.Text) Then .Fields("PBillDate").Value = Null Else .Fields("PBillDate").Value = GetDate(MhDateInput4.Text)
        .Fields("PaidAmount").Value = MhRealInput16.Value
        .Fields("PPaidAmount").Value = MhRealInput24.Value
        .Fields("AdjustmentRemarks").Value = IIf(MhRealInput9.Value <> 0 Or MhRealInput26.Value <> 0 Or MhRealInput33.Value <> 0, TxtAdNar.Text, "")
        UpdateItemList
    End With
End Sub
Private Sub LoadItemList()
    On Error GoTo ErrHandler
    rstOrderList.Open "SELECT Book As ItemCode,M.Name As ItemName,ActualQuantity,[Ups/Plate],PrintingQuantity,BillingQuantity,FrontPrintingColor,BackPrintingColor,QuantityReceived,QuantityIssued,Status FROM BookPOChild0901 T INNER JOIN BookMaster M ON T.Book=M.Code WHERE T.Code='" & VchCode & "' ORDER BY M.Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstOrderList.ActiveConnection = Nothing
    If rstOrderList.RecordCount = 0 Then Exit Sub
    Dim i As Integer
    With rstOrderList
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
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
    Dim i As Integer, Col As Variant
    For i = 1 To fpSpread1.DataRowCnt
        With rstBookPOChild0901
            fpSpread1.GetText 1, i, Col
            If Not CheckEmpty(Col, False) Then
                .AddNew
                fpSpread1.GetText 2, i, Col: .Fields("ActualQuantity").Value = Val(Col)
                fpSpread1.GetText 3, i, Col: .Fields("Ups/Plate").Value = Val(Col)
                fpSpread1.GetText 4, i, Col: .Fields("PrintingQuantity").Value = Val(Col)
                fpSpread1.GetText 5, i, Col: .Fields("BillingQuantity").Value = Val(Col)
                fpSpread1.GetText 6, i, Col: .Fields("FrontPrintingColor").Value = Val(Col)
                fpSpread1.GetText 7, i, Col: .Fields("BackPrintingColor").Value = Val(Col)
                fpSpread1.GetText 8, i, Col: .Fields("Book").Value = Col
                If IsNull(.Fields("QuantityReceived").Value) Then .Fields("QuantityReceived").Value = 0
                If IsNull(.Fields("QuantityIssued").Value) Then .Fields("QuantityIssued").Value = 0
            End If
        End With
        rstBookPOChild0901.Update
    Next
End Sub
Private Function CheckMandatoryFields() As Boolean
    If Combo1.ListIndex < 0 Then Combo1.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo11.ListIndex < 0 Then Combo11.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo2.ListIndex < 0 Then Combo2.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo22.ListIndex < 0 Then Combo22.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo3.ListIndex < 0 Then Combo3.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput16.Value <> 0 Then If MhRealInput16.Value <> MhRealInput25.Value + MhRealInput32.Value Then MhRealInput9.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput24.Value <> 0 Then If MhRealInput24.Value <> MhRealInput27.Value Then MhRealInput26.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput9.Value <> 0 Or MhRealInput26.Value <> 0 Or MhRealInput33.Value <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
End Function
Private Sub LoadMasterList()
    If rstSizeList.State = adStateOpen Then rstSizeList.Close
    rstSizeList.Open "SELECT Name As Col0, Code From GeneralMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstSizeList.ActiveConnection = Nothing
    If rstPaperList.State = adStateOpen Then rstPaperList.Close
    rstPaperList.Open "SELECT LTRIM(P.Name)+' (UOM : '+LTRIM(C.Name)+')' As Col0,[Weight/Unit] As Wt,C.Value1 As SPU,inWidth,inLength,[Form],GSM,P.Code FROM PaperMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code ORDER BY P.Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperList.ActiveConnection = Nothing
    If rstPlateMakerList.State = adStateOpen Then rstPlateMakerList.Close
    rstPlateMakerList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstPlateMakerList.ActiveConnection = Nothing
    If rstItemList.State = adStateOpen Then rstItemList.Close
    rstItemList.Open "SELECT Name As Col0,Code From BookMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstItemList.ActiveConnection = Nothing
End Sub
Private Function CheckDuplicateItem(ByVal CurRow As Double, ByVal xItem As String) As Boolean
    Dim i As Integer, Item As Variant
    With fpSpread1
        For i = CurRow To .DataRowCnt
            .GetText 1, i, Item
            If Item = xItem And i <> CurRow Then CheckDuplicateItem = True: Call DisplayError("Duplicate Item in Row #" + Trim(i)): Exit For
        Next
    End With
End Function
Private Sub CalculateConsumption()
    If SPU = 0 Or MhRealInput12.Value = 0 Then Exit Sub
    Dim C As Long, W As Long, q As Long
    If MhRealInput12.Value > 0 Then
        q = MhRealInput2.Value 'Qty (Sheets)
        W = (q * (MhRealInput11.Value + MhRealInput14.Value)) / 100  'Wastage (in Sheets)
        If W < (MhRealInput23.Value + MhRealInput29.Value) Then W = (MhRealInput23.Value + MhRealInput29.Value) 'Comparison with Minimum Wastage
        C = q + W   'Consumption With Wastage (in Sheets)
        C = C / MhRealInput12.Value
        MhRealInput13.Value = IIf(MhRealInput52.Value > 0, Round(((MhRealInput52.Value / 25.4) * inWidth * GSM) / 3100, 3), Wt) * (C / SPU)
        MhRealInput30.Value = CLng(Int(W / SPU)) + ((W Mod SPU) / 1000) 'Min Wastage Final
        MhRealInput28.Value = CLng(Int(C / SPU)) + ((C Mod SPU) / 1000)
    End If
End Sub
Private Sub CalculateTotal()
    Dim i As Integer, Qty As Variant
    With fpSpread1
        MhRealInput1.Value = 0: MhRealInput2.Value = 0: MhRealInput15.Value = 0: MhRealInput19.Value = 0: MhRealInput20.Value = 0
        For i = 1 To .DataRowCnt
            .GetText 2, i, Qty: MhRealInput1.Value = MhRealInput1.Value + Val(Qty)
            .GetText 4, i, Qty: If Val(Qty) > MhRealInput2.Value Then MhRealInput2.Value = Val(Qty)
            .GetText 5, i, Qty: If Val(Qty) > MhRealInput15.Value Then MhRealInput15.Value = Val(Qty)
            .GetText 6, i, Qty: If Val(Qty) > MhRealInput19.Value Then MhRealInput19.Value = Val(Qty)
            .GetText 7, i, Qty: If Val(Qty) > MhRealInput20.Value Then MhRealInput20.Value = Val(Qty)
        Next
        MhRealInput6.Value = MhRealInput15.Value
        MhRealInput3.Value = IIf(Combo3.ListIndex = 0, MhRealInput19.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, 0))
        MhRealInput7.Value = IIf(Combo3.ListIndex = 0, MhRealInput20.Value, IIf(MhRealInput20.Value > MhRealInput19.Value, MhRealInput20.Value, 0))
        If MhRealInput19.Value = 0 Then MhRealInput4.Value = 0: MhRealInput5.Value = 0
        If MhRealInput20.Value = 0 Then MhRealInput8.Value = 0: MhRealInput36.Value = 0
    End With
End Sub
Private Sub CalculatePrintAmount()
    Dim TotalForms As Long, TaxableAmt As Double
    TotalForms = MhRealInput6.Value * IIf(Combo3.ListIndex = 0, 1, 2)
    TaxableAmt = (MhRealInput3.Value * IIf(TotalForms < 1000, 1, TotalForms / 1000) * MhRealInput5.Value + MhRealInput7.Value * IIf(TotalForms < 1000, 1, TotalForms / 1000) * MhRealInput8.Value) + MhRealInput9.Value
    MhRealInput17.Value = TaxableAmt * MhRealInput18.Value / 100
    MhRealInput25.Value = TaxableAmt + MhRealInput17.Value
End Sub
Private Sub CalculatePlateAmount()
    Dim TaxableAmt As Double
    TaxableAmt = (MhRealInput3.Value * MhRealInput4.Value + MhRealInput7.Value * MhRealInput36.Value) + MhRealInput26.Value
    MhRealInput22.Value = TaxableAmt * MhRealInput21.Value / 100
    MhRealInput27.Value = TaxableAmt + MhRealInput22.Value
End Sub
Private Sub CalculatePaperAmount()
    Dim TaxableAmt As Double
    TaxableAmt = (MhRealInput31.Value * MhRealInput13.Value) + MhRealInput33.Value
    MhRealInput35.Value = TaxableAmt * MhRealInput34.Value / 100    'Tax
    MhRealInput32.Value = TaxableAmt + MhRealInput35.Value
End Sub
Private Sub GetPrinterRates(ByVal RateType As String, Optional ByVal Position As String)
    If MhRealInput6.Value = 0 Or CheckEmpty(SizeCode, False) Or MhRealInput19.Value + MhRealInput20.Value = 0 Then Exit Sub
    Dim frontPlateRate As Double, backPlateRate As Double, frontPrintRate As Double, backPrintRate As Double, frontPaperWastageRate As Double, backPaperWastageRate As Double, frontPaperWastageMin As Long, backPaperWastageMin As Long, Col As String
    On Error GoTo ErrorHandler
    'Fetching Front Rates
    If MhRealInput19.Value > 0 Then
        Col = IIf(MhRealInput19.Value <= 2, MhRealInput19.Value, IIf(MhRealInput19.Value <= 4, "4", "6"))
        If rstFetchRate.State = adStateOpen Then rstFetchRate.Close
        rstFetchRate.Open "SELECT TOP 1 P.* FROM AccountChild05 P INNER JOIN SizeGroupChild C ON P.[Size]=C.Code WHERE P.Code='" & PartyCode & "' AND C.[Size]='" & SizeCode & "' AND Range" & Col & ">=" & MhRealInput6.Value & " ORDER BY Range" & Col, cnDatabase, adOpenKeyset, adLockReadOnly
        If rstFetchRate.RecordCount = 0 Then
            If rstFetchRate.State = adStateOpen Then rstFetchRate.Close
            rstFetchRate.Open "SELECT TOP 1 C1.* FROM (AccountMaster P INNER JOIN AccountChild05 C1 ON P.Code=C1.Code) INNER JOIN SizeGroupChild C2 ON C1.[Size]=C2.Code WHERE Name LIKE '%Rate%' AND C2.[Size]='" & SizeCode & "' AND Range" & Col & ">=" & MhRealInput6.Value & " ORDER BY Range" & Col, cnDatabase, adOpenKeyset, adLockReadOnly
        End If
        If rstFetchRate.RecordCount > 0 Then
            If RateType = "L" Then  'Plate Rate
                frontPlateRate = Val(rstFetchRate.Fields(Choose(Combo2.ListIndex + 1, "DeepatchPlateRate", "PSPlateRate", "WipeonPlateRate", "CTPPlateRate") & Col).Value)
            ElseIf RateType = "P" Then  'Print Rate
                frontPrintRate = Val(rstFetchRate.Fields("PrintRate" & Col).Value)
            ElseIf RateType = "W" Then  'Paper Wastage (Percentage)
                frontPaperWastageRate = Val(rstFetchRate.Fields("PaperWastageRate" & Col).Value)
            ElseIf RateType = "M" Then  'Paper Wastage (Minimum Sheets)
                frontPaperWastageMin = Val(rstFetchRate.Fields("PaperWastageMin" & Col).Value)
            End If
        End If
    End If
    'Fetching Back Rates
    If MhRealInput20.Value > 0 Then
        Col = IIf(MhRealInput20.Value <= 2, MhRealInput20.Value, IIf(MhRealInput20.Value <= 4, "4", "6"))
        If rstFetchRate.State = adStateOpen Then rstFetchRate.Close
        rstFetchRate.Open "SELECT TOP 1 P.* FROM AccountChild05 P INNER JOIN SizeGroupChild C ON P.[Size]=C.Code WHERE P.Code='" & PartyCode & "' AND C.[Size]='" & SizeCode & "' AND Range" & Col & ">=" & MhRealInput6.Value & " ORDER BY Range" & Col, cnDatabase, adOpenKeyset, adLockReadOnly
        If rstFetchRate.RecordCount = 0 Then
            If rstFetchRate.State = adStateOpen Then rstFetchRate.Close
            rstFetchRate.Open "SELECT TOP 1 C1.* FROM (AccountMaster P INNER JOIN AccountChild05 C1 ON P.Code=C1.Code) INNER JOIN SizeGroupChild C2 ON C1.[Size]=C2.Code WHERE Name LIKE '%Rate%' AND C2.[Size]='" & SizeCode & "' AND Range" & Col & ">=" & MhRealInput6.Value & " ORDER BY Range" & Col, cnDatabase, adOpenKeyset, adLockReadOnly
        End If
        If rstFetchRate.RecordCount > 0 Then
            If RateType = "L" Then  'Plate Rate
                backPlateRate = Val(rstFetchRate.Fields(Choose(Combo22.ListIndex + 1, "DeepatchPlateRate", "PSPlateRate", "WipeonPlateRate", "CTPPlateRate") & Col).Value)
            ElseIf RateType = "P" Then  'Print Rate
                backPrintRate = Val(rstFetchRate.Fields("PrintRate" & Col).Value)
            ElseIf RateType = "W" Then  'Paper Wastage (Percentage)
                backPaperWastageRate = Val(rstFetchRate.Fields("PaperWastageRate" & Col).Value)
            ElseIf RateType = "M" Then  'Paper Wastage (Minimum Sheets)
                backPaperWastageMin = Val(rstFetchRate.Fields("PaperWastageMin" & Col).Value)
            End If
        End If
    End If
    If RateType = "L" Then
        If Position = "F" Then
            If IIf(Combo3.ListIndex = 0, MhRealInput19.Value > 0, MhRealInput19.Value > MhRealInput20.Value) Then
                If Combo1.ListIndex > 0 Then
                    If frontPlateRate > 0 Then
                        If MhRealInput4.Value = 0 Then
                            MhRealInput4.Value = frontPlateRate
                        ElseIf MhRealInput4.Value <> frontPlateRate Then
                            If MsgBox("Front Plate Rate [" & Trim(MhRealInput4.Value) & "] is different from that in Master [" & Trim(Format(frontPlateRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput4.Value = frontPlateRate
                        End If
                    End If
                Else
                    If InStr(1, "1_3", Trim(Combo2.ListIndex)) > 0 Then MhRealInput4.Value = 0
                End If
            Else
                MhRealInput4.Value = 0
            End If
        End If
        If Position = "B" Then
            If IIf(Combo3.ListIndex = 0, MhRealInput20.Value > 0, MhRealInput20.Value > MhRealInput19.Value) Then
                If Combo11.ListIndex > 0 Then
                    If backPlateRate > 0 Then
                        If MhRealInput36.Value = 0 Then
                            MhRealInput36.Value = backPlateRate
                        ElseIf MhRealInput36.Value <> backPlateRate Then
                            If MsgBox("Back Plate Rate [" & Trim(MhRealInput36.Value) & "] is different from that in Master [" & Trim(Format(backPlateRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput36.Value = backPlateRate
                        End If
                    End If
                Else
                    If InStr(1, "1_3", Trim(Combo22.ListIndex)) > 0 Then MhRealInput36.Value = 0
                End If
            Else
                MhRealInput36.Value = 0
            End If
        End If
    ElseIf RateType = "P" Then
        If Position = "F" Then
            If IIf(Combo3.ListIndex = 0, MhRealInput19.Value > 0, MhRealInput19.Value > MhRealInput20.Value) Then
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
            If IIf(Combo3.ListIndex = 0, MhRealInput20.Value > 0, MhRealInput20.Value > MhRealInput19.Value) Then
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
            If IIf(Combo3.ListIndex = 0, MhRealInput19.Value > 0, MhRealInput19.Value > MhRealInput20.Value) Then
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
            If IIf(Combo3.ListIndex = 0, MhRealInput20.Value > 0, MhRealInput20.Value > MhRealInput19.Value) Then
                If backPaperWastageRate > 0 Then
                    If MhRealInput14.Value = 0 Then
                        MhRealInput14.Value = backPaperWastageRate
                    ElseIf MhRealInput14.Value <> backPaperWastageRate Then
                        If MsgBox("Back Paper Wastage Rate [" & Trim(MhRealInput14.Value) & "] is different from that in Master [" & Trim(Format(backPaperWastageRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput14.Value = backPaperWastageRate
                    End If
                End If
            Else
                MhRealInput14.Value = 0
            End If
        End If
    ElseIf RateType = "M" Then
        If Position = "F" Then
            If IIf(Combo3.ListIndex = 0, MhRealInput19.Value > 0, MhRealInput19.Value > MhRealInput20.Value) Then
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
        If Position = "B" Then
            If IIf(Combo3.ListIndex = 0, MhRealInput20.Value > 0, MhRealInput20.Value > MhRealInput19.Value) Then
                If backPaperWastageMin > 0 Then
                    If MhRealInput29.Value = 0 Then
                        MhRealInput29.Value = backPaperWastageMin
                    ElseIf MhRealInput29.Value <> backPaperWastageMin Then
                        If MsgBox("Back Paper Wastage Min [" & Trim(MhRealInput29.Value) & "] is different from that in Master [" & Trim(Format(backPaperWastageMin, "#0")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput29.Value = backPaperWastageMin
                    End If
                End If
            Else
                MhRealInput29.Value = 0
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub
