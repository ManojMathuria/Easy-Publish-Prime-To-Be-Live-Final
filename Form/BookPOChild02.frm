VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmBookPOChild02 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Sheet Digital Printing Order Details"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11160
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
   ScaleHeight     =   9765
   ScaleWidth      =   11160
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008BD6FE&
      Height          =   375
      Left            =   10575
      Picture         =   "BookPOChild02.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      BackColor       =   &H008BD6FE&
      Height          =   375
      Left            =   10575
      Picture         =   "BookPOChild02.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   9525
      Left            =   120
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   105
      Width           =   10335
      _Version        =   65536
      _ExtentX        =   18230
      _ExtentY        =   16801
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
      Picture         =   "BookPOChild02.frx":0204
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel33 
         Height          =   330
         Left            =   6840
         TabIndex        =   91
         Top             =   960
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
         Caption         =   " Size"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":0220
         Picture         =   "BookPOChild02.frx":023C
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput36 
         Height          =   330
         Left            =   1800
         TabIndex        =   44
         ToolTipText     =   "Plate"
         Top             =   6030
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":0258
         Caption         =   "BookPOChild02.frx":0278
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":02E4
         Keys            =   "BookPOChild02.frx":0302
         Spin            =   "BookPOChild02.frx":034C
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
         MinValue        =   -999999999999.99
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel42 
         Height          =   330
         Left            =   6840
         TabIndex        =   102
         Top             =   8835
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
         Caption         =   " Total Amt-Plate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":0374
         Picture         =   "BookPOChild02.frx":0390
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   9120
         Width           =   1695
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   5055
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   7875
         Width           =   5055
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
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
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
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   53
         Top             =   7410
         Width           =   8415
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
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   52
         Top             =   7095
         Width           =   8415
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   105
         Width           =   1695
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   46
         Top             =   6525
         Width           =   1695
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   31
         Top             =   3135
         Width           =   5055
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   645
         Width           =   8415
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
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1230
         Width           =   8415
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   3480
         TabIndex        =   57
         Top             =   105
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
         Caption         =   " Order Date"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":03AC
         Picture         =   "BookPOChild02.frx":03C8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   3480
         TabIndex        =   59
         Top             =   1545
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
         Caption         =   " Billing Qty"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":03E4
         Picture         =   "BookPOChild02.frx":0400
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   3480
         TabIndex        =   62
         Top             =   8565
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
         Caption         =   " Plate Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":041C
         Picture         =   "BookPOChild02.frx":0438
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   3480
         TabIndex        =   63
         Top             =   2640
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
         Caption         =   " Print Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":0454
         Picture         =   "BookPOChild02.frx":0470
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   6840
         TabIndex        =   67
         Top             =   8565
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
         Caption         =   " Plate Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":048C
         Picture         =   "BookPOChild02.frx":04A8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   6840
         TabIndex        =   68
         Top             =   2640
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
         Caption         =   " Print Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":04C4
         Picture         =   "BookPOChild02.frx":04E0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   6840
         TabIndex        =   69
         Top             =   5760
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
         Caption         =   " Total Amt-Ptg."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":04FC
         Picture         =   "BookPOChild02.frx":0518
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Left            =   3480
         TabIndex        =   72
         Top             =   3450
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
         Caption         =   " Consumption"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":0534
         Picture         =   "BookPOChild02.frx":0550
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
         Height          =   330
         Left            =   6840
         TabIndex        =   73
         Top             =   3765
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
         Caption         =   " Total Consmptn"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":056C
         Picture         =   "BookPOChild02.frx":0588
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   6840
         TabIndex        =   75
         Top             =   6525
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
         Caption         =   " Paid Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":05A4
         Picture         =   "BookPOChild02.frx":05C0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   3480
         TabIndex        =   76
         Top             =   6525
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
         Caption         =   " Bill Date"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":05DC
         Picture         =   "BookPOChild02.frx":05F8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   6840
         TabIndex        =   77
         Top             =   105
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
         Caption         =   " Target Date"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":0614
         Picture         =   "BookPOChild02.frx":0630
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Left            =   3480
         TabIndex        =   79
         Top             =   8250
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
         Caption         =   " Plate Type"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":064C
         Picture         =   "BookPOChild02.frx":0668
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
         Height          =   330
         Left            =   3480
         TabIndex        =   82
         Top             =   5760
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
         Caption         =   " GST-Ptg."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":0684
         Picture         =   "BookPOChild02.frx":06A0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel29 
         Height          =   330
         Left            =   6840
         TabIndex        =   83
         Top             =   7875
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
         Caption         =   " Plate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":06BC
         Picture         =   "BookPOChild02.frx":06D8
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   5160
         TabIndex        =   1
         Top             =   105
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calendar        =   "BookPOChild02.frx":06F4
         Caption         =   "BookPOChild02.frx":080C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":0878
         Keys            =   "BookPOChild02.frx":0896
         Spin            =   "BookPOChild02.frx":08F4
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
         Left            =   8520
         TabIndex        =   2
         Top             =   105
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calendar        =   "BookPOChild02.frx":091C
         Caption         =   "BookPOChild02.frx":0A34
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":0AA0
         Keys            =   "BookPOChild02.frx":0ABE
         Spin            =   "BookPOChild02.frx":0B1C
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
         Left            =   5160
         TabIndex        =   47
         Top             =   6525
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calendar        =   "BookPOChild02.frx":0B44
         Caption         =   "BookPOChild02.frx":0C5C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":0CC8
         Keys            =   "BookPOChild02.frx":0CE6
         Spin            =   "BookPOChild02.frx":0D44
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel30 
         Height          =   330
         Left            =   6840
         TabIndex        =   84
         Top             =   2370
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
         Caption         =   " Forms/Sheet"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":0D6C
         Picture         =   "BookPOChild02.frx":0D88
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel31 
         Height          =   330
         Left            =   6840
         TabIndex        =   85
         Top             =   3135
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
         Caption         =   " Forms/Sheet"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":0DA4
         Picture         =   "BookPOChild02.frx":0DC0
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   1800
         TabIndex        =   9
         Top             =   1545
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":0DDC
         Caption         =   "BookPOChild02.frx":0DFC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":0E68
         Keys            =   "BookPOChild02.frx":0E86
         Spin            =   "BookPOChild02.frx":0ED0
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
         ForeColor       =   -2147483640
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   5160
         TabIndex        =   10
         ToolTipText     =   "One Color"
         Top             =   1545
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":0EF8
         Caption         =   "BookPOChild02.frx":0F18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":0F84
         Keys            =   "BookPOChild02.frx":0FA2
         Spin            =   "BookPOChild02.frx":0FEC
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
         ForeColor       =   -2147483640
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
         Height          =   330
         Left            =   7680
         TabIndex        =   11
         ToolTipText     =   "Double & Four Color"
         Top             =   1545
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":1014
         Caption         =   "BookPOChild02.frx":1034
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":10A0
         Keys            =   "BookPOChild02.frx":10BE
         Spin            =   "BookPOChild02.frx":1108
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
         ForeColor       =   -2147483640
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
         Height          =   330
         Left            =   8520
         TabIndex        =   18
         Top             =   2370
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":1130
         Caption         =   "BookPOChild02.frx":1150
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":11BC
         Keys            =   "BookPOChild02.frx":11DA
         Spin            =   "BookPOChild02.frx":1224
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   2
         MinValue        =   0.5
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171638789
         Value           =   1
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   1800
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "¼ Form"
         Top             =   8565
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":124C
         Caption         =   "BookPOChild02.frx":126C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":12D8
         Keys            =   "BookPOChild02.frx":12F6
         Spin            =   "BookPOChild02.frx":1340
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
         ForeColor       =   -2147483640
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput23 
         Height          =   330
         Left            =   2400
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "½ Form"
         Top             =   8565
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":1368
         Caption         =   "BookPOChild02.frx":1388
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":13F4
         Keys            =   "BookPOChild02.frx":1412
         Spin            =   "BookPOChild02.frx":145C
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
         ForeColor       =   -2147483640
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
         Height          =   330
         Left            =   2760
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "1 Form"
         Top             =   8565
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":1484
         Caption         =   "BookPOChild02.frx":14A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":1510
         Keys            =   "BookPOChild02.frx":152E
         Spin            =   "BookPOChild02.frx":1578
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
         ForeColor       =   -2147483640
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   5160
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   8565
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":15A0
         Caption         =   "BookPOChild02.frx":15C0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":162C
         Keys            =   "BookPOChild02.frx":164A
         Spin            =   "BookPOChild02.frx":1694
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
         ValueVT         =   171638789
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   8520
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   8565
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":16BC
         Caption         =   "BookPOChild02.frx":16DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":1748
         Keys            =   "BookPOChild02.frx":1766
         Spin            =   "BookPOChild02.frx":17B0
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171638789
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   1800
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "¼ Form"
         Top             =   2640
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":17D8
         Caption         =   "BookPOChild02.frx":17F8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":1864
         Keys            =   "BookPOChild02.frx":1882
         Spin            =   "BookPOChild02.frx":18CC
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
         ValueVT         =   171638789
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
         Height          =   330
         Left            =   2445
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "½ Form"
         Top             =   2640
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":18F4
         Caption         =   "BookPOChild02.frx":1914
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":1980
         Keys            =   "BookPOChild02.frx":199E
         Spin            =   "BookPOChild02.frx":19E8
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
         ValueVT         =   171638789
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
         Height          =   330
         Left            =   2970
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "1 Form"
         Top             =   2640
         Width           =   525
         _Version        =   65536
         _ExtentX        =   926
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":1A10
         Caption         =   "BookPOChild02.frx":1A30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":1A9C
         Keys            =   "BookPOChild02.frx":1ABA
         Spin            =   "BookPOChild02.frx":1B04
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
         ValueVT         =   171638789
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   5160
         TabIndex        =   28
         Top             =   2640
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":1B2C
         Caption         =   "BookPOChild02.frx":1B4C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":1BB8
         Keys            =   "BookPOChild02.frx":1BD6
         Spin            =   "BookPOChild02.frx":1C20
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
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
         Height          =   330
         Left            =   8520
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":1C48
         Caption         =   "BookPOChild02.frx":1C68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":1CD4
         Keys            =   "BookPOChild02.frx":1CF2
         Spin            =   "BookPOChild02.frx":1D3C
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
         Height          =   330
         Left            =   4440
         TabIndex        =   41
         Top             =   5760
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":1D64
         Caption         =   "BookPOChild02.frx":1D84
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":1DF0
         Keys            =   "BookPOChild02.frx":1E0E
         Spin            =   "BookPOChild02.frx":1E58
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
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
         Height          =   330
         Left            =   5160
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   5760
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":1E80
         Caption         =   "BookPOChild02.frx":1EA0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":1F0C
         Keys            =   "BookPOChild02.frx":1F2A
         Spin            =   "BookPOChild02.frx":1F74
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   330
         Left            =   1800
         TabIndex        =   40
         ToolTipText     =   "Print"
         Top             =   5760
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":1F9C
         Caption         =   "BookPOChild02.frx":1FBC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":2028
         Keys            =   "BookPOChild02.frx":2046
         Spin            =   "BookPOChild02.frx":2090
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
         MinValue        =   -999999999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
         Height          =   330
         Left            =   8520
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   5760
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":20B8
         Caption         =   "BookPOChild02.frx":20D8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":2144
         Keys            =   "BookPOChild02.frx":2162
         Spin            =   "BookPOChild02.frx":21AC
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
         Height          =   330
         Left            =   8520
         TabIndex        =   32
         Top             =   3135
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":21D4
         Caption         =   "BookPOChild02.frx":21F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":2260
         Keys            =   "BookPOChild02.frx":227E
         Spin            =   "BookPOChild02.frx":22C8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   2
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   1800
         TabIndex        =   33
         ToolTipText     =   "Wastage %"
         Top             =   3450
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":22F0
         Caption         =   "BookPOChild02.frx":2310
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":237C
         Keys            =   "BookPOChild02.frx":239A
         Spin            =   "BookPOChild02.frx":23E4
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
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
         Height          =   330
         Left            =   5160
         TabIndex        =   35
         Top             =   3450
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":240C
         Caption         =   "BookPOChild02.frx":242C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":2498
         Keys            =   "BookPOChild02.frx":24B6
         Spin            =   "BookPOChild02.frx":2500
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.000"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.000"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   8520
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3765
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":2528
         Caption         =   "BookPOChild02.frx":2548
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":25B4
         Keys            =   "BookPOChild02.frx":25D2
         Spin            =   "BookPOChild02.frx":261C
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.000"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.000"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
         Height          =   330
         Left            =   8520
         TabIndex        =   48
         Top             =   6525
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":2644
         Caption         =   "BookPOChild02.frx":2664
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":26D0
         Keys            =   "BookPOChild02.frx":26EE
         Spin            =   "BookPOChild02.frx":2738
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
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   1335
         Left            =   120
         TabIndex        =   39
         Top             =   4260
         Width           =   10095
         _Version        =   524288
         _ExtentX        =   17806
         _ExtentY        =   2355
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
         GridColor       =   33023
         MaxCols         =   32
         MaxRows         =   3
         OperationMode   =   2
         SpreadDesigner  =   "BookPOChild02.frx":2760
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   3480
         TabIndex        =   89
         Top             =   8835
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
         Caption         =   " GST-Plate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":3D2B
         Picture         =   "BookPOChild02.frx":3D47
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput28 
         Height          =   330
         Left            =   4440
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   8835
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":3D63
         Caption         =   "BookPOChild02.frx":3D83
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":3DEF
         Keys            =   "BookPOChild02.frx":3E0D
         Spin            =   "BookPOChild02.frx":3E57
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
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput29 
         Height          =   330
         Left            =   5160
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   8835
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":3E7F
         Caption         =   "BookPOChild02.frx":3E9F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":3F0B
         Keys            =   "BookPOChild02.frx":3F29
         Spin            =   "BookPOChild02.frx":3F73
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel37 
         Height          =   330
         Left            =   6840
         TabIndex        =   95
         Top             =   9120
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
         Caption         =   " Paid Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":3F9B
         Picture         =   "BookPOChild02.frx":3FB7
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
         Height          =   330
         Left            =   3480
         TabIndex        =   96
         Top             =   9120
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
         Caption         =   " Bill Date"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":3FD3
         Picture         =   "BookPOChild02.frx":3FEF
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput4 
         Height          =   330
         Left            =   5160
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   9120
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calendar        =   "BookPOChild02.frx":400B
         Caption         =   "BookPOChild02.frx":4123
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":418F
         Keys            =   "BookPOChild02.frx":41AD
         Spin            =   "BookPOChild02.frx":420B
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput30 
         Height          =   330
         Left            =   8520
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   9120
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":4233
         Caption         =   "BookPOChild02.frx":4253
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":42BF
         Keys            =   "BookPOChild02.frx":42DD
         Spin            =   "BookPOChild02.frx":4327
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
         ValueVT         =   171900933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput31 
         Height          =   330
         Left            =   2640
         TabIndex        =   34
         ToolTipText     =   " Wastage Min. Sheets"
         Top             =   3450
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":434F
         Caption         =   "BookPOChild02.frx":436F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":43DB
         Keys            =   "BookPOChild02.frx":43F9
         Spin            =   "BookPOChild02.frx":4443
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput32 
         Height          =   330
         Left            =   1800
         TabIndex        =   36
         Top             =   3765
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":446B
         Caption         =   "BookPOChild02.frx":448B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":44F7
         Keys            =   "BookPOChild02.frx":4515
         Spin            =   "BookPOChild02.frx":455F
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
         ValueVT         =   171900933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel40 
         Height          =   330
         Left            =   3480
         TabIndex        =   98
         Top             =   3765
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
         Caption         =   " Paper Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":4587
         Picture         =   "BookPOChild02.frx":45A3
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput33 
         Height          =   330
         Left            =   5160
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3765
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":45BF
         Caption         =   "BookPOChild02.frx":45DF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":464B
         Keys            =   "BookPOChild02.frx":4669
         Spin            =   "BookPOChild02.frx":46B3
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171900933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput34 
         Height          =   330
         Left            =   1800
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Plate"
         Top             =   8835
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":46DB
         Caption         =   "BookPOChild02.frx":46FB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":4767
         Keys            =   "BookPOChild02.frx":4785
         Spin            =   "BookPOChild02.frx":47CF
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
         MinValue        =   -999999999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171900933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput35 
         Height          =   330
         Left            =   8520
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   8835
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":47F7
         Caption         =   "BookPOChild02.frx":4817
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":4883
         Keys            =   "BookPOChild02.frx":48A1
         Spin            =   "BookPOChild02.frx":48EB
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171900933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel44 
         Height          =   330
         Left            =   6840
         TabIndex        =   104
         Top             =   6030
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
         Caption         =   " Total Amt-Paper"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":4913
         Picture         =   "BookPOChild02.frx":492F
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel45 
         Height          =   330
         Left            =   3480
         TabIndex        =   105
         Top             =   6030
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
         Caption         =   " GST-Paper"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":494B
         Picture         =   "BookPOChild02.frx":4967
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput37 
         Height          =   330
         Left            =   4440
         TabIndex        =   45
         Top             =   6030
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":4983
         Caption         =   "BookPOChild02.frx":49A3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":4A0F
         Keys            =   "BookPOChild02.frx":4A2D
         Spin            =   "BookPOChild02.frx":4A77
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
         ValueVT         =   171900933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput38 
         Height          =   330
         Left            =   5160
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   6030
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":4A9F
         Caption         =   "BookPOChild02.frx":4ABF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":4B2B
         Keys            =   "BookPOChild02.frx":4B49
         Spin            =   "BookPOChild02.frx":4B93
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171900933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput39 
         Height          =   330
         Left            =   8520
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   6030
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":4BBB
         Caption         =   "BookPOChild02.frx":4BDB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":4C47
         Keys            =   "BookPOChild02.frx":4C65
         Spin            =   "BookPOChild02.frx":4CAF
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   171835397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput59 
         Height          =   330
         Left            =   3120
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Revised Plates"
         Top             =   8565
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":4CD7
         Caption         =   "BookPOChild02.frx":4CF7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":4D63
         Keys            =   "BookPOChild02.frx":4D81
         Spin            =   "BookPOChild02.frx":4DCB
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
         ForeColor       =   -2147483640
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel46 
         Height          =   330
         Left            =   6840
         TabIndex        =   108
         Top             =   3450
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
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
         Picture         =   "BookPOChild02.frx":4DF3
         Picture         =   "BookPOChild02.frx":4E0F
         Begin VB.CheckBox chkPaper 
            BackColor       =   &H00000000&
            Caption         =   "Check2"
            Height          =   210
            Left            =   2340
            TabIndex        =   30
            Top             =   60
            Value           =   1  'Checked
            Width           =   210
         End
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel43 
         Height          =   330
         Left            =   120
         TabIndex        =   103
         Top             =   6030
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
         Caption         =   " Adj-Paper"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":4E2B
         Picture         =   "BookPOChild02.frx":4E47
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel41 
         Height          =   330
         Left            =   120
         TabIndex        =   100
         Top             =   8835
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
         Caption         =   " Adj-Plate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":4E63
         Picture         =   "BookPOChild02.frx":4E7F
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   120
         TabIndex        =   99
         Top             =   5760
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
         Caption         =   " Adj-Printing"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":4E9B
         Picture         =   "BookPOChild02.frx":4EB7
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel39 
         Height          =   330
         Left            =   120
         TabIndex        =   97
         Top             =   3765
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
         Caption         =   " Paper Rate/Kg."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":4ED3
         Picture         =   "BookPOChild02.frx":4EEF
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel36 
         Height          =   330
         Left            =   120
         TabIndex        =   94
         Top             =   9120
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Picture         =   "BookPOChild02.frx":4F0B
         Picture         =   "BookPOChild02.frx":4F27
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel35 
         Height          =   330
         Left            =   120
         TabIndex        =   93
         Top             =   960
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
         Caption         =   " Party Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":4F43
         Picture         =   "BookPOChild02.frx":4F5F
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel34 
         Height          =   330
         Left            =   120
         TabIndex        =   92
         Top             =   7875
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
         Caption         =   " Plate Party"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":4F7B
         Picture         =   "BookPOChild02.frx":4F97
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Left            =   120
         TabIndex        =   88
         Top             =   7410
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
         Caption         =   " Adj.Remarks"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":4FB3
         Picture         =   "BookPOChild02.frx":4FCF
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   81
         Top             =   7095
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
         Caption         =   " Remarks"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":4FEB
         Picture         =   "BookPOChild02.frx":5007
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   80
         Top             =   105
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
         Caption         =   " Order No."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":5023
         Picture         =   "BookPOChild02.frx":503F
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   78
         Top             =   645
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
         Caption         =   " Item Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":505B
         Picture         =   "BookPOChild02.frx":5077
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   74
         Top             =   6525
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Picture         =   "BookPOChild02.frx":5093
         Picture         =   "BookPOChild02.frx":50AF
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Left            =   120
         TabIndex        =   71
         Top             =   3450
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
         Caption         =   " Wastage (%)"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":50CB
         Picture         =   "BookPOChild02.frx":50E7
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   120
         TabIndex        =   70
         Top             =   3135
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
         Caption         =   " Paper Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":5103
         Picture         =   "BookPOChild02.frx":511F
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   66
         Top             =   2640
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
         Caption         =   " Total Forms"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":513B
         Picture         =   "BookPOChild02.frx":5157
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   120
         TabIndex        =   65
         Top             =   2370
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
         Caption         =   " Pages/Forms"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":5173
         Picture         =   "BookPOChild02.frx":518F
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   120
         TabIndex        =   64
         Top             =   1230
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
         Caption         =   " Ref.No."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":51AB
         Picture         =   "BookPOChild02.frx":51C7
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   61
         Top             =   8565
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
         Caption         =   " Total Plates"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":51E3
         Picture         =   "BookPOChild02.frx":51FF
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   60
         Top             =   2055
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
         Caption         =   " Printing Type"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":521B
         Picture         =   "BookPOChild02.frx":5237
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   58
         Top             =   1545
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
         Caption         =   " Actual Qty"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild02.frx":5253
         Picture         =   "BookPOChild02.frx":526F
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
         Height          =   330
         Left            =   6000
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "1 Form"
         Top             =   2370
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":528B
         Caption         =   "BookPOChild02.frx":52AB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":5317
         Keys            =   "BookPOChild02.frx":5335
         Spin            =   "BookPOChild02.frx":537F
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999
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
         Height          =   330
         Left            =   5160
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "½ Form"
         Top             =   2370
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":53A7
         Caption         =   "BookPOChild02.frx":53C7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":5433
         Keys            =   "BookPOChild02.frx":5451
         Spin            =   "BookPOChild02.frx":549B
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   3480
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "¼ Form"
         Top             =   2370
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":54C3
         Caption         =   "BookPOChild02.frx":54E3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":554F
         Keys            =   "BookPOChild02.frx":556D
         Spin            =   "BookPOChild02.frx":55B7
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999
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
         Height          =   330
         Left            =   1800
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Two & Four Color"
         Top             =   2370
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild02.frx":55DF
         Caption         =   "BookPOChild02.frx":55FF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild02.frx":566B
         Keys            =   "BookPOChild02.frx":5689
         Spin            =   "BookPOChild02.frx":56D3
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
         ValueVT         =   171638789
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   10350
         Y1              =   6435
         Y2              =   6435
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   0
         X2              =   10350
         Y1              =   1965
         Y2              =   1965
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   330
         Left            =   8520
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   7875
         Width           =   1695
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2990;582"
         ListRows        =   3
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   0
         X2              =   10350
         Y1              =   5670
         Y2              =   5670
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10350
         Y1              =   6950
         Y2              =   6950
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   10350
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10350
         Y1              =   4170
         Y2              =   4170
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   5160
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   8250
         Width           =   1695
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2990;582"
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
         X2              =   10350
         Y1              =   3045
         Y2              =   3045
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Left            =   1800
         TabIndex        =   12
         Top             =   2055
         Width           =   8415
         VariousPropertyBits=   545282073
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "14843;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "FrmBookPOChild02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstBookPOChild05 As New ADODB.Recordset
Public PrinterCode As String, VchType As String, VchDate As Date
Dim rstPaperList As New ADODB.Recordset
Dim rstSizeList As New ADODB.Recordset
'Dim rstRefList As New ADODB.Recordset
Dim rstPlateMakerList As New ADODB.Recordset
Dim rstPrinterRates As New ADODB.Recordset
Dim BookCode As String
Dim SizeCode As String
'Dim  RefCode As String
Dim PlateMakerCode As String
Dim PaperCode As String
Dim SPU As Variant, Wt As Variant, CutOffSize As Variant
Private Sub Form_Load()
    Dim Cnt As Integer, Pages As Variant
    On Error GoTo ErrorHandler
    CenterForm Me
    'Me.Left = (MdiMainMenu.ScaleWidth - Me.Width) \ 2
    BusySystemIndicator True
    DisableCloseButton Me
    For Cnt = 11 To 27
        fpSpread1.Col = Cnt
        fpSpread1.ColHidden = True
    Next
    AbortPO = False
    BookCode = FrmBookPrintOrder.rstBookList.Fields("Code").Value
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text)
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text)
    Text9.Text = Trim(FrmBookPrintOrder.Text5.Text)
    Combo1.AddItem "1 Color", 0
    Combo1.AddItem "2 Color", 1
    Combo1.AddItem "4 Color", 2
    Combo2.AddItem "Deepatch", 0
    Combo2.AddItem "PS", 1
    Combo2.AddItem "Wipeon", 2
    Combo2.AddItem "CTP", 3
    Combo3.AddItem "Old", 0
    Combo3.AddItem "New", 1
    Combo3.AddItem "Revised", 2
    ClearFields
'    If IsNull(rstBookPOChild05.Fields("Code").Value) Then MhRealInput5.ReadOnly = True Else MhRealInput5.ReadOnly = False
'    Call LoadRefList(BookCode, CheckNull(rstBookPOChild05.Fields("Code").Value))
    LoadMasterList
    If Val(CheckNull(rstBookPOChild05.Fields("ActualQuantity").Value)) = 0 Then
        MhRealInput1.Value = FrmBookPrintOrder.MhRealInput3.Value
        PlateMakerCode = PrinterCode
        If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
        rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
        If Not rstPlateMakerList.EOF Then Text7.Text = rstPlateMakerList.Fields("Col0").Value
        SizeCode = FrmBookPrintOrder.rstBookList.Fields("SizeCode").Value
        With fpSpread1
            For Cnt = 1 To .MaxRows
                .SetText 1, Cnt, IIf(Left(FrmBookPrintOrder.BookPOType, 1) = "R" And FrmBookPrintOrder.rstBookList.Fields("Type").Value = "R", 0, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPages").Value))
                .SetText 2, Cnt, IIf(Left(FrmBookPrintOrder.BookPOType, 1) = "R" And FrmBookPrintOrder.rstBookList.Fields("Type").Value = "R", 0, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorForms").Value))
                .SetText 3, Cnt, IIf(Left(FrmBookPrintOrder.BookPOType, 1) = "R" And FrmBookPrintOrder.rstBookList.Fields("Type").Value = "R", 0, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color¼Forms").Value))
                .SetText 4, Cnt, IIf(Left(FrmBookPrintOrder.BookPOType, 1) = "R" And FrmBookPrintOrder.rstBookList.Fields("Type").Value = "R", 0, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color½Forms").Value))
                fpSpread1.SetText 7, Cnt, 0#
                fpSpread1.SetText 8, Cnt, 0#
                fpSpread1.SetText 9, Cnt, 0#
                fpSpread1.SetText 10, Cnt, 0#
                fpSpread1.SetText 23, Cnt, SizeCode
                fpSpread1.SetText 24, Cnt, 0#
                fpSpread1.SetText 25, Cnt, 0#
                fpSpread1.SetText 26, Cnt, 0#
                fpSpread1.SetText 31, Cnt, 1
                .SetText 5, Cnt, IIf(Left(FrmBookPrintOrder.BookPOType, 1) = "R" And FrmBookPrintOrder.rstBookList.Fields("Type").Value = "R", 0, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color1F/BForms").Value) + Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color1W/TForms").Value))
                If Left(FrmBookPrintOrder.BookPOType, 1) = "R" And FrmBookPrintOrder.rstBookList.Fields("Type").Value = "R" Then .SetText 6, Cnt, "Wipeon": Combo2.ListIndex = 2 Else .SetText 6, Cnt, IIf(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "1", "Deepatch", IIf(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "2", "PS", IIf(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "3", "Wipeon", "CTP")))
            Next
        End With
        MhDateInput1.Value = FrmBookPrintOrder.MhDateInput1.Value
        MhDateInput3.Value = DateAdd("d", 2, MhDateInput1.Value)
    Else
        LoadFields
    End If
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 1, Cnt, Pages
        If Val(Pages) > 0 Then
            fpSpread1.SetActiveCell 1, Cnt
            fpSpread1_DblClick 1, Cnt
            Exit For
        End If
    Next
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
       Sendkeys "{TAB}"
       KeyCode = 0
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
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstSizeList)
'    Call CloseRecordset(rstRefList)
    Call CloseRecordset(rstPlateMakerList)
    Call CloseRecordset(rstPrinterRates)
End Sub
Private Sub ClearFields()
    MhDateInput1.Value = Date
    MhDateInput3.Value = DateAdd("d", 2, MhDateInput1.Value)
    Combo3.ListIndex = 0                'Processing
    Text4.Text = ""
    Text3.Text = ""                     'Ref.No.
    Text7.Text = ""
    MhRealInput1.Value = 0              'Actual Quantity
    MhRealInput2.Value = 0              'Billing Quantity (Single Color)
    MhRealInput19.Value = 0             'Billing Quantity (Double & Four)
    Combo1.ListIndex = 0                'Printing Type
    MhRealInput15.Value = 0             'Pages
    MhRealInput17.Value = 0             'Qtr Form
    MhRealInput20.Value = 0             'Half Form
    MhRealInput21.Value = 0             'Full Form
    Combo2.ListIndex = 0                    'Plate Type
    MhRealInput22.Value = 1#          'Forms/Sheet For Printing Purpose
    MhRealInput3.Value = 0              'Total Plates-¼F
    MhRealInput23.Value = 0             'Total Plates-½F
    MhRealInput24.Value = 0             'Total Plates-1F
    MhRealInput59.Value = 0             'Revised Plates
    MhRealInput4.Value = 0#             'Plate Rate
    MhRealInput7.Value = 0#             'Plate Amount
    chkPaper.Value = 1                   'Paper By Party
    MhRealInput6.Value = 0              'Total Forms-¼F
    MhRealInput25.Value = 0             'Total Forms-½F
    MhRealInput26.Value = 0             'Total Forms-1F
    MhRealInput5.Value = 0#             'Print Rate
    MhRealInput8.Value = 0#             'Print Amount
    MhRealInput14.Value = 0#            'GST %
    MhRealInput18.Value = 0#            'GST Amount
    MhRealInput9.Value = 0#             'Adjustment
    MhRealInput34.Value = 0#            'Plate Adjustment
    MhRealInput36.Value = 0#            'Paper Adjustment
    MhRealInput35.Value = 0#            'Total Amount (Plate)
    MhRealInput39.Value = 0#            'Total Amount (Paper)
    MhRealInput10.Value = 0#            'Total Plate Amount
    Text1.Text = ""                                    'Paper Name
    MhRealInput27.Value = 1#            'Forms/Sheet For Paper Purpose
    MhRealInput11.Value = 0#            'Paper Wastage (in %)
    MhRealInput31.Value = 0#            'Paper Wastage (Min)
    MhRealInput32.Value = 0#            'Paper Rate
    MhRealInput33.Value = 0#            'Paper Amount
    MhRealInput12.Value = 0#            'Paper Consumption
    MhRealInput13.Value = 0#            'Total Paper Consumption
    Text8.Text = ""                                     'Bill No.
    MhDateInput2.Text = "  -  -    "        'Bill Date
    MhRealInput16.Value = 0#    'Bill Amount
    Text10.Text = ""    'Plate Bill No.
    MhDateInput4.Text = "  -  -    "    'Plate Bill Date
    MhRealInput30.Value = 0#    'Bill Amount
    Text6.Text = ""                     'Remarks
    MhRealInput28.Value = 0
    MhRealInput37.Value = 0
    MhRealInput29.Value = 0
    MhRealInput38.Value = 0
    TxtAdNar.Text = ""
'    RefCode = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    PaperCode = "": SizeCode = "": PlateMakerCode = ""
End Sub
Private Sub LoadFields()
    Dim Cnt As Integer
    If rstBookPOChild05.RecordCount = 0 Then Exit Sub
    MhDateInput1.Value = rstBookPOChild05.Fields("OrderDate").Value
    MhDateInput3.Text = Format(rstBookPOChild05.Fields("TargetDate").Value, "dd-MM-yyyy")
    Combo3.ListIndex = IIf(rstBookPOChild05.Fields("Processing").Value = "O", 0, IIf(rstBookPOChild05.Fields("Processing").Value = "N", 1, 2))
    Text3.Text = rstBookPOChild05.Fields("Ref").Value
'    RefCode = rstBookPOChild05.Fields("Ref").Value
'    If Right(FrmBookPrintOrder.BookPOType, 1) = "S" Then
'        Text3.Text = RefCode
'    Else
'        If rstRefList.RecordCount > 0 Then rstRefList.MoveFirst
'        rstRefList.Find "[Code] = '" & RefCode & "'"
'        If Not rstRefList.EOF Then Text3.Text = Trim(rstRefList.Fields("Name").Value)
'    End If
    PlateMakerCode = rstBookPOChild05.Fields("PlateMaker").Value
    If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
    rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
    If Not rstPlateMakerList.EOF Then Text7.Text = Trim(rstPlateMakerList.Fields("Col0").Value)
    MhRealInput1.Text = Format(Val(rstBookPOChild05.Fields("ActualQuantity").Value), "0")
    MhRealInput2.Text = Format(Val(rstBookPOChild05.Fields("BillingQuantity01").Value), "0")
    MhRealInput19.Text = Format(Val(rstBookPOChild05.Fields("BillingQuantity02").Value), "0")
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.SetText 1, Cnt, Val(rstBookPOChild05.Fields("Pages" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 2, Cnt, Val(rstBookPOChild05.Fields("Forms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 3, Cnt, Val(rstBookPOChild05.Fields("Forms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value)
        fpSpread1.SetText 4, Cnt, Val(rstBookPOChild05.Fields("Forms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value)
        fpSpread1.SetText 5, Cnt, Val(rstBookPOChild05.Fields("Forms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value)
        fpSpread1.SetText 6, Cnt, IIf(rstBookPOChild05.Fields("PlateType" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = "1", "Deepatch", IIf(rstBookPOChild05.Fields("PlateType" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = "2", "PS", IIf(rstBookPOChild05.Fields("PlateType" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = "3", "Wipeon", "CTP")))
        fpSpread1.SetText 7, Cnt, Val(rstBookPOChild05.Fields("PlateAmount" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 8, Cnt, Val(rstBookPOChild05.Fields("PrintAmount" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 9, Cnt, Val(rstBookPOChild05.Fields("PaperWastage" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "%").Value)
        fpSpread1.SetText 24, Cnt, Val(rstBookPOChild05.Fields("PaperWastageMin" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 25, Cnt, Val(rstBookPOChild05.Fields("PaperRate" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 26, Cnt, Val(rstBookPOChild05.Fields("PaperAmount" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 10, Cnt, Val(rstBookPOChild05.Fields("PaperConsumptionOther" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 11, Cnt, Val(rstBookPOChild05.Fields("TotalPlates" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value)
        fpSpread1.SetText 12, Cnt, Val(rstBookPOChild05.Fields("TotalPlates" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value)
        fpSpread1.SetText 13, Cnt, Val(rstBookPOChild05.Fields("TotalPlates" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value)
        fpSpread1.SetText 14, Cnt, Val(rstBookPOChild05.Fields("PlateRate" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 15, Cnt, Val(rstBookPOChild05.Fields("TotalForms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value)
        fpSpread1.SetText 16, Cnt, Val(rstBookPOChild05.Fields("TotalForms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value)
        fpSpread1.SetText 17, Cnt, Val(rstBookPOChild05.Fields("TotalForms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value)
        fpSpread1.SetText 18, Cnt, Val(rstBookPOChild05.Fields("PrintRate" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
        rstPaperList.Find "[Code] = '" & rstBookPOChild05.Fields("Paper" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value & "'"
        If Not rstPaperList.EOF Then
            fpSpread1.SetText 19, Cnt, rstPaperList.Fields("Col0").Value
            fpSpread1.SetText 28, Cnt, Val(rstPaperList.Fields("SPU").Value)
            fpSpread1.SetText 29, Cnt, Val(rstPaperList.Fields("Wt").Value)
        End If
        fpSpread1.SetText 20, Cnt, rstBookPOChild05.Fields("Paper" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value
        fpSpread1.SetText 21, Cnt, Val(rstBookPOChild05.Fields("Forms/Sheet1-" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 22, Cnt, Val(rstBookPOChild05.Fields("Forms/Sheet2-" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 23, Cnt, rstBookPOChild05.Fields("Size" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value
        fpSpread1.SetText 30, Cnt, Val(rstBookPOChild05.Fields("PaperWastageFinal" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 31, Cnt, IIf(rstBookPOChild05.Fields("PaperByParty" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value, 1, 0)
        fpSpread1.SetText 32, Cnt, Val(rstBookPOChild05.Fields("CutOffSize" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
    Next
    MhRealInput14.Text = Format(Val(rstBookPOChild05.Fields("VAT%").Value), "0.00")
    MhRealInput18.Text = Format(Val(rstBookPOChild05.Fields("VAT").Value), "0.00")
    MhRealInput28.Value = rstBookPOChild05.Fields("PVAT%").Value
    MhRealInput37.Value = rstBookPOChild05.Fields("RVAT%").Value
    MhRealInput29.Value = rstBookPOChild05.Fields("PVAT").Value
    MhRealInput38.Value = rstBookPOChild05.Fields("RVAT").Value
    MhRealInput9.Text = Format(Val(rstBookPOChild05.Fields("Adjustment").Value), "0.00")
    MhRealInput34.Text = Format(Val(rstBookPOChild05.Fields("PAdjustment").Value), "0.00")
    MhRealInput36.Value = Val(rstBookPOChild05.Fields("RAdjustment").Value)
    MhRealInput35.Text = Format(Val(rstBookPOChild05.Fields("PBillAmount").Value), "0.00")
    MhRealInput39.Text = Format(Val(rstBookPOChild05.Fields("RBillAmount").Value), "0.00")
    MhRealInput10.Text = Format(Val(rstBookPOChild05.Fields("BillAmount").Value), "0.00")
    MhRealInput13.Text = Format(Val(rstBookPOChild05.Fields("TotalPaperConsumption").Value), "0.000")
    Text8.Text = rstBookPOChild05.Fields("BillNo").Value
    If Not IsNull(rstBookPOChild05.Fields("BillDate").Value) Then MhDateInput2.Text = Format(rstBookPOChild05.Fields("BillDate").Value, "dd-MM-yyyy")
    MhRealInput16.Text = Format(Val(rstBookPOChild05.Fields("PaidAmount").Value), "0.00")
    Text10.Text = rstBookPOChild05.Fields("PBillNo").Value
    If Not IsNull(rstBookPOChild05.Fields("PBillDate").Value) Then MhDateInput4.Text = Format(rstBookPOChild05.Fields("PBillDate").Value, "dd-MM-yyyy")
    MhRealInput30.Text = Format(Val(rstBookPOChild05.Fields("PPaidAmount").Value), "0.00")
    Text6.Text = rstBookPOChild05.Fields("Remarks").Value
    TxtAdNar.Text = rstBookPOChild05.Fields("AdjustmentRemarks").Value
End Sub
Private Sub SaveFields()
    Dim Cnt As Integer, Content As Variant
    rstBookPOChild05.Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
    rstBookPOChild05.Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
    rstBookPOChild05.Fields("Processing").Value = IIf(Combo3.ListIndex = 0, "O", IIf(Combo3.ListIndex = 1, "N", "R"))
'    rstBookPOChild05.Fields("Ref").Value = RefCode
    rstBookPOChild05.Fields("Ref").Value = Text3.Text
    rstBookPOChild05.Fields("PlateMaker").Value = PlateMakerCode
    rstBookPOChild05.Fields("ActualQuantity").Value = Val(MhRealInput1.Text)
    rstBookPOChild05.Fields("BillingQuantity01").Value = Val(MhRealInput2.Text)
    rstBookPOChild05.Fields("BillingQuantity02").Value = Val(MhRealInput19.Text)
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 1, Cnt, Content
        rstBookPOChild05.Fields("Pages" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 2, Cnt, Content
        rstBookPOChild05.Fields("Forms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 3, Cnt, Content
        rstBookPOChild05.Fields("Forms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value = Val(Content)
        fpSpread1.GetText 4, Cnt, Content
        rstBookPOChild05.Fields("Forms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value = Val(Content)
        fpSpread1.GetText 5, Cnt, Content
        rstBookPOChild05.Fields("Forms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value = Val(Content)
        fpSpread1.GetText 6, Cnt, Content
        rstBookPOChild05.Fields("PlateType" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = IIf(Content = "Deepatch", "1", IIf(Content = "PS", "2", IIf(Content = "Wipeon", "3", "4")))
        fpSpread1.GetText 7, Cnt, Content
        rstBookPOChild05.Fields("PlateAmount" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 8, Cnt, Content
        rstBookPOChild05.Fields("PrintAmount" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 9, Cnt, Content
        rstBookPOChild05.Fields("PaperWastage" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "%").Value = Val(Content)
        fpSpread1.GetText 24, Cnt, Content
        rstBookPOChild05.Fields("PaperWastageMin" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 25, Cnt, Content
        rstBookPOChild05.Fields("PaperRate" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 26, Cnt, Content
        rstBookPOChild05.Fields("PaperAmount" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 10, Cnt, Content
        rstBookPOChild05.Fields("PaperConsumptionOther" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 28, Cnt, SPU
        rstBookPOChild05.Fields("PaperConsumptionSheets" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Int(Val(Content)) * Val(SPU) + (Val(Content) - Int(Val(Content))) * 1000
        fpSpread1.GetText 11, Cnt, Content
        rstBookPOChild05.Fields("TotalPlates" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value = Val(Content)
        fpSpread1.GetText 12, Cnt, Content
        rstBookPOChild05.Fields("TotalPlates" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value = Val(Content)
        fpSpread1.GetText 13, Cnt, Content
        rstBookPOChild05.Fields("TotalPlates" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value = Val(Content)
        fpSpread1.GetText 27, Cnt, Content
        rstBookPOChild05.Fields("RevisedPlates" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 14, Cnt, Content
        rstBookPOChild05.Fields("PlateRate" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 15, Cnt, Content
        rstBookPOChild05.Fields("TotalForms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value = Val(Content)
        fpSpread1.GetText 16, Cnt, Content
        rstBookPOChild05.Fields("TotalForms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value = Val(Content)
        fpSpread1.GetText 17, Cnt, Content
        rstBookPOChild05.Fields("TotalForms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value = Val(Content)
        fpSpread1.GetText 18, Cnt, Content
        rstBookPOChild05.Fields("PrintRate" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 20, Cnt, Content
        rstBookPOChild05.Fields("Paper" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Content
        fpSpread1.GetText 21, Cnt, Content
        rstBookPOChild05.Fields("Forms/Sheet1-" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 22, Cnt, Content
        rstBookPOChild05.Fields("Forms/Sheet2-" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 23, Cnt, Content
        rstBookPOChild05.Fields("Size" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Content
        fpSpread1.GetText 30, Cnt, Content
        rstBookPOChild05.Fields("PaperWastageFinal" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 31, Cnt, Content
        rstBookPOChild05.Fields("PaperByParty" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 32, Cnt, Content
        rstBookPOChild05.Fields("CutOffSize" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
    Next
    rstBookPOChild05.Fields("VAT%").Value = Format(Val(MhRealInput14.Text), "0.00")
    rstBookPOChild05.Fields("VAT").Value = Format(Val(MhRealInput18.Text), "0.00")
    rstBookPOChild05.Fields("PVAT%").Value = MhRealInput28.Value
    rstBookPOChild05.Fields("RVAT%").Value = MhRealInput37.Value
    rstBookPOChild05.Fields("PVAT").Value = MhRealInput29.Value
    rstBookPOChild05.Fields("RVAT").Value = MhRealInput38.Value
    rstBookPOChild05.Fields("Adjustment").Value = Format(Val(MhRealInput9.Text), "0.00")
    rstBookPOChild05.Fields("PAdjustment").Value = Format(Val(MhRealInput34.Text), "0.00")
    rstBookPOChild05.Fields("RAdjustment").Value = Format(Val(MhRealInput36.Text), "0.00")
    rstBookPOChild05.Fields("PBillAmount").Value = Format(Val(MhRealInput35.Text), "0.00")
    rstBookPOChild05.Fields("RBillAmount").Value = Format(Val(MhRealInput39.Text), "0.00")
    rstBookPOChild05.Fields("BillAmount").Value = Format(Val(MhRealInput10.Text), "0.00")
    rstBookPOChild05.Fields("TotalPaperConsumption").Value = Format(Val(MhRealInput13.Text), "0.000")
    rstBookPOChild05.Fields("BillNo").Value = Text8.Text
    If MhDateInput2.ValueIsNull Then rstBookPOChild05.Fields("BillDate").Value = Null Else rstBookPOChild05.Fields("BillDate").Value = GetDate(MhDateInput2.Text)
    rstBookPOChild05.Fields("PaidAmount").Value = Format(Val(MhRealInput16.Text), "0.00")
    rstBookPOChild05.Fields("PBillNo").Value = Text10.Text
    If Not IsDate(MhDateInput4.Text) Then rstBookPOChild05.Fields("PBillDate").Value = Null Else rstBookPOChild05.Fields("PBillDate").Value = GetDate(MhDateInput4.Text)
    rstBookPOChild05.Fields("PPaidAmount").Value = MhRealInput30.Value
    rstBookPOChild05.Fields("Remarks").Value = Text6.Text
    rstBookPOChild05.Fields("AdjustmentRemarks").Value = IIf(MhRealInput9.Value <> 0 Or MhRealInput34.Value <> 0 Or MhRealInput36.Value <> 0, TxtAdNar.Text, "")
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild05.Fields("BillFeedDate").Value) Then rstBookPOChild05.Fields("BillFeedDate").Value = Now()
    Dim lpBuff As String * 1024
    GetComputerName lpBuff, Len(lpBuff)
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild05.Fields("ComputerName").Value) Then rstBookPOChild05.Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf Val(CheckNull(rstBookPOChild05.Fields("ActualQuantity").Value)) = 0 Then
        MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
    End If
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If MhDateInput2.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput2.Text)) Then
        Cancel = True
'    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput2.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
'        Cancel = True
    End If
End Sub
Private Sub MhDateInput3_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput3.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput3.Text), "yyyymmdd") <= Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
        DisplayError ("Target Date cann't be prior to Order Date")
        MhDateInput3.SetFocus
        Cancel = True
    End If
End Sub
Private Sub Combo3_Validate(Cancel As Boolean)
    If Combo3.ListIndex = -1 Then Cancel = True
    If Combo3.ListIndex <= 1 Then MhRealInput59.Value = 0: MhRealInput59_Validate False: MhRealInput59.Enabled = True Else MhRealInput59.Enabled = True    'Old/New
    If Combo3.ListIndex = 0 Then MhRealInput4.Value = 0: MhRealInput4_Validate False    'Old
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
        If Not CheckEmpty(SizeCode, False) Then LoadMasterList: fpSpread1.SetText 23, fpSpread1.ActiveRow, SizeCode: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        SizeCode = "": Text4.Text = ""
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    Dim Ups As Integer
    If CheckEmpty(Text4.Text, False) Then
        Cancel = True
    Else
        Ups = CalUps()
        If Ups = 0 Then Cancel = True: Exit Sub
        If Ups > 0 Then
            If Ups <> MhRealInput27.Value And MhRealInput27.Value <> 0 Then
                If MsgBox("Calculated [Ups/Sheet] are different from existing Ups ! Change Ups?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput27.Value = Ups: MhRealInput27_Validate False
            Else
                MhRealInput27.Value = Ups: MhRealInput27_Validate False
            End If
        End If
    End If
End Sub
'Private Sub Text3_Change()
'    If Right(FrmBookPrintOrder.BookPOType, 1) = "S" Then Exit Sub
'    If Text3.Text = " " Then
'        Text3.Text = "?": SendKeys "{TAB}"
'    ElseIf CheckEmpty(Text3, False) Then
'        RefCode = ""
'    End If
'End Sub
'Private Sub Text3_Validate(Cancel As Boolean)
'    If Right(FrmBookPrintOrder.BookPOType, 1) = "S" Then RefCode = Text3.Text: Exit Sub
'    Dim SearchString As String
'    If CheckEmpty(Text3, False) Then Exit Sub
'    SearchString = FixQuote(Text3.Text)
'    If rstRefList.RecordCount = 0 Then DisplayError ("No Pending Reference"): Cancel = True: Exit Sub Else rstRefList.MoveFirst
'    rstRefList.Find "[Name] = '" & Pad(Trim(SearchString), Space(1), 10, "L") & "'"
'    If rstRefList.EOF Then
'        SelectionType = "S"
'        RefCode = ""
'        Call LoadSelectionList(rstRefList, "List of References...", "Ref.No.")
'        SearchOrder = 0
'        Call DisplaySelectionList(Text3, RefCode)
'        Call CloseForm(FrmSelectionList)
'        If CheckEmpty(Text3.Text, False) Then Text3.Text = "?"
'        If RTrim(RefCode) <> "" Then SendKeys "{TAB}"
'        Cancel = True
'    Else
'        RefCode = rstRefList.Fields("Code").Value
'        Text3.Text = Trim(rstRefList.Fields("Name").Value)
'        If Val(CheckNull(rstBookPOChild05.Fields("ActualQuantity").Value)) = 0 Then
'            MhRealInput1.Text = Format(Val(rstRefList.Fields("BalanceQuantity").Value), "0")
'            CalculateAQD
'        End If
'    End If
'End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = PlateMakerCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        PlateMakerCode = slCode: Text7.Text = slName
        If Not CheckEmpty(PlateMakerCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        Text5.Text = "": PlateMakerCode = ""
    End If
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
    If CheckEmpty(Text7.Text, False) Then Cancel = True
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    If Val(MhRealInput1.Text) <= 0 Then Cancel = True Else CalculateAQD
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    If Val(MhRealInput2.Text) <= 0 Then
        Cancel = True
'    ElseIf Val(MhRealInput2.Text) Mod 1000 <> 0 Then
        MhRealInput2.SetFocus
        Cancel = True
'    Else
        CalculateBQD (fpSpread1.ActiveRow)
    End If
End Sub
Private Sub MhRealInput19_Validate(Cancel As Boolean)
    If Val(MhRealInput2.Text) <= 0 Then
        Cancel = True
'    ElseIf Val(MhRealInput19.Text) Mod 1000 <> 0 Then
        MhRealInput19.SetFocus
        Cancel = True
'    Else
        CalculateBQD (fpSpread1.ActiveRow)
    End If
End Sub
Private Sub MhRealInput15_Validate(Cancel As Boolean)
    Dim Pages As Variant, Forms As Double
    With fpSpread1
        Pages = MhRealInput15.Value
        .SetText 1, .ActiveRow, Pages
        If Val(Pages) > 0 Then
            Forms = Val(Pages) / Choose(Val(FrmBookPrintOrder.rstBookList.Fields("FormType").Value), 8, 16, 4, 12, 24, 32, 64, 6, 2)
            .SetText 2, .ActiveRow, Val(Forms)
            If MhRealInput21.Value = 0 Then MhRealInput21.Value = (Int(Forms / 2) * 2) + (Int(Forms) - Int(Forms / 2) * 2): .SetText 5, .ActiveRow, MhRealInput21.Value: MhRealInput21_Validate False
            Forms = Forms - Int(Forms)
            If MhRealInput17.Value = 0 Then MhRealInput17.Value = IIf(Forms = 0.25, 1, IIf(Forms = 0.75, 1, IIf(Forms = 0.375, 1, IIf(Forms = 0.875, 1, 0)))): .SetText 3, .ActiveRow, MhRealInput17.Value: MhRealInput17_Validate False
            If MhRealInput20.Value = 0 Then MhRealInput20.Value = IIf(Forms = 0.5, 1, IIf(Forms = 0.75, 1, IIf(Forms = 0.625, 1, IIf(Forms = 0.875, 1, IIf(Forms = (5 / 6), 1, 0))))): .SetText 4, .ActiveRow, MhRealInput20.Value: MhRealInput20_Validate False
        Else
            MhRealInput17.Value = 0
            MhRealInput17_Validate False
            MhRealInput20.Value = 0
            MhRealInput20_Validate False
            MhRealInput21.Value = 0
            MhRealInput21_Validate False
            .SetText 3, .ActiveRow, 0
            .SetText 4, .ActiveRow, 0
            .SetText 5, .ActiveRow, 0
        End If
    End With
End Sub
Private Sub MhRealInput17_Validate(Cancel As Boolean)   '¼ Forms
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant
    With fpSpread1
        .SetText 3, .ActiveRow, Val(MhRealInput17.Text)
        Call CalculateTotalPlates(IIf(.ActiveRow = 1, "1", IIf(.ActiveRow = 2, "2", "4")), Val(MhRealInput17.Text), "¼")
        Call CalculateTotalForms(IIf(.ActiveRow = 1, "1", IIf(.ActiveRow = 2, "2", "4")), Val(MhRealInput17.Text), "¼")
        CalculateAmount
        Call CalculateConsumption(IIf(.ActiveRow = 1, "1", IIf(.ActiveRow = 2, "2", "4")))
        .GetText 3, .ActiveRow, Forms¼
        .GetText 4, .ActiveRow, Forms½
        .GetText 5, .ActiveRow, Forms1
        .SetText 2, .ActiveRow, Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1
    End With
End Sub
Private Sub MhRealInput20_Validate(Cancel As Boolean)   '½ Forms
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant
    With fpSpread1
        .SetText 4, .ActiveRow, Val(MhRealInput20.Text)
        Call CalculateTotalPlates(IIf(.ActiveRow = 1, "1", IIf(.ActiveRow = 2, "2", "4")), Val(MhRealInput20.Text), "½")
        Call CalculateTotalForms(IIf(.ActiveRow = 1, "1", IIf(.ActiveRow = 2, "2", "4")), Val(MhRealInput20.Text), "½")
        CalculateAmount
        Call CalculateConsumption(IIf(.ActiveRow = 1, "1", IIf(.ActiveRow = 2, "2", "4")))
        .GetText 3, .ActiveRow, Forms¼
        .GetText 4, .ActiveRow, Forms½
        .GetText 5, .ActiveRow, Forms1
        .SetText 2, .ActiveRow, Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1
    End With
End Sub
Private Sub MhRealInput21_Validate(Cancel As Boolean)   '1 Forms
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant
    With fpSpread1
        .SetText 5, .ActiveRow, Val(MhRealInput21.Text)
        Call CalculateTotalPlates(IIf(.ActiveRow = 1, "1", IIf(.ActiveRow = 2, "2", "4")), Val(MhRealInput21.Text), "1")
        Call CalculateTotalForms(IIf(.ActiveRow = 1, "1", IIf(.ActiveRow = 2, "2", "4")), Val(MhRealInput21.Text), "1")
        CalculateAmount
        Call CalculateConsumption(IIf(.ActiveRow = 1, "1", IIf(.ActiveRow = 2, "2", "4")))
        .GetText 3, .ActiveRow, Forms¼
        .GetText 4, .ActiveRow, Forms½
        .GetText 5, .ActiveRow, Forms1
        .SetText 2, .ActiveRow, Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1
    End With
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)  'Plate Type
    fpSpread1.SetText 6, fpSpread1.ActiveRow, IIf(Combo2.ListIndex = 0, "Deepatch", IIf(Combo2.ListIndex = 1, "PS", IIf(Combo2.ListIndex = 2, "Wipeon", "CTP")))
    GetPrinterRates IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), "L"  'Get Plate Rates
    CalculateAmount
    If Combo2.ListIndex = 1 Or Combo2.ListIndex = 3 Then    'PS/CTP Plate Details
        On Error Resume Next
        FrmPSPlateRegister.ItemCode = BookCode
        FrmPSPlateRegister.ItemName = Trim(Text2.Text)
        FrmPSPlateRegister.OrderCode = IIf(CheckNull(rstBookPOChild05.Fields("Code").Value) = "", "999999", rstBookPOChild05.Fields("Code").Value)
        FrmPSPlateRegister.OrderDate = GetDate(MhDateInput1.Text)
        FrmPSPlateRegister.TblSuffix = "05"
        FrmPSPlateRegister.OrderType = VchType
        FrmPSPlateRegister.PlateType = IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4"))
        Load FrmPSPlateRegister
        If Err.Number <> 364 Then FrmPSPlateRegister.Show vbModal
        On Error GoTo 0
    End If
End Sub
Private Sub MhRealInput22_Validate(Cancel As Boolean)   'Forms/Sheet For Printing Purpose
    Dim Forms As Variant
    If Val(MhRealInput22.Text) <> 0.5 And Val(MhRealInput22.Text) <> 1 And Val(MhRealInput22.Text) <> 2 Then
        Cancel = True
    Else
        fpSpread1.SetText 21, fpSpread1.ActiveRow, Val(MhRealInput22.Text)
        fpSpread1.GetText 3, fpSpread1.ActiveRow, Forms   '¼ Forms
        Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "¼")
        Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "¼")
        fpSpread1.GetText 4, fpSpread1.ActiveRow, Forms   '½ Forms
        Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "½")
        Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "½")
        fpSpread1.GetText 5, fpSpread1.ActiveRow, Forms   '1 Forms
        Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "1")
        Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "1")
        CalculateAmount
    End If
    Dim Ups As Integer
    Ups = CalUps()
    If Ups = 0 Then Cancel = True: Exit Sub
    If Ups > 0 Then
        If Ups <> MhRealInput27.Value And MhRealInput27.Value <> 0 Then
            If MsgBox("Calculated [Ups/Sheet] are different from existing Ups ! Change Ups?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput27.Value = Ups: MhRealInput27_Validate False
        Else
            MhRealInput27.Value = Ups: MhRealInput27_Validate False
        End If
    End If
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)    'Plate Rate
    fpSpread1.SetText 14, fpSpread1.ActiveRow, Val(MhRealInput4.Text)
    CalculateAmount
End Sub
Private Sub MhRealInput5_Validate(Cancel As Boolean)    'Print Rate
    fpSpread1.SetText 18, fpSpread1.ActiveRow, Val(MhRealInput5.Text)
    CalculateAmount
End Sub
Private Sub MhRealInput14_Validate(Cancel As Boolean)   'VAT
    CalculateTotalAmount
End Sub
Private Sub MhRealInput28_Validate(Cancel As Boolean)   'PVAT%
    CalculateTotalAmount
End Sub
Private Sub MhRealInput37_Validate(Cancel As Boolean)   'RVAT%
    CalculateTotalAmount
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean)    'Adjustment
    CalculateTotalAmount
End Sub
Private Sub MhRealInput34_Validate(Cancel As Boolean)   'Plate Adjustment
    CalculateTotalAmount
End Sub
Private Sub MhRealInput36_Validate(Cancel As Boolean)   'Paper Adjustment
    CalculateTotalAmount
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
        With fpSpread1
            .SetText 19, .ActiveRow, "" 'Paper Name
            .SetText 20, .ActiveRow, "" 'Paper Code
            .SetText 28, .ActiveRow, "" 'SPU
            .SetText 29, .ActiveRow, "" 'Wt
        End With
    End If
End Sub
Private Sub Text1_Validate(Cancel As Boolean)
    Dim Ups As Integer
    If CheckEmpty(Text1, False) Then
        If Val(MhRealInput8.Text) = 0 Then 'Print Amount
            PaperCode = ""
            With fpSpread1
                .SetText 19, .ActiveRow, "" 'Paper Name
                .SetText 20, .ActiveRow, "" 'Paper Code
                .SetText 28, .ActiveRow, "" 'SPU
                .SetText 29, .ActiveRow, "" 'Wt
            End With
        Else
            Cancel = True
        End If
        Exit Sub
    End If
    fpSpread1.SetText 19, fpSpread1.ActiveRow, Trim(Text1.Text) 'Paper Name
    rstPaperList.Filter = adFilterNone
    If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
    rstPaperList.Find "[Col0] LIKE '" & Text1.Text & "%'"
    PaperCode = rstPaperList.Fields("Code").Value
    fpSpread1.SetText 20, fpSpread1.ActiveRow, PaperCode
    fpSpread1.SetText 28, fpSpread1.ActiveRow, Val(rstPaperList.Fields("SPU").Value)
    fpSpread1.SetText 29, fpSpread1.ActiveRow, Val(rstPaperList.Fields("Wt").Value)
    If rstPaperList.Fields("Form").Value = "R" Then
        fpSpread1.GetText 32, fpSpread1.ActiveRow, CutOffSize
        Do While True
            CutOffSize = InputBox("Reel Cut Off (mm)", "Easy Publish", Val(CutOffSize))
            If Val(CutOffSize) = 0 Then DisplayError ("Reel Cut off Size cann't be zero"): Cancel = True Else Exit Do
        Loop
        fpSpread1.SetText 32, fpSpread1.ActiveRow, CutOffSize
    Else
        fpSpread1.SetText 32, fpSpread1.ActiveRow, 0
    End If
    Ups = CalUps()
    If Ups = 0 Then
        Cancel = True
    ElseIf Ups > 0 Then
        If Ups <> MhRealInput27.Value And MhRealInput27.Value <> 0 Then
            If MsgBox("Calculated [Ups/Sheet] are different from existing Ups ! Change Ups?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput27.Value = Ups: MhRealInput27_Validate False
        Else
            MhRealInput27.Value = Ups: MhRealInput27_Validate False
        End If
    End If
End Sub
Private Sub MhRealInput27_Validate(Cancel As Boolean)   'Forms/Sheet For Paper Purpose
'    If Val(MhRealInput27.Text) <> 0.5 And Val(MhRealInput27.Text) <> 1 And Val(MhRealInput27.Text) <> 2 Then
'        Cancel = True
'    Else
    fpSpread1.SetText 22, fpSpread1.ActiveRow, Val(MhRealInput27.Text)
    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
'    End If
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)   'Paper Wastage Rate
    fpSpread1.SetText 9, fpSpread1.ActiveRow, Val(MhRealInput11.Text)
    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
End Sub
Private Sub MhRealInput31_Validate(Cancel As Boolean)   'Paper Wastage Minimum
    fpSpread1.SetText 24, fpSpread1.ActiveRow, Val(MhRealInput31.Text)
    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
End Sub
Private Sub MhRealInput32_Validate(Cancel As Boolean)
    With fpSpread1
        .GetText 28, .ActiveRow, SPU: .GetText 29, .ActiveRow, Wt
        If Val(SPU) = 0 Then SPU = 500
        MhRealInput33.Value = MhRealInput32.Value * Val(Wt) * ((Int(MhRealInput12.Value) * Val(SPU) + (MhRealInput12.Value - Int(MhRealInput12.Value)) * 1000) / Val(SPU))
        .SetText 25, .ActiveRow, Val(MhRealInput32.Text)
        .SetText 26, .ActiveRow, Val(MhRealInput33.Text)
    End With
    CalculateTotalAmount
End Sub
Private Sub fpSpread1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    fpSpread1.SetActiveCell 1, NewRow
    fpSpread1_DblClick 1, NewRow
    Text4.SetFocus
End Sub
Private Sub fpSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim Content As Variant
    Combo1.ListIndex = IIf(Row = 1, 0, IIf(Row = 2, 1, 2))  'Printing Type
    fpSpread1.GetText 1, Row, Content   'Pages
    MhRealInput15.Text = Format(Val(Content), "0")
    fpSpread1.GetText 3, Row, Content   '¼ F
    MhRealInput17.Text = Format(Val(Content), "0")
    fpSpread1.GetText 4, Row, Content   '½ F
    MhRealInput20.Text = Format(Val(Content), "0")
    fpSpread1.GetText 5, Row, Content   '1 F
    MhRealInput21.Text = Format(Val(Content), "0")
    fpSpread1.GetText 6, Row, Content   'Plate Type
    Combo2.ListIndex = IIf(Content = "Deepatch", 0, IIf(Content = "PS", 1, IIf(Content = "Wipeon", 2, 3)))
    fpSpread1.GetText 7, Row, Content   'Plate Amount
    MhRealInput7.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 8, Row, Content   'Print Amount
    MhRealInput8.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 9, Row, Content   'Paper Wastage
    MhRealInput11.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 24, Row, Content   'Paper Wastage Min
    MhRealInput31.Text = Format(Val(Content), "0")
    fpSpread1.GetText 25, Row, Content   'Paper Rate
    MhRealInput32.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 26, Row, Content   'Paper Amount
    MhRealInput33.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 10, Row, Content   'Paper Consumption (Reams)
    MhRealInput12.Text = Format(Val(Content), "0.000")
    fpSpread1.GetText 11, Row, Content   'Total Plates - ¼F
    MhRealInput3.Text = Format(Val(Content), "0")
    fpSpread1.GetText 12, Row, Content   'Total Plates - ½F
    MhRealInput23.Text = Format(Val(Content), "0")
    fpSpread1.GetText 13, Row, Content   'Total Plates - 1F
    MhRealInput24.Text = Format(Val(Content), "0")
    fpSpread1.GetText 27, Row, Content   'Revised Plates
    MhRealInput59.Text = Format(Val(Content), "0")
    fpSpread1.GetText 14, Row, Content   'Plate Rate
    MhRealInput4.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 15, Row, Content   'Total Forms - ¼F
    MhRealInput6.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 16, Row, Content   'Total Forms - ½F
    MhRealInput25.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 17, Row, Content   'Total Forms - 1F
    MhRealInput26.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 18, Row, Content   'Print Rate
    MhRealInput5.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 19, Row, Content   'Paper Name
    Text1.Text = Content
    fpSpread1.GetText 21, Row, Content   'Forms/Sheet - For Printing Purpose
    MhRealInput22.Text = Format(IIf(Val(Content) = 0, 1, Val(Content)), "0.00")
    fpSpread1.GetText 22, Row, Content   'Forms/Sheet - For Paper Purpose
    MhRealInput27.Text = Format(IIf(Val(Content) = 0, 1, Val(Content)), "0.00")
    fpSpread1.GetText 23, Row, Content   'Size Code
    If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
    SizeCode = Content
    rstSizeList.Find "[Code] = '" & SizeCode & "'"
    If Not rstSizeList.EOF Then Text4.Text = rstSizeList.Fields("Col0").Value Else Text4.Text = ""
    fpSpread1.GetText 31, Row, Content   'Paper by party
    chkPaper.Value = Val(Content)
End Sub
'Private Sub LoadRefList(ByVal strBookCode As String, ByVal strOrderCode As String)
'    Dim BalanceQuantity As Long
'    On Error GoTo ErrorHandler
'    If rstRefList.State = adStateOpen Then rstRefList.Close
'    If DatabaseType = "MS SQL" Then
'        rstRefList.Open "SELECT P.Name,Quantity As PlannedQuantity,(SELECT SUM(ActualQuantity) FROM BookPOChild05 C1 INNER JOIN BookPOParent P1 ON P1.Code=C1.Code WHERE C1.Ref=P.Code AND P1.Book=C.Book And C1.Code<>'" & strOrderCode & "') As PrintedQuantity,Quantity As BalanceQuantity,Remarks As Col0,[PaperWastage%],P.Code FROM PrintPVParent P INNER JOIN PrintPVChild C ON P.Code=C.Code WHERE P.PlanningType ='1' And C.Book='" & strBookCode & "' ORDER BY P.Name", cnDatabase, adOpenKeyset, adLockOptimistic
'    Else
'        rstRefList.Open "Select P.Name,Quantity As PlannedQuantity,Format((Select Sum(ActualQuantity) From BookPOChild05,BookPOParent Where BookPOChild05.Ref=P.Code And BookPOParent.Code=BookPOChild05.Code And BookPOParent.Book=C.Book And BookPOChild05.Code<>'" & strOrderCode & "'),0) As PrintedQuantity,Quantity As BalanceQuantity,Remarks As Col0,[PaperWastage%],P.Code From PrintPVParent P,PrintPVChild C Where P.Code=C.Code And P.PlanningType ='1' And C.Book='" & strBookCode & "' Order By P.Name", cnDatabase, adOpenKeyset, adLockOptimistic
'    End If
'    rstRefList.ActiveConnection = Nothing
'    Do While Not rstRefList.EOF
'        BalanceQuantity = (Val(CheckNull(rstRefList.Fields("PlannedQuantity").Value)) - Val(CheckNull(rstRefList.Fields("PrintedQuantity").Value)))
'        If BalanceQuantity <> 0 Then
'            rstRefList.Fields("Col0").Value = Trim(rstRefList.Fields("Name").Value) + " Quantity : " + Format(Str(BalanceQuantity), "#0")
'            rstRefList.Fields("BalanceQuantity").Value = BalanceQuantity
'            rstRefList.Update
'        Else
'            rstRefList.Delete
'        End If
'        rstRefList.MoveNext
'    Loop
'    Exit Sub
'ErrorHandler:
'    DisplayError ("Failed to Load Ref List")
'End Sub
Private Sub GetPrinterRates(ByVal xPrintingType As String, ByVal xRateType As String)   'xRateType : 'B'-Both Plate & Print Rate 'L'-Only Plate Rate
    On Error GoTo ErrorHandler
    Dim PrintRate As Double, PlateRate As Double, PaperWastageRate As Double, PaperWastageMin As Integer, CurrentRate As Variant, PlateType As Variant, i As Integer
    i = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) 'Active Row
    'Fetch Data
    If rstPrinterRates.State = adStateOpen Then rstPrinterRates.Close
    rstPrinterRates.Open "SELECT TOP 1 * FROM AccountChild05 WHERE Code = '" & PrinterCode & "' And [Size]=(SELECT Code FROM SizeGroupChild WHERE [Size]='" & SizeCode & "') And Range" & Trim(xPrintingType) & " >= " & IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) & " ORDER BY Range" & Trim(xPrintingType), cnDatabase, adOpenKeyset, adLockReadOnly
    If rstPrinterRates.RecordCount = 0 Then
        If rstPrinterRates.State = adStateOpen Then rstPrinterRates.Close
        rstPrinterRates.Open "SELECT TOP 1 * FROM AccountMaster P INNER JOIN AccountChild05 C ON P.Code=C.Code WHERE Name Like '%Rate%' And [Size]=(SELECT Code FROM SizeGroupChild WHERE [Size]='" & SizeCode & "') AND Range" & Trim(xPrintingType) & " >= " & IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) & " ORDER BY Range" & Trim(xPrintingType), cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstPrinterRates.RecordCount > 0 Then
        fpSpread1.GetText 6, i, PlateType
        PlateRate = IIf(Combo3.ListIndex = 0, 0, Val(rstPrinterRates.Fields(PlateType & "PlateRate" & Trim(xPrintingType)).Value))
        PrintRate = rstPrinterRates.Fields("PrintRate" & Trim(xPrintingType)).Value
        PrintRate = PrintRate + IIf(PrintRate > 0, Val(FrmBookPrintOrder.rstBookList.Fields("AddOnRate01").Value), 0)
        PaperWastageRate = Val(rstPrinterRates.Fields("PaperWastageRate" & Trim(xPrintingType)))
        PaperWastageMin = Val(rstPrinterRates.Fields("PaperWastageMin" & Trim(xPrintingType)))
    End If
    'Plate Rate
    fpSpread1.GetText 14, i, CurrentRate
    If CurrentRate <> PlateRate And CurrentRate <> 0 Then
        If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Plate rate is different from that in Master ! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then fpSpread1.SetText 14, i, PlateRate
    Else
        fpSpread1.SetText 14, i, PlateRate
    End If
    If xRateType = "B" Then
        'Print Rate
        fpSpread1.GetText 18, i, CurrentRate  'Print Rate
        If CurrentRate <> PrintRate And CurrentRate <> 0 Then
            If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Printing Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then fpSpread1.SetText 18, i, PrintRate
        Else
            fpSpread1.SetText 18, i, PrintRate
        End If
        'Paper Wastage Rate
        fpSpread1.GetText 9, i, CurrentRate
        If CurrentRate <> PaperWastageRate And CurrentRate <> 0 Then
            If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Paper Wastage is different from that in Master ! Change Wastage?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then fpSpread1.SetText 9, i, PaperWastageRate
        Else
            fpSpread1.SetText 9, i, PaperWastageRate
        End If
        'Paper Wastage Min
        fpSpread1.GetText 24, i, CurrentRate
        If CurrentRate <> PaperWastageMin And CurrentRate <> 0 Then
            If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Paper Wastage (Min) is different from that in Master ! Change Wastage?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then fpSpread1.SetText 24, i, PaperWastageMin
        Else
            fpSpread1.SetText 24, i, PaperWastageMin
        End If
    End If
    If fpSpread1.ActiveRow = i Then
        fpSpread1.GetText 14, fpSpread1.ActiveRow, CurrentRate: MhRealInput4.Value = CurrentRate    'Plate Rate
        fpSpread1.GetText 18, fpSpread1.ActiveRow, CurrentRate: MhRealInput5.Text = Format(CurrentRate, "0.00") 'Print Rate
        fpSpread1.GetText 9, fpSpread1.ActiveRow, CurrentRate: MhRealInput11.Text = Format(CurrentRate, "0.00") 'Paper Wastage Rate
        fpSpread1.GetText 24, fpSpread1.ActiveRow, CurrentRate: MhRealInput31.Text = Format(CurrentRate, "0")   'Paper Wastage Min
    End If
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub
Private Sub CalculateAQD()   'Calculate Actual Quantity Dependents
    Dim Q1 As Double, Q24 As Double
    
    'For Single Color : Actual Quantity = Billing Quantity + 10 % of Billing Quantity + 99
    Q1 = Val(MhRealInput1.Text) * 100 / (10 + 100) 'Mod 1000
    Q1 = IIf(Val(MhRealInput1.Text) > 99 And Q1 > 0 And Int(Q1) <= 90, Val(MhRealInput1.Text) - 99, Val(MhRealInput1.Text))  'New Actual Quantity
    Q1 = Int(Q1 * 100 / (10 + 100) / 1000) * 1000 + IIf(Q1 * 100 / (10 + 100) Mod 1000 = 0, 0, 1000)
    'For Double/Four Color : Actual Quantity = Billing Quantity - 200 + 99 OR Actual Quantity = Billing Quantity - 500 + 99
    Q24 = IIf(Int(Val(MhRealInput1.Text) / 1000) = 0, 1000, Int(Val(MhRealInput1.Text) / 1000) * 1000) + IIf(Val(MhRealInput1.Text) Mod 1000 <= IIf(Val(MhRealInput1.Text) <= 10000, 299, 599), 0, 1000)
    If Val(MhRealInput2.Text) = 0 Then
        MhRealInput2.Text = Val(MhRealInput1.Text) 'Format(Q1, "0")
    ElseIf Val(MhRealInput2.Text) <> Q1 Then
        If MsgBox("Variation (Single Color) between Billing Quantity (" & MhRealInput2.Text & ") Vs Calculated Billing Quantity (" & Trim(Str(Q1)) & ") ! Change Quantity ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
            MhRealInput2.Text = Format(Q1, "0")
        End If
    End If
    If Val(MhRealInput19.Text) = 0 Then
        MhRealInput19.Text = Val(MhRealInput1.Text) 'Q24
    ElseIf Val(MhRealInput19.Text) <> Q24 Then
        If MsgBox("Variation (Double & Four Color) between Billing Quantity (" & MhRealInput19.Text & ") Vs Calculated Billing Quantity (" & Trim(Str(Q24)) & ") ! Change Quantity ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
            MhRealInput19.Text = Format(Q24, "0")
        End If
    End If
    CalculateBQD (fpSpread1.ActiveRow)
    Call CalculateConsumption("1"): Call CalculateConsumption("2"): Call CalculateConsumption("4")
End Sub
Private Sub CalculateBQD(ByVal xPrintingType As Integer)    'Calculate Billing Quantity Dependents
    Dim Content As Variant, Forms As Variant
    fpSpread1.GetText 1, xPrintingType, Content   'Pages
    If Val(Content) <> 0 Then GetPrinterRates IIf(xPrintingType = 1, "1", IIf(xPrintingType = 2, "2", "4")), "B"             'Get Both Plate & Printing Rates
    fpSpread1.GetText 3, xPrintingType, Forms
    Call CalculateTotalForms(IIf(xPrintingType = 1, "1", IIf(xPrintingType = 2, "2", "4")), Val(Forms), "¼")
    fpSpread1.GetText 4, xPrintingType, Forms
    Call CalculateTotalForms(IIf(xPrintingType = 1, "1", IIf(xPrintingType = 2, "2", "4")), Val(Forms), "½")
    fpSpread1.GetText 5, xPrintingType, Forms
    Call CalculateTotalForms(IIf(xPrintingType = 1, "1", IIf(xPrintingType = 2, "2", "4")), Val(Forms), "1")
    CalculateAmount
End Sub
Private Function CalculateConsumption(ByVal xPrintingType As String) As Double
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant, WastageRate As Variant, WastageMin As Variant, CurrentPaperConsumption As Variant, Cnt As Integer, FS As Variant, W As Double, Forms As Double
    fpSpread1.GetText 3, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms¼
    fpSpread1.GetText 4, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms½
    fpSpread1.GetText 5, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms1
    fpSpread1.GetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), WastageRate
    fpSpread1.GetText 24, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), WastageMin
    fpSpread1.GetText 28, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), SPU
    If Val(SPU) = 0 Then SPU = 500
    Forms = (Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1)
    If Forms = 0 Then Exit Function
    CalculateConsumption = (MhRealInput1.Value / 1000) * Forms  'Consumption (in Reams)
    CalculateConsumption = (CalculateConsumption * 500) / Val(SPU)
    W = (((MhRealInput1.Value / 1000) * 1 * Val(WastageRate)) / 100) * 500 'Wastage for Single Form (in Sheets)
    If W < Val(WastageMin) Then W = Val(WastageMin) 'Comparison with Minimum Wastage
    W = (W * Forms) / Val(SPU)  'Wastage for Total Forms (in Reams)
    CalculateConsumption = (CalculateConsumption + W) * Val(SPU) 'Consumption With Wastage (in Sheets)
    fpSpread1.GetText 22, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateConsumption = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * CalculateConsumption
    CalculateConsumption = Format(CLng(Int(Val(CalculateConsumption) / Val(SPU))) + ((Val(CalculateConsumption) Mod Val(SPU)) / 1000), "0.000")
    'Wasage Final
    W = W * Val(SPU)    'Wastage (in Sheets)
    W = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * W
    fpSpread1.SetText 30, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Format(CLng(Int(Val(W) / Val(SPU))) + ((Val(W) Mod Val(SPU)) / 1000), "0.000")
    fpSpread1.SetText 10, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateConsumption
    If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then MhRealInput12.Value = Val(CalculateConsumption)
    'Total Consumption Calculation
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 10, Cnt, CurrentPaperConsumption
        MhRealInput13.Text = Format(IIf(Cnt = 1, 0, Val(MhRealInput13.Text)) + CLng((Int(Val(CurrentPaperConsumption)) * Val(SPU)) + ((Val(CurrentPaperConsumption) - Int(Val(CurrentPaperConsumption))) * 1000)), "0.000")
    Next
    MhRealInput13.Text = Format(CLng(Int(Val(MhRealInput13.Text) / Val(SPU))) + ((Val(MhRealInput13.Text) Mod Val(SPU)) / 1000), "0.000")
End Function
Private Sub CalculateAmount()
    Dim Cnt As Integer, TotalPlates¼ As Variant, TotalPlates½ As Variant, TotalPlates1 As Variant, RevisedPlates As Variant, PlateRate As Variant, TotalForms¼ As Variant, TotalForms½ As Variant, TotalForms1 As Variant, PrintRate As Variant
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 11, Cnt, TotalPlates¼
        fpSpread1.GetText 12, Cnt, TotalPlates½
        fpSpread1.GetText 13, Cnt, TotalPlates1
        fpSpread1.GetText 27, Cnt, RevisedPlates
        fpSpread1.GetText 14, Cnt, PlateRate
        fpSpread1.GetText 15, Cnt, TotalForms¼
        fpSpread1.GetText 16, Cnt, TotalForms½
        fpSpread1.GetText 17, Cnt, TotalForms1
        fpSpread1.GetText 18, Cnt, PrintRate
        If Val(RevisedPlates) = 0 Then
            fpSpread1.SetText 7, Cnt, IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalPlates¼) + Val(TotalPlates½) + Val(TotalPlates1)) * Val(PlateRate)
        Else
            fpSpread1.SetText 7, Cnt, IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * Val(RevisedPlates) * Val(PlateRate)
        End If
        fpSpread1.SetText 8, Cnt, IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalForms¼) + Val(TotalForms½) + Val(TotalForms1)) * Val(PrintRate)
        If fpSpread1.ActiveRow = Cnt Then
            If Val(RevisedPlates) = 0 Then
                MhRealInput7.Text = Format(IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalPlates¼) + Val(TotalPlates½) + Val(TotalPlates1)) * Val(PlateRate), "0.00")  'Plate Amount
            Else
                MhRealInput7.Text = Format(IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * Val(RevisedPlates) * Val(PlateRate), "0.00")  'Plate Amount
            End If
            MhRealInput8.Text = Format(IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalForms¼) + Val(TotalForms½) + Val(TotalForms1)) * Val(PrintRate), "0.00")     'Print Amount
        End If
    Next
    CalculateTotalAmount
End Sub
Private Function CalculateTotalAmount() As Double
    Dim i As Integer, PlateAmount As Variant, PrintAmount As Variant, PaperAmount As Variant, TotalPrintAmount As Double, TotalPlateAmount As Double, TotalPaperAmount As Variant
    With fpSpread1
        For i = 1 To .MaxRows
            .GetText 7, i, PlateAmount: .GetText 8, i, PrintAmount: .GetText 26, i, PaperAmount
            TotalPlateAmount = TotalPlateAmount + PlateAmount: TotalPrintAmount = TotalPrintAmount + PrintAmount: TotalPaperAmount = TotalPaperAmount + PaperAmount
        Next
    End With
    MhRealInput29.Value = (TotalPlateAmount + MhRealInput34.Value) * MhRealInput28.Value / 100      'GST Plate
    MhRealInput18.Value = (TotalPrintAmount + MhRealInput9.Value) * MhRealInput14.Value / 100       'GST Printing
    MhRealInput38.Value = (TotalPaperAmount + MhRealInput36.Value) * MhRealInput37.Value / 100      'GST Paper
    MhRealInput10.Value = Round(TotalPrintAmount + MhRealInput9.Value + MhRealInput18.Value, 0)     'Total Printing Amount
    MhRealInput35.Value = Round(TotalPlateAmount + MhRealInput29.Value + MhRealInput34.Value, 0)    'Total Plate Amount
    MhRealInput39.Value = Round(TotalPaperAmount + MhRealInput38.Value + MhRealInput36.Value, 0)    'Total Paper Amount
End Function
Private Function CalculateTotalForms(ByVal xPrintingType As String, ByVal Forms As Double, ByVal FormType As String) As Double
    Dim FS As Variant
    fpSpread1.GetText 21, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateTotalForms = IIf(((IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) / 1000) - (Int(IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) / 1000))) = 0, (Int(IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) * IIf(FormType = "¼", 0.25, IIf(FormType = "½", 0.5, 1)) / 1000) + IIf(IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) * IIf(FormType = "¼", 0.25, IIf(FormType = "½", 0.5, 1)) Mod 1000 = 0, 0, 1)) * Forms, (IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) * IIf(FormType = "¼", 0.25, IIf(FormType = "½", 0.5, 1)) / 1000) * Forms)
    CalculateTotalForms = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * Val(CalculateTotalForms)
    If FrmBookPrintOrder.rstBookList.Fields("DuplexPrinting").Value = "N" Then CalculateTotalForms = 0.5 * CalculateTotalForms
    CalculateTotalForms = IIf(((IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) / 1000) - (Int(IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) / 1000))) = 0, Int(Val(CalculateTotalForms)) + IIf(Val(CalculateTotalForms) - Int(Val(CalculateTotalForms)) = 0, 0, 1), (Val(CalculateTotalForms)) + IIf(Val(CalculateTotalForms) - Int(Val(CalculateTotalForms)) = 0, 0, 0))
    If FormType = "¼" Then
        fpSpread1.SetText 15, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then MhRealInput6.Text = Format(CalculateTotalForms, "0.00")
    ElseIf FormType = "½" Then
        fpSpread1.SetText 16, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then MhRealInput25.Text = Format(CalculateTotalForms, "0.00")
    Else
        fpSpread1.SetText 17, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then MhRealInput26.Text = Format(CalculateTotalForms, "0.00")
    End If
End Function
Private Function CalculateTotalPlates(ByVal xPrintingType As String, ByVal Forms As Double, ByVal FormType As String) As Double
    Dim FS As Variant
    fpSpread1.GetText 21, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateTotalPlates = Forms
    CalculateTotalPlates = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * Val(CalculateTotalPlates)
    If FrmBookPrintOrder.rstBookList.Fields("DuplexPrinting").Value = "N" Then CalculateTotalPlates = 0.5 * CalculateTotalPlates
    CalculateTotalPlates = Int(Val(CalculateTotalPlates)) + IIf(Val(CalculateTotalPlates) - Int(Val(CalculateTotalPlates)) = 0.5, 1, 0)
    If FormType = "¼" Then
        fpSpread1.SetText 11, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalPlates
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            MhRealInput3.Text = Format(CalculateTotalPlates, "0")
        End If
    ElseIf FormType = "½" Then
        fpSpread1.SetText 12, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalPlates
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            MhRealInput23.Text = Format(CalculateTotalPlates, "0")
        End If
    Else
        fpSpread1.SetText 13, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalPlates
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            MhRealInput24.Text = Format(CalculateTotalPlates, "0")
        End If
    End If
End Function
Private Sub cmdProceed_Click()
    Dim Cnt As Integer, PaperBalance As Double, PaperCode As Variant, PaperName As Variant, PaperStock As Variant, PaperConsumption As Variant, PaperRate As Variant, PaperByParty As Variant, VchDate As Date
    VchDate = MhDateInput1
    If CheckMandatoryFields Then Exit Sub
    If Combo3.ListIndex = 0 Then MhRealInput59.Value = 0: MhRealInput59_Validate False
    If Left(FrmBookPrintOrder.BookPOType, 1) <> "O" Then
        For Cnt = 1 To fpSpread1.MaxRows
            fpSpread1.SetActiveCell 1, Cnt
            fpSpread1_DblClick 1, Cnt
            fpSpread1.GetText 19, Cnt, PaperName
            fpSpread1.GetText 20, Cnt, PaperCode
            fpSpread1.GetText 25, Cnt, PaperRate
            If Not CheckEmpty(PaperCode, False) Then
                fpSpread1.GetText 28, Cnt, SPU
                fpSpread1.GetText 31, Cnt, PaperByParty
                fpSpread1.GetText 10, Cnt, PaperConsumption: PaperConsumption = Fix(Val(PaperConsumption)) * Val(SPU) + Round(Val(PaperConsumption) - Fix(Val(PaperConsumption)), 3) * 1000
                PaperStock = CalculatePaperBalance(IIf(PaperByParty = "1", PrinterCode, "000000"), PaperCode, CheckNull(rstBookPOChild05.Fields("Code").Value), "PO", VchDate): PaperStock = Fix(Val(PaperStock)) * Val(SPU) + Round(Val(PaperStock) - Fix(Val(PaperStock)), 3) * 1000
                PaperBalance = PaperStock - PaperConsumption
                If PaperBalance < 0 Then
                    PaperBalance = Format(CLng(Fix(0 - Abs(PaperBalance) / Val(SPU))) + ((0 - Abs(PaperBalance) Mod Val(SPU)) / 1000), "0.000")
                    If UserLevel <= 2 Then
                        If MsgBox("Stock (" & Format(PaperBalance, "0.000") & ") of the Paper - " & Trim(PaperName) & vbCrLf & " is going negative ! Would you like to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbNo Then Exit Sub
                    Else
                        Call DisplayError("Cann't Save ! Stock (" & Format(PaperBalance, "0.000") & ") of the Paper - " & Trim(PaperName) & " is going negative"): AbortPO = True: Exit Sub
                    End If
                End If
            End If
        Next
    End If
    SaveFields
    rstBookPOChild05.Update
    Call CloseForm(Me)
End Sub
Private Function CheckMandatoryFields() As Boolean
    Dim Cnt As Integer, Pages As Variant, Paper As Variant, Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant, TotalForms As Variant
    If Combo2.ListIndex < 0 Then Combo2.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo3.ListIndex < 0 Then Combo3.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput16.Text) <> 0 Then If Val(MhRealInput16.Text) <> Val(MhRealInput10.Text) + Val(MhRealInput39.Text) Then MhRealInput9.SetFocus: CheckMandatoryFields = True: Exit Function
    'If Val(MhRealInput32.Text) = 0 And chkPaper.Value = 0 Then MhRealInput32.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput30.Text) <> 0 Then If Val(MhRealInput30.Text) <> Val(MhRealInput35.Text) Then MhRealInput34.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput9.Value <> 0 Or MhRealInput34.Value <> 0 Or MhRealInput36.Value <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.SetActiveCell 1, Cnt
        fpSpread1_DblClick 1, Cnt
        fpSpread1.GetText 1, Cnt, Pages
        fpSpread1.GetText 20, Cnt, Paper
        If Pages <> 0 Then
            If CheckNull(Paper) = "" Then
                Text4.SetFocus
                CheckMandatoryFields = True
                Exit For
            End If
        End If
        fpSpread1.GetText 2, Cnt, TotalForms
        fpSpread1.GetText 3, Cnt, Forms¼
        fpSpread1.GetText 4, Cnt, Forms½
        fpSpread1.GetText 5, Cnt, Forms1
        If Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1 <> TotalForms Then
            DisplayError ("Variation between Total Forms Vs Bifurcated Forms")
            MhRealInput17.SetFocus
            CheckMandatoryFields = True
            Exit For
        End If
    Next
End Function
Private Sub cmdCancel_Click()
    rstBookPOChild05.CancelUpdate
    Call CloseForm(Me)
End Sub
Private Sub MhRealInput59_Validate(Cancel As Boolean)
    fpSpread1.SetText 27, fpSpread1.ActiveRow, MhRealInput59.Value
    CalculateAmount
End Sub
Private Function CalUps() As Integer
        If CheckEmpty(PaperCode, False) Or CheckEmpty(SizeCode, False) Then CalUps = -1: Exit Function
        Dim FL As Double, FR As Double, PL As Double, PW As Double
        rstPaperList.MoveFirst
        rstPaperList.Find "[Code]='" & PaperCode & "'"
        fpSpread1.GetText 32, fpSpread1.ActiveRow, CutOffSize
        FL = Val(Left(Text4.Text, InStr(1, Text4.Text, "X") - 1)): FR = Val(Mid(Text4.Text, InStr(1, Text4.Text, "X") + 1)) 'Printing Size Left & Right
        PL = IIf(rstPaperList.Fields("Form").Value = "R", Val(CutOffSize) / 25.4, Val(rstPaperList.Fields("inLength").Value)): PW = Val(rstPaperList.Fields("inWidth").Value) 'Paper Area Length & Width
        If Abs(FL - PL) <= 1 Then PL = FL
        If Abs(FR - PL) <= 1 Then PL = FR
        If Abs(FL - PW) <= 1 Then PW = FL
        If Abs(FR - PW) <= 1 Then PW = FR
        CalUps = CalcUps(PL * PW, FL * FR)
End Function
Private Sub LoadMasterList(Optional ByVal LoadSelected As Boolean)
    If rstSizeList.State = adStateOpen Then rstSizeList.Close
    rstSizeList.Open "SELECT Name As Col0, Code From GeneralMaster Where Type = '1' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstSizeList.ActiveConnection = Nothing
    If rstPaperList.State = adStateOpen Then rstPaperList.Close
    If LoadSelected Then
        rstPaperList.Open "SELECT * FROM (SELECT LTRIM(P.Name)+' (UOM : '+LTRIM(C.Name)+'='+LTRIM(C.Value1)+')' As Col0,FORMAT(dbo.ufnGetPaperStock('" & IIf(chkPaper.Value, PrinterCode, "000000") & "',P.Code,'PO','" & CheckNull(rstBookPOChild05.Fields("Code").Value) & "','" & GetDate(MhDateInput1.Text) & "'),'#0.000') As Col1,C.Name As UOM,GSM,inWidth,inLength,P.Code,C.Value1 As SPU,[Form],[Weight/Unit] As Wt,LTRIM(Q.Name) As Quality,Grade FROM (PaperMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code) INNER JOIN GeneralMaster Q ON P.Quality=Q.Code) As Tbl WHERE CONVERT(DECIMAL(12,3),Col1)<>0 ORDER BY Col0", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstPaperList.Open "SELECT LTRIM(P.Name)+' (UOM : '+LTRIM(C.Name)+'='+LTRIM(C.Value1)+')' As Col0,FORMAT(0,'#0.000') As Col1,C.Name As UOM,GSM,inWidth,inLength,P.Code,C.Value1 As SPU,[Form],[Weight/Unit] As Wt,LTRIM(Q.Name) As Quality,Grade FROM (PaperMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code) INNER JOIN GeneralMaster Q ON P.Quality=Q.Code ORDER BY Col0", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    rstPaperList.ActiveConnection = Nothing
    If rstPlateMakerList.State = adStateOpen Then rstPlateMakerList.Close
    rstPlateMakerList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstPlateMakerList.ActiveConnection = Nothing
End Sub
Private Sub chkPaper_Click()
    fpSpread1.SetText 31, fpSpread1.ActiveRow, Val(chkPaper.Value)
End Sub
Private Sub chkPaper_Validate(Cancel As Boolean)
'            If MhRealInput32.Value = 0 Then
'               MhRealInput32.Value = 0         'Paper Rate Must Be Zero
            If MhRealInput32.Value <> 0 And chkPaper.Value = 1 Then
               If MsgBox("Paper By Party Selected  So That Paper Rate Rs.(" & MhRealInput32.Text & ") Must Be Zero  ! Change Rate ?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Change !") = vbYes Then MhRealInput32.Value = 0
            ElseIf MhRealInput32.Value = 0 And chkPaper.Value = 0 Then
               If MsgBox("Paper Not Supplied By Party Selected  So That Paper Rate Rs.(" & MhRealInput32.Text & ") Must Be Required  ! Change Rate ?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Change !") = vbYes Then MhRealInput32.Value = 0
            End If
'    fpSpread1.SetText 31, fpSpread1.ActiveRow, Val(chkPaper.Value)
'    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
End Sub
