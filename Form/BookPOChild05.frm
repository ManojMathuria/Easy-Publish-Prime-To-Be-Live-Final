VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmBookPOChild05 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Form Format Order Details"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
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
   ScaleHeight     =   10095
   ScaleWidth      =   11640
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   11160
      Picture         =   "BookPOChild05.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   9860
      Left            =   120
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   120
      Width           =   10935
      _Version        =   65536
      _ExtentX        =   19288
      _ExtentY        =   17392
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
      Picture         =   "BookPOChild05.frx":0102
      Begin VB.TextBox Text16 
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
         Left            =   9240
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   13
         Top             =   1590
         Width           =   1575
      End
      Begin VB.TextBox Text13 
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
         TabIndex        =   15
         Top             =   2425
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   290
         Index           =   4
         Left            =   1380
         Picture         =   "BookPOChild05.frx":011E
         Style           =   1  'Graphical
         TabIndex        =   131
         TabStop         =   0   'False
         ToolTipText     =   "Cancel [Esc]"
         Top             =   5280
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   290
         Index           =   3
         Left            =   1065
         Picture         =   "BookPOChild05.frx":0220
         Style           =   1  'Graphical
         TabIndex        =   130
         TabStop         =   0   'False
         ToolTipText     =   "Save [Ctrl+S]"
         Top             =   5280
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   290
         Index           =   2
         Left            =   750
         Picture         =   "BookPOChild05.frx":0322
         Style           =   1  'Graphical
         TabIndex        =   129
         TabStop         =   0   'False
         ToolTipText     =   "Delete [Ctrl+D]"
         Top             =   5280
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   290
         Index           =   1
         Left            =   435
         Picture         =   "BookPOChild05.frx":0424
         Style           =   1  'Graphical
         TabIndex        =   128
         TabStop         =   0   'False
         ToolTipText     =   "Edit [Ctrl+E]"
         Top             =   5280
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   290
         Index           =   0
         Left            =   120
         Picture         =   "BookPOChild05.frx":0956
         Style           =   1  'Graphical
         TabIndex        =   127
         TabStop         =   0   'False
         ToolTipText     =   "Add [Ctrl+A]"
         Top             =   5280
         Width           =   315
      End
      Begin VB.TextBox Text14 
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
         TabIndex        =   6
         Top             =   960
         Width           =   5655
      End
      Begin VB.TextBox Text11 
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
         Left            =   9240
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   7
         Top             =   960
         Width           =   1575
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
         Left            =   9240
         MaxLength       =   40
         TabIndex        =   5
         Top             =   645
         Width           =   1575
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
         Left            =   9240
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1275
         Width           =   1575
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput36 
         Height          =   330
         Left            =   1800
         TabIndex        =   59
         ToolTipText     =   "Plate"
         Top             =   7740
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":0E88
         Caption         =   "BookPOChild05.frx":0EA8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":0F14
         Keys            =   "BookPOChild05.frx":0F32
         Spin            =   "BookPOChild05.frx":0F7C
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
         Left            =   7560
         TabIndex        =   111
         Top             =   7430
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
         Picture         =   "BookPOChild05.frx":0FA4
         Picture         =   "BookPOChild05.frx":0FC0
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
         TabIndex        =   66
         Top             =   8580
         Width           =   2775
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
         MaxLength       =   60
         TabIndex        =   9
         Top             =   1275
         Width           =   5655
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
         TabIndex        =   10
         Top             =   1590
         Width           =   5655
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
         TabIndex        =   70
         Top             =   9420
         Width           =   9015
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
         TabIndex        =   69
         Top             =   9105
         Width           =   9015
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   110
         Width           =   2775
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
         TabIndex        =   63
         Top             =   8260
         Width           =   2775
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
         TabIndex        =   40
         Top             =   3900
         Width           =   5655
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   645
         Width           =   5655
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   4560
         TabIndex        =   73
         Top             =   105
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Picture         =   "BookPOChild05.frx":0FDC
         Picture         =   "BookPOChild05.frx":0FF8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   74
         Top             =   1905
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
         Picture         =   "BookPOChild05.frx":1014
         Picture         =   "BookPOChild05.frx":1030
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   4560
         TabIndex        =   75
         Top             =   1905
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Picture         =   "BookPOChild05.frx":104C
         Picture         =   "BookPOChild05.frx":1068
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   76
         Top             =   2425
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
         Caption         =   " Printing Color"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":1084
         Picture         =   "BookPOChild05.frx":10A0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   77
         Top             =   3065
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
         Picture         =   "BookPOChild05.frx":10BC
         Picture         =   "BookPOChild05.frx":10D8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   4560
         TabIndex        =   78
         Top             =   3060
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Picture         =   "BookPOChild05.frx":10F4
         Picture         =   "BookPOChild05.frx":1110
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   4560
         TabIndex        =   79
         Top             =   3375
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Picture         =   "BookPOChild05.frx":112C
         Picture         =   "BookPOChild05.frx":1148
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   7440
         TabIndex        =   80
         Top             =   645
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
         Caption         =   " Ref No."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":1164
         Picture         =   "BookPOChild05.frx":1180
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   120
         TabIndex        =   81
         Top             =   2750
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
         Caption         =   " Pages && Forms"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":119C
         Picture         =   "BookPOChild05.frx":11B8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   82
         Top             =   3380
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
         Picture         =   "BookPOChild05.frx":11D4
         Picture         =   "BookPOChild05.frx":11F0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   7440
         TabIndex        =   83
         Top             =   3065
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
         Caption         =   " Plate Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":120C
         Picture         =   "BookPOChild05.frx":1228
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   7440
         TabIndex        =   84
         Top             =   3380
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
         Caption         =   " Print Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":1244
         Picture         =   "BookPOChild05.frx":1260
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   7560
         TabIndex        =   85
         Top             =   7110
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
         Picture         =   "BookPOChild05.frx":127C
         Picture         =   "BookPOChild05.frx":1298
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   120
         TabIndex        =   86
         Top             =   3900
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
         Picture         =   "BookPOChild05.frx":12B4
         Picture         =   "BookPOChild05.frx":12D0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Left            =   4560
         TabIndex        =   87
         Top             =   4215
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Picture         =   "BookPOChild05.frx":12EC
         Picture         =   "BookPOChild05.frx":1308
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Left            =   4560
         TabIndex        =   88
         Top             =   4535
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Picture         =   "BookPOChild05.frx":1324
         Picture         =   "BookPOChild05.frx":1340
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
         Height          =   330
         Left            =   7440
         TabIndex        =   89
         Top             =   4535
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
         Caption         =   " Consumption-UOM"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":135C
         Picture         =   "BookPOChild05.frx":1378
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   90
         Top             =   8260
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
         Picture         =   "BookPOChild05.frx":1394
         Picture         =   "BookPOChild05.frx":13B0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   7560
         TabIndex        =   91
         Top             =   8260
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
         Picture         =   "BookPOChild05.frx":13CC
         Picture         =   "BookPOChild05.frx":13E8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   4560
         TabIndex        =   92
         Top             =   8260
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Picture         =   "BookPOChild05.frx":1404
         Picture         =   "BookPOChild05.frx":1420
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   7440
         TabIndex        =   93
         Top             =   110
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
         Caption         =   " Target Date"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":143C
         Picture         =   "BookPOChild05.frx":1458
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   94
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
         Picture         =   "BookPOChild05.frx":1474
         Picture         =   "BookPOChild05.frx":1490
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Left            =   7440
         TabIndex        =   95
         Top             =   1590
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
         Caption         =   " Plate Type"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":14AC
         Picture         =   "BookPOChild05.frx":14C8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   96
         Top             =   110
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
         Picture         =   "BookPOChild05.frx":14E4
         Picture         =   "BookPOChild05.frx":1500
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   97
         Top             =   9105
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
         Picture         =   "BookPOChild05.frx":151C
         Picture         =   "BookPOChild05.frx":1538
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
         Height          =   330
         Left            =   3480
         TabIndex        =   98
         Top             =   7110
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
         Caption         =   " GST"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":1554
         Picture         =   "BookPOChild05.frx":1570
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   6000
         TabIndex        =   2
         Top             =   105
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "BookPOChild05.frx":158C
         Caption         =   "BookPOChild05.frx":16A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1710
         Keys            =   "BookPOChild05.frx":172E
         Spin            =   "BookPOChild05.frx":178C
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
         Left            =   9240
         TabIndex        =   3
         Top             =   110
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild05.frx":17B4
         Caption         =   "BookPOChild05.frx":18CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1938
         Keys            =   "BookPOChild05.frx":1956
         Spin            =   "BookPOChild05.frx":19B4
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
         Left            =   6000
         TabIndex        =   64
         Top             =   8260
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild05.frx":19DC
         Caption         =   "BookPOChild05.frx":1AF4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1B60
         Keys            =   "BookPOChild05.frx":1B7E
         Spin            =   "BookPOChild05.frx":1BDC
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel31 
         Height          =   330
         Left            =   120
         TabIndex        =   99
         Top             =   4215
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
         Caption         =   " Ups/Sheet"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":1C04
         Picture         =   "BookPOChild05.frx":1C20
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   1800
         TabIndex        =   11
         Top             =   1905
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":1C3C
         Caption         =   "BookPOChild05.frx":1C5C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1CC8
         Keys            =   "BookPOChild05.frx":1CE6
         Spin            =   "BookPOChild05.frx":1D30
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
         ValueVT         =   1638405
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   6000
         TabIndex        =   12
         ToolTipText     =   "One Color"
         Top             =   1910
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":1D58
         Caption         =   "BookPOChild05.frx":1D78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1DE4
         Keys            =   "BookPOChild05.frx":1E02
         Spin            =   "BookPOChild05.frx":1E4C
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
         ValueVT         =   1638405
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
         Height          =   330
         Left            =   1800
         TabIndex        =   18
         ToolTipText     =   "Pages"
         Top             =   2750
         Width           =   520
         _Version        =   65536
         _ExtentX        =   917
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":1E74
         Caption         =   "BookPOChild05.frx":1E94
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1F00
         Keys            =   "BookPOChild05.frx":1F1E
         Spin            =   "BookPOChild05.frx":1F68
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   3000
         TabIndex        =   20
         ToolTipText     =   "¼ Form"
         Top             =   2750
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":1F90
         Caption         =   "BookPOChild05.frx":1FB0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":201C
         Keys            =   "BookPOChild05.frx":203A
         Spin            =   "BookPOChild05.frx":2084
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
         Height          =   330
         Left            =   6000
         TabIndex        =   24
         Top             =   2745
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":20AC
         Caption         =   "BookPOChild05.frx":20CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2138
         Keys            =   "BookPOChild05.frx":2156
         Spin            =   "BookPOChild05.frx":21A0
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
         ValueVT         =   5
         Value           =   1
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   1800
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "¼ Form"
         Top             =   3065
         Width           =   520
         _Version        =   65536
         _ExtentX        =   917
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":21C8
         Caption         =   "BookPOChild05.frx":21E8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2254
         Keys            =   "BookPOChild05.frx":2272
         Spin            =   "BookPOChild05.frx":22BC
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput23 
         Height          =   330
         Left            =   2310
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "½ Form"
         Top             =   3065
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":22E4
         Caption         =   "BookPOChild05.frx":2304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2370
         Keys            =   "BookPOChild05.frx":238E
         Spin            =   "BookPOChild05.frx":23D8
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
         Height          =   330
         Left            =   3000
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "1 Form-F&B"
         Top             =   3060
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2400
         Caption         =   "BookPOChild05.frx":2420
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":248C
         Keys            =   "BookPOChild05.frx":24AA
         Spin            =   "BookPOChild05.frx":24F4
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   6000
         TabIndex        =   32
         Top             =   3065
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":251C
         Caption         =   "BookPOChild05.frx":253C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":25A8
         Keys            =   "BookPOChild05.frx":25C6
         Spin            =   "BookPOChild05.frx":2610
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   9240
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3065
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2638
         Caption         =   "BookPOChild05.frx":2658
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":26C4
         Keys            =   "BookPOChild05.frx":26E2
         Spin            =   "BookPOChild05.frx":272C
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
         ForeColor       =   255
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   1800
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "¼ Form"
         Top             =   3375
         Width           =   520
         _Version        =   65536
         _ExtentX        =   917
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2754
         Caption         =   "BookPOChild05.frx":2774
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":27E0
         Keys            =   "BookPOChild05.frx":27FE
         Spin            =   "BookPOChild05.frx":2848
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
         Height          =   330
         Left            =   2310
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "½ Form"
         Top             =   3375
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2870
         Caption         =   "BookPOChild05.frx":2890
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":28FC
         Keys            =   "BookPOChild05.frx":291A
         Spin            =   "BookPOChild05.frx":2964
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
         Height          =   330
         Left            =   3000
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "1 Form-F&B"
         Top             =   3375
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":298C
         Caption         =   "BookPOChild05.frx":29AC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2A18
         Keys            =   "BookPOChild05.frx":2A36
         Spin            =   "BookPOChild05.frx":2A80
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   6000
         TabIndex        =   38
         Top             =   3375
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2AA8
         Caption         =   "BookPOChild05.frx":2AC8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2B34
         Keys            =   "BookPOChild05.frx":2B52
         Spin            =   "BookPOChild05.frx":2B9C
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
         Height          =   330
         Left            =   9240
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3380
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2BC4
         Caption         =   "BookPOChild05.frx":2BE4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2C50
         Keys            =   "BookPOChild05.frx":2C6E
         Spin            =   "BookPOChild05.frx":2CB8
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
         ForeColor       =   255
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
         Height          =   330
         Left            =   4320
         TabIndex        =   52
         Top             =   7110
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2CE0
         Caption         =   "BookPOChild05.frx":2D00
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2D6C
         Keys            =   "BookPOChild05.frx":2D8A
         Spin            =   "BookPOChild05.frx":2DD4
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
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
         Height          =   330
         Left            =   6000
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   7110
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2DFC
         Caption         =   "BookPOChild05.frx":2E1C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2E88
         Keys            =   "BookPOChild05.frx":2EA6
         Spin            =   "BookPOChild05.frx":2EF0
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
         ForeColor       =   255
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   330
         Left            =   1800
         TabIndex        =   51
         ToolTipText     =   "Print"
         Top             =   7110
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2F18
         Caption         =   "BookPOChild05.frx":2F38
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2FA4
         Keys            =   "BookPOChild05.frx":2FC2
         Spin            =   "BookPOChild05.frx":300C
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
         Height          =   330
         Left            =   9240
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   7110
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":3034
         Caption         =   "BookPOChild05.frx":3054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":30C0
         Keys            =   "BookPOChild05.frx":30DE
         Spin            =   "BookPOChild05.frx":3128
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
         ForeColor       =   255
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
         Height          =   330
         Left            =   1800
         TabIndex        =   42
         Top             =   4215
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":3150
         Caption         =   "BookPOChild05.frx":3170
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":31DC
         Keys            =   "BookPOChild05.frx":31FA
         Spin            =   "BookPOChild05.frx":3244
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   5880
         TabIndex        =   43
         Top             =   4215
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":326C
         Caption         =   "BookPOChild05.frx":328C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":32F8
         Keys            =   "BookPOChild05.frx":3316
         Spin            =   "BookPOChild05.frx":3360
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
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   9240
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Total Consumption (Units)"
         Top             =   4535
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":3388
         Caption         =   "BookPOChild05.frx":33A8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":3414
         Keys            =   "BookPOChild05.frx":3432
         Spin            =   "BookPOChild05.frx":347C
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
         ForeColor       =   255
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
         Left            =   9240
         TabIndex        =   65
         Top             =   8260
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":34A4
         Caption         =   "BookPOChild05.frx":34C4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":3530
         Keys            =   "BookPOChild05.frx":354E
         Spin            =   "BookPOChild05.frx":3598
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
         ValueVT         =   1638405
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Left            =   120
         TabIndex        =   100
         Top             =   9420
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
         Picture         =   "BookPOChild05.frx":35C0
         Picture         =   "BookPOChild05.frx":35DC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   3480
         TabIndex        =   101
         Top             =   7430
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
         Caption         =   " GST"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":35F8
         Picture         =   "BookPOChild05.frx":3614
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput28 
         Height          =   330
         Left            =   4320
         TabIndex        =   56
         Top             =   7430
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":3630
         Caption         =   "BookPOChild05.frx":3650
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":36BC
         Keys            =   "BookPOChild05.frx":36DA
         Spin            =   "BookPOChild05.frx":3724
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
         ValueVT         =   2009726981
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput29 
         Height          =   330
         Left            =   6000
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   7430
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":374C
         Caption         =   "BookPOChild05.frx":376C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":37D8
         Keys            =   "BookPOChild05.frx":37F6
         Spin            =   "BookPOChild05.frx":3840
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
         ForeColor       =   255
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel34 
         Height          =   330
         Left            =   120
         TabIndex        =   102
         Top             =   1590
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
         Picture         =   "BookPOChild05.frx":3868
         Picture         =   "BookPOChild05.frx":3884
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel35 
         Height          =   330
         Left            =   120
         TabIndex        =   103
         Top             =   1275
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
         Caption         =   " Print Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":38A0
         Picture         =   "BookPOChild05.frx":38BC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel36 
         Height          =   330
         Left            =   120
         TabIndex        =   104
         Top             =   8580
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
         Caption         =   " Plate Party Bill No."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":38D8
         Picture         =   "BookPOChild05.frx":38F4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel37 
         Height          =   330
         Left            =   7560
         TabIndex        =   105
         Top             =   8580
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
         Picture         =   "BookPOChild05.frx":3910
         Picture         =   "BookPOChild05.frx":392C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
         Height          =   330
         Left            =   4560
         TabIndex        =   106
         Top             =   8580
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Picture         =   "BookPOChild05.frx":3948
         Picture         =   "BookPOChild05.frx":3964
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput4 
         Height          =   330
         Left            =   6000
         TabIndex        =   67
         Top             =   8580
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild05.frx":3980
         Caption         =   "BookPOChild05.frx":3A98
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":3B04
         Keys            =   "BookPOChild05.frx":3B22
         Spin            =   "BookPOChild05.frx":3B80
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
         Left            =   9240
         TabIndex        =   68
         Top             =   8580
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":3BA8
         Caption         =   "BookPOChild05.frx":3BC8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":3C34
         Keys            =   "BookPOChild05.frx":3C52
         Spin            =   "BookPOChild05.frx":3C9C
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
         ValueVT         =   1638405
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel39 
         Height          =   330
         Left            =   120
         TabIndex        =   107
         Top             =   4845
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
         Caption         =   " Paper Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":3CC4
         Picture         =   "BookPOChild05.frx":3CE0
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput32 
         Height          =   330
         Left            =   1800
         TabIndex        =   48
         Top             =   4845
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":3CFC
         Caption         =   "BookPOChild05.frx":3D1C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":3D88
         Keys            =   "BookPOChild05.frx":3DA6
         Spin            =   "BookPOChild05.frx":3DF0
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel40 
         Height          =   330
         Left            =   4560
         TabIndex        =   108
         Top             =   4845
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Picture         =   "BookPOChild05.frx":3E18
         Picture         =   "BookPOChild05.frx":3E34
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput33 
         Height          =   330
         Left            =   6000
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   4845
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":3E50
         Caption         =   "BookPOChild05.frx":3E70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":3EDC
         Keys            =   "BookPOChild05.frx":3EFA
         Spin            =   "BookPOChild05.frx":3F44
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
         ForeColor       =   255
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput34 
         Height          =   330
         Left            =   1800
         TabIndex        =   55
         ToolTipText     =   "Plate"
         Top             =   7430
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":3F6C
         Caption         =   "BookPOChild05.frx":3F8C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":3FF8
         Keys            =   "BookPOChild05.frx":4016
         Spin            =   "BookPOChild05.frx":4060
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   120
         TabIndex        =   109
         Top             =   7110
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
         Picture         =   "BookPOChild05.frx":4088
         Picture         =   "BookPOChild05.frx":40A4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel41 
         Height          =   330
         Left            =   120
         TabIndex        =   110
         Top             =   7430
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
         Picture         =   "BookPOChild05.frx":40C0
         Picture         =   "BookPOChild05.frx":40DC
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput35 
         Height          =   330
         Left            =   9240
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   7430
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":40F8
         Caption         =   "BookPOChild05.frx":4118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":4184
         Keys            =   "BookPOChild05.frx":41A2
         Spin            =   "BookPOChild05.frx":41EC
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
         ForeColor       =   255
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel43 
         Height          =   330
         Left            =   120
         TabIndex        =   112
         Top             =   7740
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
         Picture         =   "BookPOChild05.frx":4214
         Picture         =   "BookPOChild05.frx":4230
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel44 
         Height          =   330
         Left            =   7560
         TabIndex        =   113
         Top             =   7740
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
         Picture         =   "BookPOChild05.frx":424C
         Picture         =   "BookPOChild05.frx":4268
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel45 
         Height          =   330
         Left            =   3480
         TabIndex        =   114
         Top             =   7740
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
         Caption         =   " GST"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":4284
         Picture         =   "BookPOChild05.frx":42A0
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput37 
         Height          =   330
         Left            =   4320
         TabIndex        =   60
         Top             =   7740
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":42BC
         Caption         =   "BookPOChild05.frx":42DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":4348
         Keys            =   "BookPOChild05.frx":4366
         Spin            =   "BookPOChild05.frx":43B0
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
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput38 
         Height          =   330
         Left            =   6000
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   7740
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":43D8
         Caption         =   "BookPOChild05.frx":43F8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":4464
         Keys            =   "BookPOChild05.frx":4482
         Spin            =   "BookPOChild05.frx":44CC
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
         ForeColor       =   255
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput39 
         Height          =   330
         Left            =   9240
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   7740
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":44F4
         Caption         =   "BookPOChild05.frx":4514
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":4580
         Keys            =   "BookPOChild05.frx":459E
         Spin            =   "BookPOChild05.frx":45E8
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
         ForeColor       =   255
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput59 
         Height          =   330
         Left            =   4140
         TabIndex        =   31
         ToolTipText     =   "Revised Plates"
         Top             =   3060
         Width           =   432
         _Version        =   65536
         _ExtentX        =   762
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":4610
         Caption         =   "BookPOChild05.frx":4630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":469C
         Keys            =   "BookPOChild05.frx":46BA
         Spin            =   "BookPOChild05.frx":4704
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel29 
         Height          =   330
         Left            =   7440
         TabIndex        =   115
         Top             =   1910
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
         Caption         =   " Plate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":472C
         Picture         =   "BookPOChild05.frx":4748
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel33 
         Height          =   330
         Left            =   7440
         TabIndex        =   116
         Top             =   1275
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
         Caption         =   " Printing Size"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":4764
         Picture         =   "BookPOChild05.frx":4780
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel46 
         Height          =   330
         Left            =   7440
         TabIndex        =   117
         Top             =   960
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
         Caption         =   " Finish Size"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":479C
         Picture         =   "BookPOChild05.frx":47B8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel30 
         Height          =   330
         Left            =   4560
         TabIndex        =   118
         Top             =   2745
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Picture         =   "BookPOChild05.frx":47D4
         Picture         =   "BookPOChild05.frx":47F0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel47 
         Height          =   330
         Left            =   4560
         TabIndex        =   119
         Top             =   2430
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Pages/Ptg Form"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":480C
         Picture         =   "BookPOChild05.frx":4828
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput40 
         Height          =   330
         Left            =   6000
         TabIndex        =   16
         ToolTipText     =   "One Color"
         Top             =   2430
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":4844
         Caption         =   "BookPOChild05.frx":4864
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":48D0
         Keys            =   "BookPOChild05.frx":48EE
         Spin            =   "BookPOChild05.frx":4938
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
         ValueVT         =   1638405
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel48 
         Height          =   330
         Left            =   7440
         TabIndex        =   120
         Top             =   2745
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
         Caption         =   " Duplex Printing"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":4960
         Picture         =   "BookPOChild05.frx":497C
      End
      Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
         Height          =   330
         Left            =   9240
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   2745
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
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
         Picture         =   "BookPOChild05.frx":4998
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Yes"
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
            Left            =   240
            TabIndex        =   25
            Top             =   60
            Value           =   -1  'True
            Width           =   585
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No"
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
            Left            =   915
            TabIndex        =   26
            Top             =   60
            Width           =   615
         End
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput41 
         Height          =   330
         Left            =   4140
         TabIndex        =   23
         ToolTipText     =   "1 Form-W&T"
         Top             =   2745
         Width           =   432
         _Version        =   65536
         _ExtentX        =   762
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":49B4
         Caption         =   "BookPOChild05.frx":49D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":4A40
         Keys            =   "BookPOChild05.frx":4A5E
         Spin            =   "BookPOChild05.frx":4AA8
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput46 
         Height          =   330
         Left            =   3720
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "1 Form-W&T"
         Top             =   3065
         Width           =   432
         _Version        =   65536
         _ExtentX        =   762
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":4AD0
         Caption         =   "BookPOChild05.frx":4AF0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":4B5C
         Keys            =   "BookPOChild05.frx":4B7A
         Spin            =   "BookPOChild05.frx":4BC4
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
         Height          =   330
         Left            =   3360
         TabIndex        =   21
         ToolTipText     =   "½ Form"
         Top             =   2750
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":4BEC
         Caption         =   "BookPOChild05.frx":4C0C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":4C78
         Keys            =   "BookPOChild05.frx":4C96
         Spin            =   "BookPOChild05.frx":4CE0
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
         Height          =   330
         Left            =   3720
         TabIndex        =   22
         ToolTipText     =   "1 Form-F&B"
         Top             =   2750
         Width           =   432
         _Version        =   65536
         _ExtentX        =   762
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":4D08
         Caption         =   "BookPOChild05.frx":4D28
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":4D94
         Keys            =   "BookPOChild05.frx":4DB2
         Spin            =   "BookPOChild05.frx":4DFC
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel49 
         Height          =   330
         Left            =   7440
         TabIndex        =   122
         Top             =   4845
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
         Caption         =   " Consumption-Kgs"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":4E24
         Picture         =   "BookPOChild05.frx":4E40
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput49 
         Height          =   330
         Left            =   9240
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Total Consumption (Kgs)"
         Top             =   4845
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":4E5C
         Caption         =   "BookPOChild05.frx":4E7C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":4EE8
         Keys            =   "BookPOChild05.frx":4F06
         Spin            =   "BookPOChild05.frx":4F50
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
         ForeColor       =   255
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel52 
         Height          =   330
         Left            =   120
         TabIndex        =   123
         Top             =   960
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
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
         Caption         =   " Element Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":4F78
         Picture         =   "BookPOChild05.frx":4F94
      End
      Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
         Height          =   330
         Left            =   1800
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   4535
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
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
         Picture         =   "BookPOChild05.frx":4FB0
         Begin VB.CheckBox chkPaper 
            Caption         =   "Check1"
            Height          =   210
            Left            =   1290
            TabIndex        =   45
            Top             =   80
            Width           =   210
         End
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel50 
         Height          =   330
         Left            =   7440
         TabIndex        =   125
         Top             =   4215
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
         Caption         =   " Wastage Sheet"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":4FCC
         Picture         =   "BookPOChild05.frx":4FE8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel51 
         Height          =   330
         Left            =   120
         TabIndex        =   126
         Top             =   4535
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
         Caption         =   " Paper By Party"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":5004
         Picture         =   "BookPOChild05.frx":5020
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput48 
         Height          =   330
         Left            =   6000
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   4535
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":503C
         Caption         =   "BookPOChild05.frx":505C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":50C8
         Keys            =   "BookPOChild05.frx":50E6
         Spin            =   "BookPOChild05.frx":5130
         AlignHorizontal =   1
         AlignVertical   =   2
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput47 
         Height          =   330
         Left            =   9240
         TabIndex        =   44
         ToolTipText     =   "Wastage Min(Sheets)"
         Top             =   4215
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":5158
         Caption         =   "BookPOChild05.frx":5178
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":51E4
         Keys            =   "BookPOChild05.frx":5202
         Spin            =   "BookPOChild05.frx":524C
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1365
         Left            =   120
         TabIndex        =   0
         Top             =   5565
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   2408
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   9164542
         HeadLines       =   1
         RowHeight       =   18
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "ElementName"
            Caption         =   "Element"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "FinishSizeName"
            Caption         =   "Finish Size"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "PrintSizeName"
            Caption         =   "Printing Size"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Pages/PrintingForm"
            Caption         =   "Pgs/Ptg Form"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Pages"
            Caption         =   "Pages"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Color"
            Caption         =   "ColorName"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Forms"
            Caption         =   "Forms"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Forms-¼"
            Caption         =   "¼F"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Forms-½"
            Caption         =   "½F"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Forms-1-F&B"
            Caption         =   "1F-F&B"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Forms-1-W&T"
            Caption         =   "1F-W&T"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "PaperReqd"
            Caption         =   "Paper Req"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   524.976
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   315.213
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   269.858
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   285.165
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   915.024
            EndProperty
         EndProperty
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput42 
         Height          =   330
         Left            =   9240
         TabIndex        =   17
         ToolTipText     =   "One Color"
         Top             =   2430
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":5274
         Caption         =   "BookPOChild05.frx":5294
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":5300
         Keys            =   "BookPOChild05.frx":531E
         Spin            =   "BookPOChild05.frx":5368
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
         ValueVT         =   1638405
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel53 
         Height          =   330
         Left            =   7440
         TabIndex        =   132
         Top             =   2430
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
         Caption         =   " Pages/Form"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":5390
         Picture         =   "BookPOChild05.frx":53AC
      End
      Begin VB.TextBox Text12 
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
         Left            =   9240
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   1275
         Width           =   1575
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput43 
         Height          =   330
         Left            =   3720
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "1 Form-W&T"
         Top             =   3375
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":53C8
         Caption         =   "BookPOChild05.frx":53E8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":5454
         Keys            =   "BookPOChild05.frx":5472
         Spin            =   "BookPOChild05.frx":54BC
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput45 
         Height          =   330
         Left            =   2310
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2750
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":54E4
         Caption         =   "BookPOChild05.frx":5504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":5570
         Keys            =   "BookPOChild05.frx":558E
         Spin            =   "BookPOChild05.frx":55D8
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
         ForeColor       =   255
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel54 
         Height          =   330
         Left            =   7440
         TabIndex        =   134
         Top             =   3900
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
         Caption         =   " Reel Cut Off (mm)"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild05.frx":5600
         Picture         =   "BookPOChild05.frx":561C
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput52 
         Height          =   330
         Left            =   9240
         TabIndex        =   41
         Top             =   3900
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":5638
         Caption         =   "BookPOChild05.frx":5658
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":56C4
         Keys            =   "BookPOChild05.frx":56E2
         Spin            =   "BookPOChild05.frx":572C
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
         ForeColor       =   0
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   2
         Left            =   5700
         TabIndex        =   135
         Top             =   5280
         Width           =   5115
         _Version        =   65536
         _ExtentX        =   9022
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
         Caption         =   "Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D/F9->Delete  Ctrl+S/F2->Save"
         AutoSize        =   -1  'True
         FillColor       =   8421504
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":5754
         Multiline       =   -1  'True
         GlobalMem       =   -1  'True
         Picture         =   "BookPOChild05.frx":5770
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   10950
         Y1              =   7020
         Y2              =   7020
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   330
         Left            =   9240
         TabIndex        =   14
         Top             =   1910
         Width           =   1575
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2778;582"
         ListRows        =   3
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
         X2              =   10950
         Y1              =   8155
         Y2              =   8155
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   10950
         Y1              =   2325
         Y2              =   2325
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10950
         Y1              =   9005
         Y2              =   9005
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   10950
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10950
         Y1              =   5260
         Y2              =   5260
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10950
         Y1              =   3790
         Y2              =   3790
      End
   End
End
Attribute VB_Name = "FrmBookPOChild05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchCode As String, VchType As String, PartyCode As String, RoundOffQty As Boolean, rstBookPOChild05 As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset, rstGeneralList As New ADODB.Recordset, rstPlateMakerList As New ADODB.Recordset, rstFetchRate As New ADODB.Recordset, rstElementList As New ADODB.Recordset, WithEvents rstBookPOChild05c As ADODB.Recordset
Attribute rstBookPOChild05c.VB_VarHelpID = -1
Dim ItemCode As String, FinishSizeCode As String, SizeCode As String, TextSizeCode As String, PlateMakerCode As String, ElementCode As String, PaperCode As String, ColorCode As Variant, PlateCode As Variant, Plate As Integer, Color As Integer
Dim SPU As Long, Wt As Double, inLength As Double, inWidth As Double, GSM As Double, PaperForm As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    DisableCloseButton Me
    ItemCode = FrmBookPrintOrder.rstBookList.Fields("Code").Value
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text) 'Order No.
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text) 'Item Name
    Combo3.AddItem "Old", 0
    Combo3.AddItem "New", 1
    Combo3.AddItem "Revised", 2
    LoadMasterList
    ClearFields
    Set rstBookPOChild05c = New ADODB.Recordset
    cnDatabase.Execute "IF OBJECT_ID('tempdb.dbo.#T', 'U') IS NOT NULL DROP TABLE #T"
    cnDatabase.Execute "SELECT * INTO #T FROM (" & _
                                              "SELECT Element,E.Name As ElementName,ElementPrintName,FinishSize,FS.Name As FinishSizeName,[Size],PS.Name As PrintSizeName,P.DuplexPrinting,[Pages/PrintingForm],[Pages/Form],[Color],C.Name As ColorName,P.Pages,Forms,[Forms-¼],[Forms-½],[Forms-1-F&B],[Forms-1-W&T],PlateType,[Forms/Sheet1] As Ups,PaperConsumptionOther As PaperReqd FROM (((BookPOChild05 P INNER JOIN ElementMaster E ON P.[Element]=E.Code) INNER JOIN GeneralMaster FS ON P.FinishSize=FS.Code) INNER JOIN GeneralMaster PS ON P.[Size]=PS.Code) INNER JOIN GeneralMaster C ON P.Color=C.Code WHERE P.Code='" & VchCode & "' UNION " & _
                                              "SELECT Element,E.Name As ElementName,ElementPrintName,FinishSize,FS.Name As FinishSizeName,[Size],PS.Name As PrintSizeName,P.DuplexPrinting,[Pages/PrintingForm],[Pages/Form],[Color],C.Name As ColorName,P.Pages,Forms,[Forms-¼],[Forms-½],[Forms-1-F&B],[Forms-1-W&T],PlateType,Ups,0 As PaperReqd FROM (((BookChild05 P INNER JOIN ElementMaster E ON P.[Element]=E.Code) INNER JOIN GeneralMaster FS ON P.FinishSize=FS.Code) INNER JOIN GeneralMaster PS ON P.[Size]=PS.Code) INNER JOIN GeneralMaster C ON P.Color=C.Code WHERE P.Code='" & ItemCode & "' AND Element NOT IN (SELECT Element FROM BookPOChild05 WHERE Code='" & VchCode & "')" & _
                                              ") As Tbl ORDER BY ElementName,FinishSizeName,PrintSizeName"
                                              ' P.[Type]='" & VchType & "' AND
    rstBookPOChild05c.Open "SELECT * FROM #T", cnDatabase, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = rstBookPOChild05c
    rstBookPOChild05c.ActiveConnection = Nothing
    LockFields True
    SetButtons True
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Form_Activate()
    If Command1(0).Enabled Then If rstBookPOChild05c.RecordCount = 0 Then Command1_Click (0)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA And Command1(0).Enabled Then
        If Command1(0).Enabled Then Command1_Click (0)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Command1(1).Enabled Then
        If Command1(1).Enabled Then Command1_Click (1)
        KeyCode = 0
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF9)) And Command1(2).Enabled Then
        If Command1(2).Enabled Then Command1_Click (2)
        KeyCode = 0
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2)) And Command1(3).Enabled Then
        If Command1(3).Enabled Then Command1_Click (3)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        If Command1(4).Enabled Then
              If MsgBox("Are you sure to abandon the changes?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") = vbYes Then Command1_Click (4)
        Else
            cmdProceed_Click
        End If
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        If Not MhDateInput1.ReadOnly Then Sendkeys "{TAB}"
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstGeneralList)
    Call CloseRecordset(rstFetchRate)
    Call CloseRecordset(rstElementList)
    Call CloseRecordset(rstPlateMakerList)
    Call CloseRecordset(rstBookPOChild05c)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf CheckEmpty(VchCode, False) Then
        MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
    End If
End Sub
Private Sub MhDateInput3_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If Format(GetDate(MhDateInput3.Text), "yyyymmdd") <= Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then DisplayError ("Target Date cann't be prior to Order Date"): MhDateInput3.SetFocus: Cancel = True
End Sub
Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer) 'Element
    If MhDateInput1.ReadOnly Then Exit Sub
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text14.Text)
        If rstElementList.RecordCount = 0 Then DisplayError ("No Record in Element Master"): Text14.SetFocus: Exit Sub Else rstElementList.MoveFirst
        rstElementList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": ElementCode = ""
        Call LoadSelectionList(rstElementList, "List of Element(s)...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text14, ElementCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(ElementCode) <> "" Then Sendkeys "{TAB}" Else Text14.Text = ""
    End If
End Sub
Private Sub Text14_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If CheckEmpty(Text14.Text, False) Then
        Cancel = True
    ElseIf CheckDuplicateElement() Then
        Call DisplayError("Duplicate Element"): Cancel = True
    End If
End Sub
Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer) 'Finish Size
    If MhDateInput1.ReadOnly Then Exit Sub
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        With FrmFinishSizeMaster
            .SL = True
            .MasterCode = FinishSizeCode
            Load FrmFinishSizeMaster
            If Err.Number <> 364 Then .Show vbModal
            On Error GoTo 0
            FinishSizeCode = slCode: Text11.Text = slName
            If Not CheckEmpty(FinishSizeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
        End With
    End If
End Sub
Private Sub Text11_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If CheckEmpty(Text11.Text, False) Then
        Cancel = True
    Else
        With rstFetchRate
            If .State = adStateOpen Then .Close
            .Open "SELECT S.Name+'|'+'Pgs/Ptg Form: '+IIF([Ups/Form]<10,'0','')+LTRIM([Ups/Form]) As Col0,S.Name,S.Code FROM FinishSizeChild C INNER JOIN GeneralMaster S ON C.[TextSize]=S.Code WHERE C.Code='" & FinishSizeCode & "' ORDER BY S.Name,[Ups/Form]", cnDatabase, adOpenKeyset, adLockReadOnly
            SelectionType = "S": TextSizeCode = ""
            If Not CheckEmpty(Text4.Text, False) And .RecordCount > 0 Then
                .Find "[Name] = '" & RTrim(Text4.Text) & "'"
                If .EOF Then .MoveFirst Else Text12.Text = .Fields("Col0").Value 'Text4 is Printing Size & Text12 is Printing Size with some prefix (if any)
            End If
            Call LoadSelectionList(rstFetchRate, "List of Printing Sizes...", "Name", "")
            SearchOrder = 0
            Call DisplaySelectionList(Text12, TextSizeCode)
            Call CloseForm(FrmSelectionList)
            If Not CheckEmpty(Trim(TextSizeCode), False) Then
                .MoveFirst
                .Find "[Code] = '" & TextSizeCode & "'"
                Text4.Tag = .Fields("Name").Value & Right(.Fields("Col0").Value, 2)
            End If
        End With
    End If
End Sub
Private Sub Text4_GotFocus()
    If MhDateInput1.ReadOnly Then Exit Sub
    If Not CheckEmpty(Trim(TextSizeCode), False) Then
        If CheckEmpty(Text4.Text, False) Then Text4.Text = Mid(Text4.Tag, 1, Len(Text4.Tag) - 2): SizeCode = TextSizeCode
    End If
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer) 'Size
    If MhDateInput1.ReadOnly Then Exit Sub
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        With FrmGeneralMaster
            .SL = True
            .MasterType = "1"
            .MasterCode = SizeCode
            Load FrmGeneralMaster
            If Err.Number <> 364 Then .Show vbModal
            On Error GoTo 0
            SizeCode = slCode: Text4.Text = slName
            If Not CheckEmpty(SizeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
        End With
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If CheckEmpty(Text4.Text, False) Then
        Cancel = True
    Else
        If Not CheckEmpty(Trim(TextSizeCode), False) Then
            If Text4.Text <> Mid(Text4.Tag, 1, Len(Text4.Tag) - 2) Then If MsgBox("Printing Size [" & Trim(Text4.Text) & "] is different from that in Master [" & Trim(Mid(Text4.Tag, 1, Len(Text4.Tag) - 2)) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then Text4.Text = Mid(Text4.Tag, 1, Len(Text4.Tag) - 2): SizeCode = TextSizeCode
        End If
        If CheckEmpty(FinishSizeCode, False) Then Exit Sub
        Dim FL As Double, FR As Double, PL As Double, PR As Double, Ups01 As Integer, Ups02 As Integer
        PL = Val(Left(Text4.Text, InStr(1, Text4.Text, "X") - 1)) + 1: PR = Val(Mid(Text4.Text, InStr(1, Text4.Text, "X") + 1, 5)) + 1: FL = Val(Left(Text11.Text, InStr(1, Text11.Text, "X") - 1)): FR = Val(Mid(Text11.Text, InStr(1, Text11.Text, "X") + 1, 5))
        If Val(PL) * Val(PR) < Val(FL) * Val(FR) Then If MsgBox("Printing Size is smaller than Finish Size. Would you like to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Proceed !") = vbNo Then Cancel = True
    End If
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer) 'Plate Maker
    If MhDateInput1.ReadOnly Then Exit Sub
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        With FrmAccountMaster
            .SL = True
            .AccountType = "01": .AccountGroup = ""
            .MasterCode = PlateMakerCode
            Load FrmAccountMaster
            If Err.Number <> 364 Then .Show vbModal
            On Error GoTo 0
            PlateMakerCode = slCode: Text7.Text = slName
            If Not CheckEmpty(PlateMakerCode, False) Then LoadMasterList: Sendkeys "{TAB}"
        End With
    End If
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If CheckEmpty(Text7.Text, False) Then Cancel = True
End Sub
Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer) 'plate type
    If MhDateInput1.ReadOnly Then Exit Sub
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        With FrmGeneralMaster
            .SL = True
            .MasterType = "24"
            .MasterCode = PlateCode
            Load FrmGeneralMaster
            If Err.Number <> 364 Then .Show vbModal
        End With
        On Error GoTo 0
        PlateCode = slCode: Text16.Text = slName
        If Not CheckEmpty(PlateCode, False) Then
            LoadMasterList
            With rstGeneralList
                .MoveFirst
                .Find "[Code] = '" & PlateCode & "'"
                Plate = Val(.Fields("Value1").Value)
            End With
            Sendkeys "{TAB}"
        End If
    End If
End Sub
Private Sub Text16_Validate(Cancel As Boolean) 'plate type
    If MhDateInput1.ReadOnly Then Exit Sub
    If Left(VchType, 1) = "O" Then Exit Sub
    If Plate Then 'PS/CTP plate details
        On Error Resume Next
        With FrmPSPlateRegister
            .ItemCode = ItemCode
            .ItemName = Trim(Text2.Text)
            .ElementCode = ElementCode
            .ElementName = Trim(Text14.Text)
            .OrderCode = IIf(CheckEmpty(VchCode, False), "999999", VchCode)
            .OrderDate = GetDate(MhDateInput1.Text)
            .TblSuffix = "05"
            .OrderType = VchType
            Load FrmPSPlateRegister
            If Err.Number <> 364 Then .Show vbModal
        End With
        On Error GoTo 0
    End If
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean) 'Actual quantity
    If MhDateInput1.ReadOnly Then Exit Sub
    If MhRealInput1.Value = 0 Then Cancel = True: Exit Sub
    Call CalculateConsumption
End Sub
Private Sub MhRealInput2_GotFocus() 'Billing quantity
    If MhDateInput1.ReadOnly Then Exit Sub
    Dim q As Double
    q = MhRealInput1.Value \ 1000 'integer of quotient
    q = IIf(q = 0, 1000, q * 1000) + IIf(MhRealInput1.Value Mod 1000 <= IIf(MhRealInput1.Value <= 20000, 299, 599), 0, 1000)
    If MhRealInput2.Value = 0 Then MhRealInput2.Value = q
    MhRealInput2.Tag = q
'    If MhDateInput1.ReadOnly Then Exit Sub
'    CalculateBillingQty
End Sub
Private Sub CalculateBillingQty()
    Dim BillQty As Double
    If MhRealInput40.Value > 0 Then 'Pages/Form
        BillQty = MhRealInput1.Value  'BillQty
        If BillQty > 0 Then
            If RoundOffQty Then
                If BillQty < 1000 Then BillQty = 1000
                BillQty = IIf(Int(BillQty / 1000) = 0, 1000, Int(BillQty / 1000) * 1000) + IIf(BillQty Mod 1000 <= IIf(BillQty <= 20000, 299, 599), 0, 1000)
            End If
            If MhRealInput2.Value = 0 Then
                MhRealInput2.Value = BillQty
            ElseIf MhRealInput2.Value <> BillQty Then
                If MsgBox("Variation in Billing Qty. [" & Trim(BillQty) & "] and Existing [" & Trim(MhRealInput2.Value) & "] Billing Qty. ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput2.Value = BillQty
            End If
        End If
    End If
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
'    If MhRealInput2.Value = 0 Or MhRealInput2.Value Mod 1000 <> 0 Then Cancel = True: Exit Sub
    If MhRealInput2.Value = 0 Then Cancel = True: Exit Sub
    If Val(MhRealInput2.Tag) <> MhRealInput2.Value And Val(MhRealInput2.Tag) <> 0 Then
        If MsgBox("Variation in Calculated [" & Trim(MhRealInput2.Tag) & "] and Existing [" & Trim(MhRealInput2.Value) & "] Billing Quantity ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput2.Value = Val(MhRealInput2.Tag)
    End If
    CalculateTotalForms
End Sub
Private Sub Combo3_Validate(Cancel As Boolean) 'Plate
    If Combo3.ListIndex = -1 Then Cancel = True
    If Combo3.ListIndex = 0 Then If Plate Then MhRealInput4.Value = 0
End Sub
Private Sub Text13_KeyDown(KeyCode As Integer, Shift As Integer) 'Color
    If MhDateInput1.ReadOnly Then Exit Sub
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        With FrmGeneralMaster
            .SL = True
            .MasterType = "23"
            .MasterCode = ColorCode
            Load FrmGeneralMaster
            If Err.Number <> 364 Then .Show vbModal
            On Error GoTo 0
            ColorCode = slCode: Text13.Text = slName
            Color = slValue1
            If Not CheckEmpty(ColorCode, False) Then
                LoadMasterList
                With rstGeneralList
                    .MoveFirst
                    .Find "[Code] = '" & ColorCode & "'"
                End With
                Sendkeys "{TAB}"
            End If
        End With
    End If
End Sub
Private Sub Text13_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If CheckEmpty(Text13.Text, False) Then Cancel = True Else Call CalculateTotalPlates
End Sub
Private Sub MhRealInput40_GotFocus() 'Pages/Printing Form
    If MhDateInput1.ReadOnly Then Exit Sub
    Dim Ups As Integer
    If Not CheckEmpty(Trim(TextSizeCode), False) Then Ups = Val(Right(Text4.Tag, 2))
    If Ups = 0 Then Ups = MaxUps("T")
    If MhRealInput40.Value = 0 Then MhRealInput40.Value = Ups
    MhRealInput40.Tag = Ups 'Calculated value
End Sub
Private Sub MhRealInput40_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If MhRealInput40.Value = 0 Then
        Cancel = True
    ElseIf Val(MhRealInput40.Tag) <> MhRealInput40.Value And Val(MhRealInput40.Tag) <> 0 Then
        If MsgBox("Variation in Calculated [" & Trim(MhRealInput40.Tag) & "] and Existing [" & Trim(MhRealInput40.Value) & "] Pages/Printing ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput40.Value = Val(MhRealInput40.Tag)
    End If
End Sub
Private Sub MhRealInput42_GotFocus() 'Pages/Form
    If MhDateInput1.ReadOnly Then Exit Sub
    If MhRealInput42.Value = 0 Then MhRealInput42.Value = MhRealInput40.Value
End Sub
Private Sub MhRealInput42_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If MhRealInput42.Value = 0 Then
        Cancel = True
    ElseIf MhRealInput40.Value <> MhRealInput42.Value And MhRealInput40.Value <> 0 Then
           If MsgBox("Variation in Pages/Printing Form [" & Trim(MhRealInput40.Value) & "] and Pages/Form [" & Trim(MhRealInput42.Value) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput42.Value = MhRealInput40.Value
   End If
End Sub
Private Sub MhRealInput15_Validate(Cancel As Boolean) 'Pages
    If MhDateInput1.ReadOnly Then Exit Sub
    If MhRealInput15.Value = 0 Then
        Cancel = True
    Else
        If MhRealInput42.Value > 0 Then MhRealInput45.Value = MhRealInput15.Value / MhRealInput42.Value
    End If
End Sub
Private Sub MhRealInput17_GotFocus() '¼ Forms
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateForms "Q"
End Sub
Private Sub MhRealInput17_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    Call CalculateTotalForms: CalculateTotalPlates
End Sub
Private Sub MhRealInput20_GotFocus() '½ Forms
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateForms "H"
End Sub
Private Sub MhRealInput20_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    Call CalculateTotalForms: CalculateTotalPlates
End Sub
Private Sub MhRealInput21_GotFocus() '1 Forms-F&B
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateForms "F"
End Sub
Private Sub MhRealInput21_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    Call CalculateTotalForms: CalculateTotalPlates
End Sub
Private Sub MhRealInput41_GotFocus() '1 Forms-W&T
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateForms "W"
End Sub
Private Sub MhRealInput41_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    Call CalculateTotalForms: CalculateTotalPlates
End Sub
Private Sub MhRealInput22_GotFocus() 'Forms/Sheet
    If MhDateInput1.ReadOnly Then Exit Sub
    Dim FS As Double
    If MhRealInput40.Value > 0 Then FS = MhRealInput42.Value / MhRealInput40.Value
    If MhRealInput22.Value = 0 Then MhRealInput22.Value = FS
    MhRealInput22.Tag = FS
End Sub
Private Sub MhRealInput22_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If MhRealInput22.Value = 0 Then
        Cancel = True
    ElseIf Val(MhRealInput22.Tag) <> MhRealInput22.Value And Val(MhRealInput22.Tag) <> 0 Then
        If MsgBox("Variation in Calculated [" & Trim(MhRealInput22.Tag) & "] and Existing [" & Trim(MhRealInput22.Value) & "] Forms/Sheet ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput22.Value = Val(MhRealInput22.Tag)
    End If
End Sub
Private Sub Option1_Validate(Cancel As Boolean) 'Duplex Printing-Yes
    Call CalculateTotalForms: CalculateTotalPlates
End Sub
Private Sub Option2_Validate(Cancel As Boolean) 'Duplex Printing-No
    Call CalculateTotalForms: CalculateTotalPlates
End Sub
Private Sub MhRealInput59_Validate(Cancel As Boolean) 'Revised Plates
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateAmount
End Sub
Private Sub MhRealInput4_GotFocus() 'Plate Rate
    If MhDateInput1.ReadOnly Then Exit Sub
    Call GetPartyRates("L")
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateAmount
End Sub
Private Sub MhRealInput5_GotFocus() 'Print Rate
    If MhDateInput1.ReadOnly Then Exit Sub
    Call GetPartyRates("P")
End Sub
Private Sub MhRealInput5_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateAmount
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer) 'Paper
    If MhDateInput1.ReadOnly Then Exit Sub
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
    End If
End Sub
Private Sub Text1_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If CheckEmpty(Text1.Text, False) Then
        Cancel = True
    Else
        With rstPaperList
            .MoveFirst
            .Find "[Code]='" & PaperCode & "'"
            Text1.Text = .Fields("Col0").Value: SPU = Val(.Fields("SPU").Value): Wt = Val(.Fields("Wt").Value): inWidth = Val(.Fields("inWidth").Value): inLength = Val(.Fields("inLength").Value): GSM = Val(.Fields("GSM").Value): PaperForm = .Fields("Form").Value
            If PaperForm = "S" Then MhRealInput52.Value = 0: CalculateConsumption: CheckPaperSize
        End With
    End If
End Sub
Private Sub MhRealInput52_GotFocus() 'Reel cut off
    If MhDateInput1.ReadOnly Then Exit Sub
    If PaperForm = "R" Then If MhRealInput52.Value = 0 Then MhRealInput52.Value = inLength
End Sub
Private Sub MhRealInput52_Validate(Cancel As Boolean) 'Reel cut off
    If MhDateInput1.ReadOnly Then Exit Sub
    If PaperForm = "R" Then
        If MhRealInput52.Value = 0 Then
            DisplayError ("Reel cut off Size (mm) can't be zero as you are using paper reel"): Cancel = True
        Else
            If Val(inLength) <> MhRealInput52.Value And inLength <> 0 Then
                If MsgBox("Reel cut off [" & Trim(MhRealInput52.Value) & "] is different from that in Master [" & Trim(Format(inLength, "#0")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput52.Value = inLength
            End If
            CalculateConsumption
            CheckPaperSize
        End If
    End If
End Sub
Private Sub MhRealInput27_GotFocus() 'Ups/Sheet
    If MhDateInput1.ReadOnly Then Exit Sub
    Dim Ups As Integer
    Ups = MaxUps("B")
    If MhRealInput27.Value = 0 Then MhRealInput27.Value = Ups
    MhRealInput27.Tag = Ups 'Calculated value
End Sub
Private Sub MhRealInput27_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    If MhRealInput27.Value = 0 Then
        Cancel = True: Exit Sub
    ElseIf Val(MhRealInput27.Tag) <> MhRealInput27.Value And Val(MhRealInput27.Tag) <> 0 Then
        If MsgBox("Variation in Calculated [" & Trim(MhRealInput27.Tag) & "] and Existing [" & Trim(MhRealInput27.Value) & "] Ups/Sheet ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput27.Value = Val(MhRealInput27.Tag)
    End If
    CalculateConsumption
End Sub
Private Sub MhRealInput11_GotFocus() 'Wastage %
    If MhDateInput1.ReadOnly Then Exit Sub
    Call GetPartyRates("W")
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateConsumption
End Sub
Private Sub MhRealInput47_GotFocus() 'Wastage Min
    If MhDateInput1.ReadOnly Then Exit Sub
    Call GetPartyRates("M")
End Sub
Private Sub MhRealInput47_Validate(Cancel As Boolean)
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateConsumption
End Sub
Private Sub MhRealInput32_Validate(Cancel As Boolean) 'Paper Rate
    If MhDateInput1.ReadOnly Then Exit Sub
    MhRealInput33.Value = MhRealInput32.Value * MhRealInput49.Value
    CalculateTotalAmount
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean) 'Adjustment
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateTotalAmount
End Sub
Private Sub MhRealInput14_Validate(Cancel As Boolean) 'GST %
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateTotalAmount
End Sub
Private Sub MhRealInput34_Validate(Cancel As Boolean) 'Plate Adjustment
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateTotalAmount
End Sub
Private Sub MhRealInput28_Validate(Cancel As Boolean) 'Plate GST%
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateTotalAmount
End Sub
Private Sub MhRealInput36_Validate(Cancel As Boolean) 'Paper Adjustment
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateTotalAmount
End Sub
Private Sub MhRealInput37_Validate(Cancel As Boolean) 'Paper GST%
    If MhDateInput1.ReadOnly Then Exit Sub
    CalculateTotalAmount
End Sub
Private Sub cmdProceed_Click()
    If Not Command1(4).Enabled Then 'Cancel button disabled
        With rstBookPOChild05c
            If .RecordCount > 0 Then
                .MoveFirst
                Do Until .EOF
                    If rstBookPOChild05c.Fields("PaperReqd").Value = 0 Then If MsgBox("[" & .Fields("ElementName").Value & "] Element has not been processed ! Process?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Process !") = vbYes Then Command1_Click (1): Exit Sub
                    .MoveNext
                Loop
            End If
        End With
        If Not CheckEmpty(Text8.Text, False) Or Not CheckEmpty(Text10.Text, False) Then
            With rstBookPOChild05
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do Until .EOF()
                        .Fields("BillNo").Value = Text8.Text
                        If MhDateInput2.ValueIsNull Then .Fields("BillDate").Value = Null Else .Fields("BillDate").Value = GetDate(MhDateInput2.Text)
                        .Fields("PBillNo").Value = Text10.Text
                        If MhDateInput4.ValueIsNull Then .Fields("PBillDate").Value = Null Else .Fields("PBillDate").Value = GetDate(MhDateInput4.Text)
                        .Fields("PaidAmount").Value = MhRealInput16.Value
                        .Fields("PPaidAmount").Value = MhRealInput30.Value
                        .Update
                        .MoveNext
                    Loop
                End If
            End With
        End If
        FrmBookPrintOrder.Command1.Enabled = False: Call CloseForm(Me)
    Else
        Command1_Click (3)
    End If
End Sub
Private Sub Command1_Click(Index As Integer)
    With rstBookPOChild05c
        Select Case Index
            Case 0
                ClearFields
                Me.Tag = "A"
                SetButtons False
                LockFields False
                MhDateInput1.Text = Format(Date, "dd-MM-yyyy"): MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
                FinishSizeCode = FrmBookPrintOrder.rstBookList.Fields("FinishSize").Value
                If rstGeneralList.RecordCount > 0 Then rstGeneralList.MoveFirst
                rstGeneralList.Find "[Code] = '" & FinishSizeCode & "'"
                If Not rstGeneralList.EOF Then Text11.Text = rstGeneralList.Fields("Col0").Value
                PlateMakerCode = PartyCode
                If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
                rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
                If Not rstPlateMakerList.EOF Then Text7.Text = rstPlateMakerList.Fields("Col0").Value
                MhRealInput1.Value = FrmBookPrintOrder.MhRealInput3.Value 'Final Quantity
                MhDateInput1.SetFocus
            Case 1
                If .RecordCount > 0 Then
                    ClearFields
                    Me.Tag = "E"
                    MhDateInput1.Text = Format(Date, "dd-MM-yyyy"): MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
                    ElementCode = .Fields("Element").Value
                    If rstElementList.RecordCount > 0 Then rstElementList.MoveFirst
                    rstElementList.Find "[Code] = '" & ElementCode & "'"
                    If Not rstElementList.EOF Then Text14.Text = rstElementList.Fields("Col0").Value
                    If Not IsNull(.Fields("ElementPrintName").Value) Then Text9.Text = .Fields("ElementPrintName").Value
                    FinishSizeCode = .Fields("FinishSize").Value
                    If rstGeneralList.RecordCount > 0 Then rstGeneralList.MoveFirst
                    rstGeneralList.Find "[Code] = '" & FinishSizeCode & "'"
                    If Not rstGeneralList.EOF Then Text11.Text = rstGeneralList.Fields("Col0").Value
                    SizeCode = .Fields("Size").Value
                    If rstGeneralList.RecordCount > 0 Then rstGeneralList.MoveFirst
                    rstGeneralList.Find "[Code] = '" & SizeCode & "'"
                    If Not rstGeneralList.EOF Then Text4.Text = rstGeneralList.Fields("Col0").Value
                    Option1.Value = IIf(.Fields("DuplexPrinting").Value, True, False): Option2.Value = IIf(.Fields("DuplexPrinting").Value, False, True)
                    PlateMakerCode = PartyCode
                    If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
                    rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
                    If Not rstPlateMakerList.EOF Then Text7.Text = rstPlateMakerList.Fields("Col0").Value
                    MhRealInput40.Value = Val(.Fields("Pages/PrintingForm").Value)
                    MhRealInput42.Value = Val(.Fields("Pages/Form").Value)
                    ColorCode = .Fields("Color").Value
                    If rstGeneralList.RecordCount > 0 Then rstGeneralList.MoveFirst
                    rstGeneralList.Find "[Code] = '" & ColorCode & "'"
                    If Not rstGeneralList.EOF Then Color = rstGeneralList.Fields("Value1").Value
                    If Not rstGeneralList.EOF Then Text13.Text = Trim(rstGeneralList.Fields("Col0").Value)
                    MhRealInput15.Value = Val(.Fields("Pages").Value)
                    MhRealInput45.Value = Val(.Fields("Forms").Value)
                    MhRealInput17.Value = Val(.Fields("Forms-¼").Value)
                    MhRealInput20.Value = Val(.Fields("Forms-½").Value)
                    MhRealInput21.Value = Val(.Fields("Forms-1-F&B").Value)
                    MhRealInput41.Value = Val(.Fields("Forms-1-W&T").Value)
                    PlateCode = .Fields("PlateType").Value
                    If rstGeneralList.RecordCount > 0 Then rstGeneralList.MoveFirst
                    rstGeneralList.Find "[Code] = '" & PlateCode & "'"
                    If Not rstGeneralList.EOF Then Text16.Text = rstGeneralList.Fields("Col0").Value: Plate = rstGeneralList.Fields("Value1").Value
                    MhRealInput22.Value = Val(.Fields("Ups").Value)
                    LoadFields
                    If MhRealInput1.Value = 0 Then MhRealInput1.Value = FrmBookPrintOrder.MhRealInput3.Value 'Final Quantity
                    SetButtons False
                    LockFields False
                    DataGrid1.Enabled = False
                    MhDateInput1.SetFocus
                End If
            Case 2
                If .RecordCount > 0 Then
                    If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
                        If rstBookPOChild05.RecordCount > 0 Then
                            rstBookPOChild05.MoveFirst
                            rstBookPOChild05.Find "[Element]='" & .Fields("Element").Value & "'"
                            If Not rstBookPOChild05.EOF Then rstBookPOChild05.Delete: rstBookPOChild05.MoveNext
                        End If
                        Me.Tag = "D"
                        .Delete: .MoveNext
                        Me.Tag = ""
                    End If
                End If
            Case 3
                If CheckMandatoryFields Then Exit Sub
                If Me.Tag = "A" Then Call AddRecord(rstBookPOChild05c)
                .Fields("Element").Value = ElementCode
                .Fields("ElementName").Value = Text14.Text
                .Fields("FinishSize").Value = FinishSizeCode
                .Fields("FinishSizeName").Value = Text11.Text
                .Fields("Size").Value = SizeCode
                .Fields("PrintSizeName").Value = Text4.Text
                .Fields("DuplexPrinting").Value = IIf(Option1.Value, 1, 0)
                .Fields("Pages/PrintingForm").Value = MhRealInput40.Value
                .Fields("Pages/Form").Value = MhRealInput42.Value
                .Fields("Color").Value = ColorCode
                .Fields("Pages").Value = MhRealInput15.Value
                .Fields("Forms").Value = MhRealInput45.Value
                .Fields("Forms-¼").Value = MhRealInput17.Value
                .Fields("Forms-½").Value = MhRealInput20.Value
                .Fields("Forms-1-F&B").Value = MhRealInput21.Value
                .Fields("Forms-1-W&T").Value = MhRealInput41.Value
                .Fields("PlateType").Value = PlateCode
                .Fields("Ups").Value = MhRealInput22.Value
                .Fields("PaperReqd").Value = MhRealInput49.Value
                .Update
                If InStr(1, "A_E1", Me.Tag) > 0 Then Call AddRecord(rstBookPOChild05)
                SaveFields
                rstBookPOChild05.Update
                SetButtons True
                LockFields True
                DataGrid1.Enabled = True
                DataGrid1.SetFocus
                If Left(Me.Tag, 1) = "E" Then
                    Me.Tag = ""
                    rstBookPOChild05c.MoveNext
                    If rstBookPOChild05c.EOF Then
                        rstBookPOChild05c.MoveLast
                            If MsgBox("All Element has been processed..!!! Do you want to Exit the Process?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Quit !") = vbYes Then cmdProceed_Click
                    Else
                        Command1_Click (1)
                    End If
                Else
                    Me.Tag = ""
                End If
'                Me.Tag = ""
            Case 4  'Cancel
                ClearFields
                SetButtons True
                LockFields True
                If .RecordCount > 0 Then LoadFields
                Me.Tag = ""
                DataGrid1.Enabled = True
                DataGrid1.SetFocus
        End Select
    End With
End Sub
Private Sub rstBookPOChild05c_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If CheckEmpty(Me.Tag, False) Then
        On Error Resume Next
        ClearFields
        If Not (rstBookPOChild05c.EOF Or rstBookPOChild05c.BOF) Then LoadFields
    End If
End Sub
Private Sub ClearFields()
    MhDateInput1.Value = Null
    MhDateInput3.Value = Null
    Text14.Text = "": ElementCode = ""
    Text11.Text = "": FinishSizeCode = "": TextSizeCode = ""
    Text4.Text = "": SizeCode = "'"
    Option1.Value = True: Option2.Value = False
    Combo3.ListIndex = 0
    Text3.Text = ""
    Text7.Text = "": PlateMakerCode = ""
    MhRealInput1.Value = 0
    MhRealInput2.Value = 0
    MhRealInput40.Value = 0
    MhRealInput42.Value = 0
    Text13.Text = "": ColorCode = "": Color = 0
    MhRealInput15.Value = 0
    MhRealInput45.Value = 0
    MhRealInput17.Value = 0
    MhRealInput20.Value = 0
    MhRealInput21.Value = 0
    MhRealInput41.Value = 0
    Text16.Text = "": PlateCode = "*24004"
    With rstGeneralList
        .MoveFirst
        .Find "[Code] = '" & PlateCode & "'"
        Text16.Text = .Fields("Col0").Value
        Plate = Val(.Fields("Value1").Value)
    End With
    MhRealInput6.Value = 0
    MhRealInput25.Value = 0
    MhRealInput26.Value = 0
    MhRealInput43.Value = 0
    MhRealInput3.Value = 0
    MhRealInput23.Value = 0
    MhRealInput24.Value = 0
    MhRealInput46.Value = 0
    MhRealInput59.Value = 0
    MhRealInput5.Value = 0
    MhRealInput8.Value = 0
    MhRealInput4.Value = 0
    MhRealInput7.Value = 0
    chkPaper.Value = 1
    Text1.Text = "": PaperCode = "": SPU = 0: Wt = 0: inLength = 0: inWidth = 0: GSM = 0: PaperForm = ""
    MhRealInput52.Value = 0
    MhRealInput11.Value = 0
    MhRealInput47.Value = 0
    MhRealInput48.Value = 0
    MhRealInput13.Value = 0
    MhRealInput49.Value = 0
    MhRealInput32.Value = 0
    MhRealInput33.Value = 0
    MhRealInput22.Value = 1
    MhRealInput27.Value = 0
    Text6.Text = ""
    Text9.Text = ""
    Text8.Text = ""
    MhDateInput2.Value = Null
    Text10.Text = ""
    MhDateInput4.Value = Null
    MhRealInput9.Value = 0
    MhRealInput34.Value = 0
    MhRealInput36.Value = 0
    MhRealInput14.Value = 0
    MhRealInput18.Value = 0
    MhRealInput28.Value = 0
    MhRealInput29.Value = 0
    MhRealInput37.Value = 0
    MhRealInput38.Value = 0
    MhRealInput10.Value = 0
    MhRealInput35.Value = 0
    MhRealInput39.Value = 0
    MhRealInput16.Value = 0
    MhRealInput30.Value = 0
    TxtAdNar.Text = ""
    Text12.Text = ""
End Sub
Private Sub LoadFields()
    With rstBookPOChild05
        If Me.Tag = "E" Then Me.Tag = "E1"
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        .Find "[Element]='" & rstBookPOChild05c.Fields("Element").Value & "'"
        If Not .EOF Then
            If Me.Tag = "E1" Then Me.Tag = "E2"
            MhDateInput1.Text = Format(.Fields("OrderDate").Value, "dd-MM-yyyy")
            MhDateInput3.Text = Format(.Fields("TargetDate").Value, "dd-MM-yyyy")
            ElementCode = .Fields("Element").Value
            If rstElementList.RecordCount > 0 Then rstElementList.MoveFirst
            rstElementList.Find "[Code] = '" & ElementCode & "'"
            If Not rstElementList.EOF Then Text14.Text = Trim(rstElementList.Fields("Col0").Value)
            FinishSizeCode = .Fields("FinishSize").Value
            If rstGeneralList.RecordCount > 0 Then rstGeneralList.MoveFirst
            rstGeneralList.Find "[Code] = '" & FinishSizeCode & "'"
            If Not rstGeneralList.EOF Then Text11.Text = rstGeneralList.Fields("Col0").Value
            SizeCode = .Fields("Size").Value
            If rstGeneralList.RecordCount > 0 Then rstGeneralList.MoveFirst
            rstGeneralList.Find "[Code] = '" & SizeCode & "'"
            If Not rstGeneralList.EOF Then Text4.Text = rstGeneralList.Fields("Col0").Value
            Option1.Value = IIf(.Fields("DuplexPrinting").Value, True, False): Option2.Value = IIf(.Fields("DuplexPrinting").Value, False, True)
            Combo3.ListIndex = IIf(.Fields("Processing").Value = "O", 0, IIf(.Fields("Processing").Value = "N", 1, 2))  'O:Old N:New R:Revised
            Text3.Text = .Fields("Ref").Value
            PlateMakerCode = .Fields("PlateMaker").Value
            If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
            rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
            If Not rstPlateMakerList.EOF Then Text7.Text = Trim(rstPlateMakerList.Fields("Col0").Value)
            MhRealInput1.Value = Val(.Fields("ActualQuantity").Value)
            MhRealInput2.Value = Val(.Fields("BillingQuantity").Value)
            MhRealInput40.Value = Val(.Fields("Pages/PrintingForm").Value)
            MhRealInput42.Value = Val(.Fields("Pages/Form").Value)
            ColorCode = .Fields("Color").Value
            If rstGeneralList.RecordCount > 0 Then rstGeneralList.MoveFirst
            rstGeneralList.Find "[Code] = '" & ColorCode & "'"
            Color = rstGeneralList.Fields("Value1").Value
            If Not rstGeneralList.EOF Then Text13.Text = Trim(rstGeneralList.Fields("Col0").Value)
            MhRealInput15.Value = Val(.Fields("Pages").Value)
            MhRealInput45.Value = Val(.Fields("Forms").Value)
            MhRealInput17.Value = Val(.Fields("Forms-¼").Value)
            MhRealInput20.Value = Val(.Fields("Forms-½").Value)
            MhRealInput21.Value = Val(.Fields("Forms-1-F&B").Value)
            MhRealInput41.Value = Val(.Fields("Forms-1-W&T").Value)
            PlateCode = .Fields("PlateType").Value
            If rstGeneralList.RecordCount > 0 Then rstGeneralList.MoveFirst
            rstGeneralList.Find "[Code] = '" & PlateCode & "'"
            If Not rstGeneralList.EOF Then Text16.Text = rstGeneralList.Fields("Col0").Value: Plate = rstGeneralList.Fields("Value1").Value
            MhRealInput6.Value = Val(.Fields("TotalForms-¼").Value)
            MhRealInput25.Value = Val(.Fields("TotalForms-½").Value)
            MhRealInput26.Value = Val(.Fields("TotalForms-1-F&B").Value)
            MhRealInput43.Value = Val(.Fields("TotalForms-1-W&T").Value)
            MhRealInput3.Value = Val(.Fields("TotalPlates-¼").Value)
            MhRealInput23.Value = Val(.Fields("TotalPlates-½").Value)
            MhRealInput24.Value = Val(.Fields("TotalPlates-1-F&B").Value)
            MhRealInput46.Value = Val(.Fields("TotalPlates-1-W&T").Value)
            MhRealInput59.Value = Val(.Fields("RevisedPlates").Value)
            MhRealInput5.Value = Val(.Fields("PrintRate").Value)
            MhRealInput8.Value = Val(.Fields("PrintAmount").Value)
            MhRealInput4.Value = Val(.Fields("PlateRate").Value)
            MhRealInput7.Value = Val(.Fields("PlateAmount").Value)
            chkPaper.Value = IIf(.Fields("PaperByParty").Value, 1, 0)
            PaperCode = .Fields("Paper").Value
            If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
            rstPaperList.Find "[Code] = '" & PaperCode & "'"
            If Not rstPaperList.EOF Then Text1.Text = rstPaperList.Fields("Col0").Value: SPU = Val(rstPaperList.Fields("SPU").Value): Wt = Val(rstPaperList.Fields("Wt").Value): inLength = Val(rstPaperList.Fields("inLength").Value): inWidth = Val(rstPaperList.Fields("inWidth").Value): GSM = Val(rstPaperList.Fields("GSM").Value): PaperForm = rstPaperList.Fields("Form").Value
            MhRealInput52.Value = Val(.Fields("CutOffSize").Value)
            MhRealInput11.Value = Val(.Fields("PaperWastage%").Value)
            MhRealInput47.Value = Val(.Fields("PaperWastageMin").Value)
            MhRealInput48.Value = Val(.Fields("PaperWastageFinal").Value)
            MhRealInput13.Value = Val(.Fields("PaperConsumptionOther").Value)
            MhRealInput49.Value = Val(.Fields("PaperConsumptionKg").Value)
            MhRealInput32.Value = Val(.Fields("PaperRate").Value)
            MhRealInput33.Value = Val(.Fields("PaperAmount").Value)
            MhRealInput22.Value = Val(.Fields("Forms/Sheet1").Value)
            MhRealInput27.Value = Val(.Fields("Forms/Sheet2").Value)
            Text6.Text = .Fields("Remarks").Value
            Text9.Text = .Fields("ElementPrintName").Value
            Text8.Text = .Fields("BillNo").Value
            If Not IsNull(.Fields("BillDate").Value) Then MhDateInput2.Text = Format(.Fields("BillDate").Value, "dd-MM-yyyy")
            Text10.Text = .Fields("PBillNo").Value
            If Not IsNull(.Fields("PBillDate").Value) Then MhDateInput4.Text = Format(.Fields("PBillDate").Value, "dd-MM-yyyy")
            MhRealInput9.Value = .Fields("Adjustment").Value
            MhRealInput34.Value = .Fields("PAdjustment").Value
            MhRealInput36.Value = .Fields("RAdjustment").Value
            MhRealInput14.Value = .Fields("VAT%").Value
            MhRealInput18.Value = .Fields("VAT").Value
            MhRealInput28.Value = .Fields("PVAT%").Value
            MhRealInput29.Value = .Fields("PVAT").Value
            MhRealInput37.Value = .Fields("RVAT%").Value
            MhRealInput38.Value = .Fields("RVAT").Value
            MhRealInput10.Value = .Fields("BillAmount").Value
            MhRealInput35.Value = .Fields("PBillAmount").Value
            MhRealInput39.Value = .Fields("RBillAmount").Value
            MhRealInput16.Value = .Fields("PaidAmount").Value
            MhRealInput30.Value = .Fields("PPaidAmount").Value
            TxtAdNar.Text = .Fields("AdjustmentRemarks").Value
        End If
    End With
End Sub
Private Sub SaveFields()
    With rstBookPOChild05
        .Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
        .Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
        .Fields("SubItem").Value = ItemCode
        .Fields("Element").Value = ElementCode
        .Fields("ElementPrintName").Value = Text9.Text
        .Fields("FinishSize").Value = FinishSizeCode
        .Fields("Size").Value = SizeCode
        .Fields("DuplexPrinting").Value = IIf(Option1.Value, 1, 0)
        .Fields("Processing").Value = IIf(Combo3.ListIndex = 0, "O", IIf(Combo3.ListIndex = 1, "N", "R"))
        .Fields("Ref").Value = Text3.Text
        .Fields("PlateMaker").Value = PlateMakerCode
        .Fields("ActualQuantity").Value = MhRealInput1.Value
        .Fields("BillingQuantity").Value = MhRealInput2.Value
        .Fields("Pages/PrintingForm").Value = MhRealInput40.Value
        .Fields("Pages/Form").Value = MhRealInput42.Value
        .Fields("Color").Value = ColorCode
        .Fields("Pages").Value = Val(MhRealInput15.Value)
        .Fields("Forms").Value = Val(MhRealInput45.Value)
        .Fields("Forms-¼").Value = MhRealInput17.Value
        .Fields("Forms-½").Value = MhRealInput20.Value
        .Fields("Forms-1-F&B").Value = MhRealInput21.Value
        .Fields("Forms-1-W&T").Value = MhRealInput41.Value
        .Fields("PlateType").Value = PlateCode
        .Fields("TotalForms-¼").Value = MhRealInput6.Value
        .Fields("TotalForms-½").Value = MhRealInput25.Value
        .Fields("TotalForms-1-F&B").Value = MhRealInput26.Value
        .Fields("TotalForms-1-W&T").Value = MhRealInput43.Value
        .Fields("TotalPlates-¼").Value = MhRealInput3.Value
        .Fields("TotalPlates-½").Value = MhRealInput23.Value
        .Fields("TotalPlates-1-F&B").Value = MhRealInput24.Value
        .Fields("TotalPlates-1-W&T").Value = MhRealInput46.Value
        .Fields("RevisedPlates").Value = MhRealInput59.Value
        .Fields("aTotalPlates-¼").Value = MhRealInput3.Value
        .Fields("aTotalPlates-½").Value = MhRealInput23.Value
        .Fields("aTotalPlates-1-F&B").Value = MhRealInput24.Value
        .Fields("aTotalPlates-1-W&T").Value = MhRealInput46.Value
        .Fields("aRevisedPlates").Value = MhRealInput59.Value
        .Fields("PrintRate").Value = MhRealInput5.Value
        .Fields("PrintAmount").Value = MhRealInput8.Value
        .Fields("PlateRate").Value = MhRealInput4.Value
        .Fields("PlateAmount").Value = MhRealInput7.Value
        .Fields("PaperByParty").Value = chkPaper.Value
        .Fields("Paper").Value = PaperCode
        .Fields("CutOffSize").Value = MhRealInput52.Value
        .Fields("PaperWastage%").Value = MhRealInput11.Value
        .Fields("PaperWastageMin").Value = MhRealInput47.Value
        .Fields("Wastage/Set").Value = MhRealInput47.Value
        .Fields("PaperWastageFinal").Value = MhRealInput48.Value
        .Fields("PaperConsumptionOther").Value = MhRealInput13.Value
        .Fields("PaperConsumptionSheets").Value = CLng(Int(MhRealInput13.Value) * SPU) + ((MhRealInput13.Value - Int(MhRealInput13.Value)) * 1000)
        .Fields("PaperConsumptionKg").Value = MhRealInput49.Value
        .Fields("aPaperWastage%").Value = MhRealInput11.Value
        .Fields("aPaperWastageMin").Value = MhRealInput47.Value
        .Fields("aWastage/Set").Value = MhRealInput47.Value
        .Fields("aPaperWastageFinal").Value = MhRealInput48.Value
        .Fields("aPaperConsumptionOther").Value = MhRealInput13.Value
        .Fields("aPaperConsumptionSheets").Value = CLng(Int(MhRealInput13.Value) * SPU) + ((MhRealInput13.Value - Int(MhRealInput13.Value)) * 1000)
        .Fields("aPaperConsumptionKg").Value = MhRealInput49.Value
        .Fields("PaperRate").Value = MhRealInput32.Value
        .Fields("PaperAmount").Value = MhRealInput33.Value
        .Fields("Forms/Sheet1").Value = MhRealInput22.Value
        .Fields("Forms/Sheet2").Value = MhRealInput27.Value
        .Fields("Remarks").Value = Text6.Text
        .Fields("BillNo").Value = Text8.Text
        If MhDateInput2.ValueIsNull Then .Fields("BillDate").Value = Null Else .Fields("BillDate").Value = GetDate(MhDateInput2.Text)
        .Fields("PBillNo").Value = Text10.Text
        If MhDateInput4.ValueIsNull Then .Fields("PBillDate").Value = Null Else .Fields("PBillDate").Value = GetDate(MhDateInput4.Text)
        .Fields("Adjustment").Value = MhRealInput9.Value
        .Fields("PAdjustment").Value = MhRealInput34.Value
        .Fields("RAdjustment").Value = MhRealInput36.Value
        .Fields("VAT%").Value = MhRealInput14.Value
        .Fields("VAT").Value = MhRealInput18.Value
         .Fields("PVAT%").Value = MhRealInput28.Value
        .Fields("PVAT").Value = MhRealInput29.Value
        .Fields("RVAT%").Value = MhRealInput37.Value
        .Fields("RVAT").Value = MhRealInput38.Value
        .Fields("BillAmount").Value = MhRealInput10.Value
        .Fields("PBillAmount").Value = MhRealInput35.Value
        .Fields("RBillAmount").Value = MhRealInput39.Value
        .Fields("PaidAmount").Value = MhRealInput16.Value
        .Fields("PPaidAmount").Value = MhRealInput30.Value
        .Fields("AdjustmentRemarks").Value = IIf(MhRealInput9.Value <> 0 Or MhRealInput34.Value <> 0 Or MhRealInput36.Value <> 0, TxtAdNar.Text, "")
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text11.Text, False) Then Text11.SetFocus: CheckMandatoryFields = True: Exit Function 'Finish Size
    If CheckEmpty(Text4.Text, False) Then Text4.SetFocus: CheckMandatoryFields = True: Exit Function 'Printing Size
    If CheckEmpty(Text7.Text, False) Then Text7.SetFocus: CheckMandatoryFields = True: Exit Function 'Plate Party
    If Combo3.ListIndex < 0 Then Combo3.SetFocus: CheckMandatoryFields = True: Exit Function
    If CheckEmpty(Text16.Text, False) Then Text16.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput16.Value <> 0 Then If MhRealInput16.Value <> MhRealInput10.Value + MhRealInput39.Value Then MhRealInput9.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput30.Value <> 0 Then If MhRealInput30.Value <> MhRealInput35.Value Then MhRealInput34.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput9.Value <> 0 Or MhRealInput34.Value <> 0 Or MhRealInput36.Value <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
End Function
Private Sub GetPartyRates(ByVal RateType As String)
    If MhRealInput2.Value = 0 Or CheckEmpty(SizeCode, False) Or CheckEmpty(Text13.Text, False) Then Exit Sub
    Dim PlateRate As Double, PrintRate As Double, PaperWastageRate As Double, PaperWastageMin As Long
    On Error GoTo ErrorHandler
    'Fetching Rates
    With rstFetchRate
            If .State = adStateOpen Then .Close
        If RateType = "L" Then  'Plate Rate
'Size
                .Open "SELECT TOP 1 P.* FROM AccountChild06 P INNER JOIN SizeGroupChild C ON P.[SizeGroup]=C.Code WHERE P.Code='" & PartyCode & "' AND C.[Size]='" & SizeCode & "' AND [Plate]='" & PlateCode & "' AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
''Area
'            If .RecordCount = 0 Then
'            If .State = adStateOpen Then .Close
'                .Open "SELECT TOP 1 P.* FROM AccountChild06 P WHERE P.Code='" & PartyCode & "' AND  Convert(Real,Left((Select Name From GeneralMaster Where Code=P.[SizeGroup]),5))*Convert(Real,Substring((Select Name From GeneralMaster Where Code=P.[SizeGroup]),7,5))>=Convert(Real,Left((Select Name From GeneralMaster Where Code='" & SizeCode & "'),5))*Convert(Real,Substring((Select Name From GeneralMaster Where Code='" & SizeCode & "'),7,5))  AND [Plate]='" & PlateCode & "' AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
'            End If
'Size
            If .RecordCount = 0 Then
            If .State = adStateOpen Then .Close
                .Open "SELECT TOP 1 C1.* FROM (AccountMaster P INNER JOIN AccountChild06 C1 ON P.Code=C1.Code) INNER JOIN SizeGroupChild C2 ON C1.[SizeGroup]=C2.Code WHERE [Name] LIKE '%Rate%' AND C2.[Size]='" & SizeCode & "' AND [Plate]='" & PlateCode & "' AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
            End If
''Area
'            If .RecordCount = 0 Then
'            If .State = adStateOpen Then .Close
'                .Open "SELECT TOP 1 C1.* FROM (AccountMaster P INNER JOIN AccountChild06 C1 ON P.Code=C1.Code) WHERE [Name] LIKE '%Rate%' AND  Convert(Real,Left((Select Name From GeneralMaster Where Code=C1.[SizeGroup]),5))*Convert(Real,Substring((Select Name From GeneralMaster Where Code=C1.[SizeGroup]),7,5))>=Convert(Real,Left((Select Name From GeneralMaster Where Code='" & SizeCode & "'),5))*Convert(Real,Substring((Select Name From GeneralMaster Where Code='" & SizeCode & "'),7,5))  AND [Plate]='" & PlateCode & "' AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
'            End If
            If .RecordCount > 0 Then PlateRate = Val(.Fields("Rate").Value)
        Else
    'Printing Rate
    'Size
                .Open "SELECT TOP 1 P.* FROM AccountChild05 P INNER JOIN SizeGroupChild C ON P.[SizeGroup]=C.Code WHERE P.Code='" & PartyCode & "' AND C.[Size]='" & SizeCode & "' AND [Color]='" & ColorCode & "' AND [Range]>=" & MhRealInput2.Value & " AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY [Range],wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
'    'Area
'            If .RecordCount = 0 Then
'                If .State = adStateOpen Then .Close
'                .Open "SELECT TOP 1 P.* FROM AccountChild05 P WHERE P.Code='" & PartyCode & "' AND Convert(Real,Left((Select Name From GeneralMaster Where Code=P.[SizeGroup]),5))*Convert(Real,Substring((Select Name From GeneralMaster Where Code=P.[SizeGroup]),7,5))>=Convert(Real,Left((Select Name From GeneralMaster Where Code='" & SizeCode & "'),5))*Convert(Real,Substring((Select Name From GeneralMaster Where Code='" & SizeCode & "'),7,5)) AND [Color]='" & ColorCode & "' AND [Range]>=" & MhRealInput2.Value & " AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY [Range],wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
'            End If
    'Size
            If .RecordCount = 0 Then
                If .State = adStateOpen Then .Close
                .Open "SELECT TOP 1 C1.* FROM (AccountMaster P INNER JOIN AccountChild05 C1 ON P.Code=C1.Code) INNER JOIN SizeGroupChild C2 ON C1.[SizeGroup]=C2.Code WHERE [Name] LIKE '%Rate%' AND C2.[Size]='" & SizeCode & "' AND [Color]='" & ColorCode & "' AND [Range]>=" & MhRealInput2.Value & " AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY [Range],wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
            End If
'    'Area
'            If .RecordCount = 0 Then
'                If .State = adStateOpen Then .Close
'                .Open "SELECT TOP 1 C1.* FROM (AccountMaster P INNER JOIN AccountChild05 C1 ON P.Code=C1.Code) WHERE [Name] LIKE '%Rate%' AND  Convert(Real,Left((Select Name From GeneralMaster Where Code=P.[SizeGroup]),5))*Convert(Real,Substring((Select Name From GeneralMaster Where Code=P.[SizeGroup]),7,5))>=Convert(Real,Left((Select Name From GeneralMaster Where Code='" & SizeCode & "'),5))*Convert(Real,Substring((Select Name From GeneralMaster Where Code='" & SizeCode & "'),7,5)) AND [Color]='" & ColorCode & "' AND [Range]>=" & MhRealInput2.Value & " AND wef<='" & GetDate(MhDateInput1.Text) & "' ORDER BY [Range],wef DESC", cnDatabase, adOpenKeyset, adLockReadOnly
'            End If
            If .RecordCount > 0 Then
                If RateType = "P" Then  'Print Rate
                    PrintRate = Val(.Fields("PrintingRate").Value)
                ElseIf RateType = "W" Then  'Paper Wastage (Percentage)
                    PaperWastageRate = Val(.Fields("PaperWastageRate").Value)
                ElseIf RateType = "M" Then  'Paper Wastage (Minimum Sheets)
                    PaperWastageMin = Val(.Fields("PaperWastageMin").Value)
                End If
            End If
        End If
    End With
    If RateType = "L" Then
        If Combo3.ListIndex > 0 Or MhRealInput59.Value > 0 Then 'not old
            If PlateRate > 0 Then
                If MhRealInput4.Value = 0 Then
                    MhRealInput4.Value = PlateRate
                ElseIf MhRealInput4.Value <> PlateRate Then
                    If MsgBox("Plate Rate [" & Trim(MhRealInput4.Value) & "] is different from that in Master [" & Trim(Format(PlateRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput4.Value = PlateRate
                End If
            End If
        Else
            If Plate Then MhRealInput4.Value = 0 'old 3 times usable plates
        End If
    ElseIf RateType = "P" Then
        If PrintRate > 0 Then
            If MhRealInput5.Value = 0 Then
                MhRealInput5.Value = PrintRate
            ElseIf MhRealInput5.Value <> PrintRate Then
                If MsgBox("Print Rate [" & Trim(MhRealInput5.Value) & "] is different from that in Master [" & Trim(Format(PrintRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput5.Value = PrintRate
            End If
        End If
    ElseIf RateType = "W" Then
        If PaperWastageRate > 0 Then
            If MhRealInput11.Value = 0 Then
                MhRealInput11.Value = PaperWastageRate
            ElseIf MhRealInput11.Value <> PaperWastageRate Then
                If MsgBox("Paper Wastage Rate [" & Trim(MhRealInput11.Value) & "] is different from that in Master [" & Trim(Format(PaperWastageRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput11.Value = PaperWastageRate
            End If
        End If
    ElseIf RateType = "M" Then
        If PaperWastageMin > 0 Then
            If MhRealInput47.Value = 0 Then
                MhRealInput47.Value = PaperWastageMin
            ElseIf MhRealInput47.Value <> PaperWastageMin Then
                If MsgBox("Paper Wastage Min [" & Trim(MhRealInput47.Value) & "] is different from that in Master [" & Trim(Format(PaperWastageMin, "#0")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput47.Value = PaperWastageMin
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub
Private Sub CalculateForms(ByVal FormType As String)
    Dim Forms As Double
    Forms = MhRealInput45.Value
    If FormType = "Q" Then '¼ Forms
        Forms = Forms - Int(Forms)
        If Forms > 0 Then
            Forms = IIf(InStr(1, "0.25_0.75_0.375_0.875", Forms) > 0, 1, 0)
            If MhRealInput17.Value = 0 Then
                MhRealInput17.Value = Forms
            ElseIf Forms <> MhRealInput17.Value Then
                If MsgBox("Variation in Calculated [" & Trim(Forms) & "] and Existing [" & Trim(MhRealInput17.Value) & "] ¼ Forms ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput17.Value = Forms
            End If
        End If
    ElseIf FormType = "H" Then '½ Forms
        Forms = Forms - Int(Forms)
        If Forms > 0 Then
            Forms = IIf(InStr(1, "0.5_0.75_0.625_0.875", Forms) > 0 Or Forms = (5 / 6), 1, 0)
            If MhRealInput20.Value = 0 Then
                MhRealInput20.Value = Forms
            ElseIf Forms <> MhRealInput20.Value Then
                If MsgBox("Variation in Calculated [" & Trim(Forms) & "] and Existing [" & Trim(MhRealInput20.Value) & "] ½ Forms ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput20.Value = Forms
            End If
        End If
    ElseIf FormType = "F" Then '1 Form-F&B
        Forms = Int(Forms / 2) * 2
        If MhRealInput21.Value = 0 Then
            MhRealInput21.Value = Forms
        ElseIf Forms <> MhRealInput21.Value Then
            If MsgBox("Variation in Calculated [" & Trim(Forms) & "] and Existing [" & Trim(MhRealInput21.Value) & "] 1 Forms-F&B ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput21.Value = Forms
        End If
    ElseIf FormType = "W" Then '1 Form-W&T
        Forms = Int(Forms) - Int(Forms / 2) * 2
        If MhRealInput41.Value = 0 Then
            MhRealInput41.Value = Forms
        ElseIf Forms <> MhRealInput41.Value Then
            If MsgBox("Variation in Calculated [" & Trim(Forms) & "] and Existing [" & Trim(MhRealInput41.Value) & "] 1 Forms-W&T ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput41.Value = Forms
        End If
    End If
End Sub
Private Sub CalculateTotalForms()
    If MhRealInput22.Value = 0 Then Exit Sub
    Dim TotalForms As Double
    TotalForms = MhRealInput2.Value / MhRealInput22.Value
    TotalForms = IIf(Option2.Value, 0.5, 1) * TotalForms
If RoundOffQty Then
    MhRealInput6.Value = (Int((TotalForms * 0.25) / 1000) + IIf((TotalForms * 0.25) Mod 1000 = 0, 0, 1)) * MhRealInput17.Value
    MhRealInput25.Value = (Int((TotalForms * 0.5) / 1000) + IIf((TotalForms * 0.5) Mod 1000 = 0, 0, 1)) * MhRealInput20.Value
    MhRealInput26.Value = (Int((TotalForms * 1) / 1000) + IIf((TotalForms * 1) Mod 1000 = 0, 0, 1)) * MhRealInput21.Value
    MhRealInput43.Value = (Int((TotalForms * 1) / 1000) + IIf((TotalForms * 1) Mod 1000 = 0, 0, 1)) * MhRealInput41.Value
Else
    MhRealInput6.Value = (((TotalForms * 0.25) / 1000)) * MhRealInput17.Value
    MhRealInput25.Value = (((TotalForms * 0.5) / 1000)) * MhRealInput20.Value
    MhRealInput26.Value = (((TotalForms * 1) / 1000)) * MhRealInput21.Value
    MhRealInput43.Value = (((TotalForms * 1) / 1000)) * MhRealInput41.Value
End If
End Sub
Private Sub CalculateTotalPlates()
    If MhRealInput22.Value = 0 Then Exit Sub
    Dim TotalPlates As Double, i As Byte
    For i = 1 To 4
        TotalPlates = Choose(i, MhRealInput17.Value, MhRealInput20.Value, MhRealInput21.Value, MhRealInput41.Value)
        TotalPlates = TotalPlates / MhRealInput22.Value
        TotalPlates = IIf(Option2.Value, 0.5, 1) * TotalPlates
        TotalPlates = Int(TotalPlates) + IIf(TotalPlates - Int(TotalPlates) = 0.5, 1, 0)
        Choose(i, MhRealInput3, MhRealInput23, MhRealInput24, MhRealInput46).Value = TotalPlates * Color
    Next
End Sub
Private Sub CalculateConsumption()
    If SPU = 0 Or MhRealInput27.Value = 0 Then Exit Sub
    Dim Forms As Double, W As Double, Consumption As Double
    Forms = MhRealInput17.Value * 0.25 + MhRealInput20.Value * 0.5 + (MhRealInput21.Value + MhRealInput41.Value) * 1
    'Consumption & Wastage Calculation (For Single Form)
    Consumption = (MhRealInput1.Value / 1000) * 1 * 500 'Sheets
    W = (Consumption * MhRealInput11.Value) / 100 'Sheets
    If W < MhRealInput47.Value Then W = MhRealInput47.Value 'Comparison with Minimum Wastage
    'Final Consumption & Wastage (for All Forms)
    W = W * Forms 'Sheets
    Consumption = Consumption * Forms + W 'Sheets
    Consumption = Consumption / MhRealInput27.Value 'Sheets
    MhRealInput49.Value = IIf(MhRealInput52.Value > 0, Round(((MhRealInput52.Value / 25.4) * inWidth * GSM) / 3100, 3), Wt) * (Consumption / SPU) 'Consumption-Kg
    MhRealInput13.Value = CLng(Int(Consumption / SPU)) + ((Consumption Mod SPU) / 1000) 'UOM
    MhRealInput48.Value = CLng(Int(W / SPU)) + ((W Mod SPU) / 1000)
End Sub
Private Sub CalculateAmount()
    MhRealInput7.Value = IIf(Combo3.ListIndex > 0, MhRealInput3.Value + MhRealInput23.Value + MhRealInput24.Value + MhRealInput46.Value + MhRealInput59.Value, MhRealInput59.Value) * MhRealInput4.Value 'Plate Amount
    MhRealInput8.Value = Val(Left(Text13.Text, 2)) * (MhRealInput6.Value + MhRealInput25.Value + MhRealInput26.Value + MhRealInput43.Value) * MhRealInput5.Value 'Print Amount
    CalculateTotalAmount
End Sub
Private Sub CalculateTotalAmount()
    MhRealInput29.Value = (MhRealInput7.Value + MhRealInput34.Value) * MhRealInput28.Value / 100 'GST Plate
    MhRealInput18.Value = (MhRealInput8.Value + MhRealInput9.Value) * MhRealInput14.Value / 100 'GST Printing
    MhRealInput38.Value = (MhRealInput33.Value + MhRealInput36.Value) * MhRealInput37.Value / 100 'GST Paper
    MhRealInput10.Value = Round(MhRealInput8.Value + MhRealInput9.Value + MhRealInput18.Value, 0) 'Total Printing Amount
    MhRealInput35.Value = Round(MhRealInput7.Value + MhRealInput29.Value + MhRealInput34.Value, 0) 'Total Plate Amount
    MhRealInput39.Value = Round(MhRealInput33.Value + MhRealInput38.Value + MhRealInput36.Value, 0) 'Total Paper Amount
End Sub
Private Sub LoadMasterList(Optional ByVal LoadSelected As Boolean)
    If rstGeneralList.State = adStateOpen Then rstGeneralList.Close 'Size/Color/Plate Master List
    rstGeneralList.Open "SELECT Name As Col0,Value1,Code From GeneralMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstGeneralList.ActiveConnection = Nothing
    If rstPaperList.State = adStateOpen Then rstPaperList.Close
    If LoadSelected Then
        rstPaperList.Open "SELECT * FROM (SELECT LTRIM(P.Name)+' (UOM : '+LTRIM(C.Name)+'='+LTRIM(C.Value1)+')' As Col0,FORMAT(dbo.ufnGetPaperStock('" & IIf(chkPaper.Value, PartyCode, "000000") & "',P.Code,'PO','" & VchCode & "','" & GetDate(MhDateInput1.Text) & "'),'#0.000') As Col1,C.Name As UOM,GSM,inWidth,inLength,P.Code,C.Value1 As SPU,[Form],[Weight/Unit] As Wt,LTRIM(Q.Name) As Quality,Grade FROM (PaperMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code) INNER JOIN GeneralMaster Q ON P.Quality=Q.Code) As Tbl WHERE CONVERT(DECIMAL(12,3),Col1)<>0 ORDER BY Col0", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstPaperList.Open "SELECT LTRIM(P.Name)+' (UOM : '+LTRIM(C.Name)+'='+LTRIM(C.Value1)+')' As Col0,FORMAT(0,'#0.000') As Col1,C.Name As UOM,GSM,inWidth,inLength,P.Code,C.Value1 As SPU,[Form],[Weight/Unit] As Wt,LTRIM(Q.Name) As Quality,Grade FROM (PaperMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code) INNER JOIN GeneralMaster Q ON P.Quality=Q.Code ORDER BY Col0", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    rstPaperList.ActiveConnection = Nothing
    If rstPlateMakerList.State = adStateOpen Then rstPlateMakerList.Close
    rstPlateMakerList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstPlateMakerList.ActiveConnection = Nothing
    If rstElementList.State = adStateOpen Then rstElementList.Close
    rstElementList.Open "SELECT Name As Col0,Code FROM ElementMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstElementList.ActiveConnection = Nothing
End Sub
Private Sub LockFields(ByVal bVal As Boolean)
    Dim O As Object
    For Each O In Me
        If TypeName(O) = "TextBox" Then
            O.Locked = bVal
        ElseIf TypeName(O) = "TDBNumber" Then
            O.ReadOnly = bVal
        ElseIf TypeName(O) = "TDBDate" Then
            O.ReadOnly = bVal
        ElseIf TypeName(O) = "ComboBox" Or TypeName(O) = "OptionButton" Then
            O.Enabled = Not bVal
        End If
    Next
    If Not bVal Then Text5.Locked = True: Text2.Locked = True: Text7.Locked = True: Text14.Locked = True: Text11.Locked = True: Text4.Locked = True: Text1.Locked = True: Text13.Locked = True: Text16.Locked = True: MhRealInput45.ReadOnly = True: MhRealInput3.ReadOnly = True: MhRealInput23.ReadOnly = True: MhRealInput24.ReadOnly = True: MhRealInput46.ReadOnly = True: MhRealInput7.ReadOnly = True: MhRealInput6.ReadOnly = True: MhRealInput25.ReadOnly = True: MhRealInput26.ReadOnly = True: MhRealInput43.ReadOnly = True: MhRealInput8.ReadOnly = True: MhRealInput48.ReadOnly = True: MhRealInput13.ReadOnly = True: MhRealInput33.ReadOnly = True: MhRealInput49.ReadOnly = True: MhRealInput18.ReadOnly = True: MhRealInput10.ReadOnly = True: MhRealInput29.ReadOnly = True: MhRealInput35.ReadOnly = True: MhRealInput38.ReadOnly = True: MhRealInput39.ReadOnly = True
End Sub
Private Sub SetButtons(ByVal bVal As Boolean)
    Command1(0).Enabled = bVal
    Command1(1).Enabled = bVal
    Command1(2).Enabled = bVal
    Command1(3).Enabled = Not bVal
    Command1(4).Enabled = Not bVal
End Sub
Private Function CheckDuplicateElement() As Boolean
    Dim dblBookMark As Double
    With rstBookPOChild05c
        If .RecordCount = 0 Then Exit Function
        If Not (.EOF Or .BOF) Then dblBookMark = .Bookmark 'current record no.
        .MoveFirst
        Do Until .EOF
            If Me.Tag = "A" Then 'Add
                If .Fields("ElementName").Value = Trim(Text14.Text) Then CheckDuplicateElement = True: Exit Do
            ElseIf Left(Me.Tag, 1) = "E" Then 'Edit
                If .Fields("ElementName").Value = Trim(Text14.Text) And .Bookmark <> dblBookMark Then CheckDuplicateElement = True: Exit Do
            End If
            .MoveNext
        Loop
        If dblBookMark <> 0 Then .Bookmark = dblBookMark Else .MoveLast
    End With
End Function
Private Function MaxUps(ByVal Position As String) As Integer
    Dim FL As Double, FR As Double, PL As Double, PR As Double, PW As Double, Ups01 As Integer, Ups02 As Integer, Ups03 As Integer
    If Position = "T" Then
        If CheckEmpty(FinishSizeCode, False) Or CheckEmpty(SizeCode, False) Then MaxUps = 0: Exit Function
        PL = Val(Left(Text4.Text, InStr(1, Text4.Text, "X") - 1)) + 1: PR = Val(Mid(Text4.Text, InStr(1, Text4.Text, "X") + 1, 5)) + 1: FL = Val(Left(Text11.Text, InStr(1, Text11.Text, "X") - 1)): FR = Val(Mid(Text11.Text, InStr(1, Text11.Text, "X") + 1, 5))
        Ups01 = Int(IIf(PL > PR, PL, PR) / IIf(FL > FR, FL, FR)) * Int(IIf(PL < PR, PL, PR) / IIf(FL < FR, FL, FR)): Ups02 = Int(IIf(PL < PR, PL, PR) / IIf(FL > FR, FL, FR)) * Int(IIf(PL > PR, PL, PR) / IIf(FL < FR, FL, FR))
        MaxUps = IIf(Ups01 > Ups02, Ups01, Ups02)
    Else
            If CheckEmpty(PaperCode, False) Or CheckEmpty(SizeCode, False) Then MaxUps = 0: Exit Function
            FL = Val(Left(Text4.Text, InStr(1, Text4.Text, "X") - 1)): FR = Val(Mid(Text4.Text, InStr(1, Text4.Text, "X") + 1, 5)): PL = IIf(PaperForm = "R", MhRealInput52.Value / 25.4, inLength): PW = inWidth 'Printing Size Left & Right + Paper Length & Width
            If Abs(FL - PL) <= 1 Then PL = FL: If Abs(FR - PL) <= 1 Then PL = FR: If Abs(FL - PW) <= 1 Then PW = FL: If Abs(FR - PW) <= 1 Then PW = FR
            Ups01 = Int(IIf(PW > PL, PW, PL) / IIf(FL > FR, FL, FR)) * Int(IIf(PW < PL, PW, PL) / IIf(FL < FR, FL, FR)): Ups02 = Int(IIf(PW > PL, PW, PL) / IIf(FL < FR, FL, FR)) * Int(IIf(PW < PL, PW, PL) / IIf(FL > FR, FL, FR)): Ups03 = Int((PW * PL) / (FL * FR))
            MaxUps = IIf(Ups03 > IIf(Ups01 > Ups02, Ups01, Ups02), Ups03, IIf(Ups01 > Ups02, Ups01, Ups02))
            If MaxUps = 0 Then MaxUps = 1
    End If
End Function
Private Sub CheckPaperSize()
    If CheckEmpty(SizeCode, False) Then Exit Sub 'Printing Size
    Dim FL As Double, FR As Double, PL As Double, PW As Double
    FL = Val(Left(Text4.Text, InStr(1, Text4.Text, "X") - 1)): FR = Val(Mid(Text4.Text, InStr(1, Text4.Text, "X") + 1, 5)): PL = IIf(PaperForm = "R", MhRealInput52.Value / 25.4, inLength): PW = inWidth 'Printing Size Left & Right + Paper Length & Width
    If Abs(FL - PL) <= 1 Then PL = FL: If Abs(FR - PL) <= 1 Then PL = FR: If Abs(FL - PW) <= 1 Then PW = FL: If Abs(FR - PW) <= 1 Then PW = FR
    Call CalcUps(PL * PW, FL * FR)
End Sub
