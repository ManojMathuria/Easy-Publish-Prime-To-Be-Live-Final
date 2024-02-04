VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmBookPOChild01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Elements Digital Printing  Order Details"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11730
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
   ScaleHeight     =   8640
   ScaleWidth      =   11730
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   10870
      Picture         =   "BookPOChild01.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   128
      TabStop         =   0   'False
      ToolTipText     =   "Delete Item Pic"
      Top             =   360
      Width           =   375
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   930
      Index           =   2
      Left            =   10200
      TabIndex        =   127
      Top             =   1320
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   1640
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
      Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D->Delete  Ctrl+S->Save"
      AutoSize        =   -1  'True
      FillColor       =   8421504
      TextColor       =   16777215
      Picture         =   "BookPOChild01.frx":0102
      Multiline       =   -1  'True
      GlobalMem       =   -1  'True
      Picture         =   "BookPOChild01.frx":011E
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   11250
      Picture         =   "BookPOChild01.frx":013A
      Style           =   1  'Graphical
      TabIndex        =   125
      ToolTipText     =   "Save"
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdUpload 
      Height          =   375
      Left            =   10120
      Picture         =   "BookPOChild01.frx":023C
      Style           =   1  'Graphical
      TabIndex        =   124
      TabStop         =   0   'False
      ToolTipText     =   "Upload Item Pic"
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdView 
      Height          =   375
      Left            =   10500
      Picture         =   "BookPOChild01.frx":057E
      Style           =   1  'Graphical
      TabIndex        =   123
      TabStop         =   0   'False
      ToolTipText     =   "View Item Pic"
      Top             =   360
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   8400
      Left            =   120
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   120
      Width           =   9975
      _Version        =   65536
      _ExtentX        =   17595
      _ExtentY        =   14817
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
      Picture         =   "BookPOChild01.frx":0AB0
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Left            =   3360
         TabIndex        =   129
         Top             =   1610
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
         Picture         =   "BookPOChild01.frx":0ACC
         Picture         =   "BookPOChild01.frx":0AE8
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
         Height          =   330
         Left            =   5040
         TabIndex        =   13
         Top             =   1610
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":0B04
         Caption         =   "BookPOChild01.frx":0B24
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":0B90
         Keys            =   "BookPOChild01.frx":0BAE
         Spin            =   "BookPOChild01.frx":0BF8
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel30 
         Height          =   330
         Left            =   120
         TabIndex        =   130
         Top             =   1610
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
         Caption         =   " Imposition"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":0C20
         Picture         =   "BookPOChild01.frx":0C3C
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   290
         Index           =   0
         Left            =   120
         Picture         =   "BookPOChild01.frx":0C58
         Style           =   1  'Graphical
         TabIndex        =   120
         TabStop         =   0   'False
         ToolTipText     =   "Add [Ctrl+A]"
         Top             =   4485
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   290
         Index           =   1
         Left            =   435
         Picture         =   "BookPOChild01.frx":118A
         Style           =   1  'Graphical
         TabIndex        =   119
         TabStop         =   0   'False
         ToolTipText     =   "Edit [Ctrl+E]"
         Top             =   4485
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   290
         Index           =   2
         Left            =   750
         Picture         =   "BookPOChild01.frx":16BC
         Style           =   1  'Graphical
         TabIndex        =   118
         TabStop         =   0   'False
         ToolTipText     =   "Delete [Ctrl+D]"
         Top             =   4485
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   290
         Index           =   3
         Left            =   1065
         Picture         =   "BookPOChild01.frx":17BE
         Style           =   1  'Graphical
         TabIndex        =   117
         TabStop         =   0   'False
         ToolTipText     =   "Save [Ctrl+S]"
         Top             =   4485
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   290
         Index           =   4
         Left            =   1380
         Picture         =   "BookPOChild01.frx":18C0
         Style           =   1  'Graphical
         TabIndex        =   116
         TabStop         =   0   'False
         ToolTipText     =   "Cancel [Esc]"
         Top             =   4485
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
         Top             =   980
         Width           =   4815
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput35 
         Height          =   330
         Left            =   5820
         TabIndex        =   35
         ToolTipText     =   "Wastage Rate %"
         Top             =   3405
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":19C2
         Caption         =   "BookPOChild01.frx":19E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":1A4E
         Keys            =   "BookPOChild01.frx":1A6C
         Spin            =   "BookPOChild01.frx":1AB6
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
         ValueVT         =   98435077
         Value           =   4
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   93
         Top             =   7635
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
         Picture         =   "BookPOChild01.frx":1ADE
         Picture         =   "BookPOChild01.frx":1AFA
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Left            =   120
         TabIndex        =   94
         Top             =   7950
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
         Picture         =   "BookPOChild01.frx":1B16
         Picture         =   "BookPOChild01.frx":1B32
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
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1290
         Width           =   1575
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
         Top             =   665
         Width           =   4815
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   6600
         TabIndex        =   87
         Top             =   2235
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
         Caption         =   " Impressions/Set"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":1B4E
         Picture         =   "BookPOChild01.frx":1B6A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   92
         Top             =   115
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
         Picture         =   "BookPOChild01.frx":1B86
         Picture         =   "BookPOChild01.frx":1BA2
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
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   9990
         Width           =   1575
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1290
         Width           =   4815
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
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   9
         Top             =   1610
         Width           =   1575
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
         TabIndex        =   62
         Top             =   7950
         Width           =   8055
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
         TabIndex        =   61
         Top             =   7635
         Width           =   8055
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
         Top             =   115
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   55
         Top             =   7155
         Width           =   1575
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
         MaxLength       =   80
         TabIndex        =   31
         Top             =   3090
         Width           =   8055
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
         Left            =   8280
         MaxLength       =   40
         TabIndex        =   5
         Top             =   665
         Width           =   1575
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   3360
         TabIndex        =   64
         Top             =   115
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
         Picture         =   "BookPOChild01.frx":1BBE
         Picture         =   "BookPOChild01.frx":1BDA
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   6600
         TabIndex        =   65
         Top             =   1610
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
         Caption         =   " Printing Size"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":1BF6
         Picture         =   "BookPOChild01.frx":1C12
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   120
         TabIndex        =   66
         Top             =   1920
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
         Caption         =   " Quantity"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":1C2E
         Picture         =   "BookPOChild01.frx":1C4A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   6600
         TabIndex        =   69
         Top             =   9060
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
         Picture         =   "BookPOChild01.frx":1C66
         Picture         =   "BookPOChild01.frx":1C82
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   6600
         TabIndex        =   70
         Top             =   2535
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
         Picture         =   "BookPOChild01.frx":1C9E
         Picture         =   "BookPOChild01.frx":1CBA
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   6600
         TabIndex        =   71
         Top             =   6345
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
         Caption         =   " Total Amt-Ptg"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":1CD6
         Picture         =   "BookPOChild01.frx":1CF2
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Left            =   120
         TabIndex        =   72
         Top             =   3405
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
         Picture         =   "BookPOChild01.frx":1D0E
         Picture         =   "BookPOChild01.frx":1D2A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
         Height          =   330
         Left            =   6600
         TabIndex        =   73
         Top             =   3720
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
         Picture         =   "BookPOChild01.frx":1D46
         Picture         =   "BookPOChild01.frx":1D62
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   6600
         TabIndex        =   74
         Top             =   7155
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
         Picture         =   "BookPOChild01.frx":1D7E
         Picture         =   "BookPOChild01.frx":1D9A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   3360
         TabIndex        =   75
         Top             =   7155
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
         Picture         =   "BookPOChild01.frx":1DB6
         Picture         =   "BookPOChild01.frx":1DD2
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   6600
         TabIndex        =   76
         Top             =   115
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
         Picture         =   "BookPOChild01.frx":1DEE
         Picture         =   "BookPOChild01.frx":1E0A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
         Height          =   330
         Left            =   3360
         TabIndex        =   77
         Top             =   6345
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
         Picture         =   "BookPOChild01.frx":1E26
         Picture         =   "BookPOChild01.frx":1E42
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   8280
         TabIndex        =   23
         Top             =   2235
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":1E5E
         Caption         =   "BookPOChild01.frx":1E7E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":1EEA
         Keys            =   "BookPOChild01.frx":1F08
         Spin            =   "BookPOChild01.frx":1F52
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   1800
         TabIndex        =   27
         ToolTipText     =   "Print Rate Front"
         Top             =   2535
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":1F7A
         Caption         =   "BookPOChild01.frx":1F9A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":2006
         Keys            =   "BookPOChild01.frx":2024
         Spin            =   "BookPOChild01.frx":206E
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   1800
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Plate Rate Front"
         Top             =   9060
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":2096
         Caption         =   "BookPOChild01.frx":20B6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":2122
         Keys            =   "BookPOChild01.frx":2140
         Spin            =   "BookPOChild01.frx":218A
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   330
         Left            =   1800
         TabIndex        =   43
         ToolTipText     =   "Print"
         Top             =   6345
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":21B2
         Caption         =   "BookPOChild01.frx":21D2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":223E
         Keys            =   "BookPOChild01.frx":225C
         Spin            =   "BookPOChild01.frx":22A6
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   8280
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   9060
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":22CE
         Caption         =   "BookPOChild01.frx":22EE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":235A
         Keys            =   "BookPOChild01.frx":2378
         Spin            =   "BookPOChild01.frx":23C2
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
         Height          =   330
         Left            =   8280
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2535
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":23EA
         Caption         =   "BookPOChild01.frx":240A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":2476
         Keys            =   "BookPOChild01.frx":2494
         Spin            =   "BookPOChild01.frx":24DE
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
         Height          =   330
         Left            =   8280
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   6345
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":2506
         Caption         =   "BookPOChild01.frx":2526
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":2592
         Keys            =   "BookPOChild01.frx":25B0
         Spin            =   "BookPOChild01.frx":25FA
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
         Height          =   330
         Left            =   8280
         TabIndex        =   57
         Top             =   7155
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":2622
         Caption         =   "BookPOChild01.frx":2642
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":26AE
         Keys            =   "BookPOChild01.frx":26CC
         Spin            =   "BookPOChild01.frx":2716
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   1800
         TabIndex        =   15
         ToolTipText     =   "Actual"
         Top             =   1920
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":273E
         Caption         =   "BookPOChild01.frx":275E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":27CA
         Keys            =   "BookPOChild01.frx":27E8
         Spin            =   "BookPOChild01.frx":2832
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
         Height          =   330
         Left            =   1800
         TabIndex        =   33
         Top             =   3405
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":285A
         Caption         =   "BookPOChild01.frx":287A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":28E6
         Keys            =   "BookPOChild01.frx":2904
         Spin            =   "BookPOChild01.frx":294E
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   5040
         TabIndex        =   34
         ToolTipText     =   "Wastage Rate %"
         Top             =   3405
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":2976
         Caption         =   "BookPOChild01.frx":2996
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":2A02
         Keys            =   "BookPOChild01.frx":2A20
         Spin            =   "BookPOChild01.frx":2A6A
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
         ValueVT         =   159711237
         Value           =   4
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   5040
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   6345
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":2A92
         Caption         =   "BookPOChild01.frx":2AB2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":2B1E
         Keys            =   "BookPOChild01.frx":2B3C
         Spin            =   "BookPOChild01.frx":2B86
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
         Height          =   330
         Left            =   4200
         TabIndex        =   44
         Top             =   6345
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":2BAE
         Caption         =   "BookPOChild01.frx":2BCE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":2C3A
         Keys            =   "BookPOChild01.frx":2C58
         Spin            =   "BookPOChild01.frx":2CA2
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
         ValueVT         =   159776773
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   5040
         TabIndex        =   2
         Top             =   115
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild01.frx":2CCA
         Caption         =   "BookPOChild01.frx":2DE2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":2E4E
         Keys            =   "BookPOChild01.frx":2E6C
         Spin            =   "BookPOChild01.frx":2ECA
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
         Left            =   8280
         TabIndex        =   3
         Top             =   115
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild01.frx":2EF2
         Caption         =   "BookPOChild01.frx":300A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":3076
         Keys            =   "BookPOChild01.frx":3094
         Spin            =   "BookPOChild01.frx":30F2
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
         Left            =   5040
         TabIndex        =   56
         Top             =   7155
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild01.frx":311A
         Caption         =   "BookPOChild01.frx":3232
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":329E
         Keys            =   "BookPOChild01.frx":32BC
         Spin            =   "BookPOChild01.frx":331A
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   3360
         TabIndex        =   79
         Top             =   9660
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
         Picture         =   "BookPOChild01.frx":3342
         Picture         =   "BookPOChild01.frx":335E
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
         Height          =   330
         Left            =   5040
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   9660
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":337A
         Caption         =   "BookPOChild01.frx":339A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":3406
         Keys            =   "BookPOChild01.frx":3424
         Spin            =   "BookPOChild01.frx":346E
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
         Height          =   330
         Left            =   4200
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   9660
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":3496
         Caption         =   "BookPOChild01.frx":34B6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":3522
         Keys            =   "BookPOChild01.frx":3540
         Spin            =   "BookPOChild01.frx":358A
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
         ValueVT         =   159842309
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput23 
         Height          =   330
         Left            =   8280
         TabIndex        =   36
         ToolTipText     =   "Wastage Min(Sheets)"
         Top             =   3405
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":35B2
         Caption         =   "BookPOChild01.frx":35D2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":363E
         Keys            =   "BookPOChild01.frx":365C
         Spin            =   "BookPOChild01.frx":36A6
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel35 
         Height          =   330
         Left            =   6600
         TabIndex        =   80
         Top             =   9990
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
         Picture         =   "BookPOChild01.frx":36CE
         Picture         =   "BookPOChild01.frx":36EA
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel36 
         Height          =   330
         Left            =   3360
         TabIndex        =   81
         Top             =   9990
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
         Picture         =   "BookPOChild01.frx":3706
         Picture         =   "BookPOChild01.frx":3722
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
         Height          =   330
         Left            =   8280
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   9990
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":373E
         Caption         =   "BookPOChild01.frx":375E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":37CA
         Keys            =   "BookPOChild01.frx":37E8
         Spin            =   "BookPOChild01.frx":3832
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput4 
         Height          =   330
         Left            =   5040
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   9990
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild01.frx":385A
         Caption         =   "BookPOChild01.frx":3972
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":39DE
         Keys            =   "BookPOChild01.frx":39FC
         Spin            =   "BookPOChild01.frx":3A5A
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
         Height          =   330
         Left            =   1800
         TabIndex        =   40
         Top             =   4035
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":3A82
         Caption         =   "BookPOChild01.frx":3AA2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":3B0E
         Keys            =   "BookPOChild01.frx":3B2C
         Spin            =   "BookPOChild01.frx":3B76
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
         Height          =   330
         Left            =   1800
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Plate"
         Top             =   9660
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":3B9E
         Caption         =   "BookPOChild01.frx":3BBE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":3C2A
         Keys            =   "BookPOChild01.frx":3C48
         Spin            =   "BookPOChild01.frx":3C92
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput28 
         Height          =   330
         Left            =   8280
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   9660
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":3CBA
         Caption         =   "BookPOChild01.frx":3CDA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":3D46
         Keys            =   "BookPOChild01.frx":3D64
         Spin            =   "BookPOChild01.frx":3DAE
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel40 
         Height          =   330
         Left            =   6600
         TabIndex        =   82
         Top             =   9660
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
         Picture         =   "BookPOChild01.frx":3DD6
         Picture         =   "BookPOChild01.frx":3DF2
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput33 
         Height          =   330
         Left            =   5040
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":3E0E
         Caption         =   "BookPOChild01.frx":3E2E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":3E9A
         Keys            =   "BookPOChild01.frx":3EB8
         Spin            =   "BookPOChild01.frx":3F02
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel46 
         Height          =   330
         Left            =   6600
         TabIndex        =   101
         Top             =   1290
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
         Caption         =   " Finish Size"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":3F2A
         Picture         =   "BookPOChild01.frx":3F46
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput34 
         Height          =   330
         Left            =   5040
         TabIndex        =   28
         ToolTipText     =   "Print Rate Back"
         Top             =   2535
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":3F62
         Caption         =   "BookPOChild01.frx":3F82
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":3FEE
         Keys            =   "BookPOChild01.frx":400C
         Spin            =   "BookPOChild01.frx":4056
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel45 
         Height          =   330
         Left            =   3360
         TabIndex        =   102
         Top             =   2535
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
         Caption         =   " Print Rate-Back"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":407E
         Picture         =   "BookPOChild01.frx":409A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   103
         Top             =   8760
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
         Caption         =   " Total Plates-Front"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":40B6
         Picture         =   "BookPOChild01.frx":40D2
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   1800
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   8760
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":40EE
         Caption         =   "BookPOChild01.frx":410E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":417A
         Keys            =   "BookPOChild01.frx":4198
         Spin            =   "BookPOChild01.frx":41E2
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   8280
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":420A
         Caption         =   "BookPOChild01.frx":422A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":4296
         Keys            =   "BookPOChild01.frx":42B4
         Spin            =   "BookPOChild01.frx":42FE
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
         Height          =   330
         Left            =   5040
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   4035
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":4326
         Caption         =   "BookPOChild01.frx":4346
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":43B2
         Keys            =   "BookPOChild01.frx":43D0
         Spin            =   "BookPOChild01.frx":441A
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
         Height          =   330
         Left            =   3360
         TabIndex        =   104
         Top             =   4035
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
         Picture         =   "BookPOChild01.frx":4442
         Picture         =   "BookPOChild01.frx":445E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel48 
         Height          =   330
         Left            =   3360
         TabIndex        =   105
         Top             =   3720
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
         Caption         =   " Paper Wastage"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":447A
         Picture         =   "BookPOChild01.frx":4496
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel50 
         Height          =   330
         Left            =   120
         TabIndex        =   106
         Top             =   3720
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
         Picture         =   "BookPOChild01.frx":44B2
         Picture         =   "BookPOChild01.frx":44CE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel39 
         Height          =   330
         Left            =   120
         TabIndex        =   99
         Top             =   9660
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
         Picture         =   "BookPOChild01.frx":44EA
         Picture         =   "BookPOChild01.frx":4506
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel37 
         Height          =   330
         Left            =   120
         TabIndex        =   98
         Top             =   4035
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
         Picture         =   "BookPOChild01.frx":4522
         Picture         =   "BookPOChild01.frx":453E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel34 
         Height          =   330
         Left            =   120
         TabIndex        =   97
         Top             =   9990
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
         Picture         =   "BookPOChild01.frx":455A
         Picture         =   "BookPOChild01.frx":4576
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   90
         Top             =   7155
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
         Picture         =   "BookPOChild01.frx":4592
         Picture         =   "BookPOChild01.frx":45AE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Left            =   3360
         TabIndex        =   89
         Top             =   3405
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
         Caption         =   " Wastage %-F&&B"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":45CA
         Picture         =   "BookPOChild01.frx":45E6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   120
         TabIndex        =   88
         Top             =   3090
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
         Picture         =   "BookPOChild01.frx":4602
         Picture         =   "BookPOChild01.frx":461E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   120
         TabIndex        =   85
         Top             =   6345
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
         Picture         =   "BookPOChild01.frx":463A
         Picture         =   "BookPOChild01.frx":4656
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   91
         Top             =   665
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
         Picture         =   "BookPOChild01.frx":4672
         Picture         =   "BookPOChild01.frx":468E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel31 
         Height          =   330
         Left            =   120
         TabIndex        =   95
         Top             =   1290
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
         Picture         =   "BookPOChild01.frx":46AA
         Picture         =   "BookPOChild01.frx":46C6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   120
         TabIndex        =   68
         Top             =   2535
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
         Caption         =   " Print Rate-Front"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":46E2
         Picture         =   "BookPOChild01.frx":46FE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   120
         TabIndex        =   67
         Top             =   9060
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
         Caption         =   " Plate Rate-Front"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":471A
         Picture         =   "BookPOChild01.frx":4736
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel33 
         Height          =   330
         Left            =   120
         TabIndex        =   96
         Top             =   8445
         Visible         =   0   'False
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
         Picture         =   "BookPOChild01.frx":4752
         Picture         =   "BookPOChild01.frx":476E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   6600
         TabIndex        =   86
         Top             =   660
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
         Caption         =   " Ref No."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":478A
         Picture         =   "BookPOChild01.frx":47A6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel44 
         Height          =   330
         Left            =   6600
         TabIndex        =   107
         Top             =   4035
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
         Caption         =   " Consumption-Kgs"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":47C2
         Picture         =   "BookPOChild01.frx":47DE
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput39 
         Height          =   330
         Left            =   8280
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   4035
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":47FA
         Caption         =   "BookPOChild01.frx":481A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":4886
         Keys            =   "BookPOChild01.frx":48A4
         Spin            =   "BookPOChild01.frx":48EE
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
         Height          =   330
         Left            =   1800
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   3720
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
         Picture         =   "BookPOChild01.frx":4916
         Begin VB.CheckBox chkPaper 
            Caption         =   "Check1"
            Height          =   210
            Left            =   690
            TabIndex        =   30
            Top             =   80
            Width           =   210
         End
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
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   1610
         Width           =   1575
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput36 
         Height          =   330
         Left            =   9060
         TabIndex        =   37
         ToolTipText     =   "Wastage Min(Sheets)"
         Top             =   3405
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":4932
         Caption         =   "BookPOChild01.frx":4952
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":49BE
         Keys            =   "BookPOChild01.frx":49DC
         Spin            =   "BookPOChild01.frx":4A26
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel49 
         Height          =   330
         Left            =   6600
         TabIndex        =   110
         Top             =   3405
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
         Caption         =   " Wastage Sheet-F&&B"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":4A4E
         Picture         =   "BookPOChild01.frx":4A6A
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput38 
         Height          =   330
         Left            =   5040
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Plate Rate Front"
         Top             =   9060
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":4A86
         Caption         =   "BookPOChild01.frx":4AA6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":4B12
         Keys            =   "BookPOChild01.frx":4B30
         Spin            =   "BookPOChild01.frx":4B7A
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel51 
         Height          =   330
         Left            =   3360
         TabIndex        =   111
         Top             =   9060
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
         Caption         =   " Plate Rate-Back"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":4BA2
         Picture         =   "BookPOChild01.frx":4BBE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel52 
         Height          =   330
         Left            =   120
         TabIndex        =   112
         Top             =   980
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
         Picture         =   "BookPOChild01.frx":4BDA
         Picture         =   "BookPOChild01.frx":4BF6
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1365
         Left            =   120
         TabIndex        =   0
         Top             =   4785
         Width           =   9750
         _ExtentX        =   17198
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
         ColumnCount     =   5
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
            DataField       =   "Ups"
            Caption         =   "Ups"
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
            DataField       =   "PrintAmount"
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
               ColumnWidth     =   3449.764
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   2220.094
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   2220.094
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   374.74
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   915.024
            EndProperty
         EndProperty
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel53 
         Height          =   330
         Left            =   6600
         TabIndex        =   113
         Top             =   1920
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
         Caption         =   " No. of Set"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":4C12
         Picture         =   "BookPOChild01.frx":4C2E
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput40 
         Height          =   330
         Left            =   8280
         TabIndex        =   14
         ToolTipText     =   "No. of Sets"
         Top             =   1920
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":4C4A
         Caption         =   "BookPOChild01.frx":4C6A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":4CD6
         Keys            =   "BookPOChild01.frx":4CF4
         Spin            =   "BookPOChild01.frx":4D3E
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
         ValueVT         =   5
         Value           =   1
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel54 
         Height          =   330
         Left            =   6600
         TabIndex        =   114
         Top             =   980
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
         Caption         =   " No. of Pages"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":4D66
         Picture         =   "BookPOChild01.frx":4D82
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput41 
         Height          =   330
         Left            =   8280
         TabIndex        =   7
         Top             =   980
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":4D9E
         Caption         =   "BookPOChild01.frx":4DBE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":4E2A
         Keys            =   "BookPOChild01.frx":4E48
         Spin            =   "BookPOChild01.frx":4E92
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
         Format          =   "####0"
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
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel55 
         Height          =   330
         Left            =   3360
         TabIndex        =   115
         Top             =   1920
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
         Caption         =   " Actual Ptg. Sheets"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":4EBA
         Picture         =   "BookPOChild01.frx":4ED6
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput42 
         Height          =   330
         Left            =   5040
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":4EF2
         Caption         =   "BookPOChild01.frx":4F12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":4F7E
         Keys            =   "BookPOChild01.frx":4F9C
         Spin            =   "BookPOChild01.frx":4FE6
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel56 
         Height          =   330
         Left            =   3360
         TabIndex        =   121
         Top             =   8745
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
         Caption         =   " Total Plates-Back"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":500E
         Picture         =   "BookPOChild01.frx":502A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel57 
         Height          =   330
         Left            =   6600
         TabIndex        =   122
         Top             =   9330
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
         Caption         =   " Reel Cut Off (mm)"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":5046
         Picture         =   "BookPOChild01.frx":5062
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput52 
         Height          =   330
         Left            =   8280
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   9330
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":507E
         Caption         =   "BookPOChild01.frx":509E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":510A
         Keys            =   "BookPOChild01.frx":5128
         Spin            =   "BookPOChild01.frx":5172
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
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
         MaxLength       =   40
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   8445
         Width           =   4815
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
         Height          =   330
         Left            =   1800
         TabIndex        =   131
         ToolTipText     =   "Front"
         Top             =   2235
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":519A
         Caption         =   "BookPOChild01.frx":51BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":5226
         Keys            =   "BookPOChild01.frx":5244
         Spin            =   "BookPOChild01.frx":528E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   8
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   6600
         TabIndex        =   132
         Top             =   8445
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
         Caption         =   " Plate Type-F&&B"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":52B6
         Picture         =   "BookPOChild01.frx":52D2
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
         Height          =   330
         Left            =   5040
         TabIndex        =   133
         ToolTipText     =   "Back"
         Top             =   2235
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":52EE
         Caption         =   "BookPOChild01.frx":530E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":537A
         Keys            =   "BookPOChild01.frx":5398
         Spin            =   "BookPOChild01.frx":53E2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   8
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1878589445
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel47 
         Height          =   330
         Left            =   3360
         TabIndex        =   134
         Top             =   2235
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
         Caption         =   " Ptg. Color-Back"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":540A
         Picture         =   "BookPOChild01.frx":5426
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   135
         Top             =   2235
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
         Caption         =   " Ptg. Color-Front"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":5442
         Picture         =   "BookPOChild01.frx":545E
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput37 
         Height          =   330
         Left            =   5040
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   8745
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":547A
         Caption         =   "BookPOChild01.frx":549A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":5506
         Keys            =   "BookPOChild01.frx":5524
         Spin            =   "BookPOChild01.frx":556E
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel29 
         Height          =   330
         Left            =   6600
         TabIndex        =   78
         Top             =   8745
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
         Caption         =   " Plate-F&&B"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild01.frx":5596
         Picture         =   "BookPOChild01.frx":55B2
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel42 
         Height          =   330
         Left            =   120
         TabIndex        =   100
         Top             =   6615
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
         Picture         =   "BookPOChild01.frx":55CE
         Picture         =   "BookPOChild01.frx":55EA
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel43 
         Height          =   330
         Left            =   6600
         TabIndex        =   84
         Top             =   6615
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
         Picture         =   "BookPOChild01.frx":5606
         Picture         =   "BookPOChild01.frx":5622
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput32 
         Height          =   330
         Left            =   8280
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   6615
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":563E
         Caption         =   "BookPOChild01.frx":565E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":56CA
         Keys            =   "BookPOChild01.frx":56E8
         Spin            =   "BookPOChild01.frx":5732
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput29 
         Height          =   330
         Left            =   1800
         TabIndex        =   51
         ToolTipText     =   "Plate"
         Top             =   6600
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":575A
         Caption         =   "BookPOChild01.frx":577A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":57E6
         Keys            =   "BookPOChild01.frx":5804
         Spin            =   "BookPOChild01.frx":584E
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput30 
         Height          =   330
         Left            =   4200
         TabIndex        =   52
         Top             =   6615
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":5876
         Caption         =   "BookPOChild01.frx":5896
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":5902
         Keys            =   "BookPOChild01.frx":5920
         Spin            =   "BookPOChild01.frx":596A
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
         ValueVT         =   160169989
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput31 
         Height          =   330
         Left            =   5040
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   6615
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild01.frx":5992
         Caption         =   "BookPOChild01.frx":59B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild01.frx":5A1E
         Keys            =   "BookPOChild01.frx":5A3C
         Spin            =   "BookPOChild01.frx":5A86
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
         ValueVT         =   1920925701
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel41 
         Height          =   330
         Left            =   3360
         TabIndex        =   83
         Top             =   6615
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
         Picture         =   "BookPOChild01.frx":5AAE
         Picture         =   "BookPOChild01.frx":5ACA
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   330
         Left            =   1800
         TabIndex        =   12
         Top             =   1610
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
         ParagraphAlign  =   3
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   9970
         Y1              =   2970
         Y2              =   2970
      End
      Begin MSForms.ComboBox Combo22 
         Height          =   330
         Left            =   9060
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   8445
         Width           =   795
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1402;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   9970
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   9970
         Y1              =   7555
         Y2              =   7555
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   9970
         Y1              =   560
         Y2              =   560
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9970
         Y1              =   7050
         Y2              =   7050
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   8280
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   8445
         Width           =   795
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1402;582"
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
         X2              =   9970
         Y1              =   4470
         Y2              =   4470
      End
      Begin MSForms.ComboBox Combo11 
         Height          =   330
         Left            =   9060
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   8745
         Width           =   795
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1402;582"
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
         Left            =   8280
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   8745
         Width           =   795
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1402;582"
         ListRows        =   3
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSComDlg.CommonDialog cdUpload 
      Left            =   10680
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   375
      Left            =   10200
      TabIndex        =   126
      Top             =   2400
      Width           =   1395
      BackColor       =   9164542
      Caption         =   "Update Master"
      Size            =   "2469;661"
      FontName        =   "Arial"
      FontEffects     =   1073741827
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "FrmBookPOChild01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchCode As String, VchType As String, PartyCode As String, RoundOffQty As Boolean, FinalQuantity As Long, rstBookPOChild06 As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset, rstSizeList As New ADODB.Recordset, rstPlateMakerList As New ADODB.Recordset, rstFetchRate As New ADODB.Recordset, rstElementList As New ADODB.Recordset, WithEvents rstBookPOChild06c As ADODB.Recordset
Attribute rstBookPOChild06c.VB_VarHelpID = -1
Dim ItemCode As String, FinishSizeCode As String, SizeCode As String, TitleSizeCode As String, PlateMakerCode As String, ElementCode As String, PaperCode As String
Dim PaperBalance As Double, SPU As Long, Wt As Double, inWidth As Double, GSM As Double
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
'    Me.Left = (MdiMainMenu.ScaleWidth - Me.Width) \ 2
    BusySystemIndicator True
    DisableCloseButton Me
    ItemCode = FrmBookPrintOrder.rstBookList.Fields("Code").Value
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text)
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text)
    Text7.Text = Trim(FrmBookPrintOrder.Text6.Text)
    Combo1.AddItem "Old", 0
    Combo1.AddItem "New", 1
    Combo1.AddItem "Revised", 2
    Combo11.AddItem "Old", 0
    Combo11.AddItem "New", 1
    Combo11.AddItem "Revised", 2
    Combo2.AddItem "Deepatch", 0
    Combo2.AddItem "PS", 1
    Combo2.AddItem "Wipeon", 2
    Combo2.AddItem "CTP", 3
    Combo22.AddItem "Deepatch", 0
    Combo22.AddItem "PS", 1
    Combo22.AddItem "Wipeon", 2
    Combo22.AddItem "CTP", 3
    Combo3.AddItem "F&B", 0
    Combo3.AddItem "W&T", 1
    ClearFields
    LoadMasterList
'    rstBookPOChild06.MoveFirst
'        If Val(CheckNull(rstBookPOChild06.Fields("ActualQuantity").Value)) = 0 Then
    MhRealInput1.Value = FrmBookPrintOrder.MhRealInput3.Value
    If Not CheckEmpty(FrmBookPrintOrder.imgFile, False) Then cmdUpload.Enabled = False
'        Else
'        LoadFields
'    End If
    Set rstBookPOChild06c = New ADODB.Recordset
    If DatabaseType = "MS SQL" Then
        cnDatabase.Execute "IF OBJECT_ID('tempdb.dbo.#T', 'U') IS NOT NULL  DROP TABLE #T"
        cnDatabase.Execute "SELECT * INTO #T FROM (SELECT Element,E.Name As ElementName,FinishSize,FS.Name As FinishSizeName,[Size],PS.Name As PrintSizeName,Imposition,FrontPrintingType,BackPrintingType,PlateType,PlateTypeBack,P.Pages,DuplexPrinting,[Titles/sheet1] As Ups,PaperConsumptionOther As PrintAmount FROM ((BookPOChild06 P INNER JOIN ElementMaster E ON P.[Element]=E.Code) INNER JOIN GeneralMaster FS ON P.FinishSize=FS.Code) INNER JOIN GeneralMaster PS ON P.[Size]=PS.Code WHERE P.Code='" & VchCode & "' UNION " & _
                                    "SELECT Element,E.Name As ElementName,FinishSize,FS.Name As FinishSizeName,[Size],PS.Name As PrintSizeName,Imposition,FrontPrintingType,BackPrintingType,PlateType,PlateTypeBack,P.Pages,DuplexPrinting,[Titles/sheet1] As Ups,0 As PrintAmount FROM ((BookChild06 P INNER JOIN ElementMaster E ON P.[Element]=E.Code) INNER JOIN GeneralMaster FS ON P.FinishSize=FS.Code) INNER JOIN GeneralMaster PS ON P.[Size]=PS.Code WHERE P.Code='" & ItemCode & "' AND P.[Type]='" & VchType & "' AND Element NOT IN (SELECT Element FROM BookPOChild06 WHERE Code='" & VchCode & "')) As Tbl ORDER BY ElementName,FinishSizeName,PrintSizeName"
        rstBookPOChild06c.Open "SELECT * FROM #T", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild06c.Open "SELECT Element,E.Name As ElementName,FinishSize,FS.Name As FinishSizeName,[Size],PS.Name As PrintSizeName,Imposition,FrontPrintingType,BackPrintingType,PlateType,PlateTypeBack,P.Pages,DuplexPrinting,[Titles/sheet1] As Ups,PrintAmount FROM ((BookPOChild06 P INNER JOIN ElementMaster E ON P.[Element]=E.Code) INNER JOIN GeneralMaster FS ON P.FinishSize=FS.Code) INNER JOIN GeneralMaster PS ON P.[Size]=PS.Code WHERE P.Code='" & VchCode & "' UNION " & _
                                         "SELECT Element,E.Name As ElementName,FinishSize,FS.Name As FinishSizeName,[Size],PS.Name As PrintSizeName,Imposition,FrontPrintingType,BackPrintingType,PlateType,PlateTypeBack,P.Pages,DuplexPrinting,[Titles/sheet1] As Ups,0 As PrintAmount FROM ((BookChild06 P INNER JOIN ElementMaster E ON P.[Element]=E.Code) INNER JOIN GeneralMaster FS ON P.FinishSize=FS.Code) INNER JOIN GeneralMaster PS ON P.[Size]=PS.Code WHERE P.Code='" & ItemCode & "' AND Element NOT IN (SELECT Element FROM BookPOChild06 WHERE Code='" & VchCode & "') ORDER BY ElementName,FinishSizeName,PrintSizeName", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
    Set DataGrid1.DataSource = rstBookPOChild06c
    rstBookPOChild06c.ActiveConnection = Nothing
    LockFields True
    SetButtons True
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Form_Activate()
    If Command1(0).Enabled Then If rstBookPOChild06c.RecordCount = 0 Then Command1_Click (0)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA And Command1(0).Enabled Then
        If Command1(0).Enabled Then Command1_Click (0)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Command1(1).Enabled Then
        If Command1(1).Enabled Then Command1_Click (1)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD And Command1(2).Enabled Then
        If Command1(2).Enabled Then Command1_Click (2)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Command1(3).Enabled Then
        If Command1(3).Enabled Then Command1_Click (3)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        If Command1(4).Enabled Then Command1_Click (4) Else cmdProceed_Click
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
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstFetchRate)
    Call CloseRecordset(rstElementList)
    Call CloseRecordset(rstPlateMakerList)
    Call CloseRecordset(rstBookPOChild06c)
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
    If MhDateInput3.ReadOnly Then Exit Sub
    If Format(GetDate(MhDateInput3.Text), "yyyymmdd") <= Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then DisplayError ("Target Date cann't be prior to Order Date"): MhDateInput3.SetFocus: Cancel = True
End Sub
Private Sub MhRealInput52_Validate(Cancel As Boolean)
    If MhRealInput52.Value = 0 And rstPaperList.Fields("Form").Value = "R" Then MsgBox "[Reel Cut off Size (mm.) ] can not be zero. You are using paper [ Reel ]": Cancel = True
End Sub
Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
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
    If CheckEmpty(Text14.Text, False) Then
        Cancel = True
    ElseIf CheckDuplicateElement() Then
        Call DisplayError("Duplicate Element"): Cancel = True
    End If
End Sub
Private Sub MhRealInput41_GotFocus()
    If CheckEmpty(ElementCode, False) Then Exit Sub
    Dim Pages As Integer, dblBookMark As Double
    With rstBookPOChild06c
        If .RecordCount > 0 Then
            dblBookMark = .Bookmark
            .MoveFirst
            .Find "[Element]='" & ElementCode & "'"
            If Not .EOF Then Pages = Val(.Fields("Pages").Value)
            If dblBookMark <> 0 Then .Bookmark = dblBookMark
        End If
    End With
    If Pages = 0 Then
        With rstElementList
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "[Code]='" & ElementCode & "'"
                If Not .EOF Then Pages = Val(.Fields("Pages").Value)
            End If
        End With
    End If
    If Pages > 0 Then
        If MhRealInput41.Value = 0 Then
            MhRealInput41.Value = Pages
        ElseIf Pages <> MhRealInput41.Value Then
            If MsgBox("Pages [" & Trim(MhRealInput41.Value) & "] are different from that in Master [" & Trim(Format(Pages, "#0")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput41.Value = Pages
        End If
    End If
End Sub
Private Sub MhRealInput41_Validate(Cancel As Boolean)   'Number of Pages
    If MhRealInput41.Value = 0 Then Cancel = True
End Sub
Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "11"
        FrmGeneralMaster.MasterCode = FinishSizeCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        FinishSizeCode = slCode: Text11.Text = slName
        If Not CheckEmpty(FinishSizeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        FinishSizeCode = "": Text11.Text = ""
    End If
End Sub
Private Sub Text11_Validate(Cancel As Boolean)
    If CheckEmpty(Text11.Text, False) Then
        Cancel = True
    Else
        If rstFetchRate.State = adStateOpen Then rstFetchRate.Close
        rstFetchRate.Open "SELECT DISTINCT 'Printing Size: '+M1.Name As Col0,C.Code+C.TitleSize As Code,M1.Name As TitleSizeName,TitleSize FROM FinishSizeChild C INNER JOIN GeneralMaster M1 ON C.TitleSize=M1.Code WHERE C.Code='" & FinishSizeCode & "' ORDER BY M1.Name", cnDatabase, adOpenKeyset, adLockReadOnly
        SelectionType = "S"
        Call LoadSelectionList(rstFetchRate, "List of Printing Sizes...", "Name", "")
        SearchOrder = 0
        Call DisplaySelectionList(Text12, TitleSizeCode)
        Call CloseForm(FrmSelectionList)
        If Not CheckEmpty(Trim(TitleSizeCode), False) Then
            rstFetchRate.MoveFirst
            rstFetchRate.Find "[Code]='" & TitleSizeCode & "'"
            If CheckEmpty(Text4.Text, False) Then
                Text4.Text = rstFetchRate.Fields("TitleSizeName").Value: SizeCode = rstFetchRate.Fields("TitleSize").Value
            ElseIf Text4.Text <> rstFetchRate.Fields("TitleSizeName").Value Then
                If MsgBox("Variation in Current (" & Text4.Text & ") and Master (" & rstFetchRate.Fields("TitleSizeName").Value & ") Printing Size ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then Text4.Text = rstFetchRate.Fields("TitleSizeName").Value: SizeCode = rstFetchRate.Fields("TitleSize").Value
            End If
        End If
    End If
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
        If Not CheckEmpty(SizeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        SizeCode = "": Text4.Text = ""
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)   'Size
    If CheckEmpty(Text4.Text, False) Then Cancel = True
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1.ListIndex = -1 Then Cancel = True
    If Combo1.ListIndex = 0 Then If InStr(1, "1_3", Trim(Combo2.ListIndex)) > 0 Then MhRealInput4.Value = 0
End Sub
Private Sub Combo11_Validate(Cancel As Boolean)
    If Combo11.ListIndex = -1 Then Cancel = True
    If Combo11.ListIndex = 0 Then If InStr(1, "1_3", Trim(Combo22.ListIndex)) > 0 Then MhRealInput38.Value = 0
End Sub
Private Sub Combo3_Validate(Cancel As Boolean)  'Imposition
    'Plates
    MhRealInput3.Value = IIf(Combo3.ListIndex = 0, MhRealInput19.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value)) * MhRealInput40.Value
    MhRealInput37.Value = IIf(Combo3.ListIndex = 0, MhRealInput20.Value, 0) * MhRealInput40.Value
    'Plate Rate
    MhRealInput38.Value = IIf(Combo3.ListIndex = 0, MhRealInput38.Value, 0)
    'Print Rate
    MhRealInput34.Value = IIf(Combo3.ListIndex = 0, MhRealInput34.Value, 0)
End Sub
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = PlateMakerCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        PlateMakerCode = slCode: Text9.Text = slName
        If Not CheckEmpty(PlateMakerCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        PlateMakerCode = "": Text9.Text = ""
    End If
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
    If CheckEmpty(Text9.Text, False) Then Cancel = True
End Sub
Private Sub MhRealInput19_Validate(Cancel As Boolean)   'Front Color
    MhRealInput3.Value = IIf(Combo3.ListIndex = 0, MhRealInput19.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, MhRealInput20.Value)) * MhRealInput40.Value
    If MhRealInput19.Value = 0 Then MhRealInput4.Value = 0: MhRealInput5.Value = 0
End Sub
Private Sub MhRealInput20_Validate(Cancel As Boolean)   'Back Color
    MhRealInput37.Value = IIf(Combo3.ListIndex = 0, MhRealInput20.Value, 0) * MhRealInput40.Value
    If MhRealInput20.Value = 0 Then MhRealInput38.Value = 0: MhRealInput34.Value = 0
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)  'Front Plate Type
    If Left(VchType, 1) = "O" Then Exit Sub
    If InStr(1, "1_3", Trim(Combo2.ListIndex)) > 0 Then 'PS/CTP Plate Details
        On Error Resume Next
        FrmPSPlateRegister.ItemCode = ItemCode
        FrmPSPlateRegister.ItemName = Trim(Text2.Text)
        FrmPSPlateRegister.ElementCode = ElementCode
        FrmPSPlateRegister.ElementName = Trim(Text14.Text)
        FrmPSPlateRegister.OrderCode = IIf(CheckEmpty(VchCode, False), "999999", VchCode)
        FrmPSPlateRegister.OrderDate = GetDate(MhDateInput1.Text)
        FrmPSPlateRegister.TblSuffix = "06"
        FrmPSPlateRegister.OrderType = VchType
        FrmPSPlateRegister.PlateType = "F"
        Load FrmPSPlateRegister
        If Err.Number <> 364 Then FrmPSPlateRegister.Show vbModal
        On Error GoTo 0
    End If
End Sub
Private Sub Combo22_Validate(Cancel As Boolean)  'Back Plate Type
    If Left(VchType, 1) = "O" Then Exit Sub
    If InStr(1, "1_3", Trim(Combo22.ListIndex)) > 0 Then 'PS/CTP Plate Details
        On Error Resume Next
        FrmPSPlateRegister.ItemCode = ItemCode
        FrmPSPlateRegister.ItemName = Trim(Text2.Text)
        FrmPSPlateRegister.ElementCode = ElementCode
        FrmPSPlateRegister.ElementName = Trim(Text14.Text)
        FrmPSPlateRegister.OrderCode = IIf(CheckEmpty(VchCode, False), "999999", VchCode)
        FrmPSPlateRegister.OrderDate = GetDate(MhDateInput1.Text)
        FrmPSPlateRegister.TblSuffix = "06"
        FrmPSPlateRegister.OrderType = VchType
        FrmPSPlateRegister.PlateType = "B"
        Load FrmPSPlateRegister
        If Err.Number <> 364 Then FrmPSPlateRegister.Show vbModal
        On Error GoTo 0
    End If
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)    'Actual Quantity
    If MhRealInput1.Value = 0 Then Cancel = True: Exit Sub
    Call CalculateConsumption
End Sub
Private Sub MhRealInput6_GotFocus()
    CalculateTotalForms
End Sub
Private Sub MhRealInput6_Validate(Cancel As Boolean)
    Call CalculatePrintAmount
End Sub
Private Sub MhRealInput15_GotFocus()
    If CheckEmpty(FinishSizeCode, False) Or CheckEmpty(SizeCode, False) Or MhRealInput41.Value = 0 Then Exit Sub
    Dim FL As Double, FR As Double, PL As Double, PR As Double, Ups01 As Integer, Ups02 As Integer, Ups As Integer, MaxUps As Integer, BalPgs As Integer, Sets As Integer
    PL = Val(Left(Text4.Text, InStr(1, Text4.Text, "X") - 1)) + 1: PR = Val(Mid(Text4.Text, InStr(1, Text4.Text, "X") + 1, 5)) + 1
    FL = Val(Left(Text11.Text, InStr(1, Text11.Text, "X") - 1)): FR = Val(Mid(Text11.Text, InStr(1, Text11.Text, "X") + 1, 5))
    Ups01 = Int(IIf(PL > PR, PL, PR) / IIf(FL > FR, FL, FR)) * Int(IIf(PL < PR, PL, PR) / IIf(FL < FR, FL, FR)): Ups02 = Int(IIf(PL < PR, PL, PR) / IIf(FL > FR, FL, FR)) * Int(IIf(PL > PR, PL, PR) / IIf(FL < FR, FL, FR))
'    MaxUps = IIf(MhRealInput41.Value = 2, 2, IIf(Ups01 > Ups02, Ups01, Ups02))
    MaxUps = IIf(Ups01 > Ups02, Ups01, Ups02)
    If MaxUps > 0 And MhRealInput41.Value > (2 * MaxUps) Then Ups = 1 'Ups calculation
    If Ups = 0 Then Ups = Int((2 * MaxUps) / MhRealInput41.Value)
    Sets = Int(MhRealInput41.Value / MaxUps * IIf(Combo3.ListIndex = 0, 0.5, 1))
    If Sets = 0 Then Sets = 1
    If MhRealInput40.Value = 0 Then
        MhRealInput40.Value = Sets
    ElseIf Sets <> MhRealInput40.Value Then
        If MsgBox("Variation in Calculated [" & Trim(Sets) & "] and Existing [" & Trim(MhRealInput40.Value) & "] Sets ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput40.Value = Sets
    End If
    BalPgs = MhRealInput41.Value - (MhRealInput40.Value * MaxUps * IIf(Combo3.ListIndex = 0, 2, 1)) 'Bal Pages
    If BalPgs > 0 Then DisplayError ("Please note that [" & BalPgs & "] pages are pending for processing"): MhRealInput41.Value = MhRealInput41.Value - BalPgs
    
    
'    Ups = (2 * MaxUps) / MhRealInput41.Value
    If Ups > 0 Then
        If MhRealInput15.Value = 0 Then
            MhRealInput15.Value = Ups
        ElseIf Ups <> MhRealInput15.Value Then
            If MsgBox("Variation in Calculated [" & Trim(Ups) & "] and Existing [" & Trim(MhRealInput15.Value) & "] Pages/Printing Form ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput15.Value = Ups
        End If
    End If
End Sub
Private Sub MhRealInput15_Validate(Cancel As Boolean)   'Ups (Printing)
    Call CalculateConsumption
End Sub
Private Sub MhRealInput40_Validate(Cancel As Boolean)
    CalculateConsumption
End Sub
Private Sub MhRealInput5_GotFocus()
    Call GetPrinterRates("P", "F")
End Sub
Private Sub MhRealInput5_Validate(Cancel As Boolean)    'Front Print Rate
    CalculatePrintAmount
End Sub
Private Sub MhRealInput34_GotFocus()
    Call GetPrinterRates("P", "B")
End Sub
Private Sub MhRealInput34_Validate(Cancel As Boolean)    'Back Print Rate
    CalculatePrintAmount
End Sub
Private Sub MhRealInput4_GotFocus()
    Call GetPrinterRates("L", "F")
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)    'Front Plate Rate
    CalculatePlateAmount
End Sub
Private Sub MhRealInput38_GotFocus()
    Call GetPrinterRates("L", "B")
End Sub
Private Sub MhRealInput38_Validate(Cancel As Boolean)    'Back Plate Rate
    CalculatePlateAmount
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
    If CheckEmpty(Text1.Text, False) Then
        Cancel = True
    Else
        With rstPaperList
            .MoveFirst
            .Find "[Code]='" & PaperCode & "'"
            Text1.Text = .Fields("Col0").Value: SPU = Val(.Fields("SPU").Value): Wt = Val(.Fields("Wt").Value): inWidth = Val(.Fields("inWidth").Value): GSM = Val(.Fields("GSM").Value)
            If .Fields("Form").Value = "S" Then MhRealInput52.Value = 0
            CalculateConsumption
            If CheckEmpty(SizeCode, False) Then Exit Sub
            Dim FL As Double, FR As Double, PL As Double, PW As Double
            FL = Val(Left(Text4.Text, InStr(1, Text4.Text, "X") - 1)): FR = Val(Mid(Text4.Text, InStr(1, Text4.Text, "X") + 1, 5)) 'Printing Size Left & Right
            PL = IIf(.Fields("Form").Value = "R", MhRealInput52.Value / 25.4, Val(.Fields("inLength").Value)): PW = Val(.Fields("inWidth").Value) 'Paper Area Length & Width
            If Abs(FL - PL) <= 1 Then PL = FL
            If Abs(FR - PL) <= 1 Then PL = FR
            If Abs(FL - PW) <= 1 Then PW = FL
            If Abs(FR - PW) <= 1 Then PW = FR
            Call CalcUps(PL * PW, FL * FR)
        End With
    End If
End Sub
Private Sub MhRealInput12_GotFocus()
    Dim FL As Double, FR As Double, PL As Double, PR As Double, Ups01 As Integer, Ups02 As Integer, Ups03 As Integer, Ups As Integer
    If CheckEmpty(PaperCode, False) Or CheckEmpty(SizeCode, False) Then Exit Sub
    rstPaperList.MoveFirst
    rstPaperList.Find "[Code]='" & PaperCode & "'"
    
    PL = Val(Left(Text4.Text, InStr(1, Text4.Text, "X") - 1)): PR = Val(Mid(Text4.Text, InStr(1, Text4.Text, "X") + 1, 5))
           
    FL = Val(rstPaperList.Fields("inWidth").Value): FR = IIf(rstPaperList.Fields("Form").Value = "R", MhRealInput52.Value / 25.4, Val(rstPaperList.Fields("inLength").Value)) 'Val(rstPaperList.Fields("inLength").Value)
    
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
Private Sub MhRealInput35_GotFocus()
    Call GetPrinterRates("W", "B")
End Sub
Private Sub MhRealInput35_Validate(Cancel As Boolean)   'Wastage Percentage - Back
    CalculateConsumption
End Sub
Private Sub MhRealInput23_GotFocus()
    Call GetPrinterRates("M", "F")
End Sub
Private Sub MhRealInput23_Validate(Cancel As Boolean)   'Wastage Min - Front
    CalculateConsumption
End Sub
Private Sub MhRealInput36_GotFocus()
    Call GetPrinterRates("M", "B")
End Sub
Private Sub MhRealInput36_Validate(Cancel As Boolean)   'Wastage Min - Back
    CalculateConsumption
End Sub
Private Sub MhRealInput25_Validate(Cancel As Boolean)   'Paper Rate
    MhRealInput26.Value = MhRealInput25.Value * MhRealInput39.Value
    CalculateTotalAmount
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean)    'Adjustment
    CalculateTotalAmount
End Sub
Private Sub MhRealInput27_Validate(Cancel As Boolean)   'Plate Adjustment
    CalculateTotalAmount
End Sub
Private Sub MhRealInput29_Validate(Cancel As Boolean)   'Paper Adjustment
    CalculateTotalAmount
End Sub
Private Sub MhRealInput18_Validate(Cancel As Boolean)   'GST%
    CalculateTotalAmount
End Sub
Private Sub MhRealInput21_Validate(Cancel As Boolean)   'PGST%
    CalculateTotalAmount
End Sub
Private Sub MhRealInput30_Validate(Cancel As Boolean)   'RGST%
    CalculateTotalAmount
End Sub
Private Sub cmdProceed_Click()
    Dim elProcessed As Boolean
    If Not Command1(4).Enabled Then
        With rstBookPOChild06c
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    elProcessed = True
                    If rstBookPOChild06c.Fields("PrintAmount").Value = 0 Then elProcessed = False
                    If Not elProcessed Then If MsgBox("[" & .Fields("ElementName").Value & "] Element has not been processed ! Process?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Process !") = vbYes Then Command1_Click (1): Exit Sub
                    .MoveNext
                Loop
            End If
        End With
        If Not CheckEmpty(Text8.Text, False) Or Not CheckEmpty(Text10.Text, False) Then
            With rstBookPOChild06
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF()
                        .Fields("BillNo").Value = Text8.Text
                        If Not IsDate(MhDateInput2.Text) Then .Fields("BillDate").Value = Null Else .Fields("BillDate").Value = GetDate(MhDateInput2.Text)
                        .Fields("PBillNo").Value = Text10.Text
                        If Not IsDate(MhDateInput4.Text) Then .Fields("PBillDate").Value = Null Else .Fields("PBillDate").Value = GetDate(MhDateInput4.Text)
                        .Fields("PaidAmount").Value = MhRealInput16.Value
                        .Fields("PPaidAmount").Value = MhRealInput24.Value
                        .Update
                        .MoveNext
                    Loop
                End If
            End With
        End If
        FrmBookPrintOrder.Command5.Enabled = False: Call CloseForm(Me)
    Else
        Command1_Click (3)
    End If
End Sub
Private Sub Command1_Click(Index As Integer)
    With rstBookPOChild06c
        Select Case Index
            Case 0
                ClearFields
                Me.Tag = "A"
                SetButtons False
                LockFields False
                MhDateInput1.Text = Format(Date, "dd-MM-yyyy"): MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
                FinishSizeCode = FrmBookPrintOrder.rstBookList.Fields("FinishSize").Value
                If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
                rstSizeList.Find "[Code] = '" & FinishSizeCode & "'"
                If Not rstSizeList.EOF Then Text11.Text = rstSizeList.Fields("Col0").Value
                PlateMakerCode = PartyCode
                If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
                rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
                If Not rstPlateMakerList.EOF Then Text9.Text = rstPlateMakerList.Fields("Col0").Value
                MhRealInput3.Value = MhRealInput19.Value: MhRealInput37.Value = MhRealInput20.Value
                MhRealInput1.Value = FinalQuantity
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
                    MhRealInput41.Value = Val(.Fields("Pages").Value)
                    FinishSizeCode = .Fields("FinishSize").Value
                    If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
                    rstSizeList.Find "[Code] = '" & FinishSizeCode & "'"
                    If Not rstSizeList.EOF Then Text11.Text = rstSizeList.Fields("Col0").Value
                    SizeCode = .Fields("Size").Value
                    If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
                    rstSizeList.Find "[Code] = '" & SizeCode & "'"
                    If Not rstSizeList.EOF Then Text4.Text = rstSizeList.Fields("Col0").Value
                    Combo3.ListIndex = IIf(.Fields("Imposition").Value = "F", 0, 1)
                    PlateMakerCode = PartyCode
                    If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
                    rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
                    If Not rstPlateMakerList.EOF Then Text9.Text = rstPlateMakerList.Fields("Col0").Value
                    MhRealInput19.Value = Val(.Fields("FrontPrintingType").Value)
                    MhRealInput20.Value = Val(.Fields("BackPrintingType").Value)
                    Combo2.ListIndex = Val(.Fields("PlateType").Value) - 1
                    Combo22.ListIndex = Val(.Fields("PlateTypeBack").Value) - 1
                    MhRealInput15.Value = Val(.Fields("Ups").Value)
                    MhRealInput3.Value = IIf(Combo3.ListIndex = 0, MhRealInput19.Value, IIf(MhRealInput19.Value > MhRealInput20.Value, MhRealInput19.Value, 0)) * MhRealInput40.Value
                    MhRealInput37.Value = IIf(Combo3.ListIndex = 0, MhRealInput20.Value, IIf(MhRealInput20.Value > MhRealInput19.Value, MhRealInput20.Value, 0)) * MhRealInput40.Value
                    LoadFields
                    If MhRealInput1.Value = 0 Then MhRealInput1.Value = FinalQuantity
                    SetButtons False
                    LockFields False
                    DataGrid1.Enabled = False
                    MhDateInput1.SetFocus
                End If
            Case 2
                If .RecordCount > 0 Then
                    If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
                        If rstBookPOChild06.RecordCount > 0 Then
                            rstBookPOChild06.MoveFirst
                            rstBookPOChild06.Find "[Element]='" & .Fields("Element").Value & "'"
                            If Not rstBookPOChild06.EOF Then rstBookPOChild06.Delete: rstBookPOChild06.MoveNext
                        End If
                        Me.Tag = "D"
                        .Delete: .MoveNext
                        Me.Tag = ""
                    End If
                End If
            Case 3
                If CheckMandatoryFields Then Exit Sub
                If Left(VchType, 1) <> "O" Then
                    Dim Stock As Double, VchDate As Date
                    VchDate = MhDateInput1.Value
                    Stock = CalculatePaperBalance(IIf(chkPaper.Value, PartyCode, "000000"), PaperCode, VchCode, "PO", VchDate): Stock = Fix(Val(Stock)) * Val(SPU) + Round(Val(Stock) - Fix(Val(Stock)), 3) * 1000
                    PaperBalance = Stock - (CLng(Int(MhRealInput13.Value) * SPU) + (MhRealInput13.Value - Int(MhRealInput13.Value)) * 1000)
                    If PaperBalance < 0 Then
                        If UserLevel <= 2 Then
                            If MsgBox("Stock (" & Trim(Format(CLng(Fix(0 - Abs(PaperBalance) / Val(SPU))) + ((0 - Abs(PaperBalance) Mod Val(SPU)) / 1000), "0.000")) & ") (" & Trim(Format((PaperBalance / SPU) * Wt, "0.000")) & " Kg) of the Paper - " & Trim(Text1.Text) & " )" & vbCrLf & " is going negative ! Would you like to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbNo Then Exit Sub
                        Else
                            Call DisplayError("Cann't Save !! Stock (" & Trim(Format(CLng(Fix(0 - Abs(PaperBalance) / Val(SPU))) + ((0 - Abs(PaperBalance) Mod Val(SPU)) / 1000), "0.000")) & ") (" & Trim(Format((PaperBalance / SPU) * Wt, "0.000")) & " Kg) of the Paper - " & Trim(Text1.Text) & " )" & " is going negative "): AbortPO = True: Exit Sub
                        End If
                    End If
                End If
                If Me.Tag = "A" Then Call AddRecord(rstBookPOChild06c)
                .Fields("Element").Value = ElementCode
                .Fields("ElementName").Value = Text14.Text
                .Fields("FinishSize").Value = FinishSizeCode
                .Fields("FinishSizeName").Value = Text11.Text
                .Fields("Size").Value = SizeCode
                .Fields("PrintSizeName").Value = Text4.Text
                .Fields("Imposition").Value = IIf(Combo3.ListIndex = 0, "F", "W")
                .Fields("FrontPrintingType").Value = MhRealInput19.Value
                .Fields("BackPrintingType").Value = MhRealInput20.Value
                .Fields("PlateType").Value = Trim(Str(Combo2.ListIndex + 1))
                .Fields("PlateTypeBack").Value = Trim(Str(Combo22.ListIndex + 1))
                .Fields("Pages").Value = MhRealInput41.Value
                .Fields("DuplexPrinting").Value = IIf(MhRealInput19.Value = 0 Or MhRealInput20.Value = 0, 0, 1)
                .Fields("Ups").Value = MhRealInput15.Value
                .Fields("PrintAmount").Value = MhRealInput13.Value
                .Update
                If InStr(1, "A_E1", Me.Tag) > 0 Then Call AddRecord(rstBookPOChild06)
                SaveFields
                rstBookPOChild06.Update
                SetButtons True
                LockFields True
                DataGrid1.Enabled = True
                DataGrid1.SetFocus
                If Left(Me.Tag, 1) = "E" Then
                    Me.Tag = ""
                    rstBookPOChild06c.MoveNext
                    If rstBookPOChild06c.EOF Then
                        rstBookPOChild06c.MoveLast
                    Else
                        Command1_Click (1)
                    End If
                Else
                    Me.Tag = ""
                End If
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
Private Sub rstBookPOChild06c_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If CheckEmpty(Me.Tag, False) Then
        On Error Resume Next
        ClearFields
        If Not (rstBookPOChild06c.EOF Or rstBookPOChild06c.BOF) Then LoadFields
    End If
End Sub
Private Sub ClearFields()
    MhDateInput1.Text = "  -  -    "
    MhDateInput3.Text = "  -  -    "
    Text14.Text = "": ElementCode = ""
    MhRealInput41.Value = 0
    Text11.Text = "": FinishSizeCode = ""
    Text4.Text = "": SizeCode = "'"
    Combo1.ListIndex = 0
    Combo11.ListIndex = 0
    Combo3.ListIndex = 0
    Text3.Text = ""
    Text9.Text = "": PlateMakerCode = ""
    MhRealInput19.Value = 4
    MhRealInput20.Value = 0
    Combo2.ListIndex = 3
    Combo22.ListIndex = 3
    MhRealInput1.Value = 0
    MhRealInput6.Value = 0
    MhRealInput15.Value = 0
    MhRealInput40.Value = 1
    MhRealInput3.Value = 0
    MhRealInput37.Value = 0
    MhRealInput5.Value = 0
    MhRealInput34.Value = 0
    MhRealInput8.Value = 0
    MhRealInput4.Value = 0
    MhRealInput38.Value = 0
    MhRealInput7.Value = 0
    chkPaper.Value = 1
    Text1.Text = "": PaperCode = "": SPU = 0: Wt = 0: inWidth = 0: GSM = 0
    MhRealInput52.Value = 0
    MhRealInput12.Value = 0
    MhRealInput11.Value = 0
    MhRealInput35.Value = 0
    MhRealInput23.Value = 0
    MhRealInput36.Value = 0
    MhRealInput33.Value = 0
    MhRealInput13.Value = 0
    MhRealInput39.Value = 0
    MhRealInput25.Value = 0
    MhRealInput26.Value = 0
    Text6.Text = ""
    Text8.Text = ""
    MhDateInput2.Text = "  -  -    "
    Text10.Text = ""
    MhDateInput4.Text = "  -  -    "
    MhRealInput9.Value = 0
    MhRealInput27.Value = 0
    MhRealInput29.Value = 0
    MhRealInput18.Value = 0
    MhRealInput17.Value = 0
    MhRealInput21.Value = 0
    MhRealInput30.Value = 0
    MhRealInput22.Value = 0
    MhRealInput31.Value = 0
    MhRealInput10.Value = 0
    MhRealInput28.Value = 0
    MhRealInput32.Value = 0
    MhRealInput16.Value = 0
    MhRealInput24.Value = 0
    TxtAdNar.Text = ""
    Text12.Text = ""
    MhRealInput42.Value = 0
End Sub
Private Sub LoadFields()
    With rstBookPOChild06
        If Me.Tag = "E" Then Me.Tag = "E1"
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        .Find "[Element]='" & rstBookPOChild06c.Fields("Element").Value & "'"
        If Not .EOF Then
            If Me.Tag = "E1" Then Me.Tag = "E2"
            MhDateInput1.Text = Format(.Fields("OrderDate").Value, "dd-MM-yyyy")
            MhDateInput3.Text = Format(.Fields("TargetDate").Value, "dd-MM-yyyy")
            ElementCode = .Fields("Element").Value
            If rstElementList.RecordCount > 0 Then rstElementList.MoveFirst
            rstElementList.Find "[Code] = '" & ElementCode & "'"
            If Not rstElementList.EOF Then Text14.Text = Trim(rstElementList.Fields("Col0").Value)
            MhRealInput41.Value = Val(.Fields("Pages").Value)
            FinishSizeCode = .Fields("FinishSize").Value
            If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
            rstSizeList.Find "[Code] = '" & FinishSizeCode & "'"
            If Not rstSizeList.EOF Then Text11.Text = rstSizeList.Fields("Col0").Value
            SizeCode = .Fields("Size").Value
            If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
            rstSizeList.Find "[Code] = '" & SizeCode & "'"
            If Not rstSizeList.EOF Then Text4.Text = rstSizeList.Fields("Col0").Value
            Combo1.ListIndex = IIf(.Fields("Processing").Value = "O", 0, IIf(.Fields("Processing").Value = "N", 1, 2))  'O:Old N:New R:Revised
            Combo11.ListIndex = IIf(.Fields("ProcessingBack").Value = "O", 0, IIf(.Fields("ProcessingBack").Value = "N", 1, 2))  'O:Old N:New R:Revised
            Combo3.ListIndex = IIf(.Fields("Imposition").Value = "F", 0, 1) 'F:Front-Back W:Work-Turn
            Text3.Text = .Fields("Ref").Value
            PlateMakerCode = .Fields("PlateMaker").Value
            If rstPlateMakerList.RecordCount > 0 Then rstPlateMakerList.MoveFirst
            rstPlateMakerList.Find "[Code] = '" & PlateMakerCode & "'"
            If Not rstPlateMakerList.EOF Then Text9.Text = Trim(rstPlateMakerList.Fields("Col0").Value)
            MhRealInput19.Value = Val(.Fields("FrontPrintingType").Value)
            MhRealInput20.Value = Val(.Fields("BackPrintingType").Value)
            Combo2.ListIndex = Val(.Fields("PlateType").Value) - 1
            Combo22.ListIndex = Val(.Fields("PlateTypeBack").Value) - 1
            MhRealInput1.Value = Val(.Fields("ActualQuantity").Value)
            MhRealInput6.Value = Val(.Fields("BillingQuantity").Value)
            MhRealInput15.Value = Val(.Fields("Titles/Sheet1").Value)
            MhRealInput40.Value = Val(.Fields("Sets").Value)
            MhRealInput3.Value = Val(.Fields("TotalPlates").Value)
            MhRealInput37.Value = Val(.Fields("TotalPlatesBack").Value)
            MhRealInput5.Value = Val(.Fields("PrintRate").Value)
            MhRealInput34.Value = Val(.Fields("PrintRateBack").Value)
            MhRealInput8.Value = Val(.Fields("PrintAmount").Value)
            MhRealInput4.Value = Val(.Fields("PlateRate").Value)
            MhRealInput38.Value = Val(.Fields("PlateRateBack").Value)
            MhRealInput7.Value = Val(.Fields("PlateAmount").Value)
            chkPaper.Value = IIf(.Fields("PaperByParty").Value, 1, 0)
            PaperCode = .Fields("Paper").Value
            If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
            rstPaperList.Find "[Code] = '" & PaperCode & "'"
            If Not rstPaperList.EOF Then Text1.Text = rstPaperList.Fields("Col0").Value: SPU = Val(rstPaperList.Fields("SPU").Value): Wt = Val(rstPaperList.Fields("Wt").Value): inWidth = Val(rstPaperList.Fields("inWidth").Value): GSM = Val(rstPaperList.Fields("GSM").Value)
            MhRealInput52.Value = Val(.Fields("CutOffSize").Value)
            MhRealInput12.Value = Val(.Fields("Titles/Sheet2").Value)
            MhRealInput11.Value = Val(.Fields("PaperWastage%").Value)
            MhRealInput35.Value = Val(.Fields("PaperWastage%Back").Value)
            MhRealInput23.Value = Val(.Fields("PaperWastageMin").Value)
            MhRealInput36.Value = Val(.Fields("PaperWastageMinBack").Value)
            MhRealInput33.Value = Val(.Fields("PaperWastageFinal").Value)
            MhRealInput13.Value = Val(.Fields("PaperConsumptionOther").Value)
            MhRealInput39.Value = Val(.Fields("PaperConsumptionKg").Value)
            MhRealInput25.Value = Val(.Fields("PaperRate").Value)
            MhRealInput26.Value = Val(.Fields("PaperAmount").Value)
            Text6.Text = .Fields("Remarks").Value
            Text8.Text = .Fields("BillNo").Value
            If Not IsNull(.Fields("BillDate").Value) Then MhDateInput2.Text = Format(.Fields("BillDate").Value, "dd-MM-yyyy")
            Text10.Text = .Fields("PBillNo").Value
            If Not IsNull(.Fields("PBillDate").Value) Then MhDateInput4.Text = Format(.Fields("PBillDate").Value, "dd-MM-yyyy")
            MhRealInput9.Value = Val(.Fields("Adjustment").Value)
            MhRealInput27.Value = Val(.Fields("PAdjustment").Value)
            MhRealInput29.Value = Val(.Fields("RAdjustment").Value)
            MhRealInput18.Value = Val(.Fields("VAT%").Value)
            MhRealInput17.Value = Val(.Fields("VAT").Value)
            MhRealInput21.Value = Val(.Fields("PVAT%").Value)
            MhRealInput30.Value = Val(.Fields("RVAT%").Value)
            MhRealInput22.Value = Val(.Fields("PVAT").Value)
            MhRealInput31.Value = Val(.Fields("RVAT").Value)
            MhRealInput10.Value = Val(.Fields("BillAmount").Value)
            MhRealInput28.Value = Val(.Fields("PBillAmount").Value)
            MhRealInput32.Value = Val(.Fields("RBillAmount").Value)
            MhRealInput16.Value = Val(.Fields("PaidAmount").Value)
            MhRealInput24.Value = Val(.Fields("PPaidAmount").Value)
            TxtAdNar.Text = .Fields("AdjustmentRemarks").Value
            MhRealInput42.Value = MhRealInput1.Value / MhRealInput15.Text: MhRealInput42.Value = MhRealInput42.Value + IIf(MhRealInput42.Value - Int(MhRealInput42.Value) > 0, 1, 0)
        End If
    End With
End Sub
Private Sub SaveFields()
    With rstBookPOChild06
        .Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
        .Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
        .Fields("Element").Value = ElementCode
        .Fields("Pages").Value = Val(MhRealInput41.Value)
        .Fields("FinishSize").Value = FinishSizeCode
        .Fields("Size").Value = SizeCode
        .Fields("Processing").Value = IIf(Combo1.ListIndex = 0, "O", IIf(Combo1.ListIndex = 1, "N", "R"))
        .Fields("ProcessingBack").Value = IIf(Combo11.ListIndex = 0, "O", IIf(Combo11.ListIndex = 1, "N", "R"))
        .Fields("Imposition").Value = IIf(Combo3.ListIndex = 0, "F", "W")
        .Fields("Ref").Value = Text3.Text
        .Fields("PlateMaker").Value = PlateMakerCode
        .Fields("FrontPrintingType").Value = MhRealInput19.Value
        .Fields("BackPrintingType").Value = MhRealInput20.Value
        .Fields("PlateType").Value = Trim(Str(Combo2.ListIndex + 1))
        .Fields("PlateTypeBack").Value = Trim(Str(Combo22.ListIndex + 1))
        .Fields("ActualQuantity").Value = MhRealInput1.Value
        .Fields("BillingQuantity").Value = MhRealInput6.Value
        .Fields("Totalforms").Value = MhRealInput6.Value
        .Fields("Titles/Sheet1").Value = MhRealInput15.Value
        .Fields("Sets").Value = MhRealInput40.Value
        .Fields("TotalPlates").Value = MhRealInput3.Value
        .Fields("TotalPlatesBack").Value = MhRealInput37.Value
        .Fields("PrintRate").Value = MhRealInput5.Value
        .Fields("PrintRateBack").Value = MhRealInput34.Value
        .Fields("PrintAmount").Value = MhRealInput8.Value
        .Fields("PlateRate").Value = MhRealInput4.Value
        .Fields("PlateRateBack").Value = MhRealInput38.Value
        .Fields("PlateAmount").Value = MhRealInput7.Value
        .Fields("PaperByParty").Value = chkPaper.Value
        .Fields("Paper").Value = PaperCode
        .Fields("CutOffSize").Value = MhRealInput52.Value
        .Fields("Titles/Sheet2").Value = MhRealInput12.Value
        .Fields("PaperWastage%").Value = MhRealInput11.Value
        .Fields("PaperWastage%Back").Value = MhRealInput35.Value
        .Fields("PaperWastageMin").Value = MhRealInput23.Value
        .Fields("PaperWastageMinBack").Value = MhRealInput36.Value
        .Fields("PaperWastageFinal").Value = MhRealInput33.Value
        .Fields("PaperConsumptionOther").Value = MhRealInput13.Value
        .Fields("PaperConsumptionSheets").Value = CLng(Int(MhRealInput13.Value) * SPU) + ((MhRealInput13.Value - Int(MhRealInput13.Value)) * 1000)
        .Fields("PaperConsumptionKg").Value = MhRealInput39.Value
        .Fields("PaperRate").Value = MhRealInput25.Value
        .Fields("PaperAmount").Value = MhRealInput26.Value
        .Fields("Remarks").Value = Text6.Text
        .Fields("BillNo").Value = Text8.Text
        If Not IsDate(MhDateInput2.Text) Then .Fields("BillDate").Value = Null Else .Fields("BillDate").Value = GetDate(MhDateInput2.Text)
        .Fields("PBillNo").Value = Text10.Text
        If Not IsDate(MhDateInput4.Text) Then .Fields("PBillDate").Value = Null Else .Fields("PBillDate").Value = GetDate(MhDateInput4.Text)
        .Fields("Adjustment").Value = MhRealInput9.Value
        .Fields("PAdjustment").Value = MhRealInput27.Value
        .Fields("RAdjustment").Value = MhRealInput29.Value
        .Fields("VAT%").Value = MhRealInput18.Value
        .Fields("VAT").Value = MhRealInput17.Value
        .Fields("PVAT%").Value = MhRealInput21.Value
        .Fields("PVAT").Value = MhRealInput22.Value
        .Fields("RVAT%").Value = MhRealInput30.Value
        .Fields("RVAT").Value = MhRealInput31.Value
        .Fields("BillAmount").Value = MhRealInput10.Value
        .Fields("PBillAmount").Value = MhRealInput28.Value
        .Fields("RBillAmount").Value = MhRealInput32.Value
        .Fields("PaidAmount").Value = MhRealInput16.Value
        .Fields("PPaidAmount").Value = MhRealInput24.Value
        .Fields("AdjustmentRemarks").Value = IIf(MhRealInput9.Value <> 0 Or MhRealInput27.Value <> 0 Or MhRealInput29.Value <> 0, TxtAdNar.Text, "")
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    'If Text1.Text <> "" Then MhRealInput52.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo1.ListIndex < 0 Then Combo1.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo11.ListIndex < 0 Then Combo11.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo2.ListIndex < 0 Then Combo2.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo22.ListIndex < 0 Then Combo22.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo3.ListIndex < 0 Then Combo3.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput16.Value <> 0 Then If MhRealInput16.Value <> MhRealInput10.Value + MhRealInput32.Value Then MhRealInput9.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput24.Value <> 0 Then If MhRealInput24.Value <> MhRealInput28.Value Then MhRealInput27.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput9.Value <> 0 Or MhRealInput27.Value <> 0 Or MhRealInput29.Value <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
End Function
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
            If MhRealInput3.Value > 0 Then
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
            If MhRealInput37.Value > 0 Then
                If Combo11.ListIndex > 0 Then
                    If backPlateRate > 0 Then
                        If MhRealInput38.Value = 0 Then
                            MhRealInput38.Value = backPlateRate
                        ElseIf MhRealInput38.Value <> backPlateRate Then
                            If MsgBox("Back Plate Rate [" & Trim(MhRealInput38.Value) & "] is different from that in Master [" & Trim(Format(backPlateRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput38.Value = backPlateRate
                        End If
                    End If
                Else
                    If InStr(1, "1_3", Trim(Combo22.ListIndex)) > 0 Then MhRealInput38.Value = 0
                End If
            Else
                MhRealInput38.Value = 0
            End If
        End If
    ElseIf RateType = "P" Then
        If Position = "F" Then
            If MhRealInput3.Value > 0 Then
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
            If MhRealInput37.Value > 0 Then
                If backPrintRate > 0 Then
                    If MhRealInput34.Value = 0 Then
                        MhRealInput34.Value = backPrintRate
                    ElseIf MhRealInput34.Value <> backPrintRate Then
                        If MsgBox("Back Print Rate [" & Trim(MhRealInput34.Value) & "] is different from that in Master [" & Trim(Format(backPrintRate, "#0.00")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput34.Value = backPrintRate
                    End If
                End If
            Else
                MhRealInput34.Value = 0
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
Private Sub CalculateTotalForms()
    Dim TotalForms As Double
    If MhRealInput15.Value > 0 Then 'Ups/Plate
        TotalForms = MhRealInput1.Value / MhRealInput15.Text  'Total Forms
        TotalForms = TotalForms + IIf(TotalForms - Int(TotalForms) > 0, 1, 0)
        If TotalForms > 0 Then
            MhRealInput42.Value = TotalForms * MhRealInput40
            If RoundOffQty Then
                If TotalForms < 1000 Then TotalForms = TotalForms '1000
                'TotalForms = IIf(Int(TotalForms) = 0, 1000, Int(TotalForms / 1000) * 1000) + IIf(TotalForms Mod 1000 <= IIf(TotalForms <= 20000, 299, 599), 0, 1000)
            End If
            If MhRealInput6.Value = 0 Then
                MhRealInput6.Value = TotalForms  'TotalForms * MhRealInput40
            ElseIf MhRealInput6.Value <> TotalForms Then
                If MsgBox("Variation in Calculated [" & Trim(TotalForms) & "] and Existing [" & Trim(MhRealInput6.Value) & "] Impressions/Set ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput6.Value = TotalForms
            End If
        End If
    End If
End Sub
Private Sub CalculatePlateAmount()
    MhRealInput7.Value = MhRealInput3.Value * MhRealInput4.Value + MhRealInput37.Value * MhRealInput38.Value
    CalculateTotalAmount
End Sub
Private Sub CalculatePrintAmount()
    Dim TotalForms As Long
    TotalForms = MhRealInput6.Value * IIf(Combo3.ListIndex = 0, 1, 2)
    MhRealInput8.Value = IIf(TotalForms < 1000, TotalForms, TotalForms) * MhRealInput5.Value * MhRealInput40 + IIf(TotalForms < 1000, TotalForms, TotalForms) * MhRealInput34.Value * MhRealInput40
    CalculateTotalAmount
End Sub
Private Sub CalculateTotalAmount()
    MhRealInput22.Value = (MhRealInput7.Value + MhRealInput27.Value) * MhRealInput21.Value / 100    'GST-Plate
    MhRealInput17.Value = (MhRealInput8.Value + MhRealInput9.Value) * MhRealInput18.Value / 100     'GST-Ptg
    MhRealInput31.Value = (MhRealInput26.Value + MhRealInput29.Value) * MhRealInput30.Value / 100   'GST-Paper
    MhRealInput10.Value = Round(MhRealInput8.Value + MhRealInput9.Value + MhRealInput17.Value, 0)
    MhRealInput28.Value = Round(MhRealInput7.Value + MhRealInput22.Value + MhRealInput27.Value, 0)
    MhRealInput32.Value = Round(MhRealInput26.Value + MhRealInput29.Value + MhRealInput31.Value, 0)
End Sub
Private Sub CalculateConsumption()
    If SPU = 0 Or MhRealInput12.Value = 0 Then Exit Sub
    Dim C As Long, W As Long, q As Long
    If MhRealInput12.Value > 0 Then
        q = MhRealInput1.Value / MhRealInput15.Value    'Qty (Sheets)
        W = (q * (MhRealInput11.Value + MhRealInput35.Value)) / 100  'Wastage (in Sheets)
        If W < (MhRealInput23.Value + MhRealInput36.Value) Then W = (MhRealInput23.Value + MhRealInput36.Value) 'Comparison with Minimum Wastage
        C = q + W   'Consumption With Wastage (in Sheets)
        C = C / MhRealInput12.Value
        MhRealInput39.Value = IIf(MhRealInput52.Value > 0, Round(((MhRealInput52.Value / 25.4) * inWidth * GSM) / 3100, 3), Wt) * (C / SPU)
        MhRealInput33.Value = CLng(Int((W * MhRealInput40.Value) / SPU)) + (((W * MhRealInput40.Value) Mod SPU) / 1000) 'Min Wastage Final
        MhRealInput13.Value = CLng(Int((C * MhRealInput40.Value) / SPU)) + (((C * MhRealInput40.Value) Mod SPU) / 1000)
    End If
End Sub
Private Sub SetButtons(ByVal bVal As Boolean)
    Command1(0).Enabled = bVal
    Command1(1).Enabled = bVal
    Command1(2).Enabled = bVal
    Command1(3).Enabled = Not bVal
    Command1(4).Enabled = Not bVal
End Sub
Private Sub LoadMasterList(Optional ByVal LoadSelected As Boolean)
    If rstSizeList.State = adStateOpen Then rstSizeList.Close
    rstSizeList.Open "SELECT Name As Col0, Code From GeneralMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstSizeList.ActiveConnection = Nothing
    If rstPaperList.State = adStateOpen Then rstPaperList.Close
    If LoadSelected Then
        rstPaperList.Open "SELECT * FROM (SELECT LTRIM(P.Name)+' (UOM : '+LTRIM(C.Name)+'='+LTRIM(C.Value1)+')' As Col0,FORMAT(dbo.ufnGetPaperStock('" & IIf(chkPaper.Value, PartyCode, "000000") & "',P.Code,'PO','" & VchCode & "','" & GetDate(MhDateInput1.Text) & "'),'#0.000') As Col1,C.Name As UOM,GSM,inWidth,inLength,P.Code,C.Value1 As SPU,[Form],[Weight/Unit] As Wt,LTRIM(Q.Name) As Quality,Grade FROM (PaperMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code) INNER JOIN GeneralMaster Q ON P.Quality=Q.Code) As Tbl WHERE CONVERT(DECIMAL(12,3),Col1)<>0 ORDER BY Col0", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstPaperList.Open "SELECT LTRIM(P.Name)+' (UOM : '+LTRIM(C.Name)+'='+LTRIM(C.Value1)+')' As Col0,FORMAT(0,'#0.000') As Col1,C.Name As UOM,GSM,inWidth,inLength,P.Code,C.Value1 As SPU,[Form],[Weight/Unit] As Wt,LTRIM(Q.Name) As Quality,Grade FROM (PaperMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code) INNER JOIN GeneralMaster Q ON P.Quality=Q.Code ORDER BY Col0", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstPlateMakerList.State = adStateOpen Then rstPlateMakerList.Close
    rstPlateMakerList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstPlateMakerList.ActiveConnection = Nothing
    If rstElementList.State = adStateOpen Then rstElementList.Close
    rstElementList.Open "SELECT Name As Col0,Pages,Code FROM ElementMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
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
        End If
    Next
    If Not bVal Then Text5.Locked = True: Text2.Locked = True: Text7.Locked = True: Text9.Locked = True: Text14.Locked = True: Text11.Locked = True: Text4.Locked = True: Text1.Locked = True: MhRealInput42.ReadOnly = True: MhRealInput3.ReadOnly = True: MhRealInput37.ReadOnly = True: MhRealInput7.ReadOnly = True: MhRealInput8.ReadOnly = True: MhRealInput33.ReadOnly = True: MhRealInput13.ReadOnly = True: MhRealInput26.ReadOnly = True: MhRealInput39.ReadOnly = True: MhRealInput17.ReadOnly = True: MhRealInput10.ReadOnly = True: MhRealInput22.ReadOnly = True: MhRealInput28.ReadOnly = True: MhRealInput31.ReadOnly = True: MhRealInput32.ReadOnly = True
End Sub
Private Function CheckDuplicateElement() As Boolean
    Dim dblBookMark As Double
    With rstBookPOChild06c
        If .RecordCount = 0 Then Exit Function
        If Not (.EOF Or .BOF) Then dblBookMark = .Bookmark
        .MoveFirst
        Do While Not .EOF
            If Me.Tag = "A" Then
                If .Fields("ElementName").Value = Trim(Text14.Text) Then CheckDuplicateElement = True: Exit Do
            ElseIf Left(Me.Tag, 1) = "E" Then
                If .Fields("ElementName").Value = Trim(Text14.Text) And .Bookmark <> dblBookMark Then CheckDuplicateElement = True: Exit Do
            End If
            .MoveNext
        Loop
        If dblBookMark <> 0 Then .Bookmark = dblBookMark Else .MoveLast
    End With
End Function
Private Sub cmdUpload_Click() 'Load Pic
    On Error Resume Next
    With cdUpload
        .CancelError = True
        .DialogTitle = "Open Image"
        .Filter = "All Picture Files|*.jpg;*.jpeg;*.bmp;*.gif;*.png"
        .ShowOpen
        If Err.Number = 0 Then FrmBookPrintOrder.imgFile = .FileName: cmdUpload.Enabled = False 'Ok Selected
    End With
End Sub
Private Sub cmdDelete_Click() 'Delete Pic
    If Not CheckEmpty(FrmBookPrintOrder.imgFile, False) Then
        If MsgBox("Are you sure to delete the Picture?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then FrmBookPrintOrder.imgFile = "": cmdUpload.Enabled = True
    End If
End Sub
Private Sub cmdView_Click() 'View Pic
    If CheckEmpty(FrmBookPrintOrder.imgFile, False) Then DisplayError ("No image exists") Else Call ShellExecute(Me.hwnd, "open", FrmBookPrintOrder.imgFile, "", "", 1)
End Sub
Private Sub CommandButton1_Click()
    On Error GoTo ErrHandler
    cnDatabase.Execute "DELETE FROM BookChild06 WHERE Type='" & VchType & "' AND Code='" & ItemCode & "'"
    cnDatabase.Execute "INSERT INTO BookChild06 SELECT TOP 1 Book As Code,Element,Pages,[FinishSize],[Size],Imposition,FrontPrintingType,BackPrintingType,PlateType,PlateTypeBack,[Titles/sheet1],P.Type FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type='" & VchType & "' AND P.Book='" & ItemCode & "' ORDER BY P.Code DESC"
ErrHandler:
End Sub

