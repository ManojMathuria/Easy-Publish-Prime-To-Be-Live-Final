VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmAccountChild05 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Rate Detail"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
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
   ScaleHeight     =   3900
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   7853
      Picture         =   "AccountChild05.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   7853
      Picture         =   "AccountChild05.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   3705
      Left            =   120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   105
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   6535
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
      Picture         =   "AccountChild05.frx":0204
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
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   35
         Top             =   100
         Width           =   4815
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
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   0
         Top             =   425
         Width           =   4815
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   120
         TabIndex        =   37
         Top             =   420
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Size Group"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":0220
         Picture         =   "AccountChild05.frx":023C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   105
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Picture         =   "AccountChild05.frx":0258
         Picture         =   "AccountChild05.frx":0274
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   39
         Top             =   1680
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " PS Plate Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":0290
         Picture         =   "AccountChild05.frx":02AC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   120
         TabIndex        =   40
         Top             =   1995
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Depatch Plate Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":02C8
         Picture         =   "AccountChild05.frx":02E4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   120
         TabIndex        =   41
         Top             =   2940
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Paper Wastage %  Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":0300
         Picture         =   "AccountChild05.frx":031C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   42
         Top             =   735
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Picture         =   "AccountChild05.frx":0338
         Picture         =   "AccountChild05.frx":0354
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   2160
         TabIndex        =   43
         Top             =   735
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
         Caption         =   " 1 Color "
         Alignment       =   1
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":0370
         Picture         =   "AccountChild05.frx":038C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   4560
         TabIndex        =   44
         Top             =   735
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
         Caption         =   " 4 Color "
         Alignment       =   1
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":03A8
         Picture         =   "AccountChild05.frx":03C4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   3360
         TabIndex        =   45
         Top             =   735
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
         Caption         =   " 2 Color "
         Alignment       =   1
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":03E0
         Picture         =   "AccountChild05.frx":03FC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   120
         TabIndex        =   46
         Top             =   1365
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Picture         =   "AccountChild05.frx":0418
         Picture         =   "AccountChild05.frx":0434
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   5760
         TabIndex        =   47
         Top             =   735
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
         Caption         =   " Spl. Color "
         Alignment       =   1
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":0450
         Picture         =   "AccountChild05.frx":046C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   48
         Top             =   1050
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Range"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":0488
         Picture         =   "AccountChild05.frx":04A4
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   2160
         TabIndex        =   1
         Top             =   1050
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":04C0
         Caption         =   "AccountChild05.frx":04E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":054C
         Keys            =   "AccountChild05.frx":056A
         Spin            =   "AccountChild05.frx":05B4
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
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   3360
         TabIndex        =   2
         Top             =   1050
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":05DC
         Caption         =   "AccountChild05.frx":05FC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0668
         Keys            =   "AccountChild05.frx":0686
         Spin            =   "AccountChild05.frx":06D0
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
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   4560
         TabIndex        =   3
         Top             =   1050
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":06F8
         Caption         =   "AccountChild05.frx":0718
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0784
         Keys            =   "AccountChild05.frx":07A2
         Spin            =   "AccountChild05.frx":07EC
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
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   5760
         TabIndex        =   4
         Top             =   1050
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0814
         Caption         =   "AccountChild05.frx":0834
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":08A0
         Keys            =   "AccountChild05.frx":08BE
         Spin            =   "AccountChild05.frx":0908
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
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   2160
         TabIndex        =   5
         Top             =   1365
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0930
         Caption         =   "AccountChild05.frx":0950
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":09BC
         Keys            =   "AccountChild05.frx":09DA
         Spin            =   "AccountChild05.frx":0A24
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   3360
         TabIndex        =   6
         Top             =   1365
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0A4C
         Caption         =   "AccountChild05.frx":0A6C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0AD8
         Keys            =   "AccountChild05.frx":0AF6
         Spin            =   "AccountChild05.frx":0B40
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   4560
         TabIndex        =   7
         Top             =   1365
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0B68
         Caption         =   "AccountChild05.frx":0B88
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0BF4
         Keys            =   "AccountChild05.frx":0C12
         Spin            =   "AccountChild05.frx":0C5C
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
         Height          =   330
         Left            =   5760
         TabIndex        =   8
         Top             =   1365
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0C84
         Caption         =   "AccountChild05.frx":0CA4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0D10
         Keys            =   "AccountChild05.frx":0D2E
         Spin            =   "AccountChild05.frx":0D78
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   2160
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0DA0
         Caption         =   "AccountChild05.frx":0DC0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0E2C
         Keys            =   "AccountChild05.frx":0E4A
         Spin            =   "AccountChild05.frx":0E94
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
         Height          =   330
         Left            =   3360
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0EBC
         Caption         =   "AccountChild05.frx":0EDC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0F48
         Keys            =   "AccountChild05.frx":0F66
         Spin            =   "AccountChild05.frx":0FB0
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   330
         Left            =   4560
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0FD8
         Caption         =   "AccountChild05.frx":0FF8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1064
         Keys            =   "AccountChild05.frx":1082
         Spin            =   "AccountChild05.frx":10CC
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
         Height          =   330
         Left            =   5760
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":10F4
         Caption         =   "AccountChild05.frx":1114
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1180
         Keys            =   "AccountChild05.frx":119E
         Spin            =   "AccountChild05.frx":11E8
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
         Height          =   330
         Left            =   2160
         TabIndex        =   13
         Top             =   1995
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1210
         Caption         =   "AccountChild05.frx":1230
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":129C
         Keys            =   "AccountChild05.frx":12BA
         Spin            =   "AccountChild05.frx":1304
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   3360
         TabIndex        =   14
         Top             =   1995
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":132C
         Caption         =   "AccountChild05.frx":134C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":13B8
         Keys            =   "AccountChild05.frx":13D6
         Spin            =   "AccountChild05.frx":1420
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
         Height          =   330
         Left            =   4560
         TabIndex        =   15
         Top             =   1995
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1448
         Caption         =   "AccountChild05.frx":1468
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":14D4
         Keys            =   "AccountChild05.frx":14F2
         Spin            =   "AccountChild05.frx":153C
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
         Height          =   330
         Left            =   5760
         TabIndex        =   16
         Top             =   1995
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1564
         Caption         =   "AccountChild05.frx":1584
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":15F0
         Keys            =   "AccountChild05.frx":160E
         Spin            =   "AccountChild05.frx":1658
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
         Height          =   330
         Left            =   3360
         TabIndex        =   18
         Top             =   2315
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1680
         Caption         =   "AccountChild05.frx":16A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":170C
         Keys            =   "AccountChild05.frx":172A
         Spin            =   "AccountChild05.frx":1774
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
         Height          =   330
         Left            =   4560
         TabIndex        =   19
         Top             =   2315
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":179C
         Caption         =   "AccountChild05.frx":17BC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1828
         Keys            =   "AccountChild05.frx":1846
         Spin            =   "AccountChild05.frx":1890
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
         Height          =   330
         Left            =   5760
         TabIndex        =   20
         Top             =   2315
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":18B8
         Caption         =   "AccountChild05.frx":18D8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1944
         Keys            =   "AccountChild05.frx":1962
         Spin            =   "AccountChild05.frx":19AC
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
         Height          =   330
         Left            =   2160
         TabIndex        =   25
         Top             =   2945
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":19D4
         Caption         =   "AccountChild05.frx":19F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1A60
         Keys            =   "AccountChild05.frx":1A7E
         Spin            =   "AccountChild05.frx":1AC8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
         Height          =   330
         Left            =   3360
         TabIndex        =   26
         Top             =   2945
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1AF0
         Caption         =   "AccountChild05.frx":1B10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1B7C
         Keys            =   "AccountChild05.frx":1B9A
         Spin            =   "AccountChild05.frx":1BE4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput23 
         Height          =   330
         Left            =   4560
         TabIndex        =   27
         Top             =   2945
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1C0C
         Caption         =   "AccountChild05.frx":1C2C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1C98
         Keys            =   "AccountChild05.frx":1CB6
         Spin            =   "AccountChild05.frx":1D00
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
         Height          =   330
         Left            =   5760
         TabIndex        =   28
         Top             =   2945
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1D28
         Caption         =   "AccountChild05.frx":1D48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1DB4
         Keys            =   "AccountChild05.frx":1DD2
         Spin            =   "AccountChild05.frx":1E1C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   120
         TabIndex        =   49
         Top             =   2310
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Wipeon Plate Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":1E44
         Picture         =   "AccountChild05.frx":1E60
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   120
         TabIndex        =   50
         Top             =   2625
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " CTP Plate Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":1E7C
         Picture         =   "AccountChild05.frx":1E98
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
         Height          =   330
         Left            =   2160
         TabIndex        =   21
         Top             =   2630
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1EB4
         Caption         =   "AccountChild05.frx":1ED4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1F40
         Keys            =   "AccountChild05.frx":1F5E
         Spin            =   "AccountChild05.frx":1FA8
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
         Height          =   330
         Left            =   3360
         TabIndex        =   22
         Top             =   2630
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1FD0
         Caption         =   "AccountChild05.frx":1FF0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":205C
         Keys            =   "AccountChild05.frx":207A
         Spin            =   "AccountChild05.frx":20C4
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
         Height          =   330
         Left            =   4560
         TabIndex        =   23
         Top             =   2630
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":20EC
         Caption         =   "AccountChild05.frx":210C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":2178
         Keys            =   "AccountChild05.frx":2196
         Spin            =   "AccountChild05.frx":21E0
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput28 
         Height          =   330
         Left            =   5760
         TabIndex        =   24
         Top             =   2630
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":2208
         Caption         =   "AccountChild05.frx":2228
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":2294
         Keys            =   "AccountChild05.frx":22B2
         Spin            =   "AccountChild05.frx":22FC
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   120
         TabIndex        =   51
         Top             =   3255
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Paper Wastage Min"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "AccountChild05.frx":2324
         Picture         =   "AccountChild05.frx":2340
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput29 
         Height          =   330
         Left            =   2160
         TabIndex        =   29
         Top             =   3260
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":235C
         Caption         =   "AccountChild05.frx":237C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":23E8
         Keys            =   "AccountChild05.frx":2406
         Spin            =   "AccountChild05.frx":2450
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
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput30 
         Height          =   330
         Left            =   3360
         TabIndex        =   30
         Top             =   3255
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":2478
         Caption         =   "AccountChild05.frx":2498
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":2504
         Keys            =   "AccountChild05.frx":2522
         Spin            =   "AccountChild05.frx":256C
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
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput31 
         Height          =   330
         Left            =   4560
         TabIndex        =   31
         Top             =   3260
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":2594
         Caption         =   "AccountChild05.frx":25B4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":2620
         Keys            =   "AccountChild05.frx":263E
         Spin            =   "AccountChild05.frx":2688
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
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput32 
         Height          =   330
         Left            =   5760
         TabIndex        =   32
         Top             =   3260
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":26B0
         Caption         =   "AccountChild05.frx":26D0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":273C
         Keys            =   "AccountChild05.frx":275A
         Spin            =   "AccountChild05.frx":27A4
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
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   2160
         TabIndex        =   17
         Top             =   2310
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":27CC
         Caption         =   "AccountChild05.frx":27EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":2858
         Keys            =   "AccountChild05.frx":2876
         Spin            =   "AccountChild05.frx":28C0
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
         ValueVT         =   88866821
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   930
      Index           =   2
      Left            =   7320
      TabIndex        =   52
      Top             =   960
      Width           =   1440
      _Version        =   65536
      _ExtentX        =   2540
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
      Caption         =   "Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D->Delete  Ctrl+S->Save"
      AutoSize        =   -1  'True
      FillColor       =   8421504
      TextColor       =   16777215
      Picture         =   "AccountChild05.frx":28E8
      Multiline       =   -1  'True
      GlobalMem       =   -1  'True
      Picture         =   "AccountChild05.frx":2904
   End
End
Attribute VB_Name = "FrmAccountChild05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rstAccountChild As New ADODB.Recordset
Public rstSizeList As New ADODB.Recordset
Public AccountName As String
Dim SizeCode As String
Private Sub Form_Load()
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
    Text2.Text = Trim(AccountName)
    ClearFields
    LoadFields
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
    If UnloadMode = 0 Then
        Call CloseForm(FrmAccountChild05)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set rstAccountChild = Nothing
    Set rstSizeList = Nothing
End Sub
Private Sub ClearFields()
    Text3.Text = ""
    MhRealInput1.Text = "0"
    MhRealInput2.Text = "0"
    MhRealInput3.Text = "0"
    MhRealInput4.Text = "0.00"
    MhRealInput5.Text = "0.00"
    MhRealInput6.Text = "0.00"
    MhRealInput7.Text = "0.00"
    MhRealInput8.Text = "0.00"
    MhRealInput9.Text = "0.00"
    MhRealInput10.Text = "0.00"
    MhRealInput11.Text = "0.00"
    MhRealInput12.Text = "0.00"
    MhRealInput13.Text = "0"
    MhRealInput14.Text = "0.00"
    MhRealInput15.Text = "0.00"
    MhRealInput16.Text = "0.00"
    MhRealInput17.Text = "0.00"
    MhRealInput18.Text = "0.00"
    MhRealInput19.Text = "0.00"
    MhRealInput20.Text = "0.00"
    MhRealInput21.Text = "0.00"
    MhRealInput22.Text = "0.00"
    MhRealInput23.Text = "0.00"
    MhRealInput24.Text = "0.00"
    MhRealInput25.Text = "0.00"
    MhRealInput26.Text = "0.00"
    MhRealInput27.Text = "0.00"
    MhRealInput28.Text = "0.00"
    MhRealInput29.Text = "0"
    MhRealInput30.Text = "0"
    MhRealInput31.Text = "0"
    MhRealInput32.Text = "0"
End Sub
Private Sub LoadFields()
    If rstAccountChild.RecordCount = 0 Then Exit Sub
    If Not CheckEmpty(rstAccountChild.Fields("Size").Value, False) Then
        Text3.Text = rstAccountChild.Fields("SizeName").Value
        MhRealInput1.Text = Format(Val(rstAccountChild.Fields("Range1").Value), "0")
        MhRealInput2.Text = Format(Val(rstAccountChild.Fields("Range2").Value), "0")
        MhRealInput3.Text = Format(Val(rstAccountChild.Fields("Range4").Value), "0")
        MhRealInput13.Text = Format(Val(rstAccountChild.Fields("Range6").Value), "0")
        MhRealInput4.Text = Format(Val(rstAccountChild.Fields("PrintRate1").Value), "0.00")
        MhRealInput5.Text = Format(Val(rstAccountChild.Fields("PrintRate2").Value), "0.00")
        MhRealInput6.Text = Format(Val(rstAccountChild.Fields("PrintRate4").Value), "0.00")
        MhRealInput14.Text = Format(Val(rstAccountChild.Fields("PrintRate6").Value), "0.00")
        MhRealInput7.Text = Format(Val(rstAccountChild.Fields("PSPlateRate1").Value), "0.00")
        MhRealInput8.Text = Format(Val(rstAccountChild.Fields("PSPlateRate2").Value), "0.00")
        MhRealInput9.Text = Format(Val(rstAccountChild.Fields("PSPlateRate4").Value), "0.00")
        MhRealInput15.Text = Format(Val(rstAccountChild.Fields("PSPlateRate6").Value), "0.00")
        MhRealInput10.Text = Format(Val(rstAccountChild.Fields("DeepatchPlateRate1").Value), "0.00")
        MhRealInput11.Text = Format(Val(rstAccountChild.Fields("DeepatchPlateRate2").Value), "0.00")
        MhRealInput12.Text = Format(Val(rstAccountChild.Fields("DeepatchPlateRate4").Value), "0.00")
        MhRealInput16.Text = Format(Val(rstAccountChild.Fields("DeepatchPlateRate6").Value), "0.00")
        MhRealInput17.Text = Format(Val(rstAccountChild.Fields("WipeonPlateRate1").Value), "0.00")
        MhRealInput18.Text = Format(Val(rstAccountChild.Fields("WipeonPlateRate2").Value), "0.00")
        MhRealInput19.Text = Format(Val(rstAccountChild.Fields("WipeonPlateRate4").Value), "0.00")
        MhRealInput20.Text = Format(Val(rstAccountChild.Fields("WipeonPlateRate6").Value), "0.00")
        MhRealInput25.Text = Format(Val(rstAccountChild.Fields("CTPPlateRate1").Value), "0.00")
        MhRealInput26.Text = Format(Val(rstAccountChild.Fields("CTPPlateRate2").Value), "0.00")
        MhRealInput27.Text = Format(Val(rstAccountChild.Fields("CTPPlateRate4").Value), "0.00")
        MhRealInput28.Text = Format(Val(rstAccountChild.Fields("CTPPlateRate6").Value), "0.00")
        MhRealInput21.Text = Format(Val(rstAccountChild.Fields("PaperWastageRate1").Value), "0.00")
        MhRealInput22.Text = Format(Val(rstAccountChild.Fields("PaperWastageRate2").Value), "0.00")
        MhRealInput23.Text = Format(Val(rstAccountChild.Fields("PaperWastageRate4").Value), "0.00")
        MhRealInput24.Text = Format(Val(rstAccountChild.Fields("PaperWastageRate6").Value), "0.00")
        MhRealInput29.Text = Format(Val(rstAccountChild.Fields("PaperWastageMin1").Value), "0")
        MhRealInput30.Text = Format(Val(rstAccountChild.Fields("PaperWastageMin2").Value), "0")
        MhRealInput31.Text = Format(Val(rstAccountChild.Fields("PaperWastageMin4").Value), "0")
        MhRealInput32.Text = Format(Val(rstAccountChild.Fields("PaperWastageMin6").Value), "0")
    End If
End Sub
Private Sub SaveFields()
    rstAccountChild.Fields("Size").Value = SizeCode
    rstAccountChild.Fields("SizeName").Value = Trim(Text3.Text)
    rstAccountChild.Fields("Range1").Value = Val(MhRealInput1.Text)
    rstAccountChild.Fields("Range2").Value = Val(MhRealInput2.Text)
    rstAccountChild.Fields("Range4").Value = Val(MhRealInput3.Text)
    rstAccountChild.Fields("Range6").Value = Val(MhRealInput13.Text)
    rstAccountChild.Fields("PrintRate1").Value = Val(MhRealInput4.Text)
    rstAccountChild.Fields("PrintRate2").Value = Val(MhRealInput5.Text)
    rstAccountChild.Fields("PrintRate4").Value = Val(MhRealInput6.Text)
    rstAccountChild.Fields("PrintRate6").Value = Val(MhRealInput14.Text)
    rstAccountChild.Fields("PSPlateRate1").Value = Val(MhRealInput7.Text)
    rstAccountChild.Fields("PSPlateRate2").Value = Val(MhRealInput8.Text)
    rstAccountChild.Fields("PSPlateRate4").Value = Val(MhRealInput9.Text)
    rstAccountChild.Fields("PSPlateRate6").Value = Val(MhRealInput15.Text)
    rstAccountChild.Fields("DeepatchPlateRate1").Value = Val(MhRealInput10.Text)
    rstAccountChild.Fields("DeepatchPlateRate2").Value = Val(MhRealInput11.Text)
    rstAccountChild.Fields("DeepatchPlateRate4").Value = Val(MhRealInput12.Text)
    rstAccountChild.Fields("DeepatchPlateRate6").Value = Val(MhRealInput16.Text)
    rstAccountChild.Fields("WipeonPlateRate1").Value = Val(MhRealInput17.Text)
    rstAccountChild.Fields("WipeonPlateRate2").Value = Val(MhRealInput18.Text)
    rstAccountChild.Fields("WipeonPlateRate4").Value = Val(MhRealInput19.Text)
    rstAccountChild.Fields("WipeonPlateRate6").Value = Val(MhRealInput20.Text)
    rstAccountChild.Fields("CTPPlateRate1").Value = Val(MhRealInput25.Text)
    rstAccountChild.Fields("CTPPlateRate2").Value = Val(MhRealInput26.Text)
    rstAccountChild.Fields("CTPPlateRate4").Value = Val(MhRealInput27.Text)
    rstAccountChild.Fields("CTPPlateRate6").Value = Val(MhRealInput28.Text)
    rstAccountChild.Fields("PaperWastageRate1").Value = Val(MhRealInput21.Text)
    rstAccountChild.Fields("PaperWastageRate2").Value = Val(MhRealInput22.Text)
    rstAccountChild.Fields("PaperWastageRate4").Value = Val(MhRealInput23.Text)
    rstAccountChild.Fields("PaperWastageRate6").Value = Val(MhRealInput24.Text)
    rstAccountChild.Fields("PaperWastageMin1").Value = Val(MhRealInput29.Text)
    rstAccountChild.Fields("PaperWastageMin2").Value = Val(MhRealInput30.Text)
    rstAccountChild.Fields("PaperWastageMin4").Value = Val(MhRealInput31.Text)
    rstAccountChild.Fields("PaperWastageMin6").Value = Val(MhRealInput32.Text)
End Sub
Private Sub Text3_Change()
    If Text3.Text = " " Then Text3.Text = "?": Sendkeys "{TAB}"
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text3.Text)
    If rstSizeList.RecordCount = 0 Then
       DisplayError ("No Record in Size Master")
       Cancel = True
       Exit Sub
    Else
       rstSizeList.MoveFirst
    End If
    rstSizeList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstSizeList.EOF Then
        SelectionType = "S"
        SizeCode = ""
        Call LoadSelectionList(rstSizeList, "List of Sizes...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, SizeCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(SizeCode) <> "" Then Sendkeys "{TAB}"
        Cancel = True
        Exit Sub
    ElseIf (rstAccountChild.Fields("SizeName").Value <> Trim(Text3.Text)) Or (CheckEmpty(rstAccountChild.Fields("SizeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            FocusSelect Me.ActiveControl
            Cancel = True
            Exit Sub
        End If
    End If
    SizeCode = rstSizeList.Fields("Code").Value
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    If (Val(CheckNull(rstAccountChild.Fields("Range1").Value)) <> Val(MhRealInput1.Text)) Or (CheckEmpty(rstAccountChild.Fields("SizeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Me.SetFocus
            Cancel = True
        End If
    End If
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    If (Val(CheckNull(rstAccountChild.Fields("Range2").Value)) <> Val(MhRealInput2.Text)) Or (CheckEmpty(rstAccountChild.Fields("SizeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Me.SetFocus
            Cancel = True
        End If
    End If
End Sub
Private Sub MhRealInput3_Validate(Cancel As Boolean)
    If (Val(CheckNull(rstAccountChild.Fields("Range4").Value)) <> Val(MhRealInput3.Text)) Or (CheckEmpty(rstAccountChild.Fields("SizeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Me.SetFocus
            Cancel = True
        End If
    End If
End Sub
Private Sub MhRealInput13_Validate(Cancel As Boolean)
    If (Val(CheckNull(rstAccountChild.Fields("Range6").Value)) <> Val(MhRealInput13.Text)) Or (CheckEmpty(rstAccountChild.Fields("SizeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Me.SetFocus
            Cancel = True
        End If
    End If
End Sub
Private Sub cmdProceed_Click()
    Dim Control As Object
    If CheckMandatoryFields Then Exit Sub
    SaveFields
    Me.Tag = "T"
    For Each Control In Me
        If Left(Control.Name, 6) = "MhReal" Then If Val(Control.Text) <> 0 Then Me.Tag = "F"
    Next
    If Me.Tag = "T" Then rstAccountChild.Fields("Size").Value = ""
    rstAccountChild.Update
    Call CloseForm(FrmAccountChild05)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(FrmAccountChild05)
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstSizeList, SizeCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf Val(MhRealInput1.Text) < 0 Or Val(MhRealInput1.Text) > 999999 Then
        MhRealInput1.SetFocus
        FocusSelect Me.ActiveControl
        CheckMandatoryFields = True
    ElseIf Val(MhRealInput2.Text) < 0 Or Val(MhRealInput2.Text) > 999999 Then
        MhRealInput2.SetFocus
        FocusSelect Me.ActiveControl
        CheckMandatoryFields = True
    ElseIf Val(MhRealInput3.Text) < 0 Or Val(MhRealInput3.Text) > 999999 Then
        MhRealInput3.SetFocus
        FocusSelect Me.ActiveControl
        CheckMandatoryFields = True
    ElseIf Val(MhRealInput13.Text) < 0 Or Val(MhRealInput13.Text) > 999999 Then
        MhRealInput13.SetFocus
        FocusSelect Me.ActiveControl
        CheckMandatoryFields = True
    End If
End Function
Private Function CheckDuplicateEntry() As Boolean
    Dim dblBookMark As Double
    If rstAccountChild.RecordCount = 0 Then Exit Function
    If Not (rstAccountChild.EOF Or rstAccountChild.BOF) Then dblBookMark = rstAccountChild.Bookmark
    rstAccountChild.MoveFirst
    Do While Not rstAccountChild.EOF
          If rstAccountChild.Fields("SizeName").Value = Trim(Text3.Text) And Val(CheckNull(rstAccountChild.Fields("Range1").Value)) = Val(MhRealInput1.Text) And Val(CheckNull(rstAccountChild.Fields("Range2").Value)) = Val(MhRealInput2.Text) And Val(CheckNull(rstAccountChild.Fields("Range4").Value)) = Val(MhRealInput3.Text) And Val(CheckNull(rstAccountChild.Fields("Range6").Value)) = Val(MhRealInput13.Text) Then CheckDuplicateEntry = True: Exit Do
          rstAccountChild.MoveNext
    Loop
    If dblBookMark <> 0 Then rstAccountChild.Bookmark = dblBookMark Else rstAccountChild.MoveLast
End Function
