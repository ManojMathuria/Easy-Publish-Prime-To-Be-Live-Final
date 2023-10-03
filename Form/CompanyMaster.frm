VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmCompanyMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Master"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Height          =   375
      Left            =   8270
      Picture         =   "CompanyMaster.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Save"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   8270
      Picture         =   "CompanyMaster.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6415
      Left            =   105
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   105
      Width           =   8085
      _Version        =   65536
      _ExtentX        =   14261
      _ExtentY        =   11333
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
      Picture         =   "CompanyMaster.frx":0204
      Begin VB.TextBox Text25 
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
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   54
         Top             =   1920
         Width           =   1725
      End
      Begin VB.TextBox Text24 
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
         Left            =   5400
         MaxLength       =   60
         TabIndex        =   8
         Top             =   2315
         Width           =   2565
      End
      Begin VB.TextBox Text23 
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
         IMEMode         =   3  'DISABLE
         Left            =   6240
         MaxLength       =   60
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   5990
         Width           =   1725
      End
      Begin VB.TextBox Text22 
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
         MaxLength       =   60
         TabIndex        =   24
         Top             =   5990
         Width           =   3015
      End
      Begin VB.TextBox Text21 
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
         Left            =   6240
         MaxLength       =   60
         TabIndex        =   23
         Top             =   5665
         Width           =   1725
      End
      Begin VB.TextBox Text17 
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
         MaxLength       =   60
         TabIndex        =   22
         Top             =   5670
         Width           =   3015
      End
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
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   15
         Top             =   3885
         Width           =   1725
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4230
         TabIndex        =   47
         Top             =   5200
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1950
         TabIndex        =   46
         Top             =   5200
         Value           =   -1  'True
         Width           =   225
      End
      Begin VB.TextBox Text15 
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
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   21
         Top             =   5145
         Width           =   1725
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
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   5145
         Width           =   765
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   5140
         Width           =   765
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
         MaxLength       =   40
         TabIndex        =   12
         Top             =   3575
         Width           =   6285
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
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   9
         Top             =   2630
         Width           =   6285
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   40
         Top             =   2315
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
         Caption         =   " Mobile"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0220
         Picture         =   "CompanyMaster.frx":023C
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
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   7
         Top             =   2315
         Width           =   3045
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
         MaxLength       =   40
         TabIndex        =   11
         Top             =   3260
         Width           =   6285
      End
      Begin VB.TextBox Text20 
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
         MaxLength       =   60
         TabIndex        =   18
         Top             =   4830
         Width           =   6285
      End
      Begin VB.TextBox Text19 
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
         MaxLength       =   60
         TabIndex        =   17
         Top             =   4515
         Width           =   6285
      End
      Begin VB.TextBox Text18 
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
         MaxLength       =   60
         TabIndex        =   16
         Top             =   4205
         Width           =   6285
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   3000
         TabIndex        =   33
         Top             =   3885
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
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
         Caption         =   " - "
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0258
         Picture         =   "CompanyMaster.frx":0274
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   3360
         TabIndex        =   14
         Top             =   3885
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   582
         Calendar        =   "CompanyMaster.frx":0290
         Caption         =   "CompanyMaster.frx":03A8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "CompanyMaster.frx":0414
         Keys            =   "CompanyMaster.frx":0432
         Spin            =   "CompanyMaster.frx":0490
         AlignHorizontal =   2
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   120
         TabIndex        =   34
         Top             =   3260
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
         Caption         =   " URL"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":04B8
         Picture         =   "CompanyMaster.frx":04D4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   120
         TabIndex        =   35
         Top             =   4205
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
         Caption         =   " Bank Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":04F0
         Picture         =   "CompanyMaster.frx":050C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   120
         TabIndex        =   36
         Top             =   4515
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
         Caption         =   " A/c No."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0528
         Picture         =   "CompanyMaster.frx":0544
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   37
         Top             =   4830
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
         Caption         =   " Branch && IFSC"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0560
         Picture         =   "CompanyMaster.frx":057C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   120
         TabIndex        =   32
         Top             =   1995
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
         Caption         =   " Phone"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0598
         Picture         =   "CompanyMaster.frx":05B4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   38
         Top             =   2945
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
         Caption         =   " E-Mail"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":05D0
         Picture         =   "CompanyMaster.frx":05EC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   1270
         Left            =   120
         TabIndex        =   31
         Top             =   740
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   2249
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Address"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0608
         Picture         =   "CompanyMaster.frx":0624
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   30
         Top             =   425
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
         Caption         =   " Print Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0640
         Picture         =   "CompanyMaster.frx":065C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   29
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
         Caption         =   " Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0678
         Picture         =   "CompanyMaster.frx":0694
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
         MaxLength       =   60
         TabIndex        =   0
         Top             =   105
         Width           =   6285
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
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   1
         Top             =   425
         Width           =   6285
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
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   2
         Top             =   740
         Width           =   6285
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
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1055
         Width           =   6285
      End
      Begin VB.TextBox Text5 
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
         TabIndex        =   4
         Top             =   1370
         Width           =   6285
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
         TabIndex        =   5
         Top             =   1685
         Width           =   6285
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
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1995
         Width           =   6285
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
         MaxLength       =   50
         TabIndex        =   10
         Top             =   2945
         Width           =   6285
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   120
         TabIndex        =   39
         Top             =   3890
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
         Caption         =   " Financial Year"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":06B0
         Picture         =   "CompanyMaster.frx":06CC
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   1680
         TabIndex        =   13
         Top             =   3885
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   582
         Calendar        =   "CompanyMaster.frx":06E8
         Caption         =   "CompanyMaster.frx":0800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "CompanyMaster.frx":086C
         Keys            =   "CompanyMaster.frx":088A
         Spin            =   "CompanyMaster.frx":08E8
         AlignHorizontal =   2
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   120
         TabIndex        =   41
         Top             =   2630
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
         Caption         =   " Fax"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0910
         Picture         =   "CompanyMaster.frx":092C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   120
         TabIndex        =   42
         Top             =   3575
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
         Caption         =   " GSTIN"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0948
         Picture         =   "CompanyMaster.frx":0964
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   120
         TabIndex        =   43
         Top             =   5140
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
         Caption         =   " TallyIntegration"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0980
         Picture         =   "CompanyMaster.frx":099C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   2400
         TabIndex        =   44
         Top             =   5145
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
         Caption         =   "  BusyIntegration"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":09B8
         Picture         =   "CompanyMaster.frx":09D4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   4680
         TabIndex        =   45
         Top             =   5145
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
         Caption         =   "  Company Alias"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":09F0
         Picture         =   "CompanyMaster.frx":0A0C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Left            =   4680
         TabIndex        =   48
         Top             =   3885
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
         Caption         =   "  FY Code"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0A28
         Picture         =   "CompanyMaster.frx":0A44
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Left            =   120
         TabIndex        =   49
         Top             =   5665
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
         Caption         =   " SMTP Server"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0A60
         Picture         =   "CompanyMaster.frx":0A7C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   4680
         TabIndex        =   50
         Top             =   5665
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
         Caption         =   "  Port"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0A98
         Picture         =   "CompanyMaster.frx":0AB4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   120
         TabIndex        =   51
         Top             =   5990
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
         Caption         =   " User Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0AD0
         Picture         =   "CompanyMaster.frx":0AEC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
         Height          =   330
         Left            =   4680
         TabIndex        =   52
         Top             =   5990
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
         Caption         =   " Password"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0B08
         Picture         =   "CompanyMaster.frx":0B24
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
         Height          =   330
         Left            =   4680
         TabIndex        =   53
         Top             =   2310
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
         Caption         =   "  State"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "CompanyMaster.frx":0B40
         Picture         =   "CompanyMaster.frx":0B5C
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   8085
         Y1              =   5570
         Y2              =   5570
      End
   End
End
Attribute VB_Name = "FrmCompanyMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strCreateCompany As String
Public ActionCancelled As Boolean
Dim StateCode As String
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstStateList As New ADODB.Recordset
Private Sub Form_Load()
    CenterForm Me
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    Text25.Visible = False
    If Trim(ReadFromFile("Super User")) = "EasyPublish" Then
    MhDateInput1.ReadOnly = False: MhDateInput2.ReadOnly = False: Text13.Locked = False: Text14.Locked = False: Text15.Locked = False: Text16.Locked = False: Option1.Enabled = True: Option2.Enabled = True
    End If
    If strCreateCompany = "Y" Then
        BusySystemIndicator True
        cnDatabase.CursorLocation = adUseClient
        If DatabaseType = "MS SQL" Then
            cnDatabase.CommandTimeout = 300
            ConnectionString = "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=EP000;Data Source=" & ServerName
        Else
            ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\EasyPublish." & CompCode & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
        End If
        cnDatabase.Open ConnectionString
        FrmCompanyMaster.Caption = "Company Creation"
        MhDateInput1.Text = Format(CDate("Apr-01-" & Year(Date)), "dd-MM-yyyy")
        MhDateInput2.Text = Format(CDate("Mar-31-" & Str(Year(Date) + 1)), "dd-MM-yyyy")
        BusySystemIndicator False
    Else
        FrmCompanyMaster.Caption = "Company Modification"
        rstCompanyMaster.CursorLocation = adUseServer
        BusySystemIndicator True
        rstCompanyMaster.Open "SELECT *,(Select Name From GeneralMaster Where Type=56 And Code=State) As StateName,State As StateCode FROM CompanyMaster Where FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockPessimistic
        rstStateList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '56'  Order By Name", cnDatabase, adOpenKeyset, adLockReadOnly
        LoadFields
        BusySystemIndicator False
        MdiMainMenu.MousePointer = vbHourglass
        rstCompanyMaster.Fields("PrintStatus") = "N"
        MdiMainMenu.MousePointer = vbNormal
    End If
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the Company")
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstStateList)
    BusySystemIndicator False
    MdiMainMenu.MousePointer = vbNormal
    Call CloseForm(Me)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
       cmdExit_Click
       If strCreateCompany = "Y" Then Cancel = 1
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
       Sendkeys "{TAB}"
       KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
       cmdSave_Click
       KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
       cmdExit_Click
       KeyCode = 0
    End If
End Sub
Private Sub LoadFields()
    If rstCompanyMaster.EOF Or rstCompanyMaster.BOF Then Exit Sub
    Text1.Text = rstCompanyMaster.Fields("Name").Value
    Text2.Text = rstCompanyMaster.Fields("PrintName").Value
    Text3.Text = rstCompanyMaster.Fields("Address1").Value
    Text4.Text = CheckNull(rstCompanyMaster.Fields("Address2").Value)
    Text5.Text = CheckNull(rstCompanyMaster.Fields("Address3").Value)
    Text6.Text = CheckNull(rstCompanyMaster.Fields("Address4").Value)
    Text7.Text = CheckNull(rstCompanyMaster.Fields("Phone").Value)
    Text11.Text = CheckNull(rstCompanyMaster.Fields("Mobile").Value)
    Text12.Text = CheckNull(rstCompanyMaster.Fields("Fax").Value)
    Text8.Text = CheckNull(rstCompanyMaster.Fields("EMail").Value)
    Text9.Text = CheckNull(rstCompanyMaster.Fields("Website").Value)
    MhDateInput1.Text = Format(rstCompanyMaster.Fields("FinancialYearFrom").Value, "dd-MM-yyyy")
    MhDateInput2.Text = Format(rstCompanyMaster.Fields("FinancialYearTo").Value, "dd-MM-yyyy")
    Text10.Text = rstCompanyMaster.Fields("GSTIN").Value
    Text18.Text = rstCompanyMaster.Fields("BankName").Value
    Text19.Text = rstCompanyMaster.Fields("AccountNo").Value
    Text20.Text = rstCompanyMaster.Fields("IFSC").Value
    Text15.Text = rstCompanyMaster.Fields("Alias").Value
    Text16.Text = rstCompanyMaster.Fields("FYCode").Value
    Option1.Value = rstCompanyMaster.Fields("TallyIntegration").Value
    Option2.Value = rstCompanyMaster.Fields("BusyIntegration").Value
    Text17.Text = rstCompanyMaster.Fields("SMTPServer").Value
    Text21.Text = rstCompanyMaster.Fields("Port").Value
    Text22.Text = rstCompanyMaster.Fields("UserName").Value
    Text23.Text = rstCompanyMaster.Fields("Password").Value
    Text24.Text = rstCompanyMaster.Fields("StateName").Value
    Text25.Text = rstCompanyMaster.Fields("StateCode").Value
    StateCode = rstCompanyMaster.Fields("StateCode").Value
End Sub
Private Sub SaveFields()
    If rstCompanyMaster.EOF Or rstCompanyMaster.BOF Then Exit Sub
    rstCompanyMaster.Fields("Name").Value = Left(Trim(Text1.Text), 40)
    rstCompanyMaster.Fields("PrintName").Value = Left(Trim(Text2.Text), 40)
    rstCompanyMaster.Fields("Address1").Value = Trim(Text3.Text)
    rstCompanyMaster.Fields("Address2").Value = Trim(Text4.Text)
    rstCompanyMaster.Fields("Address3").Value = Trim(Text5.Text)
    rstCompanyMaster.Fields("Address4").Value = Trim(Text6.Text)
    rstCompanyMaster.Fields("Phone").Value = Trim(Text7.Text)
    rstCompanyMaster.Fields("Mobile").Value = Trim(Text11.Text)
    rstCompanyMaster.Fields("Fax").Value = Trim(Text12.Text)
    rstCompanyMaster.Fields("EMail").Value = Trim(Text8.Text)
    rstCompanyMaster.Fields("Website").Value = Trim(Text9.Text)
    rstCompanyMaster.Fields("FinancialYearFrom").Value = GetDate(MhDateInput1.Text)
    rstCompanyMaster.Fields("FinancialYearTo").Value = GetDate(MhDateInput2.Text)
    rstCompanyMaster.Fields("GSTIN").Value = Trim(Text10.Text)
    rstCompanyMaster.Fields("BankName").Value = Trim(Text18.Text)
    rstCompanyMaster.Fields("AccountNo").Value = Trim(Text19.Text)
    rstCompanyMaster.Fields("IFSC").Value = Trim(Text20.Text)
    rstCompanyMaster.Fields("Printstatus").Value = "N"
    rstCompanyMaster.Fields("Alias").Value = Trim(Text15.Text)
    rstCompanyMaster.Fields("FYCode").Value = Trim(Text16.Text)
    rstCompanyMaster.Fields("TallyIntegration") = Option1.Value
    rstCompanyMaster.Fields("BusyIntegration") = Option2.Value
    rstCompanyMaster.Fields("SMTPServer") = Text17.Text
    rstCompanyMaster.Fields("Port") = Text21.Text
    rstCompanyMaster.Fields("UserName") = Text22.Text
    rstCompanyMaster.Fields("Password") = Text23.Text
    rstCompanyMaster.Fields("State") = StateCode
End Sub
Private Sub Text1_Validate(Cancel As Boolean)   'Name
    If CheckEmpty(Text1, True) Then
        Cancel = True
    ElseIf strCreateCompany = "Y" Then
        Text2.Text = Text1.Text
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)   'Print Name
    If CheckEmpty(Text2.Text, True) Then Cancel = True
End Sub
Private Sub Text3_Validate(Cancel As Boolean)   'Address
    If CheckEmpty(Text3, True) Then Cancel = True
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput2.Text)) Then
        Cancel = True
    ElseIf (Format(GetDate(MhDateInput2.Text), "yyyymmdd") <= Format(GetDate(MhDateInput1.Text), "yyyymmdd")) Then
       Cancel = True
    End If
End Sub
Private Sub Text10_Validate(Cancel As Boolean)  'GSTIN
    If CheckEmpty(Text10, True) Then Cancel = True
End Sub
Private Sub cmdExit_Click()
    If strCreateCompany <> "Y" Then
        If CancelRecordUpdate(rstCompanyMaster) Then Call CloseRecordset(rstCompanyMaster): Call CloseForm(FrmCompanyMaster)
    Else
        ActionCancelled = True: Me.Hide
    End If
End Sub
Private Sub cmdSave_Click()
    If CheckMandatoryFields Then Exit Sub
    If strCreateCompany = "Y" Then
        Me.Hide
    Else
        SaveFields
        If UpdateRecord(rstCompanyMaster) Then
            rstCompanyMaster.Close
            If DatabaseType = "MS SQL" Then
            rstCompanyMaster.Open "SELECT Name,'-Financial Year From '+REPLACE(CONVERT(VARCHAR(11),FinancialYearFrom,106),' ','-')+' To '+REPLACE(CONVERT(VARCHAR(11),FinancialYearTo,106),' ','-'),* FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
            Else
            rstCompanyMaster.Open "SELECT Name,'-Financial Year From '+REPLACE(CONVERT(VARCHAR(11),FinancialYearFrom,106),' ','-')+' To '+REPLACE(CONVERT(VARCHAR(11),FinancialYearTo,106),' ','-'),* FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
            End If
            MdiMainMenu.Caption = "Easy Publish 19 Rel 4.0-Production Management System [" & Trim(rstCompanyMaster.Fields("Name").Value) & Trim(rstCompanyMaster.Fields(1).Value) & "]"
            Call CloseRecordset(rstCompanyMaster)
            Call CloseForm(FrmCompanyMaster)
        Else
            Call DisplayError("Failed to Edit the Company")
            cmdExit_Click
        End If
    End If
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text1.Text, False) Then
        Text1.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text2.Text, False) Then
        Text2.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus:  CheckMandatoryFields = True
    ElseIf (Format(GetDate(MhDateInput2.Text), "yyyymmdd") <= Format(GetDate(MhDateInput1.Text), "yyyymmdd")) Then
        MhDateInput2.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text10.Text, False) Then
        Text10.SetFocus: CheckMandatoryFields = True
    ElseIf CheckEmpty(Text24.Text, False) Then
        Text24.SetFocus: CheckMandatoryFields = True
    End If
End Function
'        cnDatabase.Execute "INSERT INTO CompanyMaster (Code,Name,PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,Fax,eMail,Website,GSTIN,CreatedFrom,MCGroup,MCPrimary,MCRepair,FinancialYearFrom,FinancialYearTo,Printstatus,TitleCombo,BankName,AccountNo,IFSC,TallyIntegration,BusyIntegration,FYCode,Alias,SMTPServer,Port,UserName,Password) VALUES ('000001','" & Trim(FrmCompanyMaster.Text1.Text) & "','" & Trim(FrmCompanyMaster.Text2.Text) & "','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text12.Text) & "'" & _
                                                  ",'" & Trim(FrmCompanyMaster.Text8.Text) & "','" & Trim(FrmCompanyMaster.Text9.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & CompCode & "','0','0','0','" & Format(GetDate(FrmCompanyMaster.MhDateInput1.Text), "mm-dd-yyyy") & "','" & Format(GetDate(FrmCompanyMaster.MhDateInput2.Text), "mm-dd-yyyy") & "','N','1','" & Trim(FrmCompanyMaster.Text18.Text) & "','" & Trim(FrmCompanyMaster.Text19.Text) & "','" & Trim(FrmCompanyMaster.Text20.Text) & "','" & Trim(FrmCompanyMaster.Option1.Value) & "','" & Trim(FrmCompanyMaster.Option2.Value) & "','" & Trim(FrmCompanyMaster.Text16.Text) & "','" & Trim(FrmCompanyMaster.Text15.Text) & "','" & Trim(FrmCompanyMaster.Text17.Text) & "','" & Trim(FrmCompanyMaster.Text21.Text) & "','" & Trim(FrmCompanyMaster.Text22.Text) & "','" & Trim(FrmCompanyMaster.Text23.Text) & "')"
Private Sub Text24_Change()
    If Text24.Text = " " Then Text24.Text = "?": Sendkeys "{TAB}"
End Sub
Private Sub Text24_Validate(Cancel As Boolean)
    Dim SearchString As String
    If rstStateList.State = 1 Then rstStateList.Close
    rstStateList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '56'  Order By Name", cnDatabase, adOpenKeyset, adLockReadOnly
    SearchString = FixQuote(Text24.Text)
    If rstStateList.RecordCount = 0 Then
       DisplayError ("No Record in State Master")
       Cancel = True
       Exit Sub
    Else
       rstStateList.MoveFirst
    End If
    rstStateList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstStateList.EOF Then
       SelectionType = "S"
       StateCode = ""
       Call LoadSelectionList(rstStateList, "List of States...", "Name")
       SearchOrder = 0
       Call DisplaySelectionList(Text24, StateCode)
       Text25.Text = StateCode
       Call CloseForm(FrmSelectionList)
       If CheckEmpty(Text24.Text, False) Then Text24.Text = "?"
       If RTrim(StateCode) <> "" Then Sendkeys "{TAB}"
       Cancel = True
    Else
       StateCode = rstStateList.Fields("Code").Value
       Text25.Text = rstStateList.Fields("Code").Value
    End If
End Sub

