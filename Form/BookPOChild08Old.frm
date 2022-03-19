VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmBookPOChild08 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binding Order Details"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
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
   ScaleHeight     =   6555
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008BD6FE&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10733
      Picture         =   "BookPOChild08.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      BackColor       =   &H008BD6FE&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10733
      Picture         =   "BookPOChild08.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   6315
      Left            =   120
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   105
      Width           =   9975
      _Version        =   65536
      _ExtentX        =   17595
      _ExtentY        =   11139
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Picture         =   "BookPOChild08.frx":0204
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   48
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
         Caption         =   " Item Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0220
         Picture         =   "BookPOChild08.frx":023C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   49
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
         Caption         =   " Actual Quantity"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0258
         Picture         =   "BookPOChild08.frx":0274
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   50
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
         Caption         =   " Folding Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0290
         Picture         =   "BookPOChild08.frx":02AC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   120
         TabIndex        =   52
         Top             =   3480
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
         Caption         =   " Rate/Piece"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":02C8
         Picture         =   "BookPOChild08.frx":02E4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   57
         Top             =   5010
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
         Caption         =   " Bill No."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0300
         Picture         =   "BookPOChild08.frx":031C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   61
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
         Caption         =   " Party Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0338
         Picture         =   "BookPOChild08.frx":0354
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   63
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
         Picture         =   "BookPOChild08.frx":0370
         Picture         =   "BookPOChild08.frx":038C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   64
         Top             =   5535
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
         Picture         =   "BookPOChild08.frx":03A8
         Picture         =   "BookPOChild08.frx":03C4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   3360
         TabIndex        =   65
         Top             =   3480
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
         Caption         =   " Cartage/Box"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":03E0
         Picture         =   "BookPOChild08.frx":03FC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   67
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
         Caption         =   " Binding Form"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0418
         Picture         =   "BookPOChild08.frx":0434
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   120
         TabIndex        =   70
         Top             =   4100
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
         Caption         =   " Adjustment"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0450
         Picture         =   "BookPOChild08.frx":046C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
         Height          =   330
         Left            =   120
         TabIndex        =   72
         Top             =   2220
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
         Caption         =   " Folding Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0488
         Picture         =   "BookPOChild08.frx":04A4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Left            =   120
         TabIndex        =   75
         Top             =   5850
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
         Picture         =   "BookPOChild08.frx":04C0
         Picture         =   "BookPOChild08.frx":04DC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel30 
         Height          =   330
         Left            =   120
         TabIndex        =   76
         Top             =   3170
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
         Caption         =   " Total Pkt."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":04F8
         Picture         =   "BookPOChild08.frx":0514
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel31 
         Height          =   330
         Left            =   6600
         TabIndex        =   77
         Top             =   3170
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
         Caption         =   " Total Boxes"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0530
         Picture         =   "BookPOChild08.frx":054C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel37 
         Height          =   330
         Left            =   120
         TabIndex        =   82
         Top             =   2850
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
         Caption         =   " Loose Qty/Box"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0568
         Picture         =   "BookPOChild08.frx":0584
      End
      Begin VB.TextBox TxtAdNar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   41
         Top             =   5850
         Width           =   8055
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   40
         Top             =   5535
         Width           =   4815
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   105
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   640
         Width           =   8055
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   37
         Top             =   5010
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   965
         Width           =   8055
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   5040
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1275
         Width           =   4815
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   3360
         TabIndex        =   47
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
         Picture         =   "BookPOChild08.frx":05A0
         Picture         =   "BookPOChild08.frx":05BC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   3360
         TabIndex        =   51
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
         Caption         =   " Stitching Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":05D8
         Picture         =   "BookPOChild08.frx":05F4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   3360
         TabIndex        =   53
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
         Caption         =   " Binding Type"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0610
         Picture         =   "BookPOChild08.frx":062C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   3360
         TabIndex        =   54
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
         Caption         =   " Billing Quantity"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0648
         Picture         =   "BookPOChild08.frx":0664
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   6600
         TabIndex        =   55
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
         Caption         =   " Binding Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0680
         Picture         =   "BookPOChild08.frx":069C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   120
         TabIndex        =   56
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
         Caption         =   " Qty/Pkt."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":06B8
         Picture         =   "BookPOChild08.frx":06D4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   6600
         TabIndex        =   58
         Top             =   5010
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
         Picture         =   "BookPOChild08.frx":06F0
         Picture         =   "BookPOChild08.frx":070C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   3360
         TabIndex        =   59
         Top             =   5010
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
         Picture         =   "BookPOChild08.frx":0728
         Picture         =   "BookPOChild08.frx":0744
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   6600
         TabIndex        =   60
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
         Picture         =   "BookPOChild08.frx":0760
         Picture         =   "BookPOChild08.frx":077C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Left            =   6600
         TabIndex        =   62
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
         Caption         =   " Adj.Quantity"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0798
         Picture         =   "BookPOChild08.frx":07B4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   6600
         TabIndex        =   66
         Top             =   4410
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
         Caption         =   " Total Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":07D0
         Picture         =   "BookPOChild08.frx":07EC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   3360
         TabIndex        =   69
         Top             =   4410
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
         Caption         =   " GST"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0808
         Picture         =   "BookPOChild08.frx":0824
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Left            =   3360
         TabIndex        =   43
         Top             =   3170
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
         Caption         =   " Pkt./Box"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0840
         Picture         =   "BookPOChild08.frx":085C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Left            =   6600
         TabIndex        =   71
         Top             =   5535
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
         Caption         =   " Received Qty"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0878
         Picture         =   "BookPOChild08.frx":0894
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput3 
         Height          =   330
         Left            =   8280
         TabIndex        =   2
         Top             =   105
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild08.frx":08B0
         Caption         =   "BookPOChild08.frx":09C8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":0A34
         Keys            =   "BookPOChild08.frx":0A52
         Spin            =   "BookPOChild08.frx":0AB0
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
         Height          =   330
         Left            =   3360
         TabIndex        =   73
         Top             =   2220
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
         Caption         =   " Stitching Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0AD8
         Picture         =   "BookPOChild08.frx":0AF4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel29 
         Height          =   330
         Left            =   6600
         TabIndex        =   74
         Top             =   2220
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
         Caption         =   " Binding Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":0B10
         Picture         =   "BookPOChild08.frx":0B2C
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   1800
         TabIndex        =   5
         Top             =   1275
         Width           =   840
         _Version        =   65536
         _ExtentX        =   1482
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":0B48
         Caption         =   "BookPOChild08.frx":0B68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":0BD4
         Keys            =   "BookPOChild08.frx":0BF2
         Spin            =   "BookPOChild08.frx":0C3C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0"
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
         MaxValue        =   99999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2001141765
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
         Height          =   330
         Left            =   2625
         TabIndex        =   6
         Top             =   1275
         Width           =   750
         _Version        =   65536
         _ExtentX        =   1323
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":0C64
         Caption         =   "BookPOChild08.frx":0C84
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":0CF0
         Keys            =   "BookPOChild08.frx":0D0E
         Spin            =   "BookPOChild08.frx":0D58
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0"
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
         MaxValue        =   99999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2001141765
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   1800
         TabIndex        =   8
         Top             =   1590
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":0D80
         Caption         =   "BookPOChild08.frx":0DA0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":0E0C
         Keys            =   "BookPOChild08.frx":0E2A
         Spin            =   "BookPOChild08.frx":0E74
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   8280
         TabIndex        =   10
         Top             =   1590
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":0E9C
         Caption         =   "BookPOChild08.frx":0EBC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":0F28
         Keys            =   "BookPOChild08.frx":0F46
         Spin            =   "BookPOChild08.frx":0F90
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   5040
         TabIndex        =   24
         Top             =   3170
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":0FB8
         Caption         =   "BookPOChild08.frx":0FD8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1044
         Keys            =   "BookPOChild08.frx":1062
         Spin            =   "BookPOChild08.frx":10AC
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
         ValueVT         =   2001141765
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
         Height          =   330
         Left            =   8280
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3170
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":10D4
         Caption         =   "BookPOChild08.frx":10F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1160
         Keys            =   "BookPOChild08.frx":117E
         Spin            =   "BookPOChild08.frx":11C8
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   1800
         TabIndex        =   11
         Top             =   1905
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":11F0
         Caption         =   "BookPOChild08.frx":1210
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":127C
         Keys            =   "BookPOChild08.frx":129A
         Spin            =   "BookPOChild08.frx":12E4
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   8280
         TabIndex        =   13
         Top             =   1905
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":130C
         Caption         =   "BookPOChild08.frx":132C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1398
         Keys            =   "BookPOChild08.frx":13B6
         Spin            =   "BookPOChild08.frx":1400
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
         Height          =   330
         Left            =   1800
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1428
         Caption         =   "BookPOChild08.frx":1448
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":14B4
         Keys            =   "BookPOChild08.frx":14D2
         Spin            =   "BookPOChild08.frx":151C
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   36241413
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
         Height          =   330
         Left            =   8280
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1544
         Caption         =   "BookPOChild08.frx":1564
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":15D0
         Keys            =   "BookPOChild08.frx":15EE
         Spin            =   "BookPOChild08.frx":1638
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
         Height          =   330
         Left            =   1800
         TabIndex        =   26
         Top             =   3480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1660
         Caption         =   "BookPOChild08.frx":1680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":16EC
         Keys            =   "BookPOChild08.frx":170A
         Spin            =   "BookPOChild08.frx":1754
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   5040
         TabIndex        =   27
         Top             =   3480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":177C
         Caption         =   "BookPOChild08.frx":179C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1808
         Keys            =   "BookPOChild08.frx":1826
         Spin            =   "BookPOChild08.frx":1870
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
         Height          =   330
         Left            =   5040
         TabIndex        =   34
         Top             =   4410
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1898
         Caption         =   "BookPOChild08.frx":18B8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1924
         Keys            =   "BookPOChild08.frx":1942
         Spin            =   "BookPOChild08.frx":198C
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
         ValueVT         =   2001141765
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   5760
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   4410
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":19B4
         Caption         =   "BookPOChild08.frx":19D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1A40
         Keys            =   "BookPOChild08.frx":1A5E
         Spin            =   "BookPOChild08.frx":1AA8
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
         ReadOnly        =   1
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
         TabIndex        =   32
         Top             =   4095
         Width           =   8055
         _Version        =   65536
         _ExtentX        =   14208
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1AD0
         Caption         =   "BookPOChild08.frx":1AF0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1B5C
         Keys            =   "BookPOChild08.frx":1B7A
         Spin            =   "BookPOChild08.frx":1BC4
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
         MinValue        =   -9999999999.99
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
         Left            =   8280
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4410
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1BEC
         Caption         =   "BookPOChild08.frx":1C0C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1C78
         Keys            =   "BookPOChild08.frx":1C96
         Spin            =   "BookPOChild08.frx":1CE0
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
         Height          =   330
         Left            =   8280
         TabIndex        =   39
         Top             =   5010
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1D08
         Caption         =   "BookPOChild08.frx":1D28
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1D94
         Keys            =   "BookPOChild08.frx":1DB2
         Spin            =   "BookPOChild08.frx":1DFC
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
         Height          =   330
         Left            =   8280
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   5535
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1E24
         Caption         =   "BookPOChild08.frx":1E44
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1EB0
         Keys            =   "BookPOChild08.frx":1ECE
         Spin            =   "BookPOChild08.frx":1F18
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel33 
         Height          =   330
         Left            =   3360
         TabIndex        =   78
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
         Caption         =   " Pkt. Packing Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":1F40
         Picture         =   "BookPOChild08.frx":1F5C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel34 
         Height          =   330
         Left            =   6600
         TabIndex        =   79
         Top             =   3480
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
         Caption         =   " Box Packing Rate"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":1F78
         Picture         =   "BookPOChild08.frx":1F94
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel35 
         Height          =   330
         Left            =   6600
         TabIndex        =   80
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
         Caption         =   "Pkt. Packing Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":1FB0
         Picture         =   "BookPOChild08.frx":1FCC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel36 
         Height          =   330
         Left            =   6600
         TabIndex        =   81
         Top             =   3800
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
         Caption         =   " Box Pack Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":1FE8
         Picture         =   "BookPOChild08.frx":2004
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
         Height          =   330
         Left            =   1800
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3170
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2020
         Caption         =   "BookPOChild08.frx":2040
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":20AC
         Keys            =   "BookPOChild08.frx":20CA
         Spin            =   "BookPOChild08.frx":2114
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
         Height          =   330
         Left            =   3360
         TabIndex        =   83
         Top             =   2850
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
         Caption         =   " Extra Loose Qty"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":213C
         Picture         =   "BookPOChild08.frx":2158
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel39 
         Height          =   330
         Left            =   6600
         TabIndex        =   84
         Top             =   2850
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
         Caption         =   " Total Loose Qty"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":2174
         Picture         =   "BookPOChild08.frx":2190
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput29 
         Height          =   330
         Left            =   1800
         TabIndex        =   20
         Top             =   2850
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":21AC
         Caption         =   "BookPOChild08.frx":21CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2238
         Keys            =   "BookPOChild08.frx":2256
         Spin            =   "BookPOChild08.frx":22A0
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
         ValueVT         =   2001141765
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   5040
         TabIndex        =   1
         Top             =   105
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild08.frx":22C8
         Caption         =   "BookPOChild08.frx":23E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":244C
         Keys            =   "BookPOChild08.frx":246A
         Spin            =   "BookPOChild08.frx":24C8
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
         TabIndex        =   38
         Top             =   5010
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild08.frx":24F0
         Caption         =   "BookPOChild08.frx":2608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2674
         Keys            =   "BookPOChild08.frx":2692
         Spin            =   "BookPOChild08.frx":26F0
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
         Height          =   330
         Left            =   8280
         TabIndex        =   28
         Top             =   3480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2718
         Caption         =   "BookPOChild08.frx":2738
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":27A4
         Keys            =   "BookPOChild08.frx":27C2
         Spin            =   "BookPOChild08.frx":280C
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   5040
         TabIndex        =   9
         Top             =   1590
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2834
         Caption         =   "BookPOChild08.frx":2854
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":28C0
         Keys            =   "BookPOChild08.frx":28DE
         Spin            =   "BookPOChild08.frx":2928
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   5040
         TabIndex        =   12
         Top             =   1905
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2950
         Caption         =   "BookPOChild08.frx":2970
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":29DC
         Keys            =   "BookPOChild08.frx":29FA
         Spin            =   "BookPOChild08.frx":2A44
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
         Height          =   330
         Left            =   5040
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2A6C
         Caption         =   "BookPOChild08.frx":2A8C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2AF8
         Keys            =   "BookPOChild08.frx":2B16
         Spin            =   "BookPOChild08.frx":2B60
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
         Left            =   1800
         TabIndex        =   17
         Top             =   2535
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2B88
         Caption         =   "BookPOChild08.frx":2BA8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2C14
         Keys            =   "BookPOChild08.frx":2C32
         Spin            =   "BookPOChild08.frx":2C7C
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
         ValueVT         =   2001141765
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
         Height          =   330
         Left            =   5040
         TabIndex        =   18
         Top             =   2535
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2CA4
         Caption         =   "BookPOChild08.frx":2CC4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2D30
         Keys            =   "BookPOChild08.frx":2D4E
         Spin            =   "BookPOChild08.frx":2D98
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput30 
         Height          =   330
         Left            =   5040
         TabIndex        =   21
         Top             =   2850
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2DC0
         Caption         =   "BookPOChild08.frx":2DE0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2E4C
         Keys            =   "BookPOChild08.frx":2E6A
         Spin            =   "BookPOChild08.frx":2EB4
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
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel40 
         Height          =   330
         Left            =   120
         TabIndex        =   86
         Top             =   3800
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
         Caption         =   " Amount( Rate/Pec)"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":2EDC
         Picture         =   "BookPOChild08.frx":2EF8
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput32 
         Height          =   330
         Left            =   1800
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3800
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2F14
         Caption         =   "BookPOChild08.frx":2F34
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2FA0
         Keys            =   "BookPOChild08.frx":2FBE
         Spin            =   "BookPOChild08.frx":3008
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput28 
         Height          =   330
         Left            =   8280
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3800
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":3030
         Caption         =   "BookPOChild08.frx":3050
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":30BC
         Keys            =   "BookPOChild08.frx":30DA
         Spin            =   "BookPOChild08.frx":3124
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   3360
         TabIndex        =   68
         Top             =   3800
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
         Caption         =   " Cartage Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":314C
         Picture         =   "BookPOChild08.frx":3168
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
         Height          =   330
         Left            =   5040
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3800
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":3184
         Caption         =   "BookPOChild08.frx":31A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":3210
         Keys            =   "BookPOChild08.frx":322E
         Spin            =   "BookPOChild08.frx":3278
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel41 
         Height          =   330
         Left            =   120
         TabIndex        =   87
         Top             =   4410
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
         Caption         =   " Taxable Amount"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild08.frx":32A0
         Picture         =   "BookPOChild08.frx":32BC
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput33 
         Height          =   330
         Left            =   1800
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4410
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":32D8
         Caption         =   "BookPOChild08.frx":32F8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":3364
         Keys            =   "BookPOChild08.frx":3382
         Spin            =   "BookPOChild08.frx":33CC
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
         Height          =   330
         Left            =   8280
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2535
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":33F4
         Caption         =   "BookPOChild08.frx":3414
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":3480
         Keys            =   "BookPOChild08.frx":349E
         Spin            =   "BookPOChild08.frx":34E8
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput31 
         Height          =   330
         Left            =   8280
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2850
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":3510
         Caption         =   "BookPOChild08.frx":3530
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":359C
         Keys            =   "BookPOChild08.frx":35BA
         Spin            =   "BookPOChild08.frx":3604
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2001141765
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10000
         Y1              =   5430
         Y2              =   5430
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   10000
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10000
         Y1              =   4880
         Y2              =   4880
      End
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   2
      Left            =   10200
      TabIndex        =   88
      Top             =   4920
      Width           =   1440
      _Version        =   65536
      _ExtentX        =   2540
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
      Caption         =   "Ctrl+S->Save"
      AutoSize        =   -1  'True
      FillColor       =   8421504
      TextColor       =   16777215
      Picture         =   "BookPOChild08.frx":362C
      Multiline       =   -1  'True
      GlobalMem       =   -1  'True
      Picture         =   "BookPOChild08.frx":3648
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   375
      Left            =   10200
      TabIndex        =   85
      Top             =   5400
      Width           =   1440
      BackColor       =   9164542
      Caption         =   "Update Master"
      Size            =   "2540;661"
      FontName        =   "Arial"
      FontEffects     =   1073741827
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "FrmBookPOChild08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstBookPOChild08 As New ADODB.Recordset
Dim rstBindingTypeList As New ADODB.Recordset
Dim rstBinderRates As New ADODB.Recordset
Public BinderCode As String
Public BookPrinterQuantity As Long
Dim FormType As String
Dim SizeCode As String
Dim BindingTypeCode As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    DisableCloseButton Me
    FormType = FrmBookPrintOrder.rstBookList.Fields("FormType").Value
    SizeCode = FrmBookPrintOrder.rstBookList.Fields("SizeCode").Value
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text)
    Text4.Text = Trim(FrmBookPrintOrder.Text8.Text)
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text)
    ClearFields
    rstBindingTypeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '6' Order By Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstBindingTypeList.ActiveConnection = Nothing
'    If IsNull(rstBookPOChild08.Fields("Code").Value) Then MhRealInput2.ReadOnly = True: MhRealInput5.ReadOnly = True: MhRealInput6.ReadOnly = True: MhRealInput8.ReadOnly = True Else MhRealInput2.ReadOnly = False: MhRealInput5.ReadOnly = False: MhRealInput6.ReadOnly = False: MhRealInput8.ReadOnly = False
    If Val(CheckNull(rstBookPOChild08.Fields("ActualQuantity").Value)) = 0 Then
        MhRealInput1.Value = FrmBookPrintOrder.MhRealInput3.Value
'        MhRealInput1.Text = Format(BookPrinterQuantity, "0")
        MhRealInput7.Text = Format(Val(FrmBookPrintOrder.rstBookList.Fields("BindingForms01").Value), "0")
        MhRealInput18.Text = Format(Val(FrmBookPrintOrder.rstBookList.Fields("BindingForms02").Value), "0")
        MhRealInput23.Text = Format(Val(FrmBookPrintOrder.rstBookList.Fields("Qty/Pkt").Value), "0")
        MhRealInput11.Text = Format(Val(FrmBookPrintOrder.rstBookList.Fields("Pkt/Box").Value), "0")
        MhRealInput29.Text = Format(Val(FrmBookPrintOrder.rstBookList.Fields("LooseQty/Box").Value), "0")
        BindingTypeCode = FrmBookPrintOrder.rstBookList.Fields("BindingType").Value
        If rstBindingTypeList.RecordCount > 0 Then rstBindingTypeList.MoveFirst
        rstBindingTypeList.Find "[Code] = '" & BindingTypeCode & "'"
        If Not rstBindingTypeList.EOF Then Text3.Text = rstBindingTypeList.Fields("Col0").Value
        MhDateInput1.Text = Format(GetDate(FrmBookPrintOrder.MhDateInput1.Text), "dd-MM-yyyy")
        MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
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
       SendKeys "{TAB}"
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
       Call CloseForm(Me)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    BookPrinterQuantity = 0
    Call CloseRecordset(rstBindingTypeList)
    Call CloseRecordset(rstBinderRates)
End Sub
Private Sub ClearFields()
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput2.Text = "  -  -    "
    MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
    Text3.Text = ""
    Text6.Text = ""
    Text8.Text = ""
    MhRealInput1.Text = "0"
    MhRealInput2.Text = "0.00"
    MhRealInput3.Text = "0"
    MhRealInput4.Text = "0"
    MhRealInput5.Text = "0.00"
    MhRealInput6.Text = "0.00"
    MhRealInput7.Text = "0"
    MhRealInput8.Text = "0.00"
    MhRealInput9.Text = "0.00"
    MhRealInput10.Text = "0.00"
    MhRealInput11.Text = "0"
    MhRealInput12.Text = "0"
    MhRealInput13.Text = "0.00"
    MhRealInput14.Text = "0.00"
    MhRealInput15.Text = "0.00"
    MhRealInput16.Text = "0.00"
    MhRealInput17.Text = "0.00"
    MhRealInput18.Text = "0"
    MhRealInput19.Text = "0"
    MhRealInput20.Text = "0.00"
    MhRealInput21.Text = "0.00"
    MhRealInput22.Text = "0.00"
    MhRealInput23.Text = "0"
    MhRealInput24.Text = "0"
    MhRealInput25.Text = "0.00"
    MhRealInput26.Text = "0.00"
    MhRealInput27.Text = "0.00"
    MhRealInput28.Text = "0.00"
    MhRealInput29.Text = "0"
    MhRealInput30.Text = "0"
    MhRealInput31.Text = "0"
    TxtAdNar.Text = ""
End Sub
Private Sub LoadFields()
    If rstBookPOChild08.RecordCount = 0 Then Exit Sub
    MhDateInput1.Text = Format(rstBookPOChild08.Fields("OrderDate").Value, "dd-MM-yyyy")
    MhDateInput3.Text = Format(rstBookPOChild08.Fields("TargetDate").Value, "dd-MM-yyyy")
    MhRealInput7.Text = Format(Val(rstBookPOChild08.Fields("BindingForms").Value), "0")
    MhRealInput18.Text = Format(Val(rstBookPOChild08.Fields("ExtraForms").Value), "0")
    MhRealInput1.Text = Format(Val(rstBookPOChild08.Fields("ActualQuantity").Value), "0")
    MhRealInput3.Text = Format(Val(rstBookPOChild08.Fields("BillingQuantity").Value), "0")
    MhRealInput4.Text = Format(Val(rstBookPOChild08.Fields("AdjustQuantity").Value), "0")
    MhRealInput2.Text = Format(Val(rstBookPOChild08.Fields("FormFoldRate").Value), "0.00")
    MhRealInput5.Text = Format(Val(rstBookPOChild08.Fields("FormStitchRate").Value), "0.00")
    MhRealInput6.Text = Format(Val(rstBookPOChild08.Fields("FormPasteRate").Value), "0.00")
    MhRealInput8.Text = Format(Val(rstBookPOChild08.Fields("Rate/Book").Value), "0.00")
    MhRealInput29.Text = Format(Val(rstBookPOChild08.Fields("LooseQty/Box").Value), "0")
    MhRealInput30.Text = Format(Val(rstBookPOChild08.Fields("ExtraLooseQty").Value), "0")
    MhRealInput31.Text = Format(Val(rstBookPOChild08.Fields("TotalLooseQty").Value), "0")
    MhRealInput23.Text = Format(Val(rstBookPOChild08.Fields("Qty/Pkt").Value), "0")
    MhRealInput24.Text = Format(Val(rstBookPOChild08.Fields("TotalPkts").Value), "0")
    MhRealInput11.Text = Format(Val(rstBookPOChild08.Fields("Pkt/Box").Value), "0")
    MhRealInput12.Text = Format(Val(rstBookPOChild08.Fields("TotalBoxes").Value), "0")
    MhRealInput25.Text = Format(Val(rstBookPOChild08.Fields("PktPackRate").Value), "0.00")
    MhRealInput26.Text = Format(Val(rstBookPOChild08.Fields("BoxPackRate").Value), "0.00")
    MhRealInput13.Text = Format(Val(rstBookPOChild08.Fields("CartageRate").Value), "0.00")
    MhRealInput9.Text = Format(Val(rstBookPOChild08.Fields("Adjustment").Value), "0.00")
    MhRealInput10.Text = Format(Val(rstBookPOChild08.Fields("BillAmount").Value), "0.00")
    BindingTypeCode = rstBookPOChild08.Fields("BindingType").Value
    If rstBindingTypeList.RecordCount > 0 Then rstBindingTypeList.MoveFirst
    rstBindingTypeList.Find "[Code] = '" & BindingTypeCode & "'"
    If Not rstBindingTypeList.EOF Then Text3.Text = rstBindingTypeList.Fields("Col0").Value
    Text8.Text = rstBookPOChild08.Fields("BillNo").Value
    If Not IsNull(rstBookPOChild08.Fields("BillDate").Value) Then MhDateInput2.Text = Format(rstBookPOChild08.Fields("BillDate").Value, "dd-MM-yyyy")
    MhRealInput15.Text = Format(Val(rstBookPOChild08.Fields("VAT%").Value), "0.00")
    MhRealInput17.Text = Format(Val(rstBookPOChild08.Fields("VAT").Value), "0.00")
    MhRealInput16.Text = Format(Val(rstBookPOChild08.Fields("PaidAmount").Value), "0.00")
    Text6.Text = rstBookPOChild08.Fields("Remarks").Value
    TxtAdNar.Text = rstBookPOChild08.Fields("AdjustmentRemarks").Value
    CalculateTotalAmount
End Sub
Private Sub SaveFields()
    rstBookPOChild08.Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
    rstBookPOChild08.Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
    rstBookPOChild08.Fields("BindingType").Value = BindingTypeCode
    rstBookPOChild08.Fields("BindingForms").Value = Format(Val(MhRealInput7.Text), "0")
    rstBookPOChild08.Fields("ExtraForms").Value = Format(Val(MhRealInput18.Text), "0")
    rstBookPOChild08.Fields("ActualQuantity").Value = Format(Val(MhRealInput1.Text), "0")
    rstBookPOChild08.Fields("BillingQuantity").Value = Format(Val(MhRealInput3.Text), "0")
    rstBookPOChild08.Fields("AdjustQuantity").Value = Format(Val(MhRealInput4.Text), "0")
    rstBookPOChild08.Fields("FormFoldRate").Value = Format(Val(MhRealInput2.Text), "0.00")
    rstBookPOChild08.Fields("FormStitchRate").Value = Format(Val(MhRealInput5.Text), "0.00")
    rstBookPOChild08.Fields("FormPasteRate").Value = Format(Val(MhRealInput6.Text), "0.00")
    rstBookPOChild08.Fields("Rate/Book").Value = Format(Val(MhRealInput8.Text), "0.00")
    rstBookPOChild08.Fields("LooseQty/Box").Value = Format(Val(MhRealInput29.Text), "0")
    rstBookPOChild08.Fields("ExtraLooseQty").Value = Format(Val(MhRealInput30.Text), "0")
    rstBookPOChild08.Fields("TotalLooseQty").Value = Format(Val(MhRealInput31.Text), "0")
    rstBookPOChild08.Fields("Qty/Pkt").Value = Format(Val(MhRealInput23.Text), "0")
    rstBookPOChild08.Fields("TotalPkts").Value = Format(Val(MhRealInput24.Text), "0")
    rstBookPOChild08.Fields("Pkt/Box").Value = Format(Val(MhRealInput11.Text), "0")
    rstBookPOChild08.Fields("TotalBoxes").Value = Format(Val(MhRealInput12.Text), "0")
    rstBookPOChild08.Fields("PktPackRate").Value = Format(Val(MhRealInput25.Text), "0.00")
    rstBookPOChild08.Fields("BoxPackRate").Value = Format(Val(MhRealInput26.Text), "0.00")
    rstBookPOChild08.Fields("CartageRate").Value = Format(Val(MhRealInput13.Text), "0.00")
    rstBookPOChild08.Fields("Adjustment").Value = Format(Val(MhRealInput9.Text), "0.00")
    rstBookPOChild08.Fields("BillAmount").Value = Format(Val(MhRealInput10.Text), "0.00")
    rstBookPOChild08.Fields("BillNo").Value = Text8.Text
    If Not IsDate(MhDateInput2.Text) Then rstBookPOChild08.Fields("BillDate").Value = Null Else rstBookPOChild08.Fields("BillDate").Value = GetDate(MhDateInput2.Text)
    rstBookPOChild08.Fields("VAT%").Value = Format(Val(MhRealInput15.Text), "0.00")
    rstBookPOChild08.Fields("VAT").Value = Format(Val(MhRealInput17.Text), "0.00")
    rstBookPOChild08.Fields("PaidAmount").Value = Format(Val(MhRealInput16.Text), "0.00")
    rstBookPOChild08.Fields("Remarks").Value = Text6.Text
    rstBookPOChild08.Fields("AdjustmentRemarks").Value = IIf(Val(MhRealInput9.Text) <> 0, TxtAdNar.Text, "")
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild08.Fields("BillFeedDate").Value) Then rstBookPOChild08.Fields("BillFeedDate").Value = Now()
    Dim lpBuff As String * 1024
    GetComputerName lpBuff, Len(lpBuff)
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild08.Fields("ComputerName").Value) Then rstBookPOChild08.Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
End Sub

Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf Val(CheckNull(rstBookPOChild08.Fields("ActualQuantity").Value)) = 0 Then
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

Private Sub MhRealInput32_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub

Private Sub MhRealInput33_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub

Private Sub Text3_Change()
    If Text3.Text = " " Then
        Text3.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text3.Text)
    If rstBindingTypeList.RecordCount = 0 Then
        DisplayError ("No Record in Binding Type Master")
        Cancel = True
        Exit Sub
    Else
        rstBindingTypeList.MoveFirst
    End If
    rstBindingTypeList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstBindingTypeList.EOF Then
        SelectionType = "S"
        BindingTypeCode = ""
        Call LoadSelectionList(rstBindingTypeList, "List of Binding Types...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, BindingTypeCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(BindingTypeCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        BindingTypeCode = rstBindingTypeList.Fields("Code").Value
        Call GetBinderRates: CalculateTotalAmount
    End If
End Sub
Private Sub MhRealInput7_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    If Val(MhRealInput3.Text) = 0 Then MhRealInput3.Text = Format(Val(MhRealInput1.Text), "0"): Exit Sub
    If Val(MhRealInput3.Text) <> Val(MhRealInput1.Text) Then If MsgBox("Alter billing quantity?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Alter !") = vbYes Then MhRealInput3.Text = Format(Val(MhRealInput1.Text), "0")
End Sub
Private Sub MhRealInput3_Validate(Cancel As Boolean)
    CalculateBundle
    CalculateTotalAmount
End Sub
Private Sub MhRealInput29_Validate(Cancel As Boolean)   'Loose/Box
    CalculateBundle
End Sub
Private Sub MhRealInput30_Validate(Cancel As Boolean)   'Extra Loose
    CalculateBundle
End Sub
Private Sub MhRealInput23_Validate(Cancel As Boolean)   'Qty/Pkt
    CalculateBundle
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)   'Pkt/Box
    CalculateBundle
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput5_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput6_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput8_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput25_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput26_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput13_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput14_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput15_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub CalculateBundle()
    Dim CalcPkt As Long, CalcBox As Long, CalcLoose As Long, TotalBox As Long
    'Total box Calculation
    If Val(MhRealInput23.Text) * Val(MhRealInput11.Text) + Val(MhRealInput29.Text) > 0 Then TotalBox = Int((Val(MhRealInput3.Text) - Val(MhRealInput30.Text)) / (Val(MhRealInput23.Text) * Val(MhRealInput11.Text) + Val(MhRealInput29.Text)))   'int((billing quantity - extra loose quantity) / quantity per box)
    MhRealInput12.Text = Format(Round(TotalBox, 0), "0")
    'Total Loose Calculation
    If Val(MhRealInput23.Text) > 0 Then 'qty per packet
        CalcLoose = Val(MhRealInput30.Text) Mod Val(MhRealInput23.Text) 'Loose qty remaining from extra loose qty after packet making
        CalcLoose = CalcLoose + (Val(MhRealInput3.Text) - Val(MhRealInput30.Text) - (TotalBox * Val(MhRealInput29.Text))) Mod Val(MhRealInput23.Text)
    End If
    CalcLoose = CalcLoose + TotalBox * Val(MhRealInput29.Text)
    MhRealInput31.Text = Format(CalcLoose, "0")
    'Total Packet Calculation
    If Val(MhRealInput23.Text) > 0 Then 'qty per packet
        CalcPkt = Int(Val(MhRealInput30.Text) / Val(MhRealInput23.Text))
        CalcPkt = CalcPkt + Int((Val(MhRealInput3.Text) - Val(MhRealInput30.Text) - (TotalBox * Val(MhRealInput29.Text))) / Val(MhRealInput23.Text))
    End If
    MhRealInput24.Text = Format(Round(CalcPkt, 0), "0")
    CalculateTotalAmount
End Sub
Private Sub CalculateTotalAmount()
    MhRealInput20.Value = (MhRealInput2.Value * MhRealInput3.Value * (MhRealInput7.Value + MhRealInput18.Value)) / 1000 'Folding Amount
    MhRealInput21.Value = (MhRealInput5.Value * MhRealInput3.Value * (MhRealInput7.Value + MhRealInput18.Value)) / 1000 'Stitching Amount
    MhRealInput22.Value = (MhRealInput6.Value * MhRealInput3.Value) / 1000  'Pasting Amount
    MhRealInput27.Value = MhRealInput24.Value * MhRealInput25.Value 'Pkt Packing Amount
    MhRealInput28.Value = MhRealInput12.Value * MhRealInput26.Value 'Box Packing Amount
    MhRealInput14.Value = MhRealInput12.Value * MhRealInput13.Value 'Cartage
    MhRealInput32.Value = (MhRealInput8.Value * MhRealInput3.Value) 'Binding per Item Rtae Amount
    MhRealInput17.Value = (MhRealInput20.Value + MhRealInput21.Value + MhRealInput22.Value + MhRealInput27.Value + MhRealInput28.Value + MhRealInput9.Value + MhRealInput14.Value + (MhRealInput8.Value * MhRealInput3.Value)) * MhRealInput15.Value / 100  'GST
    MhRealInput33.Value = Round(MhRealInput20.Value + MhRealInput21.Value + MhRealInput22.Value + MhRealInput27.Value + MhRealInput28.Value + (MhRealInput8.Value * MhRealInput3.Value) + MhRealInput14.Value + MhRealInput9.Value, 0) 'Total Taxable Amount
    MhRealInput10.Value = Round(MhRealInput20.Value + MhRealInput21.Value + MhRealInput22.Value + MhRealInput27.Value + MhRealInput28.Value + (MhRealInput8.Value * MhRealInput3.Value) + MhRealInput14.Value + MhRealInput17.Value + MhRealInput9.Value, 0) 'Total Amount
End Sub
Private Sub cmdProceed_Click()
    If CheckMandatoryFields Then Exit Sub
    SaveFields
    rstBookPOChild08.Update
    Call CloseForm(Me)
End Sub
Private Sub cmdCancel_Click()
    rstBookPOChild08.CancelUpdate
    Call CloseForm(Me)
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text3.Text, False) Then Text3.SetFocus: CheckMandatoryFields = True: Exit Function
    If Not CheckExists(Text3, "Col0", rstBindingTypeList, BindingTypeCode) Then Text3.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput16.Text) <> 0 Then If Val(MhRealInput16.Text) <> Val(MhRealInput10.Text) Then MhRealInput9.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput9.Text) <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
End Function
Private Sub GetBinderRates()
    Dim FoldingRate As Double, StitchingRate As Double, PastingRate As Double, RPB As Double, PktPackRate As Double, BoxPackRate As Double, CartagePerBox As Double
    On Error GoTo ErrorHandler
    If rstBinderRates.State = adStateOpen Then rstBinderRates.Close
    rstBinderRates.Open "Select Top 1 * From AccountChild08 Where Code = '" & BinderCode & "' And [Size]=(SELECT Code FROM SizeGroupChild WHERE [Size]='" & SizeCode & "') And BindingType = '" & BindingTypeCode & "' And Range" & Choose(Val(FormType), "08", "16", "04", "12", "24", "32", "64", "06", "02") & " >= " & Val(MhRealInput7.Text) + Val(MhRealInput18.Text) & " Order By Range" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64")))))), cnDatabase, adOpenKeyset, adLockReadOnly
    If rstBinderRates.RecordCount = 0 Then
        If rstBinderRates.State = adStateOpen Then rstBinderRates.Close
        rstBinderRates.Open "Select Top 1 * From AccountMaster,AccountChild08 Where AccountMaster.Code = AccountChild08.Code And [Name] Like '%Rate%' And [Size]=(SELECT Code FROM SizeGroupChild WHERE [Size]='" & SizeCode & "') And BindingType = '" & BindingTypeCode & "' And Range" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64")))))) & " >= " & Val(MhRealInput7.Text) + Val(MhRealInput18.Text) & " Order By Range" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64")))))), cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstBinderRates.RecordCount > 0 Then
        FoldingRate = rstBinderRates.Fields("FormFoldRate" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64"))))))).Value
        StitchingRate = rstBinderRates.Fields("FormStitchRate" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64"))))))).Value
        PastingRate = rstBinderRates.Fields("FormPasteRate" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64"))))))).Value
        If Val(MhRealInput7.Text) + Val(MhRealInput18.Text) > 25 Then PastingRate = PastingRate + IIf(PastingRate > 0, (Val(MhRealInput7.Text) + Val(MhRealInput18.Text) - 25) * Val(FrmBookPrintOrder.rstBookList.Fields("AddOnRate02").Value) * 1000, 0)
        RPB = rstBinderRates.Fields("Rate/Book" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64"))))))).Value
        PktPackRate = rstBinderRates.Fields("PktPackRate").Value
        BoxPackRate = rstBinderRates.Fields("BoxPackRate").Value
        CartagePerBox = rstBinderRates.Fields("Cartage/Box").Value
    End If
    If Val(MhRealInput2.Text) <> FoldingRate And Val(MhRealInput2.Text) <> 0 Then
        If MsgBox("Folding Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput2.Text = Format(FoldingRate, "0.00")
    Else
        MhRealInput2.Text = Format(FoldingRate, "0.00")
    End If
    If Val(MhRealInput5.Text) <> StitchingRate And Val(MhRealInput5.Text) <> 0 Then
        If MsgBox("Stitching Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput5.Text = Format(StitchingRate, "0.00")
    Else
        MhRealInput5.Text = Format(StitchingRate, "0.00")
    End If
    If Val(MhRealInput6.Text) <> PastingRate And Val(MhRealInput6.Text) <> 0 Then
        If MsgBox("Pasting Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput6.Text = Format(PastingRate, "0.00")
    Else
        MhRealInput6.Text = Format(PastingRate, "0.00")
    End If
    If Val(MhRealInput8.Text) <> RPB And Val(MhRealInput8.Text) <> 0 Then
        If MsgBox("Rate/Book is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput8.Text = Format(RPB, "0.00")
    Else
        MhRealInput8.Text = Format(RPB, "0.00")
    End If
    If Val(MhRealInput25.Text) <> PktPackRate And Val(MhRealInput25.Text) <> 0 Then
        If MsgBox("Pkt Packing Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput25.Text = Format(PktPackRate, "0.00")
    Else
        MhRealInput25.Text = Format(PktPackRate, "0.00")
    End If
    If Val(MhRealInput26.Text) <> BoxPackRate And Val(MhRealInput26.Text) <> 0 Then
        If MsgBox("Box Packing Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput26.Text = Format(BoxPackRate, "0.00")
    Else
        MhRealInput26.Text = Format(BoxPackRate, "0.00")
    End If
    If Val(MhRealInput13.Text) <> CartagePerBox And Val(MhRealInput13.Text) <> 0 Then
        If MsgBox("Cartage/Box is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput13.Text = Format(CartagePerBox, "0.00")
    Else
        MhRealInput13.Value = CartagePerBox
    End If
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Fetch Binder Rates")
End Sub
