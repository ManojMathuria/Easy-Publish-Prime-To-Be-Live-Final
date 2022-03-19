VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmBookPOChild04 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plate Making Order Details"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
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
   ScaleHeight     =   7305
   ScaleWidth      =   8760
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   8295
      Picture         =   "BookPOChild04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   8295
      Picture         =   "BookPOChild04.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   7100
      Left            =   120
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   105
      Width           =   8055
      _Version        =   65536
      _ExtentX        =   14208
      _ExtentY        =   12524
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
      Picture         =   "BookPOChild04.frx":0204
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
         Left            =   4320
         MaxLength       =   40
         TabIndex        =   2
         Top             =   645
         Width           =   1095
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
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   32
         Top             =   6650
         Width           =   6615
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
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   31
         Top             =   6325
         Width           =   6615
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   61
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
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   28
         Top             =   5790
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
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   24
         Top             =   3390
         Width           =   4095
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   645
         Width           =   1575
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   2880
         TabIndex        =   38
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Order Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0220
         Picture         =   "BookPOChild04.frx":023C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   2880
         TabIndex        =   39
         Top             =   960
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Actual Quantity"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0258
         Picture         =   "BookPOChild04.frx":0274
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   645
         Left            =   5400
         TabIndex        =   40
         Top             =   960
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   1138
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
         Caption         =   " Billing Quantity"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0290
         Picture         =   "BookPOChild04.frx":02AC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   41
         Top             =   1280
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
         Caption         =   " Printing Type"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":02C8
         Picture         =   "BookPOChild04.frx":02E4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   42
         Top             =   1910
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
         Caption         =   " Total Plates"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0300
         Picture         =   "BookPOChild04.frx":031C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   2880
         TabIndex        =   43
         Top             =   1910
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Plate Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0338
         Picture         =   "BookPOChild04.frx":0354
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   2880
         TabIndex        =   44
         Top             =   2220
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Print Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0370
         Picture         =   "BookPOChild04.frx":038C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   645
         Left            =   2880
         TabIndex        =   45
         Top             =   2540
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   1138
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
         Caption         =   " Adjustment"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":03A8
         Picture         =   "BookPOChild04.frx":03C4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   120
         TabIndex        =   46
         Top             =   1590
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
         Caption         =   " Pages/Forms"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":03E0
         Picture         =   "BookPOChild04.frx":03FC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   47
         Top             =   2220
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
         Caption         =   " Total Forms"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0418
         Picture         =   "BookPOChild04.frx":0434
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   5400
         TabIndex        =   48
         Top             =   1910
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Plate Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0450
         Picture         =   "BookPOChild04.frx":046C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   5400
         TabIndex        =   49
         Top             =   2220
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Print Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0488
         Picture         =   "BookPOChild04.frx":04A4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   645
         Left            =   5400
         TabIndex        =   50
         Top             =   2540
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   1138
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
         Caption         =   " Total Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":04C0
         Picture         =   "BookPOChild04.frx":04DC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   120
         TabIndex        =   51
         Top             =   3390
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
         Caption         =   " Paper Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":04F8
         Picture         =   "BookPOChild04.frx":0514
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Left            =   120
         TabIndex        =   52
         Top             =   3710
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
         Caption         =   " Wastage (%)"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0530
         Picture         =   "BookPOChild04.frx":054C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Left            =   2880
         TabIndex        =   53
         Top             =   3710
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Consumption"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0568
         Picture         =   "BookPOChild04.frx":0584
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
         Height          =   330
         Left            =   5400
         TabIndex        =   54
         Top             =   3710
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Total Consmptn"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":05A0
         Picture         =   "BookPOChild04.frx":05BC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   55
         Top             =   5790
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   32896
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
         Caption         =   " Bill No."
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":05D8
         Picture         =   "BookPOChild04.frx":05F4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   5400
         TabIndex        =   56
         Top             =   5790
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Paid Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0610
         Picture         =   "BookPOChild04.frx":062C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   2880
         TabIndex        =   57
         Top             =   5790
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Bill Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0648
         Picture         =   "BookPOChild04.frx":0664
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   5400
         TabIndex        =   58
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Target Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0680
         Picture         =   "BookPOChild04.frx":069C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   59
         Top             =   645
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
         Caption         =   " Book Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":06B8
         Picture         =   "BookPOChild04.frx":06D4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Left            =   2880
         TabIndex        =   60
         Top             =   1590
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Plate Type"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":06F0
         Picture         =   "BookPOChild04.frx":070C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   62
         Top             =   105
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
         Caption         =   " Order No."
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0728
         Picture         =   "BookPOChild04.frx":0744
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   63
         Top             =   6325
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
         Caption         =   " Remarks"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0760
         Picture         =   "BookPOChild04.frx":077C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
         Height          =   330
         Left            =   120
         TabIndex        =   64
         Top             =   2535
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
         Caption         =   " GST (Printing)"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0798
         Picture         =   "BookPOChild04.frx":07B4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel29 
         Height          =   330
         Left            =   5400
         TabIndex        =   65
         Top             =   645
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Plate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":07D0
         Picture         =   "BookPOChild04.frx":07EC
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   4320
         TabIndex        =   0
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild04.frx":0808
         Caption         =   "BookPOChild04.frx":0920
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":098C
         Keys            =   "BookPOChild04.frx":09AA
         Spin            =   "BookPOChild04.frx":0A08
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
         Left            =   6840
         TabIndex        =   1
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild04.frx":0A30
         Caption         =   "BookPOChild04.frx":0B48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":0BB4
         Keys            =   "BookPOChild04.frx":0BD2
         Spin            =   "BookPOChild04.frx":0C30
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
         Left            =   4320
         TabIndex        =   29
         Top             =   5790
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild04.frx":0C58
         Caption         =   "BookPOChild04.frx":0D70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":0DDC
         Keys            =   "BookPOChild04.frx":0DFA
         Spin            =   "BookPOChild04.frx":0E58
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
         Left            =   5400
         TabIndex        =   66
         Top             =   1590
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Forms/Sheet"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0E80
         Picture         =   "BookPOChild04.frx":0E9C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel31 
         Height          =   330
         Left            =   5400
         TabIndex        =   67
         Top             =   3390
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Forms/Sheet"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":0EB8
         Picture         =   "BookPOChild04.frx":0ED4
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   4320
         TabIndex        =   4
         Top             =   960
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":0EF0
         Caption         =   "BookPOChild04.frx":0F10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":0F7C
         Keys            =   "BookPOChild04.frx":0F9A
         Spin            =   "BookPOChild04.frx":0FE4
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   6840
         TabIndex        =   5
         ToolTipText     =   "One Color"
         Top             =   960
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":100C
         Caption         =   "BookPOChild04.frx":102C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":1098
         Keys            =   "BookPOChild04.frx":10B6
         Spin            =   "BookPOChild04.frx":1100
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
         Height          =   330
         Left            =   6840
         TabIndex        =   6
         ToolTipText     =   "Double & Four Color"
         Top             =   1280
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":1128
         Caption         =   "BookPOChild04.frx":1148
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":11B4
         Keys            =   "BookPOChild04.frx":11D2
         Spin            =   "BookPOChild04.frx":121C
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         ToolTipText     =   "Two & Four Color"
         Top             =   1590
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":1244
         Caption         =   "BookPOChild04.frx":1264
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":12D0
         Keys            =   "BookPOChild04.frx":12EE
         Spin            =   "BookPOChild04.frx":1338
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
         Left            =   1800
         TabIndex        =   8
         ToolTipText     =   "¼ Form"
         Top             =   1590
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":1360
         Caption         =   "BookPOChild04.frx":1380
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":13EC
         Keys            =   "BookPOChild04.frx":140A
         Spin            =   "BookPOChild04.frx":1454
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
         Left            =   2160
         TabIndex        =   9
         ToolTipText     =   "½ Form"
         Top             =   1590
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":147C
         Caption         =   "BookPOChild04.frx":149C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":1508
         Keys            =   "BookPOChild04.frx":1526
         Spin            =   "BookPOChild04.frx":1570
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
         Height          =   330
         Left            =   2520
         TabIndex        =   10
         ToolTipText     =   "1 Form"
         Top             =   1590
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":1598
         Caption         =   "BookPOChild04.frx":15B8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":1624
         Keys            =   "BookPOChild04.frx":1642
         Spin            =   "BookPOChild04.frx":168C
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
         Height          =   330
         Left            =   6840
         TabIndex        =   12
         Top             =   1590
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":16B4
         Caption         =   "BookPOChild04.frx":16D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":1740
         Keys            =   "BookPOChild04.frx":175E
         Spin            =   "BookPOChild04.frx":17A8
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
         ValueVT         =   1245189
         Value           =   1
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   1320
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "¼ Form"
         Top             =   1910
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":17D0
         Caption         =   "BookPOChild04.frx":17F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":185C
         Keys            =   "BookPOChild04.frx":187A
         Spin            =   "BookPOChild04.frx":18C4
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput23 
         Height          =   330
         Left            =   1845
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "½ Form"
         Top             =   1910
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":18EC
         Caption         =   "BookPOChild04.frx":190C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":1978
         Keys            =   "BookPOChild04.frx":1996
         Spin            =   "BookPOChild04.frx":19E0
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
         Height          =   330
         Left            =   2370
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "1 Form"
         Top             =   1910
         Width           =   525
         _Version        =   65536
         _ExtentX        =   917
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":1A08
         Caption         =   "BookPOChild04.frx":1A28
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":1A94
         Keys            =   "BookPOChild04.frx":1AB2
         Spin            =   "BookPOChild04.frx":1AFC
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   4320
         TabIndex        =   16
         Top             =   1910
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":1B24
         Caption         =   "BookPOChild04.frx":1B44
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":1BB0
         Keys            =   "BookPOChild04.frx":1BCE
         Spin            =   "BookPOChild04.frx":1C18
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
         ValueVT         =   1966407685
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   6840
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1910
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":1C40
         Caption         =   "BookPOChild04.frx":1C60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":1CCC
         Keys            =   "BookPOChild04.frx":1CEA
         Spin            =   "BookPOChild04.frx":1D34
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   1320
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "¼ Form"
         Top             =   2220
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":1D5C
         Caption         =   "BookPOChild04.frx":1D7C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":1DE8
         Keys            =   "BookPOChild04.frx":1E06
         Spin            =   "BookPOChild04.frx":1E50
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
         Height          =   330
         Left            =   1845
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "½ Form"
         Top             =   2220
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":1E78
         Caption         =   "BookPOChild04.frx":1E98
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":1F04
         Keys            =   "BookPOChild04.frx":1F22
         Spin            =   "BookPOChild04.frx":1F6C
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
         Height          =   330
         Left            =   2370
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "1 Form"
         Top             =   2220
         Width           =   525
         _Version        =   65536
         _ExtentX        =   917
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":1F94
         Caption         =   "BookPOChild04.frx":1FB4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":2020
         Keys            =   "BookPOChild04.frx":203E
         Spin            =   "BookPOChild04.frx":2088
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   4320
         TabIndex        =   20
         Top             =   2220
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":20B0
         Caption         =   "BookPOChild04.frx":20D0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":213C
         Keys            =   "BookPOChild04.frx":215A
         Spin            =   "BookPOChild04.frx":21A4
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
         Left            =   6840
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":21CC
         Caption         =   "BookPOChild04.frx":21EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":2258
         Keys            =   "BookPOChild04.frx":2276
         Spin            =   "BookPOChild04.frx":22C0
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
         Height          =   330
         Left            =   1320
         TabIndex        =   21
         Top             =   2540
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":22E8
         Caption         =   "BookPOChild04.frx":2308
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":2374
         Keys            =   "BookPOChild04.frx":2392
         Spin            =   "BookPOChild04.frx":23DC
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
         Left            =   2120
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2540
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":2404
         Caption         =   "BookPOChild04.frx":2424
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":2490
         Keys            =   "BookPOChild04.frx":24AE
         Spin            =   "BookPOChild04.frx":24F8
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   645
         Left            =   4320
         TabIndex        =   23
         Top             =   2535
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1147
         Calculator      =   "BookPOChild04.frx":2520
         Caption         =   "BookPOChild04.frx":2540
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":25AC
         Keys            =   "BookPOChild04.frx":25CA
         Spin            =   "BookPOChild04.frx":2614
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
         Height          =   645
         Left            =   6840
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2540
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1138
         Calculator      =   "BookPOChild04.frx":263C
         Caption         =   "BookPOChild04.frx":265C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":26C8
         Keys            =   "BookPOChild04.frx":26E6
         Spin            =   "BookPOChild04.frx":2730
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
         Height          =   330
         Left            =   6840
         TabIndex        =   25
         Top             =   3390
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":2758
         Caption         =   "BookPOChild04.frx":2778
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":27E4
         Keys            =   "BookPOChild04.frx":2802
         Spin            =   "BookPOChild04.frx":284C
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   1320
         TabIndex        =   26
         Top             =   3710
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":2874
         Caption         =   "BookPOChild04.frx":2894
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":2900
         Keys            =   "BookPOChild04.frx":291E
         Spin            =   "BookPOChild04.frx":2968
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
         Height          =   330
         Left            =   4320
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3710
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":2990
         Caption         =   "BookPOChild04.frx":29B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":2A1C
         Keys            =   "BookPOChild04.frx":2A3A
         Spin            =   "BookPOChild04.frx":2A84
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   6840
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   3710
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":2AAC
         Caption         =   "BookPOChild04.frx":2ACC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":2B38
         Keys            =   "BookPOChild04.frx":2B56
         Spin            =   "BookPOChild04.frx":2BA0
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
         Height          =   330
         Left            =   6840
         TabIndex        =   30
         Top             =   5790
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":2BC8
         Caption         =   "BookPOChild04.frx":2BE8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":2C54
         Keys            =   "BookPOChild04.frx":2C72
         Spin            =   "BookPOChild04.frx":2CBC
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   1335
         Left            =   120
         TabIndex        =   27
         Top             =   4245
         Width           =   7815
         _Version        =   524288
         _ExtentX        =   13785
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
         MaxCols         =   23
         MaxRows         =   3
         OperationMode   =   2
         SpreadDesigner  =   "BookPOChild04.frx":2CE4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Left            =   120
         TabIndex        =   74
         Top             =   6650
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
         Caption         =   " Adj.Remarks"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":3A6A
         Picture         =   "BookPOChild04.frx":3A86
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   75
         Top             =   2855
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
         Caption         =   " GST (Plate)"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":3AA2
         Picture         =   "BookPOChild04.frx":3ABE
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput28 
         Height          =   330
         Left            =   1320
         TabIndex        =   22
         Top             =   2855
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":3ADA
         Caption         =   "BookPOChild04.frx":3AFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":3B66
         Keys            =   "BookPOChild04.frx":3B84
         Spin            =   "BookPOChild04.frx":3BCE
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput29 
         Height          =   330
         Left            =   2115
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   2855
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   582
         Calculator      =   "BookPOChild04.frx":3BF6
         Caption         =   "BookPOChild04.frx":3C16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild04.frx":3C82
         Keys            =   "BookPOChild04.frx":3CA0
         Spin            =   "BookPOChild04.frx":3CEA
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel33 
         Height          =   330
         Left            =   2880
         TabIndex        =   77
         Top             =   645
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Size"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild04.frx":3D12
         Picture         =   "BookPOChild04.frx":3D2E
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   330
         Left            =   6840
         TabIndex        =   3
         Top             =   645
         Width           =   1095
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1931;582"
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
         X1              =   0
         X2              =   8055
         Y1              =   5685
         Y2              =   5685
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   8055
         Y1              =   6215
         Y2              =   6215
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   8055
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8055
         Y1              =   4140
         Y2              =   4140
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   4320
         TabIndex        =   11
         Top             =   1590
         Width           =   1095
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1931;582"
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
         X2              =   8055
         Y1              =   3290
         Y2              =   3290
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Left            =   1320
         TabIndex        =   35
         Top             =   1280
         Width           =   4095
         VariousPropertyBits=   545282073
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "7223;582"
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
Attribute VB_Name = "FrmBookPOChild04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstBookPOChild05 As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim rstSizeList As New ADODB.Recordset
Dim rstPrinterRates As New ADODB.Recordset
Public PrinterCode As String
Dim BookCode As String
Dim SizeCode As String
Dim PaperCode As String
Private Sub Form_Load()
    Dim Cnt As Integer, Pages As Variant
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    DisableCloseButton Me
    For Cnt = 11 To 24
        fpSpread1.Col = Cnt
        fpSpread1.ColHidden = True
    Next
    AbortPO = False
    BookCode = FrmBookPrintOrder.rstBookList.Fields("Code").Value
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text)
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text)
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
    rstPaperList.Open "Select Name As Col0, Code From PaperMaster Where PaperMaster.Type = '1' Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstSizeList.Open "SELECT Name As Col0, Code From GeneralMaster Where Type = '1' ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperList.ActiveConnection = Nothing
    rstSizeList.ActiveConnection = Nothing
    If IsNull(rstBookPOChild05.Fields("Code").Value) Then MhRealInput5.ReadOnly = True Else MhRealInput5.ReadOnly = False
    If Val(CheckNull(rstBookPOChild05.Fields("ActualQuantity").Value)) = 0 Then
        For Cnt = 1 To fpSpread1.MaxRows
            fpSpread1.SetText 1, Cnt, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPages").Value)
            fpSpread1.SetText 2, Cnt, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorForms").Value)
            fpSpread1.SetText 3, Cnt, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color¼Forms").Value)
            fpSpread1.SetText 4, Cnt, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color½Forms").Value)
            fpSpread1.SetText 5, Cnt, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color1F/BForms").Value) + Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color1W/TForms").Value)
            fpSpread1.SetText 6, Cnt, IIf(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "1", "Deepatch", IIf(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "2", "PS", IIf(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "3", "Wipeon", "CTP")))
            fpSpread1.SetText 7, Cnt, 0#
            fpSpread1.SetText 8, Cnt, 0#
            fpSpread1.SetText 9, Cnt, 0#
            fpSpread1.SetText 10, Cnt, 0#
        Next
        MhDateInput1.Text = Format(GetDate(FrmBookPrintOrder.MhDateInput1.Text), "dd-MM-yyyy")
        MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        SizeCode = FrmBookPrintOrder.rstBookList.Fields("SizeCode").Value
        If rstSizeList.RecordCount > 0 Then rstSizeList.MoveFirst
        rstSizeList.Find "[Code] = '" & SizeCode & "'"
        If Not rstSizeList.EOF Then Text4.Text = rstSizeList.Fields("Col0").Value: fpSpread1.SetText 23, fpSpread1.ActiveRow, SizeCode
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
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstPrinterRates)
End Sub
Private Sub ClearFields()
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy") 'Order Date
    MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy") 'Target Date
    Combo3.ListIndex = 0                'Processing
    Text4.Text = ""
    MhRealInput1.Value = 0              'Actual Quantity
    Combo1.ListIndex = 0                'Printing Type
    MhRealInput15.Value = 0             'Pages
    MhRealInput17.Value = 0             'Qtr Form
    MhRealInput20.Value = 0             'Half Form
    MhRealInput21.Value = 0             'Full Form
    Combo2.ListIndex = 0                'Plate Type
    MhRealInput22.Value = 1#            'Forms/Sheet For Printing Purpose
    MhRealInput3.Value = 0              'Total Plates-¼F
    MhRealInput23.Value = 0             'Total Plates-½F
    MhRealInput24.Value = 0             'Total Plates-1F
    MhRealInput4.Value = 0#             'Plate Rate
    MhRealInput7.Value = 0#             'Plate Amount
    MhRealInput6.Value = 0              'Total Forms-¼F
    MhRealInput25.Value = 0             'Total Forms-½F
    MhRealInput26.Value = 0             'Total Forms-1F
    MhRealInput5.Value = 0#             'Print Rate
    MhRealInput8.Value = 0#             'Print Amount
    MhRealInput14.Value = 0#            'GST %
    MhRealInput18.Value = 0#            'GST Amount
    MhRealInput9.Value = 0#             'Adjustment
    MhRealInput10.Value = 0#            'Total Amount
    Text1.Text = ""                     'Paper Name
    MhRealInput27.Value = 1#            'Forms/Sheet For Paper Purpose
    MhRealInput11.Value = 0#            'Paper Wastage (in %)
    MhRealInput12.Value = 0#            'Paper Consumption
    MhRealInput13.Value = 0#            'Total Paper Consumption
    Text8.Text = ""                     'Bill No.
    MhDateInput2.Text = "  -  -    "    'Bill Date
    MhRealInput16.Value = 0#            'Bill Amount
    Text6.Text = ""                     'Remarks
    MhRealInput28.Value = 0
    MhRealInput29.Value = 0
    TxtAdNar.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
End Sub
Private Sub LoadFields()
    Dim Cnt As Integer
    If rstBookPOChild05.RecordCount = 0 Then Exit Sub
    MhDateInput1.Text = Format(rstBookPOChild05.Fields("OrderDate").Value, "dd-MM-yyyy")
    MhDateInput3.Text = Format(rstBookPOChild05.Fields("TargetDate").Value, "dd-MM-yyyy")
    Combo3.ListIndex = IIf(rstBookPOChild05.Fields("Processing").Value = "O", 0, IIf(rstBookPOChild05.Fields("Processing").Value = "N", 1, 2))
    MhRealInput1.Text = Format(Val(rstBookPOChild05.Fields("ActualQuantity").Value), "0")
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
        If Not rstPaperList.EOF Then fpSpread1.SetText 19, Cnt, rstPaperList.Fields("Col0").Value
        fpSpread1.SetText 20, Cnt, rstBookPOChild05.Fields("Paper" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value
        fpSpread1.SetText 21, Cnt, Val(rstBookPOChild05.Fields("Forms/Sheet1-" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 22, Cnt, Val(rstBookPOChild05.Fields("Forms/Sheet2-" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 23, Cnt, rstBookPOChild05.Fields("Size" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value
    Next
    MhRealInput14.Text = Format(Val(rstBookPOChild05.Fields("VAT%").Value), "0.00")
    MhRealInput18.Text = Format(Val(rstBookPOChild05.Fields("VAT").Value), "0.00")
    MhRealInput28.Value = rstBookPOChild05.Fields("PVAT%").Value
    MhRealInput29.Value = rstBookPOChild05.Fields("PVAT").Value
    MhRealInput9.Text = Format(Val(rstBookPOChild05.Fields("Adjustment").Value), "0.00")
    MhRealInput10.Text = Format(Val(rstBookPOChild05.Fields("BillAmount").Value), "0.00")
    MhRealInput13.Text = Format(Val(rstBookPOChild05.Fields("TotalPaperConsumption").Value), "0.000")
    Text8.Text = rstBookPOChild05.Fields("BillNo").Value
    If Not IsNull(rstBookPOChild05.Fields("BillDate").Value) Then MhDateInput2.Text = Format(rstBookPOChild05.Fields("BillDate").Value, "dd-MM-yyyy")
    MhRealInput16.Text = Format(Val(rstBookPOChild05.Fields("PaidAmount").Value), "0.00")
    Text6.Text = rstBookPOChild05.Fields("Remarks").Value
    TxtAdNar.Text = rstBookPOChild05.Fields("AdjustmentRemarks").Value
End Sub
Private Sub SaveFields()
    Dim Cnt As Integer, Content As Variant
    rstBookPOChild05.Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
    rstBookPOChild05.Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
    rstBookPOChild05.Fields("Processing").Value = IIf(Combo3.ListIndex = 0, "O", IIf(Combo3.ListIndex = 1, "N", "R"))
    rstBookPOChild05.Fields("ActualQuantity").Value = Val(MhRealInput1.Text)
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
        fpSpread1.GetText 10, Cnt, Content
        rstBookPOChild05.Fields("PaperConsumptionOther" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        rstBookPOChild05.Fields("PaperConsumptionSheets" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Int(Val(Content)) * 500 + (Val(Content) - Int(Val(Content))) * 1000
        fpSpread1.GetText 11, Cnt, Content
        rstBookPOChild05.Fields("TotalPlates" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value = Val(Content)
        fpSpread1.GetText 12, Cnt, Content
        rstBookPOChild05.Fields("TotalPlates" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value = Val(Content)
        fpSpread1.GetText 13, Cnt, Content
        rstBookPOChild05.Fields("TotalPlates" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value = Val(Content)
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
    Next
    rstBookPOChild05.Fields("VAT%").Value = Format(Val(MhRealInput14.Text), "0.00")
    rstBookPOChild05.Fields("VAT").Value = Format(Val(MhRealInput18.Text), "0.00")
    rstBookPOChild05.Fields("PVAT%").Value = MhRealInput28.Value
    rstBookPOChild05.Fields("PVAT").Value = MhRealInput29.Value
    rstBookPOChild05.Fields("Adjustment").Value = Format(Val(MhRealInput9.Text), "0.00")
    rstBookPOChild05.Fields("BillAmount").Value = Format(Val(MhRealInput10.Text), "0.00")
    rstBookPOChild05.Fields("TotalPaperConsumption").Value = Format(Val(MhRealInput13.Text), "0.000")
    rstBookPOChild05.Fields("BillNo").Value = Text8.Text
    If Not IsDate(MhDateInput2.Text) Then
         rstBookPOChild05.Fields("BillDate").Value = Null
    Else
         rstBookPOChild05.Fields("BillDate").Value = GetDate(MhDateInput2.Text)
    End If
    rstBookPOChild05.Fields("PaidAmount").Value = Format(Val(MhRealInput16.Text), "0.00")
    rstBookPOChild05.Fields("Remarks").Value = Text6.Text
    rstBookPOChild05.Fields("AdjustmentRemarks").Value = IIf(Val(MhRealInput9.Text) <> 0, TxtAdNar.Text, "")
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
Private Sub Text4_Change()
    If Text4.Text = " " Then
        Text4.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text4.Text)
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
        Call DisplaySelectionList(Text4, SizeCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text4.Text, False) Then
            Text4.Text = "?"
        End If
        If RTrim(SizeCode) <> "" Then SendKeys "{TAB}"
        Cancel = True
    Else
        SizeCode = rstSizeList.Fields("Code").Value
        fpSpread1.SetText 23, fpSpread1.ActiveRow, SizeCode
    End If
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    If Val(MhRealInput1.Text) <= 0 Then
        Cancel = True
    Else
        CalculateAQD
    End If
End Sub
Private Sub MhRealInput15_Validate(Cancel As Boolean)
    fpSpread1.SetText 1, fpSpread1.ActiveRow, Val(MhRealInput15.Text)
    fpSpread1.SetText 2, fpSpread1.ActiveRow, Val(MhRealInput15.Text) / IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "1", "08", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "2", "16", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "3", "04", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "4", "12", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "5", "24", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "6", "32", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "7", "64", "06")))))))
End Sub
Private Sub MhRealInput17_Validate(Cancel As Boolean)   '¼ Forms
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant

    fpSpread1.SetText 3, fpSpread1.ActiveRow, Val(MhRealInput17.Text)
    Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput17.Text), "¼")
    Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput17.Text), "¼")
    CalculateAmount
    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
    fpSpread1.GetText 3, fpSpread1.ActiveRow, Forms¼
    fpSpread1.GetText 4, fpSpread1.ActiveRow, Forms½
    fpSpread1.GetText 5, fpSpread1.ActiveRow, Forms1
    fpSpread1.SetText 2, fpSpread1.ActiveRow, Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1
End Sub
Private Sub MhRealInput20_Validate(Cancel As Boolean)   '½ Forms
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant

    fpSpread1.SetText 4, fpSpread1.ActiveRow, Val(MhRealInput20.Text)
    Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput20.Text), "½")
    Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput20.Text), "½")
    CalculateAmount
    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
    fpSpread1.GetText 3, fpSpread1.ActiveRow, Forms¼
    fpSpread1.GetText 4, fpSpread1.ActiveRow, Forms½
    fpSpread1.GetText 5, fpSpread1.ActiveRow, Forms1
    fpSpread1.SetText 2, fpSpread1.ActiveRow, Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1
End Sub
Private Sub MhRealInput21_Validate(Cancel As Boolean)   '1 Forms
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant

    fpSpread1.SetText 5, fpSpread1.ActiveRow, Val(MhRealInput21.Text)
    Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput21.Text), "1")
    Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput21.Text), "1")
    CalculateAmount
    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
    fpSpread1.GetText 3, fpSpread1.ActiveRow, Forms¼
    fpSpread1.GetText 4, fpSpread1.ActiveRow, Forms½
    fpSpread1.GetText 5, fpSpread1.ActiveRow, Forms1
    fpSpread1.SetText 2, fpSpread1.ActiveRow, Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)  'Plate Type
    fpSpread1.SetText 6, fpSpread1.ActiveRow, IIf(Combo2.ListIndex = 0, "Deepatch", IIf(Combo2.ListIndex = 1, "PS", IIf(Combo2.ListIndex = 2, "Wipeon", "CTP")))
    GetPrinterRates IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), "L"  'Get Plate Rates
    CalculateAmount
    If Combo2.ListIndex = 1 Or Combo2.ListIndex = 3 Then    'PS/CTP Plate Details
        On Error Resume Next
        FrmPSPlateRegister.BookCode = BookCode
        FrmPSPlateRegister.BookName = Trim(Text2.Text)
        FrmPSPlateRegister.OrderCode = IIf(CheckNull(rstBookPOChild05.Fields("Code").Value) = "", "999999", rstBookPOChild05.Fields("Code").Value)
        FrmPSPlateRegister.OrderDate = GetDate(MhDateInput1.Text)
        FrmPSPlateRegister.OrderType = "05"
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
Private Sub MhRealInput9_Validate(Cancel As Boolean)    'Adjustment
    CalculateTotalAmount
End Sub
Private Sub Text1_Change()  'Paper
    If Text1.Text = " " Then
        Text1.Text = "?"
        SendKeys "{TAB}"
    ElseIf CheckEmpty(Text1, False) Then
        PaperCode = ""
        fpSpread1.SetText 19, fpSpread1.ActiveRow, ""
        fpSpread1.SetText 20, fpSpread1.ActiveRow, ""
    End If
End Sub
Private Sub Text1_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    If Val(MhRealInput8.Text) = 0 Then
        If CheckEmpty(Text1, False) Then Exit Sub
    End If
    SearchString = FixQuote(Text1.Text)
    If rstPaperList.RecordCount = 0 Then
        DisplayError ("No Record in Paper Master")
        Cancel = True
        Exit Sub
    Else
        rstPaperList.MoveFirst
    End If
    rstPaperList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstPaperList.EOF Then
        SelectionType = "S"
        PaperCode = ""
        Call LoadSelectionList(rstPaperList, "List of Papers...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text1, PaperCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text1.Text, False) Then
            Text1.Text = "?"
        End If
        If RTrim(PaperCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        PaperCode = rstPaperList.Fields("Code").Value
        fpSpread1.SetText 19, fpSpread1.ActiveRow, Trim(Text1.Text)
        fpSpread1.SetText 20, fpSpread1.ActiveRow, PaperCode
    End If
End Sub
Private Sub MhRealInput27_Validate(Cancel As Boolean)   'Forms/Sheet For Paper Purpose
    If Val(MhRealInput27.Text) <> 0.5 And Val(MhRealInput27.Text) <> 1 And Val(MhRealInput27.Text) <> 2 Then
        Cancel = True
    Else
        fpSpread1.SetText 22, fpSpread1.ActiveRow, Val(MhRealInput27.Text)
        Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
    End If
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)   'Paper Wastage Rate
    fpSpread1.SetText 9, fpSpread1.ActiveRow, Val(MhRealInput11.Text)
    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
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
    fpSpread1.GetText 10, Row, Content   'Paper Consumption (Reams)
    MhRealInput12.Text = Format(Val(Content), "0.000")
    fpSpread1.GetText 11, Row, Content   'Total Plates - ¼F
    MhRealInput3.Text = Format(Val(Content), "0")
    fpSpread1.GetText 12, Row, Content   'Total Plates - ½F
    MhRealInput23.Text = Format(Val(Content), "0")
    fpSpread1.GetText 13, Row, Content   'Total Plates - 1F
    MhRealInput24.Text = Format(Val(Content), "0")
    fpSpread1.GetText 14, Row, Content   'Plate Rate
    MhRealInput4.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 15, Row, Content   'Total Forms - ¼F
    MhRealInput6.Text = Format(Val(Content), "0")
    fpSpread1.GetText 16, Row, Content   'Total Forms - ½F
    MhRealInput25.Text = Format(Val(Content), "0")
    fpSpread1.GetText 17, Row, Content   'Total Forms - 1F
    MhRealInput26.Text = Format(Val(Content), "0")
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
End Sub
Private Sub GetPrinterRates(ByVal xPrintingType As String, ByVal xRateType As String)   'xRateType : 'B'-Both Plate & Print Rate 'L'-Only Plate Rate
    Dim PrintRate As Double, PlateRate As Double, PaperWastageRate As Double, CurrentRate As Variant, PlateType As Variant
    On Error GoTo ErrorHandler
    If rstPrinterRates.State = adStateOpen Then rstPrinterRates.Close
    rstPrinterRates.Open "Select Top 1 * From AccountChild05 Where Code = '" & PrinterCode & "' And [Size]=(SELECT Code FROM SizeGroupChild WHERE [Size]='" & SizeCode & "') And Range" & Trim(xPrintingType) & " >= " & IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) & " Order By Range" & Trim(xPrintingType), CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstPrinterRates.RecordCount = 0 Then
        If rstPrinterRates.State = adStateOpen Then rstPrinterRates.Close
        rstPrinterRates.Open "Select Top 1 * From AccountMaster,AccountChild05 Where AccountMaster.Code = AccountChild05.Code And [Name] Like '%Rate%' And [Size]=(SELECT Code FROM SizeGroupChild WHERE [Size]='" & SizeCode & "') And Range" & Trim(xPrintingType) & " >= " & IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) & " Order By Range" & Trim(xPrintingType), CxnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstPrinterRates.RecordCount > 0 Then
        fpSpread1.GetText 6, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PlateType
        PlateRate = rstPrinterRates.Fields(PlateType & "PlateRate" & Trim(xPrintingType)).Value
        PrintRate = rstPrinterRates.Fields("PrintRate" & Trim(xPrintingType)).Value
        PrintRate = PrintRate + IIf(PrintRate > 0, Val(FrmBookPrintOrder.rstBookList.Fields("AddOnRate01").Value), 0)
        PaperWastageRate = Val(rstPrinterRates.Fields("PaperWastageRate" & Trim(xPrintingType)))
    End If
    fpSpread1.GetText 14, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate  'Plate Rate
    If CurrentRate <> PlateRate Then
        If Val(CheckNull(rstBookPOChild05.Fields("ActualQuantity").Value)) <> 0 Then
            If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Plate rate is different from that in Master ! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
                fpSpread1.SetText 14, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PlateRate
            End If
        Else
            fpSpread1.SetText 14, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PlateRate
        End If
    End If
    If xRateType = "B" Then
        fpSpread1.GetText 18, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate  'Print Rate
        If CurrentRate <> PrintRate And CurrentRate <> 0 Then
            If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Printing Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
                fpSpread1.SetText 18, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PrintRate
            End If
        Else
            fpSpread1.SetText 18, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PrintRate
        End If
        fpSpread1.GetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate   'Paper Wastage Rate
        If CurrentRate <> PaperWastageRate Then
            If Val(CheckNull(rstBookPOChild05.Fields("ActualQuantity").Value)) <> 0 Then
                If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Paper Wastage is different from that in Master ! Change Wastage?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
                    fpSpread1.SetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PaperWastageRate
                End If
            Else
                fpSpread1.SetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PaperWastageRate
            End If
        End If
    End If
    If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        fpSpread1.GetText 14, fpSpread1.ActiveRow, CurrentRate  'Plate Rate
        MhRealInput4.Text = Format(CurrentRate, "0.00")
        fpSpread1.GetText 18, fpSpread1.ActiveRow, CurrentRate  'Print Rate
        MhRealInput5.Text = Format(CurrentRate, "0.00")
        fpSpread1.GetText 9, fpSpread1.ActiveRow, CurrentRate   'Paper Wastage Rate
        MhRealInput11.Text = Format(CurrentRate, "0.00")
    End If
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Fetch Printer Rates")
End Sub
Private Sub CalculateBQD(ByVal xPrintingType As String)    'Calculate Billing Quantity Dependents
    Dim Cnt As Integer, Content As Variant, Forms As Variant
    For Cnt = IIf(xPrintingType = "S", 1, 2) To IIf(xPrintingType = "S", 1, fpSpread1.MaxRows)
        fpSpread1.GetText 1, Cnt, Content   'Pages
        If Val(Content) <> 0 Then GetPrinterRates IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), "B"    'Get Both Plate & Printing Rates
        fpSpread1.GetText 3, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "¼")
        fpSpread1.GetText 4, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "½")
        fpSpread1.GetText 5, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "1")
    Next
    CalculateAmount
End Sub
Private Function CalculateConsumption(ByVal xPrintingType As String) As Double
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant, WastageRate As Variant, CurrentPaperConsumption As Variant, Cnt As Integer, FS As Variant
    
    fpSpread1.GetText 3, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms¼
    fpSpread1.GetText 4, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms½
    fpSpread1.GetText 5, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms1
    fpSpread1.GetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), WastageRate
    CalculateConsumption = CLng(Val(MhRealInput1.Text) * (Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1) * ((100 + Val(WastageRate)) / 100))
    CalculateConsumption = CLng(Val(CalculateConsumption) / 2)
    fpSpread1.GetText 22, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateConsumption = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * CalculateConsumption
    CalculateConsumption = Format(CLng(Int(Val(CalculateConsumption) / 500)) + ((Val(CalculateConsumption) Mod 500) / 1000), "0.000")
    fpSpread1.SetText 10, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateConsumption
    If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        MhRealInput12.Text = Format(Val(CalculateConsumption), "0.000")
    End If
    'Total Consumption Calculation
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 10, Cnt, CurrentPaperConsumption
        MhRealInput13.Text = Format(IIf(Cnt = 1, 0, Val(MhRealInput13.Text)) + CLng((Int(Val(CurrentPaperConsumption)) * 500) + ((Val(CurrentPaperConsumption) - Int(Val(CurrentPaperConsumption))) * 1000)), "0.000")
    Next
    MhRealInput13.Text = Format(CLng(Int(Val(MhRealInput13.Text) / 500)) + ((Val(MhRealInput13.Text) Mod 500) / 1000), "0.000")
End Function
Private Sub CalculateAmount()
    Dim Cnt As Integer, TotalPlates¼ As Variant, TotalPlates½ As Variant, TotalPlates1 As Variant, PlateRate As Variant, TotalForms¼ As Variant, TotalForms½ As Variant, TotalForms1 As Variant, PrintRate As Variant

    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 11, Cnt, TotalPlates¼
        fpSpread1.GetText 12, Cnt, TotalPlates½
        fpSpread1.GetText 13, Cnt, TotalPlates1
        fpSpread1.GetText 14, Cnt, PlateRate
        fpSpread1.GetText 15, Cnt, TotalForms¼
        fpSpread1.GetText 16, Cnt, TotalForms½
        fpSpread1.GetText 17, Cnt, TotalForms1
        fpSpread1.GetText 18, Cnt, PrintRate
        fpSpread1.SetText 7, Cnt, IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalPlates¼) + Val(TotalPlates½) + Val(TotalPlates1)) * Val(PlateRate)
        fpSpread1.SetText 8, Cnt, IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalForms¼) + Val(TotalForms½) + Val(TotalForms1)) * Val(PrintRate)
        If fpSpread1.ActiveRow = Cnt Then
            MhRealInput7.Text = Format(IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalPlates¼) + Val(TotalPlates½) + Val(TotalPlates1)) * Val(PlateRate), "0.00")
            MhRealInput8.Text = Format(IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalForms¼) + Val(TotalForms½) + Val(TotalForms1)) * Val(PrintRate), "0.00")
        End If
    Next
    CalculateTotalAmount
End Sub
Private Function CalculateTotalAmount() As Double
    Dim Cnt As Integer, PlateAmount As Variant, PrintAmount As Variant, TotalAmount As Double
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 7, Cnt, PlateAmount
        fpSpread1.GetText 8, Cnt, PrintAmount
        TotalAmount = TotalAmount + PlateAmount + PrintAmount
    Next
    TotalAmount = TotalAmount + MhRealInput9.Value
    MhRealInput29.Value = MhRealInput7.Value * MhRealInput28.Value / 100    'GST Plate
    MhRealInput18.Value = MhRealInput8.Value * MhRealInput14.Value / 100    'GST Printing
    MhRealInput10.Value = Round(TotalAmount + MhRealInput18.Value + MhRealInput29.Value, 0) 'Total Amount
End Function
Private Function CalculateTotalForms(ByVal xPrintingType As String, ByVal Forms As Double, ByVal FormType As String) As Double
    Dim FS As Variant
    fpSpread1.GetText 21, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateTotalForms = (Int(IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) * IIf(FormType = "¼", 0.25, IIf(FormType = "½", 0.5, 1)) / 1000) + IIf(IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) * IIf(FormType = "¼", 0.25, IIf(FormType = "½", 0.5, 1)) Mod 1000 = 0, 0, 1)) * Forms
    CalculateTotalForms = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * Val(CalculateTotalForms)
    If FrmBookPrintOrder.rstBookList.Fields("DuplexPrinting").Value = "N" Then CalculateTotalForms = 0.5 * CalculateTotalForms
    CalculateTotalForms = Int(Val(CalculateTotalForms)) + IIf(Val(CalculateTotalForms) - Int(Val(CalculateTotalForms)) = 0, 0, 1)
    If FormType = "¼" Then
        fpSpread1.SetText 15, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then MhRealInput6.Text = Format(CalculateTotalForms, "0")
    ElseIf FormType = "½" Then
        fpSpread1.SetText 16, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then MhRealInput25.Text = Format(CalculateTotalForms, "0")
    Else
        fpSpread1.SetText 17, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then MhRealInput26.Text = Format(CalculateTotalForms, "0")
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
    Dim Cnt As Integer, PaperBalance As Double, PaperCode As Variant, PaperName As Variant, PaperStock As Variant, PaperConsumption As Variant
    If CheckMandatoryFields Then Exit Sub
    If FrmBookPrintOrder.BookPOType <> "O" Then
        For Cnt = 1 To fpSpread1.MaxRows
            fpSpread1.SetActiveCell 1, Cnt
            fpSpread1_DblClick 1, Cnt
            fpSpread1.GetText 20, Cnt, PaperCode
            PaperStock = CalculatePaperBalance(PrinterCode, PaperCode, CheckNull(rstBookPOChild05.Fields("Code").Value), "BPOB")
            fpSpread1.GetText 10, Cnt, PaperConsumption
            If Not CheckEmpty(PaperCode, False) Then
                PaperBalance = Val(PaperStock) - Int(Val(PaperConsumption)) * 500 - Round((Val(PaperConsumption) - Int(Val(PaperConsumption))), 3) * 1000
                If PaperBalance < 0 Then
                    PaperBalance = Format(CLng(Int(Val(Abs(PaperBalance)) / 500)) + ((Val(Abs(PaperBalance)) Mod 500) / 1000), "0.000")
                    fpSpread1.GetText 19, Cnt, PaperName
                    If UserLevel <= 2 Then
                        If MsgBox("Stock (" & Format(0 - PaperBalance, "0.000") & ") of the Paper - " & Trim(PaperName) & vbCrLf & " is going negative ! Would you like to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbNo Then
                            Exit Sub
                        End If
                    Else
                        Call DisplayError("Cann't Save ! Stock (" & Format(0 - PaperBalance, "0.000") & ") of the Paper - " & Trim(PaperName) & " is going negative"): AbortPO = True: Exit Sub
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
    If Val(MhRealInput16.Text) <> 0 Then If Val(MhRealInput16.Text) <> Val(MhRealInput10.Text) Then MhRealInput9.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput9.Text) <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
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
