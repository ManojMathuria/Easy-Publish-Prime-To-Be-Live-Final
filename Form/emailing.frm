VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmEmailing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19845
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   19845
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   9150
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   19650
      _Version        =   65536
      _ExtentX        =   34660
      _ExtentY        =   16140
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
      Picture         =   "emailing.frx":0000
      Begin VB.CommandButton cmdRefresh 
         Height          =   450
         Left            =   11765
         Picture         =   "emailing.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Refresh"
         Top             =   40
         Width           =   450
      End
      Begin VB.CommandButton cmdProceed 
         Height          =   450
         Left            =   12240
         Picture         =   "emailing.frx":0166
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Save"
         Top             =   40
         Width           =   450
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   450
         Left            =   12720
         Picture         =   "emailing.frx":0268
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Cancel"
         Top             =   40
         Width           =   450
      End
      Begin MSComDlg.CommonDialog cdUpload 
         Left            =   2400
         Top             =   3360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdUpload 
         Height          =   450
         Left            =   10800
         Picture         =   "emailing.frx":036A
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Upload Item Pic"
         Top             =   40
         Width           =   450
      End
      Begin VB.CommandButton Command1 
         Height          =   450
         Left            =   11280
         Picture         =   "emailing.frx":0A6C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   40
         Width           =   450
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
         Left            =   7320
         MaxLength       =   40
         TabIndex        =   20
         Text            =   " SEDEX Audited Printing Press"
         ToolTipText     =   "Find And Search"
         Top             =   120
         Width           =   3390
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   18150
         TabIndex        =   13
         Top             =   110
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Print Preview"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Print"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Mail"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Exit"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
         Height          =   330
         Left            =   2790
         TabIndex        =   14
         Top             =   120
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   582
         _StockProps     =   77
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
         Picture         =   "emailing.frx":116E
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
            Height          =   330
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
            _ExtentY        =   582
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TintColor       =   16711935
            Caption         =   " &Show"
            Alignment       =   0
            FillColor       =   9164542
            TextColor       =   0
            Picture         =   "emailing.frx":118A
            Picture         =   "emailing.frx":11A6
         End
         Begin MSForms.OptionButton Option1 
            Height          =   300
            Left            =   720
            TabIndex        =   17
            Top             =   10
            Width           =   615
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "1085;529"
            Value           =   "0"
            Caption         =   "All "
            FontName        =   "Calibri"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.OptionButton Option2 
            Height          =   300
            Left            =   1440
            TabIndex        =   15
            Top             =   10
            Width           =   1095
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "1931;529"
            Value           =   "1"
            Caption         =   "Pending"
            FontName        =   "Calibri"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.OptionButton Option3 
            Height          =   300
            Left            =   2640
            TabIndex        =   16
            Top             =   10
            Width           =   855
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "1508;529"
            Value           =   "0"
            Caption         =   "Close"
            FontName        =   "Calibri"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.CommandButton Preview 
         Caption         =   "&Print Preview"
         Height          =   330
         Left            =   14400
         TabIndex        =   12
         Top             =   8760
         Width           =   1215
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
         Left            =   3240
         MaxLength       =   40
         TabIndex        =   3
         ToolTipText     =   "Find And Search"
         Top             =   8760
         Width           =   3390
      End
      Begin VB.CommandButton cmdFilter 
         Height          =   320
         Left            =   6720
         Picture         =   "emailing.frx":11C2
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Filter"
         Top             =   8760
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   320
         Left            =   7200
         Picture         =   "emailing.frx":1504
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Search"
         Top             =   8760
         Width           =   375
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   330
         Left            =   1200
         TabIndex        =   4
         Top             =   8760
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "emailing.frx":1846
         Caption         =   "emailing.frx":1866
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "emailing.frx":18CA
         Keys            =   "emailing.frx":18E8
         Spin            =   "emailing.frx":1932
         AlignHorizontal =   2
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   255
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   8760
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " Data Count"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "emailing.frx":195A
         Picture         =   "emailing.frx":1976
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   18330
         TabIndex        =   6
         Top             =   8760
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
         Caption         =   " Print Data"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "emailing.frx":1992
         Picture         =   "emailing.frx":19AE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   17040
         TabIndex        =   7
         Top             =   8760
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
         Caption         =   " Export Data"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "emailing.frx":19CA
         Picture         =   "emailing.frx":19E6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   2520
         TabIndex        =   8
         Top             =   8760
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
         Caption         =   " Find"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "emailing.frx":1A02
         Picture         =   "emailing.frx":1A1E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel 
         Height          =   330
         Left            =   9840
         TabIndex        =   9
         Top             =   8760
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   "Ctrl+F->Search ->  F5->Refresh"
         FillColor       =   8421504
         TextColor       =   16777215
         Picture         =   "emailing.frx":1A3A
         Picture         =   "emailing.frx":1A56
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   15735
         TabIndex        =   11
         Top             =   8760
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
         Caption         =   "Import Data"
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "emailing.frx":1A72
         Picture         =   "emailing.frx":1A8E
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
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "emailing.frx":1AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "emailing.frx":1FEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "emailing.frx":2102
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "emailing.frx":2214
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   8085
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   19410
         _Version        =   524288
         _ExtentX        =   34237
         _ExtentY        =   14261
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
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
         MaxCols         =   11
         MaxRows         =   5000
         SelectBlockOptions=   11
         SpreadDesigner  =   "emailing.frx":2326
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   6480
         TabIndex        =   21
         Top             =   120
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
         Caption         =   " Subject:"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "emailing.frx":2DB5
         Picture         =   "emailing.frx":2DD1
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
         Height          =   330
         Left            =   1440
         TabIndex        =   22
         Top             =   120
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "emailing.frx":2DED
         Caption         =   "emailing.frx":2E0D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "emailing.frx":2E71
         Keys            =   "emailing.frx":2E8F
         Spin            =   "emailing.frx":2ED9
         AlignHorizontal =   2
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   255
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   5
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   23
         Top             =   120
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
         Caption         =   " Set Max. &Email"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "emailing.frx":2F01
         Picture         =   "emailing.frx":2F1D
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   330
         Left            =   7680
         TabIndex        =   10
         Top             =   8760
         Width           =   2085
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3678;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   19810
         Y1              =   540
         Y2              =   540
      End
   End
End
Attribute VB_Name = "FrmEmailing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VchType As String, OutputTo As String
Dim oOutlook As New Outlook.Application
Dim ToID As Variant, ToName As Variant
Dim cnEmail As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstEmail As New ADODB.Recordset
Public rstEmailTable As New ADODB.Recordset
Dim EditMode As Boolean
Dim FileName As String
Dim Reset As Long, LR As Integer, UpdateFlag As Boolean
Private Sub cmdFilter_Click()
Total_Click
End Sub
Private Sub Command1_Click()
fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.MaxCols = 11
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT * FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    cnEmail.CursorLocation = adUseClient
    cnEmail.Open cnDatabase.ConnectionString
    rstEmailTable.Open "Select Code,Company,ContactPerson,Mobile,email,Address,PIN,CITY,Category,State,Status From dbo.Email ORDER BY Code", cnEmail, adOpenKeyset, adLockOptimistic
    Reset = 0:
    'Combo3.AddItem
    Combo3.Clear
    Combo3.AddItem " Company", 0
    Combo3.AddItem " Contact Person", 1
    Combo3.AddItem " Mobile", 2
    Combo3.AddItem " Email", 3
    Combo3.AddItem " Address", 4
    Combo3.AddItem " City", 5
    Combo3.AddItem " Category", 6
    Combo3.AddItem " State", 6
    Combo3.AddItem " Status", 7
    Combo3.ListIndex = 1
    Reset = 1
    cmdRefresh_Click
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Combo2_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Combo4_Change()
If Reset = 1 Then Call cmdRefresh_Click
End Sub
Private Sub Check2_Click()
Call cmdRefresh_Click
End Sub
Private Sub fpSpread1_BeforeUserSort(ByVal Col As Long, ByVal State As FPSpreadADO.BeforeUserSortStateConstants, DefaultAction As FPSpreadADO.BeforeUserSortDefaultActionConstants)
    Dim n As Integer
    With fpSpread1
    If .DataRowCnt = 0 Then Exit Sub
    n = .DataRowCnt:
    .RowHeight(n) = 12.75
    .DeleteRows n, 1
    End With
    End Sub
Private Sub fpSpread1_AfterUserSort(ByVal Col As Long)
With fpSpread1
If .DataRowCnt = 0 Then Exit Sub
End With
'Call Total_Click
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}": KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        If Not EditMode Then cmdProceed_Click: KeyCode = 0
    ElseIf KeyCode = vbKeyF5 Then
        Call cmdRefresh_Click
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        If Not EditMode Then cmdCancel_Click: KeyCode = 0
    
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
        If Text1.Text = "" Then
            MsgBox "Please Provide Search Input", vbInformation
            Text1.SetFocus
        ElseIf Text1.Text <> "" Then
        Call Command2_Click
        End If
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstEmailTable)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(cnEmail)
    End Sub
Private Sub cmdProceed_Click()
        SaveFields
End Sub
Private Sub cmdCancel_Click()
Call CloseForm(Me)
End Sub
Private Sub SaveFields()
    Dim i As Integer, Code As Variant, Company As Variant, ContactPerson As Variant, Mobile As Variant, Email As Variant, Address As Variant, PIN As Variant, CITY As Variant, Category As Variant, State As Variant, Status As Variant, SrNo As Variant, n As Integer, SNo As Variant
         i = 0: SrNo = "000000"
         SNo = GenerateCode(cnEmail, "Select Top(1) ISNULL(Code,1) From Email Order By Code DESC", 6, 0)
         If IsNull(SNo) Then SNo = 1
         SNo = Val(SNo) - 1
    With rstEmail
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
        cnEmail.Execute "DELETE FROM Email WHERE Code='" & rstEmail.Fields("Code").Value & "'"
        .MoveNext
        Loop
    End With
    With fpSpread1
        For i = 1 To .DataRowCnt
                    .GetText 1, i, Company
                    .GetText 2, i, ContactPerson
                    .GetText 3, i, Mobile
                    .GetText 4, i, Email
                    .GetText 5, i, Address
                    .GetText 6, i, PIN
                    .GetText 7, i, CITY
                    .GetText 8, i, Category
                    .GetText 9, i, State
                    .GetText 10, i, Status
                    .GetText 11, i, Code
                With rstEmailTable
                    .AddNew
                    If Trim(Code) = "" Then SNo = Pad((Val(SNo) + 1), "0", 6, "L")
                    .Fields("Code").Value = IIf(Trim(Code) = "", SNo, Trim(Code))
                    .Fields("Company").Value = Trim(Company)
                    .Fields("ContactPerson").Value = Trim(ContactPerson)
                    .Fields("Mobile").Value = Trim(Mobile)
                    .Fields("email").Value = Trim(Email)
                    .Fields("Address").Value = Trim(Address)
                    .Fields("PIN").Value = PIN
                    .Fields("CITY").Value = Trim(CITY)
                    .Fields("Category").Value = Trim(Category)
                    .Fields("State").Value = Trim(State)
                    .Fields("Status").Value = Trim(Status)
                    .Update
                End With
        Next
    End With
        cmdRefresh_Click
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub cmdRefresh_Click()
Dim mcN As String, MC As String, C As Long, cVal(1 To 7) As Variant, Col As Variant, K As Long
Text1.Text = "": TDBNumber2 = 0:
fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.MaxCols = 11
fpSpread1.RowHeadersAutoText = DispBlank
fpSpread1.ColHeadersAutoText = DispBlank
    FormatCols
    On Error GoTo ErrHandler
    Dim SQL As String, i As Long
    'MF1
  SQL = "Select Code,Company,ContactPerson,Mobile,email,Address,PIN,CITY,Category,State,Status From dbo.Email Where " & IIf(Option1.Value, "1=1", IIf(Option2.Value, "Status=''", "Status<>''")) & _
              " ORDER BY Code,Company,ContactPerson,CITY,Category,State"
    Screen.MousePointer = vbHourglass
    If rstEmail.State = adStateOpen Then rstEmail.Close
    rstEmail.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
    If rstEmail.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
    With fpSpread1
       .ClearRange -1, 1, .MaxCols, .MaxRows, False
       .ColsFrozen = 0
       rstEmail.MoveFirst
       Do While Not rstEmail.EOF
            i = i + 1
            TDBNumber2 = TDBNumber2 + 1 'Data Count
                K = K + 1
                .SetText .RowHeaderCols - 1, i, K
                .SetText 1, i, Trim(rstEmail.Fields("Company").Value)
                .SetText 2, i, Trim(rstEmail.Fields("ContactPerson").Value)
                .SetText 3, i, Trim(rstEmail.Fields("Mobile").Value)
                .SetText 4, i, Trim(rstEmail.Fields("email").Value)
                .SetText 5, i, Trim(rstEmail.Fields("Address").Value)
                .SetText 6, i, Trim(rstEmail.Fields("PIN").Value)
                .SetText 7, i, Trim(rstEmail.Fields("City").Value)
                .SetText 8, i, Trim(rstEmail.Fields("Category").Value)
                .SetText 9, i, Trim(rstEmail.Fields("State").Value)
                .SetText 10, i, Trim(rstEmail.Fields("Status").Value)
                .SetText 11, i, Trim(rstEmail.Fields("Code").Value)
            rstEmail.MoveNext
        Loop
        .ColsFrozen = 0
    End With
Screen.MousePointer = vbNormal
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Sub
Private Sub FormatCols()
Dim C As Long
    With fpSpread1
        .ColWidth(1) = 30
        .ColWidth(2) = 17
        .ColWidth(3) = 10
        .ColWidth(4) = 21
        .ColWidth(5) = 28.875
        .ColWidth(6) = 9
        .ColWidth(7) = 9
        .ColWidth(9) = 11
        .ColWidth(10) = 10.25
    'Header Text
        .Row = SpreadHeader
        .Col = 1: .Text = "Company"
        .Col = 2: .Text = "Contact Person Name"
        .Col = 3: .Text = "Mobile"
        .Col = 4: .Text = "Email Address"
        .Col = 5: .Text = "Address"
        .Col = 6: .Text = "Pin Code"
        .Col = 7: .Text = "City"
        .Col = 8: .Text = "Category "
        .Col = 9: .Text = "State"
        .Col = 10: .Text = "Status"
        .Col = 11: .Text = "Code"
    End With
End Sub
Private Sub Command2_Click()
  Dim i As Integer, cVal As Variant, R As Long, C As Long
    With fpSpread1
    If Text1.Text = "" Then Exit Sub
            If .DataRowCnt = 0 Then Exit Sub
                For i = 1 To .DataRowCnt 'Unhide All
                .Row = i: .RowHidden = False
            Next
            R = IIf(.ActiveRow + 1 <> LR, .ActiveRow + 1, 1)
            LR = R
            C = Combo3.ListIndex + 1
            For i = R To .DataRowCnt
            If Combo3.ListIndex >= 0 Then .GetText C, i, cVal
                        If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
                        ElseIf Combo3.ListIndex >= 0 Then
                        .SetActiveCell C, i: Exit Sub
                        End If
            Next
    End With
End Sub
Private Sub Option1_Click()
Call cmdRefresh_Click
End Sub
Private Sub Option2_Click()
Call cmdRefresh_Click
End Sub
Private Sub Option3_Click()
Call cmdRefresh_Click
End Sub
Private Sub Preview_Click()
Dim PrintHeader As String
Dim R As Long, C As Long, i As Long
Dim JQty As Variant
Dim answer As Integer
'*********************************************************
With fpSpread1
.ColsFrozen = 0
PrintHeader = Me.Caption
.LockBackColor = vbWhite
'' These are 8.5" X 11" paper dimensions in TWIPS  12240  15840
Const PaperWidth = 12240
Const PaperHeight = 15840
'Printer.PaperSize = vbPRPSA4
' Set printing options for sheet
fpSpread1.PrintAbortMsg = "Printing - Click Cancel to .Quit"
fpSpread1.PrintJobName = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") '& ".pdf"
'fpSpread1.PrintHeader = "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy"): fpSpread1.PrintHeader=: .Font = 20 '"/cPrint Header/rPage # ./p/n2nd Line"
fpSpread1.PrintFooter = "        Export Data Company : " & rstCompanyMaster.Fields("PrintName").Value & " _(" & CompCode & "_" & PrintHeader & ")" & "  From [" + Format(Date, "dd-MM-yyyy") + "] " & " /rPage # ./p " & " Print Date : ( " & Format(Date, "dd-MMM-yyyy") & " )         ": .FontSize = 16  '& ".pdf" ' "/cPrint Footer/rPage # ./p/n2nd Line"
fpSpread1.PrintBorder = True
fpSpread1.PrintColHeaders = True
fpSpread1.PrintColor = True
fpSpread1.PrintGrid = True
fpSpread1.PrintMarginTop = 200 '750 '1440
fpSpread1.PrintMarginBottom = 200 '500 '1440
fpSpread1.PrintMarginLeft = 100 '720
fpSpread1.PrintMarginRight = 100 '720
'fpSpread1.PrintType = SPRD_PRINT_ALL
fpSpread1.PrintRowHeaders = True
'fpSpread1.PrintShadows = True
fpSpread1.PrintUseDataMax = True
' Center vertically
fpSpread1.PrintCenterOnPageV = False
' Center horizontally
fpSpread1.PrintCenterOnPageH = True
' Perform the printing action
' Set the sheet to print
fpSpread1.Sheet = 1
' Set scaling method
fpSpread1.PrintScalingMethod = PrintScalingMethodZoom
' Set zoom factor
fpSpread1.PrintZoomFactor = 0.75
' Print
'fpSpread1.PrintSheet 0
fpSpread1.PrintOrientation = PrintOrientationLandscape
'fpSpread1.PrintSheet
'.LockBackColor = RGB(245, 255, 230) '(250, 255, 242) '
   
   'If a cell is currently active, turn off edit mode
    If fpSpread1.EditMode = True Then
        fpSpread1.EditMode = False
        DoEvents
    End If
    Set spreadpreview.frm = Me
    Set pagesetup.frmPageSetup = Me
    Set PrintDlg.frmPrintDlg = Me
    Set headerfooter.frmHeaderFooter = Me
MsgBox "Processing Complete"
          spreadpreview.Show
 End With
End Sub
Private Sub Mh3dLabel7_Click()
Dim x As Boolean, FileName As String, SheetName As String, LogFileName As String
Dim R As Long, C As Long, Header1 As String
Dim JQty As Variant
    '"Export Data" &
    With fpSpread1
        .ColsFrozen = 0
        fpSpread1.InsertRows 1, 5
              For R = 1 To 5
                    For C = 1 To .MaxCols
                    .Col = C: .Row = R: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .SetText C, R, ""
                Next C
            Next R
'                'Header-1
                .AddCellSpan 1, 1, 34, 1:
                R = 1
                .Col = 1: .Row = R: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .FontBold = True: .FontSize = 20: .ForeColor = RGB(1, 106, 106): .FontUnderline = True:
                .GetText 1, SpreadHeader, JQty
                .SetText 1, R, JQty
                'Header-2
                .AddCellSpan 1, 2, 34, 1:
                R = 2
                .Col = 1: .Row = R: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .FontBold = True: .FontSize = 16: .ForeColor = RGB(1, 106, 106): .FontUnderline = True:
                .GetText 1, SpreadHeader + 1, JQty
                .SetText 1, R, JQty
               .AddCellSpan 1, 3, 34, 1
               .AddCellSpan 1, 4, 34, 1
                'Header-3
                R = 5
                For C = 1 To .MaxCols
                .Col = C: .Row = R: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .FontBold = True: .FontSize = 10: ForeColor = RGB(1, 106, 106): .FontUnderline = True:
                .GetText C, SpreadHeader + 2, JQty
                .SetText C, R, JQty
                Next
    End With
    If Dir(App.Path & "\Export", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Export"
    '
    ' Export Excel file and set result to x
     FileName = App.Path & "\Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"
    SheetName = "Sheet1" '"(" & Me.Caption & ")"
    LogFileName = "Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".txt"
    x = fpSpread1.ExportToExcelEx(FileName, SheetName, LogFileName, ExcelSaveFlagNone)
    ' Display result to user based on T/F value of x
    If x = True Then
    MsgBox "Export complete.", vbInformation, "Easy Publish...Export !!! "
        Dim oExcel As Object
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open (FileName)
        oExcel.Visible = True
        oExcel.Sheets("Sheet1").Select
        oExcel.Sheets("Sheet1").Unprotect
         Set oExcel = Nothing
    Else
    MsgBox "Export did not succeed.", vbInformation, "Easy Publish...Export !!!"
    End If
    '
    With fpSpread1
    fpSpread1.DeleteRows 1, 5
    End With
End Sub
Private Sub Mh3dLabel6_Click()
Dim i As Long
With fpSpread1
Dim PrintHeader As String
Dim R As Long, C As Long
Dim JQty As Variant
Dim answer As Integer
    .ColsFrozen = 0
    .MaxRows = .MaxRows + 2
    .InsertRows 1, 2
    .SetText .RowHeaderCols - 1, 1, " ": .SetText .RowHeaderCols - 1, 2, " "
    .Col = 9: fpSpread1.ColHidden = True
    For i = .DataRowCnt To .MaxRows
    .SetText .RowHeaderCols - 1, i, " "
    Next i
    For R = 1 To 2
        For C = 1 To .MaxCols
            .Col = C: .Row = R:  .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: '.BackColor = &H8000000F: .FontBold = True: .FontSize = 20: .FontUnderline = True: .ForeColor = vbRed:
            .SetText C, R, ""
        Next C
    Next R
    'Col Span
    .AddCellSpan 1, 1, 30, 1
    .AddCellSpan 1, 2, 30, 1
    
     .Col = 1: .Row = 1: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .FontBold = True: .FontSize = 20: .FontUnderline = True: .ForeColor = RGB(1, 106, 106): .SetText 1, 1, rstCompanyMaster.Fields("PrintName").Value:
     .Col = 1: .Row = 2: .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter:  .FontBold = True: .FontSize = 16: .ForeColor = RGB(20, 106, 106): .SetText 1, 2, Me.Caption + "  Dated: [" + Format(GetDate(Date), "dd-MM-yyyy") + "] ":
    R = 1
PrintHeader = Me.Caption
.LockBackColor = vbWhite
' These are 8.5" X 11" paper dimensions in TWIPS  12240  15840
Const PaperWidth = 12240
Const PaperHeight = 15840
' Set printing options for sheet
fpSpread1.PrintAbortMsg = "Printing - Click Cancel to .Quit"
fpSpread1.PrintJobName = "Export Data" & "(" & CompCode & "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy") '& ".pdf"
'fpSpread1.PrintHeader = "_" & PrintHeader & ")" & Format(Date, "dd-MMM-yyyy"): fpSpread1.PrintHeader=: .Font = 20 '"/cPrint Header/rPage # ./p/n2nd Line"
fpSpread1.PrintFooter = "        Export Data Company : " & rstCompanyMaster.Fields("PrintName").Value & " _(" & CompCode & "_" & PrintHeader & ")" & "  Dated: [" + Format(GetDate(Date), "dd-MM-yyyy") + "] " & " /rPage # ./p " & " Print Date : ( " & Format(Date, "dd-MMM-yyyy") & " )         ": .FontSize = 16 '& ".pdf" ' "/cPrint Footer/rPage # ./p/n2nd Line"
fpSpread1.PrintBorder = True
fpSpread1.PrintColHeaders = True
fpSpread1.PrintColor = True
fpSpread1.PrintGrid = True
fpSpread1.PrintMarginTop = 750 '1440
fpSpread1.PrintMarginBottom = 500 '1440
fpSpread1.PrintMarginLeft = 100 '720
fpSpread1.PrintMarginRight = 100 '720
'fpSpread1.PrintType = SPRD_PRINT_ALL
fpSpread1.PrintRowHeaders = True
fpSpread1.PrintShadows = True
fpSpread1.PrintUseDataMax = True
' Center vertically
fpSpread1.PrintCenterOnPageV = False
' Center horizontally
fpSpread1.PrintCenterOnPageH = True
' Perform the printing action
' Set the sheet to print
fpSpread1.Sheet = 1
' Set scaling method
fpSpread1.PrintScalingMethod = PrintScalingMethodZoom
' Set zoom factor
fpSpread1.PrintZoomFactor = 0.75
' Print
'fpSpread1.PrintSheet 0
MsgBox "Processing Complete"
fpSpread1.PrintOrientation = PrintOrientationLandscape
fpSpread1.PrintSheet
    'Delete Row
    .DeleteRows 1, 2
    .MaxRows = .MaxRows - 2
    .Col = 9: fpSpread1.ColHidden = False
 End With
End Sub
Private Sub Mh3dLabel9_Click()
Dim C As Long, R As Long
Dim JQty As Variant
Dim FileName As String
If Dir(App.Path & "\Export", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Export"
 FileName = App.Path & "\Export\Export Data" & "(" & CompCode & "_" & Me.Caption & ")" & Format(Date, "dd-MMM-yyyy") & ".xls"
' Save to xls file type
' Load an Excel-formatted file
fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, False
'MsgBox
    MsgBox "Import Processing....", vbInformation, "Easy Publish...Import !!! "
fpSpread1.ImportExcelBook FileName, ""
cmdRefresh.Visible = False
    With fpSpread1
            fpSpread1.DeleteRows .DataRowCnt, 1
    End With
End Sub
Private Sub cmdUpload_Click() 'Load Pic
    On Error Resume Next
    With cdUpload
        .CancelError = True
        .DialogTitle = "Attach File"
        .Filter = "All Files|*.jpg;*.jpeg;*.bmp;*.gif;*.png;*.pdf;*.xlsx;*.xlsm;.xltx;*.xlsb;*.xltm;*.xls;*.xlt"
        .ShowOpen
        If Err.Number = 0 Then FileName = .FileName: cmdUpload.Enabled = False 'Ok Selected
    End With
End Sub
Private Sub Total_Click()
    Dim i As Integer, j As Integer, cVal As Variant, n As Integer, R As Long, C As Long, Cols As Long, nVal As Variant
    Dim cValue(1 To 33) As Variant, cValTotal(1 To 33) As Variant
    With fpSpread1
    n = 0
        If .DataRowCnt = 0 Then Exit Sub
        n = .DataRowCnt:
    For i = 1 To .DataRowCnt 'Unhide All
        .GetText 4, i, cVal
        .Row = i: .RowHidden = False
        If cVal = "Grand Total" Then fpSpread1.DeleteRows i, 1: n = n - 1
    Next

    .MaxRows = n + 1
    C = Combo3.ListIndex + 1
    'Get Value
For i = 1 To .DataRowCnt
        If Combo3.ListIndex >= 0 Then .GetText C, i, cVal
        For j = 1 To 33
            .GetText j, i, cValue(j)
            .GetText 7, i, cValue(7)
        Next j
                 
            .GetText C, i, cVal
            .GetText 2, i, nVal
            If nVal = "" Then n = n - 1
            .GetText 4, i, nVal
            If nVal = "Grand Total" Then fpSpread1.DeleteRows .DataRowCnt, 1:
'Filter
            .GetText 2, i, nVal
    If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then
            If cVal <> "" And nVal <> "" Then .Row = i: .RowHidden = True: n = n - 1 'Hide Filter
    Else
            .Row = i
            If Not .RowHidden Then
                For j = 7 To 27
                    cValTotal(j) = Val(cValTotal(j)) + Val(cValue(j)) '(7 To 27)
                Next j
                If cValue(28) <> "" Then cValTotal(28) = cValTotal(28) + Left(cValue(28), (Len(cValue(28)) - 3)) + Right(cValue(28), 2) / 60 '28
                For j = 32 To 33
                    cValTotal(j) = Val(cValTotal(j)) + Val(cValue(j)) '(32 To 33)
                Next j
            End If
    End If
Next
    TDBNumber2 = n 'Data Count
    End With
    End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 4 Then CloseForm Me: Exit Sub
    OutputTo = Choose(Button.Index, "S", "P", "M")
    List
End Sub
Private Sub List()
Dim i As Integer, Code As Variant, iCount As Integer
    On Error GoTo NXT
If OutputTo = "S" Then
Preview_Click
ElseIf OutputTo = "P" Then
Preview_Click
ElseIf OutputTo = "M" Then
        For i = 1 To (TDBNumber1.Value + 1)
        i = i - iCount
        fpSpread1.Row = i
        If Not fpSpread1.RowHidden Then
            fpSpread1.GetText 2, i, ToName: ToName = Trim(ToName)
            fpSpread1.GetText 4, i, ToID: ToID = Trim(ToID)
            fpSpread1.GetText 10, i, Code: Code = Trim(Code)
            If ToID = "" Or Code = "Sent" Then GoTo NXT
            If iCount = 0 Then
'            If UpdateFlag = False Then MsgBox "Sending-email TO: " & ToName & "   @" & ToID & " " & i & "!!!"
            End If
            If Code <> "Sent" Then Call SendEmail
            If UpdateFlag = False Then iCount = 1 Else iCount = 0
            fpSpread1.GetText 11, i, Code: Code = Trim(Code)
            If UpdateFlag = True Then fpSpread1.SetText 10, i, "Sent"
            If UpdateFlag = True Then cnEmail.Execute "Update Email Set Status='Sent' Where Code= '" & Code & "'"
            UpdateFlag = False
        End If
NXT:
        Next
End If
    Screen.MousePointer = vbNormal
End Sub
Private Sub SendEmail()
On Error GoTo NXT
    Dim oOutlookMsg As Outlook.MailItem, iCount As Integer
    Screen.MousePointer = vbHourglass
    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
    With oOutlookMsg
                .Display
                .To = ToID
                .Subject = Trim(Text2.Text)
                .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir/Madam,<Br>&nbsp;" & .HTMLBody
                If FileName <> "" Then
                .Attachments.Add (FileName)
                End If
                .Importance = olImportanceHigh
                .ReadReceiptRequested = True
                If CheckEmpty(.To, False) Then .Display Else .Send: UpdateFlag = True
            End With
NXT:
    Set oOutlookMsg = Nothing
    Set oOutlook = Nothing
End Sub
