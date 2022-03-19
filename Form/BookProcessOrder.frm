VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBookProcessOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Process Order"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13740
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   13740
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   8250
      Left            =   15
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   13715
      _Version        =   65536
      _ExtentX        =   24192
      _ExtentY        =   14552
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
      Picture         =   "BookProcessOrder.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8030
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Width           =   13485
         _ExtentX        =   23786
         _ExtentY        =   14155
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&List"
         TabPicture(0)   =   "BookProcessOrder.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Mh3dLabel1(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "DataGrid1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Text1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "BookProcessOrder.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).ControlCount=   1
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   605
            MaxLength       =   40
            TabIndex        =   17
            Top             =   7590
            Width           =   8220
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7070
            Left            =   120
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   450
            Width           =   13260
            _ExtentX        =   23389
            _ExtentY        =   12462
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
               Weight          =   700
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
               DataField       =   "Name"
               Caption         =   "    Order No."
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
            BeginProperty Column01 
               DataField       =   "Date"
               Caption         =   "Order Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd-MM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "ProcessorName"
               Caption         =   "Processor Name"
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
               DataField       =   "BookName"
               Caption         =   " Item Name"
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
            BeginProperty Column04 
               DataField       =   "BillAmount"
               Caption         =   "     Order Amount"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   0
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
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   4529.764
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   4545.071
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  Locked          =   -1  'True
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7045
            Left            =   -74880
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   480
            Width           =   13260
            _Version        =   65536
            _ExtentX        =   23389
            _ExtentY        =   12427
            _StockProps     =   77
            Enabled         =   0   'False
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
            Picture         =   "BookProcessOrder.frx":0054
            Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
               Height          =   300
               Left            =   7350
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   5295
               Width           =   1300
               _Version        =   65536
               _ExtentX        =   2293
               _ExtentY        =   529
               Calculator      =   "BookProcessOrder.frx":0070
               Caption         =   "BookProcessOrder.frx":0090
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookProcessOrder.frx":00FC
               Keys            =   "BookProcessOrder.frx":011A
               Spin            =   "BookProcessOrder.frx":0164
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
               ValueVT         =   2005925893
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
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
               Left            =   7080
               MaxLength       =   40
               TabIndex        =   9
               Top             =   6105
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
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   4
               Top             =   945
               Width           =   11475
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
               TabIndex        =   5
               Top             =   1260
               Width           =   11475
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
               MaxLength       =   60
               TabIndex        =   3
               Top             =   630
               Width           =   11475
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
               Height          =   330
               Left            =   1680
               TabIndex        =   8
               Top             =   6105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "BookProcessOrder.frx":018C
               Caption         =   "BookProcessOrder.frx":01AC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookProcessOrder.frx":0218
               Keys            =   "BookProcessOrder.frx":0236
               Spin            =   "BookProcessOrder.frx":0280
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
               ValueVT         =   2005925893
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   330
               Left            =   1680
               TabIndex        =   7
               Top             =   5790
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "BookProcessOrder.frx":02A8
               Caption         =   "BookProcessOrder.frx":02C8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookProcessOrder.frx":0334
               Keys            =   "BookProcessOrder.frx":0352
               Spin            =   "BookProcessOrder.frx":039C
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
               TabIndex        =   10
               Top             =   6620
               Width           =   1575
            End
            Begin VB.TextBox Text2 
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
               Left            =   1680
               MaxLength       =   10
               TabIndex        =   0
               Top             =   105
               Width           =   1575
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   20
               Top             =   6600
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
               Caption         =   " Bill No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookProcessOrder.frx":03C4
               Picture         =   "BookProcessOrder.frx":03E0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   21
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
               Picture         =   "BookProcessOrder.frx":03FC
               Picture         =   "BookProcessOrder.frx":0418
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   5565
               TabIndex        =   22
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
               Picture         =   "BookProcessOrder.frx":0434
               Picture         =   "BookProcessOrder.frx":0450
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   120
               TabIndex        =   23
               Top             =   5790
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
               Caption         =   " GST (%)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookProcessOrder.frx":046C
               Picture         =   "BookProcessOrder.frx":0488
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   10020
               TabIndex        =   24
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
               Caption         =   " Delivery Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookProcessOrder.frx":04A4
               Picture         =   "BookProcessOrder.frx":04C0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Left            =   120
               TabIndex        =   25
               Top             =   6105
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
               Picture         =   "BookProcessOrder.frx":04DC
               Picture         =   "BookProcessOrder.frx":04F8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   10020
               TabIndex        =   26
               Top             =   6105
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
               Caption         =   " Net Amount"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookProcessOrder.frx":0514
               Picture         =   "BookProcessOrder.frx":0530
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   10020
               TabIndex        =   27
               Top             =   5790
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
               Caption         =   " VAT"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookProcessOrder.frx":054C
               Picture         =   "BookProcessOrder.frx":0568
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
               Height          =   330
               Left            =   10020
               TabIndex        =   28
               Top             =   6600
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
               Picture         =   "BookProcessOrder.frx":0584
               Picture         =   "BookProcessOrder.frx":05A0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
               Height          =   330
               Left            =   5565
               TabIndex        =   29
               Top             =   6615
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
               Caption         =   " Bill Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookProcessOrder.frx":05BC
               Picture         =   "BookProcessOrder.frx":05D8
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   7080
               TabIndex        =   1
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "BookProcessOrder.frx":05F4
               Caption         =   "BookProcessOrder.frx":070C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookProcessOrder.frx":0778
               Keys            =   "BookProcessOrder.frx":0796
               Spin            =   "BookProcessOrder.frx":07F4
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
               Left            =   11580
               TabIndex        =   2
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "BookProcessOrder.frx":081C
               Caption         =   "BookProcessOrder.frx":0934
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookProcessOrder.frx":09A0
               Keys            =   "BookProcessOrder.frx":09BE
               Spin            =   "BookProcessOrder.frx":0A1C
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
               Left            =   7080
               TabIndex        =   11
               Top             =   6615
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "BookProcessOrder.frx":0A44
               Caption         =   "BookProcessOrder.frx":0B5C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookProcessOrder.frx":0BC8
               Keys            =   "BookProcessOrder.frx":0BE6
               Spin            =   "BookProcessOrder.frx":0C44
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
               Height          =   330
               Left            =   11580
               TabIndex        =   12
               Top             =   6620
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "BookProcessOrder.frx":0C6C
               Caption         =   "BookProcessOrder.frx":0C8C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookProcessOrder.frx":0CF8
               Keys            =   "BookProcessOrder.frx":0D16
               Spin            =   "BookProcessOrder.frx":0D60
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
               ValueVT         =   2005925893
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   30
               Top             =   630
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
               Caption         =   " Item Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookProcessOrder.frx":0D88
               Picture         =   "BookProcessOrder.frx":0DA4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   31
               Top             =   945
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
               Caption         =   " Processor Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookProcessOrder.frx":0DC0
               Picture         =   "BookProcessOrder.frx":0DDC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   32
               Top             =   1260
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
               Picture         =   "BookProcessOrder.frx":0DF8
               Picture         =   "BookProcessOrder.frx":0E14
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   330
               Left            =   11580
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   5790
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "BookProcessOrder.frx":0E30
               Caption         =   "BookProcessOrder.frx":0E50
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookProcessOrder.frx":0EBC
               Keys            =   "BookProcessOrder.frx":0EDA
               Spin            =   "BookProcessOrder.frx":0F24
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
               Height          =   330
               Left            =   11580
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   6105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "BookProcessOrder.frx":0F4C
               Caption         =   "BookProcessOrder.frx":0F6C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookProcessOrder.frx":0FD8
               Keys            =   "BookProcessOrder.frx":0FF6
               Spin            =   "BookProcessOrder.frx":1040
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
               Height          =   330
               Left            =   5565
               TabIndex        =   35
               Top             =   6105
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
               Picture         =   "BookProcessOrder.frx":1068
               Picture         =   "BookProcessOrder.frx":1084
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   3525
               Left            =   120
               TabIndex        =   6
               Top             =   1785
               Width           =   13035
               _Version        =   524288
               _ExtentX        =   22992
               _ExtentY        =   6218
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
               GridColor       =   4227327
               MaxCols         =   13
               MaxRows         =   15
               OperationMode   =   2
               ScrollBars      =   2
               SpreadDesigner  =   "BookProcessOrder.frx":10A0
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
               Left            =   960
               MaxLength       =   100
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   3440
               Width           =   11595
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   285
               Left            =   120
               TabIndex        =   36
               Top             =   5300
               Width           =   13035
               _Version        =   65536
               _ExtentX        =   22992
               _ExtentY        =   494
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
               Picture         =   "BookProcessOrder.frx":1A71
               Picture         =   "BookProcessOrder.frx":1A8D
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   13240
               Y1              =   1680
               Y2              =   1680
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   13240
               Y1              =   6515
               Y2              =   6515
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   13240
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   13240
               Y1              =   5685
               Y2              =   5685
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   8805
            TabIndex        =   39
            Top             =   7590
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
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
            Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  F8->Delete  Ctrl+S->Save"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "BookProcessOrder.frx":1AA9
            Picture         =   "BookProcessOrder.frx":1AC5
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H008BD6FE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Find"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   120
            TabIndex        =   18
            Top             =   7590
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filter"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Preview"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Mail"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "First"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Previous"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Next"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Last"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4
      Left            =   2760
      Top             =   2280
   End
End
Attribute VB_Name = "FrmBookProcessOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnBookOutputOrder As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstBookOOList As New ADODB.Recordset
Dim rstBookOOParent As New ADODB.Recordset
Dim rstBookOOChild As New ADODB.Recordset
Dim rstProcessorList As New ADODB.Recordset
Dim rstBookList As New ADODB.Recordset
Dim rstSizeList As New ADODB.Recordset
Dim rstBindingTypeList As New ADODB.Recordset
Dim BookCode As String, ProcessorCode As String, SizeCode As String, BindingTypeCode As String
Dim EditMode As Boolean
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim oOutlook As New Outlook.Application
Dim OutputTo As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    CxnBookOutputOrder.CursorLocation = adUseClient
    CxnBookOutputOrder.Open cnDatabase.ConnectionString
    rstCompanyMaster.Open "SELECT PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website FROM CompanyMaster", CxnBookOutputOrder, adOpenKeyset, adLockReadOnly
    rstProcessorList.Open "SELECT Name As Col0,P.Code,NegativeOnePcRate,NegativeCutPcRate,NegativePastingRate,PositiveOnePcRate,PositiveCutPcRate,PositivePastingRate FROM AccountMaster P INNER JOIN AccountChild04 C ON P.Code=C.Code ORDER BY Name", CxnBookOutputOrder, adOpenKeyset, adLockReadOnly
    rstBookList.Open "SELECT Name As Col0,Code FROM BookMaster ORDER BY Name", CxnBookOutputOrder, adOpenKeyset, adLockOptimistic
    rstBookOOList.Open "SELECT T.Code,T.Name,T.Date,A.Name As ProcessorName,B.Name As BookName,T.BillAmount FROM (BookOOParent T INNER JOIN AccountMaster A ON T.Processor=A.Code) INNER JOIN BookMaster B ON T.Book=B.Code WHERE FYCode='" & FYCode & "' ORDER BY T.Name", CxnBookOutputOrder, adOpenKeyset, adLockOptimistic
    rstSizeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='1' ORDER BY Name", CxnBookOutputOrder, adOpenKeyset, adLockReadOnly
    rstBindingTypeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='6' ORDER BY Name", CxnBookOutputOrder, adOpenKeyset, adLockReadOnly
    rstBookOOParent.CursorLocation = adUseClient
    rstBookOOList.Filter = adFilterNone
    If rstBookOOList.RecordCount > 0 Then rstBookOOList.MoveLast
    Set DataGrid1.DataSource = rstBookOOList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstBookOOList.EOF Or rstBookOOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstBookOOList.ActiveConnection = Nothing
    rstProcessorList.ActiveConnection = Nothing
    rstBookList.ActiveConnection = Nothing
    rstSizeList.ActiveConnection = Nothing
    rstBindingTypeList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True
    MdiMainMenu.mnuBookProcessOrder.Enabled = False
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If SSTab1.Tab = 0 Then
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If Not EditMode Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
            If Not EditMode Then KeyCode = 0
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF8 And Toolbar1.Buttons.Item(3).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(9)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(10)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyM And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(11)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(13)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(14)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyN And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(15)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(16)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        If Toolbar1.Buttons.Item(1).Enabled Then
            SSTab1.Tab = 1: SSTab1.SetFocus
        Else
           If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBookOOList)
    Call CloseRecordset(rstBookOOParent)
    Call CloseRecordset(rstBookOOChild)
    Call CloseRecordset(rstProcessorList)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstBindingTypeList)
    Call CloseConnection(CxnBookOutputOrder)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuBookProcessOrder.Enabled = True
End Sub
Private Sub Text1_Change()
On Error Resume Next
    With rstBookOOList
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        If Not CheckEmpty(Text1.Text, False) Then
            If SortOrder = "Name" Then .Filter = "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'" Else .Filter = "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
            If .EOF Then
            .Filter = adFilterNone
                .MoveFirst
                If PrevStr <> "" And Len(Text1.Text) > 1 Then If dblBookMark <> 0 Then .Bookmark = dblBookMark Else PrevStr = ""
                Beep
                DisplayError ("Spelling Error")
                Text1.Text = PrevStr
                Sendkeys "{End}"
            Else
                PrevStr = Text1.Text
                dblBookMark = DataGrid1.Bookmark
            End If
        Else
            .Filter = adFilterNone
            PrevStr = ""
        End If
        If Not (.EOF Or .BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
    End With
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean

    If rstBookOOList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstBookOOList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstBookOOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstBookOOList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstBookOOList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstBookOOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstBookOOList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstBookOOList
            .MoveLast
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    End If
    If KeyProcessed Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
        KeyProcessed = False
        KeyCode = 0
    End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If Toolbar1.Buttons.Item(1).Enabled Then
        If SSTab1.Tab = 1 Then
            ViewRecord
        Else
            If Not (rstBookOOList.EOF Or rstBookOOList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            Text1.SetFocus
        End If
        SSTab1.TabEnabled(0) = True
    Else
        SSTab1.TabEnabled(0) = False
        Text2.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer, i As Integer, x As Integer
    Dim CellVal01 As Variant, CellVal02 As Variant, CellVal03 As Variant
    If Button.Index = 1 Then
        If rstBookOOParent.State = adStateOpen Then rstBookOOParent.Close
        rstBookOOParent.Open "SELECT * FROM BookOOParent WHERE Code = ''", CxnBookOutputOrder, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstBookOOParent) Then
            Text2.Text = GenerateCode(CxnBookOutputOrder, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,Name))", "VAL(Name))") & "  FROM BookOOParent WHERE FYCode='" & FYCode & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            CxnBookOutputOrder.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstBookOOList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstBookOOList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnBookOutputOrder.Execute "DELETE FROM BookOOParent WHERE Code = '" & rstBookOOList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstBookOOList.Delete
                rstBookOOList.MoveNext
                If rstBookOOList.RecordCount > 0 And rstBookOOList.EOF Then rstBookOOList.MoveLast
                ShowProgressInStatusBar True
                Timer1.Enabled = True
            Else
                DisplayError ("Failed to delete the record")
            End If
            On Error GoTo 0
        End If
        SetButtons (True)
        SetButtonsForNoRecord
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 4 Then
        If CheckMandatoryFields Then Exit Sub
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstBookOOParent) Then
            If UpdateItemList("D", 0) Then
                UpdateFlag = 1: x = 0
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.SetActiveCell 9, i
                    fpSpread1.GetText 9, i, CellVal01
                    fpSpread1.GetText 12, i, CellVal02
                    fpSpread1.GetText 13, i, CellVal03
                    If Val(CellVal01) <> 0 And CellVal02 <> "" And CellVal03 <> "" Then
                        x = x + 1
                        If Not UpdateItemList("I", x) Then UpdateFlag = 0: Exit For
                    End If
                Next
            End If
        End If
        If UpdateFlag Then
            AddToList
            CxnBookOutputOrder.CommitTrans
            If rstBookOOParent.State = adStateOpen Then rstBookOOParent.Close
            rstBookOOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            Call MsgBox("Record updated !!!", vbInformation, App.Title)
            SSTab1.Tab = 0
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstBookOOParent) Then
            CxnBookOutputOrder.RollbackTrans
            If rstBookOOParent.State = adStateOpen Then rstBookOOParent.Close
            rstBookOOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstBookOOList.Filter = adFilterNone
        rstBookOOList.ActiveConnection = CxnBookOutputOrder
        Do While Not RefreshRecord(rstBookOOList)
        Loop
        Set DataGrid1.DataSource = rstBookOOList
        rstBookOOList.ActiveConnection = Nothing
        If rstBookOOList.RecordCount > 0 Then rstBookOOList.MoveLast
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Processor", 0
            .Combo1.AddItem "Book", 1
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstBookOOList.RecordCount = 0 Then Exit Sub
        OutputTo = "P"
        PrintBookProcessOrder rstBookOOList.Fields("Code").Value
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstBookOOList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        PrintBookProcessOrder rstBookOOList.Fields("Code").Value
        HiLiteRecord = True
    ElseIf Button.Index = 11 Then
        If rstBookOOList.RecordCount = 0 Then Exit Sub
        OutputTo = "M"
        PrintBookProcessOrder rstBookOOList.Fields("Code").Value
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstBookOOList.RecordCount > 0 Then rstBookOOList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstBookOOList.RecordCount > 0 Then
            rstBookOOList.MovePrevious
            If rstBookOOList.BOF Then rstBookOOList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstBookOOList.RecordCount > 0 Then
            rstBookOOList.MoveNext
            If rstBookOOList.EOF Then rstBookOOList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstBookOOList.RecordCount > 0 Then rstBookOOList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstBookOOList.EOF Or rstBookOOList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        Text1.SetFocus
    End If
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstBookOOList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstBookOOList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstBookOOList.EOF Or rstBookOOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub SetButtons(bVal As Boolean)
    Toolbar1.Buttons.Item(1).Enabled = bVal
    Toolbar1.Buttons.Item(2).Enabled = bVal
    Toolbar1.Buttons.Item(3).Enabled = bVal
    Toolbar1.Buttons.Item(4).Enabled = Not bVal
    Toolbar1.Buttons.Item(5).Enabled = Not bVal
    Toolbar1.Buttons.Item(6).Enabled = bVal
    Toolbar1.Buttons.Item(7).Enabled = bVal
    Toolbar1.Buttons.Item(9).Enabled = bVal
    Toolbar1.Buttons.Item(10).Enabled = bVal
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstBookOOList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(9).Enabled = False
        Toolbar1.Buttons.Item(10).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstBookOOParent.EOF Or rstBookOOParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnBookOutputOrder, "BookOOParent", "Code", "[Name]", Trim(Text2.Text), rstBookOOParent.Fields("Code").Value, False, FYCode) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf Not blnRecordExist Then
        MhDateInput3.Text = Format(DateAdd("d", 1, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
    End If
End Sub
Private Sub Text5_Change()
    If Text5.Text = " " Then Text5.Text = "?": Sendkeys "{TAB}"
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text5.Text)
    If rstBookList.RecordCount = 0 Then DisplayError ("No Record in Book Master"): Cancel = True: Exit Sub Else rstBookList.MoveFirst
    rstBookList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstBookList.EOF Then
        SelectionType = "S"
        BookCode = ""
        Call LoadSelectionList(rstBookList, "List of Books...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text5, BookCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text5.Text, False) Then Text5.Text = "?"
        If RTrim(BookCode) <> "" Then Sendkeys "{TAB}"
        Cancel = True
    Else
        BookCode = rstBookList.Fields("Code").Value
    End If
End Sub
Private Sub Text3_Change()
    If Text3.Text = " " Then Text3.Text = "?": Sendkeys "{TAB}"
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text3.Text)
    If rstProcessorList.RecordCount = 0 Then DisplayError ("No Record in Processor Master"): Cancel = True: Exit Sub Else rstProcessorList.MoveFirst
    rstProcessorList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstProcessorList.EOF Then
        SelectionType = "S"
        ProcessorCode = ""
        Call LoadSelectionList(rstProcessorList, "List of Processors...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, ProcessorCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then Text3.Text = "?"
        If RTrim(ProcessorCode) <> "" Then Sendkeys "{TAB}"
        Cancel = True
    Else
        ProcessorCode = rstProcessorList.Fields("Code").Value
    End If
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If MhDateInput2.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput2.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput3_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput3.Text)) Then Cancel = True
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)   'VAT (%)
    MhRealInput12.Value = MhRealInput19.Value * MhRealInput11.Value / 100   'VAT
    Call CalculateTotal("N")    'VAT Changed
End Sub
Private Sub MhRealInput14_Validate(Cancel As Boolean)   'Adjustment
    Call CalculateTotal("N")    'Adjustment Changed
End Sub
Private Sub CalculateTotal(ByVal strType As String)
    Dim Amt As Variant, i As Integer
    If strType = "G" Then   'Calculate VAT
        MhRealInput19.Value = 0
        With fpSpread1
            For i = 1 To .DataRowCnt
                .GetText 9, i, Amt
                MhRealInput19.Value = MhRealInput19.Value + Amt
            Next
        End With
        MhRealInput12.Value = MhRealInput19.Value * MhRealInput11.Value / 100   'GST
    Else
        MhRealInput15.Value = Round(MhRealInput19.Value + MhRealInput12.Value + MhRealInput14.Value, 0)
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstBookOOList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstBookOOParent.State = adStateOpen Then rstBookOOParent.Close
    rstBookOOParent.Open "SELECT * FROM BookOOParent WHERE Code = '" & FixQuote(rstBookOOList.Fields("Code").Value) & "'", CxnBookOutputOrder, adOpenKeyset, adLockOptimistic
    If rstBookOOParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput2.Text = "  -  -    "    'Bill Date
    MhDateInput3.Text = Format(DateAdd("d", 1, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")    'Delivery Date
    MhRealInput19.Value = 0 'Total Gross Amount
    MhRealInput11.Value = 0 'VAT (%)
    MhRealInput12.Value = 0 'VAT
    MhRealInput14.Value = 0 'VAT (%)
    MhRealInput15.Value = 0 'VAT
    MhRealInput16.Value = 0
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
End Sub
Private Sub LoadFields()
    If rstBookOOParent.EOF Or rstBookOOParent.BOF Then Exit Sub
    Text2.Text = rstBookOOParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstBookOOParent.Fields("Date").Value, "dd-MM-yyyy")
    MhDateInput3.Text = Format(rstBookOOParent.Fields("DeliveryDate").Value, "dd-MM-yyyy")
    ProcessorCode = rstBookOOParent.Fields("Processor").Value
    If rstProcessorList.RecordCount > 0 Then rstProcessorList.MoveFirst
    rstProcessorList.Find "[Code] = '" & ProcessorCode & "'"
    If Not rstProcessorList.EOF Then Text3.Text = rstProcessorList.Fields("Col0").Value
    BookCode = rstBookOOParent.Fields("Book").Value
    If rstBookList.RecordCount > 0 Then rstBookList.MoveFirst
    rstBookList.Find "[Code] = '" & BookCode & "'"
    If Not rstBookList.EOF Then Text5.Text = rstBookList.Fields("Col0").Value
    Text4.Text = rstBookOOParent.Fields("Remarks").Value
    MhRealInput11.Value = rstBookOOParent.Fields("VAT%").Value
    MhRealInput12.Value = rstBookOOParent.Fields("VAT").Value
    MhRealInput14.Value = rstBookOOParent.Fields("Adjustment").Value
    TxtAdNar.Text = rstBookOOParent.Fields("AdjustmentRemarks").Value
    MhRealInput15.Value = rstBookOOParent.Fields("BillAmount").Value
    Text8.Text = rstBookOOParent.Fields("BillNo").Value
    If Not IsNull(rstBookOOParent.Fields("BillDate").Value) Then MhDateInput2.Text = Format(rstBookOOParent.Fields("BillDate").Value, "dd-MM-yyyy")
    MhRealInput16.Value = rstBookOOParent.Fields("PaidAmount").Value
    Call LoadItemList(rstBookOOParent.Fields("Code").Value)
    CalculateTotal ("G")
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstBookOOParent.RecordCount = 0 Then Exit Sub
    If rstBookOOParent.State = adStateOpen Then rstBookOOParent.Close
    rstBookOOParent.CursorLocation = adUseServer
    rstBookOOParent.Open "SELECT * FROM BookOOParent WHERE Code = '" & FixQuote(rstBookOOList.Fields("Code").Value) & "'", CxnBookOutputOrder, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstBookOOParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    If AllowTransactionsModification = 0 Then
        If Not CheckEmpty(Text8.Text, False) Then LockFields (True)
        Text1.Locked = False
    End If
    CxnBookOutputOrder.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstBookOOParent.EOF Or rstBookOOParent.BOF Then Exit Sub
    Dim lpBuff As String * 1024
    If Not blnRecordExist Then
        rstBookOOParent.Fields("Code").Value = GenerateCode(CxnBookOutputOrder, "SELECT MAX(Code) FROM BookOOParent", 6, "0")
        rstBookOOParent.Fields("CreatedBy").Value = UserCode
        rstBookOOParent.Fields("CreatedOn").Value = Now()
        rstBookOOParent.Fields("Recordstatus").Value = "N"
    Else
        rstBookOOParent.Fields("ModifiedBy").Value = UserCode
        rstBookOOParent.Fields("ModifiedOn").Value = Now()
        rstBookOOParent.Fields("Recordstatus").Value = "M"
    End If
    rstBookOOParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstBookOOParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstBookOOParent.Fields("DeliveryDate").Value = GetDate(MhDateInput3.Text)
    rstBookOOParent.Fields("Book").Value = BookCode
    rstBookOOParent.Fields("Processor").Value = ProcessorCode
    rstBookOOParent.Fields("VAT%").Value = MhRealInput11.Value
    rstBookOOParent.Fields("VAT").Value = MhRealInput12.Value
    rstBookOOParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstBookOOParent.Fields("BillNo").Value = Trim(Text8.Text)
    If Not IsDate(MhDateInput2.Text) Then rstBookOOParent.Fields("BillDate").Value = Null Else rstBookOOParent.Fields("BillDate").Value = GetDate(MhDateInput2.Text)
    rstBookOOParent.Fields("Adjustment").Value = MhRealInput14.Value
    rstBookOOParent.Fields("BillAmount").Value = MhRealInput15.Value
    rstBookOOParent.Fields("PaidAmount").Value = MhRealInput16.Value
    rstBookOOParent.Fields("AdjustmentRemarks").Value = IIf(MhRealInput14.Value <> 0, TxtAdNar.Text, "")
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookOOParent.Fields("BillFeedDate").Value) Then rstBookOOParent.Fields("BillFeedDate").Value = Now()
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookOOParent.Fields("ComputerName").Value) Then rstBookOOParent.Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    rstBookOOParent.Fields("FYCode").Value = FYCode
    rstBookOOParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstBookOOList.MoveFirst
    rstBookOOList.Find "[Code] = '" & rstBookOOParent.Fields("Code").Value & "'"
    If rstBookOOList.EOF Then
       rstBookOOList.AddNew
       rstBookOOList.Fields("Code").Value = rstBookOOParent.Fields("Code").Value
    End If
    rstBookOOList.Fields("Name").Value = Pad(rstBookOOParent.Fields("Name").Value, Space(1), 10, "L")
    rstBookOOList.Fields("Date").Value = rstBookOOParent.Fields("Date").Value
    rstProcessorList.MoveFirst
    rstProcessorList.Find "[Code] = '" & rstBookOOParent.Fields("Processor").Value & "'"
    rstBookOOList.Fields("ProcessorName").Value = Trim(rstProcessorList.Fields("Col0").Value)
    rstProcessorList.MoveFirst
    rstBookList.Find "[Code] = '" & rstBookOOParent.Fields("Book").Value & "'"
    rstBookOOList.Fields("BookName").Value = Trim(rstBookList.Fields("Col0").Value)
    rstBookOOList.Update
    rstBookOOList.Sort = SortOrder & " Asc"
    rstBookOOList.Find "[Code] = '" & rstBookOOParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Order No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
       Text3.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstProcessorList, ProcessorCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text5.Text, False) Then
       Text5.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text5, "Col0", rstBookList, BookCode) Then
        Text5.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnBookOutputOrder, "BookOOParent", "Code", "[Name]", Trim(Text2.Text), rstBookOOParent.Fields("Code").Value, False, FYCode) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    End If
    If MhRealInput14.Value <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
    If MhRealInput16.Value <> 0 Then If MhRealInput16.Value <> MhRealInput15.Value Then MhRealInput14.SetFocus: CheckMandatoryFields = True: Exit Function
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Processor" Then
        rstBookOOList.Filter = "[ProcessorName] Like '%" & SrchText & "%'"
    ElseIf SrchFor = "Book" Then
        rstBookOOList.Filter = "[BookName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1: fpSpread1.SetFocus
            CalculateTotal ("G"): CalculateTotal ("N")
        End If
    ElseIf KeyCode = vbKeySpace Then
        Dim CurrentText As Variant
        With fpSpread1
            If .ActiveCol = 2 Then
                .GetText .ActiveCol, .ActiveRow, CurrentText
                Text6.Text = FixQuote(CurrentText)
                If rstSizeList.RecordCount = 0 Then DisplayError ("No Record in Size Master"): .SetActiveCell 2, .ActiveRow: .SetFocus: Exit Sub Else rstSizeList.MoveFirst
                rstSizeList.Find "[Col0] = '" & RTrim(CurrentText) & "'"
                SelectionType = "S"
                SizeCode = ""
                Call LoadSelectionList(rstSizeList, "List of Sizes...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text6, SizeCode)
                Call CloseForm(FrmSelectionList)
                If SizeCode = "" Then
                    .SetActiveCell .ActiveCol, .ActiveRow
                Else
                    rstSizeList.MoveFirst: rstSizeList.Find "[Code] ='" & SizeCode & "'"
                    .SetText .ActiveCol, .ActiveRow, Text6.Text
                    .SetText .ActiveCol + 10, .ActiveRow, SizeCode
                    Sendkeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 10 Then
                .GetText .ActiveCol, .ActiveRow, CurrentText
                Text6.Text = FixQuote(CurrentText)
                If rstBindingTypeList.RecordCount = 0 Then DisplayError ("No Record in Binding Type Master"): .SetActiveCell 10, .ActiveRow: .SetFocus: Exit Sub Else rstBindingTypeList.MoveFirst
                rstBindingTypeList.Find "[Col0] = '" & RTrim(CurrentText) & "'"
                SelectionType = "S"
                BindingTypeCode = ""
                Call LoadSelectionList(rstBindingTypeList, "List of Binding Types...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text6, BindingTypeCode)
                Call CloseForm(FrmSelectionList)
                If BindingTypeCode = "" Then
                    .SetActiveCell .ActiveCol, .ActiveRow
                Else
                    rstBindingTypeList.MoveFirst: rstBindingTypeList.Find "[Code] ='" & BindingTypeCode & "'"
                    .SetText .ActiveCol, .ActiveRow, Text6.Text
                    .SetText .ActiveCol + 3, .ActiveRow, BindingTypeCode
                    Sendkeys "{ENTER}"
                End If
            End If
        End With
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim OutputSize As Variant, Forms As Variant, Colors As Variant, Rate As Variant, ProcType As Variant, OutputFormat As Variant, xPos As Integer
    With fpSpread1
        If Col <> 9 And Col <> 11 Then
            .GetText Col, Row, OutputSize: If OutputSize = "" Then Cancel = True: Exit Sub
        End If
        .GetText 2, Row, OutputSize
        .GetText 4, Row, Forms
        .GetText 5, Row, Colors
        .GetText 6, Row, ProcType
        .GetText 7, Row, OutputFormat
        .GetText 8, Row, Rate
        If Col = 6 Or Col = 7 Then
            If ProcType <> "" And OutputFormat <> "" And Rate = 0 Then
                If ProcessorCode <> "" Then
                    rstProcessorList.MoveFirst
                    rstProcessorList.Find "[Code] = '" & ProcessorCode & "'"
                    If Not rstProcessorList.EOF Then
                        Rate = Val(rstProcessorList.Fields(ProcType + Replace(OutputFormat, Space(1), "") + "Rate").Value)
                        .SetText 8, Row, Rate
                    End If
                End If
            End If
        End If
        If OutputSize <> "" Then xPos = InStr(1, LCase(OutputSize), "x", vbTextCompare): .SetText 9, Row, Left(OutputSize, xPos - 1) * Mid(OutputSize, xPos + 1) * Forms * Colors * Rate: CalculateTotal ("G"): CalculateTotal ("N")
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub LoadItemList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstBookOOChild.State = adStateOpen Then rstBookOOChild.Close
    rstBookOOChild.Open "SELECT OutputType,OutputSize,S.Name As SizeName,Ups,Forms,Colors,ProcessingType,OutputFormat,Rate,Amount,BindingType,B.Name As BindingTypeName,Remarks FROM (BookOOChild T INNER JOIN GeneralMaster S ON T.OutputSize=S.Code) INNER JOIN GeneralMaster B ON T.BindingType=B.Code WHERE T.Code = '" & strOrderCode & "' ORDER BY SerialNo", CxnBookOutputOrder, adOpenKeyset, adLockOptimistic
    rstBookOOChild.ActiveConnection = Nothing
    If rstBookOOChild.RecordCount > 0 Then rstBookOOChild.MoveFirst
    i = 0
    Do While Not rstBookOOChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, IIf(rstBookOOChild.Fields("OutputType").Value = "1", "Text", "Title")
            .SetText 2, i, rstBookOOChild.Fields("SizeName").Value
            .SetText 3, i, Val(rstBookOOChild.Fields("Ups").Value)
            .SetText 4, i, Val(rstBookOOChild.Fields("Forms").Value)
            .SetText 5, i, Val(rstBookOOChild.Fields("Colors").Value)
            .SetText 6, i, IIf(rstBookOOChild.Fields("ProcessingType").Value = "1", "Positive", "Negative")
            .SetText 7, i, IIf(rstBookOOChild.Fields("OutputFormat").Value = "1", "One Pc", IIf(rstBookOOChild.Fields("OutputFormat").Value = "1", "Cut Pc", "Pasting"))
            .SetText 8, i, Val(rstBookOOChild.Fields("Rate").Value)
            .SetText 9, i, Val(rstBookOOChild.Fields("Amount").Value)
            .SetText 10, i, rstBookOOChild.Fields("BindingTypeName").Value
            .SetText 11, i, rstBookOOChild.Fields("Remarks").Value
            .SetText 12, i, rstBookOOChild.Fields("OutputSize").Value
            .SetText 13, i, rstBookOOChild.Fields("BindingType").Value
        End With
        rstBookOOChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Item List")
End Sub
Private Function UpdateItemList(ByVal ActionType As String, ByVal SerialNo As Integer) As Boolean
    Dim CellVal(1 To 11) As Variant
    On Error GoTo ErrorHandler
    UpdateItemList = True
    If ActionType = "D" Then
        If Not blnRecordExist Then Exit Function
        CxnBookOutputOrder.Execute "DELETE FROM BookOOChild WHERE Code = '" & rstBookOOParent.Fields("Code").Value & "'"
    Else
        With fpSpread1
            .GetText 1, .ActiveRow, CellVal(1)      'Output Type
            .GetText 12, .ActiveRow, CellVal(2)     'Output Size
            .GetText 3, .ActiveRow, CellVal(3)      'Ups
            .GetText 4, .ActiveRow, CellVal(4)      'Forms
            .GetText 5, .ActiveRow, CellVal(5)      'Colors
            .GetText 6, .ActiveRow, CellVal(6)      'Processing Type
            .GetText 7, .ActiveRow, CellVal(7)      'Output Format
            .GetText 8, .ActiveRow, CellVal(8)      'Rate
            .GetText 9, .ActiveRow, CellVal(9)      'Amount
            .GetText 13, .ActiveRow, CellVal(10)    'Binding Type
            .GetText 11, .ActiveRow, CellVal(11)    'Remarks
        End With
        CxnBookOutputOrder.Execute "INSERT INTO BookOOChild VALUES ('" & rstBookOOParent.Fields("Code").Value & "','" & IIf(CellVal(1) = "Text", "1", "2") & "','" & CellVal(2) & "'," & Val(CellVal(3)) & "," & Val(CellVal(4)) & "," & Val(CellVal(5)) & ",'" & IIf(CellVal(6) = "Positive", "1", "2") & "','" & IIf(CellVal(7) = "One Pc", "1", IIf(CellVal(7) = "Cut Pc", "2", "3")) & "'," & Val(CellVal(8)) & "," & Val(CellVal(9)) & ",'" & CellVal(10) & "','" & CellVal(11) & "'," & SerialNo & ")"
    End If
    Exit Function
ErrorHandler:
    UpdateItemList = False
End Function
Private Sub PrintBookProcessOrder(ByVal OrderCode As String)
    Dim rstCompanyMaster As New ADODB.Recordset, rstBookProcessOrder As New ADODB.Recordset, Prefix As String
    Dim oOutlookMsg As Outlook.MailItem, FileName As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Prefix = "OO/" + Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/"
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,eMail FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    rstBookProcessOrder.Open "SELECT '" & Prefix & "'+LTRIM(P.Name) As OrderNo,Date,DeliveryDate As TargetDate,B.PrintName As BookName,A.PrintName As ProcessName,[VAT%],VAT,Adjustment,BillAmount,P.Remarks As OrderRemarks,IIF(OutputType='1','Text','Title') As OPType,G1.Name As OutputSizeName,Ups,C.Forms,Colors,IIF(ProcessingType='1','Positive','Negative') As PRType,IIF(OutputFormat='1','One Piece',IIF(OutputFormat='2','Cut Piece','Pasting')) As OPFormat,Rate,Amount,G2.Name As BindingTypeName,C.Remarks As ItemRemarks,LTRIM(A.eMail) As ProcessorMail FROM ((((BookOOParent P INNER JOIN BookOOChild C ON P.Code=C.Code) INNER JOIN BookMaster B ON P.Book=B.Code) INNER JOIN AccountMaster A ON P.Processor=A.Code) INNER JOIN GeneralMaster G1 ON C.OutputSize=G1.Code) INNER JOIN GeneralMaster G2 ON C.BindingType=G2.Code WHERE P.Code='" & OrderCode & "' ORDER BY C.SerialNo", cnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    rstBookProcessOrder.ActiveConnection = Nothing
    rptBookProcessOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptBookProcessOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) + Space(1) + Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptBookProcessOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    rptBookProcessOrder.Text28.SetText " (" & Trim(NumberToWords(rstBookProcessOrder.Fields("BillAmount").Value, True)) & ")"
    rptBookProcessOrder.Text27.SetText "for " & Trim(rstBookProcessOrder.Fields("ProcessName").Value)
    rptBookProcessOrder.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptBookProcessOrder.Database.SetDataSource rstBookProcessOrder, 3, 1
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptBookProcessOrder
        FrmReportViewer.Show vbModal
    ElseIf OutputTo = "P" Then
        rptBookProcessOrder.PaperSource = crPRBinAuto
        rptBookProcessOrder.PrintOut False    'Print Report Without Prompt
    Else
        Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
        With oOutlookMsg
            .To = rstBookProcessOrder.Fields("ProcessorMail").Value
            .Subject = "Book Process Order #" & Trim(rstBookProcessOrder.Fields("OrderNo").Value)
            .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith PO #" & Trim(rstBookProcessOrder.Fields("OrderNo").Value) & " for doing the needful at your end. An early execution of the order will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of execution of order.<Br><Br>Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
            rptBookProcessOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
            rptBookProcessOrder.ExportOptions.DestinationType = crEDTDiskFile
            FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
            rptBookProcessOrder.ExportOptions.DiskFileName = FileName
            rptBookProcessOrder.Export False
            .Attachments.Add (FileName)
            .Importance = olImportanceHigh
            .ReadReceiptRequested = True
            If CheckEmpty(.To, False) Then .Display Else .Send
        End With
        Set oOutlookMsg = Nothing
    End If
    Set rptBookProcessOrder = Nothing
    Call CloseRecordset(rstBookProcessOrder): Call CloseRecordset(rstCompanyMaster)
    On Error GoTo 0
End Sub
Private Sub LockFields(ByVal bVal As Boolean)
    Dim O As Object
    For Each O In Me
        If TypeName(O) = "TextBox" Then
            O.Locked = bVal
        ElseIf TypeName(O) = "TDBNumber" Then
            O.ReadOnly = bVal
        ElseIf TypeName(O) = "fpSpread" Then
            O.Enabled = Not bVal
        End If
    Next
End Sub
