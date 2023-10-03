VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmPaperPurchaseOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Purchase Order"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17745
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
   ScaleHeight     =   8985
   ScaleWidth      =   17745
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   8970
      Left            =   15
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   17715
      _Version        =   65536
      _ExtentX        =   31247
      _ExtentY        =   15822
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
      Picture         =   "PaperPurchaseOrder.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8745
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   120
         Width           =   17490
         _ExtentX        =   30850
         _ExtentY        =   15425
         _Version        =   393216
         Style           =   1
         Tabs            =   2
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
         TabPicture(0)   =   "PaperPurchaseOrder.frx":001C
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
         TabPicture(1)   =   "PaperPurchaseOrder.frx":0038
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
            Left            =   720
            MaxLength       =   40
            TabIndex        =   25
            Top             =   8310
            Width           =   12105
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7785
            Left            =   120
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   450
            Width           =   17265
            _ExtentX        =   30454
            _ExtentY        =   13732
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
            ColumnCount     =   4
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
               DataField       =   "SupplierName"
               Caption         =   "Supplier Name"
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
            BeginProperty Column03 
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
                  ColumnWidth     =   1049.953
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   13110.24
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  Locked          =   -1  'True
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8160
            Left            =   -74880
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   480
            Width           =   17265
            _Version        =   65536
            _ExtentX        =   30454
            _ExtentY        =   14393
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
            Picture         =   "PaperPurchaseOrder.frx":0054
            Begin VB.CommandButton cmdDelete 
               Height          =   375
               Left            =   16780
               Picture         =   "PaperPurchaseOrder.frx":0070
               Style           =   1  'Graphical
               TabIndex        =   63
               TabStop         =   0   'False
               ToolTipText     =   "Delete Item Pic"
               Top             =   75
               Width           =   375
            End
            Begin VB.CommandButton cmdUpload 
               Height          =   375
               Left            =   16015
               Picture         =   "PaperPurchaseOrder.frx":0172
               Style           =   1  'Graphical
               TabIndex        =   58
               TabStop         =   0   'False
               ToolTipText     =   "Upload Bill"
               Top             =   75
               Width           =   375
            End
            Begin VB.CommandButton cmdView 
               Height          =   375
               Left            =   16405
               Picture         =   "PaperPurchaseOrder.frx":04B4
               Style           =   1  'Graphical
               TabIndex        =   57
               TabStop         =   0   'False
               ToolTipText     =   "View Bill"
               Top             =   75
               Width           =   375
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
               Left            =   1560
               MaxLength       =   30
               TabIndex        =   18
               Top             =   7725
               Width           =   1530
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
               Left            =   9330
               MaxLength       =   40
               TabIndex        =   11
               Top             =   6360
               Width           =   1575
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
               Left            =   1560
               MaxLength       =   255
               TabIndex        =   12
               Top             =   6885
               Width           =   1530
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
               Left            =   1560
               MaxLength       =   30
               TabIndex        =   15
               Top             =   7200
               Width           =   1530
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
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   0
               Top             =   105
               Width           =   1530
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
               Left            =   1560
               MaxLength       =   40
               TabIndex        =   4
               Top             =   945
               Width           =   15595
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
               Left            =   1560
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   3
               Top             =   630
               Width           =   15595
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   28
               Top             =   7200
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
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
               Picture         =   "PaperPurchaseOrder.frx":09E6
               Picture         =   "PaperPurchaseOrder.frx":0A02
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   29
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
               Caption         =   " Order No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0A1E
               Picture         =   "PaperPurchaseOrder.frx":0A3A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   7605
               TabIndex        =   30
               Top             =   105
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "PaperPurchaseOrder.frx":0A56
               Picture         =   "PaperPurchaseOrder.frx":0A72
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   120
               TabIndex        =   31
               Top             =   6045
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
               Caption         =   " GST (%)"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0A8E
               Picture         =   "PaperPurchaseOrder.frx":0AAA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   32
               Top             =   630
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
               Caption         =   " Supplier Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0AC6
               Picture         =   "PaperPurchaseOrder.frx":0AE2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   13620
               TabIndex        =   33
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
               Picture         =   "PaperPurchaseOrder.frx":0AFE
               Picture         =   "PaperPurchaseOrder.frx":0B1A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   34
               Top             =   950
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0B36
               Picture         =   "PaperPurchaseOrder.frx":0B52
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   35
               Top             =   5730
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
               Caption         =   " Cartage/Kg"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0B6E
               Picture         =   "PaperPurchaseOrder.frx":0B8A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   14625
               TabIndex        =   36
               Top             =   6360
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
               Caption         =   " Net Amount"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0BA6
               Picture         =   "PaperPurchaseOrder.frx":0BC2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   120
               TabIndex        =   37
               Top             =   6360
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
               Caption         =   " Adjustment"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0BDE
               Picture         =   "PaperPurchaseOrder.frx":0BFA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   14625
               TabIndex        =   38
               Top             =   5730
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
               Caption         =   " Total Cartage"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0C16
               Picture         =   "PaperPurchaseOrder.frx":0C32
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   14625
               TabIndex        =   39
               Top             =   6045
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
               Caption         =   " GST"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0C4E
               Picture         =   "PaperPurchaseOrder.frx":0C6A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
               Height          =   330
               Left            =   14625
               TabIndex        =   40
               Top             =   7200
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
               Caption         =   " Paid Amount"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0C86
               Picture         =   "PaperPurchaseOrder.frx":0CA2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
               Height          =   330
               Left            =   7605
               TabIndex        =   41
               Top             =   7200
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "PaperPurchaseOrder.frx":0CBE
               Picture         =   "PaperPurchaseOrder.frx":0CDA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Left            =   120
               TabIndex        =   42
               Top             =   6885
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
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
               Caption         =   " E-Way Bill No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0CF6
               Picture         =   "PaperPurchaseOrder.frx":0D12
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
               Height          =   330
               Left            =   7605
               TabIndex        =   43
               Top             =   6885
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Delivery Start Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0D2E
               Picture         =   "PaperPurchaseOrder.frx":0D4A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
               Height          =   330
               Left            =   14625
               TabIndex        =   44
               Top             =   6885
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
               Caption         =   " Dlv End Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":0D66
               Picture         =   "PaperPurchaseOrder.frx":0D82
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
               Height          =   330
               Left            =   7605
               TabIndex        =   45
               Top             =   6360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "PaperPurchaseOrder.frx":0D9E
               Picture         =   "PaperPurchaseOrder.frx":0DBA
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput3 
               Height          =   330
               Left            =   14820
               TabIndex        =   2
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "PaperPurchaseOrder.frx":0DD6
               Caption         =   "PaperPurchaseOrder.frx":0EEE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":0F5A
               Keys            =   "PaperPurchaseOrder.frx":0F78
               Spin            =   "PaperPurchaseOrder.frx":0FD6
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
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   9330
               TabIndex        =   1
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "PaperPurchaseOrder.frx":0FFE
               Caption         =   "PaperPurchaseOrder.frx":1116
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":1182
               Keys            =   "PaperPurchaseOrder.frx":11A0
               Spin            =   "PaperPurchaseOrder.frx":11FE
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
               Height          =   1640
               Left            =   120
               TabIndex        =   5
               Top             =   1470
               Width           =   17040
               _Version        =   524288
               _ExtentX        =   30057
               _ExtentY        =   2893
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
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "PaperPurchaseOrder.frx":1226
            End
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   1640
               Left            =   120
               TabIndex        =   6
               Top             =   3630
               Width           =   17040
               _Version        =   524288
               _ExtentX        =   30057
               _ExtentY        =   2893
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
               MaxCols         =   14
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "PaperPurchaseOrder.frx":1EE7
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput4 
               Height          =   330
               Left            =   9330
               TabIndex        =   13
               Top             =   6885
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "PaperPurchaseOrder.frx":2BEA
               Caption         =   "PaperPurchaseOrder.frx":2D02
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":2D6E
               Keys            =   "PaperPurchaseOrder.frx":2D8C
               Spin            =   "PaperPurchaseOrder.frx":2DEA
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
               Left            =   9330
               TabIndex        =   16
               Top             =   7200
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "PaperPurchaseOrder.frx":2E12
               Caption         =   "PaperPurchaseOrder.frx":2F2A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":2F96
               Keys            =   "PaperPurchaseOrder.frx":2FB4
               Spin            =   "PaperPurchaseOrder.frx":3012
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
            Begin TDBDate6Ctl.TDBDate MhDateInput5 
               Height          =   330
               Left            =   15945
               TabIndex        =   14
               Top             =   6885
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   582
               Calendar        =   "PaperPurchaseOrder.frx":303A
               Caption         =   "PaperPurchaseOrder.frx":3152
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":31BE
               Keys            =   "PaperPurchaseOrder.frx":31DC
               Spin            =   "PaperPurchaseOrder.frx":323A
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
               Left            =   15945
               TabIndex        =   17
               Top             =   7200
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrder.frx":3262
               Caption         =   "PaperPurchaseOrder.frx":3282
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":32EE
               Keys            =   "PaperPurchaseOrder.frx":330C
               Spin            =   "PaperPurchaseOrder.frx":3356
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
               ForeColor       =   -2147483640
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
               Left            =   1560
               TabIndex        =   7
               Top             =   5730
               Width           =   1530
               _Version        =   65536
               _ExtentX        =   2699
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrder.frx":337E
               Caption         =   "PaperPurchaseOrder.frx":339E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":340A
               Keys            =   "PaperPurchaseOrder.frx":3428
               Spin            =   "PaperPurchaseOrder.frx":3472
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
               ForeColor       =   -2147483640
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
               Left            =   1560
               TabIndex        =   9
               Top             =   6045
               Width           =   1530
               _Version        =   65536
               _ExtentX        =   2699
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrder.frx":349A
               Caption         =   "PaperPurchaseOrder.frx":34BA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":3526
               Keys            =   "PaperPurchaseOrder.frx":3544
               Spin            =   "PaperPurchaseOrder.frx":358E
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
               ValueVT         =   1845493765
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   330
               Left            =   15945
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   6045
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrder.frx":35B6
               Caption         =   "PaperPurchaseOrder.frx":35D6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":3642
               Keys            =   "PaperPurchaseOrder.frx":3660
               Spin            =   "PaperPurchaseOrder.frx":36AA
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
               Height          =   330
               Left            =   15945
               TabIndex        =   8
               Top             =   5730
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrder.frx":36D2
               Caption         =   "PaperPurchaseOrder.frx":36F2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":375E
               Keys            =   "PaperPurchaseOrder.frx":377C
               Spin            =   "PaperPurchaseOrder.frx":37C6
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
               ForeColor       =   -2147483640
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
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
               Height          =   330
               Left            =   1560
               TabIndex        =   10
               Top             =   6360
               Width           =   1530
               _Version        =   65536
               _ExtentX        =   2699
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrder.frx":37EE
               Caption         =   "PaperPurchaseOrder.frx":380E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":387A
               Keys            =   "PaperPurchaseOrder.frx":3898
               Spin            =   "PaperPurchaseOrder.frx":38E2
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
               ForeColor       =   -2147483640
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
               MinValue        =   -9999999.99
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
               Left            =   15945
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   6360
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrder.frx":390A
               Caption         =   "PaperPurchaseOrder.frx":392A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":3996
               Keys            =   "PaperPurchaseOrder.frx":39B4
               Spin            =   "PaperPurchaseOrder.frx":39FE
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   48
               Top             =   3090
               Width           =   17040
               _Version        =   65536
               _ExtentX        =   30048
               _ExtentY        =   503
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
               Picture         =   "PaperPurchaseOrder.frx":3A26
               Picture         =   "PaperPurchaseOrder.frx":3A42
               Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
                  Height          =   285
                  Left            =   14085
                  TabIndex        =   50
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1005
                  _Version        =   65536
                  _ExtentX        =   1782
                  _ExtentY        =   503
                  Calculator      =   "PaperPurchaseOrder.frx":3A5E
                  Caption         =   "PaperPurchaseOrder.frx":3A7E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperPurchaseOrder.frx":3AEA
                  Keys            =   "PaperPurchaseOrder.frx":3B08
                  Spin            =   "PaperPurchaseOrder.frx":3B52
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
               Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
                  Height          =   285
                  Left            =   10545
                  TabIndex        =   51
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
                  _ExtentY        =   503
                  Calculator      =   "PaperPurchaseOrder.frx":3B7A
                  Caption         =   "PaperPurchaseOrder.frx":3B9A
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperPurchaseOrder.frx":3C06
                  Keys            =   "PaperPurchaseOrder.frx":3C24
                  Spin            =   "PaperPurchaseOrder.frx":3C6E
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "#####0.000"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "#####0.000"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   999999.999
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
                  Height          =   285
                  Left            =   11520
                  TabIndex        =   59
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1035
                  _Version        =   65536
                  _ExtentX        =   1834
                  _ExtentY        =   503
                  Calculator      =   "PaperPurchaseOrder.frx":3C96
                  Caption         =   "PaperPurchaseOrder.frx":3CB6
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperPurchaseOrder.frx":3D22
                  Keys            =   "PaperPurchaseOrder.frx":3D40
                  Spin            =   "PaperPurchaseOrder.frx":3D8A
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
                  Height          =   285
                  Left            =   15900
                  TabIndex        =   60
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   885
                  _Version        =   65536
                  _ExtentX        =   1570
                  _ExtentY        =   503
                  Calculator      =   "PaperPurchaseOrder.frx":3DB2
                  Caption         =   "PaperPurchaseOrder.frx":3DD2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperPurchaseOrder.frx":3E3E
                  Keys            =   "PaperPurchaseOrder.frx":3E5C
                  Spin            =   "PaperPurchaseOrder.frx":3EA6
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
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   285
               Left            =   120
               TabIndex        =   49
               Top             =   5250
               Width           =   17040
               _Version        =   65536
               _ExtentX        =   30048
               _ExtentY        =   503
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
               Picture         =   "PaperPurchaseOrder.frx":3ECE
               Picture         =   "PaperPurchaseOrder.frx":3EEA
               Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
                  Height          =   285
                  Left            =   13520
                  TabIndex        =   52
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
                  _ExtentY        =   503
                  Calculator      =   "PaperPurchaseOrder.frx":3F06
                  Caption         =   "PaperPurchaseOrder.frx":3F26
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperPurchaseOrder.frx":3F92
                  Keys            =   "PaperPurchaseOrder.frx":3FB0
                  Spin            =   "PaperPurchaseOrder.frx":3FFA
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "#####0.000"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "#####0.000"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   999999.999
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
               Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
                  Height          =   285
                  Left            =   15510
                  TabIndex        =   53
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   665
                  _Version        =   65536
                  _ExtentX        =   1173
                  _ExtentY        =   503
                  Calculator      =   "PaperPurchaseOrder.frx":4022
                  Caption         =   "PaperPurchaseOrder.frx":4042
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperPurchaseOrder.frx":40AE
                  Keys            =   "PaperPurchaseOrder.frx":40CC
                  Spin            =   "PaperPurchaseOrder.frx":4116
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
                  ForeColor       =   255
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
                  ReadOnly        =   1
                  Separator       =   ""
                  ShowContextMenu =   1
                  ValueVT         =   5
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
                  Height          =   285
                  Left            =   14490
                  TabIndex        =   61
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1035
                  _Version        =   65536
                  _ExtentX        =   1826
                  _ExtentY        =   503
                  Calculator      =   "PaperPurchaseOrder.frx":413E
                  Caption         =   "PaperPurchaseOrder.frx":415E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperPurchaseOrder.frx":41CA
                  Keys            =   "PaperPurchaseOrder.frx":41E8
                  Spin            =   "PaperPurchaseOrder.frx":4232
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
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
               Height          =   330
               Left            =   120
               TabIndex        =   54
               Top             =   7725
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
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
               Caption         =   " Bilty No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":425A
               Picture         =   "PaperPurchaseOrder.frx":4276
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   330
               Left            =   14625
               TabIndex        =   55
               Top             =   7725
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
               Caption         =   " Bilty Amount"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":4292
               Picture         =   "PaperPurchaseOrder.frx":42AE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
               Height          =   330
               Left            =   7605
               TabIndex        =   56
               Top             =   7725
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Bilty Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "PaperPurchaseOrder.frx":42CA
               Picture         =   "PaperPurchaseOrder.frx":42E6
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput6 
               Height          =   330
               Left            =   9330
               TabIndex        =   19
               Top             =   7725
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "PaperPurchaseOrder.frx":4302
               Caption         =   "PaperPurchaseOrder.frx":441A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":4486
               Keys            =   "PaperPurchaseOrder.frx":44A4
               Spin            =   "PaperPurchaseOrder.frx":4502
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
               Height          =   330
               Left            =   15945
               TabIndex        =   20
               Top             =   7725
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrder.frx":452A
               Caption         =   "PaperPurchaseOrder.frx":454A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrder.frx":45B6
               Keys            =   "PaperPurchaseOrder.frx":45D4
               Spin            =   "PaperPurchaseOrder.frx":461E
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
               ForeColor       =   -2147483640
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
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin VB.Line Line6 
               X1              =   0
               X2              =   17240
               Y1              =   7620
               Y2              =   7620
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   17240
               Y1              =   6780
               Y2              =   6780
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   17240
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   17240
               Y1              =   1370
               Y2              =   1370
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   17240
               Y1              =   3490
               Y2              =   3490
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   17240
               Y1              =   5640
               Y2              =   5640
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   12810
            TabIndex        =   62
            Top             =   8310
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
            Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D->Delete  Ctrl+S->Save"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "PaperPurchaseOrder.frx":4646
            Picture         =   "PaperPurchaseOrder.frx":4662
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
            TabIndex        =   26
            Top             =   8310
            Width           =   615
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   17745
      _ExtentX        =   31300
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
   Begin MSComDlg.CommonDialog cdUpload 
      Left            =   6600
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmPaperPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnPaperPurchaseOrder As New ADODB.Connection
Dim rstPaperPOList As New ADODB.Recordset, rstPaperPOParent As New ADODB.Recordset, rstPaperPOChild As New ADODB.Recordset, rstPaperList As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstLastPurchaseRate As New ADODB.Recordset, srmPicMgr As New ADODB.Stream
Dim SupplierCode As String, AccountCode As Variant, PaperCode As Variant
Dim SortOrder, PrevStr
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim oOutlook As New Outlook.Application
Dim EditMode As Boolean
Dim EMailID, Attachment, Message, imgFile As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    cnPaperPurchaseOrder.CursorLocation = adUseClient
    cnPaperPurchaseOrder.Open cnDatabase.ConnectionString
    rstPaperPOList.Open "SELECT T.Code,T.Name,Date,M.Name As SupplierName,BillAmount FROM PaperPOParent T INNER JOIN AccountMaster M ON T.Supplier=M.Code WHERE OrderType='P' AND FYCode='" & FYCode & "' ORDER BY T.Name", cnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstPaperPOParent.CursorLocation = adUseClient
    rstPaperPOList.Filter = adFilterNone
    If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
    Set DataGrid1.DataSource = rstPaperPOList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
        End If
    rstPaperPOList.ActiveConnection = Nothing
    LoadMasterList
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    MdiMainMenu.mnuPaperModule(1).Enabled = False
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
           If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstPaperPOList)
    Call CloseRecordset(rstPaperPOParent)
    Call CloseRecordset(rstPaperPOChild)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstLastPurchaseRate)
    Call CloseConnection(cnPaperPurchaseOrder)
    If srmPicMgr.State = adStateOpen Then srmPicMgr.Close
    Set srmPicMgr = Nothing
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuPaperModule(1).Enabled = True
End Sub
Private Sub Text1_Change()
On Error Resume Next
    With rstPaperPOList
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
    If rstPaperPOList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstPaperPOList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstPaperPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstPaperPOList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstPaperPOList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstPaperPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstPaperPOList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstPaperPOList
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
            If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
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
        Text3.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    Dim CellVal01 As Variant, CellVal02 As Variant, CellVal03 As Variant, CellVal04 As Variant, i As Integer
    If Button.Index = 1 Then
        If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
        rstPaperPOParent.Open "SELECT * FROM PaperPOParent WHERE Code=''", cnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstPaperPOParent) Then
            Text2.Text = GenerateCode(cnPaperPurchaseOrder, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,Name))", "VAL(Name))") & "  FROM PaperPOParent WHERE OrderType='P' AND FYCode='" & FYCode & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text3.SetFocus
            blnRecordExist = False
            cnPaperPurchaseOrder.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
                cnPaperPurchaseOrder.Execute "DELETE FROM PaperPOParent WHERE Code IN (Select Distinct Code From PaperIOChild Where (Ref= '" & rstPaperPOList.Fields("Code").Value & "' Or Code= '" & rstPaperPOList.Fields("Code").Value & "'))"
                cnPaperPurchaseOrder.Execute "DELETE FROM PaperIOChild WHERE Code IN (Select Distinct Code From PaperIOChild Where (Ref= '" & rstPaperPOList.Fields("Code").Value & "' Or Code= '" & rstPaperPOList.Fields("Code").Value & "'))"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstPaperPOList.Delete
                rstPaperPOList.MoveNext
                If rstPaperPOList.RecordCount > 0 And rstPaperPOList.EOF Then rstPaperPOList.MoveLast
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
        If UpdateRecord(rstPaperPOParent) Then
            If UpdatePaperList("D") Then
                UpdateFlag = 1
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 9, i
                        .GetText 9, i, CellVal01    'Amount
                        .GetText 12, i, CellVal02   'Paper
                        If Val(CellVal01) <> 0 And CellVal02 <> "" Then
                            If Not UpdatePaperList("I1") Then UpdateFlag = 0: Exit For
                        End If
                    Next
                End With
                If UpdateFlag = 1 Then
                    With fpSpread2
                        For i = 1 To .DataRowCnt
                            .SetActiveCell 3, i
                            .GetText 3, i, CellVal01    'Quantity
                            .GetText 8, i, CellVal03    'Printer
                            .GetText 9, i, CellVal02    'Paper
                            .GetText 14, i, CellVal04 'VchCode
                            If Val(CellVal01) <> 0 And CellVal02 <> "" And CellVal03 <> "" And (CheckEmpty(CellVal04, False) Or CellVal04 = rstPaperPOParent.Fields("Code").Value) Then If Not UpdatePaperList("I2") Then UpdateFlag = 0: Exit For
                        Next
                    End With
                End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnPaperPurchaseOrder.CommitTrans
            If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
            rstPaperPOParent.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstPaperPOParent) Then
            cnPaperPurchaseOrder.RollbackTrans
            If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
            rstPaperPOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstPaperPOList.Filter = adFilterNone
        rstPaperPOList.ActiveConnection = cnPaperPurchaseOrder
        Do While Not RefreshRecord(rstPaperPOList)
        Loop
        Set DataGrid1.DataSource = rstPaperPOList
        rstPaperPOList.ActiveConnection = Nothing
        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Supplier", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        Call PrintPaperPurchaseOrder(rstPaperPOList.Fields("Code").Value, "", "P")
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        Call PrintPaperPurchaseOrder(rstPaperPOList.Fields("Code").Value, "", "S")
        HiLiteRecord = True
    ElseIf Button.Index = 11 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        Call PrintPaperPurchaseOrder(rstPaperPOList.Fields("Code").Value, "", "M")
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstPaperPOList.RecordCount > 0 Then
            rstPaperPOList.MovePrevious
            If rstPaperPOList.BOF Then rstPaperPOList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstPaperPOList.RecordCount > 0 Then
            rstPaperPOList.MoveNext
            If rstPaperPOList.EOF Then rstPaperPOList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
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
        rstPaperPOList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstPaperPOList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
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
    Toolbar1.Buttons.Item(11).Enabled = bVal
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstPaperPOList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(9).Enabled = False
        Toolbar1.Buttons.Item(10).Enabled = False
        Toolbar1.Buttons.Item(11).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstPaperPOParent.EOF Or rstPaperPOParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnPaperPurchaseOrder, "PaperPOParent", "Code", "[Name]+[OrderType]", Trim(Text2.Text) & "P", rstPaperPOParent.Fields("Code").Value, False, FYCode) Then
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
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = SupplierCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        SupplierCode = slCode: Text3.Text = slName
        If Not CheckEmpty(SupplierCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If CheckEmpty(Text3.Text, False) Then Cancel = True
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If MhDateInput2.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput2.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput3_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput3.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput4_Validate(Cancel As Boolean)
    If MhDateInput4.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput4.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput5_Validate(Cancel As Boolean)
    If MhDateInput5.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput5.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput6_Validate(Cancel As Boolean)
    If MhDateInput6.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput6.Text)) Then Cancel = True
End Sub
Private Sub MhRealInput10_Validate(Cancel As Boolean)   'Cartage/Kg
    On Error Resume Next
    If MhRealInput10.Value <> 0 Then MhRealInput13.Value = MhRealInput18.Value * MhRealInput10.Value    'Total Cartage
    MhRealInput13_Validate (False)
End Sub
Private Sub MhRealInput13_Validate(Cancel As Boolean)   'Cartage
    If Not blnRecordExist Then MhRealInput22.Value = MhRealInput13.Value
    Call CalculateTotal("N")    'Cartage Changed
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)   'GST (%)
    Call CalculateTotal("N")    'GST Changed
End Sub
Private Sub MhRealInput14_Validate(Cancel As Boolean)   'Adjustment
    MhRealInput11_Validate (False)
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstPaperPOList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
    rstPaperPOParent.Open "SELECT * FROM PaperPOParent WHERE Code='" & FixQuote(rstPaperPOList.Fields("Code").Value) & "'", cnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    If rstPaperPOParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput2.Text = "  -  -    "    'Bill Date
    MhDateInput3.Text = Format(DateAdd("d", 1, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")    'Delivery Date
    MhDateInput4.Text = "  -  -    "    'Delivery Start Date
    MhDateInput5.Text = "  -  -    "    'Delivery End Date
    MhDateInput6.Text = "  -  -    "    'Bilty Date
    MhRealInput18.Value = 0 'Total Quantity (Kg)
    MhRealInput19.Value = 0 'Total Gross Amount
    MhRealInput10.Value = 0 'Cartage/Kg
    MhRealInput11.Value = 12    'GST (%)
    MhRealInput12.Value = 0 'GST
    MhRealInput13.Value = 0 'Total Cartage
    MhRealInput14.Value = 0 'Adjustment
    MhRealInput15.Value = 0 'Net Amount
    MhRealInput16.Value = 0 'Paid Amount
    MhRealInput20.Value = 0 'Total Quantity (Ream) - To be issued
    MhRealInput21.Value = 0 'Total Tat
    MhRealInput22.Value = 0 'Bilty Amount
    TxtAdNar.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    fpSpread2.ClearRange 1, 1, fpSpread2.MaxCols, fpSpread2.MaxRows, True: fpSpread2.SetActiveCell 1, 1
    AccountCode = "": SupplierCode = "": imgFile = "": cmdUpload.Enabled = True
End Sub
Private Sub LoadFields()
    With rstPaperPOParent
        If .EOF Or .BOF Then Exit Sub
        Text2.Text = .Fields("Name").Value
        MhDateInput1.Text = Format(.Fields("Date").Value, "dd-MM-yyyy")
        MhDateInput3.Text = Format(.Fields("DeliveryDate").Value, "dd-MM-yyyy")
        SupplierCode = .Fields("Supplier").Value
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        rstAccountList.Find "[Code] = '" & SupplierCode & "'"
        If Not rstAccountList.EOF Then Text3.Text = rstAccountList.Fields("Col0").Value
        Text4.Text = .Fields("Remarks").Value
        MhRealInput10.Value = Val(.Fields("Cartage/Kg").Value)
        MhRealInput11.Value = Val(.Fields("GST%").Value)
        MhRealInput12.Value = Val(.Fields("GST").Value)
        MhRealInput13.Value = Val(.Fields("Cartage").Value)
        MhRealInput14.Value = Val(.Fields("Adjustment").Value)
        MhRealInput15.Value = Val(.Fields("BillAmount").Value)
        Text8.Text = .Fields("BillNo").Value
        Text9.Text = .Fields("ChallanNo").Value
        If Not IsNull(.Fields("BillDate").Value) Then MhDateInput2.Text = Format(.Fields("BillDate").Value, "dd-MM-yyyy")
        If Not IsNull(.Fields("DeliveryStartDate").Value) Then MhDateInput4.Text = Format(.Fields("DeliveryStartDate").Value, "dd-MM-yyyy")
        If Not IsNull(.Fields("DeliveryEndDate").Value) Then MhDateInput5.Text = Format(.Fields("DeliveryEndDate").Value, "dd-MM-yyyy")
        MhRealInput16.Value = Val(.Fields("PaidAmount").Value)
        TxtAdNar.Text = .Fields("AdjustmentRemarks").Value
        Text5.Text = .Fields("BiltyNo").Value
        If Not IsNull(.Fields("BiltyDate").Value) Then MhDateInput6.Text = Format(.Fields("BiltyDate").Value, "dd-MM-yyyy")
        MhRealInput22.Value = Val(.Fields("BiltyAmount").Value)
        If Dir(App.Path & "\Pic\", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Pic\"
        If Dir(App.Path & "\Pic\PPO" & FinancialYear & CompCode, vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Pic\PPO" & FinancialYear & CompCode
        If Not CheckEmpty(.Fields("PicData"), False) Then imgFile = App.Path & "\Pic\PPO" & FinancialYear & CompCode & "\" & FinancialYear & CompCode & .Fields("Code").Value & "." & .Fields("PicType").Value: RetrievePic .Fields("PicData").Value, imgFile, srmPicMgr: cmdUpload.Enabled = False
        Call LoadPaperList(.Fields("Code").Value)
    End With
    CalculateTotal ("G")
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstPaperPOParent.RecordCount = 0 Then Exit Sub
    If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
    rstPaperPOParent.CursorLocation = adUseServer
    rstPaperPOParent.Open "SELECT * FROM PaperPOParent WHERE Code='" & FixQuote(rstPaperPOList.Fields("Code").Value) & "'", cnPaperPurchaseOrder, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstPaperPOParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text3.SetFocus
    blnRecordExist = True
    If AllowTransactionsModification = 0 Then
        If Not CheckEmpty(Text8.Text, False) Then LockFields (True)
        Text1.Locked = False
    End If
    cnPaperPurchaseOrder.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    With rstPaperPOParent
        If .EOF Or .BOF Then Exit Sub
        Dim lpBuff As String * 1024
        GetComputerName lpBuff, Len(lpBuff)
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnPaperPurchaseOrder, "SELECT MAX(Code) FROM PaperPOParent", 6, "0")
            .Fields("CreatedBy").Value = UserCode
            .Fields("CreatedOn").Value = Now()
            .Fields("Recordstatus").Value = "N"
        Else
            .Fields("ModifiedBy").Value = UserCode
            .Fields("ModifiedOn").Value = Now()
            .Fields("Recordstatus").Value = "M"
        End If
        If Not CheckEmpty(imgFile, False) Then
            If srmPicMgr.State = adStateOpen Then srmPicMgr.Close
            srmPicMgr.Type = adTypeBinary
            srmPicMgr.Open
            srmPicMgr.LoadFromFile imgFile
            If srmPicMgr.Size > 0 Then .Fields("PicData").Value = srmPicMgr.Read: .Fields("PicType").Value = UCase(FSO.GetExtensionName(FSO.GetFileName(imgFile))) Else .Fields("PicData").Value = Null: .Fields("PicType").Value = Null
        Else
            .Fields("PicData").Value = Null: .Fields("PicType").Value = Null
        End If
        .Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
        .Fields("Date").Value = GetDate(MhDateInput1.Text)
        .Fields("Supplier").Value = SupplierCode
        .Fields("DeliveryDate").Value = GetDate(MhDateInput3.Text)
        .Fields("Remarks").Value = Trim(Text4.Text)
        .Fields("Cartage/Kg").Value = MhRealInput10.Value
        .Fields("GST%").Value = MhRealInput11.Value
        .Fields("GST").Value = MhRealInput12.Value
        .Fields("Cartage").Value = MhRealInput13.Value
        .Fields("Adjustment").Value = MhRealInput14.Value
        .Fields("BillAmount").Value = MhRealInput15.Value
        .Fields("BillNo").Value = Trim(Text8.Text)
        .Fields("ChallanNo").Value = Trim(Text9.Text)
        If Not IsDate(MhDateInput2.Text) Then .Fields("BillDate").Value = Null Else .Fields("BillDate").Value = GetDate(MhDateInput2.Text)
        If Not IsDate(MhDateInput4.Text) Then .Fields("DeliveryStartDate").Value = Null Else .Fields("DeliveryStartDate").Value = GetDate(MhDateInput4.Text)
        If Not IsDate(MhDateInput5.Text) Then .Fields("DeliveryEndDate").Value = Null Else .Fields("DeliveryEndDate").Value = GetDate(MhDateInput5.Text)
        .Fields("PaidAmount").Value = MhRealInput16.Value
        .Fields("AdjustmentRemarks").Value = IIf(MhRealInput14.Value <> 0, TxtAdNar.Text, "")
        .Fields("BiltyNo").Value = Trim(Text5.Text)
        If Not IsDate(MhDateInput6.Text) Then .Fields("BiltyDate").Value = Null Else .Fields("BiltyDate").Value = GetDate(MhDateInput6.Text)
        .Fields("BiltyAmount").Value = MhRealInput22.Value
        If Not CheckEmpty(Text8.Text, False) Then If IsNull(.Fields("BillFeedDate").Value) Then .Fields("BillFeedDate").Value = Now()
        If Not CheckEmpty(Text8.Text, False) Then If IsNull(.Fields("ComputerName").Value) Then .Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
        .Fields("OrderType").Value = "P"
        .Fields("FYCode").Value = FYCode
        .Fields("PrintStatus").Value = "N"
    End With
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstPaperPOList.MoveFirst
    rstPaperPOList.Find "[Code] = '" & rstPaperPOParent.Fields("Code").Value & "'"
    If rstPaperPOList.EOF Then rstPaperPOList.AddNew
    rstPaperPOList.Fields("Code").Value = rstPaperPOParent.Fields("Code").Value
    rstPaperPOList.Fields("Name").Value = Pad(rstPaperPOParent.Fields("Name").Value, Space(1), 10, "L")
    rstPaperPOList.Fields("Date").Value = rstPaperPOParent.Fields("Date").Value
    rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & rstPaperPOParent.Fields("Supplier").Value & "'"
    rstPaperPOList.Fields("SupplierName").Value = Trim(rstAccountList.Fields("Col0").Value)
    rstPaperPOList.Fields("BillAmount").Value = rstPaperPOParent.Fields("BillAmount").Value
    rstPaperPOList.Update
    rstPaperPOList.Sort = SortOrder & " Asc"
    rstPaperPOList.Find "[Code] = '" & rstPaperPOParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        DisplayError ("Order No. cannot be blank")
        Text2.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnPaperPurchaseOrder, "PaperPOParent", "Code", "[Name]+[OrderType]", Trim(Text2.Text) & "P", rstPaperPOParent.Fields("Code").Value, False, FYCode) Then
        Text2.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf Not chkPaper() Then
        fpSpread2.SetFocus
        CheckMandatoryFields = True: Exit Function
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
Private Sub LoadPaperList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstPaperPOChild.State = adStateOpen Then rstPaperPOChild.Close
    rstPaperPOChild.Open "SELECT M.Name As PaperName,Quantity,U.Name As UOMName,T.[Weight/Unit],QuantityKg,QuantitySheets,T.[Rate/Kg],T.[Rate/Unit],Amount,T.[Units/Bundle],TotalBundles,Paper As PaperCode,U.Value1 As SPU FROM (PaperPOChild T INNER JOIN PaperMaster M ON T.Paper=M.Code) INNER JOIN GeneralMaster U ON M.UOM=U.Code WHERE T.Code='" & strOrderCode & "' ORDER BY M.Name", cnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstPaperPOChild.ActiveConnection = Nothing
    If rstPaperPOChild.RecordCount > 0 Then rstPaperPOChild.MoveFirst
    i = 0
    Do While Not rstPaperPOChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstPaperPOChild.Fields("PaperName").Value
            .SetText 2, i, Val(rstPaperPOChild.Fields("Quantity").Value)
            .SetText 3, i, rstPaperPOChild.Fields("UOMName").Value
            .SetText 4, i, Val(rstPaperPOChild.Fields("Weight/Unit").Value)
            .SetText 5, i, Val(rstPaperPOChild.Fields("QuantityKg").Value)
            .SetText 6, i, Val(rstPaperPOChild.Fields("QuantitySheets").Value)
            .SetText 7, i, Val(rstPaperPOChild.Fields("Rate/Kg").Value)
            .SetText 8, i, Val(rstPaperPOChild.Fields("Rate/Unit").Value)
            .SetText 9, i, Val(rstPaperPOChild.Fields("Amount").Value)
            .SetText 10, i, Val(rstPaperPOChild.Fields("Units/Bundle").Value)
            .SetText 11, i, Val(rstPaperPOChild.Fields("TotalBundles").Value)
            .SetText 12, i, rstPaperPOChild.Fields("PaperCode").Value
            .SetText 13, i, Val(rstPaperPOChild.Fields("SPU").Value)
        End With
        rstPaperPOChild.MoveNext
    Loop
    If rstPaperPOChild.State = adStateOpen Then rstPaperPOChild.Close
    rstPaperPOChild.Open "SELECT Paper As PaperCode,M1.Name As PaperName,Account As AccountCode,M2.Name As AccountName,Quantity,QuantitySheets,QuantityKg,T.[Units/Bundle],TotalBundles,Narration,T.[Weight/Unit],U.Value1 As SPU,Ref,T.Code As VchCode FROM ((PaperIOChild T INNER JOIN PaperMaster M1 ON T.Paper=M1.Code) INNER JOIN AccountMaster M2 ON T.Account=M2.Code) INNER JOIN GeneralMaster U ON M1.UOM=U.Code WHERE T.Ref='" & strOrderCode & "' UNION ALL SELECT Paper As PaperCode,M1.Name As PaperName,Account As AccountCode,M2.Name As AccountName,Quantity,QuantitySheets,QuantityKg,T.[Units/Bundle],TotalBundles,Narration,T.[Weight/Unit],U.Value1 As SPU,Ref,T.Code As VchCode FROM ((PaperIOChild T INNER JOIN PaperMaster M1 ON T.Paper=M1.Code) INNER JOIN AccountMaster M2 ON T.Account=M2.Code) INNER JOIN GeneralMaster U ON M1.UOM=U.Code WHERE T.Code='" & strOrderCode & "' AND (Ref IS NULL OR Ref='') ORDER BY PaperName,AccountName", cnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstPaperPOChild.ActiveConnection = Nothing
    If rstPaperPOChild.RecordCount > 0 Then rstPaperPOChild.MoveFirst
    i = 0
    Do While Not rstPaperPOChild.EOF
        i = i + 1
        With fpSpread2
            .SetText 1, i, rstPaperPOChild.Fields("PaperName").Value
            .SetText 2, i, rstPaperPOChild.Fields("AccountName").Value
            .SetText 3, i, Val(rstPaperPOChild.Fields("Quantity").Value)
            .SetText 4, i, Val(rstPaperPOChild.Fields("QuantityKg").Value)
            .SetText 5, i, Val(rstPaperPOChild.Fields("QuantitySheets").Value)
            .SetText 6, i, Val(rstPaperPOChild.Fields("TotalBundles").Value)
            .SetText 7, i, IIf(IsNull(rstPaperPOChild.Fields("Ref").Value), 0, 1)
            .SetText 8, i, rstPaperPOChild.Fields("AccountCode").Value
            .SetText 9, i, rstPaperPOChild.Fields("PaperCode").Value
            .SetText 10, i, rstPaperPOChild.Fields("Narration").Value
            .SetText 11, i, Val(rstPaperPOChild.Fields("Weight/Unit").Value)
            .SetText 12, i, Val(rstPaperPOChild.Fields("Units/Bundle").Value)
            .SetText 13, i, Val(rstPaperPOChild.Fields("SPU").Value)
            .SetText 14, i, rstPaperPOChild.Fields("VchCode").Value
        End With
        rstPaperPOChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Paper List")
End Sub
Private Sub CalculateTotal(ByVal strType As String)
    Dim Qty01 As Variant, Qty02 As Variant, Amt As Variant, Bdls As Variant, i As Integer
    If strType = "G" Then   'Calculate Cartage & GST
        MhRealInput18.Value = 0: MhRealInput19.Value = 0: MhRealInput23.Value = 0: MhRealInput24.Value = 0: Bdls = 0
        With fpSpread1
            For i = 1 To .DataRowCnt
                .GetText 5, i, Qty01: .GetText 6, i, Qty02: .GetText 9, i, Amt: .GetText 10, i, Bdls
                MhRealInput18.Value = MhRealInput18.Value + Qty01
                MhRealInput23.Value = MhRealInput23.Value + Qty02
                MhRealInput19.Value = MhRealInput19.Value + Amt
                MhRealInput24.Value = MhRealInput24.Value + Bdls
            Next
        End With
        MhRealInput20.Value = 0: MhRealInput25.Value = 0: MhRealInput21.Value = 0
        With fpSpread2
            For i = 1 To .DataRowCnt
                .GetText 4, i, Qty01: .GetText 5, i, Qty02: .GetText 6, i, Bdls
                MhRealInput20.Value = MhRealInput20.Value + Qty01
                MhRealInput25.Value = MhRealInput25.Value + Qty02
                MhRealInput21.Value = MhRealInput21.Value + Bdls
            Next
        End With
        MhRealInput10_Validate False 'Calculate Cartage
    Else
        MhRealInput12.Value = (MhRealInput19.Value + MhRealInput13.Value + MhRealInput14.Value) * MhRealInput11.Value / 100 'GST
        MhRealInput15.Value = Round(MhRealInput19.Value + MhRealInput12.Value + MhRealInput13.Value + MhRealInput14.Value, 0)
    End If
End Sub
Private Function GetLastPurchaseRate() As String
    On Error GoTo ErrorHandler
    If rstLastPurchaseRate.State = adStateOpen Then rstLastPurchaseRate.Close
    rstLastPurchaseRate.Open "SELECT TOP 1 [Rate/Kg],[Rate/Unit] FROM PaperPOParent P INNER JOIN PaperPOChild C ON P.Code=C.Code WHERE Paper='" & PaperCode & "' AND P.Code < '" & IIf(IsNull(rstPaperPOParent.Fields("Code").Value), "999999", rstPaperPOParent.Fields("Code").Value) & "' ORDER BY P.Name DESC", cnPaperPurchaseOrder, adOpenKeyset, adLockReadOnly
    If rstLastPurchaseRate.RecordCount > 0 Then GetLastPurchaseRate = "Kg-" & Trim(rstLastPurchaseRate.Fields("Rate/Kg").Value) & Space(1) & "Unit-" & Trim(rstLastPurchaseRate.Fields("Rate/Unit").Value)
    Exit Function
ErrorHandler:
    DisplayError ("Failed to fetch Last Purchase Rate")
End Function
Private Function UpdatePaperList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 10) As Variant
    On Error GoTo ErrorHandler
    UpdatePaperList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        cnPaperPurchaseOrder.Execute "DELETE FROM PaperPOChild WHERE Code='" & rstPaperPOParent.Fields("Code").Value & "'"
        cnPaperPurchaseOrder.Execute "DELETE FROM PaperIOChild WHERE Code='" & rstPaperPOParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I1" Then
        With fpSpread1
            .GetText 2, .ActiveRow, CellVal(1)  'Qty in Units
            .GetText 4, .ActiveRow, CellVal(10)  'Weight/Unit
            .GetText 5, .ActiveRow, CellVal(2)  'Qty in Weight
            .GetText 6, .ActiveRow, CellVal(3)  'Qty in Sheets
            .GetText 7, .ActiveRow, CellVal(4)  'Rate/Kg
            .GetText 8, .ActiveRow, CellVal(5)  'Rate/Unit
            .GetText 9, .ActiveRow, CellVal(6)  'Amount
            .GetText 10, .ActiveRow, CellVal(7)  'Units/Bundle
            .GetText 11, .ActiveRow, CellVal(8) 'Total Bundles
            .GetText 12, .ActiveRow, CellVal(9) 'Paper
        End With
        cnPaperPurchaseOrder.Execute "INSERT INTO PaperPOChild VALUES ('" & rstPaperPOParent.Fields("Code").Value & "','" & CellVal(9) & "'," & Val(CellVal(10)) & "," & Val(CellVal(1)) & "," & Val(CellVal(3)) & "," & Val(CellVal(2)) & "," & Val(CellVal(4)) & "," & Val(CellVal(5)) & "," & Val(CellVal(6)) & "," & Val(CellVal(7)) & "," & Val(CellVal(8)) & ")"
    Else
        With fpSpread2
            .GetText 3, .ActiveRow, CellVal(1)  'Quantity
            .GetText 4, .ActiveRow, CellVal(2)  'Qty in Weight
            .GetText 5, .ActiveRow, CellVal(3)  'Qty in Sheets
            .GetText 6, .ActiveRow, CellVal(4)  'Bundles
            .GetText 7, .ActiveRow, CellVal(10)  'Issued/Planned
            .GetText 8, .ActiveRow, CellVal(5)  'Account
            .GetText 9, .ActiveRow, CellVal(6)  'Paper
            .GetText 10, .ActiveRow, CellVal(7)  'Narration
            .GetText 11, .ActiveRow, CellVal(9) 'Weight/Unit
            .GetText 12, .ActiveRow, CellVal(8) 'Units/Bundle
        End With
        cnPaperPurchaseOrder.Execute "INSERT INTO PaperIOChild VALUES ('" & rstPaperPOParent.Fields("Code").Value & "','" & CellVal(6) & "','" & CellVal(5) & "'," & Val(CellVal(9)) & "," & Val(CellVal(1)) & "," & Val(CellVal(3)) & "," & Val(CellVal(2)) & "," & Val(CellVal(8)) & "," & Val(CellVal(4)) & ",'" & CellVal(7) & "'," & IIf(Val(CellVal(10)), "'" & rstPaperPOParent.Fields("Code").Value & "'", "Null") & ")"
    End If
    Exit Function
ErrorHandler:
    UpdatePaperList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Supplier" Then rstPaperPOList.Filter = "[SupplierName] Like '%" & SrchText & "%'"
End Sub
Public Sub PrintPaperPurchaseOrder(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String)
    Dim rstCompanyMaster As New ADODB.Recordset, rstPurchaseOrder As New ADODB.Recordset, rstPurchaseOrderChild As New ADODB.Recordset, Prefix As String
    Dim oOutlookMsg As Outlook.MailItem, FileName As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Prefix = "PPO/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/"
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,eMail,GSTIN FROM CompanyMaster Where FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    rstPurchaseOrder.Open "SELECT ('" & Prefix & "'+LTRIM(P.Name)) As OrderNo,[Date] As OrderDate,DeliveryDate,LTRIM(M1.PrintName) As SupplierName,[GST%],GST,P.Cartage,Adjustment,BillAmount,Remarks,LTRIM(M2.PrintName) As PaperName,'',Quantity As QuantityOther,C.[Weight/Unit],QuantityKg,C.[Rate/Kg],(SELECT TOP 1 '" & Prefix & "'+LTRIM(P1.Name)+'/'+FORMAT(P1.Date,'dd-MM-yyyy')+'/'+FORMAT([Rate/Kg],'0.00') FROM PaperPOParent P1 INNER JOIN PaperPOChild C1 ON P1.Code=C1.Code WHERE C1.Paper=C.Paper AND P1.Code<P.Code ORDER BY P1.Name DESC) As LastPurchaseRate,Amount,LTRIM(eMail) As SupplierMail,(SELECT TOP 1 LTRIM([Rate/Kg])+'-'+P1.Supplier FROM PaperPOParent P1 INNER JOIN PaperPOChild C1 ON P1.Code=C1.Code WHERE C1.Paper=C.Paper AND P1.Code<P.Code ORDER BY P1.Name DESC) As LastSupplier,P.Supplier As CurrentSupplier,'('+LTRIM(G.PrintName)+'='+LTRIM(G.Value1)+')' As UOM,Form " & _
                                                      "FROM (((PaperPOParent P LEFT JOIN PaperPOChild C ON P.Code=C.Code) LEFT JOIN AccountMaster M1 ON M1.Code=P.Supplier) LEFT JOIN PaperMaster M2 ON M2.Code=C.Paper) LEFT JOIN GeneralMaster G ON M2.UOM=G.Code WHERE P.Code='" & OrderCode & "' ORDER BY M2.PrintName", cnDatabase, adOpenKeyset, adLockOptimistic
    rstPurchaseOrderChild.Open "SELECT ('" & Prefix & "'+LTRIM(P.Name)) As OrderNo,[Date] As OrderDate,LTRIM(M3.PrintName) As Godown,LTRIM(M2.PrintName) As PaperName,LTRIM(M1.PrintName) As PrinterName,'' As RefNo,IIF(Form='S',Quantity,QuantityKg) As Quantity,0 As Tat,'' As Remarks,M1.Address1 As PrinterAdd1,M1.Address2 As PrinterAdd2,M1.Address3 As PrinterAdd3,M1.Address4 As PrinterAdd4,LTRIM(M1.eMail) As PrinterMail,M1.TIN As GSTIN,M1.Mobile,IIF(Form='S',LTRIM(G.PrintName),'Kilograms') As UOM,M2.Form,C.QuantityKg FROM ((((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN AccountMaster M3 ON P.Supplier=M3.Code) INNER JOIN GeneralMaster G ON G.Code=M2.UOM WHERE P.Code='" & OrderCode & "' AND (P.Code=Ref OR Ref IS NULL OR Ref='') ORDER BY M2.PrintName", cnDatabase, adOpenKeyset, adLockOptimistic
    'rstPurchaseOrderChild.Open "SELECT '" & Prefix & "'+LTRIM(P.Name) As OrderNo,[Date] As OrderDate,LTRIM(M3.PrintName) As Godown,LTRIM(M2.PrintName) As PaperName,LTRIM(M1.PrintName) As PrinterName,'' As RefNo,IIF(Form='S',Quantity,QuantityKg) As Quantity,0 As Tat,'' As Remarks,M1.Address1 As PrinterAdd1,M1.Address2 As PrinterAdd2,M1.Address3 As PrinterAdd3,M1.Address4 As PrinterAdd4,LTRIM(M1.eMail) As PrinterMail,M1.TIN As GSTIN,M1.Mobile,IIF(Form='S','('+LTRIM(G.PrintName)+'='+LTRIM(G.Value1)+')','Kg') As UOM,M2.Form,C.QuantityKg FROM ((((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN AccountMaster M3 ON P.Supplier=M3.Code) INNER JOIN GeneralMaster G ON G.Code=M2.UOM WHERE P.Code='" & OrderCode & "' AND (P.Code=Ref OR Ref IS NULL OR Ref='') ORDER BY M2.PrintName", cnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    
    rstPurchaseOrder.ActiveConnection = Nothing: rstPurchaseOrderChild.ActiveConnection = Nothing
    If MsgBox("Print Last Purchase Ref. Detail?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm Quit !") = vbNo Then rptPaperPurchaseOrder.Section12.Suppress = True
    
    With rptPaperPurchaseOrder
            If Logo = "S" Then
                .Picture1.Width = LogoW
                .Picture1.Height = LogoH
            End If
'        .Text2.Width = Header '9000 '7800
'        .Text2.Left = HeaderL '1000 '1680
'        .Text2.Top = 400 '240
        If LogoLine = "N" Then
        .Picture1.LeftLineStyle = crLSNoLine
        .Picture1.RightLineStyle = crLSNoLine
        .Picture1.TopLineStyle = crLSNoLine
        .Picture1.BottomLineStyle = crLSNoLine
        End If
            If Len(LTrim(rstCompanyMaster.Fields("PrintName").Value)) <= 30 Then
                .Text2.Font.Size = 20
            ElseIf Len(LTrim(rstCompanyMaster.Fields("PrintName").Value)) <= 40 Then
                .Text2.Font.Size = 18
            ElseIf Len(LTrim(rstCompanyMaster.Fields("PrintName").Value)) <= 50 Then
                .Text2.Font.Size = 16
            ElseIf Len(LTrim(rstCompanyMaster.Fields("PrintName").Value)) <= 60 Then
                .Text2.Font.Size = 14
            End If
    End With
    
    rptPaperPurchaseOrder.Text1.SetText "Paper Purchase Order"
    rptPaperPurchaseOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPaperPurchaseOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptPaperPurchaseOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value) & Space(1) & "GSTIN : " & Trim(rstCompanyMaster.Fields("GSTIN").Value)
    rptPaperPurchaseOrder.Text20.SetText "Add : GST @" + Format(rstPurchaseOrder.Fields("GST%").Value, "0.00") + "%"
    rptPaperPurchaseOrder.Text28.SetText " (" & Trim(NumberToWords(rstPurchaseOrder.Fields("BillAmount").Value, True)) & ")"
    rptPaperPurchaseOrder.Text27.SetText "for " & Trim(rstPurchaseOrder.Fields("SupplierName").Value)
    rptPaperPurchaseOrder.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPaperPurchaseOrder.Database.SetDataSource rstPurchaseOrder, 3, 1
    rptPaperPurchaseOrder.Subreport1.OpenSubreport.Database.SetDataSource rstPurchaseOrderChild, 3, 1
    
    If OutputType = "S" Then
        Set FrmReportViewer.Report = rptPaperPurchaseOrder
        FrmReportViewer.Show vbModal
    ElseIf OutputType = "P" Then
        rptPaperPurchaseOrder.PaperSource = crPRBinAuto
        rptPaperPurchaseOrder.PrintOut False    'Print Report Without Prompt
    Else
        Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
        With oOutlookMsg
            .To = rstPurchaseOrder.Fields("SupplierMail").Value
            .Subject = "Paper Purchase Order #" & Trim(rstPurchaseOrder.Fields("OrderNo").Value)
            .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith PO #" & Trim(rstPurchaseOrder.Fields("OrderNo").Value) & " for doing the needful at your end. An early execution of the order will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of execution of order.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & "Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
            rptPaperPurchaseOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
            rptPaperPurchaseOrder.ExportOptions.DestinationType = crEDTDiskFile
            FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
            rptPaperPurchaseOrder.ExportOptions.DiskFileName = FileName
            rptPaperPurchaseOrder.Export False
            .Attachments.Add (FileName)
            .Importance = olImportanceHigh
            .ReadReceiptRequested = True
            If CheckEmpty(.To, False) Then .Display Else .Send
        End With
        Set oOutlookMsg = Nothing
    End If
    Set rptPaperPurchaseOrder = Nothing
    Call CloseRecordset(rstPurchaseOrder): Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstPurchaseOrderChild)
    On Error GoTo 0
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1: fpSpread1.SetFocus
            CalculateTotal ("G"): CalculateTotal ("N")
        End If
    ElseIf KeyCode = vbKeySpace Then
        Dim LastPurchaseRate As String
        With fpSpread1
            If .ActiveCol = 1 Then
                .GetText 12, .ActiveRow, PaperCode
                On Error Resume Next
                FrmPaperMaster.SL = True
                FrmPaperMaster.MasterCode = PaperCode
                Load FrmPaperMaster
                If Err.Number <> 364 Then FrmPaperMaster.Show vbModal
                On Error GoTo 0
                .SetText 1, .ActiveRow, slName
                .SetText 12, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then
                    LoadMasterList
                    rstPaperList.MoveFirst: rstPaperList.Find "[Code] ='" & slCode & "'"
                    .SetText 3, .ActiveRow, rstPaperList.Fields("UOMName").Value
                    .SetText 4, .ActiveRow, Val(rstPaperList.Fields("Weight/Unit").Value)
                    .SetText 10, .ActiveRow, Val(rstPaperList.Fields("Units/Bundle").Value)
                    .SetText 13, .ActiveRow, Val(rstPaperList.Fields("SPU").Value)
                    PaperCode = slCode: LastPurchaseRate = GetLastPurchaseRate
                    If Not CheckEmpty(GetLastPurchaseRate, False) Then MsgBox "Last Purchase Rate : " & LastPurchaseRate & " !!!", vbInformation, App.Title
                    .SetFocus
                    Sendkeys "{ENTER}"
                Else
                    .SetActiveCell 1, .ActiveRow
                End If
            End If
        End With
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Paper As Variant, Qty As Variant, Wt As Variant, GrWt As Variant, Rate As Variant, UnitRate As Variant, UnitsPerBdl As Variant, ShtsPerUnit As Variant
    Dim QtyShts As Long, TotalBdl As Double
    With fpSpread1
        If Col = 1 Or Col = 2 Or Col = 4 Or Col = 5 Or Col = 7 Or Col = 8 Or Col = 10 Then
            .GetText 1, Row, Paper
            .GetText 2, Row, Qty
            .GetText 4, Row, Wt
            .GetText 5, Row, GrWt
            .GetText 7, Row, Rate
            .GetText 8, Row, UnitRate
            .GetText 10, Row, UnitsPerBdl
            .GetText 13, Row, ShtsPerUnit
            If Paper <> "" Then
                If (Col = 2 Or Col = 4) And Qty > 0 Then
                    GrWt = (Fix(Qty) * Wt) + ((Qty - Fix(Qty)) * 1000) * (Wt / ShtsPerUnit)
                ElseIf Col = 5 And GrWt > 0 Then
                    Qty = GrWt / Wt: Qty = Fix(Val(Qty)) + ((Val(Qty) - Fix(Val(Qty))) * ShtsPerUnit) / 1000
                End If
                If Col = 7 And Rate > 0 Then
                    .SetText 8, Row, Rate * Wt
                ElseIf Col = 8 And UnitRate > 0 Then
                    .SetText 7, Row, UnitRate / Wt
                End If
                .GetText 7, Row, Rate
                If Qty > 0 Then QtyShts = Fix(Qty) * ShtsPerUnit + (Qty - Fix(Qty)) * 1000
                If UnitsPerBdl > 0 Then TotalBdl = GrWt / (Wt * UnitsPerBdl): TotalBdl = Fix(TotalBdl) + IIf(TotalBdl - Fix(TotalBdl) > 0, 1, 0)
                .SetText 2, Row, Round(Qty, 3): .SetText 5, Row, GrWt: .SetText 6, Row, QtyShts: .SetText 9, Row, GrWt * Rate: .SetText 11, Row, TotalBdl: CalculateTotal ("G"): CalculateTotal ("N")
            Else
                .SetText 2, Row, "": .SetText 4, Row, "": .SetText 5, Row, "": .SetText 6, Row, "": .SetText 7, Row, "": .SetText 8, Row, "": .SetText 10, Row, ""
            End If
        End If
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub fpSpread2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Or KeyCode = vbKeySpace Then
        Dim VchCode As Variant
        fpSpread2.GetText 14, fpSpread2.ActiveRow, VchCode
        If Not CheckEmpty(VchCode, False) And VchCode <> rstPaperPOParent.Fields("Code").Value Then DisplayError ("This Entry cann't be edited or deleted"): fpSpread2.SetFocus: Exit Sub
    End If
    If KeyCode = vbKeyF9 Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread2.DeleteRows fpSpread2.ActiveRow, 1: fpSpread2.SetFocus
            CalculateTotal ("G")
        End If
    ElseIf KeyCode = vbKeySpace Then
        Dim Paper As Variant, Account As Variant
        With fpSpread2
            .GetText 1, .ActiveRow, Paper
            If .ActiveCol = 1 Then
                If Paper = "" Then
                    fpSpread1.GetText 1, fpSpread1.ActiveRow, Paper     'Paper Name
                    .SetText 1, .ActiveRow, Paper
                    If Paper <> "" Then Sendkeys "{ENTER}"
                    fpSpread1.GetText 12, fpSpread1.ActiveRow, Paper    'Paper Code
                    .SetText 9, .ActiveRow, Paper
                    fpSpread1.GetText 4, fpSpread1.ActiveRow, Paper     'Wt/Unit
                    .SetText 11, .ActiveRow, Paper
                    fpSpread1.GetText 10, fpSpread1.ActiveRow, Paper     'Unit/Bdl
                    .SetText 12, .ActiveRow, Paper
                    fpSpread1.GetText 13, fpSpread1.ActiveRow, Paper     'Sheets/Unit
                    .SetText 13, .ActiveRow, Paper
                End If
            ElseIf .ActiveCol = 2 Then
                If Paper <> "" Then
                    .GetText 8, .ActiveRow, AccountCode
                    On Error Resume Next
                    FrmAccountMaster.SL = True
                    FrmAccountMaster.MasterCode = AccountCode
                    Load FrmAccountMaster
                    If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
                    On Error GoTo 0
                    AccountCode = slCode
                    If Not CheckEmpty(AccountCode, False) Then
                        LoadMasterList
                        rstAccountList.MoveFirst: rstAccountList.Find "[Code] ='" & AccountCode & "'"
                        .SetText 2, .ActiveRow, rstAccountList.Fields("Col0").Value
                        .SetText 8, .ActiveRow, AccountCode
                        Sendkeys "{ENTER}"
                    Else
                        .SetText 2, .ActiveRow, slName
                        .SetActiveCell 2, .ActiveRow
                    End If
                End If
            End If
        End With
    End If
End Sub
Private Sub fpSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Paper As Variant, Qty As Variant, Wt As Variant, GrWt As Variant, UnitsPerBdl As Variant, ShtsPerUnit As Variant
    Dim QtyShts As Long, TotalBdl As Double
    With fpSpread2
        If Col = 3 Or Col = 4 Then
            .GetText 1, Row, Paper
            .GetText 3, Row, Qty
            .GetText 11, Row, Wt
            .GetText 4, Row, GrWt
            .GetText 12, Row, UnitsPerBdl
            .GetText 13, Row, ShtsPerUnit
            If Paper <> "" Then
                If Col = 3 And Qty > 0 Then
                    GrWt = (Fix(Qty) * Wt) + ((Qty - Fix(Qty)) * 1000) * (Wt / ShtsPerUnit)
                ElseIf Col = 4 And GrWt > 0 Then
                    Qty = GrWt / Wt: Qty = Fix(Qty) + ((Qty - Fix(Qty)) * ShtsPerUnit) / 1000
                End If
                If Qty > 0 Then QtyShts = Fix(Qty) * ShtsPerUnit + (Qty - Fix(Qty)) * 1000
                If UnitsPerBdl > 0 Then TotalBdl = GrWt / (Wt * UnitsPerBdl): TotalBdl = Fix(TotalBdl) + IIf(TotalBdl - Fix(TotalBdl) > 0, 1, 0)
                .SetText 3, Row, Round(Qty, 3): .SetText 4, Row, GrWt: .SetText 5, Row, QtyShts: .SetText 6, Row, TotalBdl: CalculateTotal ("G")
            Else
                .SetText 3, Row, "": .SetText 4, Row, "": .SetText 5, Row, "": .SetText 6, Row, ""
            End If
        End If
    End With
End Sub
Private Sub fpSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Function chkPaper() As Boolean
    Dim i As Integer, K As Integer, Paper01 As Variant, Qty01 As Variant, Paper02 As Variant, Qty02 As Variant, Price As Variant, Issued As Variant, Qty As Double
    chkPaper = True
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.GetText 1, i, Paper01
        fpSpread1.GetText 5, i, Qty01
        fpSpread1.GetText 7, i, Price
        If Val(Price) = 0 Then DisplayError ("Price of Paper at row #" & Trim(Str(i)) & " is zero"): chkPaper = False: Exit Function
        If fpSpread2.DataRowCnt = 0 Then chkPaper = True: Exit Function
        Qty = 0
        With fpSpread2
            For K = 1 To .DataRowCnt
                .GetText 1, K, Paper02
                .GetText 7, K, Issued
                If Paper01 = Paper02 And Val(Issued) = 1 Then
                    .GetText 4, K, Qty02
                    Qty = Qty + Val(Qty02)
                End If
            Next
        End With
        If Val(Qty) <> 0 Then If Val(Qty01) < Qty Then DisplayError ("Purchased vs Issued quantity difference for Paper - " & Paper01): chkPaper = False: Exit Function
    Next
End Function
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
Private Sub LoadMasterList()
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "SELECT LTRIM(Name) As Col0,Code FROM AccountMaster ORDER BY Name", cnPaperPurchaseOrder, adOpenKeyset, adLockReadOnly
    If rstPaperList.State = adStateOpen Then rstPaperList.Close
    rstPaperList.Open "SELECT P.Name As Col0,C.Name As UOMName,C.Value1 As SPU,[Weight/Unit],[Units/Bundle],P.Code FROM PaperMaster P INNER JOIN GeneralMaster C ON P.UOM=C.Code ORDER BY P.Name", cnPaperPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    rstPaperList.ActiveConnection = Nothing
End Sub
Private Sub cmdView_Click() 'View Pic
    If CheckEmpty(imgFile, False) Then DisplayError ("No image exists") Else Call ShellExecute(Me.hwnd, "open", imgFile, "", "", 1)
End Sub
Private Sub cmdUpload_Click() 'Load Pic
    On Error Resume Next
    With cdUpload
        .CancelError = True
        .DialogTitle = "Open Image"
        .Filter = "All Picture Files|*.jpg;*.jpeg;*.bmp;*.gif;*.png"
        .ShowOpen
        If Err.Number = 0 Then imgFile = .FileName: cmdUpload.Enabled = False 'Ok Selected
    End With
End Sub
Private Sub cmdDelete_Click() 'Delete Pic
    If Not CheckEmpty(imgFile, False) Then
        If MsgBox("Are you sure to delete the Picture?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then imgFile = "": cmdUpload.Enabled = True
    End If
End Sub
