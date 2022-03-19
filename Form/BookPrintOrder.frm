VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0F1F1508-C40A-101B-AD04-00AA00575482}#1.0#0"; "mhrinp32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBookPrintOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Print Order"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17610
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   17610
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7715
      Left            =   15
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   0
      Width           =   17580
      _Version        =   65536
      _ExtentX        =   31009
      _ExtentY        =   13608
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
      Picture         =   "BookPrintOrder.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   7485
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   120
         Width           =   17370
         _ExtentX        =   30639
         _ExtentY        =   13203
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
         TabPicture(0)   =   "BookPrintOrder.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Mh3dLabel1(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Mh3dLabel1(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Text1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "DataGrid1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdProceed"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "BookPrintOrder.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame7"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Mh3dFrame3"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Mh3dFrame5"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Mh3dFrame2"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Mh3dFrame6"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin VB.CommandButton cmdProceed 
            Caption         =   " Show Combo Item List Only"
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
            Left            =   14745
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "Save"
            Top             =   7020
            Width           =   2520
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame6 
            Height          =   855
            Left            =   -74880
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   5145
            Width           =   9675
            _Version        =   65536
            _ExtentX        =   17066
            _ExtentY        =   1517
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
            TextColor       =   0
            WallPaper       =   0
            NoPrefix        =   0   'False
            FormatString    =   ""
            Caption         =   ""
            Picture         =   "BookPrintOrder.frx":0054
            Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
               Height          =   330
               Left            =   2040
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   420
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":0070
               Caption         =   "BookPrintOrder.frx":0090
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":00FC
               Keys            =   "BookPrintOrder.frx":011A
               Spin            =   "BookPrintOrder.frx":0164
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
               Height          =   330
               Left            =   2040
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   105
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":018C
               Caption         =   "BookPrintOrder.frx":01AC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":0218
               Keys            =   "BookPrintOrder.frx":0236
               Spin            =   "BookPrintOrder.frx":0280
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   120
               TabIndex        =   54
               Top             =   105
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Estimated Unit Rate"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":02A8
               Picture         =   "BookPrintOrder.frx":02C4
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
               Height          =   330
               Left            =   3005
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   105
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":02E0
               Caption         =   "BookPrintOrder.frx":0300
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":036C
               Keys            =   "BookPrintOrder.frx":038A
               Spin            =   "BookPrintOrder.frx":03D4
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   330
               Left            =   3960
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   105
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":03FC
               Caption         =   "BookPrintOrder.frx":041C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":0488
               Keys            =   "BookPrintOrder.frx":04A6
               Spin            =   "BookPrintOrder.frx":04F0
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   330
               Left            =   4920
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   105
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":0518
               Caption         =   "BookPrintOrder.frx":0538
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":05A4
               Keys            =   "BookPrintOrder.frx":05C2
               Spin            =   "BookPrintOrder.frx":060C
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   120
               TabIndex        =   60
               Top             =   420
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Estimated Amount"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":0634
               Picture         =   "BookPrintOrder.frx":0650
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
               Height          =   330
               Left            =   3005
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   420
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":066C
               Caption         =   "BookPrintOrder.frx":068C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":06F8
               Keys            =   "BookPrintOrder.frx":0716
               Spin            =   "BookPrintOrder.frx":0760
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
               Height          =   330
               Left            =   3960
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   420
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":0788
               Caption         =   "BookPrintOrder.frx":07A8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":0814
               Keys            =   "BookPrintOrder.frx":0832
               Spin            =   "BookPrintOrder.frx":087C
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
               Height          =   330
               Left            =   4920
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   420
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":08A4
               Caption         =   "BookPrintOrder.frx":08C4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":0930
               Keys            =   "BookPrintOrder.frx":094E
               Spin            =   "BookPrintOrder.frx":0998
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
               Height          =   330
               Left            =   5880
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   105
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":09C0
               Caption         =   "BookPrintOrder.frx":09E0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":0A4C
               Keys            =   "BookPrintOrder.frx":0A6A
               Spin            =   "BookPrintOrder.frx":0AB4
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
               Height          =   330
               Left            =   5880
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   420
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":0ADC
               Caption         =   "BookPrintOrder.frx":0AFC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":0B68
               Keys            =   "BookPrintOrder.frx":0B86
               Spin            =   "BookPrintOrder.frx":0BD0
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
               Height          =   645
               Left            =   7785
               TabIndex        =   85
               Top             =   105
               Width           =   1785
               _Version        =   65536
               _ExtentX        =   3149
               _ExtentY        =   1138
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Variable Qnty Detail"
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":0BF8
               Picture         =   "BookPrintOrder.frx":0C14
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput29 
               Height          =   330
               Left            =   6840
               TabIndex        =   89
               TabStop         =   0   'False
               Top             =   105
               Width           =   960
               _Version        =   65536
               _ExtentX        =   1693
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":0C30
               Caption         =   "BookPrintOrder.frx":0C50
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":0CBC
               Keys            =   "BookPrintOrder.frx":0CDA
               Spin            =   "BookPrintOrder.frx":0D24
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput30 
               Height          =   330
               Left            =   6840
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   420
               Width           =   960
               _Version        =   65536
               _ExtentX        =   1693
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":0D4C
               Caption         =   "BookPrintOrder.frx":0D6C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":0DD8
               Keys            =   "BookPrintOrder.frx":0DF6
               Spin            =   "BookPrintOrder.frx":0E40
               AlignHorizontal =   1
               AlignVertical   =   0
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
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   6495
            Left            =   120
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   450
            Width           =   17145
            _ExtentX        =   30242
            _ExtentY        =   11456
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   9164542
            Enabled         =   -1  'True
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
            ColumnCount     =   14
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "   Order No."
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
               DataField       =   "RefNo"
               Caption         =   "Ref No."
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
               Caption         =   "Item Name"
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
               DataField       =   "UnitRate"
               Caption         =   "Unit Rate"
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
            BeginProperty Column05 
               DataField       =   "BPODStatus"
               Caption         =   "MP"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "v"
                  FalseValue      =   "x"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "TPODStatus"
               Caption         =   "SP"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "v"
                  FalseValue      =   "x"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "TLODStatus"
               Caption         =   "L"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "v"
                  FalseValue      =   "x"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "BBODStatus"
               Caption         =   "IB"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "v"
                  FalseValue      =   "x"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "DeliveredQuantity"
               Caption         =   "    Recd Qty"
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
               DataField       =   "BookPrinterName"
               Caption         =   "Multi Format"
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
            BeginProperty Column11 
               DataField       =   "TitlePrinterName"
               Caption         =   "Spread Format"
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
            BeginProperty Column12 
               DataField       =   "LaminatorName"
               Caption         =   "Misc Operation"
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
            BeginProperty Column13 
               DataField       =   "BinderName"
               Caption         =   "Binding Process"
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
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   4380.095
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   840.189
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   345.26
               EndProperty
               BeginProperty Column06 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   285.165
               EndProperty
               BeginProperty Column07 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   285.165
               EndProperty
               BeginProperty Column08 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   285.165
               EndProperty
               BeginProperty Column09 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column10 
                  Locked          =   -1  'True
                  ColumnWidth     =   1289.764
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  ColumnWidth     =   1335.118
               EndProperty
               BeginProperty Column12 
                  Locked          =   -1  'True
                  ColumnWidth     =   989.858
               EndProperty
               BeginProperty Column13 
                  Locked          =   -1  'True
                  ColumnWidth     =   1950.236
               EndProperty
            EndProperty
         End
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
            Left            =   600
            TabIndex        =   32
            Top             =   7020
            Width           =   9945
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   3645
            Left            =   -74880
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   480
            Width           =   9675
            _Version        =   65536
            _ExtentX        =   17066
            _ExtentY        =   6429
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
            Picture         =   "BookPrintOrder.frx":0E68
            Begin VB.CommandButton Command6 
               Caption         =   "BOM"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   9000
               TabIndex        =   99
               ToolTipText     =   "Binding Process Order"
               Top             =   2850
               Width           =   570
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   98
               ToolTipText     =   " Binding Process Party"
               Top             =   2840
               Width           =   6960
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   18
               ToolTipText     =   "Multi Form Format Party"
               Top             =   3140
               Width           =   7530
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   10
               ToolTipText     =   "Multi Elements Format Party"
               Top             =   1580
               Width           =   6960
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   12
               ToolTipText     =   " Combo Format Party"
               Top             =   1890
               Width           =   6960
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   16
               ToolTipText     =   " Binding Process Party"
               Top             =   2520
               Width           =   6960
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   14
               ToolTipText     =   "Misc. Operation Party"
               Top             =   2210
               Width           =   6960
            End
            Begin MhinrelLib.MhRealInput MhRealInput2 
               Height          =   330
               Left            =   7995
               TabIndex        =   7
               TabStop         =   0   'False
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
               FillColor       =   16777215
               MaxReal         =   999999
               MinReal         =   0
               ReadOnly        =   -1  'True
               SpinChangeReal  =   0
               CaretColor      =   -2147483642
               DecimalPlaces   =   0
               VAlignment      =   2
            End
            Begin VB.CommandButton Command4 
               Caption         =   "BP"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   9000
               TabIndex        =   17
               ToolTipText     =   "Binding Process Order"
               Top             =   2530
               Width           =   570
            End
            Begin VB.CommandButton Command3 
               Caption         =   "MO"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   9000
               TabIndex        =   15
               ToolTipText     =   "Misc. Operation Order"
               Top             =   2220
               Width           =   570
            End
            Begin VB.CommandButton Command2 
               Caption         =   "CF"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   9000
               TabIndex        =   13
               ToolTipText     =   " Combo Format Order"
               Top             =   1900
               Width           =   570
            End
            Begin VB.CommandButton Command5 
               Caption         =   "ME"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   9000
               TabIndex        =   11
               ToolTipText     =   " Multi Elements Format Order"
               Top             =   1590
               Width           =   570
            End
            Begin VB.CommandButton Command1 
               Caption         =   "MF"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   9000
               TabIndex        =   9
               ToolTipText     =   " Multi Form Format Order"
               Top             =   1270
               Width           =   570
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
               Left            =   7995
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   630
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   950
               Width           =   2250
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
               Left            =   2040
               MaxLength       =   10
               TabIndex        =   1
               Top             =   105
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   3
               Top             =   630
               Width           =   4410
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   35
               Top             =   105
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "BookPrintOrder.frx":0E84
               Picture         =   "BookPrintOrder.frx":0EA0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   6435
               TabIndex        =   36
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
               Picture         =   "BookPrintOrder.frx":0EBC
               Picture         =   "BookPrintOrder.frx":0ED8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   37
               Top             =   630
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "BookPrintOrder.frx":0EF4
               Picture         =   "BookPrintOrder.frx":0F10
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   42
               Top             =   950
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Item Size"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":0F2C
               Picture         =   "BookPrintOrder.frx":0F48
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   6435
               TabIndex        =   43
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
               Caption         =   " Pages"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":0F64
               Picture         =   "BookPrintOrder.frx":0F80
            End
            Begin MhinrelLib.MhRealInput MhRealInput1 
               Height          =   330
               Left            =   5235
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   945
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
               FillColor       =   16777215
               MaxReal         =   999999
               MinReal         =   0
               ReadOnly        =   -1  'True
               SpinChangeReal  =   0
               CaretColor      =   -2147483642
               DecimalPlaces   =   2
               VAlignment      =   2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Left            =   4275
               TabIndex        =   44
               Top             =   945
               Width           =   1020
               _Version        =   65536
               _ExtentX        =   1799
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Forms"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":0F9C
               Picture         =   "BookPrintOrder.frx":0FB8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Left            =   6435
               TabIndex        =   45
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
               Caption         =   " Color"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":0FD4
               Picture         =   "BookPrintOrder.frx":0FF0
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   7995
               TabIndex        =   2
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "BookPrintOrder.frx":100C
               Caption         =   "BookPrintOrder.frx":1124
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":1190
               Keys            =   "BookPrintOrder.frx":11AE
               Spin            =   "BookPrintOrder.frx":120C
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   38
               Top             =   1260
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Multi Form Format "
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":1234
               Picture         =   "BookPrintOrder.frx":1250
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   120
               TabIndex        =   39
               Top             =   1575
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Multi Element Format"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":126C
               Picture         =   "BookPrintOrder.frx":1288
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   120
               TabIndex        =   40
               Top             =   2205
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
               Caption         =   " Misc.Operation"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":12A4
               Picture         =   "BookPrintOrder.frx":12C0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   120
               TabIndex        =   41
               Top             =   2520
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Binding Process"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":12DC
               Picture         =   "BookPrintOrder.frx":12F8
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
               Height          =   330
               Left            =   120
               TabIndex        =   87
               Top             =   1890
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Combo Format "
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":1314
               Picture         =   "BookPrintOrder.frx":1330
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   8
               ToolTipText     =   "Multi Form Format Party"
               Top             =   1260
               Width           =   6960
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel29 
               Height          =   330
               Left            =   120
               TabIndex        =   95
               Top             =   3140
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Material Centre"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":134C
               Picture         =   "BookPrintOrder.frx":1368
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel30 
               Height          =   330
               Left            =   120
               TabIndex        =   100
               Top             =   2840
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " BOM"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":1384
               Picture         =   "BookPrintOrder.frx":13A0
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   9700
               Y1              =   525
               Y2              =   525
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame5 
            Height          =   525
            Left            =   -74880
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   4110
            Width           =   9675
            _Version        =   65536
            _ExtentX        =   17066
            _ExtentY        =   926
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
            Picture         =   "BookPrintOrder.frx":13BC
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   120
               TabIndex        =   47
               Top             =   105
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Email Status"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":13D8
               Picture         =   "BookPrintOrder.frx":13F4
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
               Height          =   330
               Left            =   2040
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   105
               Width           =   7530
               _Version        =   65536
               _ExtentX        =   13282
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
               Picture         =   "BookPrintOrder.frx":1410
               Begin VB.CheckBox Check1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "CF"
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
                  Left            =   3240
                  TabIndex        =   88
                  Top             =   60
                  Width           =   1260
               End
               Begin VB.CheckBox chkBP 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "MF"
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
                  Left            =   465
                  TabIndex        =   19
                  Top             =   60
                  Width           =   1260
               End
               Begin VB.CheckBox chkTP 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "SF"
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
                  Left            =   1815
                  TabIndex        =   20
                  Top             =   60
                  Width           =   1260
               End
               Begin VB.CheckBox chkTL 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "MO"
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
                  Left            =   4635
                  TabIndex        =   21
                  Top             =   60
                  Width           =   1140
               End
               Begin VB.CheckBox chkBB 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "BP"
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
                  Left            =   6150
                  TabIndex        =   22
                  Top             =   60
                  Width           =   1140
               End
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame3 
            Height          =   540
            Left            =   -74880
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   4620
            Width           =   9675
            _Version        =   65536
            _ExtentX        =   17066
            _ExtentY        =   944
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
            TextColor       =   0
            WallPaper       =   0
            NoPrefix        =   0   'False
            FormatString    =   ""
            Caption         =   ""
            Picture         =   "BookPrintOrder.frx":142C
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   120
               TabIndex        =   50
               Top             =   105
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Estimation Quantity"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":1448
               Picture         =   "BookPrintOrder.frx":1464
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   330
               Left            =   2040
               TabIndex        =   0
               Top             =   105
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":1480
               Caption         =   "BookPrintOrder.frx":14A0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":150C
               Keys            =   "BookPrintOrder.frx":152A
               Spin            =   "BookPrintOrder.frx":1574
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
               Left            =   3005
               TabIndex        =   23
               Top             =   105
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":159C
               Caption         =   "BookPrintOrder.frx":15BC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":1628
               Keys            =   "BookPrintOrder.frx":1646
               Spin            =   "BookPrintOrder.frx":1690
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
               Left            =   3960
               TabIndex        =   24
               Top             =   105
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":16B8
               Caption         =   "BookPrintOrder.frx":16D8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":1744
               Keys            =   "BookPrintOrder.frx":1762
               Spin            =   "BookPrintOrder.frx":17AC
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   330
               Left            =   4920
               TabIndex        =   25
               Top             =   105
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":17D4
               Caption         =   "BookPrintOrder.frx":17F4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":1860
               Keys            =   "BookPrintOrder.frx":187E
               Spin            =   "BookPrintOrder.frx":18C8
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
               Height          =   330
               Left            =   5880
               TabIndex        =   26
               Top             =   105
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":18F0
               Caption         =   "BookPrintOrder.frx":1910
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":197C
               Keys            =   "BookPrintOrder.frx":199A
               Spin            =   "BookPrintOrder.frx":19E4
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   8640
               TabIndex        =   51
               ToolTipText     =   "Profit Margin %"
               Top             =   105
               Width           =   930
               _Version        =   65536
               _ExtentX        =   1640
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":1A0C
               Caption         =   "BookPrintOrder.frx":1A2C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":1A98
               Keys            =   "BookPrintOrder.frx":1AB6
               Spin            =   "BookPrintOrder.frx":1B00
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###0.00"
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
               ValueVT         =   1922891781
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   7785
               TabIndex        =   52
               Top             =   105
               Width           =   870
               _Version        =   65536
               _ExtentX        =   1535
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   "  Profit %"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":1B28
               Picture         =   "BookPrintOrder.frx":1B44
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput33 
               Height          =   330
               Left            =   6840
               TabIndex        =   27
               Top             =   105
               Width           =   960
               _Version        =   65536
               _ExtentX        =   1693
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":1B60
               Caption         =   "BookPrintOrder.frx":1B80
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":1BEC
               Keys            =   "BookPrintOrder.frx":1C0A
               Spin            =   "BookPrintOrder.frx":1C54
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
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame7 
            Height          =   1200
            Left            =   -74880
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   5985
            Width           =   9675
            _Version        =   65536
            _ExtentX        =   17066
            _ExtentY        =   2108
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
            TextColor       =   0
            WallPaper       =   0
            NoPrefix        =   0   'False
            FormatString    =   ""
            Caption         =   ""
            Picture         =   "BookPrintOrder.frx":1C7C
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
               Height          =   330
               Left            =   5880
               TabIndex        =   82
               Top             =   120
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   "BP "
               Alignment       =   1
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":1C98
               Picture         =   "BookPrintOrder.frx":1CB4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
               Height          =   330
               Left            =   4920
               TabIndex        =   83
               Top             =   120
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   "MO "
               Alignment       =   1
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":1CD0
               Picture         =   "BookPrintOrder.frx":1CEC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   330
               Left            =   120
               TabIndex        =   67
               Top             =   440
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Estimated UnitRate"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":1D08
               Picture         =   "BookPrintOrder.frx":1D24
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
               Height          =   330
               Left            =   2040
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   440
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":1D40
               Caption         =   "BookPrintOrder.frx":1D60
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":1DCC
               Keys            =   "BookPrintOrder.frx":1DEA
               Spin            =   "BookPrintOrder.frx":1E34
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
               Height          =   330
               Left            =   3005
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   440
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":1E5C
               Caption         =   "BookPrintOrder.frx":1E7C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":1EE8
               Keys            =   "BookPrintOrder.frx":1F06
               Spin            =   "BookPrintOrder.frx":1F50
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput23 
               Height          =   330
               Left            =   3960
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   440
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":1F78
               Caption         =   "BookPrintOrder.frx":1F98
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":2004
               Keys            =   "BookPrintOrder.frx":2022
               Spin            =   "BookPrintOrder.frx":206C
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
               Height          =   330
               Left            =   4920
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   440
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":2094
               Caption         =   "BookPrintOrder.frx":20B4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":2120
               Keys            =   "BookPrintOrder.frx":213E
               Spin            =   "BookPrintOrder.frx":2188
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
               Height          =   330
               Left            =   5880
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   440
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":21B0
               Caption         =   "BookPrintOrder.frx":21D0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":223C
               Keys            =   "BookPrintOrder.frx":225A
               Spin            =   "BookPrintOrder.frx":22A4
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
               Height          =   330
               Left            =   120
               TabIndex        =   73
               Top             =   750
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Estimated Amount"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":22CC
               Picture         =   "BookPrintOrder.frx":22E8
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
               Height          =   330
               Left            =   2040
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   750
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":2304
               Caption         =   "BookPrintOrder.frx":2324
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":2390
               Keys            =   "BookPrintOrder.frx":23AE
               Spin            =   "BookPrintOrder.frx":23F8
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
               Height          =   330
               Left            =   3005
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   750
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":2420
               Caption         =   "BookPrintOrder.frx":2440
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":24AC
               Keys            =   "BookPrintOrder.frx":24CA
               Spin            =   "BookPrintOrder.frx":2514
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
               Height          =   330
               Left            =   3960
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   750
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":253C
               Caption         =   "BookPrintOrder.frx":255C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":25C8
               Keys            =   "BookPrintOrder.frx":25E6
               Spin            =   "BookPrintOrder.frx":2630
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
               Height          =   330
               Left            =   4920
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   750
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":2658
               Caption         =   "BookPrintOrder.frx":2678
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":26E4
               Keys            =   "BookPrintOrder.frx":2702
               Spin            =   "BookPrintOrder.frx":274C
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput28 
               Height          =   330
               Left            =   5880
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   750
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":2774
               Caption         =   "BookPrintOrder.frx":2794
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":2800
               Keys            =   "BookPrintOrder.frx":281E
               Spin            =   "BookPrintOrder.frx":2868
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   330
               Left            =   2040
               TabIndex        =   79
               Top             =   120
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " MF "
               Alignment       =   1
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":2890
               Picture         =   "BookPrintOrder.frx":28AC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
               Height          =   330
               Left            =   3005
               TabIndex        =   80
               Top             =   120
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
               Caption         =   "SF "
               Alignment       =   1
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":28C8
               Picture         =   "BookPrintOrder.frx":28E4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
               Height          =   330
               Left            =   3960
               TabIndex        =   81
               Top             =   120
               Width           =   970
               _Version        =   65536
               _ExtentX        =   1711
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   "CF "
               Alignment       =   1
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":2900
               Picture         =   "BookPrintOrder.frx":291C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
               Height          =   330
               Left            =   120
               TabIndex        =   84
               Top             =   120
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   ""
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":2938
               Picture         =   "BookPrintOrder.frx":2954
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
               Height          =   960
               Left            =   7785
               TabIndex        =   86
               Top             =   120
               Width           =   1785
               _Version        =   65536
               _ExtentX        =   3149
               _ExtentY        =   1693
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Base Qnty Detail"
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":2970
               Picture         =   "BookPrintOrder.frx":298C
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput31 
               Height          =   330
               Left            =   6840
               TabIndex        =   91
               TabStop         =   0   'False
               Top             =   440
               Width           =   960
               _Version        =   65536
               _ExtentX        =   1693
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":29A8
               Caption         =   "BookPrintOrder.frx":29C8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":2A34
               Keys            =   "BookPrintOrder.frx":2A52
               Spin            =   "BookPrintOrder.frx":2A9C
               AlignHorizontal =   1
               AlignVertical   =   0
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
               Height          =   330
               Left            =   6840
               TabIndex        =   92
               Top             =   120
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
               Caption         =   "BOM "
               Alignment       =   1
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookPrintOrder.frx":2AC4
               Picture         =   "BookPrintOrder.frx":2AE0
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput32 
               Height          =   330
               Left            =   6840
               TabIndex        =   93
               TabStop         =   0   'False
               Top             =   750
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   582
               Calculator      =   "BookPrintOrder.frx":2AFC
               Caption         =   "BookPrintOrder.frx":2B1C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":2B88
               Keys            =   "BookPrintOrder.frx":2BA6
               Spin            =   "BookPrintOrder.frx":2BF0
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   10530
            TabIndex        =   94
            Top             =   7020
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
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
            Picture         =   "BookPrintOrder.frx":2C18
            Picture         =   "BookPrintOrder.frx":2C34
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   1
            Left            =   13440
            TabIndex        =   96
            Top             =   0
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
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
            Caption         =   " F5-> Refresh-> F12-> Create Duplicate Posting"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "BookPrintOrder.frx":2C50
            Picture         =   "BookPrintOrder.frx":2C6C
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
            TabIndex        =   33
            Top             =   7020
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   17610
      _ExtentX        =   31062
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
Attribute VB_Name = "FrmBookPrintOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FI As Object, rstEasyPublish As DAO.Recordset
Public BookPOType As String, DisplayListType As String, ItemCode As String
Dim cnBookPrintOrder As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset, rstBookPOList As New ADODB.Recordset, rstBookPOParent As New ADODB.Recordset, rstBookPOChild05 As New ADODB.Recordset, rstBookPOChild06 As New ADODB.Recordset, rstBookPOChild09 As New ADODB.Recordset, rstBookPOChild10 As New ADODB.Recordset, rstBookPOChild0901 As New ADODB.Recordset, rstBookPOChild07 As New ADODB.Recordset, rstBookPOChild08 As New ADODB.Recordset, rstBookPOChild0801 As New ADODB.Recordset, rstCorrections As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstMaterialCentreList As New ADODB.Recordset, srmPicMgr As New ADODB.Stream
Public rstBookList As New ADODB.Recordset, imgFile As String
Dim PaperCode As String, BookPrinterCode As String, TitlePrinterCode As String, LaminatorCode As String, BinderCode As String, MaterialCentreCode As String
Dim Amount As Double, TaxableAmount As Double, UnitRate As Double, UnitRateBT As Double, PMTaxableAmount As Double, PMUnitRateBT As Double, PMUnitRate As Double, PMTotalTax As Double
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim oOutlook As New Outlook.Application
Dim EMailID As String, Attachment As String, Message As String, OutputTo As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
'    Set FI = CreateObject("Busy2L16.CFixedInterface")
 '   FI.OpenCSDB Trim(ReadFromFile("BusyPath")), Trim(ReadFromFile("Server Name")), "sa", Trim(ReadFromFile("Server Password")), Mid(cnBusy.DefaultDatabase, 5, 8)
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    Unload FrmBookPOPrintUtility
    CenterForm Me
    WheelHook DataGrid1
    Me.Caption = IIf(BookPOType = "DS", "Sales Order [Digital Printing]", IIf(BookPOType = "DP", "Purchase Order [Digital Printing]", IIf(BookPOType = "FP", "Purchase Order [Finished Goods]", IIf(BookPOType = "RP", "Purchase Order [Unfinished Goods]", IIf(BookPOType = "OP", "Cost Sheet", IIf(BookPOType = "FS", "Sales Order [Finished Goods]", "Sales Order [Unfinished Goods]"))))))
    If Left(BookPOType, 1) = "O" Then Mh3dFrame5.Visible = False: Mh3dFrame3.Top = 3750: Mh3dFrame6.Top = 4280: Mh3dFrame7.Top = 5120 Else Mh3dLabel14.Caption = " Final Quantity": Mh3dLabel15.Caption = " Final Unit Rate": Mh3dLabel16.Caption = " Final Amount": Mh3dLabel17.Caption = " Unit Rate": Mh3dLabel18.Caption = " Amount": Mh3dLabel26.Caption = " Final Qnty Detail": MhRealInput3.Width = 5780: MhRealInput9.Width = 7530: MhRealInput14.Width = 7530: MhRealInput4.Visible = False: MhRealInput5.Visible = False: MhRealInput6.Visible = False: MhRealInput7.Visible = False: MhRealInput33.Visible = False: Mh3dLabel25.Visible = False
    cnBookPrintOrder.CursorLocation = adUseClient
    cnBookPrintOrder.Open cnDatabase.ConnectionString
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,EMail,Website From CompanyMaster", cnBookPrintOrder, adOpenKeyset, adLockReadOnly
    rstBookPOParent.CursorLocation = adUseClient
    SSTab1.Tab = 0
'    SortOrder = "Name"
    If FrmStockLedger.dSortBy = True Then
    SortOrder = "Code"
    Else
    SortOrder = "NAME"
    End If
    DisplayListType = "O"
    Call RefreshList("")
    LoadMasterList
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True
    MdiMainMenu.mnuPurchaseOrderJobWork.Enabled = False: MdiMainMenu.mnuSalesOrderJobWork.Enabled = False: MdiMainMenu.mnuCostSheet.Enabled = False
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If SSTab1.Tab = 0 Then
            CloseForm Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                    Me.ActiveControl.SetFocus
                Else
                    Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                End If
            End If
            KeyCode = 0
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
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF12 Then
        If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then DuplicateRecord
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(9)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(10)
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
            SSTab1.Tab = 1
            SSTab1.SetFocus
        Else
            Sendkeys "{TAB}"
        End If
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1 Else CloseForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBookPOList)
    Call CloseRecordset(rstBookPOParent)
    Call CloseRecordset(rstBookPOChild05)
    Call CloseRecordset(rstBookPOChild06)
    Call CloseRecordset(rstBookPOChild09)
    Call CloseRecordset(rstBookPOChild0901)
    Call CloseRecordset(rstBookPOChild07)
    Call CloseRecordset(rstBookPOChild08)
    Call CloseRecordset(rstBookPOChild0801)
    Call CloseRecordset(rstCorrections)
    Call CloseRecordset(rstBookList)
    Call CloseConnection(cnBookPrintOrder)
    If srmPicMgr.State = adStateOpen Then srmPicMgr.Close
    Set srmPicMgr = Nothing
    OutputTo = ""
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuPurchaseOrderJobWork.Enabled = True: MdiMainMenu.mnuSalesOrderJobWork.Enabled = True: MdiMainMenu.mnuCostSheet.Enabled = True
End Sub
Private Sub MhRealInput3_Validate(Cancel As Boolean) 'EstQty01
    If MhRealInput3.Value = 0 Then Cancel = True Else MhRealInput8_Validate False
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)
    If MhRealInput4.Value > 0 Then MhRealInput8_Validate False
End Sub
Private Sub MhRealInput5_Validate(Cancel As Boolean)
    If MhRealInput5.Value > 0 Then MhRealInput8_Validate False
End Sub
Private Sub MhRealInput6_Validate(Cancel As Boolean)
    If MhRealInput6.Value > 0 Then MhRealInput8_Validate False
End Sub
Private Sub MhRealInput7_Validate(Cancel As Boolean)
    If MhRealInput7.Value > 0 Then MhRealInput8_Validate False
End Sub
Private Sub MhRealInput33_Validate(Cancel As Boolean)
    If MhRealInput33.Value > 0 Then MhRealInput8_Validate False
End Sub
Private Sub Text1_Change()
On Error Resume Next
    With rstBookPOList
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
    
    If rstBookPOList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstBookPOList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstBookPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstBookPOList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstBookPOList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstBookPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstBookPOList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstBookPOList
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
            Me.Width = 10230
            Mh3dFrame1.Width = 10125
            SSTab1.Width = 9885
            CenterForm Me
            ViewRecord
        Else
            Me.Width = 17700
            Mh3dFrame1.Width = 17580
            SSTab1.Width = 17370
            CenterForm Me
            If Not (rstBookPOList.EOF Or rstBookPOList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            Text1.SetFocus
        End If
        SSTab1.TabEnabled(0) = True
    Else
        Me.Width = 10230
        Mh3dFrame1.Width = 10125
        SSTab1.Width = 9885
        CenterForm Me
        SSTab1.TabEnabled(0) = False
        MhRealInput3.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer, i As Integer, CellVal As Variant
    
    If Button.Index = 1 Then
        If rstBookPOParent.State = adStateOpen Then rstBookPOParent.Close
        rstBookPOParent.Open "SELECT * FROM BookPOParent WHERE Code=''", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
        ClearFields
        Call LoadOrder("")
        If rstBookPOChild05.State = adStateClosed Or rstBookPOChild06.State = adStateClosed Or rstBookPOChild09.State = adStateClosed Or rstBookPOChild0901.State = adStateClosed Or rstBookPOChild07.State = adStateClosed Or rstBookPOChild08.State = adStateClosed Or rstBookPOChild0801.State = adStateClosed Then SSTab1.Tab = 0: Exit Sub
        Me.Tag = "A"
        If AddRecord(rstBookPOParent) Then
            Text2.Text = GenerateCode(cnBookPrintOrder, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,Name))", "VAL(Name))") & "  FROM BookPOParent WHERE Type='" & BookPOType & "' AND FYCode='" & FYCode & "' AND LEFT(Code,1)<>'*'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            MhRealInput3.SetFocus
            blnRecordExist = False
            cnBookPrintOrder.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstBookPOList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        Me.Tag = "E"
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstBookPOList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            If Right(BookPOType, 1) = "S" Then If chkBilled Then DisplayError ("Deletion failed due to bill generation"): Exit Sub
            If Right(BookPOType, 1) = "P" Then If chkBilled Then DisplayError ("Deletion failed due to bill generation"): Exit Sub
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnBookPrintOrder.Execute "Delete From BookPOParent Where Code = '" & rstBookPOList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstBookPOList.Delete
                rstBookPOList.MoveNext
                If rstBookPOList.RecordCount > 0 And rstBookPOList.EOF Then rstBookPOList.MoveLast
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
        'If blnRecordExist Then If chkBilled Then DisplayError ("Updation failed due to bill generation"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(5): Exit Sub
        If CheckMandatoryFields Then Exit Sub
        Set FrmBookPOChild0801.rstBookPOChild0801 = rstBookPOChild0801  'Outsource Material
        FrmBookPOChild0801.BinderCode = BinderCode
        FrmBookPOChild0801.BookCode = ItemCode
        FrmBookPOChild0801.OrderCode = CheckNull(rstBookPOParent.Fields("Code").Value)
        Load FrmBookPOChild0801
        FrmBookPOChild0801.Show vbModal
        MhRealInput8_Validate False
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstBookPOParent) Then
            Amount = 0
            If UpdateOrder("D") Then
                UpdateFlag = 1
                With rstBookPOChild05   'Text Printing
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            If Val(.Fields("ActualQuantity").Value) <> 0 Then
                                If Not UpdateOrder("U", "1") Then UpdateFlag = 0: Exit Do
                            End If
                            .MoveNext
                        Loop
                    End If
                End With
                If UpdateFlag Then  'Binding
                    With rstBookPOChild08
                        If .RecordCount <> 0 Then
                            .MoveFirst
                            Do While Not .EOF
                                If Val(.Fields("Quantity").Value) <> 0 Then
                                    If Not UpdateOrder("U", "4") Then UpdateFlag = 0: Exit Do
                                End If
                                .MoveNext
                            Loop
                        End If
                    End With
                End If
                If UpdateFlag Then  'Title Printing
                    With rstBookPOChild06
                        If .RecordCount <> 0 Then
                            .MoveFirst
                            Do While Not .EOF
                                If Val(.Fields("ActualQuantity").Value) <> 0 Then
                                    If Not UpdateOrder("U", "21") Then UpdateFlag = 0: Exit Do
                                End If
                                .MoveNext
                            Loop
                        End If
                    End With
                End If
                If UpdateFlag Then  'Combo Printing
                    If rstBookPOChild09.RecordCount <> 0 Then
                        Dim Qty As Long
                        With rstBookPOChild0901
                            .MoveFirst
                            Do While Not .EOF
                                Qty = Qty + .Fields("ActualQuantity").Value
                                .MoveNext
                            Loop
                            .MoveFirst
                        End With
                        With rstBookPOChild09
                            Do While Not .EOF
                                If Qty <> 0 Then If Not UpdateOrder("U", "22") Then UpdateFlag = 0: Exit Do
                                .MoveNext
                            Loop
                        End With
                    End If
                End If
                If UpdateFlag Then  'Misc Operations
                    With rstBookPOChild07
                        If .RecordCount <> 0 Then
                            .MoveFirst
                            Do While Not .EOF
                                If Val(.Fields("BillAmount").Value) <> 0 Then
                                    If Not UpdateOrder("U", "3") Then UpdateFlag = 0: Exit Do
                                End If
                                .MoveNext
                            Loop
                        End If
                    End With
                    If InStr(1, "FP_RP_FS_RS", BookPOType) > 0 Then If Not Save2Master() Then UpdateFlag = 0
                End If
                If UpdateFlag Then  'Outsource Material
                    With rstBookPOChild0801
                        If .RecordCount > 0 Then .MoveFirst
                        Do While Not .EOF
                            If Val(.Fields("Consumption/Item").Value) <> 0 Then
                                If Not UpdateOrder("U", "0") Then UpdateFlag = 0: Exit Do
                            End If
                            .MoveNext
                        Loop
                    End With
                End If
                If UpdateFlag Then
                    cnBookPrintOrder.Execute "UPDATE BookPOChild05 SET BilledMFC=C.BilledMFC,BilledMFB=C.BilledMFB,DeliveredQuantityC=C.DeliveredQuantityC,DeliveredQuantityB=C.DeliveredQuantityB FROM BookPOChild05 P INNER JOIN #T05 C ON P.Code=C.Code"
                    cnBookPrintOrder.Execute "UPDATE BookPOChild06 SET BilledMEC=C.BilledMEC,BilledMEB=C.BilledMEB,DeliveredQuantityC=C.DeliveredQuantityC,DeliveredQuantityB=C.DeliveredQuantityB FROM BookPOChild06 P INNER JOIN #T06 C ON P.Code=C.Code"
                    cnBookPrintOrder.Execute "UPDATE BookPOChild07 SET BilledMOC=C.BilledMOC,BilledMOB=C.BilledMOB,DeliveredQuantityC=C.DeliveredQuantityC,DeliveredQuantityB=C.DeliveredQuantityB FROM BookPOChild07 P INNER JOIN #T07 C ON P.Code=C.Code"
                    cnBookPrintOrder.Execute "UPDATE BookPOChild08 SET BilledBNC=C.BilledBNC,BilledBNB=C.BilledBNB,DeliveredQuantityC=C.DeliveredQuantityC,DeliveredQuantityB=C.DeliveredQuantityB FROM BookPOChild08 P INNER JOIN #T08 C ON P.Code=C.Code"
                    cnBookPrintOrder.Execute "UPDATE BookPOChild0801 SET BilledBMC=C.BilledBMC,BilledBMB=C.BilledBMB,DeliveredQuantityC=C.DeliveredQuantityC,DeliveredQuantityB=C.DeliveredQuantityB FROM BookPOChild0801 P INNER JOIN #T0801 C ON P.Code=C.Code"
                    cnBookPrintOrder.Execute "UPDATE BookPOChild0901 SET BilledCFC=C.BilledCFC,BilledCFB=C.BilledCFB,DeliveredQuantityC=C.DeliveredQuantityC,DeliveredQuantityB=C.DeliveredQuantityB FROM BookPOChild0901 P INNER JOIN #T0901 C ON P.Code=C.Code"
                    With FrmCorrectionRegister
                        Dim SNo As Variant
                        i = 1
                        On Error Resume Next
                        For i = 1 To .fpSpread3.DataRowCnt
                            .fpSpread3.GetText 1, i, SNo
                            If Val(SNo) = 1 Then
                                .fpSpread3.GetText 4, i, SNo
                                cnBookPrintOrder.Execute "UPDATE BookChild02 SET Remarks='PO #" & Trim(Text2.Text) & " Dated-" & Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") & "' WHERE Code='" & ItemCode & "' AND SNo=" & Val(SNo)
                                If Err.Number <> 0 Then UpdateFlag = 0: Exit For
                            End If
                        Next
                        On Error GoTo 0
                    End With
                End If
            End If
        End If
        Call CloseForm(FrmCorrectionRegister)
        If UpdateFlag Then
            Call RefreshList(rstBookPOParent.Fields("Code").Value)
            If DatabaseType = "MS SQL" Then cnBookPrintOrder.Execute "UPDATE BookPOParent SET UnitRate=" & Round((Amount + ((Amount * MhRealInput8.Value) / 100)) / Val(rstBookPOParent.Fields("EstQty01").Value), 3) & " WHERE Code='" & rstBookPOParent.Fields("Code").Value & "'"
            cnBookPrintOrder.CommitTrans
            If DatabaseType <> "MS SQL" Then cnBookPrintOrder.Execute "UPDATE BookPOParent SET UnitRate=" & Round((Amount + ((Amount * MhRealInput8.Value) / 100)) / Val(rstBookPOParent.Fields("EstQty01").Value), 3) & " WHERE Code='" & rstBookPOParent.Fields("Code").Value & "'"
            UpdateLastPrinterBinder
            If rstBookPOParent.State = adStateOpen Then rstBookPOParent.Close
            rstBookPOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            Call MsgBox("Record updated !!!", vbInformation, App.Title)
            If Left(BookPOType, 1) <> "" Then
                If BusyIntegration Or TallyIntegration Then
                    If MsgBox("Are you sure to export the Voucher?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Export !") = vbYes Then PushVch
                End If
            End If
            SSTab1.Tab = 0
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstBookPOParent) Then
            cnBookPrintOrder.RollbackTrans
            If rstBookPOParent.State = adStateOpen Then rstBookPOParent.Close
            rstBookPOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
            Call CloseForm(FrmCorrectionRegister)
        End If
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
        OutputTo = "P"
        DisplayMenu
'        If Left(BookPOType, 1) <> "O" Then DisplayMenu Else PrintPlanning (rstBookPOList.Fields("Code").Value)
    ElseIf Button.Index = 10 Then
        OutputTo = "S"
        DisplayMenu
'        If Left(BookPOType, 1) <> "O" Then DisplayMenu Else PrintPlanning (rstBookPOList.Fields("Code").Value)
    ElseIf Button.Index = 13 Then
        If rstBookPOList.RecordCount > 0 Then rstBookPOList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstBookPOList.RecordCount > 0 Then
            rstBookPOList.MovePrevious
            If rstBookPOList.BOF Then rstBookPOList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstBookPOList.RecordCount > 0 Then
            rstBookPOList.MoveNext
            If rstBookPOList.EOF Then
                rstBookPOList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstBookPOList.RecordCount > 0 Then rstBookPOList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        CloseForm Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstBookPOList.EOF Or rstBookPOList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        Text1.SetFocus
    End If
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
    End If
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstBookPOList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstBookPOList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstBookPOList.EOF Or rstBookPOList.BOF) Then
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
    Mh3dFrame5.Enabled = Not bVal
    Mh3dFrame3.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstBookPOList.RecordCount = 0 Then
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
    If rstBookPOParent.EOF Or rstBookPOParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnBookPrintOrder, "BookPOParent", "Code", "[Name]+[Type]", Trim(Text2.Text) & BookPOType, rstBookPOParent.Fields("Code").Value, False, FYCode) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If CheckEmpty(Text3.Text, False) Then Cancel = True
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmBookMaster.SL = True
        FrmBookMaster.ItemType = IIf(Left(BookPOType, 1) = "D", "F", IIf(Left(BookPOType, 1) = "O", "F", Left(BookPOType, 1)))
        FrmBookMaster.MasterCode = ItemCode
        Load FrmBookMaster
        If Err.Number <> 364 Then FrmBookMaster.Show vbModal
        On Error GoTo 0
        ItemCode = slCode: Text3.Text = slName
        If Not CheckEmpty(ItemCode, False) Then
            LoadMasterList
            rstBookList.MoveFirst
            rstBookList.Find "[Code] = '" & ItemCode & "'"
            If Val(rstBookList.Fields("TwoColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
                Text10.Text = "1 Color"
            ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
                Text10.Text = "2 Color"
            ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("TwoColorPages").Value) = 0 Then
                Text10.Text = "4 Color"
            ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("TwoColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
                Text10.Text = "6 Color"
            Else
                Text10.Text = "Multi Color"
            End If
            Text4.Text = rstBookList.Fields("SizeName").Value & "/" & Choose(Val(rstBookList.Fields("FormType").Value), "08", "16", "04", "12", "24", "32", "64", "06", "02")
            MhRealInput1.Text = Val(rstBookList.Fields("Forms").Value)
            MhRealInput2.Text = Val(rstBookList.Fields("Pages").Value)
            If Me.Tag = "A" Then
                If Trim(rstBookList.Fields("BookPrinter")) <> "" Then
                    rstAccountList.MoveFirst
                    BookPrinterCode = rstBookList.Fields("BookPrinter")
                    rstAccountList.Find "[Code] = '" & RTrim(BookPrinterCode) & "'"
                    If Not rstAccountList.EOF Then Text5.Text = rstAccountList.Fields("Col0").Value
                End If
                If Trim(rstBookList.Fields("TitlePrinter")) <> "" Then
                    rstAccountList.MoveFirst
                    TitlePrinterCode = rstBookList.Fields("TitlePrinter")
                    rstAccountList.Find "[Code] = '" & RTrim(TitlePrinterCode) & "'"
                    If Not rstAccountList.EOF Then Text6.Text = rstAccountList.Fields("Col0").Value: Text9.Text = rstAccountList.Fields("Col0").Value
                End If
                If Trim(rstBookList.Fields("Laminator")) <> "" Then
                    rstAccountList.MoveFirst
                    LaminatorCode = rstBookList.Fields("Laminator")
                    rstAccountList.Find "[Code] = '" & RTrim(LaminatorCode) & "'"
                    If Not rstAccountList.EOF Then Text7.Text = rstAccountList.Fields("Col0").Value
                End If
                If InStr(1, "F_O", Left(BookPOType, 1)) > 0 Then
                    If Trim(rstBookList.Fields("BinderFresh")) <> "" Then
                        rstAccountList.MoveFirst
                        BinderCode = rstBookList.Fields("BinderFresh")
                        rstAccountList.Find "[Code] = '" & RTrim(BinderCode) & "'"
                        If Not rstAccountList.EOF Then Text8.Text = rstAccountList.Fields("Col0").Value
                    End If
                ElseIf Left(BookPOType, 1) = "R" Then
                    If Trim(rstBookList.Fields("BinderRepair")) <> "" Then
                        rstAccountList.MoveFirst
                        BinderCode = rstBookList.Fields("BinderRepair")
                        rstAccountList.Find "[Code] = '" & RTrim(BinderCode) & "'"
                        If Not rstAccountList.EOF Then Text8.Text = rstAccountList.Fields("Col0").Value
                    End If
                End If
                Sendkeys "{TAB}"
            End If
        End If
    ElseIf KeyCode = vbKeyDelete Then
        Text3.Text = "": ItemCode = ""
    End If
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = BookPrinterCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        BookPrinterCode = slCode: Text5.Text = slName
        If Not CheckEmpty(BookPrinterCode, False) Then
            LoadMasterList
            If Not CheckEmpty(Text5.Text, False) Then Command1_Click
            Sendkeys "{TAB}"
        End If
    ElseIf KeyCode = vbKeyDelete Then
        Text5.Text = "": BookPrinterCode = ""
    End If
End Sub
Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = TitlePrinterCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        TitlePrinterCode = slCode: Text6.Text = slName
        If Not CheckEmpty(TitlePrinterCode, False) Then
            LoadMasterList
            If Not CheckEmpty(Text6.Text, False) Then Command5_Click
            Sendkeys "{TAB}"
        End If
    ElseIf KeyCode = vbKeyDelete Then
        Text6.Text = "": TitlePrinterCode = ""
    End If
End Sub
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = TitlePrinterCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        TitlePrinterCode = slCode: Text9.Text = slName
        If Not CheckEmpty(TitlePrinterCode, False) Then
            LoadMasterList
            If Not CheckEmpty(Text9.Text, False) Then Command2_Click
            Sendkeys "{TAB}"
        End If
    ElseIf KeyCode = vbKeyDelete Then
        Text9.Text = "": TitlePrinterCode = ""
    End If
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = LaminatorCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        LaminatorCode = slCode: Text7.Text = slName
        If Not CheckEmpty(LaminatorCode, False) Then
            LoadMasterList
            If Not CheckEmpty(Text7.Text, False) Then Command3_Click
            Sendkeys "{TAB}"
        End If
    ElseIf KeyCode = vbKeyDelete Then
        Text7.Text = "": LaminatorCode = ""
    End If
End Sub
Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = BinderCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        BinderCode = slCode: Text8.Text = slName
        If Not CheckEmpty(BinderCode, False) Then
            LoadMasterList
            If Not CheckEmpty(Text8.Text, False) Then Command4_Click
            Sendkeys "{TAB}"
        End If
    ElseIf KeyCode = vbKeyDelete Then
        Text8.Text = "": BinderCode = ""
    End If
End Sub
Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = "*99999"
        FrmAccountMaster.MasterCode = MaterialCentreCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        MaterialCentreCode = slCode: Text11.Text = slName
        If Not CheckEmpty(MaterialCentreCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text11_Validate(Cancel As Boolean)
    If CheckEmpty(Text11.Text, False) Then Cancel = True
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstBookPOList.EOF Then
        If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
        If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
        If rstBookPOChild09.State = adStateOpen Then rstBookPOChild09.Close
        If rstBookPOChild0901.State = adStateOpen Then rstBookPOChild0901.Close
        If rstBookPOChild07.State = adStateOpen Then rstBookPOChild07.Close
        If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
        If rstBookPOChild0801.State = adStateOpen Then rstBookPOChild0801.Close
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstBookPOParent.State = adStateOpen Then rstBookPOParent.Close
    rstBookPOParent.Open "Select * From BookPOParent Where Code = '" & FixQuote(rstBookPOList.Fields("Code").Value) & "'", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
    If rstBookPOParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text9.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhRealInput1.Text = 0#
    MhRealInput2.Text = 0
    MhRealInput3.Value = 0
    MhRealInput4.Value = 0
    MhRealInput5.Value = 0
    MhRealInput6.Value = 0
    MhRealInput7.Value = 0
    MhRealInput33.Value = 0
    MhRealInput8.Value = 0
    BookPrinterCode = ""
    TitlePrinterCode = ""
    LaminatorCode = ""
    BinderCode = ""
    chkBP.Value = 0
    chkTP.Value = 0
    chkTL.Value = 0
    chkBB.Value = 0
    Me.Tag = ""
    PaperCode = "": ItemCode = "": BookPrinterCode = "": TitlePrinterCode = "": LaminatorCode = "": BinderCode = "": imgFile = ""
    Command1.Enabled = True: Command3.Enabled = True: Command5.Enabled = True: Command4.Enabled = True:
End Sub
Private Sub LoadFields()
    With rstBookPOParent
        If .EOF Or .BOF Then Exit Sub
        Text2.Text = .Fields("Name").Value
        MhDateInput1.Text = Format(.Fields("Date").Value, "dd-MM-yyyy")
        ItemCode = .Fields("Book").Value
        If rstBookList.RecordCount > 0 Then rstBookList.MoveFirst
        rstBookList.Find "[Code] = '" & ItemCode & "'"
        If Not rstBookList.EOF Then
            Text3.Text = rstBookList.Fields("Col0").Value
            If Val(rstBookList.Fields("TwoColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
                Text10.Text = "1 Color"
            ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
                Text10.Text = "2 Color"
            ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("TwoColorPages").Value) = 0 Then
                Text10.Text = "4 Color"
            ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("TwoColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
                Text10.Text = "6 Color"
            Else
                Text10.Text = "Multi Color"
            End If
            Text4.Text = rstBookList.Fields("SizeName").Value & "/" & Choose(Val(rstBookList.Fields("FormType").Value), "08", "16", "04", "12", "24", "32", "64", "06", "02")
            MhRealInput1.Text = Val(rstBookList.Fields("Forms").Value)
            MhRealInput2.Text = Val(rstBookList.Fields("Pages").Value)
        End If
        BookPrinterCode = .Fields("BookPrinter").Value
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        rstAccountList.Find "[Code] = '" & BookPrinterCode & "'"
        If Not rstAccountList.EOF Then Text5.Text = rstAccountList.Fields("Col0").Value
        TitlePrinterCode = .Fields("TitlePrinter").Value
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        rstAccountList.Find "[Code] = '" & TitlePrinterCode & "'"
        If Not rstAccountList.EOF Then Text6.Text = rstAccountList.Fields("Col0").Value: Text9.Text = rstAccountList.Fields("Col0").Value
        LaminatorCode = .Fields("Laminator").Value
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        rstAccountList.Find "[Code] = '" & LaminatorCode & "'"
        If Not rstAccountList.EOF Then Text7.Text = rstAccountList.Fields("Col0").Value
        BinderCode = .Fields("Binder").Value
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        rstAccountList.Find "[Code] = '" & BinderCode & "'"
        If Not rstAccountList.EOF Then Text8.Text = rstAccountList.Fields("Col0").Value
        MaterialCentreCode = .Fields("MaterialCentre").Value
        If rstMaterialCentreList.RecordCount > 0 Then rstMaterialCentreList.MoveFirst
        rstMaterialCentreList.Find "[Code] = '" & MaterialCentreCode & "'"
        If Not rstMaterialCentreList.EOF Then Text11.Text = rstMaterialCentreList.Fields("Col0").Value
        chkBP.Value = IIf(.Fields("BPODStatus").Value, 1, 0)
        chkTP.Value = IIf(.Fields("TPODStatus").Value, 1, 0)
        chkTL.Value = IIf(.Fields("TLODStatus").Value, 1, 0)
        chkBB.Value = IIf(.Fields("BBODStatus").Value, 1, 0)
        MhRealInput3.Value = Val(.Fields("EstQty01").Value)
        MhRealInput4.Value = Val(.Fields("EstQty02").Value)
        MhRealInput5.Value = Val(.Fields("EstQty03").Value)
        MhRealInput6.Value = Val(.Fields("EstQty04").Value)
        MhRealInput7.Value = Val(.Fields("EstQty05").Value)
        MhRealInput33.Value = Val(.Fields("EstQty06").Value)
        MhRealInput8.Value = Val(.Fields("ProfitMargin").Value)
        If Dir(App.Path & "\Pic\", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Pic\"
        If Dir(App.Path & "\Pic\" & FinancialYear & CompCode, vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Pic\" & FinancialYear & CompCode
        If Not CheckEmpty(.Fields("PicData"), False) Then imgFile = App.Path & "\Pic\" & FinancialYear & CompCode & "\" & FinancialYear & CompCode & .Fields("Code").Value & "." & .Fields("PicType").Value: RetrievePic .Fields("PicData").Value, imgFile, srmPicMgr
        Call LoadOrder(.Fields("Code").Value)
    End With
    MhRealInput8_Validate False
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstBookPOParent.RecordCount = 0 Then Exit Sub
    If rstBookPOChild05.State = adStateClosed Or rstBookPOChild06.State = adStateClosed Or rstBookPOChild09.State = adStateClosed Or rstBookPOChild0901.State = adStateClosed Or rstBookPOChild07.State = adStateClosed Or rstBookPOChild08.State = adStateClosed Or rstBookPOChild0801.State = adStateClosed Then SSTab1.Tab = 0: Exit Sub
    If rstBookPOParent.State = adStateOpen Then rstBookPOParent.Close
    rstBookPOParent.CursorLocation = adUseServer
    rstBookPOParent.Open "Select * From BookPOParent Where Code = '" & FixQuote(rstBookPOList.Fields("Code").Value) & "'", cnBookPrintOrder, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstBookPOParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    MhRealInput3.SetFocus
    blnRecordExist = True
    If AllowTransactionsModification = 0 Then
        LockFields (True)
        Text1.Locked = False: Text5.Locked = False: Text6.Locked = False: Text9.Locked = False: Text7.Locked = False: Text8.Locked = False: Text11.Locked = False
    End If
    cnBookPrintOrder.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    With rstBookPOParent
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnBookPrintOrder, "SELECT MAX(Code) FROM BookPOParent", 6, "0")
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
        .Fields("Book").Value = ItemCode
        .Fields("BookPrinter").Value = BookPrinterCode
        .Fields("TitlePrinter").Value = TitlePrinterCode
        .Fields("Laminator").Value = LaminatorCode
        .Fields("Binder").Value = BinderCode
        .Fields("MaterialCentre").Value = MaterialCentreCode
        .Fields("BPODStatus").Value = chkBP.Value
        .Fields("TPODStatus").Value = chkTP.Value
        .Fields("TLODStatus").Value = chkTL.Value
        .Fields("BBODStatus").Value = chkBB.Value
        .Fields("EstQty01").Value = MhRealInput3.Value
        .Fields("EstQty02").Value = MhRealInput4.Value
        .Fields("EstQty03").Value = MhRealInput5.Value
        .Fields("EstQty04").Value = MhRealInput6.Value
        .Fields("EstQty05").Value = MhRealInput7.Value
        .Fields("EstQty06").Value = MhRealInput33.Value
        .Fields("ProfitMargin").Value = MhRealInput8.Value
        .Fields("ProfitMargin").Value = MhRealInput8.Value
        .Fields("UnitRate-MF").Value = MhRealInput19.Value
        .Fields("UnitRate-SF").Value = MhRealInput21.Value
        .Fields("UnitRate-CF").Value = MhRealInput23.Value
        .Fields("UnitRate-MO").Value = MhRealInput25.Value
        .Fields("UnitRate-BP").Value = MhRealInput27.Value
        .Fields("UnitRate-BOM").Value = MhRealInput31.Value
        .Fields("Type").Value = BookPOType
        .Fields("FYCode").Value = FYCode
        .Fields("PrintStatus").Value = "N"
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(MhRealInput3.Text, False) Then
       DisplayError ("Final Quantity cannot be blank")
       MhRealInput3.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckEmpty(Text2.Text, False) Then
       DisplayError ("Order No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckDuplicate(cnBookPrintOrder, "BookPOParent", "Code", "[Name]+[Type]", Trim(Text2.Text) & BookPOType, rstBookPOParent.Fields("Code").Value, False, FYCode) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
       Text3.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckEmpty(Text11.Text, False) Then 'Material Centre
       Text11.SetFocus
       CheckMandatoryFields = True
    End If
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Private Sub LoadOrder(ByVal strOrderCode As String)
    On Error GoTo ErrorHandler
    If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    rstBookPOChild05.Open "SELECT * FROM BookPOChild05 WHERE Code = '" & strOrderCode & "'", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookPOChild05.ActiveConnection = Nothing
    If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
    rstBookPOChild06.Open "SELECT * FROM BookPOChild06 WHERE Code = '" & strOrderCode & "'", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookPOChild06.ActiveConnection = Nothing
    If rstBookPOChild09.State = adStateOpen Then rstBookPOChild09.Close
    rstBookPOChild09.Open "SELECT * FROM BookPOChild09 WHERE Code = '" & strOrderCode & "'", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookPOChild09.ActiveConnection = Nothing
    If rstBookPOChild0901.State = adStateOpen Then rstBookPOChild0901.Close
    rstBookPOChild0901.Open "SELECT Book As ItemCode,M.Name As ItemName,ActualQuantity,[Ups/Plate],PrintingQuantity,BillingQuantity,FrontPrintingColor,BackPrintingColor FROM BookPOChild0901 T INNER JOIN BookMaster M ON T.Book=M.Code WHERE T.Code='" & strOrderCode & "' ORDER BY M.Name", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookPOChild0901.ActiveConnection = Nothing
    If rstBookPOChild07.State = adStateOpen Then rstBookPOChild07.Close
    rstBookPOChild07.Open "SELECT * FROM BookPOChild07 WHERE Code = '" & strOrderCode & "'", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookPOChild07.ActiveConnection = Nothing
    If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
    rstBookPOChild08.Open "SELECT * FROM BookPOChild08 WHERE Code='" & strOrderCode & "'", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookPOChild08.ActiveConnection = Nothing
    If rstBookPOChild0801.State = adStateOpen Then rstBookPOChild0801.Close
    rstBookPOChild0801.Open "SELECT * FROM BookPOChild0801 T WHERE Code = '" & strOrderCode & "'", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookPOChild0801.ActiveConnection = Nothing
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Print Order")
End Sub
Private Function UpdateOrder(ByVal strOption As String, Optional ByVal POType As String) As Boolean
    On Error GoTo ErrorHandler
    Dim VchCode As String
    VchCode = rstBookPOParent.Fields("Code").Value
    UpdateOrder = True
    If strOption = "D" Then
        With cnBookPrintOrder
            .Execute "IF OBJECT_ID('tempdb.dbo.#T05', 'U') IS NOT NULL DROP TABLE #T05" 'Multi Form Printing
            .Execute "SELECT * INTO #T05 FROM BookPOChild05 WHERE Code='" & VchCode & "'"
            .Execute "IF OBJECT_ID('tempdb.dbo.#T06', 'U') IS NOT NULL DROP TABLE #T06" 'Multi Element Printing
            .Execute "SELECT * INTO #T06 FROM BookPOChild06 WHERE Code='" & VchCode & "'"
            .Execute "IF OBJECT_ID('tempdb.dbo.#T07', 'U') IS NOT NULL DROP TABLE #T07" 'Miscellaneous Operations
            .Execute "SELECT * INTO #T07 FROM BookPOChild07 WHERE Code='" & VchCode & "'"
            .Execute "IF OBJECT_ID('tempdb.dbo.#T08', 'U') IS NOT NULL DROP TABLE #T08" 'Binding
            .Execute "SELECT * INTO #T08 FROM BookPOChild08 WHERE Code='" & VchCode & "'"
            .Execute "IF OBJECT_ID('tempdb.dbo.#T0801', 'U') IS NOT NULL DROP TABLE #T0801" 'BOM
            .Execute "SELECT * INTO #T0801 FROM BookPOChild0801 WHERE Code='" & VchCode & "'"
            .Execute "IF OBJECT_ID('tempdb.dbo.#T0901', 'U') IS NOT NULL DROP TABLE #T0901" 'Combo Format
            .Execute "SELECT * INTO #T0901 FROM BookPOChild0901 WHERE Code='" & VchCode & "'"
            .Execute "DELETE FROM BookPOChild05 WHERE Code='" & VchCode & "'"
            .Execute "DELETE FROM BookPOChild06 WHERE Code='" & VchCode & "'"
            .Execute "DELETE FROM BookPOChild07 WHERE Code='" & VchCode & "'"
            .Execute "DELETE FROM BookPOChild08 WHERE Code='" & VchCode & "'"
            .Execute "DELETE FROM BookPOChild0801 WHERE Code='" & VchCode & "'"
            .Execute "DELETE FROM BookPOChild09 WHERE Code='" & VchCode & "'"
            .Execute "DELETE FROM BookPOChild0901 WHERE Code='" & VchCode & "'"
        End With
    Else
        If POType = "1" And Not CheckEmpty(Text5.Text, False) Then 'Multi Form Printing
'            With rstBookPOChild05
'                cnBookPrintOrder.Execute "INSERT INTO BookPOChild05 VALUES ('" & VchCode & "','" & Format(.Fields("OrderDate").Value, "dd-MMM-yyyy") & "','" & Format(.Fields("TargetDate").Value, "dd-MMM-yyyy") & "','" & .Fields("Size1").Value & "','" & .Fields("Size2").Value & "','" & .Fields("Size4").Value & "','" & .Fields("Processing").Value & "','" & .Fields("Ref").Value & "','" & .Fields("PlateMaker").Value & "'," & Val(.Fields("ActualQuantity").Value) & "," & Val(.Fields("BillingQuantity01").Value) & "," & Val(.Fields("BillingQuantity02").Value) & "," & Val(.Fields("Pages1").Value) & "," & Val(.Fields("Forms1").Value) & "," & Val(.Fields("Forms1-¼").Value) & "," & Val(.Fields("Forms1-½").Value) & "," & Val(.Fields("Forms1-1").Value) & ",'" & .Fields("PlateType1").Value & "'," & _
'                                                                        Val(.Fields("TotalForms1-¼").Value) & "," & Val(.Fields("TotalForms1-½").Value) & "," & Val(.Fields("TotalForms1-1").Value) & "," & Val(.Fields("TotalPlates1-¼").Value) & "," & Val(.Fields("TotalPlates1-½").Value) & "," & Val(.Fields("TotalPlates1-1").Value) & "," & Val(.Fields("RevisedPlates1").Value) & "," & Val(.Fields("PrintRate1").Value) & "," & Val(.Fields("PrintAmount1").Value) & "," & Val(.Fields("PlateRate1").Value) & "," & Val(.Fields("PlateAmount1").Value) & "," & IIf(.Fields("PaperByParty1").Value, 1, 0) & ",'" & _
'                                                                        .Fields("Paper1").Value & "','" & IIf(.Fields("PaperByParty1").Value, BookPrinterCode, "000000") & "'," & Val(.Fields("PaperWastage1%").Value) & "," & Val(.Fields("PaperWastageMin1").Value) & "," & Val(.Fields("PaperWastageFinal1").Value) & "," & Val(.Fields("PaperConsumptionOther1").Value) & "," & Val(.Fields("PaperConsumptionSheets1").Value) & "," & Val(.Fields("PaperRate1").Value) & "," & Val(.Fields("PaperAmount1").Value) & "," & Val(.Fields("Forms/Sheet1-1").Value) & "," & Val(.Fields("Forms/Sheet2-1").Value) & "," & _
'                                                                        Val(.Fields("Pages2").Value) & "," & Val(.Fields("Forms2").Value) & "," & Val(.Fields("Forms2-¼").Value) & "," & Val(.Fields("Forms2-½").Value) & "," & Val(.Fields("Forms2-1").Value) & ",'" & .Fields("PlateType2").Value & "'," & Val(.Fields("TotalForms2-¼").Value) & "," & Val(.Fields("TotalForms2-½").Value) & "," & Val(.Fields("TotalForms2-1").Value) & "," & Val(.Fields("TotalPlates2-¼").Value) & "," & Val(.Fields("TotalPlates2-½").Value) & "," & Val(.Fields("TotalPlates2-1").Value) & "," & Val(.Fields("RevisedPlates2").Value) & "," & Val(.Fields("PrintRate2").Value) & "," & Val(.Fields("PrintAmount2").Value) & "," & Val(.Fields("PlateRate2").Value) & "," & Val(.Fields("PlateAmount2").Value) & "," & IIf(.Fields("PaperByParty2").Value, 1, 0) & ",'" & _
'                                                                        .Fields("Paper2").Value & "','" & IIf(.Fields("PaperByParty2").Value, BookPrinterCode, "000000") & "'," & Val(.Fields("PaperWastage2%").Value) & "," & Val(.Fields("PaperWastageMin2").Value) & "," & Val(.Fields("PaperWastageFinal2").Value) & "," & Val(.Fields("PaperConsumptionOther2").Value) & "," & Val(.Fields("PaperConsumptionSheets2").Value) & "," & Val(.Fields("PaperRate2").Value) & "," & Val(.Fields("PaperAmount2").Value) & "," & Val(.Fields("Forms/Sheet1-2").Value) & "," & Val(.Fields("Forms/Sheet2-2").Value) & "," & Val(.Fields("Pages4").Value) & "," & Val(.Fields("Forms4").Value) & "," & Val(.Fields("Forms4-¼").Value) & "," & Val(.Fields("Forms4-½").Value) & "," & Val(.Fields("Forms4-1").Value) & ",'" & .Fields("PlateType4").Value & "'," & _
'                                                                        Val(.Fields("TotalForms4-¼").Value) & "," & Val(.Fields("TotalForms4-½").Value) & "," & Val(.Fields("TotalForms4-1").Value) & "," & Val(.Fields("TotalPlates4-¼").Value) & "," & Val(.Fields("TotalPlates4-½").Value) & "," & Val(.Fields("TotalPlates4-1").Value) & "," & Val(.Fields("RevisedPlates4").Value) & "," & Val(.Fields("PrintRate4").Value) & "," & Val(.Fields("PrintAmount4").Value) & "," & Val(.Fields("PlateRate4").Value) & "," & Val(.Fields("PlateAmount4").Value) & "," & IIf(.Fields("PaperByParty4").Value, 1, 0) & ",'" & _
'                                                                        .Fields("Paper4").Value & "','" & IIf(.Fields("PaperByParty4").Value, BookPrinterCode, "000000") & "'," & Val(.Fields("PaperWastage4%").Value) & "," & Val(.Fields("PaperWastageMin4").Value) & "," & Val(.Fields("PaperWastageFinal4").Value) & "," & Val(.Fields("PaperConsumptionOther4").Value) & "," & Val(.Fields("PaperConsumptionSheets4").Value) & "," & Val(.Fields("PaperRate4").Value) & "," & Val(.Fields("PaperAmount4").Value) & "," & Val(.Fields("Forms/Sheet1-4").Value) & "," & Val(.Fields("Forms/Sheet2-4").Value) & "," & _
'                                                                        Val(.Fields("TotalPaperConsumption").Value) & ",'" & .Fields("Remarks").Value & "','" & .Fields("BillNo").Value & "'," & IIf(IsNull(.Fields("BillDate").Value), "Null", "'" & Format(.Fields("BillDate").Value, "dd-MMM-yyyy") & "'") & ",'" & .Fields("PBillNo").Value & "'," & IIf(IsNull(.Fields("PBillDate").Value), "Null", "'" & Format(.Fields("PBillDate").Value, "dd-MMM-yyyy") & "'") & "," & Val(.Fields("Adjustment").Value) & "," & Val(.Fields("PAdjustment").Value) & "," & Val(.Fields("RAdjustment").Value) & "," & Val(.Fields("VAT%").Value) & "," & Val(.Fields("VAT").Value) & "," & Val(.Fields("PVAT%").Value) & "," & Val(.Fields("PVAT").Value) & "," & Val(.Fields("RVAT%").Value) & "," & _
'                                                                        Val(.Fields("RVAT").Value) & "," & Val(.Fields("BillAmount").Value) & "," & Val(.Fields("PBillAmount").Value) & "," & Val(.Fields("RBillAmount").Value) & "," & Val(.Fields("PaidAmount").Value) & "," & Val(.Fields("PPaidAmount").Value) & ",'" & .Fields("Status").Value & "','" & .Fields("Narration").Value & "'," & IIf(IsNull(.Fields("BillFeedDate").Value), "Null", "'" & Format(.Fields("BillFeedDate").Value, "dd-MMM-yyyy") & "'") & ",'" & _
'                                                                        .Fields("AdjustmentRemarks").Value & "'," & IIf(IsNull(.Fields("ComputerName").Value), "Null", "'" & .Fields("ComputerName").Value & "'") & "," & Val(.Fields("CutOffSize1").Value) & "," & Val(.Fields("CutOffSize2").Value) & "," & Val(.Fields("CutOffSize4").Value) & ",0,0,0,0)"
'                Amount = Amount + Val(.Fields("PrintAmount1").Value) + Val(.Fields("PrintAmount2").Value) + Val(.Fields("PrintAmount4").Value) + Val(.Fields("PlateAmount1").Value) + Val(.Fields("PlateAmount2").Value) + Val(.Fields("PlateAmount4").Value) + Val(.Fields("PaperAmount1").Value) + Val(.Fields("PaperAmount2").Value) + Val(.Fields("PaperAmount4").Value) + Val(.Fields("Adjustment").Value) + Val(.Fields("PAdjustment").Value) + Val(.Fields("RAdjustment").Value)
'            End With
            With rstBookPOChild05
                cnBookPrintOrder.Execute "INSERT INTO BookPOChild05 VALUES ('" & rstBookPOParent.Fields("Code").Value & "','" & Format(.Fields("OrderDate").Value, "dd-MMM-yyyy") & "','" & Format(.Fields("TargetDate").Value, "dd-MMM-yyyy") & "','" & .Fields("SubItem").Value & "','" & .Fields("Element").Value & "','" & .Fields("ElementPrintName").Value & "','" & .Fields("FinishSize").Value & "','" & .Fields("Size").Value & "'," & IIf(.Fields("DuplexPrinting").Value, 1, 0) & ",'" & .Fields("Processing").Value & "','" & .Fields("Ref").Value & "','" & .Fields("PlateMaker").Value & "'," & Val(.Fields("ActualQuantity").Value) & "," & Val(.Fields("BillingQuantity").Value) & "," & Val(.Fields("Pages/PrintingForm").Value) & "," & Val(.Fields("Pages/Form").Value) & ",'" & .Fields("Color").Value & "'," & Val(.Fields("Pages").Value) & "," & _
                                                                        Val(.Fields("Forms").Value) & "," & Val(.Fields("Forms-¼").Value) & "," & Val(.Fields("Forms-½").Value) & "," & Val(.Fields("Forms-1-F&B").Value) & "," & Val(.Fields("Forms-1-W&T").Value) & ",'" & .Fields("PlateType").Value & "'," & Val(.Fields("TotalForms-¼").Value) & "," & Val(.Fields("TotalForms-½").Value) & "," & Val(.Fields("TotalForms-1-F&B").Value) & "," & Val(.Fields("TotalForms-1-W&T").Value) & "," & _
                                                                        Val(.Fields("TotalPlates-¼").Value) & "," & Val(.Fields("TotalPlates-½").Value) & "," & Val(.Fields("TotalPlates-1-F&B").Value) & "," & Val(.Fields("TotalPlates-1-W&T").Value) & "," & Val(.Fields("RevisedPlates").Value) & "," & Val(.Fields("aTotalPlates-¼").Value) & "," & Val(.Fields("aTotalPlates-½").Value) & "," & Val(.Fields("aTotalPlates-1-F&B").Value) & "," & Val(.Fields("aTotalPlates-1-W&T").Value) & "," & Val(.Fields("aRevisedPlates").Value) & "," & _
                                                                        Val(.Fields("PrintRate").Value) & "," & Val(.Fields("PrintAmount").Value) & "," & Val(.Fields("PlateRate").Value) & "," & Val(.Fields("PlateAmount").Value) & "," & IIf(.Fields("PaperByParty").Value, 1, 0) & ",'" & .Fields("Paper").Value & "','" & IIf(.Fields("PaperByParty").Value, BookPrinterCode, "000000") & "'," & Val(.Fields("CutOffSize").Value) & "," & _
                                                                        Val(.Fields("PaperWastage%").Value) & "," & Val(.Fields("PaperWastageMin").Value) & "," & Val(.Fields("Wastage/Set").Value) & "," & Val(.Fields("PaperWastageFinal").Value) & "," & Val(.Fields("PaperConsumptionOther").Value) & "," & Val(.Fields("PaperConsumptionSheets").Value) & "," & Val(.Fields("PaperConsumptionKg").Value) & "," & Val(.Fields("aPaperWastage%").Value) & "," & Val(.Fields("aPaperWastageMin").Value) & "," & Val(.Fields("aWastage/Set").Value) & "," & Val(.Fields("aPaperWastageFinal").Value) & "," & Val(.Fields("aPaperConsumptionOther").Value) & "," & Val(.Fields("aPaperConsumptionSheets").Value) & "," & _
                                                                        Val(.Fields("aPaperConsumptionKg").Value) & "," & Val(.Fields("PaperRate").Value) & "," & Val(.Fields("PaperAmount").Value) & "," & Val(.Fields("Forms/Sheet1").Value) & "," & Val(.Fields("Forms/Sheet2").Value) & ",'" & .Fields("Remarks").Value & "','" & _
                                                                        .Fields("BillNo").Value & "'," & IIf(IsNull(.Fields("BillDate").Value), "Null", "'" & Format(.Fields("BillDate").Value, "dd-MMM-yyyy") & "'") & ",'" & .Fields("PBillNo").Value & "'," & IIf(IsNull(.Fields("PBillDate").Value), "Null", "'" & Format(.Fields("PBillDate").Value, "dd-MMM-yyyy") & "'") & "," & Val(.Fields("Adjustment").Value) & "," & Val(.Fields("PAdjustment").Value) & "," & Val(.Fields("RAdjustment").Value) & "," & Val(.Fields("VAT%").Value) & "," & Val(.Fields("VAT").Value) & "," & Val(.Fields("PVAT%").Value) & "," & Val(.Fields("PVAT").Value) & "," & Val(.Fields("RVAT%").Value) & "," & Val(.Fields("RVAT").Value) & "," & Val(.Fields("BillAmount").Value) & "," & Val(.Fields("PBillAmount").Value) & "," & _
                                                                        Val(.Fields("RBillAmount").Value) & "," & Val(.Fields("PaidAmount").Value) & "," & Val(.Fields("PPaidAmount").Value) & ",'" & .Fields("Status").Value & "','" & .Fields("Narration").Value & "','" & .Fields("AdjustmentRemarks").Value & "'," & Val(.Fields("DeliveredQuantityC").Value) & "," & Val(.Fields("DeliveredQuantityB").Value) & "," & Val(.Fields("BilledMFC").Value) & "," & Val(.Fields("BilledMFB").Value) & ")"
                Amount = Amount + Val(.Fields("PrintAmount").Value) + Val(.Fields("PlateAmount").Value) + Val(.Fields("PaperAmount").Value) + Val(.Fields("Adjustment").Value) + Val(.Fields("PAdjustment").Value) + Val(.Fields("RAdjustment").Value)
            End With
        ElseIf POType = "21" And Not CheckEmpty(Text6.Text, False) Then 'Multi Element Printing
'            With rstBookPOChild06
'                cnBookPrintOrder.Execute "INSERT INTO BookPOChild06 VALUES ('" & VchCode & "','" & Format(.Fields("OrderDate").Value, "dd-MMM-yyyy") & "','" & Format(.Fields("TargetDate").Value, "dd-MMM-yyyy") & "','" & .Fields("Element").Value & "'," & .Fields("Pages").Value & ",'" & .Fields("FinishSize").Value & "','" & .Fields("Size").Value & "','" & .Fields("Processing").Value & "','" & .Fields("ProcessingBack").Value & "','" & .Fields("Imposition").Value & "','" & .Fields("Ref").Value & "','" & .Fields("PlateMaker").Value & "'," & .Fields("FrontPrintingType").Value & "," & .Fields("BackPrintingType").Value & ",'" & .Fields("PlateType").Value & "','" & .Fields("PlateTypeBack").Value & "'," & _
'                                                                        Val(.Fields("ActualQuantity").Value) & "," & Val(.Fields("BillingQuantity").Value) & "," & Val(.Fields("Ups").Value) & "," & Val(.Fields("Sets").Value) & "," & Val(.Fields("TotalForms").Value) & "," & Val(.Fields("TotalPlates").Value) & "," & Val(.Fields("TotalPlatesBack").Value) & "," & Val(.Fields("PrintRate").Value) & "," & Val(.Fields("PrintRateBack").Value) & "," & Val(.Fields("PrintAmount").Value) & "," & Val(.Fields("PlateRate").Value) & "," & Val(.Fields("PlateRateBack").Value) & "," & Val(.Fields("PlateAmount").Value) & "," & IIf(.Fields("PaperByParty").Value, 1, 0) & ",'" & .Fields("Paper").Value & "','" & IIf(.Fields("PaperByParty").Value, TitlePrinterCode, "000000") & "'," & _
'                                                                        Val(.Fields("CutOffSize").Value) & "," & Val(.Fields("Titles/Sheet2").Value) & "," & Val(.Fields("PaperWastage%").Value) & "," & Val(.Fields("PaperWastage%Back").Value) & "," & Val(.Fields("PaperWastageMin").Value) & "," & Val(.Fields("PaperWastageMinBack").Value) & "," & Val(.Fields("PaperWastageFinal").Value) & "," & Val(.Fields("PaperConsumptionOther").Value) & "," & Val(.Fields("PaperConsumptionSheets").Value) & "," & Val(.Fields("PaperConsumptionKg").Value) & "," & Val(.Fields("PaperRate").Value) & "," & Val(.Fields("PaperAmount").Value) & ",'" & .Fields("Remarks").Value & "','" & .Fields("BillNo").Value & "'," & IIf(IsNull(.Fields("BillDate").Value), "Null", "'" & Format(.Fields("BillDate").Value, "dd-MMM-yyyy") & "'") & ",'" & .Fields("PBillNo").Value & "'," & IIf(IsNull(.Fields("PBillDate").Value), "Null", "'" & Format(.Fields("PBillDate").Value, "dd-MMM-yyyy") & "'") & "," & _
'                                                                        Val(.Fields("Adjustment").Value) & "," & Val(.Fields("PAdjustment").Value) & "," & Val(.Fields("RAdjustment").Value) & "," & Val(.Fields("VAT%").Value) & "," & Val(.Fields("VAT").Value) & "," & Val(.Fields("PVAT%").Value) & "," & Val(.Fields("PVAT").Value) & "," & Val(.Fields("RVAT%").Value) & "," & Val(.Fields("RVAT").Value) & "," & Val(.Fields("BillAmount").Value) & "," & Val(.Fields("PBillAmount").Value) & "," & Val(.Fields("RBillAmount").Value) & "," & Val(.Fields("PaidAmount").Value) & "," & Val(.Fields("PPaidAmount").Value) & ",'" & _
'                                                                        .Fields("Status").Value & "','" & .Fields("Narration").Value & "','" & .Fields("AdjustmentRemarks").Value & "'," & IIf(IsNull(.Fields("ComputerName").Value), "Null", "'" & .Fields("ComputerName").Value & "'") & ",0,0,0,0)"
'                Amount = Amount + Val(.Fields("PrintAmount").Value) + Val(.Fields("PlateAmount").Value) + Val(.Fields("PaperAmount").Value) + Val(.Fields("Adjustment").Value) + Val(.Fields("PAdjustment").Value) + Val(.Fields("RAdjustment").Value)
'            End With
            With rstBookPOChild06
                cnBookPrintOrder.Execute "INSERT INTO BookPOChild06 VALUES ('" & VchCode & "','" & Format(.Fields("OrderDate").Value, "dd-MMM-yyyy") & "','" & Format(.Fields("TargetDate").Value, "dd-MMM-yyyy") & "','" & .Fields("SubItem").Value & "','" & .Fields("Element").Value & "','" & .Fields("ElementPrintName").Value & "'," & .Fields("Pages").Value & ",'" & .Fields("FinishSize").Value & "','" & .Fields("Size").Value & "','" & .Fields("Processing").Value & "','" & .Fields("ProcessingBack").Value & "','" & .Fields("Imposition").Value & "','" & .Fields("Ref").Value & "','" & .Fields("PlateMaker").Value & "','" & .Fields("FrontPrintingType").Value & "','" & .Fields("BackPrintingType").Value & "','" & .Fields("PlateType").Value & "','" & .Fields("PlateTypeBack").Value & "'," & _
                                                                        Val(.Fields("ActualQuantity").Value) & "," & Val(.Fields("BillingQuantity").Value) & "," & Val(.Fields("Ups").Value) & "," & Val(.Fields("Sets").Value) & "," & Val(.Fields("TotalForms").Value) & "," & Val(.Fields("TotalPlates").Value) & "," & Val(.Fields("TotalPlatesBack").Value) & "," & Val(.Fields("aTotalPlates").Value) & "," & Val(.Fields("aTotalPlatesBack").Value) & "," & Val(.Fields("PrintRate").Value) & "," & Val(.Fields("PrintRateBack").Value) & "," & Val(.Fields("PrintAmount").Value) & "," & Val(.Fields("PlateRate").Value) & "," & Val(.Fields("PlateRateBack").Value) & "," & Val(.Fields("PlateAmount").Value) & "," & IIf(.Fields("PaperByParty").Value, 1, 0) & ",'" & .Fields("Paper").Value & "','" & IIf(.Fields("PaperByParty").Value, TitlePrinterCode, "000000") & "'," & _
                                                                        Val(.Fields("CutOffSize").Value) & "," & Val(.Fields("Titles/Sheet2").Value) & "," & Val(.Fields("PaperWastage%").Value) & "," & Val(.Fields("PaperWastage%Back").Value) & "," & Val(.Fields("PaperWastageMin").Value) & "," & Val(.Fields("PaperWastageMinBack").Value) & "," & Val(.Fields("Wastage/Set").Value) & "," & Val(.Fields("PaperWastageFinal").Value) & "," & Val(.Fields("PaperConsumptionOther").Value) & "," & Val(.Fields("PaperConsumptionSheets").Value) & "," & Val(.Fields("PaperConsumptionKg").Value) & "," & _
                                                                        Val(.Fields("aPaperWastage%").Value) & "," & Val(.Fields("aPaperWastage%Back").Value) & "," & Val(.Fields("aPaperWastageMin").Value) & "," & Val(.Fields("aPaperWastageMinBack").Value) & "," & Val(.Fields("aPaperWastage/Set").Value) & "," & Val(.Fields("aPaperWastageFinal").Value) & "," & Val(.Fields("aPaperConsumptionOther").Value) & "," & Val(.Fields("aPaperConsumptionSheets").Value) & "," & Val(.Fields("aPaperConsumptionKg").Value) & "," & _
                                                                        Val(.Fields("PaperRate").Value) & "," & Val(.Fields("PaperAmount").Value) & ",'" & .Fields("Remarks").Value & "','" & .Fields("BillNo").Value & "'," & IIf(IsNull(.Fields("BillDate").Value), "Null", "'" & Format(.Fields("BillDate").Value, "dd-MMM-yyyy") & "'") & ",'" & .Fields("PBillNo").Value & "'," & IIf(IsNull(.Fields("PBillDate").Value), "Null", "'" & Format(.Fields("PBillDate").Value, "dd-MMM-yyyy") & "'") & "," & _
                                                                        Val(.Fields("Adjustment").Value) & "," & Val(.Fields("PAdjustment").Value) & "," & Val(.Fields("RAdjustment").Value) & "," & Val(.Fields("VAT%").Value) & "," & Val(.Fields("VAT").Value) & "," & Val(.Fields("PVAT%").Value) & "," & Val(.Fields("PVAT").Value) & "," & Val(.Fields("RVAT%").Value) & "," & Val(.Fields("RVAT").Value) & "," & Val(.Fields("BillAmount").Value) & "," & Val(.Fields("PBillAmount").Value) & "," & Val(.Fields("RBillAmount").Value) & "," & Val(.Fields("PaidAmount").Value) & "," & Val(.Fields("PPaidAmount").Value) & ",'" & _
                                                                        .Fields("Status").Value & "','" & .Fields("Narration").Value & "','" & .Fields("AdjustmentRemarks").Value & "'," & IIf(IsNull(.Fields("ComputerName").Value), "Null", "'" & .Fields("ComputerName").Value & "'") & "," & Val(.Fields("DeliveredQuantityC").Value) & "," & Val(.Fields("DeliveredQuantityB").Value) & "," & Val(.Fields("BilledMEC").Value) & "," & Val(.Fields("BilledMEB").Value) & ")"
                Amount = Amount + Val(.Fields("PrintAmount").Value) + Val(.Fields("PlateAmount").Value) + Val(.Fields("PaperAmount").Value) + Val(.Fields("Adjustment").Value) + Val(.Fields("PAdjustment").Value) + Val(.Fields("RAdjustment").Value)
            End With
        ElseIf POType = "22" And Not CheckEmpty(Text9.Text, False) Then 'Combo Format Printing
            With rstBookPOChild09
                cnBookPrintOrder.Execute "INSERT INTO BookPOChild09 VALUES ('" & VchCode & "','" & Format(.Fields("OrderDate").Value, "dd-MMM-yyyy") & "','" & Format(.Fields("TargetDate").Value, "dd-MMM-yyyy") & "'," & Val(.Fields("ActualQuantity").Value) & "," & Val(.Fields("MaxPrintingQuantity").Value) & ",'" & .Fields("Plate").Value & "','" & .Fields("PlateType").Value & "','" & .Fields("Imposition").Value & "','" & .Fields("Calculation").Value & "','" & .Fields("Size").Value & "','" & .Fields("PlateMaker").Value & "'," & Val(.Fields("BillingQuantity").Value) & "," & _
                                                                        Val(.Fields("TotalFormsFront").Value) & "," & Val(.Fields("TotalFormsBack").Value) & "," & Val(.Fields("PrintRateFront").Value) & "," & Val(.Fields("PrintRateBack").Value) & "," & Val(.Fields("PrintAmountBT").Value) & "," & Val(.Fields("Adjustment").Value) & "," & Val(.Fields("GST%").Value) & "," & Val(.Fields("GST").Value) & "," & Val(.Fields("PrintAmount").Value) & "," & _
                                                                        Val(.Fields("FrontPrintingColor").Value) & "," & Val(.Fields("BackPrintingColor").Value) & "," & Val(.Fields("TotalPlates").Value) & "," & Val(.Fields("PlateRate").Value) & "," & Val(.Fields("PlateAmountBT").Value) & "," & Val(.Fields("PAdjustment").Value) & "," & Val(.Fields("PGST%").Value) & "," & Val(.Fields("PGST").Value) & "," & Val(.Fields("PlateAmount").Value) & "," & IIf(.Fields("PaperByParty").Value, 1, 0) & ",'" & .Fields("Paper").Value & "','" & IIf(.Fields("PaperByParty").Value, TitlePrinterCode, "000000") & "'," & Val(.Fields("Ups/Sheet").Value) & "," & Val(.Fields("PaperConsumption").Value) & "," & Val(.Fields("PaperWastage%").Value) & "," & _
                                                                        Val(.Fields("PaperWastageMin").Value) & "," & Val(.Fields("PaperWastageMax").Value) & "," & Val(.Fields("PaperWastageFinal").Value) & "," & Val(.Fields("PaperConsumptionSheets").Value) & "," & Val(.Fields("PaperConsumptionOther").Value) & "," & Val(.Fields("PaperRate").Value) & "," & Val(.Fields("PaperAmountBT").Value) & "," & Val(.Fields("RAdjustment").Value) & "," & Val(.Fields("RGST%").Value) & "," & Val(.Fields("RGST").Value) & "," & Val(.Fields("PaperAmount").Value) & ",'" & _
                                                                        .Fields("Remarks").Value & "','" & .Fields("BillNo").Value & "'," & IIf(IsNull(.Fields("BillDate").Value), "Null", "'" & Format(.Fields("BillDate").Value, "dd-MMM-yyyy") & "'") & ",'" & .Fields("PBillNo").Value & "'," & IIf(IsNull(.Fields("PBillDate").Value), "Null", "'" & Format(.Fields("PBillDate").Value, "dd-MMM-yyyy") & "'") & "," & Val(.Fields("PaidAmount").Value) & "," & Val(.Fields("PPaidAmount").Value) & ",'" & _
                                                                        .Fields("Status").Value & "','" & .Fields("Narration").Value & "'," & IIf(IsNull(.Fields("BillFeedDate").Value), "Null", "'" & Format(.Fields("BillFeedDate").Value, "dd-MMM-yyyy") & "'") & ",'" & .Fields("ComputerName").Value & "'," & Val(.Fields("CutOffSize").Value) & ")"
                Amount = Amount + Val(.Fields("PrintAmount").Value) + Val(.Fields("PlateAmount").Value) + Val(.Fields("PaperAmount").Value) + Val(.Fields("Adjustment").Value) + Val(.Fields("PAdjustment").Value) + Val(.Fields("RAdjustment").Value)
            End With
            With rstBookPOChild0901
                If .RecordCount > 0 Then .MoveFirst
                Do Until .EOF
                    cnBookPrintOrder.Execute "INSERT INTO BookPOChild0901 VALUES ('" & VchCode & "','" & .Fields("ItemCode").Value & "'," & Val(.Fields("ActualQuantity").Value) & "," & Val(.Fields("Ups/Plate").Value) & "," & Val(.Fields("PrintingQuantity").Value) & "," & Val(.Fields("BillingQuantity").Value) & "," & Val(.Fields("FrontPrintingColor").Value) & "," & Val(.Fields("BackPrintingColor").Value) & ",0,0,0,0)"
                    .MoveNext
                Loop
            End With
        ElseIf POType = "3" And Not CheckEmpty(Text7.Text, False) Then 'Miscellaneous Operations
            With rstBookPOChild07
                cnBookPrintOrder.Execute "INSERT INTO BookPOChild07 VALUES ('" & VchCode & "','" & Format(.Fields("OrderDate").Value, "dd-MMM-yyyy") & "','" & Format(.Fields("TargetDate").Value, "dd-MMM-yyyy") & "','" & .Fields("SubItem").Value & "','" & .Fields("Element").Value & "','" & .Fields("Operation").Value & "'," & Val(.Fields("Number").Value) & ",'" & .Fields("OperationCountName").Value & "','" & .Fields("Size").Value & "'," & Val(.Fields("Quantity").Value) & ",'" & .Fields("CalcMode").Value & "'," & Val(.Fields("CalcValue").Value) & "," & Val(.Fields("Rate").Value) & "," & Val(.Fields("Amount").Value) & "," & Val(.Fields("Adjustment").Value) & "," & Val(.Fields("GST%").Value) & "," & Val(.Fields("GST").Value) & "," & Val(.Fields("BillAmount").Value) & ",'" & .Fields("Remarks").Value & "','" & .Fields("BillNo").Value & "'," & _
                                                                        IIf(IsNull(.Fields("BillDate").Value), "Null", "'" & Format(.Fields("BillDate").Value, "dd-MMM-yyyy") & "'") & "," & Val(.Fields("PaidAmount").Value) & ",'" & .Fields("Status").Value & "','" & .Fields("Narration").Value & "',0,0,0,0)"
                Amount = Amount + Val(.Fields("Amount").Value) + Val(.Fields("Adjustment").Value)
            End With
        ElseIf POType = "4" And Not CheckEmpty(Text8.Text, False) Then 'Binding
            With rstBookPOChild08
'                cnBookPrintOrder.Execute "INSERT INTO BookPOChild08 VALUES ('" & VchCode & "','" & Format(.Fields("OrderDate").Value, "dd-MMM-yyyy") & "','" & Format(.Fields("TargetDate").Value, "dd-MMM-yyyy") & "','" & .Fields("BindingType").Value & "'," & Val(.Fields("BindingForms").Value) & "," & Val(.Fields("ExtraForms").Value) & "," & Val(.Fields("ActualQuantity").Value) & "," & Val(.Fields("BillingQuantity").Value) & "," & Val(.Fields("AdjustQuantity").Value) & "," & Val(.Fields("FormFoldRate").Value) & "," & Val(.Fields("FormStitchRate").Value) & "," & Val(.Fields("FormPasteRate").Value) & "," & _
'                                                                        Val(.Fields("Rate/Book").Value) & "," & Val(.Fields("LooseQty/Box").Value) & "," & Val(.Fields("ExtraLooseQty").Value) & "," & Val(.Fields("TotalLooseQty").Value) & "," & Val(.Fields("Qty/Pkt").Value) & "," & Val(.Fields("TotalPkts").Value) & "," & Val(.Fields("Pkt/Box").Value) & "," & Val(.Fields("TotalBoxes").Value) & "," & _
'                                                                        Val(.Fields("PktPackRate").Value) & "," & Val(.Fields("BoxPackRate").Value) & "," & Val(.Fields("CartageRate").Value) & ",'" & .Fields("Remarks").Value & "','" & .Fields("BillNo").Value & "'," & IIf(IsNull(.Fields("BillDate").Value), "Null", "'" & Format(.Fields("BillDate").Value, "dd-MMM-yyyy") & "'") & "," & Val(.Fields("Adjustment").Value) & "," & Val(.Fields("VAT%").Value) & "," & Val(.Fields("VAT").Value) & "," & Val(.Fields("BillAmount").Value) & "," & Val(.Fields("PaidAmount").Value) & ",'" & .Fields("Status").Value & "','" & .Fields("Narration").Value & "','" & .Fields("DNDetails").Value & "','" & _
'                                                                        .Fields("CNDetails").Value & "'," & IIf(IsNull(.Fields("BillFeedDate").Value), "Null", "'" & Format(.Fields("BillFeedDate").Value, "dd-MMM-yyyy") & "'") & ",'" & .Fields("AdjustmentRemarks").Value & "'," & IIf(IsNull(.Fields("ComputerName").Value), "Null", "'" & .Fields("ComputerName").Value & "'") & ",0,0,0,0)"
'                Amount = Amount + (Val(.Fields("FormFoldRate").Value) * Val(.Fields("ActualQuantity").Value) * (Val(.Fields("BindingForms").Value) + Val(.Fields("ExtraForms").Value))) / 1000 + (Val(.Fields("FormPasteRate").Value) * Val(.Fields("ActualQuantity").Value)) / 1000 + (Val(.Fields("FormStitchRate").Value) * Val(.Fields("ActualQuantity").Value) * (Val(.Fields("BindingForms").Value) + Val(.Fields("ExtraForms").Value))) / 1000 + Val(.Fields("Rate/Book").Value) * Val(.Fields("ActualQuantity").Value) + Val(.Fields("TotalPkts").Value) * Val(.Fields("PktPackRate").Value) + Val(.Fields("TotalBoxes").Value) * Val(.Fields("BoxPackRate").Value) + Val(.Fields("TotalBoxes").Value) * Val(.Fields("CartageRate").Value)
'                Amount = Amount + Val(.Fields("Adjustment").Value)
                cnBookPrintOrder.Execute "INSERT INTO BookPOChild08 VALUES ('" & VchCode & "','" & Format(.Fields("OrderDate").Value, "dd-MMM-yyyy") & "','" & Format(.Fields("TargetDate").Value, "dd-MMM-yyyy") & "','" & .Fields("SubItem").Value & "','" & .Fields("BindingType").Value & "','" & .Fields("Element").Value & "','" & .Fields("BinderyProcess").Value & "'," & Val(.Fields("Number").Value) & ",'" & .Fields("OperationCountName").Value & "','" & .Fields("Size").Value & "'," & Val(.Fields("Quantity").Value) & ",'" & .Fields("CalcMode").Value & "'," & Val(.Fields("CalcValue").Value) & "," & Val(.Fields("Rate").Value) & "," & Val(.Fields("Amount").Value) & "," & Val(.Fields("Adjustment").Value) & "," & Val(.Fields("GST%").Value) & "," & Val(.Fields("GST").Value) & "," & Val(.Fields("BillAmount").Value) & ",'" & .Fields("Remarks").Value & "','" & .Fields("BillNo").Value & "'," & _
                                                                        IIf(IsNull(.Fields("BillDate").Value), "Null", "'" & Format(.Fields("BillDate").Value, "dd-MMM-yyyy") & "'") & "," & Val(.Fields("PaidAmount").Value) & ",'" & .Fields("Status").Value & "','" & .Fields("Narration").Value & "',0,0,0,0)"
                Amount = Amount + Val(.Fields("Amount").Value) + Val(.Fields("Adjustment").Value)
            End With
        ElseIf POType = "0" Then 'BOM
            With rstBookPOChild0801
                cnBookPrintOrder.Execute "INSERT INTO BookPOChild0801 Values ('" & VchCode & "','" & .Fields("Category").Value & "','" & .Fields("Item").Value & "'," & Val(.Fields("Consumption/Item").Value) & "," & Val(.Fields("OrderQuantity").Value) & "," & Val(.Fields("TotalConsumption").Value) & ",'" & .Fields("Vendor").Value & "'," & Val(.Fields("Rate").Value) & "," & Val(.Fields("Amount").Value) & ",0,0,0,0)"
                Amount = Amount + Val(.Fields("Amount").Value)
            End With
        End If
    End If
    Exit Function
ErrorHandler:
    DisplayError (Err.Description)
    UpdateOrder = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Book" Then
        rstBookPOList.Filter = "[BookName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub Command1_Click()    'Multi form format
If Left(BookPOType, 1) = "F" Or Left(BookPOType, 1) = "R" Then
    If CheckEmpty(Text5.Text, False) Then Exit Sub
    With FrmBookPOChild05
        .VchCode = CheckNull(rstBookPOParent.Fields("Code").Value)
        .VchType = BookPOType
        .PartyCode = BookPrinterCode
        Set .rstBookPOChild05 = rstBookPOChild05
        .Mh3dLabel51.Caption = IIf(Right(BookPOType, 1) = "P", " Paper Supplied", " Paper by Party")
        On Error Resume Next
        Load FrmBookPOChild05
        If Err.Number <> 364 Then .Show vbModal: MhRealInput8_Validate False
    End With

'    If CheckEmpty(Text5.Text, False) Then Exit Sub
'    If rstBookPOChild05.RecordCount = 0 Then Call AddRecord(rstBookPOChild05)
'    Set FrmBookPOChild05.rstBookPOChild05 = rstBookPOChild05
'    FrmBookPOChild05.VchType = BookPOType
'    FrmBookPOChild05.PrinterCode = BookPrinterCode
'    If Right(BookPOType, 1) = "P" Then FrmBookPOChild05.Mh3dLabel46.Caption = " Paper Supplied" Else FrmBookPOChild05.Mh3dLabel46.Caption = " Paper by Party"
'    On Error Resume Next
'    Load FrmBookPOChild05
'    If Err.Number <> 364 Then
'        If rstBookPOParent.Fields("BookPrinter").Value <> "" Then
'            If Left(BookPOType, 1) <> "O" Then
'                If blnRecordExist And AllowTransactionsModification = 0 Then
'                    If Not CheckEmpty(FrmBookPOChild05.Text8.Text, False) Then
'                        Dim O As Object
'                        For Each O In FrmBookPOChild05
'                                If TypeName(O) = "TextBox" Then
'                                    O.Locked = True
'                                ElseIf TypeName(O) = "TDBNumber" Then
'                                    O.ReadOnly = True
'                                ElseIf TypeName(O) = "ComboBox" Then
'                                    O.Enabled = False
'                                ElseIf TypeName(O) = "TDBDate" Then
'                                    O.ReadOnly = True
'                                End If
'                        Next
'                    End If
'                End If
'            End If
'        End If
'        FrmBookPOChild05.Show vbModal
'        MhRealInput8_Validate False
'    End If
'    On Error GoTo 0
'    If AbortPO Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
ElseIf Left(BookPOType, 1) = "D" Then
'    If CheckEmpty(Text5.Text, False) Then Exit Sub
'    With FrmBookPOChild02
'        .VchCode = CheckNull(rstBookPOParent.Fields("Code").Value)
'        .VchType = BookPOType
'        .PartyCode = BookPrinterCode
'        Set .rstBookPOChild05 = rstBookPOChild05
'        .Mh3dLabel51.Caption = IIf(Right(BookPOType, 1) = "P", " Paper Supplied", " Paper by Party")
'        On Error Resume Next
'        Load FrmBookPOChild02
'        If Err.Number <> 364 Then .Show vbModal: MhRealInput8_Validate False
'    End With
End If
End Sub
Private Sub Command2_Click()    'Multi Element Format
    If CheckEmpty(Text9.Text, False) Then Exit Sub
    If rstBookPOChild09.RecordCount = 0 Then Call AddRecord(rstBookPOChild09)
    Set FrmBookPOChild09.rstBookPOChild09 = rstBookPOChild09
    Set FrmBookPOChild09.rstBookPOChild0901 = rstBookPOChild0901
    FrmBookPOChild09.SizeCode = rstBookList.Fields("TitleSizeCode").Value
    FrmBookPOChild09.PrinterCode = TitlePrinterCode
    rstAccountList.MoveFirst
    rstAccountList.Find "[Code]='" & TitlePrinterCode & "'"
    FrmBookPOChild09.RoundOffQty = rstAccountList.Fields("RoundOffQty").Value
    If Right(BookPOType, 1) = "P" Then FrmBookPOChild09.Mh3dLabel50.Caption = " Paper Supplied" Else FrmBookPOChild09.Mh3dLabel50.Caption = " Paper by Party"
    On Error Resume Next
    Load FrmBookPOChild09
    If Err.Number <> 364 Then
        If rstBookPOParent.Fields("TitlePrinter").Value <> "" Then
            If Left(BookPOType, 1) <> "O" Then
                If blnRecordExist And AllowTransactionsModification = 0 Then
                    If Not CheckEmpty(FrmBookPOChild09.Text8.Text, False) Then
                        Dim O As Object
                        For Each O In FrmBookPOChild09
                            If TypeName(O) = "TextBox" Then
                                O.Locked = True
                            ElseIf TypeName(O) = "TDBNumber" Then
                                O.ReadOnly = True
                            ElseIf TypeName(O) = "ComboBox" Then
                                O.Enabled = False
                            ElseIf TypeName(O) = "TDBDate" Then
                                O.ReadOnly = True
                            End If
                        Next
                    End If
                End If
            End If
        End If
        FrmBookPOChild09.Show vbModal
        MhRealInput8_Validate False
    End If
    On Error GoTo 0
    If AbortPO Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
End Sub
Private Sub Command5_Click()
If Left(BookPOType, 1) = "F" Or Left(BookPOType, 1) = "R" Then
    If CheckEmpty(Text6.Text, False) Then Exit Sub
    With FrmBookPOChild06
        .VchCode = CheckNull(rstBookPOParent.Fields("Code").Value)
        .VchType = BookPOType
        .PartyCode = TitlePrinterCode
        With rstAccountList
            .MoveFirst
            .Find "[Code]='" & TitlePrinterCode & "'"
            FrmBookPOChild06.RoundOffQty = .Fields("RoundOffQty").Value
        End With
        Set .rstBookPOChild06 = rstBookPOChild06
        .Mh3dLabel50.Caption = IIf(Right(BookPOType, 1) = "P", " Paper Supplied", " Paper by Party")
        On Error Resume Next
        Load FrmBookPOChild06
        If Err.Number <> 364 Then .Show vbModal: MhRealInput8_Validate False
    End With
'    If CheckEmpty(Text6.Text, False) Then Exit Sub
''    If rstBookPOChild06.RecordCount = 0 Then Call AddRecord(rstBookPOChild06)
'    Set FrmBookPOChild06.rstBookPOChild06 = rstBookPOChild06
'    FrmBookPOChild06.VchCode = CheckNull(rstBookPOParent.Fields("Code").Value)
'    FrmBookPOChild06.VchType = BookPOType
'    FrmBookPOChild06.PartyCode = TitlePrinterCode
'    rstAccountList.MoveFirst
'    rstAccountList.Find "[Code]='" & TitlePrinterCode & "'"
'    FrmBookPOChild06.RoundOffQty = rstAccountList.Fields("RoundOffQty").Value
'    FrmBookPOChild06.FinalQuantity = MhRealInput3.Value
'    If Right(BookPOType, 1) = "P" Then FrmBookPOChild06.Mh3dLabel50.Caption = " Paper Supplied" Else FrmBookPOChild06.Mh3dLabel50.Caption = " Paper by Party"
'    On Error Resume Next
'    Load FrmBookPOChild06
'    If Err.Number <> 364 Then
'        If rstBookPOParent.Fields("TitlePrinter").Value <> "" Then
'            If Left(BookPOType, 1) <> "O" Then
'                If blnRecordExist And AllowTransactionsModification = 0 Then
'                    If Not CheckEmpty(FrmBookPOChild06.Text8.Text, False) Then
'                        Dim O As Object
'                        For Each O In FrmBookPOChild06
'                            If TypeName(O) = "TextBox" Then
'                                O.Locked = True
'                            ElseIf TypeName(O) = "TDBNumber" Then
'                                O.ReadOnly = True
'                            ElseIf TypeName(O) = "ComboBox" Then
'                                O.Enabled = False
'                            ElseIf TypeName(O) = "TDBDate" Then
'                                O.ReadOnly = True
'                            End If
'                        Next
'                    End If
'                End If
'            End If
'        End If
'        FrmBookPOChild06.Show vbModal
'        MhRealInput8_Validate False
'    End If
'    On Error GoTo 0
'    If AbortPO Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
ElseIf Left(BookPOType, 1) = "D" Then
    If CheckEmpty(Text6.Text, False) Then Exit Sub
    With FrmBookPOChild01
        .VchCode = CheckNull(rstBookPOParent.Fields("Code").Value)
        .VchType = BookPOType
        .PartyCode = TitlePrinterCode
        With rstAccountList
            .MoveFirst
            .Find "[Code]='" & TitlePrinterCode & "'"
            FrmBookPOChild01.RoundOffQty = .Fields("RoundOffQty").Value
        End With
        Set .rstBookPOChild06 = rstBookPOChild06
        .Mh3dLabel50.Caption = IIf(Right(BookPOType, 1) = "P", " Paper Supplied", " Paper by Party")
        On Error Resume Next
        Load FrmBookPOChild01
        If Err.Number <> 364 Then .Show vbModal: MhRealInput8_Validate False
    End With
'    If CheckEmpty(Text6.Text, False) Then Exit Sub
'    Set FrmBookPOChild01.rstBookPOChild06 = rstBookPOChild06
'    FrmBookPOChild01.VchCode = CheckNull(rstBookPOParent.Fields("Code").Value)
'    FrmBookPOChild01.VchType = BookPOType
'    FrmBookPOChild01.PartyCode = TitlePrinterCode
'    rstAccountList.MoveFirst
'    rstAccountList.Find "[Code]='" & TitlePrinterCode & "'"
'    FrmBookPOChild01.RoundOffQty = rstAccountList.Fields("RoundOffQty").Value
'    FrmBookPOChild01.FinalQuantity = MhRealInput3.Value
'    If Right(BookPOType, 1) = "P" Then FrmBookPOChild01.Mh3dLabel50.Caption = " Paper Supplied" Else FrmBookPOChild01.Mh3dLabel50.Caption = " Paper by Party"
'    On Error Resume Next
'    Load FrmBookPOChild01
'    If Err.Number <> 364 Then
'        If rstBookPOParent.Fields("TitlePrinter").Value <> "" Then
'            If Left(BookPOType, 1) <> "O" Then
'                If blnRecordExist And AllowTransactionsModification = 0 Then
'                    If Not CheckEmpty(FrmBookPOChild01.Text8.Text, False) Then
'                        Dim O As Object
'                        For Each O In FrmBookPOChild01
'                            If TypeName(O) = "TextBox" Then
'                                O.Locked = True
'                            ElseIf TypeName(O) = "TDBNumber" Then
'                                O.ReadOnly = True
'                            ElseIf TypeName(O) = "ComboBox" Then
'                                O.Enabled = False
'                            ElseIf TypeName(O) = "TDBDate" Then
'                                O.ReadOnly = True
'                            End If
'                        Next
'                    End If
'                End If
'            End If
'        End If
'        FrmBookPOChild01.Show vbModal
'        MhRealInput8_Validate False
'    End If
'    On Error GoTo 0
'    If AbortPO Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
End If
End Sub
Private Sub Command3_Click()
    If CheckEmpty(Text7.Text, False) Then Exit Sub
    If rstBookPOChild07.RecordCount = 0 Then Call AddRecord(rstBookPOChild07)
    Set FrmBookPOChild07.rstBookPOChild07 = rstBookPOChild07
    FrmBookPOChild07.PartyCode = LaminatorCode
    If rstBookPOChild06.RecordCount > 0 Then FrmBookPOChild07.titleQty = Val(rstBookPOChild06.Fields("ActualQuantity").Value) Else FrmBookPOChild07.titleQty = 0
    On Error Resume Next
    Load FrmBookPOChild07
    If Err.Number <> 364 Then
        If rstBookPOParent.Fields("Laminator").Value <> "" Then
            If Left(BookPOType, 1) <> "O" Then
                If blnRecordExist And AllowTransactionsModification = 0 Then
                    If Not CheckEmpty(FrmBookPOChild07.Text8.Text, False) Then
                        Dim O As Object
                        For Each O In FrmBookPOChild07
                                If TypeName(O) = "TextBox" Then
                                    O.Locked = True
                                ElseIf TypeName(O) = "TDBNumber" Then
                                    O.ReadOnly = True
                                ElseIf TypeName(O) = "TDBDate" Then
                                    O.ReadOnly = True
                                End If
                        Next
                    End If
                End If
            End If
        End If
        FrmBookPOChild07.Show vbModal
        MhRealInput8_Validate False
    End If
    On Error GoTo 0
End Sub
Private Sub Command4_Click()
    If CheckEmpty(Text8.Text, False) Then Exit Sub
    If rstBookPOChild08.RecordCount = 0 Then Call AddRecord(rstBookPOChild08)
    Set FrmBookPOChild08.rstBookPOChild08 = rstBookPOChild08
    FrmBookPOChild08.PartyCode = BinderCode
    If rstBookPOChild08.RecordCount > 0 Then FrmBookPOChild08.OrderQty = Val(rstBookPOChild08.Fields("Quantity").Value) Else FrmBookPOChild08.OrderQty = 0
    On Error Resume Next
    Load FrmBookPOChild08
    If Err.Number <> 364 Then
        If rstBookPOParent.Fields("Laminator").Value <> "" Then
            If Left(BookPOType, 1) <> "O" Then
                If blnRecordExist And AllowTransactionsModification = 0 Then
                    If Not CheckEmpty(FrmBookPOChild08.Text8.Text, False) Then
                        Dim O As Object
                        For Each O In FrmBookPOChild08
                                If TypeName(O) = "TextBox" Then
                                    O.Locked = True
                                ElseIf TypeName(O) = "TDBNumber" Then
                                    O.ReadOnly = True
                                ElseIf TypeName(O) = "TDBDate" Then
                                    O.ReadOnly = True
                                End If
                        Next
                    End If
                End If
            End If
        End If
        FrmBookPOChild08.Show vbModal
        MhRealInput8_Validate False
    End If
    On Error GoTo 0
'    If CheckEmpty(Text8.Text, False) Then Exit Sub
'    If rstBookPOChild08.RecordCount = 0 Then Call AddRecord(rstBookPOChild08)
'    Set FrmBookPOChild08.rstBookPOChild08 = rstBookPOChild08
'    FrmBookPOChild08.BinderCode = BinderCode
'    If rstBookPOChild06.RecordCount > 0 Then FrmBookPOChild08.BookPrinterQuantity = Val(rstBookPOChild06.Fields("ActualQuantity").Value)
'    If rstBookPOChild05.RecordCount > 0 Then FrmBookPOChild08.BookPrinterQuantity = Val(rstBookPOChild05.Fields("ActualQuantity").Value)
'    FrmBookPOChild08.MhRealInput19.Text = Val(CheckNull(rstBookPOParent.Fields("DeliveredQuantityC").Value)) + Val(CheckNull(rstBookPOParent.Fields("DeliveredQuantityB").Value))
'    On Error Resume Next
'    Load FrmBookPOChild08
'    If Err.Number <> 364 Then
'        If rstBookPOParent.Fields("Binder").Value <> "" Then
'            If Left(BookPOType, 1) <> "O" Then
'                If blnRecordExist And AllowTransactionsModification = 0 Then
'                    If Not CheckEmpty(FrmBookPOChild08.Text8.Text, False) Then
'                        Dim O As Object
'                        For Each O In FrmBookPOChild08
'                                If TypeName(O) = "TextBox" Then
'                                    O.Locked = True
'                                ElseIf TypeName(O) = "TDBNumber" Then
'                                    O.ReadOnly = True
'                                ElseIf TypeName(O) = "TDBDate" Then
'                                    O.ReadOnly = True
'                                End If
'                        Next
'                    End If
'                End If
'            End If
'        End If
'        FrmBookPOChild08.Show vbModal
'        MhRealInput8_Validate False
'    End If
'    On Error GoTo 0
'    If AbortPO Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
End Sub
Public Sub PaperSlip(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String, Optional ByVal OrderType As String, Optional ByVal BookPOType As String)
    Dim oOutlookMsg As Outlook.MailItem, HeaderPrinted As Boolean, OrderNo As String, ItemName As String, TotalTax As Double, TotalAmount As Double, BillAmount As Double
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptPaperSlip.Text1.SetText IIf(Right(BookPOType, 1) = "S", "Paper-Requisition-Slip", "Item-Paper-Requisition-Slip")
        If OrderType = "BP" Then
        rptPaperSlip.Section4.Suppress = True
        rptPaperSlip.Section20.Suppress = True
        rptPaperSlip.Section11.Suppress = True
        rptPaperSlip.Section16.Suppress = True
    ElseIf OrderType = "TP" Then
        rptPaperSlip.Section13.Suppress = True
        rptPaperSlip.Section20.Suppress = True
        rptPaperSlip.Section11.Suppress = True
        rptPaperSlip.Section16.Suppress = True
    ElseIf OrderType = "CB" Then
        rptPaperSlip.Section4.Suppress = True
        rptPaperSlip.Section11.Suppress = True
        rptPaperSlip.Section13.Suppress = True
        rptPaperSlip.Section20.Suppress = True
    ElseIf OrderType = "TL" Then
        rptPaperSlip.Section13.Suppress = True
        rptPaperSlip.Section4.Suppress = True
        rptPaperSlip.Section11.Suppress = True
        rptPaperSlip.Section16.Suppress = True
    ElseIf OrderType = "BB" Then
        rptPaperSlip.Section13.Suppress = True
        rptPaperSlip.Section4.Suppress = True
        rptPaperSlip.Section20.Suppress = True
        rptPaperSlip.Section16.Suppress = True
    End If
    If rstBookPOChild0801.State = adStateOpen Then rstBookPOChild0801.Close
    If OrderType = "BP" Then
        rstBookPOChild0801.Open "SELECT BookPrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "TP" Then
        rstBookPOChild0801.Open "SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "CB" Then
        rstBookPOChild0801.Open "SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "TL" Then
        rstBookPOChild0801.Open "SELECT Laminator As Vendor FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "BB" Then
        rstBookPOChild0801.Open "SELECT Binder As Vendor FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild0801.Open "SELECT BookPrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT Laminator As Vendor FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT Binder As Vendor FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
    If rstBookPOChild0801.RecordCount > 0 Then rstBookPOChild0801.MoveFirst
    Do While Not rstBookPOChild0801.EOF
        TotalTax = 0: TotalAmount = 0
        HeaderPrinted = False
'
'        If OrderType = "TL" Or OrderType = "ALL" Then
'            If rstBookPOParent.State = adStateOpen Then rstBookPOParent.Close
'            rstBookPOParent.Open "SELECT E.Name As Element,O.Name As Operation,[Number],OS.Name As [Size],Quantity,M.Name As CalcMode,Rate,Amount,Adjustment,[GST%],GST,BillAmount,C.Remarks,TRIM(I.PrintName)+IIF(I.Price=0,'',' (Price : Rs. '+Format(I.Price,'0.00')+')') As Item,FS.PrintName As FinishSize,TRIM(TRIM(I.Pages)+'p/'+TRIM(I.Forms)+'f('+IIF(I.OneColorForms=0,'','1C-'+TRIM(I.OneColorForms)+' ')+IIF(I.TwoColorForms=0,'','2C-'+TRIM(I.TwoColorForms)+' ')+IIF(I.FourColorForms=0,'','4C-'+TRIM(I.FourColorForms)))+')' As Forme,TRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.Laminator) As Laminator,(SELECT eMail FROM AccountMaster WHERE Code=P.Laminator) As EMailId,I.Narration " & _
'                                                               "FROM ((((((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster FS ON I.FinishSize=FS.Code) INNER JOIN GeneralMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code) LEFT JOIN GeneralMaster OS ON C.[Size]=OS.Code WHERE P.Code='" & OrderCode & "' AND P.Laminator='" & rstBookPOChild0801.Fields("Vendor").Value & "' ORDER BY E.Name,O.Name", cnDatabase, adOpenKeyset, adLockOptimistic
'                rptPaperSlip.Text22.SetText rstBookPOChild08.Fields("BindingType").Value
'                rptPaperSlip.Text29.SetText Val(rstBookPOChild08.Fields("BindingForms").Value) + Val(rstBookPOChild08.Fields("ExtraForms").Value)
'            If rstBookPOParent.RecordCount = 0 Then
'                rptPaperSlip.Section20.Suppress = True
'            Else
'                With rstBookPOParent
'                    .MoveFirst
'                    BillAmount = 0
'                    Do While Not .EOF
'                        BillAmount = BillAmount + Val(.Fields("BillAmount").Value)
'                        TotalAmount = TotalAmount + Val(.Fields("BillAmount").Value)
'                        TotalTax = TotalTax + Val(.Fields("GST").Value)
'                        .MoveNext
'                    Loop
'                End With
'                rstBookPOParent.MoveFirst
'                If Not HeaderPrinted Then
'                    rptPaperSlip.Text7.SetText rstBookPOParent.Fields("OrderNo").Value
'                    rptPaperSlip.Text42.SetText Format(rstBookPOParent.Fields("OrderDate").Value, "dd-MM-yyyy")
'                    rptPaperSlip.Text12.SetText rstBookPOParent.Fields("Laminator").Value
'                    rptPaperSlip.Text13.SetText rstBookPOParent.Fields("Item").Value
'                    rptPaperSlip.Text14.SetText rstBookPOParent.Fields("FinishSize").Value
'                    rptPaperSlip.Text64.SetText rstBookPOParent.Fields("Forme").Value
'                    rptPaperSlip.Text15.SetText rstBookPOParent.Fields("FinalQuantity").Value
'                    EMailID = rstBookPOParent.Fields("EMailId").Value
'                    OrderNo = rstBookPOParent.Fields("OrderNo").Value
'                    ItemName = rstBookPOParent.Fields("Item").Value
'                    Attachment = Trim(rstBookPOParent.Fields("OrderNo").Value)
'                    HeaderPrinted = True
'                End If
'                rptPaperSlip.Subreport4.OpenSubreport.Text25.SetText "Amount Payable : " & Trim(NumberToWords(Round(BillAmount, 0), True))
'            End If
'        End If
     If OrderType = "BP" Or OrderType = "ALL" Then
            If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
     If DatabaseType = "MS SQL" Then
            rstBookPOChild05.Open "SELECT LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'pages/'+LTrim(M1.Forms)+'f ('+IIF(M1.OneColorForms=0,'','1-Col_'+LTrim(M1.OneColorForms)+'f +')+IIF(M1.TwoColorForms=0,'',' 2-Col_'+LTrim(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'','+ 4-Col_'+LTrim(M1.FourColorForms)+'f '))+' )' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,C.ActualQuantity,M1.DuplexPrinting,BillingQuantity01,BillingQuantity02,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size1) As [Size1],Pages1,[Forms1-¼],[Forms1-½],[Forms1-1],CHOOSE(CONVERT(NUMERIC,PlateType1),'Deep-etch','PS','Wipe-on','CTP') As Plate1,PrintRate1,PrintAmount1,PlateRate1,PlateAmount1," & _
                                  "(SELECT LTrim(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper1) As Paper1Name,[PaperWastage1%],PaperConsumptionOther1,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper1) As UOM1," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size2) As [Size2],Pages2,[Forms2-¼],[Forms2-½],[Forms2-1],CHOOSE(CONVERT(NUMERIC,PlateType2),'Deep-etch','PS','Wipe-on','CTP') As Plate2,PrintRate2,PrintAmount2,PlateRate2,PlateAmount2," & _
                                  "(SELECT LTrim(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper2) As Paper2Name,[PaperWastage2%],PaperConsumptionOther2,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper2) As UOM2," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size4) As [Size4],Pages4,[Forms4-¼],[Forms4-½],[Forms4-1],CHOOSE(CONVERT(NUMERIC,PlateType4),'Deep-etch','PS','Wipe-on','CTP') As Plate4,PrintRate4,PrintAmount4,PlateRate4,PlateAmount4," & _
                                  "(SELECT LTrim(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper4) As Paper4Name,[PaperWastage4%],PaperConsumptionOther4,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper4) As UOM4," & _
                                  "TotalPaperConsumption,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT LTrim(eMail) FROM AccountMaster WHERE CODE=P.BookPrinter) As EMailId,M1.Narration,P.BookPrinter,PlateMaker,PaperWastageMin1,PaperWastageMin2,PaperWastageMin4,PaperRate1,PaperRate2,PaperRate4,PaperAmount1,PaperAmount2,PaperAmount4,RBillAmount,RAdjustment,[RVAT%],RVAT,C.Processing,P.EstQty01 As FinalQuantity,C.Ref  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
            rstBookPOChild05.Open "SELECT Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,TRIM(TRIM(M1.Pages)+'pages/'+TRIM(M1.Forms)+'f ('+IIF(M1.OneColorForms=0,'','1-Col_'+TRIM(M1.OneColorForms)+'f +')+IIF(M1.TwoColorForms=0,'',' 2-Col_'+TRIM(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'','+ 4-Col_'+TRIM(M1.FourColorForms)+'f '))+' )' As Forme,TRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,C.ActualQuantity,M1.DuplexPrinting,BillingQuantity01,BillingQuantity02,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size1) As [Size1],Pages1,[Forms1-¼],[Forms1-½],[Forms1-1],CHOOSE(VAL(PlateType1),'Deep-etch','PS','Wipe-on','CTP') As Plate1,PrintRate1,PrintAmount1,PlateRate1,PlateAmount1," & _
                                  "(SELECT TRIM(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper1) As Paper1Name,[PaperWastage1%],PaperConsumptionOther1,(SELECT '('+TRIM(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper1) As UOM1," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size2) As [Size2],Pages2,[Forms2-¼],[Forms2-½],[Forms2-1],CHOOSE(VAL(PlateType2),'Deep-etch','PS','Wipe-on','CTP') As Plate2,PrintRate2,PrintAmount2,PlateRate2,PlateAmount2," & _
                                  "(SELECT TRIM(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper2) As Paper2Name,[PaperWastage2%],PaperConsumptionOther2,(SELECT '('+TRIM(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper2) As UOM2," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size4) As [Size4],Pages4,[Forms4-¼],[Forms4-½],[Forms4-1],CHOOSE(VAL(PlateType4),'Deep-etch','PS','Wipe-on','CTP') As Plate4,PrintRate4,PrintAmount4,PlateRate4,PlateAmount4," & _
                                  "(SELECT TRIM(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper4) As Paper4Name,[PaperWastage4%],PaperConsumptionOther4,(SELECT '('+TRIM(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper4) As UOM4," & _
                                  "TotalPaperConsumption,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT TRIM(eMail) FROM AccountMaster WHERE CODE=P.BookPrinter) As EMailId,M1.Narration,P.BookPrinter,PlateMaker,PaperWastageMin1,PaperWastageMin2,PaperWastageMin4,PaperRate1,PaperRate2,PaperRate4,PaperAmount1,PaperAmount2,PaperAmount4,RBillAmount,RAdjustment,[RVAT%],RVAT,C.Processing,P.EstQty01 As FinalQuantity,C.Ref  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
                rptPaperSlip.Text22.SetText rstBookPOChild08.Fields("BindingType").Value
                rptPaperSlip.Text29.SetText Val(rstBookPOChild08.Fields("BindingForms").Value) + Val(rstBookPOChild08.Fields("ExtraForms").Value)
            If rstBookPOChild05.RecordCount = 0 Then
                rptPaperSlip.Section13.Suppress = True
            Else
                rstBookPOChild05.MoveFirst
                TotalAmount = TotalAmount + IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("PBillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value)
                TotalTax = TotalTax + IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("VAT").Value + rstBookPOChild05.Fields("RVAT").Value, rstBookPOChild05.Fields("VAT").Value + rstBookPOChild05.Fields("PVAT").Value + rstBookPOChild05.Fields("RVAT").Value)
                rptPaperSlip.Text7.SetText rstBookPOChild05.Fields("OrderNo").Value
                rptPaperSlip.Text31.SetText rstBookPOChild05.Fields("Ref").Value
                rptPaperSlip.Text42.SetText Format(rstBookPOChild05.Fields("OrderDate").Value, "dd-MM-yyyy")
                rptPaperSlip.Text12.SetText rstBookPOChild05.Fields("TextPrinter").Value
                rptPaperSlip.Text13.SetText rstBookPOChild05.Fields("Item").Value
                rptPaperSlip.Text14.SetText rstBookPOChild05.Fields("FinishSize").Value
                rptPaperSlip.Text64.SetText rstBookPOChild05.Fields("Forme").Value
                rptPaperSlip.Text15.SetText rstBookPOChild05.Fields("FinalQuantity").Value
                rptPaperSlip.Subreport7.OpenSubreport.Text2.SetText " (" & Trim(NumberToWords(IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("PBillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value), True)) & ")"
                EMailID = rstBookPOChild05.Fields("EMailId").Value
                OrderNo = rstBookPOChild05.Fields("OrderNo").Value
                ItemName = rstBookPOChild05.Fields("Item").Value
                Attachment = Trim(rstBookPOChild05.Fields("OrderNo").Value)
                HeaderPrinted = True
            End If
        End If
        If OrderType = "TP" Or OrderType = "ALL" Then
    If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
                If DatabaseType = "MS SQL" Then
            rstBookPOChild06.Open "SELECT (SELECT PrintName FROM ElementMaster WHERE Code=C.Element) As Element,LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'pages/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1-Col-'+LTrim(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','f + 2-Col-'+LTrim(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','f + 4-Col-'+LTrim(M1.FourColorForms)+'f '))+'f)' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As CoverPrinter,C.ActualQuantity,C.BillingQuantity,C.FrontPrintingType,C.BackPrintingType,Imposition,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],(SELECT PrintName FROM GeneralMaster WHERE Code=C.PlateType) As Plate,PrintRate,PrintAmount,PlateRate,PlateAmount," & _
                                  "(SELECT LTrim(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper) As PaperName,[PaperWastage%],PaperConsumptionOther,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT LTrim(eMail) FROM AccountMaster WHERE Code=P.TitlePrinter) As EMailId,M1.Narration,P.TitlePrinter,PlateMaker,PaperWastageMin,PaperRate,PaperAmount,RBillAmount,RAdjustment,[RVAT%],RVAT,C.Processing,C.ProcessingBack,C.[Ups] As Ups,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOM,P.EstQty01 As FinalQuantity,C.Sets,C.TotalPlates,C.TotalPlatesBack,C.Ref,CHOOSE(CONVERT(NUMERIC,PlateTypeBack),'Deep-etch','PS','Wipe-on','CTP') As PlateBack,C.Pages  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.TitlePrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
            rstBookPOChild06.Open "SELECT (SELECT PrintName FROM ElementMaster WHERE Code=C.Element) As Element,Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,TRIM(TRIM(M1.Pages)+'pages/'+TRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1-Col_'+TRIM(M1.OneColorForms)+'f +')+IIF(M1.TwoColorForms=0,'','2-Col_'+TRIM(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'',' + 4-Col-'+TRIM(M1.FourColorForms)+'f '))+')' As Forme,TRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As CoverPrinter,C.ActualQuantity,C.BillingQuantity,C.FrontPrintingType,C.BackPrintingType,Imposition,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],CHOOSE(VAL(PlateType),'Deep-etch','PS','Wipe-on','CTP') As Plate,PrintRate,PrintAmount,PlateRate,PlateAmount,(SELECT TRIM(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper) As PaperName," & _
                                  "[PaperWastage%],PaperConsumptionOther,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT TRIM(eMail) FROM AccountMaster WHERE Code=P.TitlePrinter) As EMailId,M1.Narration,P.TitlePrinter,PlateMaker,PaperWastageMin,PaperRate,PaperAmount,RBillAmount,RAdjustment,[RVAT%],RVAT,C.Processing,C.ProcessingBack,C.[Ups] As Ups,(SELECT '('+TRIM(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOM,P.EstQty01 As FinalQuantity,C.Sets,C.TotalPlates,C.TotalPlatesBack,C.Ref,CHOOSE(VAL(PlateTypeBack),'Deep-etch','PS','Wipe-on','CTP') As PlateBack,C.Pages  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.TitlePrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If

                rptPaperSlip.Text22.SetText rstBookPOChild08.Fields("BindingType").Value
                rptPaperSlip.Text29.SetText Val(rstBookPOChild08.Fields("BindingForms").Value) + Val(rstBookPOChild08.Fields("ExtraForms").Value)
            If rstBookPOChild06.RecordCount = 0 Then
                rptPaperSlip.Section4.Suppress = True
            Else
                rstBookPOChild06.MoveFirst
                TotalAmount = TotalAmount + IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("PBillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value)
                TotalTax = TotalTax + IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, rstBookPOChild06.Fields("VAT").Value + rstBookPOChild06.Fields("RVAT").Value, rstBookPOChild06.Fields("VAT").Value + rstBookPOChild06.Fields("PVAT").Value + rstBookPOChild06.Fields("RVAT").Value)
                If Not HeaderPrinted Then
                    rptPaperSlip.Text7.SetText rstBookPOChild06.Fields("OrderNo").Value
                    rptPaperSlip.Text31.SetText rstBookPOChild06.Fields("Ref").Value
                    rptPaperSlip.Text42.SetText Format(rstBookPOChild06.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptPaperSlip.Text12.SetText rstBookPOChild06.Fields("CoverPrinter").Value
                    rptPaperSlip.Text13.SetText rstBookPOChild06.Fields("Item").Value
                    rptPaperSlip.Text14.SetText rstBookPOChild06.Fields("FinishSize").Value
                    rptPaperSlip.Text64.SetText rstBookPOChild06.Fields("Forme").Value
                    rptPaperSlip.Text15.SetText rstBookPOChild06.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild06.Fields("EMailId").Value
                    OrderNo = rstBookPOChild06.Fields("OrderNo").Value
                    ItemName = rstBookPOChild06.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild06.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                rptPaperSlip.Subreport2.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("PBillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value), True)) & ")"
            End If
        End If
        
        If OrderType = "CB" Or OrderType = "ALL" Then
                If rstBookPOChild10.State = adStateOpen Then rstBookPOChild10.Close
    If DatabaseType = "MS SQL" Then
        rstBookPOChild10.Open "SELECT LTrim(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinter,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMaker,CHOOSE(CONVERT(NUMERIC,C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalFormsFront,C1.TotalFormsBack,C1.PrintRateFront,C1.PrintRateBack,C1.[GST%],C1.GST,C1.Adjustment,C1.PrintAmount,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,C1.[PGST%],C1.PGST,PlateAmount," & _
                          "(SELECT LTrim(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C1.Paper) As PaperName,C1.[PaperWastage%],C1.[PaperConsumptionOther],C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,C2.ActualQuantity,[Ups/Plate],C2.PrintingQuantity,C2.BillingQuantity,LTrim(C2.FrontPrintingColor)+'+'+LTrim(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType," & _
                          "P.TitlePrinter,C1.PlateMaker,PaperWastageMin,PaperRate,PaperAmountBT,RAdjustment,[RGST%],RGST,PaperAmount,PrintAmountBT,PlateAmountBT,(SELECT EMail FROM AccountMaster WHERE Code=P.TitlePrinter) As EMail,Imposition,PlateType,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C1.Paper) As UOM,(SELECT LTrim(MAX(PrintingQuantity)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxPrintingQty, " & _
                          "P.EstQty01 As FinalQuantity,P.ProfitMargin,(SELECT LTrim(MAX(FrontPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxFrontColor,(SELECT LTrim(MAX(BackPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxBackColor,C1.Calculation " & _
                          "FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild10.Open "SELECT TRIM(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinter,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMaker,CHOOSE(VAL(C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalFormsFront,C1.TotalFormsBack,C1.PrintRateFront,C1.PrintRateBack,[C1.GST%],C1.GST,C1.Adjustment,C1.PrintAmount,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,[C1.PGST%],C1.PGST,PlateAmount," & _
                          "(SELECT TRIM(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C1.Paper) As PaperName,C1.[PaperWastage%],C1.[PaperConsumptionOther],C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,C2.ActualQuantity,[Ups/Plate],C2.PrintingQuantity,C2.BillingQuantity,TRIM(C2.FrontPrintingColor)+'+'+TRIM(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType,P.TitlePrinter,C1.PlateMaker,PaperWastageMin,PaperRate,PaperAmountBT,RAdjustment,[RGST%],RGST,PaperAmount,PrintAmountBT,PlateAmountBT,(SELECT EMail FROM AccountMaster WHERE Code=P.TitlePrinter) As EMail,Imposition,PlateType,(SELECT '('+TRIM(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C1.Paper) As UOM,(SELECT TRIM(MAX(PrintingQuantity)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxPrintingQty, " & _
                          "P.EstQty01 As FinalQuantity,P.ProfitMargin,(SELECT TRIM(MAX(FrontPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxFrontColor,(SELECT TRIM(MAX(BackPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxBackColor,C1.Calculation " & _
                          "FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
                    rptPaperSlip.Text22.SetText rstBookPOChild08.Fields("BindingType").Value
                    rptPaperSlip.Text29.SetText Val(rstBookPOChild08.Fields("BindingForms").Value) + Val(rstBookPOChild08.Fields("ExtraForms").Value)
                If rstBookPOChild10.RecordCount = 0 Then
                   rptPaperSlip.Section16.Suppress = True
            Else
                rstBookPOChild10.MoveFirst
                TotalTax = TotalTax + IIf(rstBookPOChild10.Fields("P.TitlePrinter").Value <> rstBookPOChild10.Fields("C1.PlateMaker").Value, rstBookPOChild10.Fields("GST").Value + rstBookPOChild10.Fields("RGST").Value, rstBookPOChild10.Fields("GST").Value + rstBookPOChild10.Fields("PGST").Value + rstBookPOChild10.Fields("RGST").Value)
                TotalAmount = TotalAmount + IIf(rstBookPOChild10.Fields("P.TitlePrinter").Value <> rstBookPOChild10.Fields("C1.PlateMaker").Value, rstBookPOChild10.Fields("PrintAmount").Value + rstBookPOChild10.Fields("PaperAmount").Value, rstBookPOChild10.Fields("PrintAmount").Value + rstBookPOChild10.Fields("PlateAmount").Value + rstBookPOChild10.Fields("PaperAmount").Value)
                If Not HeaderPrinted Then
                    rptPaperSlip.Text7.SetText rstBookPOChild10.Fields("OrderNo").Value
                    rptPaperSlip.Text42.SetText Format(rstBookPOChild10.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptPaperSlip.Text12.SetText rstBookPOChild10.Fields("TitlePrinter").Value
                    rptPaperSlip.Text13.SetText rstBookPOChild10.Fields("Item").Value
                    rptPaperSlip.Text59.SetText ""
'                    rptPaperSlip.Text64.SetText rstBookPOChild10.Fields("Forme").Value
                    rptPaperSlip.Text9.SetText ""
                    rptPaperSlip.Text15.SetText rstBookPOChild10.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild10.Fields("EMailId").Value
                    OrderNo = rstBookPOChild10.Fields("OrderNo").Value
                    ItemName = rstBookPOChild10.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild10.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                    
                    If rstBookPOChild10.Fields("Calculation").Value = "S" Then
                    rptPaperSlip.Subreport6_Text66.SetText "Combo Form Printing Details (As per Single Set Calculation)"
                    Else
                    rptPaperSlip.Subreport6_Text66.SetText "Combo Form Printing Details (As per Individual Set Calculation)"
                    End If
                    
                
                rptPaperSlip.Subreport6.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(IIf(rstBookPOChild10.Fields("TitlePrinter").Value <> rstBookPOChild10.Fields("PlateMaker").Value, rstBookPOChild10.Fields("BillAmount").Value + rstBookPOChild10.Fields("RBillAmount").Value, rstBookPOChild10.Fields("BillAmount").Value + rstBookPOChild10.Fields("PBillAmount").Value + rstBookPOChild10.Fields("RBillAmount").Value), True)) & ")"
            End If
        End If
        If OrderType = "TL" Or OrderType = "ALL" Then
            If rstBookPOChild07.State = adStateOpen Then rstBookPOChild07.Close
            rstBookPOChild07.Open "SELECT (SELECT PrintName FROM ElementMaster WHERE Code=C.Element) As Element,O.Name As Operation,[Number],OS.Name As [Size],Quantity,M.Name As CalcMode,Rate,Amount,Adjustment,[GST%],GST,BillAmount,C.Remarks,LTrim(I.PrintName)+IIF(I.Price=0,'',' (Price : Rs. '+Format(I.Price,'0.00')+')') As Item,FS.PrintName As FinishSize,LTrim(LTrim(I.Pages)+'p/'+LTrim(I.Forms)+'f('+IIF(I.OneColorForms=0,'','1C-'+LTrim(I.OneColorForms)+' ')+IIF(I.TwoColorForms=0,'','2C-'+LTrim(I.TwoColorForms)+' ')+IIF(I.FourColorForms=0,'','4C-'+LTrim(I.FourColorForms)))+')' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.Laminator) As Laminator,(SELECT eMail FROM AccountMaster WHERE Code=P.Laminator) As EMailId,I.Narration,P.EstQty01 As FinalQuantity " & _
                                                               "FROM ((((((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster FS ON I.FinishSize=FS.Code) Left JOIN GeneralMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code) LEFT JOIN GeneralMaster OS ON C.[Size]=OS.Code WHERE P.Code='" & OrderCode & "' AND P.Laminator='" & rstBookPOChild0801.Fields("Vendor").Value & "' ORDER BY E.Name,O.Name", cnDatabase, adOpenKeyset, adLockOptimistic
                rptPaperSlip.Text22.SetText rstBookPOChild08.Fields("BindingType").Value
                rptPaperSlip.Text29.SetText Val(rstBookPOChild08.Fields("BindingForms").Value) + Val(rstBookPOChild08.Fields("ExtraForms").Value)
            If rstBookPOChild07.RecordCount = 0 Then
                rptPaperSlip.Section20.Suppress = True
            Else
                With rstBookPOChild07
                    .MoveFirst
                    BillAmount = 0
                    Do While Not .EOF
                        BillAmount = BillAmount + Val(.Fields("BillAmount").Value)
                        TotalAmount = TotalAmount + Val(.Fields("BillAmount").Value)
                        TotalTax = TotalTax + Val(.Fields("GST").Value)
                        .MoveNext
                    Loop
                End With
                rstBookPOChild07.MoveFirst
                If Not HeaderPrinted Then
                    rptPaperSlip.Text7.SetText rstBookPOChild07.Fields("OrderNo").Value
                    rptPaperSlip.Text42.SetText Format(rstBookPOChild07.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptPaperSlip.Text12.SetText rstBookPOChild07.Fields("Laminator").Value
                    rptPaperSlip.Text13.SetText rstBookPOChild07.Fields("Item").Value
                    rptPaperSlip.Text14.SetText rstBookPOChild07.Fields("FinishSize").Value
                    rptPaperSlip.Text64.SetText rstBookPOChild07.Fields("Forme").Value
                    rptPaperSlip.Text15.SetText rstBookPOChild07.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild07.Fields("EMailId").Value
                    OrderNo = rstBookPOChild07.Fields("OrderNo").Value
                    ItemName = rstBookPOChild07.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild07.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                rptPaperSlip.Subreport4.OpenSubreport.Text25.SetText "Amount Payable : " & Trim(NumberToWords(Round(BillAmount, 0), True))
            End If
        End If
        If OrderType = "BB" Or OrderType = "ALL" Then
            If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
            rstBookPOChild08.Open "SELECT LTrim(M1.PrintName)+IIF(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'p/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+LTrim(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+LTrim(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+LTrim(M1.FourColorForms)))+')' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.Binder) As Binder,C.ActualQuantity, " & _
                                  "C.BillingQuantity,(SELECT PrintName FROM GeneralMaster WHERE Code=C.BindingType) As BindingType,BindingForms,ExtraForms,FormFoldRate,FormPasteRate,FormStitchRate,[Rate/Book],TotalPkts,PktPackRate,TotalBoxes,BoxPackRate,CartageRate,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,(SELECT eMail FROM AccountMaster WHERE Code=P.Binder) As EMailId,M1.Narration,P.EstQty01 As FinalQuantity " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.Binder='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
                rptPaperSlip.Text22.SetText rstBookPOChild08.Fields("BindingType").Value
                rptPaperSlip.Text29.SetText Val(rstBookPOChild08.Fields("BindingForms").Value) + Val(rstBookPOChild08.Fields("ExtraForms").Value)
            If rstBookPOChild08.RecordCount = 0 Then
                rptPaperSlip.Section11.Suppress = True
            Else
                rstBookPOChild08.MoveFirst
                TotalAmount = TotalAmount + Val(rstBookPOChild08.Fields("BillAmount").Value)
                TotalTax = TotalTax + Val(rstBookPOChild08.Fields("VAT").Value)
                If Not HeaderPrinted Then
                    rptPaperSlip.Text7.SetText rstBookPOChild08.Fields("OrderNo").Value
                    rptPaperSlip.Text42.SetText rstBookPOChild08.Fields("OrderDate").Value
                    rptPaperSlip.Text12.SetText rstBookPOChild08.Fields("Binder").Value
                    rptPaperSlip.Text13.SetText rstBookPOChild08.Fields("Item").Value
                    rptPaperSlip.Text14.SetText rstBookPOChild08.Fields("FinishSize").Value
                    rptPaperSlip.Text64.SetText rstBookPOChild08.Fields("Forme").Value
                    rptPaperSlip.Text15.SetText rstBookPOChild08.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild08.Fields("EMailId").Value
                    OrderNo = rstBookPOChild08.Fields("OrderNo").Value
                    ItemName = rstBookPOChild08.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild08.Fields("OrderNo").Value)
                End If
                rptPaperSlip.Subreport3.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(Round(rstBookPOChild08.Fields("BillAmount").Value, 0), True)) & ")"
            End If
        End If
            If rstBookPOChild09.State = adStateOpen Then rstBookPOChild09.Close
            rstBookPOChild09.Open "SELECT Choose(Val(Category),(SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item),(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='R' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item)) As ItemName,[Consumption/Item],OrderQuantity,TotalConsumption " & _
                              "FROM BookPOChild0801 T WHERE T.Code='" & OrderCode & "' AND T.Vendor='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
            If rstBookPOChild09.RecordCount = 0 Then rptPaperSlip.Section12.Suppress = True
            Screen.MousePointer = vbNormal
            If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,EMail,Website FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
            rptPaperSlip.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptPaperSlip.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
            rptPaperSlip.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
            rptPaperSlip.Text27.SetText "for " & rptPaperSlip.Text12.Text
            rptPaperSlip.Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
            If TotalAmount = 0 Then
            rptPaperSlip.Section14.Suppress = True
        Else
            rptPaperSlip.Text17.SetText Format(TotalTax, "##0.00")
            rptPaperSlip.Text18.SetText Format(TotalAmount, "##0.00")
            rptPaperSlip.Text19.SetText " (" & Trim(NumberToWords(Round(TotalAmount, 0), True)) & ")"
        End If
        rptPaperSlip.Subreport1.OpenSubreport.Database.SetDataSource rstBookPOChild05, 3, 1
        rptPaperSlip.Subreport2.OpenSubreport.Database.SetDataSource rstBookPOChild06, 3, 1
        rptPaperSlip.Subreport4.OpenSubreport.Database.SetDataSource rstBookPOChild07, 3, 1
        rptPaperSlip.Subreport3.OpenSubreport.Database.SetDataSource rstBookPOChild08, 3, 1
        rptPaperSlip.Subreport5.OpenSubreport.Database.SetDataSource rstBookPOChild09, 3, 1
        rptPaperSlip.Subreport6.OpenSubreport.Database.SetDataSource rstBookPOChild10, 3, 1
        rptPaperSlip.Subreport7.OpenSubreport.Database.SetDataSource rstBookPOChild10, 3, 1
        Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
        Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
        If OutputTo = "S" Then
            FrmReportViewer.EMailID = EMailID
            FrmReportViewer.Subject = "Book Order #" & Trim(OrderNo) + " Book : " + Trim(ItemName)
            FrmReportViewer.Attachment = Attachment
            FrmReportViewer.Message = Message
            Set FrmReportViewer.Report = rptPaperSlip
            FrmReportViewer.Show vbModal
        Else
            If rstBookPOList.State = adStateClosed Then
                If EMailID = "" Or OutputType = "P" Then
                    rptPaperSlip.PaperSource = crPRBinAuto
                    rptPaperSlip.PrintOut False   ' Print Report Without Prompt
                Else
                    rptPaperSlip.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                    rptPaperSlip.ExportOptions.DestinationType = crEDTDiskFile
                    rptPaperSlip.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                    rptPaperSlip.Export False
                    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                    With oOutlookMsg
                        .To = EMailID
                        .Subject = "Book Order #" & Trim(OrderNo) + " Book : " + Trim(ItemName)
                        .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                        .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                        .Importance = olImportanceHigh
                        .ReadReceiptRequested = True
                        .Send
                        If Err.Number = 0 Then cnDatabase.Execute "UPDATE BookPOParent SET BBODStatus=1 WHERE Code='" & OrderCode & "'"
                    End With
                    Set oOutlookMsg = Nothing
                End If
            Else
                rptPaperSlip.PaperSource = crPRBinAuto
                rptPaperSlip.PrintOut
            End If
        End If
        Set rptPaperSlip = Nothing
        rstBookPOChild0801.MoveNext
        If OrderType = "BP" Or OrderType = "TP" Or OrderType = "CB" Or OrderType = "TL" Or OrderType = "BB" Then Exit Do
    Loop
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild05): Call CloseRecordset(rstBookPOChild06): Call CloseRecordset(rstBookPOChild07): Call CloseRecordset(rstBookPOChild08): Call CloseRecordset(rstBookPOChild0801): Call CloseRecordset(rstBookPOChild09)
    On Error GoTo 0
    Screen.MousePointer = vbNormal
End Sub
Public Sub JobCard(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String, Optional ByVal OrderType As String, Optional ByVal BookPOType As String)
    Dim oOutlookMsg As Outlook.MailItem, HeaderPrinted As Boolean, OrderNo As String, ItemName As String, TotalTax As Double, TotalAmount As Double, BillAmount As Double
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptJobCard.Text1.SetText IIf(Right(BookPOType, 1) = "S", "Job-Card ", "Job-Order")
    If OrderType = "BP" Then
        rptJobCard.Section4.Suppress = True
        rptJobCard.Section20.Suppress = True
        rptJobCard.Section11.Suppress = True
        rptJobCard.Section16.Suppress = True
    ElseIf OrderType = "TP" Then
        rptJobCard.Section13.Suppress = True
        rptJobCard.Section20.Suppress = True
        rptJobCard.Section11.Suppress = True
        rptJobCard.Section16.Suppress = True
    ElseIf OrderType = "CB" Then
        rptJobCard.Section13.Suppress = True
        rptJobCard.Section20.Suppress = True
        rptJobCard.Section11.Suppress = True
        rptJobCard.Section4.Suppress = True
    ElseIf OrderType = "TL" Then
        rptJobCard.Section13.Suppress = True
        rptJobCard.Section4.Suppress = True
        rptJobCard.Section11.Suppress = True
        rptJobCard.Section16.Suppress = True
    ElseIf OrderType = "BB" Then
        rptJobCard.Section13.Suppress = True
        rptJobCard.Section4.Suppress = True
        rptJobCard.Section20.Suppress = True
        rptJobCard.Section16.Suppress = True
    End If
    If rstBookPOChild0801.State = adStateOpen Then rstBookPOChild0801.Close
    If OrderType = "BP" Then
        rstBookPOChild0801.Open "SELECT BookPrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "TP" Then
        rstBookPOChild0801.Open "SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "CB" Then
        rstBookPOChild0801.Open "SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "TL" Then
        rstBookPOChild0801.Open "SELECT Laminator As Vendor FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "BB" Then
        rstBookPOChild0801.Open "SELECT Binder As Vendor FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild0801.Open "SELECT BookPrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT Laminator As Vendor FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT Binder As Vendor FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
    If rstBookPOChild0801.RecordCount > 0 Then rstBookPOChild0801.MoveFirst
    Do While Not rstBookPOChild0801.EOF
        TotalTax = 0: TotalAmount = 0
        HeaderPrinted = False
        If OrderType = "BP" Or OrderType = "ALL" Then
            If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    If DatabaseType = "MS SQL" Then
            rstBookPOChild05.Open "SELECT LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'pages/'+LTrim(M1.Forms)+'f ('+IIF(M1.OneColorForms=0,'','1-Col_'+LTrim(M1.OneColorForms)+'f  +')+IIF(M1.TwoColorForms=0,'',' 2-Col_'+LTrim(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'','+ 4-Col_'+LTrim(M1.FourColorForms)+'f '))+' )' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,C.ActualQuantity,M1.DuplexPrinting,BillingQuantity01,BillingQuantity02,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size1) As [Size1],Pages1,[Forms1-¼],[Forms1-½],[Forms1-1],CHOOSE(CONVERT(NUMERIC,PlateType1),'Deep-etch','PS','Wipe-on','CTP') As Plate1,PrintRate1,PrintAmount1,PlateRate1,PlateAmount1," & _
                                  "(SELECT LTrim(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper1) As Paper1Name,[PaperWastage1%],PaperConsumptionOther1,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper1) As UOM1," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size2) As [Size2],Pages2,[Forms2-¼],[Forms2-½],[Forms2-1],CHOOSE(CONVERT(NUMERIC,PlateType2),'Deep-etch','PS','Wipe-on','CTP') As Plate2,PrintRate2,PrintAmount2,PlateRate2,PlateAmount2," & _
                                  "(SELECT LTrim(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper2) As Paper2Name,[PaperWastage2%],PaperConsumptionOther2,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper2) As UOM2," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size4) As [Size4],Pages4,[Forms4-¼],[Forms4-½],[Forms4-1],CHOOSE(CONVERT(NUMERIC,PlateType4),'Deep-etch','PS','Wipe-on','CTP') As Plate4,PrintRate4,PrintAmount4,PlateRate4,PlateAmount4," & _
                                  "(SELECT LTrim(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper4) As Paper4Name,[PaperWastage4%],PaperConsumptionOther4,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper4) As UOM4," & _
                                  "TotalPaperConsumption,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT LTrim(eMail) FROM AccountMaster WHERE CODE=P.BookPrinter) As EMailId,M1.Narration,P.BookPrinter,PlateMaker,PaperWastageMin1,PaperWastageMin2,PaperWastageMin4,PaperRate1,PaperRate2,PaperRate4,PaperAmount1,PaperAmount2,PaperAmount4,RBillAmount,RAdjustment,[RVAT%],RVAT,C.Processing,P.EstQty01 As FinalQuantity,C.Ref  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
            rstBookPOChild05.Open "SELECT LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'pages/'+LTrim(M1.Forms)+'f ('+IIF(M1.OneColorForms=0,'','1-Col_'+LTrim(M1.OneColorForms)+'f  +')+IIF(M1.TwoColorForms=0,'',' 2-Col_'+LTrim(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'','+ 4-Col_'+LTrim(M1.FourColorForms)+'f '))+' )' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,C.ActualQuantity,M1.DuplexPrinting,BillingQuantity01,BillingQuantity02,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size1) As [Size1],Pages1,[Forms1-¼],[Forms1-½],[Forms1-1],CHOOSE(VAL(PlateType1),'Deep-etch','PS','Wipe-on','CTP') As Plate1,PrintRate1,PrintAmount1,PlateRate1,PlateAmount1," & _
                                  "(SELECT LTrim(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper1) As Paper1Name,[PaperWastage1%],PaperConsumptionOther1,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper1) As UOM1," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size2) As [Size2],Pages2,[Forms2-¼],[Forms2-½],[Forms2-1],CHOOSE(VAL(PlateType2),'Deep-etch','PS','Wipe-on','CTP') As Plate2,PrintRate2,PrintAmount2,PlateRate2,PlateAmount2," & _
                                  "(SELECT LTrim(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper2) As Paper2Name,[PaperWastage2%],PaperConsumptionOther2,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper2) As UOM2," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size4) As [Size4],Pages4,[Forms4-¼],[Forms4-½],[Forms4-1],CHOOSE(VAL(PlateType4),'Deep-etch','PS','Wipe-on','CTP') As Plate4,PrintRate4,PrintAmount4,PlateRate4,PlateAmount4," & _
                                  "(SELECT LTrim(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper4) As Paper4Name,[PaperWastage4%],PaperConsumptionOther4,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper4) As UOM4," & _
                                  "TotalPaperConsumption,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT LTrim(eMail) FROM AccountMaster WHERE CODE=P.BookPrinter) As EMailId,M1.Narration,P.BookPrinter,PlateMaker,PaperWastageMin1,PaperWastageMin2,PaperWastageMin4,PaperRate1,PaperRate2,PaperRate4,PaperAmount1,PaperAmount2,PaperAmount4,RBillAmount,RAdjustment,[RVAT%],RVAT,C.Processing,P.EstQty01 As FinalQuantity,C.Ref  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
            
            If rstBookPOChild05.RecordCount = 0 Then
                rptJobCard.Section13.Suppress = True
            Else
                rstBookPOChild05.MoveFirst
                TotalAmount = TotalAmount + IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("PBillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value)
                TotalTax = TotalTax + IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("VAT").Value + rstBookPOChild05.Fields("RVAT").Value, rstBookPOChild05.Fields("VAT").Value + rstBookPOChild05.Fields("PVAT").Value + rstBookPOChild05.Fields("RVAT").Value)
                rptJobCard.Text7.SetText rstBookPOChild05.Fields("OrderNo").Value
                rptJobCard.Text22.SetText rstBookPOChild05.Fields("Ref").Value
                rptJobCard.Text42.SetText Format(rstBookPOChild05.Fields("OrderDate").Value, "dd-MM-yyyy")
                rptJobCard.Text12.SetText rstBookPOChild05.Fields("TextPrinter").Value
                rptJobCard.Text13.SetText rstBookPOChild05.Fields("Item").Value
                rptJobCard.Text14.SetText rstBookPOChild05.Fields("FinishSize").Value
                rptJobCard.Text64.SetText rstBookPOChild05.Fields("Forme").Value
                rptJobCard.Text15.SetText rstBookPOChild05.Fields("FinalQuantity").Value       'FinalQuantity ("ActualQuantity").Value
                rptJobCard.Subreport1.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("PBillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value), True)) & ")"
                EMailID = rstBookPOChild05.Fields("EMailId").Value
                OrderNo = rstBookPOChild05.Fields("OrderNo").Value
                ItemName = rstBookPOChild05.Fields("Item").Value
                Attachment = Trim(rstBookPOChild05.Fields("OrderNo").Value)
                HeaderPrinted = True
            End If
        End If
        If OrderType = "TP" Or OrderType = "ALL" Then
            If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
    If DatabaseType = "MS SQL" Then
            rstBookPOChild06.Open "SELECT (SELECT PrintName FROM ElementMaster WHERE Code=C.Element) As Element,LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'pages/'+LTrim(M1.Forms)+'f ('+IIF(M1.OneColorForms=0,'','1-Col_'+LTrim(M1.OneColorForms)+'f  +')+IIF(M1.TwoColorForms=0,'',' 2-Col_'+LTrim(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'','+ 4-Col_'+LTrim(M1.FourColorForms)+'f '))+' )' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As CoverPrinter,C.ActualQuantity,C.BillingQuantity,C.FrontPrintingType,C.BackPrintingType,Imposition,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],(SELECT PrintName FROM GeneralMaster WHERE Code=C.PlateType) As Plate,PrintRate,PrintAmount,PlateRate,PlateAmount," & _
                                  "(SELECT LTrim(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper) As PaperName,[PaperWastage%],PaperConsumptionOther,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT LTrim(eMail) FROM AccountMaster WHERE Code=P.TitlePrinter) As EMailId,M1.Narration,P.TitlePrinter,PlateMaker,PaperWastageMin,PaperRate,PaperAmount,RBillAmount,RAdjustment,[RVAT%],RVAT,C.Processing,C.ProcessingBack,C.[Ups] As Ups,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOM,P.EstQty01 As FinalQuantity,C.Sets,C.TotalPlates,C.TotalPlatesBack,C.Ref,CHOOSE(CONVERT(NUMERIC,PlateTypeBack),'Deep-etch','PS','Wipe-on','CTP') As PlateBack,C.Pages  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.TitlePrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
            rstBookPOChild06.Open "SELECT (SELECT PrintName FROM ElementMaster WHERE Code=C.Element) As Element,LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'pages/'+LTrim(M1.Forms)+'f ('+IIF(M1.OneColorForms=0,'','1-Col_'+LTrim(M1.OneColorForms)+'f  +')+IIF(M1.TwoColorForms=0,'',' 2-Col_'+LTrim(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'','+ 4-Col_'+LTrim(M1.FourColorForms)+'f '))+' )' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As CoverPrinter,C.ActualQuantity,C.BillingQuantity,C.FrontPrintingType,C.BackPrintingType,Imposition,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],CHOOSE(VAL(PlateType),'Deep-etch','PS','Wipe-on','CTP') As Plate,PrintRate,PrintAmount,PlateRate,PlateAmount,(SELECT LTrim(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper) As PaperName," & _
                                  "[PaperWastage%],PaperConsumptionOther,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT LTrim(eMail) FROM AccountMaster WHERE Code=P.TitlePrinter) As EMailId,M1.Narration,P.TitlePrinter,PlateMaker,PaperWastageMin,PaperRate,PaperAmount,RBillAmount,RAdjustment,[RVAT%],RVAT,C.Processing,C.ProcessingBack,C.[Ups] As Ups,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOM,P.EstQty01 As FinalQuantity,C.Sets,C.TotalPlates,C.TotalPlatesBack,C.Ref,CHOOSE(VAL(PlateTypeBack),'Deep-etch','PS','Wipe-on','CTP') As PlateBack,C.Pages  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.TitlePrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
            If rstBookPOChild06.RecordCount = 0 Then
                rptJobCard.Section4.Suppress = True
            Else
                rstBookPOChild06.MoveFirst
                TotalAmount = TotalAmount + IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("PBillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value)
                TotalTax = TotalTax + IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, rstBookPOChild06.Fields("VAT").Value + rstBookPOChild06.Fields("RVAT").Value, rstBookPOChild06.Fields("VAT").Value + rstBookPOChild06.Fields("PVAT").Value + rstBookPOChild06.Fields("RVAT").Value)
                If Not HeaderPrinted Then
                    rptJobCard.Text7.SetText rstBookPOChild06.Fields("OrderNo").Value
                    rptJobCard.Text22.SetText rstBookPOChild06.Fields("Ref").Value
                    rptJobCard.Text42.SetText Format(rstBookPOChild06.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptJobCard.Text12.SetText rstBookPOChild06.Fields("CoverPrinter").Value
                    rptJobCard.Text13.SetText rstBookPOChild06.Fields("Item").Value
                    rptJobCard.Text14.SetText rstBookPOChild06.Fields("FinishSize").Value
                    rptJobCard.Text64.SetText rstBookPOChild06.Fields("Forme").Value
                    rptJobCard.Text15.SetText rstBookPOChild06.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild06.Fields("EMailId").Value
                    OrderNo = rstBookPOChild06.Fields("OrderNo").Value
                    ItemName = rstBookPOChild06.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild06.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                rptJobCard.Subreport2.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("PBillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value), True)) & ")"
            End If
        End If
        If OrderType = "CB" Or OrderType = "ALL" Then
                If rstBookPOChild10.State = adStateOpen Then rstBookPOChild10.Close
    If DatabaseType = "MS SQL" Then
        rstBookPOChild10.Open "SELECT LTrim(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinter,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMaker,CHOOSE(CONVERT(NUMERIC,C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalFormsFront,C1.TotalFormsBack,C1.PrintRateFront,C1.PrintRateBack,C1.[GST%],C1.GST,C1.Adjustment,C1.PrintAmount,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,C1.[PGST%],C1.PGST,PlateAmount," & _
                          "(SELECT LTRIM(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C1.Paper) As PaperName,C1.[PaperWastage%],C1.[PaperConsumptionOther],C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,C2.ActualQuantity,[Ups/Plate],C2.PrintingQuantity,C2.BillingQuantity,LTRIM(C2.FrontPrintingColor)+'+'+LTRIM(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType,P.TitlePrinter,C1.PlateMaker,PaperWastageMin,PaperRate,PaperAmountBT,RAdjustment,[RGST%],RGST,PaperAmount,PrintAmountBT,PlateAmountBT,(SELECT EMail FROM AccountMaster WHERE Code=P.TitlePrinter) As EMail,Imposition,PlateType, " & _
                        "(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C1.Paper) As UOM,(SELECT LTrim(MAX(PrintingQuantity)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxPrintingQty, " & _
                          "P.EstQty01 As FinalQuantity,P.ProfitMargin,(SELECT LTrim(MAX(FrontPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxFrontColor,(SELECT LTrim(MAX(BackPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxBackColor,C1.Calculation FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild10.Open "SELECT LTrim(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinter,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMaker,CHOOSE(VAL(C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalFormsFront,C1.TotalFormsBack,C1.PrintRateFront,C1.PrintRateBack,[C1.GST%],C1.GST,C1.Adjustment,C1.PrintAmount,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,[C1.PGST%],C1.PGST,PlateAmount," & _
                          "(SELECT TRIM(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C1.Paper) As PaperName,C1.[PaperWastage%],C1.[PaperConsumptionOther],C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,C2.ActualQuantity,[Ups/Plate],C2.PrintingQuantity,C2.BillingQuantity,TRIM(C2.FrontPrintingColor)+'+'+TRIM(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType,P.TitlePrinter,C1.PlateMaker,PaperWastageMin,PaperRate,PaperAmountBT,RAdjustment,[RGST%],RGST,PaperAmount,PrintAmountBT,PlateAmountBT,(SELECT EMail FROM AccountMaster WHERE Code=P.TitlePrinter) As EMail,Imposition,PlateType, " & _
                        "(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C1.Paper) As UOM,(SELECT LTrim(MAX(PrintingQuantity)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxPrintingQty, " & _
                          "P.EstQty01 As FinalQuantity,P.ProfitMargin,(SELECT LTrim(MAX(FrontPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxFrontColor,(SELECT LTrim(MAX(BackPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxBackColor,C1.Calculation FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
                If rstBookPOChild10.RecordCount = 0 Then
                   rptJobCard.Section16.Suppress = True
            Else
                rstBookPOChild10.MoveFirst
                TotalTax = TotalTax + IIf(rstBookPOChild10.Fields("P.TitlePrinter").Value <> rstBookPOChild10.Fields("C1.PlateMaker").Value, rstBookPOChild10.Fields("GST").Value + rstBookPOChild10.Fields("RGST").Value, rstBookPOChild10.Fields("GST").Value + rstBookPOChild10.Fields("PGST").Value + rstBookPOChild10.Fields("RGST").Value)
                TotalAmount = TotalAmount + IIf(rstBookPOChild10.Fields("P.TitlePrinter").Value <> rstBookPOChild10.Fields("C1.PlateMaker").Value, rstBookPOChild10.Fields("PrintAmount").Value + rstBookPOChild10.Fields("PaperAmount").Value, rstBookPOChild10.Fields("PrintAmount").Value + rstBookPOChild10.Fields("PlateAmount").Value + rstBookPOChild10.Fields("PaperAmount").Value)
                If Not HeaderPrinted Then
                    rptJobCard.Text7.SetText rstBookPOChild10.Fields("OrderNo").Value
                    rptJobCard.Text42.SetText Format(rstBookPOChild10.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptJobCard.Text12.SetText rstBookPOChild10.Fields("TitlePrinter").Value
                    rptJobCard.Text13.SetText rstBookPOChild10.Fields("Item").Value
                    rptJobCard.Text14.SetText rstBookPOChild10.Fields("FinishSize").Value
                    rptJobCard.Text59.SetText ""
                    rptJobCard.Text64.SetText rstBookPOChild10.Fields("Forme").Value
                    rptJobCard.Text9.SetText ""
                    rptJobCard.Text15.SetText rstBookPOChild10.Fields("FinalQuantity").Value     'IIf(OrderType = "CB", (Val(rstBookPOChild10.Fields("MaxPrintingQty").Value)), Val(rstBookPOChild10.Fields("FinalQuantity").Value))
                    EMailID = rstBookPOChild10.Fields("EMailId").Value
                    OrderNo = rstBookPOChild10.Fields("OrderNo").Value
                    ItemName = rstBookPOChild10.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild10.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                
                    If rstBookPOChild10.Fields("Calculation").Value = "S" Then
                    rptJobCard.Subreport6_Text66.SetText "Combo Format Printing Details (As per Single Set Calculation)"
                    Else
                    rptJobCard.Subreport6_Text66.SetText "Combo Format Printing Details (As per Individual Set Calculation)"
                    End If
                
                rptJobCard.Subreport6.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(IIf(rstBookPOChild10.Fields("TitlePrinter").Value <> rstBookPOChild10.Fields("PlateMaker").Value, rstBookPOChild10.Fields("BillAmount").Value + rstBookPOChild10.Fields("RBillAmount").Value, rstBookPOChild10.Fields("BillAmount").Value + rstBookPOChild10.Fields("PBillAmount").Value + rstBookPOChild10.Fields("RBillAmount").Value), True)) & ")"
            End If
        End If
        If OrderType = "TL" Or OrderType = "ALL" Then
            If rstBookPOChild07.State = adStateOpen Then rstBookPOChild07.Close
            rstBookPOChild07.Open "SELECT (SELECT PrintName FROM ElementMaster WHERE Code=C.Element) As Element,O.Name As Operation,[Number],OS.Name As [Size],Quantity,M.Name As CalcMode,Rate,Amount,Adjustment,[GST%],GST,BillAmount,C.Remarks,LTrim(I.PrintName)+IIF(I.Price=0,'',' (Price : Rs. '+Format(I.Price,'0.00')+')') As Item,FS.PrintName As FinishSize,LTrim(LTrim(I.Pages)+'p/'+LTrim(I.Forms)+'f('+IIF(I.OneColorForms=0,'','1C-'+LTrim(I.OneColorForms)+'f ')+IIF(I.TwoColorForms=0,'','2Col_'+LTrim(I.TwoColorForms)+'f ')+IIF(I.FourColorForms=0,'','4C-'+LTrim(I.FourColorForms)+'f '))+')' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.Laminator) As Laminator,(SELECT eMail FROM AccountMaster WHERE Code=P.Laminator) As EMailId,I.Narration,P.EstQty01 As FinalQuantity " & _
                                                               "FROM ((((((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) Left JOIN BookMaster I ON P.Book=I.Code) Left JOIN GeneralMaster FS ON I.FinishSize=FS.Code) Left JOIN GeneralMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code) LEFT JOIN GeneralMaster OS ON C.[Size]=OS.Code WHERE P.Code='" & OrderCode & "' AND P.Laminator='" & rstBookPOChild0801.Fields("Vendor").Value & "' ORDER BY E.Name,O.Name", cnDatabase, adOpenKeyset, adLockOptimistic
            If rstBookPOChild07.RecordCount = 0 Then
                rptJobCard.Section20.Suppress = True
            Else
                With rstBookPOChild07
                    .MoveFirst
                    BillAmount = 0
                    Do While Not .EOF
                        BillAmount = BillAmount + Val(.Fields("BillAmount").Value)
                        TotalAmount = TotalAmount + Val(.Fields("BillAmount").Value)
                        TotalTax = TotalTax + Val(.Fields("GST").Value)
                        .MoveNext
                    Loop
                End With
                rstBookPOChild07.MoveFirst
                If Not HeaderPrinted Then
                    rptJobCard.Text7.SetText rstBookPOChild07.Fields("OrderNo").Value
                    rptJobCard.Text42.SetText Format(rstBookPOChild07.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptJobCard.Text12.SetText rstBookPOChild07.Fields("Laminator").Value
                    rptJobCard.Text13.SetText rstBookPOChild07.Fields("Item").Value
                    rptJobCard.Text14.SetText rstBookPOChild07.Fields("FinishSize").Value
                    rptJobCard.Text64.SetText rstBookPOChild07.Fields("Forme").Value
                    rptJobCard.Text15.SetText rstBookPOChild07.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild07.Fields("EMailId").Value
                    OrderNo = rstBookPOChild07.Fields("OrderNo").Value
                    ItemName = rstBookPOChild07.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild07.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                rptJobCard.Subreport4.OpenSubreport.Text25.SetText "Amount Payable : " & Trim(NumberToWords(Round(BillAmount, 0), True))
            End If
        End If
        If OrderType = "BB" Or OrderType = "ALL" Then
            If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
            rstBookPOChild08.Open "SELECT LTrim(M1.PrintName)+IIF(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'p/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+LTrim(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2Col_'+LTrim(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+LTrim(M1.FourColorForms)))+')' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.Binder) As Binder,C.ActualQuantity, " & _
                                  "C.BillingQuantity,(SELECT PrintName FROM GeneralMaster WHERE Code=C.BindingType) As BindingType,BindingForms,ExtraForms,FormFoldRate,FormPasteRate,FormStitchRate,[Rate/Book],TotalPkts,PktPackRate,TotalBoxes,BoxPackRate,CartageRate,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,(SELECT eMail FROM AccountMaster WHERE Code=P.Binder) As EMailId,M1.Narration,P.EstQty01 As FinalQuantity " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.Binder='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
            If rstBookPOChild08.RecordCount = 0 Then
                rptJobCard.Section11.Suppress = True
            Else
                rstBookPOChild08.MoveFirst
                TotalAmount = TotalAmount + Val(rstBookPOChild08.Fields("BillAmount").Value)
                TotalTax = TotalTax + Val(rstBookPOChild08.Fields("VAT").Value)
                If Not HeaderPrinted Then
                    rptJobCard.Text7.SetText rstBookPOChild08.Fields("OrderNo").Value
                    rptJobCard.Text42.SetText rstBookPOChild08.Fields("OrderDate").Value
                    rptJobCard.Text12.SetText rstBookPOChild08.Fields("Binder").Value
                    rptJobCard.Text13.SetText rstBookPOChild08.Fields("Item").Value
                    rptJobCard.Text14.SetText rstBookPOChild08.Fields("FinishSize").Value
                    rptJobCard.Text64.SetText rstBookPOChild08.Fields("Forme").Value
                    rptJobCard.Text15.SetText rstBookPOChild08.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild08.Fields("EMailId").Value
                    OrderNo = rstBookPOChild08.Fields("OrderNo").Value
                    ItemName = rstBookPOChild08.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild08.Fields("OrderNo").Value)
                End If
                rptJobCard.Subreport3.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(Round(rstBookPOChild08.Fields("BillAmount").Value, 0), True)) & ")"
            End If
        End If
        If rstBookPOChild09.State = adStateOpen Then rstBookPOChild09.Close
    If DatabaseType = "MS SQL" Then
        rstBookPOChild09.Open "SELECT Choose(CONVERT(NUMERIC,Category),(SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item),(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='R' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item)) As ItemName,[Consumption/Item],OrderQuantity,TotalConsumption " & _
                              "FROM BookPOChild0801 T WHERE T.Code='" & OrderCode & "' AND T.Vendor='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild09.Open "SELECT Choose(Val(Category),(SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item),(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='R' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item)) As ItemName,[Consumption/Item],OrderQuantity,TotalConsumption " & _
                              "FROM BookPOChild0801 T WHERE T.Code='" & OrderCode & "' AND T.Vendor='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
        
        If rstBookPOChild09.RecordCount = 0 Then rptJobCard.Section12.Suppress = True
        Screen.MousePointer = vbNormal
        If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,EMail,Website FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
        rptJobCard.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptJobCard.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
        rptJobCard.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
        rptJobCard.Text27.SetText "for " & rptJobCard.Text12.Text
        rptJobCard.Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
        If TotalAmount = 0 Then
            rptJobCard.Section14.Suppress = True
        Else
            rptJobCard.Text17.SetText Format(TotalTax, "##0.00")
            rptJobCard.Text18.SetText Format(TotalAmount, "##0.00")
            rptJobCard.Text19.SetText " (" & Trim(NumberToWords(Round(TotalAmount, 0), True)) & ")"
        End If
        rptJobCard.Subreport1.OpenSubreport.Database.SetDataSource rstBookPOChild05, 3, 1
        rptJobCard.Subreport2.OpenSubreport.Database.SetDataSource rstBookPOChild06, 3, 1
        rptJobCard.Subreport4.OpenSubreport.Database.SetDataSource rstBookPOChild07, 3, 1
        rptJobCard.Subreport3.OpenSubreport.Database.SetDataSource rstBookPOChild08, 3, 1
        rptJobCard.Subreport5.OpenSubreport.Database.SetDataSource rstBookPOChild09, 3, 1
        rptJobCard.Subreport6.OpenSubreport.Database.SetDataSource rstBookPOChild10, 3, 1
        Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
        Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
        If OutputTo = "S" Then
            FrmReportViewer.EMailID = EMailID
            FrmReportViewer.Subject = "Book Order #" & Trim(OrderNo) + " Book : " + Trim(ItemName)
            FrmReportViewer.Attachment = Attachment
            FrmReportViewer.Message = Message
            Set FrmReportViewer.Report = rptJobCard
            FrmReportViewer.Show vbModal
        Else
            If rstBookPOList.State = adStateClosed Then
                If EMailID = "" Or OutputType = "P" Then
                    rptJobCard.PaperSource = crPRBinAuto
                    rptJobCard.PrintOut False   ' Print Report Without Prompt
                Else
                    rptJobCard.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                    rptJobCard.ExportOptions.DestinationType = crEDTDiskFile
                    rptJobCard.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                    rptJobCard.Export False
                    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                    With oOutlookMsg
                        .To = EMailID
                        .Subject = "Book Order #" & Trim(OrderNo) + " Book : " + Trim(ItemName)
                        .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                        .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                        .Importance = olImportanceHigh
                        .ReadReceiptRequested = True
                        .Send
                        If Err.Number = 0 Then cnDatabase.Execute "UPDATE BookPOParent SET BBODStatus=1 WHERE Code='" & OrderCode & "'"
                    End With
                    Set oOutlookMsg = Nothing
                End If
            Else
                rptJobCard.PaperSource = crPRBinAuto
                rptJobCard.PrintOut
            End If
        End If
        Set rptJobCard = Nothing
        rstBookPOChild0801.MoveNext
        If OrderType = "BP" Or OrderType = "TP" Or OrderType = "CB" Or OrderType = "TL" Or OrderType = "BB" Then Exit Do
    Loop
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild05): Call CloseRecordset(rstBookPOChild06): Call CloseRecordset(rstBookPOChild07): Call CloseRecordset(rstBookPOChild08): Call CloseRecordset(rstBookPOChild0801): Call CloseRecordset(rstBookPOChild09)
    On Error GoTo 0
    Screen.MousePointer = vbNormal
End Sub
Public Sub PrintBookPrintOrder01(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String, Optional ByVal OrderType As String, Optional ByVal BookPOType As String) 'JUC/UC
    Dim oOutlookMsg As Outlook.MailItem, HeaderPrinted As Boolean, Rate As Double, Amount As Double, FinalRate As Double, FinalAmount As Double, GST As Double, Qty As Long, OrderNo As String, ItemName As String, TmpAmount As Long
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstBookPOChild0801.State = adStateOpen Then rstBookPOChild0801.Close
    rstBookPOChild0801.Open "SELECT BookPrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT Laminator As Vendor FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT Binder As Vendor FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    If rstBookPOChild0801.RecordCount > 0 Then rstBookPOChild0801.MoveFirst
    Do While Not rstBookPOChild0801.EOF
        HeaderPrinted = False: Qty = 0: Rate = 0: Amount = 0: FinalRate = 0: FinalAmount = 0: GST = 0: TmpAmount = 0
        If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    If DatabaseType = "MS SQL" Then
        rstBookPOChild05.Open "SELECT LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,C.ActualQuantity,DuplexPrinting,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,'1' As Color,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size1) As [Size],Pages1 As Pages,LTrim([TotalForms1-¼])+'(¼)+'+LTrim([TotalForms1-½])+'(½)+'+LTrim([TotalForms1-1])+'='+LTrim([TotalForms1-¼]+[TotalForms1-½]+[TotalForms1-1])+'f' As Forms,CHOOSE(CONVERT(NUMERIC,PlateType1),'Deep-etch','PS','Wipe-on','CTP') As Plate,(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper1) As PaperName,[PaperWastage1%] As PaperWastage,PaperConsumptionOther1 As PaperConsumption,PrintAmount1+PaperAmount1 As GrossAmount,PlateAmount1 As PlateAmount,C.Remarks," & _
                              "Adjustment+RAdjustment As Adjustment,PAdjustment,VAT+RVAT As GST,PVAT,LTrim(P.Name) As OrderNo,P.Date,LTrim(LTrim(M1.Pages)+'p/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+LTrim(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+LTrim(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+LTrim(M1.FourColorForms)))+')' As Forme,(SELECT eMail FROM AccountMaster WHERE Code=P.BookPrinter) As EMailId,P.BookPrinter,PlateMaker,P.EstQty01 As FinalQuantity FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "' UNION " & _
                              "SELECT LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,C.ActualQuantity,DuplexPrinting,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,'2' As Color,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size2) As [Size],Pages2 As Pages,LTrim([TotalForms2-¼])+'(¼)+'+LTrim([TotalForms2-½])+'(½)+'+LTrim([TotalForms2-1])+'='+LTrim([TotalForms2-¼]+[TotalForms2-½]+[TotalForms2-1])+'f' As Forms,CHOOSE(CONVERT(NUMERIC,PlateType2),'Deep-etch','PS','Wipe-on','CTP') As Plate,(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper2) As PaperName,[PaperWastage2%] As PaperWastage,PaperConsumptionOther2 As PaperConsumption,PrintAmount2+PaperAmount2 As GrossAmount,PlateAmount2 As PlateAmount,C.Remarks," & _
                              "Adjustment+RAdjustment As Adjustment,PAdjustment,VAT+RVAT As GST,PVAT,LTrim(P.Name) As OrderNo,P.Date,LTrim(LTrim(M1.Pages)+'p/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+LTrim(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+LTrim(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+LTrim(M1.FourColorForms)))+')' As Forme,(SELECT eMail FROM AccountMaster WHERE Code=P.BookPrinter) As EMailId,P.BookPrinter,PlateMaker,P.EstQty01 As FinalQuantity FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "' UNION " & _
                              "SELECT LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,C.ActualQuantity,DuplexPrinting,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,'4' As Color,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size4) As [Size],Pages4 As Pages,LTrim([TotalForms4-¼])+'(¼)+'+LTrim([TotalForms4-½])+'(½)+'+LTrim([TotalForms4-1])+'='+LTrim([TotalForms4-¼]+[TotalForms4-½]+[TotalForms4-1])+'f' As Forms,CHOOSE(CONVERT(NUMERIC,PlateType4),'Deep-etch','PS','Wipe-on','CTP') As Plate,(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper4) As PaperName,[PaperWastage4%] As PaperWastage,PaperConsumptionOther4 As PaperConsumption,PrintAmount4+PaperAmount4 As GrossAmount,PlateAmount4 As PlateAmount,C.Remarks," & _
                              "Adjustment+RAdjustment As Adjustment,PAdjustment,VAT+RVAT As GST,PVAT," & _
                              "LTrim(P.Name) As OrderNo,P.Date,LTrim(LTrim(M1.Pages)+'p/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+LTrim(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+LTrim(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+LTrim(M1.FourColorForms)))+')' As Forme,(SELECT eMail FROM AccountMaster WHERE Code=P.BookPrinter) As EMailId,P.BookPrinter,PlateMaker,P.EstQty01 As FinalQuantity FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild05.Open "SELECT Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,C.ActualQuantity,DuplexPrinting,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,'1' As Color,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size1) As [Size],Pages1 As Pages,TRIM([TotalForms1-¼])+'(¼)+'+TRIM([TotalForms1-½])+'(½)+'+TRIM([TotalForms1-1])+'='+TRIM([TotalForms1-¼]+[TotalForms1-½]+[TotalForms1-1])+'f' As Forms,CHOOSE(VAL(PlateType1),'Deep-etch','PS','Wipe-on','CTP') As Plate,(SELECT TRIM(M3.PrintName)+' (UOM : '+TRIM(M4.PrintName)+'='+TRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper1) As PaperName,[PaperWastage1%] As PaperWastage,PaperConsumptionOther1 As PaperConsumption,PrintAmount1+PaperAmount1 As GrossAmount,PlateAmount1 As PlateAmount,C.Remarks," & _
                              "Adjustment+RAdjustment As Adjustment,PAdjustment,VAT+RVAT As GST,PVAT,TRIM(P.Name) As OrderNo,P.Date,TRIM(TRIM(M1.Pages)+'p/'+TRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+TRIM(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+TRIM(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+TRIM(M1.FourColorForms)))+')' As Forme,(SELECT eMail FROM AccountMaster WHERE Code=P.BookPrinter) As EMailId,P.BookPrinter,PlateMaker,P.EstQty01 As FinalQuantity FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "' UNION " & _
                              "SELECT Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,C.ActualQuantity,DuplexPrinting,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,'2' As Color,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size2) As [Size],Pages2 As Pages,TRIM([TotalForms2-¼])+'(¼)+'+TRIM([TotalForms2-½])+'(½)+'+TRIM([TotalForms2-1])+'='+TRIM([TotalForms2-¼]+[TotalForms2-½]+[TotalForms2-1])+'f' As Forms,CHOOSE(VAL(PlateType2),'Deep-etch','PS','Wipe-on','CTP') As Plate,(SELECT TRIM(M3.PrintName)+' (UOM : '+TRIM(M4.PrintName)+'='+TRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper2) As PaperName,[PaperWastage2%] As PaperWastage,PaperConsumptionOther2 As PaperConsumption,PrintAmount2+PaperAmount2 As GrossAmount,PlateAmount2 As PlateAmount,C.Remarks," & _
                              "Adjustment+RAdjustment As Adjustment,PAdjustment,VAT+RVAT As GST,PVAT,TRIM(P.Name) As OrderNo,P.Date,TRIM(TRIM(M1.Pages)+'p/'+TRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+TRIM(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+TRIM(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+TRIM(M1.FourColorForms)))+')' As Forme,(SELECT eMail FROM AccountMaster WHERE Code=P.BookPrinter) As EMailId,P.BookPrinter,PlateMaker,P.EstQty01 As FinalQuantity FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "' UNION " & _
                              "SELECT Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,C.ActualQuantity,DuplexPrinting,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,'4' As Color,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size4) As [Size],Pages4 As Pages,TRIM([TotalForms4-¼])+'(¼)+'+TRIM([TotalForms4-½])+'(½)+'+TRIM([TotalForms4-1])+'='+TRIM([TotalForms4-¼]+[TotalForms4-½]+[TotalForms4-1])+'f' As Forms,CHOOSE(VAL(PlateType4),'Deep-etch','PS','Wipe-on','CTP') As Plate,(SELECT TRIM(M3.PrintName)+' (UOM : '+TRIM(M4.PrintName)+'='+TRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper4) As PaperName,[PaperWastage4%] As PaperWastage,PaperConsumptionOther4 As PaperConsumption,PrintAmount4+PaperAmount4 As GrossAmount,PlateAmount4 As PlateAmount,C.Remarks," & _
                              "Adjustment+RAdjustment As Adjustment,PAdjustment,VAT+RVAT As GST,PVAT," & _
                              "TRIM(P.Name) As OrderNo,P.Date,TRIM(TRIM(M1.Pages)+'p/'+TRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+TRIM(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+TRIM(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+TRIM(M1.FourColorForms)))+')' As Forme,(SELECT eMail FROM AccountMaster WHERE Code=P.BookPrinter) As EMailId,P.BookPrinter,PlateMaker,P.EstQty01 As FinalQuantity FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
        
    End If
        If rstBookPOChild05.RecordCount = 0 Then
            rptBookPrintOrder01.Section13.Suppress = True: rptBookPrintOrder01.Section14.Suppress = True
        Else
            rstBookPOChild05.MoveFirst
            Qty = Val(rstBookPOChild05.Fields("FinalQuantity").Value)
            rptBookPrintOrder01.Text7.SetText rstBookPOChild05.Fields("OrderNo").Value
            rptBookPrintOrder01.Text42.SetText Format(rstBookPOChild05.Fields("Date").Value, "dd-MM-yyyy")
            rptBookPrintOrder01.Text12.SetText rstBookPOChild05.Fields("TextPrinter").Value
            rptBookPrintOrder01.Text13.SetText rstBookPOChild05.Fields("Item").Value
            rptBookPrintOrder01.Text14.SetText rstBookPOChild05.Fields("FinishSize").Value
            rptBookPrintOrder01.Text64.SetText rstBookPOChild05.Fields("Forme").Value
            rptBookPrintOrder01.Text15.SetText Qty                                                  'To be Check Later "Qty
            EMailID = rstBookPOChild05.Fields("EMailId").Value
            OrderNo = rstBookPOChild05.Fields("OrderNo").Value
            ItemName = rstBookPOChild05.Fields("Item").Value
            Attachment = Trim(rstBookPOChild05.Fields("OrderNo").Value)
            Do While Not rstBookPOChild05.EOF
                Amount = Amount + IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, Val(rstBookPOChild05.Fields("GrossAmount").Value), Val(rstBookPOChild05.Fields("GrossAmount").Value) + Val(rstBookPOChild05.Fields("PlateAmount").Value))
                rstBookPOChild05.MoveNext
            Loop
            rstBookPOChild05.MoveFirst
            Amount = Amount + IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, Val(rstBookPOChild05.Fields("Adjustment").Value), Val(rstBookPOChild05.Fields("Adjustment").Value) + Val(rstBookPOChild05.Fields("PAdjustment").Value))
            FinalAmount = FinalAmount + Amount
            GST = GST + IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, Val(rstBookPOChild05.Fields("GST").Value), Val(rstBookPOChild05.Fields("GST").Value) + Val(rstBookPOChild05.Fields("PVAT").Value))
            Rate = Round(Amount / Val(rstBookPOChild05.Fields("FinalQuantity").Value), 3)
            Amount = Val(rstBookPOChild05.Fields("FinalQuantity").Value) * Rate
            If OrderType = "UC" Then
                rptBookPrintOrder01.Section14.Suppress = True
                rptBookPrintOrder01.Subreport1.OpenSubreport.Text66.SetText "Multi-Form-Printing Details"
                rptBookPrintOrder01.Subreport1.OpenSubreport.Sections(4).Suppress = True
            Else
                rptBookPrintOrder01.Subreport1.OpenSubreport.Sections(3).Suppress = True
                rptBookPrintOrder01.Text22.SetText Format(Rate, "#0.000")
                rptBookPrintOrder01.Text25.SetText Format(Amount, "#0.00")
                TmpAmount = TmpAmount + Amount
            End If
            If Not CheckEmpty(rstBookPOChild05.Fields("Remarks").Value, False) Then
                rptBookPrintOrder01.Section22.Suppress = False
                rptBookPrintOrder01.Text54.SetText rstBookPOChild05.Fields("Remarks").Value
            End If
            FinalRate = FinalRate + Round(Amount / Qty, 3)
            HeaderPrinted = True
        End If
        Amount = 0
        If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
    If DatabaseType = "MS SQL" Then
        rstBookPOChild06.Open "SELECT LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,C.ActualQuantity,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinter,LTrim(FrontPrintingType)+'+'+LTrim(BackPrintingType) As Col,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],(SELECT PrintName FROM GeneralMaster WHERE Code=C.PlateType) As Plate,(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper) As PaperName,[PaperWastage%] As PaperWastage,PaperConsumptionOther As PaperConsumption,PrintAmount+PaperAmount As GrossAmount,PlateAmount,C.Remarks," & _
                              "Adjustment+RAdjustment As Adjustment,PAdjustment,VAT+RVAT As GST,PVAT,LTrim(P.Name) As OrderNo,P.Date,LTrim(LTrim(M1.Pages)+'p/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+LTrim(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+LTrim(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+LTrim(M1.FourColorForms)))+')' As Forms,(SELECT eMail FROM AccountMaster WHERE Code=P.BookPrinter) As EMailId,P.TitlePrinter As CoverPrinter,PlateMaker,P.EstQty01 As FinalQuantity,C1.TotalPlates FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.TitlePrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild06.Open "SELECT Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,C.ActualQuantity,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinter,TRIM(FrontPrintingType)+'+'+TRIM(BackPrintingType) As Col,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],CHOOSE(VAL(PlateType),'Deep-etch','PS','Wipe-on','CTP') As Plate,(SELECT TRIM(M3.PrintName)+' (UOM : '+TRIM(M4.PrintName)+'='+TRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper) As PaperName,[PaperWastage%] As PaperWastage,PaperConsumptionOther As PaperConsumption,PrintAmount+PaperAmount As GrossAmount,PlateAmount,C.Remarks," & _
                              "Adjustment+RAdjustment As Adjustment,PAdjustment,VAT+RVAT As GST,PVAT,TRIM(P.Name) As OrderNo,P.Date,TRIM(TRIM(M1.Pages)+'p/'+TRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+TRIM(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+TRIM(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+TRIM(M1.FourColorForms)))+')' As Forms,(SELECT eMail FROM AccountMaster WHERE Code=P.BookPrinter) As EMailId,P.TitlePrinter As CoverPrinter,PlateMaker,P.EstQty01 As FinalQuantity,C1.TotalPlates FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.TitlePrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
        If rstBookPOChild06.RecordCount = 0 Then
            rptBookPrintOrder01.Section4.Suppress = True: rptBookPrintOrder01.Section17.Suppress = False
        Else
            rstBookPOChild06.MoveFirst
            If Not HeaderPrinted Then
                rptBookPrintOrder01.Text7.SetText rstBookPOChild06.Fields("OrderNo").Value
                rptBookPrintOrder01.Text42.SetText Format(rstBookPOChild06.Fields("Date").Value, "dd-MM-yyyy")
                rptBookPrintOrder01.Text12.SetText rstBookPOChild06.Fields("TitlePrinter").Value
                rptBookPrintOrder01.Text13.SetText rstBookPOChild06.Fields("Item").Value
                rptBookPrintOrder01.Text14.SetText rstBookPOChild06.Fields("FinishSize").Value
                rptBookPrintOrder01.Text64.SetText rstBookPOChild06.Fields("Forms").Value
                rptBookPrintOrder01.Text15.SetText rstBookPOChild06.Fields("FinalQuantity").Value
                If Qty = 0 Then Qty = Val(rstBookPOChild06.Fields("FinalQuantity").Value)
                EMailID = rstBookPOChild06.Fields("EMailId").Value
                OrderNo = rstBookPOChild06.Fields("OrderNo").Value
                ItemName = rstBookPOChild06.Fields("Item").Value
                Attachment = Trim(rstBookPOChild06.Fields("OrderNo").Value)
                HeaderPrinted = True
            End If
            Do While Not rstBookPOChild06.EOF
                Amount = Amount + IIf(rstBookPOChild06.Fields("CoverPrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, Val(rstBookPOChild06.Fields("GrossAmount").Value), Val(rstBookPOChild06.Fields("GrossAmount").Value) + Val(rstBookPOChild06.Fields("PlateAmount").Value))
                rstBookPOChild06.MoveNext
            Loop
            rstBookPOChild06.MoveFirst
            Amount = Amount + IIf(rstBookPOChild06.Fields("CoverPrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, Val(rstBookPOChild06.Fields("Adjustment").Value), Val(rstBookPOChild06.Fields("Adjustment").Value) + Val(rstBookPOChild06.Fields("PAdjustment").Value))
            FinalAmount = FinalAmount + Amount
            GST = GST + IIf(rstBookPOChild06.Fields("CoverPrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, Val(rstBookPOChild06.Fields("GST").Value), Val(rstBookPOChild06.Fields("GST").Value) + Val(rstBookPOChild06.Fields("PVAT").Value))
            Rate = Round(Amount / Val(rstBookPOChild06.Fields("FinalQuantity").Value), 3): Amount = Val(rstBookPOChild06.Fields("FinalQuantity").Value) * Rate
            If OrderType = "UC" Then
                rptBookPrintOrder01.Section17.Suppress = False
                rptBookPrintOrder01.Subreport2.OpenSubreport.Text66.SetText "Single-Form-Printing Details"
                rptBookPrintOrder01.Subreport2.OpenSubreport.Sections(4).Suppress = True
            Else
                rptBookPrintOrder01.Subreport2.OpenSubreport.Sections(3).Suppress = True
                rptBookPrintOrder01.Text29.SetText Format(Rate, "#0.000")
                rptBookPrintOrder01.Text30.SetText Format(Amount, "#0.00")
                TmpAmount = TmpAmount + Amount
            End If
            If Not CheckEmpty(rstBookPOChild06.Fields("Remarks").Value, False) Then
                rptBookPrintOrder01.Section21.Suppress = False
                rptBookPrintOrder01.Text55.SetText rstBookPOChild06.Fields("Remarks").Value
            End If
            FinalRate = FinalRate + Round(Amount / Qty, 3)
        End If
        Amount = 0
                
        If rstBookPOChild10.State = adStateOpen Then rstBookPOChild10.Close
    If DatabaseType = "MS SQL" Then
        rstBookPOChild10.Open "SELECT LTrim(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinter,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMaker,CHOOSE(CONVERT(NUMERIC,C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalFormsFront,C1.TotalFormsBack,C1.PrintRateFront,C1.PrintRateBack,[C1.GST%],C1.GST+C1.RGST AS GST,C1.Adjustment+C1.RAdjustment As Adjustment,C1.PrintAmount,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,[C1.PGST%],C1.PGST,PlateAmount," & _
                              "(SELECT LTrim(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C1.Paper) As PaperName,C1.[PaperWastage%],C1.[PaperConsumptionOther],C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,C2.ActualQuantity,[Ups/Plate],C2.PrintingQuantity,C2.BillingQuantity,LTrim(C2.FrontPrintingColor)+'+'+LTrim(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType,P.TitlePrinter,C1.PlateMaker,PaperWastageMin,PaperRate,PaperAmountBT,RAdjustment,[RGST%],RGST,PaperAmount,PrintAmountBT,PlateAmountBT,(SELECT EMail FROM AccountMaster WHERE Code=P.TitlePrinter) As EMailid,Imposition,PlateType,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C1.Paper) As UOM, " & _
                              "(SELECT LTrim(MAX(PrintingQuantity)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxPrintingQty,P.EstQty01 As FinalQuantity,PaperAmountBT+PrintAmountBT As GrossAmount FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild10.Open "SELECT TRIM(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinter,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMaker,CHOOSE(VAL(C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalFormsFront,C1.TotalFormsBack,C1.PrintRateFront,C1.PrintRateBack,[C1.GST%],C1.GST+C1.RGST AS GST,C1.Adjustment+C1.RAdjustment As Adjustment,C1.PrintAmount,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,[C1.PGST%],C1.PGST,PlateAmount," & _
                              "(SELECT TRIM(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C1.Paper) As PaperName,C1.[PaperWastage%],C1.[PaperConsumptionOther],C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,C2.ActualQuantity,[Ups/Plate],C2.PrintingQuantity,C2.BillingQuantity,TRIM(C2.FrontPrintingColor)+'+'+TRIM(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType,P.TitlePrinter,C1.PlateMaker,PaperWastageMin,PaperRate,PaperAmountBT,RAdjustment,[RGST%],RGST,PaperAmount,PrintAmountBT,PlateAmountBT,(SELECT EMail FROM AccountMaster WHERE Code=P.TitlePrinter) As EMailid,Imposition,PlateType,(SELECT '('+TRIM(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C1.Paper) As UOM, " & _
                              "(SELECT TRIM(MAX(PrintingQuantity)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxPrintingQty,P.EstQty01 As FinalQuantity,PaperAmountBT+PrintAmountBT As GrossAmount FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
        If rstBookPOChild10.RecordCount = 0 Then
            rptBookPrintOrder01.Section26.Suppress = True: rptBookPrintOrder01.Section27.Suppress = True
        Else
            rstBookPOChild10.MoveFirst
            If Not HeaderPrinted Then
                rptBookPrintOrder01.Text7.SetText rstBookPOChild10.Fields("OrderNo").Value
                rptBookPrintOrder01.Text42.SetText Format(rstBookPOChild10.Fields("OrderDate").Value, "dd-MM-yyyy")
                rptBookPrintOrder01.Text12.SetText rstBookPOChild10.Fields("TitlePrinter").Value
                rptBookPrintOrder01.Text13.SetText rstBookPOChild10.Fields("Item").Value
'                rptBookPrintOrder01.Text14.SetText rstBookPOChild10.Fields("FinishSize").Value
'                rptBookPrintOrder01.Text64.SetText rstBookPOChild10.Fields("Forms").Value
                rptBookPrintOrder01.Text15.SetText rstBookPOChild10.Fields("FinalQuantity").Value
                If Qty = 0 Then Qty = Val(rstBookPOChild10.Fields("FinalQuantity").Value)
                EMailID = rstBookPOChild10.Fields("EMailId").Value
                OrderNo = rstBookPOChild10.Fields("OrderNo").Value
                ItemName = rstBookPOChild10.Fields("Item").Value
                Attachment = Trim(rstBookPOChild10.Fields("OrderNo").Value)
                HeaderPrinted = True
            End If
            'Do While Not rstBookPOChild10.EOF
                Amount = Amount + IIf(rstBookPOChild10.Fields("TitlePrinter").Value <> rstBookPOChild10.Fields("PlateMaker").Value, Val(rstBookPOChild10.Fields("GrossAmount").Value), Val(rstBookPOChild10.Fields("GrossAmount").Value) + Val(rstBookPOChild10.Fields("PlateAmountBT").Value))
                rstBookPOChild10.MoveNext
            'Loop
            rstBookPOChild10.MoveFirst
            Amount = Amount + IIf(rstBookPOChild10.Fields("TitlePrinter").Value <> rstBookPOChild10.Fields("PlateMaker").Value, Val(rstBookPOChild10.Fields("Adjustment").Value), Val(rstBookPOChild10.Fields("Adjustment").Value) + Val(rstBookPOChild10.Fields("PAdjustment").Value))
            FinalAmount = FinalAmount + Amount
            GST = GST + IIf(rstBookPOChild10.Fields("TitlePrinter").Value <> rstBookPOChild10.Fields("PlateMaker").Value, Val(rstBookPOChild10.Fields("GST").Value), Val(rstBookPOChild10.Fields("GST").Value) + Val(rstBookPOChild10.Fields("PGST").Value))
            Rate = Round(Amount / Val(rstBookPOChild10.Fields("FinalQuantity").Value), 3): Amount = Val(rstBookPOChild10.Fields("FinalQuantity").Value) * Rate
            If OrderType = "UC" Then
                rptBookPrintOrder01.Section27.Suppress = True
                rptBookPrintOrder01.Subreport4.OpenSubreport.Text66.SetText "Combo-Form-Printing Details"
                rptBookPrintOrder01.Subreport4.OpenSubreport.Sections(26).Suppress = True
            Else
                rptBookPrintOrder01.Subreport3.OpenSubreport.Sections(3).Suppress = True
                rptBookPrintOrder01.Text72.SetText Format(Rate, "#0.000")
                rptBookPrintOrder01.Text73.SetText Format(Amount, "#0.00")
                TmpAmount = TmpAmount + Amount
            End If
            If Not CheckEmpty(rstBookPOChild10.Fields("Remarks").Value, False) Then
                rptBookPrintOrder01.Section21.Suppress = False
                rptBookPrintOrder01.Text55.SetText rstBookPOChild10.Fields("Remarks").Value
            End If
            FinalRate = FinalRate + Round(Amount / Qty, 3)
        End If
        Amount = 0
        
        
        If rstBookPOChild07.State = adStateOpen Then rstBookPOChild07.Close
        rstBookPOChild07.Open "SELECT E.Name As Element,Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,O.Name As Operation,[Number],R.Name As CalcMode,C.Quantity,(SELECT PrintName FROM AccountMaster WHERE Code=P.Laminator) As Laminator,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],(SELECT PrintName FROM GeneralMaster WHERE Code=C.Operation) As LaminationType,ROUND((C.Amount+C.Adjustment)/C.Quantity,3) As UnitRate,ROUND((C.Amount+C.Adjustment)/C.Quantity,3)*C.Quantity As Amount,Amount As GrossAmount,C.Remarks,C.Adjustment,C.GST,TRIM(P.Name) As OrderNo,P.Date,TRIM(TRIM(M1.Pages)+'p/'+TRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+TRIM(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+TRIM(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+TRIM(M1.FourColorForms)))+')' As Forms,(SELECT eMail FROM AccountMaster WHERE Code=P.Laminator) As EMailId,P.EstQty01 As FinalQuantity " & _
                      "FROM (((((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN GeneralMaster E ON C.Element=E.Code)INNER JOIN GeneralMaster O ON C.Operation=O.Code)INNER JOIN GeneralMaster R ON C.CalcMode=R.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code " & _
                      "WHERE P.Code='" & OrderCode & "' AND P.Laminator='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
        If rstBookPOChild07.RecordCount = 0 Then
            rptBookPrintOrder01.Section20.Suppress = True: rptBookPrintOrder01.Section19.Suppress = True
        Else
            rstBookPOChild07.MoveFirst
            If Not HeaderPrinted Then
                rptBookPrintOrder01.Text7.SetText rstBookPOChild07.Fields("OrderNo").Value
                rptBookPrintOrder01.Text42.SetText Format(rstBookPOChild07.Fields("Date").Value, "dd-MM-yyyy")
                rptBookPrintOrder01.Text12.SetText rstBookPOChild07.Fields("Laminator").Value
                rptBookPrintOrder01.Text13.SetText rstBookPOChild07.Fields("Item").Value
                rptBookPrintOrder01.Text14.SetText rstBookPOChild07.Fields("FinishSize").Value
                rptBookPrintOrder01.Text64.SetText rstBookPOChild07.Fields("Forms").Value
                rptBookPrintOrder01.Text15.SetText rstBookPOChild07.Fields("FinalQuantity").Value
                If Qty = 0 Then Qty = Val(rstBookPOChild07.Fields("FinalQuantity").Value)
                EMailID = rstBookPOChild07.Fields("EMailId").Value
                OrderNo = rstBookPOChild07.Fields("OrderNo").Value
                ItemName = rstBookPOChild07.Fields("Item").Value
                Attachment = Trim(rstBookPOChild07.Fields("OrderNo").Value)
                HeaderPrinted = True
            End If
            rptBookPrintOrder01.Text19.SetText rstBookPOChild07.Fields("LaminationType").Value
            rptBookPrintOrder01.Text17.SetText rstBookPOChild07.Fields("Size").Value
            Do While Not rstBookPOChild07.EOF
                Amount = Amount + Val(rstBookPOChild07.Fields("GrossAmount").Value)
                rstBookPOChild07.MoveNext
            Loop
            rstBookPOChild07.MoveFirst
            Do While Not rstBookPOChild07.EOF
            Amount = Amount + Val(rstBookPOChild07.Fields("Adjustment").Value)
            rstBookPOChild07.MoveNext
            Loop
                        FinalAmount = FinalAmount + Amount
            rstBookPOChild07.MoveFirst
            Do While Not rstBookPOChild07.EOF
            GST = GST + Val(rstBookPOChild07.Fields("GST").Value)
            rstBookPOChild07.MoveNext
            Loop
            Rate = Round(Amount / Val(rstBookPOChild07.Fields("FinalQuantity").Value), 3): Amount = Val(rstBookPOChild07.Fields("FinalQuantity").Value) * Rate
            If OrderType = "UC" Then
                rptBookPrintOrder01.Section19.Suppress = True
                rptBookPrintOrder01.Text66.SetText "Misc-Operations Details"
            Else
                rptBookPrintOrder01.Text31.SetText Format(Rate, "#0.000")
                rptBookPrintOrder01.Text32.SetText Format(Amount, "#0.00")
                TmpAmount = TmpAmount + Amount
            End If
            If Not CheckEmpty(rstBookPOChild07.Fields("Remarks").Value, False) Then
                rptBookPrintOrder01.Section18.Suppress = True
                rptBookPrintOrder01.Text56.SetText rstBookPOChild07.Fields("Remarks").Value
            End If
            FinalRate = FinalRate + Round(Amount / Qty, 3)
        End If
        Amount = 0
        If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
        rstBookPOChild08.Open "SELECT Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,C.ActualQuantity,(SELECT PrintName FROM AccountMaster WHERE Code=P.Binder) As Binder,(SELECT PrintName FROM GeneralMaster WHERE Code=C.BindingType) As BindingType,(BindingForms+ExtraForms) As BindingForms,(FormFoldRate*C.BillingQuantity*(BindingForms+ExtraForms)/1000)+(FormPasteRate*C.BillingQuantity/1000)+(FormStitchRate*C.BillingQuantity*(BindingForms+ExtraForms)/1000)+[Rate/Book]*C.BillingQuantity+TotalPkts*PktPackRate+TotalBoxes*BoxPackRate+TotalBoxes*CartageRate As GrossAmount,C.Remarks,Adjustment,VAT As GST,TRIM(P.Name) As OrderNo,P.Date,TRIM(TRIM(M1.Pages)+'p/'+TRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+TRIM(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+TRIM(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+TRIM(M1.FourColorForms)))+')' As Forms," & _
                              "(SELECT eMail FROM AccountMaster WHERE Code=P.Binder) As EMailId,P.EstQty01 As FinalQuantity FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.Binder='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
        If rstBookPOChild08.RecordCount = 0 Then
            rptBookPrintOrder01.Section11.Suppress = True: rptBookPrintOrder01.Section12.Suppress = True
        Else
            rstBookPOChild08.MoveFirst
            If Not HeaderPrinted Then
                rptBookPrintOrder01.Text7.SetText rstBookPOChild08.Fields("OrderNo").Value
                rptBookPrintOrder01.Text42.SetText Format(rstBookPOChild08.Fields("Date").Value, "dd-MM-yyyy")
                rptBookPrintOrder01.Text12.SetText rstBookPOChild08.Fields("Binder").Value
                rptBookPrintOrder01.Text13.SetText rstBookPOChild08.Fields("Item").Value
                rptBookPrintOrder01.Text14.SetText rstBookPOChild08.Fields("FinishSize").Value
                rptBookPrintOrder01.Text64.SetText rstBookPOChild08.Fields("Forms").Value
                rptBookPrintOrder01.Text15.SetText rstBookPOChild08.Fields("FinalQuantity").Value
                EMailID = rstBookPOChild08.Fields("EMailId").Value
                Attachment = Trim(rstBookPOChild08.Fields("OrderNo").Value)
                OrderNo = rstBookPOChild08.Fields("OrderNo").Value
                ItemName = rstBookPOChild08.Fields("Item").Value
                If Qty = 0 Then Qty = Val(rstBookPOChild08.Fields("ActualQuantity").Value)
            End If
            rptBookPrintOrder01.Text40.SetText rstBookPOChild08.Fields("BindingType").Value
            rptBookPrintOrder01.Text41.SetText rstBookPOChild08.Fields("BindingForms").Value
            Do While Not rstBookPOChild08.EOF
                Amount = Amount + Val(rstBookPOChild08.Fields("GrossAmount").Value)
                rstBookPOChild08.MoveNext
            Loop
            rstBookPOChild08.MoveFirst
            Amount = Amount + Val(rstBookPOChild08.Fields("Adjustment").Value)
            FinalAmount = FinalAmount + Amount
            Rate = Round(Amount / Val(rstBookPOChild08.Fields("FinalQuantity").Value), 3): Amount = Val(rstBookPOChild08.Fields("FinalQuantity").Value) * Rate
            GST = GST + Val(rstBookPOChild08.Fields("GST").Value)
            If OrderType = "UC" Then
                rptBookPrintOrder01.Section12.Suppress = True
                rptBookPrintOrder01.Text37.SetText "Item Binding Details"
            Else
                rptBookPrintOrder01.Text33.SetText Format(Rate, "#0.000")
                rptBookPrintOrder01.Text34.SetText Format(Amount, "#0.00")
                TmpAmount = TmpAmount + Amount
            End If
            If Not CheckEmpty(rstBookPOChild08.Fields("Remarks").Value, False) Then
                rptBookPrintOrder01.Section15.Suppress = False
                rptBookPrintOrder01.Text57.SetText rstBookPOChild08.Fields("Remarks").Value
            End If
            FinalRate = FinalRate + Round(Amount / Qty, 3)
        End If
        Amount = 0
        If OrderType = "UC" Then
            rptBookPrintOrder01.Text43.SetText "Amount Details"
            rptBookPrintOrder01.Text47.SetText Format(FinalRate, "#0.000")
            Amount = FinalRate * Qty
        Else
            rptBookPrintOrder01.Section24.Suppress = True
            Amount = TmpAmount
        End If
        rptBookPrintOrder01.Text48.SetText Format(Amount, "#0.00")
        rptBookPrintOrder01.Text63.SetText Format(FinalAmount - Amount, "#0.00")
        rptBookPrintOrder01.Text60.SetText Format(GST, "#0.00")
        rptBookPrintOrder01.Text50.SetText Format(Round(FinalAmount + GST, 0), "#0.00")
        Screen.MousePointer = vbNormal
        If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,EMail,Website FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
        rptBookPrintOrder01.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptBookPrintOrder01.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
        rptBookPrintOrder01.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
        rptBookPrintOrder01.Text58.SetText " (" & Trim(NumberToWords(Round(FinalAmount + GST, 0), True)) & ")"
        rptBookPrintOrder01.Text27.SetText "for " & rptBookPrintOrder01.Text12.Text
        rptBookPrintOrder01.Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptBookPrintOrder01.Subreport1.OpenSubreport.Database.SetDataSource rstBookPOChild05, 3, 1
        rptBookPrintOrder01.Subreport2.OpenSubreport.Database.SetDataSource rstBookPOChild06, 3, 1
        rptBookPrintOrder01.Subreport3.OpenSubreport.Database.SetDataSource rstBookPOChild07, 3, 1
        rptBookPrintOrder01.Subreport4.OpenSubreport.Database.SetDataSource rstBookPOChild10, 3, 1
        
                Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
        Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
        If OutputTo = "S" Then
            FrmReportViewer.EMailID = EMailID
            FrmReportViewer.Subject = "Book Order #" & Trim(OrderNo) + " Book : " + Trim(ItemName)
            FrmReportViewer.Attachment = Attachment
            FrmReportViewer.Message = Message
            Set FrmReportViewer.Report = rptBookPrintOrder01
            FrmReportViewer.Show vbModal
        Else
            If rstBookPOList.State = adStateClosed Then
                If EMailID = "" Or OutputType = "P" Then
                    rptBookPrintOrder01.PaperSource = crPRBinAuto
                    rptBookPrintOrder01.PrintOut False   ' Print Report Without Prompt
                Else
                    rptBookPrintOrder01.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                    rptBookPrintOrder01.ExportOptions.DestinationType = crEDTDiskFile
                    rptBookPrintOrder01.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                    rptBookPrintOrder01.Export False
                    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                    With oOutlookMsg
                        .To = EMailID
                        .Subject = "Book Order #" & Trim(OrderNo) + " Book : " + Trim(ItemName)
                        .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                        .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                        .Importance = olImportanceHigh
                        .ReadReceiptRequested = True
                        .Send
                        If Err.Number = 0 Then cnDatabase.Execute "UPDATE BookPOParent SET BBODStatus=1 WHERE Code='" & OrderCode & "'"
                    End With
                    Set oOutlookMsg = Nothing
                End If
            Else
                rptBookPrintOrder01.PaperSource = crPRBinAuto
                rptBookPrintOrder01.PrintOut
            End If
        End If
        Set rptBookPrintOrder01 = Nothing
        rstBookPOChild0801.MoveNext
    Loop
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild05): Call CloseRecordset(rstBookPOChild06): Call CloseRecordset(rstBookPOChild07): Call CloseRecordset(rstBookPOChild08): Call CloseRecordset(rstBookPOChild0801)
    On Error GoTo 0
End Sub
Public Sub PrintQuotationFormat(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String, Optional ByVal OrderType As String, Optional ByVal BookPOType As String)
    Dim oOutlookMsg As Outlook.MailItem, HeaderPrinted As Boolean, OrderNo As String, ItemName As String, TotalTax As Double, TotalAmount As Double, BillAmount As Double
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptQuotationFormat.Text1.SetText IIf(Left(BookPOType, 1) = "D", "Digital Printing Order", IIf(Right(BookPOType, 1) = "S", "Quotation/Estimation", "Estimation/Costing"))
    If OrderType = "BP" Then
        rptQuotationFormat.Section4.Suppress = True
        rptQuotationFormat.Section20.Suppress = True
        rptQuotationFormat.Section11.Suppress = True
        rptQuotationFormat.Section16.Suppress = True
    ElseIf OrderType = "TP" Then
        rptQuotationFormat.Section13.Suppress = True
        rptQuotationFormat.Section20.Suppress = True
        rptQuotationFormat.Section11.Suppress = True
        rptQuotationFormat.Section16.Suppress = True
    ElseIf OrderType = "CB" Then
        rptQuotationFormat.Section13.Suppress = True
        rptQuotationFormat.Section20.Suppress = True
        rptQuotationFormat.Section11.Suppress = True
        rptQuotationFormat.Section4.Suppress = True
    ElseIf OrderType = "TL" Then
        rptQuotationFormat.Section13.Suppress = True
        rptQuotationFormat.Section4.Suppress = True
        rptQuotationFormat.Section11.Suppress = True
        rptQuotationFormat.Section16.Suppress = True
    ElseIf OrderType = "BB" Then
        rptQuotationFormat.Section13.Suppress = True
        rptQuotationFormat.Section4.Suppress = True
        rptQuotationFormat.Section20.Suppress = True
        rptQuotationFormat.Section16.Suppress = True
    End If
    If rstBookPOChild0801.State = adStateOpen Then rstBookPOChild0801.Close
    If OrderType = "BP" Then
        rstBookPOChild0801.Open "SELECT BookPrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "TP" Then
        rstBookPOChild0801.Open "SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "CB" Then
        rstBookPOChild0801.Open "SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "TL" Then
        rstBookPOChild0801.Open "SELECT Laminator As Vendor FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "BB" Then
        rstBookPOChild0801.Open "SELECT Binder As Vendor FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild0801.Open "SELECT BookPrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT Laminator As Vendor FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT Binder As Vendor FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
    If rstBookPOChild0801.RecordCount > 0 Then rstBookPOChild0801.MoveFirst
    Do While Not rstBookPOChild0801.EOF
        TotalTax = 0: TotalAmount = 0
        HeaderPrinted = False
        If OrderType = "BP" Or OrderType = "ALL" Then
            If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    If DatabaseType = "MS SQL" Then
            rstBookPOChild05.Open "SELECT LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'-pages/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1Col-'+LTrim(M1.OneColorForms)+'f ')+IIF(M1.TwoColorForms=0,'',' 2Col-'+LTrim(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'',' 4Col-'+LTrim(M1.FourColorForms)))+'f)' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,C.ActualQuantity,M1.DuplexPrinting,BillingQuantity01,BillingQuantity02,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size1) As [Size1],Pages1,[TotalForms1-¼],[TotalForms1-½],[TotalForms1-1],CHOOSE(CONVERT(NUMERIC,PlateType1),'Deep-etch','PS','Wipe-on','CTP') As Plate1,PrintRate1,PrintAmount1,PlateRate1,PlateAmount1," & _
                                  "(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper1) As Paper1Name,[PaperWastage1%],PaperConsumptionOther1," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size2) As [Size2],Pages2,[TotalForms2-¼],[TotalForms2-½],[TotalForms2-1],CHOOSE(CONVERT(NUMERIC,PlateType2),'Deep-etch','PS','Wipe-on','CTP') As Plate2,PrintRate2,PrintAmount2,PlateRate2,PlateAmount2," & _
                                  "(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper2) As Paper2Name,[PaperWastage2%],PaperConsumptionOther2," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size4) As [Size4],Pages4,[TotalForms4-¼],[TotalForms4-½],[TotalForms4-1],CHOOSE(CONVERT(NUMERIC,PlateType4),'Deep-etch','PS','Wipe-on','CTP') As Plate4,PrintRate4,PrintAmount4,PlateRate4,PlateAmount4," & _
                                  "(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper4) As Paper4Name,[PaperWastage4%],PaperConsumptionOther4," & _
                                  "TotalPaperConsumption,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT LTrim(eMail) FROM AccountMaster WHERE CODE=P.BookPrinter) As EMailId,M1.Narration,P.BookPrinter,PlateMaker,PaperWastageMin1,PaperWastageMin2,PaperWastageMin4,PaperRate1,PaperRate2,PaperRate4,PaperAmount1,PaperAmount2,PaperAmount4,RBillAmount,RAdjustment,[RVAT%],RVAT,P.EstQty01 As FinalQuantity,P.ProfitMargin,C.Ref  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
            rstBookPOChild05.Open "SELECT Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,TRIM(TRIM(M1.Pages)+'-pages/'+TRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1Col-'+TRIM(M1.OneColorForms)+'f ')+IIF(M1.TwoColorForms=0,'',' 2Col-'+TRIM(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'',' 4Col-'+TRIM(M1.FourColorForms)))+'f)' As Forme,TRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,C.ActualQuantity,M1.DuplexPrinting,BillingQuantity01,BillingQuantity02,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size1) As [Size1],Pages1,[TotalForms1-¼],[TotalForms1-½],[TotalForms1-1],CHOOSE(VAL(PlateType1),'Deep-etch','PS','Wipe-on','CTP') As Plate1,PrintRate1,PrintAmount1,PlateRate1,PlateAmount1," & _
                                  "(SELECT TRIM(M3.PrintName)+' (UOM : '+TRIM(M4.PrintName)+'='+TRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper1) As Paper1Name,[PaperWastage1%],PaperConsumptionOther1," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size2) As [Size2],Pages2,[TotalForms2-¼],[TotalForms2-½],[TotalForms2-1],CHOOSE(VAL(PlateType2),'Deep-etch','PS','Wipe-on','CTP') As Plate2,PrintRate2,PrintAmount2,PlateRate2,PlateAmount2," & _
                                  "(SELECT TRIM(M3.PrintName)+' (UOM : '+TRIM(M4.PrintName)+'='+TRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper2) As Paper2Name,[PaperWastage2%],PaperConsumptionOther2," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size4) As [Size4],Pages4,[TotalForms4-¼],[TotalForms4-½],[TotalForms4-1],CHOOSE(VAL(PlateType4),'Deep-etch','PS','Wipe-on','CTP') As Plate4,PrintRate4,PrintAmount4,PlateRate4,PlateAmount4," & _
                                  "(SELECT TRIM(M3.PrintName)+' (UOM : '+TRIM(M4.PrintName)+'='+TRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper4) As Paper4Name,[PaperWastage4%],PaperConsumptionOther4," & _
                                  "TotalPaperConsumption,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT TRIM(eMail) FROM AccountMaster WHERE CODE=P.BookPrinter) As EMailId,M1.Narration,P.BookPrinter,PlateMaker,PaperWastageMin1,PaperWastageMin2,PaperWastageMin4,PaperRate1,PaperRate2,PaperRate4,PaperAmount1,PaperAmount2,PaperAmount4,RBillAmount,RAdjustment,[RVAT%],RVAT,P.EstQty01 As FinalQuantity,P.ProfitMargin,C.Ref  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
            If rstBookPOChild05.RecordCount = 0 Then
                rptQuotationFormat.Section13.Suppress = True
            Else
                rstBookPOChild05.MoveFirst
                TotalAmount = TotalAmount + IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("PBillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value)
                TotalTax = TotalTax + IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("VAT").Value + rstBookPOChild05.Fields("RVAT").Value, rstBookPOChild05.Fields("VAT").Value + rstBookPOChild05.Fields("PVAT").Value + rstBookPOChild05.Fields("RVAT").Value)
                TaxableAmount = (TotalAmount - TotalTax)
                UnitRateBT = TaxableAmount / Val(rstBookPOChild05.Fields("FinalQuantity").Value)
                UnitRate = TotalAmount / Val(rstBookPOChild05.Fields("FinalQuantity").Value)
                
                PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild05.Fields("ProfitMargin").Value) / 100))
                PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild05.Fields("ProfitMargin").Value) / 100))
                PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild05.Fields("FinalQuantity").Value)
                PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild05.Fields("FinalQuantity").Value)
                
                rptQuotationFormat.Text7.SetText rstBookPOChild05.Fields("OrderNo").Value
                rptQuotationFormat.Text43.SetText rstBookPOChild05.Fields("Ref").Value
                rptQuotationFormat.Text42.SetText Format(rstBookPOChild05.Fields("OrderDate").Value, "dd-MM-yyyy")
                rptQuotationFormat.Text12.SetText rstBookPOChild05.Fields("TextPrinter").Value
                rptQuotationFormat.Text13.SetText rstBookPOChild05.Fields("Item").Value
                rptQuotationFormat.Text14.SetText rstBookPOChild05.Fields("FinishSize").Value
                rptQuotationFormat.Text64.SetText rstBookPOChild05.Fields("Forme").Value
                rptQuotationFormat.Text15.SetText rstBookPOChild05.Fields("FinalQuantity").Value
                rptQuotationFormat.Subreport1.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("PBillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value), True)) & ")"
                EMailID = rstBookPOChild05.Fields("EMailId").Value
                OrderNo = rstBookPOChild05.Fields("OrderNo").Value
                ItemName = rstBookPOChild05.Fields("Item").Value
                Attachment = Trim(rstBookPOChild05.Fields("OrderNo").Value)
                HeaderPrinted = True
            End If
        End If
        If OrderType = "TP" Or OrderType = "ALL" Then
            If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
    If DatabaseType = "MS SQL" Then
            rstBookPOChild06.Open "SELECT LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'-pages/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1Col-'+LTrim(M1.OneColorForms)+'f ')+IIF(M1.TwoColorForms=0,'',' 2Col-'+LTrim(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'',' 4Col-'+LTrim(M1.FourColorForms)))+'f)' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As CoverPrinter,C.ActualQuantity,C.BillingQuantity,C.FrontPrintingType,C.BackPrintingType,Imposition,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],(SELECT PrintName FROM GeneralMaster WHERE Code=C.PlateType) As Plate,PrintRate,PrintAmount,PlateRate,PlateAmount,(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper) As PaperName," & _
                                  "[PaperWastage%],PaperConsumptionOther,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT LTrim(eMail) FROM AccountMaster WHERE Code=P.TitlePrinter) As EMailId,M1.Narration,P.TitlePrinter,PlateMaker,PaperWastageMin,PaperRate,PaperAmount,RBillAmount,RAdjustment,[RVAT%],RVAT,P.EstQty01 As FinalQuantity,P.ProfitMargin,C.Sets,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOM,C.Ref  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.TitlePrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
            rstBookPOChild06.Open "SELECT Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,TRIM(TRIM(M1.Pages)+'-pages/'+TRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1Col-'+TRIM(M1.OneColorForms)+'f ')+IIF(M1.TwoColorForms=0,'',' 2Col-'+TRIM(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'',' 4Col-'+TRIM(M1.FourColorForms)))+'f)' As Forme,TRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As CoverPrinter,C.ActualQuantity,C.BillingQuantity,C.FrontPrintingType,C.BackPrintingType,Imposition,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],CHOOSE(VAL(PlateType),'Deep-etch','PS','Wipe-on','CTP') As Plate,PrintRate,PrintAmount,PlateRate,PlateAmount,(SELECT TRIM(M3.PrintName)+' (UOM : '+TRIM(M4.PrintName)+'='+TRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper) As PaperName," & _
                                  "[PaperWastage%],PaperConsumptionOther,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT TRIM(eMail) FROM AccountMaster WHERE Code=P.TitlePrinter) As EMailId,M1.Narration,P.TitlePrinter,PlateMaker,PaperWastageMin,PaperRate,PaperAmount,RBillAmount,RAdjustment,[RVAT%],RVAT,P.EstQty01 As FinalQuantity,P.ProfitMargin,C.Sets,(SELECT '('+TRIM(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOM,C.Ref  " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.TitlePrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
            If rstBookPOChild06.RecordCount = 0 Then
                rptQuotationFormat.Section4.Suppress = True
            Else
                rstBookPOChild06.MoveFirst
                TotalAmount = TotalAmount + IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("PBillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value)
                TotalTax = TotalTax + IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, rstBookPOChild06.Fields("VAT").Value + rstBookPOChild06.Fields("RVAT").Value, rstBookPOChild06.Fields("VAT").Value + rstBookPOChild06.Fields("PVAT").Value + rstBookPOChild06.Fields("RVAT").Value)
                TaxableAmount = (TotalAmount - TotalTax)
                UnitRateBT = TaxableAmount / Val(rstBookPOChild06.Fields("FinalQuantity").Value)
                UnitRate = TotalAmount / Val(rstBookPOChild06.Fields("FinalQuantity").Value)
                
                PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild06.Fields("ProfitMargin").Value) / 100))
                PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild06.Fields("ProfitMargin").Value) / 100))
                PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild05.Fields("FinalQuantity").Value)
                PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild06.Fields("FinalQuantity").Value)
                
                
                If Not HeaderPrinted Then
                    rptQuotationFormat.Text7.SetText rstBookPOChild06.Fields("OrderNo").Value
                    rptQuotationFormat.Text43.SetText rstBookPOChild06.Fields("Ref").Value
                    rptQuotationFormat.Text42.SetText Format(rstBookPOChild06.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptQuotationFormat.Text12.SetText rstBookPOChild06.Fields("CoverPrinter").Value
                    rptQuotationFormat.Text13.SetText rstBookPOChild06.Fields("Item").Value
                    rptQuotationFormat.Text14.SetText rstBookPOChild06.Fields("FinishSize").Value
                    rptQuotationFormat.Text64.SetText rstBookPOChild06.Fields("Forme").Value
                    rptQuotationFormat.Text15.SetText rstBookPOChild06.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild06.Fields("EMailId").Value
                    OrderNo = rstBookPOChild06.Fields("OrderNo").Value
                    ItemName = rstBookPOChild06.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild06.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                rptQuotationFormat.Subreport2.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("PBillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value), True)) & ")"
            End If
        End If
        If OrderType = "CB" Or OrderType = "ALL" Then
                If rstBookPOChild10.State = adStateOpen Then rstBookPOChild10.Close
    If DatabaseType = "MS SQL" Then
        rstBookPOChild10.Open "SELECT LTrim(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As CoverPrinter,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMakers,CHOOSE(CONVERT(NUMERIC,C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalFormsFront,C1.TotalFormsBack,C1.PrintRateFront,C1.PrintRateBack,C1.[GST%],C1.GST,C1.Adjustment,C1.PrintAmount,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,C1.[PGST%],C1.PGST,PlateAmount," & _
                          "(SELECT LTrim(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C1.Paper) As PaperName,C1.[PaperWastage%],C1.[PaperConsumptionOther],C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,C2.ActualQuantity,[Ups/Plate],C2.PrintingQuantity,C2.BillingQuantity,LTrim(C2.FrontPrintingColor)+'+'+LTrim(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType,P.TitlePrinter AS TitlePrinterCode ,C1.PlateMaker As PlateMakerCode," & _
                          "PaperWastageMin,PaperRate,PaperAmountBT,RAdjustment,[RGST%],RGST,PaperAmount,PrintAmountBT,PlateAmountBT,(SELECT EMail FROM AccountMaster WHERE Code=P.TitlePrinter) As EMail,Imposition,PlateType,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C1.Paper) As UOM,(SELECT LTrim(MAX(PrintingQuantity)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxPrintingQty, " & _
                          "P.EstQty01 As FinalQuantity,P.ProfitMargin,(SELECT LTrim(MAX(FrontPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxFrontColor,(SELECT LTrim(MAX(BackPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxBackColor FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild10.Open "SELECT LTrim(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As CoverPrinter,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMakers,CHOOSE(VAL(C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalFormsFront,C1.TotalFormsBack,C1.PrintRateFront,C1.PrintRateBack,[C1.GST%],C1.GST,C1.Adjustment,C1.PrintAmount,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,[C1.PGST%],C1.PGST,PlateAmount," & _
                          "(SELECT LTrim(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C1.Paper) As PaperName,C1.[PaperWastage%],C1.[PaperConsumptionOther],C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,C2.ActualQuantity,[Ups/Plate],C2.PrintingQuantity,C2.BillingQuantity,LTrim(C2.FrontPrintingColor)+'+'+LTrim(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType,P.TitlePrinter AS TitlePrinterCode ,C1.PlateMaker As PlateMakerCode," & _
                          "PaperWastageMin,PaperRate,PaperAmountBT,RAdjustment,[RGST%],RGST,PaperAmount,PrintAmountBT,PlateAmountBT,(SELECT EMail FROM AccountMaster WHERE Code=P.TitlePrinter) As EMail,Imposition,PlateType,(SELECT '('+LTrim(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C1.Paper) As UOM,(SELECT LTrim(MAX(PrintingQuantity)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxPrintingQty, " & _
                          "P.EstQty01 As FinalQuantity,P.ProfitMargin,(SELECT LTrim(MAX(FrontPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxFrontColor,(SELECT LTrim(MAX(BackPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxBackColor FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
                If rstBookPOChild10.RecordCount = 0 Then
                   rptQuotationFormat.Section16.Suppress = True
            Else
                rstBookPOChild10.MoveFirst
                TotalTax = TotalTax + IIf(rstBookPOChild10.Fields("TitlePrinterCode").Value <> rstBookPOChild10.Fields("PlateMakerCode").Value, rstBookPOChild10.Fields("GST").Value + rstBookPOChild10.Fields("RGST").Value, rstBookPOChild10.Fields("GST").Value + rstBookPOChild10.Fields("PGST").Value + rstBookPOChild10.Fields("RGST").Value)
                TotalAmount = TotalAmount + IIf(rstBookPOChild10.Fields("TitlePrinterCode").Value <> rstBookPOChild10.Fields("PlateMakerCode").Value, rstBookPOChild10.Fields("PrintAmount").Value + rstBookPOChild10.Fields("PaperAmount").Value, rstBookPOChild10.Fields("PrintAmount").Value + rstBookPOChild10.Fields("PlateAmount").Value + rstBookPOChild10.Fields("PaperAmount").Value)
                TaxableAmount = (TotalAmount - TotalTax)
                UnitRateBT = TaxableAmount / Val(rstBookPOChild10.Fields("FinalQuantity").Value)
                UnitRate = TotalAmount / Val(rstBookPOChild10.Fields("FinalQuantity").Value)
                
                PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild10.Fields("ProfitMargin").Value) / 100))
                PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild10.Fields("ProfitMargin").Value) / 100))
                PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild10.Fields("FinalQuantity").Value)
                PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild10.Fields("FinalQuantity").Value)
                
                If Not HeaderPrinted Then
                    rptQuotationFormat.Text7.SetText rstBookPOChild10.Fields("OrderNo").Value
                    rptQuotationFormat.Text42.SetText Format(rstBookPOChild10.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptQuotationFormat.Text12.SetText rstBookPOChild10.Fields("TitlePrinter").Value
                    rptQuotationFormat.Text13.SetText rstBookPOChild10.Fields("Item").Value
'                    rptQuotationFormat.Text14.SetText rstBookPOChild10.Fields("FinishSize").Value
                    rptQuotationFormat.Text59.SetText ""
'                    rptQuotationFormat.Text64.SetText rstBookPOChild10.Fields("Forme").Value
                    rptQuotationFormat.Text9.SetText ""
                    rptQuotationFormat.Text15.SetText rstBookPOChild10.Fields("FinalQuantity").Value     'IIf(OrderType = "CB", (Val(rstBookPOChild10.Fields("MaxPrintingQty").Value)), Val(rstBookPOChild10.Fields("FinalQuantity").Value))
                    EMailID = rstBookPOChild10.Fields("EMailId").Value
                    OrderNo = rstBookPOChild10.Fields("OrderNo").Value
                    ItemName = rstBookPOChild10.Fields("Item").Value
                    Attachment = LTrim(rstBookPOChild10.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                rptQuotationFormat.Subreport6.OpenSubreport.Text25.SetText " (" & LTrim(NumberToWords(IIf(rstBookPOChild10.Fields("TitlePrinterCode").Value <> rstBookPOChild10.Fields("PlateMakerCode").Value, rstBookPOChild10.Fields("BillAmount").Value + rstBookPOChild10.Fields("RBillAmount").Value, rstBookPOChild10.Fields("BillAmount").Value + rstBookPOChild10.Fields("PBillAmount").Value + rstBookPOChild10.Fields("RBillAmount").Value), True)) & ")"
            End If
        End If
        If OrderType = "TL" Or OrderType = "ALL" Then
            If rstBookPOChild07.State = adStateOpen Then rstBookPOChild07.Close
            rstBookPOChild07.Open "SELECT E.Name As Element,O.Name As Operation,[Number],OS.Name As [Size],Quantity,M.Name As CalcMode,Rate,Amount,Adjustment,[GST%],GST,BillAmount,C.Remarks,LTrim(I.PrintName)+IIF(I.Price=0,'',' (Price : Rs. '+Format(I.Price,'0.00')+')') As Item,FS.PrintName As FinishSize,LTrim(LTrim(I.Pages)+'-pages/'+LTrim(I.Forms)+'f('+IIF(I.OneColorForms=0,'',' 1Col-'+LTrim(I.OneColorForms)+'f ')+IIF(I.TwoColorForms=0,'',' 2Col-'+LTrim(I.TwoColorForms)+'f ')+IIF(I.FourColorForms=0,'',' 4Col-'+LTrim(I.FourColorForms)))+'f)' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.Laminator) As Laminator,(SELECT eMail FROM AccountMaster WHERE Code=P.Laminator) As EMailId,I.Narration,P.EstQty01 As FinalQuantity,P.ProfitMargin " & _
                                                               "FROM ((((((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster FS ON I.FinishSize=FS.Code) INNER JOIN GeneralMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code) LEFT JOIN GeneralMaster OS ON C.[Size]=OS.Code WHERE P.Code='" & OrderCode & "' AND P.Laminator='" & rstBookPOChild0801.Fields("Vendor").Value & "' ORDER BY E.Name,O.Name", cnDatabase, adOpenKeyset, adLockOptimistic
            If rstBookPOChild07.RecordCount = 0 Then
                rptQuotationFormat.Section20.Suppress = True
            Else
                With rstBookPOChild07
                    .MoveFirst
                    BillAmount = 0
                    Do While Not .EOF
                        BillAmount = BillAmount + Val(.Fields("BillAmount").Value)
                        TotalAmount = TotalAmount + Val(.Fields("BillAmount").Value)
                        TotalTax = TotalTax + Val(.Fields("GST").Value)
                        TaxableAmount = (TotalAmount - TotalTax)
                        UnitRateBT = TaxableAmount / Val(rstBookPOChild07.Fields("FinalQuantity").Value)
                        UnitRate = TotalAmount / Val(rstBookPOChild07.Fields("FinalQuantity").Value)
                        
                PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild07.Fields("ProfitMargin").Value) / 100))
                PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild07.Fields("ProfitMargin").Value) / 100))
                PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild07.Fields("FinalQuantity").Value)
                PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild07.Fields("FinalQuantity").Value)
                
                        .MoveNext
                    Loop
                End With
                rstBookPOChild07.MoveFirst
                If Not HeaderPrinted Then
                    rptQuotationFormat.Text7.SetText rstBookPOChild07.Fields("OrderNo").Value
                    rptQuotationFormat.Text42.SetText Format(rstBookPOChild07.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptQuotationFormat.Text12.SetText rstBookPOChild07.Fields("Laminator").Value
                    rptQuotationFormat.Text13.SetText rstBookPOChild07.Fields("Item").Value
                    rptQuotationFormat.Text14.SetText rstBookPOChild07.Fields("FinishSize").Value
                    rptQuotationFormat.Text64.SetText rstBookPOChild07.Fields("Forme").Value
                    rptQuotationFormat.Text15.SetText rstBookPOChild07.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild07.Fields("EMailId").Value
                    OrderNo = rstBookPOChild07.Fields("OrderNo").Value
                    ItemName = rstBookPOChild07.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild07.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                rptQuotationFormat.Subreport4.OpenSubreport.Text25.SetText "Amount Payable : " & Trim(NumberToWords(Round(BillAmount, 0), True))
            End If
        End If
        If OrderType = "BB" Or OrderType = "ALL" Then
            If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
            rstBookPOChild08.Open "SELECT LTrim(M1.PrintName)+IIF(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'-pages/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1Col-'+LTrim(M1.OneColorForms)+'f ')+IIF(M1.TwoColorForms=0,'',' 2Col-'+LTrim(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'',' 4Col-'+LTrim(M1.FourColorForms)))+'f)' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.Binder) As Binder,C.ActualQuantity, " & _
                                  "C.BillingQuantity,(SELECT PrintName FROM GeneralMaster WHERE Code=C.BindingType) As BindingType,BindingForms,ExtraForms,FormFoldRate,FormPasteRate,FormStitchRate,[Rate/Book],TotalPkts,PktPackRate,TotalBoxes,BoxPackRate,CartageRate,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,(SELECT eMail FROM AccountMaster WHERE Code=P.Binder) As EMailId,M1.Narration,P.EstQty01 As FinalQuantity,P.ProfitMargin " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.Binder='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
            If rstBookPOChild08.RecordCount = 0 Then
                rptQuotationFormat.Section11.Suppress = True
            Else
                rstBookPOChild08.MoveFirst
                TotalAmount = TotalAmount + Val(rstBookPOChild08.Fields("BillAmount").Value)
                TotalTax = TotalTax + Val(rstBookPOChild08.Fields("VAT").Value)
                TaxableAmount = (TotalAmount - TotalTax)
                UnitRateBT = TaxableAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
                UnitRate = TotalAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
                
                PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild08.Fields("ProfitMargin").Value) / 100))
                PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild08.Fields("ProfitMargin").Value) / 100))
                PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
                PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
                
                If Not HeaderPrinted Then
                    rptQuotationFormat.Text7.SetText rstBookPOChild08.Fields("OrderNo").Value
                    rptQuotationFormat.Text42.SetText rstBookPOChild08.Fields("OrderDate").Value
                    rptQuotationFormat.Text12.SetText rstBookPOChild08.Fields("Binder").Value
                    rptQuotationFormat.Text13.SetText rstBookPOChild08.Fields("Item").Value
                    rptQuotationFormat.Text14.SetText rstBookPOChild08.Fields("FinishSize").Value
                    rptQuotationFormat.Text64.SetText rstBookPOChild08.Fields("Forme").Value
                    rptQuotationFormat.Text15.SetText rstBookPOChild08.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild08.Fields("EMailId").Value
                    OrderNo = rstBookPOChild08.Fields("OrderNo").Value
                    ItemName = rstBookPOChild08.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild08.Fields("OrderNo").Value)
                End If
                rptQuotationFormat.Subreport3.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(Round(rstBookPOChild08.Fields("BillAmount").Value, 0), True)) & ")"
            End If
        End If
        If rstBookPOChild09.State = adStateOpen Then rstBookPOChild09.Close
    If DatabaseType = "MS SQL" Then
        rstBookPOChild09.Open "SELECT Choose(CONVERT(NUMERIC,Category),(SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item),(SELECT LTrim(M3.PrintName)+' (UOM : '+LTrim(M4.PrintName)+'='+LTrim(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='R' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item)) As ItemName,[Consumption/Item],OrderQuantity,TotalConsumption,Rate,Amount " & _
                              "FROM BookPOChild0801 T WHERE T.Code='" & OrderCode & "' AND T.Vendor='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild09.Open "SELECT Choose(Val(Category),(SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item),(SELECT TRIM(M3.PrintName)+' (UOM : '+TRIM(M4.PrintName)+'='+TRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='R' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item)) As ItemName,[Consumption/Item],OrderQuantity,TotalConsumption,Rate,Amount " & _
                              "FROM BookPOChild0801 T WHERE T.Code='" & OrderCode & "' AND T.Vendor='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
        If rstBookPOChild09.RecordCount = 0 Then rptQuotationFormat.Section12.Suppress = True
'            Else
'                rstBookPOChild08.MoveFirst
'                TotalAmount = TotalAmount + Val(rstBookPOChild08.Fields("BillAmount").Value)
'                TotalTax = TotalTax + Val(rstBookPOChild08.Fields("VAT").Value)
'                TaxableAmount = (TotalAmount - TotalTax)
'                UnitRateBT = TaxableAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
'                UnitRate = TotalAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
'
'                PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild08.Fields("ProfitMargin").Value) / 100))
'                PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild08.Fields("ProfitMargin").Value) / 100))
'                PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
'                PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
'
'
        Screen.MousePointer = vbNormal
        If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,EMail,Website FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
        rptQuotationFormat.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptQuotationFormat.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
        rptQuotationFormat.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
        rptQuotationFormat.Text27.SetText "for " & rptQuotationFormat.Text12.Text
        rptQuotationFormat.Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
        If TotalAmount = 0 Then
'            rptQuotationFormat.Section14.Suppress = True
        Else
            rptQuotationFormat.Text17.SetText Format(TotalTax, "##0.00")
            rptQuotationFormat.Text18.SetText Format(TotalAmount, "##0.00")
            rptQuotationFormat.Text22.SetText Format(TaxableAmount, "##0.00")
            rptQuotationFormat.Text29.SetText Format(UnitRateBT, "##0.000")
            rptQuotationFormat.Text33.SetText Format(UnitRate, "##0.000")
            
            rptQuotationFormat.Text37.SetText Format(PMTotalTax, "##0.00")
            rptQuotationFormat.Text18.SetText Format(TotalAmount, "##0.00")
            rptQuotationFormat.Text34.SetText Format(PMTaxableAmount, "##0.00")
            rptQuotationFormat.Text38.SetText Format(PMUnitRateBT, "##0.000")
            rptQuotationFormat.Text40.SetText Format(PMUnitRate, "##0.000")
                        
            rptQuotationFormat.Text19.SetText " (" & Trim(NumberToWords(Round(TotalAmount, 0), True)) & ")"
                        
        End If
        rptQuotationFormat.Subreport1.OpenSubreport.Database.SetDataSource rstBookPOChild05, 3, 1
        rptQuotationFormat.Subreport2.OpenSubreport.Database.SetDataSource rstBookPOChild06, 3, 1
        rptQuotationFormat.Subreport4.OpenSubreport.Database.SetDataSource rstBookPOChild07, 3, 1
        rptQuotationFormat.Subreport3.OpenSubreport.Database.SetDataSource rstBookPOChild08, 3, 1
        rptQuotationFormat.Subreport5.OpenSubreport.Database.SetDataSource rstBookPOChild09, 3, 1
        rptQuotationFormat.Subreport6.OpenSubreport.Database.SetDataSource rstBookPOChild10, 3, 1
        Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
        Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
        If OutputTo = "S" Then
            FrmReportViewer.EMailID = EMailID
            FrmReportViewer.Subject = "Book Order #" & Trim(OrderNo) + " Book : " + Trim(ItemName)
            FrmReportViewer.Attachment = Attachment
            FrmReportViewer.Message = Message
            Set FrmReportViewer.Report = rptQuotationFormat
            FrmReportViewer.Show vbModal
        Else
            If rstBookPOList.State = adStateClosed Then
                If EMailID = "" Or OutputType = "P" Then
                    rptQuotationFormat.PaperSource = crPRBinAuto
                    rptQuotationFormat.PrintOut False   ' Print Report Without Prompt
                Else
                    rptQuotationFormat.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                    rptQuotationFormat.ExportOptions.DestinationType = crEDTDiskFile
                    rptQuotationFormat.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                    rptQuotationFormat.Export False
                    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                    With oOutlookMsg
                        .To = EMailID
                        .Subject = "Book Order #" & Trim(OrderNo) + " Book : " + Trim(ItemName)
                        .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                        .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                        .Importance = olImportanceHigh
                        .ReadReceiptRequested = True
                        .Send
                        If Err.Number = 0 Then cnDatabase.Execute "UPDATE BookPOParent SET BBODStatus=1 WHERE Code='" & OrderCode & "'"
                    End With
                    Set oOutlookMsg = Nothing
                End If
            Else
                rptQuotationFormat.PaperSource = crPRBinAuto
                rptQuotationFormat.PrintOut
            End If
        End If
        Set rptQuotationFormat = Nothing
        rstBookPOChild0801.MoveNext
        If OrderType = "BP" Or OrderType = "TP" Or OrderType = "CB" Or OrderType = "TL" Or OrderType = "BB" Then Exit Do
    Loop
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild05): Call CloseRecordset(rstBookPOChild06): Call CloseRecordset(rstBookPOChild07): Call CloseRecordset(rstBookPOChild08): Call CloseRecordset(rstBookPOChild0801): Call CloseRecordset(rstBookPOChild09)
    On Error GoTo 0
    Screen.MousePointer = vbNormal
End Sub
Public Sub PrintBookPrintOrder02(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String, Optional ByVal OrderType As String, Optional ByVal BookPOType As String, Optional ByVal cRpt As String)
    Dim oOutlookMsg As Outlook.MailItem, HeaderPrinted As Boolean, OrderNo As String, ItemName As String, TotalTax As Double, TotalAmount As Double, BillAmount As Double, PBillAmount As Double, RBillAmount As Double, BOMAmount As Double, TotalPages As Double, TotalForme As Double
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    'rptBookPrintOrder02.Text1.SetText IIf(cRpt = "1", IIf(Right(BookPOType, 1) = "S", "Proforma Invoice", "Item Order"), IIf(Right(BookPOType, 1) = "S", "Quotation/Estimation", "Estimation/Costing"))
    rptBookPrintOrder02.Text1.SetText IIf(Left(BookPOType, 1) = "D", "Digital Printing Order", IIf(Right(BookPOType, 1) = "S", "Proforma Invoice", "Item Order"))
    If Right(BookPOType, 1) = "P" Then
        rptBookPrintOrder02.Section17.Suppress = True
    Else
        rptBookPrintOrder02.Section17.Suppress = False
    End If
    If OrderType = "BP" Then
        rptBookPrintOrder02.Section4.Suppress = True
        rptBookPrintOrder02.Section20.Suppress = True
        rptBookPrintOrder02.Section11.Suppress = True
        rptBookPrintOrder02.Section15.Suppress = True
        rptBookPrintOrder02.Section16.Suppress = True
        rptBookPrintOrder02.Section21.Suppress = True
    ElseIf OrderType = "TP" Then
        rptBookPrintOrder02.Section13.Suppress = True
        rptBookPrintOrder02.Section20.Suppress = True
        rptBookPrintOrder02.Section11.Suppress = True
        rptBookPrintOrder02.Section15.Suppress = True
        rptBookPrintOrder02.Section16.Suppress = True
        rptBookPrintOrder02.Section18.Suppress = True
    ElseIf OrderType = "CB" Then
        rptBookPrintOrder02.Section13.Suppress = True
        rptBookPrintOrder02.Section20.Suppress = True
        rptBookPrintOrder02.Section11.Suppress = True
        rptBookPrintOrder02.Section15.Suppress = True
        rptBookPrintOrder02.Section4.Suppress = True
        rptBookPrintOrder02.Section18.Suppress = True
        rptBookPrintOrder02.Section21.Suppress = True
    ElseIf OrderType = "TL" Then
        rptBookPrintOrder02.Section13.Suppress = True
        rptBookPrintOrder02.Section4.Suppress = True
        rptBookPrintOrder02.Section11.Suppress = True
        rptBookPrintOrder02.Section15.Suppress = True
        rptBookPrintOrder02.Section16.Suppress = True
        rptBookPrintOrder02.Section18.Suppress = True
        rptBookPrintOrder02.Section21.Suppress = True
    ElseIf OrderType = "BB" Then
        rptBookPrintOrder02.Section11.Suppress = True
        rptBookPrintOrder02.Section13.Suppress = True
        rptBookPrintOrder02.Section4.Suppress = True
        rptBookPrintOrder02.Section20.Suppress = True
        rptBookPrintOrder02.Section16.Suppress = True
        rptBookPrintOrder02.Section18.Suppress = True
        rptBookPrintOrder02.Section21.Suppress = True
    End If
    If rstBookPOChild0801.State = adStateOpen Then rstBookPOChild0801.Close
    If OrderType = "BP" Then
        rstBookPOChild0801.Open "SELECT BookPrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "TP" Then
        rstBookPOChild0801.Open "SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "CB" Then
        rstBookPOChild0801.Open "SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "TL" Then
        rstBookPOChild0801.Open "SELECT Laminator As Vendor FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "BB" Then
        rstBookPOChild0801.Open "SELECT Binder As Vendor FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild0801.Open "SELECT BookPrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT TitlePrinter As Vendor FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT Laminator As Vendor FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT Binder As Vendor FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
    If rstBookPOChild0801.RecordCount > 0 Then rstBookPOChild0801.MoveFirst
    Do While Not rstBookPOChild0801.EOF
        TotalTax = 0: TotalAmount = 0
        HeaderPrinted = False
    If OrderType = "BP" Or OrderType = "ALL" Then
            If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
                    rstBookPOChild05.Open "SELECT IIF(C.ElementPrintName<>'',C.ElementPrintName,(SELECT PrintName FROM ElementMaster WHERE Code=C.Element)) As Element,LTRIM(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTRIM(C.Forms) As Forme,LTRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinter,C.ActualQuantity,C.DuplexPrinting,BillingQuantity,(SELECT PrintName+'/'+Convert(Varchar,C.[Pages/Form]) FROM GeneralMaster WHERE Code=C.Size) As [Size],C.Pages,(SELECT PrintName FROM GeneralMaster WHERE Code=C.Color) As Col,[Forms-¼],[Forms-½],[Forms-1-F&B],[Forms-1-W&T]," & _
                                          "(SELECT PrintName FROM GeneralMaster WHERE Code=C.PlateType) As Plate,PrintRate,PrintAmount,PlateRate,PlateAmount," & _
                                          "(SELECT LTRIM(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper) As PaperName,[PaperWastage%],PaperConsumptionOther,(SELECT ''+LTrim(M6.PrintName)+' ' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOM,(SELECT LTrim(M7.[Weight/Unit]) FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOMwt,(SELECT LTrim(M6.[Value1]) FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOMqty," & _
                                          "C.PaperConsumptionsheets,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT LTRIM(eMail) FROM AccountMaster WHERE CODE=P.BookPrinter) As EMailId,M1.Narration,P.BookPrinter,PlateMaker,PaperWastageMin,PaperRate,PaperAmount,RBillAmount,RAdjustment,[RVAT%],RVAT,C.Processing,P.EstQty01 As FinalQuantity,P.ProfitMargin,C.Ref,C.[TotalForms-¼],C.[TotalForms-½],C.[TotalForms-1-F&B],C.[TotalForms-1-W&T],C.[TotalPlates-¼],C.[TotalPlates-½],C.[TotalPlates-1-F&B],C.[TotalPlates-1-W&T],C.[RevisedPlates]  " & _
                                          "FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON C.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.BookPrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
        If rstBookPOChild05.RecordCount = 0 Then
            rptBookPrintOrder02.Section13.Suppress = True
                Else
            With rstBookPOChild05
                .MoveFirst
                Do While Not .EOF
                    TotalAmount = TotalAmount + IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("PBillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value)
                    TotalTax = TotalTax + IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("VAT").Value + rstBookPOChild05.Fields("RVAT").Value, rstBookPOChild05.Fields("VAT").Value + rstBookPOChild05.Fields("PVAT").Value + rstBookPOChild05.Fields("RVAT").Value)
                    TaxableAmount = (TotalAmount - TotalTax)
                    UnitRateBT = TaxableAmount / Val(rstBookPOChild05.Fields("FinalQuantity").Value)
                    UnitRate = TotalAmount / Val(rstBookPOChild05.Fields("FinalQuantity").Value)
                    PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild05.Fields("ProfitMargin").Value) / 100))
                    PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild05.Fields("ProfitMargin").Value) / 100))
                    PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild05.Fields("FinalQuantity").Value)
                    PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild05.Fields("FinalQuantity").Value)
                    TotalPages = TotalPages + Val(rstBookPOChild05.Fields("Pages").Value)
                    TotalForme = TotalForme + Val(rstBookPOChild05.Fields("Forme").Value)
                    .MoveNext
                Loop
            End With
            rstBookPOChild05.MoveFirst
            If Not HeaderPrinted Then
                    rptBookPrintOrder02.Text7.SetText rstBookPOChild05.Fields("OrderNo").Value
                    rptBookPrintOrder02.Text43.SetText rstBookPOChild05.Fields("Ref").Value
                    rptBookPrintOrder02.Text42.SetText Format(rstBookPOChild05.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptBookPrintOrder02.Text12.SetText rstBookPOChild05.Fields("TextPrinter").Value
                    rptBookPrintOrder02.Text13.SetText rstBookPOChild05.Fields("Item").Value
                    rptBookPrintOrder02.Text14.SetText rstBookPOChild05.Fields("FinishSize").Value
                    rptBookPrintOrder02.Text64.SetText "(" & TotalPages & " Pages/" & TotalForme & " f )"
                    rptBookPrintOrder02.Text59.SetText "Pages/Forms:"
                    rptBookPrintOrder02.Text15.SetText rstBookPOChild05.Fields("FinalQuantity").Value
                    rptBookPrintOrder02.Subreport7.OpenSubreport.Text2.SetText " (" & LTrim(NumberToWords(IIf(rstBookPOChild05.Fields("BookPrinter").Value <> rstBookPOChild05.Fields("PlateMaker").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value, rstBookPOChild05.Fields("BillAmount").Value + rstBookPOChild05.Fields("PBillAmount").Value + rstBookPOChild05.Fields("RBillAmount").Value), True)) & ")"
                    EMailID = rstBookPOChild05.Fields("EMailId").Value
                    OrderNo = rstBookPOChild05.Fields("OrderNo").Value
                    ItemName = rstBookPOChild05.Fields("Item").Value
                    Attachment = LTrim(rstBookPOChild05.Fields("OrderNo").Value)
                HeaderPrinted = True
            End If
        End If
    End If
    If OrderType = "TP" Or OrderType = "ALL" Then 'Title or All Printing
        If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
                rstBookPOChild06.Open "SELECT (SELECT PrintName FROM ElementMaster WHERE Code=C.Element) As Element,C.Pages," & _
                                                                   "LTRIM(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTRIM(C.Sets) As Forme,LTRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As CoverPrinter,C.ActualQuantity,C.BillingQuantity,C.FrontPrintingType,C.BackPrintingType,Imposition,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],(SELECT PrintName FROM GeneralMaster WHERE Code=C.PlateType) As Plate,(SELECT PrintName FROM GeneralMaster WHERE Code=C.PlateTypeBack) As PlateBack,C.TotalPlates,C.TotalPlatesBack," & _
                                                                   "PrintRate,PrintRateBack,PrintAmount,PlateRate,PlateRateBack,PlateAmount,(SELECT LTRIM(M3.PrintName)+' (UOM : '+LTRIM(M4.PrintName)+'='+LTRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C.Paper) As PaperName," & _
                                                                   "[PaperWastage%],[PaperWastage%Back],PaperWastageFinal,PaperConsumptionOther,PaperConsumptionKg,Adjustment,PAdjustment,[VAT%],VAT,[PVAT%],PVAT,BillAmount,PBillAmount,C.Remarks,(SELECT LTRIM(eMail) FROM AccountMaster WHERE Code=P.TitlePrinter) As EMailId,M1.Narration,P.TitlePrinter,PlateMaker,PaperWastageMin,PaperWastageMinBack,PaperRate,PaperAmount,RBillAmount,RAdjustment,[RVAT%],RVAT,C.Processing,C.ProcessingBack,C.[Ups] As Ups,P.EstQty01 As FinalQuantity,P.ProfitMargin,C.Sets,(SELECT '('+LTRIM(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOM,C.Ref,P.Code As VchCode,C.PaperConsumptionsheets,(SELECT LTrim(M7.[Weight/Unit]) FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOMwt,(SELECT LTrim(M6.[Value1]) FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C.Paper) As UOMqty,PicData,PicType " & _
                                                                   "FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON C.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.TitlePrinter='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
        If rstBookPOChild06.RecordCount = 0 Then
            rptBookPrintOrder02.Section4.Suppress = True
                Else
            With rstBookPOChild06
                .MoveFirst
                BillAmount = 0
                Do While Not .EOF
                    BillAmount = BillAmount + Val(.Fields("BillAmount").Value)
                    PBillAmount = PBillAmount + Val(.Fields("PBillAmount").Value)
                    RBillAmount = RBillAmount + Val(.Fields("RBillAmount").Value)
                    TotalAmount = TotalAmount + IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value, rstBookPOChild06.Fields("BillAmount").Value + rstBookPOChild06.Fields("PBillAmount").Value + rstBookPOChild06.Fields("RBillAmount").Value)
                    TotalTax = TotalTax + IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, rstBookPOChild06.Fields("VAT").Value + rstBookPOChild06.Fields("RVAT").Value, rstBookPOChild06.Fields("VAT").Value + rstBookPOChild06.Fields("PVAT").Value + rstBookPOChild06.Fields("RVAT").Value)
                    TaxableAmount = (TotalAmount - TotalTax)
                    UnitRateBT = TaxableAmount / Val(rstBookPOChild06.Fields("FinalQuantity").Value)
                    UnitRate = TotalAmount / Val(rstBookPOChild06.Fields("FinalQuantity").Value)
                    PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild06.Fields("ProfitMargin").Value) / 100))
                    PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild06.Fields("ProfitMargin").Value) / 100))
                    PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild05.Fields("FinalQuantity").Value)
                    PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild06.Fields("FinalQuantity").Value)
                    TotalPages = TotalPages + Val(rstBookPOChild06.Fields("Pages").Value)
                    TotalForme = TotalForme + Val(rstBookPOChild06.Fields("Sets").Value)
                    .MoveNext
                Loop
            End With
                rstBookPOChild06.MoveFirst
                If Not HeaderPrinted Then
                    rptBookPrintOrder02.Text7.SetText rstBookPOChild06.Fields("OrderNo").Value
                    rptBookPrintOrder02.Text43.SetText rstBookPOChild06.Fields("Ref").Value
                    rptBookPrintOrder02.Text42.SetText Format(rstBookPOChild06.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptBookPrintOrder02.Text12.SetText rstBookPOChild06.Fields("CoverPrinter").Value
                    rptBookPrintOrder02.Text13.SetText rstBookPOChild06.Fields("Item").Value
                    rptBookPrintOrder02.Text14.SetText rstBookPOChild06.Fields("FinishSize").Value
                    rptBookPrintOrder02.Text64.SetText "(" & TotalPages & " Pages/" & TotalForme & " Sets )"
                    rptBookPrintOrder02.Text59.SetText "Pages/Sets:"
                    rptBookPrintOrder02.Text15.SetText rstBookPOChild06.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild06.Fields("EMailId").Value
                    OrderNo = rstBookPOChild06.Fields("OrderNo").Value
                    ItemName = rstBookPOChild06.Fields("Item").Value
                    Attachment = LTrim(rstBookPOChild06.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                If Not CheckEmpty(rstBookPOChild06.Fields("PicData").Value, False) Then
                    If Dir(App.Path & "\Pic\", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Pic\"
                    If Dir(App.Path & "\Pic\" & FinancialYear & CompCode, vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Pic\" & FinancialYear & CompCode
                    imgFile = App.Path & "\Pic\" & FinancialYear & CompCode & "\" & FinancialYear & CompCode & rstBookPOChild06.Fields("VchCode").Value & "." & rstBookPOChild06.Fields("PicType").Value: RetrievePic rstBookPOChild06.Fields("PicData").Value, imgFile, srmPicMgr
                    rptBookPrintOrder02.picFile = imgFile
                End If
                rptBookPrintOrder02.Subreport8.OpenSubreport.Text13.SetText " (" & LTrim(NumberToWords(IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, Round(BillAmount, 0) + Round(RBillAmount, 0), Round(BillAmount, 0) + Round(PBillAmount, 0) + Round(RBillAmount, 0)), True)) & ")"
        End If
    End If
    If OrderType = "CB" Or OrderType = "ALL" Then
                If rstBookPOChild10.State = adStateOpen Then rstBookPOChild10.Close
                rstBookPOChild10.Open "SELECT LTRIM(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As CoverPrinter,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMakers,CHOOSE(CONVERT(NUMERIC,C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalFormsFront,C1.TotalFormsBack,C1.PrintRateFront,C1.PrintRateBack,C1.[GST%],C1.GST,C1.Adjustment,C1.PrintAmount,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,C1.[PGST%],C1.PGST,PlateAmount," & _
                          "(SELECT LTRIM(M3.PrintName)FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C1.Paper) As PaperName,C1.[PaperWastage%],C1.[PaperConsumptionOther],C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],LTRIM(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,C2.ActualQuantity,[Ups/Plate],C2.PrintingQuantity,C2.BillingQuantity,LTRIM(C2.FrontPrintingColor)+'+'+LTRIM(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType,P.TitlePrinter AS TitlePrinterCode,C1.PlateMaker As PlateMakerCode," & _
                          "PaperWastageMin,PaperRate,PaperAmountBT,RAdjustment,[RGST%],RGST,PaperAmount,PrintAmountBT,PlateAmountBT,(SELECT EMail FROM AccountMaster WHERE Code=P.TitlePrinter) As EMail,Imposition,PlateType,(SELECT '('+LTRIM(M6.PrintName)+')' FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C1.Paper) As UOM,(SELECT LTRIM(MAX(PrintingQuantity)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxPrintingQty, " & _
                          "P.EstQty01 As FinalQuantity,P.ProfitMargin,(SELECT LTRIM(MAX(FrontPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxFrontColor,(SELECT LTRIM(MAX(BackPrintingColor)) FROM BookPOChild0901 WHERE Code=P.Code) As MaxBackColor,C1.Calculation,C1.PaperConsumptionsheets,(SELECT LTrim(M7.[Weight/Unit]) FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C1.Paper) As UOMwt,(SELECT LTrim(M6.[Value1]) FROM PaperMaster M7 INNER JOIN GeneralMaster M6 ON M7.UOM=M6.Code WHERE M7.Code=C1.Paper) As UOMqty,(Select Name From BookMaster Where Code=P.Book) As oItem FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
        If rstBookPOChild10.RecordCount = 0 Then
                    rptBookPrintOrder02.Section16.Suppress = True
                Else
            rstBookPOChild10.MoveFirst
                    TotalTax = TotalTax + IIf(rstBookPOChild10.Fields("TitlePrinterCode").Value <> rstBookPOChild10.Fields("PlateMakerCode").Value, rstBookPOChild10.Fields("GST").Value + rstBookPOChild10.Fields("RGST").Value, rstBookPOChild10.Fields("GST").Value + rstBookPOChild10.Fields("PGST").Value + rstBookPOChild10.Fields("RGST").Value)
                    TotalAmount = TotalAmount + IIf(rstBookPOChild10.Fields("TitlePrinterCode").Value <> rstBookPOChild10.Fields("PlateMakerCode").Value, rstBookPOChild10.Fields("PrintAmount").Value + rstBookPOChild10.Fields("PaperAmount").Value, rstBookPOChild10.Fields("PrintAmount").Value + rstBookPOChild10.Fields("PlateAmount").Value + rstBookPOChild10.Fields("PaperAmount").Value)
                    TaxableAmount = (TotalAmount - TotalTax)
                    UnitRateBT = TaxableAmount / Val(rstBookPOChild10.Fields("FinalQuantity").Value)
                    UnitRate = TotalAmount / Val(rstBookPOChild10.Fields("FinalQuantity").Value)
                    PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild10.Fields("ProfitMargin").Value) / 100))
                    PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild10.Fields("ProfitMargin").Value) / 100))
                    PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild10.Fields("FinalQuantity").Value)
                    PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild10.Fields("FinalQuantity").Value)
                If Not HeaderPrinted Then
                    rptBookPrintOrder02.Text7.SetText rstBookPOChild10.Fields("OrderNo").Value
                    rptBookPrintOrder02.Text42.SetText Format(rstBookPOChild10.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptBookPrintOrder02.Text12.SetText rstBookPOChild10.Fields("CoverPrinter").Value
                    rptBookPrintOrder02.Text13.SetText rstBookPOChild10.Fields("oItem").Value
'                    rptBookPrintOrder02.Text14.SetText rstBookPOChild10.Fields("FinishSize").Value
                    rptBookPrintOrder02.Text59.SetText ""
'                    rptBookPrintOrder02.Text64.SetText rstBookPOChild10.Fields("Forme").Value
                    rptBookPrintOrder02.Text9.SetText ""
                    rptBookPrintOrder02.Text15.SetText rstBookPOChild10.Fields("FinalQuantity").Value     'IIf(OrderType = "CB", (Val(rstBookPOChild10.Fields("MaxPrintingQty").Value)), Val(rstBookPOChild10.Fields("FinalQuantity").Value))
                    EMailID = rstBookPOChild10.Fields("EMailId").Value
                    OrderNo = rstBookPOChild10.Fields("OrderNo").Value
                    ItemName = rstBookPOChild10.Fields("Item").Value
                    Attachment = LTrim(rstBookPOChild10.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                    If rstBookPOChild10.Fields("Calculation").Value = "S" Then
                    rptBookPrintOrder02.Subreport6_Text66.SetText "Combo Format Printing Details (As per Single Set Calculation)"
                    Else
                    rptBookPrintOrder02.Subreport6_Text66.SetText "Combo Format Printing Details (As per Individual Set Calculation)"
                    End If
                    rptBookPrintOrder02.Subreport6.OpenSubreport.Text25.SetText " (" & LTrim(NumberToWords(IIf(rstBookPOChild10.Fields("TitlePrinterCode").Value <> rstBookPOChild10.Fields("PlateMakerCode").Value, rstBookPOChild10.Fields("BillAmount").Value + rstBookPOChild10.Fields("RBillAmount").Value, rstBookPOChild10.Fields("BillAmount").Value + rstBookPOChild10.Fields("PBillAmount").Value + rstBookPOChild10.Fields("RBillAmount").Value), True)) & ")"
        End If
    End If
    If OrderType = "TL" Or OrderType = "ALL" Then
        If rstBookPOChild07.State = adStateOpen Then rstBookPOChild07.Close
            rstBookPOChild07.Open "SELECT E.Name As Element,O.Name As Operation,[Number],OS.Name As [Size],Quantity,M.Name As CalcMode,Rate,Amount,Adjustment,[GST%],GST,BillAmount,C.Remarks,LTRIM(I.PrintName)+IIF(I.Price=0,'',' (Price : Rs. '+Format(I.Price,'0.00')+')') As Item,FS.PrintName As FinishSize,LTRIM(LTRIM(I.Pages)+'-pages/'+LTRIM(I.Forms)+'f('+IIF(I.OneColorForms=0,'',' 1Col-'+LTRIM(I.OneColorForms)+'f ')+IIF(I.TwoColorForms=0,'',' 2Col-'+LTRIM(I.TwoColorForms)+'f ')+IIF(I.FourColorForms=0,'',' 4Col-'+LTRIM(I.FourColorForms)))+'f)' As Forme,LTRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.Laminator) As Laminator,(SELECT eMail FROM AccountMaster WHERE Code=P.Laminator) As EMailId,I.Narration,P.EstQty01 As FinalQuantity,P.ProfitMargin " & _
                                                               "FROM ((((((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster FS ON I.FinishSize=FS.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code) LEFT JOIN GeneralMaster OS ON C.[Size]=OS.Code WHERE P.Code='" & OrderCode & "' AND P.Laminator='" & rstBookPOChild0801.Fields("Vendor").Value & "' ORDER BY E.Name,O.Name", cnDatabase, adOpenKeyset, adLockOptimistic
        If rstBookPOChild07.RecordCount = 0 Then
                rptBookPrintOrder02.Section20.Suppress = True
            Else
                With rstBookPOChild07
                    .MoveFirst
                    BillAmount = 0
                    Do While Not .EOF
                        BillAmount = BillAmount + Val(.Fields("BillAmount").Value)
                        TotalAmount = TotalAmount + Val(.Fields("BillAmount").Value)
                        TotalTax = TotalTax + Val(.Fields("GST").Value)
                        TaxableAmount = (TotalAmount - TotalTax)
                        UnitRateBT = TaxableAmount / Val(rstBookPOChild07.Fields("FinalQuantity").Value)
                        UnitRate = TotalAmount / Val(rstBookPOChild07.Fields("FinalQuantity").Value)
                        
                PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild07.Fields("ProfitMargin").Value) / 100))
                PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild07.Fields("ProfitMargin").Value) / 100))
                PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild07.Fields("FinalQuantity").Value)
                PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild07.Fields("FinalQuantity").Value)
                        .MoveNext
                    Loop
                End With
                rstBookPOChild07.MoveFirst
                If Not HeaderPrinted Then
                    rptBookPrintOrder02.Text7.SetText rstBookPOChild07.Fields("OrderNo").Value
                    rptBookPrintOrder02.Text42.SetText Format(rstBookPOChild07.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptBookPrintOrder02.Text12.SetText rstBookPOChild07.Fields("Laminator").Value
                    rptBookPrintOrder02.Text13.SetText rstBookPOChild07.Fields("Item").Value
                    rptBookPrintOrder02.Text14.SetText rstBookPOChild07.Fields("FinishSize").Value
                    rptBookPrintOrder02.Text64.SetText rstBookPOChild07.Fields("Forme").Value
                    rptBookPrintOrder02.Text15.SetText rstBookPOChild07.Fields("FinalQuantity").Value
                    EMailID = rstBookPOChild07.Fields("EMailId").Value
                    OrderNo = rstBookPOChild07.Fields("OrderNo").Value
                    ItemName = rstBookPOChild07.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild07.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                rptBookPrintOrder02.Subreport4.OpenSubreport.Text25.SetText "Amount Payable : " & Trim(NumberToWords(Round(BillAmount, 0), True))
            End If
        End If
        If OrderType = "BB" Or OrderType = "ALL" Then
            If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
                rstBookPOChild08.Open "SELECT E.Name As Element,O.Name As Operation,[Number],OS.Name As [Size],Quantity,M.Name As CalcMode,Rate,Amount,Adjustment,[GST%],GST,BillAmount,C.Remarks,LTRIM(I.PrintName)+IIF(I.Price=0,'',' (Price : Rs. '+Format(I.Price,'0.00')+')') As Item,FS.PrintName As FinishSize,LTRIM(LTRIM(I.Pages)+'-pages/'+LTRIM(I.Forms)+'f('+IIF(I.OneColorForms=0,'',' 1Col-'+LTRIM(I.OneColorForms)+'f ')+IIF(I.TwoColorForms=0,'',' 2Col-'+LTRIM(I.TwoColorForms)+'f ')+IIF(I.FourColorForms=0,'',' 4Col-'+LTRIM(I.FourColorForms)))+'f)' As Forme,LTRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.Binder) As Binder,(SELECT eMail FROM AccountMaster WHERE Code=P.Binder) As EMailId,I.Narration,P.EstQty01 As FinalQuantity,P.ProfitMargin " & _
                                                                   "FROM ((((((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster FS ON I.FinishSize=FS.Code) INNER JOIN ElementMaster E ON C.ElementGroup=E.Code) INNER JOIN GeneralMaster O ON C.BinderyProcess=O.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code) LEFT JOIN GeneralMaster OS ON C.[Size]=OS.Code WHERE P.Code='" & OrderCode & "' AND P.Binder='" & rstBookPOChild0801.Fields("Vendor").Value & "' ORDER BY E.Name,O.Name", cnDatabase, adOpenKeyset, adLockOptimistic
            If rstBookPOChild08.RecordCount = 0 Then
                    rptBookPrintOrder02.Section15.Suppress = True
                Else
                    With rstBookPOChild08
                        .MoveFirst
                        BillAmount = 0
                        Do While Not .EOF
                            BillAmount = BillAmount + Val(.Fields("BillAmount").Value)
                            TotalAmount = TotalAmount + Val(.Fields("BillAmount").Value)
                            TotalTax = TotalTax + Val(.Fields("GST").Value)
                            TaxableAmount = (TotalAmount - TotalTax)
                            UnitRateBT = TaxableAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
                            UnitRate = TotalAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
                            
                    PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild08.Fields("ProfitMargin").Value) / 100))
                    PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild08.Fields("ProfitMargin").Value) / 100))
                    PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
                    PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
                            .MoveNext
                        Loop
                    End With
                    rstBookPOChild08.MoveFirst
                    If Not HeaderPrinted Then
                        rptBookPrintOrder02.Text7.SetText rstBookPOChild08.Fields("OrderNo").Value
                        rptBookPrintOrder02.Text42.SetText Format(rstBookPOChild08.Fields("OrderDate").Value, "dd-MM-yyyy")
                        rptBookPrintOrder02.Text12.SetText rstBookPOChild08.Fields("Binder").Value
                        rptBookPrintOrder02.Text13.SetText rstBookPOChild08.Fields("Item").Value
                        rptBookPrintOrder02.Text14.SetText rstBookPOChild08.Fields("FinishSize").Value
                        rptBookPrintOrder02.Text64.SetText rstBookPOChild08.Fields("Forme").Value
                        rptBookPrintOrder02.Text15.SetText rstBookPOChild08.Fields("FinalQuantity").Value
                        EMailID = rstBookPOChild08.Fields("EMailId").Value
                        OrderNo = rstBookPOChild08.Fields("OrderNo").Value
                        ItemName = rstBookPOChild08.Fields("Item").Value
                        Attachment = Trim(rstBookPOChild08.Fields("OrderNo").Value)
                        HeaderPrinted = True
                    End If
                    rptBookPrintOrder02.Subreport9.OpenSubreport.Text25.SetText "Amount Payable : " & Trim(NumberToWords(Round(BillAmount, 0), True))
                End If
            End If
'            If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
'            rstBookPOChild08.Open "SELECT LTRIM(M1.PrintName)+IIF(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTRIM(LTRIM(M1.Pages)+'-pages/'+LTRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1Col-'+LTRIM(M1.OneColorForms)+'f ')+IIF(M1.TwoColorForms=0,'',' 2Col-'+LTRIM(M1.TwoColorForms)+'f ')+IIF(M1.FourColorForms=0,'',' 4Col-'+LTRIM(M1.FourColorForms)))+'f)' As Forme,LTRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=P.Binder) As Binder,C.ActualQuantity, " & _
'                                  "C.BillingQuantity,(SELECT PrintName FROM GeneralMaster WHERE Code=C.BindingType) As BindingType,BindingForms,ExtraForms,FormFoldRate,FormPasteRate,FormStitchRate,[Rate/Book],TotalPkts,PktPackRate,TotalBoxes,BoxPackRate,CartageRate,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,(SELECT eMail FROM AccountMaster WHERE Code=P.Binder) As EMailId,M1.Narration,P.EstQty01 As FinalQuantity,P.ProfitMargin " & _
'                                  "FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND P.Binder='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
'            If rstBookPOChild08.RecordCount = 0 Then
'                rptBookPrintOrder02.Section11.Suppress = True
'            Else
'                rstBookPOChild08.MoveFirst
'                TotalAmount = TotalAmount + Val(rstBookPOChild08.Fields("BillAmount").Value)
'                TotalTax = TotalTax + Val(rstBookPOChild08.Fields("VAT").Value)
'                TaxableAmount = (TotalAmount - TotalTax)
'                UnitRateBT = TaxableAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
'                UnitRate = TotalAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
'
'                PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild08.Fields("ProfitMargin").Value) / 100))
'                PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild08.Fields("ProfitMargin").Value) / 100))
'                PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
'                PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild08.Fields("FinalQuantity").Value)
'
'                If Not HeaderPrinted Then
'                    rptBookPrintOrder02.Text7.SetText rstBookPOChild08.Fields("OrderNo").Value
'                    rptBookPrintOrder02.Text42.SetText rstBookPOChild08.Fields("OrderDate").Value
'                    rptBookPrintOrder02.Text12.SetText rstBookPOChild08.Fields("Binder").Value
'                    rptBookPrintOrder02.Text13.SetText rstBookPOChild08.Fields("Item").Value
'                    rptBookPrintOrder02.Text14.SetText rstBookPOChild08.Fields("FinishSize").Value
'                    rptBookPrintOrder02.Text64.SetText rstBookPOChild08.Fields("Forme").Value
'                    rptBookPrintOrder02.Text15.SetText rstBookPOChild08.Fields("FinalQuantity").Value
'                    EMailID = rstBookPOChild08.Fields("EMailId").Value
'                    OrderNo = rstBookPOChild08.Fields("OrderNo").Value
'                    ItemName = rstBookPOChild08.Fields("Item").Value
'                    Attachment = LTrim(rstBookPOChild08.Fields("OrderNo").Value)
'                End If
'                rptBookPrintOrder02.Subreport3.OpenSubreport.Text25.SetText " (" & LTrim(NumberToWords(Round(rstBookPOChild08.Fields("BillAmount").Value, 0), True)) & ")"
'            End If
'        End If
        If rstBookPOChild09.State = adStateOpen Then rstBookPOChild09.Close
        If DatabaseType = "MS SQL" Then
        rstBookPOChild09.Open "SELECT Choose(LTrim(Category),(SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item),(SELECT LTRIM(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='R' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item)) As ItemName,[Consumption/Item],OrderQuantity,TotalConsumption,Rate,Amount,(Select EstQty01 from BookPoParent where Code=T.Code)*1 as FinalQuantity,(Select ProfitMargin from BookPoParent where Code=T.Code)*1 as ProfitMargin " & _
                              "FROM BookPOChild0801 T WHERE T.Code='" & OrderCode & "' AND T.Vendor='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
                                '+' (UOM : '+LTRIM(M4.PrintName)+'='+LTRIM(M4.Value1)+')'       +' (UOM : '+LTRIM(M4.PrintName)+'='+LTRIM(M4.Value1)+')'
        Else
        rstBookPOChild09.Open "SELECT Choose(Val(Category),(SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item),(SELECT LTRIM(M3.PrintName) FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='R' AND Code=T.Item),(SELECT Name FROM BookMaster WHERE LEFT(Type,1)='F' AND Code=T.Item)) As ItemName,[Consumption/Item],OrderQuantity,TotalConsumption,Rate,Amount,(Select EstQty01 from BookPoParent where Code=T.Code)*1 as FinalQuantity,(Select ProfitMargin from BookPoParent where Code=T.Code)*1 as ProfitMargin " & _
                              "FROM BookPOChild0801 T WHERE T.Code='" & OrderCode & "' AND T.Vendor='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
        End If
        If rstBookPOChild09.RecordCount = 0 Then rptBookPrintOrder02.Section12.Suppress = True
        
        If rstBookPOChild09.RecordCount = 0 Then
        Else
                With rstBookPOChild09
            .MoveFirst
            
                BOMAmount = 0
                Do While Not .EOF
                BOMAmount = BOMAmount + rstBookPOChild09.Fields("Amount").Value
                'PBillAmount = PBillAmount
                'RBillAmount = RBillAmount
                TotalAmount = TotalAmount + rstBookPOChild09.Fields("Amount").Value
                'TotalTax = TotalTax
                TaxableAmount = (TotalAmount - TotalTax)
                UnitRateBT = TaxableAmount / Val(rstBookPOChild09.Fields("FinalQuantity").Value)
                UnitRate = TotalAmount / Val(rstBookPOChild09.Fields("FinalQuantity").Value)
                PMTaxableAmount = TaxableAmount + (TaxableAmount * (Val(rstBookPOChild09.Fields("ProfitMargin").Value) / 100))
                PMTotalTax = TotalTax + (TotalTax * (Val(rstBookPOChild09.Fields("ProfitMargin").Value) / 100))
                PMUnitRateBT = PMTaxableAmount / Val(rstBookPOChild09.Fields("FinalQuantity").Value)
                PMUnitRate = (PMTaxableAmount + PMTotalTax) / Val(rstBookPOChild09.Fields("FinalQuantity").Value)
                    
            .MoveNext
                Loop
                End With
        End If
        
        Screen.MousePointer = vbNormal
        If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,EMail,Website FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
        rptBookPrintOrder02.Text2.SetText LTrim(rstCompanyMaster.Fields("PrintName").Value)
        rptBookPrintOrder02.Text3.SetText LTrim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & LTrim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & LTrim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & LTrim(rstCompanyMaster.Fields("Address4").Value)
        rptBookPrintOrder02.Text24.SetText "Phone : " & LTrim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & LTrim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & LTrim(rstCompanyMaster.Fields("EMail").Value)
        rptBookPrintOrder02.Text27.SetText "for " & rptBookPrintOrder02.Text12.Text
        rptBookPrintOrder02.Text28.SetText "for " & LTrim(rstCompanyMaster.Fields("PrintName").Value)
        If TotalAmount = 0 Then
'            rptBookPrintOrder02.Section14.Suppress = True
        Else
            rptBookPrintOrder02.Text17.SetText Format(TotalTax, "##0.00")
            rptBookPrintOrder02.Text18.SetText Format(TotalAmount, "##0.00")
            rptBookPrintOrder02.Text22.SetText Format(TaxableAmount, "##0.00")
            rptBookPrintOrder02.Text29.SetText Format(UnitRateBT, "##0.000")
            rptBookPrintOrder02.Text33.SetText Format(UnitRate, "##0.000")
            
            rptBookPrintOrder02.Text37.SetText Format(PMTotalTax, "##0.00")
            rptBookPrintOrder02.Text18.SetText Format(TotalAmount, "##0.00")
            rptBookPrintOrder02.Text34.SetText Format(PMTaxableAmount, "##0.00")
            rptBookPrintOrder02.Text38.SetText Format(PMUnitRateBT, "##0.000")
            rptBookPrintOrder02.Text40.SetText Format(PMUnitRate, "#0.000")
            rptBookPrintOrder02.Text19.SetText " (" & LTrim(NumberToWords(Round(TotalAmount, 0), True)) & ")"
        End If
        If rstBookPOChild05.Fields("ProfitMargin").Value = 0 Then
        rptBookPrintOrder02.Section17.Suppress = True
        Else
        rptBookPrintOrder02.Text44.SetText Format(rstBookPOChild05.Fields("ProfitMargin").Value / 100, "##0.00 %")
        End If
        rptBookPrintOrder02.Subreport1.OpenSubreport.Database.SetDataSource rstBookPOChild05, 3, 1
        rptBookPrintOrder02.Subreport2.OpenSubreport.Database.SetDataSource rstBookPOChild06, 3, 1
        rptBookPrintOrder02.Subreport4.OpenSubreport.Database.SetDataSource rstBookPOChild07, 3, 1
        rptBookPrintOrder02.Subreport3.OpenSubreport.Database.SetDataSource rstBookPOChild08, 3, 1
        rptBookPrintOrder02.Subreport5.OpenSubreport.Database.SetDataSource rstBookPOChild09, 3, 1
        rptBookPrintOrder02.Subreport6.OpenSubreport.Database.SetDataSource rstBookPOChild10, 3, 1
        rptBookPrintOrder02.Subreport7.OpenSubreport.Database.SetDataSource rstBookPOChild05, 3, 1
        rptBookPrintOrder02.Subreport8.OpenSubreport.Database.SetDataSource rstBookPOChild06, 3, 1
        rptBookPrintOrder02.Subreport9.OpenSubreport.Database.SetDataSource rstBookPOChild08, 3, 1
        
        Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
        Message = "Dear Sir,<Br>Please find attached herewith PO #" & LTrim(rstBookPOChild08.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & LTrim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & LTrim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & LTrim(rstCompanyMaster.Fields("EMail").Value) & "'>" & LTrim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
        If OutputTo = "S" Then
            FrmReportViewer.EMailID = EMailID
            FrmReportViewer.Subject = "Book Order #" & LTrim(OrderNo) + " Book : " + LTrim(ItemName)
            FrmReportViewer.Attachment = Attachment
            FrmReportViewer.Message = Message
            Set FrmReportViewer.Report = rptBookPrintOrder02
            FrmReportViewer.Show vbModal
        Else
            If rstBookPOList.State = adStateClosed Then
                If EMailID = "" Or OutputType = "P" Then
                    rptBookPrintOrder02.PaperSource = crPRBinAuto
                    rptBookPrintOrder02.PrintOut False   ' Print Report Without Prompt
                Else
                    rptBookPrintOrder02.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                    rptBookPrintOrder02.ExportOptions.DestinationType = crEDTDiskFile
                    rptBookPrintOrder02.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                    rptBookPrintOrder02.Export False
                    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                    With oOutlookMsg
                        .To = EMailID
                        .Subject = "Book Order #" & LTrim(OrderNo) + " Book : " + LTrim(ItemName)
                        .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                        .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                        .Importance = olImportanceHigh
                        .ReadReceiptRequested = True
                        .Send
                        If Err.Number = 0 Then cnDatabase.Execute "UPDATE BookPOParent SET BBODStatus=1 WHERE Code='" & OrderCode & "'"
                    End With
                    Set oOutlookMsg = Nothing
                End If
            Else
                rptBookPrintOrder02.PaperSource = crPRBinAuto
                rptBookPrintOrder02.PrintOut
            End If
        End If
        Set rptBookPrintOrder02 = Nothing
        rstBookPOChild0801.MoveNext
        If OrderType = "BP" Or OrderType = "TP" Or OrderType = "CB" Or OrderType = "TL" Or OrderType = "BB" Then Exit Do
    Loop
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild05): Call CloseRecordset(rstBookPOChild06): Call CloseRecordset(rstBookPOChild07): Call CloseRecordset(rstBookPOChild08): Call CloseRecordset(rstBookPOChild0801): Call CloseRecordset(rstBookPOChild09)
    On Error GoTo 0
    Screen.MousePointer = vbNormal
End Sub
Public Sub PrintBookPrintOrder03(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String, Optional ByVal OrderType As String, Optional ByVal BookPOType As String)
    Dim oOutlookMsg As Outlook.MailItem, HeaderPrinted As Boolean, OrderNo As String, ItemName As String, TotalTax As Double, TotalAmount As Double
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If OrderType = "BP" Then
        rptBookPlateOrder.Section4.Suppress = True
    ElseIf OrderType = "TP" Then
        rptBookPlateOrder.Section13.Suppress = True
    End If
    If rstBookPOChild0801.State = adStateOpen Then rstBookPOChild0801.Close
    If OrderType = "BP" Then
        rstBookPOChild0801.Open "SELECT PlateMaker As Vendor FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf OrderType = "TP" Then
        rstBookPOChild0801.Open "SELECT PlateMaker As Vendor FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstBookPOChild0801.Open "SELECT PlateMaker As Vendor FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' UNION SELECT PlateMaker As Vendor FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
    If rstBookPOChild0801.RecordCount > 0 Then rstBookPOChild0801.MoveFirst
    Do While Not rstBookPOChild0801.EOF
        HeaderPrinted = False
        If OrderType = "BP" Or OrderType = "ALL" Then
            If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    If DatabaseType = "MS SQL" Then
            rstBookPOChild05.Open "SELECT LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'p/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+LTrim(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+LTrim(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+LTrim(M1.FourColorForms)))+')' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=C.PlateMaker) As PlateMaker,M1.DuplexPrinting,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size1) As [Size1],Pages1,[TotalPlates1-¼],[TotalPlates1-½],[TotalPlates1-1],CHOOSE(LTrim(PlateType1),'Deep-etch','PS','Wipe-on','CTP') As Plate1,PlateRate1,PlateAmount1," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size2) As [Size2],Pages2,[TotalPlates2-¼],[TotalPlates2-½],[TotalPlates2-1],CHOOSE(CONVERT(NUMERIC,PlateType2),'Deep-etch','PS','Wipe-on','CTP') As Plate2,PlateRate2,PlateAmount2,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size4) As [Size4],Pages4,[TotalPlates4-¼],[TotalPlates4-½],[TotalPlates4-1],CHOOSE(CONVERT(NUMERIC,PlateType4),'Deep-etch','PS','Wipe-on','CTP') As Plate4,PlateRate4,PlateAmount4,PAdjustment,[PVAT%],PVAT,C.Remarks,(SELECT LTrim(eMail) FROM AccountMaster WHERE CODE=C.PlateMaker) As EMailId,M1.Narration " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND C.PlateMaker='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
            rstBookPOChild05.Open "SELECT Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,TRIM(TRIM(M1.Pages)+'p/'+TRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+TRIM(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+TRIM(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+TRIM(M1.FourColorForms)))+')' As Forme,TRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=C.PlateMaker) As PlateMaker,M1.DuplexPrinting,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size1) As [Size1],Pages1,[TotalPlates1-¼],[TotalPlates1-½],[TotalPlates1-1],CHOOSE(VAL(PlateType1),'Deep-etch','PS','Wipe-on','CTP') As Plate1,PlateRate1,PlateAmount1," & _
                                  "(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size2) As [Size2],Pages2,[TotalPlates2-¼],[TotalPlates2-½],[TotalPlates2-1],CHOOSE(VAL(PlateType2),'Deep-etch','PS','Wipe-on','CTP') As Plate2,PlateRate2,PlateAmount2,(SELECT PrintName+'/'+CHOOSE(M1.FormType,'08','16','04','12','24','32','64','06') FROM GeneralMaster WHERE Code=C.Size4) As [Size4],Pages4,[TotalPlates4-¼],[TotalPlates4-½],[TotalPlates4-1],CHOOSE(VAL(PlateType4),'Deep-etch','PS','Wipe-on','CTP') As Plate4,PlateRate4,PlateAmount4,PAdjustment,[PVAT%],PVAT,C.Remarks,(SELECT TRIM(eMail) FROM AccountMaster WHERE CODE=C.PlateMaker) As EMailId,M1.Narration " & _
                                  "FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code WHERE P.Code='" & OrderCode & "' AND C.PlateMaker='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
            If rstBookPOChild05.RecordCount = 0 Then
                rptBookPlateOrder.Section13.Suppress = True
            Else
                rstBookPOChild05.MoveFirst
                TotalAmount = TotalAmount + Round(Val(rstBookPOChild05.Fields("PlateAmount1").Value) + Val(rstBookPOChild05.Fields("PlateAmount2").Value) + Val(rstBookPOChild05.Fields("PlateAmount4").Value) + Val(rstBookPOChild05.Fields("PAdjustment").Value) + Val(rstBookPOChild05.Fields("PVAT").Value), 0)
                TotalTax = TotalTax + Val(rstBookPOChild05.Fields("PVAT").Value)
                rptBookPlateOrder.Text7.SetText rstBookPOChild05.Fields("OrderNo").Value
                rptBookPlateOrder.Text42.SetText Format(rstBookPOChild05.Fields("OrderDate").Value, "dd-MM-yyyy")
                rptBookPlateOrder.Text12.SetText rstBookPOChild05.Fields("PlateMaker").Value
                rptBookPlateOrder.Text13.SetText rstBookPOChild05.Fields("Item").Value
                rptBookPlateOrder.Text14.SetText rstBookPOChild05.Fields("FinishSize").Value
                rptBookPlateOrder.Text64.SetText rstBookPOChild05.Fields("Forme").Value
                rptBookPlateOrder.Subreport1.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(Round(Val(rstBookPOChild05.Fields("PlateAmount1").Value) + Val(rstBookPOChild05.Fields("PlateAmount2").Value) + Val(rstBookPOChild05.Fields("PlateAmount4").Value) + Val(rstBookPOChild05.Fields("PAdjustment").Value) + Val(rstBookPOChild05.Fields("PVAT").Value), 0), True)) & ")"
                EMailID = rstBookPOChild05.Fields("EMailId").Value
                OrderNo = rstBookPOChild05.Fields("OrderNo").Value
                ItemName = rstBookPOChild05.Fields("Item").Value
                Attachment = Trim(rstBookPOChild05.Fields("OrderNo").Value)
                HeaderPrinted = True
            End If
        End If
        If OrderType = "TP" Or OrderType = "ALL" Then
            If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
    If DatabaseType = "MS SQL" Then
            rstBookPOChild06.Open "SELECT LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,LTrim(LTrim(M1.Pages)+'p/'+LTrim(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+LTrim(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+LTrim(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+LTrim(M1.FourColorForms)))+')' As Forme,LTrim(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=C.PlateMaker) As PlateMaker,C.FrontPrintingType,C.BackPrintingType,Imposition,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],(SELECT PrintName FROM GeneralMaster WHERE Code=C.PlateType) As Plate,PlateRate,PlateAmount,PAdjustment,[PVAT%],PVAT,C.Remarks,(SELECT LTrim(eMail) FROM AccountMaster WHERE Code=C.PlateMaker) As EMailId,M1.Narration " & _
            "FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code " & _
                                  "WHERE P.Code='" & OrderCode & "' AND C.PlateMaker='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    Else
            rstBookPOChild06.Open "SELECT Trim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,M2.PrintName As FinishSize,TRIM(TRIM(M1.Pages)+'p/'+TRIM(M1.Forms)+'f('+IIF(M1.OneColorForms=0,'','1C-'+TRIM(M1.OneColorForms)+' ')+IIF(M1.TwoColorForms=0,'','2C-'+TRIM(M1.TwoColorForms)+' ')+IIF(M1.FourColorForms=0,'','4C-'+TRIM(M1.FourColorForms)))+')' As Forme,TRIM(P.Name) As OrderNo,P.Date As OrderDate,(SELECT PrintName FROM AccountMaster WHERE Code=C.PlateMaker) As PlateMaker,C.FrontPrintingType,C.BackPrintingType,Imposition,(SELECT PrintName FROM GeneralMaster WHERE Code=C.[Size]) As [Size],CHOOSE(VAL(PlateType),'Deep-etch','PS','Wipe-on','CTP') As Plate,PlateRate,PlateAmount,PAdjustment,[PVAT%],PVAT,C.Remarks,(SELECT TRIM(eMail) FROM AccountMaster WHERE Code=C.PlateMaker) As EMailId,M1.Narration FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.FinishSize=M2.Code " & _
                                  "WHERE P.Code='" & OrderCode & "' AND C.PlateMaker='" & rstBookPOChild0801.Fields("Vendor").Value & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
            If rstBookPOChild06.RecordCount = 0 Then
                rptBookPlateOrder.Section4.Suppress = True
            Else
                rstBookPOChild06.MoveFirst
                TotalAmount = TotalAmount + Round(Val(rstBookPOChild06.Fields("PlateAmount").Value) + Val(rstBookPOChild06.Fields("PAdjustment").Value) + Val(rstBookPOChild06.Fields("PVAT").Value), 0)
                TotalTax = TotalTax + Val(rstBookPOChild06.Fields("PVAT").Value)
                If Not HeaderPrinted Then
                    rptBookPlateOrder.Text7.SetText rstBookPOChild06.Fields("OrderNo").Value
                    rptBookPlateOrder.Text42.SetText Format(rstBookPOChild06.Fields("OrderDate").Value, "dd-MM-yyyy")
                    rptBookPlateOrder.Text12.SetText rstBookPOChild06.Fields("PlateMaker").Value
                    rptBookPlateOrder.Text13.SetText rstBookPOChild06.Fields("Item").Value
                    rptBookPlateOrder.Text14.SetText rstBookPOChild06.Fields("FinishSize").Value
                    rptBookPlateOrder.Text64.SetText rstBookPOChild06.Fields("Forme").Value
                    EMailID = rstBookPOChild06.Fields("EMailId").Value
                    OrderNo = rstBookPOChild06.Fields("OrderNo").Value
                    ItemName = rstBookPOChild06.Fields("Item").Value
                    Attachment = Trim(rstBookPOChild06.Fields("OrderNo").Value)
                    HeaderPrinted = True
                End If
                rptBookPlateOrder.Subreport2.OpenSubreport.Text25.SetText " (" & Trim(NumberToWords(Round(Val(rstBookPOChild06.Fields("PlateAmount").Value) + Val(rstBookPOChild06.Fields("PAdjustment").Value) + Val(rstBookPOChild06.Fields("PVAT").Value), 0), True)) & ")"
            End If
        End If
        Screen.MousePointer = vbNormal
        If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,EMail,Website FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
        rptBookPlateOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptBookPlateOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
        rptBookPlateOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
        rptBookPlateOrder.Text27.SetText "for " & rptBookPlateOrder.Text12.Text
        rptBookPlateOrder.Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptBookPlateOrder.Text17.SetText Format(TotalTax, "##0.00")
        rptBookPlateOrder.Text18.SetText Format(TotalAmount, "##0.00")
        rptBookPlateOrder.Text19.SetText " (" & Trim(NumberToWords(Round(TotalAmount, 0), True)) & ")"
        rptBookPlateOrder.Subreport1.OpenSubreport.Database.SetDataSource rstBookPOChild05, 3, 1
        rptBookPlateOrder.Subreport2.OpenSubreport.Database.SetDataSource rstBookPOChild06, 3, 1
        Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
        Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
        If OutputTo = "S" Then
            FrmReportViewer.EMailID = EMailID
            FrmReportViewer.Subject = "Book Plate Order #" & Trim(OrderNo) + " Book : " + Trim(ItemName)
            FrmReportViewer.Attachment = Attachment
            FrmReportViewer.Message = Message
            Set FrmReportViewer.Report = rptBookPlateOrder
            FrmReportViewer.Show vbModal
        Else
            If rstBookPOList.State = adStateClosed Then
                If EMailID = "" Or OutputType = "P" Then
                    rptBookPlateOrder.PaperSource = crPRBinAuto
                    rptBookPlateOrder.PrintOut False   ' Print Report Without Prompt
                Else
                    rptBookPlateOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                    rptBookPlateOrder.ExportOptions.DestinationType = crEDTDiskFile
                    rptBookPlateOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                    rptBookPlateOrder.Export False
                    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                    With oOutlookMsg
                        .To = EMailID
                        .Subject = "Book Plate Order #" & Trim(OrderNo) + " Book : " + Trim(ItemName)
                        .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                        .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                        .Importance = olImportanceHigh
                        .ReadReceiptRequested = True
                        .Send
                        If Err.Number = 0 Then cnDatabase.Execute "UPDATE BookPOParent SET BBODStatus=1 WHERE Code='" & OrderCode & "'"
                    End With
                    Set oOutlookMsg = Nothing
                End If
            Else
                rptBookPlateOrder.PaperSource = crPRBinAuto
                rptBookPlateOrder.PrintOut
            End If
        End If
        Set rptBookPlateOrder = Nothing
        rstBookPOChild0801.MoveNext
        If OrderType = "BP" Or OrderType = "TP" Then Exit Do
    Loop
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild05): Call CloseRecordset(rstBookPOChild06): Call CloseRecordset(rstBookPOChild0801)
    On Error GoTo 0
    Screen.MousePointer = vbNormal
End Sub
Public Sub PrintTitlePrintingOrder(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String, Optional ByVal BookPOType As String)
    Dim oOutlookMsg As Outlook.MailItem, RecordAffected As Integer, CoverPrinter As String, TotalTax As Double, TotalAmount As Double
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
    If DatabaseType = "MS SQL" Then
    rstBookPOChild06.Open "SELECT LTRIM(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinter,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMaker,CHOOSE(CONVERT(NUMERIC,C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalFormsFront,C1.TotalFormsBack,C1.PrintRateFront,C1.PrintRateBack,C1.[GST%],C1.GST,C1.Adjustment,C1.PrintAmount,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,C1.[PGST%],C1.PGST,PlateAmount," & _
                          "(SELECT LTRIM(M3.PrintName)+' (UOM : '+LTRIM(M4.PrintName)+'='+LTRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C1.Paper) As PaperName,C1.[PaperWastage%],C1.[PaperConsumptionOther],C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],LTRIM(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,C2.ActualQuantity,[Ups/Plate],C2.PrintingQuantity,C2.BillingQuantity,LTRIM(C2.FrontPrintingColor)+'+'+LTRIM(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType,P.TitlePrinter,C1.PlateMaker,PaperWastageMin,PaperRate,PaperAmountBT,RAdjustment,[RGST%],RGST,PaperAmount,PrintAmountBT,PlateAmountBT,(SELECT EMail FROM AccountMaster WHERE Code=P.TitlePrinter) As EMail " & _
                          "FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic

    Else
    rstBookPOChild06.Open "SELECT LTRIM(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinter,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMaker,CHOOSE(Val(C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalFormsFront,C1.TotalFormsBack,C1.PrintRateFront,C1.PrintRateBack,C1.[GST%],C1.GST,C1.Adjustment,C1.PrintAmount,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,C1.[PGST%],C1.PGST,PlateAmount," & _
                          "(SELECT LTRIM(M3.PrintName)+' (UOM : '+LTRIM(M4.PrintName)+'='+LTRIM(M4.Value1)+')' FROM PaperMaster M3 INNER JOIN GeneralMaster M4 ON M3.UOM=M4.Code WHERE M3.Code=C1.Paper) As PaperName,C1.[PaperWastage%],C1.[PaperConsumptionOther],C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],LTRIM(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,C2.ActualQuantity,[Ups/Plate],C2.PrintingQuantity,C2.BillingQuantity,LTRIM(C2.FrontPrintingColor)+'+'+LTRIM(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType,P.TitlePrinter,C1.PlateMaker,PaperWastageMin,PaperRate,PaperAmountBT,RAdjustment,[RGST%],RGST,PaperAmount,PrintAmountBT,PlateAmountBT,(SELECT EMail FROM AccountMaster WHERE Code=P.TitlePrinter) As EMail " & _
                          "FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    End If
    Screen.MousePointer = vbNormal
    If rstBookPOChild06.RecordCount = 0 Then On Error GoTo 0: Exit Sub
        If DatabaseType = "MS SQL" Then
    CoverPrinter = rstBookPOChild06.Fields("TitlePrinter").Value
         Else
    CoverPrinter = rstBookPOChild06.Fields("P.TitlePrinter").Value
        End If
    TotalTax = TotalTax + IIf(rstBookPOChild06.Fields("P.TitlePrinter").Value <> rstBookPOChild06.Fields("C1.PlateMaker").Value, rstBookPOChild06.Fields("GST").Value + rstBookPOChild06.Fields("RGST").Value, rstBookPOChild06.Fields("GST").Value + rstBookPOChild06.Fields("PGST").Value + rstBookPOChild06.Fields("RGST").Value)
    TotalAmount = TotalAmount + IIf(rstBookPOChild06.Fields("P.TitlePrinter").Value <> rstBookPOChild06.Fields("C1.PlateMaker").Value, rstBookPOChild06.Fields("PrintAmount").Value + rstBookPOChild06.Fields("PaperAmount").Value, rstBookPOChild06.Fields("PrintAmount").Value + rstBookPOChild06.Fields("PlateAmount").Value + rstBookPOChild06.Fields("PaperAmount").Value)
    If rstBookPOChild07.State = adStateOpen Then rstBookPOChild07.Close
    rstBookPOChild07.Open "SELECT (SELECT LTRIM(PrintName) FROM GeneralMaster WHERE Code = C.operation) As LaminationType,C.Quantity,Rate,CalcMode,Amount,Adjustment,[GST%],GST,BillAmount,C.Remarks " & _
                          "FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Code='" & OrderCode & "' AND P.Laminator='" & CoverPrinter & "'", cnDatabase, adOpenKeyset, adLockOptimistic         'AND P.Laminator='" & CoverPrinter & "'", cnDatabase, adOpenKeyset, adLockOptimistic
    If rstBookPOChild07.RecordCount = 0 Then
        rptTitlePrintingOrder.Section13.Suppress = True
        rptTitlePrintingOrder.Section15.Suppress = True
    Else
        TotalTax = TotalTax + rstBookPOChild07.Fields("VAT").Value
        TotalAmount = TotalAmount + rstBookPOChild07.Fields("BillAmount").Value
    End If
    If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,EMail,Website FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    rptTitlePrintingOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptTitlePrintingOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptTitlePrintingOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
    rptTitlePrintingOrder.Text20.SetText Format(IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, Round(rstBookPOChild06.Fields("PrintAmount").Value + rstBookPOChild06.Fields("PaperAmount").Value, 0), Round(rstBookPOChild06.Fields("PrintAmount").Value + rstBookPOChild06.Fields("PlateAmount").Value + rstBookPOChild06.Fields("PaperAmount").Value, 0)), "#0.00") + " (" & Trim(NumberToWords(IIf(rstBookPOChild06.Fields("TitlePrinter").Value <> rstBookPOChild06.Fields("PlateMaker").Value, Round(rstBookPOChild06.Fields("PrintAmount").Value + rstBookPOChild06.Fields("PaperAmount").Value, 0), Round(rstBookPOChild06.Fields("PrintAmount").Value + rstBookPOChild06.Fields("PlateAmount").Value + rstBookPOChild06.Fields("PaperAmount").Value, 0)), True)) & ")"
    rptTitlePrintingOrder.Text27.SetText "for " & Trim(rstBookPOChild06.Fields("TitlePrinter").Value)
    rptTitlePrintingOrder.Text22.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    If TotalAmount = 0 Then
        rptTitlePrintingOrder.Section15.Suppress = True
        rptTitlePrintingOrder.Section3.Suppress = True
    Else
        rptTitlePrintingOrder.Text29.SetText Format(TotalTax, "##0.00")
        rptTitlePrintingOrder.Text31.SetText Format(TotalAmount, "##0.00")
        rptTitlePrintingOrder.Text32.SetText Trim(NumberToWords(Round(TotalAmount, 0), True))
    End If
    rstBookPOChild06.MoveFirst
    rptTitlePrintingOrder.Database.SetDataSource rstBookPOChild06, 3, 1
    rptTitlePrintingOrder.Subreport1.OpenSubreport.Database.SetDataSource rstBookPOChild07, 3, 1
    EMailID = rstBookPOChild06.Fields("EMail").Value
    Attachment = Trim(rstBookPOChild06.Fields("OrderNo").Value)
    Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
    Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild06.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
    If OutputTo = "S" Then
        FrmReportViewer.EMailID = EMailID
        FrmReportViewer.Subject = "Title Printing Order #" & Trim(rstBookPOChild06.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild06.Fields("BookName").Value)
        FrmReportViewer.Attachment = Attachment
        FrmReportViewer.Message = Message
        Set FrmReportViewer.Report = rptTitlePrintingOrder
        FrmReportViewer.Show vbModal
    Else
        If rstBookPOList.State = adStateClosed Then
            If EMailID = "" Or OutputType = "P" Then
                rptTitlePrintingOrder.PaperSource = crPRBinAuto
                rptTitlePrintingOrder.PrintOut False   ' Print Report Without Prompt
            Else
                rptTitlePrintingOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptTitlePrintingOrder.ExportOptions.DestinationType = crEDTDiskFile
                rptTitlePrintingOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                rptTitlePrintingOrder.Export False
                rstBookPOChild06.MoveFirst
                Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                With oOutlookMsg
                    .To = EMailID
                    .Subject = "Title Printing Order #" & Trim(rstBookPOChild06.Fields("OrderNo").Value)
                    .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                    .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                    .Importance = olImportanceHigh
                    .ReadReceiptRequested = True
                    .Send
                    If Err.Number = 0 Then cnDatabase.Execute "UPDATE BookPOParent SET TPODStatus=1 WHERE Code='" & OrderCode & "'", RecordAffected
                    If RecordAffected = 0 Then DisplayError ("Failed to update EMail Flag (Title Print Order)")
                End With
                Set oOutlookMsg = Nothing
            End If
        Else
            rptTitlePrintingOrder.PaperSource = crPRBinAuto
            rptTitlePrintingOrder.PrintOut
        End If
    End If
    Set rptTitlePrintingOrder = Nothing
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild06)
    On Error GoTo 0
End Sub
Public Sub PrintTitlePlateOrder(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String, Optional ByVal BookPOType As String)
    Dim oOutlookMsg As Outlook.MailItem, RecordAffected As Integer
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
    rstBookPOChild06.Open "SELECT LTRIM(P.Name) As OrderNo,P.Date As OrderDate,TargetDate,Plate,(SELECT PrintName FROM AccountMaster WHERE Code=C1.PlateMaker) As PlateMaker,CHOOSE(CONVERT(NUMERIC,C1.PlateType),'Deep-etch','PS','Wipe-on','CTP') As PlateName,C1.TotalPlates,C1.PlateRate,C1.PAdjustment,C1.[PGST%],C1.PGST,PlateAmount," & _
                          "C1.Remarks,(SELECT PrintName FROM GeneralMaster WHERE Code=C1.[Size]) As [Size],LTrim(M1.PrintName)+iif(M1.Price=0,'',' (Price : Rs. '+Format(M1.Price,'0.00')+')') As Item,[Ups/Plate],LTRIM(C2.FrontPrintingColor)+'+'+LTRIM(C2.BackPrintingColor) + ' ('+IIF(C1.Imposition='F','F&B','W&T')+')' As PrintingType,(SELECT EMail FROM AccountMaster WHERE Code=C1.PlateMaker) As EMail " & _
                          "FROM ((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN BookMaster M1 ON C2.Book=M1.Code WHERE P.Code='" & OrderCode & "' And P.TitlePrinter <> ''", cnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    If rstBookPOChild06.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,EMail,Website FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    rptTitlePlateOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptTitlePlateOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptTitlePlateOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
    rptTitlePlateOrder.Text20.SetText Format(Round(rstBookPOChild06.Fields("PlateAmount").Value, 0), "#0.00") + " (" & Trim(NumberToWords(Round(rstBookPOChild06.Fields("PlateAmount").Value, 0), True)) & ")"
    rptTitlePlateOrder.Text27.SetText "for " & Trim(rstBookPOChild06.Fields("TitlePrinter").Value)
    rptTitlePlateOrder.Text22.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptTitlePlateOrder.Database.SetDataSource rstBookPOChild06, 3, 1
    EMailID = rstBookPOChild06.Fields("EMail").Value
    Attachment = Trim(rstBookPOChild06.Fields("OrderNo").Value)
    Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
    Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild06.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
    If OutputTo = "S" Then
        FrmReportViewer.EMailID = EMailID
        FrmReportViewer.Subject = "Title Plate Order #" & Trim(rstBookPOChild06.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild06.Fields("BookName").Value)
        FrmReportViewer.Attachment = Attachment
        FrmReportViewer.Message = Message
        Set FrmReportViewer.Report = rptTitlePlateOrder
        FrmReportViewer.Show vbModal
    Else
        If rstBookPOList.State = adStateClosed Then
            If EMailID = "" Or OutputType = "P" Then
                rptTitlePlateOrder.PaperSource = crPRBinAuto
                rptTitlePlateOrder.PrintOut False   ' Print Report Without Prompt
            Else
                rptTitlePlateOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptTitlePlateOrder.ExportOptions.DestinationType = crEDTDiskFile
                rptTitlePlateOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                rptTitlePlateOrder.Export False
                rstBookPOChild06.MoveFirst
                Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                With oOutlookMsg
                    .To = EMailID
                    .Subject = "Title Plate Order #" & Trim(rstBookPOChild06.Fields("OrderNo").Value)
                    .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                    .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                    .Importance = olImportanceHigh
                    .ReadReceiptRequested = True
                    .Send
                    If Err.Number = 0 Then cnDatabase.Execute "UPDATE BookPOParent SET TPODStatus=1 WHERE Code='" & OrderCode & "'", RecordAffected
                    If RecordAffected = 0 Then DisplayError ("Failed to update EMail Flag (Title Print Order)")
                End With
                Set oOutlookMsg = Nothing
            End If
        Else
            rptTitlePlateOrder.PaperSource = crPRBinAuto
            rptTitlePlateOrder.PrintOut
        End If
    End If
    Set rptTitlePlateOrder = Nothing
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild06)
    On Error GoTo 0
End Sub
Public Sub PrintCostSheet(ByVal OrderCode As String)
    Dim oExcel As Object, i As Integer
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    rstBookPOChild05.Open "SELECT DISTINCT M1.PrintName As BookName,M2.PrintName As BookSize,M1.FormType,M1.Royalty," & _
                                                       "C1.ActualQuantity As Quantity,(C1.Pages1+C1.Pages2+C1.Pages4) As Pages,C1.Forms1,C1.PrintRate1,C1.PrintAmount1,C1.PlateType1,C1.[TotalPlates1-¼]+C1.[TotalPlates1-½]+C1.[TotalPlates1-1] As TotalPlates1,C1.PlateRate1,C1.PlateAmount1,(SELECT STR([Weight/Unit]) FROM PaperMaster WHERE Code=C1.Paper1) As TextPaper1,PaperConsumptionOther1 As TextPaperConsumption1,C1.PaperRate1,C1.PaperAmount1," & _
                                                       "C1.Forms2,C1.PrintRate2,C1.PrintAmount2,C1.PlateType2,C1.[TotalPlates2-¼]+C1.[TotalPlates2-½]+C1.[TotalPlates2-1] As TotalPlates2,C1.PlateRate2,C1.PlateAmount2,(SELECT STR([Weight/Unit]) FROM PaperMaster WHERE Code=C1.Paper2) As TextPaper2,PaperConsumptionOther2 As TextPaperConsumption2,C1.PaperRate2,C1.PaperAmount2," & _
                                                       "C1.Forms4,C1.PrintRate4,C1.PrintAmount4,C1.PlateType4,C1.[TotalPlates4-¼]+C1.[TotalPlates4-½]+C1.[TotalPlates4-1] As TotalPlates4,C1.PlateRate4 ,C1.PlateAmount4,C1.Paper4,(SELECT STR([Weight/Unit]) FROM PaperMaster WHERE Code=C1.Paper4) As TextPaper4,PaperConsumptionOther4 As TextPaperConsumption4,C1.PaperRate4,C1.PaperAmount4," & _
                                                       "LTRIM(C2.FrontPrintingType)+'+'+LTRIM(C2.BackPrintingType) As TitlePrintingType,C2.PrintRate As TitlePrintRate,C2.PrintAmount As TitlePrintAmount,C2.PlateType As TitlePlateType,C2.TotalPlates As TitleTotalPlates,C2.PlateRate As TitlePlateRate,C2.PlateAmount As TitlePlateAmount,(SELECT STR([Weight/Unit]) FROM PaperMaster WHERE Code=C2.Paper) As TitlePaper,C2.PaperConsumptionOther As TitlePaperConsumption,C2.PaperRate As TitlePaperRate,C2.PaperAmount As TitlePaperAmount," & _
                                                       "(SELECT LTRIM(MAX(FrontPrintingColor))+'+'+LTRIM(MAX(BackPrintingColor)) FROM BookPOChild0901 WHERE Code=C5.Code) As ComboPrintingType,LTRIM(C5.PrintRateFront)+'+'+LTRIM(C5.PrintRateBack) As ComboPrintRate,C5.PrintAmountBT As ComboPrintAmount,C5.PlateType As ComboPlateType,C5.TotalPlates As ComboTotalPlates,C5.PlateRate As ComboPlateRate,C5.PlateAmountBT As ComboPlateAmount,(SELECT STR([Weight/Unit]) FROM PaperMaster WHERE Code=C5.Paper) As ComboPaper,C5.PaperConsumptionOther As ComboPaperConsumption,C5.PaperRate As ComboPaperRate,C5.PaperAmount As ComboPaperAmount," & _
                                                       "(SELECT PrintName FROM GeneralMaster WHERE Code=C4.BindingType) As BindingType,(C4.BindingForms+C4.ExtraForms) As BindingForms,C4.FormFoldRate,C4.FormStitchRate,C4.FormPasteRate,C4.[Rate/Book],(TotalPkts*PktPackRate)+(TotalBoxes*BoxPackRate)+(TotalBoxes*CartageRate) As [Packing&Cartage]," & _
                                                       "EstQty01,EstQty02,EstQty03,EstQty04,EstQty05," & _
                                                       "IIF(ISNULL(C1.VAT)=True,0,C1.VAT)+IIF(ISNULL(C1.PVAT)=True,0,C1.PVAT)+IIF(ISNULL(C1.RVAT)=True,0,C1.RVAT)+IIF(ISNULL(C2.VAT)=True,0,C2.VAT)+IIF(ISNULL(C2.PVAT)=True,0,C2.PVAT)+IIF(ISNULL(C2.RVAT)=True,0,C2.RVAT)+IIF(ISNULL(C4.VAT)=True,0,C4.VAT)+IIF(ISNULL(C5.GST)=True,0,C5.GST)+IIF(ISNULL(C5.PGST)=True,0,C5.PGST)+IIF(ISNULL(C5.RGST)=True,0,C5.RGST) As GST,IIF(ISNULL(C1.Adjustment)=True,0,C1.Adjustment)+IIF(ISNULL(C1.PAdjustment)=True,0,C1.PAdjustment)+IIF(ISNULL(C1.RAdjustment)=True,0,C1.RAdjustment)+IIF(ISNULL(C2.Adjustment)=True,0,C2.Adjustment)+IIF(ISNULL(C2.PAdjustment)=True,0,C2.PAdjustment)+IIF(ISNULL(C2.RAdjustment)=True,0,C2.RAdjustment)+IIF(ISNULL(C4.Adjustment)=True,0,C4.Adjustment)+IIF(ISNULL(C5.Adjustment)=True,0,C5.Adjustment)+IIF(ISNULL(C5.PAdjustment)=True,0,C5.PAdjustment)+IIF(ISNULL(C5.RAdjustment)=True,0,C5.RAdjustment) As Adjustment " & _
                                                       "FROM (((((BookPOParent P INNER JOIN BookPOChild05 C1 ON P.Code=C1.Code) LEFT JOIN BookPOChild06 C2 ON P.Code=C2.Code) LEFT JOIN BookPOChild08 C4 ON P.Code=C4.Code) LEFT JOIN BookPOChild09 C5 ON P.Code=C5.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.[FinishSize]=M2.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
    rstBookPOChild06.Open "SELECT SUM(GST) As GST,SUM(Adjustment) As Adjustment,SUM(Amount) As BillAmount FROM BookPOChild07 WHERE Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    Screen.MousePointer = vbNormal
    If rstBookPOChild05.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Cost Sheet")
    oExcel.DisplayAlerts = False
    If Dir(App.Path & "\Costing\", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Costing\"
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Cost Sheet\" & Replace(Trim(rstBookPOList.Fields("BookName").Value), "*", "") + " (" + Format(Date, "dd-MM-yyyy") + ")")
    oExcel.DisplayAlerts = True
    oExcel.Sheets("Sheet1").Select
    oExcel.Sheets("Sheet1").Unprotect ("eisi")
    '
    oExcel.Application.Cells(3, "B").Value = Trim(rstBookPOChild05.Fields("BookName").Value)
    oExcel.Application.Cells(4, "B").Value = Val(rstBookPOChild05.Fields("Pages").Value)
    oExcel.Application.Cells(53, "C").Value = Val(rstBookPOChild05.Fields("Royalty").Value) / 100
    oExcel.Application.Cells(5, "B").Value = Val(rstBookPOChild05.Fields("Quantity").Value): oExcel.Application.Cells(5, "C").Value = Val(rstBookPOChild05.Fields("EstQty01").Value): oExcel.Application.Cells(5, "D").Value = Val(rstBookPOChild05.Fields("EstQty02").Value): oExcel.Application.Cells(5, "E").Value = Val(rstBookPOChild05.Fields("EstQty03").Value): oExcel.Application.Cells(5, "F").Value = Val(rstBookPOChild05.Fields("EstQty04").Value): oExcel.Application.Cells(5, "G").Value = Val(rstBookPOChild05.Fields("EstQty05").Value)
    oExcel.Application.Cells(6, "B").Value = Trim(rstBookPOChild05.Fields("BookSize").Value) & "/" & Choose(Val(rstBookPOChild05.Fields("FormType").Value), "08", "16", "04", "12", "24", "32", "64", "06", "02")
    oExcel.Application.Cells(7, "B").Value = Val(rstBookPOChild05.Fields("Forms1").Value) + Val(rstBookPOChild05.Fields("Forms2").Value) + Val(rstBookPOChild05.Fields("Forms4").Value)
    oExcel.Application.Cells(8, "B").Value = Val(rstBookPOChild05.Fields("BindingForms").Value)
    oExcel.Application.Cells(9, "B").Value = Trim(rstBookPOChild05.Fields("LaminationType").Value)
    oExcel.Application.Cells(10, "B").Value = Trim(rstBookPOChild05.Fields("BindingType").Value)
    '
    If Not IsNull(rstBookPOChild05.Fields("TextPaper1").Value) Then
        oExcel.Application.Cells(24, "C").Value = Val(rstBookPOChild05.Fields("TextPaper1").Value)    'Weight/Unit
        oExcel.Application.Cells(24, "D").Value = Val(rstBookPOChild05.Fields("TextPaperConsumption1").Value)
        oExcel.Application.Cells(24, "E").Value = Val(rstBookPOChild05.Fields("PaperRate1").Value)
        oExcel.Application.Cells(24, "G").Value = Val(rstBookPOChild05.Fields("PaperAmount1").Value)
        '
        oExcel.Application.Cells(29, "C").Value = Choose(Val(rstBookPOChild05.Fields("PlateType1").Value), "Deepatch", "PS", "Wipeon", "CTP")
        oExcel.Application.Cells(29, "D").Value = Val(rstBookPOChild05.Fields("TotalPlates1").Value)
        oExcel.Application.Cells(29, "E").Value = Val(rstBookPOChild05.Fields("PlateRate1").Value)
        oExcel.Application.Cells(29, "G").Value = Val(rstBookPOChild05.Fields("PlateAmount1").Value)
        '
        oExcel.Application.Cells(34, "D").Value = Val(rstBookPOChild05.Fields("Forms1").Value)
        oExcel.Application.Cells(34, "E").Value = Val(rstBookPOChild05.Fields("PrintRate1").Value)
        oExcel.Application.Cells(34, "G").Value = Val(rstBookPOChild05.Fields("PrintAmount1").Value)
    End If
    If Not IsNull(rstBookPOChild05.Fields("TextPaper2").Value) Then
        oExcel.Application.Cells(25, "C").Value = Val(rstBookPOChild05.Fields("TextPaper2").Value)
        oExcel.Application.Cells(25, "D").Value = Val(rstBookPOChild05.Fields("TextPaperConsumption2").Value)
        oExcel.Application.Cells(25, "E").Value = Val(rstBookPOChild05.Fields("PaperRate2").Value)
        oExcel.Application.Cells(25, "G").Value = Val(rstBookPOChild05.Fields("PaperAmount2").Value)
        '
        oExcel.Application.Cells(30, "C").Value = Choose(Val(rstBookPOChild05.Fields("PlateType2").Value), "Deepatch", "PS", "Wipeon", "CTP")
        oExcel.Application.Cells(30, "D").Value = Val(rstBookPOChild05.Fields("TotalPlates2").Value)
        oExcel.Application.Cells(30, "E").Value = Val(rstBookPOChild05.Fields("PlateRate2").Value)
        oExcel.Application.Cells(30, "G").Value = Val(rstBookPOChild05.Fields("PlateAmount2").Value)
        '
        oExcel.Application.Cells(35, "D").Value = Val(rstBookPOChild05.Fields("Forms2").Value)
        oExcel.Application.Cells(35, "E").Value = Val(rstBookPOChild05.Fields("PrintRate2").Value)
        oExcel.Application.Cells(35, "G").Value = Val(rstBookPOChild05.Fields("PrintAmount2").Value)
    End If
    If Not IsNull(rstBookPOChild05.Fields("TextPaper4").Value) Then
        oExcel.Application.Cells(26, "C").Value = Val(rstBookPOChild05.Fields("TextPaper4").Value)
        oExcel.Application.Cells(26, "D").Value = Val(rstBookPOChild05.Fields("TextPaperConsumption4").Value)
        oExcel.Application.Cells(26, "E").Value = Val(rstBookPOChild05.Fields("PaperRate4").Value)
        oExcel.Application.Cells(26, "G").Value = Val(rstBookPOChild05.Fields("PaperAmount4").Value)
        '
        oExcel.Application.Cells(31, "C").Value = Choose(Val(rstBookPOChild05.Fields("PlateType4").Value), "Deepatch", "PS", "Wipeon", "CTP")
        oExcel.Application.Cells(31, "D").Value = Val(rstBookPOChild05.Fields("TotalPlates4").Value)
        oExcel.Application.Cells(31, "E").Value = Val(rstBookPOChild05.Fields("PlateRate4").Value)
        oExcel.Application.Cells(31, "G").Value = Val(rstBookPOChild05.Fields("PlateAmount4").Value)
        '
        oExcel.Application.Cells(36, "D").Value = Val(rstBookPOChild05.Fields("Forms4").Value)
        oExcel.Application.Cells(36, "E").Value = Val(rstBookPOChild05.Fields("PrintRate4").Value)
        oExcel.Application.Cells(36, "G").Value = Val(rstBookPOChild05.Fields("PrintAmount4").Value)
    End If
    If Not IsNull(rstBookPOChild05.Fields("TitlePaper").Value) Then
        oExcel.Application.Cells(27, "C").Value = Val(rstBookPOChild05.Fields("TitlePaper").Value)
        oExcel.Application.Cells(27, "D").Value = Val(rstBookPOChild05.Fields("TitlePaperConsumption").Value)
        oExcel.Application.Cells(27, "E").Value = Val(rstBookPOChild05.Fields("TitlePaperRate").Value)
        oExcel.Application.Cells(27, "G").Value = Val(rstBookPOChild05.Fields("TitlePaperAmount").Value)
        '
        oExcel.Application.Cells(32, "B").Value = rstBookPOChild05.Fields("TitlePrintingType").Value & " Color"
        oExcel.Application.Cells(32, "C").Value = Choose(Val(rstBookPOChild05.Fields("TitlePlateType").Value), "Deepatch", "PS", "Wipeon", "CTP")
        oExcel.Application.Cells(32, "D").Value = Val(rstBookPOChild05.Fields("TitleTotalPlates").Value)
        oExcel.Application.Cells(32, "E").Value = Val(rstBookPOChild05.Fields("TitlePlateRate").Value)
        oExcel.Application.Cells(32, "G").Value = Val(rstBookPOChild05.Fields("TitlePlateAmount").Value)
        '
        oExcel.Application.Cells(37, "B").Value = rstBookPOChild05.Fields("TitlePrintingType").Value & " Color"
        oExcel.Application.Cells(37, "E").Value = Val(rstBookPOChild05.Fields("TitlePrintRate").Value)
        oExcel.Application.Cells(37, "G").Value = Val(rstBookPOChild05.Fields("TitlePrintAmount").Value)
    End If
    If Not IsNull(rstBookPOChild05.Fields("ComboPaper").Value) Then
        oExcel.Application.Cells(28, "C").Value = Val(rstBookPOChild05.Fields("ComboPaper").Value)
        oExcel.Application.Cells(28, "D").Value = Val(rstBookPOChild05.Fields("ComboPaperConsumption").Value)
        oExcel.Application.Cells(28, "E").Value = Val(rstBookPOChild05.Fields("ComboPaperRate").Value)
        oExcel.Application.Cells(28, "G").Value = Val(rstBookPOChild05.Fields("ComboPaperAmount").Value)
        '
        oExcel.Application.Cells(33, "B").Value = rstBookPOChild05.Fields("ComboPrintingType").Value & " Color"
        oExcel.Application.Cells(33, "C").Value = Choose(Val(rstBookPOChild05.Fields("ComboPlateType").Value), "Deepatch", "PS", "Wipeon", "CTP")
        oExcel.Application.Cells(33, "D").Value = Val(rstBookPOChild05.Fields("ComboTotalPlates").Value)
        oExcel.Application.Cells(33, "E").Value = Val(rstBookPOChild05.Fields("ComboPlateRate").Value)
        oExcel.Application.Cells(33, "G").Value = Val(rstBookPOChild05.Fields("ComboPlateAmount").Value)
        '
        oExcel.Application.Cells(38, "B").Value = rstBookPOChild05.Fields("ComboPrintingType").Value & " Color"
        oExcel.Application.Cells(38, "E").Value = rstBookPOChild05.Fields("ComboPrintRate").Value
        oExcel.Application.Cells(38, "G").Value = Val(rstBookPOChild05.Fields("ComboPrintAmount").Value)
    End If
    '
    oExcel.Application.Cells(40, "E").Value = Val(rstBookPOChild05.Fields("FormFoldRate").Value)
    oExcel.Application.Cells(41, "E").Value = Val(rstBookPOChild05.Fields("FormStitchRate").Value)
    oExcel.Application.Cells(42, "E").Value = Val(rstBookPOChild05.Fields("FormPasteRate").Value) / 1000
    oExcel.Application.Cells(43, "G").Value = Val(rstBookPOChild05.Fields("Packing&Cartage").Value)
    oExcel.Application.Cells(44, "E").Value = Val(rstBookPOChild05.Fields("Rate/Book").Value)
    If rstBookPOChild06.RecordCount = 0 Then
        oExcel.Application.Cells(45, "D").Value = Val(rstBookPOChild05.Fields("Adjustment").Value): oExcel.Application.Cells(45, "F").Value = Val(rstBookPOChild05.Fields("GST").Value)
    Else
        oExcel.Application.Cells(39, "G").Value = Val(rstBookPOChild06.Fields("BillAmount").Value)
        oExcel.Application.Cells(45, "D").Value = Val(rstBookPOChild05.Fields("Adjustment").Value) + Val(rstBookPOChild06.Fields("Adjustment").Value): oExcel.Application.Cells(45, "F").Value = Val(rstBookPOChild05.Fields("GST").Value) + Val(rstBookPOChild06.Fields("GST").Value)
    End If
    Screen.MousePointer = vbHourglass
    If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    rstBookPOChild05.Open "SELECT E.PrintName As EName,O.PrintName As OName,Number,Quantity,Rate,M.PrintName As MName,Amount FROM (((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN GeneralMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code WHERE P.Code='" & OrderCode & "' ORDER BY E.PrintName,O.PrintName", cnDatabase, adOpenKeyset, adLockReadOnly
    Screen.MousePointer = vbNormal
    With rstBookPOChild05
        If .RecordCount > 0 Then
            i = 14
            Do While Not .EOF
                oExcel.Application.Cells(i, "J").Value = .Fields("EName").Value
                oExcel.Application.Cells(i, "L").Value = .Fields("OName").Value
                oExcel.Application.Cells(i, "O").Value = Val(.Fields("Number").Value)
                oExcel.Application.Cells(i, "P").Value = Val(.Fields("Quantity").Value)
                oExcel.Application.Cells(i, "Q").Value = Val(.Fields("Rate").Value)
                oExcel.Application.Cells(i, "R").Value = .Fields("MName").Value
                oExcel.Application.Cells(i, "S").Value = .Fields("Amount").Value
                .MoveNext
                i = i + 1
            Loop
        End If
    End With
    oExcel.Sheets("Sheet1").Protect ("eisi")
    oExcel.Workbooks.Item(1).Save
    If OutputTo = "S" Then
        oExcel.Range("A1").Activate
        oExcel.Application.Visible = True
    Else
        oExcel.Workbooks.Item(1).PrintOut
    End If
    Set oExcel = Nothing
    On Error GoTo 0
End Sub
Public Sub PrintPlanning(ByVal OrderCode As String)
    Dim oExcel As Object, i As Integer
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    rstBookPOChild05.Open "SELECT DISTINCT M1.PrintName As BookName,M2.PrintName As BookSize,M1.FormType,M1.Royalty," & _
                                                       "C1.ActualQuantity As Quantity,(C1.Pages1+C1.Pages2+C1.Pages4) As Pages,C1.Forms1,C1.PrintRate1,C1.PrintAmount1,C1.PlateType1,C1.[TotalPlates1-¼]+C1.[TotalPlates1-½]+C1.[TotalPlates1-1] As TotalPlates1,C1.PlateRate1,C1.PlateAmount1,(SELECT STR([Weight/Unit]) FROM PaperMaster WHERE Code=C1.Paper1) As TextPaper1,PaperConsumptionOther1 As TextPaperConsumption1,C1.PaperRate1,C1.PaperAmount1," & _
                                                       "C1.Forms2,C1.PrintRate2,C1.PrintAmount2,C1.PlateType2,C1.[TotalPlates2-¼]+C1.[TotalPlates2-½]+C1.[TotalPlates2-1] As TotalPlates2,C1.PlateRate2,C1.PlateAmount2,(SELECT STR([Weight/Unit]) FROM PaperMaster WHERE Code=C1.Paper2) As TextPaper2,PaperConsumptionOther2 As TextPaperConsumption2,C1.PaperRate2,C1.PaperAmount2," & _
                                                       "C1.Forms4,C1.PrintRate4,C1.PrintAmount4,C1.PlateType4,C1.[TotalPlates4-¼]+C1.[TotalPlates4-½]+C1.[TotalPlates4-1] As TotalPlates4,C1.PlateRate4 ,C1.PlateAmount4,C1.Paper4,(SELECT STR([Weight/Unit]) FROM PaperMaster WHERE Code=C1.Paper4) As TextPaper4,PaperConsumptionOther4 As TextPaperConsumption4,C1.PaperRate4,C1.PaperAmount4," & _
                                                       "TRIM(C2.FrontPrintingType)+'+'+TRIM(C2.BackPrintingType) As TitlePrintingType,C2.PrintRate As TitlePrintRate,C2.PrintAmount As TitlePrintAmount,C2.PlateType As TitlePlateType,C2.TotalPlates As TitleTotalPlates,C2.PlateRate As TitlePlateRate,C2.PlateAmount As TitlePlateAmount,(SELECT STR([Weight/Unit]) FROM PaperMaster WHERE Code=C2.Paper) As TitlePaper,C2.PaperConsumptionOther As TitlePaperConsumption,C2.PaperRate As TitlePaperRate,C2.PaperAmount As TitlePaperAmount," & _
                                                       "(SELECT TRIM(MAX(FrontPrintingColor))+'+'+TRIM(MAX(BackPrintingColor)) FROM BookPOChild0901 WHERE Code=C5.Code) As ComboPrintingType,TRIM(C5.PrintRateFront)+'+'+TRIM(C5.PrintRateBack) As ComboPrintRate,C5.PrintAmountBT As ComboPrintAmount,C5.PlateType As ComboPlateType,C5.TotalPlates As ComboTotalPlates,C5.PlateRate As ComboPlateRate,C5.PlateAmountBT As ComboPlateAmount,(SELECT STR([Weight/Unit]) FROM PaperMaster WHERE Code=C5.Paper) As ComboPaper,C5.PaperConsumptionOther As ComboPaperConsumption,C5.PaperRate As ComboPaperRate,C5.PaperAmount As ComboPaperAmount," & _
                                                       "(SELECT PrintName FROM GeneralMaster WHERE Code=C4.BindingType) As BindingType,(C4.BindingForms+C4.ExtraForms) As BindingForms,C4.FormFoldRate,C4.FormStitchRate,C4.FormPasteRate,C4.[Rate/Book],(TotalPkts*PktPackRate)+(TotalBoxes*BoxPackRate)+(TotalBoxes*CartageRate) As [Packing&Cartage]," & _
                                                       "EstQty01,EstQty02,EstQty03,EstQty04,EstQty05," & _
                                                       "IIF(ISNULL(C1.VAT)=True,0,C1.VAT)+IIF(ISNULL(C1.PVAT)=True,0,C1.PVAT)+IIF(ISNULL(C1.RVAT)=True,0,C1.RVAT)+IIF(ISNULL(C2.VAT)=True,0,C2.VAT)+IIF(ISNULL(C2.PVAT)=True,0,C2.PVAT)+IIF(ISNULL(C2.RVAT)=True,0,C2.RVAT)+IIF(ISNULL(C4.VAT)=True,0,C4.VAT)+IIF(ISNULL(C5.GST)=True,0,C5.GST)+IIF(ISNULL(C5.PGST)=True,0,C5.PGST)+IIF(ISNULL(C5.RGST)=True,0,C5.RGST) As GST,IIF(ISNULL(C1.Adjustment)=True,0,C1.Adjustment)+IIF(ISNULL(C1.PAdjustment)=True,0,C1.PAdjustment)+IIF(ISNULL(C1.RAdjustment)=True,0,C1.RAdjustment)+IIF(ISNULL(C2.Adjustment)=True,0,C2.Adjustment)+IIF(ISNULL(C2.PAdjustment)=True,0,C2.PAdjustment)+IIF(ISNULL(C2.RAdjustment)=True,0,C2.RAdjustment)+IIF(ISNULL(C4.Adjustment)=True,0,C4.Adjustment)+IIF(ISNULL(C5.Adjustment)=True,0,C5.Adjustment)+IIF(ISNULL(C5.PAdjustment)=True,0,C5.PAdjustment)+IIF(ISNULL(C5.RAdjustment)=True,0,C5.RAdjustment) As Adjustment,P.Name,M3.Name As Client,M4.Name As PSize, " & _
                                                       "C1.Pages1 As Pages1,C1.Pages2 As Pages2,C1.Pages4 As Pages4 " & _
                                                       "FROM (((((((BookPOParent P INNER JOIN BookPOChild05 C1 ON P.Code=C1.Code) LEFT JOIN BookPOChild06 C2 ON P.Code=C2.Code) LEFT JOIN BookPOChild08 C4 ON P.Code=C4.Code) LEFT JOIN BookPOChild09 C5 ON P.Code=C5.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.[FinishSize]=M2.Code) INNER JOIN GeneralMaster M3 ON M1.[Group]=M3.Code) INNER JOIN GeneralMaster M4 ON M1.[Size]=M4.Code WHERE P.Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
    rstBookPOChild06.Open "SELECT SUM(GST) As GST,SUM(Adjustment) As Adjustment,SUM(Amount) As BillAmount FROM BookPOChild07 WHERE Code='" & OrderCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    Screen.MousePointer = vbNormal
    If rstBookPOChild05.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Imposition")
    If Dir(App.Path & "\Imposition\", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Imposition\"
    oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Imposition\" & Replace(Trim(rstBookPOList.Fields("BookName").Value), "*", "") + " (Ref No-" + Trim(rstBookPOChild05.Fields("Name").Value) + " )" + " (" + Format(Date, "dd-MM-yyyy") + ")")
'    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Imposition\" & Replace(Trim(rstBookPOList.Fields("BookName").Value), "*", "") + " (" + Format(Date, "dd-MM-yyyy") + ")")
    oExcel.DisplayAlerts = True
    oExcel.Sheets("Sheet1").Select
    oExcel.Sheets("Sheet1").Unprotect ("eisi")
    '
    oExcel.Application.Cells(1, "B").Value = Trim(rstBookPOChild05.Fields("BookName").Value)
    oExcel.Application.Cells(2, "D").Value = Trim(rstBookPOChild05.Fields("PSize").Value)
    oExcel.Application.Cells(4, "B").Value = Trim(IIf((rstBookPOChild05.Fields("Pages1").Value) > 0, "1Col", "") & IIf((rstBookPOChild05.Fields("Pages2").Value) > 0, "2Col", "") & IIf((rstBookPOChild05.Fields("Pages4").Value) > 0, "4Col", ""))
    oExcel.Application.Cells(3, "D").Value = Trim(rstBookPOChild05.Fields("Name").Value)
    oExcel.Application.Cells(1, "H").Value = Trim(rstBookPOChild05.Fields("Client").Value)
    oExcel.Application.Cells(3, "E").Value = Val(rstBookPOChild05.Fields("Pages").Value)
'    oExcel.Application.Cells(53, "C").Value = Val(rstBookPOChild05.Fields("Royalty").Value) / 100
    oExcel.Application.Cells(3, "B").Value = Val(rstBookPOChild05.Fields("EstQty01").Value)
    oExcel.Application.Cells(2, "B").Value = Trim(rstBookPOChild05.Fields("BookSize").Value)
'    oExcel.Application.Cells(2, "D").Value = Trim(rstBookPOChild05.Fields("BookSize").Value) & "/" & Choose(Val(rstBookPOChild05.Fields("FormType").Value), "08", "16", "04", "12", "24", "32", "64", "06", "02")
    oExcel.Application.Cells(5, "E").Value = Choose(Val(rstBookPOChild05.Fields("FormType").Value), "08", "16", "04", "12", "24", "32", "64", "06", "02")
'    oExcel.Application.Cells(7, "B").Value = Val(rstBookPOChild05.Fields("Forms1").Value) + Val(rstBookPOChild05.Fields("Forms2").Value) + Val(rstBookPOChild05.Fields("Forms4").Value)
'    oExcel.Application.Cells(8, "B").Value = Val(rstBookPOChild05.Fields("BindingForms").Value)
'    oExcel.Application.Cells(9, "B").Value = Trim(rstBookPOChild05.Fields("LaminationType").Value)
    oExcel.Application.Cells(5, "G").Value = Trim(rstBookPOChild05.Fields("BindingType").Value)
    '
'    If Not IsNull(rstBookPOChild05.Fields("TextPaper1").Value) Then
'        oExcel.Application.Cells(24, "C").Value = Val(rstBookPOChild05.Fields("TextPaper1").Value)    'Weight/Unit
'        oExcel.Application.Cells(24, "D").Value = Val(rstBookPOChild05.Fields("TextPaperConsumption1").Value)
'        oExcel.Application.Cells(24, "E").Value = Val(rstBookPOChild05.Fields("PaperRate1").Value)
'        oExcel.Application.Cells(24, "G").Value = Val(rstBookPOChild05.Fields("PaperAmount1").Value)
'        '
'        oExcel.Application.Cells(29, "C").Value = Choose(Val(rstBookPOChild05.Fields("PlateType1").Value), "Deepatch", "PS", "Wipeon", "CTP")
'        oExcel.Application.Cells(29, "D").Value = Val(rstBookPOChild05.Fields("TotalPlates1").Value)
'        oExcel.Application.Cells(29, "E").Value = Val(rstBookPOChild05.Fields("PlateRate1").Value)
'        oExcel.Application.Cells(29, "G").Value = Val(rstBookPOChild05.Fields("PlateAmount1").Value)
'        '
'        oExcel.Application.Cells(34, "D").Value = Val(rstBookPOChild05.Fields("Forms1").Value)
'        oExcel.Application.Cells(34, "E").Value = Val(rstBookPOChild05.Fields("PrintRate1").Value)
'        oExcel.Application.Cells(34, "G").Value = Val(rstBookPOChild05.Fields("PrintAmount1").Value)
'    End If
'    If Not IsNull(rstBookPOChild05.Fields("TextPaper2").Value) Then
'        oExcel.Application.Cells(25, "C").Value = Val(rstBookPOChild05.Fields("TextPaper2").Value)
'        oExcel.Application.Cells(25, "D").Value = Val(rstBookPOChild05.Fields("TextPaperConsumption2").Value)
'        oExcel.Application.Cells(25, "E").Value = Val(rstBookPOChild05.Fields("PaperRate2").Value)
'        oExcel.Application.Cells(25, "G").Value = Val(rstBookPOChild05.Fields("PaperAmount2").Value)
'        '
'        oExcel.Application.Cells(30, "C").Value = Choose(Val(rstBookPOChild05.Fields("PlateType2").Value), "Deepatch", "PS", "Wipeon", "CTP")
'        oExcel.Application.Cells(30, "D").Value = Val(rstBookPOChild05.Fields("TotalPlates2").Value)
'        oExcel.Application.Cells(30, "E").Value = Val(rstBookPOChild05.Fields("PlateRate2").Value)
'        oExcel.Application.Cells(30, "G").Value = Val(rstBookPOChild05.Fields("PlateAmount2").Value)
'        '
'        oExcel.Application.Cells(35, "D").Value = Val(rstBookPOChild05.Fields("Forms2").Value)
'        oExcel.Application.Cells(35, "E").Value = Val(rstBookPOChild05.Fields("PrintRate2").Value)
'        oExcel.Application.Cells(35, "G").Value = Val(rstBookPOChild05.Fields("PrintAmount2").Value)
'    End If
'    If Not IsNull(rstBookPOChild05.Fields("TextPaper4").Value) Then
'        oExcel.Application.Cells(26, "C").Value = Val(rstBookPOChild05.Fields("TextPaper4").Value)
'        oExcel.Application.Cells(26, "D").Value = Val(rstBookPOChild05.Fields("TextPaperConsumption4").Value)
'        oExcel.Application.Cells(26, "E").Value = Val(rstBookPOChild05.Fields("PaperRate4").Value)
'        oExcel.Application.Cells(26, "G").Value = Val(rstBookPOChild05.Fields("PaperAmount4").Value)
'        '
'        oExcel.Application.Cells(31, "C").Value = Choose(Val(rstBookPOChild05.Fields("PlateType4").Value), "Deepatch", "PS", "Wipeon", "CTP")
'        oExcel.Application.Cells(31, "D").Value = Val(rstBookPOChild05.Fields("TotalPlates4").Value)
'        oExcel.Application.Cells(31, "E").Value = Val(rstBookPOChild05.Fields("PlateRate4").Value)
'        oExcel.Application.Cells(31, "G").Value = Val(rstBookPOChild05.Fields("PlateAmount4").Value)
'        '
'        oExcel.Application.Cells(36, "D").Value = Val(rstBookPOChild05.Fields("Forms4").Value)
'        oExcel.Application.Cells(36, "E").Value = Val(rstBookPOChild05.Fields("PrintRate4").Value)
'        oExcel.Application.Cells(36, "G").Value = Val(rstBookPOChild05.Fields("PrintAmount4").Value)
'    End If
'    If Not IsNull(rstBookPOChild05.Fields("TitlePaper").Value) Then
'        oExcel.Application.Cells(27, "C").Value = Val(rstBookPOChild05.Fields("TitlePaper").Value)
'        oExcel.Application.Cells(27, "D").Value = Val(rstBookPOChild05.Fields("TitlePaperConsumption").Value)
'        oExcel.Application.Cells(27, "E").Value = Val(rstBookPOChild05.Fields("TitlePaperRate").Value)
'        oExcel.Application.Cells(27, "G").Value = Val(rstBookPOChild05.Fields("TitlePaperAmount").Value)
'        '
'        oExcel.Application.Cells(32, "B").Value = rstBookPOChild05.Fields("TitlePrintingType").Value & " Color"
'        oExcel.Application.Cells(32, "C").Value = Choose(Val(rstBookPOChild05.Fields("TitlePlateType").Value), "Deepatch", "PS", "Wipeon", "CTP")
'        oExcel.Application.Cells(32, "D").Value = Val(rstBookPOChild05.Fields("TitleTotalPlates").Value)
'        oExcel.Application.Cells(32, "E").Value = Val(rstBookPOChild05.Fields("TitlePlateRate").Value)
'        oExcel.Application.Cells(32, "G").Value = Val(rstBookPOChild05.Fields("TitlePlateAmount").Value)
'        '
'        oExcel.Application.Cells(37, "B").Value = rstBookPOChild05.Fields("TitlePrintingType").Value & " Color"
'        oExcel.Application.Cells(37, "E").Value = Val(rstBookPOChild05.Fields("TitlePrintRate").Value)
'        oExcel.Application.Cells(37, "G").Value = Val(rstBookPOChild05.Fields("TitlePrintAmount").Value)
'    End If
'    If Not IsNull(rstBookPOChild05.Fields("ComboPaper").Value) Then
'        oExcel.Application.Cells(28, "C").Value = Val(rstBookPOChild05.Fields("ComboPaper").Value)
'        oExcel.Application.Cells(28, "D").Value = Val(rstBookPOChild05.Fields("ComboPaperConsumption").Value)
'        oExcel.Application.Cells(28, "E").Value = Val(rstBookPOChild05.Fields("ComboPaperRate").Value)
'        oExcel.Application.Cells(28, "G").Value = Val(rstBookPOChild05.Fields("ComboPaperAmount").Value)
'        '
'        oExcel.Application.Cells(33, "B").Value = rstBookPOChild05.Fields("ComboPrintingType").Value & " Color"
'        oExcel.Application.Cells(33, "C").Value = Choose(Val(rstBookPOChild05.Fields("ComboPlateType").Value), "Deepatch", "PS", "Wipeon", "CTP")
'        oExcel.Application.Cells(33, "D").Value = Val(rstBookPOChild05.Fields("ComboTotalPlates").Value)
'        oExcel.Application.Cells(33, "E").Value = Val(rstBookPOChild05.Fields("ComboPlateRate").Value)
'        oExcel.Application.Cells(33, "G").Value = Val(rstBookPOChild05.Fields("ComboPlateAmount").Value)
'        '
'        oExcel.Application.Cells(38, "B").Value = rstBookPOChild05.Fields("ComboPrintingType").Value & " Color"
'        oExcel.Application.Cells(38, "E").Value = rstBookPOChild05.Fields("ComboPrintRate").Value
'        oExcel.Application.Cells(38, "G").Value = Val(rstBookPOChild05.Fields("ComboPrintAmount").Value)
'    End If
'    '
'    oExcel.Application.Cells(40, "E").Value = Val(rstBookPOChild05.Fields("FormFoldRate").Value)
'    oExcel.Application.Cells(41, "E").Value = Val(rstBookPOChild05.Fields("FormStitchRate").Value)
'    oExcel.Application.Cells(42, "E").Value = Val(rstBookPOChild05.Fields("FormPasteRate").Value) / 1000
'    oExcel.Application.Cells(43, "G").Value = Val(rstBookPOChild05.Fields("Packing&Cartage").Value)
'    oExcel.Application.Cells(44, "E").Value = Val(rstBookPOChild05.Fields("Rate/Book").Value)
'    If rstBookPOChild06.RecordCount = 0 Then
'        oExcel.Application.Cells(45, "D").Value = Val(rstBookPOChild05.Fields("Adjustment").Value): oExcel.Application.Cells(45, "F").Value = Val(rstBookPOChild05.Fields("GST").Value)
'    Else
'        oExcel.Application.Cells(39, "G").Value = Val(rstBookPOChild06.Fields("BillAmount").Value)
'        oExcel.Application.Cells(45, "D").Value = Val(rstBookPOChild05.Fields("Adjustment").Value) + Val(rstBookPOChild06.Fields("Adjustment").Value): oExcel.Application.Cells(45, "F").Value = Val(rstBookPOChild05.Fields("GST").Value) + Val(rstBookPOChild06.Fields("GST").Value)
'    End If
    Screen.MousePointer = vbHourglass
    If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    rstBookPOChild05.Open "SELECT E.PrintName As EName,O.PrintName As OName,Number,Quantity,Rate,M.PrintName As MName,Amount FROM (((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN GeneralMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code WHERE P.Code='" & OrderCode & "' ORDER BY E.PrintName,O.PrintName", cnDatabase, adOpenKeyset, adLockReadOnly
    Screen.MousePointer = vbNormal
    With rstBookPOChild05
        If .RecordCount > 0 Then
            i = 14
'            Do While Not .EOF
'                oExcel.Application.Cells(i, "J").Value = .Fields("EName").Value
'                oExcel.Application.Cells(i, "L").Value = .Fields("OName").Value
'                oExcel.Application.Cells(i, "O").Value = Val(.Fields("Number").Value)
'                oExcel.Application.Cells(i, "P").Value = Val(.Fields("Quantity").Value)
'                oExcel.Application.Cells(i, "Q").Value = Val(.Fields("Rate").Value)
'                oExcel.Application.Cells(i, "R").Value = .Fields("MName").Value
'                oExcel.Application.Cells(i, "S").Value = .Fields("Amount").Value
'                .MoveNext
'                i = i + 1
'            Loop
        End If
    End With
'    oExcel.Sheets("Sheet1").Protect ("eisi")
    oExcel.Workbooks.Item(1).Save
    If OutputTo = "S" Then
        oExcel.Range("A1").Activate
        oExcel.Application.Visible = True
    Else
        oExcel.Workbooks.Item(1).PrintOut
    End If
    Set oExcel = Nothing
    On Error GoTo 0
End Sub
Private Function UpdateLastPrinterBinder() As Boolean
    If BookPOType = "FP" Then
        If DatabaseType = "MS SQL" Then
            cnBookPrintOrder.Execute "UPDATE BookMaster SET BookPrinter=T.BookPrinter FROM BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book WHERE M.Code='" & ItemCode & "' AND T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book=M.Code ORDER BY P.Code DESC)"
        Else
            cnBookPrintOrder.Execute "UPDATE BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book SET M.BookPrinter=T.BookPrinter WHERE M.Code='" & ItemCode & "' AND T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book=M.Code ORDER BY P.Code DESC)"
        End If
        rstBookList.Fields("BookPrinter").Value = BookPrinterCode
        rstBookList.Update
        If DatabaseType = "MS SQL" Then
            cnBookPrintOrder.Execute "UPDATE BookMaster SET TitlePrinter=T.TitlePrinter FROM BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book WHERE M.Code='" & ItemCode & "' AND T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book=M.Code ORDER BY P.Code DESC)"
        Else
            cnBookPrintOrder.Execute "UPDATE BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book SET M.TitlePrinter=T.TitlePrinter WHERE M.Code='" & ItemCode & "' AND T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book=M.Code ORDER BY P.Code DESC)"
        End If
        rstBookList.Fields("TitlePrinter").Value = TitlePrinterCode
        rstBookList.Update
        If DatabaseType = "MS SQL" Then
            cnBookPrintOrder.Execute "UPDATE BookMaster SET Laminator=T.Laminator FROM BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book WHERE M.Code='" & ItemCode & "' AND T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book=M.Code ORDER BY P.Code DESC)"
        Else
            cnBookPrintOrder.Execute "UPDATE BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book SET M.Laminator=T.Laminator WHERE M.Code='" & ItemCode & "' AND T.Code=(SELECT DISTINCT TOP 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book=M.Code ORDER BY P.Code DESC)"
        End If
        rstBookList.Fields("Laminator").Value = LaminatorCode
        rstBookList.Update
        If DatabaseType = "MS SQL" Then
            cnBookPrintOrder.Execute "UPDATE BookMaster SET BinderFresh=T.Binder FROM BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book WHERE M.Code='" & ItemCode & "' AND T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book=M.Code ORDER BY P.Code DESC)"
        Else
            cnBookPrintOrder.Execute "UPDATE BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book SET M.BinderFresh=T.Binder WHERE M.Code='" & ItemCode & "' AND T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book=M.Code ORDER BY P.Code DESC)"
        End If
        rstBookList.Fields("BinderFresh").Value = BinderCode
        rstBookList.Update
    ElseIf BookPOType = "RP" Then
        If DatabaseType = "MS SQL" Then
            cnBookPrintOrder.Execute "UPDATE BookMaster SET BinderRepair=T.Binder FROM BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book WHERE M.Code='" & ItemCode & "' AND T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book=M.Code ORDER BY P.Code DESC)"
        Else
            cnBookPrintOrder.Execute "UPDATE BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book SET M.BinderRepair=T.Binder WHERE M.Code='" & ItemCode & "' AND T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book=M.Code ORDER BY P.Code DESC)"
        End If
        rstBookList.Fields("BinderRepair").Value = BinderCode
        rstBookList.Update
    End If
End Function
Private Sub CheckCorrections()
    If Not blnRecordExist Then
        If rstCorrections.State = adStateOpen Then rstCorrections.Close
        If DatabaseType = "MS SQL" Then
            rstCorrections.Open "SELECT ArrivedOn,Correction,RectifiedOn,SNo FROM BookChild02 WHERE Code='" & ItemCode & "' AND (RectifiedOn='' OR ISNULL(RectifiedOn,0)<>0) ORDER BY SNo", cnBookPrintOrder, adOpenKeyset, adLockReadOnly
        Else
            rstCorrections.Open "SELECT ArrivedOn,Correction,RectifiedOn,SNo FROM BookChild02 WHERE Code='" & ItemCode & "' AND (RectifiedOn='' OR ISNULL(RectifiedOn)<>0) ORDER BY SNo", cnBookPrintOrder, adOpenKeyset, adLockReadOnly
        End If
        rstCorrections.ActiveConnection = Nothing
        If rstCorrections.RecordCount > 0 Then
            Dim i As Integer
            Load FrmCorrectionRegister
            With FrmCorrectionRegister
                .Text2.Text = Text3.Text
                .fpSpread3.ClearRange 1, 1, .fpSpread3.MaxCols, .fpSpread3.MaxRows, True
                i = 1
                rstCorrections.MoveFirst
                Do While Not rstCorrections.EOF
                    .fpSpread3.SetText 1, i, IIf(CheckEmpty(rstCorrections.Fields("RectifiedOn").Value, False), 0, 1)
                    .fpSpread3.SetText 2, i, Trim(rstCorrections.Fields("Correction").Value)
                    .fpSpread3.SetText 3, i, Format(rstCorrections.Fields("ArrivedOn").Value, "dd-mm-yyyy")
                    .fpSpread3.SetText 4, i, Format(rstCorrections.Fields("SNo").Value, "###########0")
                    i = i + 1
                    rstCorrections.MoveNext
                Loop
                .fpSpread3.SetActiveCell 1, 1
                .Show vbModal
            End With
        End If
    End If
End Sub
Private Sub LockFields(ByVal bVal As Boolean)
    Dim O As Object
    For Each O In Me
        If TypeName(O) = "TextBox" Then O.Locked = bVal
    Next
End Sub
Private Sub DisplayMenu()
    Dim menusel As String
    If rstBookPOList.RecordCount = 0 Then Exit Sub
    menusel = DisplayPopupMenu(Me.hwnd, 1)
    Select Case menusel
        Case 1
            PrintBookPrintOrder02 rstBookPOList.Fields("Code").Value, , , "BP", BookPOType
        Case 2
            PrintBookPrintOrder02 rstBookPOList.Fields("Code").Value, , , "TP", BookPOType
        Case 26
            PrintBookPrintOrder02 rstBookPOList.Fields("Code").Value, , , "BP", BookPOType
        Case 3
            PrintBookPrintOrder02 rstBookPOList.Fields("Code").Value, , , "TL", BookPOType
        Case 25
            PrintBookPrintOrder02 rstBookPOList.Fields("Code").Value, , , "TP", BookPOType
        Case 24
            PrintBookPrintOrder02 rstBookPOList.Fields("Code").Value, , , "CB", BookPOType
        Case 4
            PrintBookPrintOrder02 rstBookPOList.Fields("Code").Value, , , "BB", BookPOType
        Case 5
            PrintBookPrintOrder02 rstBookPOList.Fields("Code").Value, , , "ALL", BookPOType
        Case 6
            PrintBookPrintOrder01 rstBookPOList.Fields("Code").Value, , , "JUC", BookPOType
        Case 7
            PrintBookPrintOrder01 rstBookPOList.Fields("Code").Value, , , "UC", BookPOType
        Case 8
            PrintTitlePrintingOrder (rstBookPOList.Fields("Code").Value), , , BookPOType
        Case 9
            PrintBookPrintOrder03 rstBookPOList.Fields("Code").Value, , , "BP", BookPOType
        Case 11
            PrintBookPrintOrder03 rstBookPOList.Fields("Code").Value, , , "ALL", BookPOType
        Case 12
            PrintBookPrintOrder03 rstBookPOList.Fields("Code").Value, , , "TP", BookPOType
        Case 13
            PrintTitlePlateOrder (rstBookPOList.Fields("Code").Value), , , BookPOType
        Case 14
            JobCard rstBookPOList.Fields("Code").Value, , , "BP", BookPOType   '1
        Case 15
            JobCard rstBookPOList.Fields("Code").Value, , , "TP", BookPOType   '2
        Case 16
            JobCard rstBookPOList.Fields("Code").Value, , , "TL", BookPOType   '3
        Case 17
            JobCard rstBookPOList.Fields("Code").Value, , , "BB", BookPOType   '4
        Case 18
            JobCard rstBookPOList.Fields("Code").Value, , , "ALL", BookPOType  '5
        Case 19
            JobCard rstBookPOList.Fields("Code").Value, , , "CB", BookPOType  '6
        Case 20
            PaperSlip rstBookPOList.Fields("Code").Value, , , "BP", BookPOType   '1-14
        Case 21
            PaperSlip rstBookPOList.Fields("Code").Value, , , "TP", BookPOType   '2-15
        Case 22
            PaperSlip rstBookPOList.Fields("Code").Value, , , "ALL", BookPOType  '5-22
        Case 23
            PaperSlip rstBookPOList.Fields("Code").Value, , , "CB", BookPOType   '19
        Case 27
            PrintQuotationFormat rstBookPOList.Fields("Code").Value, , , "BP", BookPOType
        Case 28
            PrintQuotationFormat rstBookPOList.Fields("Code").Value, , , "TP", BookPOType
        Case 29
            PrintQuotationFormat rstBookPOList.Fields("Code").Value, , , "TL", BookPOType
        Case 30
            PrintQuotationFormat rstBookPOList.Fields("Code").Value, , , "CB", BookPOType
        Case 31
            PrintQuotationFormat rstBookPOList.Fields("Code").Value, , , "BB", BookPOType
        Case 32
            PrintQuotationFormat rstBookPOList.Fields("Code").Value, , , "ALL", BookPOType
        Case 33
            PrintCostSheet rstBookPOList.Fields("Code").Value
        Case 34
            PrintPlanning rstBookPOList.Fields("Code").Value
        End Select
    If Not (rstBookPOList.EOF Or rstBookPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.SetFocus
End Sub
Private Sub DuplicateRecord()
    Dim OrderSubType(1 To 6) As Boolean, ConvertTo As String, ConvertOrderType As String
    Load frmJobWorkOrderDuplication
    frmJobWorkOrderDuplication.Show vbModal
    If frmJobWorkOrderDuplication.Tag = "Cancel" Then Call CloseForm(frmJobWorkOrderDuplication): Exit Sub
    ConvertTo = Choose(frmJobWorkOrderDuplication.Combo3.ListIndex + 1, "P", "S", "E") 'Purchase/Sale/Estimation
    ConvertOrderType = IIf(ConvertTo = "E", "OP", IIf(Left(BookPOType, 1) = "O", "F", Left(BookPOType, 1)) + ConvertTo)
    Dim i As Integer
    For i = 1 To 6
        OrderSubType(i) = frmJobWorkOrderDuplication.ListView1.ListItems(i).Checked
    Next i
    Call CloseForm(frmJobWorkOrderDuplication)
    Dim TmpTbl As String
    TmpTbl = "T" & GetFileNameFromPath(GetTemporaryFileName()): TmpTbl = Left(TmpTbl, InStr(1, TmpTbl, ".", vbTextCompare) - 1)
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    Dim VchCode As String, VchNo As String
    VchCode = GenerateCode(cnBookPrintOrder, "SELECT MAX(Code) FROM BookPOParent", 6, "0")
    VchNo = GenerateCode(cnBookPrintOrder, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,Name))", "VAL(Name))") & "  FROM BookPOParent WHERE Type='" & ConvertOrderType & "' AND LEFT(Code,1)<>'*' AND FYCode='" & FYCode & "'", 10, Space(1))
    cnBookPrintOrder.BeginTrans
    cnBookPrintOrder.Execute "SELECT * INTO [" & TmpTbl & "] FROM BookPOParent Where Code = '" & rstBookPOList.Fields("Code").Value & "'"
    cnBookPrintOrder.Execute "UPDATE [" & TmpTbl & "] SET Code='" & VchCode & "',Name='" & Pad(Trim(VchNo), Space(1), 10, "L") & "',[Date]=" & IIf(DatabaseType = "MS SQL", "GETDATE()", "NOW()") & ",DeliveredQuantityC=0,DeliveredQuantityB=0,Type='" & ConvertOrderType & "',BPODStatus=0,TPODStatus=0,TLODStatus=0,BBODStatus=0,UnitRate=0,[UnitRate-MF]=0,[UnitRate-SF]=0,[UnitRate-CF]=0,[UnitRate-MO]=0,[UnitRate-BP]=0,BilledAllC=0,BilledAllB=0,[UnitRate-BOM]=0"
    cnBookPrintOrder.Execute "INSERT INTO BookPOParent SELECT * FROM " & TmpTbl
    cnBookPrintOrder.Execute "DROP TABLE " & TmpTbl
    If OrderSubType(1) Then 'Multi Form Format
        cnBookPrintOrder.Execute "SELECT * INTO [" & TmpTbl & "] FROM BookPOChild05 Where Code = '" & rstBookPOList.Fields("Code").Value & "'"
        cnBookPrintOrder.Execute "UPDATE [" & TmpTbl & "] SET Code='" & VchCode & "',OrderDate=" & IIf(DatabaseType = "MS SQL", "GETDATE()", "NOW()") & ",TargetDate=" & IIf(DatabaseType = "MS SQL", "GETDATE()", "NOW()") & "+2,Ref='',BillNo='',BillDate=Null,PBillNo='',PBillDate=Null,PaidAmount=0,PPaidAmount=0,Status='',DeliveredQuantityC=0,DeliveredQuantityB=0,BilledMFC=0,BilledMFB=0"
        cnBookPrintOrder.Execute "INSERT INTO BookPOChild05 SELECT * FROM " & TmpTbl
        cnBookPrintOrder.Execute "DROP TABLE " & TmpTbl
    End If
    If OrderSubType(2) Then 'Spread Format
        cnBookPrintOrder.Execute "SELECT * INTO [" & TmpTbl & "] FROM BookPOChild06 Where Code = '" & rstBookPOList.Fields("Code").Value & "'"
        cnBookPrintOrder.Execute "UPDATE [" & TmpTbl & "] SET Code='" & VchCode & "',OrderDate=" & IIf(DatabaseType = "MS SQL", "GETDATE()", "NOW()") & ",TargetDate=" & IIf(DatabaseType = "MS SQL", "GETDATE()", "NOW()") & "+2,Ref='',BillNo='',BillDate=Null,PBillNo='',PBillDate=Null,PaidAmount=0,PPaidAmount=0,Status='',DeliveredQuantityC=0,DeliveredQuantityB=0,BilledMEC=0,BilledMEB=0"
        cnBookPrintOrder.Execute "INSERT INTO BookPOChild06 SELECT * FROM " & TmpTbl
        cnBookPrintOrder.Execute "DROP TABLE " & TmpTbl
    End If
    If OrderSubType(3) Then 'Combo Format
        cnBookPrintOrder.Execute "SELECT * INTO [" & TmpTbl & "] FROM BookPOChild09 Where Code = '" & rstBookPOList.Fields("Code").Value & "'"
        cnBookPrintOrder.Execute "UPDATE [" & TmpTbl & "] SET Code='" & VchCode & "',OrderDate=" & IIf(DatabaseType = "MS SQL", "GETDATE()", "NOW()") & ",TargetDate=" & IIf(DatabaseType = "MS SQL", "GETDATE()", "NOW()") & "+2,BillNo='',BillDate=Null,PBillNo='',PBillDate=Null,PaidAmount=0,PPaidAmount=0,Status='',BillFeedDate=Null"
        cnBookPrintOrder.Execute "INSERT INTO BookPOChild09 SELECT * FROM " & TmpTbl
        cnBookPrintOrder.Execute "DROP TABLE " & TmpTbl
        cnBookPrintOrder.Execute "SELECT * INTO [" & TmpTbl & "] FROM BookPOChild0901 Where Code = '" & rstBookPOList.Fields("Code").Value & "'"
        cnBookPrintOrder.Execute "UPDATE [" & TmpTbl & "] SET Code='" & VchCode & "',DeliveredQuantityC=0,DeliveredQuantityB=0,BilledCFC=0,BilledCFB=0"
        cnBookPrintOrder.Execute "INSERT INTO BookPOChild0901 SELECT * FROM " & TmpTbl
        cnBookPrintOrder.Execute "DROP TABLE " & TmpTbl
    End If
    If OrderSubType(4) Then 'Misc Operations
        cnBookPrintOrder.Execute "SELECT * INTO [" & TmpTbl & "] FROM BookPOChild07 Where Code = '" & rstBookPOList.Fields("Code").Value & "'"
        cnBookPrintOrder.Execute "UPDATE [" & TmpTbl & "] SET Code='" & VchCode & "',OrderDate=" & IIf(DatabaseType = "MS SQL", "GETDATE()", "NOW()") & ",TargetDate=" & IIf(DatabaseType = "MS SQL", "GETDATE()", "NOW()") & "+2,BillNo='',BillDate=Null,PaidAmount=0,Status='',DeliveredQuantityC=0,DeliveredQuantityB=0,BilledMOC=0,BilledMOB=0"
        cnBookPrintOrder.Execute "INSERT INTO BookPOChild07 SELECT * FROM " & TmpTbl
        cnBookPrintOrder.Execute "DROP TABLE " & TmpTbl
    End If
    If OrderSubType(5) Then 'Binding Process
        cnBookPrintOrder.Execute "SELECT * INTO [" & TmpTbl & "] FROM BookPOChild08 Where Code = '" & rstBookPOList.Fields("Code").Value & "'"
        cnBookPrintOrder.Execute "UPDATE [" & TmpTbl & "] SET Code='" & VchCode & "',OrderDate=" & IIf(DatabaseType = "MS SQL", "GETDATE()", "NOW()") & ",TargetDate=" & IIf(DatabaseType = "MS SQL", "GETDATE()", "NOW()") & "+2,BillNo='',BillDate=Null,PaidAmount=0,Status='',DNDetails='',CNDetails='',BillFeedDate=Null,DeliveredQuantityC=0,DeliveredQuantityB=0,BilledBNC=0,BilledBNB=0"
        cnBookPrintOrder.Execute "INSERT INTO BookPOChild08 SELECT * FROM " & TmpTbl
        cnBookPrintOrder.Execute "DROP TABLE " & TmpTbl
    End If
    If OrderSubType(6) Then 'BOM
        cnBookPrintOrder.Execute "SELECT * INTO [" & TmpTbl & "] FROM BookPOChild0801 Where Code = '" & rstBookPOList.Fields("Code").Value & "'"
        cnBookPrintOrder.Execute "UPDATE [" & TmpTbl & "] SET Code='" & VchCode & "',DeliveredQuantityC=0,DeliveredQuantityB=0,BilledBMC=0,BilledBMB=0"
        cnBookPrintOrder.Execute "INSERT INTO BookPOChild0801 SELECT * FROM " & TmpTbl
        cnBookPrintOrder.Execute "DROP TABLE " & TmpTbl
    End If
    MdiMainMenu.MousePointer = vbNormal
    Call MsgBox("Successfully Duplicated the Record !", vbInformation, App.Title)
    cnBookPrintOrder.CommitTrans
    CloseForm Me
    FrmBookPrintOrder.BookPOType = ConvertOrderType
    Load FrmBookPrintOrder
    FrmBookPrintOrder.Text1 = Trim(VchNo)
    If Err.Number <> 364 Then FrmBookPrintOrder.Show
    FrmBookPrintOrder.Toolbar1_ButtonClick FrmBookPrintOrder.Toolbar1.Buttons.Item(2)
    Exit Sub
ErrorHandler:
    MdiMainMenu.MousePointer = vbNormal
    DisplayError ("Failed to Duplicate the Record")
    cnBookPrintOrder.RollbackTrans
End Sub
Private Function chkBilled() As Boolean
    On Error GoTo ErrHandler
    Dim rstchkBill As New ADODB.Recordset
    rstchkBill.Open "SELECT C.Code FROM JobworkBVChild C INNER Join JobworkBVParent P ON P.Code=C.Code WHERE Ref='" & rstBookPOList.Fields("Code").Value & "' And (RefCode='' Or  RefCode='XXXXXX')", cnBookPrintOrder, adOpenKeyset, adLockReadOnly
    'rstchkBill.Open "SELECT C.Code FROM JobworkBVChild C INNER Join JobworkBVParent P ON P.Code=C.Code WHERE Ref='" & rstBookPOList.Fields("Code").Value & "' And Right(Type,2)='FI' OR Right(Type,2)='SJ' OR Right(Type,2)='SC' OR Right(Type,2)='SU' OR Right(Type,2)='PU' OR Right(Type,2)='PJ' OR Right(Type,2)='PC'", cnBookPrintOrder, adOpenKeyset, adLockReadOnly
    If rstchkBill.RecordCount > 0 Then chkBilled = True
    Call CloseRecordset(rstchkBill)
    Exit Function
ErrHandler:
    Call CloseRecordset(rstchkBill)
End Function
Private Sub LoadMasterList()
    If rstBookList.State = adStateOpen Then rstBookList.Close
    rstBookList.Open "SELECT M1.Name As Col0,M2.Name As SizeName,M2.Code As SizeCode,TitleSize As TitleSizeCode,FinishSize,FormType,Forms,Pages,OneColorPages,TwoColorPages,FourColorPages,[OneColor¼Forms],[OneColor½Forms],[OneColor1F/BForms],[OneColor1W/TForms],OneColorForms,[TwoColor¼Forms],[TwoColor½Forms],[TwoColor1F/BForms],[TwoColor1W/TForms],TwoColorForms,[FourColor¼Forms],[FourColor½Forms],[FourColor1F/BForms],[FourColor1W/TForms],FourColorForms,OneColorPlateType,TwoColorPlateType,FourColorPlateType,DuplexPrinting,BindingType,LaminationType,TitlePlateType,BindingForms01,BindingForms02,TitleFrontColor,TitleBackColor,TitlePlateType,[Qty/Pkt],[Pkt/Box],[LooseQty/Box],AddOnRate01,AddOnRate02,BookPrinter,TitlePrinter,Laminator,BinderFresh,BinderRepair,M1.Type,M1.Code From BookMaster M1,GeneralMaster M2 Where M1.[Size] = M2.Code Order by M1.Name", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookList.ActiveConnection = Nothing
    If Not CheckEmpty(ItemCode, False) Then rstBookList.Find "[Code]='" & ItemCode & "'"
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "SELECT Name As Col0,RoundOffQty,Code From AccountMaster Order by Name", cnBookPrintOrder, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    If rstMaterialCentreList.State = adStateOpen Then rstMaterialCentreList.Close
    rstMaterialCentreList.Open "SELECT Name As Col0,Code From AccountMaster WHERE [Group]='*99999' ORDER BY Name", cnBookPrintOrder, adOpenKeyset, adLockReadOnly
    rstMaterialCentreList.ActiveConnection = Nothing
End Sub
Private Function Save2Master() As Boolean
    Save2Master = True
    On Error GoTo ErrHandler
'    cnBookPrintOrder.Execute "DELETE FROM BookChild07 WHERE Type='" & BookPOType & "' AND Code='" & ItemCode & "'"
'    cnBookPrintOrder.Execute "INSERT INTO BookChild07 SELECT TOP 1 Book As Code,Element,Operation,[Number],OperationCountName,[Size],CalcMode,CalcValue,P.Type FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book='" & ItemCode & "' ORDER BY P.Code DESC"
'    cnBookPrintOrder.Execute "DELETE FROM BookChild08 WHERE Type='" & BookPOType & "' AND Code='" & ItemCode & "'"
'    cnBookPrintOrder.Execute "INSERT INTO BookChild08 SELECT Book As Code,ElementGroup,BindingType,BinderyProcess,[Number],OperationCountName,[Size],CalcMode,CalcValue,P.Type FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book='" & ItemCode & "' ORDER BY P.Code DESC"
'    cnBookPrintOrder.Execute "DELETE FROM BookChild06 WHERE Type='" & BookPOType & "' AND Code='" & ItemCode & "'"
'    cnBookPrintOrder.Execute "INSERT INTO BookChild06 SELECT TOP 1 Book As Code,Element,Pages,[FinishSize],[Size],Imposition,FrontPrintingType,BackPrintingType,PlateType,PlateTypeBack,[Ups],P.Type FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type='" & BookPOType & "' AND P.Book='" & ItemCode & "' ORDER BY P.Code DESC"
    cnBookPrintOrder.Execute "INSERT INTO BookChild05 SELECT Book As Code,Book As Code,Element,[FinishSize],[Size],DuplexPrinting,[Pages/PrintingForm],[Pages/Form],Color,Pages,Forms, [Forms-¼],[Forms-½],[Forms-1-F&B],[Forms-1-W&T],PlateType,[Forms/Sheet1],0,P.Type,1 FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & rstBookPOParent.Fields("Code").Value & "' AND Element NOT IN (SELECT Element FROM BookChild05 WHERE Code=P.Book)"
    cnBookPrintOrder.Execute "INSERT INTO BookChild06 SELECT Book AsCode,Book As Code,Element,'',Pages,[FinishSize],[Size],Imposition,FrontPrintingType,BackPrintingType,PlateType,PlateTypeBack,[Ups],Sets,0,P.Type FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & rstBookPOParent.Fields("Code").Value & "' AND Element NOT IN (SELECT Element FROM BookChild06 WHERE Code=P.Book)"
    cnBookPrintOrder.Execute "INSERT INTO BookChild07 SELECT Book As Code,Book As Code,Element,Operation,[Number],OperationCountName,[Size],CalcMode,CalcValue,P.Type FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Code='" & rstBookPOParent.Fields("Code").Value & "' AND Element+Operation+Type NOT IN (SELECT Element+Operation+Type FROM BookChild07 WHERE Code=P.Book)"
    cnBookPrintOrder.Execute "INSERT INTO BookChild08 SELECT Book As Code,Book As Code,BindingType,BinderyProcess,[Number],OperationCountName,[Size],CalcMode,CalcValue,P.Type FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Code='" & rstBookPOParent.Fields("Code").Value & "' AND SubItem+BinderyProcess+Type NOT IN (SELECT SubItem+BinderyProcess+Type FROM BookChild08 WHERE Code=P.Book)"
    Exit Function
ErrHandler:
    Save2Master = False
End Function
Private Sub MhRealInput8_Validate(Cancel As Boolean)
    If MhRealInput3.Value = 0 Then Exit Sub
    Dim PlateAmount As Double, OthersAmount As Double, OthersUnitRate As Double, Amount As Double
    MhRealInput19.Value = 0: MhRealInput21.Value = 0: MhRealInput23.Value = 0: MhRealInput25.Value = 0: MhRealInput27.Value = 0: MhRealInput31.Value = 0
    MhRealInput20.Value = 0: MhRealInput22.Value = 0: MhRealInput24.Value = 0: MhRealInput26.Value = 0: MhRealInput28.Value = 0: MhRealInput32.Value = 0
'    With rstBookPOChild05
'        If .RecordCount <> 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                Amount = Val(.Fields("PlateAmount1").Value) + Val(.Fields("PlateAmount2").Value) + Val(.Fields("PlateAmount4").Value) + Val(.Fields("PAdjustment").Value) 'Plate Amount
'                PlateAmount = PlateAmount + Amount
'                MhRealInput20.Value = MhRealInput20.Value + Amount
'                Amount = Val(.Fields("PaperAmount1").Value) + Val(.Fields("PaperAmount2").Value) + Val(.Fields("PaperAmount4").Value) + Val(.Fields("RAdjustment").Value) 'Paper Amount
'                OthersAmount = OthersAmount + Amount
'                MhRealInput20.Value = MhRealInput20.Value + Amount
'                Amount = Val(.Fields("PrintAmount1").Value) + Val(.Fields("PrintAmount2").Value) + Val(.Fields("PrintAmount4").Value) + Val(.Fields("Adjustment").Value) 'Print Amount
'                OthersAmount = OthersAmount + Amount
'                MhRealInput20.Value = MhRealInput20.Value + Amount
'                .MoveNext
'            Loop
'            MhRealInput19.Value = MhRealInput20.Value / MhRealInput3.Value
'            .MoveFirst
'        End If
'    End With
    With rstBookPOChild05
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
                Amount = Val(.Fields("PlateAmount").Value) + Val(.Fields("PAdjustment").Value) 'Plate Amount
                PlateAmount = PlateAmount + Amount
                MhRealInput20.Value = MhRealInput20.Value + Amount
                Amount = Val(.Fields("PaperAmount").Value) + Val(.Fields("RAdjustment").Value) 'Paper Amount
                OthersAmount = OthersAmount + Amount
                MhRealInput20.Value = MhRealInput20.Value + Amount
                Amount = Val(.Fields("PrintAmount").Value) + Val(.Fields("Adjustment").Value) 'Print Amount
                OthersAmount = OthersAmount + Amount
                MhRealInput20.Value = MhRealInput20.Value + Amount
                .MoveNext
            Loop
            MhRealInput19.Value = MhRealInput20.Value / MhRealInput3.Value
            .MoveFirst
        End If
    End With
    With rstBookPOChild06
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
                Amount = Val(.Fields("PlateAmount").Value) + Val(.Fields("PAdjustment").Value) 'Plate Amount
                PlateAmount = PlateAmount + Amount
                MhRealInput22.Value = MhRealInput22.Value + Amount
                Amount = Val(.Fields("PaperAmount").Value) + Val(.Fields("RAdjustment").Value) 'Paper Amount
                OthersAmount = OthersAmount + Amount
                MhRealInput22.Value = MhRealInput22.Value + Amount
                Amount = Val(.Fields("PrintAmount").Value) + Val(.Fields("Adjustment").Value) 'Print Amount
                OthersAmount = OthersAmount + Amount
                MhRealInput22.Value = MhRealInput22.Value + Amount
                .MoveNext
            Loop
            MhRealInput21.Value = MhRealInput22.Value / MhRealInput3.Value
            .MoveFirst
        End If
    End With
    With rstBookPOChild07
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
                MhRealInput26.Value = MhRealInput26.Value + Val(.Fields("Amount").Value) + Val(.Fields("Adjustment").Value)
                .MoveNext
            Loop
            OthersAmount = OthersAmount + MhRealInput26.Value
            MhRealInput25.Value = MhRealInput26.Value / MhRealInput3.Value
            .MoveFirst
        End If
    End With
    With rstBookPOChild08
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
                MhRealInput28.Value = MhRealInput28.Value + Val(.Fields("Amount").Value) + Val(.Fields("Adjustment").Value)
                .MoveNext
            Loop
            MhRealInput27.Value = MhRealInput28.Value / MhRealInput3.Value
            OthersAmount = OthersAmount + MhRealInput28.Value
            .MoveFirst
        End If
    End With
    With rstBookPOChild0801
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
                MhRealInput32.Value = MhRealInput32.Value + Val(.Fields("Amount").Value)
                .MoveNext
            Loop
            OthersAmount = OthersAmount + MhRealInput32.Value
            MhRealInput31.Value = MhRealInput32.Value / MhRealInput3.Value
            .MoveFirst
        End If
    End With
    With rstBookPOChild09
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
                Amount = Val(.Fields("PlateAmount").Value) + Val(.Fields("PAdjustment").Value) 'Plate Amount
                PlateAmount = PlateAmount + Amount
                MhRealInput24.Value = MhRealInput24.Value + Amount
                Amount = Val(.Fields("PaperAmount").Value) + Val(.Fields("RAdjustment").Value) 'Paper Amount
                OthersAmount = OthersAmount + Amount
                MhRealInput24.Value = MhRealInput24.Value + Amount
                Amount = Val(.Fields("PrintAmount").Value) + Val(.Fields("Adjustment").Value) 'Print Amount
                OthersAmount = OthersAmount + Amount
                MhRealInput24.Value = MhRealInput24.Value + Amount
                .MoveNext
            Loop
            MhRealInput23.Value = MhRealInput24.Value / MhRealInput3.Value
            .MoveFirst
        End If
    End With
    OthersUnitRate = OthersAmount / MhRealInput3.Value
    If MhRealInput3.Value > 0 Then 'EstQty01
        MhRealInput9.Value = (PlateAmount / MhRealInput3.Value) + OthersUnitRate
        MhRealInput9.Value = Round(MhRealInput9.Value + ((MhRealInput9.Value * MhRealInput8.Value) / 100), 3) 'Add Profit
        MhRealInput14.Value = MhRealInput9.Value * MhRealInput3.Value
    End If
    If MhRealInput4.Value > 0 Then 'EstQty02
        MhRealInput10.Value = (PlateAmount / MhRealInput4.Value) + OthersUnitRate
        MhRealInput10.Value = Round(MhRealInput10.Value + ((MhRealInput10.Value * MhRealInput8.Value) / 100), 3)
        MhRealInput15.Value = MhRealInput10.Value * MhRealInput4.Value
    End If
    If MhRealInput5.Value > 0 Then
        MhRealInput11.Value = (PlateAmount / MhRealInput5.Value) + OthersUnitRate
        MhRealInput11.Value = Round(MhRealInput11.Value + ((MhRealInput11.Value * MhRealInput8.Value) / 100), 3)
        MhRealInput16.Value = MhRealInput11.Value * MhRealInput5.Value
    End If
    If MhRealInput6.Value > 0 Then
        MhRealInput12.Value = (PlateAmount / MhRealInput6.Value) + OthersUnitRate
        MhRealInput12.Value = Round(MhRealInput12.Value + ((MhRealInput12.Value * MhRealInput8.Value) / 100), 3)
        MhRealInput17.Value = MhRealInput12.Value * MhRealInput6.Value
    End If
    If MhRealInput7.Value > 0 Then
        MhRealInput13.Value = (PlateAmount / MhRealInput7.Value) + OthersUnitRate
        MhRealInput13.Value = Round(MhRealInput13.Value + ((MhRealInput13.Value * MhRealInput8.Value) / 100), 3)
        MhRealInput18.Value = MhRealInput13.Value * MhRealInput7.Value
    End If
    If MhRealInput33.Value > 0 Then
        MhRealInput29.Value = (PlateAmount / MhRealInput33.Value) + OthersUnitRate
        MhRealInput29.Value = Round(MhRealInput29.Value + ((MhRealInput29.Value * MhRealInput8.Value) / 100), 3)
        MhRealInput30.Value = MhRealInput29.Value * MhRealInput33.Value
    End If
End Sub
Private Sub RefreshList(ByVal VchCode As String)
    With rstBookPOList
        If .State = adStateOpen Then .Close
        BusySystemIndicator True
        If DisplayListType = "O" Then
            .Open "SELECT DISTINCT T.Code,T.Name,IIF(ISNULL(C1.Ref,'')='' AND ISNULL(C2.Ref,'')='','',IIF(ISNULL(C1.Ref,'')='',C2.Ref,IIF(ISNULL(C2.Ref,'')='',C1.Ref,C1.Ref+'-'+C2.Ref))) As RefNo,Date,I.Name As BookName,BPODStatus,TPODStatus,TLODStatus,BBODStatus,T.DeliveredQuantityC+T.DeliveredQuantityB As DeliveredQuantity,XP.Name As BookPrinterName,TP.Name As TitlePrinterName,LM.Name  As LaminatorName,BD.Name As BinderName,UnitRate FROM ((((((BookPOParent T INNER JOIN BookMaster I ON T.Book=I.Code) LEFT JOIN AccountMaster XP ON XP.Code=T.BookPrinter) LEFT JOIN AccountMaster TP ON TP.Code=T.TitlePrinter) LEFT JOIN AccountMaster LM ON LM.Code=T.Laminator) LEFT JOIN AccountMaster BD ON BD.Code=T.Binder) LEFT JOIN BookPOChild05 C1 ON T.Code=C1.Code) LEFT JOIN BookPOChild06 C2 ON T.Code=C2.Code WHERE T.Type = '" & BookPOType & "' AND FYCode='" & FYCode & "' AND LEFT(T.Code,1)<>'*' ORDER BY T.Name", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
            cmdProceed.Caption = " Show Combo Item List Only"
        Else
            .Open "SELECT DISTINCT T.Code,T.Name,'' As RefNo,Date,I.Name As BookName,BPODStatus,TPODStatus,TLODStatus,BBODStatus,T.DeliveredQuantityC+T.DeliveredQuantityB As DeliveredQuantity,XP.Name As BookPrinterName,TP.Name As TitlePrinterName,LM.Name  As LaminatorName,BD.Name As BinderName,UnitRate FROM (((((BookPOParent T INNER JOIN BookPOChild0901 C3 ON T.Code=C3.Code) INNER JOIN BookMaster I ON C3.Book=I.Code) LEFT JOIN AccountMaster XP ON XP.Code=T.BookPrinter) LEFT JOIN AccountMaster TP ON TP.Code=T.TitlePrinter) LEFT JOIN AccountMaster LM ON LM.Code=T.Laminator) LEFT JOIN AccountMaster BD ON BD.Code=T.Binder WHERE T.Type = '" & BookPOType & "' AND FYCode='" & FYCode & "' AND LEFT(T.Code,1)<>'*' ORDER BY T.Name", cnBookPrintOrder, adOpenKeyset, adLockOptimistic
            cmdProceed.Caption = " Show Genaral Item List "
        End If
        .ActiveConnection = Nothing
        .Filter = adFilterNone
        .Sort = SortOrder & " Asc"
        If CheckEmpty(VchCode, False) Then
            If .RecordCount > 0 Then .MoveLast
        Else
            .Find "[Code] = '" & VchCode & "'"
            If .EOF Then .MoveLast
        End If
    End With
    Set DataGrid1.DataSource = rstBookPOList
    If Not (rstBookPOList.EOF Or rstBookPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    BusySystemIndicator False
End Sub
Private Sub cmdProceed_Click()
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        If DisplayListType = "C" Then DisplayListType = "O" Else DisplayListType = "C"
        Call RefreshList("")
End Sub
Private Sub PushVch() 'PushPO2Busy
'To Be Confirm---FI.FormatDate,FI.GetRecordset
    Dim VchSeriesName, VchDate, VchNo, STName, AccountCode, AccountName, MCName, xmlstr
    Dim ItemCode, ItemName, Qty, Price
    AccountCode = IIf(Not CheckEmpty(BookPrinterCode, False), BookPrinterCode, BinderCode)
    If CheckEmpty(AccountCode, False) Then Set FI = Nothing: Exit Sub
    VchSeriesName = IIf(Left(BookPOType, 1) = "F", "Main", "Repair"): MCName = "Noida Godown": VchNo = Trim(Text2.Text): VchDate = FI.FormatDate(MhDateInput1.Text): Qty = Val(MhRealInput3.Value)
    Set rstEasyPublish = FI.GetRecordset("SELECT Name,GSTNo FROM Master1 P INNER JOIN MasterAddressInfo C ON P.Code=C.MasterCode WHERE Code=" & AccountCode)
    AccountName = Replace(rstEasyPublish.Fields("Name").Value, "&", "&amp;", 1)
    STName = IIf(Left(rstEasyPublish.Fields("GSTNo").Value, 2) = "07", "L/GST-Exempt", "I/GST-Exempt")
    ItemCode = Mid(ItemCode, 2, 6)
    Set rstEasyPublish = FI.GetRecordset("SELECT Name,D3 As Price FROM Master1 WHERE Code=" & ItemCode)
    ItemName = Replace(rstEasyPublish.Fields("Name").Value, "&", "&amp;", 1): Price = Val(rstEasyPublish.Fields("Price").Value)
    xmlstr = "<PurchaseOrder>"
        xmlstr = xmlstr & "<VchSeriesName>" & VchSeriesName & "</VchSeriesName><Date>" & VchDate & "</Date><VchType>13</VchType><VchNo>" & VchNo & "</VchNo><STPTName>" & STName & "</STPTName><MasterName1>" & AccountName & "</MasterName1><MasterName2>" & MCName & "</MasterName2>"
        xmlstr = xmlstr & "<ItemEntries>"
        xmlstr = xmlstr & "<ItemDetail><SrNo>1</SrNo><ItemName>" & ItemName & "</ItemName><UnitName>Nos</UnitName><Qty>" & Trim(Qty) & "</Qty><QtyMainUnit>" & Trim(Qty) & "</QtyMainUnit><QtyAltUnit>" & Trim(Qty) & "</QtyAltUnit><Price>" & Trim(Price) & "</Price><Amt>" & Trim(Qty * Price) & "</Amt><STAmount>0</STAmount><STPercent>0</STPercent><TaxBeforeSurcharge>0</TaxBeforeSurcharge><MC>" & MCName & "</MC></ItemDetail>"
        xmlstr = xmlstr & "</ItemEntries>"
        xmlstr = xmlstr & "<PendingOrders>"
            xmlstr = xmlstr & "<OrderDetail><MasterName1>" & ItemName & "</MasterName1><MasterName2>" & AccountName & "</MasterName2>"
            xmlstr = xmlstr & "<OrderRefs><Method>1</Method><SrNo>1</SrNo><RefNo>" & VchNo & "</RefNo><Date>" & VchDate & "</Date><DueDate>" & VchDate & "</DueDate><Value1>" & Trim(0 - Qty) & "</Value1><Value2>" & Trim(0 - Qty) & "</Value2><ItemSrNo>1</ItemSrNo><tmpMasterCode1>" & Trim(ItemCode) & "</tmpMasterCode1><tmpMasterCode2>" & Trim(AccountCode) & "</tmpMasterCode2></OrderRefs>"
            xmlstr = xmlstr & "</OrderDetail>"
        xmlstr = xmlstr & "</PendingOrders>"
    xmlstr = xmlstr & "</PurchaseOrder>"
    If Not FI.SaveVchFromXML(13, xmlstr, Err, True, 2) Then DisplayError (Err)
End Sub

