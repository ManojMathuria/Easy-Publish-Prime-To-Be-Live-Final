VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmBookMaster 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Master"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16455
   BeginProperty Font 
      Name            =   "Comic Sans MS"
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
   ScaleHeight     =   9150
   ScaleWidth      =   16455
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9150
      Left            =   15
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   0
      Width           =   16440
      _Version        =   65536
      _ExtentX        =   28998
      _ExtentY        =   16140
      _StockProps     =   77
      BackColor       =   -2147483633
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
      Picture         =   "BookMaster.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8910
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   120
         Width           =   16200
         _ExtentX        =   28575
         _ExtentY        =   15716
         _Version        =   393216
         Style           =   1
         Tabs            =   8
         TabsPerRow      =   8
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
         TabPicture(0)   =   "BookMaster.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtNotes"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "CmdLabel"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "CmdPrint"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "CmdExport"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Mh3dLabel1(2)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "DataGrid1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Text1"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "CommonDialog1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "ImageList1"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "BookMaster.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&BOM"
         TabPicture(2)   =   "BookMaster.frx":0054
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Mh3dFrame3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "&Editorial Component"
         TabPicture(3)   =   "BookMaster.frx":0070
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Mh3dFrame5"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Multi Form Format Element"
         TabPicture(4)   =   "BookMaster.frx":008C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Mh3dFrame7"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Multi Element Format Element"
         TabPicture(5)   =   "BookMaster.frx":00A8
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Mh3dFrame4"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Miscellaneous Operation"
         TabPicture(6)   =   "BookMaster.frx":00C4
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Mh3dFrame8"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "Binding && Finishing"
         TabPicture(7)   =   "BookMaster.frx":00E0
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Mh3dFrame9"
         Tab(7).ControlCount=   1
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   6000
            Top             =   4320
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   15
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":00FC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":0640
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":0B84
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":0C98
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":0DAC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":0EC0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":101C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":1560
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":1674
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":1BB8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":1CCC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":1DE0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":1EF4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":2008
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "BookMaster.frx":211C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   6720
            Top             =   4320
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
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
            Left            =   720
            TabIndex        =   31
            Top             =   8445
            Width           =   12240
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7440
            Left            =   120
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   480
            Width           =   15960
            _ExtentX        =   28152
            _ExtentY        =   13123
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
            ColumnCount     =   12
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   "S.No"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Name"
               Caption         =   "Name"
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
               DataField       =   "PrintName"
               Caption         =   " Item Description"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "ItemGroup"
               Caption         =   "Group"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   ""
               Caption         =   "Alias"
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
            BeginProperty Column05 
               DataField       =   ""
               Caption         =   " Code"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "ISBN"
               Caption         =   " EAN/ISBN"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "Unit"
               Caption         =   " Unit"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "MFDC"
               Caption         =   " MFDC"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "PurPrice"
               Caption         =   "Pur. Price"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "SalePrice"
               Caption         =   "Sales Price"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "Price"
               Caption         =   "Cash Price"
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
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               ScrollBars      =   3
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  Alignment       =   2
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   4440.189
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   4680
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2550.047
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1695.118
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   0
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
               EndProperty
               BeginProperty Column07 
                  Alignment       =   2
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column08 
                  Alignment       =   1
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column09 
                  Alignment       =   1
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column10 
                  Alignment       =   1
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column11 
                  Alignment       =   1
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   945.071
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8340
            Left            =   -74880
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   480
            Width           =   15960
            _Version        =   65536
            _ExtentX        =   28152
            _ExtentY        =   14711
            _StockProps     =   77
            BackColor       =   -2147483633
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
            Picture         =   "BookMaster.frx":2230
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
               Left            =   5760
               MaxLength       =   255
               TabIndex        =   26
               Top             =   4200
               Visible         =   0   'False
               Width           =   2420
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
               Left            =   1680
               MaxLength       =   255
               TabIndex        =   25
               Top             =   4200
               Visible         =   0   'False
               Width           =   2420
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
               Left            =   13800
               MaxLength       =   255
               TabIndex        =   22
               Top             =   3570
               Width           =   2040
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
               Left            =   9960
               MaxLength       =   255
               TabIndex        =   17
               Top             =   2940
               Visible         =   0   'False
               Width           =   5880
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   8160
               TabIndex        =   85
               Top             =   2940
               Visible         =   0   'False
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " S. No."
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":224C
               Picture         =   "BookMaster.frx":2268
            End
            Begin VB.Frame Frame1 
               Height          =   4330
               Left            =   9960
               TabIndex        =   75
               Top             =   3900
               Width           =   5895
               Begin VB.CommandButton Last 
                  Height          =   375
                  Left            =   1200
                  Picture         =   "BookMaster.frx":2284
                  Style           =   1  'Graphical
                  TabIndex        =   82
                  TabStop         =   0   'False
                  ToolTipText     =   "Last"
                  Top             =   3840
                  Width           =   375
               End
               Begin VB.CommandButton First 
                  Height          =   375
                  Left            =   120
                  Picture         =   "BookMaster.frx":2386
                  Style           =   1  'Graphical
                  TabIndex        =   81
                  TabStop         =   0   'False
                  ToolTipText     =   "First"
                  Top             =   3840
                  Width           =   375
               End
               Begin VB.CommandButton cmdDelete 
                  Height          =   375
                  Left            =   5400
                  Picture         =   "BookMaster.frx":2488
                  Style           =   1  'Graphical
                  TabIndex        =   80
                  TabStop         =   0   'False
                  ToolTipText     =   "Delete This Item  Pic"
                  Top             =   3840
                  Width           =   375
               End
               Begin VB.CommandButton cmdUpload 
                  Height          =   375
                  Left            =   4605
                  Picture         =   "BookMaster.frx":258A
                  Style           =   1  'Graphical
                  TabIndex        =   79
                  TabStop         =   0   'False
                  ToolTipText     =   "Upload New Image"
                  Top             =   3840
                  Width           =   375
               End
               Begin VB.CommandButton cmdView 
                  Height          =   375
                  Left            =   4995
                  Picture         =   "BookMaster.frx":28CC
                  Style           =   1  'Graphical
                  TabIndex        =   78
                  TabStop         =   0   'False
                  ToolTipText     =   "View Image"
                  Top             =   3840
                  Width           =   375
               End
               Begin VB.CommandButton Previous 
                  Height          =   375
                  Left            =   480
                  Picture         =   "BookMaster.frx":2DFE
                  Style           =   1  'Graphical
                  TabIndex        =   77
                  TabStop         =   0   'False
                  ToolTipText     =   "Previous"
                  Top             =   3840
                  Width           =   375
               End
               Begin VB.CommandButton Next 
                  Height          =   375
                  Left            =   840
                  Picture         =   "BookMaster.frx":2F00
                  Style           =   1  'Graphical
                  TabIndex        =   76
                  TabStop         =   0   'False
                  ToolTipText     =   "Next"
                  Top             =   3840
                  Width           =   375
               End
               Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
                  Height          =   375
                  Left            =   1635
                  TabIndex        =   84
                  Top             =   3840
                  Width           =   2055
                  _Version        =   65536
                  _ExtentX        =   3625
                  _ExtentY        =   661
                  _StockProps     =   77
                  BackColor       =   -2147483633
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  TintColor       =   16711935
                  Caption         =   "No Image"
                  BorderStyle     =   0
                  TextColor       =   0
                  Picture         =   "BookMaster.frx":3002
                  Picture         =   "BookMaster.frx":301E
               End
               Begin MSForms.Image Image1 
                  Height          =   3495
                  Left            =   120
                  Top             =   240
                  Width           =   5655
                  BorderColor     =   9164542
                  BackColor       =   9164542
                  Size            =   "9975;6165"
                  VariousPropertyBits=   19
               End
               Begin MSForms.CheckBox CheckBox1 
                  Height          =   375
                  Left            =   3735
                  TabIndex        =   83
                  Top             =   3840
                  Width           =   855
                  VariousPropertyBits=   746588179
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  DisplayStyle    =   4
                  Size            =   "1508;661"
                  Value           =   "0"
                  Caption         =   "Zoom"
                  FontName        =   "Comic Sans MS"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin VB.Shape Shape1 
                  BorderColor     =   &H008BD6FE&
                  FillColor       =   &H008BD6FE&
                  FillStyle       =   0  'Solid
                  Height          =   3495
                  Left            =   120
                  Top             =   240
                  Width           =   5655
               End
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
               Left            =   9960
               MaxLength       =   255
               TabIndex        =   15
               Top             =   2630
               Visible         =   0   'False
               Width           =   5880
            End
            Begin MSComDlg.CommonDialog cdUpload 
               Left            =   14880
               Top             =   960
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel30 
               Height          =   330
               Left            =   8160
               TabIndex        =   73
               Top             =   2310
               Visible         =   0   'False
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Item Marks"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":303A
               Picture         =   "BookMaster.frx":3056
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
               Left            =   9960
               MaxLength       =   255
               TabIndex        =   13
               Top             =   2310
               Visible         =   0   'False
               Width           =   5880
            End
            Begin VB.CommandButton btnNotes 
               Caption         =   " Notes"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               TabIndex        =   71
               Top             =   7920
               Width           =   1575
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   980
               Left            =   120
               TabIndex        =   70
               Top             =   720
               Visible         =   0   'False
               Width           =   15720
               _Version        =   65536
               _ExtentX        =   27728
               _ExtentY        =   1729
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   9164542
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Item Description"
               FillColor       =   9164542
               ShadowColor     =   0
               TextColor       =   0
               Picture         =   "BookMaster.frx":3072
               Picture         =   "BookMaster.frx":308E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   8160
               TabIndex        =   66
               Top             =   3570
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Purchase Price"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":30AA
               Picture         =   "BookMaster.frx":30C6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   35
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "BookMaster.frx":30E2
               Picture         =   "BookMaster.frx":30FE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   330
               Left            =   120
               TabIndex        =   48
               Top             =   1055
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "BookMaster.frx":311A
               Picture         =   "BookMaster.frx":3136
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   330
               Left            =   120
               TabIndex        =   49
               Top             =   1370
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "BookMaster.frx":3152
               Picture         =   "BookMaster.frx":316E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   120
               TabIndex        =   45
               Top             =   2630
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " HSN Code"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":318A
               Picture         =   "BookMaster.frx":31A6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
               Height          =   330
               Left            =   120
               TabIndex        =   42
               Top             =   2940
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "BookMaster.frx":31C2
               Picture         =   "BookMaster.frx":31DE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   120
               TabIndex        =   40
               Top             =   1995
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":31FA
               Picture         =   "BookMaster.frx":3216
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   120
               TabIndex        =   39
               Top             =   2310
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Group"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":3232
               Picture         =   "BookMaster.frx":324E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   120
               TabIndex        =   38
               Top             =   1680
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " MRP"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":326A
               Picture         =   "BookMaster.frx":3286
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Left            =   120
               TabIndex        =   37
               Top             =   740
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "BookMaster.frx":32A2
               Picture         =   "BookMaster.frx":32BE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   34
               Top             =   425
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "BookMaster.frx":32DA
               Picture         =   "BookMaster.frx":32F6
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
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   2
               Top             =   740
               Width           =   14160
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
               Left            =   9960
               MaxLength       =   17
               TabIndex        =   11
               Top             =   1995
               Width           =   5880
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
               MaxLength       =   60
               TabIndex        =   1
               Top             =   425
               Width           =   14160
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
               TabIndex        =   0
               Top             =   105
               Width           =   14160
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   330
               Left            =   12900
               TabIndex        =   7
               TabStop         =   0   'False
               ToolTipText     =   "Binding"
               Top             =   1365
               Width           =   2945
               _Version        =   65536
               _ExtentX        =   5195
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":3312
               Caption         =   "BookMaster.frx":3332
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":339E
               Keys            =   "BookMaster.frx":33BC
               Spin            =   "BookMaster.frx":3406
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
               ForeColor       =   255
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
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
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
               MaxLength       =   255
               TabIndex        =   16
               Top             =   2940
               Width           =   14160
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   8160
               TabIndex        =   36
               Top             =   1995
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " ISBN / EAN"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":342E
               Picture         =   "BookMaster.frx":344A
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
               Height          =   330
               Left            =   1680
               TabIndex        =   8
               ToolTipText     =   "Printing Form"
               Top             =   1680
               Width           =   6495
               _Version        =   65536
               _ExtentX        =   11456
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":3466
               Caption         =   "BookMaster.frx":3486
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":34F2
               Keys            =   "BookMaster.frx":3510
               Spin            =   "BookMaster.frx":355A
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
               ValueVT         =   187498501
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
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
               TabIndex        =   10
               Top             =   1995
               Width           =   6495
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
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   3
               Top             =   1055
               Width           =   14160
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   8160
               TabIndex        =   56
               Top             =   1365
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "BookMaster.frx":3582
               Picture         =   "BookMaster.frx":359E
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
               Height          =   330
               Left            =   1680
               TabIndex        =   4
               TabStop         =   0   'False
               ToolTipText     =   "Multi Form Format Pages"
               Top             =   1365
               Width           =   3240
               _Version        =   65536
               _ExtentX        =   5715
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":35BA
               Caption         =   "BookMaster.frx":35DA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":3646
               Keys            =   "BookMaster.frx":3664
               Spin            =   "BookMaster.frx":36AE
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
               ForeColor       =   255
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
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1870069765
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
               Height          =   330
               Left            =   9960
               TabIndex        =   6
               TabStop         =   0   'False
               ToolTipText     =   "Printing"
               Top             =   1365
               Width           =   2951
               _Version        =   65536
               _ExtentX        =   5205
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":36D6
               Caption         =   "BookMaster.frx":36F6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":3762
               Keys            =   "BookMaster.frx":3780
               Spin            =   "BookMaster.frx":37CA
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
               ForeColor       =   255
               Format          =   "###0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   187498501
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
               Height          =   330
               Left            =   4905
               TabIndex        =   5
               TabStop         =   0   'False
               ToolTipText     =   "Multi Elementi Format Pages"
               Top             =   1365
               Width           =   3270
               _Version        =   65536
               _ExtentX        =   5768
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":37F2
               Caption         =   "BookMaster.frx":3812
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":387E
               Keys            =   "BookMaster.frx":389C
               Spin            =   "BookMaster.frx":38E6
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
               ForeColor       =   255
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
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1870069765
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   8160
               TabIndex        =   59
               Top             =   1680
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Weight"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":390E
               Picture         =   "BookMaster.frx":392A
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
               Height          =   330
               Left            =   9960
               TabIndex        =   9
               ToolTipText     =   "Printing Form"
               Top             =   1680
               Width           =   5880
               _Version        =   65536
               _ExtentX        =   10372
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":3946
               Caption         =   "BookMaster.frx":3966
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":39D2
               Keys            =   "BookMaster.frx":39F0
               Spin            =   "BookMaster.frx":3A3A
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
               ForeColor       =   -2147483640
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
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   188940293
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
               Height          =   330
               Left            =   120
               TabIndex        =   63
               Top             =   3260
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Integration Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":3A62
               Picture         =   "BookMaster.frx":3A7E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel29 
               Height          =   330
               Left            =   8160
               TabIndex        =   64
               Top             =   3260
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Unit"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":3A9A
               Picture         =   "BookMaster.frx":3AB6
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
               Left            =   1680
               MaxLength       =   60
               TabIndex        =   18
               Top             =   3260
               Width           =   6490
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
               Left            =   9960
               MaxLength       =   255
               TabIndex        =   19
               Top             =   3260
               Width           =   5880
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   120
               TabIndex        =   65
               Top             =   3570
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Sale Price"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":3AD2
               Picture         =   "BookMaster.frx":3AEE
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   330
               Left            =   9945
               TabIndex        =   21
               ToolTipText     =   "Printing Form"
               Top             =   3570
               Width           =   2070
               _Version        =   65536
               _ExtentX        =   3651
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":3B0A
               Caption         =   "BookMaster.frx":3B2A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":3B96
               Keys            =   "BookMaster.frx":3BB4
               Spin            =   "BookMaster.frx":3BFE
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
               ValueVT         =   188940293
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
               Height          =   330
               Left            =   1680
               TabIndex        =   20
               ToolTipText     =   "Printing Form"
               Top             =   3570
               Width           =   6495
               _Version        =   65536
               _ExtentX        =   11456
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":3C26
               Caption         =   "BookMaster.frx":3C46
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":3CB2
               Keys            =   "BookMaster.frx":3CD0
               Spin            =   "BookMaster.frx":3D1A
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
               ValueVT         =   188940293
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   2370
               Index           =   1
               Left            =   120
               TabIndex        =   72
               Top             =   5520
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   4180
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D->Delete  Ctrl+S->Save  F5-> Refresh-> F12-> Create Duplicate Item Master"
               FillColor       =   8421504
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":3D42
               Multiline       =   -1  'True
               Picture         =   "BookMaster.frx":3D5E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Left            =   8160
               TabIndex        =   74
               Top             =   2630
               Visible         =   0   'False
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " MFCD"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":3D7A
               Picture         =   "BookMaster.frx":3D96
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
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   14
               ToolTipText     =   "Finish Size"
               Top             =   2630
               Width           =   14160
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
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   12
               Top             =   2310
               Width           =   14160
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
               Height          =   330
               Left            =   120
               TabIndex        =   86
               Top             =   3880
               Visible         =   0   'False
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " OP. Stock A"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":3DB2
               Picture         =   "BookMaster.frx":3DCE
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   330
               Left            =   1680
               TabIndex        =   23
               ToolTipText     =   "Printing Form"
               Top             =   3885
               Visible         =   0   'False
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":3DEA
               Caption         =   "BookMaster.frx":3E0A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":3E76
               Keys            =   "BookMaster.frx":3E94
               Spin            =   "BookMaster.frx":3EDE
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
               ReadOnly        =   -1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   188940293
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
               Height          =   330
               Left            =   4080
               TabIndex        =   87
               Top             =   3885
               Visible         =   0   'False
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " OP. Stock B"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":3F06
               Picture         =   "BookMaster.frx":3F22
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   330
               Left            =   5760
               TabIndex        =   24
               ToolTipText     =   "Printing Form"
               Top             =   3885
               Visible         =   0   'False
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":3F3E
               Caption         =   "BookMaster.frx":3F5E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":3FCA
               Keys            =   "BookMaster.frx":3FE8
               Spin            =   "BookMaster.frx":4032
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
               ReadOnly        =   -1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   188940293
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
               Height          =   330
               Left            =   12000
               TabIndex        =   88
               Top             =   3570
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Purchase Date"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":405A
               Picture         =   "BookMaster.frx":4076
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
               Height          =   330
               Left            =   120
               TabIndex        =   89
               Top             =   4200
               Visible         =   0   'False
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Room A"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":4092
               Picture         =   "BookMaster.frx":40AE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
               Height          =   330
               Left            =   4080
               TabIndex        =   90
               Top             =   4200
               Visible         =   0   'False
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Room B"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "BookMaster.frx":40CA
               Picture         =   "BookMaster.frx":40E6
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame3 
            Height          =   8300
            Left            =   -74880
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   480
            Width           =   15960
            _Version        =   65536
            _ExtentX        =   28152
            _ExtentY        =   14640
            _StockProps     =   77
            BackColor       =   -2147483633
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
            Picture         =   "BookMaster.frx":4102
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   8085
               Left            =   120
               TabIndex        =   57
               Top             =   105
               Width           =   15720
               _Version        =   524288
               _ExtentX        =   27728
               _ExtentY        =   14261
               _StockProps     =   64
               ButtonDrawMode  =   1
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
               MaxCols         =   4
               MaxRows         =   100
               SpreadDesigner  =   "BookMaster.frx":411E
            End
            Begin VB.TextBox Text99 
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
               Left            =   5640
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   3600
               Width           =   5775
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame5 
            Height          =   8300
            Left            =   -74880
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   480
            Width           =   15960
            _Version        =   65536
            _ExtentX        =   28152
            _ExtentY        =   14640
            _StockProps     =   77
            BackColor       =   -2147483633
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
            Picture         =   "BookMaster.frx":476E
            Begin FPSpreadADO.fpSpread fpSpread3 
               Height          =   8085
               Left            =   120
               TabIndex        =   44
               Top             =   105
               Width           =   15720
               _Version        =   524288
               _ExtentX        =   27728
               _ExtentY        =   14261
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
               MaxCols         =   7
               MaxRows         =   100
               SpreadDesigner  =   "BookMaster.frx":478A
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame7 
            Height          =   8300
            Left            =   -74880
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   480
            Width           =   15960
            _Version        =   65536
            _ExtentX        =   28152
            _ExtentY        =   14640
            _StockProps     =   77
            BackColor       =   -2147483633
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
            Picture         =   "BookMaster.frx":4E48
            Begin FPSpreadADO.fpSpread fpSpread4 
               Height          =   8085
               Left            =   120
               TabIndex        =   47
               Top             =   105
               Width           =   15720
               _Version        =   524288
               _ExtentX        =   27728
               _ExtentY        =   14261
               _StockProps     =   64
               ButtonDrawMode  =   1
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
               GridColor       =   33023
               MaxCols         =   25
               MaxRows         =   100
               SpreadDesigner  =   "BookMaster.frx":4E64
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame8 
            Height          =   8300
            Left            =   -74880
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   480
            Width           =   15960
            _Version        =   65536
            _ExtentX        =   28152
            _ExtentY        =   14640
            _StockProps     =   77
            BackColor       =   -2147483633
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
            Picture         =   "BookMaster.frx":5ECC
            Begin FPSpreadADO.fpSpread fpSpread5 
               Height          =   8085
               Left            =   120
               TabIndex        =   51
               Top             =   105
               Width           =   15720
               _Version        =   524288
               _ExtentX        =   27728
               _ExtentY        =   14261
               _StockProps     =   64
               ButtonDrawMode  =   1
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
               GridColor       =   33023
               MaxCols         =   12
               MaxRows         =   99
               SpreadDesigner  =   "BookMaster.frx":5EE8
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame9 
            Height          =   8300
            Left            =   -74880
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   480
            Width           =   15960
            _Version        =   65536
            _ExtentX        =   28152
            _ExtentY        =   14640
            _StockProps     =   77
            BackColor       =   -2147483633
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
            Picture         =   "BookMaster.frx":68BD
            Begin VB.CommandButton cmdLoadElement 
               Caption         =   ".."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   15600
               TabIndex        =   60
               ToolTipText     =   "Load Elements"
               Top             =   7945
               Width           =   240
            End
            Begin FPSpreadADO.fpSpread fpSpread6 
               Height          =   8085
               Left            =   120
               TabIndex        =   53
               Top             =   105
               Width           =   15720
               _Version        =   524288
               _ExtentX        =   27728
               _ExtentY        =   14261
               _StockProps     =   64
               ButtonDrawMode  =   1
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
               GridColor       =   33023
               MaxCols         =   13
               MaxRows         =   99
               SpreadDesigner  =   "BookMaster.frx":68D9
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
            Height          =   8300
            Left            =   -74880
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   480
            Width           =   15960
            _Version        =   65536
            _ExtentX        =   28152
            _ExtentY        =   14640
            _StockProps     =   77
            BackColor       =   -2147483633
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
            Picture         =   "BookMaster.frx":72E9
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   8085
               Left            =   120
               TabIndex        =   55
               Top             =   105
               Width           =   15720
               _Version        =   524288
               _ExtentX        =   27728
               _ExtentY        =   14261
               _StockProps     =   64
               ButtonDrawMode  =   1
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
               GridColor       =   33023
               MaxCols         =   24
               MaxRows         =   99
               SpreadDesigner  =   "BookMaster.frx":7305
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   120
            TabIndex        =   62
            Top             =   8085
            Width           =   15975
            _Version        =   65536
            _ExtentX        =   28178
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483633
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
            Caption         =   " Ctrl+A->Add  Ctrl+E->Edit  Ctrl+D->Delete  Ctrl+S->Save  F5-> Refresh-> F12-> Create Duplicate Item Master"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "BookMaster.frx":84C6
            Picture         =   "BookMaster.frx":84E2
         End
         Begin Mh3dlblLib.Mh3dLabel CmdExport 
            Height          =   330
            Left            =   14040
            TabIndex        =   67
            Top             =   8445
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TintColor       =   16711935
            Caption         =   " &Export List"
            FillColor       =   9164542
            TextColor       =   0
            Picture         =   "BookMaster.frx":84FE
            Picture         =   "BookMaster.frx":851A
         End
         Begin Mh3dlblLib.Mh3dLabel CmdPrint 
            Height          =   330
            Left            =   12980
            TabIndex        =   68
            Top             =   8445
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TintColor       =   16711935
            Caption         =   " &Print List"
            FillColor       =   9164542
            TextColor       =   0
            Picture         =   "BookMaster.frx":8536
            Picture         =   "BookMaster.frx":8552
         End
         Begin Mh3dlblLib.Mh3dLabel CmdLabel 
            Height          =   330
            Left            =   15090
            TabIndex        =   69
            Top             =   8445
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TintColor       =   16711935
            Caption         =   " &Label Print"
            FillColor       =   9164542
            TextColor       =   0
            Picture         =   "BookMaster.frx":856E
            Picture         =   "BookMaster.frx":858A
         End
         Begin VB.TextBox txtNotes 
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
            Height          =   345
            Left            =   720
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   61
            ToolTipText     =   "Open Notes"
            Top             =   7920
            Visible         =   0   'False
            Width           =   1455
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
            ForeColor       =   &H80000009&
            Height          =   330
            Left            =   120
            TabIndex        =   32
            Top             =   8445
            Width           =   615
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   16455
      _ExtentX        =   29025
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
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print Preview"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
Attribute VB_Name = "FrmBookMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PtgType As String
Public SL As Boolean, MasterCode As String, ItemType As String, MFCode As String
Dim cnItemMaster As New ADODB.Connection
Dim rstBdgList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, orstItemList As New ADODB.Recordset, rstItemMaster As New ADODB.Recordset, rstHSNCodeList As New ADODB.Recordset, rstItemGroupList As New ADODB.Recordset, rstFinishSizeList As New ADODB.Recordset, rstBindingTypeList As New ADODB.Recordset, rstItemChild As New ADODB.Recordset, rstUnfinishedItemList As New ADODB.Recordset, rstElementList As New ADODB.Recordset, rstSizeList As New ADODB.Recordset, rstUnitList As New ADODB.Recordset, rstMFList As New ADODB.Recordset, rstCompanyMaster As New ADODB.Recordset
Dim HSNCode As String, ItemGroupCode As String, FinishSizeCode1 As String, FinishSizeCode2 As Variant, BindingTypeCode As String, GeneralItemCode As Variant, PaperCode As Variant, UnfinishedItemCode As Variant, ElementCode As String, OperationCode As Variant, SizeCode As String, CalcModeCode As Variant, BinderyProcessCode As Variant, PrintingSizeCode As Variant, ColorCode As Variant, TextSizeCode As String, TitleSizeCode As String, IntegrationUnitCode As String
Dim PrevStr As String, blnRecordExist As Boolean, SortCol As String, SortOrder As String, EditMode As Boolean
Dim StartColumn As String, StartRow As String, EndColumn As String, EndRow As String, PrintFlag As Boolean
Dim srmPicMgr As New ADODB.Stream, imgFile As String, imgNo As Long, imgNoLast As Long, imgNoFirst As Long
Dim rstPicList As New ADODB.Recordset, rstPicMaster As New ADODB.Recordset
Dim oOutlook As New Outlook.Application
Dim EMailID, Attachment, MsgSubject, MsgText, Message As String
Private Sub btnNotes_Click()
    frmNotes.NotesFlag = 2
    frmNotes.Label1.Caption = "Notes : " & Text2.Text
    frmNotes.Show (vbModal)
    Text2.SetFocus
End Sub
Private Sub Form_Load()
    Dim i As Integer
    If Not SL Then MasterCode = ""
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    Me.Top = 1200
    Me.Left = 300
    BusySystemIndicator True
    Me.Caption = IIf(ItemType = "F", "Item Master [Finished]", "Item Master [Unfinished]")
    cnItemMaster.CursorLocation = adUseClient: cnItemMaster.Open cnDatabase.ConnectionString
    rstItemList.Open "SELECT ROW_NUMBER() OVER (ORDER BY C.Name) AS RowNo,P.Name,P.PrintName As PrintName,C.PrintName As ItemGroup,BusyCode As Alias,P.Code,ISBN,(Select PrintName From GeneralMaster Where Type=25 AND  Code=IntegrationUnit) As Unit,ISNULL((Select Name FROM AccountMaster Where Code=P.MFCD),'') AS MFDC,Format(P.Price,'##00.00') As Price,Format(P.salePrice,'##00.00') As SalePrice,Format(P.PurPrice,'##00.00') AS PurPrice,C.Name As GroupCode,C.PrintName As GroupName,P.Remarks,P.S_No As S_No,P.ItemMarks As Marks,Format(P.PurDate,'dd-MM-yy') AS PurDate FROM BookMaster P INNER JOIN GeneralMaster C ON P.[Group]=C.Code WHERE P.Type='" & ItemType & "' ORDER BY C.Name,P.S_No", cnItemMaster, adOpenKeyset, adLockOptimistic
    LoadMasterList
    rstItemMaster.CursorLocation = adUseClient
    rstItemList.Filter = adFilterNone
    If rstItemList.RecordCount > 0 Then
        rstItemList.MoveFirst
        If Not CheckEmpty(MasterCode, False) Then rstItemList.Find "[Code]='" & MasterCode & "'"
    End If
    
    'Data Grid
        DataGrid1.Columns(0).Caption = "S.No."
        DataGrid1.Columns(0).Width = 500
        DataGrid1.Columns(1).Width = 4950
        DataGrid1.Columns(3).Width = 2500
        DataGrid1.Columns(5).Width = 0
        DataGrid1.Columns(6).Width = 1700
        DataGrid1.Columns(7).Width = 1184
        DataGrid1.Columns(8).Width = 0
        DataGrid1.Columns(9).Width = 950
        DataGrid1.Columns(10).Width = 950
        DataGrid1.Columns(11).Width = 950
        DataGrid1.Columns(2).Visible = False

    BusySystemIndicator False
    SSTab1.Tab = 0
    SortCol = "Name"
    
    If Trim(ReadFromFile("Customer Type")) = "General" Then
        Mh3dLabel1(0).Caption = " Item Code"
        Mh3dLabel3.Caption = " Item Description"
        Mh3dLabel21.Caption = " New Item"
        Mh3dLabel12.Visible = True
        Mh3dLabel30.Visible = True
        Mh3dLabel18.Visible = True
        Mh3dLabel20.Visible = True
        Mh3dLabel22.Visible = True
        Mh3dLabel23.Visible = True
        Mh3dLabel24.Visible = True
        
        MhRealInput5.Visible = True
        MhRealInput6.Visible = True
        Text17.Visible = True
        Text18.Visible = True
        Text19.Visible = True
        Text20.Visible = True
        Text21.Visible = True
        Text22.Visible = True
        Mh3dLabel6.Visible = True
        Mh3dLabel15.Visible = True
        Text5.TabStop = False
        Text10.TabStop = False
        For i = 2 To 7
           SSTab1.TabEnabled(i) = False
            SSTab1.TabCaption(i) = " "
            SSTab1.TabHeight = 300
        Next
'DataGrid1_Caption
i = 0
        i = i + 0: DataGrid1.Columns(i).Caption = " Marks"
        i = i + 1: DataGrid1.Columns(i).Caption = " Item Code"
        i = i + 1: DataGrid1.Columns(i).Caption = " Item Description"
        i = i + 1: DataGrid1.Columns(i).Caption = " Group"
        i = i + 1: DataGrid1.Columns(i).Caption = " G. Code"
        i = i + 1: DataGrid1.Columns(i).Caption = " S.No."
        i = i + 1: DataGrid1.Columns(i).Caption = " Code"
        i = i + 1: DataGrid1.Columns(i).Caption = " Unit"
        i = i + 1: DataGrid1.Columns(i).Caption = " MFDC"
        i = i + 1: DataGrid1.Columns(i).Caption = " Pur. Price"
        i = i + 1: DataGrid1.Columns(i).Caption = " Sales Price"
        i = i + 1: DataGrid1.Columns(i).Caption = " Cash Price"

'        i = i + 0: DataGrid1.Columns(0).Caption = " Marks"
'        i = i + 1: DataGrid1.Columns(1).Caption = " Item Code"
'        i = i + 1: DataGrid1.Columns(2).Caption = " Item Description"
'        i = i + 1: DataGrid1.Columns(3).Caption = " Group"
'        i = i + 1: DataGrid1.Columns(4).Caption = " G. Code"
'        i = i + 1: DataGrid1.Columns(5).Caption = " S.No."
'        i = i + 1: DataGrid1.Columns(6).Caption = " Code"
'        i = i + 1: DataGrid1.Columns(7).Caption = " Unit"
'        i = i + 1: DataGrid1.Columns(8).Caption = " MFDC"
'        i = i + 1: DataGrid1.Columns(9).Caption = " Pur. Price"
'        i = i + 1: DataGrid1.Columns(10).Caption = " Sales Price"
'        i = i + 1: DataGrid1.Columns(11).Caption = " Cash Price"
'DataGrid1_Width
i = 0
        i = i + 0: DataGrid1.Columns(i).Width = 650
        i = i + 0: DataGrid1.Columns(i).DataField = "Marks"
        i = i + 0: DataGrid1.Columns(i).Alignment = dbgCenter
        i = i + 0: DataGrid1.Columns(i).NumberFormat = "General"
        i = i + 1: DataGrid1.Columns(i).Width = 1000
        i = i + 1: DataGrid1.Columns(i).Width = 4844
        i = i + 0: DataGrid1.Columns(i).Visible = True
        i = i + 1: DataGrid1.Columns(i).Width = 3000
        i = i + 1: DataGrid1.Columns(i).Width = 800
        i = i + 0: DataGrid1.Columns(i).DataField = "GroupCode"
        i = i + 0: DataGrid1.Columns(i).Alignment = dbgCenter
        i = i + 0: DataGrid1.Columns(i).NumberFormat = "##0"
        i = i + 1: DataGrid1.Columns(i).Width = 550
        i = i + 0: DataGrid1.Columns(i).DataField = "S_No"
        i = i + 0: DataGrid1.Columns(i).Alignment = dbgCenter
        i = i + 0: DataGrid1.Columns(i).NumberFormat = "##0"
        i = i + 1: DataGrid1.Columns(i).Width = 0
        i = i + 0: DataGrid1.Columns(i).Visible = False
        i = i + 1: DataGrid1.Columns(i).Width = 500
        i = i + 1: DataGrid1.Columns(i).Width = 1000
        i = i + 1: DataGrid1.Columns(i).Width = 1000
        i = i + 1: DataGrid1.Columns(i).Width = 1000
'        DataGrid1.Columns(0).Width = 650
'        DataGrid1.Columns(0).DataField = "Marks"
'        DataGrid1.Columns(0).Alignment = dbgCenter
'        DataGrid1.Columns(0).NumberFormat = "General"
'        DataGrid1.Columns(1).Width = 1000
'        DataGrid1.Columns(2).Width = 4844
'        DataGrid1.Columns(2).Visible = True
'        DataGrid1.Columns(3).Width = 3000
'        DataGrid1.Columns(4).Width = 800
'        DataGrid1.Columns(4).DataField = "GroupCode"
'        DataGrid1.Columns(4).Alignment = dbgCenter
'        DataGrid1.Columns(4).NumberFormat = "##0"
'        DataGrid1.Columns(5).Width = 550
'        DataGrid1.Columns(5).DataField = "S_No"
'        DataGrid1.Columns(5).Alignment = dbgCenter
'        DataGrid1.Columns(5).NumberFormat = "##0"
'        DataGrid1.Columns(6).Width = 0
'        DataGrid1.Columns(6).Visible = False
'        DataGrid1.Columns(7).Width = 500
'        DataGrid1.Columns(8).Width = 1000
'        DataGrid1.Columns(9).Width = 1000
'        DataGrid1.Columns(10).Width = 1000
'        DataGrid1.Columns(11).Width = 1000
''DataGrid1_Data_Field
'        DataGrid1.Columns(0).DataField = "Marks"
'        DataGrid1.Columns(4).DataField = "GroupCode"
'        DataGrid1.Columns(5).DataField = "S_No"
'        DataGrid1.Columns(0).NumberFormat = "General"
'        DataGrid1.Columns(4).NumberFormat = "##0"
'        DataGrid1.Columns(5).NumberFormat = "##0"
''DataGrid1_Data_Alignment
'        DataGrid1.Columns(0).Alignment = dbgCenter
'        DataGrid1.Columns(4).Alignment = dbgCenter
'        DataGrid1.Columns(5).Alignment = dbgCenter
    Else
'DataGrid1_Data_Field
        DataGrid1.Columns(0).DataField = "RowNo"
        DataGrid1.Columns(4).DataField = "BusyCode"
        DataGrid1.Columns(5).DataField = "Code"
    End If
    Set DataGrid1.DataSource = rstItemList
    If Not (rstItemList.EOF Or rstItemList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstItemList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
cmdUpload.Enabled = True
    If Shift = 0 And KeyCode = vbKeyEscape Then
        EditMode = False
        If SSTab1.Tab = 0 Then
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then 'Add button enabled
                SSTab1.Tab = 0
            Else
                If InStr(1, "fpSpread1_fpSpread2_fpSpread3_fpSpread4_fpSpread5_fpSpread6", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True
                If Not EditMode Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then Me.ActiveControl.SetFocus Else Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                End If
            End If
        End If
        If Not EditMode Then KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then 'Add
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1): KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then 'Edit
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2): KeyCode = 0
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF8)) And Toolbar1.Buttons.Item(3).Enabled Then  'Delete
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3): KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF12 And Toolbar1.Buttons.Item(1).Enabled Then 'Duplicate
        If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then DuplicateRecord
        KeyCode = 0
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2)) And Toolbar1.Buttons.Item(4).Enabled Then 'Save
        EditMode = False
        If InStr(1, "fpSpread1_fpSpread2_fpSpread3_fpSpread4_fpSpread5_fpSpread6", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then 'Refresh
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6): KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF And Toolbar1.Buttons.Item(1).Enabled Then 'First
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(13): KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then 'Previous
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(14): KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyN And Toolbar1.Buttons.Item(1).Enabled Then 'Next
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(15): KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL And Toolbar1.Buttons.Item(1).Enabled Then 'Last
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(16): KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        If Toolbar1.Buttons.Item(1).Enabled Then
            If SL Then
                If SSTab1.Tab = 0 Then Me.Tag = "S": slCode = rstItemList.Fields("Code").Value: slName = rstItemList.Fields("Name").Value: KeyCode = 0: Unload Me: Exit Sub
            Else
                SSTab1.Tab = 1: SSTab1.SetFocus
            End If
        Else 'Move to next control
            If InStr(1, "fpSpread1_fpSpread2_fpSpread3_fpSpread4_fpSpread5_fpSpread6", Me.ActiveControl.Name) = 0 Then Sendkeys "{TAB}"
        End If
        If InStr(1, "fpSpread1_fpSpread2_fpSpread3_fpSpread4_fpSpread5_fpSpread6", Me.ActiveControl.Name) = 0 Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
    Else
        If Me.Tag <> "S" Then slCode = "": slName = ""
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(orstItemList)
    Call CloseRecordset(rstPicList)
    Call CloseRecordset(rstPicMaster)
    Call CloseRecordset(rstUnitList)
    Call CloseRecordset(rstItemMaster)
    Call CloseRecordset(rstHSNCodeList)
    Call CloseRecordset(rstItemGroupList)
    Call CloseRecordset(rstFinishSizeList)
    Call CloseRecordset(rstBindingTypeList)
    Call CloseRecordset(rstItemChild)
    Call CloseRecordset(rstUnfinishedItemList)
    Call CloseRecordset(rstElementList)
    Call CloseRecordset(rstSizeList)
    Call CloseConnection(cnItemMaster)
    ShowProgressInStatusBar False
End Sub
Private Sub fpSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim cVal As Variant
fpSpread1.GetText 4, Row, cVal
    If cVal <> "" Then
        If fpSpread1.ActiveCol = 1 Then fpSpread1.SetText 4, Row, "": fpSpread1.SetText 2, Row, "": fpSpread1.SetText 3, Row, ""
    Else
    Sendkeys "{TAB}"
   End If
    Dim CurVal As Variant
    With fpSpread1
            .SetActiveCell 2, .ActiveRow
            If .ActiveCol = 2 Then
                .GetText 1, .ActiveRow, CurVal
                If CheckEmpty(CurVal, False) Then Exit Sub
                If CurVal = "BOM Item" Then
                    .GetText .ActiveCol + 2, .ActiveRow, GeneralItemCode
                    On Error Resume Next
                    FrmOutsourceItemMaster.SL = True
                    FrmOutsourceItemMaster.MasterCode = GeneralItemCode
                    Load FrmOutsourceItemMaster
                    If Err.Number <> 364 Then FrmOutsourceItemMaster.Show vbModal
                    On Error GoTo 0
                    .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 2, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                ElseIf CurVal = "Paper" Then
                    .GetText .ActiveCol + 2, .ActiveRow, PaperCode
                    On Error Resume Next
                    FrmPaperMaster.SL = True
                    FrmPaperMaster.MasterCode = PaperCode
                    Load FrmPaperMaster
                    If Err.Number <> 364 Then FrmPaperMaster.Show vbModal
                    On Error GoTo 0
                    .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 2, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                ElseIf CurVal = "Finished Item" Then
                    .GetText .ActiveCol + 2, .ActiveRow, UnfinishedItemCode
                    On Error Resume Next
                    Dim frmItemMaster As New FrmBookMaster
                    frmItemMaster.SL = True
                    frmItemMaster.ItemType = "F"
                    frmItemMaster.MasterCode = UnfinishedItemCode
                    Load frmItemMaster
                    If Err.Number <> 364 Then frmItemMaster.Show vbModal
                    On Error GoTo 0
                    .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 2, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                ElseIf CurVal = "Unfinished Item" Then
                    .GetText .ActiveCol + 2, .ActiveRow, UnfinishedItemCode
                    On Error Resume Next
                    'Dim frmItemMaster As New FrmBookMaster
                    frmItemMaster.SL = True
                    frmItemMaster.ItemType = "R"
                    frmItemMaster.MasterCode = UnfinishedItemCode
                    Load frmItemMaster
                    If Err.Number <> 364 Then frmItemMaster.Show vbModal
                    On Error GoTo 0
                    .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 2, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                End If
            End If
    End With
End Sub
Private Sub Text1_Change()
    If rstItemList.RecordCount = 0 Then Exit Sub
    rstItemList.MoveFirst
    If Len(Text1.Text) > 0 Then
        rstItemList.Filter = "[" & SortCol & "] Like '%" & FixQuote(Text1.Text) & "%'"
        If rstItemList.EOF Then  'if Spelling mistake
            rstItemList.Filter = adFilterNone
            rstItemList.MoveFirst
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            Sendkeys "{End}"
        Else    'if Spelling alright
            PrevStr = Text1.Text
        End If
    Else
        rstItemList.Filter = adFilterNone
        rstItemList.MoveFirst
        Set DataGrid1.DataSource = rstItemList
        PrevStr = ""
    End If
    If Not (rstItemList.EOF Or rstItemList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    If rstItemList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstItemList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstItemList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstItemList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstItemList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstItemList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstItemList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstItemList
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
    On Error Resume Next
    If Toolbar1.Buttons.Item(1).Enabled Then
        If SSTab1.Tab >= 1 Then
            ViewRecord
        Else
            If Not (rstItemList.EOF Or rstItemList.BOF) Then
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
        If SSTab1.Tab = 1 Then
            Mh3dFrame2.Enabled = True: Mh3dFrame3.Enabled = False: Mh3dFrame4.Enabled = False: Mh3dFrame5.Enabled = False: Mh3dFrame7.Enabled = False: Mh3dFrame8.Enabled = False: Mh3dFrame9.Enabled = False: Text2.SetFocus
        ElseIf SSTab1.Tab = 2 Then
            Mh3dFrame2.Enabled = False: Mh3dFrame3.Enabled = True: Mh3dFrame4.Enabled = False: Mh3dFrame5.Enabled = False: Mh3dFrame7.Enabled = False: Mh3dFrame8.Enabled = False: Mh3dFrame9.Enabled = False: fpSpread1.SetFocus
        ElseIf SSTab1.Tab = 3 Then
            Mh3dFrame2.Enabled = False: Mh3dFrame3.Enabled = False: Mh3dFrame4.Enabled = False: Mh3dFrame5.Enabled = True: Mh3dFrame7.Enabled = False: Mh3dFrame8.Enabled = False: Mh3dFrame9.Enabled = False: fpSpread3.SetFocus
        ElseIf SSTab1.Tab = 4 Then
            Mh3dFrame2.Enabled = False: Mh3dFrame3.Enabled = False: Mh3dFrame4.Enabled = False: Mh3dFrame5.Enabled = False: Mh3dFrame7.Enabled = True: Mh3dFrame8.Enabled = False: Mh3dFrame9.Enabled = False: fpSpread4.SetFocus
        ElseIf SSTab1.Tab = 5 Then
            Mh3dFrame2.Enabled = False: Mh3dFrame3.Enabled = False: Mh3dFrame4.Enabled = True: Mh3dFrame5.Enabled = False: Mh3dFrame7.Enabled = False: Mh3dFrame8.Enabled = False: Mh3dFrame9.Enabled = False: fpSpread2.SetFocus
        ElseIf SSTab1.Tab = 6 Then
            Mh3dFrame2.Enabled = False: Mh3dFrame3.Enabled = False: Mh3dFrame4.Enabled = False: Mh3dFrame5.Enabled = False: Mh3dFrame7.Enabled = False: Mh3dFrame8.Enabled = True: Mh3dFrame9.Enabled = False: fpSpread5.SetFocus
        ElseIf SSTab1.Tab = 7 Then
            Mh3dFrame2.Enabled = False: Mh3dFrame3.Enabled = False: Mh3dFrame4.Enabled = False: Mh3dFrame5.Enabled = False: Mh3dFrame7.Enabled = False: Mh3dFrame8.Enabled = False: Mh3dFrame9.Enabled = True: fpSpread6.SetFocus
            cmdLoadElement_Click
        End If
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim CellVal(1 To 5) As Variant, HiLiteRecord As Boolean, UpdateFlag As Integer, i As Integer
    If Button.Index = 1 Then 'Add Button
        If rstItemMaster.State = adStateOpen Then rstItemMaster.Close
        rstItemMaster.Open "SELECT * FROM BookMaster WHERE Code=''", cnItemMaster, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstItemMaster) Then
            Call SetButtons(False): SSTab1.Tab = 1: Text2.SetFocus: blnRecordExist = False
            cnItemMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then 'Edit Button
        If rstItemList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then 'Delete Button
        If rstItemList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Master"): Exit Sub
        SSTab1.Tab = 1
        If chkRef("SELECT Item FROM AccountChild0801 WHERE Category='3' AND Item='" & rstItemList.Fields("Code").Value & "'") Or Left(rstItemList.Fields("Code").Value, 1) = "*" Then
            DisplayError ("Failed to delete the record")
        ElseIf MsgBox("Are you sure to Delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            cnItemMaster.BeginTrans
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnItemMaster.Execute "DELETE FROM BookMaster WHERE Code='" & rstItemList.Fields("Code").Value & "'"
            If Err.Number = 0 Then
                cnItemMaster.CommitTrans
                rstItemList.Delete
                rstItemList.MoveNext
                If rstItemList.RecordCount > 0 And rstItemList.EOF Then rstItemList.MoveLast
                ShowProgressInStatusBar True
                Timer1.Enabled = True
                Text1.Text = ""
                rstItemList.Filter = adFilterNone
            Else
                DisplayError (Err.Description)
                cnItemMaster.RollbackTrans
            End If
            MdiMainMenu.MousePointer = vbNormal
            On Error GoTo 0
        End If
        SetButtons (True)
        SetButtonsForNoRecord
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 4 Then 'Save Button
        If ValidateForm Then Exit Sub
        If blnRecordExist And AllowMastersModification = 0 Then Call DisplayError("You don't have the rights to Edit this Master"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(5): Exit Sub
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstItemMaster) Then
            UpdateFlag = 1
            If UpdateItemList("D") Then
                SSTab1.Tab = 2
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 1, i
                        .GetText 1, i, CellVal(1) 'Category
                        .GetText 4, i, CellVal(2) 'Item
                        If Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False)) Then If Not UpdateItemList("I") Then UpdateFlag = 0: Exit For
                    Next
                End With
                SSTab1.Tab = 3
                With fpSpread3
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 1, i
                        .GetText 1, i, CellVal(1) 'Arrived ON
                        .GetText 2, i, CellVal(2) 'Correction
                        If IsDate(CellVal(1)) And (Not CheckEmpty(CellVal(2), False)) Then If Not UpdateItemList("I") Then UpdateFlag = 0: Exit For
                    Next
                End With
                SSTab1.Tab = 4
                With fpSpread4
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 1, i
                        .GetText 18, i, CellVal(1) 'Element
                        .GetText 19, i, CellVal(2) 'Finish Size
                        .GetText 20, i, CellVal(3) 'Printing Size
                        .GetText 21, i, CellVal(4) 'Color
                        .GetText 14, i, CellVal(5) 'Plate
                        If Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False) Or CheckEmpty(CellVal(3), False) Or CheckEmpty(CellVal(4), False) Or CheckEmpty(CellVal(5), False)) Then If Not UpdateItemList("I") Then UpdateFlag = 0: Exit For
                    Next
                End With
                SSTab1.Tab = 5
                With fpSpread2
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 1, i
                        .GetText 14, i, CellVal(1) 'Element
                        .GetText 15, i, CellVal(2) 'Finish Size
                        .GetText 16, i, CellVal(3) 'Printing Size
                        If Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False) Or CheckEmpty(CellVal(3), False)) Then If Not UpdateItemList("I") Then UpdateFlag = 0: Exit For
                    Next
                End With
                SSTab1.Tab = 6
                With fpSpread5
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 1, i
                        .GetText 7, i, CellVal(1) 'Element
                        .GetText 8, i, CellVal(2) 'Operation
                        .GetText 9, i, CellVal(3) 'Size
                        .GetText 10, i, CellVal(4) 'Calc Mode
                        If IIf(CheckEmpty(CellVal(3), False), Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False) Or CheckEmpty(CellVal(4), False)), Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False) Or CheckEmpty(CellVal(3), False) Or CheckEmpty(CellVal(4), False))) Then If Not UpdateItemList("I") Then UpdateFlag = 0: Exit For
                    Next
                End With
                SSTab1.Tab = 7
                With fpSpread6
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 1, i
                        .GetText 10, i, CellVal(1) 'Binding
                        .GetText 11, i, CellVal(2) 'Process
                        .GetText 12, i, CellVal(3) 'Size
                        .GetText 13, i, CellVal(4) 'Calc Mode
                        If Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False) Or CheckEmpty(CellVal(3), False) Or CheckEmpty(CellVal(4), False)) Then If Not UpdateItemList("I") Then UpdateFlag = 0: Exit For
                    Next
                End With
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnItemMaster.CommitTrans
            If rstItemMaster.State = adStateOpen Then rstItemMaster.Close
            rstItemMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            Call MsgBox("Record updated !!!", vbInformation, App.Title)
        Else
            DisplayError ("Failed to Save the Record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then 'Cancel Button
        If CancelRecordUpdate(rstItemMaster) Then
            cnItemMaster.RollbackTrans
            If rstItemMaster.State = adStateOpen Then rstItemMaster.Close
            rstItemMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then 'Refresh Button
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        RefreshData rstItemList
        Set DataGrid1.DataSource = rstItemList
        LoadMasterList
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then 'Filter Button
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Name", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then 'First Record Button
        If rstItemList.RecordCount > 0 Then rstItemList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then 'Previous Record Button
        If rstItemList.RecordCount > 0 Then
           rstItemList.MovePrevious
           If rstItemList.BOF Then rstItemList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then 'Next Record Button
        If rstItemList.RecordCount > 0 Then
           rstItemList.MoveNext
           If rstItemList.EOF Then
              rstItemList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then 'Last Record Button
        If rstItemList.RecordCount > 0 Then rstItemList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstItemList.EOF Or rstItemList.BOF) Then
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
    SortCol = DataGrid1.Columns(ColIndex).DataField
    SortOrder = IIf(SortOrder = "Asc", "Desc", "Asc")
    rstItemList.Sort = "[" + SortCol & "] " & SortOrder
    DataGrid1.ClearSelCols
    If Not (rstItemList.EOF Or rstItemList.BOF) Then
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
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
    Mh3dFrame3.Enabled = False: Mh3dFrame4.Enabled = False: Mh3dFrame5.Enabled = False: Mh3dFrame7.Enabled = False: Mh3dFrame8.Enabled = False: Mh3dFrame9.Enabled = False
End Sub
Private Sub SetButtonsForNoRecord()
    If rstItemList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstItemMaster.EOF Or rstItemMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnItemMaster, "BookMaster", "Code", "Name", Trim(Text2.Text), rstItemMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = Text2.Text
        If CheckEmpty(Text15, False) Then Text15.Text = Text2.Text
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If CheckEmpty(Text4.Text, False) Then Exit Sub
    If rstItemMaster.EOF Or rstItemMaster.BOF Then Exit Sub
    If CheckDuplicate(cnItemMaster, "BookMaster", "Code", "ISBN", Text4.Text, rstItemMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf Len(Text4.Text) = 13 Then
        If Not bVerifySum10(Text4.Text) Then Cancel = True
    ElseIf Len(Text4.Text) = 17 Then
        If Not bVerifySum13(Text4.Text) Then Cancel = True
    End If
End Sub
Private Sub Text11_Validate(Cancel As Boolean) 'Alias
    If CheckEmpty(Text11.Text, False) Then Exit Sub
    If rstItemMaster.EOF Or rstItemMaster.BOF Then Exit Sub
    If CheckDuplicate(cnItemMaster, "BookMaster", "Code", "BusyCode", Text11.Text, rstItemMaster.Fields("Code").Value, False) Then Cancel = True
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer) 'HSN Code
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "18"
        FrmGeneralMaster.MasterCode = HSNCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        HSNCode = slCode: Text7.Text = slName
        If Not CheckEmpty(HSNCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        HSNCode = "": Text7.Text = ""
    End If
End Sub
Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "5"
        FrmGeneralMaster.MasterCode = ItemGroupCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        ItemGroupCode = slCode: Text8.Text = slName
        If Not CheckEmpty(ItemGroupCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        ItemGroupCode = "": Text8.Text = ""
    End If
End Sub
Private Sub Text8_Validate(Cancel As Boolean)
    If CheckEmpty(Text8.Text, False) Then Cancel = True
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "11"
        FrmGeneralMaster.MasterCode = FinishSizeCode1
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        FinishSizeCode1 = slCode: Text5.Text = slName
        If Not CheckEmpty(FinishSizeCode1, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    If CheckEmpty(Text5.Text, False) Then Cancel = True
End Sub
Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)    'Binding Type
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "6"
        FrmGeneralMaster.MasterCode = BindingTypeCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        BindingTypeCode = slCode: Text10.Text = slName
        If Not CheckEmpty(BindingTypeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        BindingTypeCode = "": Text10.Text = ""
    End If
End Sub
Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmGeneralMaster.SL = True
        FrmGeneralMaster.MasterType = "25"
        FrmGeneralMaster.MasterCode = IntegrationUnitCode
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
        On Error GoTo 0
        IntegrationUnitCode = slCode: Text16.Text = slName
        If Not CheckEmpty(IntegrationUnitCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyDelete Then
        IntegrationUnitCode = "": Text16.Text = ""
    End If
End Sub
Private Sub Text16_Validate(Cancel As Boolean)
    If CheckEmpty(Text16.Text, False) Then Cancel = True
End Sub
Private Sub Text18_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = MFCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        MFCode = slCode: Text18.Text = slName
        If Not CheckEmpty(MFCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer) 'BOM Item
    Dim CurVal As Variant
    With fpSpread1
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to Delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 2 Then
                .GetText 1, .ActiveRow, CurVal
                If CheckEmpty(CurVal, False) Then Exit Sub
                If CurVal = "BOM Item" Then
                    .GetText .ActiveCol + 2, .ActiveRow, GeneralItemCode
                    On Error Resume Next
                    FrmOutsourceItemMaster.SL = True
                    FrmOutsourceItemMaster.MasterCode = GeneralItemCode
                    Load FrmOutsourceItemMaster
                    If Err.Number <> 364 Then FrmOutsourceItemMaster.Show vbModal
                    On Error GoTo 0
                    .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 2, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                ElseIf CurVal = "Paper" Then
                    .GetText .ActiveCol + 2, .ActiveRow, PaperCode
                    On Error Resume Next
                    FrmPaperMaster.SL = True
                    FrmPaperMaster.MasterCode = PaperCode
                    Load FrmPaperMaster
                    If Err.Number <> 364 Then FrmPaperMaster.Show vbModal
                    On Error GoTo 0
                    .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 2, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                ElseIf CurVal = "Finished Item" Then
                    .GetText .ActiveCol + 2, .ActiveRow, UnfinishedItemCode
                    On Error Resume Next
                    Dim frmItemMaster As New FrmBookMaster
                    frmItemMaster.SL = True
                    frmItemMaster.ItemType = "F"
                    frmItemMaster.MasterCode = UnfinishedItemCode
                    Load frmItemMaster
                    If Err.Number <> 364 Then frmItemMaster.Show vbModal
                    On Error GoTo 0
                    .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 2, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                ElseIf CurVal = "Unfinished Item" Then
                    .GetText .ActiveCol + 2, .ActiveRow, UnfinishedItemCode
                    On Error Resume Next
                    'Dim frmItemMaster As New FrmBookMaster
                    frmItemMaster.SL = True
                    frmItemMaster.ItemType = "R"
                    frmItemMaster.MasterCode = UnfinishedItemCode
                    Load frmItemMaster
                    If Err.Number <> 364 Then frmItemMaster.Show vbModal
                    On Error GoTo 0
                    .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 2, .ActiveRow, slCode
                    If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                End If
            End If
        End If
    End With
End Sub
Private Sub fpSpread3_KeyDown(KeyCode As Integer, Shift As Integer) 'Editorial Component
    Dim CurVal(1 To 2) As Variant
    With fpSpread3
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to Delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 5 Then 'Team Member
                .GetText 7, .ActiveRow, CurVal(1)
                On Error Resume Next
                FrmTeamMemberMaster.SL = True
                FrmTeamMemberMaster.MasterCode = CurVal(1)
                Load FrmTeamMemberMaster
                If Err.Number <> 364 Then FrmTeamMemberMaster.Show vbModal
                On Error GoTo 0
                .SetText 5, .ActiveRow, slName: .SetText 7, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then .SetActiveCell 6, .ActiveRow Else Sendkeys "{ENTER}" 'Sendkeys "{ENTER}"
            End If
        End If
    End With
End Sub
Private Sub fpSpread4_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CurVal(1 To 2) As Variant, BindingForms As Integer
    With fpSpread4
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to Delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus: CalculateTotalForms
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then 'Element
                .GetText 1, .ActiveRow, CurVal(1)
                Text99.Text = FixQuote(CurVal(1))
                If rstElementList.RecordCount = 0 Then DisplayError ("No Record in Element Master"): .SetActiveCell 1, .ActiveRow: Exit Sub Else rstElementList.MoveFirst
                rstElementList.Find "[Col0] = '" & RTrim(CurVal(1)) & "'"
                SelectionType = "S": ElementCode = ""
                Call LoadSelectionList(rstElementList, "List of Element(s)...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text99, ElementCode)
                Call CloseForm(FrmSelectionList)
                Text99.Text = IIf(CheckEmpty(ElementCode, False), "", Text99.Text)
                .SetText 1, .ActiveRow, Text99.Text
                .SetText 18, .ActiveRow, ElementCode
                If CheckEmpty(ElementCode, False) Then .SetActiveCell 1, .ActiveRow Else Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 2 Then 'Finish Size
                .GetText 19, .ActiveRow, CurVal(1)
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "11"
                FrmGeneralMaster.MasterCode = CurVal(1)
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 2, .ActiveRow, slName: .SetText 19, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then
                    With rstItemChild
                        If .State = adStateOpen Then .Close
                        .Open "SELECT S.Name+'|'+'Pages/Ptg Form: '+IIF([Ups/Form]<10,'0','')+LTRIM([Ups/Form]) As Col0,S.Name,S.Code FROM FinishSizeChild C INNER JOIN GeneralMaster S ON C.[TextSize]=S.Code WHERE C.Code='" & slCode & "' ORDER BY S.Name,[Ups/Form]", cnItemMaster, adOpenKeyset, adLockReadOnly
                        SelectionType = "S": TextSizeCode = ""
                        fpSpread4.GetText 3, fpSpread4.ActiveRow, CurVal(1) 'Printing Size
                        If Not CheckEmpty(CurVal(1), False) And .RecordCount > 0 Then 'Move Pointer
                            .Find "[Name] = '" & RTrim(CurVal(1)) & "'"
                            If .EOF Then .MoveFirst Else Text99.Text = .Fields("Col0").Value
                        End If
                        Call LoadSelectionList(rstItemChild, "List of Printing Sizes...", "Name", "")
                        SearchOrder = 0: Text99.Text = ""
                        Call DisplaySelectionList(Text99, TextSizeCode)
                        Call CloseForm(FrmSelectionList)
                        If Not CheckEmpty(TextSizeCode, False) Then
                            .MoveFirst
                            .Find "[Code] = '" & TextSizeCode & "'"
                            fpSpread4.SetText 22, fpSpread4.ActiveRow, TextSizeCode & Right(.Fields("Col0").Value, 2) & .Fields("Name").Value '6+2+40
                        Else
                            fpSpread4.SetText 22, fpSpread4.ActiveRow, "" 'M-Printing Size
                        End If
                        Sendkeys "{ENTER}"
                    End With
                End If
            ElseIf .ActiveCol = 3 Then 'Printing Size
                .GetText 20, .ActiveRow, CurVal(1)
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "1"
                FrmGeneralMaster.MasterCode = CurVal(1)
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 3, .ActiveRow, slName: .SetText 20, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 7 Then 'Color
                .GetText 21, .ActiveRow, CurVal(1)
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "23"
                FrmGeneralMaster.MasterCode = CurVal(1)
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 7, .ActiveRow, slName: .SetText 21, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 14 Then 'Plate
                .GetText 25, .ActiveRow, CurVal(1)
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "24"
                FrmGeneralMaster.MasterCode = CurVal(1)
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 14, .ActiveRow, slName: .SetText 25, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            End If
        ElseIf KeyCode = vbKeyReturn Then
            If .ActiveCol = 3 Then 'Printing Size
                .GetText 3, .ActiveRow, CurVal(1): .GetText 22, .ActiveRow, CurVal(2) 'M-Printing Size
                If Trim(CurVal(1)) <> Trim(Mid(CurVal(2), 9, 60)) And (Not CheckEmpty(CurVal(2), False)) Then
                    If MsgBox("Printing Size [" & Trim(CurVal(1)) & "] is different from that in Master [" & Trim(Mid(CurVal(2), 9, 60)) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 3, .ActiveRow, Mid(CurVal(2), 9, 60): .SetText 20, .ActiveRow, Left(CurVal(2), 6)
                End If
                .GetText 3, .ActiveRow, CurVal(1): .GetText 2, .ActiveRow, CurVal(2) 'Finishing Size
                If CheckEmpty(CurVal(1), False) Or CheckEmpty(CurVal(2), False) Then Exit Sub
                Dim FL As Double, FR As Double, PL As Double, PR As Double, Ups01 As Integer, Ups02 As Integer
                PL = Val(Left(CurVal(1), InStr(1, CurVal(1), "X") - 1)) + 1: PR = Val(Mid(CurVal(1), InStr(1, CurVal(1), "X") + 1, 5)) + 1: FL = Val(Left(CurVal(2), InStr(1, CurVal(2), "X") - 1)): FR = Val(Mid(CurVal(2), InStr(1, CurVal(2), "X") + 1, 5))
                If Val(PL) * Val(PR) < Val(FL) * Val(FR) Then DisplayError ("Printing Size is smaller than Finish Size"): .SetActiveCell 3, .ActiveRow
            ElseIf .ActiveCol = 5 Then 'Pages/Printing Form
                .GetText 5, .ActiveRow, CurVal(1)
                .GetText 23, .ActiveRow, CurVal(2) 'C-Ups
                If Val(CurVal(1)) <> Val(CurVal(2)) And Val(CurVal(2)) <> 0 Then
                    If MsgBox("Variation in Calculated [" & Trim(CurVal(2)) & "] and Existing [" & Trim(CurVal(1)) & "] Pages/Printing Form ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 5, .ActiveRow, Val(CurVal(2))
                End If
            ElseIf .ActiveCol = 6 Then 'Pages/Form
                .GetText 6, .ActiveRow, CurVal(1)
                .GetText 5, .ActiveRow, CurVal(2)
                If Val(CurVal(1)) <> Val(CurVal(2)) And Val(CurVal(2)) <> 0 Then
                    If MsgBox("Variation in Calculated [" & Trim(CurVal(2)) & "] and Existing [" & Trim(CurVal(1)) & "] Pgs/Form ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 6, .ActiveRow, CurVal(2)
                End If
            ElseIf .ActiveCol = 9 Then 'Forms
                .GetText 9, .ActiveRow, CurVal(1)
                If Val(CurVal(1)) = 0 Then .SetText 10, .ActiveRow, 0: .SetText 11, .ActiveRow, 0: .SetText 12, .ActiveRow, 0: .SetText 13, .ActiveRow, 0
            ElseIf .ActiveCol = 10 Then '¼F
                .GetText 10, .ActiveRow, CurVal(1)
                CurVal(2) = CalculateForms("Q", .ActiveRow)
                If CurVal(1) = "" Then
                    .SetText 10, .ActiveRow, Val(CurVal(2))
                ElseIf Val(CurVal(1)) <> Val(CurVal(2)) Then
                    If MsgBox("Variation in Calculated [" & Trim(CurVal(2)) & "] and Existing [" & Trim(CurVal(1)) & "] ¼ Forms ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 10, .ActiveRow, Val(CurVal(2))
                End If
            ElseIf .ActiveCol = 11 Then '½F
                .GetText 11, .ActiveRow, CurVal(1)
                CurVal(2) = CalculateForms("H", .ActiveRow)
                If CurVal(1) = "" Then
                    .SetText 11, .ActiveRow, Val(CurVal(2))
                ElseIf Val(CurVal(1)) <> Val(CurVal(2)) Then
                    If MsgBox("Variation in Calculated [" & Trim(CurVal(2)) & "] and Existing [" & Trim(CurVal(1)) & "] ½ Forms ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 11, .ActiveRow, Val(CurVal(2))
                End If
            ElseIf .ActiveCol = 12 Then '1F-F&B
                .GetText 12, .ActiveRow, CurVal(1)
                CurVal(2) = CalculateForms("F", .ActiveRow)
                If CurVal(1) = "" Then
                    .SetText 12, .ActiveRow, Val(CurVal(2))
                ElseIf Val(CurVal(1)) <> Val(CurVal(2)) Then
                    If MsgBox("Variation in Calculated [" & Trim(CurVal(2)) & "] and Existing [" & Trim(CurVal(1)) & "] 1 Forms-F&B ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 12, .ActiveRow, Val(CurVal(2))
                End If
            ElseIf .ActiveCol = 13 Then '1F-W&T
                .GetText 13, .ActiveRow, CurVal(1)
                CurVal(2) = CalculateForms("W", .ActiveRow)
                If CurVal(1) = "" Then
                    .SetText 13, .ActiveRow, Val(CurVal(2))
                ElseIf Val(CurVal(1)) <> Val(CurVal(2)) Then
                    If MsgBox("Variation in Calculated [" & Trim(CurVal(2)) & "] and Existing [" & Trim(CurVal(1)) & "] 1 Forms-W&T ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 13, .ActiveRow, Val(CurVal(2))
                End If
                .GetText 25, .ActiveRow, CurVal(1)
                On Error Resume Next
                If CheckEmpty(CurVal(1), False) Then
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "24"
                FrmGeneralMaster.MasterCode = CurVal(1)
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 14, .ActiveRow, slName: .SetText 25, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 15 Then 'Binding Forms
                .GetText 10, .ActiveRow, CurVal(1) '¼F
                BindingForms = BindingForms + Val(CurVal(1))
                .GetText 11, .ActiveRow, CurVal(1) '½F
                BindingForms = BindingForms + Val(CurVal(1))
                .GetText 12, .ActiveRow, CurVal(1) '1F-F&B
                .GetText 6, .ActiveRow, CurVal(2) 'Pages/Form
                If Val(CurVal(2)) <= 12 Then CurVal(1) = Val(CurVal(1)) / 2: CurVal(1) = Int(CurVal(1)) + IIf(CurVal(1) = Int(CurVal(1)), 0, 1)
                BindingForms = BindingForms + Val(CurVal(1))
                .GetText 13, .ActiveRow, CurVal(1) '1F-W&T
                BindingForms = BindingForms + Val(CurVal(1)) 'Calculated Binding Forms
                .GetText 15, .ActiveRow, CurVal(1) 'Binding Forms
                If CurVal(1) = "" Then
                    .SetText 15, .ActiveRow, BindingForms
                ElseIf BindingForms <> Val(CurVal(1)) Then
                    If MsgBox("Variation in Calculated [" & Trim(BindingForms) & "] and Existing [" & Trim(CurVal(1)) & "] Binding Forms ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 15, .ActiveRow, BindingForms
                End If
                    .GetText 17, .ActiveRow, CurVal(1) 'Type
                If CurVal(1) = "" Then
                    If Trim(ReadFromFile("Client ID")) = "Publisher" Then CurVal(2) = "Purchase" Else CurVal(2) = "Sale"
                    .SetText 17, .ActiveRow, CurVal(2)
                 End If
            ElseIf .ActiveCol = 16 Then 'Forms/Sheet
                .GetText 16, .ActiveRow, CurVal(1) 'Forms/Sheet
                .GetText 24, .ActiveRow, CurVal(2) 'C-Forms/Sheet
                If Val(CurVal(1)) <> Val(CurVal(2)) And Val(CurVal(2)) <> 0 Then
                    If MsgBox("Variation in Calculated [" & Trim(Format(Val(CurVal(2)), "#0.00")) & "] and Existing [" & Trim(Format(Val(CurVal(1)), "#0.00")) & "] Forms/Sheet ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 16, .ActiveRow, Val(CurVal(2))
                End If
            End If
        End If
    End With
End Sub
Private Sub fpSpread4_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim CurVal(1 To 2) As Variant, Ups As Double
    With fpSpread4
        If NewCol = 2 Then 'Finishing Size
            .GetText 19, Row, CurVal(1)
            If CheckEmpty(CurVal(1), False) Then .SetText 2, Row, Text5.Text: .SetText 19, Row, FinishSizeCode1
            .GetText 22, Row, CurVal(1) 'M-Printing Size
            If CheckEmpty(CurVal(1), False) Then
                    With rstItemChild
                        If .State = adStateOpen Then .Close
                        .Open "SELECT S.Name+'|'+'Pages/Ptg Form: '+IIF([Ups/Form]<10,'0','')+LTRIM([Ups/Form]) As Col0,S.Name,S.Code FROM FinishSizeChild C INNER JOIN GeneralMaster S ON C.[TextSize]=S.Code WHERE C.Code='" & FinishSizeCode1 & "' ORDER BY S.Name,[Ups/Form]", cnItemMaster, adOpenKeyset, adLockReadOnly
                        SelectionType = "S": TextSizeCode = ""
                        fpSpread4.GetText 3, fpSpread4.ActiveRow, CurVal(2) 'Printing Size
                        If Not CheckEmpty(CurVal(1), False) And .RecordCount > 0 Then 'Move Pointer
                            .Find "[Name] = '" & RTrim(CurVal(1)) & "'"
                            If .EOF Then .MoveFirst Else Text99.Text = .Fields("Col0").Value
                        End If
                        Call LoadSelectionList(rstItemChild, "List of Printing Sizes...", "Name", "")
                        SearchOrder = 0: Text99.Text = ""
                        Call DisplaySelectionList(Text99, TextSizeCode)
                        Call CloseForm(FrmSelectionList)
                        If Not CheckEmpty(TextSizeCode, False) Then
                            .MoveFirst
                            .Find "[Code] = '" & TextSizeCode & "'"
                            fpSpread4.SetText 22, fpSpread4.ActiveRow, TextSizeCode & Right(.Fields("Col0").Value, 2) & .Fields("Name").Value '6+2+40
                        Else
                            fpSpread4.SetText 22, fpSpread4.ActiveRow, "" 'M-Printing Size
                        End If
                        Sendkeys "{ENTER}"
                    End With
            End If
        ElseIf NewCol = 3 Then 'Printing Size
            .GetText 3, Row, CurVal(1)
            If CheckEmpty(CurVal(1), False) Then
                .GetText 22, Row, CurVal(2) 'M-Printing Size
                .SetText 3, Row, Mid(CurVal(2), 9, 60): .SetText 20, Row, Left(CurVal(2), 6) '6+2+40
            End If
        ElseIf NewCol = 5 Then 'Pages/Printing Form
            .GetText 22, Row, CurVal(2) 'M-Printing Size
            Ups = Val(Mid(CurVal(2), 7, 2))
            If Ups = 0 Then Ups = MaxUps("F")
            .GetText 5, Row, CurVal(1)
            If Val(CurVal(1)) = 0 Then .SetText 5, Row, Ups
            .SetText 23, Row, Ups 'C-Ups
        ElseIf NewCol = 6 Then 'Pages/Form
            .GetText 6, Row, CurVal(1)
            If Val(CurVal(1)) = 0 Then
                .GetText 5, Row, CurVal(2)
                .SetText 6, Row, Val(CurVal(2))
            End If
        ElseIf NewCol = 9 Then 'Forms
            .GetText 6, .ActiveRow, CurVal(1) 'Ups
            .GetText 8, .ActiveRow, CurVal(2) 'Pages
            If Val(CurVal(1)) > 0 Then .SetText 9, .ActiveRow, Val(CurVal(2)) / Val(CurVal(1)) 'Forms
        ElseIf NewCol = 16 Then 'Forms/Sheet
            .GetText 5, Row, CurVal(1)
            .GetText 6, Row, CurVal(2)
            If Val(CurVal(1)) > 0 Then Ups = Val(CurVal(2)) / Val(CurVal(1))
            .GetText 16, Row, CurVal(1)
            If Val(CurVal(1)) = 0 Then .SetText 16, Row, Ups
            .SetText 24, Row, Ups 'C-Forms/Sheet
        End If
    End With
    If Col = 8 Or Col = 15 Then CalculateTotalForms
End Sub
Private Sub fpSpread2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CurVal(1 To 2) As Variant, Pages As Variant, Sets As Integer, BalPgs As Integer
    With fpSpread2
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to Delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus: CalculateTotalForms
        ElseIf Shift = 0 And KeyCode = vbKeyDelete Then
            If .ActiveCol = 5 Or .ActiveCol = 6 Then 'Color
                .SetText .ActiveCol, .ActiveRow, "": .SetText .ActiveCol + 12, .ActiveRow, ""
            ElseIf .ActiveCol = 7 Or .ActiveCol = 8 Then 'Plate
                .SetText .ActiveCol, .ActiveRow, ""
            End If
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then 'Element
                .GetText 1, .ActiveRow, CurVal(1)
                Text99.Text = FixQuote(CurVal(1))
                If rstElementList.RecordCount = 0 Then DisplayError ("No Record in Element Master"): .SetActiveCell 1, .ActiveRow: Exit Sub Else rstElementList.MoveFirst
                rstElementList.Find "[Col0] = '" & RTrim(CurVal(1)) & "'"
                SelectionType = "S": ElementCode = ""
                Call LoadSelectionList(rstElementList, "List of Element(s)...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text99, ElementCode)
                Call CloseForm(FrmSelectionList)
                Text99.Text = IIf(CheckEmpty(ElementCode, False), "", Text99.Text)
                .SetText 1, .ActiveRow, Text99.Text
                .SetText 14, .ActiveRow, ElementCode
                If CheckEmpty(ElementCode, False) Then .SetActiveCell 1, .ActiveRow Else Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 3 Then 'Finish Size
                .GetText 15, .ActiveRow, CurVal(1)
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "11"
                FrmGeneralMaster.MasterCode = CurVal(1)
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 3, .ActiveRow, slName: .SetText 15, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then
                    With rstItemChild
                        If .State = adStateOpen Then .Close
                        .Open "SELECT DISTINCT S.Name As Col0,S.Code FROM FinishSizeChild C INNER JOIN GeneralMaster S ON C.TitleSize=S.Code WHERE C.Code='" & slCode & "' ORDER BY S.Name", cnItemMaster, adOpenKeyset, adLockReadOnly
                        SelectionType = "S": TitleSizeCode = ""
                        fpSpread2.GetText 4, fpSpread2.ActiveRow, CurVal(1) 'Printing Size
                        If Not CheckEmpty(CurVal(1), False) And .RecordCount > 0 Then 'Move Pointer
                            .Find "[Col0] = '" & RTrim(CurVal(1)) & "'"
                            If .EOF Then .MoveFirst Else Text99.Text = CurVal(1)
                        End If
                        Call LoadSelectionList(rstItemChild, "List of Printing Sizes...", "Name", "")
                        SearchOrder = 0: Text99.Text = ""
                        Call DisplaySelectionList(Text99, TitleSizeCode)
                        Call CloseForm(FrmSelectionList)
                        If Not CheckEmpty(Trim(TitleSizeCode), False) Then
                            .MoveFirst
                            .Find "[Code] = '" & TitleSizeCode & "'"
                            fpSpread2.SetText 20, fpSpread2.ActiveRow, TitleSizeCode & .Fields("Col0").Value '6+40
                        Else
                            fpSpread2.SetText 20, fpSpread2.ActiveRow, "" 'M-Printing Size
                        End If
                        Sendkeys "{ENTER}"
                    End With
                End If
            ElseIf .ActiveCol = 4 Then 'Printing Size
                .GetText 16, .ActiveRow, CurVal(1)
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "1"
                FrmGeneralMaster.MasterCode = CurVal(1)
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 4, .ActiveRow, slName: .SetText 16, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 5 Then 'Front Color
                .GetText 17, .ActiveRow, CurVal(1)
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "23"
                FrmGeneralMaster.MasterCode = CurVal(1)
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 5, .ActiveRow, slName: .SetText 17, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 6 Then 'Back Color
                .GetText 18, .ActiveRow, CurVal(1)
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "23"
                FrmGeneralMaster.MasterCode = CurVal(1)
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 6, .ActiveRow, slName: .SetText 18, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 7 Then 'Front Plate
                .GetText 23, .ActiveRow, CurVal(1)
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "24"
                FrmGeneralMaster.MasterCode = CurVal(1)
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 7, .ActiveRow, slName: .SetText 23, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 8 Then 'Back Plate
                .GetText 24, .ActiveRow, CurVal(1)
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "24"
                FrmGeneralMaster.MasterCode = CurVal(1)
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 8, .ActiveRow, slName: .SetText 24, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            End If
        ElseIf KeyCode = vbKeyReturn Then
            If .ActiveCol = 2 Then 'Pages
                .GetText 2, .ActiveRow, CurVal(1): .GetText 19, .ActiveRow, CurVal(2)
                If Val(CurVal(1)) <> Val(CurVal(2)) And Val(CurVal(2)) <> 0 Then
                    If MsgBox("Pages [" & Trim(CurVal(1)) & "] are different from that in Master [" & Trim(Format(Val(CurVal(2)), "#0")) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 2, .ActiveRow, Val(CurVal(2))
                End If
            ElseIf .ActiveCol = 4 Then 'Printing Size
                .GetText 4, .ActiveRow, CurVal(1): .GetText 20, .ActiveRow, CurVal(2)
                If Trim(CurVal(1)) <> Trim(Mid(CurVal(2), 7, 60)) And (Not CheckEmpty(Trim(Mid(CurVal(2), 7, 60)), False)) Then
                    If MsgBox("Printing Size [" & Trim(CurVal(1)) & "] is different from that in Master [" & Trim(Mid(CurVal(2), 7, 60)) & "] ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 4, .ActiveRow, Mid(CurVal(2), 7, 60): .SetText 16, .ActiveRow, Left(CurVal(2), 6)
                End If
                .GetText 4, .ActiveRow, CurVal(1): .GetText 3, .ActiveRow, CurVal(2) 'Finishing Size
                If CheckEmpty(CurVal(1), False) Or CheckEmpty(CurVal(2), False) Then Exit Sub
                Dim FL As Double, FR As Double, PL As Double, PR As Double, Ups01 As Integer, Ups02 As Integer
                PL = Val(Left(CurVal(1), InStr(1, CurVal(1), "X") - 1)) + 1: PR = Val(Mid(CurVal(1), InStr(1, CurVal(1), "X") + 1, 5)) + 1: FL = Val(Left(CurVal(2), InStr(1, CurVal(2), "X") - 1)): FR = Val(Mid(CurVal(2), InStr(1, CurVal(2), "X") + 1, 5))
                If Val(PL) * Val(PR) < Val(FL) * Val(FR) Then DisplayError ("Printing Size is smaller than Finish Size"): .SetActiveCell 4, .ActiveRow
            ElseIf .ActiveCol = 10 Then 'Ups/Sheet
                .GetText 10, .ActiveRow, CurVal(1): .GetText 21, .ActiveRow, CurVal(2)
                If Val(CurVal(1)) <> Val(CurVal(2)) And Val(CurVal(2)) <> 0 Then
                    If MsgBox("Variation in Calculated [" & Trim(CurVal(2)) & "] and Existing [" & Trim(CurVal(1)) & "] Ups/Sheet ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 10, .ActiveRow, Val(CurVal(2))
                End If
            ElseIf .ActiveCol = 11 Then 'Sets
                .GetText 11, .ActiveRow, CurVal(1): .GetText 22, .ActiveRow, CurVal(2)
                If Val(CurVal(1)) <> Val(CurVal(2)) And Val(CurVal(2)) <> 0 Then
                    If MsgBox("Variation in Calculated [" & Trim(CurVal(2)) & "] and Existing [" & Trim(CurVal(1)) & "] Sets ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 11, .ActiveRow, Val(CurVal(2))
                End If
                .GetText 2, .ActiveRow, Pages: .GetText 11, .ActiveRow, CurVal(1): .GetText 9, .ActiveRow, CurVal(2) 'Imposition
                BalPgs = Pages - (Val(CurVal(1)) * MaxUps("E") * IIf(CurVal(2) = "F&B", 2, 1)) 'Bal pages calculation
                If BalPgs > 0 Then
                    If MsgBox("[" & BalPgs & "] pages are pending for processing ! Process?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Process !") = vbYes Then
                        .CopyRowRange .ActiveRow, .ActiveRow, .DataRowCnt + 1
                        .SetText 2, .DataRowCnt, BalPgs: .SetText 10, .DataRowCnt, "":    .SetText 11, .DataRowCnt, "":  .SetText 12, .DataRowCnt, ""
                    End If
                    .SetText 2, .ActiveRow, Pages - BalPgs
                End If
            ElseIf .ActiveCol = 12 Then 'Binding Forms
                .GetText 11, .ActiveRow, CurVal(1): .GetText 12, .ActiveRow, CurVal(2)
                If CurVal(2) = "" Then
                    .SetText 12, .ActiveRow, Val(CurVal(1))
                ElseIf Val(CurVal(1)) <> Val(CurVal(2)) Then
                    If MsgBox("Variation in Calculated [" & Trim(CurVal(1)) & "] and Existing [" & Trim(CurVal(2)) & "] Binding Forms ! Change?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then .SetText 12, .ActiveRow, Val(CurVal(1))
                End If
                        .GetText 13, .ActiveRow, CurVal(1) 'Type
                If CurVal(1) = "" Then
                    If Trim(ReadFromFile("Client ID")) = "Publisher" Then CurVal(2) = "Purchase" Else CurVal(2) = "Sale"
                        .SetText 13, .ActiveRow, CurVal(2)
                End If
            End If
        End If
    End With
End Sub
Private Sub fpSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim CurVal(1 To 2) As Variant, Ups As Integer, Sets As Integer, MxUps As Integer
    With fpSpread2
        If NewCol = 2 Then 'GotFocus
            With rstElementList
                If .RecordCount > 0 Then
                    .MoveFirst
                    fpSpread2.GetText 14, Row, CurVal(1)
                    .Find "[Code]='" & CurVal(1) & "'"
                    If Not .EOF Then
                        fpSpread2.SetText 19, Row, Val(.Fields("Pages").Value) 'M-Pages
                        fpSpread2.GetText 2, Row, CurVal(1) 'Pages
                        If Val(CurVal(1)) = 0 Then fpSpread2.SetText 2, Row, Val(.Fields("Pages").Value)
                    End If
                End If
            End With
        ElseIf NewCol = 3 Then 'Finishing Size
            .GetText 15, Row, CurVal(1)
            If CheckEmpty(CurVal(1), False) Then .SetText 3, Row, Text5.Text: .SetText 15, Row, FinishSizeCode1
        ElseIf NewCol = 4 Then 'Printing Size
            .GetText 4, Row, CurVal(1)
            If CheckEmpty(Trim(CurVal(1)), False) Then
                .GetText 20, Row, CurVal(2) 'M-Printing Size
                .SetText 4, Row, Mid(CurVal(2), 7, 60): .SetText 16, Row, Left(CurVal(2), 6) '6+40
            End If
        ElseIf NewCol = 10 Then 'Ups/Sheet
            .GetText 2, Row, CurVal(1) 'Pages
            If Val(CurVal(1)) > 0 Then
                Ups = Int((2 * MaxUps("E")) / Val(CurVal(1)))
                If Ups = 0 Then Ups = 1
                .GetText 10, Row, CurVal(1)
                If Val(CurVal(1)) = 0 Then .SetText 10, Row, Ups
                .SetText 21, Row, Ups
            End If
            .GetText 9, .ActiveRow, CurVal(1)
            If CurVal(1) = "W&T" Then .SetText 6, .ActiveRow, "": .SetText 18, .ActiveRow, "": .SetText 8, .ActiveRow, ""
        ElseIf NewCol = 11 Then 'Sets
            MxUps = MaxUps("E")
            .GetText 2, Row, CurVal(1): .GetText 9, Row, CurVal(2) 'Imposition
            If MxUps > 0 Then Sets = Int((Val(CurVal(1)) / MxUps) * IIf(CurVal(2) = "F&B", 0.5, 1)) 'Sets calculation just like forms calculation
            If Sets = 0 Then Sets = 1
            .SetText 22, Row, Sets 'Calculated Sets
            .GetText 11, Row, CurVal(1)
            If Val(CurVal(1)) = 0 Then .SetText 11, Row, Sets
        End If
    End With
    If Col = 2 Or Col = 12 Then CalculateTotalForms
End Sub
Private Sub fpSpread5_KeyDown(KeyCode As Integer, Shift As Integer) 'Miscellaneous Operation
    Dim CurVal As Variant
    With fpSpread5
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to Delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then 'Element
                .GetText .ActiveCol, .ActiveRow, CurVal
                Text99.Text = FixQuote(CurVal)
                If rstElementList.RecordCount = 0 Then DisplayError ("No Record in Element Master"): .SetActiveCell .ActiveCol, .ActiveRow: Exit Sub Else rstElementList.MoveFirst
                rstElementList.Find "[Col0] = '" & RTrim(CurVal) & "'"
                SelectionType = "S": ElementCode = ""
                Call LoadSelectionList(rstElementList, "List of Element(s)...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text99, ElementCode)
                Call CloseForm(FrmSelectionList)
                .SetText .ActiveCol + 8, .ActiveRow, ElementCode
                If CheckEmpty(ElementCode, False) Then
                    .SetActiveCell .ActiveCol, .ActiveRow: .SetText .ActiveCol, .ActiveRow, ""
                Else
                   .GetText 8, .ActiveRow, CurVal
                    If CheckEmpty(CurVal, False) Then
                    If Trim(ReadFromFile("Client ID")) = "Publisher" Then CurVal = "Purchase" Else CurVal = "Sale"
                    .SetText 8, .ActiveRow, CurVal
                    End If
                    .SetText .ActiveCol, .ActiveRow, Text99.Text: Sendkeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 2 Then 'Operation
                .GetText .ActiveCol + 8, .ActiveRow, OperationCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "7"
                FrmGeneralMaster.MasterCode = OperationCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 8, .ActiveRow, slCode 'Operation
                .SetText .ActiveCol + 4, .ActiveRow, slUGroupName: .SetText .ActiveCol + 10, .ActiveRow, slUGroupCode 'CalcModeName,Code
                .SetText .ActiveCol + 5, .ActiveRow, slUGroupValue1 'CalcValue
                .GetText .ActiveCol + 1, .ActiveRow, CurVal 'Nos
                If CheckEmpty(CurVal, False) Then .SetText .ActiveCol + 1, .ActiveRow, 1
                .GetText .ActiveCol + 2, .ActiveRow, CurVal
                If CheckEmpty(CurVal, False) Then .SetText .ActiveCol + 2, .ActiveRow, "Nos"
                .GetText .ActiveCol + 3, .ActiveRow, CurVal 'Finish Size
                If CheckEmpty(CurVal, False) Then .SetText .ActiveCol + 3, .ActiveRow, Text5.Text: .SetText .ActiveCol + 9, .ActiveRow, FinishSizeCode1
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 5 Then 'Size
                .GetText .ActiveCol, .ActiveRow, CurVal
                Text99.Text = FixQuote(CurVal)
                If rstSizeList.RecordCount = 0 Then DisplayError ("No Record in Size Master"): .SetActiveCell .ActiveCol, .ActiveRow: Exit Sub Else rstSizeList.MoveFirst
                rstSizeList.Find "[Col0] = '" & RTrim(CurVal) & "'"
                SelectionType = "S": SizeCode = ""
                Call LoadSelectionList(rstSizeList, "List of Size(s)...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text99, SizeCode)
                Call CloseForm(FrmSelectionList)
                .SetText .ActiveCol + 6, .ActiveRow, SizeCode
                If CheckEmpty(SizeCode, False) Then
                    .SetActiveCell .ActiveCol, .ActiveRow: .SetText .ActiveCol, .ActiveRow, ""
                Else
                    .SetText .ActiveCol, .ActiveRow, Text99.Text: Sendkeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 6 Then ' Calc Mode
                .GetText .ActiveCol + 6, .ActiveRow, CalcModeCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "20"
                FrmGeneralMaster.MasterCode = CalcModeCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 6, .ActiveRow, slCode: .SetText .ActiveCol + 1, .ActiveRow, slValue1
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            End If
        End If
    End With
End Sub
Private Sub fpSpread6_KeyDown(KeyCode As Integer, Shift As Integer) 'Binding Element
    Dim CurVal As Variant, oBindingTypeCode As Variant
    With fpSpread6
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to Delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then
                .GetText .ActiveCol + 9, .ActiveRow, oBindingTypeCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "6"
                FrmGeneralMaster.MasterCode = oBindingTypeCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 9, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 2 Then
                .GetText .ActiveCol + 9, .ActiveRow, BinderyProcessCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "7"
                FrmGeneralMaster.MasterCode = BinderyProcessCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 9, .ActiveRow, slCode
                .SetText .ActiveCol + 2, .ActiveRow, "Nos"
                .SetText .ActiveCol + 3, .ActiveRow, Text5.Text
                .SetText .ActiveCol + 10, .ActiveRow, FinishSizeCode1
                .SetText .ActiveCol + 4, .ActiveRow, 1
                .SetText .ActiveCol + 5, .ActiveRow, slUGroupName
                .SetText .ActiveCol + 6, .ActiveRow, slUGroupValue1
                .SetText .ActiveCol + 11, .ActiveRow, slUGroupCode
                If Trim(ReadFromFile("Client ID")) = "Publisher" Then CurVal = "Purchase" Else CurVal = "Sale"
                .SetText 9, .ActiveRow, CurVal
                .GetText .ActiveCol + 8, .ActiveRow, CurVal
                If CheckEmpty(CurVal, False) Then .GetText .ActiveCol + 8, .ActiveRow - 1, CurVal: .SetText 10, .ActiveRow, CurVal
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 5 Then
                .GetText .ActiveCol, .ActiveRow, CurVal
                Text99.Text = FixQuote(CurVal)
                If rstSizeList.RecordCount = 0 Then DisplayError ("No Record in Size Master"): .SetActiveCell .ActiveCol, .ActiveRow: Exit Sub Else rstSizeList.MoveFirst
                rstSizeList.Find "[Col0] = '" & RTrim(CurVal) & "'"
                SelectionType = "S": SizeCode = ""
                Call LoadSelectionList(rstSizeList, "List of Size(s)...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text99, SizeCode)
                Call CloseForm(FrmSelectionList)
                .SetText .ActiveCol + 7, .ActiveRow, SizeCode
                If CheckEmpty(SizeCode, False) Then
                    .SetActiveCell .ActiveCol, .ActiveRow: .SetText .ActiveCol, .ActiveRow, ""
                Else
                    .SetText .ActiveCol, .ActiveRow, Text99.Text: Sendkeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 7 Then
                .GetText .ActiveCol + 6, .ActiveRow, CalcModeCode
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "20"
                FrmGeneralMaster.MasterCode = CalcModeCode
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText .ActiveCol + 6, .ActiveRow, slCode: .SetText .ActiveCol + 1, .ActiveRow, slValue1
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            End If
        End If
    End With
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstItemList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    With rstItemMaster
        If .State = adStateOpen Then .Close
        .Open "SELECT *,ISNULL((Select OpBal From BookChild Where MaterialCentre='000005' And Item=P.Code),0) opStockA,ISNULL((Select OpBal From BookChild Where MaterialCentre='005995' And Item=P.Code),0) opStockB FROM BookMaster P WHERE Code='" & FixQuote(rstItemList.Fields("Code").Value) & "'", cnItemMaster, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End With
    LoadMasterList
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text11.Text = ""
    Text7.Text = "998912"
    With rstHSNCodeList
        If .RecordCount > 0 Then .MoveFirst
        .Find "[Col0] = '" & Trim(Text7.Text) & "'"
        If Not .EOF Then HSNCode = .Fields("Code").Value
    End With
    MhRealInput1.Value = 0
    MhRealInput17.Value = 0
    Text8.Text = "": ItemGroupCode = ""
    Text5.Text = "": FinishSizeCode1 = ""
    Text10.Text = "": BindingTypeCode = ""
    Text13.Text = ""
    Text15.Text = ""
    Text16.Text = "": IntegrationUnitCode = ""
    Text17.Text = ""
    Text18.Text = "": MFCode = ""
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
    Text22.Text = ""
    MhRealInput15.Value = 0
    MhRealInput16.Value = 0
    MhRealInput7.Value = 0
    MhRealInput4.Value = 0
    MhRealInput2.Value = 0
    MhRealInput3.Value = 0
    MhRealInput5.Value = 0
    MhRealInput6.Value = 0
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    fpSpread3.ClearRange 1, 1, fpSpread3.MaxCols, fpSpread3.MaxRows, True: fpSpread3.SetActiveCell 1, 1
    fpSpread4.ClearRange 1, 1, fpSpread4.MaxCols, fpSpread4.MaxRows, True: fpSpread4.SetActiveCell 1, 1
    fpSpread2.ClearRange 1, 1, fpSpread2.MaxCols, fpSpread2.MaxRows, True: fpSpread2.SetActiveCell 1, 1
    fpSpread5.ClearRange 1, 1, fpSpread5.MaxCols, fpSpread5.MaxRows, True: fpSpread5.SetActiveCell 1, 1
    fpSpread6.ClearRange 1, 1, fpSpread6.MaxCols, fpSpread6.MaxRows, True: fpSpread6.SetActiveCell 1, 1
    imgFile = "": cmdUpload.Enabled = True: 'Image1.Picture = Nothing
End Sub
Private Sub LoadFields()
    With rstItemMaster
        If .EOF Or .BOF Then Exit Sub
        Text2.Text = .Fields("Name").Value
        Text3.Text = .Fields("PrintName").Value
        If Not IsNull(.Fields("ISBN").Value) Then Text4.Text = .Fields("ISBN").Value
        If Not IsNull(.Fields("BusyCode").Value) Then Text11.Text = .Fields("BusyCode").Value
        HSNCode = .Fields("HSNCode").Value
        If rstHSNCodeList.RecordCount > 0 Then rstHSNCodeList.MoveFirst
        rstHSNCodeList.Find "[Code] = '" & HSNCode & "'"
        If Not rstHSNCodeList.EOF Then Text7.Text = rstHSNCodeList.Fields("Col0").Value
        MhRealInput1.Value = Val(.Fields("Price").Value)
        MhRealInput17.Value = Val(.Fields("Weight").Value)
        ItemGroupCode = .Fields("Group").Value
        If rstItemGroupList.RecordCount > 0 Then rstItemGroupList.MoveFirst
        rstItemGroupList.Find "[Code] = '" & ItemGroupCode & "'"
        If Not rstItemGroupList.EOF Then Text8.Text = rstItemGroupList.Fields("Col0").Value
        FinishSizeCode1 = .Fields("FinishSize").Value
        rstFinishSizeList.MoveFirst
        rstFinishSizeList.Find "[Code] = '" & FinishSizeCode1 & "'"
        Text5.Text = rstFinishSizeList.Fields("Col0").Value
        If Not IsNull(.Fields("BindingType").Value) Then
            BindingTypeCode = .Fields("BindingType").Value
            If rstBindingTypeList.RecordCount > 0 Then rstBindingTypeList.MoveFirst
            rstBindingTypeList.Find "[Code] = '" & BindingTypeCode & "'"
            If Not rstBindingTypeList.EOF Then Text10.Text = rstBindingTypeList.Fields("Col0").Value
        End If
        If Not IsNull(.Fields("Remarks").Value) Then Text13.Text = .Fields("Remarks").Value
        If Not IsNull(.Fields("ItemMarks").Value) Then Text17.Text = .Fields("ItemMarks").Value
        If Not IsNull(.Fields("MFCD").Value) Then MFCode = .Fields("MFCD").Value
        If rstMFList.RecordCount > 0 Then rstMFList.MoveFirst
        rstMFList.Find "[Code] = '" & MFCode & "'"
        If Not rstMFList.EOF Then Text18.Text = rstMFList.Fields("Col0").Value
        If Not IsNull(.Fields("S_No").Value) Then Text19.Text = .Fields("S_No").Value
        If Not IsNull(.Fields("PurDate").Value) Then Text20.Text = Format(.Fields("PurDate").Value, "dd-MM-yyyy")
        If Not IsNull(.Fields("RoomA").Value) Then Text21.Text = .Fields("RoomA").Value
        If Not IsNull(.Fields("RoomB").Value) Then Text22.Text = .Fields("RoomB").Value
        Text15.Text = .Fields("ItemIntegrationName").Value
        If Not IsNull(.Fields("IntegrationUnit").Value) Then IntegrationUnitCode = .Fields("IntegrationUnit").Value
        If rstUnitList.RecordCount > 0 Then rstUnitList.MoveFirst
        rstUnitList.Find "[Code] = '" & IntegrationUnitCode & "'"
        If Not rstUnitList.EOF Then Text16.Text = rstUnitList.Fields("Col0").Value
        
        If Not IsNull(.Fields("Notes").Value) Then txtNotes.Text = .Fields("Notes").Value
        MhRealInput2.Value = .Fields("SalePrice").Value
        MhRealInput3.Value = .Fields("PurPrice").Value
        MhRealInput5.Value = .Fields("opStockA").Value
        MhRealInput6.Value = .Fields("opStockB").Value
        Call LoadItemList(.Fields("Code").Value)
        Call CalculateTotalForms
        If Dir(App.Path & "\Pic\Item [" & CompCode & "]", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Pic\Item [" & CompCode & "]"
        If Not (rstPicMaster.EOF) Then
            If rstPicMaster.RecordCount > 0 Then .MoveFirst
            If Not CheckEmpty(rstPicMaster.Fields("PicData"), False) Then
                imgFile = App.Path & "\Pic\Item [" & CompCode & "]" & "\" & rstPicMaster.Fields("PicName").Value & "." & rstPicMaster.Fields("PicType").Value
                RetrievePic rstPicMaster.Fields("PicData").Value, imgFile, srmPicMgr
                Call cmdView_Click
            End If
        Else
            imgFile = "": Call cmdView_Click
        End If
    End With
End Sub
Private Sub LoadItemList(ByVal MasterCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    With rstItemChild
        If .State = adStateOpen Then .Close
        .Open "SELECT * FROM (SELECT Category,Item As ItemCode,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=C.Item),IIF(Category='2',(SELECT P.Name+' (UOM : '+LTRIM(U.Name)+')' As Name FROM PaperMaster P INNER JOIN GeneralMaster U ON P.UOM=U.Code WHERE P.Code=C.Item),(SELECT Name FROM BookMaster WHERE Code=C.Item))) As ItemName,Quantity FROM BookChild01 C WHERE C.Code='" & MasterCode & "') As Tbl ORDER BY Category,ItemName", cnItemMaster, adOpenKeyset, adLockReadOnly
        i = 0
        Do Until .EOF
            i = i + 1
            fpSpread1.SetText 1, i, Choose(Val(.Fields("Category").Value), "BOM Item", "Paper", "Finished Item", "Unfinished Item"): fpSpread1.SetText 4, i, .Fields("ItemCode").Value
            fpSpread1.SetText 2, i, .Fields("ItemName").Value
            fpSpread1.SetText 3, i, Val(.Fields("Quantity").Value)
            .MoveNext
        Loop
        fpSpread1.SetActiveCell 1, 1
        If .State = adStateOpen Then .Close
        .Open "SELECT Sno,Correction,ArrivedOn,TargetDate,RectifiedOn,(Select Name From TeamMemberMaster Where Code=Member) As MemberName,Member As MemberCode,Remarks FROM BookChild0201 T WHERE Code='" & MasterCode & "'  ORDER BY ArrivedOn DESC,SNo", cnItemMaster, adOpenKeyset, adLockReadOnly
        i = 0
        Do Until .EOF
            i = i + 1
            fpSpread3.SetText 1, i, Format(.Fields("ArrivedOn").Value, "dd-MM-yyyy")
            fpSpread3.SetText 2, i, .Fields("Correction").Value
            If Not IsNull(.Fields("TargetDate").Value) Then fpSpread3.SetText 3, i, Format(.Fields("TargetDate").Value, "dd-MM-yyyy")
            If Not IsNull(.Fields("RectifiedOn").Value) Then fpSpread3.SetText 4, i, Format(.Fields("RectifiedOn").Value, "dd-MM-yyyy")
            fpSpread3.SetText 5, i, .Fields("MemberName").Value
            fpSpread3.SetText 6, i, .Fields("Remarks").Value
            fpSpread3.SetText 7, i, .Fields("MemberCode").Value
            .MoveNext
        Loop
        fpSpread3.SetActiveCell 1, 1
        If .State = adStateOpen Then .Close
        .Open "SELECT Element,E.Name As ElementName,FinishSize,F.Name As FinishSizeName,[Size] As PrintingSize,P.Name As PrintingSizeName,Color,R.Name As ColorName,PlateType,R1.Name As PlateName,C.DuplexPrinting,[Pages/PrintingForm],[Pages/Form],C.Pages,Forms,[Forms-¼],[Forms-½],[Forms-1-F&B],[Forms-1-W&T],BindingForms,Ups,C.Type FROM ((((BookChild05 C INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster F ON C.FinishSize=F.Code) INNER JOIN GeneralMaster P ON C.[Size]=P.Code) INNER JOIN GeneralMaster R ON C.Color=R.Code) INNER JOIN GeneralMaster R1 ON C.PlateType=R1.Code WHERE C.Code='" & MasterCode & "' ORDER BY C.Type,E.Name,F.Name,P.Name,R.Name", cnItemMaster, adOpenKeyset, adLockReadOnly
        i = 0
        Do Until .EOF
            i = i + 1
            fpSpread4.SetText 1, i, .Fields("ElementName").Value: fpSpread4.SetText 18, i, .Fields("Element").Value
            fpSpread4.SetText 2, i, .Fields("FinishSizeName").Value: fpSpread4.SetText 19, i, .Fields("FinishSize").Value
            fpSpread4.SetText 3, i, .Fields("PrintingSizeName").Value: fpSpread4.SetText 20, i, .Fields("PrintingSize").Value
            fpSpread4.SetText 4, i, IIf(.Fields("DuplexPrinting").Value, 1, 0)
            fpSpread4.SetText 5, i, Val(.Fields("Pages/PrintingForm").Value)
            fpSpread4.SetText 6, i, Val(.Fields("Pages/Form").Value)
            fpSpread4.SetText 7, i, .Fields("ColorName").Value: fpSpread4.SetText 21, i, .Fields("Color").Value
            fpSpread4.SetText 8, i, Val(.Fields("Pages").Value)
            fpSpread4.SetText 9, i, Val(.Fields("Forms").Value)
            fpSpread4.SetText 10, i, Val(.Fields("Forms-¼").Value)
            fpSpread4.SetText 11, i, Val(.Fields("Forms-½").Value)
            fpSpread4.SetText 12, i, Val(.Fields("Forms-1-F&B").Value)
            fpSpread4.SetText 13, i, Val(.Fields("Forms-1-W&T").Value)
            fpSpread4.SetText 14, i, .Fields("PlateName").Value: fpSpread4.SetText 25, i, .Fields("PlateType").Value 'Choose(Val(.Fields("PlateType").Value), "Deep-Etch", "PS", "Wipe-on", "CTP")
            fpSpread4.SetText 15, i, Val(.Fields("BindingForms").Value)
            fpSpread4.SetText 16, i, Val(.Fields("Ups").Value)
'            fpSpread4.SetText 17, i, IIf(.Fields("Type").Value = "S", "Sale", "Purchase")
             If InStr(1, "FS_S", Trim(.Fields("Type").Value)) > 0 Then
                fpSpread4.SetText 17, i, "Sale"
            Else 'If InStr(1, "FP_P", Trim(.Fields("Type").Value)) > 0 Then
                fpSpread4.SetText 17, i, "Purchase"
            End If
            .MoveNext
        Loop
        fpSpread4.SetActiveCell 1, 1
        If .State = adStateOpen Then .Close
        .Open "SELECT Element,E.Name As ElementName,FinishSize,F.Name As FinishSizeName,C.[Size] As PrintingSize,P.Name As PrintingSizeName,FrontPrintingType As FrontColor,R1.Name As FrontColorName,BackPrintingType As BackColor,R2.Name As BackColorName,PlateType As FrontPlateType,R3.Name As FrontPlateName,PlateTypeBack As BackPlateType,R4.Name As BackPlateName,C.Pages,Imposition,Ups,C.Sets,BindingForms,C.Type FROM ((((((BookChild06 C INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster F ON C.FinishSize=F.Code) INNER JOIN GeneralMaster P ON C.[Size]=P.Code) LEFT JOIN GeneralMaster R1 ON C.FrontPrintingType=R1.Code) LEFT JOIN GeneralMaster R2 ON C.BackPrintingType=R2.Code) LEFT JOIN GeneralMaster R3 ON C.PlateType=R3.Code) LEFT JOIN GeneralMaster R4 ON C.PlateTypeBack=R4.Code WHERE C.Code='" & MasterCode & "' ORDER BY C.Type,E.Name,F.Name,P.Name", cnItemMaster, adOpenKeyset, adLockReadOnly
        i = 0
        Do Until .EOF
            i = i + 1
            fpSpread2.SetText 1, i, .Fields("ElementName").Value: fpSpread2.SetText 14, i, .Fields("Element").Value
            fpSpread2.SetText 2, i, Val(.Fields("Pages").Value)
            fpSpread2.SetText 3, i, .Fields("FinishSizeName").Value: fpSpread2.SetText 15, i, .Fields("FinishSize").Value
            fpSpread2.SetText 4, i, .Fields("PrintingSizeName").Value: fpSpread2.SetText 16, i, .Fields("PrintingSize").Value
            fpSpread2.SetText 5, i, .Fields("FrontColorName").Value: fpSpread2.SetText 17, i, .Fields("FrontColor").Value
            fpSpread2.SetText 6, i, .Fields("BackColorName").Value: fpSpread2.SetText 18, i, .Fields("BackColor").Value
            fpSpread2.SetText 7, i, .Fields("FrontPlateName").Value: fpSpread2.SetText 23, i, .Fields("FrontPlateType").Value 'Choose(Val(.Fields("FrontPlateType").Value), "Deep-Etch", "PS", "Wipe-on", "CTP")
            fpSpread2.SetText 8, i, .Fields("BackPlateName").Value: fpSpread2.SetText 24, i, .Fields("BackPlateType").Value 'Choose(Val(.Fields("BackPlateType").Value), "Deep-Etch", "PS", "Wipe-on", "CTP")
            fpSpread2.SetText 9, i, IIf(.Fields("Imposition").Value = "F", "F&B", "W&T")
            fpSpread2.SetText 10, i, Val(.Fields("Ups").Value)
            fpSpread2.SetText 11, i, Val(.Fields("Sets").Value)
            fpSpread2.SetText 12, i, Val(.Fields("BindingForms").Value)
'            fpSpread2.SetText 13, i, IIf(.Fields("Type").Value = "S", "Sale", "Purchase")
             If InStr(1, "FS_S", Trim(.Fields("Type").Value)) > 0 Then
                fpSpread2.SetText 13, i, "Sale"
            Else 'If InStr(1, "FP_P", Trim(.Fields("Type").Value)) > 0 Then
                fpSpread2.SetText 13, i, "Purchase"
            End If
            .MoveNext
        Loop
        fpSpread2.SetActiveCell 1, 1
        If .State = adStateOpen Then .Close
        .Open "SELECT Element,E.Name As ElementName,Operation,O.Name As OperationName,CalcMode,CalcValue,M.Name As CalcModeName,C.[Size],S.Name As SizeName,Number,OperationCountName,C.Type FROM (((BookChild07 C INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) LEFT JOIN GeneralMaster S ON C.[Size]=S.Code WHERE C.Code='" & MasterCode & "' ORDER BY C.Type,E.Name,O.Name,M.Name,S.Name", cnItemMaster, adOpenKeyset, adLockReadOnly
        i = 0
        Do Until .EOF
            i = i + 1
            fpSpread5.SetText 1, i, .Fields("ElementName").Value: fpSpread5.SetText 9, i, .Fields("Element").Value
            fpSpread5.SetText 2, i, .Fields("OperationName").Value: fpSpread5.SetText 10, i, .Fields("Operation").Value
            fpSpread5.SetText 3, i, Val(.Fields("Number").Value)
            fpSpread5.SetText 4, i, .Fields("OperationCountName").Value
            fpSpread5.SetText 5, i, .Fields("SizeName").Value: fpSpread5.SetText 11, i, .Fields("Size").Value
            fpSpread5.SetText 6, i, .Fields("CalcModeName").Value: fpSpread5.SetText 12, i, .Fields("CalcMode").Value
            fpSpread5.SetText 7, i, Val(.Fields("CalcValue").Value)
'            fpSpread5.SetText 8, i, IIf(Trim(.Fields("Type").Value) = "S", "Sale", "Purchase")
             If InStr(1, "FS_S", Trim(.Fields("Type").Value)) > 0 Then
                fpSpread5.SetText 8, i, "Sale"
            Else 'If InStr(1, "FP_P", Trim(.Fields("Type").Value)) > 0 Then
                fpSpread5.SetText 8, i, "Purchase"
            End If
            .MoveNext
        Loop
        fpSpread5.SetActiveCell 1, 1
        If .State = adStateOpen Then .Close
        .Open "SELECT SubItem AS SubItemCode,E.Name As SubItemName,C.BindingType,B.Name As BindingTypeName,BinderyProcess,P.Name As BinderyProcessName,CalcMode,M.Name As CalcModeName,CalcValue,C.[Size],S.Name As SizeName,S.Type AS sType,Number,OperationCountName,C.Type FROM ((((BookChild08 C INNER JOIN GeneralMaster B ON C.BindingType=B.Code)INNER JOIN GeneralMaster P ON C.BinderyProcess=P.Code) INNER JOIN GeneralMaster M ON C.CalcMode=M.Code) INNER JOIN BOOKMaster E ON C.SubItem=E.Code) INNER JOIN GeneralMaster S ON C.[Size]=S.Code WHERE C.Code='" & MasterCode & "' ORDER BY C.Type,C.BindingType,E.Name,P.Name,M.Name,S.Name", cnItemMaster, adOpenKeyset, adLockReadOnly
        i = 0
        Do Until .EOF
            i = i + 1
Dim BindingType As String
            If i = 1 Or BindingType <> .Fields("BindingType").Value Then fpSpread6.SetText 1, i, .Fields("BindingTypeName").Value
            fpSpread6.SetText 10, i, .Fields("BindingType").Value: BindingType = .Fields("BindingType").Value
            fpSpread6.SetText 2, i, .Fields("BinderyProcessName").Value: fpSpread6.SetText 11, i, .Fields("BinderyProcess").Value
            fpSpread6.SetText 3, i, Val(.Fields("Number").Value)
            fpSpread6.SetText 4, i, .Fields("OperationCountName").Value
            fpSpread6.SetText 5, i, .Fields("SizeName").Value: fpSpread6.SetText 12, i, .Fields("Size").Value
            fpSpread6.SetText 6, i, IIf(.Fields("sType").Value = 11, 1, 0)
            fpSpread6.SetText 7, i, .Fields("CalcModeName").Value: fpSpread6.SetText 8, i, Val(.Fields("CalcValue").Value): fpSpread6.SetText 13, i, .Fields("CalcMode").Value
'            fpSpread6.SetText 9, i, IIf(.Fields("Type").Value = "FS", "Sale", "Purchase")
            If InStr(1, "FS_S", Trim(.Fields("Type").Value)) > 0 Then
                fpSpread6.SetText 9, i, "Sale"
            Else 'If InStr(1, "FP_P", Trim(.Fields("Type").Value)) > 0 Then
                fpSpread6.SetText 9, i, "Purchase"
                
            End If
            .MoveNext
        Loop
        fpSpread6.SetActiveCell 1, 1
    End With
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    With rstItemMaster
        If .RecordCount = 0 Then Exit Sub
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseServer
        .Open "SELECT * FROM BookMaster WHERE Code='" & FixQuote(rstItemList.Fields("Code").Value) & "'", cnItemMaster, adOpenKeyset, adLockPessimistic
        MdiMainMenu.MousePointer = vbHourglass
        .Fields("Printstatus") = "N"
    End With
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    cnItemMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    With rstItemMaster
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnItemMaster, "SELECT MAX(Code) FROM BookMaster", 6, "0")
            .Fields("CreatedBy").Value = UserCode
            .Fields("CreatedOn").Value = Now()
            .Fields("RecordStatus").Value = "N"
        Else
            .Fields("ModifiedBy").Value = UserCode
            .Fields("ModifiedOn").Value = Now()
            .Fields("RecordStatus").Value = "M"
        End If
        .Fields("Name").Value = Trim(Text2.Text)
        .Fields("PrintName").Value = Trim(Text3.Text)
        .Fields("ISBN").Value = Trim(Text4.Text)
        .Fields("BusyCode").Value = Trim(Text11.Text)
        .Fields("HSNCode").Value = HSNCode
        .Fields("Price").Value = MhRealInput1.Value
        .Fields("Weight").Value = MhRealInput17.Value
        .Fields("Group").Value = ItemGroupCode
        .Fields("FinishSize").Value = IIf(CheckEmpty(FinishSizeCode1, False), "*11098", FinishSizeCode1) 'FinishSizeCode1
        .Fields("BindingType").Value = IIf(CheckEmpty(BindingTypeCode, False), Null, BindingTypeCode)
        .Fields("Remarks").Value = Trim(Text13.Text)
        .Fields("ItemMarks").Value = Trim(Text17.Text)
        .Fields("S_No").Value = Trim(Text19.Text)
    If Trim(ReadFromFile("Customer Type")) = "General" Then
        .Fields("RoomA").Value = Trim(Text21.Text)
        .Fields("RoomB").Value = Trim(Text22.Text)
    Else
        .Fields("RoomA").Value = 0
        .Fields("RoomB").Value = 0
    End If
        .Fields("MFCD").Value = MFCode
        .Fields("ItemIntegrationName").Value = Trim(Text15.Text)
        .Fields("Type").Value = ItemType
        .Fields("IntegrationUnit").Value = IntegrationUnitCode
        .Fields("PrintStatus").Value = "N"
        .Fields("Notes").Value = txtNotes.Text
        .Fields("SalePrice").Value = MhRealInput2.Value
        .Fields("PurPrice").Value = MhRealInput3.Value
    End With
End Sub
Private Function UpdateItemList(ByVal ActionType As String) As Boolean
    On Error GoTo ErrorHandler
    Dim CellVal(1 To 18) As Variant, RectifiedON As String
    UpdateItemList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType <> "I" Then
        cnItemMaster.Execute "DELETE FROM BookChild01 WHERE Code='" & rstItemMaster.Fields("Code").Value & "'"
        cnItemMaster.Execute "DELETE FROM BookChild02 WHERE Code='" & rstItemMaster.Fields("Code").Value & "'"
        cnItemMaster.Execute "DELETE FROM BookChild0201 WHERE Code='" & rstItemMaster.Fields("Code").Value & "'"
        cnItemMaster.Execute "DELETE FROM BookChild05 WHERE Code='" & rstItemMaster.Fields("Code").Value & "'"
        cnItemMaster.Execute "DELETE FROM BookChild06 WHERE Code='" & rstItemMaster.Fields("Code").Value & "'"
        cnItemMaster.Execute "DELETE FROM BookChild07 WHERE Code='" & rstItemMaster.Fields("Code").Value & "'"
        cnItemMaster.Execute "DELETE FROM BookChild08 WHERE Code='" & rstItemMaster.Fields("Code").Value & "'"
    Else
        If SSTab1.Tab = 2 Then 'BOM
            With fpSpread1
                .GetText 1, .ActiveRow, CellVal(1) 'Category
                .GetText 4, .ActiveRow, CellVal(2) 'Item
                .GetText 3, .ActiveRow, CellVal(3) 'Quantity
                CellVal(1) = IIf(CellVal(1) = "BOM Item", "1", IIf(CellVal(1) = "Paper", "2", IIf(CellVal(1) = "Finished Item", "3", "4")))
            End With
            cnItemMaster.Execute "INSERT INTO BookChild01 VALUES ('" & rstItemMaster.Fields("Code").Value & "','" & CellVal(1) & "','" & CellVal(2) & "'," & Val(CellVal(3)) & ")"
        ElseIf SSTab1.Tab = 3 Then 'Editorial
            With fpSpread3
                .GetText 1, .ActiveRow, CellVal(1) 'Arrived ON
                .GetText 2, .ActiveRow, CellVal(2) 'Correction fpSpread3.ActiveRow
                .GetText 3, .ActiveRow, CellVal(3) 'Target Date
                .GetText 4, .ActiveRow, CellVal(4) 'Rectified ON
                .GetText 6, .ActiveRow, CellVal(5) 'Remarks
                .GetText 7, .ActiveRow, CellVal(6) 'Member
            End With
            If IsDate(CellVal(3)) Then RectifiedON = "'" & Format(GetDate(CellVal(3)), "MM-dd-yyyy") & "'" Else RectifiedON = "Null"
            'cnItemMaster.Execute "INSERT INTO BookChild02 VALUES ('" & rstItemMaster.Fields("Code").Value & "'," & fpSpread3.ActiveRow & ",'" & Format(GetDate(CellVal(1)), "MM-dd-yyyy") & "',Null,'" & CellVal(2) & "',Null,Null,Null," & RectifiedON & ",Null,'','P')"
            cnItemMaster.Execute "INSERT INTO BookChild0201 VALUES ('" & rstItemMaster.Fields("Code").Value & "'," & fpSpread3.ActiveRow & ",'" & CellVal(2) & "','" & Format(GetDate(CellVal(1)), "MM-dd-yyyy") & "','" & Format(GetDate(CellVal(3)), "MM-dd-yyyy") & "','" & Format(GetDate(CellVal(4)), "MM-dd-yyyy") & "','" & CellVal(6) & "','" & CellVal(5) & "')"
        ElseIf SSTab1.Tab = 4 Then 'MultiForm
            With fpSpread4
                .GetText 18, .ActiveRow, CellVal(1) 'Element
                .GetText 1, .ActiveRow, CellVal(18) 'Element Print Name
                .GetText 19, .ActiveRow, CellVal(2) 'Finish Size
                .GetText 20, .ActiveRow, CellVal(3) 'Printing Size
                .GetText 21, .ActiveRow, CellVal(4) 'Color
                .GetText 25, .ActiveRow, CellVal(5) 'Plate
                'CellVal(5) = IIf(CellVal(5) = "Deep-Etch", 1, IIf(CellVal(5) = "PS", 2, IIf(CellVal(5) = "Wipe-on", 3, 4)))
                .GetText 4, .ActiveRow, CellVal(6)  'Duplex Printing
                CellVal(6) = IIf(Val(CellVal(6)) = 1, 1, 0)
                .GetText 5, .ActiveRow, CellVal(7)  'Pages/Printing Form
                .GetText 6, .ActiveRow, CellVal(8)  'Pages/Form
                .GetText 8, .ActiveRow, CellVal(9)  'Pages
                .GetText 9, .ActiveRow, CellVal(10)  'Forms
                .GetText 10, .ActiveRow, CellVal(11) 'Forms-¼
                .GetText 11, .ActiveRow, CellVal(12) 'Forms-½
                .GetText 12, .ActiveRow, CellVal(13) 'Forms-F&B
                .GetText 13, .ActiveRow, CellVal(14) 'Forms-W&T
                .GetText 15, .ActiveRow, CellVal(15) 'Binding Forms
                .GetText 16, .ActiveRow, CellVal(16) 'Forms/Sheet
                .GetText 17, .ActiveRow, CellVal(17) 'Type
                CellVal(17) = IIf(CellVal(17) = "Sale", "S", "P")
            End With
            cnItemMaster.Execute "INSERT INTO BookChild05 VALUES ('" & rstItemMaster.Fields("Code").Value & "','" & rstItemMaster.Fields("Code").Value & "','" & CellVal(1) & "','" & CellVal(18) & "','" & CellVal(2) & "','" & CellVal(3) & "'," & CellVal(6) & "," & Val(CellVal(7)) & "," & Val(CellVal(8)) & ",'" & CellVal(4) & "'," & Val(CellVal(9)) & "," & Val(CellVal(10)) & "," & Val(CellVal(11)) & "," & Val(CellVal(12)) & "," & Val(CellVal(13)) & "," & Val(CellVal(14)) & ",'" & CellVal(5) & "'," & Val(CellVal(16)) & "," & Val(CellVal(15)) & ",'" & CellVal(17) & "')"
        ElseIf SSTab1.Tab = 5 Then 'MultiElement
            With fpSpread2
                .GetText 14, .ActiveRow, CellVal(1) 'Element
                .GetText 1, .ActiveRow, CellVal(17) 'Element Print Name
                .GetText 15, .ActiveRow, CellVal(2) 'Finish Size
                .GetText 16, .ActiveRow, CellVal(3) 'Printing Size
                .GetText 17, .ActiveRow, CellVal(4)
                    CellVal(4) = IIf(CellVal(4) = "", "NULL", "'" & CellVal(4) & "'")
                .GetText 18, .ActiveRow, CellVal(5)
                    CellVal(5) = IIf(CellVal(5) = "", "NULL", "'" & CellVal(5) & "'")
                .GetText 2, .ActiveRow, CellVal(6)  'Pages
                .GetText 23, .ActiveRow, CellVal(7) 'Plate-Front
                    CellVal(7) = IIf(CellVal(7) = "", "NULL", "'" & CellVal(7) & "'")
                .GetText 24, .ActiveRow, CellVal(8) 'Plate-Back
                    CellVal(8) = IIf(CellVal(8) = "", "NULL", "'" & CellVal(8) & "'")
                .GetText 9, .ActiveRow, CellVal(9) 'Imposition
                    CellVal(9) = IIf(CellVal(9) = "F&B", "F", "W")
                .GetText 10, .ActiveRow, CellVal(10) 'Ups/Sheet
                .GetText 11, .ActiveRow, CellVal(11) 'Sets
                .GetText 12, .ActiveRow, CellVal(12) 'Binding Forms
                .GetText 13, .ActiveRow, CellVal(13) 'Type
                    CellVal(13) = IIf(CellVal(13) = "Sale", "S", "P")
            End With
            cnItemMaster.Execute "INSERT INTO BookChild06 VALUES ('" & rstItemMaster.Fields("Code").Value & "','" & rstItemMaster.Fields("Code").Value & "','" & CellVal(1) & "','" & CellVal(17) & "'," & Val(CellVal(6)) & ",'" & CellVal(2) & "','" & CellVal(3) & "','" & CellVal(9) & "'," & CellVal(4) & "," & CellVal(5) & "," & CellVal(7) & "," & CellVal(8) & "," & Val(CellVal(10)) & "," & Val(CellVal(11)) & "," & Val(CellVal(12)) & ",'" & CellVal(13) & "')"
        ElseIf SSTab1.Tab = 6 Then 'Misc Operation
            With fpSpread5
                .GetText 9, .ActiveRow, CellVal(1) 'Element
                .GetText 10, .ActiveRow, CellVal(2) 'Operation
                .GetText 3, .ActiveRow, CellVal(5) 'Number
                .GetText 4, .ActiveRow, CellVal(7) 'OperationCountName
                .GetText 11, .ActiveRow, CellVal(3) 'Size
                .GetText 12, .ActiveRow, CellVal(4) 'Calc Mode
                .GetText 7, .ActiveRow, CellVal(8) 'Calc Value
                .GetText 8, .ActiveRow, CellVal(6) 'Type
                CellVal(6) = IIf(CellVal(6) = "Sale", "S", "P")
            End With
            cnItemMaster.Execute "INSERT INTO BookChild07 VALUES ('" & rstItemMaster.Fields("Code").Value & "','" & rstItemMaster.Fields("Code").Value & "','" & CellVal(1) & "','" & CellVal(2) & "'," & Val(CellVal(5)) & ",'" & CellVal(7) & "'," & IIf(CheckEmpty(CellVal(3), False), "Null", "'" & CellVal(3) & "'") & ",'" & CellVal(4) & "'," & Val(CellVal(8)) & ",'" & CellVal(6) & "')"
        ElseIf SSTab1.Tab = 7 Then 'Binding & Finishing
            With fpSpread6
                .GetText 10, .ActiveRow, CellVal(1) 'Binding Type
                .GetText 11, .ActiveRow, CellVal(2) 'Bindery Process
                .GetText 12, .ActiveRow, CellVal(3) 'Size
                .GetText 13, .ActiveRow, CellVal(4) 'Calc Mode
                .GetText 3, .ActiveRow, CellVal(5) 'Forms
                .GetText 6, .ActiveRow, CellVal(6) 'Fraction
                .GetText 9, .ActiveRow, CellVal(7) 'Type
                CellVal(7) = IIf(CellVal(7) = "Sale", "S", "P")
                .GetText 4, .ActiveRow, CellVal(8) 'Discribe/Relate
                .GetText 8, .ActiveRow, CellVal(9) 'Discribe/Relate
            End With
            cnItemMaster.Execute "INSERT INTO BookChild08 VALUES ('" & rstItemMaster.Fields("Code").Value & "','" & rstItemMaster.Fields("Code").Value & "','" & CellVal(1) & "','" & CellVal(2) & "','" & CellVal(5) & "','" & CellVal(8) & "','" & CellVal(3) & "','" & CellVal(4) & "'," & Val(CellVal(9)) & ",'" & ItemType & CellVal(7) & "')" '
        End If
    End If
    Exit Function
ErrorHandler:
        UpdateItemList = False
End Function
Private Sub AddToList()
    On Error Resume Next
    With rstItemList
        .MoveFirst
        .Find "[Code] = '" & rstItemMaster.Fields("Code").Value & "'"
        If .EOF Then
        .AddNew
        .Fields("Code").Value = rstItemMaster.Fields("Code").Value
        .Fields("Name").Value = rstItemMaster.Fields("Name").Value
        .Fields("Alias").Value = rstItemMaster.Fields("BusyCode").Value
        .Fields("ItemGroup").Value = Text8.Text
        .Update
        .Sort = SortCol & " " & SortOrder
        .Find "[Code] = '" & rstItemMaster.Fields("Code").Value & "'"
        End If
    End With
End Sub
Private Function ValidateForm() As Boolean
    If CheckEmpty(Text2.Text, False) Then 'Name
        SSTab1.Tab = 1: Text2.SetFocus: ValidateForm = True
    ElseIf CheckDuplicate(cnItemMaster, "BookMaster", "Code", "Name", Trim(Text2.Text), rstItemMaster.Fields("Code").Value, False) Then
        SSTab1.Tab = 1: Text2.SetFocus: ValidateForm = True
    ElseIf CheckEmpty(Text3.Text, False) Then 'Print Name
       SSTab1.Tab = 1: Text3.SetFocus: ValidateForm = True
    ElseIf CheckEmpty(Text5.Text, False) Then 'Finish Size
        If Trim(ReadFromFile("Customer Type")) <> "General" Then SSTab1.Tab = 1: Text5.SetFocus: ValidateForm = True
    ElseIf CheckEmpty(Text8.Text, False) Then 'Item Group
        SSTab1.Tab = 1: Text8.SetFocus: ValidateForm = True
    ElseIf CheckEmpty(Text7.Text, False) Then 'HSN Code
        SSTab1.Tab = 1: Text7.SetFocus: ValidateForm = True
    ElseIf CheckEmpty(Text15.Text, False) Then 'Intigration Name
        SSTab1.Tab = 1: Text15.SetFocus: ValidateForm = True
    ElseIf CheckEmpty(Text16.Text, False) Then 'Intigration Unit
        SSTab1.Tab = 1: Text16.SetFocus: ValidateForm = True
    ElseIf chkItem(4) Then
        fpSpread4.SetFocus: ValidateForm = True
    ElseIf chkItem(5) Then
        fpSpread2.SetFocus: ValidateForm = True
    End If
End Function
Private Function chkItem(ByVal TabNo As Integer) As Boolean
    Dim CellVal(1 To 2) As Variant, i As Integer
    chkItem = False
    SSTab1.Tab = TabNo
    If SSTab1.Tab = 4 Then
        With fpSpread4
            For i = 1 To .DataRowCnt
                .SetActiveCell 1, i
                .GetText 18, i, CellVal(1) 'Element
                .GetText 17, i, CellVal(2) 'Type
                If chkForms Then chkItem = True: Exit For
                If Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False)) Then If CheckDuplicateItem(i, CellVal(1) + CellVal(2)) Then chkItem = True: Exit For
            Next
        End With
    ElseIf SSTab1.Tab = 5 Then
        With fpSpread2
            For i = 1 To .DataRowCnt
                .SetActiveCell 1, i
                .GetText 14, i, CellVal(1) 'Element
                .GetText 13, i, CellVal(2) 'Type
                If Not (CheckEmpty(CellVal(1), False) Or CheckEmpty(CellVal(2), False)) Then If CheckDuplicateItem(i, CellVal(1) + CellVal(2)) Then chkItem = True: Exit For
            Next
        End With
    End If
End Function
Private Function CheckDuplicateItem(ByVal CurRow As Double, ByVal UniqueField As String) As Boolean
    Dim CellVal(1 To 2) As Variant, i As Integer
    If SSTab1.Tab = 4 Then
        With fpSpread4
            For i = 1 To .DataRowCnt
                .SetActiveCell 1, i
                .GetText 18, i, CellVal(1) 'Element
                .GetText 17, i, CellVal(2) 'Type
                If CellVal(1) + CellVal(2) = UniqueField And i <> CurRow Then CheckDuplicateItem = True: Call DisplayError("Duplicate Item in Row #" + Trim(i)): Exit For
            Next
        End With
    ElseIf SSTab1.Tab = 5 Then
        With fpSpread2
            For i = 1 To .DataRowCnt
                .SetActiveCell 1, i
                .GetText 14, i, CellVal(1) 'Element
                .GetText 13, i, CellVal(2) 'Type
                If CellVal(1) + CellVal(2) = UniqueField And i <> CurRow Then CheckDuplicateItem = True: Call DisplayError("Duplicate Item in Row #" + Trim(i)): Exit For
            Next
        End With
    End If
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then Timer1.Enabled = False: ShowProgressInStatusBar False
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstItemList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
Private Sub DuplicateRecord()
    Dim Tbl As String, ItemCode As String, ItemName As String
    Tbl = GetFileNameFromPath(GetTemporaryFileName()): Tbl = Left(Tbl, InStr(1, Tbl, ".", vbTextCompare) - 1)
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    ItemCode = GenerateCode(cnItemMaster, "SELECT MAX(Code) FROM BookMaster", 6, "0"): ItemName = Trim(Left(rstItemList.Fields("Name").Value, 76)) + " (D)"
    With cnItemMaster
        .BeginTrans
        .Execute "SELECT * INTO " & Tbl & " FROM BookMaster WHERE Code = '" & rstItemList.Fields("Code").Value & "'"
        .Execute "UPDATE  " & Tbl & " SET Code='" & ItemCode & "',Name='" & ItemName & "',PrintName='" & ItemName & "'"
        .Execute "INSERT INTO BookMaster SELECT * FROM " & Tbl
        .Execute "DROP TABLE " & Tbl
        .Execute "SELECT * INTO " & Tbl & " FROM BookChild01 WHERE Code='" & rstItemList.Fields("Code").Value & "'"
        .Execute "UPDATE  " & Tbl & " SET Code='" & ItemCode & "'"
        .Execute "INSERT INTO BookChild01 SELECT * FROM " & Tbl
        .Execute "DROP TABLE " & Tbl
        .Execute "SELECT * INTO " & Tbl & " FROM BookChild05 WHERE Code='" & rstItemList.Fields("Code").Value & "'"
        .Execute "UPDATE  " & Tbl & " SET Code='" & ItemCode & "'"
        .Execute "INSERT INTO BookChild05 SELECT * FROM " & Tbl
        .Execute "DROP TABLE " & Tbl
        .Execute "SELECT * INTO " & Tbl & " FROM BookChild06 WHERE Code='" & rstItemList.Fields("Code").Value & "'"
        .Execute "UPDATE  " & Tbl & " SET Code='" & ItemCode & "'"
        .Execute "INSERT INTO BookChild06 SELECT * FROM " & Tbl
        .Execute "DROP TABLE " & Tbl
        .Execute "SELECT * INTO " & Tbl & " FROM BookChild07 WHERE Code='" & rstItemList.Fields("Code").Value & "'"
        .Execute "UPDATE  " & Tbl & " SET Code='" & ItemCode & "'"
        .Execute "INSERT INTO BookChild07 SELECT * FROM " & Tbl
        .Execute "DROP TABLE " & Tbl
        .Execute "SELECT * INTO " & Tbl & " FROM BookChild08 WHERE Code='" & rstItemList.Fields("Code").Value & "'"
        .Execute "UPDATE  " & Tbl & " SET Code='" & ItemCode & "'"
        .Execute "INSERT INTO BookChild08 SELECT * FROM " & Tbl
        .Execute "DROP TABLE " & Tbl
        .CommitTrans
    End With
    Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    Text1.Text = Trim(ItemName): Sendkeys "{END}"
    MdiMainMenu.MousePointer = vbNormal
    Call MsgBox("Successfully duplicated the record !", vbInformation, App.Title)
    Exit Sub
ErrorHandler:
    MdiMainMenu.MousePointer = vbNormal
    DisplayError ("Failed to duplicate the record")
    cnItemMaster.RollbackTrans
End Sub
Private Function CalculateForms(ByVal FormType As String, ByVal Row As Long) As Integer
    Dim TotalForms As Variant, Forms As Variant
    With fpSpread4
        .GetText 9, Row, TotalForms
        If InStr(1, "Q_H", FormType) > 0 Then '¼ or ½ Forms
            TotalForms = Val(TotalForms) - Int(TotalForms)
            If Val(TotalForms) > 0 Then TotalForms = IIf(FormType = "Q", IIf(InStr(1, "0.25_0.75_0.375_0.875", TotalForms) > 0, 1, 0), IIf(InStr(1, "0.5_0.75_0.625_0.875", TotalForms) > 0 Or TotalForms = (5 / 6), 1, 0))
            CalculateForms = Val(TotalForms)
        ElseIf InStr(1, "F_W", FormType) > 0 Then '1 Forms-F&B or W&T
            TotalForms = IIf(FormType = "F", Int(TotalForms / 2) * 2, Int(TotalForms) - Int(TotalForms / 2) * 2)
            CalculateForms = Val(TotalForms)
        End If
    End With
End Function
Private Function chkForms() As Boolean
    Dim i As Integer, Forms As Variant, TotalForms As Double
    chkForms = False
    With fpSpread4
        For i = 1 To .DataRowCnt
            .SetActiveCell 1, i: TotalForms = 0
            .GetText 10, i, Forms '¼ Forms
            TotalForms = TotalForms + Val(Forms) * 0.25
            .GetText 11, i, Forms '½ Forms
            TotalForms = TotalForms + Val(Forms) * 0.5
            .GetText 12, i, Forms '1F-F&B
            TotalForms = TotalForms + Val(Forms)
            .GetText 13, i, Forms '1F-W&T
            TotalForms = TotalForms + Val(Forms)
            .GetText 9, i, Forms 'Forms
            If Val(Forms) <> TotalForms Then Call DisplayError("Printing forms mismatch in Row #" + Trim(i)): chkForms = True: Exit For
        Next
    End With
End Function
Private Sub CalculateTotalForms()
    Dim i As Integer, CurVal As Variant
    MhRealInput15.Value = 0: MhRealInput16.Value = 0: MhRealInput7.Value = 0: MhRealInput4.Value = 0
    With fpSpread4 'Multi Form
        For i = 1 To .DataRowCnt
            .GetText 8, i, CurVal 'Pages
            MhRealInput15.Value = MhRealInput15.Value + Val(CurVal)
            .GetText 9, i, CurVal 'Forms
            MhRealInput7.Value = MhRealInput7.Value + Val(CurVal)
            .GetText 15, i, CurVal 'Binding Forms
            MhRealInput4.Value = MhRealInput4.Value + Val(CurVal)
        Next
    End With
    With fpSpread2 'Multi Element
        For i = 1 To .DataRowCnt
            .GetText 2, i, CurVal 'Pages
            MhRealInput16.Value = MhRealInput16.Value + Val(CurVal)
            .GetText 12, i, CurVal 'Binding Forms
            MhRealInput4.Value = MhRealInput4.Value + Val(CurVal)
        Next
    End With
End Sub
Private Function MaxUps(ByVal FT As String) As Integer
    Dim FL As Double, FR As Double, PL As Double, PR As Double, Ups01 As Integer, Ups02 As Integer
    If FT = "E" Then 'Multi Element Format
        With fpSpread2
            .GetText 15, .ActiveRow, FinishSizeCode2: .GetText 16, .ActiveRow, PrintingSizeCode
            If CheckEmpty(FinishSizeCode2, False) Or CheckEmpty(PrintingSizeCode, False) Then MaxUps = 0: Exit Function
            .GetText 3, .ActiveRow, FinishSizeCode2: .GetText 4, .ActiveRow, PrintingSizeCode
        End With
    ElseIf FT = "F" Then 'Multi Form Format
        With fpSpread4
            .GetText 19, .ActiveRow, FinishSizeCode2: .GetText 20, .ActiveRow, PrintingSizeCode
            If CheckEmpty(FinishSizeCode2, False) Or CheckEmpty(PrintingSizeCode, False) Then MaxUps = 0: Exit Function
            .GetText 2, .ActiveRow, FinishSizeCode2: .GetText 3, .ActiveRow, PrintingSizeCode
        End With
    End If
    PL = Val(Left(PrintingSizeCode, InStr(1, PrintingSizeCode, "X") - 1)) + 1: PR = Val(Mid(PrintingSizeCode, InStr(1, PrintingSizeCode, "X") + 1, 5)) + 1: FL = Val(Left(FinishSizeCode2, InStr(1, FinishSizeCode2, "X") - 1)): FR = Val(Mid(FinishSizeCode2, InStr(1, FinishSizeCode2, "X") + 1, 5))
    Ups01 = Int(IIf(PL > PR, PL, PR) / IIf(FL > FR, FL, FR)) * Int(IIf(PL < PR, PL, PR) / IIf(FL < FR, FL, FR)): Ups02 = Int(IIf(PL < PR, PL, PR) / IIf(FL > FR, FL, FR)) * Int(IIf(PL > PR, PL, PR) / IIf(FL < FR, FL, FR))
    MaxUps = IIf(Ups01 > Ups02, Ups01, Ups02)
End Function
Private Sub LoadMasterList()
    If rstHSNCodeList.State = adStateOpen Then rstHSNCodeList.Close
    rstHSNCodeList.Open "SELECT Name As Col0, Code FROM GeneralMaster WHERE Type= '18' ORDER BY Name", cnItemMaster, adOpenKeyset, adLockReadOnly
    rstHSNCodeList.ActiveConnection = Nothing
    If rstItemGroupList.State = adStateOpen Then rstItemGroupList.Close
    rstItemGroupList.Open "SELECT IIF(Name<>PrintName,Name+'  '+PrintName,Name) As Col0, Code FROM GeneralMaster WHERE Type = '5' ORDER BY Name", cnItemMaster, adOpenKeyset, adLockReadOnly
    rstItemGroupList.ActiveConnection = Nothing
    If rstUnitList.State = adStateOpen Then rstUnitList.Close
    rstUnitList.Open "SELECT Name As Col0, Code FROM GeneralMaster WHERE Type = '25' ORDER BY Name", cnItemMaster, adOpenKeyset, adLockReadOnly
    rstUnitList.ActiveConnection = Nothing
    If rstFinishSizeList.State = adStateOpen Then rstFinishSizeList.Close
    rstFinishSizeList.Open "SELECT Name As Col0, Code FROM GeneralMaster WHERE Type = '11' ORDER BY Name", cnItemMaster, adOpenKeyset, adLockReadOnly
    rstFinishSizeList.ActiveConnection = Nothing
    If rstBindingTypeList.State = adStateOpen Then rstBindingTypeList.Close
    rstBindingTypeList.Open "SELECT Name As Col0, Code FROM GeneralMaster WHERE Type = '6' ORDER BY Name", cnItemMaster, adOpenKeyset, adLockReadOnly
    rstBindingTypeList.ActiveConnection = Nothing
    If rstUnfinishedItemList.State = adStateOpen Then rstUnfinishedItemList.Close
    rstUnfinishedItemList.Open "SELECT Name As Col0,Code FROM BookMaster WHERE Type='R' ORDER BY Name", cnItemMaster, adOpenKeyset, adLockReadOnly
    rstUnfinishedItemList.ActiveConnection = Nothing
    If rstElementList.State = adStateOpen Then rstElementList.Close
    rstElementList.Open "SELECT Name As Col0,Pages,Code FROM ElementMaster ORDER BY Name", cnItemMaster, adOpenKeyset, adLockReadOnly
    rstElementList.ActiveConnection = Nothing
    If rstSizeList.State = adStateOpen Then rstSizeList.Close
    rstSizeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type IN ('1','11') ORDER BY Name", cnItemMaster, adOpenKeyset, adLockReadOnly
    rstSizeList.ActiveConnection = Nothing
    If rstPicList.State = adStateOpen Then rstPicList.Close
    rstPicList.Open "SELECT * FROM ItemPicChild WHERE Item='" & FixQuote(rstItemList.Fields("Code").Value) & "'", cnItemMaster, adOpenKeyset, adLockReadOnly
    rstPicList.ActiveConnection = Nothing
    If rstPicMaster.State = adStateOpen Then rstPicMaster.Close
    rstPicMaster.Open "SELECT * FROM ItemPicChild WHERE Item='" & FixQuote(rstItemList.Fields("Code").Value) & "'", cnItemMaster, adOpenKeyset, adLockOptimistic
    rstPicList.ActiveConnection = Nothing
    If rstMFList.State = adStateOpen Then rstMFList.Close
    rstMFList.Open "SELECT IIF(Name<>PrintName,Name+' '+PrintName,PrintName) As Col0,Code FROM AccountMaster ORDER BY Name", cnItemMaster, adOpenKeyset, adLockOptimistic
    rstMFList.ActiveConnection = Nothing
    End Sub
Private Sub cmdLoadElement_Click()
    If fpSpread6.DataRowCnt > 0 Then Exit Sub
    Dim i As Integer, n As Integer, CurVal(1 To 6) As Variant
    If rstBdgList.State = adStateOpen Then rstBdgList.Close
    rstBdgList.Open "SELECT B.Code As BCode,B.Name As BName,O.Code As OCode,O.Name As OName,V.Code As CalcMode,V.Name As CalcModeName,V.Value1 As CalcValue FROM GeneralMaster B  Left JOIN BindingTypeChild B1 ON B.Code=B1.Code Left JOIN GeneralMaster O ON BinderyProcess=O.Code Left JOIN GeneralMaster V ON O.UnderGroup=V.Code WHERE B.Code='" & BindingTypeCode & "' ORDER BY O.Name", cnItemMaster, adOpenKeyset, adLockReadOnly
    rstBdgList.ActiveConnection = Nothing
    rstBdgList.MoveFirst
    Do While Not rstBdgList.EOF
    With fpSpread6 'Binding & Finishing
            CurVal(1) = rstBdgList.Fields("BName").Value: CurVal(2) = rstBdgList.Fields("BCode").Value 'BindingType
            CurVal(3) = MhRealInput4 'Binding Forms
            CurVal(4) = FinishSizeCode1: CurVal(5) = Text5 'Printing Size
            CurVal(6) = 1 'Ups
            If Val(CurVal(3)) > 0 Then
                    n = n + 1
                    If n = 1 Then .SetText 1, n, CurVal(1)
                    .SetText 10, n, CurVal(2) 'BindingType
                    .SetText 2, n, rstBdgList.Fields("OName").Value: .SetText 11, n, rstBdgList.Fields("OCode").Value 'Bindiry Process
                    .SetText 3, n, CurVal(3) 'Binding Forms
                    .SetText 4, n, "Nos"
                    .SetText 5, n, CurVal(5): .SetText 12, n, CurVal(4) 'Printing Size
                    .SetText 6, n, CurVal(6) 'Ups'
                    .SetText 7, n, rstBdgList.Fields("CalcModeName").Value
                    If Not IsNull(rstBdgList.Fields("CalcValue").Value) Then .SetText 8, n, Val(rstBdgList.Fields("CalcValue").Value)
                    .SetText 13, n, rstBdgList.Fields("CalcMode").Value
                    If Trim(ReadFromFile("Client ID")) = "Publisher" Then CurVal(6) = "Purchase" Else CurVal(6) = "Sale"
                    .SetText 9, n, CurVal(6)
            End If
    End With
        rstBdgList.MoveNext
    Loop

'    If fpSpread6.DataRowCnt > 0 Then Exit Sub
'    Dim i As Integer, n As Integer, CurVal(1 To 6) As Variant
'    With fpSpread4 'Multi Form Format
'        For i = 1 To .DataRowCnt
'            .GetText 1, i, CurVal(1): .GetText 18, i, CurVal(2) 'Element
'            .GetText 15, i, CurVal(3) 'Binding Forms
'            .GetText 3, i, CurVal(4): .GetText 20, i, CurVal(5) 'Printing Size
'            .GetText 16, i, CurVal(6) 'Ups
'            If Val(CurVal(3)) > 0 Then
'                With fpSpread6 'Binding Element
'                    n = n + 1
'                    .SetText 1, n, CurVal(1): .SetText 8, n, CurVal(2) 'Element
'                    .SetText 3, n, CurVal(3) 'Binding Forms
'                    .SetText 4, n, CurVal(4): .SetText 10, n, CurVal(5) 'Printing Size
'                    .SetText 5, n, CurVal(6) 'Ups
'                End With
'            End If
'        Next
'    End With
'    With fpSpread2 'Multi Element Format
'        For i = 1 To .DataRowCnt
'            .GetText 1, i, CurVal(1): .GetText 14, i, CurVal(2) 'Element
'            .GetText 12, i, CurVal(3) 'Binding Forms
'            .GetText 4, i, CurVal(4): .GetText 16, i, CurVal(5) 'Printing Size
'            .GetText 10, i, CurVal(6) 'Ups
'            If Val(CurVal(3)) > 0 Then
'                With fpSpread6 'Binding Element
'                    n = n + 1
'                    .SetText 1, n, CurVal(1): .SetText 8, n, CurVal(2) 'Element
'                    .SetText 3, n, CurVal(3) 'Binding Forms
'                    .SetText 4, n, CurVal(4): .SetText 10, n, CurVal(5) 'Printing Size
'                    .SetText 5, n, CurVal(6) 'Ups
'                End With
'            End If
'        Next
'    End With
End Sub
Private Sub CmdPrint_Click()
PrintFlag = True
On Error GoTo errHandler_print
PrintItemList ("S")
'cmdexport_click
    On Error GoTo 0
Exit Sub
errHandler_print:
  On Error GoTo 0
  Exit Sub
End Sub
Private Sub cmdexport_click()
Screen.MousePointer = vbHourglass
Dim oExcel As Object
Dim oPdf As Object
Dim oBook As Object
Dim oSheet As Object
Dim j As Integer, i As Integer, Cnt As Long
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   Set oSheet = oBook.Worksheets(1)
   On Error GoTo errcode
   With oBook.Worksheets("sheet1").Rows(1)
        .Font.Bold = True
        .Font.Size = 16
        oBook.Worksheets("sheet1").Cells(1, j + 1).Value = "Item List"
        .Range("A1:I1").Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        oBook.Worksheets("sheet1").Rows(2).Font.Bold = True
        For j = 0 To DataGrid1.Columns.Count - 1
            oBook.Worksheets("sheet1").Cells(2, j + 1).Value = DataGrid1.Columns(j).Caption
           MdiMainMenu.StatusBar1.Panels(2).Text = "Processed record #" & Trim(Str(Cnt)) & " of " & Trim(Str(rstItemList.RecordCount)) & " !!!"
        Cnt = Cnt + 1
        Next j
    'Next i
   End With
 'Save Data
   oSheet.Range("A3").CopyFromRecordset rstItemList
     
   With oExcel
            StartColumn = "A"
            StartRow = 1
            EndColumn = "I"
            EndRow = rstItemList.RecordCount + 1
            oBook.Activate
            oExcel.Visible = True
'            oBook.Worksheets("sheet1").Cells(2, 1).Value = " Name"
'            oBook.Worksheets("sheet1").Cells(2, 2).Value = " Print Name"
'            oBook.Worksheets("sheet1").Cells(2, 3).Value = " Alias"
'            oBook.Worksheets("sheet1").Cells(2, 4).Value = " ISBN"
'            oBook.Worksheets("sheet1").Cells(2, 5).Value = " Group"
'            oBook.Worksheets("sheet1").Cells(2, 6).Value = " Code"
'            oBook.Worksheets("sheet1").Cells(2, 7).Value = "  Price"
'            oBook.Worksheets("sheet1").Cells(2, 8).Value = "  Sales Price"
'            oBook.Worksheets("sheet1").Cells(2, 9).Value = " Purchase Price"
            .Columns("A:Z").EntireColumn.AutoFit
            .ActiveSheet.pagesetup.Orientation = xlLandscape
            .ActiveSheet.pagesetup.LeftMargin = .InchesToPoints(0.36)
            .ActiveSheet.pagesetup.RightMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.TopMargin = .InchesToPoints(0.5)
            .ActiveSheet.pagesetup.BottomMargin = .InchesToPoints(0.5)
            .ActiveSheet.pagesetup.HeaderMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.FooterMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.PrintArea = StartColumn & StartRow & ":" & EndColumn & EndRow + 1
            .ActiveSheet.pagesetup.Zoom = False
            .ActiveSheet.pagesetup.FitToPagesTall = False
            .ActiveSheet.pagesetup.FitToPagesWide = 1
            .ActiveSheet.pagesetup.PrintGridlines = True
            .ActiveSheet.pagesetup.PrintTitleRows = "$1:$2"
            
            If PrintFlag Then
            .ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF
                With CommonDialog1
                    .Copies = 1
                    .flags = &H0&
                    .ShowPrinter
                    oExcel.ActiveSheet.PrintOut
                End With
            End If
   End With
   PrintFlag = False

   Set oBook = Nothing
   Set oSheet = Nothing
   Set oExcel = Nothing
   Screen.MousePointer = vbDefault
   Exit Sub
errcode:
   MsgBox Err.Description, , Err.Source
   PrintFlag = False
   Set oBook = Nothing
   Set oSheet = Nothing
   Set oExcel = Nothing
   Screen.MousePointer = vbDefault
End Sub
Private Sub CmdLabel_Click()
On Error Resume Next
    PrintFlag = True
    Call PrintLabel
End Sub
Public Sub PrintLabel()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Dim oExcel As Object
    Dim oPdf As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim j, R As Integer, i As Integer, Cnt As Long
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    
    If rstItemMaster.State = adStateOpen Then rstItemMaster.Close
    rstItemMaster.Open "SELECT * ,(Select Name From GeneralMaster Where Type=5 And Code=[Group]) As GroupName FROM BookMaster WHERE Code='" & FixQuote(rstItemList.Fields("Code").Value) & "'", cnItemMaster, adOpenKeyset, adLockReadOnly
    If rstItemMaster.RecordCount = 0 Then On Error GoTo 0: Screen.MousePointer = vbNormal: Exit Sub
    rstItemMaster.ActiveConnection = Nothing
    
   On Error GoTo errcode
        For j = 1 To 60
      With oBook.Worksheets("sheet1").Rows(j)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 16
    End With
        Next j
   With oExcel
            oBook.Activate
            oExcel.Visible = True
For j = 1 To 2
    For i = 0 To 49
            i = i + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Item Description :"
            If rstItemMaster.Fields("Name").Value <> "" Then i = i + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = " Name: " + rstItemMaster.Fields("Name").Value
            If rstItemMaster.Fields("PrintName").Value <> "" Then i = i + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "  Description: " + rstItemMaster.Fields("PrintName").Value
            If rstItemMaster.Fields("GroupName").Value <> "" Then i = i + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = " Group: " + rstItemMaster.Fields("GroupName").Value
            If rstItemMaster.Fields("ISBN").Value <> "" Then i = i + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = " ISBN/EAN: " + rstItemMaster.Fields("ISBN").Value
            If rstItemMaster.Fields("BusyCode").Value <> "" Then i = i + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = " Alias: " + rstItemMaster.Fields("BusyCode").Value
            If rstItemMaster.Fields("Code").Value <> "" Then i = i + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = "Code: " + rstItemMaster.Fields("Code").Value
            If rstItemMaster.Fields("Price").Value <> "" Then i = i + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = " Price:  " & rstItemMaster.Fields("Price").Value
            If rstItemMaster.Fields("SalePrice").Value <> "" Then i = i + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = " Sale Price: " & rstItemMaster.Fields("SalePrice").Value
            If rstItemMaster.Fields("PurPrice").Value <> "" Then i = i + 1: oBook.Worksheets("sheet1").Cells(i, j).Value = " Purchase Price: " & rstItemMaster.Fields("PurPrice").Value
            
    Next i
Next j
            .Columns("A:B").EntireColumn.AutoFit
            .ActiveSheet.pagesetup.Orientation = xlPortrait
            .ActiveSheet.pagesetup.LeftMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.RightMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.TopMargin = .InchesToPoints(0)
            .ActiveSheet.pagesetup.BottomMargin = .InchesToPoints(0)
            .ActiveSheet.pagesetup.HeaderMargin = .InchesToPoints(0)
            .ActiveSheet.pagesetup.FooterMargin = .InchesToPoints(0)
            StartColumn = "A"
            StartRow = 1
            EndColumn = "B"
            EndRow = i
            .ActiveSheet.pagesetup.PrintArea = StartColumn & StartRow & ":" & EndColumn & EndRow - 1
            .ActiveSheet.pagesetup.Zoom = False
            .ActiveSheet.pagesetup.FitToPagesTall = 1
            .ActiveSheet.pagesetup.FitToPagesWide = 1
            .ActiveSheet.pagesetup.PrintGridlines = False
            .ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF
            If PrintFlag Then
                With CommonDialog1
                    .Copies = 1
                    .flags = &H0&
                    .ShowPrinter
                    oExcel.ActiveSheet.PrintOut
                End With
            End If
   End With
   PrintFlag = False
   Set oBook = Nothing
   Set oSheet = Nothing
   Set oExcel = Nothing
   Screen.MousePointer = vbDefault
   Exit Sub
errcode:
   MsgBox Err.Description, , Err.Source
   PrintFlag = False
   Set oBook = Nothing
   Set oSheet = Nothing
   Set oExcel = Nothing
   Screen.MousePointer = vbDefault
    Call CloseRecordset(rstItemMaster)
    On Error GoTo 0
End Sub
Public Sub PrintItemList(ByVal OutputType As String)
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Load FrmDialog: Screen.MousePointer = vbNormal:  FrmDialog.Frame1.Caption = "Select Format": FrmDialog.Flag = 8: FrmDialog.Caption = "Item List": FrmDialog.Command1.Caption = "Sale Price List": FrmDialog.Command2.Caption = "Sale & Pur. Item List": FrmDialog.Command3.Caption = "MFD-Wise Item List": FrmDialog.Command4.Caption = "Item List": FrmDialog.Command5.Visible = False: FrmDialog.Show vbModal
    Screen.MousePointer = vbHourglass
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT * FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    rstCompanyMaster.ActiveConnection = Nothing
        If rstCompanyMaster.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    If orstItemList.State = adStateOpen Then orstItemList.Close
If PtgType = "4" Then 'Export
    orstItemList.Open "SELECT ROW_NUMBER() OVER (ORDER BY P.Name) AS RowNo,P.Name,P.PrintName,C.PrintName As ItemGroup,BusyCode As Alias,P.Code,ISBN,(Select PrintName From GeneralMaster Where Type=25 AND  Code=IntegrationUnit) As Unit,ISNULL((Select Name FROM AccountMaster Where Code=P.MFCD),'') AS MFDC,Format(P.Price,'##00.00') As Price,Format(P.salePrice,'##00.00') As SalePrice,Format(P.PurPrice,'##00.00') AS PurPrice,C.Name As GroupCode,C.PrintName As GroupName,P.Remarks,P.s_no,P.ItemMarks AS Marks,Format(P.PurDate,'dd-MM-yy') AS PurDate,ISNULL((Select PrintName FROM AccountMaster Where Code=P.MFCD),'') AS MFDCName,(Select SUM(OpBal) From BookChild Where P.Code=Item) AS OP,ISNULL((Select SUM(ABS(Quantity)) From JobworkBVChild Where P.Code=Item And Left(BOM,2)='01'),0) AS PUR,ISNULL((Select SUM(ABS(Quantity)) From JobworkBVChild Where P.Code=Item And Left(BOM,2)='04'),0) AS Sale, " & _
                                  "(Select SUM(OpBal) From BookChild Where P.Code=Item)+ISNULL((Select SUM(ABS(Quantity)) From JobworkBVChild Where P.Code=Item And Left(BOM,2)='01'),0)-ISNULL((Select SUM(ABS(Quantity)) From JobworkBVChild Where P.Code=Item And Left(BOM,2)='04'),0) AS Stock  FROM BookMaster P INNER JOIN GeneralMaster C ON P.[Group]=C.Code WHERE P.Type='F' AND P.ItemMarks<>'' ORDER BY P.Name ", cnDatabase, adOpenKeyset, adLockReadOnly
Else
    orstItemList.Open "SELECT ROW_NUMBER() OVER (ORDER BY P.Name) AS RowNo,P.Name,P.PrintName,C.PrintName As ItemGroup,BusyCode As Alias,P.Code,ISBN,(Select PrintName From GeneralMaster Where Type=25 AND  Code=IntegrationUnit) As Unit,ISNULL((Select Name FROM AccountMaster Where Code=P.MFCD),'') AS MFDC,Format(P.Price,'##00.00') As Price,Format(P.salePrice,'##00.00') As SalePrice,Format(P.PurPrice,'##00.00') AS PurPrice,C.Name As GroupCode,C.PrintName As GroupName,P.Remarks,P.s_no,P.ItemMarks AS Marks,Format(P.PurDate,'dd-MM-yy') AS PurDate,ISNULL((Select PrintName FROM AccountMaster Where Code=P.MFCD),'') AS MFDCName,(Select SUM(OpBal) From BookChild Where P.Code=Item) AS OP,ISNULL((Select SUM(ABS(Quantity)) From JobworkBVChild Where P.Code=Item And Left(BOM,2)='01'),0) AS PUR,ISNULL((Select SUM(ABS(Quantity)) From JobworkBVChild Where P.Code=Item And Left(BOM,2)='04'),0) AS Sale, " & _
                                  "(Select SUM(OpBal) From BookChild Where P.Code=Item)+ISNULL((Select SUM(ABS(Quantity)) From JobworkBVChild Where P.Code=Item And Left(BOM,2)='01'),0)-ISNULL((Select SUM(ABS(Quantity)) From JobworkBVChild Where P.Code=Item And Left(BOM,2)='04'),0) AS Stock  FROM BookMaster P INNER JOIN GeneralMaster C ON P.[Group]=C.Code WHERE P.Type='F' ORDER BY P.Name ", cnDatabase, adOpenKeyset, adLockReadOnly
End If
    orstItemList.ActiveConnection = Nothing
        If orstItemList.RecordCount = 0 Then On Error GoTo 0: Exit Sub
        orstItemList.ActiveConnection = Nothing
        orstItemList.MoveFirst
        rptItemList.Database.SetDataSource orstItemList, 3, 1
        rptItemList.DiscardSavedData
        Screen.MousePointer = vbNormal
    With rptItemList
    'Section Suppress
    If PtgType = "1" Then '
        .Section1.Suppress = False
        .Section5.Suppress = False
        .Section8.Suppress = False
    ElseIf PtgType = "2" Then
        .Section6.Suppress = False
        .Section7.Suppress = False
        .Section8.Suppress = False
    ElseIf PtgType = "3" Or PtgType = "4" Then
        .Section13.Suppress = False
        .Section10.Suppress = False
        .Section16.Suppress = False
    End If
    'Text_Print Date & Time
            '.Text.SetText "Printed on " & Format(Now, "dd-MMM-yyyy") & " at " & Format(Now, "hh:mm"): .Text.Font.Size = 7: .Text.Font.Bold = False
    'Text12_Company Name
            .Text12.SetText rstCompanyMaster.Fields("PrintName").Value: .Text12.Font.Size = 16: .Text12.Font.Bold = True
    'Text5_Company Address
            .Text5.SetText rstCompanyMaster.Fields("Address1").Value & " " & rstCompanyMaster.Fields("Address2").Value & " " & rstCompanyMaster.Fields("Address3").Value & " " & rstCompanyMaster.Fields("Address4").Value: .Text5.Font.Size = 12: .Text5.Font.Bold = False
    'Text11_Report Header
            If PtgType = "1" Then .Text11.SetText " Sales PRICE LIST ": .Text11.Font.Size = 12: .Text11.Font.Bold = True:
            If PtgType = "2" Then .Text11.SetText " Full PRICE LIST ": .Text11.Font.Size = 12: .Text11.Font.Bold = True:
            If PtgType = "3" Then .Text11.SetText " MFCD-WISE ITEM LIST ": .Text11.Font.Size = 12: .Text11.Font.Bold = True
            If PtgType = "4" Then .Text11.SetText " ITEM LIST [Export] ": .Text11.Font.Size = 12: .Text11.Font.Bold = True:
    End With
        If OutputType = "S" Then
            Screen.MousePointer = vbNormal
            Set FrmReportViewer.Report = rptItemList: FrmReportViewer.Show vbModal
        ElseIf OutputType = "P" Then
            Screen.MousePointer = vbNormal
            rptItemList.PaperSource = crPRBinAuto
            rptItemList.PrintOut
        Else
            Dim oOutlookMsg As Outlook.MailItem, FileName As String
            MsgSubject = IIf(PtgType = 1, "Accounts Summary", "Account Ledger") + " from " + Format(GetDate(Date), "dd-MMM-yyyy") + " to " + Format(GetDate(Date), "dd-MMM-yyyy")
            MsgText = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith " & "Account Ledger from " + Format(GetDate(Date), "dd-MMM-yyyy") + " to " + Format(GetDate(Date), "dd-MMM-yyyy") & " for doing the needful at your end.<Br><b>Kindly do acknowledge the receipt of the mail</b>.<Br><Br>Thanks & Regards<Br>Accounts Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
            ' Export report to PDF
            rptItemList.ExportOptions.FormatType = crEFTPortableDocFormat ' Set the Export Format As .Pdf
            rptItemList.ExportOptions.DestinationType = crEDTDiskFile
            FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
            rptItemList.ExportOptions.DiskFileName = FileName
            rptItemList.Export False
            
            If MsgBox("Do You want's to Send Account Ledger by Outlook email", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbNo Then
                    Send_email rstCompanyMaster.Fields("SmtpServer").Value, rstCompanyMaster.Fields("Port").Value, rstCompanyMaster.Fields("UserName").Value, rstCompanyMaster.Fields("Password").Value, "", MsgSubject, MsgText, "", "Accounts Department", Trim(rstCompanyMaster.Fields("PrintName").Value), Trim(rstCompanyMaster.Fields("Phone").Value), Trim(rstCompanyMaster.Fields("EMail").Value), "", "", FileName
            Else
                    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                    With oOutlookMsg
                        .To = ""
                        .Subject = MsgSubject
                        .HTMLBody = MsgText
                        .Attachments.Add (FileName)
                        .Importance = olImportanceHigh
                        .ReadReceiptRequested = True
                        If CheckEmpty(.To, False) Then .Display Else .Send
                    End With
                    Set oOutlookMsg = Nothing
            End If
        End If
        Set rptItemList = Nothing
        Screen.MousePointer = vbNormal
           'If GridFlag = False Then Call CloseForm(Me) Else GridFlag = False
        Exit Sub
        On Error GoTo 0
        Screen.MousePointer = vbNormal
End Sub
Private Sub First_Click()
If rstPicList.RecordCount > 0 Then rstPicList.MoveFirst Else Exit Sub
        If Not CheckEmpty(rstPicList.Fields("PicData"), False) Then imgFile = App.Path & "\Pic\Item [" & CompCode & "]" & "\" & rstPicList.Fields("PicName").Value & "." & rstPicList.Fields("PicType").Value: RetrievePic rstPicList.Fields("PicData").Value, imgFile, srmPicMgr: Call cmdView_Click
End Sub
Private Sub Last_Click()
If rstPicList.RecordCount > 0 Then rstPicList.MoveLast Else Exit Sub
        If Not CheckEmpty(rstPicList.Fields("PicData"), False) Then imgFile = App.Path & "\Pic\Item [" & CompCode & "]" & "\" & rstPicList.Fields("PicName").Value & "." & rstPicList.Fields("PicType").Value: RetrievePic rstPicList.Fields("PicData").Value, imgFile, srmPicMgr: Call cmdView_Click
End Sub
Private Sub Next_Click()
On Error Resume Next
rstPicList.MoveNext
If Not (rstPicList.EOF) Then
        If Not CheckEmpty(rstPicList.Fields("PicData"), False) Then imgFile = App.Path & "\Pic\Item [" & CompCode & "]" & "\" & rstPicList.Fields("PicName").Value & "." & rstPicList.Fields("PicType").Value: RetrievePic rstPicList.Fields("PicData").Value, imgFile, srmPicMgr: Call cmdView_Click
End If
End Sub
Private Sub Previous_Click()
On Error Resume Next
rstPicList.MovePrevious
If Not (rstPicList.BOF) Then
        If Not CheckEmpty(rstPicList.Fields("PicData"), False) Then imgFile = App.Path & "\Pic\Item [" & CompCode & "]" & "\" & rstPicList.Fields("PicName").Value & "." & rstPicList.Fields("PicType").Value: RetrievePic rstPicList.Fields("PicData").Value, imgFile, srmPicMgr: Call cmdView_Click
End If
End Sub
Private Sub CheckBox1_Change()
    cmdView_Click
End Sub
Private Sub cmdView_Click() 'View Pic
If CheckEmpty(imgFile, False) Then
    Image1.Picture = LoadPicture(imgFile) ': DisplayError ("No image exists")
    Mh3dLabel11 = "No Image"
Else
    Image1.Picture = LoadPicture(imgFile)
    If CheckBox1.Value Then Image1.PictureSizeMode = fmPictureSizeModeZoom
    If Not CheckBox1.Value Then Image1.PictureSizeMode = fmPictureSizeModeClip
    If CheckBox1.Value Then Image1.AutoSize = True
    If Not CheckBox1.Value Then Image1.AutoSize = False
    Image1.PictureAlignment = fmPictureAlignmentCenter
    Image1.SpecialEffect = fmSpecialEffectBump
    Mh3dLabel11 = rstPicList.Fields("PicName").Value + "." & rstPicList.Fields("PicType").Value & "[" & "1" & " /" & rstPicList.RecordCount & "]"
    
End If
End Sub
Private Sub cmdUpload_Click() 'Load Pic
    On Error GoTo ErrorHandler
    With cdUpload
        .CancelError = True
        .DialogTitle = "Open Image"
        .Filter = "All Picture Files|*.jpg;*.jpeg;*.bmp;*.gif;*.png"
        .ShowOpen
        If Err.Number = 0 Then imgFile = .FileName 'Ok Selected
        If CheckEmpty(imgFile, False) Then DisplayError ("No image exists")
            Image1.Picture = LoadPicture(imgFile)
            'Image1.PictureSizeMode = fmPictureSizeModeZoom
            Image1.AutoSize = True
            Image1.PictureAlignment = fmPictureAlignmentCenter
            Image1.SpecialEffect = fmSpecialEffectBump
End With

With rstPicMaster
        If Not CheckEmpty(imgFile, False) Then
                If srmPicMgr.State = adStateOpen Then srmPicMgr.Close
                    srmPicMgr.Type = adTypeBinary
                    srmPicMgr.Open
                        srmPicMgr.LoadFromFile imgFile
        End If
            If srmPicMgr.Size > 0 Then
                    If blnRecordExist And AllowMastersModification = 0 Then
                        Call DisplayError("You don't have the rights to Edit this Master")
                        Exit Sub
                    End If
                    imgFile = cdUpload.FileTitle: Mh3dLabel11 = imgFile
                    imgFile = Left(imgFile, InStr(1, UCase(imgFile), ".") - 1)
                If .RecordCount > 0 Then .MoveFirst
                    .Find "[PicName] = '" & imgFile & "'"
                If Not .EOF Then
                    .Fields("Code").Value = .Fields("Code").Value
                    .Fields("PicName").Value = .Fields("PicName").Value
                Else
                    .AddNew
                    .Fields("Code").Value = GenerateCode(cnItemMaster, "SELECT MAX(Code) FROM ItemPicChild", 6, "0")
                End If
                    .Fields("Item").Value = rstItemMaster.Fields("Code").Value
                    .Fields("PicName").Value = imgFile
                    imgFile = cdUpload.FileName
                    rstPicMaster.Fields("PicType").Value = UCase(FSO.GetExtensionName(FSO.GetFileName(imgFile)))
                    rstPicMaster.Fields("PicData").Value = srmPicMgr.Read
'                rstPicMaster.Update
                    imgFile = cdUpload.FileName
            If UpdateRecord(rstPicMaster) Then
                Call UpdateUserAction("Item Pic Child", IIf(blnRecordExist, "M", "A"), Trim(Date), cnDatabase)
            End If
            If rstPicMaster.State = adStateOpen Then
                rstPicMaster.Close
            End If
            rstPicMaster.CursorLocation = adUseClient
        End If
            Call MsgBox("Item Pic Record updated !!!", vbInformation, App.Title)
    End With
        LoadMasterList
        Call cmdView_Click
Exit Sub
ErrorHandler:
MsgBox ("Failed to save the record" & Err.Description)
Err.Clear
Call CloseRecordset(rstPicMaster)
End Sub
Private Sub cmdDelete_Click() 'Delete Pic
    If Not CheckEmpty(imgFile, False) Then
        If MsgBox("Are you sure to delete the Picture?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            imgFile = ""
            cmdUpload.Enabled = True
            cnItemMaster.Execute "Delete From ItemPicChild where Code='" & rstPicMaster.Fields("Code").Value & "'"
        End If
    End If
End Sub
