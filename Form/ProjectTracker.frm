VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmProjectTracker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Tracker"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18870
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   18870
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9105
      Left            =   15
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   18840
      _Version        =   65536
      _ExtentX        =   33232
      _ExtentY        =   16060
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
      Picture         =   "ProjectTracker.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8895
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   18615
         _ExtentX        =   32835
         _ExtentY        =   15690
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
         TabPicture(0)   =   "ProjectTracker.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Mh3dLabel4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "TDBNumber2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "CmdPrint"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "CmdExport"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Mh3dFrame4"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Mh3dLabel1(2)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Timer2"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "DataGrid1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Text1"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmdRefresh"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "ImageList1"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "CommonDialog1"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Timer3"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Picture1"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "ProjectTracker.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame3"
         Tab(1).Control(1)=   "Mh3dFrame2"
         Tab(1).Control(2)=   "Mh3dFrame5"
         Tab(1).ControlCount=   3
         Begin VB.PictureBox Picture1 
            Height          =   2670
            Left            =   5880
            ScaleHeight     =   2610
            ScaleWidth      =   4965
            TabIndex        =   35
            Top             =   3720
            Visible         =   0   'False
            Width           =   5025
         End
         Begin VB.Timer Timer3 
            Interval        =   60000
            Left            =   4800
            Top             =   3120
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   2040
            Top             =   3120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   6960
            Top             =   2760
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
                  Picture         =   "ProjectTracker.frx":0054
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":0598
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":0ADC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":0BF0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":0D04
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":0E18
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":0F74
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":14B8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":15CC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":1B10
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":1C24
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":1D38
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":1E4C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":1F60
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectTracker.frx":2074
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdRefresh 
            Height          =   325
            Left            =   18280
            Picture         =   "ProjectTracker.frx":2188
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Refresh"
            Top             =   0
            Width           =   325
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
            Left            =   3240
            TabIndex        =   8
            Top             =   8415
            Width           =   10680
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7875
            Left            =   120
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   450
            Width           =   18355
            _ExtentX        =   32385
            _ExtentY        =   13891
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
            ColumnCount     =   12
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "Task"
               Caption         =   "Task"
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
            BeginProperty Column02 
               DataField       =   "Status"
               Caption         =   "Status"
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
               DataField       =   "AssignTo"
               Caption         =   "Task Assigned  To"
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
               DataField       =   "AssignBy"
               Caption         =   "Task Assigned  By"
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
            BeginProperty Column05 
               DataField       =   "ArrivedOn"
               Caption         =   "Assigned Date"
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
               DataField       =   "TargetDate"
               Caption         =   "Target Date"
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
               DataField       =   "Alias"
               Caption         =   "Alias"
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
               DataField       =   "Code"
               Caption         =   "Item Code"
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
               DataField       =   "StartDate"
               Caption         =   "Task Start Date"
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
            BeginProperty Column10 
               DataField       =   "EndDate"
               Caption         =   "Task End Date"
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
            BeginProperty Column11 
               DataField       =   "Remarks"
               Caption         =   "Remarks"
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
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               ScrollBars      =   3
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   5595.024
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2399.811
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   3044.977
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   3044.977
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1395.213
               EndProperty
               BeginProperty Column06 
               EndProperty
               BeginProperty Column07 
               EndProperty
               BeginProperty Column08 
               EndProperty
               BeginProperty Column09 
               EndProperty
               BeginProperty Column10 
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  ColumnWidth     =   1395.213
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame5 
            Height          =   3465
            Left            =   -74880
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   5010
            Width           =   18375
            _Version        =   65536
            _ExtentX        =   32411
            _ExtentY        =   6112
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
            Picture         =   "ProjectTracker.frx":22D2
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   3225
               Left            =   120
               TabIndex        =   11
               Top             =   120
               Width           =   18150
               _Version        =   524288
               _ExtentX        =   32015
               _ExtentY        =   5689
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
               MaxCols         =   3
               MaxRows         =   1000
               SpreadDesigner  =   "ProjectTracker.frx":22EE
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   3465
            Left            =   -74880
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1440
            Width           =   18375
            _Version        =   65536
            _ExtentX        =   32411
            _ExtentY        =   6112
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
            Picture         =   "ProjectTracker.frx":2886
            Begin FPSpreadADO.fpSpread fpSpread3 
               Height          =   3225
               Left            =   120
               TabIndex        =   0
               Top             =   120
               Width           =   18150
               _Version        =   524288
               _ExtentX        =   32015
               _ExtentY        =   5689
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
               MaxCols         =   10
               MaxRows         =   1000
               OperationMode   =   2
               SelectBlockOptions=   0
               SpreadDesigner  =   "ProjectTracker.frx":28A2
            End
            Begin VB.TextBox Text3 
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
               Left            =   2640
               TabIndex        =   13
               Top             =   2280
               Width           =   5775
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame3 
            Height          =   900
            Left            =   -74880
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   450
            Width           =   18375
            _Version        =   65536
            _ExtentX        =   32411
            _ExtentY        =   1587
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
            Picture         =   "ProjectTracker.frx":3169
            Begin VB.CommandButton cmdEnd 
               Enabled         =   0   'False
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
               Left            =   17880
               Picture         =   "ProjectTracker.frx":3185
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "End Work"
               Top             =   255
               Width           =   375
            End
            Begin VB.CommandButton cmdPause 
               Enabled         =   0   'False
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
               Left            =   17520
               Picture         =   "ProjectTracker.frx":34C7
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Pause Work"
               Top             =   255
               Width           =   375
            End
            Begin VB.CommandButton cmdStart 
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
               Left            =   17160
               Picture         =   "ProjectTracker.frx":3809
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Start Work"
               Top             =   255
               Width           =   375
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
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   440
               Width           =   15630
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
               Left            =   6320
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   120
               Width           =   10635
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
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   120
               Width           =   3930
            End
            Begin VB.TextBox Text4 
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
               Left            =   2640
               TabIndex        =   15
               Top             =   2400
               Width           =   5775
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   16
               Top             =   120
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
               Caption         =   " User Code"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ProjectTracker.frx":3B4B
               Picture         =   "ProjectTracker.frx":3B67
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   2
               Left            =   5240
               TabIndex        =   18
               Top             =   120
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
               Caption         =   " User Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ProjectTracker.frx":3B83
               Picture         =   "ProjectTracker.frx":3B9F
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   3
               Left            =   120
               TabIndex        =   20
               Top             =   440
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
               Caption         =   " Project Name"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ProjectTracker.frx":3BBB
               Picture         =   "ProjectTracker.frx":3BD7
            End
            Begin VB.Line Line1 
               X1              =   17060
               X2              =   17060
               Y1              =   0
               Y2              =   975
            End
         End
         Begin VB.Timer Timer2 
            Interval        =   150
            Left            =   8160
            Top             =   4080
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   14010
            TabIndex        =   22
            Top             =   8415
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
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
            Caption         =   " Ctrl+E->Edit   Ctrl+S->Save"
            Alignment       =   0
            FillColor       =   8421504
            TextColor       =   16777215
            Picture         =   "ProjectTracker.frx":3BF3
            Picture         =   "ProjectTracker.frx":3C0F
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
            Height          =   300
            Left            =   5290
            TabIndex        =   23
            Top             =   0
            Width           =   12975
            _Version        =   65536
            _ExtentX        =   22886
            _ExtentY        =   529
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
            Picture         =   "ProjectTracker.frx":3C2B
            Begin VB.OptionButton Option7 
               Caption         =   "All"
               Height          =   225
               Left            =   1440
               TabIndex        =   31
               Top             =   45
               Width           =   1215
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Pending"
               Height          =   225
               Left            =   2760
               TabIndex        =   30
               Top             =   45
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Hold"
               Height          =   225
               Left            =   9240
               TabIndex        =   29
               Top             =   45
               Width           =   1215
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Done"
               Height          =   225
               Left            =   11040
               TabIndex        =   28
               Top             =   45
               Width           =   1215
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Pause"
               Height          =   225
               Left            =   7680
               TabIndex        =   27
               Top             =   45
               Width           =   1215
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Running"
               Height          =   225
               Left            =   6000
               TabIndex        =   26
               Top             =   45
               Width           =   1215
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Open"
               Height          =   225
               Left            =   4320
               TabIndex        =   25
               Top             =   45
               Width           =   1215
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   300
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
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
               Caption         =   " &Show Task"
               Alignment       =   0
               FillColor       =   9164542
               TextColor       =   0
               Picture         =   "ProjectTracker.frx":3C47
               Picture         =   "ProjectTracker.frx":3C63
            End
         End
         Begin Mh3dlblLib.Mh3dLabel CmdExport 
            Height          =   330
            Left            =   17470
            TabIndex        =   33
            Top             =   8400
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
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
            Caption         =   " &Export "
            FillColor       =   9164542
            TextColor       =   0
            Picture         =   "ProjectTracker.frx":3C7F
            Picture         =   "ProjectTracker.frx":3C9B
         End
         Begin Mh3dlblLib.Mh3dLabel CmdPrint 
            Height          =   330
            Left            =   16380
            TabIndex        =   34
            Top             =   8400
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
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
            Caption         =   " &Print"
            FillColor       =   9164542
            TextColor       =   0
            Picture         =   "ProjectTracker.frx":3CB7
            Picture         =   "ProjectTracker.frx":3CD3
         End
         Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
            Height          =   330
            Left            =   1200
            TabIndex        =   36
            Top             =   8415
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   582
            Calculator      =   "ProjectTracker.frx":3CEF
            Caption         =   "ProjectTracker.frx":3D0F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "ProjectTracker.frx":3D73
            Keys            =   "ProjectTracker.frx":3D91
            Spin            =   "ProjectTracker.frx":3DDB
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
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
            Height          =   330
            Left            =   120
            TabIndex        =   37
            Top             =   8415
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
            Picture         =   "ProjectTracker.frx":3E03
            Picture         =   "ProjectTracker.frx":3E1F
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H008BD6FE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Find"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   330
            Left            =   2640
            TabIndex        =   9
            Top             =   8415
            Width           =   615
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   18870
      _ExtentX        =   33285
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
Attribute VB_Name = "FrmProjectTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Screen_Shot
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_SNAPSHOT = &H2C
Private Const VK_MENU = &H12
Private Const KEYEVENTF_KEYUP = &H2
''
Dim PrintFlag As Boolean
Dim cnProjectTracker As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstItemList As New ADODB.Recordset, rstItemChild As New ADODB.Recordset
Dim SortOrder As String, PrevStr As String, dblBookMark As String, blnRecordExist As Boolean, StartTime As Date, EntryNo As Integer
Dim MsgText As String, MsgSubject As String, Task As String, TaskComments As String, ShotFlag As Boolean
Private Sub cmdRefresh_Click()
Form_Load
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Dim SQL As String
    CenterForm Me
    Me.Top = 1200
    BusySystemIndicator True
    cnProjectTracker.CursorLocation = adUseClient
    If cnProjectTracker.State = adStateOpen Then cnProjectTracker.Close
    cnProjectTracker.Open cnDatabase.ConnectionString
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT * FROM CompanyMaster Where FYCode= '" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstItemList.State = adStateOpen Then rstItemList.Close
    If UserLevel <> "1" Then
    SQL = "AND I.CODE IN (Select Code From BookChild02 Where Member=(Select Code From TeamMemberMaster Where LoginId=" & UserCode & ")) AND (Member=(Select Code From TeamMemberMaster Where LoginId=" & UserCode & ") OR " & UserCode & "=Left(SNo,6)) "
    ElseIf UserLevel = "1" Then
    SQL = " "
    'SQL = "AND (I.CODE IN (Select Code From BookChild02 Where Member=(Select Code From TeamMemberMaster Where LoginId=" & UserCode & ")) OR " & UserCode & "=Left(SNo,6) OR " & UserCode & "='000000' OR " & UserCode & "='000001' OR " & UserCode & "='000005') "
    End If
    rstItemList.Open "SELECT I.Name,C.Correction As Task,C.Status,(Select PrintName From GeneralMaster Where Code=M.Designation)+'-'+M.PrintName As AssignTo,(Select PrintName From GeneralMaster Where Code=H.Designation)+'-'+H.PrintName As AssignBy,C.ArrivedOn,C.TargetDate,BusyCode,I.Code,C.StartDate,C.RectifiedOn,C.Remarks,M.email As FromEmail,H.Email As ToEmail FROM BookMaster I LEFT Join BookChild02 C ON C.Code=I.Code LEFT Join TeamMemberMaster M On M.Code=C.Member INNER Join TeamMemberMaster H On H.LoginId=Left(C.SNo,6) " & _
                                "WHERE [Type]='F' AND " & IIf(Option7.Value, "Status<>''", IIf(Option6.Value, "Status<>'Done'", IIf(Option1.Value, "Status='Open'", IIf(Option2.Value, "Status='Running'", IIf(Option3.Value, "Status='Pause'", IIf(Option5.Value, "Status='Hold'", IIf(Option4.Value, "Status='Done'", "1=1"))))))) & " " & SQL & " ORDER BY Name", cnProjectTracker, adOpenKeyset, adLockOptimistic
    TDBNumber2 = rstItemList.RecordCount
    rstItemList.Filter = adFilterNone
    Set DataGrid1.DataSource = rstItemList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
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
    CloseForm Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu
    Text1.SetFocus
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
                
                If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then Me.ActiveControl.SetFocus Else Call SetButtons(True): SetButtonsForNoRecord: cmdStart.Enabled = True: cmdPause.Enabled = False: cmdEnd.Enabled = False: Mh3dFrame3.Enabled = False: Me.Caption = "Project Manager": SSTab1.Tab = 0
            End If
        End If
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD And Toolbar1.Buttons.Item(3).Enabled Then
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
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
        If Toolbar1.Buttons.Item(1).Enabled Then SSTab1.Tab = 1: SSTab1.SetFocus
        KeyCode = 0
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstItemChild)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(cnProjectTracker)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Option1_Click()
Form_Load
End Sub
Private Sub Option2_Click()
Form_Load
End Sub
Private Sub Option3_Click()
Form_Load
End Sub
Private Sub Option4_Click()
Form_Load
End Sub
Private Sub Option5_Click()
Form_Load
End Sub
Private Sub Option6_Click()
Form_Load
End Sub
Private Sub Option7_Click()
Form_Load
End Sub
Private Sub Text1_Change()
On Error Resume Next
With rstItemList
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
        If SSTab1.Tab = 1 Then
            CenterForm Me
            ViewRecord
            GoTo DisplayInfo
        Else
            If Not (rstItemList.EOF Or rstItemList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            CenterForm Me
            Text1.SetFocus
        End If
        SSTab1.TabEnabled(0) = True
    Else
        SSTab1.TabEnabled(0) = False
        Mh3dFrame2.Enabled = True
        fpSpread3.SetFocus
        GoTo DisplayInfo
    End If
    Exit Sub
DisplayInfo:
   Text2.Text = UserCode: Text6.Text = Username: Text7.Text = rstItemList.Fields("Name").Value
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    If Button.Index = 2 Then
        If rstItemList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstItemList.ActiveConnection = cnProjectTracker
        Do While Not RefreshRecord(rstItemList)
        Loop
        Set DataGrid1.DataSource = rstItemList
        rstItemList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Name", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstItemList.RecordCount > 0 Then rstItemList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstItemList.RecordCount > 0 Then
           rstItemList.MovePrevious
           If rstItemList.BOF Then
              rstItemList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstItemList.RecordCount > 0 Then
           rstItemList.MoveNext
           If rstItemList.EOF Then
              rstItemList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstItemList.RecordCount > 0 Then rstItemList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Call CloseForm(FrmProjectAssigner)
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
    If Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
    End If
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstItemList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstItemList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
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
Private Sub ViewRecord()
    ClearFields
    If rstItemList.EOF Then Exit Sub
    LoadFields
End Sub
Private Sub ClearFields()
    fpSpread3.ClearRange 1, 1, fpSpread3.MaxCols, fpSpread3.MaxRows, True
End Sub
Private Sub LoadFields()
    Call LoadMaterialList(rstItemList.Fields("Code").Value, "T", "")
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    fpSpread3.SetFocus
    blnRecordExist = True
    cmdStart.Enabled = True: cmdPause.Enabled = False: cmdEnd.Enabled = False: Mh3dFrame3.Enabled = True
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstItemList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
Private Sub LoadMaterialList(ByVal strBookCode As String, ByVal LoadGrid As String, ByVal SrNo As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If LoadGrid = "T" Then  'Top Grid
        If rstItemChild.State = adStateOpen Then rstItemChild.Close
        If DatabaseType = "MS SQL" Then
            rstItemChild.Open "SELECT ArrivedOn,Status,Correction,M.Name As MemberName,TargetDate,StartDate,RectifiedOn,Member,Remarks,SNo FROM (BookChild02 T INNER JOIN TeamMemberMaster M ON T.Member=M.Code) INNER JOIN UserMaster U ON M.LoginId=U.Code WHERE T.Code='" & strBookCode & "' AND (" & IIf(UserCode = "000000", "1=1", IIf(UserCode = "000001", "1=1", IIf(UserCode = "000005", "1=1", "LEFT(SNo,6)='" & UserCode & "' OR LoginId='" & UserCode & "'"))) & ") AND 1=CASE WHEN LoginId='" & UserCode & "' THEN CASE WHEN Status<>'Done' AND Status<>'Hold' THEN 1 ELSE 2 END ELSE 1 END ORDER BY ArrivedOn,SNo", cnProjectTracker, adOpenKeyset, adLockReadOnly
        Else
            rstItemChild.Open "SELECT ArrivedOn,Status,Correction,M.Name As MemberName,TargetDate,StartDate,RectifiedOn,Member,Remarks,SNo FROM (BookChild02 T INNER JOIN TeamMemberMaster M ON T.Member=M.Code) INNER JOIN UserMaster U ON M.LoginId=U.Code WHERE T.Code='" & strBookCode & "' AND (" & IIf(UserCode = "000000", "1=1", IIf(UserCode = "000001", "1=1", IIf(UserCode = "000005", "1=1", "LEFT(SNo,6)='" & UserCode & "' OR LoginId='" & UserCode & "'"))) & ") AND IIF(LoginId='" & UserCode & "',(Status<>'Done' AND Status<>'Hold'),'1=1') ORDER BY ArrivedOn,SNo", cnProjectTracker, adOpenKeyset, adLockReadOnly
        End If
        rstItemChild.ActiveConnection = Nothing
        If rstItemChild.RecordCount > 0 Then rstItemChild.MoveFirst
        i = 0
        Do While Not rstItemChild.EOF
            i = i + 1
            With fpSpread3
                .SetText 1, i, Format(rstItemChild.Fields("ArrivedON").Value, "dd-mm-yyyy") 'Assigned ON
                .SetText 2, i, rstItemChild.Fields("Status").Value  'Status
                .SetText 3, i, rstItemChild.Fields("Correction").Value  'Assignment Remarks
                .SetText 4, i, rstItemChild.Fields("MemberName").Value  'Member Name
                .SetText 5, i, Format(rstItemChild.Fields("TargetDate").Value, "dd-mm-yyyy")    'Target Date
                .SetText 6, i, Format(rstItemChild.Fields("StartDate").Value, "dd-mm-yyyy") 'Start Date
                .SetText 7, i, Format(rstItemChild.Fields("RectifiedON").Value, "dd-mm-yyyy")   'End Date
                .SetText 8, i, rstItemChild.Fields("Member").Value  'Member Code
                .SetText 9, i, rstItemChild.Fields("Remarks").Value 'Comments
                .SetText 10, i, rstItemChild.Fields("SNo").Value
            End With
            rstItemChild.MoveNext
        Loop
    End If
    Dim ActiveCellVal As Variant
    If SrNo <> "" Then ActiveCellVal = SrNo Else fpSpread3.GetText 10, fpSpread3.ActiveRow, ActiveCellVal
    If rstItemChild.State = adStateOpen Then rstItemChild.Close
    rstItemChild.Open "SELECT StartDate,EndDate,Remarks FROM BookChild03 T WHERE T.Code='" & strBookCode & "' AND SNo='" & ActiveCellVal & "' ORDER BY StartDate", cnProjectTracker, adOpenKeyset, adLockReadOnly
    rstItemChild.ActiveConnection = Nothing
    If rstItemChild.RecordCount > 0 Then rstItemChild.MoveFirst
    i = 0
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    Do While Not rstItemChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, Format(rstItemChild.Fields("StartDate").Value, "dd-MMM-yyyy hh:mm:ss")
            .SetText 2, i, Format(rstItemChild.Fields("EndDate").Value, "dd-MMM-yyyy hh:mm:ss")
            .SetText 3, i, rstItemChild.Fields("Remarks").Value
        End With
        rstItemChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub
Private Sub cmdStart_Click()
    'If UserCode = "000000" Then Exit Sub
    ShotFlag = False
    Dim ActiveCellVal As Variant
    fpSpread3.GetText 2, fpSpread3.ActiveRow, ActiveCellVal
    If InStr(1, "Open_Running_Pause", ActiveCellVal) = 0 Or CheckEmpty(ActiveCellVal, False) Then Exit Sub
    On Error GoTo ErrorHandler
    If MsgBox("Are you sure to Start Work?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Start !") = vbYes Then
        fpSpread3.GetText 10, fpSpread3.ActiveRow, ActiveCellVal
        cnProjectTracker.BeginTrans
        cnProjectTracker.Execute "INSERT INTO BookChild03 VALUES ('" & rstItemList.Fields("Code").Value & "','" & ActiveCellVal & "','" & Format(DateTime.Now, "dd-MMM-yyyy hh:mm:ss") & "',Null,'')"
        cnProjectTracker.Execute "UPDATE BookChild02 SET Status='Running' WHERE Code='" & rstItemList.Fields("Code").Value & "' AND SNo='" & ActiveCellVal & "'"
        fpSpread3.SetText 2, fpSpread3.ActiveRow, "Running"
        cnProjectTracker.CommitTrans
        cmdStart.Enabled = False: cmdPause.Enabled = True: cmdEnd.Enabled = True: Mh3dFrame2.Enabled = False: Mh3dFrame5.Enabled = True: EntryNo = fpSpread1.DataRowCnt + 1: fpSpread1.SetActiveCell 3, EntryNo: StartTime = DateTime.Now
        fpSpread1.SetText 1, EntryNo, Format(DateTime.Now, "dd-MMM-yyyy hh:mm:ss"): fpSpread1.SetFocus
       MsgText = "I have now started my task assigned by you.This is for your kind information.": MsgSubject = "Task Start at [" & Format(DateTime.Now, "dd-MMM-yyyy hh:mm:ss") & "] (" & rstItemList.Fields("Name").Value & ")"
       Task = "Task Assigned:" & rstItemList.Fields("Task").Value & ""
       TaskComments = "Task Started "
        Call SendEmail
    End If
    Exit Sub
ErrorHandler:
    cnProjectTracker.RollbackTrans
    DisplayError (Err.Description)
End Sub
Private Sub cmdPause_Click()
    On Error GoTo ErrorHandler
    ShotFlag = False
    If MsgBox("Are you sure to Pause Work?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Pause !") = vbYes Then
        If CheckMandatoryFields() Then Exit Sub
        fpSpread1.SetText 2, EntryNo, Format(DateTime.Now, "dd-MMM-yyyy hh:mm:ss")
        Dim CellVal(1 To 3) As Variant, i As Integer, ActiveCellVal As Variant
        fpSpread3.GetText 10, fpSpread3.ActiveRow, ActiveCellVal
        cnProjectTracker.BeginTrans
        cnProjectTracker.Execute "UPDATE BookChild02 SET Status='Pause' WHERE Code='" & rstItemList.Fields("Code").Value & "' AND SNo='" & ActiveCellVal & "'"
            fpSpread3.SetText 2, fpSpread3.ActiveRow, "Pause"
        cnProjectTracker.Execute "DELETE FROM BookChild03 WHERE Code='" & rstItemList.Fields("Code").Value & "' AND SNo='" & ActiveCellVal & "'"
        With fpSpread1
            For i = 1 To .DataRowCnt
                .GetText 1, i, CellVal(1)
                .GetText 2, i, CellVal(2)
                .GetText 3, i, CellVal(3)
                cnProjectTracker.Execute "INSERT INTO BookChild03 VALUES ('" & rstItemList.Fields("Code").Value & "','" & ActiveCellVal & "','" & CellVal(1) & "'," & IIf(CheckEmpty(CellVal(2), False), "Null", "'" & CellVal(2) & "'") & ",'" & CellVal(3) & "')"
            Next
        End With
        cnProjectTracker.CommitTrans
        cmdStart.Enabled = True: cmdPause.Enabled = False: cmdEnd.Enabled = False: Mh3dFrame2.Enabled = True: Mh3dFrame5.Enabled = False
        Me.Caption = "Project Manager"
       MsgText = "I have now paused this task asigned by you.This is for your kind information.": MsgSubject = "Task Pause at [" & Format(DateTime.Now, "dd-MMM-yyyy hh:mm:ss") & "] (" & rstItemList.Fields("Name").Value & ")"
       Task = "Task Assigned:" & rstItemList.Fields("Task").Value & ""
       TaskComments = CellVal(3)
        Call SendEmail
    End If
    Exit Sub
ErrorHandler:
    cnProjectTracker.RollbackTrans
    DisplayError (Err.Description)
End Sub
Private Sub cmdEnd_Click()
    On Error GoTo ErrorHandler
    ShotFlag = False
    If MsgBox("Are you sure to End Work?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm End !") = vbYes Then
        If CheckMandatoryFields() Then Exit Sub
        fpSpread1.SetText 2, EntryNo, Format(DateTime.Now, "dd-MMM-yyyy hh:mm:ss")
        Dim CellVal(1 To 3) As Variant, i As Integer, ActiveCellVal As Variant
        fpSpread3.GetText 10, fpSpread3.ActiveRow, ActiveCellVal
        cnProjectTracker.BeginTrans
        cnProjectTracker.Execute "UPDATE BookChild02 SET Status='Done',RectifiedOn='" & Format(DateTime.Now, "dd-MMM-yyyy") & "' WHERE Code='" & rstItemList.Fields("Code").Value & "' AND SNo='" & ActiveCellVal & "'"
        fpSpread3.SetText 7, fpSpread3.ActiveRow, Format(DateTime.Now, "dd-MM-yyyy")
        fpSpread3.SetText 2, fpSpread3.ActiveRow, "Done"
        cnProjectTracker.Execute "DELETE FROM BookChild03 WHERE Code='" & rstItemList.Fields("Code").Value & "' AND SNo='" & ActiveCellVal & "'"
        With fpSpread1
            For i = 1 To .DataRowCnt
                .GetText 1, i, CellVal(1)
                .GetText 2, i, CellVal(2)
                .GetText 3, i, CellVal(3)
                cnProjectTracker.Execute "INSERT INTO BookChild03 VALUES ('" & rstItemList.Fields("Code").Value & "','" & ActiveCellVal & "','" & CellVal(1) & "'," & IIf(CheckEmpty(CellVal(2), False), "Null", "'" & CellVal(2) & "'") & ",'" & CellVal(3) & "')"
            Next
        End With
        cnProjectTracker.CommitTrans
        cmdStart.Enabled = True: cmdPause.Enabled = False: cmdEnd.Enabled = False: Mh3dFrame2.Enabled = True: Mh3dFrame5.Enabled = False
        Me.Caption = "Project Manager"
       MsgText = "I have now closed this task asigned by you.This is for your kind information.": MsgSubject = "Task Closed at [" & Format(DateTime.Now, "dd-MMM-yyyy hh:mm:ss") & "] (" & rstItemList.Fields("Name").Value & ")"
       Task = "Task Assigned:" & rstItemList.Fields("Task").Value & ""
       TaskComments = CellVal(3)
        Call SendEmail
    End If
    Exit Sub
ErrorHandler:
    cnProjectTracker.RollbackTrans
    DisplayError (Err.Description)
End Sub
Private Sub Timer2_Timer()
        If (Not cmdStart.Enabled) And Mh3dFrame3.Enabled Then
            Dim Sec As Double
            Sec = DateDiff("s", StartTime, DateTime.Now)
            Me.Caption = "Project Manager (" & Format(Int(Sec / 3600), "00") & ":" & Format(Int(Sec / 60) Mod 60, "00") & ":" & Format(Sec Mod 60, "00") & ")"
        End If
End Sub
Private Sub fpSpread3_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant
    fpSpread3.GetText 10, NewRow, ActiveCellVal
    If CheckEmpty(ActiveCellVal, False) Then ActiveCellVal = "000000000000"
    Call LoadMaterialList(rstItemList.Fields("Code").Value, "B", ActiveCellVal)
End Sub
Private Function CheckMandatoryFields() As Boolean
    Dim i As Integer, ActiveCellVal As Variant
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 3, i, ActiveCellVal
            If CheckEmpty(ActiveCellVal, False) Then CheckMandatoryFields = True: DisplayError "Data incomplete in row #" & Trim(Str(i)): .SetFocus: .SetActiveCell 3, i: Exit Function
        Next
    End With
End Function
Private Sub CmdPrint_Click()
PrintFlag = True
On Error GoTo errHandler_print
cmdexport_click
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
        oBook.Worksheets("sheet1").Cells(1, j + 1).Value = "Project List"
        .Range("A1:L1").Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        oBook.Worksheets("sheet1").Rows(2).Font.Bold = True
    'For i = 0 To rstItemList.RecordCount - 1
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
        Dim StartColumn As String, StartRow As String, EndColumn As String, EndRow As String
            StartColumn = "A"
            StartRow = 1
            EndColumn = "L"
            EndRow = rstItemList.RecordCount + 1
            oBook.Activate
            oExcel.Visible = True
            .Columns("A:L").EntireColumn.AutoFit
            .ActiveSheet.pagesetup.Orientation = xlLandscape
            .ActiveSheet.pagesetup.LeftMargin = .InchesToPoints(0.36)
            .ActiveSheet.pagesetup.RightMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.TopMargin = .InchesToPoints(0.5)
            .ActiveSheet.pagesetup.BottomMargin = .InchesToPoints(0.5)
            .ActiveSheet.pagesetup.HeaderMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.FooterMargin = .InchesToPoints(0.25)
            .ActiveSheet.pagesetup.PrintArea = StartColumn & StartRow & ":" & EndColumn & EndRow
            .ActiveSheet.pagesetup.Zoom = False
            .ActiveSheet.pagesetup.FitToPagesTall = False
            .ActiveSheet.pagesetup.FitToPagesWide = 1
            .ActiveSheet.pagesetup.PrintGridlines = True
            .ActiveSheet.pagesetup.PrintTitleRows = "$1:$9"
            .ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF
            If PrintFlag Then
                With CommonDialog1
                    .Copies = 1
                    .flags = &H0&
                    .ShowPrinter
                    oExcel.ActiveSheet.PrintOut
'                    oBook.SaveAs
'                    oBook.Close
'                    oExcel.Quit
                End With
            End If
   End With
   PrintFlag = False
   Screen.MousePointer = vbDefault
   Exit Sub
errcode:
   MsgBox Err.Description, , Err.Source
   PrintFlag = False
   Screen.MousePointer = vbDefault
End Sub
Sub SendEmail()
On Error Resume Next
Screen.MousePointer = vbHourglass
Dim cdoMsg As Object
Dim cdoConf As Object
Dim cdoFields As Object
Dim schema As String
Set cdoMsg = CreateObject("CDO.Message")
Set cdoConf = CreateObject("CDO.Configuration")
Set cdoFields = cdoConf.Fields

cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = rstCompanyMaster.Fields("SmtpServer").Value
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = rstCompanyMaster.Fields("Port").Value '465
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = rstCompanyMaster.Fields("UserName").Value '"production.easyinfosolutionsi@gmail.com"
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = rstCompanyMaster.Fields("Password").Value
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusetls") = True
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
cdoFields.Update
With cdoMsg
    .To = rstItemList.Fields("ToEmail").Value
    .From = rstCompanyMaster.Fields("UserName").Value
    .Subject = MsgSubject
    .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & MsgText & "<Br><Br><b><I>" & Task & "<Br><b>Task Status: " & TaskComments & "<Br><Br><b>Kindly do acknowledge the receipt of the mail</b>.<Br><Br>Thanks & Regards<Br>" & rstItemList.Fields("AssignTo").Value & "<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
'    If ShotFlag = False Then Call Screen_Shot
'    If Dir(App.Path & "\Screen Shot\Screen Shot.bmp", vbDirectory) <> "" Then
'    .AddAttachment App.Path & "\Screen Shot\Screen Shot.bmp"
'    ShotFlag = True
'    End If
    Set .Configuration = cdoConf
'Status Bar
    MdiMainMenu.StatusBar1.Panels(2).Text = "Sending Email..!!!"
    Call Timer
    .Send
End With
If Err.Number = 0 Then
    'MsgBox "Email Send Successfully", , "Email"
        MdiMainMenu.StatusBar1.Panels(2).Text = "Email Send Successfully..!!!"
        Call Timer: Call Timer: Call Timer: Call Timer
Else
'    MsgBox "Email Error" & Err.Number, , "Email"
    MdiMainMenu.StatusBar1.Panels(2).Text = "Email Send Failed..!!!" & Err.Number
    Call Timer: Call Timer
End If
Set cdoMsg = Nothing
Set cdoConf = Nothing
Set cdoFields = Nothing
MdiMainMenu.StatusBar1.Panels(2).Text = " "
Call Timer
   Screen.MousePointer = vbDefault
End Sub
Private Sub Screen_Shot()
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    If ShotFlag = False Then
    MdiMainMenu.StatusBar1.Panels(2).Text = "Screen Capturing..!!"
    Call Timer: Call Timer
        'Press ALT.
        keybd_event VK_MENU, 0, 0, 0
        'Press Print Screen.
        keybd_event VK_SNAPSHOT, 0, 0, 0
        'Release Print Screen.
        keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
        'Release ALT
        keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
        
        DoEvents
        
        Set Picture1.Picture = Clipboard.GetData()
        MdiMainMenu.StatusBar1.Panels(2).Text = "Screen Captured..!!"
        Call Timer
        If Dir(App.Path & "\Screen Shot", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Screen Shot"
        SavePicture Picture1.Picture, App.Path & "\Screen Shot\Screen Shot.bmp"
        MdiMainMenu.StatusBar1.Panels(2).Text = "Screen Capture Saved..!!"
        Call Timer
    End If
    Screen.MousePointer = vbDefault
ErrorHandler:
MdiMainMenu.StatusBar1.Panels(2).Text = "Screen Capturing Failed..!!"
Call Timer
Picture1 = Nothing
Screen.MousePointer = vbDefault
End Sub
Private Sub Timer3_Timer()
    ShotFlag = False
    Static T As Long
    T = T + 60000
    If T / 60000 = 60 Then
    Call Screen_Shot
        T = 0
    MsgSubject = "Live Screen Shot at [" & Format(DateTime.Now, "dd-MMM-yyyy hh:mm:ss") & "] For (" & rstItemList.Fields("Name").Value & ")"
    Call SendEmail
    End If
End Sub
Sub Timer()
Dim iCount As Long
        iCount = 1
        For iCount = 1 To 30000
        iCount = iCount + 1
        Next
End Sub


