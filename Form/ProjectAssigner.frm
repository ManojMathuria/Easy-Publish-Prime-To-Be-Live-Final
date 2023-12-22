VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmProjectAssigner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Assigner"
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
   MaxButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   18870
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9105
      Left            =   15
      TabIndex        =   1
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
      Picture         =   "ProjectAssigner.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   8895
         Left            =   120
         TabIndex        =   3
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
         TabPicture(0)   =   "ProjectAssigner.frx":001C
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
         Tab(0).Control(5)=   "Mh3dLabel1(2)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "DataGrid1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Text1"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Mh3dFrame4"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmdRefresh"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "ImageList1"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "CommonDialog1"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "ProjectAssigner.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame3"
         Tab(1).Control(1)=   "Mh3dFrame2"
         Tab(1).ControlCount=   2
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   4320
            Top             =   3600
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   9240
            Top             =   3240
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
                  Picture         =   "ProjectAssigner.frx":0054
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":0598
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":0ADC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":0BF0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":0D04
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":0E18
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":0F74
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":14B8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":15CC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":1B10
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":1C24
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":1D38
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":1E4C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":1F60
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ProjectAssigner.frx":2074
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdRefresh 
            Height          =   325
            Left            =   18300
            Picture         =   "ProjectAssigner.frx":2188
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Refresh"
            Top             =   -20
            Width           =   325
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
            Height          =   300
            Left            =   5400
            TabIndex        =   18
            Top             =   0
            Width           =   12900
            _Version        =   65536
            _ExtentX        =   22754
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
            Picture         =   "ProjectAssigner.frx":22D2
            Begin VB.OptionButton Option1 
               Caption         =   "Open"
               Height          =   225
               Left            =   4320
               TabIndex        =   25
               Top             =   45
               Width           =   1215
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Running"
               Height          =   225
               Left            =   6000
               TabIndex        =   24
               Top             =   45
               Width           =   1215
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Pause"
               Height          =   225
               Left            =   7680
               TabIndex        =   23
               Top             =   45
               Width           =   1215
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Done"
               Height          =   225
               Left            =   11040
               TabIndex        =   22
               Top             =   45
               Width           =   1215
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Hold"
               Height          =   225
               Left            =   9240
               TabIndex        =   21
               Top             =   45
               Width           =   1215
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Pending"
               Height          =   225
               Left            =   2760
               TabIndex        =   20
               Top             =   45
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton Option7 
               Caption         =   "All"
               Height          =   225
               Left            =   1440
               TabIndex        =   19
               Top             =   45
               Width           =   1215
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   300
               Left            =   0
               TabIndex        =   26
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
               Picture         =   "ProjectAssigner.frx":22EE
               Picture         =   "ProjectAssigner.frx":230A
            End
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
            TabIndex        =   5
            Top             =   8450
            Width           =   10680
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7995
            Left            =   120
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   360
            Width           =   18360
            _ExtentX        =   32385
            _ExtentY        =   14102
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
               Caption         =   "Project Name"
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
               Caption         =   "Task Assigned To"
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
               Caption         =   "Task Assigned By"
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
               DataField       =   "RectifiedOn"
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
               DataField       =   "Remark"
               Caption         =   "Remark"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd-MMM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
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
                  Alignment       =   2
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
                  ColumnWidth     =   2715.024
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7245
            Left            =   -74880
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1440
            Width           =   18375
            _Version        =   65536
            _ExtentX        =   32411
            _ExtentY        =   12779
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
            Picture         =   "ProjectAssigner.frx":2326
            Begin FPSpreadADO.fpSpread fpSpread3 
               Height          =   7005
               Left            =   120
               TabIndex        =   0
               Top             =   120
               Width           =   18150
               _Version        =   524288
               _ExtentX        =   32015
               _ExtentY        =   12356
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
               MaxCols         =   13
               MaxRows         =   1000
               SpreadDesigner  =   "ProjectAssigner.frx":2342
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
               TabIndex        =   8
               Top             =   2280
               Width           =   5775
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame3 
            Height          =   900
            Left            =   -74880
            TabIndex        =   9
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
            Picture         =   "ProjectAssigner.frx":2C93
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
               Left            =   1325
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   435
               Width           =   16950
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
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   120
               Width           =   11950
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
               TabIndex        =   12
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
               TabIndex        =   10
               Top             =   2400
               Width           =   5775
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   11
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
               Picture         =   "ProjectAssigner.frx":2CAF
               Picture         =   "ProjectAssigner.frx":2CCB
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   2
               Left            =   5240
               TabIndex        =   13
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
               Picture         =   "ProjectAssigner.frx":2CE7
               Picture         =   "ProjectAssigner.frx":2D03
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   3
               Left            =   120
               TabIndex        =   15
               Top             =   435
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
               Picture         =   "ProjectAssigner.frx":2D1F
               Picture         =   "ProjectAssigner.frx":2D3B
            End
         End
         Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
            Height          =   330
            Index           =   2
            Left            =   14040
            TabIndex        =   17
            Top             =   8445
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
            Picture         =   "ProjectAssigner.frx":2D57
            Picture         =   "ProjectAssigner.frx":2D73
         End
         Begin Mh3dlblLib.Mh3dLabel CmdExport 
            Height          =   330
            Left            =   17475
            TabIndex        =   28
            Top             =   8445
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
            Picture         =   "ProjectAssigner.frx":2D8F
            Picture         =   "ProjectAssigner.frx":2DAB
         End
         Begin Mh3dlblLib.Mh3dLabel CmdPrint 
            Height          =   330
            Left            =   16400
            TabIndex        =   29
            Top             =   8445
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
            Picture         =   "ProjectAssigner.frx":2DC7
            Picture         =   "ProjectAssigner.frx":2DE3
         End
         Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
            Height          =   330
            Left            =   1200
            TabIndex        =   30
            Top             =   8445
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   582
            Calculator      =   "ProjectAssigner.frx":2DFF
            Caption         =   "ProjectAssigner.frx":2E1F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "ProjectAssigner.frx":2E83
            Keys            =   "ProjectAssigner.frx":2EA1
            Spin            =   "ProjectAssigner.frx":2EEB
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
            TabIndex        =   31
            Top             =   8445
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
            Picture         =   "ProjectAssigner.frx":2F13
            Picture         =   "ProjectAssigner.frx":2F2F
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
            Left            =   2520
            TabIndex        =   6
            Top             =   8445
            Width           =   735
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
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
Attribute VB_Name = "FrmProjectAssigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ToEmail As Variant
Dim PrintFlag As Boolean
Dim cnProjectAssigner As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstItemList As New ADODB.Recordset, rstItemChild As New ADODB.Recordset, rstItemSubChild As New ADODB.Recordset, rstMemberList As New ADODB.Recordset
Dim MemberCode As String, SortOrder As String, dblBookMark, PrevStr As String, blnRecordExist As Boolean, EditMode As Boolean
Dim MsgText As String, MsgSubject As String, ShotFlag As Boolean, TaskComments As String
Private Sub cmdRefresh_Click()
Form_Load
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Dim SQL As String
    CenterForm Me
    Me.Top = 1200
    BusySystemIndicator True
    cnProjectAssigner.CursorLocation = adUseClient
    If cnProjectAssigner.State = adStateOpen Then cnProjectAssigner.Close
    cnProjectAssigner.Open cnDatabase.ConnectionString
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT * FROM CompanyMaster Where FYCode= '" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstItemList.State = adStateOpen Then rstItemList.Close
    If UserLevel <> "1" And UserLevel <> "2" Then
    SQL = "AND I.CODE IN (Select Code From BookChild02 Where Member=(Select Code From TeamMemberMaster Where LoginId=" & UserCode & ")) AND (Member=(Select Code From TeamMemberMaster Where LoginId=" & UserCode & ") OR " & UserCode & "=Left(SNo,6)) "
    ElseIf UserLevel = "1" Then
    SQL = "AND (I.CODE IN (Select Code From BookChild02 Where Member=(Select Code From TeamMemberMaster Where LoginId=" & UserCode & ")) OR " & UserCode & "=Left(SNo,6) OR " & UserCode & "='000000' OR " & UserCode & "='000001' OR " & UserCode & "='000005') "
    End If
    rstItemList.Open "SELECT I.Name,C.Correction As Task,C.Status,(Select PrintName From GeneralMaster Where Code=M.Designation)+'-'+M.PrintName As AssignTo,(Select PrintName From GeneralMaster Where Code=H.Designation)+'-'+H.PrintName As AssignBy,C.ArrivedOn,C.TargetDate,BusyCode,I.Code,C.StartDate,C.RectifiedOn,C.Remarks,H.Email As FromEmail,M.Email As ToEmail,Left(SNo,6) As UserCode,Right(SNo,6) As SNo  FROM BookMaster I Left Join BookChild02 C ON C.Code=I.Code Left Join TeamMemberMaster M On M.Code=C.Member Left Join TeamMemberMaster H On H.LoginId=Left(C.SNo,6) " & _
                                "WHERE [Type]='F' AND " & IIf(Option7.Value, "1=1", IIf(Option6.Value, "(C.Status IS NULL OR Status<>'Done' )", IIf(Option1.Value, "(C.Status IS NULL OR Status<>'Open' )", IIf(Option2.Value, "(C.Status IS NULL OR Status<>'Running')", IIf(Option3.Value, "(C.Status IS NULL OR Status='Pause')", IIf(Option5.Value, "(C.Status IS NULL OR Status='Hold')", IIf(Option4.Value, "(C.Status IS NULL OR Status='Done')", "1=1"))))))) & " " & IIf(Option7.Value, "AND 1=1", SQL) & " ORDER BY SNo,Name", cnProjectAssigner, adOpenKeyset, adLockOptimistic
    If rstMemberList.State = adStateOpen Then rstMemberList.Close
    'rstMemberList.Open "SELECT M.Name+' ('+D.Name+')' As Col0,M.Code,email As ToEmail FROM TeamMemberMaster M INNER JOIN GeneralMaster D ON M.Department=D.Code Where LoginId<>'" & UserCode & "' ORDER BY M.Name", cnProjectAssigner, adOpenKeyset, adLockReadOnly
    rstMemberList.Open "SELECT (Select Name From GeneralMaster Where Code = M.Designation)+'_'+M.PrintName As Col0,M.Code,email As ToEmail,M.PrintName As AssignTo FROM TeamMemberMaster M INNER JOIN GeneralMaster D ON M.Department=D.Code  ORDER BY M.Name", cnProjectAssigner, adOpenKeyset, adLockReadOnly
    TDBNumber2 = rstItemList.RecordCount
    rstItemList.Filter = adFilterNone
    Set DataGrid1.DataSource = rstItemList
    SetButtons True
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
    rstMemberList.ActiveConnection = Nothing
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
                If Not EditMode Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
        End If
        If Not EditMode Then KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD And Toolbar1.Buttons.Item(3).Enabled Then
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0:    Form_Load: Text1.SetFocus: Text1.Tag = Text1.Text: Text1.Text = "": Text1.Text = Text1.Tag
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
        If Toolbar1.Buttons.Item(1).Enabled Then
            SSTab1.Tab = 1: SSTab1.SetFocus
        Else
            If Me.ActiveControl.Name <> "fpSpread3" Then Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread3" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstMemberList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstItemChild)
    Call CloseRecordset(rstItemSubChild)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(cnProjectAssigner)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Option1_Click()
Form_Load: Text1.SetFocus: Text1.Tag = Text1.Text: Text1.Text = "": Text1.Text = Text1.Tag
End Sub
Private Sub Option2_Click()
Form_Load: Text1.SetFocus: Text1.Tag = Text1.Text: Text1.Text = "": Text1.Text = Text1.Tag
End Sub
Private Sub Option3_Click()
Form_Load: Text1.SetFocus: Text1.Tag = Text1.Text: Text1.Text = "": Text1.Text = Text1.Tag
End Sub
Private Sub Option4_Click()
Form_Load: Text1.SetFocus: Text1.Tag = Text1.Text: Text1.Text = "": Text1.Text = Text1.Tag
End Sub
Private Sub Option5_Click()
Form_Load: Text1.SetFocus: Text1.Tag = Text1.Text: Text1.Text = "": Text1.Text = Text1.Tag
End Sub
Private Sub Option6_Click()
Form_Load: Text1.SetFocus: Text1.Tag = Text1.Text: Text1.Text = "": Text1.Text = Text1.Tag
End Sub
Private Sub Option7_Click()
Form_Load: Text1.SetFocus: Text1.Tag = Text1.Text: Text1.Text = "": Text1.Text = Text1.Tag
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
    Dim HiLiteRecord As Boolean, ActiveCellVal As Variant
    Dim UpdateFlag As Integer, i As Integer
    If Button.Index = 2 Then
        If rstItemList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 4 Then
        If (blnRecordExist And AllowMastersModification = 0) Or (UserLevel <> "1" And UserLevel <> "2") Then
            Call DisplayError("You don't have the rights to Edit this Master")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        If CheckMandatoryFields Then Exit Sub
        UpdateFlag = 1
        If UpdateMaterialList("D") Then
            For i = 1 To fpSpread3.DataRowCnt
                fpSpread3.SetActiveCell 1, i
                fpSpread3.GetText 3, i, ActiveCellVal
                If CheckEmpty(ActiveCellVal, False) Then Exit For
                If Not UpdateMaterialList("I") Then
                    UpdateFlag = 0
                    Exit For
                End If
            Next
        End If
        If UpdateFlag Then
            cnProjectAssigner.CommitTrans
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        cnProjectAssigner.RollbackTrans
        Call SetButtons(True)
        SetButtonsForNoRecord
        SSTab1.Tab = 0
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstItemList.ActiveConnection = cnProjectAssigner
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
    Call LoadMaterialList(rstItemList.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    fpSpread3.SetFocus
    blnRecordExist = True
    cnProjectAssigner.BeginTrans
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
Private Sub fpSpread3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cVal As Variant, delFlag As Boolean
    delFlag = True
    On Error Resume Next
    If rstItemSubChild.State = adStateOpen Then rstItemSubChild.Close
    rstItemSubChild.Open "SELECT Distinct Code+Right(SNo,6) As Code FROM BookChild03 Where Code='" & rstItemList.Fields("Code").Value & "'", cnProjectAssigner, adOpenKeyset, adLockReadOnly
    rstItemSubChild.ActiveConnection = Nothing
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
                    If delFlag = True And rstItemSubChild.RecordCount <> 0 Then rstItemSubChild.MoveFirst
                Do While Not rstItemSubChild.EOF
                    fpSpread3.GetText 10, fpSpread3.ActiveRow, cVal
                    If Trim(rstItemSubChild.Fields("Code").Value) = rstItemList.Fields("Code").Value + cVal Then delFlag = False: Exit Do
                    rstItemSubChild.MoveNext
                Loop
            If delFlag = True Then
                If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
                    fpSpread3.DeleteRows fpSpread3.ActiveRow, 1
                    fpSpread3.SetFocus
                End If
            Else
            MsgBox (" You Can't Delete This Task ,Because This Task is Already Under Process.!! ")
            End If
    ElseIf KeyCode = vbKeySpace Then
        Dim Member As Variant
        With fpSpread3
            If .ActiveCol = 4 Then
                .GetText .ActiveCol, .ActiveRow, Member
                Text3.Text = FixQuote(Member)
                If rstMemberList.RecordCount = 0 Then DisplayError ("No Record in Editorial Team Member Master"): .SetActiveCell 4, .ActiveRow: Exit Sub Else rstMemberList.MoveFirst
                rstMemberList.Find "[Col0] = '" & RTrim(Member) & "'"
                SelectionType = "S"
                MemberCode = ""
                Call LoadSelectionList(rstMemberList, "List of Editorial Team Members...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text3, MemberCode)
                Call CloseForm(FrmSelectionList)
                If MemberCode = "" Then
                    .SetActiveCell 4, .ActiveRow
                Else
                    rstMemberList.MoveFirst: rstMemberList.Find "[Code] ='" & MemberCode & "'"
                    .SetText 4, .ActiveRow, Text3.Text
                    .SetText 8, .ActiveRow, MemberCode
                    .SetText 12, .ActiveRow, rstMemberList.Fields("ToEmail").Value
                    .SetFocus
                    Sendkeys "{ENTER}"
                End If
            End If
        End With
    End If
End Sub
Private Sub fpSpread3_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant
    With fpSpread3
        .GetText 3, Row, ActiveCellVal
        If Not CheckEmpty(ActiveCellVal, False) Then
            .GetText 1, Row, ActiveCellVal
            If CheckEmpty(ActiveCellVal, False) Then .SetText 1, Row, Format(DateTime.Date, "dd-MM-yyyy")
            .GetText 2, Row, ActiveCellVal
            If CheckEmpty(ActiveCellVal, False) Then .SetText 2, Row, "Open"
            .GetText 5, Row, ActiveCellVal
            If CheckEmpty(ActiveCellVal, False) Then .GetText 1, Row, ActiveCellVal: .SetText 5, Row, DateAdd("d", 7, GetDate(ActiveCellVal))
            .GetText 6, Row, ActiveCellVal
            If CheckEmpty(ActiveCellVal, False) Then .SetText 6, Row, Format(DateTime.Date, "dd-MM-yyyy")
        End If
    End With
End Sub
Private Sub LoadMaterialList(ByVal strBookCode As String)
    Dim i As Integer, SrNo As String, n As Integer
    On Error GoTo ErrorHandler
    If rstItemChild.State = adStateOpen Then rstItemChild.Close
    rstItemChild.Open "SELECT ArrivedOn,Status,Correction,(Select Name From GeneralMaster Where Code = Designation)+'_'+PrintName As MemberName,TargetDate,StartDate,RectifiedOn,Member,Remarks,RIGHT(SNo,6) As SrNo,Email,Left(SNo,6) As UserCode FROM BookChild02 T LEFT JOIN TeamMemberMaster M ON T.Member=M.Code WHERE T.Code='" & strBookCode & "' AND " & IIf(UserCode = "000000", "1=1", IIf(UserCode = "000001", "1=1", IIf(UserCode = "000005", "1=1", "LEFT(SNo,6)='" & UserCode & "'"))) & " ORDER BY ArrivedOn", cnProjectAssigner, adOpenKeyset, adLockReadOnly
    rstItemChild.ActiveConnection = Nothing
    If rstItemChild.RecordCount > 0 Then rstItemChild.MoveFirst
    i = 0: SrNo = "000000": n = 1
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
            .SetText 10, i, rstItemChild.Fields("SrNo").Value: SrNo = rstItemChild.Fields("SrNo").Value   'Sr No
            .SetText 11, i, "Sent"
            .SetText 12, i, rstItemChild.Fields("Email").Value 'Email
            .SetText 13, i, rstItemChild.Fields("UserCode").Value 'UserCode
        End With
        rstItemChild.MoveNext
    Loop
    With fpSpread3
        For i = .DataRowCnt + 1 To 1000
            .SetText 10, i, Pad(Trim(Val(SrNo) + n), "0", 6, "L")
            n = n + 1
        Next
    End With
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub
Private Function UpdateMaterialList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 12) As Variant
    Dim eDate As String
    On Error GoTo ErrorHandler
    UpdateMaterialList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
    'WHERE T.Code='" & strBookCode & "' AND " & IIf(UserCode = "000000", "1=1", IIf(UserCode = "000001", "1=1", IIf(UserCode = "000005", "1=1", "LEFT(SNo,6)='" & UserCode & "'"))) & "
'        cnProjectAssigner.Execute "DELETE FROM BookChild02 WHERE Code='" & rstItemList.Fields("Code").Value & "' AND Right(SNo,6)='" & CellVal(9) & "' AND (LEFT(SNo,6)='" & UserCode & "' OR '" & UserCode & "'='000000' OR '" & UserCode & "'='000001'  OR '" & UserCode & "'='000005')"
         cnProjectAssigner.Execute "DELETE FROM BookChild02 WHERE Code='" & rstItemList.Fields("Code").Value & "' AND " & IIf(UserCode = "000000", "1=1", IIf(UserCode = "000001", "1=1", IIf(UserCode = "000005", "1=1", "LEFT(SNo,6)='" & UserCode & "'"))) & ""
    Else
        With fpSpread3
            .GetText 1, .ActiveRow, CellVal(1)  'Assigned ON
            .GetText 2, .ActiveRow, CellVal(2)  'Status
            .GetText 3, .ActiveRow, CellVal(3)  'Assignment Remarks
            .GetText 4, .ActiveRow, CellVal(12)  'AssignedTo
            .GetText 5, .ActiveRow, CellVal(4)  'Target Date
            .GetText 6, .ActiveRow, CellVal(5)  'Start Date
            .GetText 7, .ActiveRow, CellVal(6)  'End Date
            .GetText 8, .ActiveRow, CellVal(7)  'Member Code
            .GetText 9, .ActiveRow, CellVal(8)  'Comments
            .GetText 10, .ActiveRow, CellVal(9)  'Sr No
            .GetText 11, .ActiveRow, CellVal(10)  'email
            .GetText 12, .ActiveRow, ToEmail  'email_ID
            .GetText 13, .ActiveRow, CellVal(11)  'UserCode
        End With
        If CellVal(6) = "" Then eDate = "Null" Else eDate = "'" & GetDate(CellVal(6)) & "'"
        cnProjectAssigner.Execute "INSERT INTO BookChild02 VALUES ('" & rstItemList.Fields("Code").Value & "','" & IIf(CellVal(11) = "", UserCode, CellVal(11)) & CellVal(9) & "','" & GetDate(CellVal(1)) & "','" & CellVal(2) & "','" & CellVal(3) & "','" & CellVal(7) & "','" & GetDate(CellVal(4)) & "','" & GetDate(CellVal(5)) & "'," & eDate & ",'E','" & CellVal(8) & "')"
    If rstMemberList.RecordCount = 0 Then
        DisplayError ("No Record in Editorial Team Member Master")
    Else
        rstMemberList.MoveFirst
        rstMemberList.Find "[Code] = '" & RTrim(CellVal(7)) & "'"
    End If
    If CellVal(10) = "" Or CellVal(10) = " " Then MsgText = "You have assigned a new task.This is for your kind information and further action.": MsgSubject = "Task Assigned at [" & Format(DateTime.Now, "dd-MMM-yyyy hh:mm:ss") & "] (" & rstItemList.Fields("Name").Value & ")"
    
        If CellVal(8) <> "" Then
            TaskComments = "Assignment:" & CellVal(3) & ">>> " & "Status:" & CellVal(2) & ">>> " & "Target Date:" & CellVal(4) & ">>>" & "Task Comments: " & CellVal(8)
        Else
            TaskComments = "Assignment:" & CellVal(3) & ">>> " & "Status:" & CellVal(2) & ">>> " & "Target Date:" & CellVal(4) & ">>>"
        End If
    If CellVal(10) = "" Or CellVal(10) = " " Then Call Send_email(rstCompanyMaster.Fields("SmtpServer").Value, rstCompanyMaster.Fields("Port").Value, rstCompanyMaster.Fields("UserName").Value, rstCompanyMaster.Fields("Password").Value, ToEmail, MsgSubject, MsgText, TaskComments, (IIf(IsNull(rstItemList.Fields("AssignBy").Value), Username, rstItemList.Fields("AssignBy").Value)), Trim(rstCompanyMaster.Fields("PrintName").Value), Trim(rstCompanyMaster.Fields("Phone").Value), Trim(rstCompanyMaster.Fields("EMail").Value), CellVal(12), "Dear " + rstMemberList.Fields("AssignTo").Value, "")       ' SendEmail
    End If
    Exit Function
ErrorHandler:
    UpdateMaterialList = False
End Function
Private Sub fpSpread3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckItem() Then SSTab1.Tab = 1: fpSpread3.SetFocus: CheckMandatoryFields = True
End Function
Private Function CheckItem() As Boolean
    Dim i As Integer, AssignedON As Variant, Remarks As Variant, Member As Variant, TargetDate As Variant, StartDate As Variant, Col As Integer
    CheckItem = False
    With fpSpread3
        For i = 1 To .DataRowCnt
            .SetActiveCell 1, i
            .GetText 1, i, AssignedON
            .GetText 3, i, Remarks
            .GetText 5, i, TargetDate
            .GetText 6, i, StartDate
            .GetText 8, i, Member
            If Not CheckEmpty(Remarks, False) Then
                'Assigned ON
                If Len(AssignedON) < 10 Then CheckItem = True: Col = 1: GoTo Err
                If (Not IsDate(GetDate(AssignedON))) Then CheckItem = True: Col = 1: GoTo Err
                'Member
                If Member = "" Then CheckItem = True: Col = 4: GoTo Err
                'Target Date
                If Len(TargetDate) < 10 Then CheckItem = True: Col = 5: GoTo Err
                If (Not IsDate(GetDate(TargetDate))) Or Format(GetDate(TargetDate), "yyyymmdd") < Format(GetDate(AssignedON), "yyyymmdd") Then CheckItem = True: Col = 5: GoTo Err
                'Start Date
                If Len(StartDate) < 10 Then CheckItem = True: Col = 6: GoTo Err
                If (Not IsDate(GetDate(StartDate))) Or Format(GetDate(StartDate), "yyyymmdd") < Format(GetDate(AssignedON), "yyyymmdd") Then CheckItem = True: Col = 6: GoTo Err
            End If
            Exit Function
Err:
            If CheckItem Then DisplayError "Data incomplete in row #" & Trim(Str(i)): .SetActiveCell Col, i: Exit Function
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
Dim cdoMsg As Object
Dim cdoConf As Object
Dim cdoFields As Object
Dim schema As String
On Error GoTo errcode
    'late binding
Set cdoMsg = CreateObject("CDO.Message")
Set cdoConf = CreateObject("CDO.Configuration")
    ' load all default configurations
    cdoConf.Load -1
Set cdoFields = cdoConf.Fields

'Set All Email Properties
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtpout.secureserver.net" 'rstCompanyMaster.Fields("SmtpServer").Value
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 'rstCompanyMaster.Fields("Port").Value '465
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "Sales@easyinfosolution.com" 'rstCompanyMaster.Fields("UserName").Value '"production.easyinfosolutionsi@gmail.com"
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "pubprint123!@#" 'rstCompanyMaster.Fields("Password").Value
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusetls") = True
cdoFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
cdoFields.Update
With cdoMsg
    .From = "Sales@easyinfosolution.com" 'rstCompanyMaster.Fields("UserName").Value
    .To = ToEmail
    '.CC
    '.BCC
    .Subject = MsgSubject
     .HTMLBody = "<Font Face='Calibri' Size='3'>Dear User,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & MsgText & " <Br><b><I>Task >> " & TaskComments & "<Br><b>Kindly do acknowledge the receipt of the mail</b>.<Br><Br>Thanks & Regards<Br>" & rstItemList.Fields("AssignBy").Value & "<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
'    If ShotFlag = False Then Call Screen_Shot
'    If Dir(App.Path & "\Screen Shot\Screen Shot.bmp", vbDirectory) <> "" Then
'    .AddAttachment App.Path & "\Screen Shot\Screen Shot.bmp"
'    ShotFlag = True
'    End If
    Set .Configuration = cdoConf
    .Send
End With
If Err.Number = 0 Then
    MsgBox "Email Send To : " & ToEmail, , "Email" '" & rstItemList.Fields("AssignTo").Value, , "Email"
'Else
'    MsgBox "Email Error " & Err.Description, , "Email"
End If
Exit_Err:
'    'Release object memory
Set cdoMsg = Nothing
Set cdoConf = Nothing
Set cdoFields = Nothing
Exit Sub

errcode:
    Select Case Err.Number
    Case -2147220973  'Could be because of Internet Connection
        MsgBox "Check your internet connection." & vbNewLine & Err.Number & ": " & Err.Description
    Case -2147220975  'Incorrect credentials User ID or password
        MsgBox "Check your login credentials and try again." & vbNewLine & Err.Number & ": " & Err.Description
    Case Else   'Report other errors
        MsgBox "Error encountered while sending email." & vbNewLine & Err.Number & ": " & Err.Description
    End Select

    Resume Exit_Err
End Sub
