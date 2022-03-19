VERSION 5.00
Object = "{01646141-065C-11D4-8ED3-00E07D815373}#1.0#0"; "MBBrowse.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MdiMainMenu 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Easy Publish  21|Rel 05 | 06.29 Version |Production & Inventory Management System"
   ClientHeight    =   8625
   ClientLeft      =   165
   ClientTop       =   705
   ClientWidth     =   11280
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "MdiMainMenu"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   5880
      Top             =   3960
   End
   Begin MBBrowse.BrowseFF BrowseFF1 
      Left            =   720
      Top             =   1440
      _ExtentX        =   1085
      _ExtentY        =   1085
      ReturnOnlyFSDirs=   -1  'True
      ShowCurrentPath =   0   'False
      ShowEditBox     =   0   'False
      ValidatePath    =   0   'False
      StartUpPosition =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Add [Ctrl+A]"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Edit [Ctrl+E]"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Delete [Ctrl+D]"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Save [Ctrl+S]"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Refresh [F5]"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Filter"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Print [Ctrl+P]"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Print Preview [Ctrl+V]"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Mail"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "First [Ctrl+F]"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Previous [Ctrl+P]"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Next [Ctrl+N]"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Last [Ctrl+L]"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   15
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   1440
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
            Picture         =   "MainMenu.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":0550
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":0A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":0BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":0CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":0DD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":0F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1470
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1584
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1AC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1E04
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":202C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11250
      TabIndex        =   1
      Top             =   7950
      Visible         =   0   'False
      Width           =   11280
      Begin VB.PictureBox picOriginal 
         Height          =   5055
         Left            =   1800
         ScaleHeight     =   4995
         ScaleWidth      =   6045
         TabIndex        =   4
         Top             =   0
         Width           =   6105
         Begin VB.PictureBox picStretched 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   7260
            Left            =   0
            ScaleHeight     =   7260
            ScaleWidth      =   4095
            TabIndex        =   5
            Top             =   0
            Width           =   4095
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8325
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6306
            MinWidth        =   6306
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6306
            MinWidth        =   6306
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4357
            MinWidth        =   4357
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1501
            MinWidth        =   1501
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   5400
      Top             =   3960
   End
   Begin VB.Menu MnuCompany 
      Caption         =   "&Company"
      Begin VB.Menu MnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "Create"
         Begin VB.Menu mnuCreate01 
            Caption         =   "With Masters"
         End
         Begin VB.Menu mnuCreate02 
            Caption         =   "Without Masters"
         End
      End
      Begin VB.Menu MnuClose 
         Caption         =   "Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuLine57 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuCompanyChild 
         Caption         =   "Edit Voucher Prifix"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuLine6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu MnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBackup 
         Caption         =   "Backup"
      End
      Begin VB.Menu MnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu MnuLicenceAgreement 
         Caption         =   "License Agreement"
      End
      Begin VB.Menu MnuYouTube 
         Caption         =   "Help Videos (You Tube)"
      End
      Begin VB.Menu MnuRemoteSupprort 
         Caption         =   "Remote Support Software"
      End
      Begin VB.Menu MnuLine344 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuMasters 
      Caption         =   "&Masters"
      Enabled         =   0   'False
      Tag             =   "01000000"
      Begin VB.Menu mnuAccountMaster 
         Caption         =   "Account"
         Tag             =   "01010000"
      End
      Begin VB.Menu mnuAccountGroupMaster 
         Caption         =   "Account Group"
         Tag             =   "01020000"
      End
      Begin VB.Menu mnuRateMaster 
         Caption         =   "Rate"
         Tag             =   "01030000"
         Begin VB.Menu mnuRate 
            Caption         =   "Processing"
            Index           =   1
            Tag             =   "01030100"
         End
         Begin VB.Menu mnuRate 
            Caption         =   "Printing"
            Index           =   2
            Tag             =   "01030200"
         End
         Begin VB.Menu mnuRate 
            Caption         =   "Plate"
            Index           =   3
            Tag             =   "01030300"
         End
         Begin VB.Menu mnuRate 
            Caption         =   "Miscellaneous Operation"
            Index           =   4
            Tag             =   "01030400"
         End
         Begin VB.Menu mnuRate 
            Caption         =   "Binding"
            Index           =   5
            Tag             =   "01030500"
         End
      End
      Begin VB.Menu mnuLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBook 
         Caption         =   "Item"
         Tag             =   "01040000"
         Begin VB.Menu mnuFreshBookMaster 
            Caption         =   "FG"
            Tag             =   "01040100"
         End
         Begin VB.Menu mnuRepairBookMaster 
            Caption         =   "UFG"
            Tag             =   "01040200"
         End
      End
      Begin VB.Menu mnuItemGroupMaster 
         Caption         =   "Item Group"
         Tag             =   "01050000"
      End
      Begin VB.Menu mnuBindingTypeMaster 
         Caption         =   "Binding Type"
         Tag             =   "01060000"
      End
      Begin VB.Menu mnuOperationMaster 
         Caption         =   "Misc. Operation"
         Tag             =   "01070000"
      End
      Begin VB.Menu MnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSizeMaster 
         Caption         =   "Size"
         Tag             =   "01080000"
      End
      Begin VB.Menu mnuFinishSizeMaster 
         Caption         =   "Finish Size"
         Tag             =   "01090000"
      End
      Begin VB.Menu mnuSizeGroupMaster 
         Caption         =   "Size Group"
         Tag             =   "01100000"
      End
      Begin VB.Menu MnuLine8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaperMaster 
         Caption         =   "Paper"
         Tag             =   "01110000"
         Begin VB.Menu mnuPaper 
            Caption         =   "Sheet"
            Index           =   1
            Tag             =   "01110100"
         End
         Begin VB.Menu mnuPaper 
            Caption         =   "Reel"
            Index           =   2
            Tag             =   "01110200"
         End
      End
      Begin VB.Menu mnuPaperUnitMaster 
         Caption         =   "Paper Unit"
         Tag             =   "01120000"
      End
      Begin VB.Menu mnuLine11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColorMaster 
         Caption         =   "Color"
         Tag             =   "01130000"
      End
      Begin VB.Menu mnuLine13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaterialCentreMaster 
         Caption         =   "Material Centre"
         Tag             =   "01140000"
      End
      Begin VB.Menu mnu777 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTaxMaster 
         Caption         =   "Tax"
         Tag             =   "01150000"
      End
      Begin VB.Menu mnuLine9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOutsourceItemMaster 
         Caption         =   "BOM Item"
         Tag             =   "01160000"
      End
      Begin VB.Menu MnuLine58 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHSNCodeMaster 
         Caption         =   "HSN Code"
         Tag             =   "01170000"
      End
      Begin VB.Menu mnuBillingNarrationMaster 
         Caption         =   "Std. Narration"
         Tag             =   "01180000"
      End
      Begin VB.Menu MnuLine15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectManagementMaster 
         Caption         =   "Project Management"
         Tag             =   "01190000"
         Begin VB.Menu mnuProjectManagement 
            Caption         =   "Department"
            Index           =   1
            Tag             =   "01190100"
         End
         Begin VB.Menu mnuProjectManagement 
            Caption         =   "Designation"
            Index           =   2
            Tag             =   "01190200"
         End
         Begin VB.Menu mnuProjectManagement 
            Caption         =   "Project Member"
            Index           =   3
            Tag             =   "01190300"
         End
      End
      Begin VB.Menu MnuLine1500 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMachineMaster 
         Caption         =   "Machine"
         Tag             =   "01200000"
      End
      Begin VB.Menu mnuLine676 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDespatchManagementParent 
         Caption         =   "Despatch Management"
         Tag             =   "01210000"
         Begin VB.Menu mnuDespatchManagement 
            Caption         =   "Packer"
            Index           =   1
            Tag             =   "01210100"
         End
         Begin VB.Menu mnuDespatchManagement 
            Caption         =   "Deliverer"
            Index           =   2
            Tag             =   "01210200"
         End
         Begin VB.Menu mnuDespatchManagement 
            Caption         =   "Transporter"
            Index           =   3
            Tag             =   "01210300"
         End
         Begin VB.Menu mnuDespatchManagement 
            Caption         =   "Booking Route"
            Index           =   4
            Tag             =   "01210400"
         End
      End
      Begin VB.Menu mnuLine001 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserMaster 
         Caption         =   "User"
         Tag             =   "01220000"
      End
   End
   Begin VB.Menu MnuTransactions 
      Caption         =   "&Transactions"
      Enabled         =   0   'False
      Tag             =   "02000000"
      Begin VB.Menu mnuPrintPlanningModule 
         Caption         =   "Print Planning"
         Tag             =   "02010000"
         Begin VB.Menu mnuPrintPlanning 
            Caption         =   "Multi Form Format"
            Index           =   1
            Tag             =   "02010100"
         End
         Begin VB.Menu mnuPrintPlanning 
            Caption         =   "Spread Form Format"
            Index           =   2
            Tag             =   "02010200"
         End
      End
      Begin VB.Menu mnuPurchaseQuotation 
         Caption         =   "Purchase Quotation"
         Tag             =   "02020000"
         Begin VB.Menu mnuPurchaseQuotationJW 
            Caption         =   "Job Work"
            Tag             =   "02020100"
            Begin VB.Menu mnuPurchaseQuotationJobWork 
               Caption         =   "Unit Cost"
               Index           =   10
               Tag             =   "02020101"
            End
            Begin VB.Menu mnuPurchaseQuotationJobWork 
               Caption         =   "Job Work Unit Cost"
               Index           =   11
               Tag             =   "02020102"
            End
            Begin VB.Menu mnuPurchaseQuotationJobWork 
               Caption         =   "Job Work"
               Index           =   12
               Tag             =   "02020103"
            End
         End
         Begin VB.Menu mnuQuotationSupplyInwardFinishedItem 
            Caption         =   "Supply Inward"
            Tag             =   "02020200"
         End
      End
      Begin VB.Menu mnuSalesQuotation 
         Caption         =   "Sale Quotation"
         Tag             =   "02030000"
         Begin VB.Menu mnuSalesQuotationJW 
            Caption         =   "Job Work"
            Tag             =   "02030100"
            Begin VB.Menu mnuSalesQuotationJobWork 
               Caption         =   "Unit Cost"
               Index           =   7
               Tag             =   "02030101"
            End
            Begin VB.Menu mnuSalesQuotationJobWork 
               Caption         =   "Job Work Unit Cost"
               Index           =   8
               Tag             =   "02030102"
            End
            Begin VB.Menu mnuSalesQuotationJobWork 
               Caption         =   "Job Work"
               Index           =   9
               Tag             =   "02030103"
            End
         End
         Begin VB.Menu mnuQuotationSupplyOutwardFinishedItem 
            Caption         =   "Supply Outward"
            Tag             =   "02030200"
         End
      End
      Begin VB.Menu mnuPurchasesOrder 
         Caption         =   "Purchase Order"
         Tag             =   "02040000"
         Begin VB.Menu mnuPurchaseOrderJobWork 
            Caption         =   "Job Work"
            Tag             =   "02040100"
            Begin VB.Menu mnuPurchaseOrderJobWorkFinishedItem 
               Caption         =   "FG Item"
               Tag             =   "02040101"
            End
            Begin VB.Menu mnuPurchaseOrderJobWorkUnfinishedItem 
               Caption         =   "UFG Item"
               Tag             =   "02040102"
            End
            Begin VB.Menu mnuPurchaseOrderJobWorkDigital 
               Caption         =   "Digital"
               Tag             =   "02040103"
            End
         End
         Begin VB.Menu mnuPurchaseOrderSupplyInward 
            Caption         =   "Supply Inward"
            Tag             =   "02040200"
            Begin VB.Menu mnuPurchaseOrderSupplyInwardFinishedItem 
               Caption         =   "FG Item"
               Tag             =   "02040201"
            End
            Begin VB.Menu mnuPurchaseOrderSupplyInwardBOMItem 
               Caption         =   "BOM Item"
               Tag             =   "02040202"
            End
         End
      End
      Begin VB.Menu mnuSalesOrder 
         Caption         =   "Sales Order"
         Tag             =   "02050000"
         Begin VB.Menu mnuSalesOrderJobWork 
            Caption         =   "Job Work"
            Tag             =   "02050100"
            Begin VB.Menu mnuSalesOrderJobWorkFinishedItem 
               Caption         =   "FG Item"
               Tag             =   "02050101"
            End
            Begin VB.Menu mnuSalesOrderJobWorkUnfinishedItem 
               Caption         =   "UFG Item"
               Tag             =   "02050102"
            End
            Begin VB.Menu mnuSalesOrderJobWorkDigital 
               Caption         =   "Digital"
               Tag             =   "02050103"
            End
         End
         Begin VB.Menu mnuSalesOrderSupplyOutwardFinishedItem 
            Caption         =   "Supply Outward"
            Tag             =   "02050200"
         End
      End
      Begin VB.Menu mnuSales 
         Caption         =   "Sales"
         Tag             =   "02060000"
         Begin VB.Menu mnuSalesJW 
            Caption         =   "Job Work"
            Tag             =   "02060100"
            Begin VB.Menu mnuSalesJobWork 
               Caption         =   "Unit Cost"
               Index           =   1
               Tag             =   "02060101"
            End
            Begin VB.Menu mnuSalesJobWork 
               Caption         =   "Job Work Unit Cost"
               Index           =   2
               Tag             =   "02060102"
            End
            Begin VB.Menu mnuSalesJobWork 
               Caption         =   "Job Work"
               Index           =   3
               Tag             =   "02060103"
            End
         End
         Begin VB.Menu mnuSalesSupplyOutwardFinishedItem 
            Caption         =   "Supply Outward"
            Tag             =   "02060200"
         End
      End
      Begin VB.Menu mnuSalesReturn 
         Caption         =   "Sales Return"
         Tag             =   "02070000"
         Begin VB.Menu mnuSalesReturnSupplyOutwardReturnFinishedItem 
            Caption         =   "Supply Outward Return"
            Tag             =   "02070100"
         End
      End
      Begin VB.Menu mnuPurchase 
         Caption         =   "Purchase"
         Tag             =   "02080000"
         Begin VB.Menu mnuPurchaseJW 
            Caption         =   "Job Work"
            Tag             =   "02080100"
            Begin VB.Menu mnuPurchaseJobWork 
               Caption         =   "Unit Cost"
               Index           =   4
               Tag             =   "02080101"
            End
            Begin VB.Menu mnuPurchaseJobWork 
               Caption         =   "Job Work Unit Cost"
               Index           =   5
               Tag             =   "02080102"
            End
            Begin VB.Menu mnuPurchaseJobWork 
               Caption         =   "Jobwork"
               Index           =   6
               Tag             =   "02080103"
            End
         End
         Begin VB.Menu mnuPurchaseSupplyInwardFinishedItem 
            Caption         =   "Supply Inward"
            Tag             =   "02080200"
         End
      End
      Begin VB.Menu mnuPurchaseReturn 
         Caption         =   "Purchase Return"
         Tag             =   "02090000"
         Begin VB.Menu mnuPurchaseReturnSupplyInwardReturnFinishedItem 
            Caption         =   "Supply Inward Return"
            Tag             =   "02090100"
         End
      End
      Begin VB.Menu MnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFinanceModuleParent 
         Caption         =   "Finance"
         Tag             =   "02100000"
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Payment"
            Index           =   1
            Tag             =   "02100100"
         End
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Receipt"
            Index           =   2
            Tag             =   "02100200"
         End
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Journal"
            Index           =   3
            Tag             =   "02100300"
         End
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Contra"
            Index           =   4
            Tag             =   "02100400"
         End
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Debit Note"
            Index           =   5
            Tag             =   "02100500"
         End
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Credit Note"
            Index           =   6
            Tag             =   "02100600"
         End
      End
      Begin VB.Menu MnuLine324 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockTranferFinishedItem 
         Caption         =   "Stock Tranfer"
         Tag             =   "02110000"
      End
      Begin VB.Menu mnuMaterialIn 
         Caption         =   "Material In"
         Tag             =   "02120000"
         Begin VB.Menu mnuMaterialInJobWork 
            Caption         =   "Job Work"
            Tag             =   "02120100"
         End
         Begin VB.Menu mnuMaterialInSupplyInward 
            Caption         =   "Supply Inward"
            Tag             =   "02120200"
         End
      End
      Begin VB.Menu mnuMaterialOut 
         Caption         =   "Material Out"
         Tag             =   "02130000"
         Begin VB.Menu mnuMaterialOutJobWork 
            Caption         =   "Job Work"
            Tag             =   "02130100"
         End
         Begin VB.Menu mnuMaterialOutSupplyOutward 
            Caption         =   "Supply Outward"
            Tag             =   "02130200"
         End
      End
      Begin VB.Menu mnuBookProcessOrder 
         Caption         =   "Item Processing Order"
         Tag             =   "02140000"
      End
      Begin VB.Menu MnuLine12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaperModuleParent 
         Caption         =   "Paper"
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Purchase Order"
            Index           =   1
            Tag             =   "02150100"
         End
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Issue"
            Index           =   2
            Tag             =   "02150200"
         End
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Receipt"
            Index           =   3
            Tag             =   "02150300"
         End
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Transfer"
            Index           =   4
            Tag             =   "02150400"
         End
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Debit Note"
            Index           =   5
            Tag             =   "02150500"
         End
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Credit Note"
            Index           =   6
            Tag             =   "02150600"
         End
      End
      Begin VB.Menu mnuLine10 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMaterialIssueOrder 
         Caption         =   "Material Issue Order"
         Tag             =   "02160000"
      End
      Begin VB.Menu MnuMaterialMovement 
         Caption         =   "BOM Item Movement"
         Tag             =   "02170000"
      End
      Begin VB.Menu MnuLine16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockJournal 
         Caption         =   "Stock Journal"
         Tag             =   "02180000"
         Begin VB.Menu mnuStockJournalRawMaterial 
            Caption         =   "Raw Material"
            Tag             =   "02180100"
         End
         Begin VB.Menu mnuStockJournalFinishedGoods 
            Caption         =   "Finished Goods"
            Tag             =   "02180200"
         End
      End
      Begin VB.Menu mnu000 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPackingSlip 
         Caption         =   "Packing Slip"
         Tag             =   "02190000"
      End
   End
   Begin VB.Menu MnuDisplay 
      Caption         =   "&Display"
      Enabled         =   0   'False
      Tag             =   "03000000"
      Begin VB.Menu MnuFinalResult 
         Caption         =   "Final Result"
         Tag             =   "03010000"
      End
      Begin VB.Menu MnuTrialBalance 
         Caption         =   "Trial Balance"
         Tag             =   "03020000"
      End
      Begin VB.Menu MnuAccountBooks 
         Caption         =   "Account Books"
         Tag             =   "03030000"
         Begin VB.Menu MnuDay 
            Caption         =   "Day Book"
            Tag             =   "03030100"
         End
         Begin VB.Menu MnuLedger 
            Caption         =   "Ledger"
            Tag             =   "03030200"
            Begin VB.Menu MnuAccountWise 
               Caption         =   "Account-Wise"
               Index           =   23
               Tag             =   "03030201"
            End
         End
      End
      Begin VB.Menu MnuAccountSummary 
         Caption         =   "Account Summary"
         Tag             =   "03040000"
      End
      Begin VB.Menu MnuCostCentre 
         Caption         =   "Cost Centre Report"
         Tag             =   "03050000"
      End
      Begin VB.Menu MnuOutStandingAnalysis 
         Caption         =   "Outstanding Analysis"
         Tag             =   "03060000"
      End
      Begin VB.Menu MnuInterestCalculation 
         Caption         =   "Interest Calculation"
         Tag             =   "03070000"
      End
      Begin VB.Menu MnuProduction 
         Caption         =   "Production Scheduling"
         Tag             =   "03080000"
         Begin VB.Menu MnuProductionScheduling 
            Caption         =   "Print Production Scheduling"
            Index           =   1
            Tag             =   "03080100"
         End
         Begin VB.Menu MnuProductionScheduling 
            Caption         =   "Plate Production Scheduling"
            Index           =   2
            Tag             =   "03080200"
         End
         Begin VB.Menu MnuProductionScheduling 
            Caption         =   "Paper Cutting Scheduling"
            Index           =   3
            Tag             =   "03080300"
         End
         Begin VB.Menu MnuProductionScheduling 
            Caption         =   "Update Dispatch"
            Index           =   4
            Tag             =   "03080400"
            Visible         =   0   'False
         End
         Begin VB.Menu MnuProductionScheduling 
            Caption         =   "Production Schedule Print"
            Index           =   5
            Tag             =   "03080500"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu MnuStockStatus 
         Caption         =   "Stock Status"
         Tag             =   "03090000"
         Begin VB.Menu MnuStockLedger 
            Caption         =   "Physical Stock Audit"
            Index           =   0
            Tag             =   "03090100"
         End
         Begin VB.Menu MnuStockLedger 
            Caption         =   "Inventory Ledger "
            Index           =   1
            Tag             =   "03090200"
         End
         Begin VB.Menu MnuStockLedger 
            Caption         =   "Closing Stock Alphabetical "
            Index           =   2
            Tag             =   "03090300"
         End
         Begin VB.Menu MnuStockLedger 
            Caption         =   "Stock List - Short Item Analysis "
            Index           =   33
            Tag             =   "03090400"
         End
      End
      Begin VB.Menu MnuLine59 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuOrderProcessingStatus 
         Caption         =   "Order Status"
         Tag             =   "03100000"
         Begin VB.Menu MnuOrdersSJW 
            Caption         =   "Job-Work"
            Tag             =   "03100100"
            Begin VB.Menu MnuOrdersPartyWise 
               Caption         =   "Purchase Orders-Party-Wise-Detailed"
               Index           =   35
               Tag             =   "03100101"
            End
            Begin VB.Menu MnuOrdersPartyWise 
               Caption         =   "Purchase Orders-Party-wise-Summarised"
               Index           =   36
               Tag             =   "03100102"
            End
            Begin VB.Menu MnuOrdersPartyWise 
               Caption         =   "-"
               Index           =   37
            End
            Begin VB.Menu MnuOrdersPartyWise 
               Caption         =   "Sales Orders-Party-Wise-Detailed"
               Index           =   38
               Tag             =   "03100103"
            End
            Begin VB.Menu MnuOrdersPartyWise 
               Caption         =   "Sales Orders-Party-wise-Summarised"
               Index           =   39
               Tag             =   "03100104"
            End
         End
         Begin VB.Menu MnuOrdersSIW 
            Caption         =   "Supply IN-Ward"
            Tag             =   "03100200"
            Begin VB.Menu MnuPOrdersPartyWise 
               Caption         =   "Purchase Orders Order-Wise"
               Index           =   39
               Tag             =   "03100201"
            End
            Begin VB.Menu MnuPOrdersPartyWise 
               Caption         =   "Purchase Orders Party-wise"
               Index           =   40
               Tag             =   "03100202"
            End
            Begin VB.Menu MnuPOrdersPartyWise 
               Caption         =   "Purchase Orders Item-wise"
               Index           =   41
               Tag             =   "03100203"
            End
         End
         Begin VB.Menu MnuOrdersSOW 
            Caption         =   "Supply Out-Ward"
            Tag             =   "03100300"
            Begin VB.Menu MnuSOrdersPartyWise 
               Caption         =   "Sale Orders Order-wise"
               Index           =   42
               Tag             =   "03100301"
            End
            Begin VB.Menu MnuSOrdersPartyWise 
               Caption         =   "Sale Orders Party-wise"
               Index           =   43
               Tag             =   "03100302"
            End
            Begin VB.Menu MnuSOrdersPartyWise 
               Caption         =   "Sale Orders Item-wise"
               Index           =   44
               Tag             =   "03100303"
            End
         End
      End
      Begin VB.Menu MenuSaleAnalysis 
         Caption         =   "Sales Analysis"
         Tag             =   "03110000"
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Item-Wise"
            Index           =   3
            Tag             =   "03110100"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Return Item-Wise"
            Index           =   4
            Tag             =   "03110200"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales And Sales Return Item-Wise"
            Index           =   5
            Tag             =   "03110300"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Net Sales Item-Wise"
            Index           =   6
            Tag             =   "03110400"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales One Party Item-Wise"
            Index           =   8
            Tag             =   "03110500"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Return One Party Item-Wise"
            Index           =   9
            Tag             =   "03110600"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales And Sales Return One Party Item-Wise"
            Index           =   10
            Tag             =   "03110700"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Net Sales One Party Item-Wise"
            Index           =   11
            Tag             =   "03110800"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "-"
            Index           =   12
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Party-Wise"
            Index           =   22
            Tag             =   "03110900"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Return Party-Wise"
            Index           =   23
            Tag             =   "03111000"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales And Sales Return Party-Wise"
            Index           =   24
            Tag             =   "03111100"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Net Sales Party-Wise"
            Index           =   25
            Tag             =   "03111200"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "-"
            Index           =   26
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales One Item Party-Wise"
            Index           =   27
            Tag             =   "03111300"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Return One Item Party-Wise"
            Index           =   28
            Tag             =   "03111400"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales And Sales Return One Item Party-Wise"
            Index           =   29
            Tag             =   "03111500"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Net Sales One Item Party-Wise"
            Index           =   30
            Tag             =   "03111600"
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "-"
            Index           =   31
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Voucher-Wise"
            Index           =   32
            Tag             =   "03111700"
         End
      End
      Begin VB.Menu MenuPurchaseAnalysis 
         Caption         =   "Purchase Analysis"
         Tag             =   "03120000"
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Item-Wise"
            Index           =   53
            Tag             =   "03120100"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Return Item-Wise"
            Index           =   54
            Tag             =   "03120200"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase And Purchase Return Item-Wise"
            Index           =   55
            Tag             =   "03120300"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Net Purchase Item-Wise"
            Index           =   56
            Tag             =   "03120400"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "-"
            Index           =   57
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase One Party Item-Wise"
            Index           =   58
            Tag             =   "03120500"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Return One Party Item-Wise"
            Index           =   59
            Tag             =   "03120600"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase And Purchase Return One Party Item-Wise"
            Index           =   60
            Tag             =   "03120700"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Net Purchase One Party Item-Wise"
            Index           =   61
            Tag             =   "03120800"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "-"
            Index           =   62
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Party-Wise"
            Index           =   63
            Tag             =   "03120900"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Return Party-Wise"
            Index           =   64
            Tag             =   "03121000"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase And Purchase Return Party-Wise"
            Index           =   65
            Tag             =   "03121100"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Net Purchase Party-Wise"
            Index           =   66
            Tag             =   "03121200"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "-"
            Index           =   67
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase One Item Party-Wise"
            Index           =   68
            Tag             =   "03121300"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Return One Item Party-Wise"
            Index           =   69
            Tag             =   "03121400"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase And Purchase Return One Item Party-Wise"
            Index           =   70
            Tag             =   "03121500"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Net Purchase One Item Party-Wise"
            Index           =   71
            Tag             =   "03121600"
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "-"
            Index           =   72
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Voucher-Wise"
            Index           =   73
            Tag             =   "03121700"
         End
      End
      Begin VB.Menu MnuLine60 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MenuPaperIssueReceipt 
         Caption         =   "Paper Ledger"
         Tag             =   "03130000"
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Receipt Party-Wise"
            Index           =   11
            Tag             =   "03130100"
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Receipt Order-Wise"
            Index           =   12
            Tag             =   "03130200"
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Receipt Without-Order"
            Index           =   13
            Tag             =   "03130300"
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Issue Party-Wise"
            Index           =   14
            Tag             =   "03130400"
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Issue Order-Wise"
            Index           =   15
            Tag             =   "03130500"
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Issue Without-Order"
            Index           =   16
            Tag             =   "03130600"
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Paper Transfer Ledger"
            Index           =   17
            Tag             =   "03130700"
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Paper Pending Order"
            Index           =   18
            Tag             =   "03130800"
         End
      End
   End
   Begin VB.Menu MnuReports 
      Caption         =   "&Reports"
      Enabled         =   0   'False
      Tag             =   "04000000"
      Begin VB.Menu MnuPrintPlanningRegister 
         Caption         =   "Print Planning Register"
         Tag             =   "04010000"
         Begin VB.Menu MnuBookPrintPlanningRegister 
            Caption         =   "Multi Form Format"
            Tag             =   "04010100"
         End
         Begin VB.Menu MnuTitlePrintPlanningRegister 
            Caption         =   "Single Form Format"
            Tag             =   "04010200"
         End
      End
      Begin VB.Menu MnuPOStatusRegister 
         Caption         =   "Order Status Register"
         Tag             =   "04020000"
         Begin VB.Menu MnuPOStatusRegister01 
            Caption         =   "Itemwise"
            Tag             =   "04020100"
         End
         Begin VB.Menu MnuPOStatusRegister05 
            Caption         =   "Orderwise"
            Tag             =   "04020200"
         End
         Begin VB.Menu MnuPOStatusRegister03 
            Caption         =   "Multi Form Partywise"
            Tag             =   "04020300"
         End
         Begin VB.Menu MnuPOStatusRegister02 
            Caption         =   "Spread Form Partywise"
            Tag             =   "04020400"
         End
         Begin VB.Menu MnuPOStatusRegister04 
            Caption         =   "Binding Partywise"
            Tag             =   "04020500"
         End
      End
      Begin VB.Menu MnuLine50 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPaperIssueRegister 
         Caption         =   "Paper Purchase Ledger"
         Tag             =   "04030000"
      End
      Begin VB.Menu MnuPaperStockRegister 
         Caption         =   "Paper Stock Ledger"
         Tag             =   "04040000"
      End
      Begin VB.Menu MnuOpBal 
         Caption         =   "Paper Opening Balance"
         Tag             =   "04050000"
      End
      Begin VB.Menu MnuLine51 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMaterialStockRegister 
         Caption         =   "BOM Item Stock Register"
         Tag             =   "04060000"
         Begin VB.Menu MnuMaterialStockRegister01 
            Caption         =   "Godownwise/Itemwise/BOM Itemwise"
            Tag             =   "04060100"
         End
         Begin VB.Menu MnuMaterialStockRegister02 
            Caption         =   "Godownwise/BOM Itemwise"
            Tag             =   "04060200"
         End
      End
      Begin VB.Menu MnuLine52 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPOStatusReg 
         Caption         =   "Purchase Order Status"
         Tag             =   "04070000"
         Begin VB.Menu MnuPOStatusReg02 
            Caption         =   "BOM Item"
            Tag             =   "04070100"
         End
         Begin VB.Menu MnuPOStatusReg03 
            Caption         =   "Printed Items (BOM)"
            Tag             =   "04070200"
            Begin VB.Menu MnuPOStatusReg0301 
               Caption         =   "FG Item"
               Tag             =   "04070201"
            End
            Begin VB.Menu MnuPOStatusReg0302 
               Caption         =   "UFG Item"
               Tag             =   "04070202"
            End
            Begin VB.Menu MnuPOStatusReg0303 
               Caption         =   "Spread Format"
               Tag             =   "04070204"
            End
         End
      End
      Begin VB.Menu MnuLine53 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrderLedger 
         Caption         =   "Order Ledger"
         Tag             =   "04080000"
         Begin VB.Menu mnuPurchaseOrder 
            Caption         =   "Purchase Order"
            Tag             =   "04080100"
            Begin VB.Menu MnuPOLedgerParty 
               Caption         =   "Party-Wise- Detailed"
               Index           =   11
               Tag             =   "04080101"
            End
            Begin VB.Menu MnuPOLedgerParty 
               Caption         =   "Party-Wise- Summrised"
               Index           =   12
               Tag             =   "04080102"
            End
         End
         Begin VB.Menu mnuSaleOrder 
            Caption         =   "Sale Order"
            Tag             =   "04080200"
            Begin VB.Menu MnuSOLedgerParty 
               Caption         =   "Party-Wise- Detailed"
               Index           =   21
               Tag             =   "04080201"
            End
            Begin VB.Menu MnuSOLedgerParty 
               Caption         =   "Party-Wise- Summrised"
               Index           =   22
               Tag             =   "04080202"
            End
         End
      End
      Begin VB.Menu MnuOrderProcessing 
         Caption         =   "Order Status"
         Tag             =   "04090000"
         Begin VB.Menu MnuPendingOrders 
            Caption         =   "Job-Work"
            Index           =   1
            Tag             =   "04090100"
            Begin VB.Menu MnuPurchaseSaleOrderParty 
               Caption         =   "Purchase Orders-Party-Wise-Detailed"
               Index           =   11
               Tag             =   "04090101"
            End
            Begin VB.Menu MnuPurchaseSaleOrderParty 
               Caption         =   "Purchase Orders-Party-wise-Summarised"
               Index           =   12
               Tag             =   "04090102"
            End
            Begin VB.Menu MnuPurchaseSaleOrderParty 
               Caption         =   "Sale Orders-Party-wise-Detailed"
               Index           =   21
               Tag             =   "04090103"
            End
            Begin VB.Menu MnuPurchaseSaleOrderParty 
               Caption         =   "Sale Orders-Party-wise-Summarised"
               Index           =   22
               Tag             =   "04090104"
            End
         End
         Begin VB.Menu MnuPendingOrdersPO 
            Caption         =   "Supply INward"
            Index           =   2
            Tag             =   "04090200"
            Begin VB.Menu MnuPurchaseSaleOrderPartyPO 
               Caption         =   "Purchase Orders Order-Wise"
               Index           =   13
               Tag             =   "04090201"
            End
            Begin VB.Menu MnuPurchaseSaleOrderPartyPO 
               Caption         =   "Purchase Orders-Party-wise"
               Index           =   14
               Tag             =   "04090202"
            End
            Begin VB.Menu MnuPurchaseSaleOrderPartyPO 
               Caption         =   "Purchase Orders-Item-wise"
               Index           =   15
               Tag             =   "04090203"
            End
         End
         Begin VB.Menu MnuPendingOrdersSO 
            Caption         =   "Supply Outward"
            Index           =   3
            Tag             =   "04090300"
            Begin VB.Menu MnuPurchaseSaleOrderPartySO 
               Caption         =   "Sale Orders Order-wise"
               Index           =   23
               Tag             =   "04090301"
            End
            Begin VB.Menu MnuPurchaseSaleOrderPartySO 
               Caption         =   "Sale Orders-Party-wise"
               Index           =   24
               Tag             =   "04090302"
            End
            Begin VB.Menu MnuPurchaseSaleOrderPartySO 
               Caption         =   "Sale Orders-Item-wise"
               Index           =   25
               Tag             =   "04090303"
            End
         End
      End
      Begin VB.Menu MenuQuotation 
         Caption         =   "Quotation Processing"
         Tag             =   "04100000"
      End
      Begin VB.Menu MnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIssueReceipt 
         Caption         =   "Issue-Receipt Analysis"
         Tag             =   "04110000"
         Begin VB.Menu MnuItemIssueReceipt 
            Caption         =   "Item-wise"
            Index           =   16
            Tag             =   "04110100"
         End
         Begin VB.Menu MnuItemIssueReceipt 
            Caption         =   "Item Party-wise"
            Index           =   17
            Tag             =   "04110200"
         End
         Begin VB.Menu MnuItemIssueReceipt 
            Caption         =   "Item Group-wise"
            Index           =   18
            Tag             =   "04110300"
         End
         Begin VB.Menu MnuItemIssueReceipt 
            Caption         =   "Item Voucher-wise"
            Index           =   19
            Tag             =   "04110400"
         End
         Begin VB.Menu MnuItemIssueReceipt 
            Caption         =   "Item Date-wise"
            Index           =   20
            Tag             =   "04110500"
         End
      End
      Begin VB.Menu MnuLine56 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPendingPaymentRegister 
         Caption         =   "Pending Payment Register"
         Tag             =   "04120000"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuPendingDNRegister 
         Caption         =   "Pending Debit Notes Register"
         Tag             =   "04130000"
      End
      Begin VB.Menu MnuLine54 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBookList1 
         Caption         =   "List of Items"
         Tag             =   "04140000"
         Begin VB.Menu MnuBookList 
            Caption         =   "Items Details"
            Index           =   1
            Tag             =   "04140100"
         End
         Begin VB.Menu MnuBookList 
            Caption         =   "Items Weight"
            Index           =   2
            Tag             =   "04140200"
         End
      End
      Begin VB.Menu MnuCorrectionList 
         Caption         =   "Project Status Report"
         Tag             =   "04150000"
      End
      Begin VB.Menu MnuLine55 
         Caption         =   "-"
      End
      Begin VB.Menu MnuProductionPlanning 
         Caption         =   "Production Fore-Casting"
         Tag             =   "04160000"
         Begin VB.Menu MnuProductionPlanning01 
            Caption         =   "Main Orders"
            Tag             =   "04160100"
         End
         Begin VB.Menu MnuProductionPlanning02 
            Caption         =   "Supplement Orders"
            Tag             =   "04160200"
         End
      End
   End
   Begin VB.Menu MnuUtilities 
      Caption         =   "&Utilities"
      Enabled         =   0   'False
      Tag             =   "05000000"
      Begin VB.Menu MnuEmailUtilities 
         Caption         =   "Email Profile"
         Tag             =   "05010000"
      End
      Begin VB.Menu MnuPrintUtilities 
         Caption         =   "Print Utilities"
         Tag             =   "05020000"
         Begin VB.Menu MnuBookPOPrintUtility1 
            Caption         =   "Item Order"
            Tag             =   "05020100"
            Begin VB.Menu MnuBookPOPrintUtility 
               Caption         =   "Jobwork And Unit Cost"
               Index           =   1
               Tag             =   "05020101"
            End
            Begin VB.Menu MnuBookPOPrintUtility 
               Caption         =   "JobCard"
               Index           =   2
               Tag             =   "05020102"
            End
            Begin VB.Menu MnuBookPOPrintUtility 
               Caption         =   "Plate Orders"
               Index           =   3
               Tag             =   "05020103"
            End
            Begin VB.Menu MnuBookPOPrintUtility 
               Caption         =   "Paper-Requisition-Slip"
               Index           =   4
               Tag             =   "05020104"
            End
            Begin VB.Menu MnuBookPOPrintUtility 
               Caption         =   "Quotation Format"
               Index           =   5
               Tag             =   "05020105"
            End
         End
         Begin VB.Menu MnuPaperPOPrintUtility 
            Caption         =   "Paper Order"
            Tag             =   "05020200"
         End
      End
      Begin VB.Menu MnuLine44 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBookReceiptBusy 
         Caption         =   "Item Receipt (Busy)"
         Tag             =   "05030000"
      End
      Begin VB.Menu mnuCostSheet 
         Caption         =   "Cost Estimation"
         Tag             =   "05040000"
      End
      Begin VB.Menu mnuItemOpBal 
         Caption         =   "Mat Centrewise Item Op Bal"
         Tag             =   "05050000"
      End
      Begin VB.Menu mnuDiscount 
         Caption         =   "Discount Structure"
         Tag             =   "05060000"
      End
      Begin VB.Menu MnuImportBal 
         Caption         =   "Import Balances"
         Tag             =   "05070000"
         Begin VB.Menu MnuImportBal01 
            Caption         =   "Order"
            Tag             =   "05070100"
         End
         Begin VB.Menu MnuImportBal02 
            Caption         =   "Paper"
            Tag             =   "05070200"
         End
         Begin VB.Menu MnuImportBal03 
            Caption         =   "BOM Item"
            Tag             =   "05070300"
         End
      End
   End
   Begin VB.Menu mnuProjectManagementParent 
      Caption         =   "&Project Management"
      Enabled         =   0   'False
      Tag             =   "06000000"
      Begin VB.Menu mnuEditorial 
         Caption         =   "Editorial"
         Tag             =   "06010000"
         Begin VB.Menu mnuProject 
            Caption         =   "Project Assigner"
            Index           =   1
            Tag             =   "06000100"
         End
         Begin VB.Menu mnuProject 
            Caption         =   "Project Tracker"
            Index           =   2
            Tag             =   "06000200"
         End
      End
   End
   Begin VB.Menu MnuHelpm 
      Caption         =   "&Help"
      Tag             =   "07000000"
      Begin VB.Menu MnuHelp 
         Caption         =   "Users Manual – Easy Publish"
         Tag             =   "07010000"
      End
   End
End
Attribute VB_Name = "MdiMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iCount As Long
Public Version As String
Private WithEvents oHuffman As clsHuffman
Attribute oHuffman.VB_VarHelpID = -1
Private oRegistry As New clsRegistry
Private Developer As String
Dim rstDBList As New ADODB.Recordset
Dim Decrypt As Variant
Private Sub Command1_Click()
    Dim R As Long
        R = ShellExecute(0, "open", "https://www.youtube.com/channel/UCW5RVD8qIBTGzCSRM03U7Cw", 0, 0, 1)
        R = ShellExecute(0, "open", "http://www.easyinfosolution.com", 0, 0, 1)
        iCount = 1
        For iCount = 1 To 1000000000
            iCount = iCount + 1
        Next
        iCount = 1
        For iCount = 1 To 100000
        iCount = iCount + 1
        Next
End Sub
Private Sub MnuLicenceAgreement_Click()
    On Error Resume Next
    Load frmLicenceAgreement
    If Err.Number <> 364 Then frmLicenceAgreement.Show
End Sub
Private Sub MnuRemoteSupprort_Click()
    Dim R As Long
    If Dir(App.Path & "\AnyDesk.exe", vbDirectory) = "" Then
            R = ShellExecute(0, "open", "https://anydesk.com/en/downloads/windows", 0, 0, 1)
    Else
            Shell (App.Path & "/AnyDesk.exe")
    End If
End Sub
Private Function Decrypted(Encrypt As Variant, Decrypt)
    Dim e As Long, i As Long, s As Long, n As Long, j As Long, te As Long, ti As Long, ts As Long, tn As Long, K As Long
    Dim eFlag As Boolean, iFlag As Boolean, sFlag As Boolean, nFlag As Boolean
    j = 0: e = 0: i = 0: s = 0: n = 0: Decrypt = "": K = 0: eFlag = True: iFlag = False: sFlag = False: nFlag = False
    
    K = Len(Trim(Encrypt)) - 1
    For j = 1 To K
     If Mid(Trim(Encrypt), j, 1) <> "§" Then
     Decrypt = Decrypt + Mid(Trim(Encrypt), j, 1)
     Else
     If nFlag = True Then n = j: nFlag = False
     If sFlag = True Then s = j: sFlag = False: nFlag = True
     If iFlag = True Then i = j: iFlag = False: sFlag = True
     If eFlag = True Then e = j: eFlag = False: iFlag = True
     End If
    Next j
    Decrypt = ""
    te = 1: ti = e + 1: ts = i + 1: tn = s + 1:
    For j = 1 To K
    te = te + 1
    If j < ((e - 2) * (4)) Then Decrypt = Decrypt + Mid(Trim(Encrypt), te, 1)
    ti = ti + 1
    If j < ((i - e - 2) * (4)) Then Decrypt = Decrypt + Mid(Trim(Encrypt), ti, 1)
    ts = ts + 1
    If j < ((s - i - 2) * (4)) Then Decrypt = Decrypt + Mid(Trim(Encrypt), ts, 1)
    tn = tn + 1
    If j < ((n - s - 2) * (4)) Then Decrypt = Decrypt + Mid(Trim(Encrypt), tn, 1)
    j = j + 3
    Next j
    Decrypted = True
End Function
Private Sub MDIForm_Load()
Version = "EasyPublish |Rel " & Format(App.Major, "00") & "." & Format(App.Minor, "00") & " Ver " & App.Minor & "." & App.Revision & " |Production & Inventory Management System"
MdiMainMenu.Caption = Version
If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    If GetSystemMetrics(SM_CXSCREEN) < 800 Or GetSystemMetrics(SM_CYSCREEN) < 600 Then Call MsgBox("Easy Publish requires atleast 800 x 600 screen resolution.", vbInformation, "Cannot Continue !"): Call CloseForm(MdiMainMenu): Exit Sub
    DatabaseType = Trim(ReadFromFile("Database Type"))
        If Dir(App.Path & "\Costing", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Costing"
        If Dir(App.Path & "\Database", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Database"
        If Dir(App.Path & "\Export", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Export"
        If Dir(App.Path & "\Imposition", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Imposition"
        If Dir(App.Path & "\Pic", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Pic"
        If Dir(App.Path & "\Report", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Report"
    If DatabaseType = "MS Access" Then
        If Dir(App.Path & "\EasyPublish.ini") = "" Then WriteToFile "Database Path", App.Path & "\Database": WriteToFile "Database Name", ""
        DatabasePath = Trim(ReadFromFile("Database Path"))
    ElseIf DatabaseType = "MS SQL" Then
        If Dir(App.Path & "\EasyPublish.ini") = "" Then WriteToFile "Server Name", "": WriteToFile "Server Password", ""
    End If
    If Decrypted(Trim(ReadFromFile("Server Name")), Decrypt) Then
        ServerName = Decrypt
    End If
    If Decrypted(Trim(ReadFromFile("Server User")), Decrypt) Then
        ServerUser = Decrypt
    End If
    If Decrypted(Trim(ReadFromFile("Server Password")), Decrypt) Then
        ServerPassword = Decrypt
    End If

    If Trim(ReadFromFile("Super User")) <> "EasyPublish" Then Call Command1_Click
    If FileExist(App.Path & "\Icon\EasyPublish.jpeg") Then Developer = "Developed by Easy Info Solutions International Mobile- +91-987-342-2907   Email ID - Easyinfosolutionsi@gmail.com " & Space(150)
            ServerID = Trim(ReadFromFile("Server ID"))
            If Trim(ReadFromFile("Server ID")) = "" Then WriteToFile "Server ID", "E3R82#I12S0#SM2E1#IA2P6#EP000#"
    Do While Trim(ReadFromFile("Server ID")) = "" Or dueDate = "" Or UniqueDate <> "28-SEP-2016" Or DaysLeft <= 0
             If Trim(ReadFromFile("Server ID")) <> "" Then
                     dYear = "20" + Mid(Trim(ReadFromFile("Server ID")), 9, 1) + Mid(Trim(ReadFromFile("Server ID")), 15, 1)
                    dMonth = Mid(Trim(ReadFromFile("Server ID")), 14, 1) + Mid(Trim(ReadFromFile("Server ID")), 20, 1) + Mid(Trim(ReadFromFile("Server ID")), 3, 1)
                     dDay = Mid(Trim(ReadFromFile("Server ID")), 2, 1) + Mid(Trim(ReadFromFile("Server ID")), 8, 1)
                     dueDate = dDay + "-" + dMonth + "-" + dYear
                     
                     dYear = Mid(Trim(ReadFromFile("Server ID")), 5, 1) + Mid(Trim(ReadFromFile("Server ID")), 11, 1) + Mid(Trim(ReadFromFile("Server ID")), 17, 1) + Mid(Trim(ReadFromFile("Server ID")), 23, 1)
                    dMonth = Mid(Trim(ReadFromFile("Server ID")), 10, 1) + Mid(Trim(ReadFromFile("Server ID")), 16, 1) + Mid(Trim(ReadFromFile("Server ID")), 22, 1)
                     dDay = Mid(Trim(ReadFromFile("Server ID")), 21, 1) + Mid(Trim(ReadFromFile("Server ID")), 4, 1)
                     UniqueDate = dDay + "-" + dMonth + "-" + dYear
             Else
                     frmLicenceAgreement.cmdOK.Visible = False: frmLicenceAgreement.Show vbModal: If LaterFlag = True Then Unload Me: If LaterFlag = True Then Exit Sub: LaterFlag = False
             End If
             ServerID = Trim(ReadFromFile("Server ID"))
    
                     DaysLeft = DateDiff("d", Format(Date, "dd-MMM-yyyy"), dueDate)
                    
             If DaysLeft <= 0 Then
                     Call MsgBox("You are using a Demo/Unlicened Version or Subscription of" & Chr(13) & "Easy Publish ERP that is expired." & Chr(13) & "If you would like to purchase or continue to Subcribe, Please contact" & Chr(13) & "Easy Info Solutions International" & Chr(13) & "E-Mail:sales@easyinfosolution.com" & Chr(13) & "Mobile:+91-987-342-2907", vbInformation, App.Title)
                    frmLicenceAgreement.cmdOK.Visible = False: frmLicenceAgreement.Show vbModal: If LaterFlag = True Then Unload Me: If LaterFlag = True Then Exit Sub: LaterFlag = False
             ElseIf DaysLeft <= 30 Then
                    Call MsgBox("You have  " & DaysLeft & " Days Left..." & Chr(13) & "You are using a Demo/Unlicened Version or Subscription of" & Chr(13) & "Easy Publish ERP that will be expired soon." & Chr(13) & "If you would like to purchase or continue to Subcribe, Please contact" & Chr(13) & "Easy Info Solutions International" & Chr(13) & "E-Mail:sales@easyinfosolution.com" & Chr(13) & "Mobile:+91-987-342-2907", vbInformation, App.Title)
             End If
             
    Loop
            If RenewFlag = True Then Call MsgBox("Your Easy Publish ERP Subscription is renewed now " & Chr(13) & " till  :" & dueDate & ". " & Chr(13) & "If you would have any query, Please contact" & Chr(13) & "Easy Info Solutions International" & Chr(13) & "E-Mail:sales@easyinfosolution.com" & Chr(13) & "Mobile:+91-987-342-2907", vbInformation, App.Title): RenewFlag = False
End Sub
Private Sub MDIForm_Resize()
    On Error Resume Next
    Dim client_rect As RECT
    Dim client_hwnd As Long
    If Trim(ReadFromFile("Client ID")) = "Publisher" Then
    picOriginal.Picture = LoadPicture(App.Path & "\Icon\EasyPublish.jpeg")
    ElseIf Trim(ReadFromFile("Client ID")) = "Printer" Then
    picOriginal.Picture = LoadPicture(App.Path & "\Icon\EasyPrint.jpeg")
    End If
    picStretched.Move 0, 0, ScaleWidth, ScaleHeight
    picStretched.PaintPicture picOriginal.Picture, -20, -40, picStretched.ScaleWidth, picStretched.ScaleHeight, -8, -8, picOriginal.ScaleWidth, picOriginal.ScaleHeight
    Picture = picStretched.Image
    client_hwnd = FindWindowEx(Me.hwnd, 0, "MDIClient", vbNullChar)
    GetClientRect client_hwnd, client_rect
    InvalidateRect client_hwnd, client_rect, 1
    If Me.WindowState <> vbMinimized Then Me.WindowState = vbMaximized
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not MnuOpen.Enabled Then MsgBox "           Cannot Quit till You have Company Open." & vbCrLf & "Kindly make sure to Close the Company before Quitting !!!", vbExclamation, "Cannot Close !": Cancel = 1: Exit Sub Else Call CloseForm(MdiMainMenu)
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo ErrorHandler
    Set oHuffman = Nothing
    Set FSO = Nothing
    Set oRegistry = Nothing
    CloseMainConnection
    If Not cnDatabase Is Nothing Then Set cnDatabase = Nothing
    Call AnimateWindow(Me.hwnd, CInt(500), AW_HIDE Or AW_BLEND)
    Exit Sub
ErrorHandler:
End Sub
Private Sub MenuPaperLedger_Click(Index As Integer)
    On Error Resume Next
    FrmItemSelectionList.VchType = Trim(Index)
    Load FrmItemSelectionList
    If Err.Number <> 364 Then FrmItemSelectionList.Show
End Sub
Private Sub mnuExit_Click()
    If MnuClose.Enabled Then mnuClose_Click
    If Forms.Count <= 1 Then Call CloseForm(MdiMainMenu)
End Sub
Private Sub mnuOpen_Click()
    Dim rstCompanyMaster As New ADODB.Recordset
    On Error GoTo OpenError
    Load FrmCompanyList
    FrmCompanyList.Show vbModal
    If CompCode <> "" Then
        BusySystemIndicator True
        CloseMainConnection
        cnDatabase.CursorLocation = adUseClient
        If DatabaseType = "MS SQL" Then
            cnDatabase.CommandTimeout = 300
            ConnectionString = "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=EP" & CompCode & ";Data Source=" & ServerName
        Else
            ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\EasyPublish." & CompCode & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
        End If
        cnDatabase.Open ConnectionString
        'Require Till 31st Aug
        ServerID = Trim(ReadFromFile("Server ID"))
        If Trim(ReadFromFile("Server ID")) = "" Or Mid(Trim(ReadFromFile("Server ID")), 27, 3) <> CompCode And Mid(Trim(ReadFromFile("Server ID")), 30, 1) <> "#" Then WriteToFile "Server ID", "E3G82#I12S0#SA1E1#IU2P6#EP" + CompCode + "#$"
        ServerID = Trim(ReadFromFile("Server ID"))
        Load FrmLogin
        FrmLogin.Show vbModal
        If LoginSuccess Then
            StatusBar1.Panels(3).Text = "User Name : " & Trim(UserName)
            SetMenuOptions (True)
            rstCompanyMaster.Open "SELECT Name,'-Financial Year From '+REPLACE(CONVERT(VARCHAR(11),FinancialYearFrom,106),' ','-')+' To '+REPLACE(CONVERT(VARCHAR(11),FinancialYearTo,106),' ','-') FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
            MdiMainMenu.Caption = Version & " [" & Trim(rstCompanyMaster.Fields("Name").Value) & Trim(rstCompanyMaster.Fields(1).Value) & "]"                       '"Easy Publish  21|Rel 05 | 06.29 Version |Production & Inventory Management System [" & Trim(rstCompanyMaster.Fields("Name").Value) & Trim(rstCompanyMaster.Fields(1).Value) & "]"
            Call CloseRecordset(rstCompanyMaster)
            Exit Sub
        End If
    End If
    CloseMainConnection
    BusySystemIndicator False
    Exit Sub
OpenError:
    If Not rstCompanyMaster Is Nothing Then Set rstCompanyMaster = Nothing
    CloseMainConnection
    BusySystemIndicator False
End Sub
Private Sub mnuClose_Click()
    Dim Form As Form
        For Each Form In Forms
        If Not TypeOf Form Is MDIForm Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form
    If Forms.Count <= 1 Then
        CloseMainConnection
        SetMenuOptions (False)
        MdiMainMenu.Caption = Version
        StatusBar1.Panels(3).Text = ""
    End If
End Sub
Private Sub SetMenuOptions(bVal As Boolean)
    Dim Object As Object
    Dim rstUserChild As New ADODB.Recordset
    On Error GoTo ErrorHandler
    
    MnuOpen.Enabled = Not bVal
    mnuCreate.Enabled = Not bVal
    MnuClose.Enabled = bVal
    mnuEdit.Enabled = bVal
    MnuCompanyChild.Enabled = bVal
    MnuDelete.Enabled = Not bVal
    MnuBackup.Enabled = Not bVal
    MnuRestore.Enabled = Not bVal
    MnuLicenceAgreement.Enabled = True
    MnuUtilities.Enabled = bVal
    mnuProjectManagementParent.Enabled = bVal
    If bVal Then
        rstUserChild.Open "Select [Module] From UserChild Where Code = '" & FixQuote(UserCode) & "' Order by [Module]", cnDatabase, adOpenKeyset, adLockReadOnly
        For Each Object In Me
            If TypeName(Object) = "Menu" Then
                If Object.Tag <> "" Then
                    If UserLevel <> "1" Then
                        rstUserChild.MoveFirst
                        rstUserChild.Find "[Module] = '" & Trim(Object.Tag) & "'"
                        Object.Enabled = IIf(rstUserChild.EOF, False, True)
                        Object.Visible = IIf(rstUserChild.EOF, False, True)
                    Else
                        Object.Visible = True
                        Object.Enabled = True
                    End If
                End If
            End If
        Next
    Else
        MnuMasters.Enabled = bVal
        MnuDisplay.Enabled = bVal
        MnuTransactions.Enabled = bVal
        MnuReports.Enabled = bVal
        mnuProjectManagementParent.Enabled = bVal
    End If
ErrorHandler:
    Call CloseRecordset(rstUserChild)
End Sub
Private Sub MnuYouTube_Click()
           Dim R As Long
              R = ShellExecute(0, "open", "https://www.youtube.com/channel/UCW5RVD8qIBTGzCSRM03U7Cw/featured", 0, 0, 1)
End Sub
Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 4 Or Panel.Index = 5 Then
        On Error Resume Next
        Shell "Control.Exe Date/Time", vbNormalFocus
    End If
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    Static Counter As Integer
    Counter = Counter + 1
    StatusBar1.Panels(2).Text = Left(Developer, Counter)
    If Counter = Len(Developer) Then Counter = 0
    StatusBar1.Panels(4).Text = WeekdayName(Weekday(Date), True, vbSunday) + ", " + MonthName(Month(Date), True) + Str$(Day(Date)) + ", " + Right(Str$(Year(Date)), 2)
    StatusBar1.Panels(5).Text = Left(Time, 8)
    End Sub
Private Sub Timer2_Timer()
    Static T As Long
    T = T + 60000
    If T / 60000 = 60 Then
        mnuBookReceiptBusy_Click
        T = 0
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index <= 17 Then
        If ActiveForm.Toolbar1.Buttons.Item(Button.Index).Enabled Then
            ActiveForm.Toolbar1_ButtonClick ActiveForm.Toolbar1.Buttons.Item(Button.Index)
        End If
    Else
        If Toolbar1.Buttons(1).Enabled Then 'Company Open
            If ActiveForm.Toolbar1.Buttons.Item(Button.Index).Enabled Then
                ActiveForm.Toolbar1_ButtonClick ActiveForm.Toolbar1.Buttons.Item(Button.Index)
            End If
        Else
            mnuExit_Click
        End If
    End If
End Sub
Private Sub oCreate_PercentDone(ByVal Percent As Integer)
    MdiMainMenu.ProgressBar1.Value = Percent
End Sub
Private Sub mnuDelete_Click()
    On Error Resume Next
    Load FrmCompanyList
    If Err.Number <> 364 Then
        FrmCompanyList.Caption = "Select Company To Delete..."
        FrmCompanyList.Show vbModal
        On Error GoTo ErrorHandler
        If CompCode <> "" Then
            Load FrmLogin
            If Err.Number <> 364 Then
                FrmLogin.Show vbModal
                If LoginSuccess Then
                    If UserLevel <> "1" Then
                        Call MsgBox("You don't have authority to Delete a Company !", vbInformation, App.Title)
                        CompCode = ""
                    End If
                End If
            End If
        End If
    End If
    CloseMainConnection
    If CompCode = "" Or (Not LoginSuccess) Then
        Exit Sub
    End If
    If MsgBox("Are you sure to delete the Company?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") <> vbYes Then
        Exit Sub
    End If
    MdiMainMenu.MousePointer = vbHourglass
    FSO.DeleteFile DatabasePath & "\EasyPublish." & CompCode
    Call MsgBox("Successfully deleted the company !", vbInformation, App.Title)
    MdiMainMenu.MousePointer = vbNormal
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to delete the company")
End Sub
Private Sub mnuBackup_Click()
    On Error Resume Next
    Dim strDestination As String
    Load FrmCompanyList
    If Err.Number <> 364 Then
        FrmCompanyList.Caption = "Select Company To Backup..."
        FrmCompanyList.Show vbModal
        On Error GoTo ErrorHandler
        If CompCode <> "" Then
            BrowseFF1.Caption = "Select Destination..."
            BrowseFF1.InitialFolder = App.Path & "\Backup"
            BrowseFF1.IncludeFiles = False
            If BrowseFF1.Browse = True Then
                strDestination = BrowseFF1.SelectedItem.Name
            End If
        End If
    End If
    CloseMainConnection
    If Len(strDestination) = 0 Or CompCode = "" Then Exit Sub
    If Right(strDestination, 1) <> "\" Then strDestination = RTrim(strDestination) & "\"
    strDestination = strDestination & CStr(Format(Date, "yyyymmdd")) & "." & CompCode
    MdiMainMenu.MousePointer = vbHourglass
    ShowProgressInStatusBar True
    If Dir(strDestination) <> "" Then Kill strDestination
    Set oHuffman = New clsHuffman
    Call oHuffman.EncodeFile(DatabasePath & "\EasyPublish." & CompCode, strDestination, False)
    Set oHuffman = Nothing
    ShowProgressInStatusBar False
    MdiMainMenu.MousePointer = vbNormal
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to backup the data of the company")
End Sub
Private Sub oHuffman_Progress(Percent As Integer)
    MdiMainMenu.ProgressBar1.Value = Percent
    DoEvents
End Sub
Private Sub mnuRestore_Click()
    On Error Resume Next
    Dim strSource As String
    Load FrmCompanyList
    If Err.Number <> 364 Then
        FrmCompanyList.Caption = "Select Company To Restore..."
        FrmCompanyList.Show vbModal
        On Error GoTo ErrorHandler
        If CompCode <> "" Then
            BrowseFF1.Caption = "Select Source..."
            BrowseFF1.InitialFolder = App.Path & "\Backup"
            BrowseFF1.IncludeFiles = True
            If BrowseFF1.Browse = True Then
                strSource = BrowseFF1.SelectedItem.Name
            End If
        End If
    End If
    CloseMainConnection
    If Len(strSource) = 0 Or CompCode = "" Then
        Exit Sub
    End If
    If Right(strSource, 3) <> CompCode Then
        DisplayError ("Failed to restore the data of the company")
        Exit Sub
    End If
    MdiMainMenu.MousePointer = vbHourglass
    ShowProgressInStatusBar True
    Set oHuffman = New clsHuffman
    Call oHuffman.DecodeFile(strSource, DatabasePath & "\EasyPublish." & CompCode)
    Set oHuffman = Nothing
    ShowProgressInStatusBar False
    MdiMainMenu.MousePointer = vbNormal
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to restore the data of the company")
End Sub
Private Sub mnuCostSheet_Click()
    On Error Resume Next
    FrmBookPrintOrder.BookPOType = "OP"
    Load FrmBookPrintOrder
    If Err.Number <> 364 Then FrmBookPrintOrder.Show
End Sub
'Private Sub mnuBookDebitNote_Click()
'    On Error Resume Next
'    Load FrmBookDebitNote
'    If Err.Number <> 364 Then FrmBookDebitNote.Show
'End Sub
'Private Sub mnuDayBook_Click()
'    On Error Resume Next
'    Load FrmDayBook
'    If Err.Number <> 364 Then FrmDayBook.Show
'End Sub
Private Sub mnuProductionScheduling_Click(Index As Integer)
    On Error Resume Next
    If Trim(Index) = 1 Then
        FrmProductionScheduling.VchType = Trim(Index)
        FrmProductionScheduling.Caption = "Print Production Scheduling"
        Load FrmProductionScheduling
        If Err.Number <> 364 Then FrmProductionScheduling.Show
    ElseIf Trim(Index) = 2 Then
        FrmProductionScheduling.Caption = "Plate Production Scheduling"
        FrmProductionScheduling.VchType = Trim(Index)
        Load FrmProductionScheduling
        If Err.Number <> 364 Then FrmProductionScheduling.Show
    ElseIf Trim(Index) = 3 Then
        FrmProductionScheduling.Caption = "Paper Cutting Scheduling"
        FrmProductionScheduling.VchType = Trim(Index)
        Load FrmProductionScheduling
        If Err.Number <> 364 Then FrmProductionScheduling.Show
    Else
        FrmProductionSchedule.VchType = Trim(Index)
        Load FrmProductionSchedule
        If Err.Number <> 364 Then FrmProductionSchedule.Show
    End If
End Sub
Private Sub mnuStockLedger_Click(Index As Integer)
    On Error Resume Next
    FrmItemSelectionList.VchType = Trim(Index)
    Load FrmItemSelectionList
    If Err.Number <> 364 Then FrmItemSelectionList.Show
End Sub
Private Sub mnuPendingPaymentRegister_Click()
    On Error Resume Next
    Load FrmPendingPaymentRegister
    If Err.Number <> 364 Then
        FrmPendingPaymentRegister.Show
    End If
End Sub
Private Sub mnuProductionPlanning01_Click()
    On Error Resume Next
    FrmProductionPlanning.OrderType = "M"
    Load FrmProductionPlanning
    If Err.Number <> 364 Then
        FrmProductionPlanning.Show
    End If
End Sub
Private Sub mnuProductionPlanning02_Click()
    On Error Resume Next
    FrmProductionPlanning.OrderType = "S"
    Load FrmProductionPlanning
    If Err.Number <> 364 Then
        FrmProductionPlanning.Show
    End If
End Sub
'Private Sub mnuIOStatusUpdation_Click()
'    Dim oExcel As Object
'    Dim i As Long
'    On Error GoTo ErrorHandler
'    If Not FileExist(App.Path & "\Report\Paper Issue Register (" & CompCode & ").xlsx") Then DisplayError ("Failed to Update the Paper Issue Order(s) Status"):          Exit Sub
'    If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
'        Screen.MousePointer = vbHourglass
'        DoEvents
'        Set oExcel = CreateObject("Excel.Application")
'        oExcel.Workbooks.Open (App.Path & "\Report\Paper Issue Register (" & CompCode & ")")
'        cnDatabase.BeginTrans
'        For i = 5 To 1048576
'            If Trim(oExcel.Application.Cells(i, 16384)) = "" Then Exit For
'            cnDatabase.Execute "UPDATE PaperIOChild SET Narration='" & Trim(oExcel.Application.Cells(i, 10)) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, 16384)) & "' AND Paper='" & oExcel.Application.Cells(i, 16382) & "' AND Account='" & oExcel.Application.Cells(i, 16383) & "'"
'        Next
'        cnDatabase.CommitTrans
'        Call MsgBox("Successfully Updated the Paper Issue Order(s) Status !", vbInformation, App.Title)
'        oExcel.Workbooks.Close
'        Set oExcel = Nothing
'        Screen.MousePointer = vbNormal
'    End If
'    Exit Sub
'ErrorHandler:
'    cnDatabase.RollbackTrans
'    oExcel.Workbooks.Close
'    Set oExcel = Nothing
'    Screen.MousePointer = vbNormal
'    DisplayError ("Failed to Update the Paper Issue Order(s) Status")
'End Sub
Private Sub mnuBookReceiptBusy_Click()
'    If Trim(ReadFromFile("Item Receipt")) = "" Or Trim(ReadFromFile("Item Receipt")) = "N" Then Exit Sub
'    Dim CxnImporter As New ADODB.Connection
'    Dim rstImporter As New ADODB.Recordset
'    Dim DatabaseName, SQL
'    Dim i As Integer
'    On Error GoTo ErrorHandler
'    DatabaseName = Trim(ReadFromFile("Busy Database Name")): i = 0
'    If ServerName = "" Or DatabaseName = "" Then Exit Sub
'    Dim lpBuff As String * 1024
'    GetComputerName lpBuff, Len(lpBuff)
'    Screen.MousePointer = vbHourglass
'    cnDatabase.BeginTrans
'    CxnImporter.CursorLocation = adUseClient
'    cnDatabase.Execute "UPDATE BookPOParent SET QuantityReceived=0"
'    Do While True
'        i = InStr(1, DatabaseName, ",")
'        If CxnImporter.State = adStateOpen Then CxnImporter.Close
'        If i = 0 Then CxnImporter.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=" & Mid(DatabaseName, 1) & ";Data Source=" & ServerName Else CxnImporter.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=" & Mid(DatabaseName, 1, i - 1) & ";Data Source=" & ServerName
'        SQL = "SELECT * FROM (" & _
'                  "SELECT RefCode,LTRIM([No]) As [No],Date,MasterCode1,SUM(ABS(Value1)) As OrderedQuantity,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType IN (2,4) AND RecType=4 AND RefCode=T.RefCode) As ReceivedQuantity FROM Tran3 T WHERE VchType=13 GROUP BY RefCode,No,Date,MasterCode1 " & _
'                  "UNION " & _
'                  "SELECT RefCode,LTRIM(C.OF2) As [No],Date,MasterCode1,SUM(ABS(Value1)) As OrderedQuantity,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType IN (2,4) AND RecType=4 AND RefCode=T.RefCode) As ReceivedQuantity FROM Tran3 T INNER JOIN VchOtherInfo C ON T.VchCode=C.VchCode WHERE VchType=13 AND OF2<>'' GROUP BY RefCode,C.OF2,Date,MasterCode1 " & _
'                  "UNION " & _
'                  "SELECT RefCode,LTRIM(C.OF3) As [No],Date,MasterCode1,SUM(ABS(Value1)) As OrderedQuantity,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType IN (2,4) AND RecType=4 AND RefCode=T.RefCode) As ReceivedQuantity FROM Tran3 T INNER JOIN VchOtherInfo C ON T.VchCode=C.VchCode WHERE VchType=13 AND OF3<>'' GROUP BY RefCode,C.OF3,Date,MasterCode1 " & _
'                  "UNION " & _
'                  "SELECT RefCode,LTRIM(C.OF4) As [No],Date,MasterCode1,SUM(ABS(Value1)) As OrderedQuantity,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType IN (2,4) AND RecType=4 AND RefCode=T.RefCode) As ReceivedQuantity FROM Tran3 T INNER JOIN VchOtherInfo C ON T.VchCode=C.VchCode WHERE VchType=13 AND OF4<>'' GROUP BY RefCode,C.OF4,Date,MasterCode1 " & _
'                  "UNION " & _
'                  "SELECT RefCode,LTRIM(C.OF5) As [No],Date,MasterCode1,SUM(ABS(Value1)) As OrderedQuantity,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType IN (2,4) AND RecType=4 AND RefCode=T.RefCode) As ReceivedQuantity FROM Tran3 T INNER JOIN VchOtherInfo C ON T.VchCode=C.VchCode WHERE VchType=13 AND OF5<>'' GROUP BY RefCode,C.OF5,Date,MasterCode1 " & _
'                  "UNION " & _
'                  "SELECT RefCode,LTRIM(C.OF6) As [No],Date,MasterCode1,SUM(ABS(Value1)) As OrderedQuantity,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType IN (2,4) AND RecType=4 AND RefCode=T.RefCode) As ReceivedQuantity FROM Tran3 T INNER JOIN VchOtherInfo C ON T.VchCode=C.VchCode WHERE VchType=13 AND OF6<>'' GROUP BY RefCode,C.OF6,Date,MasterCode1 " & _
'                  ") As Tbl WHERE ReceivedQuantity>0 ORDER BY CONVERT(DECIMAL,[No])"
'        If rstImporter.State = adStateOpen Then rstImporter.Close
'        rstImporter.Open SQL, CxnImporter, adOpenKeyset, adLockReadOnly        'MasterCode1=BookCode
'        rstImporter.ActiveConnection = Nothing
'        If rstImporter.RecordCount > 0 Then rstImporter.MoveFirst
'        Do While Not rstImporter.EOF
'            MdiMainMenu.StatusBar1.Panels(2).Text = "Updating PO #" & Trim(rstImporter.Fields("No").Value) & " !!!"
'            cnDatabase.Execute "UPDATE BookPOParent P,BookPOChild05 C SET C.Status='D' WHERE P.Code=C.Code AND IIF(LEFT(P.Code,1)='*',MID(TRIM(Name),2),TRIM(Name))='" & Trim(rstImporter.Fields("No").Value) & "' AND LEFT(P.Type,1)<>'O' AND Format(Date,'dd-MMM-yyyy')='" & Format(rstImporter.Fields("Date").Value, "dd-MMM-yyyy") & "'"
'            cnDatabase.Execute "UPDATE BookPOParent P,BookPOChild06 C SET C.Status='D' WHERE P.Code=C.Code AND IIF(LEFT(P.Code,1)='*',MID(TRIM(Name),2),TRIM(Name))='" & Trim(rstImporter.Fields("No").Value) & "' AND LEFT(P.Type,1)<>'O' AND Format(Date,'dd-MMM-yyyy')='" & Format(rstImporter.Fields("Date").Value, "dd-MMM-yyyy") & "'"
'            cnDatabase.Execute "UPDATE BookPOParent SET QuantityReceived=QuantityReceived+" & Val(rstImporter.Fields("ReceivedQuantity").Value) & ",ComputerName='" & Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1) & "' WHERE IIF(LEFT(Code,1)='*',MID(TRIM(Name),2),TRIM(Name))='" & Trim(rstImporter.Fields("No").Value) & "' AND LEFT(Type,1)<>'O' AND Format(Date,'dd-MMM-yyyy')='" & Format(rstImporter.Fields("Date").Value, "dd-MMM-yyyy") & "'"
'            cnDatabase.Execute "UPDATE BookPOParent SET BPODStatus=1,TPODStatus=1,TLODStatus=1,BBODStatus=1 WHERE IIF(LEFT(Code,1)='*',MID(TRIM(Name),2),TRIM(Name))='" & Trim(rstImporter.Fields("No").Value) & "' AND LEFT(Type,1)<>'O' AND Format(Date,'dd-MMM-yyyy')='" & Format(rstImporter.Fields("Date").Value, "dd-MMM-yyyy") & "'"
'            rstImporter.MoveNext
'        Loop
'        'Price Updation
'        If rstImporter.State = adStateOpen Then rstImporter.Close
'        rstImporter.Open "SELECT Alias,D2 FROM Master1 WHERE MasterType=6 AND Alias<>'' AND (LEFT(UPPER(Name),2)<>'Z_' AND LEFT(UPPER(Name),2)<>'Z-') ORDER BY Alias", CxnImporter, adOpenKeyset, adLockReadOnly
'        rstImporter.ActiveConnection = Nothing
'        If rstImporter.RecordCount > 0 Then rstImporter.MoveFirst
'        Do While Not rstImporter.EOF
'            cnDatabase.Execute "UPDATE BookMaster SET Price=" & Val(rstImporter.Fields("D2").Value) & " WHERE LEFT(BusyCode,6)='" & Left(rstImporter.Fields("Alias").Value, 6) & "'"
'            rstImporter.MoveNext
'        Loop
'        If i = 0 Then Exit Do Else DatabaseName = Mid(DatabaseName, i + 1): i = 0
'    Loop
'    MdiMainMenu.StatusBar1.Panels(2).Text = ""
'    cnDatabase.Execute "UPDATE BookPOParent P,BookPOChild08 C SET C.Status='' WHERE P.Code=C.Code AND LEFT(P.Type,1)<>'O' AND C.Status='D'"
'    cnDatabase.Execute "UPDATE BookPOParent P,BookPOChild08 C SET C.Status='D' WHERE P.Code=C.Code AND LEFT(P.Type,1)<>'O' AND (P.QuantityReceived+C.AdjustQuantity>=C.ActualQuantity-INT(C.ActualQuantity*0.2/100) OR INT(C.ActualQuantity*0.2/100)+C.AdjustQuantity>=C.ActualQuantity)"
'    cnDatabase.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code SET P.QuantityReceived=P.QuantityReceived+C.AdjustQuantity"
'    cnDatabase.CommitTrans
'    On Error Resume Next
'    If CxnImporter.State = adStateOpen Then
'        Dim RecordsAffected As Long
'        If rstImporter.State = adStateOpen Then rstImporter.Close
'        rstImporter.Open "SELECT TRIM(P.Name) As VchNo,Date As VchDate,M.Alias As Laminator FROM ((BookPOParent P INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild07 C2 ON P.Code=C2.Code) INNER JOIN AccountMaster M ON P.Laminator=M.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' " & _
'                                              "ORDER BY P.Name", cnDatabase, adOpenKeyset, adLockReadOnly
'        rstImporter.ActiveConnection = Nothing
'        If rstImporter.RecordCount > 0 Then rstImporter.MoveFirst
'        Do While Not rstImporter.EOF
'            MdiMainMenu.StatusBar1.Panels(2).Text = "Updating Alias for PO #" & Trim(rstImporter.Fields("VchNo").Value) & " !!!"
'            CxnImporter.Execute "UPDATE VchOtherInfo SET OF1='" & rstImporter.Fields("Laminator").Value & "' WHERE VchCode IN (SELECT VchCode FROM Tran1 WHERE VchType=13 AND LTRIM(VchNo)='" & rstImporter.Fields("VchNo").Value & "' AND Date='" & Format(rstImporter.Fields("VchDate").Value, "dd-MMM-yyyy") & "')", RecordsAffected
'            If RecordsAffected = 0 Then CxnImporter.Execute "INSERT INTO VchOtherInfo (VchCode,OF1) VALUES ((SELECT VchCode FROM Tran1 WHERE VchType=13 AND LTRIM(VchNo)='" & rstImporter.Fields("VchNo").Value & "' AND Date='" & Format(rstImporter.Fields("VchDate").Value, "dd-MMM-yyyy") & "'),'" & rstImporter.Fields("Laminator").Value & "')"
'            rstImporter.MoveNext
'        Loop
'    End If
'    Call CloseRecordset(rstImporter)
'    Call CloseConnection(CxnImporter)
'    Screen.MousePointer = vbNormal
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description
'    cnDatabase.RollbackTrans
'    Screen.MousePointer = vbNormal
'    DisplayError ("Failed to import the Book Receipts")
'    Call CloseRecordset(rstImporter)
'    Call CloseConnection(CxnImporter)
End Sub
Private Sub mnuPOStatusRegister01_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Order Status Register [Bookwise]") Then
        Dim FrmPrintOrderStatusRegister01 As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegister01.OrderType = "01"
        Load FrmPrintOrderStatusRegister01
        If Err.Number <> 364 Then
            FrmPrintOrderStatusRegister01.Show
        End If
    End If
End Sub
Private Sub mnuPOStatusRegister05_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Order Status Register [Print Orderwise]") Then
        Dim FrmPrintOrderStatusRegister05 As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegister05.OrderType = "02"
        Load FrmPrintOrderStatusRegister05
        If Err.Number <> 364 Then
            FrmPrintOrderStatusRegister05.Show
        End If
    End If
End Sub
Private Sub mnuPendingDNRegister_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Order Status Register [Debit Note]") Then
        Dim FrmPrintOrderStatusRegisterYY As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegisterYY.OrderType = "YY"
        Load FrmPrintOrderStatusRegisterYY
        If Err.Number <> 364 Then FrmPrintOrderStatusRegisterYY.Show
    End If
End Sub
Private Sub mnuPOStatusRegister02_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Order Status Register [Title Printerwise]") Then
        Dim FrmPrintOrderStatusRegister02 As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegister02.OrderType = "06"
        Load FrmPrintOrderStatusRegister02
        If Err.Number <> 364 Then FrmPrintOrderStatusRegister02.Show
    End If
End Sub
Private Sub mnuPOStatusRegister03_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Order Status Register [Book Printerwise]") Then
        Dim FrmPrintOrderStatusRegister03 As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegister03.OrderType = "05"
        Load FrmPrintOrderStatusRegister03
        If Err.Number <> 364 Then
            FrmPrintOrderStatusRegister03.Show
        End If
    End If
End Sub
Private Sub mnuPOStatusRegister04_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Order Status Register [Book Binderwise]") Then
        Dim FrmPrintOrderStatusRegister04 As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegister04.OrderType = "08"
        Load FrmPrintOrderStatusRegister04
        If Err.Number <> 364 Then
            FrmPrintOrderStatusRegister04.Show
        End If
    End If
End Sub
Private Sub mnuPOStatusReg02_Click()
    On Error Resume Next
    Load FrmOutsourceItemSupplierRegister
    If Err.Number <> 364 Then FrmOutsourceItemSupplierRegister.Show
End Sub
Private Sub mnuPOStatusReg0301_Click()
    On Error Resume Next
    If Not IsFormLoaded("Insource Item [Fresh Book] Purchase Order Status Register") Then
        Dim FrmInSourceItem03SupplierRegister As New FrmInsourceItemSupplierRegister
        FrmInSourceItem03SupplierRegister.ItemType = "3"
        Load FrmInSourceItem03SupplierRegister
        If Err.Number <> 364 Then
            FrmInSourceItem03SupplierRegister.Show
        End If
    End If
End Sub
Private Sub mnuPOStatusReg0302_Click()
    On Error Resume Next
    If Not IsFormLoaded("Insource Item [Repair Book] Purchase Order Status Register") Then
        Dim FrmInSourceItem04SupplierRegister As New FrmInsourceItemSupplierRegister
        FrmInSourceItem04SupplierRegister.ItemType = "4"
        Load FrmInSourceItem04SupplierRegister
        If Err.Number <> 364 Then
            FrmInSourceItem04SupplierRegister.Show
        End If
    End If
End Sub
Private Sub mnuPOStatusReg0303_Click()
    On Error Resume Next
    If Not IsFormLoaded("Insource Item [Title] Purchase Order Status Register") Then
        Dim FrmInSourceItem05SupplierRegister As New FrmInsourceItemSupplierRegister
        FrmInSourceItem05SupplierRegister.ItemType = "5"
        Load FrmInSourceItem05SupplierRegister
        If Err.Number <> 364 Then
            FrmInSourceItem05SupplierRegister.Show
        End If
    End If
End Sub
Private Sub mnuPaperIssueRegister_Click()
    On Error Resume Next
    Load FrmPaperIssueRegister
    If Err.Number <> 364 Then FrmPaperIssueRegister.Show
End Sub
Private Sub mnuPaperStockRegister_Click()
    On Error Resume Next
    Load FrmPaperStockRegister
    If Err.Number <> 364 Then FrmPaperStockRegister.Show
End Sub
Private Sub mnuMaterialStockRegister01_Click()
    On Error Resume Next
    If Not IsFormLoaded("Material Stock Register [Binderwise/Bookwise/Itemwise]") Then
        Dim FrmMaterialStockRegister01 As New FrmMaterialStockRegister
        FrmMaterialStockRegister01.ReportType = "1"
        Load FrmMaterialStockRegister01
        If Err.Number <> 364 Then
            FrmMaterialStockRegister01.Show
        End If
    End If
End Sub
Private Sub mnuMaterialStockRegister02_Click()
    On Error Resume Next
    If Not IsFormLoaded("Material Stock Register [Binderwise/Itemwise]") Then
        Dim FrmMaterialStockRegister02 As New FrmMaterialStockRegister
        FrmMaterialStockRegister02.ReportType = "2"
        Load FrmMaterialStockRegister02
        If Err.Number <> 364 Then
            FrmMaterialStockRegister02.Show
        End If
    End If
End Sub
Private Sub mnuPOLedgerParty_Click(Index As Integer)
    On Error Resume Next
    FrmBillRegister.VchCodeType = Trim(Index)
    Load FrmBillRegister
    If Err.Number <> 364 Then FrmBillRegister.Show
End Sub
Private Sub mnuSOLedgerParty_Click(Index As Integer)
    On Error Resume Next
    FrmBillRegister.VchCodeType = Trim(Index)
    Load FrmBillRegister
    If Err.Number <> 364 Then FrmBillRegister.Show
End Sub
'Private Sub mnuSOLedgerPartyDetail_Click()
'    On Error Resume Next
'    FrmBillRegister.VchCodeType = "PO1"
'    Load FrmBillRegister
'    If Err.Number <> 364 Then FrmBillRegister.Show
'End Sub
'Private Sub mnuSOLedgerPartySummrised_Click()
'    On Error Resume Next
'    FrmBillRegister.VchCodeType = "PO2"
'    Load FrmBillRegister
'    If Err.Number <> 364 Then FrmBillRegister.Show
'End Sub
Private Sub mnuProductionSchedule_Click()
    On Error Resume Next
    Load FrmProductionSchedule
    If Err.Number <> 364 Then FrmProductionSchedule.Show
End Sub
Private Sub mnuPurchaseSaleOrderParty_Click(Index As Integer)
    On Error Resume Next
    FrmOrderProcessing.VchCodeType = Trim(Index)
    Load FrmOrderProcessing
    If Err.Number <> 364 Then FrmOrderProcessing.Show
End Sub
Private Sub MnuPurchaseSaleOrderPartyPO_Click(Index As Integer)
    On Error Resume Next
    FrmOrderProcessing.VchCodeType = Trim(Index)
    Load FrmOrderProcessing
    If Err.Number <> 364 Then FrmOrderProcessing.Show
End Sub
Private Sub MnuPurchaseSaleOrderPartySO_Click(Index As Integer)
    On Error Resume Next
    FrmOrderProcessing.VchCodeType = Trim(Index)
    Load FrmOrderProcessing
    If Err.Number <> 364 Then FrmOrderProcessing.Show
End Sub
Private Sub MnuItemIssueReceipt_Click(Index As Integer)
    On Error Resume Next
    FrmOrderProcessing.VchCodeType = Trim(Index)
    Load FrmOrderProcessing
    If Err.Number <> 364 Then FrmOrderProcessing.Show
End Sub
'Private Sub mnuSaleOrderPartyDetailed_Click()
'    On Error Resume Next
'    FrmOrderProcessing.VchCodeType = "P1"
'    Load FrmOrderProcessing
'    If Err.Number <> 364 Then FrmOrderProcessing.Show
'End Sub
'Private Sub mnuSaleOrderPartySummarised_Click()
'    On Error Resume Next
'    FrmOrderProcessing.VchCodeType = "P2"
'    Load FrmOrderProcessing
'    If Err.Number <> 364 Then FrmOrderProcessing.Show
'End Sub
'Private Sub mnuPurchaseOrderPartyDetailed_Click()
'    On Error Resume Next
'    FrmOrderProcessing.VchCodeType = "S1"
'    Load FrmOrderProcessing
'    If Err.Number <> 364 Then FrmOrderProcessing.Show
'End Sub
'Private Sub mnuPurchaseOrderPartySummarised_Click()
'    On Error Resume Next
'    FrmOrderProcessing.VchCodeType = "S2"
'    Load FrmOrderProcessing
'    If Err.Number <> 364 Then FrmOrderProcessing.Show
'End Sub
Private Sub mnuBookPOPrintUtility_Click(Index As Integer)
    On Error Resume Next
    FrmBookPOPrintUtility.VchCode = Trim(Index)
    Load FrmBookPOPrintUtility
    If Err.Number <> 364 Then FrmBookPOPrintUtility.Show
End Sub
Private Sub mnuPaperPOPrintUtility_Click()
    On Error Resume Next
    Load FrmPaperPOPrintUtility
    If Err.Number <> 364 Then FrmPaperPOPrintUtility.Show
End Sub
Private Sub mnuOpBal_Click()
    Dim oExcel As Object
    Dim i As Long, Cnt As Long
    Dim rstPaperOpBal As New ADODB.Recordset
    Dim rstCompanyMaster As New ADODB.Recordset
    On Error Resume Next
    
    If Not FileExist(App.Path & "\Template\Opening Balance.xlsx") Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If rstPaperOpBal.State = adStateOpen Then rstPaperOpBal.Close
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperOpBal.Open "SELECT M2.Name As GodownName,TRIM(M1.Name)+' (UOM : '+TRIM(U.Name)+')' As PaperName,[Weight/Unit],C.OpBalOther,C.OpBalSheets,C.OpBalTat,U.Value1 As SPU FROM ((PaperChild C INNER JOIN PaperMaster M1 ON M1.Code=C.Code) INNER JOIN AccountMaster M2 ON M2.Code=C.Account) INNER JOIN GeneralMaster U ON M1.UOM=U.Code ORDER BY M2.Name,M1.Name", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstPaperOpBal.RecordCount = 0 Then
        Screen.MousePointer = vbNormal
        On Error GoTo 0
        Exit Sub
    End If
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Opening Balance")
    oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Opening Balance (" & CompCode & ")")
    oExcel.DisplayAlerts = True
    oExcel.Sheets("Sheet1").Select
    oExcel.Visible = False
    oExcel.Cells(1, 1).Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Cells(2, 1).Value = "Opening Balance As On [" & Format(FinancialYearFrom, "dd-mm-yyyy") & "]"
    i = 4
    Cnt = 1
    Do While Not rstPaperOpBal.EOF
        oExcel.Application.Cells(i, 1).Value = Trim(rstPaperOpBal.Fields("GodownName").Value)
        oExcel.Application.Cells(i, 2).Value = Trim(rstPaperOpBal.Fields("PaperName").Value)
        oExcel.Application.Cells(i, 3).Value = Val(rstPaperOpBal.Fields("OpBalOther").Value)
        oExcel.Application.Cells(i, 4).Value = Val(rstPaperOpBal.Fields("Weight/Unit").Value)
        oExcel.Application.Cells(i, 5).Value = Val(rstPaperOpBal.Fields("OpBalSheets").Value)
        oExcel.Application.Cells(i, 6).Value = Round(Val(rstPaperOpBal.Fields("OpBalSheets").Value) / Val(rstPaperOpBal.Fields("SPU").Value), 3)
        oExcel.Application.Cells(i, 7).Value = Val(rstPaperOpBal.Fields("TatOpBal").Value)
        oExcel.Application.Cells(i, 9).Value = Val(oExcel.Application.Cells(i, 6).Value) * Val(oExcel.Application.Cells(i, 4).Value)
        Cnt = Cnt + 1
        i = i + 1
        rstPaperOpBal.MoveNext
    Loop
    oExcel.Sheets("Sheet1").Activate
    oExcel.Columns("A:J").EntireColumn.AutoFit
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    oExcel.Range("A1").Activate
    oExcel.Visible = True
    Set oExcel = Nothing
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstPaperOpBal)
    On Error GoTo 0
End Sub
Private Sub mnuBookList_Click(Index As Integer)
    On Error Resume Next
    FrmBookList.VchCodeType = Trim(Index)
    Load FrmBookList
    If Err.Number <> 364 Then
        FrmBookList.Show
    End If
End Sub
Private Sub mnuCorrectionList_Click()
    On Error Resume Next
    Load FrmCorrectionList
    If Err.Number <> 364 Then FrmCorrectionList.Show
End Sub
Private Sub mnuImportBal01_Click()  'Print Order
    Dim CxnImporter As New ADODB.Connection
    Dim rstCompanyMaster As New ADODB.Recordset
    Dim rstImporter00 As New ADODB.Recordset
'    Dim rstImporter05 As New ADODB.Recordset
'    Dim rstImporter06 As New ADODB.Recordset
'    Dim rstImporter07 As New ADODB.Recordset
    Dim rstImporter08 As New ADODB.Recordset
    Dim i As Integer, K As Integer
    Dim SQL As String
    On Error GoTo ErrorHandler
    BusySystemIndicator True
    rstCompanyMaster.Open "Select CreatedFrom From CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstCompanyMaster.Fields("CreatedFrom").Value <> "" Then
        If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
            rstCompanyMaster.ActiveConnection = Nothing
            CxnImporter.CursorLocation = adUseClient
            CxnImporter.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\EasyPublish." & rstCompanyMaster.Fields("CreatedFrom").Value & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
            rstImporter00.Open "SELECT P.* FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE Type='F' AND LEFT(P.Code,1)<>'*' AND (C.ActualQuantity-P.QuantityReceived)>0 AND C.Status NOT IN ('E','D','W') ORDER BY P.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter00.ActiveConnection = Nothing
'            rstImporter05.Open "SELECT C05.* FROM (BookPOParent P INNER JOIN BookPOChild08 C08 ON P.Code=C08.Code) INNER JOIN BookPOChild05 C05 ON P.Code=C05.Code WHERE LEFT(Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND (C08.ActualQuantity-P.QuantityReceived)>0 AND C08.Status NOT IN ('E','D','W') ORDER BY P.Code", CxnImporter, adOpenKeyset, adLockReadOnly
'            rstImporter05.ActiveConnection = Nothing
'            rstImporter06.Open "SELECT C06.* FROM (BookPOParent P INNER JOIN BookPOChild08 C08 ON P.Code=C08.Code) INNER JOIN BookPOChild06 C06 ON P.Code=C06.Code WHERE LEFT(Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND (C08.ActualQuantity-P.QuantityReceived)>0 AND C08.Status NOT IN ('E','D','W') ORDER BY P.Code", CxnImporter, adOpenKeyset, adLockReadOnly
'            rstImporter06.ActiveConnection = Nothing
'            rstImporter07.Open "SELECT C07.* FROM (BookPOParent P INNER JOIN BookPOChild08 C08 ON P.Code=C08.Code) INNER JOIN BookPOChild07 C07 ON P.Code=C07.Code WHERE LEFT(Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND (C08.ActualQuantity-P.QuantityReceived)>0 AND C08.Status NOT IN ('E','D','W') ORDER BY P.Code", CxnImporter, adOpenKeyset, adLockReadOnly
'            rstImporter07.ActiveConnection = Nothing
            rstImporter08.Open "SELECT C.* FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE Type='F' AND LEFT(P.Code,1)<>'*' AND (C.ActualQuantity-P.QuantityReceived)>0 AND C.Status NOT IN ('E','D','W') ORDER BY P.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter08.ActiveConnection = Nothing
            cnDatabase.BeginTrans
'            cnDatabase.Execute "DELETE FROM BookPOParent WHERE LEFT(Code,1)='*'"
'            cnDatabase.Execute "DELETE FROM BookPOChild05 WHERE LEFT(Code,1)='*'"
'            cnDatabase.Execute "DELETE FROM BookPOChild06 WHERE LEFT(Code,1)='*'"
'            cnDatabase.Execute "DELETE FROM BookPOChild07 WHERE LEFT(Code,1)='*'"
'            cnDatabase.Execute "DELETE FROM BookPOChild08 WHERE LEFT(Code,1)='*'"
            Do While Not rstImporter00.EOF
                K = K + 1
                SQL = "INSERT INTO BookPOParent VALUES ('" + rstImporter00.Fields(0).Value + "','" + Pad(Trim(Str(K)), Space(1), 10, "L") + "',#" & Format(FinancialYearFrom, "dd-MMM-yyyy") & "#,"
                For i = 3 To rstImporter00.Fields.Count - 1
                    If IsNull(rstImporter00.Fields(i).Value) Then
                        SQL = SQL + "Null,"
                    ElseIf rstImporter00.Fields(i).Type = adVarWChar Then
                        SQL = SQL + "'" + rstImporter00.Fields(i).Value + "',"
                    ElseIf rstImporter00.Fields(i).Type = adDate Then
                        SQL = SQL + "#" + Format(rstImporter00.Fields(i).Value, "mm-dd-yyyy") + "#,"
                    ElseIf rstImporter00.Fields(i).Type = adNumeric Then
                        SQL = SQL + Trim(Str(rstImporter00.Fields(i).Value)) + ","
                    ElseIf rstImporter00.Fields(i).Type = adBoolean Then
                        SQL = SQL + Trim(Str(rstImporter00.Fields(i).Value)) + ","
                    End If
                Next
                SQL = Left(SQL, Len(SQL) - 1)
                SQL = SQL + ")"
                cnDatabase.Execute SQL
                rstImporter00.MoveNext
            Loop
'            Do While Not rstImporter05.EOF
'                SQL = "INSERT INTO BookPOChild05 VALUES ('*" + Right(rstImporter05.Fields(0).Value, 5) + "',"
'                For i = 1 To rstImporter05.Fields.Count - 1
'                    If IsNull(rstImporter05.Fields(i).Value) Then
'                        SQL = SQL + "Null,"
'                    ElseIf rstImporter05.Fields(i).Type = adVarWChar Then
'                        SQL = SQL + "'" + rstImporter05.Fields(i).Value + "',"
'                    ElseIf rstImporter05.Fields(i).Type = adDate Then
'                        SQL = SQL + "#" + Format(rstImporter05.Fields(i).Value, "mm-dd-yyyy") + "#,"
'                    ElseIf rstImporter05.Fields(i).Type = adNumeric Then
'                        SQL = SQL + Trim(Str(rstImporter05.Fields(i).Value)) + ","
'                    End If
'                Next
'                SQL = Left(SQL, Len(SQL) - 1)
'                SQL = SQL + ")"
'                cnDatabase.Execute SQL
'                rstImporter05.MoveNext
'            Loop
'            Do While Not rstImporter06.EOF
'                SQL = "INSERT INTO BookPOChild06 VALUES ('*" + Right(rstImporter06.Fields(0).Value, 5) + "',"
'                For i = 1 To rstImporter06.Fields.Count - 1
'                    If IsNull(rstImporter06.Fields(i).Value) Then
'                        SQL = SQL + "Null,"
'                    ElseIf rstImporter06.Fields(i).Type = adVarWChar Then
'                        SQL = SQL + "'" + rstImporter06.Fields(i).Value + "',"
'                    ElseIf rstImporter06.Fields(i).Type = adDate Then
'                        SQL = SQL + "#" + Format(rstImporter06.Fields(i).Value, "mm-dd-yyyy") + "#,"
'                    ElseIf rstImporter06.Fields(i).Type = adNumeric Then
'                        SQL = SQL + Trim(Str(rstImporter06.Fields(i).Value)) + ","
'                    End If
'                Next
'                SQL = Left(SQL, Len(SQL) - 1)
'                SQL = SQL + ")"
'                cnDatabase.Execute SQL
'                rstImporter06.MoveNext
'            Loop
'            Do While Not rstImporter07.EOF
'                SQL = "INSERT INTO BookPOChild07 VALUES ('*" + Right(rstImporter07.Fields(0).Value, 5) + "',"
'                For i = 1 To rstImporter07.Fields.Count - 1
'                    If IsNull(rstImporter07.Fields(i).Value) Then
'                        SQL = SQL + "Null,"
'                    ElseIf rstImporter07.Fields(i).Type = adVarWChar Then
'                        SQL = SQL + "'" + rstImporter07.Fields(i).Value + "',"
'                    ElseIf rstImporter07.Fields(i).Type = adDate Then
'                        SQL = SQL + "#" + Format(rstImporter07.Fields(i).Value, "mm-dd-yyyy") + "#,"
'                    ElseIf rstImporter07.Fields(i).Type = adNumeric Then
'                        SQL = SQL + Trim(Str(rstImporter07.Fields(i).Value)) + ","
'                    End If
'                Next
'                SQL = Left(SQL, Len(SQL) - 1)
'                SQL = SQL + ")"
'                cnDatabase.Execute SQL
'                rstImporter07.MoveNext
'            Loop
            Do While Not rstImporter08.EOF
                SQL = "INSERT INTO BookPOChild08 VALUES ('" + rstImporter08.Fields(0).Value + "',#" & Format(FinancialYearFrom, "dd-MMM-yyyy") & "#,#" & Format(FinancialYearFrom + 6, "dd-MMM-yyyy") & "#,"
                For i = 3 To rstImporter08.Fields.Count - 1
                    If IsNull(rstImporter08.Fields(i).Value) Then
                        SQL = SQL + "Null,"
                    ElseIf rstImporter08.Fields(i).Type = adVarWChar Then
                        SQL = SQL + "'" + rstImporter08.Fields(i).Value + "',"
                    ElseIf rstImporter08.Fields(i).Type = adDate Then
                        SQL = SQL + "#" + Format(rstImporter08.Fields(i).Value, "mm-dd-yyyy") + "#,"
                    ElseIf rstImporter08.Fields(i).Type = adNumeric Then
                        SQL = SQL + Trim(Str(rstImporter08.Fields(i).Value)) + ","
                    End If
                Next
                SQL = Left(SQL, Len(SQL) - 1)
                SQL = SQL + ")"
                cnDatabase.Execute SQL
                rstImporter08.MoveNext
            Loop
            cnDatabase.CommitTrans
            Call MsgBox("Successfully imported the Balances !", vbInformation, App.Title)
        End If
    Else
        Call MsgBox("Nothing To Import !", vbInformation, App.Title)
    End If
    Call CloseRecordset(rstImporter00)
'    Call CloseRecordset(rstImporter05)
'    Call CloseRecordset(rstImporter06)
'    Call CloseRecordset(rstImporter07)
    Call CloseRecordset(rstImporter08)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    If CxnImporter.State = adStateOpen Then cnDatabase.RollbackTrans
    BusySystemIndicator False
    DisplayError ("Failed to import the Balances")
    Call CloseRecordset(rstImporter00)
'    Call CloseRecordset(rstImporter05)
'    Call CloseRecordset(rstImporter06)
'    Call CloseRecordset(rstImporter07)
    Call CloseRecordset(rstImporter08)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
End Sub
Private Sub mnuImportBal02_Click()
    Dim CxnImporter As New ADODB.Connection
    Dim rstCompanyMaster As New ADODB.Recordset
    Dim rstImporter As New ADODB.Recordset
    Dim ClBal As Double
    On Error GoTo ErrorHandler
    BusySystemIndicator True
    rstCompanyMaster.Open "Select CreatedFrom From CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstCompanyMaster.Fields("CreatedFrom").Value <> "" Then
        If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
            CxnImporter.CursorLocation = adUseClient
            CxnImporter.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\EasyPublish." & rstCompanyMaster.Fields("CreatedFrom").Value & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
            Dim Tbl As String
            Tbl = "SELECT Code As Paper,Account FROM PaperChild WHERE Code<>'' AND Account<>'' UNION " & _
                       "SELECT Paper,Account FROM PaperIOChild WHERE Paper<>'' AND Account<>'' UNION " & _
                       "SELECT Item As Paper,Account FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item<>'' AND Account<>'' UNION " & _
                       "SELECT Paper,AccountFrom As Account FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper<>'' AND AccountFrom<>'' UNION " & _
                       "SELECT Paper,AccountTo As Account FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper<>'' AND AccountTo<>'' UNION " & _
                       "SELECT Paper1 As Paper,BookPrinter As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND PaperAmount1=0 AND Paper1<>'' AND BookPrinter<>'' UNION " & _
                       "SELECT Paper2 As Paper,BookPrinter As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND PaperAmount2=0 AND Paper2<>'' AND BookPrinter<>'' UNION " & _
                       "SELECT Paper4 As Paper,BookPrinter As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND PaperAmount4=0 AND Paper4<>'' AND BookPrinter<>'' UNION " & _
                       "SELECT Paper,TitlePrinter As Account FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND PaperAmount=0 AND Paper<>'' AND TitlePrinter<>'' UNION " & _
                       "SELECT Paper,TitlePrinter As Account FROM BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND PaperAmount=0 AND Paper<>'' AND TitlePrinter<>'' UNION " & _
                       "SELECT Item As Paper,Binder As Account FROM (BookPOParent P INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0801 C2 ON C1.Code=C2.Code WHERE C2.Category='2' AND LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Item<>'' AND Binder<>'' UNION " & _
                       "SELECT Paper,Account FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE Paper<>'' AND Account<>''"
            rstImporter.Open "SELECT Account,Paper," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(OpBalSheets)),0,SUM(OpBalSheets)) FROM PaperChild Where Code=T.Paper And Account=T.Account),0) As OpBal," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(QuantitySheets)),0,SUM(QuantitySheets)) FROM PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code WHERE Paper=T.Paper AND Account=T.Account),0) As IN1," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(Quantity)),0,SUM(FIX(Quantity)*Val(M3.Value1)+(Quantity-FIX(Quantity))*1000)) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Quantity>=0 AND Item=T.Paper AND Account=T.Account),0) As IN2," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(QuantitySheets)),0,SUM(QuantitySheets)) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=T.Paper AND AccountTo=T.Account),0) As IN3," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(Quantity)),0,SUM(FIX(Quantity)*Val(M3.Value1)+(Quantity-FIX(Quantity))*1000)) FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE Paper=T.Paper AND Account=T.Account AND Quantity>=0),0) As IN4," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(Quantity)),0,ABS(SUM(FIX(Quantity)*Val(M3.Value1)+(Quantity-FIX(Quantity))*1000))) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Quantity<0 AND Item=T.Paper AND Account=T.Account),0) As OUT1," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(QuantitySheets)),0,SUM(QuantitySheets)) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=T.Paper AND AccountFrom=T.Account),0) As OUT2," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(Round(C.TotalConsumption,0))),0,SUM(Round(C.TotalConsumption,0))) FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Category='2' AND Item=T.Paper AND Vendor=T.Account),0) As OUT3," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(PaperConsumptionSheets)),0,SUM(PaperConsumptionSheets)) FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND PaperAmount=0 AND Paper=T.Paper AND TitlePrinter=T.Account),0) As OUT4," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(PaperConsumptionSheets1)),0,SUM(PaperConsumptionSheets1)) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND PaperAmount1=0 AND Paper1=T.Paper AND BookPrinter=T.Account),0) As OUT5," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(PaperConsumptionSheets2)),0,SUM(PaperConsumptionSheets2)) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND PaperAmount2=0 AND Paper2=T.Paper AND BookPrinter=T.Account),0) As OUT6," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(PaperConsumptionSheets4)),0,SUM(PaperConsumptionSheets4)) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND PaperAmount4=0 AND Paper4=T.Paper AND BookPrinter=T.Account),0) As OUT7 ," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(Quantity)),0,ABS(SUM(FIX(Quantity)*Val(M3.Value1)+(Quantity-FIX(Quantity))*1000))) FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE Paper=T.Paper AND Account=T.Account AND Quantity<0),0) As OUT8," & _
                        "FORMAT((SELECT IIF(ISNULL(SUM(PaperConsumptionSheets)),0,SUM(PaperConsumptionSheets)) FROM BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND PaperAmount=0 AND Paper=T.Paper AND TitlePrinter=T.Account),0) As OUT9 " & _
                        "FROM (" & Tbl & ") As T INNER JOIN (PaperMaster M1 INNER JOIN GeneralMaster M3 ON M1.UOM=M3.Code) ON T.Paper=M1.Code WHERE T.Paper='000708' ORDER BY Account,Paper", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter.ActiveConnection = Nothing
            rstCompanyMaster.ActiveConnection = Nothing
            cnDatabase.BeginTrans
            cnDatabase.Execute "DELETE FROM PaperChild WHERE Imported = 'Y'"
            Do While Not rstImporter.EOF
                ClBal = Val(CheckNull(rstImporter.Fields("OpBal").Value)) + Val(CheckNull(rstImporter.Fields("IN1").Value)) + Val(CheckNull(rstImporter.Fields("IN2").Value)) + Val(CheckNull(rstImporter.Fields("IN3").Value)) + Val(CheckNull(rstImporter.Fields("IN4").Value)) - Val(CheckNull(rstImporter.Fields("OUT1").Value)) - Val(CheckNull(rstImporter.Fields("OUT2").Value)) - Val(CheckNull(rstImporter.Fields("OUT3").Value)) - Val(CheckNull(rstImporter.Fields("OUT4").Value)) - Val(CheckNull(rstImporter.Fields("OUT5").Value)) - Val(CheckNull(rstImporter.Fields("OUT6").Value)) - Val(CheckNull(rstImporter.Fields("OUT7").Value)) - Val(CheckNull(rstImporter.Fields("OUT8").Value)) - Val(CheckNull(rstImporter.Fields("OUT9").Value))
                If ClBal <> 0 Then cnDatabase.Execute "INSERT INTO PaperChild VALUES ('" & rstImporter.Fields("Paper").Value & "','" & rstImporter.Fields("Account").Value & "'," & CLng(Fix(ClBal / 500)) + ((ClBal Mod 500) / 1000) & "," & ClBal & ",0,'Y')"
                rstImporter.MoveNext
            Loop
            cnDatabase.CommitTrans
            Call MsgBox("Successfully imported the Balances !", vbInformation, App.Title)
        End If
    Else
        Call MsgBox("Nothing To Import !", vbInformation, App.Title)
    End If
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    If CxnImporter.State = adStateOpen Then cnDatabase.RollbackTrans
    BusySystemIndicator False
    DisplayError ("Failed to import the Balances")
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
End Sub
Private Sub mnuImportBal03_Click()
    Dim CxnImporter As New ADODB.Connection
    Dim rstCompanyMaster As New ADODB.Recordset
    Dim rstImporter As New ADODB.Recordset
    Dim ClBal As Double
    On Error GoTo ErrorHandler
    
    BusySystemIndicator True
    rstCompanyMaster.Open "Select CreatedFrom From CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstCompanyMaster.Fields("CreatedFrom").Value <> "" Then
        If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
            CxnImporter.CursorLocation = adUseClient
            CxnImporter.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\EasyPublish." & rstCompanyMaster.Fields("CreatedFrom").Value & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
            'Outsource Items
            rstImporter.Open "SELECT DISTINCT A.Code,O.Code," & _
                                         "FORMAT((SELECT SUM(OpBal) FROM AccountChild0801 Where Category+Item='1'+O.Code AND Code=A.Code),0) As OpBal," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialIOParent M INNER JOIN MaterialIOChild I ON M.Code=I.Code WHERE I.Category+I.Item='1'+O.Code AND I.Godown=A.Code),0) As IN1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='1'+O.Code AND M.Account=A.Code AND I.Quantity>=0),0) As IN2," & _
                                         "FORMAT((SELECT SUM(ABS(Quantity)) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='1'+O.Code AND M.Account=A.Code AND I.Quantity<0),0) As OUT1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='1'+O.Code AND M.AccountFrom=A.Code),0) As OUT2," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='1'+O.Code AND M.AccountTo=A.Code),0) As IN3," & _
                                         "FORMAT((SELECT SUM(Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code)) FROM BookPOParent M INNER JOIN BookPOChild0801 I ON M.Code=I.Code WHERE LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND I.Category+I.Item='1'+O.Code AND M.Binder=A.Code),0) As OUT3 " & _
                                         "FROM OutsourceItemMaster O,AccountMaster A WHERE A.Type In ('08','09') ORDER BY A.Code,O.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter.ActiveConnection = Nothing
            rstCompanyMaster.ActiveConnection = Nothing
            cnDatabase.BeginTrans
            cnDatabase.Execute "Delete From AccountChild0801 Where Imported = 'Y'"
            Do While Not rstImporter.EOF
                ClBal = Val(CheckNull(rstImporter.Fields("OpBal").Value)) + Val(CheckNull(rstImporter.Fields("IN1").Value)) + Val(CheckNull(rstImporter.Fields("IN2").Value)) + Val(CheckNull(rstImporter.Fields("IN3").Value)) - Val(CheckNull(rstImporter.Fields("OUT1").Value)) - Val(CheckNull(rstImporter.Fields("OUT2").Value)) - Val(CheckNull(rstImporter.Fields("OUT3").Value))
                If ClBal <> 0 Then
                    cnDatabase.Execute "Insert Into AccountChild0801 Values ('" & rstImporter.Fields("A.Code").Value & "','1','" & rstImporter.Fields("O.Code").Value & "'," & ClBal & ",'Y')"
                End If
                rstImporter.MoveNext
            Loop
            If rstImporter.State = adStateOpen Then rstImporter.Close
            'Fresh Books
            rstImporter.Open "SELECT A.Code,O.Code," & _
                                         "FORMAT((SELECT SUM(OpBal) FROM AccountChild0801 Where Category+Item='3'+O.Code AND Code=A.Code),0) As OpBal," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialIOParent M INNER JOIN MaterialIOChild I ON M.Code=I.Code WHERE I.Category+I.Item='3'+O.Code AND I.Godown=A.Code),0) As IN1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='3'+O.Code AND M.Account=A.Code AND I.Quantity>=0),0) As IN2," & _
                                         "FORMAT((SELECT SUM(ABS(Quantity)) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='3'+O.Code AND M.Account=A.Code AND I.Quantity<0),0) As OUT1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='3'+O.Code AND M.AccountFrom=A.Code),0) As OUT2," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='3'+O.Code AND M.AccountTo=A.Code),0) As IN3," & _
                                         "FORMAT((SELECT SUM(Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code)) FROM BookPOParent M INNER JOIN BookPOChild0801 I ON M.Code=I.Code WHERE LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND I.Category+I.Item='3'+O.Code AND M.Binder=A.Code),0) As OUT3 " & _
                                         "FROM BookMaster O,AccountMaster A WHERE O.Board='000000' AND A.Type In ('08','09') ORDER BY A.Code,O.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter.ActiveConnection = Nothing
            Do While Not rstImporter.EOF
                ClBal = Val(CheckNull(rstImporter.Fields("OpBal").Value)) + Val(CheckNull(rstImporter.Fields("IN1").Value)) + Val(CheckNull(rstImporter.Fields("IN2").Value)) + Val(CheckNull(rstImporter.Fields("IN3").Value)) - Val(CheckNull(rstImporter.Fields("OUT1").Value)) - Val(CheckNull(rstImporter.Fields("OUT2").Value)) - Val(CheckNull(rstImporter.Fields("OUT3").Value))
                If ClBal <> 0 Then
                    cnDatabase.Execute "Insert Into AccountChild0801 Values ('" & rstImporter.Fields("A.Code").Value & "','3','" & rstImporter.Fields("O.Code").Value & "'," & ClBal & ",'Y')"
                End If
                rstImporter.MoveNext
            Loop
            If rstImporter.State = adStateOpen Then rstImporter.Close
            'Repair Books
            rstImporter.Open "SELECT A.Code,O.Code," & _
                                         "FORMAT((SELECT SUM(OpBal) FROM AccountChild0801 Where Category+Item='4'+O.Code AND Code=A.Code),0) As OpBal," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialIOParent M INNER JOIN MaterialIOChild I ON M.Code=I.Code WHERE I.Category+I.Item='4'+O.Code AND I.Godown=A.Code),0) As IN1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='4'+O.Code AND M.Account=A.Code AND I.Quantity>=0),0) As IN2," & _
                                         "FORMAT((SELECT SUM(ABS(Quantity)) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='4'+O.Code AND M.Account=A.Code AND I.Quantity<0),0) As OUT1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='4'+O.Code AND M.AccountFrom=A.Code),0) As OUT2," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='4'+O.Code AND M.AccountTo=A.Code),0) As IN3," & _
                                         "FORMAT((SELECT SUM(Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code)) FROM BookPOParent M INNER JOIN BookPOChild0801 I ON M.Code=I.Code WHERE LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND I.Category+I.Item='4'+O.Code AND M.Binder=A.Code),0) As OUT3 " & _
                                         "FROM BookMaster O,AccountMaster A WHERE O.Type='R' AND A.Type In ('08','09') ORDER BY A.Code,O.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter.ActiveConnection = Nothing
            Do While Not rstImporter.EOF
                ClBal = Val(CheckNull(rstImporter.Fields("OpBal").Value)) + Val(CheckNull(rstImporter.Fields("IN1").Value)) + Val(CheckNull(rstImporter.Fields("IN2").Value)) + Val(CheckNull(rstImporter.Fields("IN3").Value)) - Val(CheckNull(rstImporter.Fields("OUT1").Value)) - Val(CheckNull(rstImporter.Fields("OUT2").Value)) - Val(CheckNull(rstImporter.Fields("OUT3").Value))
                If ClBal <> 0 Then
                    cnDatabase.Execute "Insert Into AccountChild0801 Values ('" & rstImporter.Fields("A.Code").Value & "','4','" & rstImporter.Fields("O.Code").Value & "'," & ClBal & ",'Y')"
                End If
                rstImporter.MoveNext
            Loop
            If rstImporter.State = adStateOpen Then rstImporter.Close
            'Title
            rstImporter.Open "SELECT A.Code,O.Code," & _
                                         "FORMAT((SELECT SUM(OpBal) FROM AccountChild0801 Where Category+Item='5'+O.Code AND Code=A.Code),0) As OpBal," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialIOParent M INNER JOIN MaterialIOChild I ON M.Code=I.Code WHERE I.Category+I.Item='5'+O.Code AND I.Godown=A.Code),0) As IN1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='5'+O.Code AND M.Account=A.Code AND I.Quantity>=0),0) As IN2," & _
                                         "FORMAT((SELECT SUM(ABS(Quantity)) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='5'+O.Code AND M.Account=A.Code AND I.Quantity<0),0) As OUT1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='5'+O.Code AND M.AccountFrom=A.Code),0) As OUT2," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='5'+O.Code AND M.AccountTo=A.Code),0) As IN3," & _
                                         "FORMAT((SELECT SUM(Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code)) FROM BookPOParent M INNER JOIN BookPOChild0801 I ON M.Code=I.Code WHERE LEFT(M.Type,1)<>'O' AND LEFT(M.Code,1)<>'*' AND I.Category+I.Item='5'+O.Code AND M.Binder=A.Code),0) As OUT3 " & _
                                         "FROM BookMaster O,AccountMaster A WHERE O.Board<>'000000' AND O.Type='F' AND A.Type In ('08','09') ORDER BY A.Code,O.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter.ActiveConnection = Nothing
            Do While Not rstImporter.EOF
                ClBal = Val(CheckNull(rstImporter.Fields("OpBal").Value)) + Val(CheckNull(rstImporter.Fields("IN1").Value)) + Val(CheckNull(rstImporter.Fields("IN2").Value)) + Val(CheckNull(rstImporter.Fields("IN3").Value)) - Val(CheckNull(rstImporter.Fields("OUT1").Value)) - Val(CheckNull(rstImporter.Fields("OUT2").Value)) - Val(CheckNull(rstImporter.Fields("OUT3").Value))
                If ClBal <> 0 Then
                    cnDatabase.Execute "Insert Into AccountChild0801 Values ('" & rstImporter.Fields("A.Code").Value & "','5','" & rstImporter.Fields("O.Code").Value & "'," & ClBal & ",'Y')"
                End If
                rstImporter.MoveNext
            Loop
            cnDatabase.CommitTrans
            Call MsgBox("Successfully imported the Balances !", vbInformation, App.Title)
        End If
    Else
        Call MsgBox("Nothing To Import !", vbInformation, App.Title)
    End If
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    If CxnImporter.State = adStateOpen Then cnDatabase.RollbackTrans
    BusySystemIndicator False
    DisplayError ("Failed to import the Balances")
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
End Sub
Private Sub CloseMainConnection()
    If cnDatabase.State = adStateOpen Then cnDatabase.Close
End Sub
Private Function IsFormLoaded(ByVal FormCaption As String) As Boolean
    Dim Form As Form
    IsFormLoaded = False
    For Each Form In Forms
        If Form.Caption = FormCaption Then IsFormLoaded = True: Exit For
    Next Form
End Function
Private Sub mnuCreate01_Click()
    CreateCompany True
End Sub
Private Sub mnuCreate02_Click()
    CreateCompany False
End Sub
Private Sub mnuCompanyChild_Click()
    On Error Resume Next
    Load FrmCompanyChild
    If Err.Number <> 364 Then
    FrmCompanyChild.Show vbModal
    End If
End Sub
Private Sub mnuEdit_Click()
    On Error Resume Next
    FrmCompanyMaster.strCreateCompany = "N"
    Load FrmCompanyMaster
    If Err.Number <> 364 Then FrmCompanyMaster.Show vbModal
End Sub
Private Function CreateCompany(ByVal WithMasters As Boolean)
    Dim strSource As String, strDestination As String, Cnt As Integer
    On Error GoTo ErrHandler
    FrmCompanyMaster.strCreateCompany = "Y"
    Load FrmCompanyMaster
    FrmCompanyMaster.Show vbModal
    If FrmCompanyMaster.ActionCancelled Then Call CloseForm(FrmCompanyMaster): Exit Function
    CompCode = ""
    Load FrmCompanyList
    FrmCompanyList.Caption = "Select company to create new company from..."
    FrmCompanyList.Show vbModal
    CloseMainConnection
    If CompCode = "" Then Call CloseForm(FrmCompanyMaster): Exit Function
    If DatabaseType = "MS SQL" Then
    cnDatabase.CursorLocation = adUseClient
    If cnDatabase.State = adStateOpen Then cnDatabase.Close
    cnDatabase.Open "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=Master;Data Source=" & ServerName
    rstDBList.Open "SELECT Top 1 convert(int,Right(Name,3))+1 As CompanyCode FROM Master.sys.Databases  WHERE LEFT(Name,2)='EP' AND Len(Name)=5 ORDER BY Name Desc", cnDatabase, adOpenKeyset, adLockReadOnly
    rstDBList.MoveFirst
    strSource = rstDBList.Fields("CompanyCode").Value
    ElseIf DatabaseType = "MS Access" Then
    strSource = DatabasePath & "\EP" & CompCode
    For Cnt = 1 To 999
        If Dir(DatabasePath & "\EP" & Pad(Cnt, "0", 3, "L")) = "" Then Exit For
    Next
    strDestination = DatabasePath & "\EP" & Pad(Cnt, "0", 3, "L")
    MdiMainMenu.MousePointer = vbHourglass
    FSO.CopyFile strSource, strDestination
    End If
     If UpdateComp(strSource, WithMasters, True, False, False) Then
     'ByVal CompanyCode As String, ByVal WithMasters As Boolean, ByVal CreateComp As Boolean, ByVal UpdateVersion As Boolean, ByVal UpdateMajor As Boolean
        Call MsgBox("Successfully Created the Company !", vbInformation, App.Title)
    Else
        DisplayError ("Failed to Create the Company")
        If Dir(strDestination) <> "" Then FSO.DeleteFile strDestination
    End If
    Call CloseForm(FrmCompanyMaster)
    MdiMainMenu.MousePointer = vbNormal
    Exit Function
ErrHandler:
    DisplayError ("Failed to Create the Company")
    CloseMainConnection
    MdiMainMenu.MousePointer = vbNormal
    Call CloseForm(FrmCompanyMaster)
End Function
Private Sub mnuPrintPlanning_Click(Index As Integer)
    On Error Resume Next
    FrmPrintPlanning.PlanningType = Choose(Index, "1", "2")
    Load FrmPrintPlanning
    If Err.Number <> 364 Then FrmPrintPlanning.Show
End Sub
Private Sub mnuBookPrintPlanningRegister_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Planning Register [Book]") Then
        Dim FrmBookPrintPlanningRegister As New FrmPrintPlanningRegister
        FrmBookPrintPlanningRegister.PlanningType = "1"
        Load FrmBookPrintPlanningRegister
        If Err.Number <> 364 Then FrmBookPrintPlanningRegister.Show
    End If
End Sub
Private Sub mnuTitlePrintPlanningRegister_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Planning Register [Title]") Then
        Dim FrmTitlePrintPlanningRegister As New FrmPrintPlanningRegister
        FrmTitlePrintPlanningRegister.PlanningType = "2"
        Load FrmTitlePrintPlanningRegister
        If Err.Number <> 364 Then FrmTitlePrintPlanningRegister.Show
    End If
End Sub
Private Sub mnuFinishSizeMaster_Click()
    On Error Resume Next
    FrmFinishSizeMaster.SL = False
    Load FrmFinishSizeMaster
    If Err.Number <> 364 Then FrmFinishSizeMaster.Show
End Sub
Private Sub mnuSizeGroupMaster_Click()
    On Error Resume Next
    FrmSizeGroupMaster.SL = False
    Load FrmSizeGroupMaster
    If Err.Number <> 364 Then FrmSizeGroupMaster.Show
End Sub
Private Sub mnuPaper_Click(Index As Integer)
    On Error Resume Next
    FrmPaperMaster.SL = False
    FrmPaperMaster.FormType = IIf(Index = 1, "S", "R")
    Load FrmPaperMaster
    If Err.Number <> 364 Then FrmPaperMaster.Show
End Sub
Private Sub mnuTaxMaster_Click()
    On Error Resume Next
    FrmTaxMaster.SL = False
    Load FrmTaxMaster
    If Err.Number <> 364 Then FrmTaxMaster.Show
End Sub
Private Sub mnuOutsourceItemMaster_Click()
    On Error Resume Next
    FrmOutsourceItemMaster.SL = False
    Load FrmOutsourceItemMaster
    If Err.Number <> 364 Then FrmOutsourceItemMaster.Show
End Sub
Private Sub mnuUserMaster_Click()
    On Error Resume Next
    Load FrmUserMaster
    If Err.Number <> 364 Then FrmUserMaster.Show
End Sub
Private Sub mnuPurchaseOrderJobWorkFinishedItem_Click()
    On Error Resume Next
    FrmBookPrintOrder.BookPOType = "FP"
    Load FrmBookPrintOrder
    If Err.Number <> 364 Then FrmBookPrintOrder.Show
End Sub
Private Sub mnuPurchaseOrderJobWorkUnfinishedItem_Click()
    On Error Resume Next
    FrmBookPrintOrder.BookPOType = "RP"
    Load FrmBookPrintOrder
    If Err.Number <> 364 Then FrmBookPrintOrder.Show
End Sub
Private Sub mnuPurchaseOrderJobWorkDigital_Click()
    On Error Resume Next
    FrmBookPrintOrder.BookPOType = "DP"
    Load FrmBookPrintOrder
    If Err.Number <> 364 Then FrmBookPrintOrder.Show
End Sub
Private Sub mnuPurchaseOrderSupplyInwardFinishedItem_Click()
    On Error Resume Next
    frmSalesOrderVoucher.VchType = "PO"
    Load frmSalesOrderVoucher
    If Err.Number <> 364 Then frmSalesOrderVoucher.Show
End Sub
Private Sub mnuPurchaseOrderSupplyInwardBOMItem_Click()
    On Error Resume Next
    Load FrmOutsourceItemPurchaseOrder
    If Err.Number <> 364 Then FrmOutsourceItemPurchaseOrder.Show
End Sub
Private Sub mnuSalesOrderJobWorkFinishedItem_Click()
    On Error Resume Next
    FrmBookPrintOrder.BookPOType = "FS"
    Load FrmBookPrintOrder
    If Err.Number <> 364 Then FrmBookPrintOrder.Show
End Sub
Private Sub mnuSalesOrderJobWorkUnfinishedItem_Click()
    On Error Resume Next
    FrmBookPrintOrder.BookPOType = "RS"
    Load FrmBookPrintOrder
    If Err.Number <> 364 Then FrmBookPrintOrder.Show
End Sub
Private Sub mnuSalesOrderJobWorkDigital_Click()
    On Error Resume Next
    FrmBookPrintOrder.BookPOType = "DS"
    Load FrmBookPrintOrder
    If Err.Number <> 364 Then FrmBookPrintOrder.Show
End Sub
Private Sub mnuSalesOrderSupplyOutwardFinishedItem_Click()
    On Error Resume Next
    frmSalesOrderVoucher.VchType = "SO"
    Load frmSalesOrderVoucher
    If Err.Number <> 364 Then frmSalesOrderVoucher.Show
End Sub
Private Sub mnuSalesSupplyOutwardFinishedItem_Click()
    On Error Resume Next
    frmSalesVoucher.VchType = "SF"
    Load frmSalesVoucher
    If Err.Number <> 364 Then frmSalesVoucher.Show
End Sub
Private Sub mnuSalesReturnSupplyOutwardReturnFinishedItem_Click()
    On Error Resume Next
    frmSalesVoucher.VchType = "TF"
    Load frmSalesVoucher
    If Err.Number <> 364 Then frmSalesVoucher.Show
End Sub
Private Sub mnuPurchaseSupplyInwardFinishedItem_Click()
    On Error Resume Next
    frmSalesVoucher.VchType = "PF"
    Load frmSalesVoucher
    If Err.Number <> 364 Then frmSalesVoucher.Show
End Sub
Private Sub mnuPurchaseReturnSupplyInwardReturnFinishedItem_Click()
    On Error Resume Next
    frmSalesVoucher.VchType = "OF"
    Load frmSalesVoucher
    If Err.Number <> 364 Then frmSalesVoucher.Show
End Sub
Private Sub mnuStockTranferFinishedItem_Click()
    On Error Resume Next
    frmSalesOrderVoucher.VchType = "ST"
    Load frmSalesOrderVoucher
    If Err.Number <> 364 Then frmSalesOrderVoucher.Show
End Sub
Private Sub mnuMaterialInJobWork_Click()
    On Error Resume Next
    frmItemIssueReceiptVoucher.VchType = "R"
    Load frmItemIssueReceiptVoucher
    If Err.Number <> 364 Then frmItemIssueReceiptVoucher.Show
End Sub
Private Sub mnuMaterialInSupplyInward_Click()
    On Error Resume Next
    frmSalesChallanVoucher.VchType = "RF"
    Load frmSalesChallanVoucher
    If Err.Number <> 364 Then frmSalesChallanVoucher.Show
End Sub
Private Sub mnuMaterialOutJobWork_Click()
    On Error Resume Next
    frmItemIssueReceiptVoucher.VchType = "I"
    Load frmItemIssueReceiptVoucher
    If Err.Number <> 364 Then frmItemIssueReceiptVoucher.Show
End Sub
Private Sub mnuMaterialOutSupplyOutward_Click()
    On Error Resume Next
    frmSalesChallanVoucher.VchType = "IF"
    Load frmSalesChallanVoucher
    If Err.Number <> 364 Then frmSalesChallanVoucher.Show
End Sub
Private Sub mnuBookProcessOrder_Click()
    On Error Resume Next
    Load FrmBookProcessOrder
    If Err.Number <> 364 Then FrmBookProcessOrder.Show
End Sub
Private Sub mnuMaterialIssueOrder_Click()
    On Error Resume Next
    Load FrmMaterialIssueOrder
    If Err.Number <> 364 Then FrmMaterialIssueOrder.Show
End Sub
Private Sub mnuMaterialMovement_Click()
    On Error Resume Next
    Load FrmMaterialMovement
    If Err.Number <> 364 Then FrmMaterialMovement.Show
End Sub
Private Sub mnuStockJournalRawMaterial_Click()
    On Error Resume Next
    Load FrmStockJournal
    If Err.Number <> 364 Then FrmStockJournal.Show
End Sub
Private Sub mnuStockJournalFinishedGoods_Click()
    On Error Resume Next
    frmStockJournalVoucher.VchType = "JR"
    Load frmStockJournalVoucher
    If Err.Number <> 364 Then frmStockJournalVoucher.Show
End Sub
Private Sub mnuPackingSlip_Click()
    On Error Resume Next
    FrmPackingSlip.VchType = "04SF"
    Load FrmPackingSlip
    If Err.Number <> 364 Then FrmPackingSlip.Show
End Sub
Private Sub mnuAccountMaster_Click()
    On Error Resume Next
    FrmAccountMaster.AccountType = "01"
    FrmAccountMaster.SL = False
    FrmAccountMaster.AccountGroup = "" 'All Accounts Master excluding Material Centre
    Load FrmAccountMaster
    If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
End Sub
Private Sub mnuRate_Click(Index As Integer)
    On Error Resume Next
    FrmAccountMaster.AccountType = Choose(Index, "04", "05", "06", "07", "08")
    FrmAccountMaster.AccountGroup = ""
    FrmAccountMaster.RateType = "S"
    FrmAccountMaster.SL = False
    Load FrmAccountMaster
    If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
End Sub
Private Sub mnuMaterialCentreMaster_Click()
    On Error Resume Next
    FrmAccountMaster.AccountType = "01"
    FrmAccountMaster.SL = False
    FrmAccountMaster.AccountGroup = "*99999" 'Material Centre
    Load FrmAccountMaster
    If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
End Sub
Private Sub mnuAccountGroupMaster_Click()
    On Error Resume Next
    FrmGeneralMaster.MasterType = "12"
    FrmGeneralMaster.SL = False
    Load FrmGeneralMaster
    If Err.Number <> 364 Then FrmGeneralMaster.Caption = "Account Group Master": FrmGeneralMaster.Show
End Sub
Private Sub mnuItemGroupMaster_Click()
    On Error Resume Next
    FrmGeneralMaster.MasterType = "5"
    FrmGeneralMaster.SL = False
    Load FrmGeneralMaster
    If Err.Number <> 364 Then FrmGeneralMaster.Caption = "Item Group Master": FrmGeneralMaster.Show
End Sub
Private Sub mnuBindingTypeMaster_Click()
    On Error Resume Next
    FrmBindingTypeMaster.SL = False
    Load FrmBindingTypeMaster
    If Err.Number <> 364 Then FrmBindingTypeMaster.Show
End Sub
Private Sub mnuOperationMaster_Click()
    On Error Resume Next
    FrmGeneralMaster.MasterType = "7"
    FrmGeneralMaster.SL = False
    Load FrmGeneralMaster
    If Err.Number <> 364 Then FrmGeneralMaster.Caption = "Operation Master": FrmGeneralMaster.Show
End Sub
Private Sub mnuSizeMaster_Click()
    On Error Resume Next
    FrmGeneralMaster.MasterType = "1"
    FrmGeneralMaster.SL = False
    Load FrmGeneralMaster
    If Err.Number <> 364 Then FrmGeneralMaster.Caption = "Size Master": FrmGeneralMaster.Show
End Sub
Private Sub mnuPaperUnitMaster_Click()
    On Error Resume Next
    FrmGeneralMaster.MasterType = "15"
    FrmGeneralMaster.SL = False
    Load FrmGeneralMaster
    If Err.Number <> 364 Then FrmGeneralMaster.Caption = "Paper Unit Master": FrmGeneralMaster.Show
End Sub
Private Sub mnuHSNCodeMaster_Click()
    On Error Resume Next
    FrmGeneralMaster.MasterType = "18"
    FrmGeneralMaster.SL = False
    Load FrmGeneralMaster
    If Err.Number <> 364 Then FrmGeneralMaster.Caption = "HSN Code Master": FrmGeneralMaster.Show
End Sub
Private Sub mnuBillingNarrationMaster_Click()
    On Error Resume Next
    FrmGeneralMaster.MasterType = "17"
    FrmGeneralMaster.SL = False
    Load FrmGeneralMaster
    If Err.Number <> 364 Then FrmGeneralMaster.Caption = "Billing Narration Master": FrmGeneralMaster.Show
End Sub
Private Sub mnuMachineMaster_Click()
    On Error Resume Next
    FrmMachineMaster.SL = False
    Load FrmMachineMaster
    If Err.Number <> 364 Then FrmMachineMaster.Caption = "Machine Master": FrmMachineMaster.Show
End Sub
Private Sub mnuFreshBookMaster_Click()
    On Error Resume Next
    FrmBookMaster.ItemType = "F"
    FrmBookMaster.SL = False
    Load FrmBookMaster
    If Err.Number <> 364 Then FrmBookMaster.Show
End Sub
Private Sub mnuRepairBookMaster_Click()
    On Error Resume Next
    FrmBookMaster.ItemType = "R"
    FrmBookMaster.SL = False
    Load FrmBookMaster
    If Err.Number <> 364 Then FrmBookMaster.Show
End Sub
Private Sub mnuFinanceModule_Click(Index As Integer)
    On Error Resume Next
    frmDebitCreditVoucher.VchType = Choose(Index, "PI", "PR", "JE", "CE", "DN", "CN")
    Load frmDebitCreditVoucher
    If Err.Number <> 364 Then frmDebitCreditVoucher.Show
End Sub
Private Sub mnuPaperModule_Click(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        Load FrmPaperPurchaseOrder
        If Err.Number <> 364 Then FrmPaperPurchaseOrder.Show
    ElseIf Index >= 2 And Index <= 4 Then
        frmPaperIssueReceiptVoucher.VchType = Choose(Index, "", "I", "R", "T")
        Load frmPaperIssueReceiptVoucher
        If Err.Number <> 364 Then frmPaperIssueReceiptVoucher.Show
    ElseIf Index = 5 Or Index = 6 Then
        FrmPaperDebitNote.VchType = IIf(Index = 5, "D", "C")
        Load FrmPaperDebitNote
        If Err.Number <> 364 Then FrmPaperDebitNote.Show
    End If
End Sub
Private Sub mnuQuotationSupplyInwardFinishedItem_Click()
    On Error Resume Next
    frmSalesOrderVoucher.VchType = "PQ"
    Load frmSalesOrderVoucher
    If Err.Number <> 364 Then frmSalesOrderVoucher.Show
End Sub
Private Sub mnuQuotationSupplyOutwardFinishedItem_Click()
    On Error Resume Next
    frmSalesOrderVoucher.VchType = "SQ"
    Load frmSalesOrderVoucher
    If Err.Number <> 364 Then frmSalesOrderVoucher.Show
End Sub
Private Sub MnuAccountWise_Click(Index As Integer)
    On Error Resume Next
    FrmAccountSelectionList.VchType = Trim(Index)
    Load FrmAccountSelectionList
    If Err.Number <> 364 Then FrmAccountSelectionList.Show
End Sub
Private Sub mnuPurchaseQuotationJobWork_Click(Index As Integer)
    On Error Resume Next
    frmJobworkBill.VchType = Trim(Index)
    Load frmJobworkBill
    If Err.Number <> 364 Then frmJobworkBill.Show
End Sub
Private Sub mnuSalesQuotationJobWork_Click(Index As Integer)
    On Error Resume Next
    frmJobworkBill.VchType = Trim(Index)
    Load frmJobworkBill
    If Err.Number <> 364 Then frmJobworkBill.Show
End Sub
Private Sub mnuSalesJobWork_Click(Index As Integer)
    On Error Resume Next
    frmJobworkBill.VchType = Trim(Index)
    Load frmJobworkBill
    If Err.Number <> 364 Then frmJobworkBill.Show
End Sub
Private Sub mnuPurchaseJobWork_Click(Index As Integer)
    On Error Resume Next
    frmJobworkBill.VchType = Trim(Index)
    Load frmJobworkBill
    If Err.Number <> 364 Then frmJobworkBill.Show
End Sub
Private Sub mnuDiscount_Click()
    On Error Resume Next
    Load FrmDiscountMaster
    If Err.Number <> 364 Then FrmDiscountMaster.Show vbModal
End Sub
Private Sub mnuItemOpBal_Click()
    On Error Resume Next
    Load FrmItemOpBal
    If Err.Number <> 364 Then FrmItemOpBal.Show vbModal
End Sub
Private Sub mnuProjectManagement_Click(Index As Integer)
    On Error Resume Next
    If Index <= 2 Then
        FrmGeneralMaster.MasterType = Choose(Index, "13", "14")
        FrmGeneralMaster.SL = False
        Load FrmGeneralMaster
        If Err.Number <> 364 Then FrmGeneralMaster.Caption = Choose(Index, "Department", "Designation") & " Master": FrmGeneralMaster.Show
    ElseIf Index = 3 Then
        Load FrmTeamMemberMaster
        If Err.Number <> 364 Then FrmTeamMemberMaster.Show
    End If
End Sub
Private Sub mnuProject_Click(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        Load FrmProjectAssigner
        If Err.Number <> 364 Then FrmProjectAssigner.Show
    ElseIf Index = 2 Then
        Load FrmProjectTracker
        If Err.Number <> 364 Then FrmProjectTracker.Show
    End If
End Sub
Private Sub mnuDespatchManagement_Click(Index As Integer)
    On Error Resume Next
    If Index <= 3 Then
        FrmAccountMaster.AccountType = "01"
        FrmAccountMaster.SL = False
        FrmAccountMaster.AccountGroup = Choose(Index, "*99997", "*99998", "*99996")
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
    ElseIf Index = 4 Then
        Load FrmBookingRouteMaster
        If Err.Number <> 364 Then FrmBookingRouteMaster.Show
    End If
End Sub
Private Sub MenuSaleLedger_Click(Index As Integer)
    On Error Resume Next
    FrmItemSelectionList.VchType = IIf(Trim(Index) <= 6, Trim(Index), IIf(Trim(Index) >= 7 And Trim(Index) <= 11, (Trim(Index) - 1), IIf(Trim(Index) >= 21 And Trim(Index) <= 25, (Trim(Index) - 1), IIf(Trim(Index) = 32, (Trim(Index) - 3), (Trim(Index) - 2)))))
    Load FrmItemSelectionList
    If Err.Number <> 364 Then FrmItemSelectionList.Show
End Sub
Private Sub MenuPurchaseLedger_Click(Index As Integer)
    On Error Resume Next
    FrmItemSelectionList.VchType = IIf(Trim(Index) <= 56, Trim(Index), IIf(Trim(Index) >= 57 And Trim(Index) <= 61, (Trim(Index) - 1), IIf(Trim(Index) >= 62 And Trim(Index) <= 66, (Trim(Index) - 2), IIf(Trim(Index) = 73, (Trim(Index) - 4), (Trim(Index) - 3)))))
    Load FrmItemSelectionList
    If Err.Number <> 364 Then FrmItemSelectionList.Show
End Sub
Private Sub MnuOrdersPartyWise_Click(Index As Integer)
    On Error Resume Next
    FrmItemSelectionList.VchType = IIf(Trim(Index) = 38, (Trim(Index) - 1), IIf(Trim(Index) = 39, (Trim(Index) - 1), (Trim(Index))))
    Load FrmItemSelectionList
    If Err.Number <> 364 Then FrmItemSelectionList.Show
End Sub
Private Sub MnuPOrdersPartyWise_Click(Index As Integer)
    On Error Resume Next
    FrmItemSelectionList.VchType = IIf(Trim(Index) <= 56, Trim(Index), IIf(Trim(Index) >= 57 And Trim(Index) <= 60, (Trim(Index) - 1), IIf(Trim(Index) >= 61 And Trim(Index) <= 64, (Trim(Index) - 2), (Trim(Index) - 3))))
    Load FrmItemSelectionList
    If Err.Number <> 364 Then FrmItemSelectionList.Show
End Sub
Private Sub MnuSOrdersPartyWise_Click(Index As Integer)
    On Error Resume Next
    FrmItemSelectionList.VchType = Trim(Index)
    Load FrmItemSelectionList
    If Err.Number <> 364 Then FrmItemSelectionList.Show
End Sub
Private Sub mnuColorMaster_Click()
    On Error Resume Next
    FrmGeneralMaster.MasterType = "23"
    FrmGeneralMaster.SL = False
    Load FrmGeneralMaster
    If Err.Number <> 364 Then FrmGeneralMaster.Caption = "Color Master": FrmGeneralMaster.Show
End Sub
Private Sub MnuEmailUtilities_Click()
    On Error Resume Next
    Load FrmEmailing
    If Err.Number <> 364 Then FrmEmailing.Caption = "Emailing ": FrmEmailing.Show
End Sub
Private Sub MnuHelp_Click()
    On Error Resume Next
    Dim R As Long
    If Dir(App.Path & "\HelpFiles\Easy Publish Prime v22.chm", vbDirectory) = "" Then
            R = ShellExecute(0, "open", "http://www.easyinfosolution.com", 0, 0, 1)
    Else
            R = ShellExecute(0, "open", App.Path & "\HelpFiles\Easy Publish Prime v22.chm", 0, 0, 1)
    End If
End Sub
