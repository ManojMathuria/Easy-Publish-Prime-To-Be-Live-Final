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
      Tag             =   "0100"
      Begin VB.Menu mnuAccountMaster 
         Caption         =   "Account"
         Tag             =   "0101"
      End
      Begin VB.Menu mnuAccountGroupMaster 
         Caption         =   "Account Group"
      End
      Begin VB.Menu mnuRateMaster 
         Caption         =   "Rate"
         Begin VB.Menu mnuRate 
            Caption         =   "Printing"
            Index           =   1
         End
         Begin VB.Menu mnuRate 
            Caption         =   "Misc Operation"
            Index           =   2
         End
         Begin VB.Menu mnuRate 
            Caption         =   "Binding Process"
            Index           =   3
         End
         Begin VB.Menu mnuRate 
            Caption         =   "Processing"
            Index           =   4
         End
      End
      Begin VB.Menu mnuLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBook 
         Caption         =   "Item"
         Tag             =   "0102"
         Begin VB.Menu mnuFreshBookMaster 
            Caption         =   "FG"
         End
         Begin VB.Menu mnuRepairBookMaster 
            Caption         =   "UFG"
         End
      End
      Begin VB.Menu mnuItemGroupMaster 
         Caption         =   "Item Group"
         Tag             =   "0106"
      End
      Begin VB.Menu mnuBindingTypeMaster 
         Caption         =   "Binding Type"
         Tag             =   "0107"
      End
      Begin VB.Menu mnuOperationMaster 
         Caption         =   "Misc. Operation"
         Tag             =   "0108"
      End
      Begin VB.Menu MnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSizeMaster 
         Caption         =   "Size"
         Tag             =   "0109"
      End
      Begin VB.Menu mnuFinishSizeMaster 
         Caption         =   "Finish Size"
      End
      Begin VB.Menu mnuSizeGroupMaster 
         Caption         =   "Size Group"
      End
      Begin VB.Menu MnuLine8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaperMaster 
         Caption         =   "Paper"
         Tag             =   "0110"
         Begin VB.Menu mnuPaper 
            Caption         =   "Sheet"
            Index           =   1
         End
         Begin VB.Menu mnuPaper 
            Caption         =   "Reel"
            Index           =   2
         End
      End
      Begin VB.Menu mnuPaperUnitMaster 
         Caption         =   "Paper Unit"
      End
      Begin VB.Menu mnuLine11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaterialCentreMaster 
         Caption         =   "Material Centre"
      End
      Begin VB.Menu mnu777 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTaxMaster 
         Caption         =   "Tax"
      End
      Begin VB.Menu mnuLine9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOutsourceItemMaster 
         Caption         =   "BOM Item"
         Tag             =   "0111"
      End
      Begin VB.Menu MnuLine58 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHSNCodeMaster 
         Caption         =   "HSN Code"
      End
      Begin VB.Menu mnuBillingNarrationMaster 
         Caption         =   "Std. Narration"
      End
      Begin VB.Menu MnuLine15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectManagementParent 
         Caption         =   "Project Management"
         Begin VB.Menu mnuProjectManagement 
            Caption         =   "Department"
            Index           =   1
         End
         Begin VB.Menu mnuProjectManagement 
            Caption         =   "Designation"
            Index           =   2
         End
         Begin VB.Menu mnuProjectManagement 
            Caption         =   "Project Member"
            Index           =   3
         End
         Begin VB.Menu mnuProjectManagement 
            Caption         =   "Project Assigner"
            Index           =   4
         End
         Begin VB.Menu mnuProjectManagement 
            Caption         =   "Project Tracker"
            Index           =   5
         End
      End
      Begin VB.Menu MnuLine1500 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMachineMaster 
         Caption         =   "Machine"
      End
      Begin VB.Menu mnuLine676 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDespatchManagementParent 
         Caption         =   "Despatch Management"
         Begin VB.Menu mnuDespatchManagement 
            Caption         =   "Packer"
            Index           =   1
         End
         Begin VB.Menu mnuDespatchManagement 
            Caption         =   "Deliverer"
            Index           =   2
         End
         Begin VB.Menu mnuDespatchManagement 
            Caption         =   "Transporter"
            Index           =   3
         End
         Begin VB.Menu mnuDespatchManagement 
            Caption         =   "Booking Route"
            Index           =   4
         End
      End
      Begin VB.Menu mnuLine001 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserMaster 
         Caption         =   "User"
         Tag             =   "0112"
      End
   End
   Begin VB.Menu MnuTransactions 
      Caption         =   "&Transactions"
      Enabled         =   0   'False
      Tag             =   "0200"
      Begin VB.Menu mnuPrintPlanningModule 
         Caption         =   "Print Planning"
         Tag             =   "0201"
         Begin VB.Menu mnuPrintPlanning 
            Caption         =   "Multi Form Format"
            Index           =   1
         End
         Begin VB.Menu mnuPrintPlanning 
            Caption         =   "Spread Form Format"
            Index           =   2
         End
      End
      Begin VB.Menu mnuPurchasesOrder 
         Caption         =   "Purchase Order"
         Tag             =   "0202"
         Begin VB.Menu mnuPurchaseOrderJobWork 
            Caption         =   "Job Work"
            Begin VB.Menu mnuPurchaseOrderJobWorkFinishedItem 
               Caption         =   "FG Item"
            End
            Begin VB.Menu mnuPurchaseOrderJobWorkUnfinishedItem 
               Caption         =   "UFG Item"
            End
         End
         Begin VB.Menu mnuPurchaseOrderSupplyInward 
            Caption         =   "Supply Inward"
            Begin VB.Menu mnuPurchaseOrderSupplyInwardFinishedItem 
               Caption         =   "FG Item"
            End
            Begin VB.Menu mnuPurchaseOrderSupplyInwardBOMItem 
               Caption         =   "BOM Item"
            End
         End
      End
      Begin VB.Menu mnuQuotation 
         Caption         =   "Quotation"
         Begin VB.Menu mnuQuotationJW 
            Caption         =   "Job Work"
            Begin VB.Menu mnuQuotationJobWork 
               Caption         =   "Unit Cost"
               Index           =   7
            End
            Begin VB.Menu mnuQuotationJobWork 
               Caption         =   "Job Work Unit Cost"
               Index           =   8
            End
            Begin VB.Menu mnuQuotationJobWork 
               Caption         =   "Job Work"
               Index           =   9
            End
         End
         Begin VB.Menu mnuQuotationSupplyOutwardFinishedItem 
            Caption         =   "Supply Outward"
         End
      End
      Begin VB.Menu mnuSalesOrder 
         Caption         =   "Sales Order"
         Begin VB.Menu mnuSalesOrderJobWork 
            Caption         =   "Job Work"
            Begin VB.Menu mnuSalesOrderJobWorkFinishedItem 
               Caption         =   "FG Item"
            End
            Begin VB.Menu mnuSalesOrderJobWorkUnfinishedItem 
               Caption         =   "UFG Item"
            End
         End
         Begin VB.Menu mnuSalesOrderSupplyOutwardFinishedItem 
            Caption         =   "Supply Outward"
         End
      End
      Begin VB.Menu mnuSales 
         Caption         =   "Sales"
         Begin VB.Menu mnuSalesJW 
            Caption         =   "Job Work"
            Begin VB.Menu mnuSalesJobWork 
               Caption         =   "Unit Cost"
               Index           =   1
            End
            Begin VB.Menu mnuSalesJobWork 
               Caption         =   "Job Work Unit Cost"
               Index           =   2
            End
            Begin VB.Menu mnuSalesJobWork 
               Caption         =   "Job Work"
               Index           =   3
            End
         End
         Begin VB.Menu mnuSalesSupplyOutwardFinishedItem 
            Caption         =   "Supply Outward"
         End
      End
      Begin VB.Menu mnuSalesReturn 
         Caption         =   "Sales Return"
         Begin VB.Menu mnuSalesReturnSupplyOutwardReturnFinishedItem 
            Caption         =   "Supply Outward Return"
         End
      End
      Begin VB.Menu mnuPurchase 
         Caption         =   "Purchase"
         Begin VB.Menu mnuPurchaseJW 
            Caption         =   "Job Work"
            Begin VB.Menu mnuPurchaseJobWork 
               Caption         =   "Unit Cost"
               Index           =   4
            End
            Begin VB.Menu mnuPurchaseJobWork 
               Caption         =   "Job Work Unit Cost"
               Index           =   5
            End
            Begin VB.Menu mnuPurchaseJobWork 
               Caption         =   "Jobwork"
               Index           =   6
            End
         End
         Begin VB.Menu mnuPurchaseSupplyInwardFinishedItem 
            Caption         =   "Supply Inward"
         End
      End
      Begin VB.Menu mnuPurchaseReturn 
         Caption         =   "Purchase Return"
         Begin VB.Menu mnuPurchaseReturnSupplyInwardReturnFinishedItem 
            Caption         =   "Supply Inward Return"
         End
      End
      Begin VB.Menu MnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFinanceModuleParent 
         Caption         =   "Finance"
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Payment"
            Index           =   1
         End
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Receipt"
            Index           =   2
         End
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Journal"
            Index           =   3
         End
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Contra"
            Index           =   4
         End
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Debit Note"
            Index           =   5
         End
         Begin VB.Menu mnuFinanceModule 
            Caption         =   "Credit Note"
            Index           =   6
         End
      End
      Begin VB.Menu MnuLine324 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockTranferFinishedItem 
         Caption         =   "Stock Tranfer"
      End
      Begin VB.Menu mnuMaterialIn 
         Caption         =   "Material In"
         Begin VB.Menu mnuMaterialInJobWork 
            Caption         =   "Job Work"
         End
         Begin VB.Menu mnuMaterialInSupplyInward 
            Caption         =   "Supply Inward"
         End
      End
      Begin VB.Menu mnuMaterialOut 
         Caption         =   "Material Out"
         Begin VB.Menu mnuMaterialOutJobWork 
            Caption         =   "Job Work"
         End
         Begin VB.Menu mnuMaterialOutSupplyOutward 
            Caption         =   "Supply Outward"
         End
      End
      Begin VB.Menu mnuBookProcessOrder 
         Caption         =   "Item Processing Order"
         Tag             =   "0204"
      End
      Begin VB.Menu MnuLine12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaperModuleParent 
         Caption         =   "Paper"
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Purchase Order"
            Index           =   1
         End
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Issue"
            Index           =   2
         End
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Receipt"
            Index           =   3
         End
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Transfer"
            Index           =   4
            Tag             =   "0210"
         End
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Debit Note"
            Index           =   5
         End
         Begin VB.Menu mnuPaperModule 
            Caption         =   "Credit Note"
            Index           =   6
         End
      End
      Begin VB.Menu mnuLine10 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMaterialIssueOrder 
         Caption         =   "Material Issue Order"
         Tag             =   "0213"
      End
      Begin VB.Menu MnuMaterialMovement 
         Caption         =   "BOM Item Movement"
         Tag             =   "0214"
      End
      Begin VB.Menu MnuLine16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockJournal 
         Caption         =   "Stock Journal"
         Tag             =   "0215"
         Begin VB.Menu mnuStockJournalRawMaterial 
            Caption         =   "Raw Material"
         End
         Begin VB.Menu mnuStockJournalFinishedGoods 
            Caption         =   "Finished Goods"
         End
      End
      Begin VB.Menu mnu000 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPackingSlip 
         Caption         =   "Packing Slip"
      End
   End
   Begin VB.Menu MnuDisplay 
      Caption         =   "&Display"
      Enabled         =   0   'False
      Tag             =   "0700"
      Begin VB.Menu MnuFinalResult 
         Caption         =   "Final Result"
      End
      Begin VB.Menu MnuTrialBalance 
         Caption         =   "Trial Balance"
      End
      Begin VB.Menu MnuAccountBooks 
         Caption         =   "Account Books"
         Begin VB.Menu MnuDay 
            Caption         =   "Day Book"
         End
         Begin VB.Menu MnuLedger 
            Caption         =   "Ledger"
            Begin VB.Menu MnuAccountWise 
               Caption         =   "Account-Wise"
               Index           =   23
            End
         End
      End
      Begin VB.Menu MnuAccountSummary 
         Caption         =   "Account Summary"
      End
      Begin VB.Menu MnuCostCentre 
         Caption         =   "Cost Centre Report"
      End
      Begin VB.Menu MnuOutStandingAnalysis 
         Caption         =   "Outstanding Analysis"
      End
      Begin VB.Menu MnuInterestCalculation 
         Caption         =   "Interest Calculation"
      End
      Begin VB.Menu MnuProduction 
         Caption         =   "Production Scheduling"
         Begin VB.Menu MnuProductionScheduling 
            Caption         =   "Production Scheduling"
         End
         Begin VB.Menu MnuProductionSchedule 
            Caption         =   "Production Schedule Print"
         End
      End
      Begin VB.Menu MnuStockStatus 
         Caption         =   "Stock Status"
         Begin VB.Menu MnuStockLedger 
            Caption         =   "Physical Stock Audit"
            Index           =   0
         End
         Begin VB.Menu MnuStockLedger 
            Caption         =   "Inventory Ledger "
            Index           =   1
         End
         Begin VB.Menu MnuStockLedger 
            Caption         =   "Closing Stock Alphabetical "
            Index           =   2
         End
         Begin VB.Menu MnuStockLedger 
            Caption         =   "Stock List - Short Item Analysis "
            Index           =   33
         End
      End
      Begin VB.Menu MnuLine59 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuOrderProcessingStatus 
         Caption         =   "Order Status"
         Begin VB.Menu MnuOrdersSJW 
            Caption         =   "Job-Work"
            Begin VB.Menu MnuOrdersPartyWise 
               Caption         =   "Purchase Orders-Party-Wise-Detailed"
               Index           =   35
            End
            Begin VB.Menu MnuOrdersPartyWise 
               Caption         =   "Purchase Orders-Party-wise-Summarised"
               Index           =   36
            End
            Begin VB.Menu MnuOrdersPartyWise 
               Caption         =   "-"
               Index           =   37
            End
            Begin VB.Menu MnuOrdersPartyWise 
               Caption         =   "Sales Orders-Party-Wise-Detailed"
               Index           =   38
            End
            Begin VB.Menu MnuOrdersPartyWise 
               Caption         =   "Sales Orders-Party-wise-Summarised"
               Index           =   39
            End
         End
         Begin VB.Menu MnuOrdersSIW 
            Caption         =   "Supply IN-Ward"
            Begin VB.Menu MnuPOrdersPartyWise 
               Caption         =   "Purchase Orders Order-Wise"
               Index           =   39
            End
            Begin VB.Menu MnuPOrdersPartyWise 
               Caption         =   "Purchase Orders Party-wise"
               Index           =   40
            End
            Begin VB.Menu MnuPOrdersPartyWise 
               Caption         =   "Purchase Orders Item-wise"
               Index           =   41
            End
         End
         Begin VB.Menu MnuOrdersSOW 
            Caption         =   "Supply Out-Ward"
            Begin VB.Menu MnuSOrdersPartyWise 
               Caption         =   "Sale Orders Order-wise"
               Index           =   42
            End
            Begin VB.Menu MnuSOrdersPartyWise 
               Caption         =   "Sale Orders Party-wise"
               Index           =   43
            End
            Begin VB.Menu MnuSOrdersPartyWise 
               Caption         =   "Sale Orders Item-wise"
               Index           =   44
            End
         End
      End
      Begin VB.Menu MenuSaleAnalysis 
         Caption         =   "Sales Analysis"
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Item-Wise"
            Index           =   3
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Return Item-Wise"
            Index           =   4
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales And Sales Return Item-Wise"
            Index           =   5
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Net Sales Item-Wise"
            Index           =   6
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales One Party Item-Wise"
            Index           =   8
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Return One Party Item-Wise"
            Index           =   9
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales And Sales Return One Party Item-Wise"
            Index           =   10
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Net Sales One Party Item-Wise"
            Index           =   11
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "-"
            Index           =   12
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Party-Wise"
            Index           =   22
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Return Party-Wise"
            Index           =   23
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales And Sales Return Party-Wise"
            Index           =   24
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Net Sales Party-Wise"
            Index           =   25
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "-"
            Index           =   26
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales One Item Party-Wise"
            Index           =   27
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Return One Item Party-Wise"
            Index           =   28
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales And Sales Return One Item Party-Wise"
            Index           =   29
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Net Sales One Item Party-Wise"
            Index           =   30
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "-"
            Index           =   31
         End
         Begin VB.Menu MenuSaleLedger 
            Caption         =   "Sales Voucher-Wise"
            Index           =   32
         End
      End
      Begin VB.Menu MenuPurchaseAnalysis 
         Caption         =   "Purchase Analysis"
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Item-Wise"
            Index           =   53
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Return Item-Wise"
            Index           =   54
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase And Purchase Return Item-Wise"
            Index           =   55
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Net Purchase Item-Wise"
            Index           =   56
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "-"
            Index           =   57
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase One Party Item-Wise"
            Index           =   58
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Return One Party Item-Wise"
            Index           =   59
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase And Purchase Return One Party Item-Wise"
            Index           =   60
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Net Purchase One Party Item-Wise"
            Index           =   61
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "-"
            Index           =   62
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Party-Wise"
            Index           =   63
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Return Party-Wise"
            Index           =   64
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase And Purchase Return Party-Wise"
            Index           =   65
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Net Purchase Party-Wise"
            Index           =   66
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "-"
            Index           =   67
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase One Item Party-Wise"
            Index           =   68
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Return One Item Party-Wise"
            Index           =   69
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase And Purchase Return One Item Party-Wise"
            Index           =   70
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Net Purchase One Item Party-Wise"
            Index           =   71
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "-"
            Index           =   72
         End
         Begin VB.Menu MenuPurchaseLedger 
            Caption         =   "Purchase Voucher-Wise"
            Index           =   73
         End
      End
      Begin VB.Menu MnuLine60 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MenuPaperIssueReceipt 
         Caption         =   "Paper Ledger"
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Receipt Party-Wise"
            Index           =   11
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Receipt Order-Wise"
            Index           =   12
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Receipt Without-Order"
            Index           =   13
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Issue Party-Wise"
            Index           =   14
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Issue Order-Wise"
            Index           =   15
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Issue Without-Order"
            Index           =   16
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Paper Transfer Ledger"
            Index           =   17
         End
         Begin VB.Menu MenuPaperLedger 
            Caption         =   "Paper Pending Order"
            Index           =   18
         End
      End
   End
   Begin VB.Menu MnuReports 
      Caption         =   "&Reports"
      Enabled         =   0   'False
      Tag             =   "0300"
      Begin VB.Menu MnuPrintPlanningRegister 
         Caption         =   "Print Planning Register"
         Tag             =   "0301"
         Begin VB.Menu MnuBookPrintPlanningRegister 
            Caption         =   "Multi Form Format"
         End
         Begin VB.Menu MnuTitlePrintPlanningRegister 
            Caption         =   "Single Form Format"
         End
      End
      Begin VB.Menu MnuPOStatusRegister 
         Caption         =   "Order Status Register"
         Tag             =   "0302"
         Begin VB.Menu MnuPOStatusRegister01 
            Caption         =   "Itemwise"
         End
         Begin VB.Menu MnuPOStatusRegister05 
            Caption         =   "Orderwise"
         End
         Begin VB.Menu MnuPOStatusRegister03 
            Caption         =   "Multi Form Partywise"
         End
         Begin VB.Menu MnuPOStatusRegister02 
            Caption         =   "Spread Form Partywise"
         End
         Begin VB.Menu MnuPOStatusRegister04 
            Caption         =   "Binding Partywise"
         End
      End
      Begin VB.Menu MnuLine50 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPaperIssueRegister 
         Caption         =   "Paper Purchase Ledger"
         Tag             =   "0303"
      End
      Begin VB.Menu MnuPaperStockRegister 
         Caption         =   "Paper Stock Ledger"
         Tag             =   "0304"
      End
      Begin VB.Menu MnuOpBal 
         Caption         =   "Paper Opening Balance"
      End
      Begin VB.Menu MnuLine51 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMaterialStockRegister 
         Caption         =   "BOM Item Stock Register"
         Tag             =   "0305"
         Begin VB.Menu MnuMaterialStockRegister01 
            Caption         =   "Godownwise/Itemwise/BOM Itemwise"
         End
         Begin VB.Menu MnuMaterialStockRegister02 
            Caption         =   "Godownwise/BOM Itemwise"
         End
      End
      Begin VB.Menu MnuLine52 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPOStatusReg 
         Caption         =   "Purchase Order Status"
         Tag             =   "0306"
         Begin VB.Menu MnuPOStatusReg02 
            Caption         =   "BOM Item"
         End
         Begin VB.Menu MnuPOStatusReg03 
            Caption         =   "Printed Items (BOM)"
            Begin VB.Menu MnuPOStatusReg0301 
               Caption         =   "FG Item"
            End
            Begin VB.Menu MnuPOStatusReg0302 
               Caption         =   "UFG Item"
            End
            Begin VB.Menu MnuPOStatusReg0303 
               Caption         =   "Spread Format"
            End
         End
      End
      Begin VB.Menu MnuLine53 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrderLedger 
         Caption         =   "Order Ledger"
         Begin VB.Menu mnuPurchaseOrder 
            Caption         =   "Purchase Order"
            Begin VB.Menu MnuPOLedgerParty 
               Caption         =   "Party-Wise- Detailed"
               Index           =   11
            End
            Begin VB.Menu MnuPOLedgerParty 
               Caption         =   "Party-Wise- Summrised"
               Index           =   12
            End
         End
         Begin VB.Menu mnuSaleOrder 
            Caption         =   "Sale Order"
            Begin VB.Menu MnuSOLedgerParty 
               Caption         =   "Party-Wise- Detailed"
               Index           =   21
               Tag             =   "0307"
            End
            Begin VB.Menu MnuSOLedgerParty 
               Caption         =   "Party-Wise- Summrised"
               Index           =   22
            End
         End
      End
      Begin VB.Menu MnuOrderProcessing 
         Caption         =   "Order Status"
         Begin VB.Menu MnuPendingOrders 
            Caption         =   "Job-Work"
            Index           =   1
            Begin VB.Menu MnuPurchaseSaleOrderParty 
               Caption         =   "Purchase Orders-Party-Wise-Detailed"
               Index           =   11
            End
            Begin VB.Menu MnuPurchaseSaleOrderParty 
               Caption         =   "Purchase Orders-Party-wise-Summarised"
               Index           =   12
            End
            Begin VB.Menu MnuPurchaseSaleOrderParty 
               Caption         =   "Sale Orders-Party-wise-Detailed"
               Index           =   21
            End
            Begin VB.Menu MnuPurchaseSaleOrderParty 
               Caption         =   "Sale Orders-Party-wise-Summarised"
               Index           =   22
            End
         End
         Begin VB.Menu MnuPendingOrdersPO 
            Caption         =   "Supply INward"
            Index           =   2
            Begin VB.Menu MnuPurchaseSaleOrderPartyPO 
               Caption         =   "Purchase Orders Order-Wise"
               Index           =   13
            End
            Begin VB.Menu MnuPurchaseSaleOrderPartyPO 
               Caption         =   "Purchase Orders-Party-wise"
               Index           =   14
            End
            Begin VB.Menu MnuPurchaseSaleOrderPartyPO 
               Caption         =   "Purchase Orders-Item-wise"
               Index           =   15
            End
         End
         Begin VB.Menu MnuPendingOrdersSO 
            Caption         =   "Supply Outward"
            Index           =   3
            Begin VB.Menu MnuPurchaseSaleOrderPartySO 
               Caption         =   "Sale Orders Order-wise"
               Index           =   23
            End
            Begin VB.Menu MnuPurchaseSaleOrderPartySO 
               Caption         =   "Sale Orders-Party-wise"
               Index           =   24
            End
            Begin VB.Menu MnuPurchaseSaleOrderPartySO 
               Caption         =   "Sale Orders-Item-wise"
               Index           =   25
            End
         End
      End
      Begin VB.Menu MenuQuotation 
         Caption         =   "Quotation Processing"
      End
      Begin VB.Menu MnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIssueReceipt 
         Caption         =   "Issue-Receipt Analysis"
         Begin VB.Menu MnuItemIssueReceipt 
            Caption         =   "Item-wise"
            Index           =   16
         End
         Begin VB.Menu MnuItemIssueReceipt 
            Caption         =   "Item Party-wise"
            Index           =   17
         End
         Begin VB.Menu MnuItemIssueReceipt 
            Caption         =   "Item Group-wise"
            Index           =   18
         End
         Begin VB.Menu MnuItemIssueReceipt 
            Caption         =   "Item Voucher-wise"
            Index           =   19
         End
         Begin VB.Menu MnuItemIssueReceipt 
            Caption         =   "Item Date-wise"
            Index           =   20
         End
      End
      Begin VB.Menu MnuLine56 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPendingPaymentRegister 
         Caption         =   "Pending Payment Register"
         Tag             =   "0308"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuPendingDNRegister 
         Caption         =   "Pending Debit Notes Register"
         Tag             =   "0309"
      End
      Begin VB.Menu MnuLine54 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBookList1 
         Caption         =   "List of Items"
         Tag             =   "0310"
         Begin VB.Menu MnuBookList 
            Caption         =   "Items Details"
            Index           =   1
         End
         Begin VB.Menu MnuBookList 
            Caption         =   "Items Weight"
            Index           =   2
         End
      End
      Begin VB.Menu MnuCorrectionList 
         Caption         =   "Project Status Report"
         Tag             =   "0311"
      End
      Begin VB.Menu MnuLine55 
         Caption         =   "-"
      End
      Begin VB.Menu MnuProductionPlanning 
         Caption         =   "Production Fore-Casting"
         Tag             =   "0313"
         Begin VB.Menu MnuProductionPlanning01 
            Caption         =   "Main Orders"
         End
         Begin VB.Menu MnuProductionPlanning02 
            Caption         =   "Supplement Orders"
         End
      End
   End
   Begin VB.Menu MnuUtilities 
      Caption         =   "&Utilities"
      Enabled         =   0   'False
      Begin VB.Menu MnuPrintUtilities 
         Caption         =   "Print Utilities"
         Tag             =   "0315"
         Begin VB.Menu MnuBookPOPrintUtility1 
            Caption         =   "Item Order"
            Begin VB.Menu MnuBookPOPrintUtility 
               Caption         =   "Jobwork And Unit Cost"
               Index           =   1
            End
            Begin VB.Menu MnuBookPOPrintUtility 
               Caption         =   "JobCard"
               Index           =   2
            End
            Begin VB.Menu MnuBookPOPrintUtility 
               Caption         =   "Plate Orders"
               Index           =   3
            End
            Begin VB.Menu MnuBookPOPrintUtility 
               Caption         =   "Paper-Requisition-Slip"
               Index           =   4
            End
            Begin VB.Menu MnuBookPOPrintUtility 
               Caption         =   "Quotation Format"
               Index           =   5
            End
         End
         Begin VB.Menu MnuPaperPOPrintUtility 
            Caption         =   "Paper Order"
         End
      End
      Begin VB.Menu MnuLine44 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBookReceiptBusy 
         Caption         =   "Item Receipt (Busy)"
         Tag             =   "0206"
      End
      Begin VB.Menu mnuCostSheet 
         Caption         =   "Cost Estimation"
      End
      Begin VB.Menu mnuItemOpBal 
         Caption         =   "Mat Centrewise Item Op Bal"
      End
      Begin VB.Menu mnuDiscount 
         Caption         =   "Discount Structure"
      End
      Begin VB.Menu MnuImportBal 
         Caption         =   "Import Balances"
         Begin VB.Menu MnuImportBal01 
            Caption         =   "Order"
         End
         Begin VB.Menu MnuImportBal02 
            Caption         =   "Paper"
         End
         Begin VB.Menu MnuImportBal03 
            Caption         =   "BOM Item"
         End
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
Version = "EasyPublish |Rel 21.05 Ver " & App.Minor & "." & App.Revision & " |Production & Inventory Management System"
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
Private Sub mnuProductionScheduling_Click()
    On Error Resume Next
    Load FrmProductionScheduling
    If Err.Number <> 364 Then FrmProductionScheduling.Show
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
     If UpdateComp(strSource, WithMasters, True, False) Then
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
'Private Function UpdateCompany(ByVal CompanyCode As String, ByVal WithMasters As Boolean) As Boolean
'    On Error GoTo ErrorHandler
'    UpdateCompany = True
'    cnDatabase.CursorLocation = adUseClient
'    If cnDatabase.State = adStateOpen Then cnDatabase.Close
'    If DatabaseType = "MS SQL" Then
'    ConnectionString = "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=EP" & CompCode & ";Data Source=" & ServerName
'    cnDatabase.Open ConnectionString
'    ElseIf DatabaseType = "MS Access" Then
'    cnDatabase.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\EasyPublish." & CompanyCode & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
'    End If
'    cnDatabase.BeginTrans
'    If DatabaseType = "MS SQL" Then
'    'BackUpDatabse
'        cnDatabase.Execute "BACKUP DATABASE [EP" & CompCode & "] TO  DISK = N'C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\Backup\EP" & CompCode & "_LogBackup_temp.bak' WITH NOFORMAT, NOINIT,  NAME = N'EP" & CompCode & " -Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10"
'    'RestoreDatabse
'        cnDatabase.Execute "RESTORE DATABASE [EP" & CompanyCode & "] FROM  DISK = N'C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\Backup\EP" & CompCode & "_LogBackup_temp.bak' WITH  FILE = 1,  MOVE N'EPM' TO N'C:\Program Files\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL\DATA\EP" & CompanyCode & "M.mdf',  MOVE N'EPL' TO N'C:\Program Files\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL\DATA\EP" & CompanyCode & "L.ldf',  NOUNLOAD,  STATS = 5"
'    End If
'    cnDatabase.CommitTrans
'    CloseMainConnection
'    CompCode = CompanyCode
'    cnDatabase.CursorLocation = adUseClient
'    If cnDatabase.State = adStateOpen Then cnDatabase.Close
'    If DatabaseType = "MS SQL" Then
'        ConnectionString = "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=EP" & CompCode & ";Data Source=" & ServerName
'    cnDatabase.Open ConnectionString
'    End If
'    cnDatabase.BeginTrans
'    cnDatabase.Execute "DELETE FROM CompanyMaster"
'        cnDatabase.Execute "INSERT INTO CompanyMaster (Code,Name,PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,Fax,eMail,Website,GSTIN,CreatedFrom,MCGroup,MCPrimary,MCRepair,FinancialYearFrom,FinancialYearTo,Printstatus,TitleCombo,BankName,AccountNo,IFSC,TallyIntegration,BusyIntegration,FYCode,Alias) VALUES ('000001','" & Trim(FrmCompanyMaster.Text1.Text) & "','" & Trim(FrmCompanyMaster.Text2.Text) & "','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text12.Text) & "'" & _
'                                          ",'" & Trim(FrmCompanyMaster.Text8.Text) & "','" & Trim(FrmCompanyMaster.Text9.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & CompCode & "','0','0','0','" & Format(GetDate(FrmCompanyMaster.MhDateInput1.Text), "mm-dd-yyyy") & "','" & Format(GetDate(FrmCompanyMaster.MhDateInput2.Text), "mm-dd-yyyy") & "','N','1','" & Trim(FrmCompanyMaster.Text18.Text) & "','" & Trim(FrmCompanyMaster.Text19.Text) & "','" & Trim(FrmCompanyMaster.Text20.Text) & "','" & Trim(FrmCompanyMaster.Option1.Value) & "','" & Trim(FrmCompanyMaster.Option2.Value) & "','" & Trim(FrmCompanyMaster.Text16.Text) & "','" & Trim(FrmCompanyMaster.Text15.Text) & "')"
'
'    'Transactions 44_Tables
'        cnDatabase.Execute "DELETE FROM BookDNChild"
'        cnDatabase.Execute "DELETE FROM BookDNParent"
'        cnDatabase.Execute "DELETE FROM BookOOChild"
'        cnDatabase.Execute "DELETE FROM BookOOParent"
'        cnDatabase.Execute "DELETE FROM BookPOChild05"
'        cnDatabase.Execute "DELETE FROM BookPOChild0501"
'        cnDatabase.Execute "DELETE FROM BookPOChild06"
'        cnDatabase.Execute "DELETE FROM BookPOChild07"
'        cnDatabase.Execute "DELETE FROM BookPOChild08"
'        cnDatabase.Execute "DELETE FROM BookPOChild0801"
'        cnDatabase.Execute "DELETE FROM BookPOChild09"
'        cnDatabase.Execute "DELETE FROM BookPOChild0901"
'        cnDatabase.Execute "DELETE FROM BookPOParent"
'        cnDatabase.Execute "DELETE FROM BookRVChild"
'        cnDatabase.Execute "DELETE FROM BookRVParent"
'        cnDatabase.Execute "DELETE FROM DebitCreditParent"
'        cnDatabase.Execute "DELETE FROM DebitCreditChild"
'        cnDatabase.Execute "DELETE FROM DebitCreditOthInf"
'        cnDatabase.Execute "DELETE FROM DebitCreditRef"
'        cnDatabase.Execute "DELETE FROM JobworkBVChild"
'        cnDatabase.Execute "DELETE FROM JobworkBVOthInf"
'        cnDatabase.Execute "DELETE FROM JobworkBVRef"
'        cnDatabase.Execute "DELETE FROM JobworkBVParent"
'        cnDatabase.Execute "DELETE FROM MaterialIOChild"
'        cnDatabase.Execute "DELETE FROM MaterialIOParent"
'        cnDatabase.Execute "DELETE FROM MaterialMVChild"
'        cnDatabase.Execute "DELETE FROM MaterialMVParent"
'        cnDatabase.Execute "DELETE FROM MaterialSVChild"
'        cnDatabase.Execute "DELETE FROM MaterialSVParent"
'        cnDatabase.Execute "DELETE FROM OutsourceItemPOChild"
'        cnDatabase.Execute "DELETE FROM OutsourceItemPOParent"
'        cnDatabase.Execute "DELETE FROM PackingSlipChild"
'        cnDatabase.Execute "DELETE FROM PackingSlipParent"
'        cnDatabase.Execute "DELETE FROM PaperDNChild"
'        cnDatabase.Execute "DELETE FROM PaperDNParent"
'        cnDatabase.Execute "DELETE FROM PaperIOChild"
'        cnDatabase.Execute "DELETE FROM PaperMVChild"
'        cnDatabase.Execute "DELETE FROM PaperMVParent"
'        cnDatabase.Execute "DELETE FROM PaperPOChild"
'        cnDatabase.Execute "DELETE FROM PaperPOParent"
'        cnDatabase.Execute "DELETE FROM PrintPVChild"
'        cnDatabase.Execute "DELETE FROM PrintPVParent"
'        cnDatabase.Execute "DELETE FROM TatRVChild"
'        cnDatabase.Execute "DELETE FROM TatRVParent"
''Without Masters
'    If Not WithMasters Then    'Delete Master
'    'Accounts Master
'        cnDatabase.Execute "DELETE FROM AccountChild04 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountChild05 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountChild06 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountChild07 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountChild08 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountChild0801 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountMaster "
'    'BookingRouteMaster
'        cnDatabase.Execute "DELETE FROM BookingRouteMaster "
'    'Book Master
'        cnDatabase.Execute "DELETE FROM BookChild01 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookChild02 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookChild03 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookChild05 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookChild06 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookChild07 Where Left(Code,1)<>'*'"
'        'cnDatabase.Execute "DELETE FROM BookChild08 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookMaster "
'    'Other Masters
'        cnDatabase.Execute "DELETE FROM DiscountMaster "
'        cnDatabase.Execute "DELETE FROM ElementMaster "
'        If MsgBox("Do You Wants to Delete 'Finish Size Masters' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
'        cnDatabase.Execute "DELETE FROM FinishSizeChild "
'        End If
'        cnDatabase.Execute "DELETE GeneralMaster "
'        cnDatabase.Execute "DELETE FROM OutsourceItemMaster "
'        If MsgBox("Do You Wants to Delete 'Paper Master' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
'        cnDatabase.Execute "DELETE FROM PaperMaster "
'        End If
'        If MsgBox("Do You Wants to Delete 'Size Group Masters' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
'        cnDatabase.Execute "DELETE FROM SizeGroupChild "
'        End If
'        cnDatabase.Execute "DELETE FROM TaxMaster "
'        cnDatabase.Execute "DELETE FROM TeamMemberMaster "
'        cnDatabase.Execute "DELETE FROM VchSeriesMaster "
'    Else
'        cnDatabase.Execute "UPDATE AccountMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N',Opening='0'"
'        cnDatabase.Execute "UPDATE BookMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'        cnDatabase.Execute "UPDATE PaperMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'        cnDatabase.Execute "UPDATE OutsourceItemMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'        cnDatabase.Execute "UPDATE TaxMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'        cnDatabase.Execute "UPDATE TeamMemberMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'        cnDatabase.Execute "UPDATE GeneralMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'    End If
'
'
'    cnDatabase.Execute "DELETE FROM BookChild"
'    cnDatabase.Execute "DELETE FROM PaperChild"
'    cnDatabase.Execute "DELETE FROM UserChild Where Code NOT IN (Select Code from UserMaster Where Level<>1)"
'    cnDatabase.Execute "DELETE FROM UserMaster Where Level<>1"
'    cnDatabase.Execute "DELETE FROM UserAction"
'    cnDatabase.Execute "DELETE FROM VchSeriesMaster Where Left(Code,1)='*'"
'    cnDatabase.Execute "UPDATE AccountMaster SET Opening='0'"
''Default Masters
''General Accounts
'    cnDatabase.Execute "DELETE FROM AccountMaster Where Left(Code,1)='*'"
'''Account Masters
'    cnDatabase.Execute "DELETE FROM AccountMaster Where Code ='000000' Or Left(Code,1)='*'"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('000000','" & Trim(FrmCompanyMaster.Text1.Text) & "','" & Trim(FrmCompanyMaster.Text2.Text) & "','000000','*12002','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00001','Rate Master','Rate Master','1002','*12002','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00002','Main Godown','Main Godown','1003','*99999','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00003','Self Transport','Self Transport','1004','*99996','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00004','Self Packer','Self Packer','1005','*99997','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00005','Direct','Direct','1006','*99998','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"

''Finance Account
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01001','Cash','Cash','1001','*26007','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01002','Development Tax','Development Tax','1002','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01003','Edu. Cess on TDS','Edu. Cess on TDS','1003','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01004','Excise Duty','Excise Duty','1004','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01005','KKC on Service Tax','KKC on Service Tax','1005','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01006','SBC on Service Tax','SBC on Service Tax','1006','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01007','Service Tax','Service Tax','1007','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01008','SHE Cess on TDS','SHE Cess on TDS','1008','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01009','TDS (Commission or Brokerage)','TDS (Commission or Brokerage)','1009','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01010','TDS (Contracts to Individuals/HUF)','TDS (Contracts to Individuals/HUF)','1010','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01011','TDS (Contracts to Others)','TDS (Contracts to Others)','1011','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01012','TDS (Contracts to Transporter)','TDS (Contracts to Transporter)','1012','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01013','TDS (Interest from a Banking Co)','TDS (Interest from a Banking Co)','1013','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01014','TDS (Interest from a NonBanking Co)','TDS (Interest from a NonBanking Co)','1014','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01015','TDS (Professionals Services)','TDS (Professionals Services)','1015','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01016','TDS (Rent of Land)','TDS (Rent of Land)','1016','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01017','TDS (Rent of Plant & Machinery)','TDS (Rent of Plant & Machinery)','1017','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01018','TDS (Salary)','TDS (Salary)','1018','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01019','Advertisement & Publicity','Advertisement & Publicity','1019','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01020','Bad Debts Written Off','Bad Debts Written Off','1020','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01021','Bank Charges','Bank Charges','1021','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01022','Books & Periodicals','Books & Periodicals','1022','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01023','Charity & Donations','Charity & Donations','1023','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01024','Commission on Sales','Commission on Sales','1024','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01025','Conveyance Expenses','Conveyance Expenses','1025','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01026','Customer Entertainment Expenses','Customer Entertainment Expenses','1026','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01027','Depreciation A/c','Depreciation A/c','1027','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01028','Freight & Forwarding Charges','Freight & Forwarding Charges','1028','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01029','Legal Expenses','Legal Expenses','1029','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01030','Miscellaneous Expenses','Miscellaneous Expenses','1030','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01031','Office Maintenance Expenses','Office Maintenance Expenses','1031','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01032','Office Rent','Office Rent','1032','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01033','Postal Expenses','Postal Expenses','1033','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01034','Printing & Stationery','Printing & Stationery','1034','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01035','Rounded Off','Rounded Off','1035','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01036','Salary','Salary','1036','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01037','Sales Promotion Expenses','Sales Promotion Expenses','1037','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01038','Service Charges Paid','Service Charges Paid','1038','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01039','Staff Welfare Expenses','Staff Welfare Expenses','1039','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01040','Telephone Expenses','Telephone Expenses','1040','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01041','Travelling Expenses','Travelling Expenses','1041','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01042','Water & Electricity Expenses','Water & Electricity Expenses','1042','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01043','Capital Equipments','Capital Equipments','1043','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01044','Computers','Computers','1044','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01045','Furniture & Fixture','Furniture & Fixture','1045','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01046','Office Equipments','Office Equipments','1046','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01047','Plant & Machinery','Plant & Machinery','1047','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01048','Service Charges Receipts','Service Charges Receipts','1048','*26018','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01049','Profit & Loss','Profit & Loss','1049','*26001','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01050','Salary & Bonus Payable','Salary & Bonus Payable','1050','*26024','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01051','Purchase','Purchase','1051','*26025','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01052','Sales','Sales','1052','*26027','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01053','Earnest Money','Earnest Money','1053','*26029','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01054','Stock','Stock','1054','*26003','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01055','Easy Info Solutions International','Easy Info Solutions International','1055','*26030','E-461, Vijay Marg,Jagjeet Nagar','Delhi-110053','','','','+91-987-342-2907','','sales@easyinfosolution.com ','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01056','XXX Bank','XXX Bank','1056','*26004','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'
''Booking Route Master
'        cnDatabase.Execute "DELETE FROM BookingRouteMaster Where Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into BookingRouteMaster VALUES ('*00001','NOIDA-NOIDA','NOIDA-NOIDA','24.5','N')"
'        cnDatabase.Execute "Insert Into BookingRouteMaster VALUES ('*00002','NOIDA-DELHI','NOIDA-DELHI','40','N')"
'        cnDatabase.Execute "Insert Into BookingRouteMaster VALUES ('*00003','DELHI-DELHI','DELHI-DELHI','30','N')"
'
''Element Master
'    cnDatabase.Execute "DELETE FROM ElementMaster Where Left(Code,1)='*'"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00011','Text-1','Text-1','Single Sheet','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00012','Text-2','Text-2','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00013','Text-3','Text-3','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00014','Single Form','Single Form','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00015','Combo Form','Combo Form','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00016','FG','FG','FG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00017','UFG','UFG','UFG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00018','Separator','Separator','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00019','End Paper','End Paper','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00020','Cover','Cover','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00027','Title','Title','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00028','Title(GateFold)','Title(GateFold)','Single Sheet','6','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00029','PLC','PLC','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00030','Calendar Fly Leaf','Calendar Fly Leaf','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00031','Calendar Leaf','Calendar Leaf','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00032','Annual Report','Annual Report','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00033','Label','Label','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00034','Letter Head','Letter Head','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00035','Leaflet','Leaflet','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00036','Poster','Poster','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00037','Sticker','Sticker','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00038','Folders','Folders','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00039','Dust Cover','Dust Cover','Single Sheet','6','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00040','Danglar','Danglar','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00041','Carton','Carton','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00042','Carton [Inner]','Carton [Inner]','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00043','Carton [Outer]','Carton [Outer]','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00044','Card','Card','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00045','Envelope','Envelope','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'
''Finish Size Master
'    cnDatabase.Execute "DELETE FROM FinishSizeChild Where Left(Code,1)='*'"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11011','*01039','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11012','*01030','16','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11012','*01064','32','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11013','*01039','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11014','*01039','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11015','*01055','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11016','*01048','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11017','*01051','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11018','*01058','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11019','*01056','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11020','*01028','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11020','*01060','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11021','*01067','16','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11033','*01039','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11023','*01031','8','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11023','*01067','16','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11024','*01033','8','16','*01033')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11024','*01068','16','16','*01033')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11025','*01068','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11026','*01037','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11026','*01070','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11027','*01054','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11028','*01072','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11029','*01038','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11029','*01072','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11030','*01055','12','24','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11031','*01039','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11032','*01046','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11034','*01048','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11035','*01063','12','24','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11036','*01048','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11037','*01048','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11038','*01055','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11039','*01067','12','24','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11040','*01050','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11041','*01058','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11042','*01027','4','8','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11043','*01070','12','24','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11044','*01060','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11046','*01060','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11047','*01039','6','12','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11049','*01073','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11050','*01068','8','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11051','*01055','6','12','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11052','*01072','6','12','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11053','*01060','4','8','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11054','*01068','4','8','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11055','*01070','6','12','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11004','*01073','4','8','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11004','*01039','2','2','*01012')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11004','*01012','1','1','*01012')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11057','*01055','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11048','*01028','4','8','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11048','*01059','8','16','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11058','*01058','8','16','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11045','*01063','8','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11085','*01060','12','24','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11005','*01011','2','2','*01011')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11092','*01065','8','16','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11091','*01028','4','8','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11091','*01067','8','16','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11094','*01028','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11094','*01068','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11022','*01072','16','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11095','*01045','8','8','*01047')"
'
''Genral Master
''Size Master_Type-1
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='1' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01001','05.25X10.00','05.25X10.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01002','10.00X29.00','10.00X29.00','1','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01003','11.00X14.00','11.00X14.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01004','11.50X18.00','11.50X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01005','12.00X18.00','12.00X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01006','12.00X23.00','12.00X23.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01007','12.50X18.00','12.50X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01008','13.00X19.00','13.00X19.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01009','14.00X19.00','14.00X19.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01010','14.00X22.00','14.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01011','15.00X10.00 (CARD)','15.00X10.00 (CARD)','1','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01012','15.00X20.00 (CARD)','15.00X20.00 (CARD)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01013','15.00X21.00','15.00X21.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01014','15.00X27.50','15.00X27.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01015','15.50X20.00','15.50X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01016','15.50X20.50','15.50X20.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01017','15.50X21.00','15.50X21.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01018','15.50X21.50','15.50X21.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01019','16.00X20.00','16.00X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01020','16.00X20.50','16.00X20.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01021','16.00X22.00','16.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01022','16.00X24.00','16.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01023','16.00X25.00','16.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01024','16.00X30.00','16.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01025','16.50X10.50','16.50X10.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01026','17.00X22.00','17.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01027','17.00X24.00','17.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01028','18.00X23.00','18.00X23.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01029','18.00X23.00 (Card)','18.00X23.00 (Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01030','18.00X24.00','18.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01031','18.00X25.00','18.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01032','19.00X20.00','19.00X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01033','19.00X25.00','19.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01034','19.00X38.00','19.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01035','20.00X24.00','20.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01036','20.00X25.00','20.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01037','20.00X26.00','20.00X26.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01038','20.00X28.00','20.00X28.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01039','20.00X30.00','20.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01040','20.00X30.00(A/P)','20.00X30.00(A/P)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01041','20.00X30.00(Card)','20.00X30.00(Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01042','20.00X31.00','20.00X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01043','20.50X24.00','20.50X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01044','20.50X31.00','20.50X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01045','21.00X29.70 (A4)','21.00X29.70 (A4)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01046','21.00X30.00','21.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01047','21.00X30.00(CARD)','21.00X30.00(CARD)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01048','21.00X31.00','21.00X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01049','21.00X32.00','21.00X32.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01050','21.00X33.00','21.00X33.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01051','21.00X34.00','21.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01052','21.00X35.00','21.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01053','21.50X28.50','21.50X28.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01054','22.00X28.00','22.00X28.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01055','22.00X32.00','22.00X32.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01056','22.00X34.00','22.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01057','23.00X30.00','23.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01058','23.00X33.00','23.00X33.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01059','23.00X35.00','23.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01060','23.00X36.00','23.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01061','23.00X36.00(A/P)','23.00X36.00(A/P)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01062','23.00X36.00(Card)','23.00X36.00(Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01063','24.00X34.00','24.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01064','24.00X36.00','24.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01065','24.13X24.13','24.13X24.13','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01066','25.00X30.00','25.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01067','25.00X36.00','25.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01068','25.00X38.00','25.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01069','26.00X38.00','26.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01070','26.00X40.00','26.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01071','28.00X35.00','28.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01072','28.00X40.00','28.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01073','30.00X40.00','30.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01074','31.50X41.50','31.50X41.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'
''Item Group Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='5' AND Left(Code,1)='*'"
'If Trim(ReadFromFile("Client ID")) = "Publisher" Then
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05001','Activity Book','Activity Book','5','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05002','Box','Box','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05003','CARD','CARD','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05004','CATALOGUE','CATALOGUE','5','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05005','General','General','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05006','GRADE 1','GRADE 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05007','GRADE 2','GRADE 2','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05008','GRADE 3','GRADE 3','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05009','GRADE 4','GRADE 4','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05010','GRADE 5','GRADE 5','5','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05011','JUNIOR','JUNIOR','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05012','LEVEL 1','LEVEL 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05013','LEVEL 2','LEVEL 2','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05014','LEVEL 3','LEVEL 3','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05015','LEVEL 4','LEVEL 4','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05016','LEVEL 5','LEVEL 5','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05017','LEVEL A','LEVEL A','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05018','LEVEL B','LEVEL B','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05019','LEVEL C','LEVEL C','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05020','NURSERY','NURSERY','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05021','SECONDARY STD VI','SECONDARY STD','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05022','SENIOR','SENIOR','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05023','SET 1','SET 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'End If
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05024','Item Group','Item Group','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'
''Binding Type
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='6' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06001','Die_Cutting','Die_Cutting','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06002','Die_Perforation','Die_Perforation','6','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06003','Hard Bound','Hard Bound','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06004','Perfect Binding With Sewing','Perfect Binding With Sewing','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06005','Perfect Binding With Sewing(CD-Insert)','Perfect Binding With Sewing(CD-Insert)','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06006','Spiral Binding','Spiral Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06007','Wirro Binding','Wirro Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06008','Cutting & Packing','Cutting & Packing','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06009','Cutting Only','Cutting Only','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06010','Half Die Cut','Half Die Cut','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06011','Loose','Loose','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06012','Pad Gumming','Pad Gumming','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06013','Pakki Binding','Pakki Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06014','Kachchi Binding','Kachchi Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06015','Center Pinning (BOX)','Center Pinning (BOX)','6','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06016','Center Pin Binding','Center Pin Binding','6','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06017','None','None','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06018','Perfect Binding','Perfect Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Finishing Type
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='7' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07001','BOPP Gloss','BOPP Gloss','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07002','BOPP Matt','BOPP Matt','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07003','Box Packing','Box Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07004','Center Pin Binding','Center Pin Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07005','Counting & Fabrication','Counting & Fabrication','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07006','Creasing+Folding+Packing','Creasing+Folding+Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07007','Cutting and Packing','Cutting and Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07008','Cutting Leaflet Only','Cutting Leaflet Only','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07009','Die Cutting Charges','Die Cutting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07010','Die Making Charges','Die Making Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07011','Digital Print','Digital Print','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07012','Embossing','Embossing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07013','Foiling Charges','Foiling Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07014','Folding & Packing','Folding & Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07015','Graning','Graning','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07016','Half Die Cutting Charges','Half Die Cutting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07017','Hardbound Binding','Hardbound Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07018','Hologram','Hologram','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07019','Matt + Spot UV','Matt + Spot UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07020','Matt + Spot UV + Foiling + Embossing','Matt + Spot UV + Foiling + Embossing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07021','Matt + Spot UV+Glitter UV','Matt + Spot UV+Glitter UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07022','Matt Both Side','Matt Both Side','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07023','MINI Offset JOB','MINI Offset JOB','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07024','None','None','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07025','Packing Shrink','Packing Shrink','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07026','Paper Cost','Paper Cost','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07027','Pasting Charges','Pasting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07028','Perfect Binding','Perfect Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07029','Plate','Plate','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07030','Printing 4 Col','Printing 4 Col','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07031','PVC','PVC','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07032','Spot UV','Spot UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07033','Thermal Matt','Thermal Matt','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07034','UV Hybraid','UV Hybraid','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07035','Varnising','Varnising','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Project Member/Editorial Team Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='8' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08002','Author_ABC','Author_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08003','DTP_ABC','DTP_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08005','Editor_ABC','Editor_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08007','Graphic_ABC','Graphic_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08008','PPQ_ABC','PPQ_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08009','Processing_S.R.K','Processing_S.R.K','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08010','Proof Reader_ABC','Proof Reader_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08011','Type Setting_ABC','Type Setting_Sanjay Khanna','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Plate Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='9' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*09001','CTP_Plates','CTP_Plates','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*09002','Nagative-Cut Pieces','Nagative-Cut Pieces','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*09003','Nagative-One Pieces','Nagative-One Pieces','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Size Group Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='10' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10016','Extra Large-28''''X40''''','Extra Large-28''''X40''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10018','Extra Large-28''''X40''''-(Card)','Extra Large-28''''X40''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10001','Extra Large-28''''X40''''-A/P','Extra Large-28''''X40''''-A/P','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10002','Extra Large-28''''X40''''-A/P_SPL','Extra Large-28''''X40''''-A/P_SPL','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10003','Extra Large-30''''X40''''','Extra Large-30''''X40''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10004','Extra Large-30''''X40''''-(A/P)','Extra Large-30''''X40''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10005','Extra Large-30''''X40''''-(Card)','Extra Large-30''''X40''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10006','LARGE-23''''X36''''','LARGE-23''''X36''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10007','LARGE-23''''X36''''-(A/P)','LARGE-23''''X36''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10008','LARGE-23''''X36''''-(Card)','LARGE-23''''X36''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10019','Little-11.50''''X18.00''''','Little-11.50''''X18.00''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10021','Little-11.50''''X18.00''''-(A/P)','Little-11.50''''X18.00''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10020','Little-11.50''''X18.00''''-(Card)','Little-11.50''''X18.00''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10009','Medium-20''''X30''''','Medium-20''''X30''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10010','Medium-20''''X30''''(A/P)','Medium-20''''X30''''(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10011','Medium-20''''X30''''(Card)','Medium-20''''X30''''(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10012','Small-19''''X26''''','Small-19''''X26''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10017','Small-19''''X26''''-(A/P)','Small-19''''X26''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10013','Small-19''''X26''''(Card)','Small-19''''X26''''(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10014','Web-508mm','Web-508mm','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10015','Web-578mm','Web-578mm','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Finish Size Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='11' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11001','05.25x10.00','05.25x10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11002','12.00X18.00','12.00X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11003','12.00X23.00','12.00X23.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11004','14.00X19.00','14.00X19.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11005','15.00X10.00 (CARD)','15.00X10.00 (CARD)','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11006','15.50X20.50','15.50X20.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11007','16.00x20.00','16.00x20.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11008','16.00X24.00','16.00X24.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11009','16.50X10.50','16.50X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11010','17.00X22.00','17.00X22.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11011','04.00X06.87','04.00X06.87','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11012','04.25X05.50','04.25X05.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11013','04.25X07.00','04.25X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11014','04.37X07.00','04.37X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11015','04.72X07.48','04.72X07.48','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11016','05.00X07.00','05.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11017','05.00X08.00','05.00X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11018','05.06X07.81','05.06X07.81','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11019','05.25X08.00','05.25X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11020','05.50X08.50','05.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11021','05.83X08.27','05.83X08.27','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11022','06.00X08.25','06.00X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11023','06.00X08.50','06.00X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11024','06.00X09.00','06.00X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11025','06.14X09.21','06.14X09.21','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11026','06.25X09.50','06.25X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11027','06.63X10.25','06.63X10.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11028','06.69X09.61','06.69X09.61','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11029','06.75X09.50','06.75X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11030','07.00X07.00','07.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11031','07.00X09.00','07.00X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11032','07.00X10.00','07.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11033','07.25X09.50','07.25X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11034','07.44X09.69','07.44X09.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11035','07.50X07.50','07.50X07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11036','07.50X09.25','07.50X09.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11037','07.50X09.50','07.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11038','07.75X10.50','07.75X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11039','08.00X08.00','08.00X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11040','08.00X10.00','08.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11041','08.00X10.88','08.00X10.88','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11042','08.00X11.25','08.00X11.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11043','08.25X08.25','08.25X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11044','08.25X11.00','08.25X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11045','08.27X11.69','08.27X11.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11046','08.50X08.50','08.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11047','08.50X09.00','08.50X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11048','08.50X11.00','08.50X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11049','09.00X07.00','09.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11050','09.00X12.00','09.00X12.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11051','10.00X10.00','10.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11052','11.00X13.00','11.00X13.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11053','11.00X17.00','11.00X17.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11054','11.00X18.00','11.00X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11055','12.00X12.00','12.00X12.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11056','18.00X23.00','18.00X23.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11057','07.75X11.25','07.75X11.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11058','08.00X11.00','08.00X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11059','04.50X01.75','04.50X01.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11060','11.00x15.75','11.00x15.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11061','11.00X16.00','11.00X16.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11062','08.25X11.75','08.25X11.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11063','04.00X06.00','04.00X06.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11064','20.00X30.00','20.00X30.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11065','17.50X22.50','17.50X22.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11066','11.50X08.00','11.50X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11067','21.00X31.00','21.00X31.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11068','05.30X08.30','05.30X08.30','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11069','11.50X10.75','11.50X10.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11070','08.50X10.75','08.50X10.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11071','02.00X03.00','02.00X03.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11072','11.50X07.00','11.50X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11073','05.50X19.00','05.50X19.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11074','10.25X07.50','10.25X07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11075','07.50X13.75','07.50X13.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11076','07.00X02.50','07.00X02.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11077','06.50X09.50','06.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11078','04.00x07.50','04.00x07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11079','23.00X36.00','23.00X36.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11080','15.00X20.00','15.00X20.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11081','25.00X36.00','25.00X36.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11082','09.00X14.00','09.00X14.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11083','05.25X07.00','05.25X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11084','08.00X10.50','08.00X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11085','07.50X08.50','07.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11086','03.25X04.75','03.25X04.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11087','09.75X11.00','09.75X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11088','13.50X18.00','13.50X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11089','07.62X11.00','07.62X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11090','07.36X11.00','07.36X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11091','08.26X11.69','08.26X11.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11092','09.50X09.50','09.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11093','11.69X05.20','11.69X05.20','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11094','05.75X08.25','05.75X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11095','21.00X29.70','21.00X29.70','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Genral Accounts Groups
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='12' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12002','Account Group','Account Group','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*99996','Transporter','Transporter','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*99997','Packer','Transporter','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*99998','Deliverer','Deliverer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*99999','Material Centre','Material Centre','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12001','Binders','Binders','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12003','Box Supplier','Box Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12004','CD Suppliers','CD Suppliers','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12005','FG Godown','FG Godown','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12006','Laminator','Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12007','Packaging Supplier','Packaging Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12008','Paper Suppliers','Paper Suppliers','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12009','Printer','Printer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12010','Printer & Binder','Printer & Binder','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12011','Printer, Binder & Laminator','Printer, Binder & Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12012','Processor & Printer','Processor & Printer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12013','Processor, Printer & Laminator','Processor, Printer & Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12014','UFG Godown','UFG Godown','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12015','Publisher','Publisher','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12016','Clients','Clients','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12017','Cons. Supplier','Cons. Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12018','Plate Maker','Plate Maker','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
''Departments
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='13' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13001','Editorial Department','Editorial Department','13','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13002','Production Department','Production Department','13','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13003','Sales Department','Sales Department','13','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13004','Contracts Department and Legal Department','Contracts Department and Legal Department','13','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13005','Managing Editorial and Production','Managing Editorial and Production','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13006','Creative Departments','Creative Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13007','Subsidiary Rights Departments','Subsidiary Rights Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13008','Marketing, Promotion, and Advertising Departments','Marketing, Promotion, and Advertising Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13009','Publicity Department','Publicity Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13010','Publisher Website Maintenance','Publisher Website Maintenance','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13011','Finance and Accounting','Finance and Accounting','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13012','Information Technology (IT)','Information Technology (IT)','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13013','Human Resources (HR)','Human Resources (HR)','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Designation
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='14' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14001','Editor-in-Chief','Editor-in-Chief','14','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14002','Managing editor','Managing editor','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14003','Editors','Editors','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14004','Author/Writers','Author/Writers','14','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14005','Fact-checkers','Fact-checkers','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14006','Graphic Designer','Graphic Designer','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14007','Production manager','Production manager','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14008','DTP-Operator','DTP-Operator','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14009','Proof Reader','Proof Reader','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Paper Unit Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='15' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15001','Gross','Gross','15','144','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15002','Packet(100)','Packet(100)','15','100','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15003','Packet(150)','Packet(150)','15','150','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15004','Ream','Ream','15','500','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15005','Reel','Reel','15','500','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15006','Bundle (700)','Bundle (700)','15','700','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15007','Packet(200)','Packet(200)','15','200','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15008','PACKET','PACKET','15','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15009','Sheet','Sheet','15','1','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15010','Packet (250)','Packet (250)','15','250','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
''Paper Quality Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='16' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*16001','Coated Matt','Coated Matt','16','0.95','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*16002','Coated Gloss','Coated Gloss','16','0.9','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*16003','Uncoated','Uncoated','16','1.35','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*16004','High Bulk','High Bulk','16','1.4','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Narration
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='17' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17001','1. Printing & Finishing Charges of','Printing & Finishing Charges of','17','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17002','1. Text Printing Charges of','Text Printing Charges of','17','2','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17003','2. Title Printing Charges of','Title Printing Charges of','17','3','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17004','3. Combo Title Printing Charges of','Combo Title Printing Charges of','17','4','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17005','4. Finishing Charges of','Finishing Charges of','17','5','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17006','5. Binding Charges of','Binding Charges of','17','6','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17007','7. Title Printing & Finishing Charges of','Title Printing & Finishing Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17008','6. Text Printing & Finishing Charges of','Text Printing & Finishing Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17009','8. Unit Cost Charges of','Unit Cost Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17010','9. Unit Cost','.','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17011','10 Lamination Charges','Lamination Charges','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17012','11 Printed Book','Printed Book','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''HSN MASTER
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='18' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*18001','998812','998812','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*18002','998912','998912','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*18003','4901','4901','18','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*18004','49011010','49011010','18','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
''Elements MASTER
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='19'"
'        'eLEMENT mASTER mOVED TO eLEMENT mASTER
''Calculation Units MASTER
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='20' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*20001','Per Unit','Per Unit','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*20002','Per Inch²','Per Inch²','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*20003','100 Inch²','100 Inch²','20','100','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*20004','1000 Inch²','1000 Inch²','20','1000','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*20005','Per 1000','Per 1000','20','1000','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Machine Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='21' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*21046','Machine To Be Decide','Machine To Be Decide','21','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*21047','RYOBI - 4 Col','RYOBI - 4 Col','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*21048','SM 102 28x40','SM 102 28x40','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*21049','SM 74 20x29','SM 74 20x29','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*21050','Heidel 2 Col','Heidel 2 Col','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''General  Unit Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='25' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25001','Kilogram','kg.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25002','Gram','gm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25003','Milligram','mg.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25004','Liter','ltr.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25005','Milliliter','ml.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25006','Feet','ft.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25007','Inch','in.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25008','Meter','mtr.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25009','Centimeter','cm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25010','Millimeter','mm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25011','Piece','pec.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25012','Bags','bags','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25013','Roll','roll','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25014','Sets','sets','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25015','Packets','packets','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25016','Gross','gross','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25017','Dozen','dozen','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25018','Tonn','tonn','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Account Group
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='26' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26001','Profit & Loss','Profit & Loss','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26002','Revenue Accounts','Revenue Accounts','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26003','Stock-in-hand','Stock-in-hand','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26004','Bank Accounts','Bank Accounts','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26005','Bank O/D Account','Bank O/D Account','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26006','Capital Account','Capital Account','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26007','Cash-in-hand','Cash-in-hand','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26008','Current Assets','Current Assets','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26009','Current Liabilities','Current Liabilities','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26010','Depreciation Res On Machine','Depreciation Res On Machine','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26016')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26011','Duties & Taxes','Duties & Taxes','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26012','Expenses (Direct/Mfg.)','Expenses (Direct/Mfg.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26013','Expenses (Indirect/Admn.)','Expenses (Indirect/Admn.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26014','File-Sundry Creditors','File-Sundry Creditors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26015','File-Sundry Debtors','File-Sundry Debtors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26016','Fixed Assets','Fixed Assets','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26017','Income (Direct/Opr.)','Income (Direct/Opr.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26018','Income (Indirect)','Income (Indirect)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26019','Income Tax Advance','Income Tax Advance','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26021')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26020','Investments','Investments','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26021','Loans & Advances (Asset)','Loans & Advances (Asset)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26022','Loans (Liability)','Loans (Liability)','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26023','Pre-Operative Expenses','Pre-Operative Expenses','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26024','Provisions/Expenses Payable','Provisions/Expenses Payable','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26025','Purchase','Purchase','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26026','Reserves & Surplus','Reserves & Surplus','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26006')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26027','Sale','Sale','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26028','Secured Loans','Secured Loans','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26029','Securities & Deposits (Asset)','Securities & Deposits (Asset)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26030','Sundry Creditors','Sundry Creditors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26031','Sundry Debtors','Sundry Debtors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26032','Suspense Account','Suspense Account','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26033','Unsecured Loans','Unsecured Loans','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"
''Paper Master
'        cnDatabase.Execute "DELETE FROM PaperMaster Where Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00001','Art Card-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-Gloss','Art Card-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-Gloss','S','B','50.8','76.2','20','30','*15002','200','Art Card','Gloss','7.742','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00002','Art Card-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-Gloss','Art Card-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-Gloss','S','B','50.8','76.2','20','30','*15002','210','Art Card','Gloss','8.129','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00003','Art Card-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-Gloss','Art Card-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-Gloss','S','B','50.8','76.2','20','30','*15002','220','Art Card','Gloss','8.516','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00004','Art Card-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-Gloss','Art Card-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-Gloss','S','B','50.8','76.2','20','30','*15002','250','Art Card','Gloss','9.677','5','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00005','Art Card-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-Gloss','Art Card-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-Gloss','S','B','58.42','91.44','23','36','*15002','200','Art Card','Gloss','10.684','5','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00006','Art Card-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-Gloss','Art Card-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-Gloss','S','B','58.42','91.44','23','36','*15002','210','Art Card','Gloss','11.218','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00007','Art Card-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-Gloss','Art Card-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-Gloss','S','B','58.42','91.44','23','36','*15002','220','Art Card','Gloss','11.752','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00008','Art Card-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-Gloss','Art Card-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-Gloss','S','B','58.42','91.44','23','36','*15002','250','Art Card','Gloss','13.355','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00009','Art Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Gloss','Art Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Gloss','S','P','50.8','76.2','20','30','*15004','70','Art Paper','Gloss','13.548','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00010','Art Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Gloss','Art Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Gloss','S','P','50.8','76.2','20','30','*15004','80','Art Paper','Gloss','15.484','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00011','Art Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Gloss','Art Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Gloss','S','P','50.8','76.2','20','30','*15004','90','Art Paper','Gloss','17.419','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00012','Art Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Gloss','Art Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Gloss','S','P','50.8','76.2','20','30','*15004','100','Art Paper','Gloss','19.355','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00013','Art Paper-130gsm-20.00X30.00in²-(50.80X76.20cm²)-25.161kg-Gloss','Art Paper-130gsm-20.00X30.00in²-(50.80X76.20cm²)-25.161kg-Gloss','S','P','50.8','76.2','20','30','*15004','130','Art Paper','Gloss','25.161','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00014','Art Paper-170gsm-20.00X30.00in²-(50.80X76.20cm²)-32.903kg-Gloss','Art Paper-170gsm-20.00X30.00in²-(50.80X76.20cm²)-32.903kg-Gloss','S','P','50.8','76.2','20','30','*15004','170','Art Paper','Gloss','32.903','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00015','Art Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Gloss','Art Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Gloss','S','P','58.42','91.44','23','36','*15004','70','Art Paper','Gloss','18.697','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00016','Art Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Gloss','Art Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Gloss','S','P','58.42','91.44','23','36','*15004','80','Art Paper','Gloss','21.368','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00017','Art Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Gloss','Art Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Gloss','S','P','58.42','91.44','23','36','*15004','90','Art Paper','Gloss','24.039','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00018','Art Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Gloss','Art Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Gloss','S','P','58.42','91.44','23','36','*15004','100','Art Paper','Gloss','26.71','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00019','Art Paper-130gsm-23.00X36.00in²-(58.42X91.44cm²)-34.723kg-Gloss','Art Paper-130gsm-23.00X36.00in²-(58.42X91.44cm²)-34.723kg-Gloss','S','P','58.42','91.44','23','36','*15004','130','Art Paper','Gloss','34.723','1','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00020','Art Paper-170gsm-23.00X36.00in²-(58.42X91.44cm²)-45.406kg-Gloss','Art Paper-170gsm-23.00X36.00in²-(58.42X91.44cm²)-45.406kg-Gloss','S','P','58.42','91.44','23','36','*15004','170','Art Paper','Gloss','45.406','1','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00021','Paper-60gsm-20.00X30.00in²-(50.80X76.20cm²)-11.613kg-Maplitho','Paper-60gsm-20.00X30.00in²-(50.80X76.20cm²)-11.613kg-Maplitho','S','P','50.8','76.2','20','30','*15004','60','Paper','Maplitho','11.613','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00022','Paper-64gsm-20.00X30.00in²-(50.80X76.20cm²)-12.387kg-Maplitho','Paper-64gsm-20.00X30.00in²-(50.80X76.20cm²)-12.387kg-Maplitho','S','P','50.8','76.2','20','30','*15004','64','Paper','Maplitho','12.387','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00023','Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Maplitho','Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Maplitho','S','P','50.8','76.2','20','30','*15004','70','Paper','Maplitho','13.548','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00024','Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Maplitho','Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Maplitho','S','P','50.8','76.2','20','30','*15004','80','Paper','Maplitho','15.484','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00025','Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Maplitho','Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Maplitho','S','P','50.8','76.2','20','30','*15004','90','Paper','Maplitho','17.419','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00026','Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Maplitho','Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Maplitho','S','P','50.8','76.2','20','30','*15004','100','Paper','Maplitho','19.355','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00027','Paper-120gsm-20.00X30.00in²-(50.80X76.20cm²)-23.226kg-Maplitho','Paper-120gsm-20.00X30.00in²-(50.80X76.20cm²)-23.226kg-Maplitho','S','P','50.8','76.2','20','30','*15004','120','Paper','Maplitho','23.226','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00028','Paper-60gsm-23.00X36.00in²-(58.42X91.44cm²)-16.026kg-Maplitho','Paper-60gsm-23.00X36.00in²-(58.42X91.44cm²)-16.026kg-Maplitho','S','P','58.42','91.44','23','36','*15004','60','Paper','Maplitho','16.026','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00029','Paper-64gsm-23.00X36.00in²-(58.42X91.44cm²)-17.094kg-Maplitho','Paper-64gsm-23.00X36.00in²-(58.42X91.44cm²)-17.094kg-Maplitho','S','P','58.42','91.44','23','36','*15004','64','Paper','Maplitho','17.094','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00030','Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Maplitho','Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Maplitho','S','P','58.42','91.44','23','36','*15004','70','Paper','Maplitho','18.697','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00031','Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Maplitho','Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Maplitho','S','P','58.42','91.44','23','36','*15004','80','Paper','Maplitho','21.368','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00032','Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Maplitho','Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Maplitho','S','P','58.42','91.44','23','36','*15004','90','Paper','Maplitho','24.039','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00033','Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Maplitho','Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Maplitho','S','P','58.42','91.44','23','36','*15004','100','Paper','Maplitho','26.71','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00034','Paper-120gsm-23.00X36.00in²-(58.42X91.44cm²)-32.052kg-Maplitho','Paper-120gsm-23.00X36.00in²-(58.42X91.44cm²)-32.052kg-Maplitho','S','P','58.42','91.44','23','36','*15004','120','Paper','Maplitho','32.052','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00035','SBS-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-C1S','SBS-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-C1S','S','B','50.8','76.2','20','30','*15002','200','SBS','C1S','7.742','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00036','SBS-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-C1S','SBS-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-C1S','S','B','50.8','76.2','20','30','*15002','210','SBS','C1S','8.129','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00037','SBS-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-C1S','SBS-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-C1S','S','B','50.8','76.2','20','30','*15002','220','SBS','C1S','8.516','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00038','SBS-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-C1S','SBS-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-C1S','S','B','50.8','76.2','20','30','*15002','250','SBS','C1S','9.677','5','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00039','SBS-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-C1S','SBS-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-C1S','S','B','58.42','91.44','23','36','*15002','200','SBS','C1S','10.684','5','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00040','SBS-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-C1S','SBS-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-C1S','S','B','58.42','91.44','23','36','*15002','210','SBS','C1S','11.218','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00041','SBS-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-C1S','SBS-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-C1S','S','B','58.42','91.44','23','36','*15002','220','SBS','C1S','11.752','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00042','SBS-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-C1S','SBS-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-C1S','S','B','58.42','91.44','23','36','*15002','250','SBS','C1S','13.355','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'
''Size Group Master
'        cnDatabase.Execute "DELETE FROM SizeGroupChild Where Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10003','*01067')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10003','*01068')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10003','*01070')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10003','*01072')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10003','*01073')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10007','*01061')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10011','*01047')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01050')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01051')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01056')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01058')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01060')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01063')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01064')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01059')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01017')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01020')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01021')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01027')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01028')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01030')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01031')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01033')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01013')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01012')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01015')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01016')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01019')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01029')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01018')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10018','*01069')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01036')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01037')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01038')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01039')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01046')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01048')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01054')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01057')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10020','*01011')"
''Tax Master
'        cnDatabase.Execute "DELETE FROM TaxMaster Where Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00001','Local GST 12%','Local GST 12%','L','6','6',0,'000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00002','IGST 12%','IGST 12%','I','0','0',12,'000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00003','IGST 5%','IGST 5%','I','0','0',5,'000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00004','Local GST 5%','Local GST 5%','L','2.5','2.5',0,'000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00005','Local GST 18%','Local GST 18%','L','9','9',0,'000006',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00006','IGST 18%','IGST 18%','I','0','0',18,'000006',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00007','Local GST NIL','Local GST NIL','L','0','0',0,'000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00008','IGST NIL','IGST NIL','I','0','0',0,'000001',GetDate(),'NULL',NULL,'N','N')"
''Vch Series Master
'        cnDatabase.Execute "DELETE FROM VchSeriesMaster Where Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00101','Main','01PF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Purc','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00102','Main','01PU','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrJU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00103','Main','01PC','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrJC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00104','Main','01PJ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00201','Main','02OF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRt','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00202','Main','02OU','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRtJU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00203','Main','02OC','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRtJC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00204','Main','02OJ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRtJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00301','Main','03TF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRt','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00302','Main','03TU','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRtJU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00303','Main','03TC','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRtJC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00304','Main','03TJ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRtJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00401','Main','04SF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Sale','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00402','Main','04SU','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlJU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00403','Main','04SC','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlJC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00404','Main','04SJ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00501','Main','05RF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/MtRc','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00502','Main','05FR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/MtRcJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00601','Main','06IF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRtC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00602','Main','06FI','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRtCJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00701','Main','07RF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRtC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00702','Main','07FR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRtCJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00801','Main','08IF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/MtIs','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00802','Main','08FI','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/MtIsJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*01701','Main','17PO','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PO','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*01801','Main','18SO','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SO','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*01901','Main','19ST','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/STrn','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02001','Main','20JR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SJrnl','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02101','Main','21JR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SJrnl','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02201','Main','22JR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SJrnl','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02301','Main','23PQ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PQ','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02302','Main','23ZU','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PQU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02303','Main','23ZC','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PQC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02304','Main','23ZJ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PQJ','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02305','Main','24SQ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SQ','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02306','Main','24QU','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SQU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02307','Main','24QC','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SQC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02308','Main','24QJ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SQJ','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05101','Main','51PI','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Pymt','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05201','Main','52PR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Rcpt','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05301','Main','53JE','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Jrnl','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05401','Main','54CE','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Cntr','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05501','Main','55CN','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/CrNt','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05601','Main','56DN','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/DrNt','A')"
''CompChild
'        cnDatabase.Execute "DELETE FROM CompChild "
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','01','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','SEPL/Pur/','/20-21','Purchase')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','02','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','SEPL/PR/','/20-21','Purchase Return')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','03','1. All disputes are subject to Our Jurisdiction Only','2. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection.','3. Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','','','SEPL/SR/','/20-21','Sale Return')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','04','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/Sale/','/20-21','Sale')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','05','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','SEPL/PC/','/20-21','Purchase Challan IN')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','06','','','','','','','','SEPL/PRC/','/20-21','Purchase Challan Out')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','07','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SRC/','/20-21','Sale Challan IN')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','08','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Sale Challan Out')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','09','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SJ/','/20-21','Sale Jobwork')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','10','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Sale Jobwork Unit Cost')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','11','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/DN/','/20-21','Challan Revesal IN')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','12','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Delhi Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SFAPL/PU/','/20-21','Challan Revesal Out')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','13','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan TO Be Billed IN')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','14','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan TO Be Billed OUT')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','15','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan Not TO Be Billed IN')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','16','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan Not TO Be Billed IOUT')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','17','1. The Deliverables shall be delivered or performed on the ','date and at the place specified in the Purchase Order.','2. Prices shall be as specified in the  Purchase  Order.','3. No increase in price shall be made or accepted unless ',' agreed in writing by Accenture.','4. The  Deliverables must conform in all respects with the','   Specifications and must be of sound.','SEPL/PO/','/20-21','Purchase Order')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','18','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SO/','/20-21','Sale Order')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','19','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/ST/','/20-21','Stock Tranfer')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','20','','','','','','','','SEPL/RN/','/20-21','Stock Genral')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','21','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Delhi Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SFAPL/SU/','/20-21','Promotional Sale Challan Out')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','22','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SQ/','/20-21','--')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','23','1. The price set for in Supplier’s Quotation (“Price”) are',' in  INDIA INR.','2. All Taxes shall be paid by Customer in addition to the ',' Price.','3.  Quotation (“Prices”) are valid for 30 days only.','','','SEPL/QP/','/20-21','Purchase Quotation')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','24','1. The price set for in Supplier’s Quotation (“Price”) are',' in  INDIA INR.','2. All Taxes shall be paid by Customer in addition to the ',' Price.','3.  Quotation (“Prices”) are valid for 30 days only.','','','SEPL/QS/','/20-21','Sales Quotation')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','25','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','26','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','27','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','28','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','29','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','30','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','31','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','32','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','33','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','34','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','35','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','36','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','37','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','38','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','39','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','40','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','41','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','42','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','43','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','44','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','45','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','46','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','47','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','48','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','49','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','50','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','51','','','','','','','','SEPL/PI/','/20-21','Payment')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','52','','','','','','','','SEPL/PR/','/20-21','Receipt')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','53','','','','','','','','SEPL/JE/','/20-21','Journal')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','54','','','','','','','','SEPL/CE/','/20-21','Contra')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','55','','','','','','','','SEPL/DN/','/20-21','Debit Note')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','56','','','','','','','','SEPL/CN/','/20-21','Credit Note')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','57','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','58','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','59','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','60','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','61','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','62','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','63','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','64','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','65','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','66','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','67','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','68','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','69','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','70','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','71','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','72','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','73','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','74','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','75','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','76','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','77','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','78','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','79','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','80','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','81','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','82','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','83','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','84','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','85','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','86','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','87','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','88','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','89','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','90','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','91','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','92','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','93','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','94','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','95','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','96','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','97','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','98','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','99','','','','','','','','','','')"
'    cnDatabase.CommitTrans
'    CloseMainConnection
'    Exit Function
'ErrorHandler:
'    UpdateCompany = False
'    cnDatabase.RollbackTrans
'    CloseMainConnection
'End Function
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
    FrmAccountMaster.AccountType = Choose(Index, "05", "07", "08", "04")
    FrmAccountMaster.SL = False
    FrmAccountMaster.Caption = Choose(Index, "Printing", "Misc Operation", "Binding", "Processing") & " Rate Master"
    Load FrmAccountMaster
    If Err.Number <> 364 Then FrmAccountMaster.Show
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
    FrmGeneralMaster.MasterType = "21"
    FrmGeneralMaster.SL = False
    Load FrmGeneralMaster
    If Err.Number <> 364 Then FrmGeneralMaster.Caption = "Machine Master": FrmGeneralMaster.Show
End Sub
Private Sub mnuFreshBookMaster_Click()
    On Error Resume Next
    FrmBookMaster.BookType = "F"
    FrmBookMaster.SL = False
    Load FrmBookMaster
    If Err.Number <> 364 Then FrmBookMaster.Show
End Sub
Private Sub mnuRepairBookMaster_Click()
    On Error Resume Next
    FrmBookMaster.BookType = "R"
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
Private Sub mnuQuotationJobWork_Click(Index As Integer)
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
    ElseIf Index = 4 Then
        Load FrmProjectAssigner
        If Err.Number <> 364 Then FrmProjectAssigner.Show
    ElseIf Index = 5 Then
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
