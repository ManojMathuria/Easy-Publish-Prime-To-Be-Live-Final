VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmReportViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Viewer"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15045
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
   MaxButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   15045
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   2
      Left            =   7995
      TabIndex        =   2
      Top             =   60
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TintColor       =   16711935
      Caption         =   "  ALT+ F4-> Exit Report"
      Alignment       =   0
      FillColor       =   8421504
      TextColor       =   16777215
      Picture         =   "ReportViewer.frx":0000
      Picture         =   "ReportViewer.frx":001C
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   7560
      Picture         =   "ReportViewer.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Alt+F4 ->> Exit Report"
      Top             =   40
      Width           =   375
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   9165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14925
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Menu MnuEMail 
      Caption         =   "E-Mail"
   End
End
Attribute VB_Name = "FrmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Report As New Report
Public Subject As String
Public Message As String
Public EMailID As String
Public CCID As String
Public Attachment As String

Private Sub Command1_Click()
       Call CloseForm(FrmReportViewer)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = vbAltMask And KeyCode = vbKeyF1 Then ' Close
       Call Command1_Click
   End If
End Sub
Private Sub Form_Load()
    With CRViewer1
        .ReportSource = Report
        .Zoom 100
        .ViewReport
    End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call CloseForm(FrmReportViewer)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub
Private Sub Form_Resize()
    With CRViewer1
        .Top = 0
        .Left = 0
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub
Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    Report.PaperSource = crPRBinAuto
    UseDefault = True
End Sub
Private Sub CRViewer1_RefreshButtonClicked(UseDefault As Boolean)
    Report.PrinterSetup (FrmReportViewer.hwnd)
    UseDefault = True
End Sub
Private Sub MnuEMail_Click()
    Dim oOutlook As New Outlook.Application
    Dim oOutlookMsg As Outlook.MailItem
    If EMailID = "" Then
        Report.PaperSource = crPRBinAuto
        Report.PrintOut True   ' Print Report With Prompt
        Exit Sub
    End If
    On Error Resume Next
    Report.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
    Report.ExportOptions.DestinationType = crEDTDiskFile
    Report.ExportOptions.DiskFileName = App.Path & "\Report\" & Trim(Attachment) & ".Pdf"
    Report.Export False
    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
    With oOutlookMsg
        .To = EMailID
        If CCID <> "" Then .CC = CCID
        .Subject = Subject
        .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</Font>"
        .Attachments.Add (App.Path & "\Report\" & Trim(Attachment) & ".Pdf")
        .Importance = olImportanceHigh
        .ReadReceiptRequested = True
        .Display
    End With
    Set oOutlookMsg = Nothing
    Set oOutlook = Nothing
End Sub
