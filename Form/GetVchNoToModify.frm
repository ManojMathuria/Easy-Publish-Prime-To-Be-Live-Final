VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmGetVchNoToModify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Voucher"
   ClientHeight    =   1125
   ClientLeft      =   7725
   ClientTop       =   4905
   ClientWidth     =   4440
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4440
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3960
      Picture         =   "GetVchNoToModify.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   3960
      Picture         =   "GetVchNoToModify.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Proceed"
      Top             =   120
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   885
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   1561
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
      Picture         =   "GetVchNoToModify.frx":06C0
      Begin VB.TextBox Text2 
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
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   60
         TabIndex        =   0
         Top             =   120
         Width           =   2070
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   120
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
         Caption         =   " Voucher No."
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "GetVchNoToModify.frx":06DC
         Picture         =   "GetVchNoToModify.frx":06F8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   435
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
         Caption         =   " Voucher Date"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "GetVchNoToModify.frx":0714
         Picture         =   "GetVchNoToModify.frx":0730
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   1560
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   435
         Width           =   2070
         _Version        =   65536
         _ExtentX        =   3651
         _ExtentY        =   582
         Calendar        =   "GetVchNoToModify.frx":074C
         Caption         =   "GetVchNoToModify.frx":0864
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "GetVchNoToModify.frx":08D0
         Keys            =   "GetVchNoToModify.frx":08EE
         Spin            =   "GetVchNoToModify.frx":094C
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
         ReadOnly        =   -1
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "  -  -    "
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   39849
         CenturyMode     =   0
      End
   End
End
Attribute VB_Name = "FrmGetVchNoToModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TblName As String, ModName As String 'Table & Module Name
Dim rstVchSeriesList As New ADODB.Recordset
Dim VchCode As String
Private Sub Form_Load()
    Dim i As Integer
    On Error GoTo ErrorHandler
    Me.Caption = "Select Voucher To Modify (" & ModName & ")"
    CenterForm Me
    BusySystemIndicator True
    rstVchSeriesList.Open "SELECT TOP 1 Code,Name,Date FROM " & TblName & " ORDER BY Name DESC", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstVchSeriesList.RecordCount > 0 Then Text2.Text = Trim(rstVchSeriesList.Fields("Name").Value): VchCode = rstVchSeriesList.Fields("Code").Value: MhDateInput1.Text = Format(rstVchSeriesList.Fields("Date").Value, "dd-mm-yyyy")
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Form_Activate()
    Text2.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}", True
        KeyCode = 0
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then cmdCancel_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstVchSeriesList)
End Sub
Private Sub cmdProceed_Click()
    If chkExist Then
        Text2.Text = Trim(rstVchSeriesList.Fields("Name").Value)
        VchCode = rstVchSeriesList.Fields("Code").Value
        MhDateInput1.Text = Format(rstVchSeriesList.Fields("Date").Value, "dd-mm-yyyy")
        Me.Hide
        Select Case ModName
            Case "Paper Credit Note"
                        FrmPaperDebitNote.VchCode = VchCode
                        Load FrmPaperDebitNote
                        If Err.Number <> 364 Then FrmPaperDebitNote.Show
            Case "Paper Debit Note"
                        FrmPaperDebitNote.VchCode = VchCode
                        Load FrmPaperDebitNote
                        If Err.Number <> 364 Then FrmPaperDebitNote.Show
            Case "Jobwork Bill"
                        frmJobworkBill.VchCode = VchCode
                        Load frmJobworkBill
                        If Err.Number <> 364 Then frmJobworkBill.Show
            Case "Stock Journal"
                        FrmStockJournal.VchCode = VchCode
                        Load FrmStockJournal
                        If Err.Number <> 364 Then FrmStockJournal.Show
            Case "Material Issue Order"
                        FrmMaterialIssueOrder.VchCode = VchCode
                        Load FrmMaterialIssueOrder
                        If Err.Number <> 364 Then FrmMaterialIssueOrder.Show
            Case "Material Movement"
                        FrmMaterialMovement.VchCode = VchCode
                        Load FrmMaterialMovement
                        If Err.Number <> 364 Then FrmMaterialMovement.Show
            Case "Paper Receipt"
                        frmPaperIssueReceiptVoucher.VchCode = VchCode
                        frmPaperIssueReceiptVoucher.VchType = "R"
                        Load frmPaperIssueReceiptVoucher
                        If Err.Number <> 364 Then frmPaperIssueReceiptVoucher.Show
            Case "Paper Movement"
                        frmPaperIssueReceiptVoucher.VchCode = VchCode
                        frmPaperIssueReceiptVoucher.VchType = "I"
                        Load frmPaperIssueReceiptVoucher
                        If Err.Number <> 364 Then frmPaperIssueReceiptVoucher.Show
            Case "Item Receipt"
                        frmItemIssueReceiptVoucher.VchCode = VchCode
                        frmItemIssueReceiptVoucher.VchType = "R"
                        Load frmItemIssueReceiptVoucher
                        If Err.Number <> 364 Then frmItemIssueReceiptVoucher.Show
            Case "Item Issue"
                        frmItemIssueReceiptVoucher.VchCode = VchCode
                        frmItemIssueReceiptVoucher.VchType = "I"
                        Load frmItemIssueReceiptVoucher
                        If Err.Number <> 364 Then frmItemIssueReceiptVoucher.Show
            Case "Item Process Order"
                        FrmBookProcessOrder.VchCode = VchCode
                        Load FrmBookProcessOrder
                        If Err.Number <> 364 Then FrmBookProcessOrder.Show
            Case "Print Planning (MF)"
                        FrmPrintPlanning.VchCode = VchCode
                        FrmPrintPlanning.PlanningType = "1"
                        Load FrmPrintPlanning
                        If Err.Number <> 364 Then FrmPrintPlanning.Show
            Case "Print Planning (SF)"
                        FrmPrintPlanning.VchCode = VchCode
                        FrmPrintPlanning.PlanningType = "2"
                        Load FrmPrintPlanning
                        If Err.Number <> 364 Then FrmPrintPlanning.Show
            Case "Paper Purchase Order"
                        FrmPaperPurchaseOrder.VchCode = VchCode
                        Load FrmPaperPurchaseOrder
                        If Err.Number <> 364 Then FrmPaperPurchaseOrder.Show
            Case "BOM Purchase Order"
                        FrmOutsourceItemPurchaseOrder.VchCode = VchCode
                        Load FrmOutsourceItemPurchaseOrder
                        If Err.Number <> 364 Then FrmOutsourceItemPurchaseOrder.Show
            Case "Purchase Order [Finished Goods]"
                        FrmBookPrintOrder.VchCode = VchCode
                        FrmBookPrintOrder.BookPOType = "FP"
                        Load FrmBookPrintOrder
                        If Err.Number <> 364 Then FrmBookPrintOrder.Show
            Case "Purchase Order [Unfinished Goods]"
                        FrmBookPrintOrder.VchCode = VchCode
                        FrmBookPrintOrder.BookPOType = "RP"
                        Load FrmBookPrintOrder
                        If Err.Number <> 364 Then FrmBookPrintOrder.Show
            Case "Sales Order [Finished Goods]"
                        FrmBookPrintOrder.VchCode = VchCode
                        FrmBookPrintOrder.BookPOType = "FS"
                        Load FrmBookPrintOrder
                        If Err.Number <> 364 Then FrmBookPrintOrder.Show
            Case "Sales Order [Unfinished Goods]"
                        FrmBookPrintOrder.VchCode = VchCode
                        FrmBookPrintOrder.BookPOType = "RS"
                        Load FrmBookPrintOrder
                        If Err.Number <> 364 Then FrmBookPrintOrder.Show
            Case "Cost Sheet"
                        FrmBookPrintOrder.VchCode = VchCode
                        FrmBookPrintOrder.BookPOType = "OP"
                        Load FrmBookPrintOrder
                        If Err.Number <> 364 Then FrmBookPrintOrder.Show
        End Select
    End If
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Function chkExist() As Boolean
    chkExist = True
    If rstVchSeriesList.State = adStateOpen Then rstVchSeriesList.Close
    rstVchSeriesList.Open "SELECT Code,Name,Date FROM " & TblName & " AND Name='" & Pad(Trim(Text2.Text), Space(1), 10, "L") & "' ORDER BY Name DESC", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstVchSeriesList.RecordCount = 0 Then Call DisplayError("Invalid Voucher No."): Text2.SetFocus: chkExist = False
End Function
