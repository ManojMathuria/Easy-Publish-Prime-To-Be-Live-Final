VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProductionSchedule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Machine Schedule"
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6540
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   6540
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mail"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ProductionSchedule.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ProductionSchedule.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ProductionSchedule.frx":0658
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ProductionSchedule.frx":076A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6255
      Left            =   45
      TabIndex        =   7
      Top             =   345
      Width           =   6450
      _Version        =   65536
      _ExtentX        =   11377
      _ExtentY        =   11033
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
      Picture         =   "ProductionSchedule.frx":087C
      Begin VB.OptionButton Option3 
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3765
         TabIndex        =   2
         Top             =   10
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5775
         TabIndex        =   4
         Top             =   10
         Width           =   630
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4785
         TabIndex        =   3
         Top             =   10
         Width           =   990
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5940
         Left            =   0
         TabIndex        =   5
         Top             =   315
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   10478
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
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
         Caption         =   " From"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ProductionSchedule.frx":0898
         Picture         =   "ProductionSchedule.frx":08B4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   9
         Top             =   0
         Width           =   645
         _Version        =   65536
         _ExtentX        =   1138
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
         Caption         =   " To"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ProductionSchedule.frx":08D0
         Picture         =   "ProductionSchedule.frx":08EC
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "ProductionSchedule.frx":0908
         Caption         =   "ProductionSchedule.frx":0A20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ProductionSchedule.frx":0A8C
         Keys            =   "ProductionSchedule.frx":0AAA
         Spin            =   "ProductionSchedule.frx":0B08
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
         Left            =   2550
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "ProductionSchedule.frx":0B30
         Caption         =   "ProductionSchedule.frx":0C48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ProductionSchedule.frx":0CB4
         Keys            =   "ProductionSchedule.frx":0CD2
         Spin            =   "ProductionSchedule.frx":0D30
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
   End
End
Attribute VB_Name = "FrmProductionSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstProductionSchedule As New ADODB.Recordset
'Dim rstAccountList As New ADODB.Recordset
Dim rstMachineList As New ADODB.Recordset
Dim AccountType As String, ReportType As Byte
Dim oOutlook As New Outlook.Application
Dim OutputTo As String
Public VchType As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster WHERE FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    Option3.Value = True
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    MhDateInput2.Text = IIf(Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd"), Format(FinancialYearTo, "dd-mm-yyyy"), Format(Date, "dd-mm-yyyy"))
'    rstAccountList.Open "SELECT Name,Code FROM AccountMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstMachineList.Open "SELECT Name,Code FROM GeneralMaster Where Type= '21' ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    Call FillList(ListView1, "List of Machine...", rstMachineList)
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Sendkeys "{TAB}", True: KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4): KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyM Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3): KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2): KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1): KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then CloseForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstMachineList)
    Call CloseRecordset(rstProductionSchedule)
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then Cancel = True
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 4 Then CloseForm Me: Exit Sub
    OutputTo = Choose(Button.Index, "S", "P", "M")
    PrintProductionSchedule
End Sub
Private Sub PrintProductionSchedule()
    On Error Resume Next
    Dim i As Integer, iCount As Integer
    For i = 1 To ListView1.ListItems.Count
       If ListView1.ListItems(i).Checked Then iCount = iCount + 1
    Next
    Screen.MousePointer = vbHourglass
    rptProductionSchedule.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptProductionSchedule.Text9.SetText IIf(Option3.Value, "Pending Schedule ", IIf(Option1.Value, "Close Schedule ", "All Schedule")) & " From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]"
    If rstProductionSchedule.State = adStateOpen Then rstProductionSchedule.Close
    rstProductionSchedule.Open "SELECT P.Name As VchNo,RIGHT(P.Type,1) As VchType,LEFT(P.Type,1) As GoodsType,RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-MF','UFG-MF')+'/'+TRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],'1' As Col,A.Name As Party,C.Forms1 As TotalForms,C1.formsPrinted,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='P',P.QuantityReceivedC+P.QuantityReceivedB,P.QuantityIssuedC+P.QuantityIssuedB) As QtyIssued,IIF(PlateType1='1','Deepatch',IIF(PlateType1='2','PS',IIF(PlateType1='3','Wipeon','CTP'))) As Plate,IIF(Processing='N',[TotalPlates1-¼]+[TotalPlates1-½]+[TotalPlates1-1]+[RevisedPlates1],IIF(Processing='O',0,[RevisedPlates1])) As TotalPlates,C1.platesIssued, " & _
                                    " TRIM(R.Name) As Paper,INT(PaperWastageFinal1)*U.Value1+(PaperWastageFinal1-INT(PaperWastageFinal1))*1000 As Wastage,PaperConsumptionsheets1 As TotalPaper,C1.paperIssued,P.Code+'MF1' As RefCode,(Select Name From Machine where Code=C1.Machine) AS MAC,(Select Code From Machine where Code=C1.Machine) AS MCode,(Select MRT From Machine where Code=C1.Machine) AS MRT,(Select IMP From Machine where Code=C1.Machine) AS IMP,C1.SNo " & _
                                    " FROM ((((((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.Size1=S.Code) INNER JOIN PaperMaster R ON C.Paper1=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'MF1'=C1.Code) LEFT JOIN AccountMaster A ON P.BookPrinter=A.Code WHERE Forms1<>0 AND LEFT(P.Type,1)<>'O' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "#  AND P.Date<=#" & GetDate(MhDateInput2.Text) & "#  AND " & IIf(Option3.Value, "C1.formsPrinted<C.Forms1", IIf(Option1.Value, "C1.formsPrinted=C.Forms1", "1=1")) & " AND (Select Code From GeneralMaster where Code=C1.Machine) IN (" & SelectedItems(ListView1) & ") UNION " & _
                                    " SELECT P.Name As VchNo,RIGHT(P.Type,1) As VchType,LEFT(P.Type,1) As GoodsType,RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-MF','UFG-MF')+'/'+TRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],'2' As Col,A.Name As Party,C.Forms2 As TotalForms,C1.formsPrinted,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='P',P.QuantityReceivedC+P.QuantityReceivedB,P.QuantityIssuedC+P.QuantityIssuedB) As QtyIssued,IIF(PlateType2='1','Deepatch',IIF(PlateType2='2','PS',IIF(PlateType2='3','Wipeon','CTP'))) As Plate,IIF(Processing='N',[TotalPlates2-¼]+[TotalPlates2-½]+[TotalPlates2-1]+[RevisedPlates2],IIF(Processing='O',0,[RevisedPlates2])) As TotalPlates,C1.platesIssued, " & _
                                    " TRIM(R.Name) As Paper,INT(PaperWastageFinal2)*U.Value1+(PaperWastageFinal2-INT(PaperWastageFinal2))*1000 As Wastage,PaperConsumptionsheets2 As TotalPaper,C1.paperIssued,P.Code+'MF2' As RefCode,(Select Name From Machine where Code=C1.Machine) AS MAC,(Select Code From Machine where Code=C1.Machine) AS MCode,(Select MRT From Machine where Code=C1.Machine) AS MRT,(Select IMP From Machine where Code=C1.Machine) AS IMP,C1.SNo " & _
                                    " FROM ((((((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.Size2=S.Code) INNER JOIN PaperMaster R ON C.Paper2=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'MF2'=C1.Code) LEFT JOIN AccountMaster A ON P.BookPrinter=A.Code WHERE Forms2<>0 AND LEFT(P.Type,1)<>'O' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "#  AND P.Date<=#" & GetDate(MhDateInput2.Text) & "#  AND " & IIf(Option3.Value, "C1.formsPrinted<C.Forms2", IIf(Option1.Value, "C1.formsPrinted=C.Forms2", "1=1")) & " AND (Select Code From GeneralMaster where Code=C1.Machine) IN (" & SelectedItems(ListView1) & ") UNION " & _
                                    " SELECT P.Name As VchNo,RIGHT(P.Type,1) As VchType,LEFT(P.Type,1) As GoodsType,RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-MF','UFG-MF')+'/'+TRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],'4' As Col,A.Name As Party,C.Forms4 As TotalForms,C1.formsPrinted,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='P',P.QuantityReceivedC+P.QuantityReceivedB,P.QuantityIssuedC+P.QuantityIssuedB) As QtyIssued,IIF(PlateType4='1','Deepatch',IIF(PlateType4='2','PS',IIF(PlateType4='3','Wipeon','CTP'))) As Plate,IIF(Processing='N',[TotalPlates4-¼]+[TotalPlates4-½]+[TotalPlates4-1]+[RevisedPlates4],IIF(Processing='O',0,[RevisedPlates4])) As TotalPlates,C1.platesIssued, " & _
                                    " TRIM(R.Name) As Paper,INT(PaperWastageFinal4)*U.Value1+(PaperWastageFinal4-INT(PaperWastageFinal4))*1000 As Wastage,PaperConsumptionsheets4 As TotalPaper,C1.paperIssued,P.Code+'MF4' As RefCode,(Select Name From Machine where Code=C1.Machine) AS MAC,(Select Code From Machine where Code=C1.Machine) AS MCode,(Select MRT From Machine where Code=C1.Machine) AS MRT,(Select IMP From Machine where Code=C1.Machine) AS IMP,C1.SNo" & _
                                    " FROM ((((((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.Size4=S.Code) INNER JOIN PaperMaster R ON C.Paper4=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'MF4'=C1.Code) LEFT JOIN AccountMaster A ON P.BookPrinter=A.Code WHERE Forms4<>0 AND LEFT(P.Type,1)<>'O' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "#  AND P.Date<=#" & GetDate(MhDateInput2.Text) & "#  AND " & IIf(Option3.Value, "C1.formsPrinted<C.Forms4", IIf(Option1.Value, "C1.formsPrinted=C.Forms4", "1=1")) & " AND (Select Code From GeneralMaster where Code=C1.Machine) IN (" & SelectedItems(ListView1) & ") UNION " & _
                                    " SELECT P.Name As VchNo,RIGHT(P.Type,1) As VchType,LEFT(P.Type,1) As GoodsType,RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-SF','UFG-SF')+'/'+TRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],IIF(TRIM(BackPrintingType)=0,TRIM(FrontPrintingType),TRIM(FrontPrintingType)+' + '+TRIM(BackPrintingType)) As Col,A.Name As Party,Sets As TotalForms,C1.formsPrinted,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='P',P.QuantityReceivedC+P.QuantityReceivedB,P.QuantityIssuedC+P.QuantityIssuedB) As QtyIssued,IIF(PlateType='1','Deepatch',IIF(PlateType='2','PS',IIF(PlateType='3','Wipeon','CTP'))) As Plate,IIF(Processing='N',(FrontPrintingType)+(BackPrintingType),IIF(Processing='O',0,0)) As TotalPlates,C1.platesIssued, " & _
                                    " TRIM(R.Name) As Paper,INT(PaperWastageFinal)*U.Value1+(PaperWastageFinal-INT(PaperWastageFinal))*1000 As Wastage,PaperConsumptionsheets As TotalPaper,C1.paperIssued,P.Code+'SF1' As RefCode,(Select Name From GeneralMaster where Code=C1.Machine) AS MAC," & _
                                    " (Select Code From GeneralMaster where Code=C1.Machine) AS MCode,C1.SNo FROM ((((((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.[Size]=S.Code) INNER JOIN PaperMaster R ON C.Paper=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'SF1'=C1.Code) LEFT JOIN AccountMaster A ON P.TitlePrinter=A.Code WHERE Sets<>0 AND LEFT(P.Type,1)<>'O' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "#  AND P.Date<=#" & GetDate(MhDateInput2.Text) & "#  AND " & IIf(Option3.Value, "C1.formsPrinted<C.Sets", IIf(Option1.Value, "C1.formsPrinted=Sets", "1=1")) & " AND (Select Code From GeneralMaster where Code=C1.Machine) IN (" & SelectedItems(ListView1) & ") UNION " & _
                                    " SELECT P.Name As VchNo,RIGHT(P.Type,1) As VchType,LEFT(P.Type,1) As GoodsType,RIGHT(P.Type,1)+'O/'+IIF(LEFT(P.Type,1)='F','FG-CF','UFG-CF')+'/'+TRIM(P.Name) As RefNo,P.Date As RefDate,I.Name As Item,S.Name As [Size],(Select Max(C9.FrontPrintingColor) From BookPOChild0901 Where C.code=C9.Code) As Col,A.Name As Party,'1' As TotalForms,C1.formsPrinted,C.ActualQuantity As TotalQty,IIF(RIGHT(P.Type,1)='P',P.QuantityReceivedC+P.QuantityReceivedB,P.QuantityIssuedC+P.QuantityIssuedB) As QtyIssued,IIF(PlateType='1','Deepatch',IIF(PlateType='2','PS',IIF(PlateType='3','Wipeon','CTP'))) As Plate,IIF(Plate='N',(C9.FrontPrintingColor)+(C9.BackPrintingColor),IIF(Plate='O',0,0)) As Total_whPlates,C1.PlatesIssued, " & _
                                    " TRIM(R.Name) As Paper,INT(PaperWastageFinal)*U.Value1+(PaperWastageFinal-INT(PaperWastageFinal))*1000 As Wastage,PaperConsumptionsheets As TotalPaper,C1.paperIssued,P.Code+'CF1' As RefCode,(Select Name From Machine where Code=C1.Machine) AS MAC,(Select Code From Machine where Code=C1.Machine) AS MCode,(Select MRT From Machine where Code=C1.Machine) AS MRT,(Select IMP From Machine where Code=C1.Machine) AS IMP,C1.SNo" & _
                                    " FROM (((((((BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code)INNER JOIN BookPOChild0901 C9 ON C.Code=C9.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster S ON C.[Size]=S.Code) INNER JOIN PaperMaster R ON C.Paper=R.Code) INNER JOIN GeneralMaster U ON R.UOM=U.Code) LEFT JOIN BookPOChild0501 C1 ON C.Code+'CF1'=C1.Code) LEFT JOIN AccountMaster A ON P.TitlePrinter=A.Code WHERE LEFT(P.Type,1)<>'O' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "#  AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "C1.formsPrinted= 0", IIf(Option1.Value, "C1.formsPrinted=1", "1=1")) & " AND (Select Code From GeneralMaster where Code=C1.Machine) IN (" & SelectedItems(ListView1) & ") " & _
                                    " ORDER BY C1.SNo,VchType,GoodsType,VchNo,Col", cnDatabase, adOpenKeyset, adLockOptimistic
    rstProductionSchedule.ActiveConnection = Nothing
    Screen.MousePointer = vbNormal
    If rstProductionSchedule.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rstProductionSchedule.MoveFirst
    rptProductionSchedule.Database.SetDataSource rstProductionSchedule, 3, 1
    rptProductionSchedule.DiscardSavedData
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptProductionSchedule: FrmReportViewer.Show vbModal
    ElseIf OutputTo = "P" Then
        rptProductionSchedule.PaperSource = crPRBinAuto
        rptProductionSchedule.PrintOut
    Else
        If iCount >= 1 Then
            Dim oOutlookMsg As Outlook.MailItem, FileName As String
            Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
            With oOutlookMsg
                '.To = rstPaperDebitNote.Fields("EMail").Value
                .Subject = IIf(Option3.Value, "Pending Payment  ", IIf(Option1.Value, "Paid Payment", "")) & "Account Ledger"
                .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith " & IIf(Option3.Value, "Pending Payment ", IIf(Option1.Value, "Paid Payment", "")) & "Account Ledger from " + Format(GetDate(MhDateInput1.Text), "dd-MMM-yyyy") + " to " + Format(GetDate(MhDateInput2.Text), "dd-MMM-yyyy") & " for doing the needful at your end.<Br><b>Kindly do acknowledge the receipt of the mail</b>.<Br><Br>Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
                rptProductionSchedule.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptProductionSchedule.ExportOptions.DestinationType = crEDTDiskFile
                FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
                rptProductionSchedule.ExportOptions.DiskFileName = FileName
                rptProductionSchedule.Export False
                .Attachments.Add (FileName)
                .Importance = olImportanceHigh
                .ReadReceiptRequested = True
                If CheckEmpty(.To, False) Then .Display Else .Send
            End With
            Set oOutlookMsg = Nothing
        End If
    End If
    Set rptProductionSchedule = Nothing
    On Error GoTo 0
End Sub
