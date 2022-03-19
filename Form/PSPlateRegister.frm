VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmPSPlateRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CTP/ PS Plate Register"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14175
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
   ScaleHeight     =   7425
   ScaleWidth      =   14175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Height          =   375
      Left            =   13680
      Picture         =   "PSPlateRegister.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   7200
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   105
      Width           =   13455
      _Version        =   65536
      _ExtentX        =   23733
      _ExtentY        =   12700
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
      Picture         =   "PSPlateRegister.frx":0102
      Begin VB.TextBox Text1 
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
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   105
         Width           =   5175
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
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   105
         Width           =   5655
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   6465
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   630
         Width           =   13215
         _Version        =   524288
         _ExtentX        =   23310
         _ExtentY        =   11404
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
         MaxCols         =   9
         OperationMode   =   2
         SpreadDesigner  =   "PSPlateRegister.frx":011E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   105
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
         Caption         =   " Item Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "PSPlateRegister.frx":0889
         Picture         =   "PSPlateRegister.frx":08A5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   6840
         TabIndex        =   6
         Top             =   105
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
         Caption         =   " Element Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "PSPlateRegister.frx":08C1
         Picture         =   "PSPlateRegister.frx":08DD
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   14000
         Y1              =   525
         Y2              =   525
      End
   End
End
Attribute VB_Name = "FrmPSPlateRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ItemCode As String, ItemName As String, ElementCode As String, ElementName As String, OrderCode As String, OrderDate As Date, OrderType As String, PlateType As String, TblSuffix As String
Dim cnImporter As New ADODB.Connection, rstImporter As New ADODB.Recordset, rstPSPlateRegister As New ADODB.Recordset, LYDBName As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    Text1.Text = ElementName: Text2.Text = ItemName
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    LYDBName = Trim(ReadFromFile("LY Database Name"))
    Screen.MousePointer = vbHourglass
    cnImporter.CursorLocation = adUseClient
    If cnImporter.State = adStateOpen Then cnImporter.Close
    
    If Not CheckEmpty(LYDBName, False) Then
        cnImporter.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\" & LYDBName & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
        If TblSuffix = "06" Then
            rstImporter.Open "SELECT P.Name As OrderNo,C.OrderDate,A.Name As PartyName,C.Processing," & IIf(PlateType = "F", "C.PlateType", "C.PlateTypeBack") & " As PlateType,C.ActualQuantity As Quantity,C.PlateRate As Rate,C.BillNo,C.BillDate,I.Name As ItemName,E.Name As ElementName " & _
                                                 "FROM (((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster A ON P.TitlePrinter=A.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code " & _
                                                 "WHERE P.Type='" & OrderType & "' AND C.Element='" & ElementCode & "' AND P.Book='" & ItemCode & "' AND C.OrderDate<='" & GetDate(Format(DateAdd("d", -365, CDate(OrderDate)), "dd-mm-yyyy")) & "'  " & _
                                                 "ORDER BY A.Name,P.Name", cnImporter, adOpenKeyset, adLockReadOnly
'        Else
'            rstImporter.Open "SELECT P.Name As OrderNo,C1.OrderDate,M1.PrintName As PrinterName,C1.Plate As Processing,C1.PlateType,C2.ActualQuantity As Quantity,C1.PlateRate As Rate,C1.BillNo,C1.BillDate,TRIM(M2.PrintName)+' (Combo)' As ItemName " & _
'                             "FROM (((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN AccountMaster M1 ON P.TitlePrinter=M1.Code) INNER JOIN BookMaster M2 ON C2.Book=M2.Code " & _
'                             "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND C2.Code IN (SELECT Code FROM BookPOChild0901 WHERE Book IN (" & ItemCode & ")) AND C1.OrderDate<='" & GetDate(Format(DateAdd("d", -365, CDate(OrderDate)), "dd-mm-yyyy")) & "' " & _
'                             "UNION ALL " & _
'                             "SELECT P.Name As OrderNo,C.OrderDate,M1.PrintName As PrinterName,C.Processing,C.PlateType,C.ActualQuantity As Quantity,C.PlateRate As Rate,C.BillNo,C.BillDate,M2.PrintName As ItemName " & _
'                             "FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.TitlePrinter=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code " & _
'                             "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND C.Code IN (SELECT Code FROM BookPOParent WHERE Book IN (" & ItemCode & ")) AND C.OrderDate<='" & GetDate(Format(DateAdd("d", -365, CDate(OrderDate)), "dd-mm-yyyy")) & "' " & _
'                             "ORDER BY PrinterName,OrderNo", cnImporter, adOpenKeyset, adLockReadOnly
        Else
            rstImporter.Open "SELECT P.Name As OrderNo,C.OrderDate,M1.PrintName As PrinterName,C.Processing,C.PlateType" & PlateType & ",C.ActualQuantity As Quantity,C.PlateRate" & PlateType & " As Rate,C.BillNo,C.BillDate,M2.PrintName As ItemName " & _
                             "FROM ((BookPOParent P INNER JOIN BookPOChild" & TblSuffix & " C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P." & IIf(TblSuffix = "06", "TitlePrinter", "BookPrinter") & "=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code " & _
                             "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND M2.Code IN (" & ItemCode & ") AND C.OrderDate>='" & GetDate(Format(DateAdd("d", -365, CDate(OrderDate)), "dd-mm-yyyy")) & "' " & _
                             "ORDER BY M1.PrintName,P.Name", cnImporter, adOpenKeyset, adLockReadOnly
        End If
        rstImporter.ActiveConnection = Nothing
    End If
    If TblSuffix = "06" Then
            rstPSPlateRegister.Open "SELECT P.Name As OrderNo,C.OrderDate,A.Name As PartyName,C.Processing," & IIf(PlateType = "F", "C.PlateType", "C.PlateTypeBack") & " As PlateType,C.ActualQuantity As Quantity,C.PlateRate As Rate,C.BillNo,C.BillDate,I.Name As ItemName,E.Name As ElementName " & _
                                                                "FROM (((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster A ON P.TitlePrinter=A.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code " & _
                                                                "WHERE P.Type='" & OrderType & "' AND C.Element='" & ElementCode & "' AND P.Book='" & ItemCode & "' AND C.Code<'" & IIf(OrderCode = "", "999999", OrderCode) & "' AND C.OrderDate<='" & GetDate(Format(OrderDate, "dd-mm-yyyy")) & "' " & _
                                                                "ORDER BY A.Name,P.Name", cnDatabase, adOpenKeyset, adLockReadOnly
'    Else
'        rstPSPlateRegister.Open "SELECT P.Name As OrderNo,C1.OrderDate,M1.PrintName As PrinterName,C1.Plate As Processing,C1.PlateType,C2.ActualQuantity As Quantity,C1.PlateRate As Rate,C1.BillNo,C1.BillDate,TRIM(M2.PrintName)+' (Combo)' As ItemName " & _
'                                "FROM (((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C2 ON C1.Code=C2.Code) INNER JOIN AccountMaster M1 ON P.TitlePrinter=M1.Code) INNER JOIN BookMaster M2 ON C2.Book=M2.Code " & _
'                                "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND C2.Code IN (SELECT Code FROM BookPOChild0901 WHERE Book IN (" & ItemCode & ")) AND C1.Code<'" & IIf(OrderCode = "", "999999", OrderCode) & "' AND C1.OrderDate<='" & GetDate(Format(OrderDate, "dd-mm-yyyy")) & "' " & _
'                                "UNION ALL " & _
'                                "SELECT P.Name As OrderNo,C.OrderDate,M1.PrintName As PrinterName,C.Processing,C.PlateType,C.ActualQuantity As Quantity,C.PlateRate As Rate,C.BillNo,C.BillDate,M2.PrintName As ItemName " & _
'                                "FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.TitlePrinter=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code " & _
'                                "WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND C.Code IN (SELECT Code FROM BookPOParent WHERE Book IN (" & ItemCode & ")) AND C.Code<'" & IIf(OrderCode = "", "999999", OrderCode) & "' AND C.OrderDate<='" & GetDate(Format(OrderDate, "dd-mm-yyyy")) & "' " & _
'                                "ORDER BY PrinterName,OrderNo", cnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstPSPlateRegister.Open "SELECT P.Name As OrderNo,C.OrderDate,M1.PrintName As PartyName,C.Processing,C.PlateType" & PlateType & ",C.ActualQuantity As Quantity,C.PlateRate" & PlateType & " As Rate,C.BillNo,C.BillDate,M2.PrintName As ItemName " & _
                                "FROM ((BookPOParent P INNER JOIN BookPOChild" & TblSuffix & " C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P." & IIf(TblSuffix = "06", "TitlePrinter", "BookPrinter") & "=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code " & _
                                "WHERE P.Type='" & OrderType & "' AND LEFT(P.Code,1)<>'*' AND M2.Code IN (" & ItemCode & ") AND C.Code<'" & IIf(OrderCode = "", "999999", OrderCode) & "' AND C.OrderDate<='" & GetDate(Format(OrderDate, "dd-mm-yyyy")) & "' " & _
                                "ORDER BY M1.PrintName,P.Name", cnDatabase, adOpenKeyset, adLockReadOnly
    End If
    rstPSPlateRegister.ActiveConnection = Nothing
    i = 1
    If Not CheckEmpty(LYDBName, False) Then
        If rstImporter.RecordCount > 0 Then rstImporter.MoveFirst
        Do While Not rstImporter.EOF
            fpSpread1.SetText 1, i, Trim(rstImporter.Fields("OrderNo").Value)
            fpSpread1.SetText 2, i, Format(rstImporter.Fields("OrderDate").Value, "dd-mm-yyyy")
            fpSpread1.SetText 3, i, Trim(rstImporter.Fields("PartyName").Value)
            fpSpread1.SetText 4, i, IIf(rstImporter.Fields("Processing").Value = "O", "", "New")
            If TblSuffix = "06" Then
                fpSpread1.SetText 5, i, IIf(rstImporter.Fields("PlateType").Value = "1", "Deepatch", IIf(rstImporter.Fields("PlateType").Value = "2", "PS", IIf(rstImporter.Fields("PlateType").Value = "3", "Wipeon", "CTP")))
            Else
                fpSpread1.SetText 5, i, IIf(rstImporter.Fields("PlateType" & PlateType).Value = "1", "Deepatch", IIf(rstImporter.Fields("PlateType" & PlateType).Value = "2", "PS", IIf(rstImporter.Fields("PlateType" & PlateType).Value = "3", "Wipeon", "CTP")))
            End If
            fpSpread1.SetText 6, i, Val(rstImporter.Fields("Quantity").Value)
            fpSpread1.SetText 7, i, Val(rstImporter.Fields("Rate").Value)
            fpSpread1.SetText 8, i, Trim(rstImporter.Fields("BillNo").Value)
            fpSpread1.SetText 9, i, Format(rstImporter.Fields("BillDate").Value, "dd-mm-yyyy")
            i = i + 1
            rstImporter.MoveNext
        Loop
    End If
    If rstPSPlateRegister.RecordCount > 0 Then rstPSPlateRegister.MoveFirst
    Do While Not rstPSPlateRegister.EOF
        fpSpread1.SetText 1, i, Trim(rstPSPlateRegister.Fields("OrderNo").Value)
        fpSpread1.SetText 2, i, Format(rstPSPlateRegister.Fields("OrderDate").Value, "dd-mm-yyyy")
        fpSpread1.SetText 3, i, Trim(rstPSPlateRegister.Fields("PartyName").Value)
        fpSpread1.SetText 4, i, IIf(rstPSPlateRegister.Fields("Processing").Value = "O", "", "New")
        If TblSuffix = "06" Then
            fpSpread1.SetText 5, i, IIf(rstPSPlateRegister.Fields("PlateType").Value = "1", "Deepatch", IIf(rstPSPlateRegister.Fields("PlateType").Value = "2", "PS", IIf(rstPSPlateRegister.Fields("PlateType").Value = "3", "Wipeon", "CTP")))
        Else
            fpSpread1.SetText 5, i, IIf(rstPSPlateRegister.Fields("PlateType" & PlateType).Value = "1", "Deepatch", IIf(rstPSPlateRegister.Fields("PlateType" & PlateType).Value = "2", "PS", IIf(rstPSPlateRegister.Fields("PlateType" & PlateType).Value = "3", "Wipeon", "CTP")))
        End If
        fpSpread1.SetText 6, i, Val(rstPSPlateRegister.Fields("Quantity").Value)
        fpSpread1.SetText 7, i, Val(rstPSPlateRegister.Fields("Rate").Value)
        fpSpread1.SetText 8, i, Trim(rstPSPlateRegister.Fields("BillNo").Value)
        fpSpread1.SetText 9, i, Format(rstPSPlateRegister.Fields("BillDate").Value, "dd-mm-yyyy")
        i = i + 1
        rstPSPlateRegister.MoveNext
    Loop
    Screen.MousePointer = vbNormal
    Exit Sub
ErrorHandler:
    Screen.MousePointer = vbNormal
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        cmdExit_Click
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstPSPlateRegister)
    Call CloseConnection(cnImporter)
End Sub
Private Sub cmdExit_Click()
    Call CloseForm(Me)
End Sub
