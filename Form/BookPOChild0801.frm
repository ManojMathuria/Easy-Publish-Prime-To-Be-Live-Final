VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmBookPOChild0801 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BOM Details"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17160
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
   ScaleHeight     =   7695
   ScaleWidth      =   17160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008BD6FE&
      Height          =   375
      Left            =   16095
      Picture         =   "BookPOChild0801.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      BackColor       =   &H008BD6FE&
      Height          =   375
      Left            =   16095
      Picture         =   "BookPOChild0801.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Proceed"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   7485
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   105
      Width           =   15375
      _Version        =   65536
      _ExtentX        =   27120
      _ExtentY        =   13203
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
      Picture         =   "BookPOChild0801.frx":0204
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   105
         Width           =   13695
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   6735
         Left            =   120
         TabIndex        =   0
         Top             =   645
         Width           =   15135
         _Version        =   524288
         _ExtentX        =   26696
         _ExtentY        =   11880
         _StockProps     =   64
         ButtonDrawMode  =   1
         EditEnterAction =   5
         EditModePermanent=   -1  'True
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
         MaxCols         =   11
         MaxRows         =   100
         OperationMode   =   2
         SpreadDesigner  =   "BookPOChild0801.frx":0220
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   105
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
         Caption         =   " Item Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "BookPOChild0801.frx":0C3A
         Picture         =   "BookPOChild0801.frx":0C56
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   15400
         Y1              =   540
         Y2              =   540
      End
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   1170
      Index           =   2
      Left            =   15540
      TabIndex        =   6
      Top             =   960
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   2064
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
      Caption         =   " Ctrl+E->Edit Row  F2->Edit Row  F9->Delete Row  Ctrl+D->Delete Row  Ctrl+S->Save"
      AutoSize        =   -1  'True
      FillColor       =   8421504
      TextColor       =   16777215
      Picture         =   "BookPOChild0801.frx":0C72
      Multiline       =   -1  'True
      GlobalMem       =   -1  'True
      Picture         =   "BookPOChild0801.frx":0C8E
   End
End
Attribute VB_Name = "FrmBookPOChild0801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstBookPOChild0801 As New ADODB.Recordset, BinderCode As String, BookCode As String, OrderCode As Variant
Dim rstOutsourceItemList As New ADODB.Recordset, rstPaperList As New ADODB.Recordset, rstVendorList As New ADODB.Recordset, rstBOMItemList As New ADODB.Recordset, OutsourceItem As String, Paper As String, Vendor As String, EditMode As Boolean
Private Sub Form_Load()
    Dim i As Integer
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    DisableCloseButton Me
    AbortPO = False
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text)
    rstOutsourceItemList.Open "SELECT Name,'1'+Code As NCode FROM OutsourceItemMaster UNION ALL SELECT Name,'3'+Code As NCode FROM BookMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockOptimistic
    rstPaperList.Open "SELECT Name,'2'+Code As NCode FROM PaperMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockOptimistic
    rstVendorList.Open "SELECT Name,Code FROM AccountMaster ORDER BY Name", cnDatabase, adOpenKeyset, adLockOptimistic
    rstOutsourceItemList.ActiveConnection = Nothing: rstPaperList.ActiveConnection = Nothing: rstVendorList.ActiveConnection = Nothing
    Call RefreshDropDownList("A")
    With fpSpread1
        .ClearRange 1, 1, .MaxCols, .MaxRows, True
        If CheckNull(OrderCode) = "" Then
            With rstBOMItemList
                If .State = adStateOpen Then .Close
                .Open "SELECT Category,Item,Quantity FROM BookChild01 WHERE Code='" & BookCode & "'", cnDatabase, adOpenKeyset, adLockOptimistic
                .ActiveConnection = Nothing
                If .RecordCount > 0 Then
                    .MoveFirst
                    i = 0
                    Do Until .EOF
                        i = i + 1
                        fpSpread1.SetText 1, i, IIf(Val(.Fields("Category").Value) = 2, "Paper", "BOM Item")
                        fpSpread1.Col = 2: fpSpread1.TypeComboBoxList = IIf(Val(.Fields("Category").Value) = 2, Paper, OutsourceItem)
                        fpSpread1.Col = 9: fpSpread1.TypeComboBoxList = Vendor
                        If .Fields("Category").Value = "2" Then fpSpread1.SetText 6, i, "Sheets" Else fpSpread1.SetText 6, i, "Nos."
                        If .Fields("Category").Value = "1" Or .Fields("Category").Value = "3" Then
                           If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
                           rstOutsourceItemList.Find "[NCode]='" & FixQuote(.Fields("Category").Value + .Fields("Item").Value) & "'"
                           If Not rstOutsourceItemList.EOF Then fpSpread1.SetText 2, i, rstOutsourceItemList.Fields("Name").Value
                        ElseIf .Fields("Category").Value = "2" Then
                           If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
                           rstPaperList.Find "[NCode]='" & FixQuote(.Fields("Category").Value + .Fields("Item").Value) & "'"
                           If Not rstPaperList.EOF Then fpSpread1.SetText 2, i, rstPaperList.Fields("Name").Value
                        End If
                        fpSpread1.SetText 4, i, Val(.Fields("Quantity").Value)
                        fpSpread1.SetText 10, i, .Fields("Category").Value + .Fields("Item").Value
                        .MoveNext
                    Loop
                End If
            End With
        End If
        With rstBookPOChild0801
            If .RecordCount > 0 Then
                .MoveFirst
                i = 0
                Do Until .EOF
                    i = i + 1
                    fpSpread1.SetText 1, i, IIf(Val(.Fields("Category").Value) = 2, "Paper", "BOM Item")
                    fpSpread1.Col = 2: fpSpread1.TypeComboBoxList = IIf(Val(.Fields("Category").Value) = 2, Paper, OutsourceItem)
                    fpSpread1.Col = 9: fpSpread1.TypeComboBoxList = Vendor
                    If .Fields("Category").Value = "2" Then fpSpread1.SetText 6, i, "Sheets" Else fpSpread1.SetText 6, i, "Nos."
                    If .Fields("Category").Value = "1" Or .Fields("Category").Value = "3" Then
                       If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
                       rstOutsourceItemList.Find "[NCode]='" & FixQuote(.Fields("Category").Value + .Fields("Item").Value) & "'"
                       If Not rstOutsourceItemList.EOF Then fpSpread1.SetText 2, i, rstOutsourceItemList.Fields("Name").Value
                    ElseIf .Fields("Category").Value = "2" Then
                       If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
                       rstPaperList.Find "[NCode]='" & FixQuote(.Fields("Category").Value + .Fields("Item").Value) & "'"
                       If Not rstPaperList.EOF Then fpSpread1.SetText 2, i, rstPaperList.Fields("Name").Value
                    End If
                    fpSpread1.SetText 3, i, Val(.Fields("OrderQuantity").Value)
                    fpSpread1.SetText 4, i, Val(.Fields("Consumption/Item").Value)
                    fpSpread1.SetText 5, i, Val(.Fields("TotalConsumption").Value)
                    fpSpread1.SetText 6, i, IIf(.Fields("Category").Value = "2", "Sheets", "Nos.")
                    fpSpread1.SetText 7, i, Val(.Fields("Rate").Value)
                    fpSpread1.SetText 8, i, Val(.Fields("Amount").Value)
                    If rstVendorList.RecordCount > 0 Then rstVendorList.MoveFirst
                    rstVendorList.Find "[Code]='" & FixQuote(.Fields("Vendor").Value) & "'"
                    If Not rstVendorList.EOF Then fpSpread1.SetText 9, i, rstVendorList.Fields("Name").Value
                    fpSpread1.SetText 10, i, .Fields("Category").Value + .Fields("Item").Value
                    fpSpread1.SetText 11, i, .Fields("Vendor").Value
                    .MoveNext
                Loop
            End If
        End With
    End With
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}": KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        If Not EditMode Then cmdProceed_Click: KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        If Not EditMode Then cmdCancel_Click: KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstVendorList)
    Call CloseRecordset(rstBOMItemList)
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckItem() Then
       fpSpread1.SetFocus
       CheckMandatoryFields = True
    End If
End Function
Private Sub cmdProceed_Click()
    If CheckMandatoryFields Then Exit Sub
    SaveFields
    Call CloseForm(Me)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Sub SaveFields()
    Dim i As Integer, Item As Variant, ConsumptionPerItem As Variant, OrderQuantity As Variant, TotalConsumption As Variant, Vendor As Variant, Rate As Variant, Amount As Variant
    With rstBookPOChild0801
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            .Delete adAffectCurrent
            .MoveNext
        Loop
    End With
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 3, i, OrderQuantity
            .GetText 4, i, ConsumptionPerItem
            .GetText 5, i, TotalConsumption
            .GetText 7, i, Rate
            .GetText 8, i, Amount
            If Val(TotalConsumption) > 0 Then
                .GetText 10, i, Item
                .GetText 11, i, Vendor
                With rstBookPOChild0801
                    .AddNew
                    .Fields("Category").Value = Left(Item, 1)
                    .Fields("Item").Value = Right(Item, 6)
                    .Fields("Consumption/Item").Value = ConsumptionPerItem
                    .Fields("OrderQuantity").Value = OrderQuantity
                    .Fields("TotalConsumption").Value = TotalConsumption
                    .Fields("Rate").Value = Rate
                    .Fields("Amount").Value = Amount
                    .Fields("Vendor").Value = Vendor
                    .Update
                End With
            End If
        Next
    End With
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF9) Then
        If UserLevel = "3" Then Call DisplayError("You don't have the rights to delete BOM Item"): Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1
            fpSpread1.SetFocus
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        Call RefreshDropDownList("R")
    End If
End Sub
Private Sub fpSpread1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    Dim Vendor As Variant
    fpSpread1.GetText 9, Row, Vendor
    If Not CheckEmpty(Vendor, False) Then
        If rstVendorList.RecordCount > 0 Then rstVendorList.MoveFirst
        rstVendorList.Find "[Name]='" & FixQuote(Vendor) & "'"
        If Not rstVendorList.EOF Then fpSpread1.SetText 11, Row, rstVendorList.Fields("Code").Value
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant, Category As Variant, ConsumptionPerItem As Variant, OrderQuantity As Variant, Rate As Variant
    fpSpread1.GetText Col, Row, ActiveCellVal
    fpSpread1.GetText 1, Row, Category
    If Col = 1 Then
        fpSpread1.Col = 2: fpSpread1.TypeComboBoxList = IIf(Category = "BOM Item", OutsourceItem, Paper)
        fpSpread1.Col = 9: fpSpread1.TypeComboBoxList = Vendor
        If Category = "Paper" Then fpSpread1.SetText 6, Row, "Sheets" Else fpSpread1.SetText 6, Row, "Nos."
    ElseIf Col = 2 Then
        If Category = "BOM Item" Then
           If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
           rstOutsourceItemList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstOutsourceItemList.EOF Then fpSpread1.SetText 10, Row, rstOutsourceItemList.Fields("NCode").Value
        ElseIf Category = "Paper" Then
           If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
           rstPaperList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstPaperList.EOF Then fpSpread1.SetText 10, Row, rstPaperList.Fields("NCode").Value
        End If
    ElseIf Col = 4 Or Col = 5 Or Col = 7 Then
        fpSpread1.GetText 3, Row, OrderQuantity
        fpSpread1.GetText 4, Row, ConsumptionPerItem
        fpSpread1.SetText 5, Row, Val(ConsumptionPerItem) * Val(OrderQuantity)
        fpSpread1.GetText 7, Row, Rate
        fpSpread1.SetText 8, Row, Val(Rate) * Val(ConsumptionPerItem) * Val(OrderQuantity)
    End If
End Sub
Private Function CheckItem() As Boolean
    Dim i As Integer, Item As Variant, Category As Variant, Qty As Variant, BalanceQuantity As Double, VchDate As Date
    CheckItem = False
    VchDate = FrmBookPrintOrder.MhDateInput1.Value
    With fpSpread1
        For i = 1 To .DataRowCnt
            .SetActiveCell 1, i
            .GetText 10, i, Item
            .GetText 1, i, Category
            .GetText 5, i, Qty
            If Not CheckEmpty(Category, False) Then
                If Category = "BOM Item" Then
                    If Left(Item, 1) <> "1" And Left(Item, 1) <> "3" Then CheckItem = True
                ElseIf Category = "Paper" Then
                    If Left(Item, 1) <> "2" Then CheckItem = True
                End If
                If CheckItem Then DisplayError "Data mismatch in row #" & Trim(Str(i)): Exit For
            End If
            If Category = "Paper" Then BalanceQuantity = CalculatePaperBalance(BinderCode, Right(Item, 6), CheckNull(OrderCode), "BPOB", VchDate) Else BalanceQuantity = CalculateMaterialBalance(BinderCode, Category, Right(Item, 6), CheckNull(OrderCode), "PO")
            If Left(FrmBookPrintOrder.BookPOType, 1) <> "O" Then
                If Val(Qty) > Round(Val(BalanceQuantity), 3) Then
                    If UserLevel <= 2 Then
                        If MsgBox("Stock (" & Format(Val(BalanceQuantity - Val(Qty)), "0.000") & ") of the Item in row #" & Trim(Str(i)) & " is going negative ! Would you like to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbNo Then CheckItem = True: Exit Function
                    Else
                        Call DisplayError("Cann't Save ! Stock (" & Format(Val(BalanceQuantity - Val(Qty)), "0.000") & ") of the Item in row #" & Trim(Str(i)) & " is going negative"): AbortPO = True: Exit Function
                    End If
                End If
            End If
        Next
    End With
End Function
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub RefreshDropDownList(ByVal xType As String)
    If xType = "R" Then
        rstOutsourceItemList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstOutsourceItemList): Loop
        rstOutsourceItemList.ActiveConnection = Nothing
        rstPaperList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstPaperList): Loop
        rstPaperList.ActiveConnection = Nothing
        rstVendorList.ActiveConnection = cnDatabase
        Do While Not RefreshRecord(rstVendorList): Loop
        rstVendorList.ActiveConnection = Nothing
        OutsourceItem = "": Paper = "": Vendor = ""
    End If
    With rstOutsourceItemList
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If OutsourceItem = "" Then OutsourceItem = .Fields("Name").Value Else OutsourceItem = OutsourceItem + Chr$(9) + .Fields("Name").Value
            .MoveNext
        Loop
    End With
    With rstPaperList
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If Paper = "" Then Paper = .Fields("Name").Value Else Paper = Paper + Chr$(9) + .Fields("Name").Value
            .MoveNext
        Loop
    End With
    With rstVendorList
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If Vendor = "" Then Vendor = .Fields("Name").Value Else Vendor = Vendor + Chr$(9) + .Fields("Name").Value
            .MoveNext
        Loop
    End With
End Sub
