VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmItemOpBal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Centrewise Item Opening Bal"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8430
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
   ScaleHeight     =   7980
   ScaleWidth      =   8430
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   7775
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   105
      Width           =   8210
      _Version        =   65536
      _ExtentX        =   14482
      _ExtentY        =   13714
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
      Picture         =   "ItemOpBal.frx":0000
      Begin VB.CommandButton cmdExit 
         Height          =   375
         Left            =   7725
         Picture         =   "ItemOpBal.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Exit"
         Top             =   80
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   375
         Left            =   7370
         Picture         =   "ItemOpBal.frx":011E
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Refresh"
         Top             =   80
         Width           =   375
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   375
         Left            =   7010
         Picture         =   "ItemOpBal.frx":0268
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Delete"
         Top             =   80
         Width           =   375
      End
      Begin VB.CommandButton cmdSave 
         Height          =   375
         Left            =   6650
         Picture         =   "ItemOpBal.frx":036A
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Save"
         Top             =   80
         Width           =   375
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
         TabIndex        =   0
         Top             =   105
         Width           =   5415
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   6
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
         Caption         =   " Mat Centre"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "ItemOpBal.frx":046C
         Picture         =   "ItemOpBal.frx":0488
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   7035
         Left            =   120
         TabIndex        =   1
         Top             =   630
         Width           =   7980
         _Version        =   524288
         _ExtentX        =   14076
         _ExtentY        =   12409
         _StockProps     =   64
         ButtonDrawMode  =   8
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
         MaxCols         =   3
         MaxRows         =   1000
         ScrollBars      =   2
         SpreadDesigner  =   "ItemOpBal.frx":04A4
      End
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
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   8
         Top             =   630
         Width           =   5175
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8190
         Y1              =   530
         Y2              =   530
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   4
      Left            =   6000
      Top             =   3720
   End
End
Attribute VB_Name = "FrmItemOpBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnItemChild As New ADODB.Connection, rstBalList As New ADODB.Recordset
Dim MatCentreCode As String, EditMode As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    DisableCloseButton Me
    cnItemChild.CursorLocation = adUseClient: cnItemChild.Open cnDatabase.ConnectionString
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If InStr(1, "fpSpread1", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True Else EditMode = False
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If Not EditMode Then
            If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then Me.ActiveControl.SetFocus Else CloseForm Me
        End If
        If Not EditMode Then KeyCode = 0
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2) Then
        If Not EditMode Then cmdSave_Click
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        cmdRefresh_Click
        KeyCode = 0
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstBalList)
    Call CloseConnection(cnItemChild)
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        With FrmAccountMaster
            .SL = True
            .AccountType = "01": .AccountGroup = "*99999"
            .MasterCode = MatCentreCode
            Load FrmAccountMaster
            If Err.Number <> 364 Then .Show vbModal
        End With
        On Error GoTo 0
        MatCentreCode = slCode: Text2.Text = slName
        If Not CheckEmpty(MatCentreCode, False) Then Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If CheckEmpty(Text2.Text, False) Then Cancel = True: Exit Sub
    On Error GoTo ErrorHandler
    Dim i As Integer
    With rstBalList
        If .State = adStateOpen Then .Close
        .Open "SELECT Item As ItemId,I.Name As Item,OpBal FROM BookChild O INNER JOIN BookMaster I ON O.Item=I.Code WHERE O.MaterialCentre='" & MatCentreCode & "' AND FYCode='" & FYCode & "' ORDER BY I.Name", cnDatabase, adOpenKeyset, adLockReadOnly
        .ActiveConnection = Nothing
        If .RecordCount > 0 Then .MoveFirst
        fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
        fpSpread1.SetActiveCell 1, 1
        Do Until .EOF
            i = i + 1
            fpSpread1.SetText 1, i, .Fields("Item").Value
            fpSpread1.SetText 2, i, Val(.Fields("OpBal").Value)
            fpSpread1.SetText 3, i, .Fields("ItemId").Value
            .MoveNext
        Loop
    End With
    fpSpread1.SetActiveCell 1, 1
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Item As Variant
    With fpSpread1
        If .EditMode Then Exit Sub
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF9) Then
            If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then
                .GetText 5, .ActiveRow, Item
                On Error Resume Next
                FrmBookMaster.SL = True
                FrmBookMaster.ItemType = "F"
                FrmBookMaster.MasterCode = Item
                Load FrmBookMaster
                If Err.Number <> 364 Then FrmBookMaster.Show vbModal
                On Error GoTo 0
                .SetText .ActiveCol, .ActiveRow, slName: .SetText 3, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then Sendkeys "{ENTER}"
            End If
        End If
    End With
End Sub
Private Sub cmdSave_Click()
    If CheckEmpty(Text2.Text, False) Then Exit Sub
    Dim CellVal(2) As Variant, i As Integer
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    cnItemChild.BeginTrans
    cnItemChild.Execute "DELETE FROM BookChild WHERE MaterialCentre='" & MatCentreCode & "' AND FYCode='" & FYCode & "'"
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 2, i, CellVal(1) 'Op Bal
            .GetText 3, i, CellVal(2) 'Item
            cnItemChild.Execute "INSERT INTO BookChild VALUES ('" & MatCentreCode & "','" & CellVal(2) & "'," & Val(CellVal(1)) & ",'" & FYCode & "')"
        Next
    End With
    MdiMainMenu.MousePointer = vbNormal
    cnItemChild.CommitTrans
    ShowProgressInStatusBar True
    Timer1.Enabled = True
    cmdRefresh_Click
    Exit Sub
ErrorHandler:
    MdiMainMenu.MousePointer = vbNormal
    DisplayError (Err.Description)
    cnItemChild.RollbackTrans
End Sub
Private Sub cmdDelete_Click()
    If CheckEmpty(Text2.Text, False) Then Exit Sub
    If AllowMastersDeletion = 0 Then Call DisplayError("You don't have the rights to delete this Master"): Exit Sub
    If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
        On Error Resume Next
        cnItemChild.BeginTrans
        cnItemChild.Execute "DELETE FROM BookChild WHERE MaterialCentre='" & MatCentreCode & "' AND FYCode='" & FYCode & "'"
        If Err.Number = 0 Then
            cnItemChild.CommitTrans
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            cmdRefresh_Click
        Else
            cnItemChild.RollbackTrans
            DisplayError (Err.Description)
        End If
        On Error GoTo 0
    End If
End Sub
Private Sub cmdRefresh_Click()
    Text2.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    Text2.SetFocus
End Sub
Private Sub cmdExit_Click()
    Form_KeyDown vbKeyEscape, 0
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then Timer1.Enabled = False: ShowProgressInStatusBar False
End Sub
