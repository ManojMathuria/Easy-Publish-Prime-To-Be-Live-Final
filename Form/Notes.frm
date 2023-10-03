VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form frmNotes 
   BackColor       =   &H00FFFEF2&
   ClientHeight    =   7890
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14160
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   14160
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   12600
      Picture         =   "Notes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Save"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   12960
      Picture         =   "Notes.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Delete"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   375
      Left            =   13320
      Picture         =   "Notes.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Refresh"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Height          =   375
      Left            =   13680
      Picture         =   "Notes.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Exit"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7200
      TabIndex        =   4
      Top             =   7440
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13935
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   6315
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   13425
         _cx             =   23680
         _cy             =   11139
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   4
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   0   'False
         SubtotalPosition=   0
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.Timer Timer1 
            Interval        =   4
            Left            =   960
            Top             =   240
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6315
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   13410
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   " Notes: Easy Info Solutions International"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   13920
      End
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
      Height          =   330
      Left            =   1200
      TabIndex        =   5
      Top             =   7440
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   582
      Calculator      =   "Notes.frx":0450
      Caption         =   "Notes.frx":0470
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Notes.frx":04D4
      Keys            =   "Notes.frx":04F2
      Spin            =   "Notes.frx":053C
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
      TabIndex        =   6
      Top             =   7440
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   8421376
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
      FillColor       =   8421376
      ShadowColor     =   0
      TextColor       =   16777215
      Picture         =   "Notes.frx":0564
      Picture         =   "Notes.frx":0580
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NotesFlag As Long, BalFlag As Boolean, EditMode As Boolean, rCount As Long
Dim nSort As Boolean, VSFlexFlag As Boolean, FontFlag As Boolean
Dim rstClientAccount As New ADODB.Recordset
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    CenterForm Me
    BusySystemIndicator True
    VSFlexGrid1.Visible = False
    If DatabaseType = "MS SQL" Then
        cnClientAccount.CommandTimeout = 300
        ConnectionString = "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=Company ;Data Source=" & ServerName
        If cnClientAccount.State = 1 Then cnClientAccount.Close
        cnClientAccount.Open ConnectionString
    End If
    If NotesFlag = 1 Then Text1.Text = FrmAccountMaster.txtNotes.Text
    If NotesFlag = 2 Then Text1.Text = FrmBookMaster.txtNotes.Text
    If NotesFlag = 3 Then Text1.Text = frmDebitCreditVoucher.txtNotes.Text
    If NotesFlag = 4 Then Text1.Text = frmSalesChallanVoucher.txtNotes.Text
    If NotesFlag = 5 Then Text1.Text = frmSalesOrderVoucher.txtNotes.Text
    If NotesFlag = 6 Then Text1.Text = frmSalesVoucher.txtNotes.Text
    If NotesFlag = 7 Then Text1.Text = frmItemIssueReceiptVoucher.txtNotes.Text
    If rstClientAccount.State = 1 Then rstClientAccount.Close
    rstClientAccount.Open "SELECT SuM(Debit)-Sum(Credit) As Bal FROM ClientAccount Where Left(UUID,5)='EP'+'" & CompCode & "'", cnClientAccount, adOpenKeyset, adLockReadOnly
    If ((Trim(ReadFromFile("Client User")) = "EasyPublish") Or (rstClientAccount.Fields("Bal").Value <> "Null" And rstClientAccount.Fields("Bal").Value <> 0 And BalFlag = True)) And NotesFlag = 0 Then
            VSFlexGrid1.Visible = True
            btnSave.Visible = False
            Call PublishGrid
    ElseIf IsNull(rstClientAccount.Fields("Bal").Value) And BalFlag = True Then
        If Trim(ReadFromFile("Client User")) = "EasyPublish" Then
        
        Else
            BalFlag = False
            VSFlexGrid1.Visible = False
            btnSave.Visible = False
            TDBNumber2.Visible = False
            Mh3dLabel4.Visible = False
        End If
    ElseIf IsNull(rstClientAccount.Fields("Bal").Value) And (Val(rstClientAccount.Fields("Bal").Value)) = 0 And BalFlag = True Then
        If Trim(ReadFromFile("Client User")) = "EasyPublish" Then
        
        Else
            BalFlag = False
            VSFlexGrid1.Visible = False
            btnSave.Visible = False
            TDBNumber2.Visible = False
            Mh3dLabel4.Visible = False
        End If
    Else
        VSFlexGrid1.Visible = False
        cmdDelete.Visible = False
        cmdSave.Visible = False
        cmdExit.Visible = False
        cmdRefresh.Visible = False
        TDBNumber2.Visible = False
        Mh3dLabel4.Visible = False
    End If
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If InStr(1, "VSFlexGrid1", Me.ActiveControl.Name) > 0 And VSFlexGrid1.Visible Then If Me.ActiveControl.Editable > flexEDNone Then EditMode = True Else EditMode = False
    If InStr(1, "VSFlexGrid1", Me.ActiveControl.Name) > 0 And VSFlexGrid1.Visible Then EditMode = True Else EditMode = False
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If Not EditMode Then
            If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then Me.ActiveControl.SetFocus Else CloseForm Me
        End If
        If Not EditMode Then KeyCode = 0
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2) Then
        If Not EditMode Then UpdateRecords
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        cmdRefresh_Click
        KeyCode = 0
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    NotesFlag = 0: BalFlag = False
    Call CloseRecordset(rstClientAccount)
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
       If Shift = 0 And KeyCode = vbKeyEscape Then
            If Not EditMode Then
                If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then Me.ActiveControl.SetFocus: Call btnSave_Click Else CloseForm Me
            End If
            If Not EditMode Then KeyCode = 0
            'Call btnSave_Click: KeyCode = 0
        ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2) Then
        Call btnSave_Click: KeyCode = 0
        ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF9) Then
        Call cmdDelete_Click: KeyCode = 0
        End If
End Sub
Private Sub btnSave_Click()
    Call cmdSave_Click
    Unload Me
End Sub
Private Sub cmdSave_Click()
    Call rowCount
If BalFlag = False Or NotesFlag > 0 Then
    SaveFields
Else
    Call UpdateRecords
End If
    Call MsgBox("Records updated !!!", vbInformation, App.Title)
    If Not BalFlag Then Unload Me
End Sub
Private Sub SaveFields()
    If NotesFlag = 1 Then FrmAccountMaster.txtNotes.Text = Text1.Text
    If NotesFlag = 2 Then FrmBookMaster.txtNotes.Text = Text1.Text
    If NotesFlag = 3 Then frmDebitCreditVoucher.txtNotes.Text = Text1.Text
    If NotesFlag = 4 Then frmSalesChallanVoucher.txtNotes.Text = Text1.Text
    If NotesFlag = 5 Then frmSalesOrderVoucher.txtNotes.Text = Text1.Text
    If NotesFlag = 6 Then frmSalesVoucher.txtNotes.Text = Text1.Text
    If NotesFlag = 7 Then frmItemIssueReceiptVoucher.txtNotes.Text = Text1.Text
    EditMode = False
End Sub
Private Sub UpdateRecords()
    If VSFlexGrid1.TextMatrix(i, 2) = "" Then Exit Sub
    Dim CellVal(9) As Variant
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    cnClientAccount.BeginTrans
    cnClientAccount.Execute "DELETE FROM ClientAccount Where Left(UUID,5)='EP'+'" & CompCode & "'"
    With VSFlexGrid1
    
        For i = 1 To rCount
            CellVal(1) = "EP" + CompCode 'UUID"
            CellVal(2) = GetDate(.TextMatrix(i, 1)) 'VchDate
            CellVal(3) = Trim(.TextMatrix(i, 2)) 'Particulars
            CellVal(4) = .TextMatrix(i, 3) 'Debit
            CellVal(5) = .TextMatrix(i, 4) 'Credit
            CellVal(6) = Trim(.TextMatrix(i, 7)) 'Remarks
            CellVal(7) = "NULL" '.TextMatrix(i, 8) 'Document
            CellVal(8) = "NULL" '.TextMatrix(i, 9) 'Document Type
            CellVal(9) = "AGPL" 'ClientName
    
            cnClientAccount.Execute "INSERT INTO ClientAccount VALUES ('" & CellVal(1) & "','" & CellVal(2) & "','" & CellVal(3) & "'," & Val(CellVal(4)) & "," & Val(CellVal(5)) & ",'" & CellVal(6) & "',Convert(varbinary(MAX),'" & CellVal(7) & "'),'" & CellVal(8) & "','" & CellVal(9) & "')"
        Next
    
    End With
    MdiMainMenu.MousePointer = vbNormal
    cnClientAccount.CommitTrans
    ShowProgressInStatusBar True
    Timer1.Enabled = True
    cmdRefresh_Click
    Exit Sub
ErrorHandler:
    MdiMainMenu.MousePointer = vbNormal
    DisplayError (Err.Description)
    cnClientAccount.RollbackTrans
End Sub
Private Sub cmdRefresh_Click()
Call PublishGrid
    Call MsgBox("Record Refresh !!!", vbInformation, App.Title)
If Not Text1.Visible Then Exit Sub
If Not VSFlexGrid1.Visible Then
    If NotesFlag = 1 Then Text1.Text = FrmAccountMaster.txtNotes.Text
    If NotesFlag = 2 Then Text1.Text = FrmBookMaster.txtNotes.Text
    If NotesFlag = 3 Then Text1.Text = frmDebitCreditVoucher.txtNotes.Text
    If NotesFlag = 4 Then Text1.Text = frmSalesChallanVoucher.txtNotes.Text
    If NotesFlag = 5 Then Text1.Text = frmSalesOrderVoucher.txtNotes.Text
    If NotesFlag = 6 Then Text1.Text = frmSalesVoucher.txtNotes.Text
    If NotesFlag = 7 Then Text1.Text = frmItemIssueReceiptVoucher.txtNotes.Text
End If

End Sub
Private Sub cmdDelete_Click()
If Not Text1.Visible Then Exit Sub
Text1.Text = ""
    Call MsgBox("Record Deleted !!!", vbInformation, App.Title)
End Sub
Private Sub cmdExit_Click()
    Call CloseRecordset(rstClientAccount)
    Form_KeyDown vbKeyEscape, 0
End Sub
Private Function PublishGrid()
Dim Credit, Debit, Bal As Double
Dim i, dPrint As Long
    If rstClientAccount.State = 1 Then rstClientAccount.Close
    rstClientAccount.Open "SELECT * FROM ClientAccount Where Left(UUID,5)='EP'+'" & CompCode & "'  Order By VchDate ", cnClientAccount, adOpenKeyset, adLockReadOnly
If rstClientAccount.RecordCount = 0 Then GoTo n
With VSFlexGrid1
    .Clear
        .Cols = 8
        .Rows = rstClientAccount.RecordCount + 1
        rstClientAccount.MoveFirst
        Do While Not rstClientAccount.EOF
        i = i + 1
                .TextMatrix(i, 0) = Trim(rstClientAccount.Fields("UUID").Value)
                .TextMatrix(i, 1) = Format(rstClientAccount.Fields("VchDate").Value, "dd-MM-yyyy")
                .TextMatrix(i, 2) = Trim(rstClientAccount.Fields("Particulars").Value)
                .TextMatrix(i, 3) = rstClientAccount.Fields("Debit").Value
                Debit = rstClientAccount.Fields("Debit").Value
                .TextMatrix(i, 4) = rstClientAccount.Fields("Credit").Value
                Credit = rstClientAccount.Fields("Credit").Value
                Bal = Bal + Debit - Credit
                .TextMatrix(i, 5) = Bal
                .TextMatrix(i, 6) = IIf(Bal > 0, "Dr.", "Cr."): If .TextMatrix(i, 5) < 0 Then .Cell(flexcpForeColor, i, 6) = vbRed Else .Cell(flexcpForeColor, i, 5) = vbBlack
                If Not IsNull(Trim(rstClientAccount.Fields("Remarks").Value)) Then
                .TextMatrix(i, 7) = Trim(Trim(rstClientAccount.Fields("Remarks").Value))
                Else
                .TextMatrix(i, 7) = ""
                End If
            dPrint = dPrint + 1
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updated record # " & dPrint & " of " & rstClientAccount.RecordCount & " !!!"
            TDBNumber2 = dPrint
                rstClientAccount.MoveNext
            If MdiMainMenu.ProgressBar1.Value + Round((100 / rstClientAccount.RecordCount), 2) <= 100 Then
                MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + Round((100 / rstClientAccount.RecordCount), 2)
                
            End If
            Loop
            .Rows = i + 1
'        Set VSFlexGrid1.DataSource = rstClientAccount
End With
n:
Call VSFlexGrid_Format_Headers
Call VSFlexGrid1_AfterDataRefresh
    Timer1.Enabled = False
    ShowProgressInStatusBar False
    MdiMainMenu.MousePointer = vbNormal
    Screen.MousePointer = vbNormal
    Exit Function
ErrHandler:
    Timer1.Enabled = False
    ShowProgressInStatusBar False
    MdiMainMenu.MousePointer = vbNormal
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Function
Private Function VSFlexGrid_Format_Headers()
    On Error GoTo ErrHandler
       ' Set VSFlexGrid1.DataSource = rstClientAccount
    With VSFlexGrid1
        Dim i As Long
        .Cols = 8
        .Rows = 50
        .ColWidth(0) = 1100
        .ColWidth(1) = 950
        .TextMatrix(i, 1) = "Date"
        .ColWidth(2) = 3000
        .TextMatrix(i, 2) = "Particulars"
        .ColWidth(3) = 880
        .TextMatrix(i, 3) = "Debit"
        .ColWidth(4) = 880
        .TextMatrix(i, 4) = "Credit"
        .ColWidth(5) = 880
        .TextMatrix(i, 5) = "Balance"
        .TextMatrix(i, 6) = "Dr./Cr."
        .ColWidth(7) = 3090
        .TextMatrix(i, 7) = "Commitment Remarks"
        Text1.Visible = False
        If Trim(ReadFromFile("Client User")) <> "EasyPublish" Then
        .Editable = flexEDNone
        End If
    End With
    Screen.MousePointer = vbNormal
Exit Function
ErrHandler:
    Screen.MousePointer = vbNormal
    DisplayError (Err.Description)
End Function
Private Sub VSFlexGrid1_AfterDataRefresh()
Dim C As Variant
Dim T As Long
Dim GroupOn As Long

nSort = False
'Subtotal
With VSFlexGrid1
.SubtotalPosition = flexSTBelow
If nSort = True Then
    .MultiTotals = True
    .Subtotal flexSTClear
        For C = 1 To .Cols - 1
        Err.Number = 0
        On Error Resume Next
        T = .TextMatrix(1, C)
        If Err.Number = 0 Then
            If InStr(1, "Item_MRP_PRICE_RATE_Balance", .TextMatrix(0, C)) > 0 Then
                .Subtotal flexSTAverage, 1, C, "(#,##0)", RGB(240, 230, 247), RGB(128, 0, 64), True, "Sub Total", , True
                For T = 1 To .Rows - 1
                If .TextMatrix(T, 3) = "" Then .TextMatrix(T, 3) = "Sub Total"
                Next
            Else
                .Subtotal flexSTSum, GroupOn, C, "(#,##0)", RGB(240, 230, 247), RGB(128, 0, 64), True, "Sub Total", GroupOn, True
                .Subtotal flexSTSum, GroupOn, C, "(#,##0)", RGB(241, 248, 248), vbRed, True, "Grand Total", GroupOn, True
            End If
        End If
        Next
    nSort = False
ElseIf nSort = False Then
    .MultiTotals = True
    .Subtotal flexSTClear
        For C = 1 To .Cols - 1
        Err.Number = 0
        On Error Resume Next
        T = .TextMatrix(1, C)
        If Err.Number = 0 Then
            If InStr(1, "Item_MRP_PRICE_RATE_Balance", .TextMatrix(0, C)) > 0 Then
        
            Else

                .Subtotal flexSTSum, GroupOn, C, "(#,##0)", RGB(240, 230, 247), vbBlue, True, "Grand Total", GroupOn, True
                .Subtotal flexSTSum, GroupOn, C, "(#,##0)", RGB(240, 230, 247), vbBlue, True, "Grand Total", GroupOn, True
            End If
        End If
        Next
    nSort = True
End If
    For C = 1 To (.Cols - 1)
'        .AutoSize C
        .ExplorerBar = flexExSort
        .ColSort(C) = flexSortCustom
        .AllowUserResizing = flexResizeBoth
    Next
    .Col = 7
End With
End Sub
Private Sub VSFlexGrid1_AfterSort(ByVal Col As Long, Order As Integer)
Dim C As Variant
Dim T As Long
VSFlexGrid1.SubtotalPosition = flexSTBelow
With VSFlexGrid1

If nSort = True Then
    .MultiTotals = True
    .Subtotal flexSTClear
        
        For C = 1 To .Cols - 1
        Err.Number = 0
        On Error Resume Next
        T = .TextMatrix(1, C)
        If Err.Number = 0 Then
            If InStr(1, "Item_MRP_PRICE_RATE_Balance", .TextMatrix(0, C)) > 0 Then
                '.Subtotal flexSTAverage, .Col, C, "(#,##0)", RGB(240, 230, 247), RGB(128, 0, 64), True, "Sub Total", .Col, True
                
'                For T = 1 To .Rows - 1
'                If .TextMatrix(T, 3) = "" Then .TextMatrix(T, 3) = "Sub Total"
'                Next
            
            Else
                .Subtotal flexSTSum, .Col, C, "(#,##0)", RGB(240, 230, 247), RGB(128, 0, 64), True, "Sub Total", .Col, True
                .Subtotal flexSTSum, 1, C, "(#,##0)", RGB(240, 230, 247), vbBlue, True, "Grand Total", 1, True
            End If
        End If
        Next
        
    nSort = False

ElseIf nSort = False Then
    .MultiTotals = True
    .Subtotal flexSTClear
        For C = 1 To .Cols - 1
        Err.Number = 0
        On Error Resume Next
        T = .TextMatrix(1, C)
        If Err.Number = 0 Then
            If InStr(1, "Item_MRP_PRICE_RATE_Balance", .TextMatrix(0, C)) > 0 Then
        
            Else
                .Subtotal flexSTSum, 1, C, "(#,##0)", RGB(240, 230, 247), vbBlue, True, "Grand Total", 1, True
            End If
        End If
        Next
    nSort = True
End If
    
    For C = 1 To (.Cols - 1)
        .AutoSize C
        .ExplorerBar = flexExSort
        .ColSort(C) = flexSortCustom
        .AllowUserResizing = flexResizeBoth
    Next
End With
End Sub
Sub rowCount()
Dim n, i As Long
If Not VSFlexGrid1.Visible Then Exit Sub
VSFlexGrid1.Subtotal flexSTClear
rCount = 0
    i = 1
    For i = 1 To VSFlexGrid1.Rows - 1
        If VSFlexGrid1.TextMatrix(i, 2) <> "" Then
            rCount = rCount + 1
        End If
    Next
End Sub
Private Sub VSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim nextRow As Integer
        Dim nextCol As Integer
        Dim PreRow As Integer
        Dim PreCol As Integer
        
        ' Determine the current cell coordinates
        Dim currentRow As Integer
        Dim currentCol As Integer
        currentRow = VSFlexGrid1.Row
        currentCol = VSFlexGrid1.Col
        ' Calculate the next cell coordinates
    If Shift = 0 And KeyCode = vbKeyReturn Then
        nextRow = currentRow
        nextCol = currentCol + 1
        ' Check if the next cell is within the grid's bounds
        If nextCol >= VSFlexGrid1.Cols Then
            ' Move to the next row if the last column is reached
            nextRow = currentRow + 1
            nextCol = 0
        End If
        ' Set the focus to the next cell
        VSFlexGrid1.Row = nextRow
        VSFlexGrid1.Col = nextCol
    ElseIf Shift = 0 And KeyCode = vbKeyUp Then
        PreRow = currentRow - 1
        PreCol = currentCol
        ' Check if the next cell is within the grid's bounds
        If currentRow <= 1 Then
            ' Move to the next row if the last column is reached
            PreRow = VSFlexGrid1.Rows - 1
            PreCol = currentCol - 1
        End If
        If currentCol <= 1 And currentRow <= 1 Then
            PreCol = VSFlexGrid1.Cols - 1
        End If
        ' Set the focus to the next cell
        VSFlexGrid1.Row = PreRow
        VSFlexGrid1.Col = PreCol
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        nextRow = currentRow + 1
        nextCol = currentCol
        ' Check if the next cell is within the grid's bounds
        If nextRow >= VSFlexGrid1.Rows Then
            ' Move to the next row if the last column is reached
            nextRow = 1
            nextCol = currentCol + 1
        End If
        If nextCol >= VSFlexGrid1.Cols Then
            nextCol = 1
        End If
        ' Set the focus to the next cell
        VSFlexGrid1.Row = nextRow: SetFocus
        VSFlexGrid1.Col = nextCol
    ElseIf Shift = 0 And KeyCode = vbKeyRight Then
        nextRow = currentRow
        nextCol = currentCol + 1
        ' Check if the next cell is within the grid's bounds
        If nextCol >= VSFlexGrid1.Cols Then
            nextRow = currentRow + 1
            nextCol = 1
        End If
        If nextRow >= VSFlexGrid1.Rows Then
            nextRow = 1
        End If
        
        ' Set the focus to the next cell
        VSFlexGrid1.Row = nextRow
        VSFlexGrid1.Col = nextCol
    ElseIf Shift = 0 And KeyCode = vbKeyLeft Then
        PreRow = currentRow
        PreCol = currentCol - 1
        ' Check if the next cell is within the grid's bounds
        If currentCol <= 1 Then
            PreCol = VSFlexGrid1.Cols - 1
            PreRow = currentRow - 1
        End If
        If currentRow <= 1 And currentCol <= 1 Then
            PreRow = VSFlexGrid1.Rows - 1
        End If
        
        ' Set the focus to the next cell
        VSFlexGrid1.Row = PreRow
        VSFlexGrid1.Col = PreCol
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        If Not EditMode Then
            If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then Me.ActiveControl.SetFocus Else CloseForm Me
        End If
        If Not EditMode Then KeyCode = 0
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2) Then
        Call cmdSave_Click
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyC) Or (Shift = 0 And KeyCode = vbKeyF12) Then
        Call CopyToClipboard
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyV) Or (Shift = 0 And KeyCode = vbKeyF12) Then
        Call PasteFromClipboard
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        cmdRefresh_Click
    End If
        KeyCode = 0
End Sub
Private Sub CopyToClipboard()
    Dim selectedData As String
    Dim i As Integer

    ' Get the selected data from the grid
    For i = VSFlexGrid1.RowSel To VSFlexGrid1.Row
        selectedData = selectedData & VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColSel) & vbCrLf
    Next i

    ' Copy the selected data to the clipboard
    Clipboard.SetText selectedData
End Sub
Private Sub PasteFromClipboard()
    Dim clipboardData As String
    Dim dataRows() As String
    Dim i As Integer

    ' Get the data from the clipboard
    clipboardData = Clipboard.GetText

    ' Split the clipboard data into individual rows
    dataRows = Split(clipboardData, vbCrLf)

    ' Paste the data into the grid
    For i = 0 To UBound(dataRows)
        VSFlexGrid1.TextMatrix(VSFlexGrid1.Row + i, VSFlexGrid1.Col) = dataRows(i)
    Next i
End Sub


