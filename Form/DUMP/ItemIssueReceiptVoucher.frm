VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemIssueReceiptVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Issue Receipt Voucher"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13740
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   13740
End
Attribute VB_Name = "frmItemIssueReceiptVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Vch Type=NNNNFI/NNNNUI/NNNNFR/NNNNUR (F-Finished U-Unfinished I-Issue R-Receipt) & BOM=NNNNXXXXXXXXXXXXFI/NNNNXXXXXXXXXXXXMF (MF/ME/CF/MO/BN/BM) & 05-Purchase Challan 06-Purchase Return Challan 07-Sales Return Challan 08-Sales Challan
Public VchType As String 'R-Item Receipt I-Item Issue
Dim cnDeliveryChallan As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset, rstDeliveryCVList As New ADODB.Recordset, rstDeliveryCVParent As New ADODB.Recordset, rstDeliveryCVChild As New ADODB.Recordset, rstAccountList As New ADODB.Recordset, rstTaxList As New ADODB.Recordset, rstItemList As New ADODB.Recordset, rstNarrationList As New ADODB.Recordset, rstHSNCodeList As New ADODB.Recordset, rstOrderList As New ADODB.Recordset, rstVchSeriesList As New ADODB.Recordset, rstMaterialCentreList As New ADODB.Recordset
Dim PartyCode As String, ConsigneeCode As String, TaxCode As String, ItemCode As String, RefCode As String, NarrationCode As String, HSNCode As String, MaterialCentreCode As String, VchPrefix As String, TranType As String, VchNumbering As String, VchSeriesCode As String, AutoVchNo As String, oVchSeriesCode As String, oVchNo As String
Dim SortOrder, PrevStr, dblBookMark As Double, blnRecordExist As Boolean, EditMode As Boolean
Private Sub btnNotes_Click()
    frmNotes.NotesFlag = 7
    frmNotes.Label1.Caption = "Notes : Voucher No. : " & Text2.Text
    frmNotes.Show (vbModal)
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    Me.Caption = "Item " & IIf(VchType = "I", "Issued to", "Received from") & " Party Voucher"
    Mh3dLabel12.Caption = IIf(VchType = "I", " Mat Centre", " Mat Centre")
    Mh3dLabel14.Caption = IIf(VchType = "I", " Issue", " Receipt") & " Type"
    DataGrid1.Columns(3).Caption = IIf(VchType = "I", "Consignee Name", "Material Centre")
    TranType = IIf(VchType = "I", "08", "05")
    cnDeliveryChallan.CursorLocation = adUseClient
    cnDeliveryChallan.Open cnDatabase.ConnectionString
    LoadMasterList
    rstDeliveryCVList.Open "SELECT T.Code,LTRIM(T.Name) As Name,V.Code As VchSeriesCode,V.Name As VchSeriesName,Date,T.Type,M1.Name As PartyName,M2.Name As MaterialCentreName,ChallanNo,ChallanDate,Amount,T.AutoVchNo FROM (JobworkBVParent T INNER JOIN AccountMaster M1 ON T.Party=M1.Code) INNER JOIN AccountMaster M2 ON " & IIf(VchType = "I", "T.Consignee", "T.MaterialCentre") & "=M2.Code INNER JOIN VchSeriesMaster V ON T.VchSeries=V.Code WHERE LEFT(Type,2) IN (" & IIf(VchType = "R", "'05','07'", "'06','08'") & ") AND RIGHT(Type,1)='" & VchType & "' AND FYCode='" & FYCode & "' ORDER BY T.AutoVchNo", cnDeliveryChallan, adOpenKeyset, adLockPessimistic
    rstDeliveryCVParent.CursorLocation = adUseClient
    rstDeliveryCVList.Filter = adFilterNone
    If rstDeliveryCVList.RecordCount > 0 Then rstDeliveryCVList.MoveLast
    Set DataGrid1.DataSource = rstDeliveryCVList
    BusySystemIndicator False
    SSTab1.Tab = 0
    If FrmStockLedger.dSortBy = True Then
    SortOrder = "Code"
    Else
    SortOrder = "AutoVchNo"
    End If
    If Not (rstDeliveryCVList.EOF Or rstDeliveryCVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstDeliveryCVList.ActiveConnection = Nothing
    cmbItemType.AddItem "Finished Goods", 0
    cmbItemType.AddItem "Unfinished Goods", 1
    cmbChallanType.AddItem IIf(VchType = "I", "Sale", "Purchase") & " Challan", 0
    cmbChallanType.AddItem IIf(VchType = "I", "Purchase Return", "Sale Return") & " Challan", 1
'    cmbChallanType.AddItem "Challan Reversal", 2
'    cmbChallanType.AddItem "To be Billed", 3
'    cmbChallanType.AddItem "Not to be Billed", 4
    SetButtonsForNoRecord
    fpSpread1.TextTip = TextTipFloating
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    With MdiMainMenu
        .mnuMaterialOutJobWork.Enabled = False: .mnuMaterialInJobWork.Enabled = False
    End With
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        EditMode = False
        If SSTab1.Tab = 0 Then  'List
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If InStr(1, "fpSpread1", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True
                If Not EditMode Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
            If Not EditMode Then KeyCode = 0
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyD) Or (Shift = 0 And KeyCode = vbKeyF8)) And Toolbar1.Buttons.Item(3).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf ((Shift = vbCtrlMask And KeyCode = vbKeyS) Or (Shift = 0 And KeyCode = vbKeyF2)) And Toolbar1.Buttons.Item(4).Enabled Then 'Save
        EditMode = False
        If InStr(1, "fpSpread1", Me.ActiveControl.Name) > 0 Then If Me.ActiveControl.EditMode Then EditMode = True
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(9)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(10)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyM And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(11)
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
           If Me.ActiveControl.Name <> "fpSpread1" Then Sendkeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstDeliveryCVList)
    Call CloseRecordset(rstDeliveryCVParent)
    Call CloseRecordset(rstDeliveryCVChild)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstTaxList)
    Call CloseRecordset(rstNarrationList)
    Call CloseRecordset(rstHSNCodeList)
    Call CloseRecordset(rstItemList)
    Call CloseRecordset(rstOrderList)
    Call CloseConnection(cnDeliveryChallan)
    ShowProgressInStatusBar False
    DisableChildMenu
    MdiMainMenu.mnuMaterialOutJobWork.Enabled = True
    MdiMainMenu.mnuMaterialInJobWork.Enabled = True
End Sub
Private Sub Text1_Change()
    'If FrmStockLedger.dSortBy = True Then SortOrder = "Code" Else SortOrder = "AutoVchNo"
    On Error Resume Next
    With rstDeliveryCVList
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
    If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstDeliveryCVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstDeliveryCVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstDeliveryCVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstDeliveryCVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstDeliveryCVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstDeliveryCVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstDeliveryCVList
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
    If Toolbar1.Buttons.Item(1).Enabled Then
        If SSTab1.Tab = 1 Then
            ViewRecord
        Else
            If Not (rstDeliveryCVList.EOF Or rstDeliveryCVList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            Text1.SetFocus
        End If
        SSTab1.TabEnabled(0) = True
    Else
        SSTab1.TabEnabled(0) = False
        Text10.SetFocus
    End If
End Sub
Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Dim SearchString As String
        SearchString = FixQuote(Text10.Text)
        If rstVchSeriesList.RecordCount = 0 Then DisplayError ("No Record in Voucher Series Master"): Text10.SetFocus: Exit Sub Else rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Col0] = '" & RTrim(SearchString) & "'"
        SelectionType = "S": VchSeriesCode = ""
        Call LoadSelectionList(rstVchSeriesList, "List of Voucher Series...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text10, VchSeriesCode)
        Call CloseForm(FrmSelectionList)
        If RTrim(VchSeriesCode) <> "" Then Sendkeys "{TAB}" Else Text10.Text = ""
    End If
End Sub
Private Sub Text10_Validate(Cancel As Boolean)
    If CheckEmpty(Text10.Text, False) Then
        Cancel = True
    Else
        rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        VchNumbering = rstVchSeriesList.Fields("VchNumbering").Value
        If VchNumbering = "A" Then Text2.Locked = True Else Text2.Locked = False
        If Not blnRecordExist Then 'Vch-New
            If VchNumbering = "A" Then
                AutoVchNo = GenerateCode(cnDeliveryChallan, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & "F" + VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
            End If
        Else 'Vch-Old
            If VchSeriesCode = oVchSeriesCode Then
                Text2.Text = Text2.Text 'oVchNo
            Else
                If VchNumbering = "A" Then
                    AutoVchNo = GenerateCode(cnDeliveryChallan, "SELECT MAX(" & IIf(DatabaseType = "MS SQL", "CONVERT(INT,AutoVchNo))", "VAL(AutoVchNo))") & "  FROM  JobworkBVParent WHERE RIGHT(Type,2)='" & "F" + VchType & "' AND VchSeries='" & VchSeriesCode & "' AND FYCode='" & FYCode & "'", 10, Space(1))
                    Text2.Text = Trim(rstVchSeriesList.Fields("Prefix").Value) + Trim(AutoVchNo) + Trim(rstVchSeriesList.Fields("Suffix").Value)
                End If
            End If
        End If
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    Dim CellVal As Variant, i As Integer
    If Button.Index = 1 Then
        If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
        rstDeliveryCVParent.Open "SELECT * FROM JobworkBVParent WHERE Code=''", cnDeliveryChallan, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstDeliveryCVParent) Then
            Text2.Text = GenerateCode(cnDeliveryChallan, "SELECT MAX(CONVERT(INT,AutoVchNo)) FROM JobworkBVParent WHERE Right(Type,2)='" & "F" + VchType & "' AND FYCode='" & FYCode & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text10.SetFocus
            blnRecordExist = False
            cnDeliveryChallan.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnDeliveryChallan.BeginTrans
            With rstDeliveryCVChild
                If .State = adStateOpen Then
                    If .RecordCount > 0 Then .MoveFirst
                    Do While Not .EOF
                        If Not CheckEmpty(.Fields("VchCode").Value, False) Then Call UpdateStatus(.Fields("VchCode").Value, .Fields("Quantity").Value, "-")
                        .MoveNext
                    Loop
                End If
            End With
            cnDeliveryChallan.Execute "DELETE FROM JobworkBVParent WHERE Code='" & rstDeliveryCVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstDeliveryCVList.Delete
                rstDeliveryCVList.MoveNext
                If rstDeliveryCVList.RecordCount > 0 And rstDeliveryCVList.EOF Then rstDeliveryCVList.MoveLast
                cnDeliveryChallan.CommitTrans
                ShowProgressInStatusBar True
                Timer1.Enabled = True
                Text1.Text = ""
                rstDeliveryCVList.Filter = adFilterNone
            Else
                DisplayError (Err.Description)
                cnDeliveryChallan.RollbackTrans
            End If
            On Error GoTo 0
        End If
        SetButtons (True)
        SetButtonsForNoRecord
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 4 Then
        If CheckMandatoryFields Then Exit Sub
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstDeliveryCVParent) Then
            If UpdateItemList("D", 0) Then
                UpdateFlag = 1
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 1, i
                        .GetText 4, i, CellVal
                        If Val(CellVal) <> 0 Then
                            If Not UpdateItemList("I", i) Then UpdateFlag = 0: Exit For
                        End If
                    Next
                End With
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnDeliveryChallan.CommitTrans
            If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
            rstDeliveryCVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            Call MsgBox("Record updated !!!", vbInformation, App.Title)
            SSTab1.Tab = 0
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstDeliveryCVParent) Then
            cnDeliveryChallan.RollbackTrans
            If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
            rstDeliveryCVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstDeliveryCVList.Filter = adFilterNone
        rstDeliveryCVList.ActiveConnection = cnDeliveryChallan
        Do While Not RefreshRecord(rstDeliveryCVList): Loop
        Set DataGrid1.DataSource = rstDeliveryCVList
        rstDeliveryCVList.ActiveConnection = Nothing
        If rstDeliveryCVList.RecordCount > 0 Then rstDeliveryCVList.MoveLast
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem IIf(VchType = "P", "Material Centre", "Consignee"), 0
            .Combo1.AddItem "Party", 1
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
        DisplayMenu "P"
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
        DisplayMenu "S"
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstDeliveryCVList.RecordCount > 0 Then rstDeliveryCVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstDeliveryCVList.RecordCount > 0 Then
            rstDeliveryCVList.MovePrevious
            If rstDeliveryCVList.BOF Then rstDeliveryCVList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstDeliveryCVList.RecordCount > 0 Then
            rstDeliveryCVList.MoveNext
            If rstDeliveryCVList.EOF Then rstDeliveryCVList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstDeliveryCVList.RecordCount > 0 Then rstDeliveryCVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstDeliveryCVList.EOF Or rstDeliveryCVList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        Text1.SetFocus
    End If
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstDeliveryCVList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstDeliveryCVList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstDeliveryCVList.EOF Or rstDeliveryCVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub SetButtons(bVal As Boolean)
    With Toolbar1.Buttons
        .Item(1).Enabled = bVal
        .Item(2).Enabled = bVal
        .Item(3).Enabled = bVal
        .Item(4).Enabled = Not bVal
        .Item(5).Enabled = Not bVal
        .Item(6).Enabled = bVal
        .Item(7).Enabled = bVal
        .Item(9).Enabled = bVal
        .Item(10).Enabled = bVal
        .Item(11).Enabled = bVal
        .Item(13).Enabled = bVal
        .Item(14).Enabled = bVal
        .Item(15).Enabled = bVal
        .Item(16).Enabled = bVal
        .Item(18).Enabled = bVal
    End With
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstDeliveryCVList.RecordCount = 0 Then
        With Toolbar1.Buttons
            .Item(2).Enabled = False
            .Item(3).Enabled = False
            .Item(9).Enabled = False
            .Item(10).Enabled = False
            .Item(11).Enabled = False
            .Item(13).Enabled = False
            .Item(14).Enabled = False
            .Item(15).Enabled = False
            .Item(16).Enabled = False
        End With
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstDeliveryCVParent.EOF Or rstDeliveryCVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnDeliveryChallan, "JobworkBVParent", "Code", "[Name]+RIGHT([Type],1)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, rstDeliveryCVParent.Fields("Code").Value, False, FYCode) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = PartyCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        PartyCode = slCode: Text3.Text = slName
        If Not CheckEmpty(PartyCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If CheckEmpty(Text3.Text, False) Then Cancel = True
    If CheckEmpty(Text8.Text, False) Then Text8.Text = Text3.Text: ConsigneeCode = PartyCode
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = IIf(VchType = "I", "*99999", "*99999")
        FrmAccountMaster.MasterCode = MaterialCentreCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        MaterialCentreCode = slCode: Text7.Text = slName
        If Not CheckEmpty(MaterialCentreCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
    If CheckEmpty(Text7.Text, False) Then Cancel = True
End Sub
Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmAccountMaster.SL = True
        FrmAccountMaster.AccountType = "01": FrmAccountMaster.AccountGroup = ""
        FrmAccountMaster.MasterCode = ConsigneeCode
        Load FrmAccountMaster
        If Err.Number <> 364 Then FrmAccountMaster.Show vbModal
        On Error GoTo 0
        ConsigneeCode = slCode: Text8.Text = slName
        If Not CheckEmpty(ConsigneeCode, False) Then LoadMasterList: Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text8_Validate(Cancel As Boolean)
    If CheckEmpty(Text8.Text, False) Then Cancel = True
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        On Error Resume Next
        FrmTaxMaster.SL = True
        FrmTaxMaster.MasterCode = TaxCode
        Load FrmTaxMaster
        If Err.Number <> 364 Then FrmTaxMaster.Show vbModal
        On Error GoTo 0
        TaxCode = slCode: Text5.Text = slName
        If Not CheckEmpty(TaxCode, False) Then
            rstTaxList.MoveFirst: rstTaxList.Find "[Code] = '" & TaxCode & "'"
            If Val(rstTaxList.Fields("SGST%").Value) > 0 Then   'Intra-State GST
                MhRealInput7.Value = Val(rstTaxList.Fields("CGST%").Value)
                MhRealInput9.Value = Val(rstTaxList.Fields("SGST%").Value)
            Else    'Inter-State GST
                MhRealInput7.Value = Val(rstTaxList.Fields("IGST%").Value)
                MhRealInput9.Value = 0
            End If
            CalculateTotal
            LoadMasterList
            Sendkeys "{TAB}"
        End If
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    If CheckEmpty(Text5.Text, False) Then Cancel = True
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)    'Discount
    CalculateTotal
End Sub
Private Sub MhRealInput6_Validate(Cancel As Boolean)    'Freight
    CalculateTotal
End Sub
Private Sub MhRealInput12_Validate(Cancel As Boolean)   'Adjustment
    CalculateTotal
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstDeliveryCVList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
    rstDeliveryCVParent.Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstDeliveryCVList.Fields("Code").Value) & "'", cnDeliveryChallan, adOpenKeyset, adLockOptimistic
    If rstDeliveryCVParent.RecordCount = 0 Then Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset"): Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
End Sub
Private Sub ClearFields()
    Text2.Text = "" 'Vch No.
    Text9.Text = "" 'Challan No
    MhDateInput2.Text = "  -  -    " 'Challan Date
    MhRealInput13.Value = 0 'Box
    Text3.Text = "": PartyCode = "" 'Party Name
    Text7.Text = "": MaterialCentreCode = "" 'Material Centre Name
    Text5.Text = "": TaxCode = "" 'Tax Name
    Text4.Text = "" 'Remarks
    cmbItemType.ListIndex = 0: cmbItemType.Enabled = True
    cmbChallanType.ListIndex = 0: cmbChallanType.Enabled = True: cmbChallanType_Click
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhRealInput1.Value = 0
    MhRealInput2.Value = 0
    MhRealInput3.Value = 0
    MhRealInput4.Value = 0
    MhRealInput5.Value = 0
    MhRealInput6.Value = 0
    MhRealInput12.Value = 0
    MhRealInput7.Value = 0
    MhRealInput8.Value = 0
    MhRealInput9.Value = 0
    MhRealInput10.Value = 0
    MhRealInput11.Value = 0
    Text10.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    PartyCode = "": ConsigneeCode = "": MaterialCentreCode = "": TaxCode = "": VchSeriesCode = "": oVchSeriesCode = "": oVchNo = "": AutoVchNo = ""
End Sub
Private Sub LoadFields()
    With rstDeliveryCVParent
        If .EOF Or .BOF Then Exit Sub
        Text2.Text = Trim(.Fields("Name").Value)
        MhDateInput1.Text = Format(.Fields("Date").Value, "dd-MM-yyyy")
        Text9.Text = CheckNull(.Fields("ChallanNo").Value)
        If Not IsNull(.Fields("ChallanDate").Value) Then MhDateInput2.Text = Format(.Fields("ChallanDate").Value, "dd-MM-yyyy")
        MhRealInput13.Value = Val(.Fields("Box").Value)
       
       PartyCode = .Fields("Party").Value
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        rstAccountList.Find "[Code] = '" & PartyCode & "'"
        If Not rstAccountList.EOF Then Text3.Text = rstAccountList.Fields("Col0").Value
        
'        If VchType = "R" Then MaterialCentreCode = .Fields("MaterialCentre").Value Else MaterialCentreCode = .Fields("Consignee").Value
'        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
'        rstAccountList.Find "[Code] = '" & MaterialCentreCode & "'"
'        If Not rstAccountList.EOF Then Text7.Text = rstAccountList.Fields("Col0").Value
        
        MaterialCentreCode = .Fields("MaterialCentre").Value
        If rstMaterialCentreList.RecordCount > 0 Then rstMaterialCentreList.MoveFirst
        rstMaterialCentreList.Find "[Code] = '" & MaterialCentreCode & "'"
        If Not rstMaterialCentreList.EOF Then Text7.Text = rstMaterialCentreList.Fields("Col0").Value
        
        ConsigneeCode = .Fields("Consignee").Value
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        rstAccountList.Find "[Code] = '" & ConsigneeCode & "'"
        If Not rstAccountList.EOF Then Text8.Text = rstAccountList.Fields("Col0").Value
        
        TaxCode = .Fields("Tax").Value
        If rstTaxList.RecordCount > 0 Then rstTaxList.MoveFirst
        rstTaxList.Find "[Code] = '" & TaxCode & "'"
        If Not rstTaxList.EOF Then Text5.Text = rstTaxList.Fields("Col0").Value
        VchSeriesCode = .Fields("VchSeries").Value: oVchSeriesCode = VchSeriesCode
        If rstVchSeriesList.RecordCount > 0 Then rstVchSeriesList.MoveFirst
        rstVchSeriesList.Find "[Code] = '" & VchSeriesCode & "'"
        If Not rstVchSeriesList.EOF Then Text10.Text = rstVchSeriesList.Fields("Col0").Value
        Text4.Text = .Fields("Remarks").Value
        cmbItemType.ListIndex = IIf(Mid(.Fields("Type").Value, 5, 1) = "F", 0, 1)
        cmbChallanType.ListIndex = IIf(VchType = "I", IIf(Left(.Fields("Type").Value, 2) = "08", 0, IIf(Left(.Fields("Type").Value, 2) = "06", 1, IIf(Left(.Fields("Type").Value, 2) = "11", 2, IIf(Left(.Fields("Type").Value, 2) = "13", 3, 4)))), IIf(Left(.Fields("Type").Value, 2) = "05", 0, IIf(Left(.Fields("Type").Value, 2) = "07", 1, IIf(Left(.Fields("Type").Value, 2) = "12", 2, IIf(Left(.Fields("Type").Value, 2) = "14", 3, 4)))))
        MhRealInput4.Value = Val(.Fields("Rebate%").Value)
        MhRealInput5.Value = Val(.Fields("Rebate").Value)
        MhRealInput6.Value = Val(.Fields("Freight").Value)
        MhRealInput12.Value = Val(.Fields("Adjustment").Value)
        If Val(.Fields("SGST%").Value) > 0 Then  'Intra-State Supply
            MhRealInput7.Value = Val(.Fields("CGST%").Value)
            MhRealInput8.Value = Val(.Fields("CGST").Value)
            MhRealInput9.Value = Val(.Fields("SGST%").Value)
            MhRealInput10.Value = Val(.Fields("SGST").Value)
        Else    'Inter-State Supply
            MhRealInput7.Value = Val(.Fields("IGST%").Value)
            MhRealInput8.Value = Val(.Fields("IGST").Value)
        End If
        MhRealInput11.Value = Val(.Fields("Amount").Value)
        Call LoadItemList(.Fields("Code").Value)
    End With
    CalculateTotal
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstDeliveryCVParent.RecordCount = 0 Then Exit Sub
    If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
    rstDeliveryCVParent.CursorLocation = adUseServer
    rstDeliveryCVParent.Open "SELECT * FROM JobworkBVParent WHERE Code='" & FixQuote(rstDeliveryCVList.Fields("Code").Value) & "'", cnDeliveryChallan, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstDeliveryCVParent.Fields("RecordStatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    cmbItemType.Enabled = False
    cmbChallanType.Enabled = False
    Text10.SetFocus
    blnRecordExist = True
    cnDeliveryChallan.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    With rstDeliveryCVParent
        If .EOF Or .BOF Then Exit Sub
        If Not blnRecordExist Then
            .Fields("Code").Value = GenerateCode(cnDeliveryChallan, "SELECT MAX(Code) FROM JobworkBVParent", 6, "0")
            .Fields("CreatedBy").Value = UserCode
            .Fields("CreatedOn").Value = Now()
            .Fields("Recordstatus").Value = "N"
        Else
            .Fields("ModifiedBy").Value = UserCode
            .Fields("ModifiedOn").Value = Now()
            .Fields("Recordstatus").Value = "M"
        End If
        .Fields("VchSeries").Value = VchSeriesCode
        .Fields("AutoVchNo").Value = Pad(Trim(AutoVchNo), Space(1), 10, "L")
        .Fields("Name").Value = Trim(Text2.Text)
        .Fields("Date").Value = GetDate(MhDateInput1.Text)
        .Fields("ChallanNo").Value = Text9.Text
        If MhDateInput2.ValueIsNull Then .Fields("ChallanDate").Value = Null Else .Fields("ChallanDate").Value = GetDate(MhDateInput2.Text)
        .Fields("Box").Value = MhRealInput13.Value
        .Fields("Party").Value = PartyCode
        .Fields("Consignee").Value = ConsigneeCode
        .Fields("MaterialCentre").Value = MaterialCentreCode
        .Fields("Tax").Value = TaxCode
        .Fields("Remarks").Value = Trim(Text4.Text)
        .Fields("Rebate%").Value = MhRealInput4.Value
        .Fields("Rebate").Value = MhRealInput5.Value
        .Fields("Freight").Value = MhRealInput6.Value
        .Fields("Adjustment").Value = MhRealInput12.Value
        .Fields("TaxableAmount").Value = MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput12.Value
        If MhRealInput9.Value > 0 Then  'Intra-State Supply
            .Fields("CGST%").Value = MhRealInput7.Value
            .Fields("CGST").Value = MhRealInput8.Value
            .Fields("SGST%").Value = MhRealInput9.Value
            .Fields("SGST").Value = MhRealInput10.Value
            .Fields("IGST%").Value = 0
            .Fields("IGST").Value = 0
        Else    'Inter-State Supply
            .Fields("CGST%").Value = 0
            .Fields("CGST").Value = 0
            .Fields("SGST%").Value = 0
            .Fields("SGST").Value = 0
            .Fields("IGST%").Value = MhRealInput7.Value
            .Fields("IGST").Value = MhRealInput8.Value
        End If
        .Fields("Amount").Value = MhRealInput11.Value
        .Fields("Type").Value = VchPrefix & IIf(cmbItemType.ListIndex = 0, "F", "U") & VchType
        .Fields("FYCode").Value = FYCode
        .Fields("RecordStatus").Value = "N"
        .Fields("Notes").Value = txtNotes.Text
        .Fields("SalesType").Value = ""
    End With
End Sub
Private Sub AddToList()
    On Error Resume Next
    With rstDeliveryCVList
        .MoveFirst
        .Find "[Code] = '" & rstDeliveryCVParent.Fields("Code").Value & "'"
        If .EOF Then .AddNew
        .Fields("VchSeriesName").Value = Text10.Text
        .Fields("Code").Value = rstDeliveryCVParent.Fields("Code").Value
        .Fields("Name").Value = rstDeliveryCVParent.Fields("Name").Value
        .Fields("Date").Value = rstDeliveryCVParent.Fields("Date").Value
        .Fields("PartyName").Value = Trim(Text3.Text)
        .Fields("MaterialCentreName").Value = Trim(Text7.Text)
        .Fields("Consignee").Value = Trim(Text8.Text)
        .Fields("Type").Value = rstDeliveryCVParent.Fields("Type").Value
        .Fields("Amount").Value = MhRealInput11.Value
        .Fields("ChallanNo").Value = rstDeliveryCVParent.Fields("ChallanNo").Value
        .Fields("ChallanDate").Value = rstDeliveryCVParent.Fields("ChallanDate").Value
        .Update
        .Sort = SortOrder & " Asc"
        .Find "[Code] = '" & rstDeliveryCVParent.Fields("Code").Value & "'"
    End With
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        DisplayError ("Voucher No. cannot be blank"): Text2.SetFocus: CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(cnDeliveryChallan, "JobworkBVParent", "Code", "[Name]+RIGHT([Type],1)+VchSeries", Trim(Text2.Text) & VchType & VchSeriesCode, rstDeliveryCVParent.Fields("Code").Value, False, FYCode) Then
        If Not blnRecordExist Then
            Dim VchNo As String
            VchNo = GenerateCode(cnDeliveryChallan, "SELECT MAX(CONVERT(INT,Name)) FROM JobworkBVParent WHERE RIGHT(Type,1)='" & VchType & "' AND FYCode='" & FYCode & "'", 10, Space(1))
            If Trim(VchNo) <> Trim(Text2.Text) Then DisplayError ("Vch No. changed from " & Trim(Text2.Text) & " to " & Trim(VchNo))
            Text2.Text = VchNo: Exit Function
        Else
            Text2.SetFocus: CheckMandatoryFields = True: Exit Function
        End If
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text7.Text, False) Then
        Text7.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text8.Text, False) Then
        Text8.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text5.Text, False) Then
        Text5.SetFocus:   CheckMandatoryFields = True: Exit Function
    ElseIf fpSpread1.DataRowCnt = 0 Then
        DisplayError ("Blank Voucher cannot be saved"): fpSpread1.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf fpSpread1.DataRowCnt > 0 Then
        Dim i As Integer, CellVal As Variant
        With fpSpread1
                For i = 1 To .DataRowCnt
                    .GetText 7, i, CellVal
                    If CheckEmpty(CellVal, False) Then DisplayError ("Narration at row #" & Trim(Str(i)) & " is blank"): CheckMandatoryFields = True: .SetFocus: Exit Function
                Next
        End With
    End If
End Function
Private Sub LoadItemList(ByVal VchNo As String)
    Dim i As Integer, SQL As String
    On Error GoTo ErrorHandler
    If rstDeliveryCVChild.State = adStateOpen Then rstDeliveryCVChild.Close
    If cmbChallanType.ListIndex < 2 Then
        If cmbItemType.ListIndex = 0 Then 'Finished with ref
            SQL = "SELECT I.Code As ItemCode,I.Name As ItemName,H.Code As HSNCode,H.Name As HSNName,T.Quantity,T.Quantity+R.EstQty01-R.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Code As NarrationCode,N.Name As NarrationName,SrNo,CASE WHEN Ref IS NULL THEN '' ELSE R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) END As VchCode,CASE WHEN Ref IS NULL THEN '' ELSE LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) END As VchNo FROM (((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,14)=R.Code+'XXXXXXXXXXXXFI') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "'"
        Else 'Unfinished with Ref
'           SQL = "SELECT I.Code As ItemCode,I.Name+'_'+E.Name+'_Printing' As ItemName,H.Code As HSNCode,H.Name As HSNName,T.Quantity,T.Quantity+C.ActualQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild05 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+'XXXXXXMF') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE T.Code='" & VchNo & "'"
            SQL = "SELECT I.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.ActualQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,14)=R.Code+'XXXXXXXXXXXXMF') INNER JOIN BookPOChild05 C ON R.Code=C.Code) INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "'"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.ActualQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild06 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+'XXXXXXME') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code  WHERE T.Code='" & VchNo & "'"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+E.Name+'_Printing' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.ActualQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+E.Code+'XXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild0901 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Book+'XXXXXXCF') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN BookMaster E ON C.Book=E.Code WHERE T.Code='" & VchNo & "'"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+E.Name+'_'+O.Name As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.Quantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+E.Code+O.Code+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild07 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Element+C.Operation+'MO') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code WHERE T.Code='" & VchNo & "'"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_Binding' As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.ActualQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM ((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,14)=R.Code+'XXXXXXXXXXXXBN') INNER JOIN BookPOChild08 C ON R.Code=C.Code) INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "'"
            SQL = SQL + " UNION ALL "
            SQL = SQL + "SELECT I.Name+'_'+CASE WHEN C.Category='1' THEN O.Name WHEN C.Category='2' THEN P.Name ELSE U.Name END As ItemName,I.Code As ItemCode,H.Name As HSNName,H.Code As HSNCode,T.Quantity,T.Quantity+C.OrderQuantity-C.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Name As NarrationName,N.Code As NarrationCode,SrNo,R.Code+C.Item+'XXXXX'+C.Category+RIGHT(T.BOM,2) As VchCode,LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) As VchNo FROM (((((((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) INNER JOIN BookPOParent R ON T.Ref=R.Code) INNER JOIN BookPOChild0801 C ON T.Ref+SUBSTRING(T.BOM,5,14)=C.Code+C.Item+'XXXXX'+C.Category+'BM') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code) LEFT JOIN OutsourceItemMaster O ON C.Category+C.Item='1'+O.Code) LEFT JOIN PaperMaster P ON C.Category+C.Item='2'+P.Code) LEFT JOIN BookMaster U ON C.Category+C.Item='3'+U.Code WHERE T.Code='" & VchNo & "'"
        End If
    Else 'Finished without ref
        SQL = "SELECT I.Code As ItemCode,I.Name As ItemName,H.Code As HSNCode,H.Name As HSNName,T.Quantity,T.Quantity+R.EstQty01-R.DeliveredQuantityC As PendingQty,T.Rate,T.Amount,N.Code As NarrationCode,N.Name As NarrationName,SrNo,CASE WHEN Ref IS NULL THEN '' ELSE R.Code+'XXXXXXXXXXXX'+RIGHT(T.BOM,2) END As VchCode,CASE WHEN Ref IS NULL THEN '' ELSE LTRIM(R.Name)+'/'+RIGHT(R.Type,1)+'O/'+RIGHT(T.BOM,2) END As VchNo FROM (((JobworkBVChild T INNER JOIN BookMaster I ON T.Item=I.Code) LEFT JOIN BookPOParent R ON T.Ref+SUBSTRING(T.BOM,5,12)=R.Code+'XXXXXXXXXXXX') INNER JOIN GeneralMaster N ON T.Narration=N.Code) INNER JOIN GeneralMaster H ON T.HSNCode=H.Code WHERE T.Code='" & VchNo & "' AND RIGHT(T.BOM,2)='FI'"
    End If
    SQL = SQL + " ORDER BY SrNo"
    rstDeliveryCVChild.Open SQL, cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstDeliveryCVChild.ActiveConnection = Nothing
    If rstDeliveryCVChild.RecordCount > 0 Then rstDeliveryCVChild.MoveFirst
    i = 0
    Do Until rstDeliveryCVChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstDeliveryCVChild.Fields("ItemName").Value
            .SetText 2, i, rstDeliveryCVChild.Fields("HSNName").Value
            .SetText 3, i, rstDeliveryCVChild.Fields("VchNo").Value
            .SetText 4, i, Val(rstDeliveryCVChild.Fields("Quantity").Value)
            .SetText 5, i, Val(rstDeliveryCVChild.Fields("Rate").Value)
            .SetText 6, i, Val(rstDeliveryCVChild.Fields("Amount").Value)
            .SetText 7, i, rstDeliveryCVChild.Fields("NarrationName").Value
            .SetText 8, i, rstDeliveryCVChild.Fields("NarrationCode").Value
            .SetText 9, i, rstDeliveryCVChild.Fields("VchCode").Value
            .SetText 10, i, rstDeliveryCVChild.Fields("ItemCode").Value
            .SetText 11, i, rstDeliveryCVChild.Fields("HSNCode").Value
            .SetText 12, i, Val(CheckNull(rstDeliveryCVChild.Fields("PendingQty").Value))
        End With
        rstDeliveryCVChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Item List")
End Sub
Private Function UpdateItemList(ByVal ActionType As String, ByVal SrNo As Integer) As Boolean
    Dim CellVal(1 To 7) As Variant, BOM As String
    On Error GoTo ErrorHandler
    UpdateItemList = True
    If ActionType = "D" Then
        If Not blnRecordExist Then Exit Function
        With rstDeliveryCVChild
            If .State = adStateOpen Then
                If .RecordCount > 0 Then .MoveFirst
                Do While Not .EOF
                    If Not CheckEmpty(.Fields("VchCode").Value, False) Then Call UpdateStatus(.Fields("VchCode").Value, .Fields("Quantity").Value, "-")
                    .MoveNext
                Loop
            End If
        End With
        cnDeliveryChallan.Execute "DELETE FROM JobworkBVChild WHERE Code='" & rstDeliveryCVParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I" Then
        With fpSpread1
            .GetText 4, .ActiveRow, CellVal(1) 'Qnty
            .GetText 5, .ActiveRow, CellVal(2)  'Rate
            .GetText 6, .ActiveRow, CellVal(3)  'Amnt
            .GetText 8, .ActiveRow, CellVal(4)  'Narration Code
            .GetText 9, .ActiveRow, CellVal(5)  'VchCode=SOCode+Element+Operation+ItemType for Sales/Purchase Challan & Null for Others
            .GetText 10, .ActiveRow, CellVal(6)  'Item Code
            .GetText 11, .ActiveRow, CellVal(7)  'HSN Code
        End With
        BOM = VchPrefix + IIf(CheckEmpty(CellVal(5), False), "XXXXXXXXXXXXFI", Right(CellVal(5), 14)) 'BOM='0410'+Element+Operation+ItemType
        cnDeliveryChallan.Execute "INSERT INTO JobworkBVChild VALUES ('" & rstDeliveryCVParent.Fields("Code").Value & "','" & Left(CellVal(5), 6) & "','" & BOM & "','" & CellVal(6) & "','" & CellVal(7) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ",'" & CellVal(4) & "'," & SrNo & ",'','','','','',0,'XXXXXX')"
        If Not CheckEmpty(CellVal(5), False) Then Call UpdateStatus(CellVal(5), Val(CellVal(1)), "+")
    End If
    Exit Function
ErrorHandler:
    UpdateItemList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Party" Then
        rstDeliveryCVList.Filter = "[PartyName] Like '%" & SrchText & "%'"
    ElseIf SrchFor = IIf(VchType = "P", "Material Centre", "Consignee") Then
        rstDeliveryCVList.Filter = "[MaterialCentreName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cVal As Variant
    With fpSpread1
        If (Shift = vbCtrlMask And KeyCode = vbKeyD) Or KeyCode = vbKeyF9 Then
            If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then .DeleteRows .ActiveRow, 1: .SetFocus: CalculateTotal
        ElseIf KeyCode = vbKeySpace Then
            If .ActiveCol = 1 Then 'Item
                If cmbChallanType.ListIndex < 2 And fpSpread1.ActiveCol = 1 Then Exit Sub
                .GetText 1, .ActiveRow, cVal 'Item
                Text6.Text = FixQuote(cVal)
                If rstItemList.RecordCount = 0 Then DisplayError ("No record in Item Master"): .SetActiveCell 1, .ActiveRow: .SetFocus: Exit Sub Else rstItemList.MoveFirst
                rstItemList.Find "[Col0] = '" & FixQuote(Trim(cVal)) & "'"
                SelectionType = "S": ItemCode = ""
                Call LoadSelectionList(rstItemList, "List of Items...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text6, ItemCode)
                Call CloseForm(FrmSelectionList)
                If CheckEmpty(ItemCode, False) Then
                    .SetActiveCell 1, .ActiveRow
                Else
                    rstItemList.MoveFirst: rstItemList.Find "[Code] ='" & ItemCode & "'"
                    .SetText 1, .ActiveRow, Text6.Text 'Item Name
                    .SetText 10, .ActiveRow, ItemCode
                    .SetText 5, .ActiveRow, Val(rstItemList.Fields("Price").Value)
                    .GetText 11, .ActiveRow, cVal 'HSN
                    If CheckEmpty(cVal, False) Then .SetText 2, .ActiveRow, rstItemList.Fields("HSNName").Value: .SetText 11, .ActiveRow, rstItemList.Fields("HSNCode").Value
                    .SetFocus
                    Sendkeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 2 Then
                .GetText 11, .ActiveRow, cVal 'HSN Code
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "18"
                FrmGeneralMaster.MasterCode = cVal
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 2, .ActiveRow, slName: .SetText 11, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then LoadMasterList: Sendkeys "{ENTER}"
            ElseIf .ActiveCol = 7 Then
                .GetText 8, .ActiveRow, cVal 'Short Narration
                On Error Resume Next
                FrmGeneralMaster.SL = True
                FrmGeneralMaster.MasterType = "17"
                FrmGeneralMaster.MasterCode = cVal
                Load FrmGeneralMaster
                If Err.Number <> 364 Then FrmGeneralMaster.Show vbModal
                On Error GoTo 0
                .SetText 7, .ActiveRow, slName: .SetText 8, .ActiveRow, slCode
                If Not CheckEmpty(slCode, False) Then LoadMasterList: Sendkeys "{ENTER}"
            End If
        ElseIf KeyCode = vbKeyF11 Then
            If .DataRowCnt = 0 And cmbChallanType.ListIndex < 2 Then LoadOrderList
        End If
        If .DataRowCnt > 0 Then cmbItemType.Enabled = False: cmbChallanType.Enabled = False Else cmbItemType.Enabled = True: cmbChallanType.Enabled = True
    End With
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Item As Variant, Qty As Variant, Rate As Variant
    With fpSpread1
        If Col = 4 Or Col = 5 Then
            .GetText 1, Row, Item
            .GetText 4, Row, Qty
            .GetText 5, Row, Rate
            If Not CheckEmpty(Item, False) Then .SetText 6, Row, Qty * Rate: CalculateTotal Else .SetText 4, Row, "": .SetText 5, Row, "": .SetText 6, Row, ""
        End If
    End With
End Sub
Private Sub fpSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As FPSpreadADO.TextTipFetchMultilineConstants, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim PendingQty As Variant
    fpSpread1.GetText 12, Row, PendingQty
    If Val(PendingQty) = 0 Then Exit Sub
    If Col = 4 Then
        fpSpread1.SetTextTipAppearance "Calibri", 10, False, False, &HC0FFFF, &H80000008
        TipText = "Pending : " & Trim(PendingQty)
        ShowTip = True
    End If
End Sub
Private Sub CalculateTotal()
    Dim i As Integer, Qty As Variant, Amt As Variant
    MhRealInput1.Value = 0: MhRealInput2.Value = 0
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 4, i, Qty
            .GetText 6, i, Amt
            MhRealInput1.Value = MhRealInput1.Value + Val(Qty)
            MhRealInput2.Value = MhRealInput2.Value + Val(Amt)
        Next
        MhRealInput3.Value = MhRealInput2.Value
        MhRealInput5.Value = (MhRealInput3.Value * MhRealInput4.Value) / 100
        MhRealInput8.Value = ((MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput12.Value) * MhRealInput7.Value) / 100
        MhRealInput10.Value = ((MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput12.Value) * MhRealInput9.Value) / 100
        MhRealInput11.Value = Round(MhRealInput3.Value - MhRealInput5.Value + MhRealInput6.Value + MhRealInput8.Value + MhRealInput10.Value + MhRealInput12.Value, 0)
    End With
End Sub
Private Sub DisplayMenu(ByVal OutputTo As String)   'Original/Duplicate/Triplicate
    Dim menusel As String
    If rstDeliveryCVList.RecordCount = 0 Then Exit Sub
    menusel = DisplayPopupMenu(Me.hwnd, 2)
    Call PrintItemIRVch(rstDeliveryCVList.Fields("Code").Value, rstDeliveryCVList.Fields("Type").Value, Choose(menusel, "O", "D", "T"), OutputTo) 'Original/Duplicate/Triplicate
    If Not (rstDeliveryCVList.EOF Or rstDeliveryCVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.SetFocus
End Sub
Private Sub UpdateStatus(ByVal VchCode As String, ByVal Quantity As Long, ByVal Operation As String)
    If InStr(1, "FI", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOParent SET DeliveredQuantityC=DeliveredQuantityC" + Operation + Trim(Quantity) + " WHERE Code+'XXXXXXXXXXXXFI'='" + VchCode + "'"
    End If
    If InStr(1, "FI_MF", Right(VchCode, 2)) > 0 Then
'       cnDeliveryChallan.Execute "UPDATE BookPOChild05 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+Element+'XXXXXXMF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
        cnDeliveryChallan.Execute "UPDATE BookPOChild05 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+'XXXXXXXXXXXXMF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
    If InStr(1, "FI_ME", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOChild06 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+Element+'XXXXXXME'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
    If InStr(1, "FI_CF", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOChild0901 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+Book+'XXXXXXCF'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
    If InStr(1, "FI_MO", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOChild07 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+Element+Operation+'MO'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
    If InStr(1, "FI_BN", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOChild08 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+'XXXXXXXXXXXXBN'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
    If InStr(1, "FI_BM", Right(VchCode, 2)) > 0 Then
        cnDeliveryChallan.Execute "UPDATE BookPOChild0801 SET DeliveredQuantityC=DeliveredQuantityC" & Operation & Trim(Quantity) & " WHERE (Code+Item+'XXXXX'+Category+'BM'='" & VchCode & "' OR Code+'XXXXXXXXXXXXFI'='" & VchCode & "')"
    End If
End Sub
Private Sub LoadOrderList()
    Dim SQL As String
    If rstOrderList.State = adStateOpen Then rstOrderList.Close
    If cmbItemType.ListIndex = 0 Then 'Finished Item
        SQL = "SELECT DISTINCT P.Code+'XXXXXXXXXXXXFI' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FI' As VchNo,P.Date As VchDate,I.Name As Item,P.EstQty01 As OrderedQty,P.DeliveredQuantityC As ChallanQty,P.DeliveredQuantityB As DirectQty FROM (BookPOParent P INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN BookPOChild0801 C ON P.Code=C.Code WHERE (P.BookPrinter='" & PartyCode & "' OR P.TitlePrinter='" & PartyCode & "' OR P.Laminator='" & PartyCode & "' OR P.Binder='" & PartyCode & "' OR C.Vendor='" & PartyCode & "') AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND P.EstQty01-P.DeliveredQuantityB-P.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)<>'FI' AND LEFT(BOM,2)='" & TranType & "') ORDER BY I.Name,P.Date,VchNo"
    Else 'Unfinished Item
'       SQL = "SELECT P.Code+E.Code+'XXXXXXMF' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.BookPrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL "
        SQL = "SELECT P.Code+'XXXXXXXXXXXXMF' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As VchNo,P.Date As VchDate,I.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.BookPrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL " & _
                    "SELECT P.Code+E.Code+'XXXXXXME' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/ME' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL " & _
                    "SELECT P.Code+E.Code+'XXXXXXCF' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CF' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_Printing' As Item,C.ActualQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN BookMaster E ON C.Book=E.Code WHERE P.TitlePrinter='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL " & _
                    "SELECT P.Code+E.Code+O.Code+'MO' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MO' As VchNo,P.Date As VchDate,I.Name+'_'+E.Name+'_'+O.Name As Item,C.Quantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM (((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code WHERE P.Laminator='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.Quantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL " & _
                    "SELECT P.Code+'XXXXXXXXXXXXBN' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/BN' As VchNo,P.Date As VchDate,I.Name+'_Binding' As Item,C.ActualQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code WHERE P.Binder='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "') UNION ALL " & _
                    "SELECT P.Code+C.Item+'XXXXX'+C.Category+'BM' As VchCode,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/BM' As VchNo,P.Date As VchDate,I.Name+'_'+CASE WHEN C.Category='1' THEN O.Name WHEN C.Category='2' THEN R.Name ELSE U.Name END As Item,C.OrderQuantity As OrderedQty,C.DeliveredQuantityC As ChallanQty,C.DeliveredQuantityB As DirectQty FROM ((((BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN OutsourceItemMaster O ON C.Category+C.Item='1'+O.Code) LEFT JOIN PaperMaster R ON C.Category+C.Item='2'+R.Code) LEFT JOIN BookMaster U ON C.Category+C.Item='3'+U.Code WHERE C.Vendor='" & PartyCode & "' AND LEFT(P.Type,1)<>'O' AND RIGHT(P.Type,1)='" & IIf(VchType = "I", "S", "P") & "' AND C.OrderQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC>0 AND P.Code NOT IN (SELECT Ref FROM JobworkBVChild WHERE Ref=P.Code AND RIGHT(BOM,2)='FI' AND LEFT(BOM,2)='" & TranType & "')" & _
                    " ORDER BY Item,VchDate,VchNo"
    End If
    rstOrderList.Open SQL, cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstOrderList.ActiveConnection = Nothing
    If rstOrderList.RecordCount = 0 Then DisplayError ("No Pending Order Exists"): fpSpread1.SetFocus: Exit Sub
    With FrmOrderList.fpSpread1
        .Row = SpreadHeader + 1
        .Col = 5: .Text = "Dlvrd-Bill"
        .Col = 6: .Text = "Dlvrd-Challan"
    End With
    Load FrmOrderList
    FrmOrderList.Text2 = Text3.Text
    Dim i As Integer, Delivered As Long, UnitRate As Double
    With rstOrderList
        For i = 1 To .RecordCount
            With FrmOrderList.fpSpread1
                .MaxRows = .MaxRows + 1
                .InsertRows i, 1
            End With
        Next
        i = 0
        Do While Not .EOF
            i = i + 1
            FrmOrderList.fpSpread1.SetText 1, i, .Fields("Item").Value
            FrmOrderList.fpSpread1.SetText 2, i, .Fields("VchNo").Value: FrmOrderList.fpSpread1.SetText 10, i, .Fields("VchCode").Value
            FrmOrderList.fpSpread1.SetText 3, i, Format(.Fields("VchDate").Value, "dd-MMM-yy")
            FrmOrderList.fpSpread1.SetText 4, i, Val(.Fields("OrderedQty").Value) 'Ordered
            FrmOrderList.fpSpread1.SetText 5, i, Val(.Fields("DirectQty").Value) 'Delivered-Bill
            FrmOrderList.fpSpread1.SetText 6, i, Val(.Fields("ChallanQty").Value) 'Delivered-Challan
            FrmOrderList.fpSpread1.SetText 7, i, Val(.Fields("OrderedQty").Value) - Val(.Fields("DirectQty").Value) - Val(.Fields("ChallanQty").Value) 'Pending
            Delivered = Val(.Fields("ChallanQty").Value) + Val(.Fields("DirectQty").Value)
            FrmOrderList.fpSpread1.SetText 8, i, IIf(Delivered = 0, "Undelivered", IIf(Delivered < Val(.Fields("OrderedQty").Value), "Under Delivery", "Delivered"))
            FrmOrderList.fpSpread1.SetText 9, i, 0
            .MoveNext
        Loop
        FrmOrderList.fpSpread1.SetActiveCell 9, 1
    End With
    FrmOrderList.Check2 = 0
    FrmOrderList.Check1.Visible = False
    CenterForm FrmOrderList
    FrmOrderList.Show vbModal
    If Not CheckEmpty(FrmOrderList.VchCodeList, False) Then
        If rstOrderList.State = adStateOpen Then rstOrderList.Close
        If cmbItemType.ListIndex = 0 Then    'Finished Item
            SQL = "SELECT I.Code As ItemCode,I.Name As ItemName,P.UnitRate,100 As ProfitMargin,H.Code As HSNCode,H.Name As HSNName,P.EstQty01-P.DeliveredQuantityB-P.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/FI' As VchNo,P.Code+'XXXXXXXXXXXXFI' As VchCode FROM (BookPOParent P INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+'XXXXXXXXXXXXFI' IN (" & FrmOrderList.VchCodeList & ") ORDER BY I.Name,VchNo"
        Else 'Unfinished Item
'           SQL = "SELECT I.Code As ItemCode,I.Name+'_'+E.Name+'_Printing' As ItemName,ROUND((C.PrintAmount+C.Adjustment+C.PlateAmount+C.PAdjustment+C.PaperAmount+C.RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As VchNo,P.Code+E.Code+'XXXXXXMF' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.Code+E.Code+'XXXXXXMF' IN (" & FrmOrderList.VchCodeList & ") UNION ALL "
            SQL = "SELECT I.Code As ItemCode,I.Name+'_Printing' As ItemName,ROUND((C.PrintAmount1+C.PrintAmount2+C.PrintAmount4+C.Adjustment+C.PlateAmount1+C.PlateAmount2+C.PlateAmount4+C.PAdjustment+C.PaperAmount1+C.PaperAmount2+C.PaperAmount4+C.RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MF' As VchNo,P.Code+'XXXXXXXXXXXXMF' As VchCode FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+'XXXXXXXXXXXXMF' IN (" & FrmOrderList.VchCodeList & ") UNION ALL " & _
                        "SELECT I.Code As ItemCode,I.Name+'_'+E.Name+'_Printing' As ItemName,ROUND((C.PrintAmount+C.Adjustment+C.PlateAmount+C.PAdjustment+C.PaperAmount+C.RAdjustment)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/ME' As VchNo,P.Code+E.Code+'XXXXXXME' As VchCode FROM (((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code) INNER JOIN ElementMaster E ON C.Element=E.Code WHERE P.Code+E.Code+'XXXXXXME' IN (" & FrmOrderList.VchCodeList & ") UNION ALL " & _
                        "SELECT I.Code As ItemCode,I.Name+'_'+E.Name+'_Printing' As ItemName,ROUND(((C1.PrintAmount+C1.Adjustment+C1.PlateAmount+C1.PAdjustment+C1.PaperAmount+C1.RAdjustment)/(SELECT SUM(ActualQuantity) FROM BookPOChild0901 WHERE Code=P.Code))*C.ActualQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/CF' As VchNo,P.Code+E.Code+'XXXXXXCF' As VchCode FROM ((((BookPOParent P INNER JOIN BookPOChild09 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0901 C ON C.Code=C1.Code) INNER JOIN BookMaster I ON  P.Book=I.Code) INNER JOIN BookMaster E ON C.Book=E.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+'XXXXXXCF' IN (" & FrmOrderList.VchCodeList & ") UNION ALL " & _
                        "SELECT I.Code As ItemCode,I.Name+'_'+E.Name+'_'+O.Name As ItemName,ROUND((C.Amount+C.Adjustment)/C.Quantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.Quantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/MO' As VchNo,P.Code+E.Code+O.Code+'MO' As VchCode FROM ((((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN ElementMaster E ON C.Element=E.Code) INNER JOIN GeneralMaster O ON C.Operation=O.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+E.Code+O.Code+'MO' IN (" & FrmOrderList.VchCodeList & ") UNION ALL " & _
                        "SELECT I.Code As ItemCode,I.Name+'_Binding' As ItemName,ROUND((C.BillAmount-C.VAT)/C.ActualQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.ActualQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/BN' As VchNo,P.Code+'XXXXXXXXXXXXBN' As VchCode FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+'XXXXXXXXXXXXBN' IN (" & FrmOrderList.VchCodeList & ") UNION ALL " & _
                        "SELECT I.Code As ItemCode,I.Name+'_'+CASE WHEN C.Category='1' THEN O.Name WHEN C.Category='2' THEN R.Name ELSE U.Name END As ItemName,ROUND(C.Amount/C.OrderQuantity,3) As UnitRate,P.ProfitMargin,H.Code As HSNCode,H.Name As HSNName,C.OrderQuantity-C.DeliveredQuantityB-C.DeliveredQuantityC As BalQty,(SELECT Code+'-'+Name FROM GeneralMaster WHERE Type='17' AND Value1=1) As Narration,LTRIM(P.Name)+'/'+RIGHT(P.Type,1)+'O/BM' As VchNo,P.Code+C.Item+'XXXXX'+C.Category+'BM' As VchCode FROM (((((BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code) INNER JOIN BookMaster I ON P.Book=I.Code) LEFT JOIN OutsourceItemMaster O ON C.Category+C.Item='1'+O.Code) LEFT JOIN PaperMaster R ON C.Category+C.Item='2'+R.Code) LEFT JOIN BookMaster U ON C.Category+C.Item='3'+U.Code) INNER JOIN GeneralMaster H ON I.HSNCode=H.Code WHERE P.Code+C.Item+'XXXXX'+C.Category+'BM' IN (" & FrmOrderList.VchCodeList & ") " & _
                        "ORDER BY ItemName,VchNo"
        End If
        rstOrderList.Open SQL, cnDeliveryChallan, adOpenKeyset, adLockReadOnly
        If rstOrderList.RecordCount > 0 Then
            i = 0
            With fpSpread1
                Do While Not rstOrderList.EOF
                    i = i + 1
                    .SetText 1, i, rstOrderList.Fields("ItemName").Value
                    .SetText 2, i, rstOrderList.Fields("HSNName").Value: .SetText 11, i, rstOrderList.Fields("HSNCode").Value
                    .SetText 3, i, rstOrderList.Fields("VchNo").Value
                    .SetText 4, i, Val(rstOrderList.Fields("BalQty").Value)
                    UnitRate = Val(rstOrderList.Fields("UnitRate").Value) + (Val(rstOrderList.Fields("UnitRate").Value) * Val(rstOrderList.Fields("ProfitMargin").Value)) / 100
                    .SetText 5, i, Round(UnitRate, 3)
                    .SetText 6, i, Val(rstOrderList.Fields("BalQty").Value) * Round(UnitRate, 3) 'quantity * rate
                    .SetText 7, i, Mid(rstOrderList.Fields("Narration").Value, InStr(1, rstOrderList.Fields("Narration").Value, "-") + 1, 40)
                    .SetText 8, i, Left(rstOrderList.Fields("Narration").Value, InStr(1, rstOrderList.Fields("Narration").Value, "-") - 1)
                    .SetText 9, i, rstOrderList.Fields("VchCode").Value
                    .SetText 10, i, rstOrderList.Fields("ItemCode").Value
                    rstOrderList.MoveNext
                Loop
                Call CalculateTotal
            End With
        End If
    End If
    CloseForm FrmOrderList
End Sub
Private Sub cmbChallanType_Click()
    VchPrefix = IIf(VchType = "I", Choose(cmbChallanType.ListIndex + 1, "08", "06", "12", "14", "16"), Choose(cmbChallanType.ListIndex + 1, "05", "07", "11", "13", "15")) & IIf(cmbChallanType.ListIndex < 2, "10", "00") 'Challan Reversal for To/Not To be billed only
End Sub
Private Sub LoadMasterList()
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "SELECT Name As Col0,Code FROM AccountMaster ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    If rstMaterialCentreList.State = adStateOpen Then rstMaterialCentreList.Close
    rstMaterialCentreList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE [Group]='*99999' ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstMaterialCentreList.ActiveConnection = Nothing
    If rstTaxList.State = adStateOpen Then rstTaxList.Close
    rstTaxList.Open "SELECT Name As Col0,[IGST%],[SGST%],[CGST%],Code FROM TaxMaster ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstTaxList.ActiveConnection = Nothing
    If rstNarrationList.State = adStateOpen Then rstNarrationList.Close
    rstNarrationList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='17' ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstNarrationList.ActiveConnection = Nothing
    If rstHSNCodeList.State = adStateOpen Then rstHSNCodeList.Close
    rstHSNCodeList.Open "SELECT Name As Col0,Code FROM GeneralMaster WHERE Type='18' ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstHSNCodeList.ActiveConnection = Nothing
    If rstItemList.State = adStateOpen Then rstItemList.Close
    rstItemList.Open "SELECT I.Name As Col0,I.Price,I.Code,H.Code As HSNCode,H.Name As HSNName FROM BookMaster I INNER JOIN GeneralMaster H ON I.HSNCode=H.Code ORDER BY I.Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstItemList.ActiveConnection = Nothing
    If rstVchSeriesList.State = adStateOpen Then rstVchSeriesList.Close
    rstVchSeriesList.Open "SELECT Name As Col0,Prefix,Suffix,VchNumbering,Code FROM VchSeriesMaster WHERE VchType='" & IIf("F" & VchType = "FR", "05", IIf("F" & VchType = "FI", "08", IIf("F" & VchType = "FR", "07", "06"))) & "F" & VchType & "' ORDER BY Name", cnDeliveryChallan, adOpenKeyset, adLockReadOnly
    rstVchSeriesList.ActiveConnection = Nothing
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Public Sub PrintItemIRVch(ByVal VchCode As String, ByVal VchType As String, ByVal BillType As String, Optional ByVal OutputType As String)
Dim ChallanType As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstDeliveryCVParent.State = adStateOpen Then rstDeliveryCVParent.Close
    rstDeliveryCVParent.Open "SELECT TYPE FROM JobworkBVParent WHERE Code='" + Left(VchCode, 6) + "' ", cnDeliveryChallan, adOpenKeyset, adLockOptimistic
    If rstDeliveryCVParent.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    ChallanType = (rstDeliveryCVParent.Fields("TYPE").Value)
    rstDeliveryCVParent.ActiveConnection = Nothing
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,EMail,Website,GSTIN,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,Prefix,Suffix FROM CompanyMaster P INNER JOIN CompChild C ON P.Code=C.Code WHERE VchType= " & IIf(ChallanType = "0510FR", 5, IIf(ChallanType = "0710FR", 7, IIf(ChallanType = "0610FI", 6, IIf(ChallanType = "0810FI", 8, 0)))), cnDeliveryChallan, adOpenKeyset, adLockOptimistic
    rstDeliveryCVChild.Open "SELECT LTRIM(P1.Name) As BillNo,P1.Date As BillDate,A.PrintName As Party,A.Address1 As PartyAddress1,A.Address2 As PartyAddress2,A.Address3 As PartyAddress3,A.Address4 As PartyAddress4,A.TIN As PartyGSTIN,IIf(P1.Type= '0810FI',C.PrintName,IIf(P1.Type= '0610FI',C.PrintName,IIf(P1.Type= '0510FR',M.PrintName,IIf(P1.Type= '0710FR',M.PrintName,'')))) As Consignee,IIf(P1.Type= '0810FI',C.Address1,IIf(P1.Type= '0610FI',C.Address1,IIf(P1.Type= '0510FR',M.Address1,IIf(P1.Type= '0710FR',M.Address1,'')))) As ConsigneeAddress1,IIf(P1.Type= '0810FI',C.Address2,IIf(P1.Type= '0610FI',C.Address2,IIf(P1.Type= '0510FR',M.Address2,IIf(P1.Type= '0710FR',M.Address2,'')))) As ConsigneeAddress2, " & _
                                                  "IIf(P1.Type= '0810FI',C.Address3,IIf(P1.Type= '0610FI',C.Address3,IIf(P1.Type= '0510FR',M.Address3,IIf(P1.Type= '0710FR',M.Address3,'')))) As ConsigneeAddress3,IIf(P1.Type= '0810FI',C.Address4,IIf(P1.Type= '0610FI',C.Address4,IIf(P1.Type= '0510FR',M.Address4,IIf(P1.Type= '0710FR',M.Address4,'')))) As ConsigneeAddress4,IIf(P1.Type= '0810FI',C.TIN,IIf(P1.Type= '0610FI',C.TIN,IIf(P1.Type= '0510FR',M.TIN,IIf(P1.Type= '0710FR',M.TIN,'')))) As ConsigneeGSTIN,P1.[Rebate%],P1.Rebate,P1.Freight,P1.Adjustment,P1.TaxableAmount,P1.[IGST%],P1.IGST,P1.[SGST%],P1.SGST,P1.[CGST%],P1.CGST,P1.Amount As TotalAmount,P1.Remarks,N.PrintName As Narration,I.PrintName As Item,H.PrintName As HSNCode," & _
                                                  "'Finish Size: '+LTRIM(S.PrintName) As FinishSize," & _
                                                  "C1.Quantity,C1.Rate,C1.Amount,N.Name As SrNo,'' As cmbTitle,LTRIM(C1.Code)+LTRIM(C1.SrNo) As Ref,M.PrintName As MC FROM (((((((JobworkBVParent P1 INNER JOIN JobworkBVChild C1 ON P1.Code=C1.Code)INNER JOIN AccountMaster A ON P1.Party=A.Code)INNER JOIN AccountMaster C ON P1.Consignee=C.Code)INNER JOIN BookMaster I ON C1.Item=I.Code)LEFT JOIN AccountMaster M ON P1.MaterialCentre=M.Code)LEFT JOIN GeneralMaster N ON C1.Narration=N.Code)LEFT JOIN GeneralMaster H ON C1.HSNCode=H.Code)LEFT JOIN GeneralMaster S ON I.FinishSize=S.Code WHERE P1.Code='" + Left(VchCode, 6) + "' ORDER BY I.PrintName,N.Name", cnDeliveryChallan, adOpenKeyset, adLockOptimistic
    
    '+', '+LTRIM(Case When C2.Pages1 IS NULL Then I.Pages Else (C2.Pages1+C2.Pages2+C2.Pages4)End)+' pages/'+LTRIM( Case When C2.Pages1 IS NULL Then I.Forms Else (C2.Forms1+C2.Forms2+C2.Forms4)End)+'f ('+ Case When C2.Pages1 IS NULL Then LTRIM(IIF(I.OneColorForms<>0,LTRIM(I.OneColorForms)+'f-1Col','')+' '+IIF(I.TwoColorForms<>0,LTRIM(I.TwoColorForms)+'f-2Col','')+' '+IIF(I.FourColorForms<>0,LTRIM(I.FourColorForms)+'f-4Col','')) Else LTRIM(IIF(C2.Forms1<>0,LTRIM(C2.Forms1)+'f-1Col','')+' '+IIF(C2.Forms2<>0,LTRIM(C2.Forms2)+'f-2Col','')+' '+IIF(C2.Forms4<>0,LTRIM(C2.Forms4)+'f-4Col','')) End+')'
    If rstDeliveryCVChild.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rstDeliveryCVChild.ActiveConnection = Nothing
    rptItemIssueReceiptVoucher.Text1.SetText IIf(ChallanType = "0710FR", "Sales Return", IIf(ChallanType = "0510FR", "Purchase", IIf(ChallanType = "0810FI", "Sales ", IIf(ChallanType = "0610FI", "Purchase Return", "Stock Transfer")))) & " Challan"
    rptItemIssueReceiptVoucher.Text13.SetText IIf(ChallanType = "0710FR", "Buyer :", IIf(ChallanType = "0510FR", "Supplier :", IIf(ChallanType = "0810FI", "Buyer :", IIf(ChallanType = "0610FI", "Supplier :", "From: Material Centre"))))
    rptItemIssueReceiptVoucher.Text7.SetText IIf(ChallanType = "0710FR", "Material Centre :", IIf(ChallanType = "0510FR", "Material Centre :", IIf(ChallanType = "0810FI", "Consignee :", IIf(ChallanType = "0610FI", "Consignee :", "TO: Material Centre"))))
    rptItemIssueReceiptVoucher.Text35.SetText "Printed on " & Format(Now, "dd-MMM-yyyy") & " at " & Format(Now, "hh:mm")
    rptItemIssueReceiptVoucher.Text40.SetText IIf(BillType = "O", "(ORIGINAL FOR RECIPIENT)", IIf(BillType = "D", "(DUPLICATE FOR SUPPLIER)", "(TRIPLICATE FOR SUPPLIER)"))
    rptItemIssueReceiptVoucher.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptItemIssueReceiptVoucher.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False)) Then
        rptItemIssueReceiptVoucher.Text4.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value) & Space(1) & "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptItemIssueReceiptVoucher.Text4.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) + ", " + Trim(rstCompanyMaster.Fields("Mobile").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("eMail").Value, False) Then
        rptItemIssueReceiptVoucher.Text4.SetText "E-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
    End If
    rptItemIssueReceiptVoucher.Text8.SetText "GSTIN/UIN : " & Trim(rstCompanyMaster.Fields("GSTIN").Value)
    rptItemIssueReceiptVoucher.Text10.SetText "(" & UCase(Trim(NumberToWords(rstDeliveryCVChild.Fields("TotalAmount").Value, False))) & ")"
    rptItemIssueReceiptVoucher.Text11.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptItemIssueReceiptVoucher.Text26.SetText CheckNull(rstCompanyMaster.Fields("Declaration01").Value)
    rptItemIssueReceiptVoucher.Text25.SetText CheckNull(rstCompanyMaster.Fields("Declaration02").Value)
    rptItemIssueReceiptVoucher.Text22.SetText CheckNull(rstCompanyMaster.Fields("Declaration03").Value)
    rptItemIssueReceiptVoucher.Text12.SetText CheckNull(rstCompanyMaster.Fields("Declaration04").Value)
    rptItemIssueReceiptVoucher.Text9.SetText CheckNull(rstCompanyMaster.Fields("Declaration05").Value)
    rptItemIssueReceiptVoucher.Text30.SetText CheckNull(rstCompanyMaster.Fields("Declaration06").Value)
    rptItemIssueReceiptVoucher.Text31.SetText CheckNull(rstCompanyMaster.Fields("Declaration07").Value)
    rptItemIssueReceiptVoucher.Database.SetDataSource rstDeliveryCVChild, 3, 1
    rptItemIssueReceiptVoucher.DiscardSavedData
    Screen.MousePointer = vbNormal
    If OutputType = "S" Then
        Set FrmReportViewer.Report = rptItemIssueReceiptVoucher
        FrmReportViewer.Show vbModal
    Else
        If rstDeliveryCVList.State = adStateClosed Then  'For Print Utility
            rptItemIssueReceiptVoucher.PaperSource = crPRBinAuto
            rptItemIssueReceiptVoucher.PrintOut False
        Else
            rptItemIssueReceiptVoucher.PaperSource = crPRBinAuto
            rptItemIssueReceiptVoucher.PrintOut
        End If
    End If
    Set rptItemIssueReceiptVoucher = Nothing
    If rstDeliveryCVList.State = adStateClosed Then  'For Print Utility
        Call CloseRecordset(rstCompanyMaster)
    End If
    Call CloseRecordset(rstDeliveryCVChild)
    On Error GoTo 0
End Sub
