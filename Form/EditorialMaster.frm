VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEditorialMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manuscript Status Master"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18630
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   18630
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7665
      Left            =   15
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18600
      _Version        =   65536
      _ExtentX        =   32808
      _ExtentY        =   13520
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Picture         =   "EditorialMaster.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   7455
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   18375
         _ExtentX        =   32411
         _ExtentY        =   13150
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&List"
         TabPicture(0)   =   "EditorialMaster.frx":001C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Text1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "EditorialMaster.frx":0038
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Mh3dFrame5"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.TextBox Text1 
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
            Left            =   -74400
            TabIndex        =   4
            Top             =   6975
            Width           =   17655
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   6435
            Left            =   -74880
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   450
            Width           =   18135
            _ExtentX        =   31988
            _ExtentY        =   11351
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   9164542
            HeadLines       =   1
            RowHeight       =   18
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "Name"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "BusyCode"
               Caption         =   "Alias"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               ScrollBars      =   3
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   13140.28
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   4394.835
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame5 
            Height          =   6795
            Left            =   120
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   480
            Width           =   18135
            _Version        =   65536
            _ExtentX        =   31988
            _ExtentY        =   11986
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
            Picture         =   "EditorialMaster.frx":0054
            Begin FPSpreadADO.fpSpread fpSpread3 
               Height          =   6585
               Left            =   120
               TabIndex        =   7
               Top             =   105
               Width           =   17910
               _Version        =   524288
               _ExtentX        =   31591
               _ExtentY        =   11615
               _StockProps     =   64
               Enabled         =   0   'False
               ButtonDrawMode  =   1
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
               MaxRows         =   1000
               SpreadDesigner  =   "EditorialMaster.frx":0070
            End
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
               Left            =   2640
               TabIndex        =   8
               Top             =   2400
               Width           =   5775
            End
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H008BD6FE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Find"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   330
            Left            =   -74880
            TabIndex        =   5
            Top             =   6975
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   18630
      _ExtentX        =   32861
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filter"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print Preview"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Mail"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "First"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Previous"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Next"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Last"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4
      Left            =   2760
      Top             =   2280
   End
End
Attribute VB_Name = "FrmEditorialMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnBookMaster As New ADODB.Connection
Dim rstBookList As New ADODB.Recordset
Dim rstBookChild As New ADODB.Recordset
Dim rstBookMaster As New ADODB.Recordset
Dim rstMemberList As New ADODB.Recordset
Dim MemberCode As String
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim OutsourceItem As String
Dim EditMode As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    CxnBookMaster.CursorLocation = adUseClient
    CxnBookMaster.Open cnDatabase.ConnectionString
    rstMemberList.Open "SELECT M.Name+' ('+D.Name+')' As Col0,M.Code FROM TeamMemberMaster M INNER JOIN GeneralMaster D ON M.Department=D.Code ORDER BY M.Name", CxnBookMaster, adOpenKeyset, adLockReadOnly
    rstBookList.Open "SELECT Name,BusyCode,Code FROM BookMaster WHERE Type='F' ORDER BY Name", CxnBookMaster, adOpenKeyset, adLockOptimistic
    rstBookMaster.CursorLocation = adUseClient
    rstBookList.Filter = adFilterNone
    Set DataGrid1.DataSource = rstBookList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstBookList.EOF Or rstBookList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstMemberList.ActiveConnection = Nothing
    rstBookList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu
    Text1.SetFocus
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If SSTab1.Tab = 0 Then
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If Not EditMode Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
        End If
        If Not EditMode Then KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD And Toolbar1.Buttons.Item(3).Enabled Then
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
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
            SSTab1.Tab = 1
            SSTab1.SetFocus
        Else
            If Me.ActiveControl.Name <> "fpSpread3" Then SendKeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread3" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    Else
        CloseForm Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstMemberList)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstBookChild)
    Call CloseRecordset(rstBookMaster)
    Call CloseConnection(CxnBookMaster)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
    If rstBookList.RecordCount = 0 Then Exit Sub
    rstBookList.MoveFirst
    If Text1.Text <> "" Then
        rstBookList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        If rstBookList.EOF Then
            rstBookList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstBookList.Bookmark = dblBookMark
                End If
            Else
                PrevStr = ""
            End If
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            SendKeys "{End}"
        Else
            PrevStr = Text1.Text
            dblBookMark = DataGrid1.Bookmark
        End If
    Else
        PrevStr = ""
    End If
    If Not (rstBookList.EOF Or rstBookList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    If rstBookList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstBookList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstBookList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstBookList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstBookList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstBookList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstBookList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstBookList
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
    On Error Resume Next
    If Toolbar1.Buttons.Item(1).Enabled Then
        If SSTab1.Tab = 1 Then
            CenterForm Me
            ViewRecord
        Else
            If Not (rstBookList.EOF Or rstBookList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            CenterForm Me
            Text1.SetFocus
        End If
        SSTab1.TabEnabled(0) = True
    Else
        SSTab1.TabEnabled(0) = False
        Mh3dFrame5.Enabled = True
        fpSpread3.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer, i As Integer
     Dim CellVal As Variant
    If Button.Index = 2 Then
        If rstBookList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 4 Then
        If blnRecordExist And AllowMastersModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Master")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        If CheckMandatoryFields Then Exit Sub
        If UpdateRecord(rstBookMaster) Then
            UpdateFlag = 1
            If UpdateMaterialList("D") Then
                For i = 1 To fpSpread3.DataRowCnt
                    fpSpread3.SetActiveCell 1, i
                    If Not UpdateMaterialList("I") Then
                        UpdateFlag = 0
                        Exit For
                    End If
                Next
            End If
        End If
        If UpdateFlag Then
            CxnBookMaster.CommitTrans
            If rstBookMaster.State = adStateOpen Then rstBookMaster.Close
            rstBookMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstBookMaster) Then
            CxnBookMaster.RollbackTrans
            If rstBookMaster.State = adStateOpen Then rstBookMaster.Close
            rstBookMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstBookList.ActiveConnection = CxnBookMaster
        Do While Not RefreshRecord(rstBookList)
        Loop
        Set DataGrid1.DataSource = rstBookList
        rstBookList.ActiveConnection = Nothing
        rstMemberList.ActiveConnection = CxnBookMaster
        Do While Not RefreshRecord(rstMemberList)
        Loop
        rstMemberList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Name", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstBookList.RecordCount > 0 Then rstBookList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstBookList.RecordCount > 0 Then
           rstBookList.MovePrevious
           If rstBookList.BOF Then
              rstBookList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstBookList.RecordCount > 0 Then
           rstBookList.MoveNext
           If rstBookList.EOF Then
              rstBookList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstBookList.RecordCount > 0 Then rstBookList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Call CloseForm(FrmBookMaster)
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstBookList.EOF Or rstBookList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        Text1.SetFocus
    End If
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
    End If
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
        If SortOrder <> "Name" Then
            SortOrder = "Name"
            rstBookList.Sort = "Name Asc"
        End If
    ElseIf ColIndex = 1 Then
        If SortOrder <> "BusyCode" Then
            SortOrder = "BusyCode"
            rstBookList.Sort = "BusyCode Asc"
        End If
    End If
    DataGrid1.ClearSelCols
    If Not (rstBookList.EOF Or rstBookList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub SetButtons(bVal As Boolean)
    Toolbar1.Buttons.Item(1).Enabled = bVal
    Toolbar1.Buttons.Item(2).Enabled = bVal
    Toolbar1.Buttons.Item(3).Enabled = bVal
    Toolbar1.Buttons.Item(4).Enabled = Not bVal
    Toolbar1.Buttons.Item(5).Enabled = Not bVal
    Toolbar1.Buttons.Item(6).Enabled = bVal
    Toolbar1.Buttons.Item(7).Enabled = bVal
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame5.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstBookList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstBookList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstBookMaster.State = adStateOpen Then rstBookMaster.Close
    rstBookMaster.Open "Select * From BookMaster Where Code = '" & FixQuote(rstBookList.Fields("Code").Value) & "'", CxnBookMaster, adOpenKeyset, adLockOptimistic
    If rstBookMaster.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    fpSpread3.ClearRange 1, 1, fpSpread3.MaxCols, fpSpread3.MaxRows, True
End Sub
Private Sub LoadFields()
    If rstBookMaster.EOF Or rstBookMaster.BOF Then Exit Sub
    Call LoadMaterialList(rstBookMaster.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstBookMaster.RecordCount = 0 Then Exit Sub
    If rstBookMaster.State = adStateOpen Then rstBookMaster.Close
    rstBookMaster.CursorLocation = adUseServer
    rstBookMaster.Open "Select * From BookMaster Where Code = '" & FixQuote(rstBookList.Fields("Code").Value) & "'", CxnBookMaster, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstBookMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    fpSpread3.SetFocus
    blnRecordExist = True
    CxnBookMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then rstBookList.Filter = "[Name] Like '%" & SrchText & "%'"
End Sub
Private Sub fpSpread3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread3.DeleteRows fpSpread3.ActiveRow, 1
            fpSpread3.SetFocus
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If Not EditMode Then fpSpread3.SetText fpSpread3.ActiveCol, fpSpread3.ActiveRow, ""
    ElseIf KeyCode = vbKeySpace Then
        Dim Member As Variant
        With fpSpread3
            If .ActiveCol = 4 Then
                .GetText .ActiveCol, .ActiveRow, Member
                Text2.Text = FixQuote(Member)
                If rstMemberList.RecordCount = 0 Then DisplayError ("No Record in Editorial Team Member Master"): .SetActiveCell 4, .ActiveRow: Exit Sub Else rstMemberList.MoveFirst
                rstMemberList.Find "[Col0] = '" & RTrim(Member) & "'"
                SelectionType = "S"
                MemberCode = ""
                Call LoadSelectionList(rstMemberList, "List of Editorial Team Members...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text2, MemberCode)
                Call CloseForm(FrmSelectionList)
                If MemberCode = "" Then
                    .SetActiveCell 4, .ActiveRow
                Else
                    rstMemberList.MoveFirst: rstMemberList.Find "[Code] ='" & MemberCode & "'"
                    .SetText 4, .ActiveRow, Text2.Text
                    .SetText 8, .ActiveRow, MemberCode
                    .SetFocus
                    SendKeys "{ENTER}"
                End If
            End If
        End With
    End If
End Sub
Private Sub fpSpread3_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant
    fpSpread3.GetText Col, Row, ActiveCellVal
    'Check Date
    If Col = 1 Then
        If Not CheckEmpty(ActiveCellVal, False) Then
            If Len(ActiveCellVal) < 10 Then Cancel = True: Exit Sub
            If Not IsDate(ActiveCellVal) Then Cancel = True: Exit Sub
        End If
    ElseIf Col = 5 Then
        If Not CheckEmpty(ActiveCellVal, False) Then
            If Len(ActiveCellVal) < 10 Then Cancel = True: Exit Sub
            If Not IsDate(ActiveCellVal) Then Cancel = True: Exit Sub
        End If
    ElseIf Col = 6 Then
        If Not CheckEmpty(ActiveCellVal, False) Then
            If Len(ActiveCellVal) < 10 Then Cancel = True: Exit Sub
            If Not IsDate(ActiveCellVal) Then Cancel = True: Exit Sub
        End If
    End If
End Sub
Private Sub LoadMaterialList(ByVal strBookCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstBookChild.State = adStateOpen Then rstBookChild.Close
    rstBookChild.Open "SELECT ArrivedOn,Status,Correction,Name As MemberName,TargetDate,StartDate,RectifiedOn,Member,Remarks FROM BookChild02 T LEFT JOIN TeamMemberMaster M ON T.Member=M.Code WHERE T.Code='" & strBookCode & "' AND " & IIf(UserLevel = 1, "1=1", "RectifiedON=''") & " ORDER BY ArrivedOn,SNo", CxnBookMaster, adOpenKeyset, adLockReadOnly
    rstBookChild.ActiveConnection = Nothing
    If rstBookChild.RecordCount > 0 Then rstBookChild.MoveFirst
    i = 0
    Do While Not rstBookChild.EOF
        i = i + 1
        With fpSpread3
            .SetText 1, i, Format(rstBookChild.Fields("ArrivedON").Value, "dd-mm-yyyy")
            .SetText 2, i, rstBookChild.Fields("Status").Value
            .SetText 3, i, rstBookChild.Fields("Correction").Value
            .SetText 4, i, rstBookChild.Fields("MemberName").Value
            .SetText 5, i, Format(rstBookChild.Fields("TargetDate").Value, "dd-mm-yyyy")
            .SetText 6, i, Format(rstBookChild.Fields("StartDate").Value, "dd-mm-yyyy")
            .SetText 7, i, Format(rstBookChild.Fields("RectifiedON").Value, "dd-mm-yyyy")
            .SetText 8, i, rstBookChild.Fields("Member").Value
            .SetText 9, i, rstBookChild.Fields("Remarks").Value
        End With
        rstBookChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Manuscript Status")
End Sub
Private Function UpdateMaterialList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 9) As Variant
    Dim sDate As String, eDate As String
    On Error GoTo ErrorHandler
    UpdateMaterialList = True
    If Left(ActionType, 1) = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        CxnBookMaster.Execute "DELETE FROM BookChild02 WHERE Code='" & rstBookMaster.Fields("Code").Value & "'"
    Else
        With fpSpread3
            .GetText 1, .ActiveRow, CellVal(1)  'Assigned ON
            .GetText 2, .ActiveRow, CellVal(2)  'Status
            .GetText 3, .ActiveRow, CellVal(3)  'Assignment Remarks
            .GetText 5, .ActiveRow, CellVal(5)  'Target Date
            .GetText 6, .ActiveRow, CellVal(6)  'Start Date
            .GetText 7, .ActiveRow, CellVal(7)  'End Date
            .GetText 8, .ActiveRow, CellVal(8)  'Member Code
            .GetText 9, .ActiveRow, CellVal(9)  'Comments
        End With
        If CellVal(6) = "" Then sDate = "" Else sDate = GetDate(CellVal(6))
        If CellVal(7) = "" Then eDate = "" Else eDate = GetDate(CellVal(7))
        CxnBookMaster.Execute "INSERT INTO BookChild02 VALUES ('" & rstBookMaster.Fields("Code").Value & "'," & fpSpread3.ActiveRow & ",'" & GetDate(CellVal(1)) & "','" & CellVal(2) & "','" & CellVal(3) & "','" & CellVal(8) & "','" & GetDate(CellVal(5)) & "','" & sDate & "','" & eDate & "','E','" & CellVal(9) & "')"
    End If
    Exit Function
ErrorHandler:
    UpdateMaterialList = False
End Function
Private Sub fpSpread3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckItem() Then SSTab1.Tab = 1: fpSpread3.SetFocus: CheckMandatoryFields = True
End Function
Private Function CheckItem() As Boolean
    Dim EndDate As Variant, TargetDate As Variant, AssignedON As Variant, i As Integer, Member As Variant, Status As Variant, Remarks As Variant, StartDate As Variant, Comments As Variant
    CheckItem = False
    For i = 1 To fpSpread3.DataRowCnt
        fpSpread3.SetActiveCell 1, i
        fpSpread3.GetText 1, i, AssignedON
        fpSpread3.GetText 2, i, Status
        fpSpread3.GetText 3, i, Remarks
        fpSpread3.GetText 5, i, TargetDate
        fpSpread3.GetText 6, i, StartDate
        fpSpread3.GetText 7, i, EndDate
        fpSpread3.GetText 8, i, Member
        fpSpread3.GetText 9, i, Comments
        If AssignedON = "" Then CheckItem = True: GoTo Err
        If Len(AssignedON) < 10 Or (Not IsDate(AssignedON)) Then CheckItem = True: GoTo Err
        If Member = "" Then CheckItem = True:   GoTo Err
        If TargetDate = "" Then CheckItem = True: GoTo Err
        If Len(TargetDate) < 10 Or (Not IsDate(TargetDate)) Or Format(TargetDate, "yyyymmdd") < Format(AssignedON, "yyyymmdd") Then CheckItem = True: GoTo Err
        If StartDate <> "" Then If Len(StartDate) < 10 Or (Not IsDate(StartDate)) Or Format(StartDate, "yyyymmdd") < Format(AssignedON, "yyyymmdd") Then CheckItem = True: GoTo Err
        If EndDate <> "" Then If Len(EndDate) < 10 Or (Not IsDate(EndDate)) Or Format(EndDate, "yyyymmdd") < Format(AssignedON, "yyyymmdd") Then CheckItem = True: GoTo Err
        Exit Function
Err:
        If CheckItem Then DisplayError "Data imcomplete in row #" & Trim(Str(i))
    Next
End Function
