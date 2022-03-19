VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmCompanyChild 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CompanyChild"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18315
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
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   18315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Height          =   375
      Left            =   17835
      Picture         =   "CompanyChild.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Refresh"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   17835
      Picture         =   "CompanyChild.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   825
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   17835
      Picture         =   "CompanyChild.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save"
      Top             =   465
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   6270
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   105
      Width           =   17655
      _Version        =   65536
      _ExtentX        =   31141
      _ExtentY        =   11060
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
      Picture         =   "CompanyChild.frx":034E
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   6060
         Left            =   120
         TabIndex        =   0
         Top             =   90
         Width           =   17415
         _Version        =   524288
         _ExtentX        =   30718
         _ExtentY        =   10689
         _StockProps     =   64
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Yu Gothic UI Semilight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   4227327
         MaxCols         =   12
         MaxRows         =   100
         SpreadDesigner  =   "CompanyChild.frx":036A
      End
   End
End
Attribute VB_Name = "FrmCompanyChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnCompanyChild As New ADODB.Connection
Dim rstCompanyChild As New ADODB.Recordset
Dim EditMode As Boolean

Private Sub Command1_Click()

End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    cnCompanyChild.CursorLocation = adUseClient
    cnCompanyChild.Open cnDatabase.ConnectionString
    rstCompanyChild.Open "SELECT Code,Prefix,VchType,Suffix,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,VchName FROM CompChild ORDER BY VchType", cnCompanyChild, adOpenKeyset, adLockReadOnly
     cmdRefresh_Click
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
    Call CloseRecordset(rstCompanyChild)
    Call CloseConnection(cnCompanyChild)
    End Sub
Private Sub cmdProceed_Click()
    SaveFields
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Sub SaveFields()
    Dim i As Integer, Code As Variant, CellVal(1 To 12) As Variant, ActiveCellVal As Variant
     i = 0:
                With rstCompanyChild
                If rstCompanyChild.State = adStateOpen Then rstCompanyChild.Close
                
                rstCompanyChild.Open "SELECT Code,Prefix,VchType,Suffix,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,VchName,Code+VchType As KeyCode FROM CompChild ORDER BY VchType", cnCompanyChild, adOpenKeyset, adLockReadOnly
                
                    cnCompanyChild.Execute "DELETE FROM CompChild" 'WHERE Code + VchType ='" & rstCompanyChild.Fields("KeyCode").Value & "'"
                End With
    With fpSpread1
        For i = 1 To .DataRowCnt
                    .GetText 1, i, CellVal(1)
                    .GetText 2, i, CellVal(2)
                    .GetText 3, i, CellVal(3)
                    .GetText 4, i, CellVal(4)
                    .GetText 5, i, CellVal(5)
                    .GetText 6, i, CellVal(6)
                    .GetText 7, i, CellVal(7)
                    .GetText 8, i, CellVal(8)
                    .GetText 9, i, CellVal(9)
                    .GetText 10, i, CellVal(10)
                    .GetText 11, i, CellVal(11)
                    .GetText 12, i, CellVal(12)
                With rstCompanyChild
                     cnCompanyChild.Execute "INSERT INTO CompChild VALUES ('" & CellVal(1) & "','" & (CellVal(2)) & "','" & CellVal(6) & "','" & CellVal(7) & "','" & CellVal(8) & "','" & CellVal(9) & "','" & CellVal(10) & "','" & CellVal(11) & "','" & CellVal(12) & "','" & CellVal(4) & "','" & CellVal(5) & "','" & CellVal(3) & "')"
                End With
        Next
    End With
End Sub
Private Sub cmdRefresh_Click()
    If EditMode = True Then cmdProceed_Click
    On Error GoTo ErrHandler
    Dim SQL As String, i As Long
    SQL = "SELECT Code,Prefix,VchType,Suffix,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,VchName FROM CompChild"
    Screen.MousePointer = vbNormal
    If rstCompanyChild.State = adStateOpen Then rstCompanyChild.Close
    rstCompanyChild.Open SQL, cnCompanyChild, adOpenKeyset, adLockOptimistic
    If rstCompanyChild.RecordCount = 0 Then Screen.MousePointer = vbNormal: Exit Sub
        With fpSpread1
            .ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
        rstCompanyChild.MoveFirst
            
            Do While Not rstCompanyChild.EOF
            i = i + 1
            .SetText 1, i, rstCompanyChild.Fields("Code").Value 'Code
            .SetText 2, i, rstCompanyChild.Fields("VchType").Value  'VchType
            .SetText 3, i, rstCompanyChild.Fields("VchName").Value  'VchName
            .SetText 4, i, rstCompanyChild.Fields("Prefix").Value  'Prefix
            .SetText 5, i, rstCompanyChild.Fields("Suffix").Value  'Suffix
            .SetText 6, i, rstCompanyChild.Fields("Declaration01").Value  'Declaration01
            .SetText 7, i, rstCompanyChild.Fields("Declaration02").Value  'Declaration02
            .SetText 8, i, rstCompanyChild.Fields("Declaration03").Value  'Declaration03
            .SetText 9, i, rstCompanyChild.Fields("Declaration04").Value  'Declaration04
            .SetText 10, i, rstCompanyChild.Fields("Declaration05").Value  'Declaration05
            .SetText 11, i, rstCompanyChild.Fields("Declaration06").Value  'Declaration06
            .SetText 12, i, rstCompanyChild.Fields("Declaration07").Value  'Declaration07
        
        rstCompanyChild.MoveNext
            Loop
            
            fpSpread1.ColUserSortIndicator(2) = ColUserSortIndicatorAscending
            fpSpread1.ColHeadersUserSortIndex = 2
            fpSpread1.UserColAction = UserColActionSort
        End With
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHandler:
        Screen.MousePointer = vbNormal
        DisplayError (Err.Description)
        EditMode = True
End Sub
Private Sub LoadMasterList()
    If rstCompanyChild.State = adStateOpen Then rstCompanyChild.Close
    rstCompanyChild.Open "SELECT Code,Prefix,VchType,Suffix,Declaration01,Declaration02,Declaration03,Declaration04,Declaration05,Declaration06,Declaration07,VchName FROM CompChild ORDER BY Code", cnCompanyChild, adOpenKeyset, adLockOptimistic
    rstCompanyChild.ActiveConnection = Nothing
End Sub
