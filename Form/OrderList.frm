VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmOrderList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Pending Orders..."
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16110
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
   ScaleHeight     =   6165
   ScaleWidth      =   16110
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
      Left            =   590
      MaxLength       =   40
      TabIndex        =   8
      Top             =   5780
      Width           =   15025
   End
   Begin VB.CommandButton cmdExit 
      Height          =   375
      Left            =   15660
      Picture         =   "OrderList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   120
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   5580
      Left            =   105
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   15495
      _Version        =   65536
      _ExtentX        =   27331
      _ExtentY        =   9842
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
      Picture         =   "OrderList.frx":0102
      Begin VB.CheckBox Check1 
         Caption         =   "Show All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   13200
         TabIndex        =   6
         Top             =   158
         Width           =   1065
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   14310
         TabIndex        =   5
         Top             =   158
         Width           =   1065
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
         Left            =   1205
         MaxLength       =   40
         TabIndex        =   3
         Top             =   105
         Width           =   11805
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
         Caption         =   " Party Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "OrderList.frx":011E
         Picture         =   "OrderList.frx":013A
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   4845
         Left            =   120
         TabIndex        =   0
         Top             =   630
         Width           =   15255
         _Version        =   524288
         _ExtentX        =   26908
         _ExtentY        =   8546
         _StockProps     =   64
         EditEnterAction =   2
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
         MaxCols         =   10
         MaxRows         =   0
         RowHeaderDisplay=   0
         SpreadDesigner  =   "OrderList.frx":0156
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   15480
         Y1              =   525
         Y2              =   525
      End
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Top             =   5780
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
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
      Caption         =   " Find"
      Alignment       =   0
      FillColor       =   9164542
      TextColor       =   0
      Picture         =   "OrderList.frx":0BFD
      Picture         =   "OrderList.frx":0C19
   End
End
Attribute VB_Name = "FrmOrderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VchCodeList As String
Private Sub Form_Load()
    CenterForm Me
End Sub
Private Sub Form_Activate()
    Text1.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}": KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        KeyCode = 0: Me.Hide
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub
Private Sub cmdExit_Click()
    Dim i As Integer, Selection As Variant, VchCode As Variant
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 9, i, Selection
            .GetText 10, i, VchCode
            If Selection = "1" Then VchCodeList = VchCodeList & IIf(VchCodeList = "", "'", ", '") & VchCode & "'"
        Next
    End With
    Me.Hide
End Sub
Private Sub Check1_Click() 'Show All
    Dim i As Integer, CellVal As Variant
    With fpSpread1
        For i = 1 To .DataRowCnt
            If Check1.Value Then
                .Row = i: .RowHidden = False
            Else
                .GetText 7, i, CellVal 'Billable
                If Val(CellVal) = 0 Then .Row = i: .RowHidden = True
            End If
        Next
    End With
End Sub
Private Sub Check2_Click() 'Select All
    Dim i As Integer, CellVal As Variant
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 7, i, CellVal 'Billable
            If Val(CellVal) > 0 Then .SetText 9, i, Check2.Value
        Next
    End With
End Sub

Private Sub Text1_Change()
    Dim i As Integer, cVal As Variant
    With fpSpread1
        For i = 1 To .DataRowCnt 'Unhide All
            .Row = i: .RowHidden = False
        Next
        If CheckEmpty(Text1.Text, False) Then Exit Sub
        .SetActiveCell 1, 1
        For i = 1 To .DataRowCnt
                .GetText 1, i, cVal
                If InStr(StrConv(cVal, vbUpperCase), StrConv(Text1.Text, vbUpperCase)) = 0 Then .Row = i: .RowHidden = True Else .SetActiveCell 1, i
        Next
    End With
End Sub
