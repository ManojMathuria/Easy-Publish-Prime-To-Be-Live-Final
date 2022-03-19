VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmUnpackedBillsList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Bills..."
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
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
   ScaleHeight     =   5790
   ScaleWidth      =   4845
   Begin VB.CommandButton cmdExit 
      Height          =   375
      Left            =   4360
      Picture         =   "UnpackedBillsList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   5580
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   105
      Width           =   4135
      _Version        =   65536
      _ExtentX        =   7294
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
      Picture         =   "UnpackedBillsList.frx":0102
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
         Left            =   600
         MaxLength       =   40
         TabIndex        =   6
         Top             =   5130
         Width           =   3390
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
         Width           =   2795
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
         Picture         =   "UnpackedBillsList.frx":011E
         Picture         =   "UnpackedBillsList.frx":013A
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   4305
         Left            =   120
         TabIndex        =   0
         Top             =   630
         Width           =   3900
         _Version        =   524288
         _ExtentX        =   6879
         _ExtentY        =   7594
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
         MaxCols         =   4
         MaxRows         =   0
         RowHeaderDisplay=   0
         ScrollBars      =   2
         SpreadDesigner  =   "UnpackedBillsList.frx":0156
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   5130
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
         Caption         =   " &Find"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "UnpackedBillsList.frx":0784
         Picture         =   "UnpackedBillsList.frx":07A0
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   9000
         Y1              =   5030
         Y2              =   5030
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9000
         Y1              =   530
         Y2              =   530
      End
   End
End
Attribute VB_Name = "FrmUnpackedBillsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VchCodeList As String
Public VchNoList As String
Private Sub Form_Load()
    CenterForm Me
    With fpSpread1
        .Col = 4
        .ColHidden = True
    End With
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name <> "Text1" Then SendKeys "{TAB}": KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        cmdExit_Click
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub
Private Sub cmdExit_Click()
    Dim i As Integer, Status As Variant, VchCode As Variant, VchNo As Variant, K As Integer
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 3, i, Status
            .GetText 4, i, VchCode
            .GetText 1, i, VchNo
            If Status = "1" Then
                VchCodeList = VchCodeList & IIf(VchCodeList = "", "'", ", '") & VchCode & "'"
                VchNoList = VchNoList & IIf(VchNoList = "", "", "+") & VchNo
                K = K + 1
            End If
        Next
        Call MsgBox(Trim(Str(K)) + " bill(s) selected !!!", vbInformation, App.Title)
    End With
    Me.Hide
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim i As Integer, VchNo As Variant
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.GetText 1, i, VchNo
            If StrConv(VchNo, vbUpperCase) = StrConv(Trim(Text1.Text), vbUpperCase) Then fpSpread1.SetActiveCell 3, i: fpSpread1.SetFocus: Exit For
        Next
    End If
End Sub
