VERSION 5.00
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmItemSearchList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Items..."
   ClientHeight    =   7830
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10545
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
   ScaleHeight     =   7830
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   7035
      Left            =   45
      TabIndex        =   1
      Top             =   345
      Width           =   10455
      _Version        =   524288
      _ExtentX        =   18441
      _ExtentY        =   12409
      _StockProps     =   64
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
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
      GridColor       =   4227327
      MaxCols         =   11
      MaxRows         =   1000
      ScrollBars      =   2
      SpreadDesigner  =   "ItemSearchList.frx":0000
   End
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
      Left            =   530
      TabIndex        =   0
      ToolTipText     =   "Find"
      Top             =   7445
      Width           =   7815
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Proceed"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   2
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemSearchList.frx":0B0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemSearchList.frx":0C1F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   7445
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   9164542
      BackColor       =   0
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
      Picture         =   "ItemSearchList.frx":0D31
      Picture         =   "ItemSearchList.frx":0D4D
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Index           =   2
      Left            =   8325
      TabIndex        =   4
      Top             =   7440
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
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
      Caption         =   " F2->Proceed  Esc->Cancel"
      Alignment       =   0
      FillColor       =   8421504
      TextColor       =   16777215
      Picture         =   "ItemSearchList.frx":0D69
      Picture         =   "ItemSearchList.frx":0D85
   End
End
Attribute VB_Name = "FrmItemSearchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstItemSearchList As New ADODB.Recordset, LoadItems As Boolean
Dim PrevStr As String
Private Sub Form_Load()
    rstItemSearchList.Filter = adFilterNone
    Set fpSpread1.DataSource = rstItemSearchList
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        KeyCode = 0: Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        KeyCode = 0: Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub
Private Sub Text1_Change()
    With rstItemSearchList
        If .RecordCount = 0 Then Exit Sub
        If Len(Text1.Text) > 0 Then
            .MoveFirst
            .Filter = "[Col0] Like '%" & Text1.Text & "%'"
            If Not .EOF Then
                PrevStr = Text1.Text
            Else
                .Filter = adFilterNone: .MoveFirst
                Text1.Text = PrevStr
                Sendkeys "{End}"
            End If
        Else
            .Filter = adFilterNone: .MoveFirst
            PrevStr = ""
        End If
'        fpSpread1.SetActiveCell 3, 1
    End With
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    With rstItemSearchList
        If .RecordCount = 0 Then Exit Sub
        If Shift = 0 And KeyCode = vbKeyBack Then
            .MoveFirst
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyUp Then
            .MovePrevious
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyDown Then
            .MoveNext
            If .EOF Then .MoveLast
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
            .Move -20
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
            .Move 20
            If .EOF Then .MoveLast
            KeyProcessed = True
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyHome Then
            .MoveFirst
            If .BOF Then .MoveFirst
            KeyProcessed = True
        ElseIf Shift = vbCtrlMask And KeyCode = vbKeyEnd Then
            .MoveLast
            If .EOF Then .MoveLast
            KeyProcessed = True
        End If
    End With
    If KeyProcessed Then KeyCode = 0
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    LoadItems = IIf(Button.Index = 1, True, False)
    Set rstItemSearchList = Nothing
    Me.Hide
End Sub
