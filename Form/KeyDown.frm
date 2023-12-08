VERSION 5.00
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form KeyDown 
   Caption         =   " List of  Master"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10431.9
   ScaleMode       =   0  'User
   ScaleWidth      =   6783.66
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   9714
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CheckBox Check1 
         Caption         =   "Unhide All"
         Height          =   195
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   9195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3135
         _Version        =   524288
         _ExtentX        =   5530
         _ExtentY        =   16219
         _StockProps     =   64
         ColHeaderDisplay=   1
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   42
         MaxRows         =   100
         ScrollBars      =   2
         SpreadDesigner  =   "KeyDown.frx":0000
         UserResize      =   0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   2
         Left            =   45
         TabIndex        =   3
         Top             =   120
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   "   F8->Un-Hide   F9->Hide  "
         Alignment       =   0
         FillColor       =   8421504
         TextColor       =   16777215
         Picture         =   "KeyDown.frx":48C6
         Picture         =   "KeyDown.frx":48E2
      End
   End
End
Attribute VB_Name = "KeyDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HideFlag As Boolean
Private Sub Form_Load()
        fpSpread1.Height = 9150
'    Me.Top = 1200 '(MdiMainMenu.ScaleHeight - Me.Height) \ 2 + 1000
 '   Me.Left = 0 '(MdiMainMenu.ScaleWidth - Me.Width) - 760
End Sub
Private Sub Form_Activate()
    Format_Grid
    MdiMainMenu.oExitFlage = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Me.Hide: MdiMainMenu.oExitFlage = True
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, StringNo, mString As String, cVal As Variant
'    Me.Left = (MdiMainMenu.ScaleWidth - Me.Width)
    If Shift = 0 And KeyCode = vbKeyEscape Then 'Exit
        Me.Hide: fpSpread1.SetActiveCell fpSpread1.Col, 0
        If Shift = 0 And KeyCode = vbKeyEscape Then MdiMainMenu.oExitFlage = False
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then 'Select
        Me.Hide
    ElseIf Shift = 0 And KeyCode = vbKeyF9 Then 'Hide Row
                StringNo = "Custom String_" & MdiMainMenu.oButtonIndex
                mString = Trim(ReadFromFile(StringNo))
                If mString = "" Then WriteToFile StringNo, ""
                 cVal = fpSpread1.ActiveRow
                        If InStr(1, mString, Format(Trim(cVal), "00")) = 0 Then
                            If mString <> "" Then WriteToFile StringNo, mString & "_" & Format(Trim(cVal), "00") Else WriteToFile StringNo, Format(Trim(cVal), "00")
                            fpSpread1.SetText 9, fpSpread1.ActiveRow, 1
                            fpSpread1.Row = fpSpread1.ActiveRow: fpSpread1.RowHidden = True
                        End If
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then 'Refresh Data
        Format_Grid
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF8 Then  'Refresh Data
         Check1.Value = 1
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF9 Then  'Refresh Data
        Check1.Value = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF8 Then 'UnHide Row
                StringNo = "Custom String_" & MdiMainMenu.oButtonIndex
                mString = Trim(ReadFromFile(StringNo))
                If mString = "" Then WriteToFile StringNo, ""
                cVal = fpSpread1.ActiveRow
            If InStr(1, mString, Format(Trim(cVal), "00")) > 0 Then
                            fpSpread1.Col = fpSpread1.ActiveCol: fpSpread1.Row = fpSpread1.ActiveRow: fpSpread1.FontBold = True: fpSpread1.FontSize = 10:  fpSpread1.ForeColor = vbBlue:
                            fpSpread1.SetText 9, fpSpread1.ActiveRow, 0
                            mString = ""
                For i = 1 To fpSpread1.DataRowCnt
                fpSpread1.GetText 9, i, cVal
                If cVal = 1 Then cVal = i: mString = mString & Format(Trim(cVal), "00") & "_"
                Next i
                            WriteToFile StringNo, mString
            End If
    End If
End Sub
Private Sub Format_Grid()
Dim StringNo, mString As String, cVal As Variant, i As Integer
        StringNo = "Custom String_" & MdiMainMenu.oButtonIndex
        mString = Trim(ReadFromFile(StringNo))
        If mString = "" Then WriteToFile StringNo, ""
        Me.Height = 10155
        Me.Width = 4360
'        Me.Left = 13200
        Me.Top = 990
        Me.Frame1.Width = 4300
        Me.Check1.Left = 3010
        fpSpread1.Width = 3900
        fpSpread1.ColWidth(1) = 30
        fpSpread1.ColWidth(5) = 30
        fpSpread1.RowHeadersShow = True
    For i = 1 To fpSpread1.DataRowCnt
        cVal = i
        If InStr(1, mString, Format(Trim(cVal), "00")) > 0 Then
            fpSpread1.SetText 9, i, 1
            fpSpread1.Row = i: fpSpread1.RowHidden = True
        Else
            fpSpread1.GetText 9, i, cVal
            If cVal <> 1 And Check1.Value Then fpSpread1.Row = i: fpSpread1.RowHidden = False
        End If
    Next i
fpSpread1.SetActiveCell fpSpread1.ActiveCol, fpSpread1.ActiveRow: fpSpread1.SetFocus
End Sub
Private Sub Check1_Click()
Dim StringNo, mString As String, cVal As Variant, i As Integer
If Check1.Value Then If MsgBox("Do You wants to Reset Default value Setting", vbYesNo) = vbYes Then HideFlag = True

    If HideFlag = False Then
            StringNo = "Custom String_" & MdiMainMenu.oButtonIndex
            mString = Trim(ReadFromFile(StringNo))
            If mString = "" Then WriteToFile StringNo, ""
    ElseIf HideFlag = True Then
            StringNo = "Default String_" & MdiMainMenu.oButtonIndex
            mString = Trim(ReadFromFile(StringNo))
            If mString = "" Then WriteToFile StringNo, ""
    End If
    
    If Check1.Value Then
        For i = 1 To fpSpread1.DataRowCnt
                If HideFlag = True Then fpSpread1.GetText fpSpread1.ActiveCol, i, cVal Else cVal = i
                
                If HideFlag = True And InStr(1, mString, Format(Trim(cVal), "00")) > 0 Then
                    fpSpread1.SetText 9, i, 0
                    fpSpread1.Row = i: fpSpread1.RowHidden = False
                ElseIf InStr(1, mString, Format(Trim(cVal), "00")) > 0 Then
                    fpSpread1.SetText 9, i, 1
                    fpSpread1.Row = i: fpSpread1.RowHidden = False
                 End If
                 
                 fpSpread1.GetText 9, i, cVal
                 fpSpread1.Col = fpSpread1.ActiveCol: fpSpread1.Row = i:
                    If cVal = 1 Then
                       fpSpread1.ForeColor = vbRed
                    ElseIf cVal = 0 Then
                        fpSpread1.ForeColor = vbBlack
                    Else
                        fpSpread1.ForeColor = vbBlue
                    End If

        Next i
        HideFlag = False
    Else
        Format_Grid
    End If
End Sub
