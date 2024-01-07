VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmDialog 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Select Printing Format"
      Height          =   2295
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton Command5 
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2400
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sheet"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Format-1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3000
         TabIndex        =   4
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Format-1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   5415
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "9551;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "FrmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Flag As Long, uInput As Variant
Private Sub Command1_Click()
If Flag = 1 Then frmSalesVoucher.PtgType = 1: Me.ActiveControl.SetFocus: Unload Me
If Flag = 6 Then FrmQuery.PtgType = 1: Me.ActiveControl.SetFocus: Unload Me
If Flag = 7 Then FrmQuery.PtgType = 0: Me.ActiveControl.SetFocus: Unload Me
End Sub
Private Sub Command2_Click()
If Flag = 1 Then frmSalesVoucher.PtgType = 2: Me.ActiveControl.SetFocus: Unload Me
If Flag = 6 Then FrmQuery.PtgType = 2: Me.ActiveControl.SetFocus: Unload Me
If Flag = 7 Then FrmQuery.PtgType = 1: Me.ActiveControl.SetFocus: Unload Me
End Sub
Private Sub Command3_Click()
If Flag = 1 Then frmSalesVoucher.PtgType = 3: Me.ActiveControl.SetFocus: Unload Me
If Flag = 6 Then FrmQuery.PtgType = 3: Me.ActiveControl.SetFocus: Unload Me
End Sub
Private Sub Command4_Click()
If Flag = 1 Then frmSalesVoucher.PtgType = 4: Me.ActiveControl.SetFocus: Unload Me
If Flag = 6 Then FrmQuery.PtgType = 4: Me.ActiveControl.SetFocus: Unload Me
End Sub
Private Sub Command5_Click()
Me.ActiveControl.SetFocus: Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo ErrorHandler
If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
CenterForm Me
    BusySystemIndicator True
        If Flag = 1 Then
            ComboBox1.Visible = False
            Command5.Visible = False
        ElseIf Flag = 2 Then 'Supply Type
            Command1.Visible = False: Command2.Visible = False: Command3.Visible = False: Command4.Visible = False
            ComboBox1.Clear
            ComboBox1.Visible = True
            ComboBox1.AddItem "Business to Business", 0  'B2B
            ComboBox1.AddItem "Business to Customer", 1  'B2C
            ComboBox1.AddItem "to SEZ with payment", 2  'SEZWP
            ComboBox1.AddItem "to SEZ without payment", 3  'SEZWOP
            ComboBox1.AddItem "Export with Payment", 4  'EXPWP
            ComboBox1.AddItem "Export without Payment", 5  'EXPWOP
            ComboBox1.AddItem "Deemed Export", 6  'DEXP
            ComboBox1.AddItem "for Invoice", 7   'INV
            ComboBox1.ListIndex = 0
        ElseIf Flag = 3 Then 'IGST on Intra
            Command1.Visible = False: Command2.Visible = False: Command3.Visible = False: Command4.Visible = False
            ComboBox1.Clear
            ComboBox1.Visible = True
            ComboBox1.AddItem "IGST on Intra (Yes)", 0  'Yes
            ComboBox1.AddItem "IGST on Intra (No)", 1  'No
            ComboBox1.ListIndex = 1
        ElseIf Flag = 4 Then 'Reverse Charge Mechanism [RCM]
            Command1.Visible = False: Command2.Visible = False: Command3.Visible = False: Command4.Visible = False
            ComboBox1.Clear
            ComboBox1.Visible = True
            ComboBox1.AddItem "Reverse Charge Mechanism [RCM] (Yes)", 0  'Yes
            ComboBox1.AddItem "Reverse Charge Mechanism [RCM] (No)", 1  'No
            ComboBox1.ListIndex = 1
        ElseIf Flag = 5 Then 'E-Commerce GST
            Command1.Visible = False: Command2.Visible = False: Command3.Visible = False: Command4.Visible = False
            ComboBox1.Clear
            ComboBox1.Visible = True
            ComboBox1.AddItem "E-Commerce GST (Yes)", 0  'Yes
            ComboBox1.AddItem "E-Commerce GST (No)", 1  'No
            ComboBox1.ListIndex = 1
        ElseIf Flag = 6 Or Flag = 7 Then
            ComboBox1.Visible = False
            Command5.Visible = False
        End If
ErrorHandler:
    BusySystemIndicator False
End Sub
Public Function Get_Code()
        uInput = ComboBox1.Text
        'MsgBox uInput, vbInformation, "Message"
If Flag = 2 Then
    If uInput = "Business to Business" Then
        uInput = "B2B"
    ElseIf uInput = "Business to Customer" Then
        uInput = "B2C"
    ElseIf uInput = "to SEZ with payment" Then
        uInput = "SEZWP"
    ElseIf uInput = "to SEZ without payment" Then
        uInput = "SEZWOP"
    ElseIf uInput = "Export with Payment" Then
        uInput = "EXPWP"
    ElseIf uInput = "Export without Payment" Then
        uInput = "EXPWOP"
    ElseIf uInput = "Deemed Export" Then
        uInput = "DEXP"
    ElseIf uInput = "for Invoice" Then
        uInput = "INV"
    End If
ElseIf Flag = 3 Then
    If uInput = "IGST on Intra (Yes)" Then
        uInput = "Y"
    ElseIf uInput = "IGST on Intra (No)" Then
        uInput = "N"
    End If
ElseIf Flag = 4 Then
    If uInput = "Reverse Charge Mechanism [RCM] (Yes)" Then
        uInput = "Y"
    ElseIf uInput = "Reverse Charge Mechanism [RCM] (No)" Then
        uInput = "N"
    End If
ElseIf Flag = 5 Then
    If uInput = "E-Commerce GST (Yes)" Then
        uInput = "Yes"
    ElseIf uInput = "E-Commerce GST (No)" Then
        uInput = "Null"
    End If
End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Get_Code
Me.ActiveControl.SetFocus: Unload Me
End Sub
