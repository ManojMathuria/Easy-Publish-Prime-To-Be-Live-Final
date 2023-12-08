VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmFY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Financial Year ..."
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4440
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
   ScaleHeight     =   1110
   ScaleWidth      =   4440
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3960
      Picture         =   "FY.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdLogin 
      Height          =   375
      Left            =   3960
      Picture         =   "FY.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Login"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtUserName 
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
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   14
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   560
      Width           =   2295
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   900
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   1587
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
      Picture         =   "FY.frx":057E
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
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
         Caption         =   " &User Name"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "FY.frx":059A
         Picture         =   "FY.frx":05B6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   435
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
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
         Caption         =   " &Password"
         Alignment       =   0
         FillColor       =   9164542
         TextColor       =   0
         Picture         =   "FY.frx":05D2
         Picture         =   "FY.frx":05EE
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   1440
      Top             =   120
   End
End
Attribute VB_Name = "FrmFY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oEncrypt As New clsBlowFish, oRegistry As New clsRegistry
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    txtUserName.Text = oRegistry.GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Easy Publish", "Last User", "")
    LoginSuccess = False
    Exit Sub
ErrorHandler:
    CloseForm Me
End Sub
Private Sub Form_Activate()
    If txtUserName.Text <> "" Then txtPassword.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Sendkeys "{TAB}", True
       KeyCode = 0
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    BusySystemIndicator False
    Set oRegistry = Nothing
    Set oEncrypt = Nothing
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(FrmFY)
End Sub
Private Sub cmdLogin_Click()
    If LoginSuccess Then Exit Sub
    Dim IsUserNameOk As Boolean, IsPasswordOk As Boolean
    On Error GoTo ErrorHandler
            IsUserNameOk = True
            If UCase("pubprint123!@#") = UCase(Trim(txtPassword.Text)) Then IsPasswordOk = True: Call CreateFY
        If Not IsUserNameOk Then
           txtUserName.SetFocus
        ElseIf Not IsPasswordOk Then
           txtPassword.SetFocus
        Else
           LoginSuccess = True
           Me.Caption = "Creating New FY Successful !!!"
           cmdCancel.ToolTipText = "Proceed"
           Call oRegistry.SetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Easy Publish", "Last User", Trim(txtUserName.Text))
           cmdLogin.Picture = LoadPicture(App.Path & "\Icon\Green.Bmp")
        End If
    Exit Sub
ErrorHandler:
End Sub
Private Sub cmdCancel_Click()
     Call CloseForm(FrmFY)
End Sub
Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub
Private Sub txtUserName_Validate(Cancel As Boolean)
    If CheckEmpty(txtUserName, False) Then Cancel = True
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub
Private Sub Timer1_Timer()
    If FrmFY.Caption = " " Then
       If LoginSuccess Then Me.Caption = "Financial Year Created Successful !" Else Me.Caption = "Create Financial Year...!!!"
    Else
        FrmFY.Caption = " "
    End If
    If LoginSuccess Then CloseForm Me
End Sub
