VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmLicenceAgreement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Licence Agreement"
   ClientHeight    =   9075
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   13440
   ClipControls    =   0   'False
   Icon            =   "LicenceAgreement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   13440
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2880
      TabIndex        =   16
      Top             =   8040
      Width           =   4695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Update Major"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   13
      Top             =   8520
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Update Version"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   12
      Top             =   8520
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Agree to Terms && Conditions"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   11
      Top             =   7680
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Activate &Later"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   10
      Top             =   8040
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Activate Key"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7680
      TabIndex        =   9
      Top             =   8040
      Visible         =   0   'False
      Width           =   3645
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   5415
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   13215
      _Version        =   524288
      _ExtentX        =   23310
      _ExtentY        =   9551
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxRows         =   506
      OperationMode   =   1
      SelectBlockOptions=   1
      SpreadDesigner  =   "LicenceAgreement.frx":000C
      TabEnhancedShape=   1
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   1020
      Left            =   106
      Picture         =   "LicenceAgreement.frx":2093
      ScaleHeight     =   674.24
      ScaleMode       =   0  'User
      ScaleWidth      =   674.24
      TabIndex        =   1
      Top             =   135
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   0
      Top             =   7065
      Width           =   1740
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   2
      Top             =   7515
      Width           =   1725
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   405
      Left            =   2880
      TabIndex        =   17
      Top             =   8040
      Visible         =   0   'False
      Width           =   4695
      VariousPropertyBits=   545282075
      BackColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "8281;714"
      MatchEntry      =   0
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   405
      Left            =   120
      TabIndex        =   15
      Top             =   8040
      Width           =   2685
      VariousPropertyBits=   545282075
      BackColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4736;714"
      MatchEntry      =   0
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3000
      TabIndex        =   14
      Top             =   9000
      Width           =   8205
   End
   Begin VB.Label Label1 
      Caption         =   " Renewal Key :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   8
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -30
      X2              =   19200
      Y1              =   6885
      Y2              =   6885
   End
   Begin VB.Label lblDescription 
      Caption         =   "Website: http://www.easyinfosolution.com/   email: sales@easyinfosolution.com"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1170
      TabIndex        =   3
      Top             =   960
      Width           =   12045
   End
   Begin VB.Label lblTitle 
      Caption         =   "Easy Info Solutions International"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1170
      TabIndex        =   5
      Top             =   240
      Width           =   12045
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   -15
      X2              =   19200
      Y1              =   6900
      Y2              =   6900
   End
   Begin VB.Label lblVersion 
      Caption         =   "Easy Publish  21|Rel 05 | 06.29 Version |Production & Inventory Management System"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1170
      TabIndex        =   6
      Top             =   660
      Width           =   12045
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"LicenceAgreement.frx":2BF0
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   135
      TabIndex        =   4
      Top             =   6945
      Width           =   10350
   End
End
Attribute VB_Name = "frmLicenceAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Decrypt As Variant
Dim EncryptFlag As Boolean
' Website hyperlink...
Private WithEvents oHuffman As clsHuffman
Attribute oHuffman.VB_VarHelpID = -1
Private oRegistry As New clsRegistry
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long
' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Dim VchType As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Dim rstCompanyMaster As New ADODB.Recordset
Private Sub Check1_Click()
    If Check1.Value Then Command1.Visible = True
    If Check1.Value = 0 Then Command1.Visible = False
End Sub
Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub Combo1_Change()
If Combo1.ListIndex = 0 Then 'Super User
    Text1.Visible = False: Combo2.Visible = True
            Combo2.Clear
            Combo2.AddItem "EasyPublish", 0
            Combo2.AddItem "Admin", 1
        If Trim(ReadFromFile(Combo1.Text)) = "EasyPublish" Then
            Combo2.ListIndex = 0
        ElseIf Trim(ReadFromFile(Combo1.Text)) = "Admin" Then
            Combo2.ListIndex = 1
        End If
ElseIf Combo1.ListIndex = 2 Then 'Data Base Type
    Text1.Visible = False: Combo2.Visible = True
            Combo2.Clear
            Combo2.AddItem "MS Access", 0
            Combo2.AddItem "MS SQL", 1
        If Trim(ReadFromFile(Combo1.Text)) = "MS Access" Then
            Combo2.ListIndex = 0
        ElseIf Trim(ReadFromFile(Combo1.Text)) = "MS SQL" Then
            Combo2.ListIndex = 1
        End If
ElseIf Combo1.ListIndex = 8 Then 'Client ID
    Text1.Visible = False: Combo2.Visible = True
            Combo2.Clear
            Combo2.AddItem "Printer", 0
            Combo2.AddItem "Publisher", 1
        If Trim(ReadFromFile(Combo1.Text)) = "Printer" Then
            Combo2.ListIndex = 0
        ElseIf Trim(ReadFromFile(Combo1.Text)) = "Publisher" Then
            Combo2.ListIndex = 1
        End If
End If
If Combo1.ListIndex = 1 Or Combo1.ListIndex = 3 Or Combo1.ListIndex = 4 Or Combo1.ListIndex = 5 Or Combo1.ListIndex = 6 Or Combo1.ListIndex = 7 Then
    Text1.Visible = True: Combo2.Visible = False
End If
Text1.Text = Trim(ReadFromFile(Combo1.Text))
EncryptFlag = False
Command1.Caption = IIf(Combo1.ListIndex = 0, "Update Super User", IIf(Combo1.ListIndex = 1, "Activate Key", IIf(Combo1.ListIndex = 2, "Update Database Type", IIf(Combo1.ListIndex = 3, "Update Server Name", IIf(Combo1.ListIndex = 4, "Update Server User", IIf(Combo1.ListIndex = 5, "Update Server Passward", IIf(Combo1.ListIndex = 6, "Update Tally Port", IIf(Combo1.ListIndex = 7, "Update Server ID", IIf(Combo1.ListIndex = 8, "Update Client ID", " ")))))))))
End Sub
Private Sub Command2_Click()
    Unload Me
    LaterFlag = True
End Sub
Public Function Encrypted(Decrypted As Variant, Encrypt As Variant)
Dim K As Long, n As Long
Dim Flag As Boolean
Dim Key As String
Dim Key1 As String
Dim Key2 As String
Dim Key3 As String
Dim Key4 As String
Dim dueDate As Date, sNow As Date
Key1 = "": Key2 = "": Key3 = "": Key4 = "":
K = 0:
'Company Fundation Date 28-SEP-2-16
n = Len(Text1.Text)

'Check For Stoping Existing Encryption
Do While n <> 0 And EncryptFlag = False And Key <> "§"
If Flag = False Then Key1 = "E"
If n <> 0 Then K = K + 1: Key = (Mid(Text1.Text, K, 1)): n = n - 1:
Loop
If Key = "§" Then EncryptFlag = True

n = Len(Text1.Text): K = 0

Do While n <> 0 And EncryptFlag = False
If Flag = False Then Key1 = "E"
If n <> 0 Then K = K + 1: Key1 = Key1 + (Mid(Text1.Text, K, 1)): n = n - 1:

If Flag = False Then Key2 = Key2 + "§I"
If n <> 0 Then K = K + 1: Key2 = Key2 + Mid(Text1.Text, K, 1): n = n - 1:

If Flag = False Then Key3 = Key3 + "§S"
If n <> 0 Then K = K + 1: Key3 = Key3 + Mid(Text1.Text, K, 1): n = n - 1:

If Flag = False Then Key4 = Key4 + "§I"
If n <> 0 Then K = K + 1: Key4 = Key4 + Mid(Text1.Text, K, 1): n = n - 1:
Flag = True
Loop
If EncryptFlag = False Then Encrypt = Key1 + Key2 + Key3 + Key4 + "§" + " " + "§"
If EncryptFlag = False Then sNow = Format(Now(), "DD-MMM-YYYY")
If EncryptFlag = False Then Encrypted = True
End Function
Private Sub Command1_Click()
Decrypt = "":
If Combo1.ListIndex <> 0 And Combo1.ListIndex <> 1 And Combo1.ListIndex <> 2 And Combo1.ListIndex <> 6 And Combo1.ListIndex <> 7 And Combo1.ListIndex <> 8 Then
    If Encrypted(Trim(Text1.Text), Decrypt) Then
        If EncryptFlag = False And Text1.Text <> "" Then Text1.Text = Decrypt: EncryptFlag = True
    End If
End If
If Combo1.ListIndex = 0 Then 'Supper User
    WriteToFile Combo1.Text, Combo2.Text
ElseIf Combo1.ListIndex = 1 Then 'Renewal Key
    If Text1.Text <> "" Then
        WriteToFile "Server ID", Text1.Text + "@" + ServerID
        RenewFlag = True
        cmdOK_Click
    Else
        Text1.SetFocus
    End If
ElseIf Combo1.ListIndex = 2 Then 'Database Type
    WriteToFile Combo1.Text, Combo2.Text
ElseIf EncryptFlag = True And Combo1.ListIndex = 3 Then 'Server Name
    WriteToFile Combo1.Text, Text1.Text
ElseIf EncryptFlag = True And Combo1.ListIndex = 4 Then 'Server User
    WriteToFile Combo1.Text, Text1.Text
ElseIf EncryptFlag = True And Combo1.ListIndex = 5 Then 'Server Password
    WriteToFile Combo1.Text, Text1.Text
ElseIf EncryptFlag = True And Combo1.ListIndex = 6 Then 'Tally Port
    WriteToFile Combo1.Text, Text1.Text
 ElseIf Combo1.ListIndex = 7 Then 'Server ID
    'WriteToFile Combo1.Text, Text1.Text
 ElseIf Combo1.ListIndex = 8 Then 'Cleint ID
    WriteToFile Combo1.Text, Combo2.Text
 End If
    Command1.Visible = False
    Check1.Value = 0
End Sub
Private Sub Form_Load()
    Dim R As Long, C As Long
    CenterForm Me
    Me.Caption = "License Agreement"   '"About " & App.Title
    If Trim(ReadFromFile("Super User")) = "EasyPublish" Then Command3.Visible = True Else Me.Height = 9540: Command3.Visible = False
    If Trim(ReadFromFile("Super User")) = "EasyPublish" Then Command4.Visible = True Else Me.Height = 9540: Command4.Visible = False
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    lblVersion.Caption = "Easy Publish |Rel  21.05 |Version " & App.Major & "." & App.Minor & "." & App.Revision & " |Production && Inventory Management System"
    lblTitle.Caption = "Easy Info Solutions International" 'App.Title
    
    With fpSpread1
    .MaxRows = 57: .MaxCols = 2
    fpSpread1.RowHeadersShow = False
    fpSpread1.ColHeadersShow = False
    
    For R = 1 To 22
    C = 1
        fpSpread1.Col = C: fpSpread1.Row = R: fpSpread1.CellType = CellTypeEdit: fpSpread1.TypeHAlign = TypeHAlignCenter
    Next
        fpSpread1.Col = 2: fpSpread1.Row = 1: fpSpread1.CellType = CellTypeEdit: fpSpread1.TypeHAlign = TypeHAlignCenter: fpSpread1.RowsFrozen = 7
    For R = 3 To 22
    C = 2
        fpSpread1.Col = C: fpSpread1.Row = R: fpSpread1.CellType = CellTypeEdit: fpSpread1.TypeHAlign = TypeHAlignLeft: fpSpread1.TypeTextWordWrap = True: fpSpread1.RowMerge = MergeAlways
    Next
    End With
    Combo1.AddItem "Super User", 0
    Combo1.AddItem "Renewal Key", 1
    Combo1.AddItem "Database Type", 2
    Combo1.AddItem "Server Name", 3
    Combo1.AddItem "Server User", 4
    Combo1.AddItem "Server Password", 5
    Combo1.AddItem "Tally Port", 6
    Combo1.AddItem "Server ID", 7
    Combo1.AddItem "Client ID", 8
    Combo1.ListIndex = 1
If Combo1.ListIndex = 0 Then
    Combo2.AddItem "EasyPublish", 0
    Combo2.AddItem "Admin", 1
    Combo2.ListIndex = 1
ElseIf Combo1.ListIndex = 2 Then
    Combo2.AddItem "MS Access", 0
    Combo2.AddItem "MS SQL", 1
    Combo2.ListIndex = 1
ElseIf Combo1.ListIndex = 8 Then
    Combo2.AddItem "Printer", 0
    Combo2.AddItem "Publisher", 1
    Combo2.ListIndex = 1
End If
    Text1.Text = Trim(ReadFromFile(Combo1.Text))
    Command1.Caption = IIf(Combo1.ListIndex = 0, "Update Super User", IIf(Combo1.ListIndex = 1, "Activate Key", IIf(Combo1.ListIndex = 2, "Update Database Type", IIf(Combo1.ListIndex = 3, "Update Server Name", IIf(Combo1.ListIndex = 4, "Update Server User", IIf(Combo1.ListIndex = 5, "Update Server Passward", IIf(Combo1.ListIndex = 6, "Update Tally Port", IIf(Combo1.ListIndex = 7, "Update Server ID", IIf(Combo1.ListIndex = 8, "Update Client ID", "")))))))))
End Sub
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function
Private Sub lblDescription_Click()
Dim R As Long
      R = ShellExecute(0, "open", "http://www.easyinfosolution.com", 0, 0, 1)
End Sub
Private Sub Command4_Click()
        MajorFlag = True
    If UpdateComp(CompCode, False, False, True, True) Then
        Call MsgBox("Successfully Updated Masters !", vbInformation, App.Title)
        MajorFlag = False
    Else
        DisplayError ("Failed to Updated Master")
    End If
End Sub
Private Sub Command3_Click()
        MajorFlag = False
    If UpdateComp(CompCode, False, False, True, False) Then
    If CompCode = "" Then Call MsgBox("Please Login Company !!!", vbInformation, App.Title): Exit Sub
        Call MsgBox("Successfully Updated Version !!!", vbInformation, App.Title)
    Else
        DisplayError ("Failed to Update Version")
    End If
End Sub
