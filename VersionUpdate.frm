VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmVersionUpdate 
   Caption         =   "Software Support"
   ClientHeight    =   5475
   ClientLeft      =   1185
   ClientTop       =   1545
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   9720
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   555
      Left            =   8640
      TabIndex        =   24
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2565
      Left            =   4200
      TabIndex        =   20
      ToolTipText     =   "30 Days Free Trial"
      Top             =   2840
      Width           =   5415
      Begin VB.CommandButton btnCopyFile 
         Caption         =   "Update Software Version "
         BeginProperty Font 
            Name            =   "ArialMT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1380
         TabIndex        =   5
         Top             =   1680
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Download Version "
         BeginProperty Font 
            Name            =   "ArialMT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1380
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Check Version "
         BeginProperty Font 
            Name            =   "ArialMT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1380
         TabIndex        =   2
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Step : 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Step : 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Step : 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2565
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "30 Days Free Trial"
      Top             =   2840
      Width           =   3975
      Begin VB.CommandButton Command4 
         Caption         =   "Easy Publish Prime Setup  v23.06.08"
         BeginProperty Font 
            Name            =   "ArialMT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Free Under Subscription"
         Top             =   1245
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Download EasyPublish Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   3615
      End
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   1020
      Left            =   120
      Picture         =   "VersionUpdate.frx":0000
      ScaleHeight     =   674.24
      ScaleMode       =   0  'User
      ScaleWidth      =   674.24
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   1020
   End
   Begin VB.TextBox txtSourceFileUrl 
      Height          =   315
      HideSelection   =   0   'False
      Left            =   1920
      TabIndex        =   12
      Text            =   $"VersionUpdate.frx":0B5D
      Top             =   1440
      Width           =   6615
   End
   Begin VB.TextBox txtDestFileName 
      Height          =   315
      HideSelection   =   0   'False
      Left            =   1920
      TabIndex        =   10
      Top             =   2520
      Width           =   6615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnSelectFolder 
      Caption         =   "..."
      Height          =   315
      Left            =   8640
      TabIndex        =   8
      ToolTipText     =   "Select Destination Folder"
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txtDestFolder 
      Height          =   315
      HideSelection   =   0   'False
      Left            =   1920
      TabIndex        =   7
      Text            =   "Select Destination Folder..."
      Top             =   2160
      Width           =   6615
   End
   Begin VB.CommandButton btnSelectFile 
      Caption         =   "..."
      Height          =   315
      Left            =   8640
      TabIndex        =   6
      ToolTipText     =   "Select File"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtSource 
      Height          =   315
      HideSelection   =   0   'False
      Left            =   1920
      TabIndex        =   0
      Text            =   "Select Source File..."
      Top             =   1800
      Width           =   6615
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
      Left            =   1320
      TabIndex        =   17
      Top             =   960
      Width           =   8445
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
      Left            =   1320
      TabIndex        =   16
      Top             =   600
      Width           =   8325
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
      Left            =   1320
      TabIndex        =   15
      Top             =   120
      Width           =   12045
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Source Url:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Destination File Name:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Destination Folder:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Source File:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
End
Attribute VB_Name = "FrmVersionUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DestinationFileVersion, SourceFileVersion
Dim dMajor, dMinor, dRevision, dRelease
Dim sMajor, sMinor, sRevision, sRelease
Dim R As Long
Dim UserName As String
Dim lpBuff As String * 25
Dim ret As Long
Dim SourceFileUrl, SourceFileFolder, DestinationFolder
Dim DestinationFile
Dim dFlag As Boolean

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_EDITBOX = &H10
Private Const BIF_VALIDATE = &H20
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const MAX_PATH = 260

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SetCurrentDirectory Lib "kernel32" _
    Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Private Declare Function GetCurrentDirectory Lib "kernel32" _
    Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

'Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
        
Private Declare Function lstrcat Lib "kernel32" _
    Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'Down loade folder Setting
'Option Explicit

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_WRITE = &H20006
Private Const REG_SZ = 1
Private Sub Command2_Click()

    Dim objMessage As Object
    Set objMessage = CreateObject("CDO.Message")

    ' Configuration for the SMTP server (replace with your email server settings)
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtpout.secureserver.net" '"smtp.example.com"
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    
    
    ' Enable SSL/TLS encryption
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 ' cdoBasic
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "Sales@easyinfosolution.com"
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "pubprint123!@#"

    
    objMessage.Configuration.Fields.Update

    ' Email settings
    objMessage.Subject = "Test Email"
    objMessage.From = "Sales@easyinfosolution.com" '"your_email@example.com"
    objMessage.To = "Manoj.Mathuria@gmail.com" '"recipient@example.com"
    objMessage.TextBody = "This is a test email sent using VB6 and CDO with SSL/TLS encryption."

    On Error Resume Next
    objMessage.Send

    If Err.Number = 0 Then
        MsgBox "Email sent successfully!", vbInformation
    Else
        MsgBox "Error sending email: " & Err.Description, vbExclamation
    End If

    Set objMessage = Nothing
End Sub

Private Sub Command4_Click()
Dim Setup
Setup = "https://onedrive.live.com/?authkey=%21ANp5Xkhjk9F9Ums&cid=A3BEF1B4FF3CDACB&id=A3BEF1B4FF3CDACB%2146057&parId=A3BEF1B4FF3CDACB%2146056&o=OneUp"
        Shell "C:\WINDOWS\explorer.exe """ & Setup & "", vbNormalFocus
End Sub
Private Sub Form_Load()
    ' Get the user name minus any trailing spaces found in the name.
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    SourceFileFolder = "C:\Users\" & UserName & "\Downloads\EasyPublish.exe"
    
    DestinationFolder = App.Path
    DestinationFile = DestinationFolder & "\EasyPublish.exe"
    
If Dir(DestinationFile, vbDirectory) <> "" Then
    GetFileVersion (DestinationFile)
    DestinationFileVersion = FileVersion
    sMajor = Format(Major, "00")
    sMinor = Format(Minor, "00")
    sRevision = Format(Revision, "00")
    sRelease = Format(Release, "00")
    
    lblVersion.Caption = "Easy Publish |Rel  " & sMajor & "." & sRelease & " |Version " & Major & "." & Minor & "." & Revision & "." & Release & " |Production && Inventory Management System"
Else
    lblVersion.Caption = "Easy Publish |Rel  " & App.Major & "." & App.Revision & " |Version " & App.Major & "." & App.Minor & "." & App.Revision & " |Production & Inventory Management System"
End If
                
                txtSourceFileUrl.Text = SourceFileUrl
    'txtSourceFileUrl.SelStart = 0
    'txtSourceFileUrl.SelLength = Len(txtSourceFileUrl.Text)
    
                txtSource.Text = SourceFileFolder
    'txtSource.SelStart = 0
    'txtSource.SelLength = Len(txtSource.Text)
    
                txtDestFolder.Text = App.Path
    'txtDestFolder.SelStart = 0
    'txtDestFolder.SelLength = Len(txtDestFolder.Text)
                
                txtDestFileName.Text = "EasyPublish.exe"

    Command3.Enabled = False
    Label1(5).Enabled = False
    
    btnCopyFile.Enabled = False
    Label1(6).Enabled = False
       Sendkeys "{TAB}", True
End Sub
Private Sub btnSelectFile_Click() 'Source
On Error GoTo ErrHandler
    With CommonDialog1
        .CancelError = True
        .Flags = cdlOFNExplorer
        .ShowOpen
        If Not .FileName = "" Then
            txtSource.Text = .FileName
            txtDestFileName.Text = Mid(Trim(txtSource.Text), InStrRev(Trim(txtSource.Text), "\") + 1)
        Else
            txtSource.Text = "Select Source File..."
            txtDestFileName.Text = ""
        End If
    End With
    Exit Sub
ErrHandler:
    Err.Clear
    txtSource.Text = "Select Source File..."
    txtDestFileName.Text = ""
End Sub
Private Sub btnSelectFolder_Click() 'Destination
'===================================
Dim lRet As Long
Dim sBuffer As String
Dim sTitle As String
Dim tBrowseInfo As BrowseInfo
Dim sCurDir As String
Dim lPidl As Long

    sTitle = "Select Destination Folder"
    
    With tBrowseInfo
        .hWndOwner = Me.hwnd
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
                   BIF_EDITBOX Or BIF_VALIDATE Or BIF_NEWDIALOGSTYLE
    End With
    
    lRet = SHBrowseForFolder(tBrowseInfo)
    
    If lRet > 0 Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lRet, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        txtDestFolder.Text = sBuffer
    End If

End Sub
Private Sub Command1_Click()  '
GetDownloadedFolderPath
    On Error Resume Next
    Load FrmModule
    FrmModule.Command2.SetFocus
    If Err.Number <> 364 Then FrmModule.Show
End Sub
Private Sub Command3_Click() 'Check Software Version Update
    ' Get the user name minus any trailing spaces found in the name.
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    
    DestinationFolder = App.Path
    DestinationFile = DestinationFolder & "\EasyPublish.exe"
        SourceFileFolder = FrmVersionUpdate.txtSource.Text '"C:\Users\" & UserName & "\Downloads"
        If Mid(Trim(txtSource.Text), InStrRev(Trim(txtSource.Text), "\") + 1) = "EasyPublish.exe" Then
            SourceFile = FrmVersionUpdate.txtSource.Text
        Else
        SourceFile = SourceFileFolder & "\EasyPublish.exe"
        End If
    If Dir(SourceFile, vbDirectory) <> "" Then
        GetFileVersion (SourceFile)
        SourceFileVersion = FileVersion
        dMajor = Major
        dMinor = Minor
        dRevision = Revision
        dRelease = Release
    Else
        If MsgBox("        <<<   Source File EasyPublish.exe is Missing   >>> " & Chr(13) & Chr(13) & "       Click on the Retry button to Download Again !!! " & Chr(13) & "              Or" & Chr(13) & "       Click on the Cancel button to Select Source File Manually !!!", vbRetryCancel, " Source File Status") = vbRetry Then
            DownloadGoogleDriveWithFilename (SourceFileUrl)
            Exit Sub
        Else
            txtSource.SelStart = 0
            txtSource.SelLength = Len(txtSource.Text)
            txtSource.SetFocus
            Exit Sub
        End If
    End If
        
    If Dir(DestinationFile, vbDirectory) <> "" Then
        GetFileVersion (DestinationFile)
        DestinationFileVersion = FileVersion
        sMajor = Major
        sMinor = Minor
        sRevision = Revision
        sRelease = Release
    Else
    If MsgBox("         <<<   Destination File EasyPublish.exe is Missing.   >>>" & Chr(13) & Chr(13) & "         Click On the Retry button to recheck the Software Version !!! " & Chr(13) & "          OR" & Chr(13) & "         Click on the Cancel button to Update Software Version !!!", vbRetryCancel, " Destination File Status") = vbRetry Then
         Exit Sub
    Else
            dFlag = True
            Update_Versions
            Command3.Enabled = False
            Label1(5).Enabled = False
            btnCopyFile.Enabled = False
            Label1(6).Enabled = False
            Unload Me: Exit Sub
      End If
    End If
    If SourceFileVersion = DestinationFileVersion Then
        If SourceFileVersion <> "" Or DestinationFileVersion <> "" Then
        MsgBox "<<<   EasyPublish Version Up-to-date   >>>" & Chr(13) & "Latest Version :  " & SourceFileVersion, vbInformation, " Version Up To Date"
        Shell "C:\WINDOWS\explorer.exe """ & DestinationFile & "", vbNormalFocus
        End If
        Unload Me: Exit Sub
    Else
        MsgBox "<<<   EasyPublish Version out-of-date   >>>" & Chr(13) & "Existing Version:  " & DestinationFileVersion & Chr(13) & "Latest Version :  " & SourceFileVersion, vbCritical, " Version Out of Date"
    End If
    Command3.Enabled = False
    Label1(5).Enabled = False
    
    btnCopyFile.Enabled = True
    btnCopyFile.SetFocus
    Label1(6).Enabled = True
End Sub
Private Sub btnCopyFile_Click()
Update_Versions
End Sub
Sub Update_Versions()
    Dim msg As String
    msg = "Destination folder already contains file with the same name.!!" & vbNewLine
    msg = msg & "Select YES !!! " & vbNewLine & "if you wish to overwrite existing file." & vbNewLine
    msg = msg & "Otherwise select NO !!!" & vbNewLine & "and change destination file name."
    ' Get the user name minus any trailing spaces found in the name.
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    DestinationFolder = App.Path
    DestinationFile = DestinationFolder & "\EasyPublish.exe"
    
    SourceFileFolder = FrmVersionUpdate.txtSource.Text '"C:\Users\" & UserName & "\Downloads"
    SourceFile = SourceFileFolder & "\EasyPublish.exe"

    If Dir(DestinationFile, vbDirectory) <> "" Then
        GetFileVersion (SourceFile)
        DestinationFileVersion = FileVersion
        dMajor = Major
        dMinor = Minor
        dRevision = Revision
        dRelease = Release
    Else
    
    If dFlag = False Then MsgBox "<<<   Destination File EasyPublish.exe is Missing   >>>" & Chr(13) & Chr(13) & "         Click on Update Software Version ", vbInformation, " Destination File Status"
    End If
    
    If Dir(SourceFile, vbDirectory) <> "" Then
        GetFileVersion (DestinationFile)
        SourceFileVersion = FileVersion
        sMajor = Major
        sMinor = Minor
        sRevision = Revision
        sRelease = Release
    Else
        If MsgBox("<<<   Source File EasyPublish.exe is Missing   >>>" & Chr(13) & Chr(13) & "       Do you want's Retry to Download Again ? ", vbRetryCancel, " Source File Status") = vbRetry Then
            DownloadGoogleDriveWithFilename
            Exit Sub
        Else
            txtSource.SelStart = 0
            txtSource.SelLength = Len(txtSource.Text)
            txtSource.SetFocus
            Exit Sub
        End If
    End If
    If Not Dir(Trim(txtSource.Text & "\EasyPublish.exe")) = "" Then 'Not Missing Source File
        If Not Dir(Trim(txtDestFolder.Text), vbDirectory) = "" Then 'Not Missing Destination Folder
                If Not Right(Trim(txtDestFolder.Text), 1) = "\" Then
                    txtDestFolder.Text = Trim(txtDestFolder.Text) & "\"
                End If
                If Not Dir(SourceFile) = "" Then
                    If SourceFileVersion <> "" And DestinationFileVersion <> "" And (SourceFileVersion <> DestinationFileVersion) Then
                        If txtDestFileName.Text = "EasyPublish.exe" Then
                            If MsgBox(msg, vbInformation + vbYesNo, " File Exists") = vbYes Then
                                Kill DestinationFile
                                FileCopy Trim(txtSource.Text & "\EasyPublish.exe"), DestinationFile
                                MsgBox "EasyPublish Version Updated !!!", vbInformation, " Updated"
                                Shell "C:\WINDOWS\explorer.exe """ & DestinationFile & "", vbNormalFocus
                                Unload Me: Exit Sub
                                
                            Else
                                txtDestFileName.SelStart = 0
                                txtDestFileName.SelLength = Len(txtDestFileName.Text)
                                txtDestFileName.SetFocus
                                Exit Sub
                            End If
                        Else
                            DestinationFile = DestinationFolder & "\" & txtDestFileName
                            FileCopy Trim(txtSource.Text), DestinationFile
                             MsgBox "Software Updated !!!", vbInformation, " Updated"
                            Shell "C:\WINDOWS\explorer.exe """ & DestinationFile & "", vbNormalFocus
                            If dFlag = False Then Unload Me: Exit Sub
                        End If
                    ElseIf Dir(DestinationFile) = "" Then
                        FileCopy Trim(txtSource.Text), DestinationFile
                         MsgBox "Software Updated !!!", vbInformation, " Updated"
                        Shell "C:\WINDOWS\explorer.exe """ & DestinationFile & "", vbNormalFocus
                        If dFlag = False Then Unload Me: Exit Sub
                    Else
                        Shell "C:\WINDOWS\explorer.exe """ & DestinationFile & "", vbNormalFocus
                        Unload Me: Exit Sub
                    End If
                End If
        Else
            MsgBox "Please select destination folder.", vbExclamation, " Missing Destination Folder"
        End If
    Else    'Sourse File Missing
        MsgBox "Please select source file.", vbExclamation, " Missing Source File"
    End If
End Sub
Private Sub ChangeChromeDefaultDownloadFolder(ByVal folderPath As String)
    Dim regKey As Long
    Dim disposition As Long
    
    ' Open the registry key for Chrome's download folder setting
    If RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", 0, vbNullString, 0, KEY_WRITE, 0, regKey, disposition) = 0 Then
        ' Set the new default download folder path
        RegSetValueEx regKey, "{374DE290-123F-4565-9164-39C4925E467B}", 0, REG_SZ, ByVal folderPath & "\", Len(folderPath) + 1
        
        ' Close the registry key
        RegCloseKey regKey
    End If
End Sub
Private Sub Command_Click()
    ' Specify the new default download folder path
    Dim newFolderPath As String
    newFolderPath = "C:\Users\" & UserName & "\Downloads"
    ' Call the subroutine to change the default download folder setting
    ChangeChromeDefaultDownloadFolder newFolderPath
    MsgBox "Default Download Folder setting {" & newFolderPath & "} updated."
End Sub

