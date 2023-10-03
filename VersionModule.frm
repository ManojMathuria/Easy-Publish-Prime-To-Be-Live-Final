VERSION 5.00
Begin VB.Form FrmModule 
   Caption         =   "Select Download"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2805
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "30 Days Free Trial"
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton Command2 
         Caption         =   "EasyPublish 23.4 "
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
         TabIndex        =   2
         ToolTipText     =   "30 Days Free Trial"
         Top             =   1920
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "EasyPublish 22.12 "
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
         TabIndex        =   1
         ToolTipText     =   "Free Under Subscription"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Select EasyPublish Module"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
   End
End
Attribute VB_Name = "FrmModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Sub Command1_Click() 'EasyPublish To Be Live 22.12
SourceFileUrl = "https://onedrive.live.com/?authkey=%21AHRTNqbmWcbrE7A&cid=A3BEF1B4FF3CDACB&id=A3BEF1B4FF3CDACB%2145851&parId=A3BEF1B4FF3CDACB%2145707&o=OneUp"
FrmVersionUpdate.txtSourceFileUrl.Text = SourceFileUrl
DownloadGoogleDriveWithFilename
Unload Me
FrmVersionUpdate.Command3.SetFocus
    Sendkeys "{TAB}", True
    KeyCode = 0
End Sub
Private Sub Command2_Click() 'EasyPublish To Be Live 23.4
SourceFileUrl = "https://onedrive.live.com/?authkey=%21AAi6Oa8yimMPjUs&cid=A3BEF1B4FF3CDACB&id=A3BEF1B4FF3CDACB%2145852&parId=A3BEF1B4FF3CDACB%2145708&o=OneUp"
FrmVersionUpdate.txtSourceFileUrl.Text = SourceFileUrl
DownloadGoogleDriveWithFilename
Unload Me
End Sub
Private Sub Form_Load()
    ' Get the user name minus any trailing spaces found in the name.
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    DestinationFolder = App.Path
    DestinationFile = DestinationFolder & "\EasyPublish.exe"
If Dir(DestinationFile, vbDirectory) <> "" Then
    GetFileVersion (DestinationFile)
    DestinationFileVersion = FileVersion
    sMajor = Format(Major, "00")
    sMinor = Format(Minor, "00")
    sRevision = Format(Revision, "00")
    sRelease = Format(Release, "00")
End If

If sMajor + "." + sMinor <> "22.12" Then
FrmModule.Command1.FontStrikethru = True
       Sendkeys "{TAB}", True
       KeyCode = 0
ElseIf sMajor + "." + sMinor <> "23.4" Then
FrmModule.Command2.FontStrikethru = True
       Sendkeys "{TAB}", True
       KeyCode = 0
End If
End Sub
