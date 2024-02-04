Attribute VB_Name = "ModVersionUpdate"
Option Explicit
    
    Public SourceFile
    Dim lpBuff As String * 25
    Dim ret As Long
    Dim UserName As String
    Dim SourceFileUrl, SourceFileFolder, DestinationFolder


    Private Declare Function GetCurrentDirectory Lib "kernel32" _
        Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    
    
    Public FileVersion, Major, Minor, Revision, Release
    Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    
    Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
    Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
    Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
    
    Public Type VS_FIXEDFILEINFO
       dwSignature As Long
       dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
       dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
       dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
       dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
       dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
       dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
       dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
       dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
       dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
       dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
       dwFileFlagsMask As Long        '  = &h3F for version "0.42"
       dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
       dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
       dwFileType As Long             '  e.g. VFT_DRIVER
       dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
       dwFileDateMS As Long           '  e.g. 0
       dwFileDateLS As Long           '  e.g. 0
    End Type
Public Function GetFileVersion(ByVal FileName As String) As String
'Windows API function declarations
   Dim nDummy As Long
   Dim sBuffer()         As Byte
   Dim nBufferLen        As Long
   Dim lplpBuffer       As Long
   Dim udtVerBuffer      As VS_FIXEDFILEINFO
   Dim puLen     As Long
      
   nBufferLen = GetFileVersionInfoSize(FileName, nDummy)
   
   If nBufferLen > 0 Then
   
        ReDim sBuffer(nBufferLen) As Byte
        Call GetFileVersionInfo(FileName, 0&, nBufferLen, sBuffer(0))
        Call VerQueryValue(sBuffer(0), "\", lplpBuffer, puLen)
        Call CopyMemory(udtVerBuffer, ByVal lplpBuffer, Len(udtVerBuffer))
        
        GetFileVersion = udtVerBuffer.dwFileVersionMSh & "." & udtVerBuffer.dwFileVersionMSl & "." & udtVerBuffer.dwFileVersionLSh & "." & udtVerBuffer.dwFileVersionLSl
        FileVersion = GetFileVersion
        Major = udtVerBuffer.dwFileVersionMSh
        Minor = udtVerBuffer.dwFileVersionMSl
        Revision = udtVerBuffer.dwFileVersionLSh
        Release = udtVerBuffer.dwFileVersionLSl
    End If
End Function
Public Sub Sendkeys(Text As Variant, Optional Wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(Text), Wait
        Set WshShell = Nothing
End Sub
Function DownloadGoogleDriveWithFilename()
    ' Get the user name minus any trailing spaces found in the name.
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    SourceFileFolder = "C:\Users\" & UserName & "\Downloads"
    SourceFile = SourceFileFolder & "\EasyPublish.exe"
    SourceFileUrl = FrmVersionUpdate.txtSourceFileUrl.Text
If Dir(SourceFile, vbDirectory) <> "" Then
    Kill SourceFile
End If
If Dir(SourceFile, vbDirectory) = "" Then
    Shell "C:\WINDOWS\explorer.exe """ & SourceFileUrl & "", vbNormalFocus
End If
    FrmVersionUpdate.Command1.Enabled = False
    FrmVersionUpdate.Label1(4).Enabled = False
    FrmVersionUpdate.Command3.Enabled = True
    FrmVersionUpdate.Label1(5).Enabled = True
End Function


