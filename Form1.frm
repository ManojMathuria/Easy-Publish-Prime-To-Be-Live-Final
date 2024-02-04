VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Text            =   $"Form1.frx":0000
      Top             =   360
      Width           =   10935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim ie As InternetExplorer
    Set ie = New InternetExplorer

    ' Make sure the URL is provided in the text box.
    If Text1.Text = "" Then
        MsgBox "Please enter the OneDrive file URL."
        Set ie = Nothing
        Exit Sub
    End If

    ' Navigate to the OneDrive file.
    ie.Visible = True ' You can set this to True for debugging.
    ie.navigate Text1.Text
     ret = GetUserName(lpBuff, 25)
    SourceFileUrl = "https://onedrive.live.com/?authkey=%21AAi6Oa8yimMPjUs&cid=A3BEF1B4FF3CDACB&id=A3BEF1B4FF3CDACB%2145852&parId=A3BEF1B4FF3CDACB%2145708&o=OneUp"
    Shell "C:\WINDOWS\explorer.exe """ & SourceFileUrl & "", vbNormalFocus
    ' Wait for the page to load (you can adjust the delay as needed).
    Do Until ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop

    ' Find and click the download button (if available).
    Dim doc As HTMLDocument
    Set doc = ie.document

    ' Replace 'ButtonText' with the text on the download button.
    Dim downloadButton As Object
    For Each downloadButton In doc.getElementsByTagName("Download")
        If downloadButton.innerText = "Download" Then
            downloadButton.Click
            Exit For
        End If
    Next

    ' Wait for the download to complete (you can adjust the delay as needed).
    ' You might need to handle file dialog boxes if they appear.
    
    ' Close Internet Explorer.
    ie.Quit
    Set ie = Nothing
End Sub

