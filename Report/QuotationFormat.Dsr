VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rptQuotationFormat 
   ClientHeight    =   9870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16095
   OleObjectBlob   =   "QuotationFormat.dsx":0000
End
Attribute VB_Name = "rptQuotationFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Section2_Format(ByVal pFormattingInfo As Object)
    On Error Resume Next
    With Section2.ReportObjects
        Set .Item("Picture1").FormattedPicture = LoadPicture(IIf(FileExist(App.Path & "\Icon\Logo" & CompCode & ".jpg"), App.Path & "\Icon\Logo" & CompCode & ".jpg", ""))
    End With
End Sub
