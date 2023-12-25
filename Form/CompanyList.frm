VERSION 5.00
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCompanyList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Companies..."
   ClientHeight    =   4800
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10935
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   4380
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
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
      Caption         =   " Find"
      Alignment       =   0
      FillColor       =   9164542
      TextColor       =   0
      Picture         =   "CompanyList.frx":0000
      Picture         =   "CompanyList.frx":001C
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
      Top             =   4380
      Width           =   10365
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "CompanyList.frx":0038
      Height          =   4215
      Left            =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   9164542
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Col0"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Col1"
         Caption         =   "Financial Year"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Col5"
         Caption         =   "Cost Center"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   5940.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2610.142
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1725.165
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCompanyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrevStr As String, dblBookMark As Double, SortOrder As String, CompanyExists As Boolean
Dim rstDBList As New ADODB.Recordset, rstCompanyMaster As New ADODB.Recordset, rstCompanyList As New ADODB.Recordset
Private Sub Form_Load()
    Dim Cnt As Integer
    On Error GoTo ErrorHandler
    BusySystemIndicator True
    CenterForm Me
    cnDatabase.CursorLocation = adUseClient
    If cnDatabase.State = adStateOpen Then cnDatabase.Close
    cnDatabase.Open "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=Master;Data Source=" & ServerName
    rstDBList.Open "SELECT Name FROM Master.sys.Databases  WHERE LEFT(Name,2)='EP' AND Len(Name)=5 ORDER BY Name", cnDatabase, adOpenKeyset, adLockReadOnly
    rstDBList.MoveFirst
    Do While Not rstDBList.EOF
        If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
        rstCompanyMaster.Open "SELECT Name+' [" & Mid(rstDBList.Fields("Name").Value, 3, 3) & "]' As Col01,REPLACE(CONVERT(CHAR(11),FinancialYearFrom,113),' ','-')+' To '+REPLACE(CONVERT(CHAR(11),FinancialYearTo,113),' ','-') As Col02, '" & Mid(rstDBList.Fields("Name").Value, 3, 3) & "' As Col03,LTRIM(YEAR(FinancialYearFrom)) As Col04,Code As Col05,* FROM " & rstDBList.Fields("Name").Value & "..CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
        'MCGroup,FinancialYearFrom,FinancialYearTo,TallyIntegration,BusyIntegration,FYCode
        If rstCompanyList.State = adStateClosed Then
            rstCompanyList.Open "SELECT Name As Col01,Name As Col02,Name As Col03,Name As Col04,Code As Col05,* FROM " & rstDBList.Fields("Name").Value & "..CompanyMaster WHERE Code=''", cnDatabase, adOpenKeyset, adLockOptimistic
            'MCGroup,FinancialYearFrom,FinancialYearTo,TallyIntegration,BusyIntegration,FYCode
            Set rstCompanyList.ActiveConnection = Nothing
        End If
        CompanyExists = True
        With rstCompanyList
            rstCompanyMaster.MoveFirst
            Do Until rstCompanyMaster.EOF
                .AddNew
                .Fields("Col01").Value = rstCompanyMaster.Fields("Col01").Value
                .Fields("Col02").Value = rstCompanyMaster.Fields("Col02").Value
                .Fields("Col03").Value = rstCompanyMaster.Fields("Col03").Value 'Company No.
                .Fields("Col04").Value = rstCompanyMaster.Fields("Col04").Value 'Financial Year
                .Fields("MCGroup").Value = rstCompanyMaster.Fields("MCGroup").Value
                .Fields("FinancialYearFrom").Value = rstCompanyMaster.Fields("FinancialYearFrom").Value
                .Fields("FinancialYearTo").Value = rstCompanyMaster.Fields("FinancialYearTo").Value
                .Fields("TallyIntegration").Value = rstCompanyMaster.Fields("TallyIntegration").Value
                .Fields("BusyIntegration").Value = rstCompanyMaster.Fields("BusyIntegration").Value
                .Fields("FYCode").Value = rstCompanyMaster.Fields("FYCode").Value
                .Fields("Col05").Value = rstCompanyMaster.Fields("Col05").Value
                .Update
                rstCompanyMaster.MoveNext
            Loop
        End With
        rstDBList.MoveNext
    Loop
    If Not CompanyExists Then DisplayError ("No Company Exists"): Call CloseForm(Me): Exit Sub
    rstCompanyList.Sort = "Col01,FinancialYearFrom DESC"
    DataGrid1.Columns(0).DataField = rstCompanyList.Fields(0).Name
    DataGrid1.Columns(1).DataField = rstCompanyList.Fields(1).Name
    DataGrid1.Columns(2).DataField = rstCompanyList.Fields(4).Name
    Set DataGrid1.DataSource = rstCompanyList
    If (Not rstCompanyList.EOF) And (Not rstCompanyList.BOF) Then
        With DataGrid1.SelBookmarks
            .Add rstCompanyList.Bookmark
            If .Count <> 0 Then .Remove 0
            .Add rstCompanyList.Bookmark
        End With
    End If
    SortOrder = "Col01"
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    DisplayError ("Failed to connect to database")
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        DataGrid1_DblClick
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        CompCode = ""
        Call CloseForm(Me)
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstCompanyList)
End Sub
Private Sub Text1_Change()
On Error Resume Next
    With rstCompanyList
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        If Not CheckEmpty(Text1.Text, False) Then
            If SortOrder = "Name" Then .Filter = "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'" Else .Filter = "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
            If .EOF Then
            .Filter = adFilterNone
                .MoveFirst
                If PrevStr <> "" And Len(Text1.Text) > 1 Then If dblBookMark <> 0 Then .Bookmark = dblBookMark Else PrevStr = ""
                Beep
                DisplayError ("Spelling Error")
                Text1.Text = PrevStr
                Sendkeys "{End}"
            Else
                PrevStr = Text1.Text
                dblBookMark = DataGrid1.Bookmark
            End If
        Else
            .Filter = adFilterNone
            PrevStr = ""
        End If
        If Not (.EOF Or .BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
    End With
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    If rstCompanyList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstCompanyList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstCompanyList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstCompanyList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstCompanyList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstCompanyList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstCompanyList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstCompanyList
            .MoveLast
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    End If
    If KeyProcessed Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
                .Add rstCompanyList.Bookmark
        End With
        KeyProcessed = False
        KeyCode = 0
    End If
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Static AD As String
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    If AD = "Asc" Then
        rstCompanyList.Sort = "[" + SortOrder & "] Desc"
        AD = "Desc"
    Else
        rstCompanyList.Sort = "[" + SortOrder & "] Asc"
        AD = "Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstCompanyList.EOF Or rstCompanyList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub DataGrid1_DblClick()
    On Error Resume Next
    Dim GoodsType As String
    With rstCompanyList
        If (Not .EOF) And (Not .BOF) Then
            CompCode = .Fields("Col03").Value
            MCGroup = .Fields("MCGroup").Value
            FinancialYearFrom = .Fields("FinancialYearFrom").Value
            FinancialYearTo = .Fields("FinancialYearTo").Value
            FinancialYear = .Fields("Col04").Value
            TallyIntegration = .Fields("TallyIntegration").Value
            BusyIntegration = .Fields("BusyIntegration").Value
            FYCode = .Fields("FYCode").Value
            CostCenter = .Fields("Col05").Value
            If Trim(ReadFromFile("Goods Type")) = "" Then
               GoodsType = InputBox("Good Types", , "Goods")
               WriteToFile "Goods Type", GoodsType
            End If
        End If
    End With
    Call CloseForm(Me)
End Sub
