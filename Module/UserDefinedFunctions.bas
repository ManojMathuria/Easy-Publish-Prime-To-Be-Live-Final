Attribute VB_Name = "ModUserDefinedFunctions"
Option Explicit
Dim CompAlias As String, compName As String, ClientID As Variant
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstEasyPublishVersion As New ADODB.Recordset
Public MajorFlag As Boolean
Public RenewFlag As Boolean, LaterFlag As Boolean
Public dueDate As String, DaysLeft As Variant, dDay As String, dMonth As String, dYear As String
Public ServerID As String, UniqueDate As String
Public Major As Variant, Minor As Variant, Revision As Variant
Public slCode As String, slName As String   'Selection List Code & Name
Public cnDatabase As New ADODB.Connection
Public cnCompany As New ADODB.Connection
Public cnBusy As New ADODB.Connection
Public FinancialYearFrom As Date, FinancialYearTo As Date, FinancialYear As String, FYCode As String
Global FSO As New FileSystemObject
Global CompCode As String, MCGroup As String
Global DatabasePath As String
Global DatabaseType
Global SearchOrder As Integer
Global SelectionType As String
Global LoginSuccess As Boolean
Global UserCode As String
Global UserName As String
Global UserLevel As String
Global AllowMastersModification As Integer
Global AllowMastersDeletion As Integer
Global AllowTransactionsModification As Integer
Global AllowTransactionsDeletion As Integer
Global ServerName As String
Global ServerUser As String
Global ServerPassword As String
Global ConnectionString As String
Global LoginPassword As String
Global AbortPO As Boolean
Global LYDataExist As Boolean
Global VchApprovalRights As Boolean
Global BusyIntegration As Boolean, TallyIntegration As Boolean
Dim LocalHwnd As Long
Dim LocalPrevWndProc As Long
Dim MyControl As Object
Dim LeftHand_Odd() As Variant
Dim LeftHand_Even() As Variant
Dim Right_Hand() As Variant
Dim Parity() As Variant
Dim BarH As Long
Dim xObj As Object
Dim xPos As Long, xTop As Long
Public Const MF_POPUP = &H10&
Public Const WM_SETTEXT = &HC
Public Const WM_CLOSE = &H10
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWNORMAL = 1
Public Const SC_RESTORE = &HF120&
Public Const WM_SYSCOMMAND = &H112
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETITEMHEIGHT = &H154
Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_TYPE = &H10
Public Const MFT_SEPARATOR = &H800
'Public Const MFT_STRING = &H0
'Public Const MFS_ENABLED = &H0
Public Const TPM_LEFTALIGN = &H0
Public Const TPM_TOPALIGN = &H0
Public Const TPM_NONOTIFY = &H80
Public Const TPM_RETURNCMD = &H100
Public Const TPM_LEFTBUTTON = &H0
Public Const MF_BYPOSITION = &H400&
Public Const AW_BLEND = &H80000 ' Uses a fade effect. This flag can be used only if hwnd is a top-level window.
Public Const AW_HIDE = &H10000 ' Hides the window. By default, the window is shown.
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Public Type POINTAPI
        x As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type
Public Type POINT_TYPE
    x As Long
    Y As Long
End Type
'Public Type MENUITEMINFO
'        cbSize As Long
'        fMask As Long
'        fType As Long
'        fState As Long
'        wID As Long
'        hSubMenu As Long
'        hbmpChecked As Long
'        hbmpUnchecked As Long
'        dwItemData As Long
'        dwTypeData As String
'        cch As Long
'End Type
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Public Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
'Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal uFlags As Long, ByVal x As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal prcRect As Long) As Long
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As Byte) As Long
Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Function CheckEmpty(ByVal strExpression As Variant, ByVal xDspMsg As Boolean) As Boolean
    If LTrim(RTrim(strExpression)) = "" Or IsNull(strExpression) Then
       If xDspMsg Then DisplayError ("Mandatory Field")
       CheckEmpty = True
    End If
End Function
Public Function CheckNull(ByVal Expression As Variant) As String
    If IsNull(Expression) Then CheckNull = "" Else CheckNull = Expression
End Function
Public Sub LoadSelectionList(ByRef xRecordset As Recordset, ByVal xListHeader As String, ByVal xColumn1Header As String, Optional ByVal xColumn2Header As String, Optional ByVal xColumn3Header As String, Optional ByVal xColumn4Header As String)
    Set FrmSelectionList.rstSelectionList = xRecordset
    FrmSelectionList.rstSelectionList.Sort = "Col0 Asc"
    FrmSelectionList.Caption = xListHeader
    FrmSelectionList.DataGrid1.Columns(0).Caption = xColumn1Header
    If xColumn2Header = "" Then
       FrmSelectionList.Width = 7320
       FrmSelectionList.DataGrid1.Width = 7130
       FrmSelectionList.Text1.Width = 6650
       FrmSelectionList.DataGrid1.Columns(1).Visible = False
    Else
       FrmSelectionList.Width = 10510 '8710
       FrmSelectionList.DataGrid1.Width = 10330 '8530
       FrmSelectionList.Text1.Width = 9850 '8050
       FrmSelectionList.DataGrid1.Columns(0).Width = 8500
       FrmSelectionList.DataGrid1.Columns(1).Width = 1200
       FrmSelectionList.DataGrid1.Columns(1).Visible = True
       FrmSelectionList.DataGrid1.Columns(1).Caption = xColumn2Header
    End If
    If xColumn3Header = "" Then
        FrmSelectionList.DataGrid1.Columns(2).Visible = False
    Else
        FrmSelectionList.DataGrid1.Columns(2).Visible = True
        FrmSelectionList.DataGrid1.Columns(2).Caption = xColumn3Header
    End If
    If xColumn4Header = "" Then
        FrmSelectionList.DataGrid1.Columns(3).Visible = False
    Else
        FrmSelectionList.DataGrid1.Columns(3).Visible = True
        FrmSelectionList.DataGrid1.Columns(3).Caption = xColumn4Header
    End If
    FrmSelectionList.FindFieldName = "Col0"
    FrmSelectionList.Top = (Screen.Height - FrmSelectionList.Height) / 2
    FrmSelectionList.Left = Screen.Width - FrmSelectionList.Width - 130
    Load FrmSelectionList
End Sub
Public Sub DisplaySelectionList(ByRef xNameTextBox As TextBox, ByRef xCode As String)
    FrmSelectionList.Text1.Text = ""
    FrmSelectionList.txtName = xNameTextBox.Text
    FrmSelectionList.TxtCode = xCode
    FrmSelectionList.Top = (Screen.Height - FrmSelectionList.Height) / 2
    FrmSelectionList.Left = Screen.Width - FrmSelectionList.Width - 130
    FrmSelectionList.Show vbModal
    xNameTextBox.Text = FrmSelectionList.txtName
    xCode = FrmSelectionList.TxtCode
End Sub
Public Function CheckExists(ByVal xTextBox As TextBox, ByVal xField As String, ByVal xRecordset As Recordset, ByRef xCode As String) As Boolean
    xTextBox.Text = LTrim(RTrim(xTextBox.Text))
    If xRecordset.RecordCount > 0 Then xRecordset.MoveFirst
    xRecordset.Find "[" & xField & "] = '" & FixQuote(xTextBox.Text) & "'"
    If Not xRecordset.EOF Then
       CheckExists = True
       xCode = xRecordset.Fields("Code").Value
    End If
End Function
Public Function FixAPIString(ByVal strInput As String) As String
    ' strips Trailing Nulls From strings returned by API Function
    Dim intPosition As Integer
    
    intPosition = InStr(1, strInput, Chr(0))
    If intPosition Then
        FixAPIString = Left(strInput, intPosition - 1)
    Else
        FixAPIString = strInput     ' No Nulls Found
    End If
End Function
Public Function GetWindowsSystemDirectory() As String
    Dim DirectoryPath As String * 255
    
    Call GetSystemDirectory(DirectoryPath, 255)
    DirectoryPath = FixAPIString(DirectoryPath)
    GetWindowsSystemDirectory = RTrim(DirectoryPath)
End Function
Public Function GetWindowsTempDirectory() As String
    Dim DirectoryPath As String * 255
    Call GetTempPath(255, DirectoryPath)
    DirectoryPath = FixAPIString(DirectoryPath)
    GetWindowsTempDirectory = RTrim(DirectoryPath)
End Function
Public Sub WriteToFile(ByVal xKeyName As String, ByVal xString As String)
    WritePrivateProfileString "EasyPublish", xKeyName, xString, App.Path + "\" + IIf(CheckEmpty(Command$, False), "EasyPublish", Command$) + ".ini"
End Sub
Public Function ReadFromFile(ByVal xKeyName As String) As String
    Dim sReturn As String * 255
    GetPrivateProfileString "EasyPublish", xKeyName, "", sReturn, 255, App.Path + "\" + IIf(CheckEmpty(Command$, False), "EasyPublish", Command$) + ".ini"
    sReturn = FixAPIString(sReturn)
    ReadFromFile = RTrim(sReturn)
End Function
Public Sub BusySystemIndicator(ByVal bVal As Boolean)
    If bVal Then MdiMainMenu.MousePointer = vbHourglass Else MdiMainMenu.MousePointer = vbNormal
End Sub
Public Function NumberToWords(ByVal xNumber As Double, Optional ByVal blnPrintPaise As Boolean) As String
    Dim Amount As String, Paise As String
    Dim Crore As String, Lakh As String, Thousand As String, Hundred As String, Ten As String
    
    NumberToWords = "Rupees "
    Amount = Format(xNumber, "000000000.00")
    Paise = Format((xNumber - Int(xNumber)) * 100, "00")
    Crore = Mid(Amount, 1, 2)
    Lakh = Mid(Amount, 3, 2)
    Thousand = Mid(Amount, 5, 2)
    Hundred = Mid(Amount, 7, 1)
    Ten = Mid(Amount, 8, 2)
    If Val(Crore) > 0 Then
       NumberToWords = NumberToWords + Words(Crore) + " Crore "
    End If
    If Val(Lakh) > 0 Then
       NumberToWords = NumberToWords + Words(Lakh) + " Lakh "
    End If
    If Val(Thousand) > 0 Then
       NumberToWords = NumberToWords + Words(Thousand) + " Thousand "
    End If
    If Val(Hundred) > 0 Then
       NumberToWords = NumberToWords + Words(Hundred) + " Hundred "
    End If
    If Val(Ten) > 0 Then
       NumberToWords = NumberToWords + Words(Ten) + Space(1)
    End If
    If Val(Crore) = 0 And Val(Lakh) = 0 And Val(Thousand) = 0 And Val(Hundred) = 0 And Val(Ten) = 0 Then
       NumberToWords = NumberToWords + "Nil "
    End If
    If blnPrintPaise Then
       If Val(Paise) > 0 Then
          NumberToWords = NumberToWords + "And Paise "
          NumberToWords = NumberToWords + Words(Paise) + Space(1)
       Else
          'NumberToWords = NumberToWords + "And Paise Nil "
       End If
    End If
    NumberToWords = NumberToWords + "Only"
End Function
Public Function Words(ByVal xNumber As String) As String
    Const Ones = "One   Two   Three Four  Five  Six   Seven Eight Nine"
    Const Tens = "Ten     Twenty  Thirty  Forty   Fifty   Sixty   Seventy Eighty  Ninety"
    Const Teens = "Eleven    Twelve    Thirteen  Fourteen  Fifteen   Sixteen   Seventeen Eighteen  Nineteen"
    
    If Val(xNumber) >= 1 And Val(xNumber) <= 9 Then
       Words = Words + LTrim(RTrim(Mid(Ones, Val(xNumber) + (Val(xNumber) - 1) * 5, 6)))
    ElseIf Val(xNumber) >= 11 And Val(xNumber) <= 19 Then
       Words = Words + LTrim(RTrim(Mid(Teens, Val(Mid(xNumber, 2, 1)) + (Val(Mid(xNumber, 2, 1)) - 1) * 9, 10)))
    ElseIf (Val(xNumber) >= 20 And Val(xNumber) <= 99) Or Val(xNumber) = 10 Then
       Words = Words + LTrim(RTrim(Mid(Tens, Val(Mid(xNumber, 1, 1)) + (Val(Mid(xNumber, 1, 1)) - 1) * 7, 8)))
       If Mid(xNumber, 2, 1) <> "0" Then
          Words = Words + Space(1) + LTrim(RTrim(Mid(Ones, Val(Mid(xNumber, 2, 1)) + (Val(Mid(xNumber, 2, 1)) - 1) * 5, 6)))
       End If
    End If
End Function
Public Function FileExist(ByVal szFileName As String) As Boolean
    Dim nFileNumber As Integer
    On Error Resume Next
    nFileNumber = FreeFile
    Open szFileName For Input As nFileNumber
    If Err.Number = 0 Then FileExist = True
    Close nFileNumber
    Err.Clear
End Function
Public Function RestoreForm(ByVal wHandle As Long) As Boolean
  If IsIconic(wHandle) Then
    Call PostMessage(wHandle, WM_SYSCOMMAND, SC_RESTORE, 0)
    RestoreForm = True
  End If
End Function
Public Sub SetComboBoxDroppedWidth(ByVal xForm As Form, ByVal xComboBox As ComboBox, ByVal NumItemsToDisplay As Integer, ByVal DroppedWidth As Integer, ByVal ShowDropDown As Boolean)
    Dim pt As POINTAPI
    Dim rc As RECT
    Dim cWidth As Long
    Dim newHeight As Long
    Dim oldScaleMode As Long
    Dim itemHeight As Long
    
    If TypeOf xComboBox.Parent Is Frame Then Exit Sub
    oldScaleMode = xForm.ScaleMode
    xForm.ScaleMode = vbPixels
    cWidth = xComboBox.Width
    itemHeight = SendMessage(xComboBox.hwnd, CB_GETITEMHEIGHT, 0, ByVal 0)
    newHeight = itemHeight * (NumItemsToDisplay + 2)
    Call GetWindowRect(xComboBox.hwnd, rc)
    pt.x = rc.Left
    pt.Y = rc.Top
    Call SendMessage(xComboBox.hwnd, CB_SETDROPPEDWIDTH, ByVal DroppedWidth, ByVal 0)
    Call ScreenToClient(xForm.hwnd, pt)
    Call MoveWindow(xComboBox.hwnd, pt.x, pt.Y, xComboBox.Width, newHeight, True)
    Call SendMessage(xComboBox.hwnd, CB_SHOWDROPDOWN, ShowDropDown, ByVal 0)
    xForm.ScaleMode = oldScaleMode
End Sub
Public Function TimeDiff(STime As Date, ETime As Date) As String
    'Example : TimeDiff(Time, "05:45:00 PM")
    Dim TimeSecs, Hrs As Double
    Dim strSeconds As String
    Dim strMinutes As String
    Dim strHours As String
    If ETime < STime Then Exit Function
    'Get Total Number of seconds difference
    TimeSecs = DateDiff("S", STime, ETime)
    strHours = Int(TimeSecs / 3600)
    strMinutes = Int((TimeSecs Mod 3600) / 60)
    strSeconds = (TimeSecs Mod 3600) Mod 60
    TimeDiff = IIf(Len(strHours) = 1, String(2 - Len(strHours), "0") + strHours, strHours) + ":" + String(2 - Len(strMinutes), "0") + strMinutes + ":" + String(2 - Len(strSeconds), "0") + strSeconds
End Function
Public Function ProperCase(ByVal strInput As String) As String
     ProperCase = StrConv(strInput, vbProperCase)
End Function
Public Sub ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean)
    Dim tRC As RECT

    If bShowProgressBar Then
        'Get the size of the Panel Rectangle from the status Bar
        SendMessage MdiMainMenu.StatusBar1.hwnd, SB_GETRECT, 0, tRC
        'And convert it to twips....
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
        'Now Reparent the ProgressBar to the statusbar
        With MdiMainMenu.ProgressBar1
            SetParent .hwnd, MdiMainMenu.StatusBar1.hwnd
            .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
            .Value = 0
        End With
    Else
        'Reparent the progress bar back to the form and hide it
        SetParent MdiMainMenu.ProgressBar1.hwnd, MdiMainMenu.Picture1.hwnd
        MdiMainMenu.ProgressBar1.Visible = False
    End If
End Sub
Public Function FixQuote(ByVal strInput As String) As String
    If InStr(1, strInput, "'") > 0 Then
       FixQuote = Replace(strInput, "'", "''")
    Else
       FixQuote = strInput
    End If
End Function
Public Sub CloseMsgBox(ByVal MsgTitle As String)
   Static hwnd As Long
   Static Ticks As Long
   
   If hwnd = 0 Then
      hwnd = FindWindow(vbNullString, MsgTitle)
   End If
   Ticks = Ticks + 1
   Call SendMessage(hwnd, WM_SETTEXT, 0, ByVal MsgTitle)
   If Ticks >= 5 Then
      Call SendMessage(hwnd, WM_CLOSE, 0, ByVal 0&)
      hwnd = 0
      Ticks = 0
   End If
End Sub
Public Sub DisplayError(ByVal strErrorMsg As String)
    On Error Resume Next
    Beep
    MsgBox RTrim(LTrim(strErrorMsg)) & " !!!", vbExclamation, "Error !"
    Err.Clear
End Sub
Public Sub CloseForm(ByRef xForm As Form)
    On Error GoTo ErrorHandler
    Unload xForm
    Set xForm = Nothing
    Exit Sub
ErrorHandler:
End Sub
Public Function GenerateCode(ByVal xConnection As ADODB.Connection, ByVal strSQL As String, intLen, ByVal strFillChar As String) As Variant
    On Error GoTo ErrorHandler
    Dim rstGenerateCode As New ADODB.Recordset
    Dim xCode As String
    rstGenerateCode.Open strSQL, xConnection, adOpenKeyset, adLockReadOnly
    If IsNull(rstGenerateCode.Fields(0).Value) Then xCode = "0" Else xCode = Val(rstGenerateCode.Fields(0).Value)
    GenerateCode = Pad(RTrim(Val(xCode) + 1), strFillChar, intLen, "L")
    rstGenerateCode.Close
    Set rstGenerateCode = Nothing
    Exit Function
ErrorHandler:
    If rstGenerateCode.State = adStateOpen Then rstGenerateCode.Close
    Set rstGenerateCode = Nothing
    GenerateCode = Null
End Function
Public Function CheckDuplicate(ByVal xConnection As ADODB.Connection, ByVal TableName As String, ByVal SelectField As String, ByVal SearchField As String, ByVal SearchValue As String, Optional ByVal CheckValue As Variant, Optional ByVal AskToContinue As Boolean, Optional ByVal xFYCode As String) As Boolean
    On Error GoTo ErrorHandler
    Dim Rs As New ADODB.Recordset, SQL As String
    SQL = "SELECT " & SelectField & " FROM " & TableName & " WHERE LTRIM(RTrim(" & SearchField & ")) = '" & FixQuote(RTrim(LTrim(SearchValue))) & "'"
    If Not CheckEmpty(xFYCode, False) Then SQL = SQL + " AND FYCode='" & xFYCode & "'"
    Rs.Open SQL, xConnection, adOpenKeyset, adLockReadOnly
    If Rs.RecordCount <> 0 Then
       If (CheckValue <> Rs.Fields(0).Value) Or CheckEmpty(CheckValue, False) Then
          If AskToContinue Then
              Beep
              If MsgBox("      Duplicate Entry !" & vbCrLf & "Would you like to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then CheckDuplicate = False Else CheckDuplicate = True
          Else
              DisplayError ("Duplicate Entry"): CheckDuplicate = True
          End If
       End If
    End If
    Call CloseRecordset(Rs)
    Exit Function
ErrorHandler:
    Call CloseRecordset(Rs)
    CheckDuplicate = True
End Function
Public Function AddRecord(ByRef Rs As ADODB.Recordset) As Boolean
    On Error GoTo ErrorHandler
    Rs.AddNew
    AddRecord = True
    Exit Function
ErrorHandler:
End Function
Public Sub DeleteRecord(ByRef Rs As ADODB.Recordset, ByVal xCode As String)
    On Error GoTo ErrorHandler
    If Rs.EOF And Rs.BOF Then Exit Sub
    If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") <> vbYes Then Exit Sub
    Rs.MoveFirst
    Rs.Find "[Code] ='" & FixQuote(xCode) & "'"
    If Not Rs.EOF Then
        MdiMainMenu.MousePointer = vbHourglass
        Rs.Delete
        Rs.MoveNext
    End If
    MdiMainMenu.MousePointer = vbNormal
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to delete the record")
    Rs.CancelUpdate
    MdiMainMenu.MousePointer = vbNormal
End Sub
Public Function UpdateRecord(ByRef Rs As ADODB.Recordset) As Boolean
    Dim strErrorMessage As String, oError As Error, blnAdd As Boolean
    On Error Resume Next 'This also clears the Error object
    UpdateRecord = True 'No Error
    blnAdd = Rs.EditMode = adEditAdd
    Rs.ActiveConnection.Errors.Clear
    Screen.MousePointer = vbHourglass 'The update might take a while
    Rs.Update
    Screen.MousePointer = vbNormal
    Select Case Err.Number
        Case 0 'Check the underlying Connection Object for Errors too. Provider-specific Errors don't show up in the above Error Trap
            If Rs.ActiveConnection.Errors.Count = 0 Then 'No provider-specific errors - save was successful
               If blnAdd Then
                  If Rs.CursorLocation = adUseClient Then
                     Rs.Resync adAffectCurrent 'show default field values that may have been entered by the database
                  End If
               End If
            ElseIf Rs.ActiveConnection.Errors.Count <> 0 Then
               For Each oError In Rs.ActiveConnection.Errors
                     strErrorMessage = strErrorMessage & oError.Description & vbCr
               Next
               If Rs.ActiveConnection.Errors.Count >= 1 Then strErrorMessage = "The following error" & IIf(Rs.ActiveConnection.Errors.Count > 1, "s were", " was") & " reported by the provider: " & vbCr & strErrorMessage
               Call DisplayError(strErrorMessage) 'display all the errors
               UpdateRecord = False 'Error
                'Leave the save and Cancel Buttons showing so that user can backtrack
            End If
        Case Else
           UpdateRecord = False 'Error
    End Select
End Function
Public Function CancelRecordUpdate(ByRef Rs As ADODB.Recordset) As Boolean
    On Error GoTo ErrorHandler
    If Rs.EditMode = adEditAdd Or Rs.EditMode = adEditInProgress Then Rs.CancelUpdate
    CancelRecordUpdate = True
    Exit Function
ErrorHandler:
End Function
Public Function RefreshRecord(ByRef Rs As ADODB.Recordset) As Boolean
    On Error GoTo ErrorHandler
    ShowProgressInStatusBar (True)
    MdiMainMenu.ProgressBar1.Value = 50
    Screen.MousePointer = vbHourglass
    Rs.Requery
    MdiMainMenu.ProgressBar1.Value = 100
    ShowProgressInStatusBar (False)
    Screen.MousePointer = vbNormal
    RefreshRecord = True
  Exit Function
ErrorHandler:
    Screen.MousePointer = vbNormal
End Function
Function Pad(ByVal strExpression As String, ByVal strFillChar As String, ByVal intLength As Integer, ByVal strAlignment As String) As String
    If Len(strExpression) > intLength Then strExpression = Left(strExpression, intLength)
    If StrConv(strAlignment, vbUpperCase) = "C" Then
       Dim lPad, rPad As Integer
       lPad = Int((intLength - Len(strExpression)) / 2)
       rPad = intLength - Len(strExpression) - lPad
       Pad = String(lPad, strFillChar) & strExpression & String(rPad, strFillChar)
    ElseIf StrConv(strAlignment, vbUpperCase) = "R" Then
        Pad = strExpression & String(intLength - Len(strExpression), strFillChar)
    ElseIf StrConv(strAlignment, vbUpperCase) = "L" Then
        Pad = String(intLength - Len(strExpression), strFillChar) & strExpression
    End If
End Function
Public Sub CloseRecordset(ByRef xRecordset As ADODB.Recordset)
    On Error GoTo ErrorHandler
    If xRecordset.State = adStateOpen Then xRecordset.Close
    Set xRecordset = Nothing
    Exit Sub
ErrorHandler:
End Sub
Public Sub CloseConnection(ByVal xConnection As ADODB.Connection)
    On Error GoTo ErrorHandler
    If xConnection.State = adStateOpen Then xConnection.Close
    Set xConnection = Nothing
    Exit Sub
ErrorHandler:
End Sub
Public Function DirExist(ByVal strDir As String) As Boolean
    On Error Resume Next
    ChDir strDir
    If Err.Number <> 76 Then DirExist = True
End Function
Public Sub SetMdiButtons(bVal As Boolean, Optional ByVal EnablePrintButtons As Boolean, Optional ByVal EnableMailButton As Boolean)
    Dim Ctr As Integer
    For Ctr = 1 To 17
        MdiMainMenu.Toolbar1.Buttons(Ctr).Enabled = bVal
        If Ctr = 9 Or Ctr = 10 Then
            MdiMainMenu.Toolbar1.Buttons(8).Visible = EnablePrintButtons
            MdiMainMenu.Toolbar1.Buttons(Ctr).Visible = EnablePrintButtons
            MdiMainMenu.Toolbar1.Buttons(Ctr).Enabled = EnablePrintButtons
        End If
    Next
    MdiMainMenu.Toolbar1.Buttons(11).Visible = EnableMailButton
    MdiMainMenu.Toolbar1.Buttons(11).Enabled = EnableMailButton
End Sub
Public Sub EnableChildMenu(Optional ByVal EnablePrintButtons As Boolean, Optional ByVal EnableMailButton As Boolean)
    MdiMainMenu.Toolbar1.Buttons(18).ToolTipText = "Close"
    Call SetMdiButtons(True, EnablePrintButtons, EnableMailButton)
End Sub
Public Sub DisableChildMenu()
    Call SetMdiButtons(False)
    MdiMainMenu.Toolbar1.Buttons(18).ToolTipText = "Exit"
End Sub
Public Sub CenterForm(frm As Form)
  Dim m_lngRetVal As Long
  Dim ClientRect As RECT     'Holds the area that the form is to be centered in
  Dim TaskBarRect As RECT     'Holds the TaskBar area if in Win95
  Dim x As Variant  'temp LeftPosition
  Dim Y As Variant  'temp TopPosition

  If frm.MDIChild Then ' Check if the form is a MDIChild.
      ' Center it in the MDIParent.
      GetClientRect GetParent(frm.hwnd), ClientRect
  Else  'Center it in the available desktop area.
      ' Get the Desktop area
      Call GetClientRect(GetDesktopWindow(), ClientRect)
      ' Check for the Task Bar.
      m_lngRetVal = FindWindow("Shell_TrayWnd", vbNullString)
      ' If there is a taskbar, ie WIN95 then adjust the ClientRect.
      If m_lngRetVal Then
          Call GetWindowRect(m_lngRetVal, TaskBarRect)
          If (TaskBarRect.Right - TaskBarRect.Left) > (TaskBarRect.Bottom - TaskBarRect.Top) Then
              ' TaskBar at the Top of Screen.
              If TaskBarRect.Top <= 0 Then
                  ClientRect.Top = ClientRect.Top + TaskBarRect.Bottom
                  ' TaskBar at the Bottom of Screen.
              Else
                  ClientRect.Bottom = ClientRect.Bottom - (TaskBarRect.Bottom - TaskBarRect.Top)
              End If
          Else
              ' TaskBar is on the Left side of the Screen.
              If TaskBarRect.Left <= 0 Then
                  ClientRect.Left = ClientRect.Left + TaskBarRect.Right
                  ' TaskBar is on the Right side of the Screen.
              Else
                  ClientRect.Right = ClientRect.Right - (TaskBarRect.Right - TaskBarRect.Left)
              End If
          End If   '[TaskBar on Top of Screen?]
      End If
  End If
' Center the Form
  With frm
       x = (((ClientRect.Right - ClientRect.Left) * Screen.TwipsPerPixelX) - .Width) / 2
       Y = (((ClientRect.Bottom - ClientRect.Top) * Screen.TwipsPerPixelY) - .Height) / 2
       .Move x, Y
  End With
End Sub
Public Function SetPrinterMode(ByVal strCode As String, ByVal blnFlag As Boolean) As String
    If strCode = "cndn" Then
       If blnFlag Then
          SetPrinterMode = Chr(15)
      Else
          SetPrinterMode = Chr(18)
      End If
    ElseIf strCode = "12cpi" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(77)
      Else
          SetPrinterMode = Chr(27) & Chr(80)
      End If
    ElseIf strCode = "15cpi" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(103)
      Else
          SetPrinterMode = Chr(27) & Chr(80)
      End If
    ElseIf strCode = "19cpi" Then '12CPI Condensed
       If blnFlag Then
          SetPrinterMode = Chr(15) & Chr(27) & Chr(77)
      Else
          SetPrinterMode = Chr(15) & Chr(27) & Chr(80)
      End If
    ElseIf strCode = "bold" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(69)
      Else
          SetPrinterMode = Chr(27) & Chr(70)
      End If
    ElseIf strCode = "dblstrk" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(71)
      Else
          SetPrinterMode = Chr(27) & Chr(72)
      End If
    ElseIf strCode = "ulin" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(45) & Chr(49)
      Else
          SetPrinterMode = Chr(27) & Chr(45) & Chr(48)
      End If
    ElseIf strCode = "dwth" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(87) & Chr(49)
      Else
          SetPrinterMode = Chr(27) & Chr(87) & Chr(48)
      End If
    ElseIf strCode = "8lpi" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(48)
      Else
          SetPrinterMode = Chr(27) & Chr(50)
      End If
    ElseIf strCode = "init" Then
      SetPrinterMode = Chr(27) & Chr(64)
    ElseIf strCode = "ejec" Then
      SetPrinterMode = Chr(12)
    End If
End Function
Public Function GetTemporaryFileName() As String
    Dim lngReturnVal As Long
    Dim strTempPath As String * 255
    Dim strTempFileName As String * 255
    On Error GoTo TempNameErr
    lngReturnVal = GetTempPath(254, strTempPath)
    lngReturnVal = GetTempFileName(strTempPath & "\", "", 0, strTempFileName)
    GetTemporaryFileName = strTempFileName
    Exit Function
TempNameErr:
    Call DisplayError("Cannot retrieve Temporary FileName")
End Function
Public Function DisplayPopupMenu(ByVal hwnd As Long, Optional ByVal intMenu As Integer) As Integer
    Dim hSubMenu11 As Long, hSubMenu10 As Long, hSubMenu09 As Long, hSubMenu08 As Long, hSubMenu07 As Long, hSubMenu06 As Long, hSubMenu05 As Long, hSubMenu04 As Long, hSubMenu03 As Long, hSubMenu02 As Long, hSubMenu01 As Long, hMenu As Long ' Handle to the popup menu to display
    Dim CurPos As POINT_TYPE ' Holds the current mouse coordinates
    Dim menusel As Long ' ID of what the user selected in the popup menu
    Dim retVal As Long ' Generic return value
    If intMenu = 0 Then
        hMenu = CreatePopupMenu
        AppendMenu hMenu, 0, 1, "Add Record"
        AppendMenu hMenu, 0, 2, "Edit Record"
        AppendMenu hMenu, 0, 3, "Delete Record"
    ElseIf intMenu = 1 Then
        hSubMenu05 = CreatePopupMenu
        AppendMenu hSubMenu05, 0, 12, "Spread-Form-Format"      'TP_Single_Plate03
        AppendMenu hSubMenu05, 0, 13, "Combo-Form-Format"       'TP_Combo_Plate03
        hSubMenu04 = CreatePopupMenu
        AppendMenu hSubMenu04, 0, 9, "Multi-Form-Format"         'BP_Plate03
        AppendMenu hSubMenu04, MF_POPUP, hSubMenu05, "Single-Form-Format"
        AppendMenu hSubMenu04, 0, 11, "All"         'All_BP_TP_Plate03
        hSubMenu03 = CreatePopupMenu
        AppendMenu hSubMenu03, 0, 2, "Spread-Form-Format"       'TP_Single_Order02
        AppendMenu hSubMenu03, 0, 8, "Combo-Form-Format"        'TP_Combo_Order02
        hSubMenu02 = CreatePopupMenu
        AppendMenu hSubMenu02, 0, 1, "Multi-Form-Format"         'BP_Order02
        AppendMenu hSubMenu02, MF_POPUP, hSubMenu03, "Spread-Form-Format"
        hSubMenu01 = CreatePopupMenu
        AppendMenu hSubMenu01, MF_POPUP, hSubMenu02, "Printing"
        AppendMenu hSubMenu01, 0, 26, "Multi-Form-Format"               'CB_Order02
        AppendMenu hSubMenu01, 0, 25, "Spread-Form-Format"               'TL_Order02
        AppendMenu hSubMenu01, 0, 24, "Combo-Form-Format"               'CB_Order02
        AppendMenu hSubMenu01, 0, 3, "MISC Operation"               'TL_Order02
        AppendMenu hSubMenu01, 0, 4, "Binding Process"                  'BB_Order02
        AppendMenu hSubMenu01, 0, 5, "All"                      'ALL_BP_TP_TL_BB_Order02
        hSubMenu08 = CreatePopupMenu
        AppendMenu hSubMenu08, 0, 15, "Spread-Form-Format"                  'TP_Single_JobCard
        AppendMenu hSubMenu08, 0, 19, "Combo-Form-Format"                   'TP_Combo_Order
        hSubMenu07 = CreatePopupMenu
        AppendMenu hSubMenu07, 0, 14, "Multi-Form-Format"                    'BP_JobCard
        AppendMenu hSubMenu07, MF_POPUP, hSubMenu08, "Spread Form Format"
        hSubMenu06 = CreatePopupMenu
        AppendMenu hSubMenu06, MF_POPUP, hSubMenu07, "Printing"
        AppendMenu hSubMenu06, 0, 14, "Multi-Form-Format"       'TP_Multi_Form_JobCard
        AppendMenu hSubMenu06, 0, 15, "Spread-Form-Format"      'TP_Single_JobCard
        AppendMenu hSubMenu06, 0, 19, "Combo-Form-Format"       'Combo_JobCard
        AppendMenu hSubMenu06, 0, 16, "MISC Operation"              'TL_JobCard
        AppendMenu hSubMenu06, 0, 17, "Binding Process"                 'BB_JobCard
        AppendMenu hSubMenu06, 0, 18, "All"                     'ALL_BP_TP_TL_BB_JObCard
        hSubMenu09 = CreatePopupMenu
        AppendMenu hSubMenu09, 0, 20, "Multi-Form-Format"       'BP_Paper-Slip
        AppendMenu hSubMenu09, 0, 21, "Spread-Form-Format"      'TP_Paper-Slip
        AppendMenu hSubMenu09, 0, 23, "Combo-Form-Format"       'Combo_Paper-Slip
        AppendMenu hSubMenu09, 0, 22, "All"                     'ALL_BP_TP_Paper-Slip
        hSubMenu10 = CreatePopupMenu
        AppendMenu hSubMenu10, 0, 27, "Quotation-MF"       'BP_Quotation
        AppendMenu hSubMenu10, 0, 28, "Quotation-SF"       'TP_Quotation
        AppendMenu hSubMenu10, 0, 30, "Quotation-CF"       'Combo_Quotation
        AppendMenu hSubMenu10, 0, 29, "Quotation-OP"       'TL_Quotation
        AppendMenu hSubMenu10, 0, 31, "Quotation-BB"       'BB_Quotation
        AppendMenu hSubMenu10, 0, 32, "Quotation-All"      'ALL_BP_TP_TL_CB_Quotation
        AppendMenu hSubMenu10, 0, 33, "Excel Format"
        hSubMenu11 = CreatePopupMenu
        AppendMenu hSubMenu11, 0, 34, "Printing Planning"       'BP_Print Planning
        hMenu = CreatePopupMenu
        AppendMenu hMenu, MF_POPUP, hSubMenu01, "Jobwork"
        AppendMenu hMenu, 0, 6, "Unit Cost-Jobwork"
        AppendMenu hMenu, 0, 7, "Unit Cost"
        AppendMenu hMenu, MF_POPUP, hSubMenu04, "Plate"
        AppendMenu hMenu, MF_POPUP, hSubMenu06, "JobCard"
        AppendMenu hMenu, MF_POPUP, hSubMenu09, "Paper-Requisition-Slip"
        AppendMenu hMenu, MF_POPUP, hSubMenu10, "Quotation Format"
        AppendMenu hMenu, MF_POPUP, hSubMenu11, "Printing Planning"
    ElseIf intMenu = 2 Then
        hMenu = CreatePopupMenu
        AppendMenu hMenu, 0, 1, "Original"
        AppendMenu hMenu, 0, 2, "Duplicate"
        AppendMenu hMenu, 0, 3, "Triplicate"
        AppendMenu hMenu, 0, 4, "All"
    ElseIf intMenu = 3 Then
        hMenu = CreatePopupMenu
        AppendMenu hMenu, 0, 1, "Packing Slip"
        AppendMenu hMenu, 0, 2, "Fowarding Slip"
'        AppendMenu hMenu, 0, 3, "Private Mark"
'        AppendMenu hMenu, 0, 4, "Packing Slip With Private Mark"
    ElseIf intMenu = 4 Then
        hMenu = CreatePopupMenu
        AppendMenu hMenu, 0, 1, "Original"
        AppendMenu hMenu, 0, 2, "Pending"
    End If
    retVal = GetCursorPos(CurPos)
    menusel = TrackPopupMenu(hMenu, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_LEFTALIGN Or TPM_LEFTBUTTON, CurPos.x, CurPos.Y, 0, hwnd, 0)
    retVal = DestroyMenu(hMenu)
    DisplayPopupMenu = menusel
End Function
Public Function CalculateConsumption(ByVal xPaperType As String, ByVal xQuantity As Long, ByVal xForms As Double, ByVal xWastage As Double) As Double
    If xPaperType = "1" Then    'Book
        CalculateConsumption = CLng(xQuantity * xForms * (100 + xWastage) / 100)
    Else    'Title
        CalculateConsumption = Format((xQuantity / 2) * ((100 + xWastage) / 100), "#0")
    End If
    CalculateConsumption = CLng(Val(CalculateConsumption) / 2)
    CalculateConsumption = Int(Val(CalculateConsumption) / 500) & "." & Format(Val(CalculateConsumption) Mod 500, "000")
End Function
Public Function CalculatePaperBalance(ByVal AccountCode As String, ByVal PaperCode As String, ByVal VchCode As String, ByVal VchType, VchDate) As Double
    On Error GoTo ErrorHandler
    Dim rstPaperBal As New ADODB.Recordset
    With rstPaperBal
        If .State = adStateOpen Then .Close
        .Open "SELECT dbo.ufnGetPaperStock('" & AccountCode & "',Code,'" & IIf(VchType = "PMV", "TR", "PO") & "','" & VchCode & "','" & VchDate & "') As CurStk FROM PaperMaster WHERE Code='" & PaperCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
        
        CalculatePaperBalance = Val(CheckNull(.Fields("CurStk").Value))
        Call CloseRecordset(rstPaperBal)
    End With
    Exit Function
ErrorHandler:
    Call CloseRecordset(rstPaperBal)
End Function
Public Function CalculateMaterialBalance(ByVal strAccountCode As String, ByVal strCategory As String, ByVal strItemCode As String, ByVal strVoucherCode As String, ByVal strVoucherType) As Double
    Dim rstMaterialBalance As New ADODB.Recordset
    Dim Category As String
    On Error GoTo ErrorHandler
    If rstMaterialBalance.State = adStateOpen Then rstMaterialBalance.Close
    Category = IIf(strCategory = "BOM", "1", IIf(strCategory = "FG", "3", IIf(strCategory = "UFG", "4", "5")))
    rstMaterialBalance.Open "SELECT FORMAT((SELECT SUM(Quantity) FROM MaterialIOChild WHERE Category='" & Category & "' AND Item=M.Code AND Godown='" & strAccountCode & "'),'0.000') AS Col0,FORMAT((SELECT SUM(Quantity) FROM MaterialSVParent,MaterialSVChild WHERE MaterialSVParent.Code=MaterialSVChild.Code AND Quantity>=0 AND Category='" & Category & "' AND Item=M.Code AND Account='" & strAccountCode & "'),'0.000') AS Col1,FORMAT((SELECT SUM(ABS(Quantity)) FROM MaterialSVParent,MaterialSVChild WHERE MaterialSVParent.Code=MaterialSVChild.Code AND Quantity<0 AND Category='" & Category & "' AND Item=M.Code AND Account='" & strAccountCode & "'),'0.000') AS Col2,FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent,MaterialMVChild WHERE MaterialMVParent.Code=MaterialMVChild.Code AND Category='" & Category & "' AND Item=M.Code AND AccountFrom='" & strAccountCode & "' AND " & IIf(strVoucherType = "MV", "MaterialMVParent.Code<>'" & strVoucherCode & "'", "1") & "),'0.000') AS Col3," & _
                            "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent,MaterialMVChild WHERE MaterialMVParent.Code=MaterialMVChild.Code AND Category='" & Category & "' AND Item=M.Code AND AccountTo='" & strAccountCode & "' AND " & IIf(strVoucherType = "MV", "MaterialMVParent.Code<>'" & strVoucherCode & "'", "1") & "),'0.000') AS Col4,FORMAT((SELECT OpBal FROM AccountChild0801 WHERE Category='" & Category & "' AND Item=M.Code AND Code='" & strAccountCode & "'),'0.000') AS Col5,FORMAT((SELECT SUM(TotalConsumption) FROM BookPOParent,BookPOChild0801 WHERE BookPOParent.Code=BookPOChild0801.Code AND Category='" & Category & "' AND Item=M.Code AND Binder='" & strAccountCode & "' AND " & IIf(strVoucherType = "PO", "BookPOParent.Code<>'" & strVoucherCode & "'", "1") & "),'0.000') AS Col6 From " & IIf(Category = "1", "OutsourceItemMaster", "BookMaster") & " M " & _
                            "WHERE Code='" & strItemCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstMaterialBalance.RecordCount > 0 Then
        CalculateMaterialBalance = Val(CheckNull(rstMaterialBalance.Fields("Col0").Value)) + Val(CheckNull(rstMaterialBalance.Fields("Col1").Value)) - Val(CheckNull(rstMaterialBalance.Fields("Col2").Value)) - Val(CheckNull(rstMaterialBalance.Fields("Col3").Value)) + Val(CheckNull(rstMaterialBalance.Fields("Col4").Value)) + Val(CheckNull(rstMaterialBalance.Fields("Col5").Value)) - Val(CheckNull(rstMaterialBalance.Fields("Col6").Value))
    Else
        CalculateMaterialBalance = 0
    End If
    Call CloseRecordset(rstMaterialBalance)
    Exit Function
ErrorHandler:
    Call CloseRecordset(rstMaterialBalance)
End Function
Public Sub FocusSelect(ByVal xTextBox As Object)
    On Error Resume Next
    If Len(xTextBox.Text) = 0 Then Exit Sub
    xTextBox.SelStart = 0
    xTextBox.SelLength = Len(xTextBox.Text)
End Sub
Public Sub ValidateKey(ByRef xTextBox As TextBox, ByRef KeyAscii As Integer, ByVal DecimalPlaces As Integer)
    Select Case KeyAscii
        Case vbKey0, vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKeyBack
        Case vbKeyDelete
            If DecimalPlaces = 0 Or InStr(xTextBox.Text, ".") <> 0 Then KeyAscii = 0
        Case vbKeyInsert
            If xTextBox.SelStart <> 0 Or InStr(xTextBox.Text, "-") <> 0 Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub
Public Function ValidateNumber(ByRef xTextBox As TextBox, ByVal DecimalPlaces As Integer) As Boolean
    ValidateNumber = False
    If IsNumeric(xTextBox.Text) Then
        If DecimalPlaces > 0 Then
            xTextBox.Text = Format(Val(xTextBox.Text), "0." + String(DecimalPlaces - 1, "#") + "0")
        Else
            xTextBox.Text = Format(Val(xTextBox.Text), "0")
        End If
        ValidateNumber = True
    Else
        FocusSelect xTextBox
    End If
End Function
Public Function ValidateDate(ByVal xMaskEdBox As Object, Optional AllowBlank As Boolean) As Boolean
    ValidateDate = True
    If AllowBlank = True And xMaskEdBox.Text = "  -  -    " Then Exit Function
    If Val(Left(xMaskEdBox.Text, 2)) > 0 And Val(Mid(xMaskEdBox.Text, 4, 2)) > 0 And Val(Mid(xMaskEdBox.Text, 4, 2)) <= 12 And Val(Right(xMaskEdBox.Text, 4)) > 0 And Len(Trim(xMaskEdBox.Text)) = 10 Then
        If IsDate(Left(xMaskEdBox.Text, 2) & "-" & MonthName(Mid(xMaskEdBox.Text, 4, 2), True) & "-" & Right(xMaskEdBox.Text, 4)) Then
            Exit Function
        End If
    End If
    ValidateDate = False: FocusSelect xMaskEdBox
End Function
Public Function GetDate(ByVal strInput As String) As String
    If strInput = "  -  -    " Then
        GetDate = "Null"
    Else
        GetDate = CStr(Left(strInput, 2)) & "-" & MonthName(Mid(strInput, 4, 2), True) & "-" & CStr(Right(strInput, 4))
    End If
End Function
Public Function FillList(ByVal lvwName As ListView, ByVal ColHdr As String, ByRef xRecordset As Recordset) As String
    Dim LITem As ListItem
    If xRecordset.RecordCount = 0 Then Exit Function
    DoEvents
    lvwName.ColumnHeaders.Add 1, , ColHdr
    lvwName.ColumnHeaders.Add 2, , ""
    xRecordset.MoveFirst
    Do While Not xRecordset.EOF
        Set LITem = lvwName.ListItems.Add(, , xRecordset.Fields(0).Value)
        LITem.ListSubItems.Add , , xRecordset.Fields(1).Value
        xRecordset.MoveNext
    Loop
    LockWindowUpdate lvwName.hwnd
    lvwName.ColumnHeaders(1).Width = lvwName.Width
    lvwName.ColumnHeaders(2).Width = 0
    LockWindowUpdate 0
End Function
Public Function SelectedItems(ByVal lvwName As ListView, Optional ByVal lvwWithCheckBox As Boolean = True) As String
    Dim i As Integer
    For i = 1 To lvwName.ListItems.Count
        If lvwWithCheckBox Then
            If lvwName.ListItems(i).Checked Then
                SelectedItems = SelectedItems + IIf(SelectedItems = "", "'", ", '") + lvwName.ListItems.Item(i).SubItems(1) + "'"
            End If
        Else
            If lvwName.ListItems(i).Selected Then
                SelectedItems = SelectedItems + IIf(SelectedItems = "", "'", ", '") + lvwName.ListItems.Item(i).SubItems(1) + "'"
            End If
        End If
    Next i
    SelectedItems = IIf(SelectedItems = "", "''", SelectedItems)
End Function
Public Function DisableCloseButton(frm As Form) As Boolean
    Dim lHndSysMenu As Long, lAns1 As Long, lAns2 As Long
    
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION) 'Remove close button
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION) 'Remove seperator bar
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0) 'Return True if both calls were successful
End Function
Public Function bVerifySum10(ByVal ISBN As String) As Boolean
    If Len(ISBN) < 13 Then bVerifySum10 = False: Exit Function
    If Len(Trim(ISBN)) < 13 Or Mid(Trim(ISBN), 12, 1) <> "-" Or InStr(1, "0123456789X", Right(Trim(ISBN), 1)) = 0 Or Len(Replace(ISBN, "-", "")) <> 10 Then bVerifySum10 = False: Exit Function
    ISBN = Replace(ISBN, "-", "")
    Dim i As Integer, K As Integer
    For K = 10 To 2 Step -1
        i = i + CInt(Val(Mid(ISBN, (10 - (K - 1)), 1))) * K
    Next
    If (i Mod 11) = 0 And Mid(ISBN, 10, 1) = "0" Then
        bVerifySum10 = True: Exit Function
    ElseIf UCase(Mid(ISBN, 10, 1)) = "X" Then
        i = i + 10
    Else
        i = i + CInt(Mid(ISBN, 10, 1))
    End If
    If Not ((i Mod 11) = 0) Then bVerifySum10 = False: Exit Function
    bVerifySum10 = True
    On Error GoTo 0
End Function
Public Function bVerifySum13(ByVal ISBN As String) As Boolean
    If Len(Trim(ISBN)) < 17 Or Mid(ISBN, 1, 3) <> "978" Or Mid(Trim(ISBN), 16, 1) <> "-" Or InStr(1, "0123456789", Right(Trim(ISBN), 1)) = 0 Or Len(Replace(ISBN, "-", "")) <> 13 Then bVerifySum13 = False: Exit Function
    ISBN = Replace(ISBN, "-", "")
    Dim i As Integer, K As Integer
    i = 30
    For K = 4 To 12 Step 2
        i = i + CInt(Mid(ISBN, K - 1, 1)) + (3 * CInt(Mid(ISBN, K, 1)))
    Next
    If Not (Mid(ISBN, 13, 1) = Trim(Str((10 - i Mod 10) Mod 10))) Then bVerifySum13 = False: Exit Function
    bVerifySum13 = True
    On Error GoTo 0
End Function
Public Sub UpdateUserAction(ByVal Activity As String, ByVal Action As String, ByVal Description As String, ByVal xConnection As ADODB.Connection)
    If UserName = "EasyPublish" Then Exit Sub
    On Error GoTo ErrorHandler
    Dim lpBuff As String * 1024
    GetComputerName lpBuff, Len(lpBuff)
    xConnection.Execute "INSERT INTO UserAction VALUES('" & UserName & "','" & Format(Now(), "dd-MMM-yyyy hh:mm:ss") & "','" & Activity & "','" & Action & "','" & Description & "','" & Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1) & "')"
    Exit Sub
ErrorHandler:
    Call DisplayError("Failed to update User Log")
End Sub
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MouseKeys As Long
    Dim Rotation As Long
    Dim xPos As Long
    Dim Ypos As Long
    If Lmsg = WM_MOUSEWHEEL Then
        MouseKeys = wParam And 65535
        Rotation = wParam / 65536
        xPos = lParam And 65535
        Ypos = lParam / 65536
        'determine if mouse wheel is being moved up or down
        If Rotation = -120 Then
            'call scroll method of datagrid and specify the number of columns and rows to scroll through DataGrid.Scroll colNum, rowNum
            MyControl.Scroll 0, 3
        Else
            MyControl.Scroll 0, -3
        End If
    End If
    WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, wParam, lParam)
End Function
Public Sub WheelHook(PassedControl As Object)
    On Error Resume Next
    Set MyControl = PassedControl
    LocalHwnd = PassedControl.hwnd
    LocalPrevWndProc = SetWindowLong(LocalHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub WheelUnHook()
    Dim WorkFlag As Long
    On Error Resume Next
    WorkFlag = SetWindowLong(LocalHwnd, GWL_WNDPROC, LocalPrevWndProc)
    Set MyControl = Nothing
End Sub
Public Function FMod(a As Variant, B As Variant) As Variant 'Floating Point Modulus
   FMod = a - Int(a / B) * B + CLng(Sgn(a) <> Sgn(B)) * B
End Function
Public Function GetChildGroup() As String
    On Error GoTo ErrHandler
    Dim DatabaseName, CurrentGroup, MCPrimary
    Dim rstEasyPublish As New ADODB.Recordset
    DatabaseName = Trim(ReadFromFile("Busy Database Name"))
    DatabaseName = StrReverse(Left(StrReverse(DatabaseName), InStr(1, StrReverse(DatabaseName), ",") - 1))
    If ConnectToBusy(DatabaseName) Then
        rstEasyPublish.Open "SELECT MCPrimary FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockReadOnly
        MCPrimary = rstEasyPublish.Fields("MCPrimary").Value
        If rstEasyPublish.State = adStateOpen Then rstEasyPublish.Close
        rstEasyPublish.Open "SELECT ''''+CONVERT(VARCHAR,Code)+'''' As Name FROM Master1 WHERE Name='" & MCPrimary & "'", cnBusy, adOpenKeyset, adLockReadOnly
        CurrentGroup = Trim(rstEasyPublish.Fields("Name").Value): GetChildGroup = CurrentGroup
        Do While True
            If rstEasyPublish.State = adStateOpen Then rstEasyPublish.Close
            rstEasyPublish.Open "SELECT STUFF((SELECT ','+(''''+CONVERT(VARCHAR,Code))+'''' FROM Master1 WHERE ParentGrp IN (" & CurrentGroup & ") FOR XML PATH('')),1,1,'') As CurrentGroup"
            If IsNull(rstEasyPublish.Fields("CurrentGroup").Value) Then Exit Do Else CurrentGroup = Trim(rstEasyPublish.Fields("CurrentGroup")): GetChildGroup = GetChildGroup + "," + CurrentGroup
        Loop
    End If
ErrHandler:
    Call CloseConnection(cnBusy): Call CloseRecordset(rstEasyPublish)
End Function
Public Function ConnectToBusy(ByVal DatabaseName As String) As Boolean
    cnBusy.CursorLocation = adUseClient
    If cnBusy.State = adStateOpen Then cnBusy.Close
    cnBusy.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=" & DatabaseName & ";Data Source=" & ServerName
    ConnectToBusy = True
End Function
Public Function RefreshData(ByRef rsGeneral As ADODB.Recordset)
    rsGeneral.ActiveConnection = cnDatabase
    Do While Not RefreshRecord(rsGeneral): Loop
    rsGeneral.ActiveConnection = Nothing
End Function
Public Function CalcUps(ByVal PaperArea As Double, ByVal PrintArea As Double) As Integer
    Dim Pc As Double
    CalcUps = Int(PaperArea / PrintArea)
    If CalcUps = 0 Then
        DisplayError ("Paper size selected is smaller than Printing Size")
    Else
        Pc = Round(((PaperArea - (PrintArea * CalcUps)) / (PrintArea * CalcUps)) * 100, 2)
        If Pc > 0 Then DisplayError ("Paper size selected is bigger than required Paper Size with " + Trim(Pc) + "%" + " wastage")
    End If
End Function
Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function
Public Function chkRef(ByVal SQL As String) As Boolean
    On Error GoTo ErrorHandler
    Dim rstRef As New ADODB.Recordset
    rstRef.Open SQL, cnDatabase, adOpenKeyset, adLockReadOnly
    If rstRef.RecordCount > 0 Then chkRef = True
    Call CloseRecordset(rstRef)
    Exit Function
ErrorHandler:
    Call CloseRecordset(rstRef): chkRef = True
End Function
Public Function U2S(ByVal Qty As Double, ByVal SPU As Integer) As Long 'Unit To Sheet
    U2S = CLng(Fix(Qty) * SPU) + (Qty - Fix(Qty)) * 1000
End Function
Public Function S2U(ByVal Qty As Long, ByVal SPU As Integer) As Double 'Sheet To Unit
    S2U = CLng(Fix(Qty / SPU)) + (Qty Mod SPU) / 1000
End Function
Public Sub Sendkeys(Text As Variant, Optional Wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(Text), Wait
        Set WshShell = Nothing
End Sub
Public Sub RetrievePic(ByVal PicData As Variant, ByVal imgFile As String, ByVal srmPicMgr As ADODB.Stream)
    With srmPicMgr
        If .State = adStateOpen Then .Close
        .Type = adTypeBinary
        .Open
        .Write PicData
        If .Size > 0 Then .SaveToFile imgFile, adSaveCreateOverWrite 'Check the size of the ado stream to make sure there is data
    End With
End Sub
Public Function FetchDiscount(ByVal Party As String, ByVal Item As String) As Double
    On Error GoTo ErrorHandler
    Dim rstFetchDiscount As New ADODB.Recordset
    With rstFetchDiscount
        .Open "SELECT [Disc%] FROM DiscountMaster D INNER JOIN BookMaster I ON D.ItemGroup=I.[Group] WHERE I.Code='" & Item & "' AND Party='" & Party & "' AND FYCode='" & FYCode & "'", cnDatabase, adOpenKeyset, adLockReadOnly
        If .RecordCount > 0 Then FetchDiscount = Val(.Fields("Disc%").Value)
    End With
    Call CloseRecordset(rstFetchDiscount)
    Exit Function
ErrorHandler:
    Call CloseRecordset(rstFetchDiscount)
End Function
Public Function UpdateComp(ByVal CompanyCode As String, ByVal WithMasters As Boolean, ByVal CreateComp As Boolean, ByVal UpdateVersion As Boolean, ByVal UpdateMajor As Boolean) As Boolean
    On Error GoTo ErrorHandler
If Trim(ReadFromFile("Client ID")) = "Publisher" Then ClientID = "P" Else ClientID = "S"
    If CompCode = "" Then Call MsgBox("Please Login Company !!!", vbInformation, App.Title): Exit Function
    UpdateComp = True
'Get Connection
If CreateComp = True Then
    cnDatabase.CursorLocation = adUseClient
    If cnDatabase.State = adStateOpen Then cnDatabase.Close
    If DatabaseType = "MS SQL" Then
    ConnectionString = "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=EP" & CompCode & ";Data Source=" & ServerName
    cnDatabase.Open ConnectionString
    ElseIf DatabaseType = "MS Access" Then
    cnDatabase.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\EasyPublish." & CompanyCode & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
    End If
    cnDatabase.BeginTrans
    If DatabaseType = "MS SQL" Then
    'BackUpDatabse
        cnDatabase.Execute "BACKUP DATABASE [EP" & CompCode & "] TO  DISK = N'C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\Backup\EP" & CompCode & "_LogBackup_temp.bak' WITH NOFORMAT, NOINIT,  NAME = N'EP" & CompCode & " -Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10"
    'RestoreDatabse
        cnDatabase.Execute "RESTORE DATABASE [EP" & CompanyCode & "] FROM  DISK = N'C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\Backup\EP" & CompCode & "_LogBackup_temp.bak' WITH  FILE = 1,  MOVE N'EPM' TO N'C:\Program Files\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL\DATA\EP" & CompanyCode & "M.mdf',  MOVE N'EPL' TO N'C:\Program Files\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL\DATA\EP" & CompanyCode & "L.ldf',  NOUNLOAD,  STATS = 5"
    End If
    cnDatabase.CommitTrans
    'CloseMainConnection
'Switch To New Company Created
    CompCode = CompanyCode
End If
    
'Get Connection Live Company
    cnDatabase.CursorLocation = adUseClient
    If cnDatabase.State = adStateOpen Then cnDatabase.Close
    If DatabaseType = "MS SQL" Then
    ConnectionString = "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=EP" & CompCode & ";Data Source=" & ServerName
    cnDatabase.Open ConnectionString
    ElseIf DatabaseType = "MS Access" Then
    cnDatabase.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\EasyPublish." & CompanyCode & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
    End If
'Begain
    cnDatabase.BeginTrans
'Create And Edit Company
    If CreateComp = True Then

    cnDatabase.Execute "DELETE FROM CompanyMaster"
    cnDatabase.Execute "INSERT INTO CompanyMaster (Code,Name,PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,Fax,eMail,Website,GSTIN,CreatedFrom,MCGroup,MCPrimary,MCRepair,FinancialYearFrom,FinancialYearTo,Printstatus,TitleCombo,BankName,AccountNo,IFSC,TallyIntegration,BusyIntegration,FYCode,Alias) VALUES ('000001','" & Trim(FrmCompanyMaster.Text1.Text) & "','" & Trim(FrmCompanyMaster.Text2.Text) & "','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text12.Text) & "'" & _
                                      ",'" & Trim(FrmCompanyMaster.Text8.Text) & "','" & Trim(FrmCompanyMaster.Text9.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & CompCode & "','0','0','0','" & Format(GetDate(FrmCompanyMaster.MhDateInput1.Text), "mm-dd-yyyy") & "','" & Format(GetDate(FrmCompanyMaster.MhDateInput2.Text), "mm-dd-yyyy") & "','N','1','" & Trim(FrmCompanyMaster.Text18.Text) & "','" & Trim(FrmCompanyMaster.Text19.Text) & "','" & Trim(FrmCompanyMaster.Text20.Text) & "','" & Trim(FrmCompanyMaster.Option1.Value) & "','" & Trim(FrmCompanyMaster.Option2.Value) & "','" & Trim(FrmCompanyMaster.Text16.Text) & "','" & Trim(FrmCompanyMaster.Text15.Text) & "')"
                                          
'Transactions 44_Tables
    cnDatabase.Execute "DELETE FROM BookDNChild"
    cnDatabase.Execute "DELETE FROM BookDNParent"
    cnDatabase.Execute "DELETE FROM BookOOChild"
    cnDatabase.Execute "DELETE FROM BookOOParent"
    cnDatabase.Execute "DELETE FROM BookPOChild05"
    cnDatabase.Execute "DELETE FROM BookPOChild0501"
    cnDatabase.Execute "DELETE FROM BookPOChild06"
    cnDatabase.Execute "DELETE FROM BookPOChild07"
    cnDatabase.Execute "DELETE FROM BookPOChild08"
    cnDatabase.Execute "DELETE FROM BookPOChild0801"
    cnDatabase.Execute "DELETE FROM BookPOChild09"
    cnDatabase.Execute "DELETE FROM BookPOChild0901"
    cnDatabase.Execute "DELETE FROM BookPOParent"
    cnDatabase.Execute "DELETE FROM BookRVChild"
    cnDatabase.Execute "DELETE FROM BookRVParent"
    cnDatabase.Execute "DELETE FROM DebitCreditParent"
    cnDatabase.Execute "DELETE FROM DebitCreditChild"
    cnDatabase.Execute "DELETE FROM DebitCreditOthInf"
    cnDatabase.Execute "DELETE FROM DebitCreditRef"
    cnDatabase.Execute "DELETE FROM JobworkBVChild"
    cnDatabase.Execute "DELETE FROM JobworkBVOthInf"
    cnDatabase.Execute "DELETE FROM JobworkBVRef"
    cnDatabase.Execute "DELETE FROM JobworkBVParent"
    cnDatabase.Execute "DELETE FROM MaterialIOChild"
    cnDatabase.Execute "DELETE FROM MaterialIOParent"
    cnDatabase.Execute "DELETE FROM MaterialMVChild"
    cnDatabase.Execute "DELETE FROM MaterialMVParent"
    cnDatabase.Execute "DELETE FROM MaterialSVChild"
    cnDatabase.Execute "DELETE FROM MaterialSVParent"
    cnDatabase.Execute "DELETE FROM OutsourceItemPOChild"
    cnDatabase.Execute "DELETE FROM OutsourceItemPOParent"
    cnDatabase.Execute "DELETE FROM PackingSlipChild"
    cnDatabase.Execute "DELETE FROM PackingSlipParent"
    cnDatabase.Execute "DELETE FROM PaperDNChild"
    cnDatabase.Execute "DELETE FROM PaperDNParent"
    cnDatabase.Execute "DELETE FROM PaperIOChild"
    cnDatabase.Execute "DELETE FROM PaperMVChild"
    cnDatabase.Execute "DELETE FROM PaperMVParent"
    cnDatabase.Execute "DELETE FROM PaperPOChild"
    cnDatabase.Execute "DELETE FROM PaperPOParent"
    cnDatabase.Execute "DELETE FROM PrintPVChild"
    cnDatabase.Execute "DELETE FROM PrintPVParent"
    cnDatabase.Execute "DELETE FROM TatRVChild"
    cnDatabase.Execute "DELETE FROM TatRVParent"

'Delete/Update Masters And Child
    cnDatabase.Execute "DELETE FROM BookChild"
    cnDatabase.Execute "DELETE FROM PaperChild"
    cnDatabase.Execute "DELETE FROM UserAction"
    cnDatabase.Execute "UPDATE AccountMaster SET Opening='0'"

'Without Masters
If Not WithMasters Then    'Delete Master
'Accounts Master
    cnDatabase.Execute "DELETE FROM AccountChild04 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
    cnDatabase.Execute "DELETE FROM AccountChild05 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
    cnDatabase.Execute "DELETE FROM AccountChild06 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
    cnDatabase.Execute "DELETE FROM AccountChild07 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
    cnDatabase.Execute "DELETE FROM AccountChild08 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
    cnDatabase.Execute "DELETE FROM AccountChild0801 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
    cnDatabase.Execute "DELETE FROM AccountMaster "
'BookingRouteMaster
    cnDatabase.Execute "DELETE FROM BookingRouteMaster "
'Book Master
    cnDatabase.Execute "DELETE FROM BookChild01 Where Left(Code,1)<>'*'"
    cnDatabase.Execute "DELETE FROM BookChild02 Where Left(Code,1)<>'*'"
    cnDatabase.Execute "DELETE FROM BookChild03 Where Left(Code,1)<>'*'"
    cnDatabase.Execute "DELETE FROM BookChild05 Where Left(Code,1)<>'*'"
    cnDatabase.Execute "DELETE FROM BookChild06 Where Left(Code,1)<>'*'"
    cnDatabase.Execute "DELETE FROM BookChild07 Where Left(Code,1)<>'*'"
    'cnDatabase.Execute "DELETE FROM BookChild08 Where Left(Code,1)<>'*'"
    cnDatabase.Execute "DELETE FROM BookMaster "
'Other Masters
    cnDatabase.Execute "DELETE FROM DiscountMaster "
    cnDatabase.Execute "DELETE FROM ElementMaster "
    If MsgBox("Do You Wants to Delete 'Finish Size Masters' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
    cnDatabase.Execute "DELETE FROM FinishSizeChild "
    End If
    cnDatabase.Execute "DELETE GeneralMaster "
    cnDatabase.Execute "DELETE FROM OutsourceItemMaster "
    If MsgBox("Do You Wants to Delete 'Paper Master' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
    cnDatabase.Execute "DELETE FROM PaperMaster "
    End If
    If MsgBox("Do You Wants to Delete 'Size Group Masters' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
    cnDatabase.Execute "DELETE FROM SizeGroupChild "
    End If
    cnDatabase.Execute "DELETE FROM TaxMaster "
    cnDatabase.Execute "DELETE FROM TeamMemberMaster "
    cnDatabase.Execute "DELETE FROM VchSeriesMaster "
    cnDatabase.Execute "DELETE FROM UserChild Where Code NOT IN (Select Code from UserMaster Where Level<>1)"
    cnDatabase.Execute "DELETE FROM UserMaster Where Level<>1"
    cnDatabase.Execute "DELETE FROM VchSeriesMaster Where Left(Code,1)='*'"
'With Master
    Else
        cnDatabase.Execute "UPDATE AccountMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N',Opening='0'"
        cnDatabase.Execute "UPDATE BookMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
        cnDatabase.Execute "UPDATE PaperMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
        cnDatabase.Execute "UPDATE OutsourceItemMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
        cnDatabase.Execute "UPDATE TaxMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
        cnDatabase.Execute "UPDATE TeamMemberMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
        cnDatabase.Execute "UPDATE GeneralMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
    End If
End If

If CreateComp = True Then
''Account Masters
    cnDatabase.Execute "DELETE FROM AccountMaster Where Code ='000000' Or Left(Code,1)='*'"
    cnDatabase.Execute "Insert Into AccountMaster VALUES ('000000','" & Trim(FrmCompanyMaster.Text1.Text) & "','" & Trim(FrmCompanyMaster.Text2.Text) & "','000000','*12002','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00001','Rate Master','Rate Master','1002','*12002','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00002','Main Godown','Main Godown','1003','*99999','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00003','Self Transport','Self Transport','1004','*99996','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00004','Self Packer','Self Packer','1005','*99997','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00005','Direct','Direct','1006','*99998','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
End If
'Check Version Update
    cnDatabase.Execute "IF EXISTS(SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'EasyPublishVersion' AND COLUMN_NAME = 'Date') Print 'Col_Exist' ELSE CREATE TABLE dbo.EasyPublishVersion(Date datetime NOT NULL, Major nchar(10) NOT NULL,Minor nchar(10) NOT NULL,Revision nchar(10) NOT NULL,Version nchar(10) NULL) ON [PRIMARY] ALTER TABLE dbo.EasyPublishVersion SET (LOCK_ESCALATION = TABLE)"
    cnDatabase.Execute "IF EXISTS(SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'EasyPublishVersion' AND COLUMN_NAME = 'vUPDATE') Print 'Col_Exist' ELSE Alter Table EasyPublishVersion Add vUPDATE nvarchar(100) NULL"
    If rstEasyPublishVersion.State = adStateOpen Then rstEasyPublishVersion.Close
    rstEasyPublishVersion.Open "Select * From EasyPublishVersion ", cnDatabase, adOpenKeyset, adLockReadOnly
    
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "Select * From CompanyMaster  where FYCode= '" & FYCode & "' ", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstCompanyMaster.RecordCount = 0 Then GoTo NXT
'Version Updates
'***************************************************************************************************************************************************************
    UpdateVersion = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.09.22" Then UpdateVersion = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMinor01' Where vUPDATE IS NULL AND Version='21.09.22'"
    rstEasyPublishVersion.MoveNext
    Loop
'Alias
    Dim aFlag As Boolean
    Dim j As Long, K As Long
        j = 0: K = 0: aFlag = True: CompAlias = ""
        compName = rstCompanyMaster.Fields("PrintName")
        K = Len(compName)
    For j = 1 To K
        If Mid(compName, j, 1) <> " " And aFlag = True Then
            CompAlias = CompAlias + Mid(compName, j, 1)
            aFlag = False
        ElseIf Mid(compName, j, 1) = " " Then
            aFlag = True
        End If
    Next j
'Update Alias CompanyMaster
    'cnDatabase.Execute "Update CompanyMaster Set FYCode= " & FYCode & ""
    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'Alias') IS NOT NULL PRINT 'Exists' ELSE Alter Table CompanyMaster Add Alias nvarchar(6) NOT NULL CONSTRAINT df_Alias DEFAULT '' "
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "Select * From CompanyMaster  where FYCode= '" & FYCode & "' ", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstCompanyMaster.Fields("Alias") = "" Or rstCompanyMaster.Fields("Alias") = Null Then CompAlias = CompAlias Else CompAlias = rstCompanyMaster.Fields("Alias")
    cnDatabase.Execute "IF (SELECT Alias FROM CompanyMaster WHERE FYCode= " & FYCode & " ) <>'' PRINT 'Exists' ELSE Update CompanyMaster Set Alias= '" & CompAlias & "' Where FYCode= " & FYCode & " AND Alias='' Or Alias IS Null"
    
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "Select * From CompanyMaster  where FYCode= '" & FYCode & "' ", cnDatabase, adOpenKeyset, adLockReadOnly
    If rstCompanyMaster.RecordCount = 0 Then GoTo NXT
    CompAlias = rstCompanyMaster.Fields("Alias")

'Update=01
If UpdateVersion = True Then
Call UpdateMinor01
'EasyPublishVersion
cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.09.22') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,9,22,'21.09.22')"
frmLicenceAgreement.Label2 = ""
End If
'***************************************************************************************************************************************************************
'Update=02
UpdateVersion = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.09.23" Then UpdateVersion = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMinor02' Where vUPDATE IS NULL AND Version='21.09.23'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateVersion = True Then
Call UpdateMinor02
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.09.23') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,09,23,'21.09.23')"
End If
'***************************************************************************************************************************************************************
'Update=03
UpdateVersion = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.10.01" Then UpdateVersion = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMinor03' Where vUPDATE IS NULL AND Version='21.10.01'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateVersion = True Then
Call UpdateMinor03
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.10.01') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,10,01,'21.10.01')"
End If
'***************************************************************************************************************************************************************
'Update=04
UpdateVersion = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.10.13" Then UpdateVersion = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMinor04' Where vUPDATE IS NULL AND Version='21.10.13'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateVersion = True Then
Call UpdateMinor04
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.10.13') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,10,13,'21.10.13')"
End If
'***************************************************************************************************************************************************************
'Update=05
UpdateVersion = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.10.19" Then UpdateVersion = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMinor05' Where vUPDATE IS NULL AND Version='21.10.19'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateVersion = True Then
Call UpdateMinor05
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.10.19') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,10,19,'21.10.19')"
End If
'***************************************************************************************************************************************************************
'Update=06
UpdateVersion = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.10.20" Then UpdateVersion = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMinor06' Where vUPDATE IS NULL AND Version='21.10.20'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateVersion = True Then
Call UpdateMinor06
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.10.20') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,10,20,'21.10.20','UpdateMinor06')"
End If
'***************************************************************************************************************************************************************
'Update=07
UpdateVersion = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.10.21" Then UpdateVersion = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMinor07' Where vUPDATE IS NULL AND Version='21.10.21'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateVersion = True Then
Call UpdateMinor07
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.10.21') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,10,21,'21.10.21','UpdateMinor07')"
End If
'***************************************************************************************************************************************************************
'Update=08
UpdateVersion = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.10.22" Then UpdateVersion = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMinor08' Where vUPDATE IS NULL AND Version='21.10.22'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateVersion = True Then
Call UpdateMinor08
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.10.22') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,10,22,'21.10.22','UpdateMinor08')"
End If
'***************************************************************************************************************************************************************
'Update=09
UpdateVersion = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.10.23" Then UpdateVersion = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMinor09' Where vUPDATE IS NULL AND Version='21.10.23'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateVersion = True Then
Call UpdateMinor09
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.10.23') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,10,23,'21.10.23','UpdateMinor09')"
End If
'***************************************************************************************************************************************************************
'Update=10
UpdateVersion = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.10.24" Then UpdateVersion = False
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateVersion = True Then
Call UpdateMinor10
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.10.24') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,10,24,'21.10.24','UpdateMinor10[MachineMaster,BookPOChild0501]')"
End If
'***************************************************************************************************************************************************************



'***************************************************************************************************************************************************************
'MajorUpdate=01
If MsgBox("Do You Wants to Update '21.11.01 Version'" & vbCrLf & "[Update Account Child Master & Update Color Master] Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
UpdateMajor = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.11.01" Then UpdateMajor = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMajor01[Update Account Child Master & Update Color Master]' Where vUPDATE IS NULL AND Version='21.11.01'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateMajor = True And MajorFlag = True Then
Call UpdateMajor01 'Update Account Child Master ,Update Color Master
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.11.01') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,11,01,'21.11.01','UpdateMajor01[Update Account Child Master & Update Color Master]')"
End If
End If
'***************************************************************************************************************************************************************
'MajorUpdate=02
If MsgBox("Do You Wants to Update '21.11.02 Version'" & vbCrLf & "[Update BookMaster ] Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
UpdateMajor = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.11.02" Then UpdateMajor = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMajor02[Update BookMaster ]' Where vUPDATE IS NULL AND Version='21.11.02'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateMajor = True And MajorFlag = True Then
Call UpdateMajor02
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.11.02') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,11,02,'21.11.02','UpdateMajor02[Update BookMaster ]')"
End If
End If
'***************************************************************************************************************************************************************
'MajorUpdate=03
If MsgBox("Do You Wants to Update '21.11.03 Version'" & vbCrLf & "[Update BookChild05 & BookPOChild05] Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
UpdateMajor = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.11.03" Then UpdateMajor = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMajor03[Update BookChild05 & BookPOChild05]' Where vUPDATE IS NULL AND Version='21.11.03'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateMajor = True And MajorFlag = True Then
Call UpdateMajor03
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.11.03') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,11,03,'21.11.03','UpdateMajor03[Update BookChild05 & BookPOChild05]')"
End If
End If
'***************************************************************************************************************************************************************
'MajorUpdate=04
If MsgBox("Do You Wants to Update '21.11.04 Version'" & vbCrLf & "[Update BookChild06 & BookPOChild06] Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
UpdateMajor = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.11.04" Then UpdateMajor = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMajor04[Update BookChild06 & BookPOChild06]' Where vUPDATE IS NULL AND Version='21.11.04'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateMajor = True And MajorFlag = True Then
Call UpdateMajor04
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.11.04') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,11,04,'21.11.04','UpdateMajor04[Update BookChild06 & BookPOChild06]')"
End If
End If
'***************************************************************************************************************************************************************
'MajorUpdate=05 Update AccountChild07
If MsgBox("Do You Wants to Update '21.11.05 Version'" & vbCrLf & "[Update AccountChild07] Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
UpdateMajor = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.11.05" Then UpdateMajor = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMajor05[Update AccountChild07]' Where vUPDATE IS NULL AND Version='21.11.05'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateMajor = True And MajorFlag = True Then
Call UpdateMajor05
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.11.05') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,11,05,'21.11.05','UpdateMajor05[Update AccountChild07]')"
End If
End If
'***************************************************************************************************************************************************************
'MajorUpdate=06 Update BookChild08
If MsgBox("Do You Wants to Update '21.11.06 Version'" & vbCrLf & "[Update BookChild08] Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
UpdateMajor = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.11.06" Then UpdateMajor = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMajor06[Update BookChild08] ' Where vUPDATE IS NULL AND Version='21.11.06'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateMajor = True And MajorFlag = True Then
Call UpdateMajor06
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.11.06') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,11,06,'21.11.06','UpdateMajor06[Update BookChild08] ')"
End If
End If
'***************************************************************************************************************************************************************
'MajorUpdate=07  Update BookPOChild08
If MsgBox("Do You Wants to Update '21.11.07 Version'" & vbCrLf & "[Update BookPOChild08] Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
UpdateMajor = True
    If rstEasyPublishVersion.RecordCount <> 0 Then rstEasyPublishVersion.MoveFirst
    Do While Not rstEasyPublishVersion.EOF
    If Trim(rstEasyPublishVersion.Fields("Version").Value) = "21.11.07" Then UpdateMajor = False: cnDatabase.Execute "Update EasyPublishVersion Set vUPDATE='UpdateMajor07' Where vUPDATE IS NULL AND Version='21.11.07'"
    rstEasyPublishVersion.MoveNext
    Loop
If UpdateMajor = True And MajorFlag = True Then
Call UpdateMajor07
'EasyPublishVersion
    cnDatabase.Execute "IF EXISTS (SELECT *FROM EasyPublishVersion WHERE Version='21.11.07') Print 'Version_Exist' ELSE Insert Into EasyPublishVersion VALUES (GetDate(),21,11,07,'21.11.07','UpdateMajor07')"
End If
End If
'***************************************************************************************************************************************************************
'Common Update
    'Email
    cnDatabase.Execute "IF EXISTS(SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'Email' AND COLUMN_NAME = 'Code') Print 'Col_Exist' ELSE CREATE TABLE dbo.Email(Code nchar(6) NOT NULL,Company nchar(100) NULL,ContactPerson nchar(60) NULL,Mobile nchar(60) NULL,email nchar(80) NULL,Address nchar(150) NULL,PIN nchar(10) NULL,CITY nchar(40) NULL,Category nchar(30) NULL,State nchar(30) NULL,Status nchar(10) NULL) ON [PRIMARY] ALTER TABLE dbo.Email SET (LOCK_ESCALATION = TABLE)"
    'General Master    *None
    cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11098' OR Name='(*None)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11098','(*None)','*None','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
    'Machine Master
    cnDatabase.Execute "IF EXISTS(SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'MachineMaster' AND COLUMN_NAME = 'Code') Print 'Col_Exist' ELSE CREATE TABLE dbo.MachineMaster(Code char(6) NOT NULL,Name nvarchar(40) NOT NULL,PrintName nvarchar(40) NOT NULL,Units tinyint NOT NULL,MakeReadyTime tinyint NOT NULL,Efficiency smallint NOT NULL,MinSizeWidth tinyint NOT NULL,MinSizeLength tinyint NOT NULL,MaxSizeWidth tinyint NOT NULL,MaxSizeLength tinyint NOT NULL,StartTime time(0) NOT NULL,EndTime time(0) NOT NULL,Category tinyint NOT NULL,CreatedBy char(6) NOT NULL,CreatedOn datetime NOT NULL,ModifiedBy char(6) NULL,ModifiedOn datetime NULL,RecordStatus char(1) NOT NULL,PrintStatus char(1) NOT NULL CONSTRAINT PK_MachineMaster PRIMARY KEY CLUSTERED (Code))  ON [PRIMARY]"
    'Machine Master PK_MachineMaster
    cnDatabase.Execute "IF Not EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE TABLE_NAME = 'MachineMaster' AND CONSTRAINT_NAME='PK_MachineMaster') ALTER TABLE MachineMaster ADD CONSTRAINT PK_MachineMaster PRIMARY KEY CLUSTERED (Code) ELSE Print 'CONSTRAINT_Exist'"
    'Machine Child
    cnDatabase.Execute "IF EXISTS(SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'MachineChild' AND COLUMN_NAME = 'Code') Print 'Col_Exist' ELSE CREATE TABLE dbo.MachineChild(Code char(6) NOT NULL,Qty smallint NOT NULL,Sets decimal(4, 2) NOT NULL,Hours tinyint NOT NULL,Efficiency smallint NOT NULL CONSTRAINT [FK_MachineChild] FOREIGN KEY([Code]) REFERENCES [MachineMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE) ON [PRIMARY]"
    'MachineChild FK_MachineChild
    cnDatabase.Execute "IF Not EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE TABLE_NAME = 'MachineChild' AND CONSTRAINT_NAME='FK_MachineChild') ALTER TABLE MachineChild ADD CONSTRAINT [FK_MachineChild] FOREIGN KEY([Code]) REFERENCES [MachineMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE ELSE Print 'CONSTRAINT_Exist'"
    'MachineChild FK_MachineChild
    cnDatabase.Execute "ALTER TABLE dbo.MachineChild Alter Column Sets decimal(4, 2) NOT NULL"
    'Element Masters
    cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00051' OR NAME='Mono Carton') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00051','Mono Carton','Mono Carton','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
    cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00052' OR NAME='Carton_Sheet-1') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00052','Carton_Sheet-1','Carton_Sheet-1','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
    cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00053' OR NAME='Carton_Sheet-2') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00053','Carton_Sheet-2','Carton_Sheet-2','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
    'BookPOChild06
    cnDatabase.Execute "IF Exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'BookPOChild06' AND COLUMN_NAME = 'Titles/sheet1') EXEC sp_rename 'BookPOChild06.Titles/sheet1','Ups','Column' Else Print 'Col_NOT_Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'BookPOChild06' AND COLUMN_NAME='Ups') ALTER TABLE BookPOChild06 Alter COLUMN [Ups] decimal(12, 0) ELSE Print 'Column_Not_Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'BookPOChild06' AND COLUMN_NAME='Sets') ALTER TABLE BookPOChild06 Alter COLUMN [Sets] decimal(12, 0) ELSE Print 'Column_Not_Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'BookPOChild06' AND COLUMN_NAME='Titles/sheet2') ALTER TABLE BookPOChild06 Alter COLUMN [Titles/sheet2] decimal(12, 0) ELSE Print 'Column_Not_Exist'"
    'Company SMTP Update
    cnDatabase.Execute "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'CompanyMaster' AND COLUMN_NAME='SmtpServer') Print 'Column_Exist' ELSE ALTER TABLE dbo.CompanyMaster ADD   SmtpServer nvarchar(60) NULL,Port nvarchar(60) NULL,UserName nvarchar(60) NULL,Password nvarchar(60) NULL"
    cnDatabase.Execute "IF EXISTS (SELECT SmtpServer FROM CompanyMaster) Update CompanyMaster Set SmtpServer='smtp.gmail.com' Where SmtpServer IS NULL  ELSE Print 'Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT Port FROM CompanyMaster) Update CompanyMaster Set Port='465' Where Port IS NULL  ELSE Print 'Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT UserName FROM CompanyMaster) Update CompanyMaster Set UserName='production.easyinfosolutionsi@gmail.com' Where UserName IS NULL  ELSE Print 'Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT Password FROM CompanyMaster) Update CompanyMaster Set Password='mr74eena' Where Password IS NULL  ELSE Print 'Exist'"
    'TeamMemberMaster Update
    cnDatabase.Execute "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TeamMemberMaster' AND COLUMN_NAME='email') Print 'Column_Exist' Else CREATE TABLE dbo.Tmp_TeamMemberMaster(Code nvarchar(6) NOT NULL,Name nvarchar(40) NOT NULL,PrintName nvarchar(40) NOT NULL,Department nvarchar(6) NOT NULL,Designation nvarchar(6) NOT NULL,LoginId nvarchar(6) NOT NULL,ReportingTo nvarchar(6) NULL,eMail nvarchar(50) NULL,CreatedBy nvarchar(6) NOT NULL,CreatedOn datetime NOT NULL,ModifiedBy nvarchar(6) NULL,ModifiedOn datetime NULL,Recordstatus nvarchar(1) NOT NULL,Printstatus nvarchar(1) NOT NULL)  ON [PRIMARY]"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TeamMemberMaster' AND COLUMN_NAME='email') Print 'Column_Exist' Else ALTER TABLE dbo.Tmp_TeamMemberMaster SET (LOCK_ESCALATION = TABLE)"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TeamMemberMaster' AND COLUMN_NAME='email') Print 'Column_Exist' Else  IF EXISTS(SELECT * FROM dbo.TeamMemberMaster) EXEC('INSERT INTO dbo.Tmp_TeamMemberMaster (Code, Name, PrintName, Department, Designation, LoginId, ReportingTo, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn, Recordstatus, Printstatus) SELECT Code, Name, PrintName, Department, Designation, LoginId, ReportingTo, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn, Recordstatus, Printstatus FROM dbo.TeamMemberMaster WITH (HOLDLOCK TABLOCKX)')"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TeamMemberMaster' AND COLUMN_NAME='email') Print 'Column_Exist' Else DROP TABLE dbo.TeamMemberMaster"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TeamMemberMaster' AND COLUMN_NAME='email') Print 'Column_Exist' Else EXECUTE sp_rename N'dbo.Tmp_TeamMemberMaster', N'TeamMemberMaster', 'OBJECT' "
    '***************************************************************************************************************************************************************
NXT:
    cnDatabase.CommitTrans
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstEasyPublishVersion)
    Screen.MousePointer = vbNormal
    Exit Function
ErrorHandler:
    UpdateComp = False
    cnDatabase.RollbackTrans
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstEasyPublishVersion)
    Screen.MousePointer = vbNormal
End Function
Public Function UpdateMajor01() 'AccountChild Update 04,05,06,07
'Create Table
'   cnDatabase.Execute "Delete TmpData Where Right(Code,1)<>'a'"
   cnDatabase.Execute "Drop Table TmpData"
   cnDatabase.Execute "IF Exists (Select *From INFORMATION_SCHEMA.COLUMNS Where TABLE_NAME = 'TmpData' AND COLUMN_NAME = 'CODE') Print 'Col_Exist' ELSE CREATE TABLE [TmpData](  [Code] [nvarchar](7) NULL,[Name] [nvarchar](100) NULL ) ON [PRIMARY] "
'Inser Tmp_Data
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10001','Extra Large-28''''X40''''-A/P')"
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10002','Extra Large-28''''X40''''-A/P_SPL')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10002a','Extra Large(28''''X40'''')A/P_SPL')"
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10003','Extra Large-30''''X40''''')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10003a','Extra Large-(30''''X40'''')')"
   
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10004','Extra Large-30''''X40''''-(A/P)')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10005','Extra Large-30''''X40''''-(Card)')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10006','LARGE-23''''X36''''')"
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10007','LARGE-23''''X36''''-(A/P)')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10007a','LARGE-(23''''X36'''')A/P')"
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10008','LARGE-23''''X36''''-(Card)')"
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10009','Medium-20''''X30''''')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10009a','Medium-(20''''X30'''')')"
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10010','Medium-20''''X30''''(A/P)')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10010a','Medium-(20''''X30'''')A/P')"
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10011','Medium-20''''X30''''(Card)')"
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10012','Small-19''''X26''''')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10012a','Small-(19''''X26'''')')"
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10013','Small-19''''X26''''(Card)')"
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10014','Web-508mm')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10015','Web-578mm')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10016','Extra Large-28''''X40''''')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10017','Small-19''''X26''''-(A/P)')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10018','Extra Large-28''''X40''''-(Card)')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10019','Little-11.50''''X18.00''''')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10020','Little-11.50''''X18.00''''-(Card)')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10021','Little-11.50''''X18.00''''-(A/P)')"
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10022','12.00X18.00-Digital')"
   
   cnDatabase.Execute "INSERT INTO TmpData VALUES ('*10023a','Extra Large(28''''X40'''')UC_A/P_SPL')"
'Update Data
   cnDatabase.Execute "Update GeneralMaster Set Name= '28.00X40.00-Extra Large-(A/P)',PrintName='28.00X40.00-Extra Large-(A/P)' Where NAME=(Select NAME From TmpData Where Code='*10001')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '28.00X40.00-Extra Large-(A/P_SPL)',PrintName='28.00X40.00-Extra Large-(A/P_SPL)' Where NAME=(Select NAME From TmpData Where Code='*10002')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '28.00X40.00-Extra Large-(A/P_SPL)',PrintName='28.00X40.00-Extra Large-(A/P_SPL)' Where NAME=(Select NAME From TmpData Where Code='*10002a')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '30.00X40.00-Extra Large',PrintName='30.00X40.00-Extra Large' Where NAME=(Select NAME From TmpData Where Code='*10003')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '30.00X40.00-Extra Large',PrintName='30.00X40.00-Extra Large' Where NAME=(Select NAME From TmpData Where Code='*10003a')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '30.00X40.00-Extra Large-(A/P)',PrintName='30.00X40.00-Extra Large-(A/P)' Where NAME=(Select NAME From TmpData Where Code='*10004')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '30.00X40.00-Extra Large-(Card)',PrintName='30.00X40.00-Extra Large-(Card)' Where NAME=(Select NAME From TmpData Where Code='*10005')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '23.00X36.00-LARGE',PrintName='23.00X36.00-LARGE' Where NAME=(Select NAME From TmpData Where Code='*10006')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '23.00X36.00-LARGE-(A/P)',PrintName='23.00X36.00-LARGE-(A/P)' Where NAME=(Select NAME From TmpData Where Code='*10007')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '23.00X36.00-LARGE-(A/P)',PrintName='23.00X36.00-LARGE-(A/P)' Where NAME=(Select NAME From TmpData Where Code='*10007a')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '23.00X36.00-LARGE-(Card)',PrintName='23.00X36.00-LARGE-(Card)' Where NAME=(Select NAME From TmpData Where Code='*10008')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X30.00-Medium',PrintName='20.00X30.00-Medium' Where NAME=(Select NAME From TmpData Where Code='*10009')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X30.00-Medium',PrintName='20.00X30.00-Medium' Where NAME=(Select NAME From TmpData Where Code='*10009a')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X30.00-Medium-(A/P)',PrintName='20.00X30.00-Medium-(A/P)' Where NAME=(Select NAME From TmpData Where Code='*10010')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X30.00-Medium-(A/P)',PrintName='20.00X30.00-Medium-(A/P)' Where NAME=(Select NAME From TmpData Where Code='*10010a')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X30.00-Medium-(Card)',PrintName='20.00X30.00-Medium-(Card)' Where NAME=(Select NAME From TmpData Where Code='*10011')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '19.00X26.00-Small',PrintName='19.00X26.00-Small' Where NAME=(Select NAME From TmpData Where Code='*10012')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '19.00X26.00-Small',PrintName='19.00X26.00-Small' Where NAME=(Select NAME From TmpData Where Code='*10012a')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '19.00X26.00-Small-(Card)',PrintName='19.00X26.00-Small-(Card)' Where NAME=(Select NAME From TmpData Where Code='*10013')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X30.00-Web-508mm',PrintName='20.00X30.00-Web-508mm' Where NAME=(Select NAME From TmpData Where Code='*10014')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '22.80X36.00-Web-578mm',PrintName='22.80X36.00-Web-578mm' Where NAME=(Select NAME From TmpData Where Code='*10015')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '28.00X40.00-Extra Large',PrintName='28.00X40.00-Extra Large' Where NAME=(Select NAME From TmpData Where Code='*10016')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '19.00X26.00-Small-(A/P)',PrintName='19.00X26.00-Small-(A/P)' Where NAME=(Select NAME From TmpData Where Code='*10017')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '28.00X40.00-Extra Large-(Card)',PrintName='28.00X40.00-Extra Large-(Card)' Where NAME=(Select NAME From TmpData Where Code='*10018')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '11.50.00X18.00.00-Little',PrintName='11.50.00X18.00.00-Little' Where NAME=(Select NAME From TmpData Where Code='*10019')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '11.50.00X18.00.00-Little-(Card)',PrintName='11.50.00X18.00.00-Little-(Card)' Where NAME=(Select NAME From TmpData Where Code='*10020')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '11.50.00X18.00.00-Little-(A/P)',PrintName='11.50.00X18.00.00-Little-(A/P)' Where NAME=(Select NAME From TmpData Where Code='*10021')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '12.00X18.00-Digital',PrintName='12.00X18.00-Digital' Where NAME=(Select NAME From TmpData Where Code='*10022')"
   cnDatabase.Execute "Update GeneralMaster Set Name= '28.00X40.00-Extra Large_UC_ A/P_SPL',PrintName='28.00X40.00-Extra Large_UC_ A/P_SPL' Where NAME=(Select NAME From TmpData Where Code='*10023a')"
    
   cnDatabase.Execute "Update GeneralMaster Set Name= '23.00X36.00-Mini Offset',PrintName='23.00X36.00-Mini Offset' Where NAME='Mini Offset (23X36)'"
   cnDatabase.Execute "Update GeneralMaster Set Name= '23.00X36.00-Screen Printing',PrintName='23.00X36.00-Screen Printing' Where NAME='Screen Printing'"
   cnDatabase.Execute "Update GeneralMaster Set Name= '60.00X80.00-Big Format',PrintName='60.00X80.00-Big Format' Where NAME='Big Format'"
   cnDatabase.Execute "Update GeneralMaster Set Name= '25.00X38.00(64X96.5)',PrintName='25.00X38.00(64X96.5)' Where NAME='25.00X38(64x96.5)'"


    cnDatabase.Execute "Update GeneralMaster Set Name= '28.00X40.00(A/P SPL)',PrintName='28.00X40.00(A/P SPL)' Where NAME='28.00X40(A/P SPL)'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '28.00X40.00(UC SPL)',PrintName='28.00X40.00(UC SPL)' Where NAME='28.00X40(UC SPL)'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '02.00X03.00-Visiting Card Size',PrintName='02.00X03.00-Visiting Card Size' Where NAME='Visiting Card'"

    cnDatabase.Execute "Update GeneralMaster Set Name= '11.50X18.00-Little',PrintName='11.50X18.00-Little' Where NAME='11.50.00X18.00.00-Little'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '11.50X18.00-Little-(A/P)',PrintName='11.50X18.00-Little-(A/P)' Where NAME='11.50.00X18.00.00-Little-(A/P)'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '11.50X18.00-Little-(Card)',PrintName='11.50X18.00-Little-(Card)' Where NAME='11.50.00X18.00.00-Little-(Card)'"
        
        
    cnDatabase.Execute "Update GeneralMaster Set Name= '23.00X36.00-Mini Offset_Size',PrintName='23.00X36.00-Mini Offset_Size' Where NAME='Mini Offset-23x36'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '30.00X40.00-Web_104.00 CM-Reel',PrintName='30.00X40.00-Web_104.00 CM-Reel' Where NAME='Web_104cm-Reel'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '32.00X44.00-Web_113.00 CM-Reel',PrintName='32.00X44.00-Web_113.00 CM-Reel' Where NAME='Web_113 CM Reel'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '26.00X40.00-Web_66.00 CM-Reel',PrintName='26.00X40.00-Web_66.00 CM-Reel' Where NAME='Web_66cm-Reel'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X29.52-Web_75.00 CM-Reel',PrintName='20.00X29.52-Web_75.00 CM-Reel' Where NAME='Web_75cm.-Reel'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X30.00-Web_76.20 CM-Reel',PrintName='20.00X30.00-Web_76.20 CM-Reel' Where NAME='Web_76.2 CM Reel'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X29.92-Web_76.00 CM-Reel',PrintName='20.00X29.92-Web_76.00 CM-Reel' Where NAME='Web_76cm-Reel'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '22.00X34.00-Web_86.50 CM-Reel',PrintName='22.00X34.00-Web_86.50 CM-Reel' Where NAME='Web_86.5 CM Reel'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '22.00X34.44-Web_87.50 CM-Reel',PrintName='22.00X34.44-Web_87.50 CM-Reel' Where NAME='Web_87.5 CM-REEL'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '22.00X35.00-Web_89.00 CM-Reel ',PrintName='22.00X35.00-Web_89.00 CM-Reel' Where NAME='Web_89cm-Reel'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '22.00X35.43-Web_90.00 CM-Reel',PrintName='22.00X35.43-Web_90.00 CM-Reel' Where NAME='Web_90cm-Reel'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '22.00X35.82-Web_91.00 CM-Reel',PrintName='22.00X35.82-Web_91.00 CM-Reel' Where NAME='Web_91cm-Reel'"
    


'AccountChild04
    cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'AccountChild04' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'AccountChild04', 'AccountChild04T'"
    cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'AccountChild04T' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE CREATE TABLE [AccountChild04]([Code] [nvarchar](6) NOT NULL,[NegativeOnePcRate] [decimal](12, 2) NOT NULL,[NegativeCutPcRate] [decimal](12, 2) NOT NULL,[NegativePastingRate] [decimal](12, 2) NOT NULL,[PositiveOnePcRate] [decimal](12, 2) NOT NULL,[PositiveCutPcRate] [decimal](12, 2) NOT NULL,[PositivePastingRate] [decimal](12, 2) NOT NULL,[WEF] [date] NOT NULL DEFAULT '01-APR-2021',[Type] [char](1) NOT NULL DEFAULT 'S' CONSTRAINT [FK_AccountChild04_AccountMaster_I]  FOREIGN KEY([Code]) REFERENCES [AccountMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE) ON [PRIMARY]"
    cnDatabase.Execute "INSERT INTO AccountChild04 SELECT [Code],[NegativeOnePcRate],[NegativeCutPcRate],[NegativePastingRate],[PositiveOnePcRate],[PositiveCutPcRate],[PositivePastingRate],'01-APR-2021','S' FROM [AccountChild04T]"
    cnDatabase.Execute "DROP TABLE AccountChild04T"

'AccountChild05
    cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'AccountChild05' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'AccountChild05', 'AccountChild05T'"
    cnDatabase.Execute "CREATE TABLE [AccountChild05]([Code] [nvarchar](6) NOT NULL,[SizeGroup] [nvarchar](6) NOT NULL,[Range] [decimal](6, 0) NOT NULL,[PrintingRate] [decimal](12, 2) NOT NULL,[PaperWastageRate] [decimal](5, 2) NOT NULL,[PaperWastageMin] [decimal](6, 0) NOT NULL,[PaperWastageMax] [decimal](6, 0) NOT NULL,[Color] [nvarchar](6) NOT NULL,[WEF] [date] NOT NULL DEFAULT '01-APR-2021',[Type] [char](1) NOT NULL DEFAULT 'S' CONSTRAINT [FK_AccountChild05_AccountMaster_I] FOREIGN KEY([Code]) REFERENCES [AccountMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE, CONSTRAINT [FK_AccountChild05_GeneralMaster_II] FOREIGN KEY([SizeGroup]) REFERENCES [GeneralMaster] ([Code]), CONSTRAINT [FK_AccountChild05_GeneralMaster_III] FOREIGN KEY([Color]) REFERENCES [GeneralMaster] ([Code])) ON [PRIMARY] "
'Color Master Update
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*23000' OR Name='None Color') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*23000','None Color','None Color','23','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*23001' OR Name='01-CMYK') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*23001','01-CMYK','01-CMYK','23','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*23002' OR Name='02-CMYK') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*23002','02-CMYK','02-CMYK','23','2','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*23003' OR Name='04-CMYK') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*23003','04-CMYK','04-CMYK','23','4','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*23004' OR Name='06-CMYK') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*23004','06-CMYK','06-CMYK','23','6','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*23005' OR Name='05-CMYK') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*23005','05-CMYK','05-CMYK','23','5','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*23006' OR Name='03-CMYK') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*23006','03-CMYK','03-CMYK','23','3','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*23007' OR Name='07-CMYK') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*23007','07-CMYK','07-CMYK','23','7','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*23008' OR Name='08-CMYK') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*23008','08-CMYK','08-CMYK','23','8','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
    
    cnDatabase.Execute "INSERT INTO AccountChild05 SELECT [Code],[Size],[Range1],[PrintRate1],[PaperWastageRate1],[PaperWastageMin1],999999,'*23001','01-APR-2021','S' FROM [AccountChild05T] WHERE [Range1]>0 OR [PrintRate1]>0 OR [PaperWastageRate1]>0 OR [PaperWastageMin1]>0"
    cnDatabase.Execute "INSERT INTO AccountChild05 SELECT [Code],[Size],[Range2],[PrintRate2],[PaperWastageRate2],[PaperWastageMin2],999999,'*23002','01-APR-2021','S' FROM [AccountChild05T] WHERE [Range1]>0 OR [PrintRate2]>0 OR [PaperWastageRate2]>0 OR [PaperWastageMin2]>0"
    cnDatabase.Execute "INSERT INTO AccountChild05 SELECT [Code],[Size],[Range4],[PrintRate4],[PaperWastageRate4],[PaperWastageMin4],999999,'*23003','01-APR-2021','S' FROM [AccountChild05T] WHERE [Range1]>0 OR [PrintRate4]>0 OR [PaperWastageRate4]>0 OR [PaperWastageMin4]>0"
    cnDatabase.Execute "INSERT INTO AccountChild05 SELECT [Code],[Size],[Range6],[PrintRate6],[PaperWastageRate6],[PaperWastageMin6],999999,'*23004','01-APR-2021','S' FROM [AccountChild05T] WHERE [Range1]>0 OR [PrintRate6]>0 OR [PaperWastageRate6]>0 OR [PaperWastageMin6]>0"

'AccountChild06
    cnDatabase.Execute "IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE  TABLE_NAME = 'AccountChild06') DROP TABLE AccountChild06"
    cnDatabase.Execute "IF Not EXISTS (Select * FROM sys.foreign_keys  Where Name ='FK_AccountChild06_AccountMaster_II') Print 'Col_Not_Exist' ELSE ALTER TABLE dbo.AccountChild06 DROP CONSTRAINT FK_AccountChild06_AccountMaster_II;"
    cnDatabase.Execute "CREATE TABLE [AccountChild06](  [Code] [nvarchar](6) NOT NULL,[SizeGroup] [nvarchar](6) NOT NULL,[Rate] [decimal](12, 2) NOT NULL,[Plate] [nvarchar](6) NOT NULL,[WEF] [date] NOT NULL DEFAULT '01-APR-2021',[Type] [char](1) NOT NULL DEFAULT 'S' CONSTRAINT [FK_AccountChild06_AccountMaster_I] FOREIGN KEY([Code]) REFERENCES [AccountMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE, CONSTRAINT [FK_AccountChild06_GeneralMaster_II] FOREIGN KEY([SizeGroup]) REFERENCES [GeneralMaster] ([Code]), CONSTRAINT [FK_AccountChild06_GeneralMaster_III] FOREIGN KEY([Plate]) REFERENCES [GeneralMaster] ([Code])) ON [PRIMARY] "

   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*24001' OR Name='Deep-etch') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*24001','Deep-etch','Deep-etch','24','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*24002' OR Name='Wipe-on') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*24002','Wipe-on','Wipe-on','24','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*24003' OR Name='PS') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*24003','PS','PS','24','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*24004' OR Name='CTP') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*24004','CTP','CTP','24','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"

    cnDatabase.Execute "INSERT INTO AccountChild06 SELECT DISTINCT [Code],[Size],[PSPlateRate1],'*24003','01-APR-2021','S' FROM [AccountChild05T] WHERE [PSPlateRate1]>0 UNION SELECT DISTINCT [Code],[Size],[DeepatchPlateRate1],'*24001','01-APR-2021','S' FROM [AccountChild05T] WHERE [DeepatchPlateRate1]>0 UNION SELECT DISTINCT [Code],[Size],[WipeonPlateRate1],'*24002','01-APR-2021','S' FROM [AccountChild05T] WHERE [WipeonPlateRate1]>0 UNION SELECT DISTINCT [Code],[Size],[CTPPlateRate1],'*24004','01-APR-2021','S' FROM [AccountChild05T] WHERE [CTPPlateRate1]>0 UNION " & _
                                      "SELECT DISTINCT [Code],[Size],[PSPlateRate2],'*24003','01-APR-2021','S' FROM [AccountChild05T] WHERE [PSPlateRate2]>0 UNION SELECT DISTINCT [Code],[Size],[DeepatchPlateRate2],'*24001','01-APR-2021','S' FROM [AccountChild05T] WHERE [DeepatchPlateRate2]>0 UNION SELECT DISTINCT [Code],[Size],[WipeonPlateRate2],'*24002','01-APR-2021','S' FROM [AccountChild05T] WHERE [WipeonPlateRate2]>0 UNION SELECT DISTINCT [Code],[Size],[CTPPlateRate2],'*24004','01-APR-2021','S' FROM [AccountChild05T] WHERE [CTPPlateRate2]>0 UNION " & _
                                      "SELECT DISTINCT [Code],[Size],[PSPlateRate4],'*24003','01-APR-2021','S' FROM [AccountChild05T] WHERE [PSPlateRate4]>0 UNION SELECT DISTINCT [Code],[Size],[DeepatchPlateRate4],'*24001','01-APR-2021','S' FROM [AccountChild05T] WHERE [DeepatchPlateRate4]>0 UNION SELECT DISTINCT [Code],[Size],[WipeonPlateRate4],'*24002','01-APR-2021','S' FROM [AccountChild05T] WHERE [WipeonPlateRate4]>0 UNION SELECT DISTINCT [Code],[Size],[CTPPlateRate4],'*24004','01-APR-2021','S' FROM [AccountChild05T] WHERE [CTPPlateRate4]>0 UNION " & _
                                      "SELECT DISTINCT [Code],[Size],[PSPlateRate6],'*24003','01-APR-2021','S' FROM [AccountChild05T] WHERE [PSPlateRate6]>0 UNION SELECT DISTINCT [Code],[Size],[DeepatchPlateRate6],'*24001','01-APR-2021','S' FROM [AccountChild05T] WHERE [DeepatchPlateRate6]>0 UNION SELECT DISTINCT [Code],[Size],[WipeonPlateRate6],'*24002','01-APR-2021','S' FROM [AccountChild05T] WHERE [WipeonPlateRate6]>0 UNION SELECT DISTINCT [Code],[Size],[CTPPlateRate6],'*24004','01-APR-2021','S' FROM [AccountChild05T] WHERE [CTPPlateRate6]>0 "
    cnDatabase.Execute "DROP TABLE AccountChild05T"

'AccountChild07
'Bindery Process
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07036' OR Name='BP-Unit Cost') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07036','BP-Unit Cost','BP-Unit Cost','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07037' OR Name='BP-Stitching') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07037','BP-Stitching','BP-Stitching','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07038' OR Name='BP-Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07038','BP-Binding','BP-Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07039' OR Name='BP-Folding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07039','BP-Folding','BP-Folding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07040' OR Name='BP-Shrink Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07040','BP-Shrink Packing','BP-Shrink Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07041' OR Name='BP-Box Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07041','BP-Box Packing','BP-Box Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07042' OR Name='BP-Cartage') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07042','BP-Cartage','BP-Cartage','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Calculation Mode
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20001' OR Name='Per Unit') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20001','Per Unit','Per Unit','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20002' OR Name='Per Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20002','Per Inch²','Per Inch²','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20003' OR Name='100 Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20003','100 Inch²','100 Inch²','20','100','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20004' OR Name='1000 Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20004','1000 Inch²','1000 Inch²','20','1000','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20005' OR Name='Per 1000') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20005','Per 1000','Per 1000','20','1000','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20006' OR Name='Per Packet') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20006','Per Packet','Per Packet','20','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20007' OR Name='Per Page') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20007','Per Page','Per Page','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20008' OR Name='Per Paisa Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20008','Per Paisa Inch²','Per Paisa Inch²','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20009' OR Name='Per Box') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20009','Per Box','Per Box','20','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20010' OR Name='Per Bundle') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20010','Per Bundle','Per Bundle','20','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Update Size Group
    cnDatabase.Execute "Update GeneralMaster Set Name= '02.00X03.00-Visiting Card',PrintName= '02.00X03.00-Visiting Card' Where Name='Visiting Card'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X30.00-Medium',PrintName= '20.00X30.00-Medium' Where Name='Medium(20''''X30'''')'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X30.00-Medium_A/P',PrintName= '20.00X30.00-Medium_A/P' Where Name='Medium(20''''X30'''')A/P'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '23.00X36.00-Large',PrintName= '23.00X36.00-Large' Where Name='Large(23''''X36'''')'"
    cnDatabase.Execute "Update GeneralMaster Set Name= '23.00X36.00-Large_A/P',PrintName= '23.00X36.00-Large_A/P' Where Name='Large(23''''X36'''')A/P'"
'Add Columns
    cnDatabase.Execute "IF NOT EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='AccountChild07' AND COLUMN_NAME='Type') ALTER TABLE AccountChild07 ADD [WEF] [date] NOT NULL DEFAULT ('01-APR-2021') WITH VALUES,[Type] [char](1) NOT NULL DEFAULT ('S') WITH VALUES"
     cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'AccountChild07' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'AccountChild07', 'AccountChild07T'"
     cnDatabase.Execute "CREATE TABLE dbo.AccountChild07 (Code nvarchar(6) NOT NULL,BinderyProcess nvarchar(6) NOT NULL,CalcMode nvarchar(6) NOT NULL,CalcValue decimal(12, 4) NOT NULL,Size nvarchar(6) NULL,Fraction decimal(12, 4) NOT NULL DEFAULT (1),AreaRange decimal(12, 3) NOT NULL DEFAULT (0),SectionRange decimal(12, 3) NOT NULL DEFAULT (1),QtyRange decimal(12, 3) NOT NULL DEFAULT (99999.99),Rate decimal(12, 4) NOT NULL,AddOnRate decimal(12, 4) NOT NULL DEFAULT (0),WEF date NOT NULL DEFAULT ('01-APR-2021'),Type char(1) NOT NULL DEFAULT ('S') " & _
                                       "CONSTRAINT [FK_AccountChild07_AccountMaster_I] FOREIGN KEY([Code]) REFERENCES [AccountMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE, " & _
                                       "CONSTRAINT [FK_AccountChild07_GeneralMaster_II] FOREIGN KEY([Size]) REFERENCES [GeneralMaster] ([Code]), " & _
                                       "CONSTRAINT [FK_AccountChild07_GeneralMaster_III] FOREIGN KEY([BinderyProcess]) REFERENCES [GeneralMaster] ([Code]), " & _
                                       "CONSTRAINT [FK_AccountChild07_GeneralMaster_IV] FOREIGN KEY([CalcMode]) REFERENCES [GeneralMaster] ([Code]))  ON [PRIMARY]"
cnDatabase.Execute "INSERT INTO dbo.AccountChild07 Select Code,LaminationType As BinderyProcess,CalcMode,(Select Value1 From GeneralMaster Where Code=CalcMode) As CalcValue,   Size,1 As Fraction,(Select(Convert(decimal,Left(Name,5))*Convert(decimal,SUBSTRING(Name,7,5))) From GeneralMaster Where Code=Size) As AreaRange,(1) As SectionRange,Range As QtyRange,Rate,0 As AddON,'01-APR-2021' AS WEF,'S' AS Type From dbo.AccountChild07T"
cnDatabase.Execute "INSERT INTO dbo.AccountChild07 " & _
"SELECT [Code],'*07036' AS BinderyProcess,'*20001' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20001') As CalcValue,[Size],4 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/4) As AreaRange,[Range04] As SectionRange,(99999.99) As QtyRange,[Rate/Book04] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [Rate/Book04]>0 Union All SELECT [Code],'*07036' As BinderyProcess,'*20001' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20001') As CalcValue,[Size],6 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/6) As AreaRange,[Range06] As SectionRange,(99999.99) As QtyRange,[Rate/Book06] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [Rate/Book06]>0 Union All " & _
"SELECT [Code],'*07036' AS BinderyProcess,'*20001' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20001') As CalcValue,[Size],8 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/8) As AreaRange,[Range08] As SectionRange,(99999.99) As QtyRange,[Rate/Book08] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [Rate/Book08]>0 Union All SELECT [Code],'*07036' AS BinderyProcess,'*20001' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20001') As CalcValue,[Size],12 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/12) As AreaRange,[Range12] As SectionRange,(99999.99) As QtyRange,[Rate/Book12] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [Rate/Book12]>0 Union All " & _
"SELECT [Code],'*07036' AS BinderyProcess,'*20001' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20001') As CalcValue,[Size],16 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/16) As AreaRange,[Range16] As SectionRange,(99999.99) As QtyRange,[Rate/Book16] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [Rate/Book16]>0 Union All SELECT [Code],'*07036' AS BinderyProcess,'*20001' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20001') As CalcValue,[Size],24 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/24) As AreaRange,[Range24] As SectionRange,(99999.99) As QtyRange,[Rate/Book24] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [Rate/Book24]>0 Union All " & _
"SELECT [Code],'*07036' AS BinderyProcess,'*20001' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20001') As CalcValue,[Size],32 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/32) As AreaRange,[Range32] As SectionRange,(99999.99) As QtyRange,[Rate/Book32] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [Rate/Book32]>0 Union All SELECT [Code],'*07036' AS BinderyProcess,'*20001' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20001') As CalcValue,[Size],64 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/64) As AreaRange,[Range64] As SectionRange,(99999.99) As QtyRange,[Rate/Book64] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [Rate/Book64]>0 Union All " & _
"SELECT [Code],'*07037' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],4 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/4) As AreaRange,[Range04] As SectionRange,(99999.99) As QtyRange,[FormStitchRate04] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormStitchRate04]>0 Union All SELECT [Code],'*07037' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],6 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/6) As AreaRange,[Range06] As SectionRange,(99999.99) As QtyRange,[FormStitchRate06] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormStitchRate06]>0 Union All " & _
"SELECT [Code],'*07037' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],8 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/8) As AreaRange,[Range08] As SectionRange,(99999.99) As QtyRange,[FormStitchRate08] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormStitchRate08]>0 Union All SELECT [Code],'*07037' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],12 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/12) As AreaRange,[Range12] As SectionRange,(99999.99) As QtyRange,[FormStitchRate12] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormStitchRate12]>0 Union All " & _
"SELECT [Code],'*07037' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],16 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/16) As AreaRange,[Range16] As SectionRange,(99999.99) As QtyRange,[FormStitchRate16] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormStitchRate16]>0 Union All SELECT [Code],'*07037' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],24 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/24) As AreaRange,[Range24] As SectionRange,(99999.99) As QtyRange,[FormStitchRate24] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormStitchRate24]>0 Union All " & _
"SELECT [Code],'*07037' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],32 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/32) As AreaRange,[Range32] As SectionRange,(99999.99) As QtyRange,[FormStitchRate32] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormStitchRate32]>0 Union All SELECT [Code],'*07037' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],64 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/64) As AreaRange,[Range64] As SectionRange,(99999.99) As QtyRange,[FormStitchRate64] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormStitchRate64]>0 Union All " & _
"SELECT [Code],'*07039' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],4 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/4) As AreaRange,[Range04] As SectionRange,(99999.99) As QtyRange,[FormFoldRate04] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormFoldRate04]>0 Union All SELECT [Code],'*07039' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],6 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/6) As AreaRange,[Range06] As SectionRange,(99999.99) As QtyRange,[FormFoldRate06] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormFoldRate06]>0 Union All " & _
"SELECT [Code],'*07039' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],8 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/8) As AreaRange,[Range08] As SectionRange,(99999.99) As QtyRange,[FormFoldRate08] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormFoldRate08]>0 Union All SELECT [Code],'*07039' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],12 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/12) As AreaRange,[Range12] As SectionRange,(99999.99) As QtyRange,[FormFoldRate12] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormFoldRate12]>0 Union All " & _
"SELECT [Code],'*07039' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],16 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/16) As AreaRange,[Range16] As SectionRange,(99999.99) As QtyRange,[FormFoldRate16] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormFoldRate16]>0 Union All SELECT [Code],'*07039' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],24 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/24) As AreaRange,[Range24] As SectionRange,(99999.99) As QtyRange,[FormFoldRate24] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormFoldRate24]>0 Union All " & _
"SELECT [Code],'*07039' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],32 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/32) As AreaRange,[Range32] As SectionRange,(99999.99) As QtyRange,[FormFoldRate32] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormFoldRate32]>0   Union All SELECT [Code],'*07039' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],64 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/64) As AreaRange,[Range64] As SectionRange,(99999.99) As QtyRange,[FormFoldRate64] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormFoldRate64]>0 Union All " & _
"SELECT [Code],'*07038' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],4 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/4) As AreaRange,[Range04] As SectionRange,(99999.99) As QtyRange,[FormPasteRate04] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormPasteRate04]>0  Union All SELECT [Code],'*07038' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],6 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/6) As AreaRange,[Range06] As SectionRange,(99999.99) As QtyRange,[FormPasteRate06] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormPasteRate06]>0 Union All " & _
"SELECT [Code],'*07038' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],8 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/8) As AreaRange,[Range08] As SectionRange,(99999.99) As QtyRange,[FormPasteRate08] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormPasteRate08]>0 Union All SELECT [Code],'*07038' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],12 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/12) As AreaRange,[Range12] As SectionRange,(99999.99) As QtyRange,[FormPasteRate12] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormPasteRate12]>0 Union All " & _
"SELECT [Code],'*07038' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],16 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/16) As AreaRange,[Range16] As SectionRange,(99999.99) As QtyRange,[FormPasteRate16] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormPasteRate16]>0 Union All SELECT [Code],'*07038' AS BinderyProcess,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,[Size],24 AS Fraction,(Convert(NUMERIC,Left((Select Name From GeneralMaster Where Code=Size),5))*Convert(NUMERIC,SubString((Select Name From GeneralMaster Where Code=Size),7,5))/24) As AreaRange,[Range24] As SectionRange,(99999.99) As QtyRange,[FormPasteRate24] As Rate,0 AS AddON, '01-Apr-2021' AS WEF,'S' AS Type  FROM [AccountChild08] WHERE [FormPasteRate24]>0 Order By CalcMode"

'AccountChild08
    cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'AccountChild08' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'AccountChild08', 'AccountChild08T'"
    cnDatabase.Execute "CREATE TABLE [AccountChild08](  [Code] [nvarchar](6) NOT NULL,[BindingType] [nvarchar](6) NOT NULL,[BinderyProcess] [nvarchar](6) NOT NULL,[CalcMode] [nvarchar](6) NOT NULL,[SizeGroup] [nvarchar](6) NOT NULL,[Fraction] [tinyint] NOT NULL,[Range] [decimal](12, 0) NOT NULL,[Rate] [decimal](12, 2) NOT NULL,[AddOnRate] [decimal](12, 2) NOT NULL DEFAULT (0),[WEF] [date] NOT NULL DEFAULT '01-APR-2021',[Type] [char](1) NOT NULL DEFAULT 'S' " & _
                                      "CONSTRAINT [FK_AccountChild08_AccountMaster_I] FOREIGN KEY([Code]) REFERENCES [AccountMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE, " & _
                                      "CONSTRAINT [FK_AccountChild08_GeneralMaster_II] FOREIGN KEY([BindingType]) REFERENCES [GeneralMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_AccountChild08_GeneralMaster_III] FOREIGN KEY([BinderyProcess]) REFERENCES [GeneralMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_AccountChild08_GeneralMaster_IV] FOREIGN KEY([CalcMode]) REFERENCES [GeneralMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_AccountChild08_AccountMaster_V] FOREIGN KEY([SizeGroup]) REFERENCES [GeneralMaster] ([Code])) ON [PRIMARY]"
    
    cnDatabase.Execute "INSERT INTO AccountChild08  " & _
    "SELECT [Code],[BindingType],'*07036','*20001',[Size],4,[Range04],[Rate/Book04],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [Rate/Book04]>0 Union All SELECT [Code],[BindingType],'*07036','*20001',[Size],6,[Range06],[Rate/Book06],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [Rate/Book06]>0 Union All SELECT [Code],[BindingType],'*07036','*20001',[Size],8,[Range08],[Rate/Book08],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [Rate/Book08]>0 Union All SELECT [Code],[BindingType],'*07036','*20001',[Size],12,[Range12],[Rate/Book12],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [Rate/Book12]>0 Union All SELECT [Code],[BindingType],'*07036','*20001',[Size],16,[Range16],[Rate/Book16],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [Rate/Book16]>0 Union All SELECT [Code],[BindingType],'*07036','*20001',[Size],24,[Range24],[Rate/Book24],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [Rate/Book24]>0 Union All  " & _
    "SELECT [Code],[BindingType],'*07036','*20001',[Size],32,[Range32],[Rate/Book32],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [Rate/Book32]>0 Union All SELECT [Code],[BindingType],'*07036','*20001',[Size],64,[Range64],[Rate/Book64],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [Rate/Book64]>0 Union All " & _
    "SELECT [Code],[BindingType],'*07037','*20005',[Size],4,[Range04],[FormStitchRate04],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormStitchRate04]>0 Union All SELECT [Code],[BindingType],'*07037','*20005',[Size],6,[Range06],[FormStitchRate06],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormStitchRate06]>0 Union All SELECT [Code],[BindingType],'*07037','*20005',[Size],8,[Range08],[FormStitchRate08],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormStitchRate08]>0 Union All SELECT [Code],[BindingType],'*07037','*20005',[Size],12,[Range12],[FormStitchRate12],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormStitchRate12]>0 Union All SELECT [Code],[BindingType],'*07037','*20005',[Size],16,[Range16],[FormStitchRate16],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormStitchRate16]>0 Union All SELECT [Code],[BindingType],'*07037','*20005',[Size],24,[Range24],[FormStitchRate24],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormStitchRate24]>0 Union All  " & _
    "SELECT [Code],[BindingType],'*07037','*20005',[Size],32,[Range32],[FormStitchRate32],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormStitchRate32]>0 Union All SELECT [Code],[BindingType],'*07037','*20005',[Size],64,[Range64],[FormStitchRate64],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormStitchRate64]>0 Union All   " & _
    "SELECT [Code],[BindingType],'*07039','*20005',[Size],4,[Range04],[FormFoldRate04],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormFoldRate04]>0 Union All SELECT [Code],[BindingType],'*07039','*20005',[Size],6,[Range06],[FormFoldRate06],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormFoldRate06]>0 Union All SELECT [Code],[BindingType],'*07039','*20005',[Size],8,[Range08],[FormFoldRate08],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormFoldRate08]>0 Union All SELECT [Code],[BindingType],'*07039','*20005',[Size],12,[Range12],[FormFoldRate12],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormFoldRate12]>0 Union All SELECT [Code],[BindingType],'*07039','*20005',[Size],16,[Range16],[FormFoldRate16],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormFoldRate16]>0 Union All SELECT [Code],[BindingType],'*07039','*20005',[Size],24,[Range24],[FormFoldRate24],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormFoldRate24]>0 Union All  " & _
    "SELECT [Code],[BindingType],'*07039','*20005',[Size],32,[Range32],[FormFoldRate32],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormFoldRate32]>0 Union All SELECT [Code],[BindingType],'*07039','*20005',[Size],64,[Range64],[FormFoldRate64],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormFoldRate64]>0 Union All  " & _
    "SELECT [Code],[BindingType],'*07038','*20005',[Size],4,[Range04],[FormPasteRate04],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormPasteRate04]>0 Union All SELECT [Code],[BindingType],'*07038','*20005',[Size],6,[Range06],[FormPasteRate06],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormPasteRate06]>0 Union All SELECT [Code],[BindingType],'*07038','*20005',[Size],8,[Range08],[FormPasteRate08],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormPasteRate08]>0 Union All SELECT [Code],[BindingType],'*07038','*20005',[Size],12,[Range12],[FormPasteRate12],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormPasteRate12]>0 Union All SELECT [Code],[BindingType],'*07038','*20005',[Size],16,[Range16],[FormPasteRate16],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormPasteRate16]>0 Union All SELECT [Code],[BindingType],'*07038','*20005',[Size],24,[Range24],[FormPasteRate24],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormPasteRate24]>0 Union All  " & _
    "SELECT [Code],[BindingType],'*07038','*20005',[Size],32,[Range32],[FormPasteRate32],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormPasteRate32]>0 Union All SELECT [Code],[BindingType],'*07038','*20005',[Size],64,[Range64],[FormPasteRate64],0,'01-Apr-2021','S' FROM [AccountChild08T] WHERE [FormPasteRate64]>0  "

    cnDatabase.Execute "DROP TABLE AccountChild08T"
'BindingTypeChild
    cnDatabase.Execute "IF Not Exists (Select *From INFORMATION_SCHEMA.COLUMNS Where TABLE_NAME = 'BindingTypeChild' AND COLUMN_NAME = 'CODE') Print 'Col_Not_Exist' ELSE DROP TABLE BindingTypeChild"
    cnDatabase.Execute "CREATE TABLE [BindingTypeChild]([Code] [nvarchar](6) NOT NULL,[BinderyProcess] [nvarchar](6) NOT NULL CONSTRAINT [FK_BindingTypeChild_GeneralMaster_I] FOREIGN KEY([Code]) REFERENCES [GeneralMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE,CONSTRAINT [FK_BindingTypeChild_GeneralMaster_II] FOREIGN KEY([BinderyProcess]) REFERENCES [GeneralMaster] ([Code]) ) ON [PRIMARY]"
'AccountChild0801
    cnDatabase.Execute "IF NOT EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='AccountChild0801' AND COLUMN_NAME='SubItem') ALTER TABLE AccountChild0801 ADD [SubItem] [nvarchar](6) NULL"
End Function
Public Function UpdateMajor02()
'    cnDatabase.Execute "Update GeneralMaster Set Name= '02.00X03.00-Visiting Card',PrintName= '02.00X03.00-Visiting Card' Where Name='Visiting Card'"
'    cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X30.00-Medium',PrintName= '20.00X30.00-Medium' Where Name='Medium(20''''X30'''')'"
'    cnDatabase.Execute "Update GeneralMaster Set Name= '20.00X30.00-Medium_A/P',PrintName= '20.00X30.00-Medium_A/P' Where Name='Medium(20''''X30'''')A/P'"
'    cnDatabase.Execute "Update GeneralMaster Set Name= '23.00X36.00-Large',PrintName= '23.00X36.00-Large' Where Name='Large(23''''X36'''')'"
'    cnDatabase.Execute "Update GeneralMaster Set Name= '23.00X36.00-Large_A/P',PrintName= '23.00X36.00-Large_A/P' Where Name='Large(23''''X36'''')A/P'"
''BookMaster
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='SaleLY1003') ALTER TABLE BookMaster DROP COLUMN SaleLY1003"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='SaleTY0409') ALTER TABLE BookMaster DROP COLUMN SaleTY0409"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='StockTransferLY1003') ALTER TABLE BookMaster DROP COLUMN StockTransferLY1003"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='StockTransferTY0409') ALTER TABLE BookMaster DROP COLUMN StockTransferTY0409"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='SpecimenLY1003') ALTER TABLE BookMaster DROP COLUMN SpecimenLY1003"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='SpecimenTY0409') ALTER TABLE BookMaster DROP COLUMN SpecimenTY0409"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='PendingSO') ALTER TABLE BookMaster DROP COLUMN PendingSO"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='SaleableStock') ALTER TABLE BookMaster DROP COLUMN SaleableStock"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='RepairableStock') ALTER TABLE BookMaster DROP COLUMN RepairableStock"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='POLTLY1003') ALTER TABLE BookMaster DROP COLUMN POLTLY1003"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='POLY0409') ALTER TABLE BookMaster DROP COLUMN POLY0409"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='POLY1003') ALTER TABLE BookMaster DROP COLUMN POLY1003"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='POTY0409') ALTER TABLE BookMaster DROP COLUMN POTY0409"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='PendingPO') ALTER TABLE BookMaster DROP COLUMN PendingPO"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='ESO30') ALTER TABLE BookMaster DROP COLUMN ESO30"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='ESO60') ALTER TABLE BookMaster DROP COLUMN ESO60"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='ESO90') ALTER TABLE BookMaster DROP COLUMN ESO90"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='ESO150') ALTER TABLE BookMaster DROP COLUMN ESO150"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='PSO15') ALTER TABLE BookMaster DROP COLUMN PSO15"
'    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='PSO30') ALTER TABLE BookMaster DROP COLUMN PSO30"
''    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='Royalty') ALTER TABLE BookMaster DROP COLUMN Royalty"
''    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='Qty/Pkt') ALTER TABLE BookMaster DROP COLUMN [Qty/Pkt]"
''    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='LooseQty/Box') ALTER TABLE BookMaster DROP COLUMN [LooseQty/Box]"
''    cnDatabase.Execute "IF  EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookMaster' AND COLUMN_NAME='Pkt/Box') ALTER TABLE BookMaster DROP COLUMN [Pkt/Box]"
End Function
Public Function UpdateMajor03() 'Update BookChild05,BookPOChild05
'GeneralMaster,ElementMaster
 cnDatabase.Execute "IF NOT EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS WHERE TABLE_NAME='ElementMaster' AND CONSTRAINT_TYPE='PRIMARY KEY') ALTER TABLE ElementMaster ADD PRIMARY KEY (Code)"
 cnDatabase.Execute "IF NOT EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS WHERE TABLE_NAME='GeneralMaster' AND CONSTRAINT_TYPE='PRIMARY KEY') ALTER TABLE GeneralMaster ADD PRIMARY KEY (Code)"
 cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11098' OR Name='*None') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11098','*None','*None','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"

'Update BookPOChild05
cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookPOChild05' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'BookPOChild05', 'BookPOChild05T'"
cnDatabase.Execute "CREATE TABLE dbo.BookPOChild05(Code nvarchar(6) NOT NULL,OrderDate datetime NOT NULL,TargetDate datetime NOT NULL,SubItem nvarchar(6) NOT NULL,Element nvarchar(6) NOT NULL,ElementPrintName nvarchar(60) NOT NULL DEFAULT (''),FinishSize nvarchar(6) NOT NULL,Size nvarchar(6) NOT NULL,DuplexPrinting bit NOT NULL,Processing nvarchar(1) NOT NULL,Ref nvarchar(40) NULL,PlateMaker nvarchar(6) NULL,ActualQuantity decimal(12, 0) NOT NULL,BillingQuantity decimal(12, 0) NOT NULL,[Pages/PrintingForm] decimal(3, 0) NOT NULL,[Pages/Form] decimal(3, 0) NOT NULL,Color nvarchar(6) NOT NULL,Pages decimal(4, 0) NOT NULL,Forms decimal(5, 2) NOT NULL,[Forms-¼] decimal(2, 0) NOT NULL,[Forms-½] decimal(2, 0) NOT NULL,[Forms-1-F&B] decimal(3, 0) NOT NULL,[Forms-1-W&T] decimal(3, 0) NOT NULL,PlateType nvarchar(6) NOT NULL,[TotalForms-¼] decimal(12, 0) NOT NULL,[TotalForms-½] decimal(12, 0) NOT NULL,[TotalForms-1-F&B] decimal(12, 0) NOT NULL,[TotalForms-1-W&T] decimal(12, 0) NOT NULL," & _
                                  "[TotalPlates-¼] decimal(12, 0) NOT NULL,[TotalPlates-½] decimal(12, 0) NOT NULL,[TotalPlates-1-F&B] decimal(12, 0) NOT NULL,[TotalPlates-1-W&T] decimal(12, 0) NOT NULL,RevisedPlates decimal(12, 0) NOT NULL," & _
                                  "[aTotalPlates-¼] decimal(12, 0) NOT NULL,[aTotalPlates-½] decimal(12, 0) NOT NULL,[aTotalPlates-1-F&B] decimal(12, 0) NOT NULL,[aTotalPlates-1-W&T] decimal(12, 0) NOT NULL,aRevisedPlates decimal(12, 0) NOT NULL," & _
                                  "PrintRate decimal(12, 2) NOT NULL,PrintAmount decimal(12, 2) NOT NULL,PlateRate decimal(12, 2) NOT NULL,PlateAmount decimal(12, 2) NOT NULL,PaperByParty bit NOT NULL,Paper nvarchar(6) NOT NULL,RAccount nvarchar(6) NOT NULL,CutOffSize smallint NOT NULL," & _
                                  "[PaperWastage%] decimal(4, 2) NOT NULL,PaperWastageMin smallint NOT NULL,[Wastage/Set] decimal(12, 0) NOT NULL,PaperWastageFinal decimal(12, 3) NOT NULL,PaperConsumptionOther decimal(12, 3) NOT NULL,PaperConsumptionsheets decimal(12, 0) NOT NULL,PaperConsumptionKg decimal(12, 3) NOT NULL," & _
                                  "[aPaperWastage%] decimal(4, 2) NOT NULL,aPaperWastageMin smallint NOT NULL,[aWastage/Set] decimal(12, 0) NOT NULL,aPaperWastageFinal decimal(12, 3) NOT NULL,aPaperConsumptionOther decimal(12, 3) NOT NULL,aPaperConsumptionsheets decimal(12, 0) NOT NULL,aPaperConsumptionKg decimal(12, 3) NOT NULL," & _
                                  "PaperRate decimal(12, 2) NOT NULL,PaperAmount decimal(12, 2) NOT NULL,   [Forms/Sheet1] decimal(4, 2) NOT NULL,[Forms/Sheet2] decimal(4, 2) NOT NULL,Remarks nvarchar(40) NULL,BillNo nvarchar(10) NULL,BillDate datetime NULL,PBillNo nvarchar(10) NULL,PBillDate datetime NULL,Adjustment decimal(12, 2) NOT NULL, " & _
                                  "PAdjustment decimal(12, 2) NOT NULL,RAdjustment decimal(12, 2) NOT NULL,[VAT%] decimal(4, 2) NOT NULL,VAT decimal(12, 2) NOT NULL,[PVAT%] decimal(4, 2) NOT NULL,PVAT decimal(12, 2) NOT NULL,[RVAT%] decimal(4, 2) NOT NULL,RVAT decimal(12, 2) NOT NULL,BillAmount decimal(12, 2) NOT NULL,PBillAmount decimal(12, 2) NOT NULL,RBillAmount decimal(12, 2) NOT NULL,PaidAmount decimal(12, 2) NOT NULL,PPaidAmount decimal(12, 2) NOT NULL,Status nvarchar(1) NULL,Narration nvarchar(40) NULL,AdjustmentRemarks nvarchar(40) NULL,DeliveredQuantityC decimal(12, 0) NOT NULL DEFAULT ((0)),DeliveredQuantityB decimal(12, 0) NOT NULL DEFAULT ((0)),BilledMFC decimal(12, 0) NOT NULL DEFAULT ((0)),BilledMFB decimal(12, 0) NOT NULL DEFAULT ((0)) " & _
                                  "CONSTRAINT [FK_BookPOChild05_BookPOParent_I] FOREIGN KEY([Code]) REFERENCES [BookPOParent] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE," & _
                                  "CONSTRAINT [FK_BookPOChild05_BookMaster_II] FOREIGN KEY([SubItem]) REFERENCES BookMaster ([Code])," & _
                                  "CONSTRAINT [FK_BookPOChild05_ElementMaster_III] FOREIGN KEY([Element]) REFERENCES ElementMaster ([Code])," & _
                                  "CONSTRAINT [FK_BookPOChild05_GeneralMaster_IV] FOREIGN KEY([FinishSize]) REFERENCES GeneralMaster ([Code])," & _
                                  "CONSTRAINT [FK_BookPOChild05_GeneralMaster_V] FOREIGN KEY([Size]) REFERENCES GeneralMaster ([Code])," & _
                                  "CONSTRAINT [FK_BookPOChild05_AccountMaster_VI] FOREIGN KEY([PlateMaker]) REFERENCES AccountMaster ([Code])," & _
                                  "CONSTRAINT [FK_BookPOChild05_PaperMaster_VII] FOREIGN KEY([Paper]) REFERENCES PaperMaster ([Code])," & _
                                  "CONSTRAINT [FK_BookPOChild05_AccountMaster_VIII] FOREIGN KEY([RAccount]) REFERENCES AccountMaster ([Code])," & _
                                  "CONSTRAINT [FK_BookPOChild05_GeneralMaster_IX] FOREIGN KEY([Color]) REFERENCES GeneralMaster ([Code])," & _
                                  "CONSTRAINT [FK_BookPOChild05_GeneralMaster_X] FOREIGN KEY([PlateType]) REFERENCES GeneralMaster ([Code]))  ON [PRIMARY]"
cnDatabase.Execute "ALTER TABLE dbo.BookPOChild05 SET (LOCK_ESCALATION = TABLE)"
cnDatabase.Execute "INSERT INTO BookPOChild05  " & _
    "SELECT DISTINCT C.[Code],[OrderDate],[TargetDate],P.Book AS SubItem,'*00011' As [Element],(Select Name From ElementMaster Where Code='*00011') As ElementPrintName,I.[FinishSize],[Size1] As [Size],IIF(I.[DuplexPrinting]='Y',1,0) As DuplexPrinting,[Processing],[Ref],[PlateMaker],[ActualQuantity],BillingQuantity01 As [BillingQuantity],CHOOSE(CONVERT(INT,I.FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/PrintingForm],CHOOSE(CONVERT(INT,FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/Form],'*23001' As [Color],[Pages1] As [Pages],[Forms1] As [Forms],[Forms1-¼] As [Forms-¼],[Forms1-½] As [Forms-½],CONVERT(INT,([Forms1-1]/2))*2 As [Forms-1-F&B],Convert(INT,(([Forms1-1]/2)-CONVERT(INT,([Forms1-1]/2)))*2) As [Forms-1-W&T],CHOOSE(CONVERT(INT,PlateType1),'*24001','*24003','*24002','*24004') as PlateType," & _
        "[TotalForms1-¼] As [TotalForms-¼],[TotalForms1-½] As [TotalForms-½],Convert(INT,Convert(INT,([Forms1-1]/2))*2* BillingQuantity01/1000)  As [TotalForms-1-F&B],Convert(INT,(([Forms1-1]-(Convert(INT,([Forms1-1]/2))*2))* BillingQuantity01)/1000) As [TotalForms-1-W&T]," & _
        "[TotalPlates1-¼] As [TotalPlates-¼],[TotalPlates1-½] As [TotalPlates-½],Convert(INT,Convert(INT,([TotalPlates1-1]/1/2))*2*1) As [TotalPlates-1-F&B],Convert(INT,[TotalPlates1-1]*1- Convert(INT,([TotalPlates1-1]/1/2))*2*1) As [TotalPlates-1-W&T],RevisedPlates1*1 As RevisedPlates," & _
        "[TotalPlates1-¼] As [aTotalPlates-¼],[TotalPlates1-½] As [aTotalPlates-½],Convert(INT,Convert(INT,([TotalPlates1-1]/1/2))*2*1) As [aTotalPlates-1-F&B],Convert(INT,[TotalPlates1-1]*1- Convert(INT,([TotalPlates1-1]/1/2))*2*1) As [aTotalPlates-1-W&T],RevisedPlates1*1 As aRevisedPlates," & _
        "PrintRate1 As PrintRate,PrintAmount1 As PrintAmount,PlateRate1 As PlateRate,PlateAmount1 As PlateAmount,PaperByParty1 As [PaperByParty],[Paper1] As [Paper],[RAccount1] As [RAccount],[CutOffSize1] As [CutOffSize]," & _
        "[PaperWastage1%] As [PaperWastage%],[PaperWastageMin1] As [PaperWastageMin],(Select (((PARSENAME(PaperWastageFinal1,2)*U.Value1)+(PARSENAME(PaperWastageFinal1,1)))/(Forms1)) From GeneralMaster U INNER JOIN PaperMaster P1 ON U.Code=P1.UOM Where C.Paper1=P1.Code) AS [Wastage/Set],[PaperWastageFinal1] As [PaperWastageFinal],[PaperConsumptionOther1] As [PaperConsumptionOther],[PaperConsumptionsheets1] As [PaperConsumptionsheets],(SELECT ROUND(([Weight/Unit]/U.Value1)*[PaperConsumptionsheets1],3) FROM PaperMaster R INNER JOIN GeneralMaster U ON R.UOM=U.Code WHERE R.Code=Paper1) As [PaperConsumptionKg]," & _
        "[PaperWastage1%] As [aPaperWastage%],[PaperWastageMin1] As [aPaperWastageMin],(Select (((PARSENAME(PaperWastageFinal1,2)*U.Value1)+(PARSENAME(PaperWastageFinal1,1)))/(Forms1)) From GeneralMaster U INNER JOIN PaperMaster P1 ON U.Code=P1.UOM Where C.Paper1=P1.Code) AS [aWastage/Set],[PaperWastageFinal1] As [aPaperWastageFinal],[PaperConsumptionOther1] As [aPaperConsumptionOther],[PaperConsumptionsheets1] As [aPaperConsumptionsheets],(SELECT ROUND(([Weight/Unit]/U.Value1)*[PaperConsumptionsheets1],3) FROM PaperMaster R INNER JOIN GeneralMaster U ON R.UOM=U.Code WHERE R.Code=Paper1) As [aPaperConsumptionKg]," & _
        "[PaperRate1] As [PaperRate],[PaperAmount1] As [PaperAmount],[Forms/Sheet1-1] As [Forms/Sheet1],[Forms/Sheet2-1] As [Forms/Sheet2],C.[Remarks],[BillNo],[BillDate],[PBillNo],[PBillDate],[Adjustment],[PAdjustment],[RAdjustment],[VAT%],[VAT],[PVAT%],[PVAT],[RVAT%],[RVAT],[BillAmount],[PBillAmount],[RBillAmount],[PaidAmount],[PPaidAmount],[Status],C.[Narration],[AdjustmentRemarks],C.DeliveredQuantityC,C.DeliveredQuantityB,BilledMFC,BilledMFB FROM ([dbo].[BookPOChild05T] C INNER JOIN BookPOParent P ON P.Code=C.Code) LEFT JOIN BookMaster I ON P.Book=I.Code WHERE [Pages1]<>0 UNION ALL " & _
    "SELECT DISTINCT C.[Code],[OrderDate],[TargetDate],P.Book AS SubItem,'*00012' As [Element],(Select Name From ElementMaster Where Code='*00012') As ElementPrintName,I.[FinishSize],[Size2] As [Size],IIF(I.[DuplexPrinting]='Y',1,0) As DuplexPrinting,[Processing],[Ref],[PlateMaker],[ActualQuantity],BillingQuantity02 As [BillingQuantity],CHOOSE(CONVERT(INT,I.FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/PrintingForm],CHOOSE(CONVERT(INT,FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/Form],'*23002' As [Color],[Pages2] As [Pages],[Forms2] As [Forms],[Forms2-¼] As [Forms-¼],[Forms2-½] As [Forms-½],CONVERT(INT,([Forms2-1]/2))*2 As [Forms-2-F&B],Convert(INT,(([Forms2-1]/2)-CONVERT(INT,([Forms2-1]/2)))*2) As [Forms-1-W&T],CHOOSE(CONVERT(INT,PlateType2),'*24001','*24003','*24002','*24004') as PlateType," & _
        "[TotalForms2-¼] As [TotalForms-¼],[TotalForms2-½] As [TotalForms-½],Convert(INT,Convert(INT,([Forms2-1]/2))*2* BillingQuantity02/1000)  As [TotalForms-1-F&B],Convert(INT,(([Forms2-1]-(Convert(INT,([Forms2-1]/2))*2))* BillingQuantity02)/1000) As [TotalForms-1-W&T]," & _
        "[TotalPlates2-¼] As [TotalPlates-¼],[TotalPlates2-½] As [TotalPlates-½],Convert(INT,Convert(INT,([TotalPlates2-1]/1/2))*2*2) As [TotalPlates-1-F&B],Convert(INT,[TotalPlates2-1]*2- Convert(INT,([TotalPlates2-1]/1/2))*2*2) As [TotalPlates-1-W&T],[RevisedPlates2]*2 As [RevisedPlates]," & _
        "[TotalPlates2-¼] As [aTotalPlates-¼],[TotalPlates2-½] As [aTotalPlates-½],Convert(INT,Convert(INT,([TotalPlates2-1]/1/2))*2*2) As [aTotalPlates-1-F&B],Convert(INT,[TotalPlates2-1]*2- Convert(INT,([TotalPlates2-1]/1/2))*2*2) As [aTotalPlates-1-W&T],[RevisedPlates2]*2 As [aRevisedPlates]," & _
        "[PrintRate2] As [PrintRate],[PrintAmount2] As [PrintAmount],[PlateRate2] As [PlateRate],[PlateAmount2] As [PlateAmount],PaperByParty2 As [PaperByParty],[Paper2] As [Paper],[RAccount2] As [RAccount],[CutOffSize2] As [CutOffSize]," & _
        "[PaperWastage2%] As [PaperWastage%],[PaperWastageMin2] As [PaperWastageMin],(Select (((PARSENAME(PaperWastageFinal2,2)*U.Value1)+(PARSENAME(PaperWastageFinal2,1)))/(Forms2)) From GeneralMaster U INNER JOIN PaperMaster P1 ON U.Code=P1.UOM Where C.Paper2=P1.Code) AS [Wastage/Set],[PaperWastageFinal2] As [PaperWastageFinal],[PaperConsumptionOther2] As [PaperConsumptionOther],[PaperConsumptionsheets2] As [PaperConsumptionsheets],(SELECT ROUND(([Weight/Unit]/U.Value1)*[PaperConsumptionsheets2],3) FROM PaperMaster R INNER JOIN GeneralMaster U ON R.UOM=U.Code WHERE R.Code=Paper2) As [PaperConsumptionKg]," & _
        "[PaperWastage2%] As [aPaperWastage%],[PaperWastageMin2] As [aPaperWastageMin],(Select (((PARSENAME(PaperWastageFinal2,2)*U.Value1)+(PARSENAME(PaperWastageFinal2,1)))/(Forms2)) From GeneralMaster U INNER JOIN PaperMaster P1 ON U.Code=P1.UOM Where C.Paper2=P1.Code) AS [aWastage/Set],[PaperWastageFinal2] As [aPaperWastageFinal],[PaperConsumptionOther2] As [aPaperConsumptionOther],[PaperConsumptionsheets2] As [aPaperConsumptionsheets],(SELECT ROUND(([Weight/Unit]/U.Value1)*[PaperConsumptionsheets2],3) FROM PaperMaster R INNER JOIN GeneralMaster U ON R.UOM=U.Code WHERE R.Code=Paper2) As [aPaperConsumptionKg]," & _
        "[PaperRate2] As [PaperRate],[PaperAmount2] As [PaperAmount],[Forms/Sheet1-2] As [Forms/Sheet1],[Forms/Sheet2-2] As [Forms/Sheet2],C.[Remarks],[BillNo],[BillDate],[PBillNo],[PBillDate],[Adjustment],[PAdjustment],[RAdjustment],[VAT%],[VAT],[PVAT%],[PVAT],[RVAT%],[RVAT],[BillAmount],[PBillAmount],[RBillAmount],[PaidAmount],[PPaidAmount],[Status],C.[Narration],[AdjustmentRemarks],C.DeliveredQuantityC,C.DeliveredQuantityB,BilledMFC,BilledMFB FROM ([dbo].[BookPOChild05T] C LEFT JOIN BookPOParent P ON P.Code=C.Code) LEFT JOIN BookMaster I ON P.Book=I.Code WHERE [Pages2]<>0 UNION ALL " & _
    "SELECT DISTINCT C.[Code],[OrderDate],[TargetDate],P.Book AS SubItem,'*00013' As [Element],(Select Name From ElementMaster Where Code='*00013') As ElementPrintName,I.[FinishSize],[Size4] As [Size],IIF(I.[DuplexPrinting]='Y',1,0) As DuplexPrinting,[Processing],[Ref],[PlateMaker],[ActualQuantity],BillingQuantity02 As [BillingQuantity],CHOOSE(CONVERT(INT,I.FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/PrintingForm],CHOOSE(CONVERT(INT,FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/Form],'*23003' As [Color],[Pages4] As [Pages],[Forms4] As [Forms],[Forms4-¼] As [Forms-¼],[Forms4-½] As [Forms-½],CONVERT(INT,([Forms4-1]/2))*2 As [Forms-2-F&B],Convert(INT,(([Forms4-1]/2)-CONVERT(INT,([Forms4-1]/2)))*2) As [Forms-1-W&T],CHOOSE(CONVERT(INT,PlateType4),'*24001','*24003','*24002','*24004') as PlateType," & _
        "[TotalForms4-¼] As [TotalForms-¼],[TotalForms4-½] As [TotalForms-½],Convert(INT,Convert(INT,([Forms4-1]/2))*2* BillingQuantity02/1000)  As [TotalForms-1-F&B],Convert(INT,(([Forms4-1]-(Convert(INT,([Forms4-1]/2))*2))* BillingQuantity02)/1000) As [TotalForms-1-W&T]," & _
        "[TotalPlates4-¼] As [TotalPlates-¼],[TotalPlates4-½] As [TotalPlates-½],Convert(INT,Convert(INT,([TotalPlates4-1]/1/2))*2*4) As [TotalPlates-1-F&B],Convert(INT,[TotalPlates4-1]*4- Convert(INT,([TotalPlates4-1]/1/2))*2*4) As [TotalPlates-1-W&T],[RevisedPlates4]*4 As [RevisedPlates]," & _
        "[TotalPlates4-¼] As [aTotalPlates-¼],[TotalPlates4-½] As [aTotalPlates-½],Convert(INT,Convert(INT,([TotalPlates4-1]/1/2))*2*4) As [aTotalPlates-1-F&B],Convert(INT,[TotalPlates4-1]*4- Convert(INT,([TotalPlates4-1]/1/2))*2*4) As [aTotalPlates-1-W&T],[RevisedPlates4]*4 As [aRevisedPlates]," & _
        "[PrintRate4] As [PrintRate],[PrintAmount4] As [PrintAmount],[PlateRate4] As [PlateRate],[PlateAmount4] As [PlateAmount],PaperByParty4 As [PaperByParty],[Paper4] As [Paper],[RAccount4] As [RAccount],[CutOffSize4] As [CutOffSize]," & _
        "[PaperWastage4%] As [PaperWastage%],[PaperWastageMin4] As [PaperWastageMin],(Select (((PARSENAME(PaperWastageFinal4,2)*U.Value1)+(PARSENAME(PaperWastageFinal4,1)))/(Forms4)) From GeneralMaster U INNER JOIN PaperMaster P1 ON U.Code=P1.UOM Where C.Paper4=P1.Code) AS [Wastage/Set],[PaperWastageFinal4] As [PaperWastageFinal],[PaperConsumptionOther4] As [PaperConsumptionOther],[PaperConsumptionsheets4] As [PaperConsumptionsheets],(SELECT ROUND(([Weight/Unit]/U.Value1)*[PaperConsumptionsheets4],3) FROM PaperMaster R INNER JOIN GeneralMaster U ON R.UOM=U.Code WHERE R.Code=Paper4) As [PaperConsumptionKg]," & _
        "[PaperWastage4%] As [aPaperWastage%],[PaperWastageMin4] As [aPaperWastageMin],(Select (((PARSENAME(PaperWastageFinal4,2)*U.Value1)+(PARSENAME(PaperWastageFinal4,1)))/(Forms4)) From GeneralMaster U INNER JOIN PaperMaster P1 ON U.Code=P1.UOM Where C.Paper4=P1.Code) AS [aWastage/Set],[PaperWastageFinal4] As [aPaperWastageFinal],[PaperConsumptionOther4] As [aPaperConsumptionOther],[PaperConsumptionsheets4] As [aPaperConsumptionsheets],(SELECT ROUND(([Weight/Unit]/U.Value1)*[PaperConsumptionsheets4],3) FROM PaperMaster R INNER JOIN GeneralMaster U ON R.UOM=U.Code WHERE R.Code=Paper4) As [aPaperConsumptionKg]," & _
        "[PaperRate4] As [PaperRate],[PaperAmount4] As [PaperAmount],[Forms/Sheet1-4] As [Forms/Sheet1],[Forms/Sheet2-4] As [Forms/Sheet2],C.[Remarks],[BillNo],[BillDate],[PBillNo],[PBillDate],[Adjustment],[PAdjustment],[RAdjustment],[VAT%],[VAT],[PVAT%],[PVAT],[RVAT%],[RVAT],[BillAmount],[PBillAmount],[RBillAmount],[PaidAmount],[PPaidAmount],[Status],C.[Narration],[AdjustmentRemarks],C.DeliveredQuantityC,C.DeliveredQuantityB,BilledMFC,BilledMFB FROM ([dbo].[BookPOChild05T] C LEFT JOIN BookPOParent P ON P.Code=C.Code) LEFT JOIN BookMaster I ON P.Book=I.Code WHERE [Pages4]<>0 "
cnDatabase.Execute "DROP TABLE BookPOChild05T"
'Update BookChild05
cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookChild05' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' Else DROP TABLE BookChild05"
If Trim(ReadFromFile("Client ID")) = "Publisher" Then ClientID = "P" Else ClientID = "S"
cnDatabase.Execute "IF EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookChild05' AND COLUMN_NAME = 'Code') Print 'Col_Exist' Else CREATE TABLE dbo.BookChild05 (Code nvarchar(6) NOT NULL,SubItem nvarchar(6) NOT NULL,Element nvarchar(6) NOT NULL,ElementPrintName nvarchar (60) NOT NULL DEFAULT (''),FinishSize nvarchar(6) NOT NULL,Size nvarchar(6) NOT NULL,DuplexPrinting bit NOT NULL,[Pages/PrintingForm] nvarchar(3) NOT NULL,[Pages/Form] nvarchar(3) NOT NULL,Color nvarchar(6) NOT NULL,Pages decimal(4, 0) NOT NULL,Forms decimal(5, 2) NOT NULL,[Forms-¼] decimal(2, 0) NOT NULL,[Forms-½] decimal(2, 0) NOT NULL,[Forms-1-F&B] decimal(3, 0) NOT NULL,[Forms-1-W&T] decimal(3, 0) NULL,PlateType nvarchar(6) NOT NULL,Ups decimal(4, 0) NOT NULL,BindingForms decimal(3, 0) NOT NULL,Type nvarchar(2) NOT NULL " & _
                                  "CONSTRAINT [FK_BookChild05_BookMaster_I] FOREIGN KEY([Code]) REFERENCES [BookMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE,  " & _
                                  "CONSTRAINT [FK_BookChild05_BookMaster_II] FOREIGN KEY([SubItem]) REFERENCES [BookMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild05_ElementMaster_III] FOREIGN KEY([Element]) REFERENCES [ElementMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild05_GeneralMaster_IV] FOREIGN KEY([FinishSize]) REFERENCES [GeneralMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild05_GeneralMaster_V] FOREIGN KEY([Size]) REFERENCES [GeneralMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild05_GeneralMaster_VI] FOREIGN KEY([Color]) REFERENCES [GeneralMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild05_GeneralMaster_VII] FOREIGN KEY([PlateType]) REFERENCES [GeneralMaster] ([Code])) ON [PRIMARY] "
                                  ','*00016' As ElementGroup
cnDatabase.Execute "INSERT INTO BookChild05  " & _
                                    "SELECT I.Code,I.Code As SubItem, '*00011' AS Element,(Select Name From ElementMaster Where Code='*00011') As ElementPrintName,I.FinishSize,I.Size,IIF(I.[DuplexPrinting]='Y',1,0) As DuplexPrinting,CHOOSE(CONVERT(INT,I.FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/PrintingForm], CHOOSE(CONVERT(INT,FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/Form],'*23001' As [Color],I.OneColorPages AS pages,I.OneColorForms AS Forms,I.[OneColor¼Forms] AS [Forms-¼],I.[OneColor½Forms] AS [Forms-½],I.[OneColor1F/BForms] AS [Forms-1-F&B],I.[OneColor1W/TForms] AS [Forms-1-W&T],CHOOSE(CONVERT(INT,I.OneColorPlateType),'*24001','*24003','*24002','*24004') as PlateType,'1' As Ups,CONVERT(INT,(I.[OneColor¼Forms]+I.[OneColor½Forms]+(I.[OneColor1F/BForms]/IIF(CHOOSE(CONVERT(INT,I.FormType),8, 16, 4, 12, 24, 32, 64, 6, 2)<=12, 2,1))+I.[OneColor1W/TForms])) As BindingForms,TYPE + Convert(nvarchar, '" & ClientID & "') As Type " & _
                                    "FROM BOOKMASTER I WHERE (((I.OneColorPages)<>0)) " & _
                                    "Union All " & _
                                    "SELECT I.Code,I.Code As SubItem, '*00012' AS Element,(Select Name From ElementMaster Where Code='*00012') As ElementPrintName,I.FinishSize,I.Size,IIF(I.[DuplexPrinting]='Y',1,0) As DuplexPrinting,CHOOSE(CONVERT(INT,I.FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/PrintingForm], CHOOSE(CONVERT(INT,FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/Form],'*23002' As [Color],I.TwoColorPages AS pages,I.TwoColorForms AS Forms,I.[TwoColor¼Forms] AS [Forms-¼],I.[TwoColor½Forms] AS [Forms-½],I.[TwoColor1F/BForms] AS [Forms-1-F&B],I.[TwoColor1W/TForms] AS [Forms-1-W&T],CHOOSE(CONVERT(INT,I.TwoColorPlateType),'*24001','*24003','*24002','*24004') as PlateType,'1' As Ups,CONVERT(INT,(I.[TwoColor¼Forms]+I.[TwoColor½Forms]+(I.[TwoColor1F/BForms]/IIF(CHOOSE(CONVERT(INT,I.FormType),8, 16, 4, 12, 24, 32, 64, 6, 2)<=12, 2,1))+I.[TwoColor1W/TForms])) As BindingForms,TYPE + Convert(nvarchar, '" & ClientID & "')  As Type " & _
                                    "FROM BOOKMASTER I WHERE (((I.TwoColorPages)<>0)) " & _
                                    "Union All " & _
                                    "SELECT I.Code,I.Code As SubItem, '*00013' AS Element,(Select Name From ElementMaster Where Code='*00013') As ElementPrintName,I.FinishSize,I.Size,IIF(I.[DuplexPrinting]='Y',1,0) As DuplexPrinting,CHOOSE(CONVERT(INT,I.FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/PrintingForm], CHOOSE(CONVERT(INT,FormType),8, 16, 4, 12, 24, 32, 64, 6, 2) As [Pages/Form],'*23003' As [Color],I.FourColorPages AS pages,I.FourColorForms AS Forms,I.[FourColor¼Forms] AS [Forms-¼],I.[FourColor½Forms] AS [Forms-½],I.[FourColor1F/BForms] AS [Forms-1-F&B],I.[FourColor1W/TForms] AS [Forms-1-W&T],CHOOSE(CONVERT(INT,I.FourColorPlateType),'*24001','*24003','*24002','*24004') as PlateType,'1' As Ups,CONVERT(INT,(I.[FourColor¼Forms]+I.[FourColor½Forms]+(I.[FourColor1F/BForms]/IIF(CHOOSE(CONVERT(INT,I.FormType),8, 16, 4, 12, 24, 32, 64, 6, 2)<=12, 2,1))+I.[FourColor1W/TForms])) As BindingForms,TYPE + Convert(nvarchar, '" & ClientID & "')  As Type " & _
                                    "FROM BOOKMASTER I WHERE (((I.FourColorPages)<>0))"
cnDatabase.Execute "ALTER FUNCTION [dbo].[ufnGetPaperStock](@Account CHAR(6),@Paper CHAR(6),@VchType CHAR(2),@VchCode CHAR(6),@VchDate DATE)  " & _
                                  " RETURNS Decimal(12, 3) " & _
                                  "AS " & _
                                  "BEGIN " & _
                                  "   DECLARE @CurStk DECIMAL(12,3); " & _
                                  "SELECT @CurStk= " & _
                                  "( " & _
                                 "(ISNULL((SELECT SUM(OpBalSheets) FROM PaperChild WHERE Code=I.Code AND Account=@Account),0)+ " & _
                                 "ISNULL((SELECT SUM(QuantitySheets) FROM PaperIOChild WHERE Paper=I.Code AND Account=@Account),0)+ " & _
                                 "ISNULL((SELECT SUM(PARSENAME(Quantity,2)*1)*U.Value1+SUM(PARSENAME(Quantity,1)*1) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=I.Code AND Quantity>=0 AND Account=@Account AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='JN',@VchCode,'XXXXXX')),0)+ " & _
                                 "ISNULL((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE [Type]='T' AND Paper=I.Code AND AccountTo=@Account AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='TR',@VchCode,'XXXXXX')),0)+ " & _
                                 "ISNULL((SELECT SUM(PARSENAME(Quantity,2)*1)*U.Value1+SUM(PARSENAME(Quantity,1)*1) FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE P.Account=@Account AND [Date]<=@VchDate AND C.Paper=I.Code AND Quantity>=0 AND P.Code<>IIF(@VchType='DN',@VchCode,'XXXXXX')),0))- " & _
                                 "(ISNULL((SELECT SUM(PARSENAME(0-Quantity,2)*1)*U.Value1+SUM(PARSENAME(0-Quantity,1)*1) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=I.Code AND Quantity<0 AND Account=@Account AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='JN',@VchCode,'XXXXXX')),0)+ " & _
                                 "ISNULL((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=I.Code AND AccountFrom=@Account AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='TR',@VchCode,'XXXXXX')),0)+ " & _
                                 "ISNULL((SELECT SUM(PARSENAME(0-Quantity,2)*1)*U.Value1+SUM(PARSENAME(0-Quantity,1)*1) FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE P.Account=@Account AND [Date]<=@VchDate AND C.Paper=I.Code AND Quantity<0 AND P.Code<>IIF(@VchType='DN',@VchCode,'XXXXXX')),0)+ " & _
                                 "ISNULL((SELECT SUM(PaperConsumptionSheets) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=I.Code AND RAccount=@Account AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='PO',@VchCode,'XXXXXX')),0)+ " & _
                                 "ISNULL((SELECT SUM(PaperConsumptionSheets) FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=I.Code AND RAccount=@Account AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='PO',@VchCode,'XXXXXX')),0)+ " & _
                                 "ISNULL((SELECT SUM(PaperConsumptionSheets) FROM BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=I.Code AND RAccount=@Account AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='PO',@VchCode,'XXXXXX')),0)+ " & _
                                 "ISNULL((SELECT SUM(Round(C2.TotalConsumption,0)) FROM (BookPOParent P INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0801 C2 ON C1.Code=C2.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND C2.Category='2' AND C2.Item=I.Code AND BookPrinter=@Account AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='PO',@VchCode,'XXXXXX')),0) " & _
                                 ") " & _
                                 ")/U.Value1 " & _
                                 "FROM PaperMaster I INNER JOIN GeneralMaster U ON I.UOM=U.Code WHERE I.Code=@Paper  " & _
                                 "RETURN PARSENAME(@CurStk,2)*1+(@CurStk-PARSENAME(@CurStk,2)*1)/2; " & _
                            "End"
End Function
Public Function UpdateMajor04() 'Update BookChild06,BookPOChild06
cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookPOChild06' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'BookChild06', 'BookChild06T'"
cnDatabase.Execute "CREATE TABLE dbo.BookChild06 (Code nvarchar(6) NOT NULL,SubItem nvarchar(6) NOT NULL,Element nvarchar(6) NOT NULL,ElementPrintName nvarchar (60) NOT NULL DEFAULT (''),Pages decimal(4, 0) NOT NULL,FinishSize nvarchar(6) NOT NULL,Size nvarchar(6) NOT NULL,Imposition nvarchar(1) NOT NULL,FrontPrintingType nvarchar(6) NULL,BackPrintingType nvarchar(6) NULL,PlateType nvarchar(6) NULL,PlateTypeBack nvarchar(6) NULL,Ups decimal(4, 2) NOT NULL,Sets decimal(4, 2) NOT NULL,BindingForms decimal(4, 2) NOT NULL,Type nvarchar(2) NOT NULL  " & _
                                  "CONSTRAINT [FK_BookChild06_BookMaster_I] FOREIGN KEY([Code]) REFERENCES [BookMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE,  " & _
                                  "CONSTRAINT [FK_BookChild06_BookMaster_II] FOREIGN KEY([SubItem]) REFERENCES [BookMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild06_ElementMaster_III] FOREIGN KEY([Element]) REFERENCES [ElementMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild06_GeneralMaster_IV] FOREIGN KEY([FinishSize]) REFERENCES [GeneralMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild06_GeneralMaster_V] FOREIGN KEY([Size]) REFERENCES [GeneralMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild06_GeneralMaster_VI] FOREIGN KEY([FrontPrintingType]) REFERENCES [GeneralMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild06_GeneralMaster_VII] FOREIGN KEY([BackPrintingType]) REFERENCES [GeneralMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild06_GeneralMaster_VIII] FOREIGN KEY([PlateType]) REFERENCES [GeneralMaster] ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild06_GeneralMaster_IX] FOREIGN KEY([PlateTypeBack]) REFERENCES [GeneralMaster] ([Code])) ON [PRIMARY] "
                                  ','*00016' As ElementGroup
cnDatabase.Execute "INSERT INTO BookChild06  " & _
                                  "SELECT I.Code,I.Code As SubItem, '*00020' AS Element,(Select Name From ElementMaster Where Code='*00020') As ElementPrintName,'4' Pages,I.FinishSize,ISNULL(IIF(I.TitleSize<>'',I.TitleSize,I.Size),I.Size) AS Size,'F' As Imposition,CHOOSE(CONVERT(INT,I.TitleFrontColor),'*23001','*23002','*23006','*23003','*23005','*23004','*23007','*23008') As FrontPrintingType,CHOOSE(CONVERT(INT,I.TitleBackColor),'*23001','*23002','*23006','*23003','*23005','*23004','*23007','*23008') As BackPrintingType,ISNULL(CHOOSE(CONVERT(INT,I.TitlePlateType),'*24001','*24003','*24002','*24004'),'') as PlateType,ISNULL(CHOOSE(CONVERT(INT,I.TitlePlateType),'*24001','*24003','*24002','*24004'),'') as PlateTypeBack,'2' As Ups,'1' As Sets,'0' As BindingForms,TYPE+ '" & ClientID & "' As Type  " & _
                                  "FROM BOOKMASTER I Where I.Code NOT IN (Select Code From BookChild06T)"
cnDatabase.Execute "INSERT INTO BookChild06  " & _
                                  "SELECT P.Book,P.Book As SubItem, Element,(Select Name From ElementMaster Where Code='*00020') As ElementPrintName,Pages,FinishSize,Size,Imposition,CHOOSE(CONVERT(INT,FrontPrintingType),'*23001','*23002','*23006','*23003','*23005','*23004','*23007','*23008') As FrontPrintingType,CHOOSE(CONVERT(INT,BackPrintingType),'*23001','*23002','*23006','*23003','*23005','*23004','*23007','*23008') As BackPrintingType,CHOOSE(CONVERT(INT,PlateType),'*24001','*24003','*24002','*24004') as PlateType,CHOOSE(CONVERT(INT,PlateTypeBack),'*24001','*24003','*24002','*24004') as PlateTypeBack,[Titles/sheet1] As Ups,'1' AS Sets,'0' As BindingForms,C.TYPE   As Type  " & _
                                  "From BookChild06T C INNER JOIN BookPOParent P ON P.Code=C.Code "
'Update Color & Plate Master Code
    cnDatabase.Execute "IF NOT EXISTS (SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='BookPOChild06' AND COLUMN_NAME='ElementPrintName') ALTER TABLE BookPOChild06 ADD [ElementPrintName] [nvarchar](60) NOT NULL DEFAULT ('') WITH VALUES"
    cnDatabase.Execute "ALTER TABLE dbo.BookPOChild06 Alter Column FrontPrintingType nvarchar(6) NOT NULL "
    cnDatabase.Execute "Update BookPOChild06 Set FrontPrintingType=ISNULL(CHOOSE(CONVERT(INT,FrontPrintingType),'*23001','*23002','*23006','*23003','*23005','*23004','*23007','*23008'),'') Where LEN(FrontPrintingType)<>6"
    cnDatabase.Execute "ALTER TABLE dbo.BookPOChild06 Alter Column BackPrintingType nvarchar(6) NOT NULL "
    cnDatabase.Execute "Update BookPOChild06 Set BackPrintingType=ISNULL(CHOOSE(CONVERT(INT,BackPrintingType),'*23001','*23002','*23006','*23003','*23005','*23004','*23007','*23008'),'') Where LEN(BackPrintingType)<>6"
    cnDatabase.Execute "ALTER TABLE dbo.BookPOChild06 Alter Column PlateType nvarchar(6) NOT NULL "
    cnDatabase.Execute "Update BookPOChild06 Set PlateType=ISNULL(CHOOSE(CONVERT(INT,PlateType),'*24001','*24003','*24002','*24004'),'') Where LEN(PlateType)<>6"
    cnDatabase.Execute "ALTER TABLE dbo.BookPOChild06 Alter Column PlateTypeBack nvarchar(6) NOT NULL "
    cnDatabase.Execute "Update BookPOChild06 Set PlateTypeBack=ISNULL(CHOOSE(CONVERT(INT,PlateTypeBack),'*24001','*24003','*24002','*24004'),'') Where LEN(PlateTypeBack)<>6"
    
    cnDatabase.Execute "ALTER TABLE dbo.BookPOChild06 Alter Column FrontPrintingType nvarchar(6) NULL"
    cnDatabase.Execute "ALTER TABLE dbo.BookPOChild06 Alter Column BackPrintingType nvarchar(6) NULL"
    cnDatabase.Execute "Update [BookPOChild06] Set FrontPrintingType = NULL where FrontPrintingType=''"
    cnDatabase.Execute "Update [BookPOChild06] Set BackPrintingType = NULL where BackPrintingType=''"
    cnDatabase.Execute "Update [BookPOChild06] Set RAccount = '000000' where RAccount=''"
'Update BookPOChild06
    cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookPOChild06' AND COLUMN_NAME = 'Code') Print 'Table_Not_Exist' ELSE EXEC sp_rename 'BookPOChild06', 'BookPOChild06T'"
    cnDatabase.Execute "CREATE TABLE dbo.BookPOChild06(Code nvarchar(6) NOT NULL,OrderDate datetime NOT NULL,TargetDate datetime NOT NULL,SubItem nvarchar(6) NOT NULL,Element nvarchar(6) NOT NULL,ElementPrintName nvarchar(60) NOT NULL,Pages decimal(4, 0) NULL,FinishSize nvarchar(6) NULL,Size nvarchar(6) NULL,Processing nvarchar(1) NOT NULL,ProcessingBack nvarchar(1) NULL,Imposition nvarchar(1) NOT NULL,Ref nvarchar(40) NULL,PlateMaker nvarchar(6) NULL,FrontPrintingType nvarchar(6) NULL,BackPrintingType nvarchar(6) NULL,PlateType nvarchar(6) NULL,PlateTypeBack nvarchar(6) NULL,ActualQuantity decimal(12, 0) NOT NULL,BillingQuantity decimal(12, 0) NOT NULL,[Titles/sheet1] decimal(4, 2) NOT NULL,Sets decimal(12, 0) NOT NULL,TotalForms decimal(12, 2) NOT NULL,TotalPlates decimal(12, 0) NOT NULL,TotalPlatesBack decimal(12, 0) NULL,aTotalPlates decimal(12, 0) NOT NULL,aTotalPlatesBack decimal(12, 0) NULL," & _
                                      "PrintRate decimal(12, 2) NOT NULL,PrintRateBack decimal(12, 2) NULL,PrintAmount decimal(12, 2) NOT NULL,PlateRate decimal(12, 2) NOT NULL,PlateRateBack decimal(12, 2) NULL,PlateAmount decimal(12, 2) NOT NULL," & _
                                      "PaperByParty bit NOT NULL,Paper nvarchar(6) NOT NULL,RAccount nvarchar(6) NOT NULL,CutOffSize decimal(6, 0) NULL,[Titles/sheet2] decimal(4, 2) NOT NULL," & _
                                      "[PaperWastage%] decimal(4, 2) NOT NULL,[PaperWastage%Back] decimal(4, 2) NULL,PaperWastageMin decimal(6, 0) NOT NULL,PaperWastageMinBack decimal(6, 0) NULL,[Wastage/Set] decimal(12, 0) NOT NULL,PaperWastageFinal decimal(12, 3) NOT NULL,PaperConsumptionOther decimal(12, 3) NOT NULL,PaperConsumptionsheets decimal(12, 0) NOT NULL,PaperConsumptionKg decimal(12, 3) NULL," & _
                                      "[aPaperWastage%] decimal(4, 2) NOT NULL,[aPaperWastage%Back] decimal(4, 2) NULL,aPaperWastageMin decimal(6, 0) NOT NULL,aPaperWastageMinBack decimal(6, 0) NULL,[aWastage/Set] decimal(12, 0) NOT NULL,aPaperWastageFinal decimal(12, 3) NOT NULL,aPaperConsumptionOther decimal(12, 3) NOT NULL,aPaperConsumptionsheets decimal(12, 0) NOT NULL,aPaperConsumptionKg decimal(12, 3) NULL," & _
                                      "PaperRate decimal(12, 2) NOT NULL,PaperAmount decimal(12, 2) NOT NULL,Remarks nvarchar(40) NULL,BillNo nvarchar(10) NULL,BillDate datetime NULL,PBillNo nvarchar(10) NULL,PBillDate datetime NULL,Adjustment decimal(12, 2) NOT NULL,PAdjustment decimal(12, 2) NOT NULL,RAdjustment decimal(12, 2) NOT NULL,[VAT%] decimal(4, 2) NOT NULL,VAT decimal(12, 2) NOT NULL,[PVAT%] decimal(4, 2) NOT NULL,PVAT decimal(12, 2) NOT NULL,[RVAT%] decimal(4, 2) NOT NULL,RVAT decimal(12, 2) NOT NULL,BillAmount decimal(12, 2) NOT NULL,PBillAmount decimal(12, 2) NOT NULL,RBillAmount decimal(12, 2) NOT NULL,PaidAmount decimal(12, 2) NOT NULL,PPaidAmount decimal(12, 2) NOT NULL,Status nvarchar(1) NULL,Narration nvarchar(40) NULL,AdjustmentRemarks nvarchar(40) NULL,ComputerName nvarchar(40) NULL," & _
                                      "DeliveredQuantityC decimal(12, 0) NOT NULL DEFAULT ((0)),DeliveredQuantityB decimal(12, 0) NOT NULL DEFAULT ((0)),BilledMEC decimal(12, 0) NOT NULL DEFAULT ((0)),BilledMEB decimal(12, 0) NOT NULL DEFAULT ((0)) " & _
                                      "CONSTRAINT [FK_BookPOChild06_BookPOParent_I] FOREIGN KEY([Code]) REFERENCES [BookPOParent] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE, " & _
                                      "CONSTRAINT [FK_BookPOChild06_BookMaster_II] FOREIGN KEY([SubItem]) REFERENCES [BookMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_BookPOChild06_ElementMaster_III] FOREIGN KEY([Element]) REFERENCES [ElementMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_BookPOChild06_GeneralMaster_IV] FOREIGN KEY([FinishSize]) REFERENCES [GeneralMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_BookPOChild06_GeneralMaster_V] FOREIGN KEY([Size]) REFERENCES [GeneralMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_BookPOChild06_AccountMaster_VI] FOREIGN KEY([PlateMaker]) REFERENCES [AccountMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_BookPOChild06_GeneralMaster_VII] FOREIGN KEY([FrontPrintingType]) REFERENCES [GeneralMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_BookPOChild06_GeneralMaster_VIII] FOREIGN KEY([BackPrintingType]) REFERENCES [GeneralMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_BookPOChild06_PaperMaster_IX] FOREIGN KEY([Paper]) REFERENCES [PaperMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_BookPOChild06_AccountMaster_X] FOREIGN KEY([RAccount]) REFERENCES [AccountMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_BookPOChild06_GeneralMaster_XI] FOREIGN KEY([PlateType]) REFERENCES [GeneralMaster] ([Code]), " & _
                                      "CONSTRAINT [FK_BookPOChild06_GeneralMaster_XII] FOREIGN KEY([PlateTypeBack]) REFERENCES [GeneralMaster] ([Code])) ON [PRIMARY] "
    
    cnDatabase.Execute "INSERT INTO BookPOChild06  " & _
                                    "SELECT C.Code,OrderDate,TargetDate,P1.Book As SubItem,Element,(Select Name From ElementMaster Where Code=Element) As ElementPrintName,Pages,FinishSize,Size,Processing,ProcessingBack,Imposition,Ref,PlateMaker,FrontPrintingType,BackPrintingType,PlateType,PlateTypeBack,ActualQuantity,BillingQuantity,[Titles/sheet1],Sets,TotalForms,TotalPlates,TotalPlatesBack,TotalPlates As aTotalPlates,TotalPlatesBack As aTotalPlatesBack,PrintRate,PrintRateBack,PrintAmount,PlateRate,PlateRateBack,PlateAmount,PaperByParty,Paper,RAccount,CutOffSize,[Titles/sheet2]," & _
                                    "[PaperWastage%],[PaperWastage%Back],PaperWastageMin,PaperWastageMinBack," & _
                                    "PARSENAME(PaperWastageFinal,2)*U.Value1+(PARSENAME(PaperWastageFinal,1)/Sets) AS [Wastage/Set]," & _
                                    "PaperWastageFinal,PaperConsumptionOther,PaperConsumptionsheets,PaperConsumptionKg," & _
                                    "[PaperWastage%] As [aPaperWastage%],[PaperWastage%Back] As [aPaperWastage%Back],PaperWastageMin AS aPaperWastageMin,PaperWastageMinBack AS aPaperWastageMinBack," & _
                                    "PARSENAME(PaperWastageFinal,2)*U.Value1+(PARSENAME(PaperWastageFinal,1)/Sets) AS [aWastage/Set]," & _
                                    "PaperWastageFinal As aPaperWastageFinal,PaperConsumptionOther As aPaperConsumptionOther,PaperConsumptionsheets As aPaperConsumptionsheets,PaperConsumptionKg As aPaperConsumptionKg," & _
                                    "PaperRate,PaperAmount,Remarks,BillNo,BillDate,PBillNo,PBillDate,Adjustment,PAdjustment,RAdjustment,[VAT%],VAT,[PVAT%],PVAT,[RVAT%],RVAT,BillAmount,PBillAmount,RBillAmount,PaidAmount,PPaidAmount,Status,Narration,AdjustmentRemarks,C.ComputerName,C.DeliveredQuantityC,C.DeliveredQuantityB,C.BilledMEC,C.BilledMEB FROM BookPOChild06T C INNER JOIN BookPOParent P1 ON P1.Code=C.Code INNER JOIN PaperMaster P ON C.Paper=P.Code INNER JOIN GeneralMaster U ON U.Code=P.UOM"
End Function
Public Function UpdateMajor05() 'Update AccountChild07
'Update BookPOChild07
   cnDatabase.Execute "Update BookPOChild07 Set  Size = NULL Where Size NOT IN ((Select Distinct Size From BookPOChild07 Where Size NOT IN (Select Code From GeneralMaster Where Type=11 Or Type=1)),'')"
   cnDatabase.Execute "IF NOT EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS WHERE TABLE_NAME='GeneralMaster' AND CONSTRAINT_TYPE='PRIMARY KEY') ALTER TABLE GeneralMaster ADD PRIMARY KEY (Code)"
   cnDatabase.Execute "IF NOT EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS WHERE TABLE_NAME='ElementMaster' AND CONSTRAINT_TYPE='PRIMARY KEY') ALTER TABLE ElementMaster ADD PRIMARY KEY (Code)"
   cnDatabase.Execute "Update BookPOChild07 Set  CalcMode ='*00021' Where CalcMode=''"
   cnDatabase.Execute "Update BookPOChild07  Set Size =I.Finishsize From BookPOChild07 C INNER JOIN BookPOParent P On P.Code=C.Code INNER JOIN BookMaster I ON P.Book=I.Code Where C.Size IS NULL Or C.Size =''"
'DROP CONSTRAINT
    cnDatabase.Execute "IF Not EXISTS (Select * FROM sys.foreign_keys  Where Name ='FK_BookPOChild07_BookPOParent_I') Print 'CONSTRAINT_Not_Exist' ELSE ALTER TABLE dbo.BookPOChild07 DROP CONSTRAINT FK_BookPOChild07_BookPOParent_I"
    cnDatabase.Execute "IF Not EXISTS (Select * FROM sys.foreign_keys  Where Name ='FK_BookPOChild07_GeneralMaster_II') Print 'CONSTRAINT_Not_Exist' ELSE ALTER TABLE dbo.BookPOChild07 DROP CONSTRAINT FK_BookPOChild07_GeneralMaster_II"
    cnDatabase.Execute "IF Not EXISTS (Select * FROM sys.foreign_keys  Where Name ='FK_BookPOChild07_GeneralMaster_III') Print 'CONSTRAINT_Not_Exist' ELSE ALTER TABLE dbo.BookPOChild07 DROP CONSTRAINT FK_BookPOChild07_GeneralMaster_III"
    cnDatabase.Execute "IF Not EXISTS (Select * FROM sys.foreign_keys  Where Name ='FK_BookPOChild07_GeneralMaster_IV') Print 'CONSTRAINT_Not_Exist' ELSE ALTER TABLE dbo.BookPOChild07 DROP CONSTRAINT FK_BookPOChild07_GeneralMaster_IV"
    cnDatabase.Execute "IF Not EXISTS (Select * FROM sys.foreign_keys  Where Name ='FK_BookPOChild07_GeneralMaster_V') Print 'CONSTRAINT_Not_Exist' ELSE ALTER TABLE dbo.BookPOChild07 DROP CONSTRAINT FK_BookPOChild07_GeneralMaster_V"
   
   cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookPOChild07' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'BookPOChild07', 'BookPOChild07T'"
   cnDatabase.Execute "CREATE TABLE dbo.BookPOChild07(Code nvarchar(6) NOT NULL,OrderDate datetime NOT NULL,TargetDate datetime NOT NULL,SubItem nvarchar(6) NOT NULL,Element nvarchar(6) NOT NULL,Operation nvarchar(6) NOT NULL,Number decimal(7, 3) NOT NULL,OperationCountName nvarchar(40) NOT NULL DEFAULT ('Nos'),Size nvarchar(6) NULL,Quantity decimal(12, 3) NOT NULL,CalcMode nvarchar(6) NOT NULL,CalcValue decimal(12, 3) NOT NULL,Rate decimal(12, 3) NOT NULL,Amount decimal(12, 2) NOT NULL,Adjustment decimal(12, 2) NOT NULL,[GST%] decimal(4, 2) NOT NULL,GST decimal(12, 2) NOT NULL,BillAmount decimal(12, 2) NOT NULL,Remarks nvarchar(40) NULL,BillNo nvarchar(10) NULL,BillDate datetime NULL,PaidAmount decimal(12, 2) NOT NULL,Status nvarchar(1) NULL,Narration nvarchar(40) NULL,DeliveredQuantityC decimal(12, 0) NOT NULL DEFAULT (0),DeliveredQuantityB decimal(12, 0) NOT NULL DEFAULT (0),BilledMOC decimal(12, 0) NOT NULL DEFAULT (0),BilledMOB decimal(12, 0) NOT NULL DEFAULT (0) " & _
                                     "CONSTRAINT [FK_BookPOChild07_BookPOParent_I] FOREIGN KEY([Code]) REFERENCES [BookPOParent] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE," & _
                                     "CONSTRAINT [FK_BookPOChild07_BookMaster_II] FOREIGN KEY([SubItem]) REFERENCES BookMaster ([Code])," & _
                                     "CONSTRAINT [FK_BookPOChild07_GeneralMaster_III] FOREIGN KEY([Element]) REFERENCES ElementMaster ([Code])," & _
                                     "CONSTRAINT [FK_BookPOChild07_GeneralMaster_IV] FOREIGN KEY([Operation]) REFERENCES GeneralMaster ([Code])," & _
                                     "CONSTRAINT [FK_BookPOChild07_GeneralMaster_V] FOREIGN KEY([SIZE]) REFERENCES GeneralMaster ([Code])," & _
                                     "CONSTRAINT [FK_BookPOChild07_GeneralMaster_VI] FOREIGN KEY([CalcMode]) REFERENCES GeneralMaster ([Code])) ON [PRIMARY]"
                                     'CONSTRAINT PK_BookPOChild07 PRIMARY KEY CLUSTERED (Code,Element,Operation),
    cnDatabase.Execute "INSERT INTO dbo.BookPOChild07 " & _
                                      "Select Code,OrderDate,TargetDate,(Select Book From BookPOParent Where Code=C.Code) As SubItem,Element,Operation,Number,'Nos' As OperationCountName,Size,Quantity,CalcMode,(Select Value1 From GeneralMaster Where Code=CalcMode) As CalcValue,Rate,Amount,Adjustment,[GST%],GST,BillAmount,Remarks,BillNo,BillDate,PaidAmount,Status,Narration,DeliveredQuantityC,DeliveredQuantityB,BilledMOC,BilledMOB From BookPOChild07T C "
    'cnDatabase.Execute "DROP TABLE BookPOChild07T"
'Update BookChild07
    cnDatabase.Execute "Delete FROM BookChild07 Where Code IN (Select Distinct Code From BookChild07 Where Code NOT IN (Select Code From BookMaster))"
    cnDatabase.Execute "Update BookChild07 Set  Size = NULL Where Size IN ((Select Distinct Size From BookChild07 Where Size NOT IN (Select Code From GeneralMaster)),'')"
    cnDatabase.Execute "Update BookChild07 Set  CalcMode ='*00021' Where CalcMode=''"
    cnDatabase.Execute "Update BookChild07  Set Size =I.Finishsize From BookChild07 C INNER JOIN BookMaster I ON C.Code=I.Code Where C.Size IS NULL OR C.Size =''"
    
    'SELECT TABLE_NAME,COLUMN_NAME,CONSTRAINT_NAME FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE TABLE_NAME = 'BookChild07';
    
    cnDatabase.Execute "IF Not EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE TABLE_NAME = 'BookChild07' AND CONSTRAINT_NAME='PK_BookChild07') Print 'CONSTRAINT_Not_Exist' ELSE ALTER TABLE dbo.BookChild07 DROP CONSTRAINT PK_BookChild07"
    cnDatabase.Execute "IF Not EXISTS (Select * FROM sys.foreign_keys  Where Name ='FK_BookChild07_BookMaster_I') Print 'CONSTRAINT_Not_Exist' ELSE ALTER TABLE dbo.BookChild07 DROP CONSTRAINT FK_BookChild07_BookMaster_I"
    cnDatabase.Execute "IF Not EXISTS (Select * FROM sys.foreign_keys  Where Name ='FK_BookChild07_GeneralMaster_II') Print 'CONSTRAINT_Not_Exist' ELSE ALTER TABLE dbo.BookChild07 DROP CONSTRAINT FK_BookChild07_GeneralMaster_II"
    cnDatabase.Execute "IF Not EXISTS (Select * FROM sys.foreign_keys  Where Name ='FK_BookChild07_GeneralMaster_III') Print 'CONSTRAINT_Not_Exist' ELSE ALTER TABLE dbo.BookChild07 DROP CONSTRAINT FK_BookChild07_GeneralMaster_III"
    cnDatabase.Execute "IF Not EXISTS (Select * FROM sys.foreign_keys  Where Name ='FK_BookChild07_GeneralMaster_IV') Print 'CONSTRAINT_Not_Exist' ELSE ALTER TABLE dbo.BookChild07 DROP CONSTRAINT FK_BookChild07_GeneralMaster_IV"
    cnDatabase.Execute "IF Not EXISTS (Select * FROM sys.foreign_keys  Where Name ='FK_BookChild07_GeneralMaster_V') Print 'CONSTRAINT_Not_Exist' ELSE ALTER TABLE dbo.BookChild07 DROP CONSTRAINT FK_BookChild07_GeneralMaster_V"
            
    cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookChild07' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'BookChild07', 'BookChild07T'"
    cnDatabase.Execute "CREATE TABLE dbo.BookChild07(Code nvarchar(6) NOT NULL,SubItem nvarchar(6) NOT NULL,Element nvarchar(6) NOT NULL,Operation nvarchar(6) NOT NULL,Number decimal(7, 3) NOT NULL,OperationCountName nvarchar(40) NOT NULL DEFAULT ('Nos'),Size nvarchar(6) NULL,CalcMode nvarchar(6) NOT NULL,CalcValue decimal(12, 3) NOT NULL,Type char(2) NOT NULL " & _
                                        "CONSTRAINT PK_BookChild07 PRIMARY KEY CLUSTERED (Code,Element,Operation,TYPE), " & _
                                        "CONSTRAINT [FK_BookChild07_BookMaster_I] FOREIGN KEY([Code]) REFERENCES [BookMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE, " & _
                                        "CONSTRAINT [FK_BookChild07_BookMaster_II] FOREIGN KEY([SubItem]) REFERENCES BookMaster ([Code]), " & _
                                        "CONSTRAINT [FK_BookChild07_ElementMaster_III] FOREIGN KEY([Element]) REFERENCES ElementMaster ([Code]), " & _
                                        "CONSTRAINT [FK_BookChild07_GeneralMaster_IV] FOREIGN KEY([Operation]) REFERENCES GeneralMaster ([Code]), " & _
                                        "CONSTRAINT [FK_BookChild07_GeneralMaster_V] FOREIGN KEY([SIZE]) REFERENCES GeneralMaster ([Code]), " & _
                                        "CONSTRAINT [FK_BookChild07_GeneralMaster_VI] FOREIGN KEY([CalcMode]) REFERENCES GeneralMaster ([Code])) ON [PRIMARY]"
    cnDatabase.Execute "INSERT INTO dbo.BookChild07 " & _
                                        "Select C.Code,(Select Code From BookMaster Where Code=C.Code) As SubItem,Element,Operation,Number,'Nos' As OperationCountName,Size,CalcMode,(Select Value1 From GeneralMaster Where Code=CalcMode) As CalcValue,Type From BookChild07T C "
    cnDatabase.Execute "DROP TABLE BookChild07T"
End Function
Public Function UpdateMajor06() 'Update BookChild08
cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookChild08' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'BookChild08', 'BookChild08T'"
cnDatabase.Execute "CREATE TABLE dbo.BookChild08 (Code nvarchar(6) NOT NULL,SubItem nvarchar(6) NOT NULL,BindingType nvarchar(6) NOT NULL,BinderyProcess nvarchar(6) NOT NULL,Number decimal(7, 3) NOT NULL DEFAULT (1),OperationCountName nvarchar(40) NOT NULL DEFAULT ('Nos'),Size nvarchar(6) NOT NULL,CalcMode nvarchar (6) NOT NULL,CalcValue Decimal (12,3) NOT NULL DEFAULT (0),Type char(2) NOT NULL  " & _
                                  "CONSTRAINT [FK_BookChild08_BookMaster_I] FOREIGN KEY([Code]) REFERENCES [BookMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE,  " & _
                                  "CONSTRAINT [FK_BookChild08_BookMaster_II] FOREIGN KEY([SubItem]) REFERENCES BookMaster ([Code]),  " & _
                                  "CONSTRAINT [FK_BookChild08_GeneralMaster_III] FOREIGN KEY([BindingType]) REFERENCES GeneralMaster ([Code]),  " & _
                                  "CONSTRAINT [FK_BookChild08_GeneralMaster_IV] FOREIGN KEY([BinderyProcess]) REFERENCES GeneralMaster ([Code]),  " & _
                                  "CONSTRAINT [FK_BookChild08_GeneralMaster_V] FOREIGN KEY([Size]) REFERENCES GeneralMaster ([Code]),  " & _
                                  "CONSTRAINT [FK_BookChild08_GeneralMaster_VI] FOREIGN KEY([CalcMode]) REFERENCES GeneralMaster ([Code])) ON [PRIMARY]"
    cnDatabase.Execute "ALTER TABLE dbo.BookChild08 SET (LOCK_ESCALATION = TABLE)"
    cnDatabase.Execute "INSERT INTO dbo.BookChild08  " & _
                                      "Select DISTINCT I.Code As Code,I.Code As SubItem,I.BindingType,C.BinderyProcess AS BinderyProcess,IIF(C.BinderyProcess='*07037',(I.BindingForms01+I.BindingForms02),IIF(C.BinderyProcess='*07039',(I.BindingForms01+I.BindingForms02),IIF(C.BinderyProcess='*07051',(I.BindingForms01+I.BindingForms02),'1'))) As Number,IIF(C.BinderyProcess='*07037',('Sections'),IIF(C.BinderyProcess='*07039',('Forms'),IIF(C.BinderyProcess='*07051',('Sections'),'Nos'))) As OperationCountName,I.FinishSize As Size,IIF(C.BinderyProcess='*07037',('*20005'),IIF(C.BinderyProcess='*07039',('*20005'),IIF(C.BinderyProcess='*07036',('*20001'),IIF(C.BinderyProcess='*07038',('*20005'),IIF(C.BinderyProcess='*07041',('*20009'),IIF(C.BinderyProcess='*07051',('*20005'),'*20006')))))) As CalcMode,  " & _
                                      "IIF(C.BinderyProcess='*07037',(1000),IIF(C.BinderyProcess='*07039',(1000),IIF(C.BinderyProcess='*07036',(1),IIF(C.BinderyProcess='*07038',(1),I.[Qty/Pkt])))) As CalcValue,I.TYPE + Convert(nvarchar, '" & ClientID & "') As Type  " & _
                                      "From BookMaster I  INNER JOIN BindingTypeChild C ON I.BindingType=C.Code Where I.Code NOT IN (Select Distinct Code From BookChild08) " & _
                                      "Order By Code,BinderyProcess"
End Function
Public Function UpdateMajor07() 'Update BookPOChild08
    cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookPOChild08' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'BookPOChild08', 'BookPOChild08T'"
    cnDatabase.Execute "CREATE TABLE dbo.BookPOChild08 (Code nvarchar(6) NOT NULL,OrderDate datetime NOT NULL,TargetDate datetime NOT NULL,SubItem nvarchar(6) NOT NULL ,BindingType nvarchar(6) NOT NULL,BinderyProcess nvarchar(6) NOT NULL,Number decimal(7, 3) NOT NULL DEFAULT (1),OperationCountName nvarchar(40) NOT NULL DEFAULT ('Nos'),Size nvarchar(6) NOT NULL,Fraction decimal(12, 3) NOT NULL DEFAULT (1),Quantity decimal(12, 3) NOT NULL,CalcMode nvarchar (6) NOT NULL,CalcValue Decimal (12,3) NOT NULL DEFAULT (1),Rate decimal(12, 3) NOT NULL,Amount decimal(12, 2) NOT NULL,Adjustment decimal(12, 2) NOT NULL,[GST%] decimal(4, 2) NOT NULL,GST decimal(12, 2) NOT NULL,BillAmount decimal(12, 2) NOT NULL,Remarks nvarchar(80) NULL,BillNo nvarchar(10) NULL,BillDate datetime NULL,PaidAmount decimal(12, 2) NOT NULL,Status nvarchar(1) NULL,Narration nvarchar(80) NULL, " & _
                                      "DeliveredQuantityC decimal(12, 0) NOT NULL DEFAULT (1),DeliveredQuantityB decimal(12, 0) NOT NULL DEFAULT (1),BilledBNC decimal(12, 0) NOT NULL DEFAULT (1),BilledBNB decimal(12, 0) NOT NULL DEFAULT (1)" & _
                                      "CONSTRAINT [FK_BookPOChild08_BookPOParent_I] FOREIGN KEY([Code]) REFERENCES [BookPOParent] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE,  " & _
                                      "CONSTRAINT [FK_BookPOChild08_BookMaster_II] FOREIGN KEY([SubItem]) REFERENCES BookMaster ([Code]),  " & _
                                      "CONSTRAINT [FK_BookPOChild08_GeneralMaster_III] FOREIGN KEY([BindingType]) REFERENCES GeneralMaster ([Code]),  " & _
                                      "CONSTRAINT [FK_BookPOChild08_GeneralMaster_IV] FOREIGN KEY([BinderyProcess]) REFERENCES GeneralMaster ([Code]),  " & _
                                      "CONSTRAINT [FK_BookPOChild08_GeneralMaster_V] FOREIGN KEY([Size]) REFERENCES GeneralMaster ([Code]),  " & _
                                      "CONSTRAINT [FK_BookPOChild08_GeneralMaster_VI] FOREIGN KEY([CalcMode]) REFERENCES GeneralMaster ([Code])) ON [PRIMARY]"
    cnDatabase.Execute "INSERT INTO dbo.BookPOChild08  " & _
    "Select C.Code,OrderDate,TargetDate,P.Book As SubItem,C.BindingType,'*07039' AS BinderyProcess,(BindingForms+ExtraForms) As Number,'Forms' As OperationCountName,I.FinishSize As Size,1,ActualQuantity As Quantity,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,FormFoldRate As Rate,((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20005')*FormFoldRate*(BindingForms+ExtraForms))) As Amount,'0' AS Adjustment,[Vat%] As [GST%],(((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20005')*FormFoldRate*(BindingForms+ExtraForms)))*[VAT%])/100 As GST,(((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20005')*FormFoldRate*(BindingForms+ExtraForms)))*([Vat%]+100))/100 As BillAmount,C.Remarks,BillNo,BillDate,PaidAmount,Status,C.Narration,C.DeliveredQuantityC,C.DeliveredQuantityB,BilledBNC,BilledBNB " & _
    "From BookPOChild08T C INNER JOIN BookPOParent P ON P.Code=C.Code INNER JOIN BookMaster I ON I.Code=P.Book Where FormFoldRate<>0  " & _
    "Union All  " & _
    "Select C.Code,OrderDate,TargetDate,P.Book As SubItem,C.BindingType,'*07037' AS BinderyProcess,(BindingForms+ExtraForms) As Number,'Sections' As OperationCountName,I.FinishSize As Size,1,ActualQuantity As Quantity,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,FormStitchRate As Rate,((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20005')*FormStitchRate*(BindingForms+ExtraForms))) As Amount,'0' AS Adjustment,[Vat%] As [GST%],((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20005')*FormStitchRate*(BindingForms+ExtraForms)))*[VAT%]/100 As  GST,(((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20005')*FormStitchRate*(BindingForms+ExtraForms)))*([Vat%]+100))/100 As BillAmount,C.Remarks,BillNo,BillDate,PaidAmount,Status,C.Narration,C.DeliveredQuantityC,C.DeliveredQuantityB,BilledBNC,BilledBNB " & _
    "From BookPOChild08T C INNER JOIN BookPOParent P ON P.Code=C.Code INNER JOIN BookMaster I ON I.Code=P.Book Where FormStitchRate<>0  " & _
    "Union All  " & _
    "Select C.Code,OrderDate,TargetDate,P.Book As SubItem,C.BindingType,'*07038' AS BinderyProcess,1 As Number,'Nos' As OperationCountName,I.FinishSize As Size,1,ActualQuantity As Quantity,'*20005' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20005') As CalcValue,FormPasteRate As Rate,((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20005')*FormPasteRate)) As Amount,'0' AS Adjustment,[Vat%] As [GST%],(((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20005')*FormPasteRate))*[VAT%])/100 As GST,(((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20005')*FormPasteRate))*([VAT%]+100))/100 As BillAmount,C.Remarks,BillNo,BillDate,PaidAmount,Status,C.Narration,C.DeliveredQuantityC,C.DeliveredQuantityB,BilledBNC,BilledBNB From BookPOChild08T C INNER JOIN BookPOParent P ON P.Code=C.Code INNER JOIN BookMaster I ON I.Code=P.Book Where FormPasteRate<>0  " & _
    "Union All  " & _
    "Select C.Code,OrderDate,TargetDate,P.Book As SubItem,C.BindingType,'*07036' AS BinderyProcess,1 As Number,'Nos' As OperationCountName,I.FinishSize As Size,1,ActualQuantity As Quantity,'*20001' As CalcMode,(Select Value1 From GeneralMaster Where Code='*20001') As CalcValue,[Rate/Book] As Rate,((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20001')*[Rate/Book])) As Amount,'0' AS Adjustment,[Vat%] As [GST%],(((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20001')*[Rate/Book]))*[VAT%])/100 As GST,(((BillingQuantity/(Select Value1 From GeneralMaster Where Code='*20001')*[Rate/Book]))*([VAT%]+100))/100 As BillAmount,C.Remarks,BillNo,BillDate,PaidAmount,Status,C.Narration,C.DeliveredQuantityC,C.DeliveredQuantityB,BilledBNC,BilledBNB From BookPOChild08T C INNER JOIN BookPOParent P ON P.Code=C.Code INNER JOIN BookMaster I ON I.Code=P.Book Where [Rate/Book]<>0  " & _
    "Union All  " & _
    "Select C.Code,OrderDate,TargetDate,P.Book As SubItem,C.BindingType,'*07040' AS BinderyProcess,1 As Number,'Nos' As OperationCountName,I.FinishSize As Size,1,ActualQuantity As Quantity,'*20006' As CalcMode,C.[Qty/Pkt] As CalcValue,PktPackRate As Rate,((TotalPkts*PktPackRate)) As Amount,'0' AS Adjustment,[Vat%] As [GST%],((((TotalPkts*PktPackRate))*[VAT%])/100) As GST,((((TotalPkts*PktPackRate))*([VAT%]+100))/100) As BillAmount,C.Remarks,BillNo,BillDate,PaidAmount,Status,C.Narration,C.DeliveredQuantityC,C.DeliveredQuantityB,BilledBNC,BilledBNB From BookPOChild08T C INNER JOIN BookPOParent P ON P.Code=C.Code INNER JOIN BookMaster I ON I.Code=P.Book Where PktPackRate<>0  " & _
    "Union All  " & _
    "Select C.Code,OrderDate,TargetDate,P.Book As SubItem,C.BindingType,'*07041' AS BinderyProcess,1 As Number,'Nos' As OperationCountName,I.FinishSize As Size,1,ActualQuantity As Quantity,'*20009' As CalcMode,C.[Pkt/BOX]*C.[Qty/Pkt] As CalcValue,BoxPackRate As Rate,((TotalBoxes*BoxPackRate)) As Amount,'0' AS Adjustment,[Vat%] As [GST%],(((TotalBoxes*BoxPackRate))*[VAT%])/100 As GST,(((TotalBoxes)*BoxPackRate*([VAT%]+100))/100) As BillAmount,C.Remarks,BillNo,BillDate,PaidAmount,Status,C.Narration,C.DeliveredQuantityC,C.DeliveredQuantityB,BilledBNC,BilledBNB From BookPOChild08T C INNER JOIN BookPOParent P ON P.Code=C.Code INNER JOIN BookMaster I ON I.Code=P.Book Where C.BoxPackRate<>0  " & _
    "Union All  " & _
    "Select C.Code,OrderDate,TargetDate,P.Book As SubItem,C.BindingType,'*07042' AS BinderyProcess,1 As Number,'Nos' As OperationCountName,I.FinishSize As Size,1,ActualQuantity As Quantity,'*20010' As CalcMode,C.[Pkt/BOX]*C.[Qty/Pkt] As CalcValue,CartageRate As Rate,((TotalBoxes*CartageRate)) As Amount,Adjustment,[Vat%] As [GST%],((((TotalBoxes*CartageRate))*[VAT%])/100) As GST,((((TotalBoxes*CartageRate))*([VAT%]+100))/100) As BillAmount,C.Remarks,BillNo,BillDate,PaidAmount,Status,C.Narration,C.DeliveredQuantityC,C.DeliveredQuantityB,BilledBNC,BilledBNB From BookPOChild08T C INNER JOIN BookPOParent P ON P.Code=C.Code INNER JOIN BookMaster I ON I.Code=P.Book Where C.CartageRate<>0 Or C.Adjustment<>0 " & _
    "Order By Code,BinderyProcess "
End Function
'//*****************************//
'//*****************************//
'//*****************************//
'//*****************************//
'//*****************************//

Public Function UpdateMinor01()
'GeneralMaster,'ElementMaster
 cnDatabase.Execute "IF NOT EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS WHERE TABLE_NAME='GeneralMaster' AND CONSTRAINT_TYPE='PRIMARY KEY') ALTER TABLE GeneralMaster ADD PRIMARY KEY (Code)"
 cnDatabase.Execute "IF NOT EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS WHERE TABLE_NAME='ElementMaster' AND CONSTRAINT_TYPE='PRIMARY KEY') ALTER TABLE ElementMaster ADD PRIMARY KEY (Code)"
'Account Master Update Table
frmLicenceAgreement.Label2 = "AccountMaster Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('AccountMaster', 'Notes') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE AccountMaster ADD Notes text NULL ALTER TABLE AccountMaster SET (LOCK_ESCALATION = TABLE) "
    cnDatabase.Execute "IF COL_LENGTH('AccountMaster', 'Notes') IS NULL PRINT 'NOT Exists' ELSE Update AccountMaster Set Notes='' Where Notes IS NULL"
    cnDatabase.Execute "IF COL_LENGTH('AccountMaster', 'Opening') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE AccountMaster ADD Opening decimal(12, 2) NULL ALTER TABLE AccountMaster SET (LOCK_ESCALATION = TABLE) "
    cnDatabase.Execute "IF COL_LENGTH('AccountMaster', 'Opening') IS NULL PRINT 'NOT Exists' ELSE Update AccountMaster Set Opening=0 Where Opening IS NULL"
'BindingTypeChild Create Table
frmLicenceAgreement.Label2 = "BindingTypeChild Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('BindingTypeChild', 'Code') IS NOT NULL PRINT 'Exists' ELSE  CREATE TABLE [BindingTypeChild]([Code] [nvarchar](6) NOT NULL,[BinderyProcess] [nvarchar](6) NOT NULL CONSTRAINT [FK_BindingTypeChild_GeneralMaster_I] FOREIGN KEY([Code]) REFERENCES [GeneralMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE,CONSTRAINT [FK_BindingTypeChild_GeneralMaster_II] FOREIGN KEY([BinderyProcess]) REFERENCES [GeneralMaster] ([Code]) ) ON [PRIMARY]"
'BookChild Create Table
frmLicenceAgreement.Label2 = "BookChild Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('BookChild', 'MaterialCentre') IS NOT NULL PRINT 'Exists' ELSE  CREATE TABLE BookChild(MaterialCentre nvarchar(6) NOT NULL,Item nvarchar(6) NOT NULL,OpBal int NOT NULL,FYCode nvarchar(6) NOT NULL ) ON [PRIMARY] ALTER TABLE BookChild SET (LOCK_ESCALATION = TABLE)"
'Book Master Update Table
frmLicenceAgreement.Label2 = "BookMaster Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('BookMaster', 'Notes') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookMaster ADD Notes text NULL ALTER TABLE BookMaster SET (LOCK_ESCALATION = TABLE) "
    cnDatabase.Execute "IF COL_LENGTH('BookMaster', 'Notes') IS NULL PRINT 'NOT Exists' ELSE Update BookMaster Set Notes='' Where Notes IS NULL"
'BookPOChild05
frmLicenceAgreement.Label2 = "BookPOChild05 Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild05 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild05 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'BilledMFC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild05 ADD BilledMFC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'BilledMFC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild05 SET BilledMFC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.BilledAllC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'BilledMFB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild05 ADD BilledMFB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'BilledMFB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild05 SET BilledMFB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.BilledAllB>0"
    'If UpdateVersion = True Then
        If MsgBox("Do You Wants to Update '21-05.7.1 Version' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild05 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild05 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
        Else
        End If
    'End If
    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
                                       "SET @table='BookPOChild05' " & _
                                       "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                       "BEGIN " & _
                                       "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                       "EXEC sp_executesql @sql " & _
                                       "End"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOChild05 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'Exists' "
'BookPOChild06 Table Update
frmLicenceAgreement.Label2 = "BookPOChild06 Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild06 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild06 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'BilledMEC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild06 ADD BilledMEC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'BilledMEC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild06 SET BilledMEC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.BilledAllC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'BilledMEB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild06 ADD BilledMEB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'BilledMEB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild06 SET BilledMEB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.BilledAllB>0"
    'If UpdateVersion = True Then
        If MsgBox("Do You Wants to Update '21-05.7.1 Version' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild06 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild06 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
        Else
        End If
    'End If
    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
                                       "SET @table='BookPOChild06' " & _
                                       "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                       "BEGIN " & _
                                        "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                        "EXEC sp_executesql @sql " & _
                                        "End"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOChild06 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE  PRINT 'Exists' "
'BookPOChild07 Table Update
frmLicenceAgreement.Label2 = "BookPOChild07 Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'Number') IS NOT NULL PRINT 'Exists' ELSE Alter Table BookPOChild07 Alter Column Number decimal(7, 3)"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'Rate') IS NOT NULL PRINT 'Exists' ELSE Alter Table BookPOChild07 Alter Column Rate decimal(12, 3)"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild07 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild07 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'BilledMOC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild07 ADD BilledMOC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'BilledMOC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild07 SET BilledMOC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.BilledAllC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'BilledMOB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild07 ADD BilledMOB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'BilledMOB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild07 SET BilledMOB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.BilledAllB>0"
    'If UpdateVersion = True Then
        If MsgBox("Do You Wants to Update '21-05.7.1 Version' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild07 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild07 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
        Else
        End If
    'End If
    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
                                      "SET @table='BookPOChild07' " & _
                                      "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                      "BEGIN " & _
                                          "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                          "EXEC sp_executesql @sql " & _
                                      "End"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOChild07 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'Not Exists'"
'BookPOChild08 Table Update
frmLicenceAgreement.Label2 = "BookPOChild08 Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild08 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild08 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'BilledBNC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild08 ADD BilledBNC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'BilledBNC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild08 SET BilledBNC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.BilledAllC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'BilledBNB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild08 ADD BilledBNB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'BilledBNB') IS NULL PRINT 'NOT Exists' UPDATE BookPOChild08 SET BilledBNB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.BilledAllB>0"
    'If UpdateVersion = True Then
        If MsgBox("Do You Wants to Update '21-05.7.1 Version' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild08 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild08 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
        Else
        End If
    'End If
    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
                                      "SET @table='BookPOChild08' " & _
                                      "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                      "BEGIN " & _
                                         "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                         "EXEC sp_executesql @sql " & _
                                     "End "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOChild08 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'NotExists' "
'BookPOChild0801 Table Update
frmLicenceAgreement.Label2 = "BookPOChild0801 Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0801 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0801 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'BilledBMC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0801 ADD BilledBMC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'BilledBMC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0801 SET BilledBMC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE P.BilledAllC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'BilledBMB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0801 ADD BilledBMB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'BilledBMB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0801 SET BilledBMB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE P.BilledAllB>0"
    'If UpdateVersion = True Then
        If MsgBox("Do You Wants to Update '21-05.7.1 Version' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0801 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0801 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
        Else
        End If
    'End If
    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
                                      "SET @table='BookPOChild0801' " & _
                                      "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                      "BEGIN " & _
                                          "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                          "EXEC sp_executesql @sql " & _
                                      "End "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'QuantityIssuedC') IS NOT NULL  ALTER TABLE BookPOChild0801 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'Exists' "
'BookPOChild09
frmLicenceAgreement.Label2 = "BookPOChild09 Update Going on!!! "
    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
                                      "SET @table='BookPOChild09' " & _
                                      "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                      "BEGIN " & _
                                          "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                          "EXEC sp_executesql @sql " & _
                                      "End "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild09', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOChild09 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'Exists' "
'BookPOChild0901 Table Update
frmLicenceAgreement.Label2 = "BookPOChild0901 Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0901 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0901 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'BilledCFC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0901 ADD BilledCFC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'BilledCFC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0901 SET BilledCFC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code WHERE P.BilledAllC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'BilledCFB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0901 ADD BilledCFB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'BilledCFB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0901 SET BilledCFB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code WHERE P.BilledAllB>0"
    'If UpdateVersion = True Then
        If MsgBox("Do You Wants to Update '21-05.7.1 Version' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0901 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0901 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
        Else
        End If
    'End If
    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
                                      "SET @table='BookPOChild0901' " & _
                                      "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                      "BEGIN " & _
                                          "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
                                          "EXEC sp_executesql @sql " & _
                                     "End"
    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'QuantityIssuedC') IS NOT NULL  ALTER TABLE BookPOChild0901 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'Exists' "
'BookPOParent Table Update
frmLicenceAgreement.Label2 = "BookPOParent Update Going on!!! "
'PaperPOParent
    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'PicData') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE dbo.BookPOParent ADD PicData varbinary(MAX) NULL"
    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'PicType') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE dbo.BookPOParent ADD PicType nvarchar(4) NULL"
    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOParent ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOParent ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
    'If UpdateVersion = True Then
        If MsgBox("Do You Wants to Update '21-05.7.1 Version' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOParent SET DeliveredQuantityC=QuantityIssuedC+QuantityReceivedC WHERE QuantityIssuedC+QuantityReceivedC>0"
    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOParent SET DeliveredQuantityB=QuantityIssuedB+QuantityReceivedB WHERE QuantityIssuedB+QuantityReceivedB>0"
   'DROP CONSTRAINT
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityIssuedC';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityReceivedC';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityIssuedB';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityReceivedB';EXEC sp_executesql @sql"
    
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityIssued07C';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityReceived07C';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityIssued07B';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityReceived07B';EXEC sp_executesql @sql"

    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityIssued0801C';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityReceived0801C';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityIssued0801B';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%QuantityReceived0801B';EXEC sp_executesql @sql"
    
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%BilledTextC';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%BilledTextB';EXEC sp_executesql @sql"
    
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%BilledTitleC';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%BilledTitleB';EXEC sp_executesql @sql"
    
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%BilledComboTitleC';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%BilledComboTitleB';EXEC sp_executesql @sql"
    
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%BilledLaminationC';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%BilledLaminationB';EXEC sp_executesql @sql"
    
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%BilledBOMC';EXEC sp_executesql @sql"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(MAX) = '';SELECT @sql += 'ALTER TABLE BookPOParent DROP CONSTRAINT ' + QUOTENAME(name) + ';' FROM sys.default_constraints WHERE Name Like '%BilledBOMB';EXEC sp_executesql @sql"
    'cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOParent DROP CONSTRAINT df_QuantityIssuedC,df_QuantityReceivedC,df_QuantityIssuedB,df_QuantityReceivedB,df_QuantityIssued07C,df_QuantityReceived07C,df_QuantityIssued07B,df_QuantityReceived07B,df_QuantityIssued0801C,df_QuantityReceived0801C,df_QuantityIssued0801B,df_QuantityReceived0801B,df_BilledTextC,df_BilledTextB,df_BilledTitleC,df_BilledTitleB,df_BilledComboTitleC,df_BilledComboTitleB,df_BilledLaminationC,df_BilledLaminationB,df_BilledBOMC,df_BilledBOMB ELSE PRINT 'Exists'  "
    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'QuantityIssuedC') IS NOT NULL  ALTER TABLE BookPOParent DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB,QuantityIssued07C,QuantityReceived07C,QuantityIssued07B,QuantityReceived07B,QuantityIssued0801C,QuantityReceived0801C,QuantityIssued0801B,QuantityReceived0801B,BilledTextC,BilledTextB,BilledTitleC,BilledTitleB,BilledComboTitleC,BilledComboTitleB,BilledLaminationC,BilledLaminationB,BilledBOMC,BilledBOMB ELSE PRINT 'Exists' "
        Else
        End If
    'End If
'Company Master Table Update Table
frmLicenceAgreement.Label2 = "CompanyMaster Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'TallyIntegration') IS NOT NULL PRINT 'Exists' ELSE  Alter Table CompanyMaster Add TallyIntegration bit NOT NULL CONSTRAINT df_TallyIntegration DEFAULT '' "
    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'TallyIntegration') IS NULL PRINT 'NOT Exists' ELSE Update CompanyMaster Set TallyIntegration=0"
    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'BusyIntegration') IS NOT NULL PRINT 'Exists' ELSE  Alter Table CompanyMaster Add BusyIntegration bit NOT NULL CONSTRAINT df_BusyIntegration DEFAULT '' "
    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'TallyIntegration') IS NULL PRINT 'NOT Exists' ELSE Update CompanyMaster Set BusyIntegration=0"
    If FYCode = "'" Then FYCode = "'" + "00" + Right(Date, 4) + "'"
    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'FYCode') IS NOT NULL PRINT 'Exists' ELSE Alter Table CompanyMaster Add FYCode nvarchar(6) NOT NULL CONSTRAINT df_FYCode DEFAULT ''  "
    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'FYCode') IS NULL PRINT 'NOT Exists' ELSE Update CompanyMaster Set FYCode= " & FYCode & " Where FYCode ='' OR FYCode IS NULL"
'Comp Child Update Table
    cnDatabase.Execute "IF COL_LENGTH('CompChild', 'VchName') IS NOT NULL PRINT 'Exists' ELSE  ALTER TABLE CompChild ADD VchName nvarchar(60) NOT NULL CONSTRAINT DF_CompChild_VchName DEFAULT ('')  "
    cnDatabase.Execute "IF COL_LENGTH('CompChild', 'VchName') IS NULL PRINT 'NOT Exists' ELSE  ALTER TABLE CompChild ALTER Column VchName nvarchar(60)"
    cnDatabase.Execute "IF COL_LENGTH('CompChild', 'VchName') IS NULL PRINT 'NOT Exists' ELSE  Update CompChild Set VchName=''  Where VchName IS NULL"
'DebitCreditChild Create Table
frmLicenceAgreement.Label2 = "DebitCreditChild Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('DebitCreditChild', 'Code') IS NOT NULL PRINT 'Exists' ELSE CREATE TABLE DebitCreditChild (Code nvarchar(6) NOT NULL,TOA nchar(1) NOT NULL,Ref nvarchar(6) NULL,BOM nvarchar(6) NULL,Account nvarchar(6) NOT NULL,Debit decimal(12, 2) NOT NULL,Credit decimal(12, 2) NOT NULL,ShortNarration nvarchar(100) NOT NULL,SrNo tinyint NOT NULL,RefCode nvarchar(6) NULL)  ON [PRIMARY]ALTER TABLE DebitCreditChild SET (LOCK_ESCALATION = TABLE)"
'DebitCreditParent Create Table
frmLicenceAgreement.Label2 = "DebitCreditOthInf Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('DebitCreditParent', 'Code') IS NOT NULL PRINT 'Exists' ELSE CREATE TABLE DebitCreditParent(Code nvarchar(6) NOT NULL,Name nvarchar(25) NOT NULL,Date datetime NULL,Debit decimal(12, 2) NOT NULL,Credit decimal(12, 2) NOT NULL,LongNarration nvarchar(100) NULL,Type nvarchar(6) NOT NULL,CreatedBy nvarchar(6) NOT NULL,CreatedOn datetime NOT NULL,ModifiedBy nvarchar(6) NULL,ModifiedOn datetime NULL,RecordStatus nvarchar(1) NULL,VchSeries nvarchar(6) NULL,AutoVchNo nvarchar(10) NULL,FYCode nvarchar(6) NOT NULL,Notes text NULL)  ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] ALTER TABLE DebitCreditParent ADD CONSTRAINT DF_DebitCreditParent_Debit DEFAULT ((0)) FOR Debit ALTER TABLE DebitCreditParent ADD CONSTRAINT DF_DebitCreditParent_Credit DEFAULT ((0)) FOR Credit ALTER TABLE DebitCreditParent ADD CONSTRAINT DF_DebitCreditParent_FYCode DEFAULT ('') FOR FYCode " & _
                                      "ALTER TABLE DebitCreditParent ADD CONSTRAINT PK_DebitCreditParent PRIMARY KEY CLUSTERED (Code) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ALTER TABLE DebitCreditParent SET (LOCK_ESCALATION = TABLE)"
'DebitCreditParent Update Table
frmLicenceAgreement.Label2 = "DebitCreditParent Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('DebitCreditParent', 'Notes') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE DebitCreditParent ADD Notes text NULL  ALTER TABLE DebitCreditParent SET (LOCK_ESCALATION = TABLE) "
    cnDatabase.Execute "IF COL_LENGTH('DebitCreditParent', 'Notes') IS NULL PRINT 'NOT Exists' ELSE Update DebitCreditParent Set Notes='' Where Notes IS NULL"
'DebitCreditRef Create Table
frmLicenceAgreement.Label2 = "DebitCreditRef Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('DebitCreditRef', 'RefCode') IS NOT NULL PRINT 'Exists' ELSE CREATE TABLE DebitCreditRef(RefCode nvarchar(6) NOT NULL,Method tinyint NOT NULL,VchType nvarchar(6) NOT NULL,VchCode nvarchar(6) NOT NULL,VchNo nvarchar(25) NULL,VchDate datetime NOT NULL,Account nvarchar(6) NOT NULL,Debit decimal(12, 2) NOT NULL,Credit decimal(12, 2) NOT NULL, TOA nchar(1) NOT NULL)  ON [PRIMARY] ALTER TABLE DebitCreditRef SET (LOCK_ESCALATION = TABLE)"
'DebitCreditOthInf Create Table
    cnDatabase.Execute "IF COL_LENGTH('DebitCreditOthInf', 'Code') IS NOT NULL PRINT 'Exists' ELSE CREATE TABLE DebitCreditOthInf(Code nvarchar(6) NOT NULL,BiltyNo nvarchar(30) NULL,BiltyDate datetime NULL,BiltyType nvarchar(30) NULL,Pkt smallint NOT NULL,Station nvarchar(30) NULL,Transport nvarchar(30) NULL,PktPicked bit NOT NULL)ON [PRIMARY]ALTER TABLE DebitCreditOthInf ADD CONSTRAINT  DF_DebitCreditOthInf_Pkt DEFAULT ((0)) FOR Pkt ALTER TABLE DebitCreditOthInf ADD CONSTRAINT DF_DebitCreditOthInf_PktPicked DEFAULT ((0)) FOR PktPicked ALTER TABLE DebitCreditOthInf SET (LOCK_ESCALATION = TABLE)"
'DiscountMaster Create Table
frmLicenceAgreement.Label2 = "DiscountMaster Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('DiscountMaster', 'Party') IS NOT NULL PRINT 'Exists' ELSE CREATE TABLE DiscountMaster (Party nvarchar(6) NOT NULL,ItemGroup nvarchar(6) NOT NULL,[Disc%] decimal(6, 2) NOT NULL,FYCode nvarchar(6) Not NULL )  ON [PRIMARY] ALTER TABLE DiscountMaster ADD CONSTRAINT [DF_DiscountMaster_Disc%] DEFAULT ((0)) FOR [Disc%] ALTER TABLE DiscountMaster ADD CONSTRAINT PK_DiscountMaster PRIMARY KEY CLUSTERED (Party,ItemGroup) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ALTER TABLE DiscountMaster SET (LOCK_ESCALATION = TABLE)"
'General Master Table Update Table
frmLicenceAgreement.Label2 = "GeneralMaster Update Going on!!! "
    cnDatabase.Execute "Alter Table GeneralMaster Alter Column Name nvarchar(60) NOT NULL"
    cnDatabase.Execute "Alter Table GeneralMaster Alter Column PrintName nvarchar(60) NOT NULL"
    cnDatabase.Execute "IF EXISTS(SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'GeneralMaster' AND COLUMN_NAME = 'UnderGroup') Print 'Col_Exist' ELSE ALTER TABLE GeneralMaster ADD  UnderGroup nvarchar(6) NULL  ALTER TABLE GeneralMaster SET (LOCK_ESCALATION = TABLE)"
'JobworkBVParent Update Table
frmLicenceAgreement.Label2 = "JobworkBVParent Update Going on!!! "
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'VchSeries') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD VchSeries nvarchar(6) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE)"
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'AutoVchNo') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD AutoVchNo nvarchar(10) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE)"
    
If MsgBox("Do You Wants to Update '21-05.7.1 Version' Also  (FYCode_ORDINAL_POSITION)!!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
    cnDatabase.Execute "IF (SELECT ORDINAL_POSITION FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'JobworkBVParent' AND COLUMN_NAME = 'FYCode') ='32' Print 'Col_POS_OK' Else ALTER TABLE JobworkBVParent ADD FYCodeD nvarchar(6) NOT NULL CONSTRAINT df_FYCode DEFAULT '' ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE)"
    cnDatabase.Execute "IF (SELECT ORDINAL_POSITION FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'JobworkBVParent' AND COLUMN_NAME = 'FYCode') ='32' Print 'Col_POS_OK' Else Update JobworkBVParent Set FYCodeD=FYCode"
    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
                                        "SET @table='JobworkBVParent' " & _
                                            "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'FYCODE','FYCODED') AND object_id = P.PARENT_OBJECT_ID)))  " & _
                                        "BEGIN  " & _
                                            "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'FYCODE','FYCODED') AND object_id = P.PARENT_OBJECT_ID)))  " & _
                                            "EXEC sp_executesql @sql  " & _
                                        "End"
    
    cnDatabase.Execute "IF (SELECT ORDINAL_POSITION FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'JobworkBVParent' AND COLUMN_NAME = 'FYCode') ='32' Print 'Col_POS_OK' Else ALTER TABLE JobworkBVParent Drop Column FYCode"
    cnDatabase.Execute "IF Exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'JobworkBVParent' AND COLUMN_NAME = 'FYCodeD') EXEC sp_rename 'JobworkBVParent.FYCodeD','FYCode','Column' Else Print 'Col_NOT_Exist'"
End If
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'Notes') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD Notes text NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'Notes') IS NULL PRINT 'NOT Exists' ELSE Update JobworkBVParent Set Notes='' Where Notes IS NULL"
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'SalesType') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD SalesType nvarchar(6) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'SalesType') IS NULL PRINT 'NOT Exists' ELSE Update JobworkBVParent Set SalesType='*01052' Where SalesType IS NULL"
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'GRDate') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD GRDate nvarchar(40) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'GRNo') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD GRNo nvarchar(40) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'Transport') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD Transport nvarchar(40) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'VehicleNo') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD VehicleNo datetime NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'Station') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD Station nvarchar(40) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVChild', 'BOM') IS NULL PRINT 'NOT Exists' ELSE ALTER TABLE JobworkBVChild ALTER COLUMN BOM NVARCHAR(18) NOT NULL"
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVChild', 'BOM') IS NULL PRINT 'NOT Exists' ELSE UPDATE JobworkBVChild SET BOM=LEFT(BOM,4)+'XXXXXXXXXXXX'+RIGHT(BOM,2) WHERE LEFT(BOM,2)='08' AND LEN(BOM)<18 "
    cnDatabase.Execute "IF COL_LENGTH('JobworkBVChild', 'BOM') IS NULL PRINT 'NOT Exists' ELSE UPDATE JobworkBVChild SET BOM=LEFT(BOM,4)+'XXXXXXXXXXXX'+RIGHT(BOM,2) WHERE LEFT(BOM,2)='05' AND LEN(BOM)<18 "
'PaperPOParent
    cnDatabase.Execute "IF COL_LENGTH('PaperPOParent', 'PicData') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE dbo.PaperPOParent ADD PicData varbinary(MAX) NULL"
    cnDatabase.Execute "IF COL_LENGTH('PaperPOParent', 'PicType') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE dbo.PaperPOParent ADD PicType nvarchar(4) NULL"
    
        'Default Master
    'Genral Master
'Size Master_Type-1
frmLicenceAgreement.Label2 = "Size Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01001' OR Name='05.25X10.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01001','05.25X10.00','05.25X10.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01002' OR Name='10.00X29.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01002','10.00X29.00','10.00X29.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01003' OR Name='11.00X14.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01003','11.00X14.00','11.00X14.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01004' OR Name='11.50X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01004','11.50X18.00','11.50X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01005' OR Name='12.00X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01005','12.00X18.00','12.00X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01006' OR Name='12.00X23.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01006','12.00X23.00','12.00X23.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01007' OR Name='12.50X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01007','12.50X18.00','12.50X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01008' OR Name='13.00X19.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01008','13.00X19.00','13.00X19.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01009' OR Name='14.00X19.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01009','14.00X19.00','14.00X19.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01010' OR Name='14.00X22.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01010','14.00X22.00','14.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01011' OR Name='15.00X10.00 (CARD)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01011','15.00X10.00 (CARD)','15.00X10.00 (CARD)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01012' OR Name='15.00X20.00 (CARD)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01012','15.00X20.00 (CARD)','15.00X20.00 (CARD)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01013' OR Name='15.00X21.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01013','15.00X21.00','15.00X21.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01014' OR Name='15.00X27.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01014','15.00X27.50','15.00X27.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01015' OR Name='15.50X20.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01015','15.50X20.00','15.50X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01016' OR Name='15.50X20.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01016','15.50X20.50','15.50X20.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017' OR Name='15.50X21.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01017','15.50X21.00','15.50X21.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01018' OR Name='15.50X21.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01018','15.50X21.50','15.50X21.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01019' OR Name='16.00X20.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01019','16.00X20.00','16.00X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01020' OR Name='16.00X20.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01020','16.00X20.50','16.00X20.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01021' OR Name='16.00X22.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01021','16.00X22.00','16.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01022' OR Name='16.00X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01022','16.00X24.00','16.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01023' OR Name='16.00X25.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01023','16.00X25.00','16.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01024' OR Name='16.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01024','16.00X30.00','16.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01025' OR Name='16.50X10.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01025','16.50X10.50','16.50X10.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01026' OR Name='17.00X22.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01026','17.00X22.00','17.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01027' OR Name='17.00X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01027','17.00X24.00','17.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028' OR Name='18.00X23.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01028','18.00X23.00','18.00X23.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01029' OR Name='18.00X23.00 (Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01029','18.00X23.00 (Card)','18.00X23.00 (Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01030' OR Name='18.00X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01030','18.00X24.00','18.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031' OR Name='18.00X25.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01031','18.00X25.00','18.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01032' OR Name='19.00X20.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01032','19.00X20.00','19.00X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01033' OR Name='19.00X25.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01033','19.00X25.00','19.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01034' OR Name='19.00X38.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01034','19.00X38.00','19.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01035' OR Name='20.00X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01035','20.00X24.00','20.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01036' OR Name='20.00X25.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01036','20.00X25.00','20.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01037' OR Name='20.00X26.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01037','20.00X26.00','20.00X26.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01038' OR Name='20.00X28.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01038','20.00X28.00','20.00X28.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039' OR Name='20.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01039','20.00X30.00','20.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01040' OR Name='20.00X30.00(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01040','20.00X30.00(A/P)','20.00X30.00(A/P)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01041' OR Name='20.00X30.00(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01041','20.00X30.00(Card)','20.00X30.00(Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01042' OR Name='20.00X31.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01042','20.00X31.00','20.00X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01043' OR Name='20.50X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01043','20.50X24.00','20.50X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01044' OR Name='20.50X31.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01044','20.50X31.00','20.50X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01045' OR Name='21.00X29.70 (A4)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01045','21.00X29.70 (A4)','21.00X29.70 (A4)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01046' OR Name='21.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01046','21.00X30.00','21.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01047' OR Name='21.00X30.00(CARD)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01047','21.00X30.00(CARD)','21.00X30.00(CARD)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01048' OR Name='21.00X31.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01048','21.00X31.00','21.00X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01049' OR Name='21.00X32.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01049','21.00X32.00','21.00X32.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01050' OR Name='21.00X33.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01050','21.00X33.00','21.00X33.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01051' OR Name='21.00X34.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01051','21.00X34.00','21.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01052' OR Name='21.00X35.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01052','21.00X35.00','21.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01053' OR Name='21.50X28.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01053','21.50X28.50','21.50X28.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01054' OR Name='22.00X28.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01054','22.00X28.00','22.00X28.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055' OR Name='22.00X32.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01055','22.00X32.00','22.00X32.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01056' OR Name='22.00X34.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01056','22.00X34.00','22.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01057' OR Name='23.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01057','23.00X30.00','23.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01058' OR Name='23.00X33.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01058','23.00X33.00','23.00X33.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01059' OR Name='23.00X35.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01059','23.00X35.00','23.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060' OR Name='23.00X36.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01060','23.00X36.00','23.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01061' OR Name='23.00X36.00(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01061','23.00X36.00(A/P)','23.00X36.00(A/P)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01062' OR Name='23.00X36.00(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01062','23.00X36.00(Card)','23.00X36.00(Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01063' OR Name='24.00X34.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01063','24.00X34.00','24.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01064' OR Name='24.00X36.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01064','24.00X36.00','24.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01065' OR Name='24.13X24.13') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01065','24.13X24.13','24.13X24.13','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01066' OR Name='25.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01066','25.00X30.00','25.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01067' OR Name='25.00X36.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01067','25.00X36.00','25.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068' OR Name='25.00X38.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01068','25.00X38.00','25.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01069' OR Name='26.00X38.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01069','26.00X38.00','26.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01070' OR Name='26.00X40.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01070','26.00X40.00','26.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01071' OR Name='28.00X35.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01071','28.00X35.00','28.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01072' OR Name='28.00X40.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01072','28.00X40.00','28.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01073' OR Name='30.00X40.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01073','30.00X40.00','30.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01074' OR Name='31.50X41.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01074','31.50X41.50','31.50X41.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Item Group Master_TYPE-5
frmLicenceAgreement.Label2 = "Item Group Master Update Going on!!! "
    If Trim(ReadFromFile("Client ID")) = "Publisher" Then
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05001' OR Name='Activity Book') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05001','Activity Book','Activity Book','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05002' OR Name='Box') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05002','Box','Box','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05003' OR Name='CARD') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05003','CARD','CARD','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05004' OR Name='CATALOGUE') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05004','CATALOGUE','CATALOGUE','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05005' OR Name='General') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05005','General','General','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05006' OR Name='GRADE 1') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05006','GRADE 1','GRADE 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05007' OR Name='GRADE 2') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05007','GRADE 2','GRADE 2','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05008' OR Name='GRADE 3') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05008','GRADE 3','GRADE 3','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05009' OR Name='GRADE 4') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05009','GRADE 4','GRADE 4','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05010' OR Name='GRADE 5') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05010','GRADE 5','GRADE 5','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05011' OR Name='JUNIOR') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05011','JUNIOR','JUNIOR','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05012' OR Name='LEVEL 1') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05012','LEVEL 1','LEVEL 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05013' OR Name='LEVEL 2') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05013','LEVEL 2','LEVEL 2','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05014' OR Name='LEVEL 3') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05014','LEVEL 3','LEVEL 3','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05015' OR Name='LEVEL 4') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05015','LEVEL 4','LEVEL 4','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05016' OR Name='LEVEL 5') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05016','LEVEL 5','LEVEL 5','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05017' OR Name='LEVEL A') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05017','LEVEL A','LEVEL A','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05018' OR Name='LEVEL B') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05018','LEVEL B','LEVEL B','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05019' OR Name='LEVEL C') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05019','LEVEL C','LEVEL C','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05020' OR Name='NURSERY') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05020','NURSERY','NURSERY','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05021' OR Name='SECONDARY STD VI') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05021','SECONDARY STD VI','SECONDARY STD','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05022' OR Name='SENIOR') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05022','SENIOR','SENIOR','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05023' OR Name='SET 1') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05023','SET 1','SET 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
End If
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05024' OR Name='Item Group') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05024','Item Group','Item Group','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Binding Type_Type-6
frmLicenceAgreement.Label2 = "Binding Type Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06001' OR Name='Die_Cutting') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06001','Die_Cutting','Die_Cutting','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06002' OR Name='Die_Perforation') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06002','Die_Perforation','Die_Perforation','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06003' OR Name='Hard Bound') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06003','Hard Bound','Hard Bound','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06004' OR Name='Perfect Binding With Sewing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06004','Perfect Binding With Sewing','Perfect Binding With Sewing','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06005' OR Name='Perfect Binding With Sewing(CD-Insert)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06005','Perfect Binding With Sewing(CD-Insert)','Perfect Binding With Sewing(CD-Insert)','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06006' OR Name='Spiral Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06006','Spiral Binding','Spiral Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06007' OR Name='Wirro Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06007','Wirro Binding','Wirro Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06008' OR Name='Cutting & Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06008','Cutting & Packing','Cutting & Packing','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06009' OR Name='Cutting Only') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06009','Cutting Only','Cutting Only','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06010' OR Name='Half Die Cut') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06010','Half Die Cut','Half Die Cut','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06011' OR Name='Loose') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06011','Loose','Loose','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06012' OR Name='Pad Gumming') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06012','Pad Gumming','Pad Gumming','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06013' OR Name='Pakki Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06013','Pakki Binding','Pakki Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06014' OR Name='Kachchi Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06014','Kachchi Binding','Kachchi Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06015' OR Name='Center Pinning (BOX)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06015','Center Pinning (BOX)','Center Pinning (BOX)','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06016' OR Name='Center Pin Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06016','Center Pin Binding','Center Pin Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06017' OR Name='None') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06017','None','None','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06018' OR Name='Perfect Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06018','Perfect Binding','Perfect Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Finishing Type_Type-7
frmLicenceAgreement.Label2 = "Finish Size Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07001' OR Name='BOPP Gloss') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07001','BOPP Gloss','BOPP Gloss','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07002' OR Name='BOPP Matt') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07002','BOPP Matt','BOPP Matt','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07003' OR Name='Box Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07003','Box Packing','Box Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07004' OR Name='Center Pin Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07004','Center Pin Binding','Center Pin Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07005' OR Name='Counting & Fabrication') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07005','Counting & Fabrication','Counting & Fabrication','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07006' OR Name='Creasing+Folding+Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07006','Creasing+Folding+Packing','Creasing+Folding+Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07007' OR Name='Cutting and Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07007','Cutting and Packing','Cutting and Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07008' OR Name='Cutting Leaflet Only') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07008','Cutting Leaflet Only','Cutting Leaflet Only','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07009' OR Name='Die Cutting Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07009','Die Cutting Charges','Die Cutting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07010' OR Name='Die Making Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07010','Die Making Charges','Die Making Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07011' OR Name='Digital Print') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07011','Digital Print','Digital Print','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07012' OR Name='Embossing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07012','Embossing','Embossing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07013' OR Name='Foiling Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07013','Foiling Charges','Foiling Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07014' OR Name='Folding & Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07014','Folding & Packing','Folding & Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07015' OR Name='Graning') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07015','Graning','Graning','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07016' OR Name='Half Die Cutting Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07016','Half Die Cutting Charges','Half Die Cutting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07017' OR Name='Hardbound Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07017','Hardbound Binding','Hardbound Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07018' OR Name='Hologram') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07018','Hologram','Hologram','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07019' OR Name='Matt + Spot UV') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07019','Matt + Spot UV','Matt + Spot UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07020' OR Name='Matt + Spot UV + Foiling + Embossing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07020','Matt + Spot UV + Foiling + Embossing','Matt + Spot UV + Foiling + Embossing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07021' OR Name='Matt + Spot UV+Glitter UV') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07021','Matt + Spot UV+Glitter UV','Matt + Spot UV+Glitter UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07022' OR Name='Matt Both Side') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07022','Matt Both Side','Matt Both Side','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07023' OR Name='MINI Offset JOB') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07023','MINI Offset JOB','MINI Offset JOB','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07024' OR Name='None') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07024','None','None','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07025' OR Name='Packing Shrink') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07025','Packing Shrink','Packing Shrink','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07026' OR Name='Paper Cost') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07026','Paper Cost','Paper Cost','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07027' OR Name='Pasting Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07027','Pasting Charges','Pasting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07028' OR Name='Perfect Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07028','Perfect Binding','Perfect Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07029' OR Name='Plate') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07029','Plate','Plate','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07030' OR Name='Printing 4 Col') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07030','Printing 4 Col','Printing 4 Col','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07031' OR Name='PVC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07031','PVC','PVC','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07032' OR Name='Spot UV') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07032','Spot UV','Spot UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07033' OR Name='Thermal Matt') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07033','Thermal Matt','Thermal Matt','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07034' OR Name='UV Hybraid') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07034','UV Hybraid','UV Hybraid','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07035' OR Name='Varnising') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07035','Varnising','Varnising','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07036' OR Name='BP-Unit Cost') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07036','BP-Unit Cost','BP-Unit Cost','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07037' OR Name='BP-Stitching') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07037','BP-Stitching','BP-Stitching','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07038' OR Name='BP-Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07038','BP-Binding','BP-Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07039' OR Name='BP-Folding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07039','BP-Folding','BP-Folding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07040' OR Name='BP-Shrink Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07040','BP-Shrink Packing','BP-Shrink Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07041' OR Name='BP-Box Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07041','BP-Box Packing','BP-Box Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07042' OR Name='BP-Cartage') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07042','BP-Cartage','BP-Cartage','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07043' OR Name='Digital Print_1C') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07043','Digital Print_1C','Digital Print_1C','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07044' OR Name='Digital Print_2C') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07044','Digital Print_2C','Digital Print_2C','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07045' OR Name='Digital Print_4C') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07045','Digital Print_4C','Digital Print_4C','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"

'Project Member/Editorial Team Master_Type-8
frmLicenceAgreement.Label2 = "Project Member/ Editorial Team Master Update Going on!!! "
If Trim(ReadFromFile("Client ID")) = "Publisher" Then
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08002' OR Name='Author_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08002','Author_ABC','Author_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08003' OR Name='DTP_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08003','DTP_ABC','DTP_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08005' OR Name='Editor_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08005','Editor_ABC','Editor_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08007' OR Name='Graphic_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08007','Graphic_ABC','Graphic_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08008' OR Name='PPQ_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08008','PPQ_ABC','PPQ_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08009' OR Name='Processing_S.R.K') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08009','Processing_S.R.K','Processing_S.R.K','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08010' OR Name='Proof Reader_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08010','Proof Reader_ABC','Proof Reader_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08011' OR Name='Type Setting_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08011','Type Setting_ABC','Type Setting_Sanjay Khanna','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
End If

'Plate Master_Type-9
frmLicenceAgreement.Label2 = "Plate Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*09001' OR Name='CTP_Plates') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*09001','CTP_Plates','CTP_Plates','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*09002' OR Name='Nagative-Cut Pieces') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*09002','Nagative-Cut Pieces','Nagative-Cut Pieces','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*09003' OR Name='Nagative-One Pieces') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*09003','Nagative-One Pieces','Nagative-One Pieces','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Size Group Master-10
frmLicenceAgreement.Label2 = "Size Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10022' OR Name='12.00X18.00-Digital') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10022','12.00X18.00-Digital','12.00X18.00-Digital','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10018' OR Name='Extra Large-28''''X40''''-(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10018','Extra Large-28''''X40''''-(Card)','Extra Large-28''''X40''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10001' OR Name='Extra Large-28''''X40''''-A/P') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10001','Extra Large-28''''X40''''-A/P','Extra Large-28''''X40''''-A/P','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10002' OR Name='Extra Large-28''''X40''''-A/P_SPL') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10002','Extra Large-28''''X40''''-A/P_SPL','Extra Large-28''''X40''''-A/P_SPL','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10003' OR Name='Extra Large-30''''X40''''') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10003','Extra Large-30''''X40''''','Extra Large-30''''X40''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10004' OR Name='Extra Large-30''''X40''''-(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10004','Extra Large-30''''X40''''-(A/P)','Extra Large-30''''X40''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10005' OR Name='Extra Large-30''''X40''''-(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10005','Extra Large-30''''X40''''-(Card)','Extra Large-30''''X40''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10006' OR Name='LARGE-23''''X36''''') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10006','LARGE-23''''X36''''','LARGE-23''''X36''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10007' OR Name='LARGE-23''''X36''''-(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10007','LARGE-23''''X36''''-(A/P)','LARGE-23''''X36''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10008' OR Name='LARGE-23''''X36''''-(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10008','LARGE-23''''X36''''-(Card)','LARGE-23''''X36''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10019' OR Name='Little-11.50''''X18.00''''') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10019','Little-11.50''''X18.00''''','Little-11.50''''X18.00''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10021' OR Name='Little-11.50''''X18.00''''-(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10021','Little-11.50''''X18.00''''-(A/P)','Little-11.50''''X18.00''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10020' OR Name='Little-11.50''''X18.00''''-(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10020','Little-11.50''''X18.00''''-(Card)','Little-11.50''''X18.00''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10009' OR Name='Medium-20''''X30''''') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10009','Medium-20''''X30''''','Medium-20''''X30''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10010' OR Name='Medium-20''''X30''''(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10010','Medium-20''''X30''''(A/P)','Medium-20''''X30''''(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10011' OR Name='Medium-20''''X30''''(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10011','Medium-20''''X30''''(Card)','Medium-20''''X30''''(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10012' OR Name='Small-19''''X26''''') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10012','Small-19''''X26''''','Small-19''''X26''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10017' OR Name='Small-19''''X26''''-(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10017','Small-19''''X26''''-(A/P)','Small-19''''X26''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10013' OR Name='Small-19''''X26''''(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10013','Small-19''''X26''''(Card)','Small-19''''X26''''(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10014' OR Name='Web-508mm') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10014','Web-508mm','Web-508mm','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10015' OR Name='Web-578mm') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10015','Web-578mm','Web-578mm','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Finish Size Master_TYPE-11
frmLicenceAgreement.Label2 = "Finish Size Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11001' OR Name='05.25x10.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11001','05.25x10.00','05.25x10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11002' OR Name='12.00X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11002','12.00X18.00','12.00X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11003' OR Name='12.00X23.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11003','12.00X23.00','12.00X23.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11004' OR Name='14.00X19.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11004','14.00X19.00','14.00X19.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11005' OR Name='15.00X10.00 (CARD)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11005','15.00X10.00 (CARD)','15.00X10.00 (CARD)','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11006' OR Name='15.50X20.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11006','15.50X20.50','15.50X20.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11007' OR Name='16.00x20.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11007','16.00x20.00','16.00x20.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11008' OR Name='16.00X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11008','16.00X24.00','16.00X24.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11009' OR Name='16.50X10.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11009','16.50X10.50','16.50X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11010' OR Name='17.00X22.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11010','17.00X22.00','17.00X22.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11011' OR Name='04.00X06.87') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11011','04.00X06.87','04.00X06.87','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11012' OR Name='04.25X05.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11012','04.25X05.50','04.25X05.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11013' OR Name='04.25X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11013','04.25X07.00','04.25X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11014' OR Name='04.37X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11014','04.37X07.00','04.37X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11015' OR Name='04.72X07.48') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11015','04.72X07.48','04.72X07.48','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11016' OR Name='05.00X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11016','05.00X07.00','05.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11017' OR Name='05.00X08.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11017','05.00X08.00','05.00X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11018' OR Name='05.06X07.81') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11018','05.06X07.81','05.06X07.81','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11019' OR Name='05.25X08.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11019','05.25X08.00','05.25X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11020' OR Name='05.50X08.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11020','05.50X08.50','05.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11021' OR Name='05.83X08.27') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11021','05.83X08.27','05.83X08.27','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11022' OR Name='06.00X08.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11022','06.00X08.25','06.00X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11023' OR Name='06.00X08.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11023','06.00X08.50','06.00X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11024' OR Name='06.00X09.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11024','06.00X09.00','06.00X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11025' OR Name='06.14X09.21') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11025','06.14X09.21','06.14X09.21','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11026' OR Name='06.25X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11026','06.25X09.50','06.25X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11027' OR Name='06.63X10.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11027','06.63X10.25','06.63X10.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11028' OR Name='06.69X09.61') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11028','06.69X09.61','06.69X09.61','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11029' OR Name='06.75X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11029','06.75X09.50','06.75X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11030' OR Name='07.00X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11030','07.00X07.00','07.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11031' OR Name='07.00X09.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11031','07.00X09.00','07.00X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11032' OR Name='07.00X10.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11032','07.00X10.00','07.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11033' OR Name='07.25X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11033','07.25X09.50','07.25X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11034' OR Name='07.44X09.69') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11034','07.44X09.69','07.44X09.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11035' OR Name='07.50X07.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11035','07.50X07.50','07.50X07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11036' OR Name='07.50X09.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11036','07.50X09.25','07.50X09.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11037' OR Name='07.50X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11037','07.50X09.50','07.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11038' OR Name='07.75X10.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11038','07.75X10.50','07.75X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11039' OR Name='08.00X08.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11039','08.00X08.00','08.00X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11040' OR Name='08.00X10.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11040','08.00X10.00','08.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11041' OR Name='08.00X10.88') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11041','08.00X10.88','08.00X10.88','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11042' OR Name='08.00X11.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11042','08.00X11.25','08.00X11.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11043' OR Name='08.25X08.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11043','08.25X08.25','08.25X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11044' OR Name='08.25X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11044','08.25X11.00','08.25X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11045' OR Name='08.27X11.69') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11045','08.27X11.69','08.27X11.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11046' OR Name='08.50X08.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11046','08.50X08.50','08.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11047' OR Name='08.50X09.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11047','08.50X09.00','08.50X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11048' OR Name='08.50X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11048','08.50X11.00','08.50X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11049' OR Name='09.00X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11049','09.00X07.00','09.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11050' OR Name='09.00X12.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11050','09.00X12.00','09.00X12.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11051' OR Name='10.00X10.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11051','10.00X10.00','10.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11052' OR Name='11.00X13.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11052','11.00X13.00','11.00X13.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11053' OR Name='11.00X17.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11053','11.00X17.00','11.00X17.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11054' OR Name='11.00X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11054','11.00X18.00','11.00X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11055' OR Name='12.00X12.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11055','12.00X12.00','12.00X12.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11056' OR Name='18.00X23.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11056','18.00X23.00','18.00X23.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11057' OR Name='07.75X11.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11057','07.75X11.25','07.75X11.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11058' OR Name='08.00X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11058','08.00X11.00','08.00X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11059' OR Name='04.50X01.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11059','04.50X01.75','04.50X01.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11060' OR Name='11.00x15.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11060','11.00x15.75','11.00x15.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11061' OR Name='11.00X16.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11061','11.00X16.00','11.00X16.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11062' OR Name='08.25X11.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11062','08.25X11.75','08.25X11.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11063' OR Name='04.00X06.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11063','04.00X06.00','04.00X06.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11064' OR Name='20.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11064','20.00X30.00','20.00X30.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11065' OR Name='17.50X22.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11065','17.50X22.50','17.50X22.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11066' OR Name='11.50X08.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11066','11.50X08.00','11.50X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11067' OR Name='21.00X31.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11067','21.00X31.00','21.00X31.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11068' OR Name='05.30X08.30') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11068','05.30X08.30','05.30X08.30','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11069' OR Name='11.50X10.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11069','11.50X10.75','11.50X10.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11070' OR Name='08.50X10.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11070','08.50X10.75','08.50X10.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11071' OR Name='02.00X03.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11071','02.00X03.00','02.00X03.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11072' OR Name='11.50X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11072','11.50X07.00','11.50X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11073' OR Name='05.50X19.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11073','05.50X19.00','05.50X19.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11074' OR Name='10.25X07.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11074','10.25X07.50','10.25X07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11075' OR Name='07.50X13.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11075','07.50X13.75','07.50X13.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11076' OR Name='07.00X02.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11076','07.00X02.50','07.00X02.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11077' OR Name='06.50X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11077','06.50X09.50','06.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11078' OR Name='04.00x07.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11078','04.00x07.50','04.00x07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11079' OR Name='23.00X36.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11079','23.00X36.00','23.00X36.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11080' OR Name='15.00X20.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11080','15.00X20.00','15.00X20.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11081' OR Name='25.00X36.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11081','25.00X36.00','25.00X36.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11082' OR Name='09.00X14.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11082','09.00X14.00','09.00X14.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11083' OR Name='05.25X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11083','05.25X07.00','05.25X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11084' OR Name='08.00X10.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11084','08.00X10.50','08.00X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11085' OR Name='07.50X08.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11085','07.50X08.50','07.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11086' OR Name='03.25X04.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11086','03.25X04.75','03.25X04.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11087' OR Name='09.75X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11087','09.75X11.00','09.75X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11088' OR Name='13.50X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11088','13.50X18.00','13.50X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11089' OR Name='07.62X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11089','07.62X11.00','07.62X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11090' OR Name='07.36X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11090','07.36X11.00','07.36X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11091' OR Name='08.26X11.69') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11091','08.26X11.69','08.26X11.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11092' OR Name='09.50X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11092','09.50X09.50','09.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11093' OR Name='11.69X05.20') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11093','11.69X05.20','11.69X05.20','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11094' OR Name='05.75X08.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11094','05.75X08.25','05.75X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11095' OR Name='21.00X29.70') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11095','21.00X29.70','21.00X29.70','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Genral Accounts Groups_TYPE-12
frmLicenceAgreement.Label2 = "Account Group Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12002' OR Name='Account Group') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12002','Account Group','Account Group','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*99996' OR Name='Transporter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*99996','Transporter','Transporter','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*99997' OR Name='Packer') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*99997','Packer','Transporter','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*99998' OR Name='Deliverer') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*99998','Deliverer','Deliverer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*99999' OR Name='Material Centre') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*99999','Material Centre','Material Centre','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*99995' OR Name='Sales Executive') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*99995','Sales Executive','Sales Executive','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12001' OR Name='Binders') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12001','Binders','Binders','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12003' OR Name='Box Supplier') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12003','Box Supplier','Box Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12004' OR Name='CD Suppliers') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12004','CD Suppliers','CD Suppliers','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12005' OR Name='FG Godown') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12005','FG Godown','FG Godown','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12006' OR Name='Laminator') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12006','Laminator','Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12007' OR Name='Packaging Supplier') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12007','Packaging Supplier','Packaging Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12008' OR Name='Paper Suppliers') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12008','Paper Suppliers','Paper Suppliers','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12009' OR Name='Printer') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12009','Printer','Printer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12010' OR Name='Printer & Binder') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12010','Printer & Binder','Printer & Binder','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12011' OR Name='Printer, Binder & Laminator') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12011','Printer, Binder & Laminator','Printer, Binder & Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12012' OR Name='Processor & Printer') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12012','Processor & Printer','Processor & Printer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12013' OR Name='Processor, Printer & Laminator') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12013','Processor, Printer & Laminator','Processor, Printer & Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12014' OR Name='UFG Godown') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12014','UFG Godown','UFG Godown','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12015' OR Name='Publisher') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12015','Publisher','Publisher','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12016' OR Name='Clients') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12016','Clients','Clients','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12017' OR Name='Cons. Supplier') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12017','Cons. Supplier','Cons. Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12018' OR Name='Plate Maker') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12018','Plate Maker','Plate Maker','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'Departments_TYPE-13
frmLicenceAgreement.Label2 = "Departments Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13001' OR Name='Editorial Department') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13001','Editorial Department','Editorial Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13002' OR Name='Production Department') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13002','Production Department','Production Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13003' OR Name='Sales Department') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13003','Sales Department','Sales Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13004' OR Name='Contracts Department and Legal Department') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13004','Contracts Department and Legal Department','Contracts Department and Legal Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13005' OR Name='Managing Editorial and Production') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13005','Managing Editorial and Production','Managing Editorial and Production','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13006' OR Name='Creative Departments') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13006','Creative Departments','Creative Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13007' OR Name='Subsidiary Rights Departments') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13007','Subsidiary Rights Departments','Subsidiary Rights Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13008' OR Name='Marketing, Promotion, and Advertising Departments') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13008','Marketing, Promotion, and Advertising Departments','Marketing, Promotion, and Advertising Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13009' OR Name='Publicity Department') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13009','Publicity Department','Publicity Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13010' OR Name='Publisher Website Maintenance') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13010','Publisher Website Maintenance','Publisher Website Maintenance','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13011' OR Name='Finance and Accounting') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13011','Finance and Accounting','Finance and Accounting','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13012' OR Name='Information Technology (IT)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13012','Information Technology (IT)','Information Technology (IT)','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13013' OR Name='Human Resources (HR)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13013','Human Resources (HR)','Human Resources (HR)','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Designation_TYPE-14
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14001' OR Name='Editor-in-Chief') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14001','Editor-in-Chief','Editor-in-Chief','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14002' OR Name='Managing editor') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14002','Managing editor','Managing editor','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14003' OR Name='Editors') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14003','Editors','Editors','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14004' OR Name='Author/Writers') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14004','Author/Writers','Author/Writers','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14005' OR Name='Fact-checkers') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14005','Fact-checkers','Fact-checkers','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14006' OR Name='Graphic Designer') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14006','Graphic Designer','Graphic Designer','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14007' OR Name='Production manager') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14007','Production manager','Production manager','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14008' OR Name='DTP-Operator') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14008','DTP-Operator','DTP-Operator','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14009' OR Name='Proof Reader') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14009','Proof Reader','Proof Reader','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Paper Unit Master_TYPE-15
frmLicenceAgreement.Label2 = "Paper Units Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15001' OR Name='Gross') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15001','Gross','Gross','15','144','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' OR Name='Packet(100)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15002','Packet(100)','Packet(100)','15','100','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15003' OR Name='Packet(150)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15003','Packet(150)','Packet(150)','15','150','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' OR Name='Ream') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15004','Ream','Ream','15','500','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15005' OR Name='Reel') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15005','Reel','Reel','15','500','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15006' OR Name='Bundle (700)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15006','Bundle (700)','Bundle (700)','15','700','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15007' OR Name='Packet(200)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15007','Packet(200)','Packet(200)','15','200','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15008' OR Name='PACKET') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15008','PACKET','PACKET','15','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15009' OR Name='Sheet') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15009','Sheet','Sheet','15','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15010' OR Name='Packet (250)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15010','Packet (250)','Packet (250)','15','250','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Paper Quality Master_TYPE-16
frmLicenceAgreement.Label2 = "Paper Quality Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16001' OR Name='Coated Matt') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16001','Coated Matt','Coated Matt','16','0.95','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002' OR Name='Coated Gloss') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16002','Coated Gloss','Coated Gloss','16','0.9','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003' OR Name='Uncoated') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16003','Uncoated','Uncoated','16','1.35','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16004' OR Name='High Bulk') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16004','High Bulk','High Bulk','16','1.4','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Narration Master_TYPE-17
 frmLicenceAgreement.Label2 = "Narration Master Update Going on!!! "
  cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17001' OR Name='1. Printing & Finishing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17001','1. Printing & Finishing Charges of','Printing & Finishing Charges of','17','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17002' OR Name='1. Text Printing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17002','1. Text Printing Charges of','Text Printing Charges of','17','2','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17003' OR Name='2. Title Printing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17003','2. Title Printing Charges of','Title Printing Charges of','17','3','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17004' OR Name='3. Combo Title Printing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17004','3. Combo Title Printing Charges of','Combo Title Printing Charges of','17','4','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17005' OR Name='4. Finishing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17005','4. Finishing Charges of','Finishing Charges of','17','5','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17006' OR Name='5. Binding Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17006','5. Binding Charges of','Binding Charges of','17','6','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17007' OR Name='7. Title Printing & Finishing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17007','7. Title Printing & Finishing Charges of','Title Printing & Finishing Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17008' OR Name='6. Text Printing & Finishing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17008','6. Text Printing & Finishing Charges of','Text Printing & Finishing Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17009' OR Name='8. Unit Cost Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17009','8. Unit Cost Charges of','Unit Cost Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17010' OR Name='9. Unit Cost') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17010','9. Unit Cost','.','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17011' OR Name='10 Lamination Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17011','10 Lamination Charges','Lamination Charges','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17012' OR Name='11 Printed Book') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17012','11 Printed Book','Printed Book','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'HSN MASTER_TYPE-18
 frmLicenceAgreement.Label2 = "HSN Master Update Going on!!! "
  cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*18001' OR Name='998812') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*18001','998812','998812','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*18002' OR Name='998912') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*18002','998912','998912','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*18003' OR Name='4901') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*18003','4901','4901','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*18004' OR Name='49011010') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*18004','49011010','49011010','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Elements MASTER_TYPE-19
        'eLEMENT mASTER mOVED TO eLEMENT mASTER
'Calculation Units MASTER_Type-20
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20001' OR Name='Per Unit') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20001','Per Unit','Per Unit','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20002' OR Name='Per Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20002','Per Inch²','Per Inch²','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20003' OR Name='100 Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20003','100 Inch²','100 Inch²','20','100','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20004' OR Name='1000 Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20004','1000 Inch²','1000 Inch²','20','1000','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20005' OR Name='Per 1000') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20005','Per 1000','Per 1000','20','1000','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20006' OR Name='Per Packet') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20006','Per Packet','Per Packet','20','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20007' OR Name='Per Page') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20007','Per Page','Per Page','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20008' OR Name='Per Paisa Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20008','Per Paisa Inch²','Per Paisa Inch²','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20009' OR Name='Per Box') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20009','Per Box','Per Box','20','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20010' OR Name='Per Bundle') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20010','Per Bundle','Per Bundle','20','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"

'Machine Master_TYPE-21
 frmLicenceAgreement.Label2 = "Machine Master Update Going on!!! "
  cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*21046' OR Name='Machine To Be Decide') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*21046','Machine To Be Decide','Machine To Be Decide','21','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*21047' OR Name='RYOBI - 4 Col') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*21047','RYOBI - 4 Col','RYOBI - 4 Col','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*21048' OR Name='SM 102 28x40') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*21048','SM 102 28x40','SM 102 28x40','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*21049' OR Name='SM 74 20x29') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*21049','SM 74 20x29','SM 74 20x29','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*21050' OR Name='Heidel 2 Col') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*21050','Heidel 2 Col','Heidel 2 Col','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'?_TYPE-22
'?_TYPE-23
'?_TYPE-24
'General  Unit MasterTYPE-25
 frmLicenceAgreement.Label2 = "Unit Master Update Going on!!! "
  cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25001' OR Name='Kilogram') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25001','Kilogram','kg.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25002' OR Name='Gram') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25002','Gram','gm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25003' OR Name='Milligram') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25003','Milligram','mg.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25004' OR Name='Liter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25004','Liter','ltr.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25005' OR Name='Milliliter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25005','Milliliter','ml.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25006' OR Name='Feet') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25006','Feet','ft.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25007' OR Name='Inch') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25007','Inch','in.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25008' OR Name='Meter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25008','Meter','mtr.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25009' OR Name='Centimeter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25009','Centimeter','cm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25010' OR Name='Millimeter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25010','Millimeter','mm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25011' OR Name='Piece') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25011','Piece','pec.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25012' OR Name='Bags') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25012','Bags','bags','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25013' OR Name='Roll') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25013','Roll','roll','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25014' OR Name='Sets') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25014','Sets','sets','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25015' OR Name='Packets') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25015','Packets','packets','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25016' OR Name='Gross') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25016','Gross','gross','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25017' OR Name='Dozen') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25017','Dozen','dozen','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25018' OR Name='Tonn') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25018','Tonn','tonn','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Account Group_TYPE-26
 frmLicenceAgreement.Label2 = "Account Group MAster Update Going on!!! "
  cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26001' OR Name='Profit & Loss') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26001','Profit & Loss','Profit & Loss','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26002' OR Name='Revenue Accounts') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26002','Revenue Accounts','Revenue Accounts','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26003' OR Name='Stock-in-hand') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26003','Stock-in-hand','Stock-in-hand','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26004' OR Name='Bank Accounts') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26004','Bank Accounts','Bank Accounts','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26005' OR Name='Bank O/D Account') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26005','Bank O/D Account','Bank O/D Account','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26006' OR Name='Capital Account') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26006','Capital Account','Capital Account','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26007' OR Name='Cash-in-hand') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26007','Cash-in-hand','Cash-in-hand','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26008' OR Name='Current Assets') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26008','Current Assets','Current Assets','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26009' OR Name='Current Liabilities') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26009','Current Liabilities','Current Liabilities','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26010' OR Name='Depreciation Res On Machine') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26010','Depreciation Res On Machine','Depreciation Res On Machine','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26016')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26011' OR Name='Duties & Taxes') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26011','Duties & Taxes','Duties & Taxes','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26012' OR Name='Expenses (Direct/Mfg.)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26012','Expenses (Direct/Mfg.)','Expenses (Direct/Mfg.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26013' OR Name='Expenses (Indirect/Admn.)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26013','Expenses (Indirect/Admn.)','Expenses (Indirect/Admn.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26014' OR Name='File-Sundry Creditors') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26014','File-Sundry Creditors','File-Sundry Creditors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26015' OR Name='File-Sundry Debtors') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26015','File-Sundry Debtors','File-Sundry Debtors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26016' OR Name='Fixed Assets') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26016','Fixed Assets','Fixed Assets','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26017' OR Name='Income (Direct/Opr.)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26017','Income (Direct/Opr.)','Income (Direct/Opr.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26018' OR Name='Income (Indirect)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26018','Income (Indirect)','Income (Indirect)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26019' OR Name='Income Tax Advance') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26019','Income Tax Advance','Income Tax Advance','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26021')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26020' OR Name='Investments') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26020','Investments','Investments','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26021' OR Name='Loans & Advances (Asset)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26021','Loans & Advances (Asset)','Loans & Advances (Asset)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26022' OR Name='Loans (Liability)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26022','Loans (Liability)','Loans (Liability)','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26023' OR Name='Pre-Operative Expenses') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26023','Pre-Operative Expenses','Pre-Operative Expenses','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26024' OR Name='Provisions/Expenses Payable') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26024','Provisions/Expenses Payable','Provisions/Expenses Payable','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26025' OR Name='Purchase') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26025','Purchase','Purchase','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26026' OR Name='Reserves & Surplus') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26026','Reserves & Surplus','Reserves & Surplus','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26006')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26027' OR Name='Sale') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26027','Sale','Sale','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26028' OR Name='Secured Loans') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26028','Secured Loans','Secured Loans','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26029' OR Name='Securities & Deposits (Asset)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26029','Securities & Deposits (Asset)','Securities & Deposits (Asset)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26030' OR Name='Sundry Creditors') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26030','Sundry Creditors','Sundry Creditors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26031' OR Name='Sundry Debtors') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26031','Sundry Debtors','Sundry Debtors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26032' OR Name='Suspense Account') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26032','Suspense Account','Suspense Account','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26033' OR Name='Unsecured Loans') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26033','Unsecured Loans','Unsecured Loans','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"

'Finance Master
frmLicenceAgreement.Label2 = "Account Master Update Going on!!! "
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01001' OR Name='Cash') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26007') Insert Into AccountMaster VALUES ('*01001','Cash','Cash','1001','*26007','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01002' OR Name='Development Tax') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01002','Development Tax','Development Tax','1002','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01003' OR Name='Edu. Cess on TDS') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01003','Edu. Cess on TDS','Edu. Cess on TDS','1003','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01004' OR Name='Excise Duty') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01004','Excise Duty','Excise Duty','1004','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01005' OR Name='KKC on Service Tax') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01005','KKC on Service Tax','KKC on Service Tax','1005','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01006' OR Name='SBC on Service Tax') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01006','SBC on Service Tax','SBC on Service Tax','1006','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01007' OR Name='Service Tax') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01007','Service Tax','Service Tax','1007','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01008' OR Name='SHE Cess on TDS') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01008','SHE Cess on TDS','SHE Cess on TDS','1008','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01009' OR Name='TDS (Commission or Brokerage)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01009','TDS (Commission or Brokerage)','TDS (Commission or Brokerage)','1009','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01010' OR Name='TDS (Contracts to Individuals/HUF)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01010','TDS (Contracts to Individuals/HUF)','TDS (Contracts to Individuals/HUF)','1010','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01011' OR Name='TDS (Contracts to Others)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01011','TDS (Contracts to Others)','TDS (Contracts to Others)','1011','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01012' OR Name='TDS (Contracts to Transporter)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01012','TDS (Contracts to Transporter)','TDS (Contracts to Transporter)','1012','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01013' OR Name='TDS (Interest from a Banking Co)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01013','TDS (Interest from a Banking Co)','TDS (Interest from a Banking Co)','1013','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01014' OR Name='TDS (Interest from a NonBanking Co)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01014','TDS (Interest from a NonBanking Co)','TDS (Interest from a NonBanking Co)','1014','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01015' OR Name='TDS (Professionals Services)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01015','TDS (Professionals Services)','TDS (Professionals Services)','1015','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01016' OR Name='TDS (Rent of Land)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01016','TDS (Rent of Land)','TDS (Rent of Land)','1016','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01017' OR Name='TDS (Rent of Plant & Machinery)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01017','TDS (Rent of Plant & Machinery)','TDS (Rent of Plant & Machinery)','1017','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01018' OR Name='TDS (Salary)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01018','TDS (Salary)','TDS (Salary)','1018','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01019' OR Name='Advertisement & Publicity') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01019','Advertisement & Publicity','Advertisement & Publicity','1019','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01020' OR Name='Bad Debts Written Off') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01020','Bad Debts Written Off','Bad Debts Written Off','1020','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01021' OR Name='Bank Charges') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01021','Bank Charges','Bank Charges','1021','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01022' OR Name='Books & Periodicals') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01022','Books & Periodicals','Books & Periodicals','1022','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01023' OR Name='Charity & Donations') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01023','Charity & Donations','Charity & Donations','1023','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01024' OR Name='Commission on Sales') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01024','Commission on Sales','Commission on Sales','1024','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01025' OR Name='Conveyance Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01025','Conveyance Expenses','Conveyance Expenses','1025','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01026' OR Name='Customer Entertainment Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01026','Customer Entertainment Expenses','Customer Entertainment Expenses','1026','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01027' OR Name='Depreciation A/c') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01027','Depreciation A/c','Depreciation A/c','1027','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01028' OR Name='Freight & Forwarding Charges') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01028','Freight & Forwarding Charges','Freight & Forwarding Charges','1028','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01029' OR Name='Legal Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01029','Legal Expenses','Legal Expenses','1029','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01030' OR Name='Miscellaneous Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01030','Miscellaneous Expenses','Miscellaneous Expenses','1030','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01031' OR Name='Office Maintenance Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01031','Office Maintenance Expenses','Office Maintenance Expenses','1031','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01032' OR Name='Office Rent') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01032','Office Rent','Office Rent','1032','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01033' OR Name='Postal Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01033','Postal Expenses','Postal Expenses','1033','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01034' OR Name='Printing & Stationery') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01034','Printing & Stationery','Printing & Stationery','1034','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01035' OR Name='Rounded Off') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01035','Rounded Off','Rounded Off','1035','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01036' OR Name='Salary') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01036','Salary','Salary','1036','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01037' OR Name='Sales Promotion Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01037','Sales Promotion Expenses','Sales Promotion Expenses','1037','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01038' OR Name='Service Charges Paid') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01038','Service Charges Paid','Service Charges Paid','1038','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01039' OR Name='Staff Welfare Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01039','Staff Welfare Expenses','Staff Welfare Expenses','1039','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01040' OR Name='Telephone Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01040','Telephone Expenses','Telephone Expenses','1040','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01041' OR Name='Travelling Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01041','Travelling Expenses','Travelling Expenses','1041','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01042' OR Name='Water & Electricity Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01042','Water & Electricity Expenses','Water & Electricity Expenses','1042','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01043' OR Name='Capital Equipments') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26016') Insert Into AccountMaster VALUES ('*01043','Capital Equipments','Capital Equipments','1043','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01044' OR Name='Computers') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26016') Insert Into AccountMaster VALUES ('*01044','Computers','Computers','1044','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01045' OR Name='Furniture & Fixture') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26016') Insert Into AccountMaster VALUES ('*01045','Furniture & Fixture','Furniture & Fixture','1045','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01046' OR Name='Office Equipments') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26016') Insert Into AccountMaster VALUES ('*01046','Office Equipments','Office Equipments','1046','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01047' OR Name='Plant & Machinery') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26016') Insert Into AccountMaster VALUES ('*01047','Plant & Machinery','Plant & Machinery','1047','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01048' OR Name='Service Charges Receipts') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26018') Insert Into AccountMaster VALUES ('*01048','Service Charges Receipts','Service Charges Receipts','1048','*26018','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01049' OR Name='Profit & Loss') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26001') Insert Into AccountMaster VALUES ('*01049','Profit & Loss','Profit & Loss','1049','*26001','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01050' OR Name='Salary & Bonus Payable') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26024') Insert Into AccountMaster VALUES ('*01050','Salary & Bonus Payable','Salary & Bonus Payable','1050','*26024','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01051' OR Name='Purchase') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26025') Insert Into AccountMaster VALUES ('*01051','Purchase','Purchase','1051','*26025','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01052' OR Name='Sales') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26027') Insert Into AccountMaster VALUES ('*01052','Sales','Sales','1052','*26027','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01053' OR Name='Earnest Money') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26029') Insert Into AccountMaster VALUES ('*01053','Earnest Money','Earnest Money','1053','*26029','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01054' OR Name='Stock') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26003') Insert Into AccountMaster VALUES ('*01054','Stock','Stock','1054','*26003','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01055' OR Name='Easy Info Solutions International') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26030') Insert Into AccountMaster VALUES ('*01055','Easy Info Solutions International','Easy Info Solutions International','1055','*26030','E-461, Vijay Marg,Jagjeet Nagar','Delhi-110053','','','','+91-987-342-2907','','sales@easyinfosolution.com ','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01056' OR Name='XXX Bank') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26004') Insert Into AccountMaster VALUES ('*01056','XXX Bank','XXX Bank','1056','*26004','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"

'Booking Route Master
frmLicenceAgreement.Label2 = "Booking Route Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS(SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookingRouteMaster' AND COLUMN_NAME = 'Code') Print 'Col_Exist' ELSE CREATE TABLE dbo.BookingRouteMaster (Code nvarchar(6) NOT NULL,Name nvarchar(40) NOT NULL,PrintName nvarchar(40) NOT NULL,Rate decimal(12, 2) NOT NULL,Printstatus nvarchar(1) NOT NULL)  ON [PRIMARY] ALTER TABLE dbo.BookingRouteMaster SET (LOCK_ESCALATION = TABLE)"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM BookingRouteMaster WHERE Code='*00001' OR Name='NOIDA-NOIDA') Print 'Exist' ELSE Insert Into BookingRouteMaster VALUES ('*00001','NOIDA-NOIDA','NOIDA-NOIDA','24.5','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM BookingRouteMaster WHERE Code='*00002' OR Name='NOIDA-DELHI') Print 'Exist' ELSE Insert Into BookingRouteMaster VALUES ('*00002','NOIDA-DELHI','NOIDA-DELHI','40','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM BookingRouteMaster WHERE Code='*00003' OR Name='DELHI-DELHI') Print 'Exist' ELSE Insert Into BookingRouteMaster VALUES ('*00003','DELHI-DELHI','DELHI-DELHI','30','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM BookingRouteMaster WHERE Code='*00004' OR Name='Local-Local') Print 'Exist' ELSE Insert Into BookingRouteMaster VALUES ('*00004','Local-Local','Local-Local','20','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM BookingRouteMaster WHERE Code='*00005' OR Name='Local-NCR') Print 'Exist' ELSE Insert Into BookingRouteMaster VALUES ('*00005','Local-NCR','Local-NCR','30','N')"

'Element Master
frmLicenceAgreement.Label2 = "Element Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00011' OR NAME='Text-1') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00011','Text-1','Text-1','Single Sheet','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00012' OR NAME='Text-2') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00012','Text-2','Text-2','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00013' OR NAME='Text-3') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00013','Text-3','Text-3','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00014' OR NAME='Single Form') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00014','Single Form','Single Form','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00015' OR NAME='Combo Form') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00015','Combo Form','Combo Form','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00016' OR NAME='FG') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00016','FG','FG','FG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00017' OR NAME='UFG') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00017','UFG','UFG','UFG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00018' OR NAME='Separator') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00018','Separator','Separator','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00019' OR NAME='End Paper') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00019','End Paper','End Paper','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00020' OR NAME='Cover') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00020','Cover','Cover','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00027' OR NAME='Title') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00027','Title','Title','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00028' OR NAME='Title(GateFold)') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00028','Title(GateFold)','Title(GateFold)','Single Sheet','6','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00029' OR NAME='PLC') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00029','PLC','PLC','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00030' OR NAME='Calendar Fly Leaf') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00030','Calendar Fly Leaf','Calendar Fly Leaf','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00031' OR NAME='Calendar Leaf') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00031','Calendar Leaf','Calendar Leaf','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00032' OR NAME='Annual Report') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00032','Annual Report','Annual Report','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00033' OR NAME='Label') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00033','Label','Label','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00034' OR NAME='Letter Head') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00034','Letter Head','Letter Head','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00035' OR NAME='Leaflet') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00035','Leaflet','Leaflet','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00036' OR NAME='Poster') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00036','Poster','Poster','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00037' OR NAME='Sticker') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00037','Sticker','Sticker','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00038' OR NAME='Folders') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00038','Folders','Folders','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00039' OR NAME='Dust Cover') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00039','Dust Cover','Dust Cover','Single Sheet','6','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00040' OR NAME='Danglar') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00040','Danglar','Danglar','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00041' OR NAME='Carton') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00041','Carton','Carton','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00042' OR NAME='Carton [Inner]') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00042','Carton [Inner]','Carton [Inner]','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00043' OR NAME='Carton [Outer]') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00043','Carton [Outer]','Carton [Outer]','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00044' OR NAME='Card') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00044','Card','Card','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00045' OR NAME='Envelope') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00045','Envelope','Envelope','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"

'Finish Size Child
frmLicenceAgreement.Label2 = "Finish Size Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11011*010391616*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11011','*01039','16','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01030') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11012*010301616*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11012','*01030','16','16','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01064') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11012*010643216*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11012','*01064','32','16','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11013*010391616*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11013','*01039','16','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11014*010391616*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11014','*01039','16','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11015*010551616*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11015','*01055','16','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01048') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11016*010481616*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11016','*01048','16','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01051') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11017*010511616*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11017','*01051','16','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01058') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11018*010581616*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11018','*01058','16','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01056') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11019*010561616*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11019','*01056','16','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11020*01028816*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11020','*01028','8','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11020*010601616*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11020','*01060','16','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01067') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11021*010671616*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11021','*01067','16','16','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11033*01039816*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11033','*01039','8','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11023*01031816*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11023','*01031','8','16','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01067') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11023*010671616*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11023','*01067','16','16','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01033') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01033') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11024*01033816*01033')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11024','*01033','8','16','*01033') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01033') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11024*010681616*01033')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11024','*01068','16','16','*01033') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11025*010681616*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11025','*01068','16','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01037') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11026*01037816*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11026','*01037','8','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01070') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11026*010701616*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11026','*01070','16','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01054') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11027*01054816*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11027','*01054','8','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01072') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11028*010721616*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11028','*01072','16','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01038') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11029*01038816*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11029','*01038','8','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01072') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11029*010721616*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11029','*01072','16','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11030*010551224*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11030','*01055','12','24','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11031*01039816*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11031','*01039','8','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01046') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11032*01046816*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11032','*01046','8','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01048') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11034*01048816*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11034','*01048','8','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01063') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11035*010631224*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11035','*01063','12','24','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01048') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11036*01048816*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11036','*01048','8','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01048') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11037*01048816*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11037','*01048','8','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11038*01055816*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11038','*01055','8','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01067') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11039*010671224*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11039','*01067','12','24','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01050') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11040*01050816*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11040','*01050','8','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01058') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11041*01058816*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11041','*01058','8','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01027') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11042*0102748*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11042','*01027','4','8','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01070') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11043*010701224*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11043','*01070','12','24','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11044*01060816*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11044','*01060','8','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11046*01060816*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11046','*01060','8','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11047*01039612*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11047','*01039','6','12','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01073') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11049*010731616*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11049','*01073','16','16','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11050*01068816*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11050','*01068','8','16','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11051*01055612*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11051','*01055','6','12','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01072') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11052*01072612*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11052','*01072','6','12','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11053*0106048*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11053','*01060','4','8','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11054*0106848*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11054','*01068','4','8','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01070') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11055*01070612*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11055','*01070','6','12','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01073') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11004*0107348*01017')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11004','*01073','4','8','*01017') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01012') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11004*0103922*01012')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11004','*01039','2','2','*01012') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01012') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01012') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11004*0101211*01012')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11004','*01012','1','1','*01012') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11057*01055816*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11057','*01055','8','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01029') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11048*0102848*01029')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11048','*01028','4','8','*01029') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01059') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01029') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11048*01059816*01029')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11048','*01059','8','16','*01029') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01058') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01029') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11058*01058816*01029')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11058','*01058','8','16','*01029') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01063') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11045*01063816*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11045','*01063','8','16','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01029') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11085*010601224*01029')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11085','*01060','12','24','*01029') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01011') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01011') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11005*0101122*01011')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11005','*01011','2','2','*01011') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01065') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01029') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11092*01065816*01029')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11092','*01065','8','16','*01029') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01029') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11091*0102848*01029')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11091','*01028','4','8','*01029') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01067') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01029') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11091*01067816*01029')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11091','*01067','8','16','*01029') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11094*01028816*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11094','*01028','8','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11094*010681616*01028')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11094','*01068','16','16','*01028') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01072') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11022*010721616*01031')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11022','*01072','16','16','*01031') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01045') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01047') IF EXISTS (Select Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize As FSIZE From FinishSizeChild Where (Code+[TextSIZE]+convert(nvarchar,[UPS/Form])+convert(nvarchar,[Ups/BdgForm])+TitleSize)=('*11095*0104588*01047')) Print 'Finish Size Exist' ELSE Insert Into FinishSizeChild VALUES ('*11095','*01045','8','8','*01047') ELSE Print 'Titlt Size Not Exist' Else Print 'Text Size Not Exist'"

'SizeGroupChild
frmLicenceAgreement.Label2 = "Size Group Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01067') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10003*01067') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10003','*01067')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10003*01068') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10003','*01068')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01070') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10003*01070') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10003','*01070')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01072') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10003*01072') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10003','*01072')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01073') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10003*01073') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10003','*01073')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01061') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10007*01061') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10007','*01061')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01047') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10011*01047') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10011','*01047')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01050') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10006*01050') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10006','*01050')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01051') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10006*01051') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10006','*01051')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01056') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10006*01056') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10006','*01056')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01058') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10006*01058') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10006','*01058')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10006*01060') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10006','*01060')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01063') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10006*01063') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10006','*01063')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01064') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10006*01064') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10006','*01064')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01059') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10006*01059') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10006','*01059')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10012*01017') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10012','*01017')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01020') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10012*01020') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10012','*01020')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01021') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10012*01021') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10012','*01021')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01027') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10012*01027') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10012','*01027')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10012*01028') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10012','*01028')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01030') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10012*01030') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10012','*01030')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10012*01031') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10012','*01031')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01033') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10012*01033') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10012','*01033')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01013') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10012*01013') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10012','*01013')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01012') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10013*01012') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10013','*01012')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01015') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10013*01015') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10013','*01015')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01016') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10013*01016') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10013','*01016')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01019') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10013*01019') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10013','*01019')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01029') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10013*01029') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10013','*01029')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01018') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10013*01018') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10013','*01018')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01069') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10018*01069') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10018','*01069')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01036') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10009*01036') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10009','*01036')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01037') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10009*01037') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10009','*01037')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01038') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10009*01038') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10009','*01038')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10009*01039') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10009','*01039')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01046') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10009*01046') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10009','*01046')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01048') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10009*01048') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10009','*01048')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01054') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10009*01054') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10009','*01054')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01057') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10009*01057') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10009','*01057')  Else Print 'Size NOT Exist' "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01011') IF EXISTS(Select Code+Size AS SGChild From SizeGroupChild Where Code+Size='*10020*01011') Print 'Size Group Code Exist' ELSE Insert Into SizeGroupChild VALUES ('*10020','*01011')  Else Print 'Size NOT Exist' "

'Tax Master
frmLicenceAgreement.Label2 = "Tax Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM TaxMaster WHERE Code='*00001' OR Name='Local GST 12%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00001','Local GST 12%','Local GST 12%','L','6','6',0,'000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM TaxMaster WHERE Code='*00002' OR Name='IGST 12%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00002','IGST 12%','IGST 12%','I','0','0',12,'000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM TaxMaster WHERE Code='*00003' OR Name='IGST 5%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00003','IGST 5%','IGST 5%','I','0','0',5,'000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM TaxMaster WHERE Code='*00004' OR Name='Local GST 5%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00004','Local GST 5%','Local GST 5%','L','2.5','2.5',0,'000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM TaxMaster WHERE Code='*00005' OR Name='Local GST 18%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00005','Local GST 18%','Local GST 18%','L','9','9',0,'000006',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM TaxMaster WHERE Code='*00006' OR Name='IGST 18%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00006','IGST 18%','IGST 18%','I','0','0',18,'000006',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM TaxMaster WHERE Code='*00007' OR Name='Local GST NIL') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00007','Local GST NIL','Local GST NIL','L','0','0',0,'000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM TaxMaster WHERE Code='*00008' OR Name='IGST NIL') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00008','IGST NIL','IGST NIL','I','0','0',0,'000001',GetDate(),'NULL',NULL,'N','N')"

'CompChild
frmLicenceAgreement.Label2 = "Company Child Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='01') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','01','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','" & CompAlias & "'+'/Pur/','/20-21','Purchase')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='02') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','02','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','" & CompAlias & "'+'/PR/','/20-21','Purchase Return')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='03') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','03','1. All disputes are subject to Our Jurisdiction Only','2. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection.','3. Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','','','" & CompAlias & "'+'/SR/','/20-21','Sale Return')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='04') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','04','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/Sale/','/20-21','Sale')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='05') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','05','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','" & CompAlias & "'+'/PC/','/20-21','Purchase Challan IN')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='06') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','06','','','','','','','','" & CompAlias & "'+'/PRC/','/20-21','Purchase Challan Out')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='07') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','07','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/SRC/','/20-21','Sale Challan IN')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='08') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','08','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/SC/','/20-21','Sale Challan Out')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='09') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','09','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/SJ/','/20-21','Sale Jobwork')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='10') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','10','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/SC/','/20-21','Sale Jobwork Unit Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='11') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','11','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/DN/','/20-21','Challan Revesal IN')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='12') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','12','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Delhi Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/PU/','/20-21','Challan Revesal Out')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='13') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','13','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/SC/','/20-21','Challan TO Be Billed IN')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='14') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','14','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/SC/','/20-21','Challan TO Be Billed OUT')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='15') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','15','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/SC/','/20-21','Challan Not TO Be Billed IN')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='16') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','16','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/SC/','/20-21','Challan Not TO Be Billed IOUT')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='17') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','17','1. The Deliverables shall be delivered or performed on the ','date and at the place specified in the Purchase Order.','2. Prices shall be as specified in the  Purchase  Order.','3. No increase in price shall be made or accepted unless ',' agreed in writing by Accenture.','4. The  Deliverables must conform in all respects with the','   Specifications and must be of sound.','" & CompAlias & "'+'/PO/','/20-21','Purchase Order')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='18') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','18','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/SO/','/20-21','Sale Order')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='19') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','19','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/ST/','/20-21','Stock Tranfer')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='20') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','20','','','','','','','','" & CompAlias & "'+'/RN/','/20-21','Stock Genral')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='21') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','21','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Delhi Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/SU/','/20-21','Promotional Sale Challan Out')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='22') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','22','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','" & CompAlias & "'+'/SQ/','/20-21','--')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='23') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','23','1. The price set for in Supplier’s Quotation (“Price”) are',' in  INDIA INR.','2. All Taxes shall be paid by Customer in addition to the ',' Price.','3.  Quotation (“Prices”) are valid for 30 days only.','','','" & CompAlias & "'+'/QP/','/20-21','Purchase Quotation')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='24') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','24','1. The price set for in Supplier’s Quotation (“Price”) are',' in  INDIA INR.','2. All Taxes shall be paid by Customer in addition to the ',' Price.','3.  Quotation (“Prices”) are valid for 30 days only.','','','" & CompAlias & "'+'/QS/','/20-21','Sales Quotation')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='25') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','25','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='26') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','26','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='27') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','27','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='28') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','28','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='29') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','29','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='30') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','30','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='31') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','31','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='32') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','32','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='33') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','33','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='34') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','34','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='35') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','35','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='36') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','36','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='37') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','37','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='38') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','38','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='39') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','39','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='40') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','40','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='41') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','41','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='42') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','42','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='43') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','43','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='44') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','44','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='45') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','45','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='46') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','46','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='47') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','47','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='48') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','48','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='49') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','49','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='50') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','50','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='51') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','51','','','','','','','','" & CompAlias & "'+'/PI/','/20-21','Payment')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='52') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','52','','','','','','','','" & CompAlias & "'+'/PR/','/20-21','Receipt')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='53') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','53','','','','','','','','" & CompAlias & "'+'/JE/','/20-21','Journal')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='54') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','54','','','','','','','','" & CompAlias & "'+'/CE/','/20-21','Contra')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='55') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','55','','','','','','','','" & CompAlias & "'+'/DN/','/20-21','Debit Note')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='56') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','56','','','','','','','','" & CompAlias & "'+'/CN/','/20-21','Credit Note')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='57') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','57','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='58') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','58','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='59') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','59','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='60') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','60','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='61') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','61','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='62') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','62','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='63') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','63','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='64') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','64','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='65') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','65','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='66') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','66','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='67') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','67','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='68') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','68','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='69') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','69','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='70') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','70','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='71') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','71','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='72') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','72','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='73') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','73','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='74') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','74','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='75') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','75','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='76') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','76','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='77') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','77','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='78') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','78','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='79') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','79','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='80') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','80','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='81') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','81','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='82') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','82','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='83') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','83','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='84') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','84','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='85') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','85','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='86') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','86','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='87') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','87','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='88') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','88','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='89') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','89','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='90') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','90','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='91') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','91','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='92') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','92','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='93') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','93','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='94') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','94','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='95') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','95','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='96') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','96','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='97') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','97','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='98') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','98','','','','','','','','','','')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='99') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','99','','','','','','','','','','')"
'Vch Series Master
frmLicenceAgreement.Label2 = "Voucher Series Master Update Going on!!! "
   cnDatabase.Execute "IF COL_LENGTH('VchSeriesMaster', 'VchName') IS NOT NULL PRINT 'Exists' ELSE Alter Table VchSeriesMaster Add VchName nvarchar(40) NOT NULL CONSTRAINT df_VchName DEFAULT '' "
   cnDatabase.Execute "IF EXISTS(SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'VchSeriesMaster' AND COLUMN_NAME = 'Code') Print 'Col_Exist' ELSE CREATE TABLE dbo.VchSeriesMaster (Code nvarchar(6) NOT NULL,Name nvarchar(40) NOT NULL,VchType nvarchar(4) NOT NULL,Prefix nvarchar(20) NULL,Suffix nvarchar(20) NULL,VchNumbering Char(1) NOT NULL,VchName nvarchar(40) NOT NULL CONSTRAINT df_VchName DEFAULT '' )  ON [PRIMARY] ALTER TABLE dbo.VchSeriesMaster SET (LOCK_ESCALATION = TABLE)"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00101' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00101','Main','01PF','" & CompAlias & "'+'/','/Purc','A','Purchase')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00102' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00102','Main','01PU','" & CompAlias & "'+'/','/PrJU','A','Purchase Unit Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00103' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00103','Main','01PC','" & CompAlias & "'+'/','/PrJC','A','Purchase Jobwork Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00104' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00104','Main','01PJ','" & CompAlias & "'+'/','/PrJW','A','Purchase Jobwork')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00201' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00201','Main','02OF','" & CompAlias & "'+'/','/PrRt','A','Purchase Return')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00202' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00202','Main','02OU','" & CompAlias & "'+'/','/PrRtJU','A','Purchase Return Unit Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00203' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00203','Main','02OC','" & CompAlias & "'+'/','/PrRtJC','A','Purchase Return Jobwork Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00204' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00204','Main','02OJ','" & CompAlias & "'+'/','/PrRtJW','A','Purchase Return Jobwork')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00301' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00301','Main','03TF','" & CompAlias & "'+'/','/SlRt','A','Sale Return')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00302' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00302','Main','03TU','" & CompAlias & "'+'/','/SlRtJU','A','Sale Return Unit Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00303' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00303','Main','03TC','" & CompAlias & "'+'/','/SlRtJC','A','Sale Return Jobwork Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00304' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00304','Main','03TJ','" & CompAlias & "'+'/','/SlRtJW','A','Sale Return Jobwork')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00401' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00401','Main','04SF','" & CompAlias & "'+'/','/Sale','A','Sales')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00402' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00402','Main','04SU','" & CompAlias & "'+'/','/SlJU','A','Sales Unit Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00403' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00403','Main','04SC','" & CompAlias & "'+'/','/SlJC','A','Sales Jobwork Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00404' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00404','Main','04SJ','" & CompAlias & "'+'/','/SlJW','A','Sales Jobwork')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00501' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00501','Main','05RF','" & CompAlias & "'+'/','/MtRc','A','Purchase Challan IN')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00502' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00502','Main','05FR','" & CompAlias & "'+'/','/MtRcJW','A','Purchase Challan IN (Jobwork)')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00601' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00601','Main','06IF','" & CompAlias & "'+'/','/PrRtC','A','Purchase Challan Out')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00602' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00602','Main','06FI','" & CompAlias & "'+'/','/PrRtCJW','A','Purchase Challan Out (Jobworj)')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00701' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00701','Main','07RF','" & CompAlias & "'+'/','/SlRtC','A','Sale Challan IN')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00702' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00702','Main','07FR','" & CompAlias & "'+'/','/SlRtCJW','A','Sale Challan IN (Jobwork)')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00801' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00801','Main','08IF','" & CompAlias & "'+'/','/MtIs','A','Sale Challan Out')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00802' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00802','Main','08FI','" & CompAlias & "'+'/','/MtIsJW','A','Sale Challan Out (Jobwork)')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*01701' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*01701','Main','17PO','" & CompAlias & "'+'/','/PO','A','Purchase Order')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*01801' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*01801','Main','18SO','" & CompAlias & "'+'/','/SO','A','Sale Order')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*01901' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*01901','Main','19ST','" & CompAlias & "'+'/','/STrn','A','Stock Tranfer')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02001' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02001','Main','20JR','" & CompAlias & "'+'/','/SJrnl','A','Stock Genral')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02101' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02101','Main','21JR','" & CompAlias & "'+'/','/SJrnl','A','Promotional Sale Challan Out')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02201' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02201','Main','22JR','" & CompAlias & "'+'/','/SJrnl','A','Promotional Purchase Challan Out')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02301' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02301','Main','23PQ','" & CompAlias & "'+'/','/PQ','A','Purchase Quotation')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02302' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02302','Main','23UZ','" & CompAlias & "'+'/','/PQU','A','Purchase Quotation Unit Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02303' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02303','Main','23CZ','" & CompAlias & "'+'/','/PQC','A','Purchase Quotation Jobwork Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02304' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02304','Main','23JZ','" & CompAlias & "'+'/','/PQJ','A','Purchase Quotation Jobwork')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02305' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02305','Main','24SQ','" & CompAlias & "'+'/','/SQ','A','Sales Quotation')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02306' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02306','Main','24UQ','" & CompAlias & "'+'/','/SQU','A','Sales Quotation Unit Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02307' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02307','Main','24CQ','" & CompAlias & "'+'/','/SQC','A','Sales Quotation Jobwork Cost')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02308' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02308','Main','24JQ','" & CompAlias & "'+'/','/SQJ','A','Sales Quotation Jobwork')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05101' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05101','Main','51PI','" & CompAlias & "'+'/','/Pymt','A','Payments')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05201' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05201','Main','52PR','" & CompAlias & "'+'/','/Rcpt','A','Receipts')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05301' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05301','Main','53JE','" & CompAlias & "'+'/','/Jrnl','A','Journal')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05401' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05401','Main','54CE','" & CompAlias & "'+'/','/Cntr','A','Countra')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05501' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05501','Main','55CN','" & CompAlias & "'+'/','/CrNt','A','Credit Note')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05601' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05601','Main','56DN','" & CompAlias & "'+'/','/DrNt','A','Debit Note')"
'Update Paper Unit Code
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Gross' AND TYPE=15) Update PaperMaster Set UOM='*15001' Where UOM=(Select Code From GeneralMaster Where NAME='Gross' AND Type=15) ELSE Print 'UOM NOT Exist' ;    IF EXISTS (Select * From GeneralMaster Where NAME='Gross' AND TYPE=15) Update GeneralMaster Set Code='*15001' Where Code=(Select Code From GeneralMaster Where NAME='Gross' AND Type=15) ELSE Print 'UOM NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Packet(100)' AND TYPE=15) Update PaperMaster Set UOM='*15002' Where UOM=(Select Code From GeneralMaster Where NAME='Packet(100)' AND Type=15) ELSE Print 'UOM NOT Exist' ;    IF EXISTS (Select * From GeneralMaster Where NAME='Packet(100)' AND TYPE=15) Update GeneralMaster Set Code='*15002' Where Code=(Select Code From GeneralMaster Where NAME='Packet(100)' AND Type=15) ELSE Print 'UOM NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Packet(150)' AND TYPE=15) Update PaperMaster Set UOM='*15003' Where UOM=(Select Code From GeneralMaster Where NAME='Packet(150)' AND Type=15) ELSE Print 'UOM NOT Exist' ;    IF EXISTS (Select * From GeneralMaster Where NAME='Packet(150)' AND TYPE=15) Update GeneralMaster Set Code='*15003' Where Code=(Select Code From GeneralMaster Where NAME='Packet(150)' AND Type=15) ELSE Print 'UOM NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Ream' AND TYPE=15) Update PaperMaster Set UOM='*15004' Where UOM=(Select Code From GeneralMaster Where NAME='Ream' AND Type=15) ELSE Print 'UOM NOT Exist' ;  IF EXISTS (Select * From GeneralMaster Where NAME='Ream' AND TYPE=15) Update GeneralMaster Set Code='*15004' Where Code=(Select Code From GeneralMaster Where NAME='Ream' AND Type=15) ELSE Print 'UOM NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Reel' AND TYPE=15) Update PaperMaster Set UOM='*15005' Where UOM=(Select Code From GeneralMaster Where NAME='Reel' AND Type=15) ELSE Print 'UOM NOT Exist' ;  IF EXISTS (Select * From GeneralMaster Where NAME='Reel' AND TYPE=15) Update GeneralMaster Set Code='*15005' Where Code=(Select Code From GeneralMaster Where NAME='Reel' AND Type=15) ELSE Print 'UOM NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Bundle (700)' AND TYPE=15) Update PaperMaster Set UOM='*15006' Where UOM=(Select Code From GeneralMaster Where NAME='Bundle (700)' AND Type=15) ELSE Print 'UOM NOT Exist' ;  IF EXISTS (Select * From GeneralMaster Where NAME='Bundle (700)' AND TYPE=15) Update GeneralMaster Set Code='*15006' Where Code=(Select Code From GeneralMaster Where NAME='Bundle (700)' AND Type=15) ELSE Print 'UOM NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Packet(200)' AND TYPE=15) Update PaperMaster Set UOM='*15007' Where UOM=(Select Code From GeneralMaster Where NAME='Packet(200)' AND Type=15) ELSE Print 'UOM NOT Exist' ;    IF EXISTS (Select * From GeneralMaster Where NAME='Packet(200)' AND TYPE=15) Update GeneralMaster Set Code='*15007' Where Code=(Select Code From GeneralMaster Where NAME='Packet(200)' AND Type=15) ELSE Print 'UOM NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='PACKET' AND TYPE=15) Update PaperMaster Set UOM='*15008' Where UOM=(Select Code From GeneralMaster Where NAME='PACKET' AND Type=15) ELSE Print 'UOM NOT Exist' ;  IF EXISTS (Select * From GeneralMaster Where NAME='PACKET' AND TYPE=15) Update GeneralMaster Set Code='*15008' Where Code=(Select Code From GeneralMaster Where NAME='PACKET' AND Type=15) ELSE Print 'UOM NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Sheet' AND TYPE=15) Update PaperMaster Set UOM='*15009' Where UOM=(Select Code From GeneralMaster Where NAME='Sheet' AND Type=15) ELSE Print 'UOM NOT Exist' ;    IF EXISTS (Select * From GeneralMaster Where NAME='Sheet' AND TYPE=15) Update GeneralMaster Set Code='*15009' Where Code=(Select Code From GeneralMaster Where NAME='Sheet' AND Type=15) ELSE Print 'UOM NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Packet (250)' AND TYPE=15) Update PaperMaster Set UOM='*15010' Where UOM=(Select Code From GeneralMaster Where NAME='Packet (250)' AND Type=15) ELSE Print 'UOM NOT Exist' ;  IF EXISTS (Select * From GeneralMaster Where NAME='Packet (250)' AND TYPE=15) Update GeneralMaster Set Code='*15010' Where Code=(Select Code From GeneralMaster Where NAME='Packet (250)' AND Type=15) ELSE Print 'UOM NOT Exist'"
'Update Paper Quality Code
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Coated Matt' AND TYPE=16) Update PaperMaster Set Quality='*16001' Where Quality=(Select Code From GeneralMaster Where NAME='Coated Matt' AND Type=16) ELSE Print 'Quality NOT Exist' ;    IF EXISTS (Select * From GeneralMaster Where NAME='Coated Matt' AND TYPE=16) Update GeneralMaster Set Code='*16001' Where Code=(Select Code From GeneralMaster Where NAME='Coated Matt' AND Type=16) ELSE Print 'Quality NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Coated Gloss' AND TYPE=16) Update PaperMaster Set Quality='*16002' Where Quality=(Select Code From GeneralMaster Where NAME='Coated Gloss' AND Type=16) ELSE Print 'Quality NOT Exist' ;  IF EXISTS (Select * From GeneralMaster Where NAME='Coated Gloss' AND TYPE=16) Update GeneralMaster Set Code='*16002' Where Code=(Select Code From GeneralMaster Where NAME='Coated Gloss' AND Type=16) ELSE Print 'Quality NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Uncoated' AND TYPE=16) Update PaperMaster Set Quality='*16003' Where Quality=(Select Code From GeneralMaster Where NAME='Uncoated' AND Type=16) ELSE Print 'Quality NOT Exist' ;    IF EXISTS (Select * From GeneralMaster Where NAME='Uncoated' AND TYPE=16) Update GeneralMaster Set Code='*16003' Where Code=(Select Code From GeneralMaster Where NAME='Uncoated' AND Type=16) ELSE Print 'Quality NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='High Bulk' AND TYPE=16) Update PaperMaster Set Quality='*16004' Where Quality=(Select Code From GeneralMaster Where NAME='High Bulk' AND Type=16) ELSE Print 'Quality NOT Exist' ;    IF EXISTS (Select * From GeneralMaster Where NAME='High Bulk' AND TYPE=16) Update GeneralMaster Set Code='*16004' Where Code=(Select Code From GeneralMaster Where NAME='High Bulk' AND Type=16) ELSE Print 'Quality NOT Exist'"
'Paper Master
frmLicenceAgreement.Label2 = "Paper Master Update Going on!!! "
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00001') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Card-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00001','Art Card-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-Gloss','Art Card-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-Gloss','S','B','50.8','76.2','20','30','*15002','200','Art Card','Gloss','7.742','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00002') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Card-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00002','Art Card-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-Gloss','Art Card-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-Gloss','S','B','50.8','76.2','20','30','*15002','210','Art Card','Gloss','8.129','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00003') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Card-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00003','Art Card-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-Gloss','Art Card-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-Gloss','S','B','50.8','76.2','20','30','*15002','220','Art Card','Gloss','8.516','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00004') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Card-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00004','Art Card-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-Gloss','Art Card-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-Gloss','S','B','50.8','76.2','20','30','*15002','250','Art Card','Gloss','9.677','5','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00005') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Card-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00005','Art Card-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-Gloss','Art Card-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-Gloss','S','B','58.42','91.44','23','36','*15002','200','Art Card','Gloss','10.684','5','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00006') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Card-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00006','Art Card-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-Gloss','Art Card-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-Gloss','S','B','58.42','91.44','23','36','*15002','210','Art Card','Gloss','11.218','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00007') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Card-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00007','Art Card-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-Gloss','Art Card-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-Gloss','S','B','58.42','91.44','23','36','*15002','220','Art Card','Gloss','11.752','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00008') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Card-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00008','Art Card-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-Gloss','Art Card-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-Gloss','S','B','58.42','91.44','23','36','*15002','250','Art Card','Gloss','13.355','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00009') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00009','Art Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Gloss','Art Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Gloss','S','P','50.8','76.2','20','30','*15004','70','Art Paper','Gloss','13.548','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00010') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00010','Art Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Gloss','Art Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Gloss','S','P','50.8','76.2','20','30','*15004','80','Art Paper','Gloss','15.484','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00011') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00011','Art Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Gloss','Art Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Gloss','S','P','50.8','76.2','20','30','*15004','90','Art Paper','Gloss','17.419','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00012') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00012','Art Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Gloss','Art Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Gloss','S','P','50.8','76.2','20','30','*15004','100','Art Paper','Gloss','19.355','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00013') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-130gsm-20.00X30.00in²-(50.80X76.20cm²)-25.161kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00013','Art Paper-130gsm-20.00X30.00in²-(50.80X76.20cm²)-25.161kg-Gloss','Art Paper-130gsm-20.00X30.00in²-(50.80X76.20cm²)-25.161kg-Gloss','S','P','50.8','76.2','20','30','*15004','130','Art Paper','Gloss','25.161','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00014') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-170gsm-20.00X30.00in²-(50.80X76.20cm²)-32.903kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00014','Art Paper-170gsm-20.00X30.00in²-(50.80X76.20cm²)-32.903kg-Gloss','Art Paper-170gsm-20.00X30.00in²-(50.80X76.20cm²)-32.903kg-Gloss','S','P','50.8','76.2','20','30','*15004','170','Art Paper','Gloss','32.903','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00015') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00015','Art Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Gloss','Art Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Gloss','S','P','58.42','91.44','23','36','*15004','70','Art Paper','Gloss','18.697','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00016') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00016','Art Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Gloss','Art Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Gloss','S','P','58.42','91.44','23','36','*15004','80','Art Paper','Gloss','21.368','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00017') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00017','Art Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Gloss','Art Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Gloss','S','P','58.42','91.44','23','36','*15004','90','Art Paper','Gloss','24.039','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00018') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00018','Art Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Gloss','Art Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Gloss','S','P','58.42','91.44','23','36','*15004','100','Art Paper','Gloss','26.71','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00019') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-130gsm-23.00X36.00in²-(58.42X91.44cm²)-34.723kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00019','Art Paper-130gsm-23.00X36.00in²-(58.42X91.44cm²)-34.723kg-Gloss','Art Paper-130gsm-23.00X36.00in²-(58.42X91.44cm²)-34.723kg-Gloss','S','P','58.42','91.44','23','36','*15004','130','Art Paper','Gloss','34.723','1','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00020') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Art Paper-170gsm-23.00X36.00in²-(58.42X91.44cm²)-45.406kg-Gloss') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002') Insert Into PaperMaster VALUES ('*00020','Art Paper-170gsm-23.00X36.00in²-(58.42X91.44cm²)-45.406kg-Gloss','Art Paper-170gsm-23.00X36.00in²-(58.42X91.44cm²)-45.406kg-Gloss','S','P','58.42','91.44','23','36','*15004','170','Art Paper','Gloss','45.406','1','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00021') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-60gsm-20.00X30.00in²-(50.80X76.20cm²)-11.613kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00021','Paper-60gsm-20.00X30.00in²-(50.80X76.20cm²)-11.613kg-Maplitho','Paper-60gsm-20.00X30.00in²-(50.80X76.20cm²)-11.613kg-Maplitho','S','P','50.8','76.2','20','30','*15004','60','Paper','Maplitho','11.613','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00022') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-64gsm-20.00X30.00in²-(50.80X76.20cm²)-12.387kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00022','Paper-64gsm-20.00X30.00in²-(50.80X76.20cm²)-12.387kg-Maplitho','Paper-64gsm-20.00X30.00in²-(50.80X76.20cm²)-12.387kg-Maplitho','S','P','50.8','76.2','20','30','*15004','64','Paper','Maplitho','12.387','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00023') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00023','Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Maplitho','Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Maplitho','S','P','50.8','76.2','20','30','*15004','70','Paper','Maplitho','13.548','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00024') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00024','Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Maplitho','Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Maplitho','S','P','50.8','76.2','20','30','*15004','80','Paper','Maplitho','15.484','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00025') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00025','Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Maplitho','Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Maplitho','S','P','50.8','76.2','20','30','*15004','90','Paper','Maplitho','17.419','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00026') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00026','Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Maplitho','Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Maplitho','S','P','50.8','76.2','20','30','*15004','100','Paper','Maplitho','19.355','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00027') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-120gsm-20.00X30.00in²-(50.80X76.20cm²)-23.226kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00027','Paper-120gsm-20.00X30.00in²-(50.80X76.20cm²)-23.226kg-Maplitho','Paper-120gsm-20.00X30.00in²-(50.80X76.20cm²)-23.226kg-Maplitho','S','P','50.8','76.2','20','30','*15004','120','Paper','Maplitho','23.226','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00028') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-60gsm-23.00X36.00in²-(58.42X91.44cm²)-16.026kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00028','Paper-60gsm-23.00X36.00in²-(58.42X91.44cm²)-16.026kg-Maplitho','Paper-60gsm-23.00X36.00in²-(58.42X91.44cm²)-16.026kg-Maplitho','S','P','58.42','91.44','23','36','*15004','60','Paper','Maplitho','16.026','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00029') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-64gsm-23.00X36.00in²-(58.42X91.44cm²)-17.094kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00029','Paper-64gsm-23.00X36.00in²-(58.42X91.44cm²)-17.094kg-Maplitho','Paper-64gsm-23.00X36.00in²-(58.42X91.44cm²)-17.094kg-Maplitho','S','P','58.42','91.44','23','36','*15004','64','Paper','Maplitho','17.094','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00030') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00030','Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Maplitho','Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Maplitho','S','P','58.42','91.44','23','36','*15004','70','Paper','Maplitho','18.697','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00031') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00031','Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Maplitho','Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Maplitho','S','P','58.42','91.44','23','36','*15004','80','Paper','Maplitho','21.368','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00032') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00032','Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Maplitho','Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Maplitho','S','P','58.42','91.44','23','36','*15004','90','Paper','Maplitho','24.039','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00033') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00033','Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Maplitho','Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Maplitho','S','P','58.42','91.44','23','36','*15004','100','Paper','Maplitho','26.71','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00034') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-120gsm-23.00X36.00in²-(58.42X91.44cm²)-32.052kg-Maplitho') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00034','Paper-120gsm-23.00X36.00in²-(58.42X91.44cm²)-32.052kg-Maplitho','Paper-120gsm-23.00X36.00in²-(58.42X91.44cm²)-32.052kg-Maplitho','S','P','58.42','91.44','23','36','*15004','120','Paper','Maplitho','32.052','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00035') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='SBS-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-C1S') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00035','SBS-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-C1S','SBS-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-C1S','S','B','50.8','76.2','20','30','*15002','200','SBS','C1S','7.742','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00036') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='SBS-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-C1S') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00036','SBS-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-C1S','SBS-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-C1S','S','B','50.8','76.2','20','30','*15002','210','SBS','C1S','8.129','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00037') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='SBS-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-C1S') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00037','SBS-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-C1S','SBS-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-C1S','S','B','50.8','76.2','20','30','*15002','220','SBS','C1S','8.516','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00038') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='SBS-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-C1S') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00038','SBS-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-C1S','SBS-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-C1S','S','B','50.8','76.2','20','30','*15002','250','SBS','C1S','9.677','5','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00039') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='SBS-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-C1S') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00039','SBS-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-C1S','SBS-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-C1S','S','B','58.42','91.44','23','36','*15002','200','SBS','C1S','10.684','5','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00040') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='SBS-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-C1S') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00040','SBS-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-C1S','SBS-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-C1S','S','B','58.42','91.44','23','36','*15002','210','SBS','C1S','11.218','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00041') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='SBS-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-C1S') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00041','SBS-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-C1S','SBS-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-C1S','S','B','58.42','91.44','23','36','*15002','220','SBS','C1S','11.752','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00042') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='SBS-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-C1S') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00042','SBS-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-C1S','SBS-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-C1S','S','B','58.42','91.44','23','36','*15002','250','SBS','C1S','13.355','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
frmLicenceAgreement.Label2 = " Update Done!!! "
End Function
Public Function UpdateMinor02()
'Update VchSeries And AutoNo
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00101',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='01' AND Right(Type,2)='PF' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00102',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='01' AND Right(Type,2)='PU' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00103',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='01' AND Right(Type,2)='PC' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00104',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='01' AND Right(Type,2)='PJ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00201',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='02' AND Right(Type,2)='OF' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00202',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='02' AND Right(Type,2)='OU' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00203',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='02' AND Right(Type,2)='OC' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00204',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='02' AND Right(Type,2)='OJ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00301',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='03' AND Right(Type,2)='TF' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00302',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='03' AND Right(Type,2)='TU' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00303',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='03' AND Right(Type,2)='TC' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00304',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='03' AND Right(Type,2)='TJ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00401',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='04' AND Right(Type,2)='SF' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00402',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='04' AND Right(Type,2)='SU' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00403',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='04' AND Right(Type,2)='SC' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00404',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='04' AND Right(Type,2)='SJ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00501',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='05' AND Right(Type,2)='RF' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00502',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='05' AND Right(Type,2)='FR' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00601',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='06' AND Right(Type,2)='IF' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00602',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='06' AND Right(Type,2)='FI' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00701',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='07' AND Right(Type,2)='RF' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00702',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='07' AND Right(Type,2)='FR' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00801',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='08' AND Right(Type,2)='IF' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*00802',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='08' AND Right(Type,2)='FI' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*01701',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='17' AND Right(Type,2)='PO' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*01801',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='18' AND Right(Type,2)='SO' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*01901',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='19' AND Right(Type,2)='ST' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*02001',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='20' AND Right(Type,2)='JR' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*02101',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='21' AND Right(Type,2)='JR' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*02201',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='22' AND Right(Type,2)='JR' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*02301',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='23' AND Right(Type,2)='PQ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*02302',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='23' AND Right(Type,2)='UZ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*02303',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='23' AND Right(Type,2)='CZ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*02304',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='23' AND Right(Type,2)='JZ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*02305',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='24' AND Right(Type,2)='SQ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*02306',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='24' AND Right(Type,2)='UQ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*02307',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='24' AND Right(Type,2)='CQ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*02308',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='24' AND Right(Type,2)='JQ' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*05101',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='51' AND Right(Type,2)='PI' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*05201',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='52' AND Right(Type,2)='PR' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*05301',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='53' AND Right(Type,2)='JE' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*05401',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='54' AND Right(Type,2)='CE' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*05501',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='55' AND Right(Type,2)='CN' And VchSeries is null"
   cnDatabase.Execute "UPDATE JobworkBVParent SET VchSeries='*05601',AutoVchNo=REPLICATE(' ',10-LEN(LTRIM(Name)))+LTRIM(Name) WHERE Left(Type,2)='56' AND Right(Type,2)='DN' And VchSeries is null"
''Account Masters
    cnDatabase.Execute "IF Exists (Select * From AccountMaster Where Code='000000') PRINT 'Exists' ELSE Insert Into AccountMaster VALUES ('000000','" & Trim(rstCompanyMaster.Fields("Name").Value) & "','" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "','000000','*12002','" & Trim(rstCompanyMaster.Fields("Address1").Value) & "','" & Trim(rstCompanyMaster.Fields("Address2").Value) & "','" & Trim(rstCompanyMaster.Fields("Address3").Value) & "','" & Trim(rstCompanyMaster.Fields("Address4").Value) & "','" & Trim(rstCompanyMaster.Fields("Phone").Value) & "','" & Trim(rstCompanyMaster.Fields("Mobile").Value) & "','" & Trim(rstCompanyMaster.Fields("GSTIN").Value) & "','" & Trim(rstCompanyMaster.Fields("eMail").Value) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0)"
    cnDatabase.Execute "IF Exists (Select * From AccountMaster Where Code='*00001') PRINT 'Exists' Else Insert Into AccountMaster VALUES ('*00001','Rate Master','Rate Master','1002','*12002','" & Trim(rstCompanyMaster.Fields("Address1").Value) & "','" & Trim(rstCompanyMaster.Fields("Address2").Value) & "','" & Trim(rstCompanyMaster.Fields("Address3").Value) & "','" & Trim(rstCompanyMaster.Fields("Address4").Value) & "','" & Trim(rstCompanyMaster.Fields("Phone").Value) & "','" & Trim(rstCompanyMaster.Fields("Mobile").Value) & "','" & Trim(rstCompanyMaster.Fields("GSTIN").Value) & "','" & Trim(rstCompanyMaster.Fields("eMail").Value) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0)"
    cnDatabase.Execute "IF Exists (Select * From AccountMaster Where Code='*00002') PRINT 'Exists' ELSE Insert Into AccountMaster VALUES ('*00002','Main Godown','Main Godown','1003','*99999','" & Trim(rstCompanyMaster.Fields("Address1").Value) & "','" & Trim(rstCompanyMaster.Fields("Address2").Value) & "','" & Trim(rstCompanyMaster.Fields("Address3").Value) & "','" & Trim(rstCompanyMaster.Fields("Address4").Value) & "','" & Trim(rstCompanyMaster.Fields("Phone").Value) & "','" & Trim(rstCompanyMaster.Fields("Mobile").Value) & "','" & Trim(rstCompanyMaster.Fields("GSTIN").Value) & "','" & Trim(rstCompanyMaster.Fields("eMail").Value) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0)"
    cnDatabase.Execute "IF Exists (Select * From AccountMaster Where Code='*00003') PRINT 'Exists' ELSE Insert Into AccountMaster VALUES ('*00003','Self Transport','Self Transport','1004','*99996','" & Trim(rstCompanyMaster.Fields("Address1").Value) & "','" & Trim(rstCompanyMaster.Fields("Address2").Value) & "','" & Trim(rstCompanyMaster.Fields("Address3").Value) & "','" & Trim(rstCompanyMaster.Fields("Address4").Value) & "','" & Trim(rstCompanyMaster.Fields("Phone").Value) & "','" & Trim(rstCompanyMaster.Fields("Mobile").Value) & "','" & Trim(rstCompanyMaster.Fields("GSTIN").Value) & "','" & Trim(rstCompanyMaster.Fields("eMail").Value) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0)"
    cnDatabase.Execute "IF Exists (Select * From AccountMaster Where Code='*00004') PRINT 'Exists' ELSE Insert Into AccountMaster VALUES ('*00004','Self Packer','Self Packer','1005','*99997','" & Trim(rstCompanyMaster.Fields("Address1").Value) & "','" & Trim(rstCompanyMaster.Fields("Address2").Value) & "','" & Trim(rstCompanyMaster.Fields("Address3").Value) & "','" & Trim(rstCompanyMaster.Fields("Address4").Value) & "','" & Trim(rstCompanyMaster.Fields("Phone").Value) & "','" & Trim(rstCompanyMaster.Fields("Mobile").Value) & "','" & Trim(rstCompanyMaster.Fields("GSTIN").Value) & "','" & Trim(rstCompanyMaster.Fields("eMail").Value) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0)"
    cnDatabase.Execute "IF Exists (Select * From AccountMaster Where Code='*00005') PRINT 'Exists' ELSE Insert Into AccountMaster VALUES ('*00005','Direct','Direct','1006','*99998','" & Trim(rstCompanyMaster.Fields("Address1").Value) & "','" & Trim(rstCompanyMaster.Fields("Address2").Value) & "','" & Trim(rstCompanyMaster.Fields("Address3").Value) & "','" & Trim(rstCompanyMaster.Fields("Address4").Value) & "','" & Trim(rstCompanyMaster.Fields("Phone").Value) & "','" & Trim(rstCompanyMaster.Fields("Mobile").Value) & "','" & Trim(rstCompanyMaster.Fields("GSTIN").Value) & "','" & Trim(rstCompanyMaster.Fields("eMail").Value) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0)"
'Update Material Centre
    cnDatabase.Execute "IF NOT Exists (Select * From JobworkBVParent WHERE MaterialCentre is null) PRINT 'NOT_Exists' ELSE UPDATE JobworkBVParent SET MaterialCentre='*00002' WHERE MaterialCentre is null"
'Update JobworkBV Parent Name Field CHARACTER_MAXIMUM_LENGTH
    cnDatabase.Execute "IF (SELECT CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'JobworkBVParent' AND COLUMN_NAME = 'Name')<>'25' Alter Table JobworkBVParent  Alter Column Name nvarchar(25) Else Print 'CHARACTER_MAXIMUM_LENGTH_OK'"
End Function
Public Function UpdateMinor03()
cnDatabase.Execute "ALTER FUNCTION [dbo].[ufnGetPaperStock](@Account CHAR(6),@Paper CHAR(6),@VchType CHAR(2),@VchCode CHAR(6),@VchDate DATE) " & _
"RETURNS Decimal(12, 3) AS " & _
"BEGIN " & _
    "DECLARE @CurStk DECIMAL(12,3); " & _
    "SELECT @CurStk= " & _
    "((ISNULL((SELECT SUM(OpBalSheets) FROM PaperChild WHERE Code=I.Code AND Account=@Account),0)+ " & _
    "ISNULL((SELECT SUM(QuantitySheets) FROM PaperIOChild WHERE Paper=I.Code AND Account=@Account),0)+ " & _
    "ISNULL((SELECT SUM(PARSENAME(Quantity,2)*1)*U.Value1+SUM(PARSENAME(Quantity,1)*1) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=I.Code AND Quantity>=0 AND Account=@Account AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='JN',@VchCode,'XXXXXX')),0)+ " & _
    "ISNULL((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE [Type]='T' AND Paper=I.Code AND AccountTo=@Account  AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='TR',@VchCode,'XXXXXX')),0)+ " & _
    "ISNULL((SELECT SUM(PARSENAME(Quantity,2)*1)*U.Value1+SUM(PARSENAME(Quantity,1)*1) FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE P.Account=@Account  AND [Date]<=@VchDate AND C.Paper=I.Code AND Quantity>=0 AND P.Code<>IIF(@VchType='DN',@VchCode,'XXXXXX')),0))- " & _
    "(ISNULL((SELECT SUM(PARSENAME(0-Quantity,2)*1)*U.Value1+SUM(PARSENAME(0-Quantity,1)*1) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=I.Code AND Quantity<0 AND Account=@Account  AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='JN',@VchCode,'XXXXXX')),0)+ " & _
    "ISNULL((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=I.Code AND AccountFrom=@Account  AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='TR',@VchCode,'XXXXXX')),0)+ " & _
    "ISNULL((SELECT SUM(PARSENAME(0-Quantity,2)*1)*U.Value1+SUM(PARSENAME(0-Quantity,1)*1) FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE P.Account=@Account  AND [Date]<=@VchDate AND C.Paper=I.Code AND Quantity<0 AND P.Code<>IIF(@VchType='DN',@VchCode,'XXXXXX')),0)+ " & _
    "ISNULL((SELECT SUM(PaperConsumptionSheets1) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper1=I.Code AND RAccount1=@Account  AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='PO',@VchCode,'XXXXXX')),0)+ " & _
    "ISNULL((SELECT SUM(PaperConsumptionSheets2) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper2=I.Code AND RAccount2=@Account  AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='PO',@VchCode,'XXXXXX')),0)+ " & _
    "ISNULL((SELECT SUM(PaperConsumptionSheets4) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper4=I.Code AND RAccount4=@Account  AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='PO',@VchCode,'XXXXXX')),0)+ " & _
    "ISNULL((SELECT SUM(PaperConsumptionSheets) FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=I.Code AND RAccount=@Account  AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='PO',@VchCode,'XXXXXX')),0)+ " & _
    "ISNULL((SELECT SUM(PaperConsumptionSheets) FROM BookPOParent P INNER JOIN BookPOChild09 C ON P.Code=C.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=I.Code AND RAccount=@Account  AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='PO',@VchCode,'XXXXXX')),0)+ " & _
    "ISNULL((SELECT SUM(Round(C2.TotalConsumption,0)) FROM (BookPOParent P INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0801 C2 ON C1.Code=C2.Code WHERE LEFT(P.Type,1)<>'O' AND LEFT(P.Code,1)<>'*' AND C2.Category='2' AND C2.Item=I.Code AND BookPrinter=@Account  AND [Date]<=@VchDate AND P.Code<>IIF(@VchType='PO',@VchCode,'XXXXXX')),0)))/U.Value1 " & _
    "FROM PaperMaster I INNER JOIN GeneralMaster U ON I.UOM=U.Code WHERE I.Code=@Paper " & _
    "RETURN PARSENAME(@CurStk,2)*1+(@CurStk-PARSENAME(@CurStk,2)*1)/2; " & _
"End"
    'Booking Route Master
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM BookingRouteMaster WHERE Code='*00004' OR Name='Local-Local') Print 'Exist' ELSE Insert Into BookingRouteMaster VALUES ('*00004','Local-Local','Local-Local','20','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM BookingRouteMaster WHERE Code='*00005' OR Name='Local-NCR') Print 'Exist' ELSE Insert Into BookingRouteMaster VALUES ('*00005','Local-NCR','Local-NCR','30','N')"
End Function
Public Function UpdateMinor04()
'Quotation VchSeries Update
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='23ZP') Update VchSeriesMaster Set VchType='23PQ' Where VchType='23ZP' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='23UZ') Update VchSeriesMaster Set VchType='23ZU' Where VchType='23UZ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='23CZ') Update VchSeriesMaster Set VchType='23ZC' Where VchType='23CZ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='23JZ') Update VchSeriesMaster Set VchType='23ZJ' Where VchType='23JZ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='24QS') Update VchSeriesMaster Set VchType='24SQ' Where VchType='24QS' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='24UQ') Update VchSeriesMaster Set VchType='24QU' Where VchType='24UQ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='24CQ') Update VchSeriesMaster Set VchType='24QC' Where VchType='24CQ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='24JQ') Update VchSeriesMaster Set VchType='24QJ' Where VchType='24JQ' ELSE Print 'VchType NOT Exist'"
'Paper Unit Updates
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Gross' AND TYPE=15) Update GeneralMaster Set PrintName='Gross'  Where NAME='Gross' AND Type=15 ELSE Print 'UOM NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Packet(100)' AND TYPE=15) Update GeneralMaster Set PrintName='Packet'  Where NAME='Packet(100)' AND Type=15 ELSE Print 'UOM NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Packet(150)' AND TYPE=15) Update GeneralMaster Set PrintName='Packet'  Where NAME='Packet(150)' AND Type=15 ELSE Print 'UOM NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Ream' AND TYPE=15) Update GeneralMaster Set PrintName='Ream'  Where NAME='Ream' AND Type=15 ELSE Print 'UOM NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Reel' AND TYPE=15) Update GeneralMaster Set PrintName='Reel'  Where NAME='Reel' AND Type=15 ELSE Print 'UOM NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Bundle (700)' AND TYPE=15) Update GeneralMaster Set PrintName='Bundle'  Where NAME='Bundle (700)' AND Type=15 ELSE Print 'UOM NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Packet(200)' AND TYPE=15) Update GeneralMaster Set PrintName='Packet'  Where NAME='Packet(200)' AND Type=15 ELSE Print 'UOM NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='PACKET' AND TYPE=15) Update GeneralMaster Set PrintName='PACKET'  Where NAME='PACKET' AND Type=15 ELSE Print 'UOM NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Sheet' AND TYPE=15) Update GeneralMaster Set PrintName='Sheet'  Where NAME='Sheet' AND Type=15 ELSE Print 'UOM NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Packet (250)' AND TYPE=15) Update GeneralMaster Set PrintName='Packet'  Where NAME='Packet (250)' AND Type=15 ELSE Print 'UOM NOT Exist' ;"
'Genral Master
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Account Group' AND TYPE=12) Update AccountMaster Set [Group]='*12002' where [Group] IN (Select Code From GeneralMaster Where NAME='Account Group' AND TYPE=12) ELSE Print 'NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Account Group' AND TYPE=12) Update GeneralMaster Set Code='*12002'  Where NAME='Account Group' AND Type=12 ELSE Print 'NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Debtors' AND TYPE=12) Update AccountMaster Set [Group]='*12019' where [Group] IN (Select Code From GeneralMaster Where NAME='Debtors' AND TYPE=12) ELSE Print 'NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Debtors' AND TYPE=12) Update GeneralMaster Set Code='*12019'  Where NAME='Debtors' AND Type=12 ELSE Print 'NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Creditor' AND TYPE=12) Update AccountMaster Set [Group]='*12020' where [Group] IN (Select Code From GeneralMaster Where NAME='Creditor' AND TYPE=12) ELSE Print 'NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Creditor' AND TYPE=12) Update GeneralMaster Set Code='*12020'  Where NAME='Creditor' AND Type=12 ELSE Print 'NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Printers' AND TYPE=12) Update AccountMaster Set [Group]='*12021' where [Group] IN (Select Code From GeneralMaster Where NAME='Printers' AND TYPE=12) ELSE Print 'NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Printers' AND TYPE=12) Update GeneralMaster Set Code='*12021'  Where NAME='Printers' AND Type=12 ELSE Print 'NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Binders' AND TYPE=12) Update AccountMaster Set [Group]='*12001' where [Group] IN (Select Code From GeneralMaster Where NAME='Binders' AND TYPE=12) ELSE Print 'NOT Exist' ;"
   cnDatabase.Execute "IF EXISTS (Select * From GeneralMaster Where NAME='Binders' AND TYPE=12) Update GeneralMaster Set Code='*12001'  Where NAME='Binders' AND Type=12 ELSE Print 'NOT Exist' ;"
'VchSeriesMaster Prifix Update
   cnDatabase.Execute "Update VchSeriesMaster Set Prefix=(Select Top (1)Alias From CompanyMaster Where Alias<>'')+'/'"
   
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='01PF' And Suffix<>'') Where VchType='01PF'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='01PU' And Suffix<>'') Where VchType='01PU'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='01PC' And Suffix<>'') Where VchType='01PC'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='01PJ' And Suffix<>'') Where VchType='01PJ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='02OF' And Suffix<>'') Where VchType='02OF'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='02OU' And Suffix<>'') Where VchType='02OU'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='02OC' And Suffix<>'') Where VchType='02OC'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='02OJ' And Suffix<>'') Where VchType='02OJ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='03TF' And Suffix<>'') Where VchType='03TF'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='03TU' And Suffix<>'') Where VchType='03TU'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='03TC' And Suffix<>'') Where VchType='03TC'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='03TJ' And Suffix<>'') Where VchType='03TJ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='04SF' And Suffix<>'') Where VchType='04SF'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='04SU' And Suffix<>'') Where VchType='04SU'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='04SC' And Suffix<>'') Where VchType='04SC'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='04SJ' And Suffix<>'') Where VchType='04SJ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='05RF' And Suffix<>'') Where VchType='05RF'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='05FR' And Suffix<>'') Where VchType='05FR'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='06IF' And Suffix<>'') Where VchType='06IF'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='06FI' And Suffix<>'') Where VchType='06FI'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='07RF' And Suffix<>'') Where VchType='07RF'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='07FR' And Suffix<>'') Where VchType='07FR'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='08IF' And Suffix<>'') Where VchType='08IF'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='08FI' And Suffix<>'') Where VchType='08FI'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='17PO' And Suffix<>'') Where VchType='17PO'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='18SO' And Suffix<>'') Where VchType='18SO'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='19ST' And Suffix<>'') Where VchType='19ST'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='20JR' And Suffix<>'') Where VchType='20JR'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='21JR' And Suffix<>'') Where VchType='21JR'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='22JR' And Suffix<>'') Where VchType='22JR'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='23PQ' And Suffix<>'') Where VchType='23PQ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='23UZ' And Suffix<>'') Where VchType='23UZ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='23CZ' And Suffix<>'') Where VchType='23CZ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='23JZ' And Suffix<>'') Where VchType='23JZ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='24SQ' And Suffix<>'') Where VchType='24SQ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='24UQ' And Suffix<>'') Where VchType='24UQ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='24CQ' And Suffix<>'') Where VchType='24CQ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='24JQ' And Suffix<>'') Where VchType='24JQ'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='51PI' And Suffix<>'') Where VchType='51PI'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='52PR' And Suffix<>'') Where VchType='52PR'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='53JE' And Suffix<>'') Where VchType='53JE'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='54CE' And Suffix<>'') Where VchType='54CE'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='55CN' And Suffix<>'') Where VchType='55CN'"
    cnDatabase.Execute "Update VchSeriesMaster Set Suffix=(Select Top(1) Suffix From VchSeriesMaster where VchType='56DN' And Suffix<>'') Where VchType='56DN'"
'Update VchName
    cnDatabase.Execute "Update JobworkBVParent Set Name=V.Prefix+ LTRIM(AutoVchNo)+V.Suffix From JobworkBVParent P Inner Join VchSeriesMaster V On V.Code=P.VchSeries Where P.AutoVchNo<>'' AND P.VchSeries IS NOT NULL"
'Update Paper Quality
    cnDatabase.Execute "Delete GeneralMaster where Type=16"
    cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16001' OR Name='Coated Matt') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16001','Coated Matt','Coated Matt','16','0.95','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
    cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002' OR Name='Coated Gloss') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16002','Coated Gloss','Coated Gloss','16','0.9','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
    cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003' OR Name='Uncoated') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16003','Uncoated','Uncoated','16','1.35','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
    cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16004' OR Name='High Bulk') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16004','High Bulk','High Bulk','16','1.4','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Update Vch Series Add VchName
    cnDatabase.Execute "IF COL_LENGTH('VchSeriesMaster', 'VchName') IS NOT NULL PRINT 'Exists' ELSE Alter Table VchSeriesMaster Add VchName nvarchar(40) NOT NULL CONSTRAINT df_VchName DEFAULT '' "
'
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='01PF') Update VchSeriesMaster Set VchName='Purchase' Where VchType='01PF' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='01PU') Update VchSeriesMaster Set VchName='Purchase Unit Cost' Where VchType='01PU' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='01PC') Update VchSeriesMaster Set VchName='Purchase Jobwork Cost' Where VchType='01PC' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='01PJ') Update VchSeriesMaster Set VchName='Purchase Jobwork' Where VchType='01PJ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='02OF') Update VchSeriesMaster Set VchName='Purchase Return' Where VchType='02OF' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='02OU') Update VchSeriesMaster Set VchName='Purchase Return Unit Cost' Where VchType='02OU' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='02OC') Update VchSeriesMaster Set VchName='Purchase Return Jobwork Cost' Where VchType='02OC' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='02OJ') Update VchSeriesMaster Set VchName='Purchase Return Jobwork' Where VchType='02OJ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='03TF') Update VchSeriesMaster Set VchName='Sale Return' Where VchType='03TF' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='03TU') Update VchSeriesMaster Set VchName='Sale Return Unit Cost' Where VchType='03TU' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='03TC') Update VchSeriesMaster Set VchName='Sale Return Jobwork Cost' Where VchType='03TC' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='03TJ') Update VchSeriesMaster Set VchName='Sale Return Jobwork' Where VchType='03TJ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='04SF') Update VchSeriesMaster Set VchName='Sales' Where VchType='04SF' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='04SU') Update VchSeriesMaster Set VchName='Sales Unit Cost' Where VchType='04SU' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='04SC') Update VchSeriesMaster Set VchName='Sales Jobwork Cost' Where VchType='04SC' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='04SJ') Update VchSeriesMaster Set VchName='Sales Jobwork' Where VchType='04SJ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='05RF') Update VchSeriesMaster Set VchName='Purchase Challan IN' Where VchType='05RF' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='05FR') Update VchSeriesMaster Set VchName='Purchase Challan IN (Jobwork)' Where VchType='05FR' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='06IF') Update VchSeriesMaster Set VchName='Purchase Challan Out' Where VchType='06IF' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='06FI') Update VchSeriesMaster Set VchName='Purchase Challan Out (Jobworj)' Where VchType='06FI' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='07RF') Update VchSeriesMaster Set VchName='Sale Challan IN' Where VchType='07RF' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='07FR') Update VchSeriesMaster Set VchName='Sale Challan IN (Jobwork)' Where VchType='07FR' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='08IF') Update VchSeriesMaster Set VchName='Sale Challan Out' Where VchType='08IF' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='08FI') Update VchSeriesMaster Set VchName='Sale Challan Out (Jobwork)' Where VchType='08FI' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='17PO') Update VchSeriesMaster Set VchName='Purchase Order' Where VchType='17PO' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='18SO') Update VchSeriesMaster Set VchName='Sale Order' Where VchType='18SO' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='19ST') Update VchSeriesMaster Set VchName='Stock Tranfer' Where VchType='19ST' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='20JR') Update VchSeriesMaster Set VchName='Stock Genral' Where VchType='20JR' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='21JR') Update VchSeriesMaster Set VchName='Promotional Sale Challan Out' Where VchType='21JR' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='22JR') Update VchSeriesMaster Set VchName='Promotional Purchase Challan Out' Where VchType='22JR' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='23PQ') Update VchSeriesMaster Set VchName='Purchase Quotation' Where VchType='23PQ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='23UZ') Update VchSeriesMaster Set VchName='Purchase Quotation Unit Cost' Where VchType='23UZ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='23CZ') Update VchSeriesMaster Set VchName='Purchase Quotation Jobwork Cost' Where VchType='23CZ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='23JZ') Update VchSeriesMaster Set VchName='Purchase Quotation Jobwork' Where VchType='23JZ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='24SQ') Update VchSeriesMaster Set VchName='Sales Quotation' Where VchType='24SQ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='24UQ') Update VchSeriesMaster Set VchName='Sales Quotation Unit Cost' Where VchType='24UQ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='24CQ') Update VchSeriesMaster Set VchName='Sales Quotation Jobwork Cost' Where VchType='24CQ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='24JQ') Update VchSeriesMaster Set VchName='Sales Quotation Jobwork' Where VchType='24JQ' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='51PI') Update VchSeriesMaster Set VchName='Payments' Where VchType='51PI' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='52PR') Update VchSeriesMaster Set VchName='Receipts' Where VchType='52PR' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='53JE') Update VchSeriesMaster Set VchName='Journal' Where VchType='53JE' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='54CE') Update VchSeriesMaster Set VchName='Countra' Where VchType='54CE' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='55CN') Update VchSeriesMaster Set VchName='Credit Note' Where VchType='55CN' ELSE Print 'VchType NOT Exist'"
   cnDatabase.Execute "IF EXISTS (Select * From VchSeriesMaster Where VchType='56DN') Update VchSeriesMaster Set VchName='Debit Note' Where VchType='56DN' ELSE Print 'VchType NOT Exist'"
End Function
Public Function UpdateMinor05()
'Update Bill No  Field Size
   cnDatabase.Execute "IF Exists (Select BillNo From OutsourceItemPOParent)  Alter Table OutsourceItemPOParent Alter Column BILLNO nvarchar(40) Else PRINT 'NOT Exists'"
'Plate Type
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*24001' OR Name='Deep-etch') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*24001','Deep-etch','Deep-etch','24','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*24002' OR Name='Wipe-on') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*24002','Wipe-on','Wipe-on','24','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*24003' OR Name='PS') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*24003','PS','PS','24','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*24004' OR Name='CTP') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*24004','CTP','CTP','24','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Update Plate Type Value
   cnDatabase.Execute "IF NOT EXISTS (SELECT *FROM GeneralMaster WHERE Code='*24003' OR Name='PS') Print 'NOT Exist' ELSE Update GeneralMaster Set Value1=1 Where Code='*24003'  AND Type=24"
   cnDatabase.Execute "IF NOT EXISTS (SELECT *FROM GeneralMaster WHERE Code='*24004' OR Name='CTP') Print 'NOT Exist' ELSE Update GeneralMaster Set  Value1=1 Where Code='*24004'  AND Type=24"
'Update BP-Process
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07036' OR Name='BP-Unit Cost') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07036','BP-Unit Cost','BP-Unit Cost','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07037' OR Name='BP-Stitching') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07037','BP-Stitching','BP-Stitching','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07038' OR Name='BP-Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07038','BP-Binding','BP-Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07039' OR Name='BP-Folding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07039','BP-Folding','BP-Folding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07040' OR Name='BP-Shrink Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07040','BP-Shrink Packing','BP-Shrink Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07041' OR Name='BP-Box Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07041','BP-Box Packing','BP-Box Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07042' OR Name='BP-Cartage') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07042','BP-Cartage','BP-Cartage','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
End Function
Public Function UpdateMinor06()
'Size CalcMode Update
   cnDatabase.Execute "Update GeneralMaster Set Code='*2000'+Right(Code,1) Where Type=20"
   cnDatabase.Execute "Update AccountChild07 Set CalcMode='*2000'+Right(CalcMode,1) Where CalcMode<>''"
   cnDatabase.Execute "Update BookChild07 Set CalcMode='*2000'+Right(CalcMode,1) Where CalcMode<>''"
   cnDatabase.Execute "Update BookPOChild07 Set CalcMode='*2000'+Right(CalcMode,1) Where CalcMode<>''"
'SizeGroup Update
'    cnDatabase.Execute "Update GeneralMaster Set Name='Small-19'''X26''' Where Name='Small(19''''X26'''')'"
'   cnDatabase.Execute "Update GeneralMaster Set Name='LARGE-23''''X36''''' Where Name='Large(23''''X36'''')'"
'   cnDatabase.Execute "Update GeneralMaster Set Name='LARGE-23''''X36''''-(A/P)' Where Name='Large(23''''X36'''')A/P'"
'   cnDatabase.Execute "Update GeneralMaster Set Name='Medium-20''''X30''''' Where Name='Medium(20''''X30'''')'"
'   cnDatabase.Execute "Update GeneralMaster Set Name='Medium-20''''X30''''(A/P)' Where Name='Medium(20''''X30'''')A/P'"
'   cnDatabase.Execute "Update GeneralMaster Set Name='Extra Large-28''''X40''''-A/P_SPL' Where Name='Extra Large(28''''X40'''')A/P_SPL'"
'   cnDatabase.Execute "Update GeneralMaster Set Name='Extra Large-28''''X40''''-UC_A/P_SPL' Where Name='Extra Large(28''''X40'''')UC_ A/P_SPL'"
   cnDatabase.Execute "Update GeneralMaster Set Name='Web-508mm' Where Name='Web(508mm)'"
   cnDatabase.Execute "Update GeneralMaster Set Name='Web-578mm' Where Name='Web(578mm)'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10022' OR Name='12.00X18.00-Digital') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10022','12.00X18.00-Digital','12.00X18.00-Digital','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Size Update
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01075' OR Name='12.00X18.00 (Digital)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01075','12.00X18.00 (Digital)','12.00X18.00 (Digital)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01076' OR Name='10.00X15.00 (Digital)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01076','10.00X15.00 (Digital)','10.00X15.00 (Digital)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Finish Size Update
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11096' OR Name='08.50X11.00(Digital)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11096','08.50X11.00(Digital)','08.50X11.00(Digital)','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11097' OR Name='07.25X09.50(Digital)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11097','07.25X09.50(Digital)','07.25X09.50(Digital)','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Paper Master Update
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00043') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-70gsm-12.00X18.00in²-(30.48X45.72cm²)-4.877kg-Digital') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00043','Paper-70gsm-12.00X18.00in²-(30.48X45.72cm²)-4.877kg-Digital','Paper-70gsm-12.00X18.00in²-(30.48X45.72cm²)-4.877kg-Digital','S','P','30.48','45.72','12','18','*15004','70','Paper','Digital','4.877','10','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00044') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-70gsm-10.00X15.00in²-(25.40X38.10cm²)-3.387kg-Digital') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00044','Paper-70gsm-10.00X15.00in²-(25.40X38.10cm²)-3.387kg-Digital','Paper-70gsm-10.00X15.00in²-(25.40X38.10cm²)-3.387kg-Digital','S','P','25.4','38.1','10','15','*15004','70','Paper','Digital','3.387','15','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM PaperMaster WHERE Code='*00045') Print 'Code Exist' IF EXISTS (SELECT *FROM PaperMaster WHERE NAME='Paper-70gsm-13.00X20.00in²-(33.02X50.80cm²)-5.871kg-Digital') Print 'Paper Master Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004') IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003') Insert Into PaperMaster VALUES ('*00045','Paper-70gsm-13.00X20.00in²-(33.02X50.80cm²)-5.871kg-Digital','Paper-70gsm-13.00X20.00in²-(33.02X50.80cm²)-5.871kg-Digital','S','P','33.02','50.8','13','20','*15004','70','Paper','Digital','5.871','9','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'Quality Code NOT Exist' ELSE Print 'UOM Code NOT Exist'"
'Color Master Update
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*23000' OR Name='None Color') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*23000','None Color','None Color','23','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Update GeneralMaster (CalcMode)
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20006' OR Name='Per Packet') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20006','Per Packet','Per Packet','20','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20007' OR Name='Per Page') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20007','Per Page','Per Page','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20008' OR Name='Per Paisa Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20008','Per Paisa Inch²','Per Paisa Inch²','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20009' OR Name='Per Box') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20009','Per Box','Per Box','20','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20010' OR Name='Per Bundle') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20010','Per Bundle','Per Bundle','20','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"

'Update GeneralMaster (Operation)
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07036' OR Name='BP-Unit Cost') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07036','BP-Unit Cost','BP-Unit Cost','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07037' OR Name='BP-Stitching') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07037','BP-Stitching','BP-Stitching','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07038' OR Name='BP-Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07038','BP-Binding','BP-Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07039' OR Name='BP-Folding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07039','BP-Folding','BP-Folding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07040' OR Name='BP-Shrink Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07040','BP-Shrink Packing','BP-Shrink Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07041' OR Name='BP-Box Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07041','BP-Box Packing','BP-Box Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07042' OR Name='BP-Cartage') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07042','BP-Cartage','BP-Cartage','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07043' OR Name='Digital Print_1C') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07043','Digital Print_1C','Digital Print_1C','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07044' OR Name='Digital Print_2C') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07044','Digital Print_2C','Digital Print_2C','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07045' OR Name='Digital Print_4C') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07045','Digital Print_4C','Digital Print_4C','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Update BookPOChild07
   cnDatabase.Execute "Update BookPOChild07 Set  Size = NULL Where Size NOT IN ((Select Distinct Size From BookPOChild07 Where Size NOT IN (Select Code From GeneralMaster Where Type=11 Or Type=1)),'')"
   cnDatabase.Execute "IF NOT EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS WHERE TABLE_NAME='GeneralMaster' AND CONSTRAINT_TYPE='PRIMARY KEY') ALTER TABLE GeneralMaster ADD PRIMARY KEY (Code)"
   cnDatabase.Execute "IF NOT EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS WHERE TABLE_NAME='ElementMaster' AND CONSTRAINT_TYPE='PRIMARY KEY') ALTER TABLE ElementMaster ADD PRIMARY KEY (Code)"
   cnDatabase.Execute "Update BookPOChild07 Set  CalcMode ='*00021' Where CalcMode=''"
   cnDatabase.Execute "Update BookPOChild07  Set Size =I.Finishsize From BookPOChild07 C INNER JOIN BookPOParent P On P.Code=C.Code INNER JOIN BookMaster I ON P.Book=I.Code Where C.Size IS NULL Or C.Size =''"
   cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookPOChild07' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'BookPOChild07', 'BookPOChild07T'"
   cnDatabase.Execute "CREATE TABLE dbo.BookPOChild07(Code nvarchar(6) NOT NULL,OrderDate datetime NOT NULL,TargetDate datetime NOT NULL,Element nvarchar(6) NOT NULL,Operation nvarchar(6) NOT NULL,Number decimal(7, 3) NOT NULL,OperationCountName nvarchar(40) NOT NULL DEFAULT ('Nos'),Size nvarchar(6) NULL,Quantity decimal(12, 3) NOT NULL,CalcMode nvarchar(6) NOT NULL,CalcValue decimal(12, 3) NOT NULL,Rate decimal(12, 3) NOT NULL,Amount decimal(12, 2) NOT NULL,Adjustment decimal(12, 2) NOT NULL,[GST%] decimal(4, 2) NOT NULL,GST decimal(12, 2) NOT NULL,BillAmount decimal(12, 2) NOT NULL,Remarks nvarchar(40) NULL,BillNo nvarchar(10) NULL,BillDate datetime NULL,PaidAmount decimal(12, 2) NOT NULL,Status nvarchar(1) NULL,Narration nvarchar(40) NULL,DeliveredQuantityC decimal(12, 0) NOT NULL DEFAULT (0),DeliveredQuantityB decimal(12, 0) NOT NULL DEFAULT (0),BilledMOC decimal(12, 0) NOT NULL DEFAULT (0),BilledMOB decimal(12, 0) NOT NULL DEFAULT (0) " & _
                                      "CONSTRAINT [FK_BookPOChild07_BookPOParent_I] FOREIGN KEY([Code]) REFERENCES [BookPOParent] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE," & _
                                      "CONSTRAINT [FK_BookPOChild07_GeneralMaster_II] FOREIGN KEY([Element]) REFERENCES ElementMaster ([Code])," & _
                                      "CONSTRAINT [FK_BookPOChild07_GeneralMaster_III] FOREIGN KEY([Operation]) REFERENCES GeneralMaster ([Code])," & _
                                      "CONSTRAINT [FK_BookPOChild07_GeneralMaster_IV] FOREIGN KEY([SIZE]) REFERENCES GeneralMaster ([Code])," & _
                                      "CONSTRAINT [FK_BookPOChild07_GeneralMaster_V] FOREIGN KEY([CalcMode]) REFERENCES GeneralMaster ([Code])) ON [PRIMARY]"
                                      'CONSTRAINT PK_BookPOChild07 PRIMARY KEY CLUSTERED (Code,Element,Operation),
    cnDatabase.Execute "INSERT INTO dbo.BookPOChild07 " & _
                                      "Select Code,OrderDate,TargetDate,Element,Operation,Number,'Nos' As OperationCountName,Size,Quantity,CalcMode,(Select Value1 From GeneralMaster Where Code=CalcMode) As CalcValue,Rate,Amount,Adjustment,[GST%],GST,BillAmount,Remarks,BillNo,BillDate,PaidAmount,Status,Narration,DeliveredQuantityC,DeliveredQuantityB,BilledMOC,BilledMOB From BookPOChild07T"
    cnDatabase.Execute "DROP TABLE BookPOChild07T"
'Update BookChild07
    cnDatabase.Execute "Delete FROM BookChild07 Where Code IN (Select Distinct Code From BookChild07 Where Code NOT IN (Select Code From BookMaster))"
    cnDatabase.Execute "Update BookChild07 Set  Size = NULL Where Size IN ((Select Distinct Size From BookChild07 Where Size NOT IN (Select Code From GeneralMaster)),'')"
    cnDatabase.Execute "Update BookChild07 Set  CalcMode ='*00021' Where CalcMode=''"
    cnDatabase.Execute "Update BookChild07  Set Size =I.Finishsize From BookChild07 C INNER JOIN BookMaster I ON C.Code=I.Code Where C.Size IS NULL OR C.Size =''"
    cnDatabase.Execute "IF Not EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookChild07' AND COLUMN_NAME = 'Code') Print 'Col_Not_Exist' ELSE EXEC sp_rename 'BookChild07', 'BookChild07T'"
    cnDatabase.Execute "CREATE TABLE dbo.BookChild07(Code nvarchar(6) NOT NULL,Element nvarchar(6) NOT NULL,Operation nvarchar(6) NOT NULL,Number decimal(7, 3) NOT NULL,OperationCountName nvarchar(40) NOT NULL DEFAULT ('Nos'),Size nvarchar(6) NULL,CalcMode nvarchar(6) NOT NULL,CalcValue decimal(12, 3) NOT NULL,Type char(2) NOT NULL " & _
                                  "CONSTRAINT PK_BookChild07 PRIMARY KEY CLUSTERED (Code,Element,Operation,TYPE), " & _
                                  "CONSTRAINT [FK_BookChild07_BookMaster_I] FOREIGN KEY([Code]) REFERENCES [BookMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE, " & _
                                  "CONSTRAINT [FK_BookChild07_GeneralMaster_II] FOREIGN KEY([Element]) REFERENCES ElementMaster ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild07_GeneralMaster_III] FOREIGN KEY([Operation]) REFERENCES GeneralMaster ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild07_GeneralMaster_IV] FOREIGN KEY([SIZE]) REFERENCES GeneralMaster ([Code]), " & _
                                  "CONSTRAINT [FK_BookChild07_GeneralMaster_V] FOREIGN KEY([CalcMode]) REFERENCES GeneralMaster ([Code])) ON [PRIMARY]"
    cnDatabase.Execute "INSERT INTO dbo.BookChild07 " & _
                                   "Select Code,Element,Operation,Number,'Nos' As OperationCountName,Size,CalcMode,(Select Value1 From GeneralMaster Where Code=CalcMode) As CalcValue,Type From BookChild07T"
    cnDatabase.Execute "DROP TABLE BookChild07T"
End Function
Public Function UpdateMinor07()
If MsgBox("Do You Wants to Update '21.10.21 Version'[Re_write BookPoChild06 ] Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then

End If
'Update ElementMaster FG-2 To FG-6
   cnDatabase.Execute "Update ElementMaster Set Name='FG-1',PrintName='FG-1' Where Code='*00016'"
   cnDatabase.Execute "Update ElementMaster Set ItemType='FG Group' Where Code='*00016'"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00046' OR NAME='FG-2') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00046','FG-2','FG-2','FG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00047' OR NAME='FG-3') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00047','FG-3','FG-3','FG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00048' OR NAME='FG-4') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00048','FG-4','FG-4','FG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00049' OR NAME='FG-5') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00049','FG-5','FG-5','FG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00050' OR NAME='FG-6') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00050','FG-6','FG-6','FG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"

'Update GeneralMaster Type-7 Operations
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07046' OR Name='BP-Center Pin') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07046','BP-Center Pin','BP-Center Pin','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07047' OR Name='BP-CD Pasting') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07047','BP-CD Pasting','BP-Center Pin','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07048' OR Name='BP-Creasing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07048','BP-Creasing','BP-Center Pin','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07049' OR Name='BP-CUTTING') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07049','BP-CUTTING','BP-Center Pin','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07050' OR Name='BP-Die_Cutting') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07050','BP-Die_Cutting','BP-Center Pin','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07051' OR Name='BP-Gathering') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07051','BP-Gathering','BP-Center Pin','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07052' OR Name='BP-Perforation') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07052','BP-Perforation','BP-Center Pin','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07053' OR Name='BP-Section Insert') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07053','BP-Section Insert','BP-Center Pin','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'Update BidingTypeChild
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Die_Cutting') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Die_Cutting'AND BinderyProcess ='*07050')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Die_Cutting'),'*07050')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Die_Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Die_Perforation'AND BinderyProcess ='*07052')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Die_Perforation'),'*07052')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Hard Bound') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Hard Bound'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Hard Bound'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(CD-Insert)'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Spiral Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Spiral Binding'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Spiral Binding'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Spiral Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Spiral Binding'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Spiral Binding'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Spiral Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Spiral Binding'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Spiral Binding'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wirro Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wirro Binding'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wirro Binding'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cutting & Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cutting & Packing'AND BinderyProcess ='*07049')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cutting & Packing'),'*07049')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cutting Only') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cutting Only'AND BinderyProcess ='*07049')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cutting Only'),'*07049')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Half Die Cut') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Half Die Cut'AND BinderyProcess ='*07050')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Half Die Cut'),'*07050')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Loose') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Loose'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Loose'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Pad Gumming') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Pad Gumming'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Pad Gumming'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Pakki Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Pakki Binding'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Pakki Binding'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Kachchi Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Kachchi Binding'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Kachchi Binding'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pinning (BOX)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pinning (BOX)'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pinning (BOX)'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pinning (BOX)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pinning (BOX)'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pinning (BOX)'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pinning (BOX)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pinning (BOX)'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pinning (BOX)'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pin Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pin Binding'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pin Binding'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pin Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pin Binding'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pin Binding'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pin Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pin Binding'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Center Pin Binding'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='None') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='None'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='None'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding (Back Cutting)'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITHOUT SEWING'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Z-PERFECT BINDING 13 FORM-WS') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Z-PERFECT BINDING 13 FORM-WS'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Z-PERFECT BINDING 13 FORM-WS'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wire-O Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wire-O Binding'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wire-O Binding'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Paper Pasting(Craft & Color)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Paper Pasting(Craft & Color)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Paper Pasting(Craft & Color)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding_Cutting_Packing 6-10') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding_Cutting_Packing 6-10'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding_Cutting_Packing 6-10'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding_Cutting_Packing 6-10') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding_Cutting_Packing 6-10'AND BinderyProcess ='*07049')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding_Cutting_Packing 6-10'),'*07049')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-15.50') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-15.50'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-15.50'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-15.50') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-15.50'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-15.50'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-15.50') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-15.50'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-15.50'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Z-ME N MINE FILE UPTO 8') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Z-ME N MINE FILE UPTO 8'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Z-ME N MINE FILE UPTO 8'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='ME N MINE FILE 9 TO 12') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='ME N MINE FILE 9 TO 12'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='ME N MINE FILE 9 TO 12'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)FILE ONLY') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)FILE ONLY'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)FILE ONLY'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)FILE ONLY') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)FILE ONLY'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)FILE ONLY'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)FILE ONLY') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)FILE ONLY'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HARD BOUND-9 (BIOLOGY)FILE ONLY'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding only') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding only'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding only'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='French Posters') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='French Posters'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='French Posters'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding With Sewing(Web Folding)'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding WOUT Sewing(Web Folding)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding WOUT Sewing(Web Folding)'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding WOUT Sewing(Web Folding)'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Gate Fold Cover Book') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Gate Fold Cover Book'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Gate Fold Cover Book'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate (Perfect Binding)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate (Perfect Binding)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate (Perfect Binding)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-A-C@65/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-A-C@65/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-A-C@65/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-01') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-01'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-01'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-Brochure-A-5@10/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-Brochure-A-5@10/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-Brochure-A-5@10/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wire-O') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wire-O'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wire-O'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Cover+Forms/Chapa') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Cover+Forms/Chapa'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Cover+Forms/Chapa'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Cover+Forms/Chapa') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Cover+Forms/Chapa'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Cover+Forms/Chapa'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Cover+Forms/Chapa') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Cover+Forms/Chapa'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-Cover+Forms/Chapa'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad (6pp)Ptg-.30+Bdg-2@2.30/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad (6pp)Ptg-.30+Bdg-2@2.30/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad (6pp)Ptg-.30+Bdg-2@2.30/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Book(Pack Slip)Ptg-4+Bdg-6.5@10.5/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Book(Pack Slip)Ptg-4+Bdg-6.5@10.5/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Book(Pack Slip)Ptg-4+Bdg-6.5@10.5/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Mini OffsetPrinting(Front+Back)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Mini OffsetPrinting(Front+Back)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Mini OffsetPrinting(Front+Back)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Book(Tour Adv)Ptg-1+Bdg-3+1-xtra@5/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Book(Tour Adv)Ptg-1+Bdg-3+1-xtra@5/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Book(Tour Adv)Ptg-1+Bdg-3+1-xtra@5/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(Cash Vouche)Ptg-2+Bdg-2.25@4.25') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(Cash Vouche)Ptg-2+Bdg-2.25@4.25'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(Cash Vouche)Ptg-2+Bdg-2.25@4.25'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(PurOr/Requir)Ptg-2+Bdg-2.25@4.25') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(PurOr/Requir)Ptg-2+Bdg-2.25@4.25'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(PurOr/Requir)Ptg-2+Bdg-2.25@4.25'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(Donation Prof)Ptg-4+Bdg-3.5@7.5') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(Donation Prof)Ptg-4+Bdg-3.5@7.5'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(Donation Prof)Ptg-4+Bdg-3.5@7.5'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(School Order)Ptg-2+Bdg-3.5@5.5') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(School Order)Ptg-2+Bdg-3.5@5.5'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(School Order)Ptg-2+Bdg-3.5@5.5'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(Order Form-Issue/Receipt)Bdg@4/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(Order Form-Issue/Receipt)Bdg@4/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Pad(Order Form-Issue/Receipt)Bdg@4/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Visiting Card Screen Printing@.30/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Visiting Card Screen Printing@.30/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Visiting Card Screen Printing@.30/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me N Mine File Repair') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me N Mine File Repair'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me N Mine File Repair'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Lab Manual Sc-09-10 Sticker@0.16/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Lab Manual Sc-09-10 Sticker@0.16/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Lab Manual Sc-09-10 Sticker@0.16/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-ICSE Lab Man Sc-09-10 Sticker@0.12/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-ICSE Lab Man Sc-09-10 Sticker@0.12/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-ICSE Lab Man Sc-09-10 Sticker@0.12/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Sankalp Cer..te DDNPS Gov..amGzb@4') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Sankalp Cer..te DDNPS Gov..amGzb@4'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Sankalp Cer..te DDNPS Gov..amGzb@4'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Langers Stickers@0.14/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Langers Stickers@0.14/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Langers Stickers@0.14/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Chennai Specimen (Pad)(23x36/6)@8/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Chennai Specimen (Pad)(23x36/6)@8/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Chennai Specimen (Pad)(23x36/6)@8/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Rangoli with CD Sticker@0.01855555/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Rangoli with CD Sticker@0.01855555/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Rangoli with CD Sticker@0.01855555/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-06-WB)'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-07-WB)'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-08-WB)'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-09-WB)'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding(ANB Maths-10-WB)'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate (1f Centre Pin Binding)@0.15/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate (1f Centre Pin Binding)@0.15/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate (1f Centre Pin Binding)@0.15/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB+Corner-Cut+Perforation'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Corner-Cut'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Gumming Sheet@7/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Gumming Sheet@7/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Gumming Sheet@7/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-SHPL Receipt Pad@23.25/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-SHPL Receipt Pad@23.25/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-SHPL Receipt Pad@23.25/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Tour & Travel Bill Pad@7/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Tour & Travel Bill Pad@7/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Tour & Travel Bill Pad@7/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Challan Book(Ptg+Binding)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Challan Book(Ptg+Binding)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Challan Book(Ptg+Binding)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Visiting Cards(NC Gaur)@0.90/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Visiting Cards(NC Gaur)@0.90/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Visiting Cards(NC Gaur)@0.90/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Letterhead Pads(NC Gaur)@66.85/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Letterhead Pads(NC Gaur)@66.85/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Letterhead Pads(NC Gaur)@66.85/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BDG. WITH SEWING+Pouch Packing'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me n Mine File Ist & IInd Term') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me n Mine File Ist & IInd Term'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me n Mine File Ist & IInd Term'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me n Mine File Ist / IInd Term Specimen') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me n Mine File Ist / IInd Term Specimen'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me n Mine File Ist / IInd Term Specimen'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me n Mine File Solution For Specimen') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me n Mine File Solution For Specimen'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Me n Mine File Solution For Specimen'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+CD'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Book(Credit Not)Ptg+Bdg+Paper@40/-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Book(Credit Not)Ptg+Bdg+Paper@40/-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Book(Credit Not)Ptg+Bdg+Paper@40/-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING+Poster Pasting') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING+Poster Pasting'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING+Poster Pasting'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING+Poster Pasting') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING+Poster Pasting'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING+Poster Pasting'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING+Poster Pasting') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING+Poster Pasting'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='CENTRE PIN BINDING+Poster Pasting'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate(Ptg+Binding)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate(Ptg+Binding)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate(Ptg+Binding)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sankalp Pen Screen Ptg') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sankalp Pen Screen Ptg'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sankalp Pen Screen Ptg'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sankalp Pen Screen Ptg') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sankalp Pen Screen Ptg'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sankalp Pen Screen Ptg'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sankalp Pen Screen Ptg') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sankalp Pen Screen Ptg'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sankalp Pen Screen Ptg'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Piano Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Piano Binding'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Piano Binding'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Unit Cost Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Unit Cost Binding'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Sta-Unit Cost Binding'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Cost(Perfect Binding)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Cost(Perfect Binding)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Cost(Perfect Binding)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding PNB Bio+Corner-Cut+Perforation'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Hard Board Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Hard Board Binding'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Hard Board Binding'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(4.25X3.5)24pages') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(4.25X3.5)24pages'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(4.25X3.5)24pages'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)'AND BinderyProcess ='*07037')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)'),'*07037')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PERFECT BINDING WITH SEWING+Shrink(1.75)'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Cost (Center Pin)') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Cost (Center Pin)'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Cost (Center Pin)'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(4.25X3.5)48pages') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(4.25X3.5)48pages'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(4.25X3.5)48pages'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(8.5X11)24pages') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(8.5X11)24pages'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(8.5X11)24pages'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(8.5X11)16pages') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(8.5X11)16pages'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(8.5X11)16pages'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(8.5X11)32pages') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(8.5X11)32pages'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Board Book(8.5X11)32pages'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-02') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-02'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-02'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-03') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-03'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-03'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-04') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-04'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-04'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-05') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-05'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart-05'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart A') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart A'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart A'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart B') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart B'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart B'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart C') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart C'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-My Art Cart C'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle A') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle A'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle A'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle B') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle B'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle B'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle C') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle C'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle C'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-01') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-01'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-01'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-02') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-02'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-02'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-03') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-03'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-03'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-04') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-04'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-04'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-05') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-05'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-05'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-06') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-06'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-06'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-07') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-07'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-07'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-08') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-08'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Unit Rate-Pogo Mad Lets Doodle-08'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-CD Pasting+Box Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-CD Pasting+Box Packing'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-CD Pasting+Box Packing'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-CD Pasting+Box Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-CD Pasting+Box Packing'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-CD Pasting+Box Packing'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-CD Pasting+Box Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-CD Pasting+Box Packing'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Repair-CD Pasting+Box Packing'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cd Pasting + Box Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cd Pasting + Box Packing'AND BinderyProcess ='*07047')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cd Pasting + Box Packing'),'*07047')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cd Pasting + Box Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cd Pasting + Box Packing'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cd Pasting + Box Packing'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Gathering+Shrink pac+Box Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Gathering+Shrink pac+Box Packing'AND BinderyProcess ='*07051')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Gathering+Shrink pac+Box Packing'),'*07051')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Gathering+Shrink pac+Box Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Gathering+Shrink pac+Box Packing'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Gathering+Shrink pac+Box Packing'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Gathering+Shrink pac+Box Packing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Gathering+Shrink pac+Box Packing'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Gathering+Shrink pac+Box Packing'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 1 Fold +Cutting +Creasing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 1 Fold +Cutting +Creasing'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 1 Fold +Cutting +Creasing'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 1 Fold +Cutting +Creasing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 1 Fold +Cutting +Creasing'AND BinderyProcess ='*07049')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 1 Fold +Cutting +Creasing'),'*07049')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 1 Fold +Cutting +Creasing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 1 Fold +Cutting +Creasing'AND BinderyProcess ='*07048')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 1 Fold +Cutting +Creasing'),'*07048')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 2 Fold +Cutting +Creasing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 2 Fold +Cutting +Creasing'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 2 Fold +Cutting +Creasing'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 2 Fold +Cutting +Creasing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 2 Fold +Cutting +Creasing'AND BinderyProcess ='*07049')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 2 Fold +Cutting +Creasing'),'*07049')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 2 Fold +Cutting +Creasing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 2 Fold +Cutting +Creasing'AND BinderyProcess ='*07048')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 2 Fold +Cutting +Creasing'),'*07048')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 3 Fold +Cutting +Creasing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 3 Fold +Cutting +Creasing'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 3 Fold +Cutting +Creasing'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 3 Fold +Cutting +Creasing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 3 Fold +Cutting +Creasing'AND BinderyProcess ='*07049')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 3 Fold +Cutting +Creasing'),'*07049')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 3 Fold +Cutting +Creasing') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 3 Fold +Cutting +Creasing'AND BinderyProcess ='*07048')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folder 3 Fold +Cutting +Creasing'),'*07048')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PB_Sewing- M,C,T,Ch.N,thak,sutli') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PB_Sewing- M,C,T,Ch.N,thak,sutli'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PB_Sewing- M,C,T,Ch.N,thak,sutli'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HB Round-') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HB Round-'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HB Round-'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HB-Flat') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HB-Flat'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='HB-Flat'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfact Bdg. with Full Pasting') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfact Bdg. with Full Pasting'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfact Bdg. with Full Pasting'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding-Manual') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding-Manual'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding-Manual'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding-Machine  Maplitho Paper') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding-Machine  Maplitho Paper'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding-Machine  Maplitho Paper'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding-Machine Art Paper') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding-Machine Art Paper'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Folding-Machine Art Paper'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Hologram Pasting') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Hologram Pasting'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Hologram Pasting'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='VJM - Dharamdoot icncl. All work') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='VJM - Dharamdoot icncl. All work'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='VJM - Dharamdoot icncl. All work'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='UBT CENTREPIN BINDING') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='UBT CENTREPIN BINDING'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='UBT CENTREPIN BINDING'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='BCA Magazine (Incl. All Work) - Per Pcs.') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='BCA Magazine (Incl. All Work) - Per Pcs.'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='BCA Magazine (Incl. All Work) - Per Pcs.'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='NCERT Strap Packing charge') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='NCERT Strap Packing charge'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='NCERT Strap Packing charge'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Side Pin charge') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Side Pin charge'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Side Pin charge'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Side Pin charge') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Side Pin charge'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Side Pin charge'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Side Pin charge') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Side Pin charge'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Side Pin charge'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Shrink Packing Charge') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Shrink Packing Charge'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Shrink Packing Charge'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Shrink Packing Charge') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Shrink Packing Charge'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Shrink Packing Charge'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Shrink Packing Charge') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Shrink Packing Charge'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Shrink Packing Charge'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Farma Bharai charge') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Farma Bharai charge'AND BinderyProcess ='*07053')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Farma Bharai charge'),'*07053')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Chaapa Charge') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Chaapa Charge'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Chaapa Charge'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cover folding  charge') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cover folding  charge'AND BinderyProcess ='*07048')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Cover folding  charge'),'*07048')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding Labour charge') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding Labour charge'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Binding Labour charge'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PB_- M,C,T,Ch.N,thak,sutli') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PB_- M,C,T,Ch.N,thak,sutli'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='PB_- M,C,T,Ch.N,thak,sutli'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch'AND BinderyProcess ='*07039')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch'),'*07039')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch'AND BinderyProcess ='*07038')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch'),'*07038')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch'AND BinderyProcess ='*07040')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch'),'*07040')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch'AND BinderyProcess ='*07041')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Perfect Binding with Perfortion & Punch'),'*07041')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='GROSS LAMINATION') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='GROSS LAMINATION'AND BinderyProcess ='*07001')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='GROSS LAMINATION'),'*07001')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wiero With Perforation') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wiero With Perforation'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='Wiero With Perforation'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding'),'*07036')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding With Centerpin') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding With Centerpin'AND BinderyProcess ='*07046')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding With Centerpin'),'*07046')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding With Centerpin') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding With Centerpin'AND BinderyProcess ='*07046')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding With Centerpin'),'*07046')"
    cnDatabase.Execute "IF NOT EXISTS (SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding Cover Outside') Print 'NOT Exist' Else IF EXISTS (Select Code+BinderyProcess from BindingTypeChild Where Code =(SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding Cover Outside'AND BinderyProcess ='*07036')) Print 'Exist' ELSE Insert Into BindingTypeChild VALUES ((SELECT Code FROM GeneralMaster WHERE Type=6 And Name='File Binding Cover Outside'),'*07036')"
   'Update Account Master
   cnDatabase.Execute "Update AccountMaster Set [Group]='*12002' Where [Group] IN (Select [Group] From AccountMaster Where [Group] NOT IN (Select Code From GeneralMaster))"
End Function
Public Function UpdateMinor08()
   cnDatabase.Execute "ALTER TABLE [JobworkBVParent] ADD [eWayBill] NVARCHAR(40) NULL,[eWayBillDate] [DATETIME] NULL,[ConsigneeName] NVARCHAR(40) NULL,[ConsigneeAddress1] NVARCHAR(40) NULL,[ConsigneeAddress2] NVARCHAR(40) NULL,[ConsigneeAddress3] NVARCHAR(40) NULL,[ConsigneeAddress4] NVARCHAR(40) NULL,[ConsigneeGSTIN] NVARCHAR(40) NULL"
   cnDatabase.Execute "ALTER TABLE [JobworkBVParent] ALTER COLUMN [Consignee] NVARCHAR(6) NULL"
   cnDatabase.Execute "Update JobworkBVParent Set Consignee=Party Where ConsigneeName IS NULL AND Consignee=''"
   cnDatabase.Execute "ALTER TABLE [JobworkBVParent] ADD CONSTRAINT [FK_JobworkBVParent_AccountMaster_I] FOREIGN KEY ([Party]) REFERENCES [AccountMaster] ([Code]),CONSTRAINT [FK_JobworkBVParent_AccountMaster_II] FOREIGN KEY ([Consignee]) REFERENCES [AccountMaster] ([Code]) ON UPDATE CASCADE"
   cnDatabase.Execute "ALTER TABLE [JobworkBVChild] ALTER COLUMN [Item] NVARCHAR(6) NULL"
   cnDatabase.Execute "UPDATE JobworkBVParent SET RecordStatus='O' WHERE LEFT(Type,2) IN ('05','06','07','08') AND RecordStatus NOT IN ('O','M')"
End Function
Public Function UpdateMinor09()
'Update Operation Value1
    cnDatabase.Execute "Update GeneralMaster Set Value1=1 Where Type=7"
    cnDatabase.Execute "Update GeneralMaster Set Value1=0 Where Type=7  And Name Like ('%Folding%')"
    cnDatabase.Execute "Update GeneralMaster Set Value1=0 Where Type=7  And Name Like ('%Sewing%')"
    cnDatabase.Execute "Update GeneralMaster Set Value1=0 Where Type=7  And Name Like ('%Gathering%')"
    cnDatabase.Execute "Update GeneralMaster Set Value1=0 Where Type=7  And Name Like ('%Stitching%')"
End Function
Public Function UpdateMinor10()
    'DropMachine Master
    cnDatabase.Execute "IF EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'MachineMaster' AND COLUMN_NAME = 'Code') DROP TABLE MachineMaster  Else Print 'Col_NOT_Exist'"
    'Machine Master
    cnDatabase.Execute "IF EXISTS(SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'MachineMaster' AND COLUMN_NAME = 'Code') Print 'Col_Exist' ELSE CREATE TABLE dbo.MachineMaster(Code char(6) NOT NULL,Name nvarchar(40) NOT NULL,PrintName nvarchar(40) NOT NULL,Units tinyint NOT NULL,MakeReadyTime tinyint NOT NULL,Efficiency smallint NOT NULL,MinSizeWidth tinyint NOT NULL,MinSizeLength tinyint NOT NULL,MaxSizeWidth tinyint NOT NULL,MaxSizeLength tinyint NOT NULL,StartTime time(0) NOT NULL,EndTime time(0) NOT NULL,Category tinyint NOT NULL,CreatedBy char(6) NOT NULL,CreatedOn datetime NOT NULL,ModifiedBy char(6) NULL,ModifiedOn datetime NULL,RecordStatus char(1) NOT NULL,PrintStatus char(1) NOT NULL CONSTRAINT PK_MachineMaster PRIMARY KEY CLUSTERED (Code))  ON [PRIMARY]"
    'Machine Master PK_MachineMaster
    cnDatabase.Execute "IF Not EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE TABLE_NAME = 'MachineMaster' AND CONSTRAINT_NAME='PK_MachineMaster') ALTER TABLE MachineMaster ADD CONSTRAINT PK_MachineMaster PRIMARY KEY CLUSTERED (Code) ELSE Print 'CONSTRAINT_Exist'"
    'Machine Child
    cnDatabase.Execute "IF EXISTS(SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'MachineChild' AND COLUMN_NAME = 'Code') Print 'Col_Exist' ELSE CREATE TABLE dbo.MachineChild(Code char(6) NOT NULL,Qty smallint NOT NULL,Sets decimal(4, 2) NOT NULL,Hours tinyint NOT NULL,Efficiency smallint NOT NULLCONSTRAINT [FK_MachineChild] FOREIGN KEY([Code]) REFERENCES [MachineMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE) ON [PRIMARY]"
    'MachineChild FK_MachineChild
    cnDatabase.Execute "IF Not EXISTS (SELECT CONSTRAINT_NAME FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE TABLE_NAME = 'MachineChild' AND CONSTRAINT_NAME='FK_MachineChild') ALTER TABLE MachineChild ADD CONSTRAINT [FK_MachineChild] FOREIGN KEY([Code]) REFERENCES [MachineMaster] ([Code]) ON UPDATE CASCADE ON DELETE CASCADE ELSE Print 'CONSTRAINT_Exist'"
    
    cnDatabase.Execute "IF EXISTS (SELECT *FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'BookPOChild0501' AND COLUMN_NAME = 'Code') DROP TABLE BookPOChild0501  Else Print 'Col_NOT_Exist'"
    cnDatabase.Execute "CREATE TABLE dbo.BookPOChild0501(Code nvarchar(15) NOT NULL,Color nvarchar(8) NOT NULL,Machine nvarchar(6) NULL,[Plan] decimal(5, 2) NOT NULL,formsPrinted decimal(5, 2) NOT NULL,platesIssued decimal(5, 2) NOT NULL,paperIssued decimal(12, 2) NOT NULL,SNo int NULL) ON [PRIMARY]"
    cnDatabase.Execute "Insert Into MachineMaster VALUES('*21046','Z-Machine To Be Decide','Z-Machine To Be Decide',0,15,3000,0,0,0,0,'09:00:00','17:30:00',1,'000001','2021-09-21 10:55:08.527','000001','2022-01-13 01:02:44.000','M','N')"
End Function

