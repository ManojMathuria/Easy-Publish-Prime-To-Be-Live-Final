Attribute VB_Name = "Module3"
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Public Const GWL_WNDPROC = (-4)
'Public lpPrevWndProc As Long
'Const WM_MOUSEWHEEL = &H20A
'Const WHEEL_DELTA = 120
'Dim Count As Integer
'Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    If Msg = WM_MOUSEWHEEL Then
'        Dim Delta As Long
'        Static Travel As Long
'        Delta = HiWord(wParam)
'        Travel = Travel + Delta
'        MouseWheel Travel \ WHEEL_DELTA, LoWord(lParam), HiWord(lParam)
'        Travel = Travel Mod WHEEL_DELTA
'    End If
'    WndProc = CallWindowProc(lpPrevWndProc, hWnd, Msg, wParam, lParam)
'End Function
'Sub MouseWheel(Travel As Integer, X As Long, Y As Long)
'    Count = Count + Travel
'    FrmAccountMaster.Cls
'    FrmAccountMaster.Print "Travel=" & Count, "X=" & X, "Y=" & Y
'End Sub
'Function HiWord(DWord As Long) As Integer
'    CopyMemory HiWord, ByVal VarPtr(DWord) + 2, 2
'End Function
'Function LoWord(DWord As Long) As Integer
'    CopyMemory LoWord, DWord, 2
'End Function
