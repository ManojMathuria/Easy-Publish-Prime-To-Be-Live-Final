VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form spreadpreview 
   Caption         =   "FarPoint's Spread 7 Print Preview"
   ClientHeight    =   9180
   ClientLeft      =   195
   ClientTop       =   495
   ClientWidth     =   19890
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9180
   ScaleWidth      =   19890
   WindowState     =   2  'Maximized
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   19815
      _Version        =   524288
      _ExtentX        =   34951
      _ExtentY        =   873
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Form2.frx":0000
      AppearanceStyle =   0
   End
   Begin FPSpreadADO.fpSpreadPreview fpSpreadPreview1 
      Height          =   10635
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   19845
      _Version        =   524288
      _ExtentX        =   35004
      _ExtentY        =   18759
      _StockProps     =   96
      AllowUserZoom   =   -1  'True
      GrayAreaColor   =   16121836
      GrayAreaMarginH =   720
      GrayAreaMarginType=   0
      GrayAreaMarginV =   720
      PageBorderColor =   8388608
      PageBorderWidth =   2
      PageShadowColor =   0
      PageShadowWidth =   2
      PageViewPercentage=   100
      PageViewType    =   1
      ScrollBarH      =   1
      ScrollBarV      =   1
      ScrollIncH      =   360
      ScrollIncV      =   360
      PageMultiCntH   =   1
      PageMultiCntV   =   1
      PageGutterH     =   -1
      PageGutterV     =   -1
      ScriptEnhanced  =   0   'False
   End
End
Attribute VB_Name = "spreadpreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frm As Form
Private Sub Form_Activate()
 
    'Attach preview control to Spread
    spreadpreview.fpSpreadPreview1.hWndSpread = frm.fpSpread1.hwnd
    'Update page count listing
    UpdatePageCount
End Sub
Private Sub Form_Load()
Me.Caption = frm.Caption
    SetupToolbar
    
    'Disable Previous button
    DisableButton 4, "LEFT"
        
    'Get the zoom display
    GetZoom zoomindex
    
    'Set up page numbering
    If frm.fpSpread1.PrintPageCount = 1 Then
        'Disable Next button if only one page
        DisableButton 2, "LEFT"
    End If
           
End Sub
Sub SetupToolbar()
Dim i As Integer

    'Specify whether Edit Mode is to remain on when switching between cells
    fpSpread1.EditModePermanent = True

    fpSpread1.Col = -1
    fpSpread1.Row = -1
    fpSpread1.Lock = True
    
    'Set the number of rows in the spreadsheet
    fpSpread1.MaxRows = 1
 
    'Set the height of a selected row
    fpSpread1.RowHeight(1) = 15
   
    'Set the number of columns in the spreadsheet
    fpSpread1.MaxCols = 17
 
    'Set the column widths
    For i = 1 To fpSpread1.MaxCols Step 2
        fpSpread1.ColWidth(i) = 0.3
    Next i
   
    'Resize wide column
    fpSpread1.ColWidth(14) = 15
    
    'Show or hide the column headers
    fpSpread1.DisplayColHeaders = False
    fpSpread1.DisplayRowHeaders = False
    
    'Turn off scroll bars
    fpSpread1.ScrollBars = ScrollBarsNone
    
    'Turn off border
    fpSpread1.BorderStyle = BorderStyleNone
      
    'Select row(s)
    fpSpread1.Row = 1
    fpSpread1.Col = -1

    'Determine the color of background, foreground and border color
    fpSpread1.ForeColor = RGB(0, 0, 0)
    fpSpread1.BackColor = RGB(192, 192, 192)
    fpSpread1.fontname = "MS Sans Serif"
    fpSpread1.FontSize = 8
    fpSpread1.FontBold = False

    'Select a single cell
    fpSpread1.Col = 2
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Next"
    Set fpSpread1.TypeButtonPicture = LoadPicture(App.Path & "\Icon\RIGHT.BMP")
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_LEFT
    
    'Select a single cell
    fpSpread1.Col = 4
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Previous"
    Set fpSpread1.TypeButtonPicture = LoadPicture(App.Path & "\Icon\LEFT.BMP")
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_RIGHT
    
    'Select a single cell
    fpSpread1.Col = 6
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Zoom"
    Set fpSpread1.TypeButtonPicture = LoadPicture(App.Path & "\Icon\ZOOM.BMP")
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_RIGHT
    
    'Select a single cell
    fpSpread1.Col = 8
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Print"
    Set fpSpread1.TypeButtonPicture = LoadPicture(App.Path & "\Icon\PRINT.BMP")
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_RIGHT
    
    'Select a single cell
    fpSpread1.Col = 10
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Setup"
    Set fpSpread1.TypeButtonPicture = LoadPicture(App.Path & "\Icon\SETUP.BMP")
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_RIGHT
    
    
    'Select a single cell
    fpSpread1.Col = 16
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Close"
    Set fpSpread1.TypeButtonPicture = LoadPicture(App.Path & "\Icon\CLOSE.BMP")
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_RIGHT
    fpSpread1.TextTip = TextTipFloating
    Dim bRet As Boolean
    bRet = fpSpread1.SetTextTipAppearance("MS Sans Serif", 8, 0, 0, &HC0FFFF, &H0)
    fpSpread1.CursorType = CursorTypeLockedCell
    fpSpread1.CursorStyle = CursorStyleArrow
    fpSpread1.NoBeep = True
End Sub
Sub DisableButton(Col As Long, bitmapdirection As String)
'Disable specified button
    fpSpread1.Redraw = False
    
    fpSpread1.Row = 1
    fpSpread1.Col = Col
    
    fpSpread1.Lock = True
    fpSpread1.TypeButtonTextColor = RGB(128, 128, 128)
    fpSpread1.Protect = True
    Set fpSpread1.TypeButtonPicture = LoadPicture(App.Path & "\Icon\" & bitmapdirection & "DIS.BMP")
    
    fpSpread1.Redraw = True
End Sub
Sub EnableButton(Col As Long, bitmapdirection As String)
'Enable specified button
    fpSpread1.Redraw = False
    
    fpSpread1.Row = 1
    fpSpread1.Col = Col
    
    fpSpread1.Lock = False
    fpSpread1.TypeButtonTextColor = RGB(0, 0, 0)
    fpSpread1.Protect = False
    Set fpSpread1.TypeButtonPicture = LoadPicture(App.Path & "\Icon\" & bitmapdirection & ".BMP")
    
    fpSpread1.Redraw = True
End Sub
Private Sub Form_Resize()
    fpSpread1.Move 0, 0, ScaleWidth, fpSpread1.Height
    fpSpreadPreview1.Move 0, fpSpread1.Height, ScaleWidth, ScaleHeight - fpSpread1.Height
End Sub

Private Sub fpSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    fpSpread1.Col = Col
    fpSpread1.Row = Row
    
    If fpSpread1.CellType = CellTypeButton Then
        Select Case Col
            Case 2  'Next
                If fpSpreadPreview1.PageCurrent < frm.fpSpread1.PrintPageCount Then
                    fpSpreadPreview1.PageCurrent = fpSpreadPreview1.PageCurrent + fpSpreadPreview1.PagesPerScreen
                    EnableButton Col, "RIGHT"
                    'Enable Previous button
                    EnableButton 4, "LEFT"
                   'Update page count listing
'                    UpdatePageCount
                End If
                
                 'If at last page, disable button
                    If fpSpreadPreview1.PageCurrent >= frm.fpSpread1.PrintPageCount Then
                        DisableButton Col, "RIGHT"
                    End If
            Case 4  'Previous
                If fpSpreadPreview1.PageCurrent > 1 Then
                    fpSpreadPreview1.PageCurrent = fpSpreadPreview1.PageCurrent - fpSpreadPreview1.PagesPerScreen
                    EnableButton Col, "LEFT"
                    EnableButton 2, "RIGHT"
                    'Update page count listing
'                    UpdatePageCount
                End If
                
                'If at first page, disable button
                If fpSpreadPreview1.PageCurrent = 1 Then
                    DisableButton Col, "LEFT"
                End If
                
            Case 6  'Zoom
                fpSpreadPreview1.ZoomState = 3
                
            Case 8  'Print
                PrintDlg.Show
                                 
            Case 10 'Setup
                pagesetup.Show 1
             
            Case 16 'Close
                Unload Me
        End Select
    End If
End Sub
Sub UpdatePageCount()
 'Page Count
    fpSpread1.Row = 1
    fpSpread1.Col = 14
    fpSpread1.Text = "Page " & fpSpreadPreview1.PageCurrent & " of " & frm.fpSpread1.PrintPageCount
End Sub
Private Sub fpSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As FPSpreadADO.TextTipFetchMultilineConstants, TipWidth As Long, TipText As String, ShowTip As Boolean)
    With fpSpread1
        .Col = Col
        .Row = Row
        If .CellType = CellTypeButton And Not .Lock Then
            ShowTip = True
            TipText = .TypeButtonText
        ElseIf .CellType = CellTypeEdit And .Text <> "" Then
            ShowTip = True
            TipText = .Text
        End If
    End With
End Sub
Private Sub fpSpreadPreview1_PageChange(ByVal Page As Long)
    UpdatePageCount
End Sub
