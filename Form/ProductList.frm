VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmProductList 
   BackColor       =   &H8000000B&
   Caption         =   "Product List"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleMode       =   0  'User
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Mh3dFrame1 
      FillColor       =   &H8000000F&
      FontTransparent =   0   'False
      Height          =   10935
      Left            =   0
      ScaleHeight     =   10875
      ScaleWidth      =   20235
      TabIndex        =   0
      Top             =   0
      Width           =   20295
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2640
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   100
         ImageHeight     =   100
         MaskColor       =   128
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":1889
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":28E04
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":2B406
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":2CD0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":2D05E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":2EEC1
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":313FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":517A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":538A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":54CE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProductList.frx":59508
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   9120
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   19995
         _ExtentX        =   35269
         _ExtentY        =   16087
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Product List"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   24
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8370
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "FrmProductList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim List As ListItem
Set ListView1.ColumnHeaderIcons = ImageList1
With ListView1.ColumnHeaders
.Add , , "              Cataegary", 10000, lvwColumnLeft, 0        'Width / 1
.Add , , "              PRODUCTS LIST", 10000, lvwColumnCenter, 0
'.Add , , " Finish Size", Width / 3, lvwColumnCenter, 2
'.Add , , " Details", Width / 3, lvwColumnCenter, 2
End With

Set ListView1.SmallIcons = ImageList1

Set List = ListView1.ListItems.Add(, , "    Booklets", , 1)
List.SubItems(1) = "Self Cover Booklets / Booklets with Cover"

Set List = ListView1.ListItems.Add(, , "    Books", , 4)
List.SubItems(1) = "Paper Back Books"

Set List = ListView1.ListItems.Add(, , "    Hard Bound Books", , 2)
List.SubItems(1) = "Hard Case Books"

Set List = ListView1.ListItems.Add(, , "    Hard Bound Books with Dust Cover", , 2)
List.SubItems(1) = "Hard Case Books with Jacket"

Set List = ListView1.ListItems.Add(, , "    Flyer", , 6)
List.SubItems(1) = "Folder/Leaflet/Pamplets"

Set List = ListView1.ListItems.Add(, , "    Flyer", , 9)
List.SubItems(1) = "Flyer"

Set List = ListView1.ListItems.Add(, , "    Tag", , 7)
List.SubItems(1) = "Tag"

Set List = ListView1.ListItems.Add(, , "    Label", , 8)
List.SubItems(1) = "Label"

Set List = ListView1.ListItems.Add(, , "    Dangler", , 12)
List.SubItems(1) = "Dangler"

Set List = ListView1.ListItems.Add(, , "    Backing Sheet", , 11)
List.SubItems(1) = "Backing Sheet"

Set List = ListView1.ListItems.Add(, , "    Bunting", , 10)
List.SubItems(1) = "Bunting"

'List.SubItems(1) = "BookLet"
'List.SubItems(2) = "Finish Size"
'List.SubItems(3) = "Details"
ListView1.AllowColumnReorder = True
ListView1.FlatScrollBar = False
ImageList1.ImageWidth = 100
ImageList1.ImageHeight = 100
'ListView1.TextBackground = lvwOpeque
'ListView1.BackColor = vbGreen
ListView1.Font.Size = 20
ListView1.Font.Bold = True
ListView1.Font.Name = "Bookman Old Style"
ListView1.ForeColor = vbBlue
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.Sorted = True
If ListView1.SortOrder = lvwAscending Then
ListView1.SortOrder = lvwDescending
Else
ListView1.SortOrder = lvwAscending
End If
End Sub

