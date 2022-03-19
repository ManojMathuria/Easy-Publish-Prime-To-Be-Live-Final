Attribute VB_Name = "Module2"
Global zoomindex As Integer
Sub GetZoom(zoomlabel As Integer)
'Set up the print previews zoom

        Select Case zoomlabel
            Case 0
                spreadpreview.fpSpreadPreview1.PageViewType = 0 'Page Width
'                spreadpreview.fpSpreadPreview1.PageViewType = 1 'Page Width
            Case 1
                spreadpreview.fpSpreadPreview1.PageViewType = 3 'Page Hight
            Case 2
                spreadpreview.fpSpreadPreview1.PageViewType = 4 'Whole Page
            Case 3
                spreadpreview.fpSpreadPreview1.PageViewType = 5 'Two Page
                spreadpreview.fpSpreadPreview1.PageMultiCntH = 2 'Two Page
                spreadpreview.fpSpreadPreview1.PageMultiCntV = 1 'Two Page
            Case 4
                spreadpreview.fpSpreadPreview1.PageViewType = 5 'Three Page
                spreadpreview.fpSpreadPreview1.PageMultiCntH = 3 'Three Page
                spreadpreview.fpSpreadPreview1.PageMultiCntV = 1 'Three Page
            Case 5
                spreadpreview.fpSpreadPreview1.PageViewType = 5 'Four Page
                spreadpreview.fpSpreadPreview1.PageMultiCntH = 2 'Four Page
                spreadpreview.fpSpreadPreview1.PageMultiCntV = 2 'Four Page
            Case 6
                spreadpreview.fpSpreadPreview1.PageViewType = 5 'Six Page
                spreadpreview.fpSpreadPreview1.PageMultiCntH = 3 'Six Page
                spreadpreview.fpSpreadPreview1.PageMultiCntV = 2 'Six Page
            Case 7
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 300
            Case 8
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 275
            Case 9
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 250
            Case 10
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 225
            Case 11
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 200
            Case 12
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 175
            Case 13
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 150
            Case 14
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 125
            Case 15
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 100
            Case 16
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 75
            Case 17
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 50
            Case 18
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 25
            Case 19
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 10
        End Select
      
End Sub

