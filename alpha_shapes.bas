Attribute VB_Name = "alpha_shapes"
'https://www.ozgrid.com/forum/index.php?thread/140336-shape-manager/
    'https://www.ozgrid.com/VBA/shapes.htm

Sub GetShapePropertiesAllWs()
    Dim sShapes As Shape, lLoop As Long
    Dim WsNew As Worksheet
    Dim wsLoop As Worksheet
    ''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''LIST PROPERTIES OF SHAPES'''''''''''''

    ''''''''''Dave Hawley www.ozgrid.com''''''''''''

    ''''''''''''''''''''''''''''''''''''''''''''''''
    Set WsNew = Sheets.Add
    
    'Add headings for our lists. Expand as needed
        WsNew.Range("A1:G1") = Array("Shape Name", "Shape Type", "Height", "Width", "Left", "Top", "Sheet Name")
    'Loop through all Worksheet
        For Each wsLoop In Worksheets
            'Loop through all shapes on Worksheet
                For Each sShapes In wsLoop.Shapes
                    'Increment Variable lLoop for row numbers
                        lLoop = lLoop + 1
                        With sShapes
                            'Add shape properties
                                WsNew.Cells(lLoop + 1, 1) = .Name
                                WsNew.Cells(lLoop + 1, 2) = .OLEFormat.Object.Name
                                WsNew.Cells(lLoop + 1, 3) = .Height
                                WsNew.Cells(lLoop + 1, 4) = .Width
                                WsNew.Cells(lLoop + 1, 5) = .Left
                                WsNew.Cells(lLoop + 1, 6) = .Top
                            'Follow the same pattern for more
                                WsNew.Cells(lLoop + 1, 7) = wsLoop.Name
                        End With
                Next sShapes
        Next wsLoop
    'AutoFit Columns.
        WsNew.Columns.AutoFit
End Sub

'other
    'https://www.thespreadsheetguru.com/blog/how-to-keep-track-of-your-shapes-created-with-vba-code
    'https://turbofuture.com/computers/Renaming-Reordering-and-Grouping-shapes-in-Excel-2007-and-Excel-2010
