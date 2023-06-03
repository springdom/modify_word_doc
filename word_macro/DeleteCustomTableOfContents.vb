Sub DeleteFirstImageAndTableOfContents()
    Dim img As InlineShape
    Dim shp As Shape
    Dim rng As Range
    Dim TOC As TableOfContents
    
    ' Delete the first image
    For Each img In ActiveDocument.InlineShapes
        If img.Type = wdInlineShapePicture Then
            img.Delete
            Exit For ' Exit the loop after deleting the first image
        End If
    Next img
    
    ' Delete shapes that may contain the table of contents
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoTextBox Or shp.Type = msoTextFrame Then
            shp.Delete
        End If
    Next shp
    
    ' Delete text ranges that may contain the table of contents
    For Each rng In ActiveDocument.StoryRanges
        With rng.Find
            .ClearFormatting
            .Text = "Table of Contents"
            .MatchWholeWord = True
            .MatchCase = True
            .Wrap = wdFindStop
            If .Execute Then
                rng.Delete
            End If
        End With
    Next rng
    
    ' Delete tables of contents
    For Each TOC In ActiveDocument.TablesOfContents
        TOC.Delete
    Next TOC
End Sub

