Sub add_progress_bar()
    ' 1.change strings to your own part's name
    ' 2.change pages to your own part's page number
    ' 3.change barHeight,linewidth and color to your own style
    ' 4.run this script
    Dim strings As Variant
    Dim pages As Variant
    Dim X As Integer
    Dim s As Shape
    Dim t As Shape
    Dim slideIndex As Integer
    Dim slideCount As Integer
    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim barWidth As Single
    Dim barHeight As Single
    Dim progressbarWidth As Single
    Dim progressbarHeight As Single
    Dim linewidth As Single
    Dim textBoxName As String
    Dim partStart As Single
    Dim partEnd As Single

    ' Content names
    strings = Array("Part 1", "Part 2", "Part 3")
    ' Page number array
    pages = Array(2, 5, 7)
    ' Bar height
    barHeight = 20
    linewidth = 2
    ' Error handling
    On Error Resume Next
    ' Get the total number of slides, slide width, and slide height
    slideCount = ActivePresentation.Slides.Count
    slideWidth = ActivePresentation.PageSetup.slideWidth
    slideHeight = ActivePresentation.PageSetup.slideHeight
    ' Iterate through all pages
    ' you can set Index=2 To slideCount-1 to ignore the first and last page 
    For slideIndex = 1 To slideCount
        ''' Delete old rectangles '''
        For X = ActivePresentation.Slides(slideIndex).Shapes.Count To 1 Step -1
            Set s = ActivePresentation.Slides(slideIndex).Shapes(X)
            If s.Name Like "PB*" Or s.Name Like "text*" Or s.Name Like "ProgressPB*" Then s.Delete
        Next X
        ''' Add fixed rectangles and text box '''
        For X = LBound(pages) To UBound(pages)
            ''' Determine the start and end positions of the current part '''
            If X = LBound(pages) Then
                partStart = 0
            Else
                partStart = pages(X - 1) * slideWidth / slideCount
            End If
            If X = UBound(pages) Then
                partEnd = slideWidth
            Else
                partEnd = pages(X) * slideWidth / slideCount
            End If
            ' Calculate the width of the rectangle
            barWidth = partEnd - partStart
            
            ''' Add Rectangle Of Each Part '''
            Set s = ActivePresentation.Slides(slideIndex).Shapes.AddShape(msoShapeRectangle, _
                partStart, slideHeight - barHeight, barWidth, barHeight) ' set rectangle
            ' set color
            s.Fill.ForeColor.RGB = RGB(0, 151, 218)
            ' set edge line
            s.Line.ForeColor.RGB = RGB(200, 200, 200)
            s.Name = "PB" & X ' set name of rectangles
            s.Shadow.Visible = msoFalse
            
            ''' Add Text Box Of Each Part '''
            ' Determine the name of the text box
            textBoxName = "text" & (X + 1)
            ' Add a text box and set it to center
            Set t = ActivePresentation.Slides(slideIndex).Shapes.AddTextbox(msoTextOrientationHorizontal, _
                partStart, slideHeight - 10, barWidth, 10)
            t.TextFrame.TextRange.Text = strings(X)
            t.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
            t.TextFrame.VerticalAnchor = msoAnchorMiddle
            t.Line.Visible = msoFalse ' Hide the border of the text box
            t.Name = textBoxName ' Rename the text box 
            With t.TextFrame.TextRange.Font
                .Name = "Arial" 
                .Size = 10      
            End With
            t.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            ' Adjust the position of the text box to align with the center of the rectangle
            t.Top = s.Top + (s.Height - t.Height) / 2
            t.Left = s.Left + (s.Width - t.Width) / 2
            ' Bring text to front
            t.ZOrder msoBringToFront
        Next X
        ''' add active rectangle progress bar '''
        For X = LBound(pages) To UBound(pages)
            ''' Determine the start and end positions of the current part '''
            If X = LBound(pages) Then
                partStart = 0
            Else
                partStart = pages(X - 1) * slideWidth / slideCount
            End If
            If X = UBound(pages) Then
                partEnd = slideWidth
            Else
                partEnd = pages(X) * slideWidth / slideCount
            End If
            ' Calculate the width of the active progress bar'
            If slideIndex < pages(X) Then
                barWidth = slideIndex * slideWidth / slideCount - partStart
            Else
                barWidth = partEnd - partStart
            End If
            ''' add active progress bar '''
            Set s = ActivePresentation.Slides(slideIndex).Shapes.AddShape(msoShapeRectangle, _
                partStart + linewidth / 2, slideHeight - barHeight + linewidth / 2, barWidth - linewidth, barHeight - linewidth) ' 高度为20
            s.Fill.ForeColor.RGB = RGB(55, 96, 146) ' set color
            s.Name = "ProgressPB" & X ' set name of progress bar
            ' hide the shadow and edge lines
            s.Line.Visible = msoFalse
            s.Shadow.Visible = msoFalse
            ' bring text to front
            For i = ActivePresentation.Slides(slideIndex).Shapes.Count To 1 Step -1
                Set t = ActivePresentation.Slides(slideIndex).Shapes(i)
                If t.Name Like "text*" Then t.ZOrder msoBringToFront
            Next i
        Next X
    Next slideIndex
End Sub
