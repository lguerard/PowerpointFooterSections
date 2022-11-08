Sub secfoot()
Dim str_Subsection As String

Dim b_found As Boolean
With ActivePresentation
    If .SectionProperties.Count > 0 Then
    
        SectionCount = .SectionProperties.Count - 1
        si_orig = .Slides(2).sectionIndex
        current_count = 1
        
        For X = 2 To .Slides.Count - 1
            .Slides(X).HeadersFooters.Footer.Visible = True
            
            si = .Slides(X).sectionIndex
            If si <> si_orig Then
                current_count = 1
                si_orig = si
            End If
            SectionSlidesCount = .SectionProperties.SlidesCount(si)
        
            For Each oshp In .Slides(X).Shapes
                If oshp.Type = msoPlaceholder Then
                    If oshp.PlaceholderFormat.Type = ppPlaceholderFooter Then _
                    oshp.TextFrame.TextRange = ActivePresentation.SectionProperties.Name(.Slides(X).sectionIndex) & " - " & current_count & "/" & SectionSlidesCount
                End If
            Next oshp
            current_count = current_count + 1
        Next X:
    End If
End With
End Sub
