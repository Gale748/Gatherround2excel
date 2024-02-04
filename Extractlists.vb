Sub ExtractListsWithHeadings()
    Dim srcDoc As Document
    Dim destDoc As Document
    Dim para As Paragraph
    Dim inList As Boolean
    Dim listRange As Range
    Dim tbl As Table
    Dim aRngHead As Range
    Dim headingText As String
    
    Set srcDoc = ActiveDocument
    Set destDoc = Documents.Add
    
    ' Create a table in the new document
    Set tbl = destDoc.Tables.Add(destDoc.Range, 1, 2)
    tbl.Cell(1, 1).Range.Text = "Heading Level & Text"
    tbl.Cell(1, 2).Range.Text = "List Contents"
    
    inList = False
    
    For Each para In srcDoc.Paragraphs
        If Not para.Range.ListFormat.List Is Nothing Then
            If Not inList Then
                ' Start of a new list
                Set listRange = para.Range
                inList = True
                
                ' Navigate to the nearest heading for the current list
                Set aRngHead = para.Range.GoToPrevious(wdGoToHeading)
                If Not aRngHead Is Nothing Then
                    ' Adjust the range to exclude the paragraph mark
                    aRngHead.End = aRngHead.Paragraphs(1).Range.End - 1
                    ' Combine list formatting (if any) with the heading text
                    headingText = aRngHead.ListFormat.listString & vbTab & aRngHead.Text
                    ' Trim to remove any trailing newline characters
                    headingText = Trim(headingText)
                Else
                    headingText = "No Heading"
                End If
            Else
                ' Extend the range to include the current list item
                listRange.End = para.Range.End
            End If
        ElseIf inList Then
            ' End of a list, so handle it
            tbl.Rows.Add
            tbl.Cell(tbl.Rows.Count, 1).Range.Text = headingText
            listRange.Copy
            tbl.Cell(tbl.Rows.Count, 2).Range.PasteAndFormat (wdFormatOriginalFormatting)
            inList = False ' Reset for the next list
        End If
    Next para
    
    ' Handle case where the document ends with a list
    If inList Then
        tbl.Rows.Add
        tbl.Cell(tbl.Rows.Count, 1).Range.Text = headingText
        listRange.Copy
        tbl.Cell(tbl.Rows.Count, 2).Range.PasteAndFormat (wdFormatOriginalFormatting)
    End If
End Sub