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
        ' Skip TOC and Foreword based on heading text or position
        If para.Range.Text Like "Table of Contents*" Or para.Range.Text Like "Foreword*" Then
            ' Skip this paragraph and move to the next one
            GoTo NextParagraph
        End If

        ' Check if the paragraph is part of a list and not part of a tracked change
        If Not para.Range.ListFormat.List Is Nothing And Not para.Range.Revisions.Count > 0 Then
            If Not inList Then
                ' Start of a new list
                Set listRange = para.Range
                inList = True
                
                ' Navigate to the nearest heading for the current list
                Set aRngHead = para.Range.GoToPrevious(wdGoToHeading)
                If Not aRngHead Is Nothing Then
                    aRngHead.End = aRngHead.Paragraphs(1).Range.End - 1
                    headingText = aRngHead.ListFormat.ListString & vbTab & aRngHead.Text
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
NextParagraph:
    Next para
    
    ' Handle case where the document ends with a list
    If inList Then
        tbl.Rows.Add
        tbl.Cell(tbl.Rows.Count, 1).Range.Text = headingText
        listRange.Copy
        tbl.Cell(tbl.Rows.Count, 2).Range.PasteAndFormat (wdFormatOriginalFormatting)
    End If
End Sub
