Sub GatherRoundEnhanced()
    ' Define constants for search text and column titles for maintainability
    Const SearchText As String = "[Red]"
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Text"
    
    Dim srcDoc As Document, destDoc As Document
    Dim aTbl As Table
    Dim aRng As Range, aRngHead As Range
    Dim lastHeadingText As String, currentHeadingText As String
    Dim sNum As String
    Dim aRow As Row
    
    Set srcDoc = ActiveDocument
    Set destDoc = Documents.Add
    
    ' Create a table in the new document with predefined titles
    Set aTbl = destDoc.Tables.Add(destDoc.Range, 1, 2)
    aTbl.Cell(1, 1).Range.Text = HeadingColumnTitle
    aTbl.Cell(1, 2).Range.Text = TextColumnTitle
    
    Set aRng = srcDoc.Range
    With aRng.Find
        .ClearFormatting
        .Text = SearchText
        .Forward = True
        Do While .Execute
            aRng.Start = aRng.Paragraphs(1).Range.Start
            Set aRow = aTbl.Rows.Add
            
            ' Handle list formatting if present
            If aRng.ListFormat.ListType <> wdListNoNumbering Then
                sNum = aRng.ListFormat.ListString
                aRow.Cells(2).Range.Text = sNum & vbTab & aRng.Text
            Else
                aRow.Cells(2).Range.FormattedText = aRng.FormattedText
            End If
            
            ' Use GoToPrevious to find the nearest heading only if needed
            Set aRngHead = aRng.GoToPrevious(wdGoToHeading)
            If Not aRngHead Is Nothing Then
                aRngHead.End = aRngHead.Paragraphs(1).Range.End - 1
                currentHeadingText = aRngHead.ListFormat.ListString & vbTab & Trim(aRngHead.Text)
                
                ' Check if the current heading is different from the last to minimize redundancy
                If currentHeadingText <> lastHeadingText Then
                    lastHeadingText = currentHeadingText
                End If
            Else
                currentHeadingText = "No Heading"
            End If
            aRow.Cells(1).Range.Text = currentHeadingText
            
            ' Prepare for the next search
            aRng.Collapse Direction:=wdCollapseEnd
            aRng.End = srcDoc.Range.End
        Loop
    End With
End Sub