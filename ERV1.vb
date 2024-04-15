Sub GatherRoundEnhancedToExcelOptimized()
    ' Define constants for search text and column titles for maintainability
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Text"
    Dim SearchTerms As Variant
    SearchTerms = Array("[R]", "#SPM", "#SPAM", "#ETC")  ' Array of search terms
    
    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim lastHeadingText As String, currentHeadingText As String
    Dim lastPosition As Long
    Dim aRng As Range, aRngHead As Range
    Dim iRow As Long
    Dim searchTerm As Variant
    Dim paragraphText As String
    
    Set srcDoc = ActiveDocument
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlApp.Visible = True  ' Excel visibility for debugging
    
    xlSheet.Cells(1, 1).Value = HeadingColumnTitle
    xlSheet.Cells(1, 2).Value = TextColumnTitle
    iRow = 2 ' Start from the second row for data
    lastPosition = 0
    
    For Each searchTerm In SearchTerms
        Set aRng = srcDoc.Range(Start:=lastPosition, End:=srcDoc.content.End)
        With aRng.Find
            .ClearFormatting
            .Text = searchTerm
            .Forward = True
            .Wrap = wdFindStop
            
            Do While .Execute
                If aRng.Start <= lastPosition Then
                    Exit Do
                End If

                ' Expand the range to the whole paragraph to capture complete context
                aRng.Expand Unit:=wdParagraph
                paragraphText = aRng.Text

                ' Ensure we don't repeat the same paragraph for different search terms
                If aRng.Start >= lastPosition Then
                    Set aRngHead = aRng.GoToPrevious(wdGoToHeading)
                    If Not aRngHead Is Nothing Then
                        aRngHead.End = aRngHead.Paragraphs(1).Range.End - 1
                        currentHeadingText = aRngHead.ListFormat.listString & vbTab & Trim(aRngHead.Text)
                    Else
                        currentHeadingText = "No Heading"
                    End If

                    xlSheet.Cells(iRow, 1).Value = currentHeadingText
                    xlSheet.Cells(iRow, 2).Value = paragraphText
                    iRow = iRow + 1
                    lastPosition = aRng.End ' Update last position after processing
                End If

                aRng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
    Next searchTerm
    
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True  ' Make sure Excel is visible after processing
End Sub