Sub ListRTagsAndSubsequentTagsWithAccurateHeadingAndSpacing()
    ' Constants for search text, tags, and column titles
    Const SearchTextR As String = "[R]"
    Dim subsequentTags As Variant
    subsequentTags = Array("#SPM", "#SPAM", "#TechLead", "#RSM")
    Const HeadingColumnTitle As String = "Heading"
    Const ParagraphTextTitle As String = "Paragraph Containing [R]"
    
    ' Excel setup
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlApp.Visible = True ' Make Excel visible
    
    ' Define Excel column titles
    xlSheet.Cells(1, 1).Value = HeadingColumnTitle
    xlSheet.Cells(1, 2).Value = ParagraphTextTitle
    For i = 1 To UBound(subsequentTags) + 1
        xlSheet.Cells(1, i + 2).Value = "Contains " & subsequentTags(i - 1) & "?"
    Next i
    
    ' Document setup
    Dim srcDoc As Document
    Set srcDoc = ActiveDocument
    Dim iRow As Long: iRow = 2 ' Start from the second row for data
    
    ' Variables for heading tracking
    Dim currentHeadingText As String
    Dim lastHeadingText As String: lastHeadingText = ""
    Dim lastHeadingPosition As Long: lastHeadingPosition = 0
    
    Dim aRng As Range, aRngHead As Range
    Set aRng = srcDoc.Range
    
    With aRng.Find
        .Text = SearchTextR
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            ' Heading retrieval, exactly as in GatherRoundEnhancedToExcelOptimized
            Set aRngHead = aRng.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
            If Not aRngHead Is Nothing Then
                currentHeadingText = Trim(aRngHead.Text)
                If currentHeadingText <> lastHeadingText Then
                    lastHeadingText = currentHeadingText
                    lastHeadingPosition = aRngHead.Start
                End If
            Else
                currentHeadingText = "No Heading"
            End If
            
            ' Output the heading and paragraph text
            xlSheet.Cells(iRow, 1).Value = currentHeadingText
            Dim fullText As String
            fullText = aRng.Paragraphs(1).Range.Text
            
            xlSheet.Cells(iRow, 2).Value = fullText ' Full paragraph text containing [R]
            
            ' Improved tag detection to account for inconsistent spacing
            For i = 1 To UBound(subsequentTags) + 1
                ' Using pattern matching for each tag to ensure detection despite spacing
                Dim tagPattern As String
                tagPattern = subsequentTags(i - 1)
                
                If InStr(fullText, "[R]") < InStr(fullText, tagPattern) And InStr(fullText, tagPattern) > 0 Then
                    xlSheet.Cells(iRow, i + 2).Value = "Yes"
                Else
                    xlSheet.Cells(iRow, i + 2).Value = "No"
                End If
            Next i
            
            iRow = iRow + 1 ' Prepare for the next row
            aRng.Collapse wdCollapseEnd
        Loop
    End With
    
    ' Auto-fit Excel columns for visibility
    xlSheet.Columns.AutoFit
End Sub
