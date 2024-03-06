Sub ListRTagsWithFollowingSpecificTags()
    ' Specified tags to check after [R]
    Dim subsequentTags As Variant
    subsequentTags = Array("#SPM", "#SPAM", "#TechLead", "#RSM")
    
    ' Setup Excel
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlApp.Visible = True ' Make Excel visible for debugging
    
    ' Define Excel column titles
    xlSheet.Cells(1, 1).Value = "Heading"
    xlSheet.Cells(1, 2).Value = "Full Paragraph Text"
    For i = 0 To UBound(subsequentTags)
        xlSheet.Cells(1, i + 3).Value = "Contains " & subsequentTags(i) & "?"
    Next i
    
    ' Document setup
    Dim srcDoc As Document
    Set srcDoc = ActiveDocument
    Dim iRow As Long: iRow = 2 ' Start from second row for data
    
    ' Search through document
    Dim para As Paragraph
    For Each para In srcDoc.Paragraphs
        If InStr(para.Range.Text, "[R]") > 0 Then
            ' Retrieve and output the heading for each [R] tag
            Dim headingText As String
            headingText = GetHeadingText(para.Range)
            xlSheet.Cells(iRow, 1).Value = headingText
            
            ' Output the full paragraph text
            xlSheet.Cells(iRow, 2).Value = Trim(para.Range.Text)
            
            ' Check and mark the presence of each specified tag
            Dim tagText As String, tagFound As Boolean
            For i = 0 To UBound(subsequentTags)
                tagText = subsequentTags(i)
                tagFound = CheckTagFollowingR(para.Range.Text, tagText)
                xlSheet.Cells(iRow, i + 3).Value = IIf(tagFound, "Yes", "No")
            Next i
            
            iRow = iRow + 1 ' Prepare for the next row
        End If
    Next para
    
    ' Auto-fit Excel columns for visibility
    xlSheet.Columns.AutoFit
End Sub

Function GetHeadingText(rng As Range) As String
    ' Retrieves heading text for a given range
    Dim headingRange As Range
    Set headingRange = rng.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
    If Not headingRange Is Nothing Then
        GetHeadingText = Trim(headingRange.Text)
        If Right(GetHeadingText, 1) = Chr(13) Then ' Remove trailing carriage return
            GetHeadingText = Left(GetHeadingText, Len(GetHeadingText) - 1)
        End If
    Else
        GetHeadingText = "No Heading"
    End If
End Function

Function CheckTagFollowingR(paragraphText As String, tag As String) As Boolean
    ' Check for a specific tag following [R] within a paragraph, accounting for variable spacing
    Dim searchText As String
    searchText = "[R]" & "*" & tag ' Wildcard pattern to account for variable spacing
    CheckTagFollowingR = (InStr(paragraphText, tag) > 0) And (InStr(paragraphText, "[R]") < InStr(paragraphText, tag))
End Function
