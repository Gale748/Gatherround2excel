Sub ListRTagsWithFollowingSpecificTagsAndCaptureHeadingsExactly()
    ' Tags to check for after [R]
    Dim subsequentTags As Variant
    subsequentTags = Array("#SPM", "#SPAM", "#TechLead", "#RSM")
    
    ' Excel setup
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlApp.Visible = True ' Make Excel visible
    
    ' Column titles in Excel
    xlSheet.Cells(1, 1).Value = "Heading"
    xlSheet.Cells(1, 2).Value = "Paragraph Containing [R]"
    For i = 0 To UBound(subsequentTags)
        xlSheet.Cells(1, i + 3).Value = "Contains " & subsequentTags(i) & "?"
    Next i
    
    ' Document setup
    Dim srcDoc As Document
    Set srcDoc = ActiveDocument
    Dim iRow As Long: iRow = 2 ' Start from the second row for data
    
    Dim aRng As Range
    Dim currentHeadingText As String
    Dim paragraphText As String
    
    Set aRng = srcDoc.Range
    
    With aRng.Find
        .Text = "[R]"
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            ' Extend range to include the whole paragraph
            Set aRng = aRng.Paragraphs(1).Range
            
            ' Retrieve heading text exactly as the provided script
            currentHeadingText = GetExactHeading(aRng)
            
            ' Prepare paragraph text for output
            paragraphText = Trim(aRng.Text)
            
            ' Output to Excel
            xlSheet.Cells(iRow, 1).Value = currentHeadingText
            xlSheet.Cells(iRow, 2).Value = paragraphText
            
            ' Check for each specified subsequent tag within the same paragraph
            Dim tagText As String
            For i = 0 To UBound(subsequentTags)
                tagText = subsequentTags(i)
                xlSheet.Cells(iRow, i + 3).Value = ContainsTag(aRng.Text, tagText)
            Next i
            
            iRow = iRow + 1 ' Move to the next row for the next entry
            
            aRng.Collapse wdCollapseEnd
        Loop
    End With
    
    ' Auto-fit Excel columns for better visibility
    xlSheet.Columns.AutoFit
End Sub

Function GetExactHeading(aRng As Range) As String
    ' Retrieves the heading exactly as done in the provided script
    Dim aRngHead As Range
    Set aRngHead = aRng.Duplicate
    aRngHead.Collapse wdCollapseStart
    Set aRngHead = aRngHead.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
    If Not aRngHead Is Nothing Then
        GetExactHeading = aRngHead.ListFormat.ListString & " " & Trim(aRngHead.Text)
    Else
        GetExactHeading = "No Heading"
    End If
End Function

Function ContainsTag(paragraphText As String, tag As String) As String
    ' Checks if the specified tag is present after [R] within the paragraph
    Dim pattern As String
    pattern = "\[R\][^\#]*" & tag ' Regex pattern to check for tag presence after [R]
    
    ' Use VBScript regular expression object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = pattern
    End With
    
    If regex.Test(paragraphText) Then
        ContainsTag = "Yes"
    Else
        ContainsTag = "No"
    End If
End Function
