Sub GatherRTagsWithHashesFinalVersion()
    On Error GoTo ErrorHandler

    ' Constants for search text and column titles
    Const SearchTextR As String = "[R]"
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Text"
    
    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim lastHeadingText As String, currentHeadingText As String
    Dim aRng As Range, aRngHead As Range
    Dim iRow As Long
    
    Set srcDoc = ActiveDocument
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    xlApp.Visible = True ' Excel visibility for debugging
    
    xlSheet.Cells(1, 1).Value = HeadingColumnTitle
    xlSheet.Cells(1, 2).Value = TextColumnTitle
    iRow = 2 ' Starting row for data
    
    Set aRng = srcDoc.Range
    
    With aRng.Find
        .ClearFormatting
        .Text = SearchTextR
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            ' Move to start of the paragraph to ensure complete capture
            aRng.Start = aRng.Paragraphs(1).Range.Start
            
            ' Capture the heading text as in the original script
            Set aRngHead = aRng.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
            If Not aRngHead Is Nothing Then
                ' This captures the entire heading text, including numbering if present
                currentHeadingText = aRngHead.Text
                If Right(currentHeadingText, 1) = vbCr Then
                    currentHeadingText = Left(currentHeadingText, Len(currentHeadingText) - 1) ' Remove trailing carriage return
                End If
            Else
                currentHeadingText = "No Heading" ' Fallback if no heading is found
            End If
            
            ' Avoid duplicates by checking if the heading has changed
            If currentHeadingText <> lastHeadingText Then
                lastHeadingText = currentHeadingText
            End If
            
            ' Insert data into Excel
            xlSheet.Cells(iRow, 1).Value = Trim(currentHeadingText)
            xlSheet.Cells(iRow, 2).Value = Trim(aRng.Text)
            
            iRow = iRow + 1 ' Prepare for the next row
            aRng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
    
    ' AutoFit Excel columns for a better appearance
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True ' Make sure Excel is visible after processing
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred: " & Err.Description
    xlApp.Visible = True ' Ensure Excel is visible if an error occurs
End Sub
