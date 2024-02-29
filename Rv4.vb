Sub GatherRTagsWithEnhancedHeadingAndTextExtraction()
    ' Define constants for search text and column titles for maintainability
    Const SearchTextR As String = "[R]"
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Text"
    
    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim currentHeadingText As String
    Dim aRng As Range, aRngHead As Range
    Dim iRow As Long
    
    Set srcDoc = ActiveDocument
    ' Create a new Excel instance and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    xlApp.Visible = True ' Excel visibility for debugging
    
    ' Add column titles to the first row in Excel
    xlSheet.Cells(1, 1).Value = HeadingColumnTitle
    xlSheet.Cells(1, 2).Value = TextColumnTitle
    iRow = 2 ' Start from the second row for data
    
    Set aRng = srcDoc.Range
    
    Do While aRng.Find.Execute(FindText:=SearchTextR, Forward:=True)
        ' Extend range to capture following '#' characters, accounting for possible whitespace
        Dim extendedRange As Range
        Set extendedRange = aRng.Duplicate
        extendedRange.MoveEndWhile Cset:=" ", Count:=wdForward
        extendedRange.MoveEndUntil Cset:="#", Count:=wdForward
        extendedRange.Expand Unit:=wdParagraph
        
        ' Retrieve heading using the original method for accuracy
        Set aRngHead = aRng.GoTo(wdGoToHeading, wdGoToPrevious)
        If Not aRngHead Is Nothing Then
            currentHeadingText = Trim(aRngHead.Text)
            If Right(currentHeadingText, 1) = vbCr Then
                currentHeadingText = Left(currentHeadingText, Len(currentHeadingText) - 1) ' Remove trailing carriage return
            End If
        Else
            currentHeadingText = "No Heading"
        End If
        
        ' Insert data into Excel
        xlSheet.Cells(iRow, 1).Value = currentHeadingText
        xlSheet.Cells(iRow, 2).Value = Trim(aRng.Text)
        
        iRow = iRow + 1 ' Prepare for the next row
        aRng.Collapse Direction:=wdCollapseEnd
        
        ' Ensure we move beyond the current `[R]` tag to avoid duplicate captures
        aRng.Start = extendedRange.End + 1
        If aRng.Start >= srcDoc.Content.End Then Exit Do
    Loop
    
    ' Optimize Excel appearance
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True ' Make sure Excel is visible after processing
End Sub
