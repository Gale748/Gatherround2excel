Sub GatherRTagsWithCorrectHeadingAndTextExtraction()
    ' Constants for search text and column titles
    Const SearchTextR As String = "[R]"
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Text"
    
    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim currentHeadingText As String, lastHeadingText As String
    Dim aRng As Range, aRngHead As Range
    Dim iRow As Long
    
    Set srcDoc = ActiveDocument
    ' Create a new Excel instance and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    xlApp.Visible = True ' Make Excel visible for debugging
    
    ' Add column titles to the first row in Excel
    xlSheet.Cells(1, 1).Value = HeadingColumnTitle
    xlSheet.Cells(1, 2).Value = TextColumnTitle
    iRow = 2 ' Initialize row counter for data entry
    
    Set aRng = srcDoc.Range
    
    ' Clear last heading text for comparison
    lastHeadingText = ""
    
    Do While aRng.Find.Execute(FindText:=SearchTextR, Forward:=True)
        ' Capture the paragraph where [R] is found
        Dim paraRange As Range
        Set paraRange = aRng.Paragraphs(1).Range
        
        ' Retrieve heading
        Set aRngHead = paraRange.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
        If Not aRngHead Is Nothing Then
            currentHeadingText = aRngHead.Text
            If Right(currentHeadingText, 1) = Chr(13) Then ' Remove trailing carriage return
                currentHeadingText = Left(currentHeadingText, Len(currentHeadingText) - 1)
            End If
        Else
            currentHeadingText = "No Heading"
        End If
        
        ' Check if the heading text has changed to avoid duplication
        If currentHeadingText <> lastHeadingText Or (currentHeadingText = lastHeadingText And aRng.Text <> srcDoc.Range(lastHeadingText).Text) Then
            ' Insert data into Excel
            xlSheet.Cells(iRow, 1).Value = currentHeadingText
            xlSheet.Cells(iRow, 2).Value = Trim(paraRange.Text)
            
            iRow = iRow + 1 ' Increment row counter for next entry
            lastHeadingText = currentHeadingText ' Update last heading text for comparison
        End If
        
        ' Move range to the end of the current paragraph to continue search
        aRng.Start = paraRange.End
        aRng.End = srcDoc.Content.End
        
        ' Exit condition if the end of the document is reached
        If aRng.Start >= srcDoc.Content.End Then Exit Do
    Loop
    
    ' Optimize Excel appearance
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True ' Keep Excel visible after processing
End Sub
