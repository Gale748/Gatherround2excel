Sub GatherRTagsWithFollowingHashTags()
    ' Define constants for search text and column titles for maintainability
    Const SearchTextR As String = "[R]"
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Following Tags"
    
    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim currentHeadingText As String
    Dim aRng As Range, aRngHead As Range, textRng As Range
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
    
    With aRng.Find
        .Text = SearchTextR
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            ' Extend the range to the end of the paragraph to capture potential #'s
            Set textRng = srcDoc.Range(aRng.Start, aRng.Paragraphs(1).Range.End)
            
            ' Check if there are any # tags immediately following [R]
            If InStr(textRng.Text, "#") > 0 Then
                ' Retrieve heading for the [R] tag
                Set aRngHead = aRng.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
                If Not aRngHead Is Nothing Then
                    currentHeadingText = Trim(aRngHead.Text)
                    If Right(currentHeadingText, 1) = Chr(13) Then ' Remove trailing carriage return
                        currentHeadingText = Left(currentHeadingText, Len(currentHeadingText) - 1)
                    End If
                Else
                    currentHeadingText = "No Heading"
                End If
                
                ' Insert data into Excel
                xlSheet.Cells(iRow, 1).Value = currentHeadingText
                xlSheet.Cells(iRow, 2).Value = Trim(textRng.Text)
                
                iRow = iRow + 1 ' Prepare for the next row
            End If
            
            ' Move the search range forward to continue the search
            aRng.Start = textRng.End
            aRng.End = srcDoc.Content.End
        Loop
    End With
    
    ' Optimize Excel appearance
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True ' Make sure Excel is visible after processing
End Sub
