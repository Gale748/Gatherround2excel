Sub GatherRTagsMissingHashTags()
    ' Define constants for search text and column titles for maintainability
    Const SearchTextR As String = "[R]"
    Const HashTag As String = "#"
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Potential Missing # Tag"
    
    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim currentHeadingText As String
    Dim aRng As Range, aRngHead As Range, nextCharRange As Range
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
        .ClearFormatting
        .Text = SearchTextR
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            ' Check the character immediately following the [R] tag
            Set nextCharRange = srcDoc.Range(Start:=aRng.End, End:=aRng.End + 1)
            nextCharRange.MoveEndWhile Cset:=" ", Count:=wdForward ' Skip any spaces
            If Not nextCharRange.Text Like HashTag & "*" Then
                ' No # immediately following [R], so capture this instance
                
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
                xlSheet.Cells(iRow, 2).Value = aRng.Text
                
                iRow = iRow + 1 ' Prepare for the next row
            End If
            
            ' Move search range to start after the current [R] tag
            aRng.Start = nextCharRange.End + 1
            aRng.End = srcDoc.Content.End
        Loop
    End With
    
    ' Optimize Excel appearance
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True ' Make sure Excel is visible after processing
End Sub
