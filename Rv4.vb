Sub GatherRTagsMissingHashTagsOptimized()
    ' Constants for search text and column titles
    Const SearchTextR As String = "[R]"
    Const HashTag As String = "#"
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Instance Detail"
    
    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim aRng As Range, aRngHead As Range, checkRange As Range
    Dim iRow As Long
    
    Set srcDoc = ActiveDocument
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    xlApp.Visible = True ' For debugging
    
    ' Add column titles
    xlSheet.Cells(1, 1).Value = HeadingColumnTitle
    xlSheet.Cells(1, 2).Value = TextColumnTitle
    iRow = 2 ' Start data from the second row
    
    Set aRng = srcDoc.Range
    
    With aRng.Find
        .Text = SearchTextR
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute
            ' Check for '#' following '[R]'
            Set checkRange = srcDoc.Range(Start:=aRng.End, End:=aRng.End + 1)
            checkRange.MoveEnd wdCharacter, 1
            checkRange.Expand wdWord
            
            ' If no '#' is found immediately after '[R]', capture the instance
            If Not checkRange.Text Like HashTag & "*" Then
                ' Find and set the heading related to the [R] tag
                Set aRngHead = aRng.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
                If Not aRngHead Is Nothing Then
                    xlSheet.Cells(iRow, 1).Value = Trim(aRngHead.Text)
                Else
                    xlSheet.Cells(iRow, 1).Value = "No Heading"
                End If
                
                xlSheet.Cells(iRow, 2).Value = aRng.Text
                
                iRow = iRow + 1 ' Move to the next row for the next entry
            End If
            
            ' Move the search range forward to continue the search
            aRng.SetRange Start:=checkRange.End, End:=srcDoc.Content.End
        Loop
    End With
    
    ' Auto-fit columns for better visibility
    xlSheet.Columns("A:B").AutoFit
    
    ' Ensure Excel is visible after processing
    xlApp.Visible = True
End Sub
