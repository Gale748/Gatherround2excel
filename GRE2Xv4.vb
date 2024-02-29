Sub GatherRoundEnhancedToExcelMissingHash()
    ' Define constants for search text and column titles for maintainability
    Const SearchTextR As String = "[R]"
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Text"
    
    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim lastHeadingText As String, currentHeadingText As String
    Dim lastHeadingPosition As Long
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
    lastHeadingPosition = 0 ' Initialize last heading position
    
    With aRng.Find
        .ClearFormatting
        .Text = SearchTextR
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            ' Check if the character following [R] is not #
            Dim followText As String
            If aRng.End < srcDoc.Content.End Then
                followText = srcDoc.Range(Start:=aRng.End, End:=aRng.End + 1).Text
                ' Adjust for potential space before #
                If Not followText Like "[!#]*#" Then ' Adjusted to check for absence of # immediately following [R]
                    aRng.Start = aRng.Paragraphs(1).Range.Start
                    If aRng.Start >= lastHeadingPosition Then
                        ' Retrieve heading only if we've moved past the last heading position
                        Set aRngHead = aRng.GoToPrevious(wdGoToHeading)
                        If Not aRngHead Is Nothing Then
                            aRngHead.End = aRngHead.Paragraphs(1).Range.End - 1
                            currentHeadingText = aRngHead.ListFormat.listString & vbTab & Trim(aRngHead.Text)
                            If currentHeadingText <> lastHeadingText Then
                                lastHeadingText = currentHeadingText
                                lastHeadingPosition = aRngHead.Start ' Update last heading position
                            End If
                        Else
                            currentHeadingText = "No Heading"
                        End If
                    End If
                    
                    ' Process text and insert into Excel
                    xlSheet.Cells(iRow, 1).Value = currentHeadingText
                    xlSheet.Cells(iRow, 2).Value = aRng.Text
                    
                    ' Prepare for the next search and row
                    iRow = iRow + 1
                End If
            End If
            aRng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
    
    ' Optimize Excel appearance
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True ' Make sure Excel is visible after processing
End Sub
