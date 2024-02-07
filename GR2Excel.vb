Sub GatherRoundEnhancedToExcelOptimized()
    ' Define constants for search text and column titles for maintainability
    Const SearchText As String = "[R]"
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Text"
    
    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim lastHeadingText As String, currentHeadingText As String
    Dim lastHeadingPosition As Long
    Dim sNum As String
    Dim aRng As Range, aRngHead As Range
    Dim iRow As Long
    
    Set srcDoc = ActiveDocument
    ' Create a new Excel instance and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    ' Excel visibility for debugging, set to False for speed optimization
    xlApp.Visible = True
    
    ' Add column titles to the first row in Excel
    xlSheet.Cells(1, 1).Value = HeadingColumnTitle
    xlSheet.Cells(1, 2).Value = TextColumnTitle
    iRow = 2 ' Start from the second row for data
    
    Set aRng = srcDoc.Range
    lastHeadingPosition = 0 ' Initialize last heading position
    
    With aRng.Find
        .ClearFormatting
        .Text = SearchText
        .Forward = True
        Do While .Execute
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
            If aRng.ListFormat.listType <> wdListNoNumbering Then
                sNum = aRng.ListFormat.listString
                xlSheet.Cells(iRow, 2).Value = sNum & " " & aRng.Text ' Using space instead of tab for Excel
            Else
                xlSheet.Cells(iRow, 2).Value = aRng.Text
            End If
            xlSheet.Cells(iRow, 1).Value = currentHeadingText
            
            ' Prepare for the next search and row
            aRng.Collapse Direction:=wdCollapseEnd
            iRow = iRow + 1
        Loop
    End With
    
    ' Optimize Excel appearance
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True ' Make sure Excel is visible after processing
End Sub