Sub GatherRTagsWithHashesOptimized()
    ' Define constants for search text and column titles
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
            ' Extend the range to capture following '#' characters, accounting for possible whitespace
            Dim extendedRange As Range
            Set extendedRange = aRng.Duplicate
            extendedRange.MoveEndWhile Cset:=" ", Count:=wdForward
            extendedRange.MoveEndUntil Cset:="#", Count:=wdForward
            extendedRange.Expand Unit:=wdParagraph

            ' Check if extended range actually found additional '#' text
            If InStr(extendedRange.Text, "#") > 0 Then
                aRng.SetRange Start:=aRng.Start, End:=extendedRange.End
            End If

            ' Retrieve heading
            aRng.Start = aRng.Paragraphs(1).Range.Start
            If aRng.Start >= lastHeadingPosition Then
                Set aRngHead = aRng.GoToPrevious(wdGoToHeading)
                If Not aRngHead Is Nothing Then
                    aRngHead.End = aRngHead.Paragraphs(1).Range.End - 1
                    currentHeadingText = aRngHead.ListFormat.ListString & vbTab & Trim(aRngHead.Text)
                    If currentHeadingText <> lastHeadingText Then
                        lastHeadingText = currentHeadingText
                        lastHeadingPosition = aRngHead.Start
                    End If
                Else
                    currentHeadingText = "No Heading"
                End If
            End If

            ' Insert data into Excel
            xlSheet.Cells(iRow, 1).Value = currentHeadingText
            xlSheet.Cells(iRow, 2).Value = Trim(aRng.Text)

            ' Prepare for the next search and row
            iRow = iRow + 1
            aRng.Collapse Direction:=wdCollapseEnd
        Loop
    End With

    ' Optimize Excel appearance
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True ' Make sure Excel is visible after processing
End Sub
