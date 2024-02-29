Sub GatherRTagsWithHashesOptimizedAndErrorHandled()
    On Error GoTo ErrorHandler

    ' Define constants for search text and column titles
    Const SearchTextR As String = "[R]"
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Text"

    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim currentHeadingText As String
    Dim aRng As Range, aRngHead As Range
    Dim iRow As Long, documentEnd As Long

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
    documentEnd = srcDoc.Content.End

    With aRng.Find
        .ClearFormatting
        .Text = SearchTextR
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            ' Extend the range to capture following '#' characters, accounting for possible whitespace
            Dim extendedRange As Range
            Set extendedRange = aRng.Duplicate
            extendedRange.MoveEndWhile Cset:=" ", Count:=wdForward
            extendedRange.MoveEndUntil Cset:="#", Count:=wdForward
            extendedRange.Expand Unit:=wdParagraph

            If InStr(extendedRange.Text, "#") > 0 Then
                aRng.SetRange Start:=aRng.Start, End:=extendedRange.End
            End If

            ' Retrieve heading
            aRng.Start = aRng.Paragraphs(1).Range.Start
            Set aRngHead = aRng.GoTo(wdGoToHeading, wdGoToPrevious)
            If Not aRngHead Is Nothing Then
                ' Combine the heading number (if any) and the heading text
                Dim headingNumber As String
                If aRngHead.ListFormat.ListType <> wdListNoNumbering Then
                    headingNumber = aRngHead.ListFormat.ListString & " "
                Else
                    headingNumber = ""
                End If
                currentHeadingText = headingNumber & Trim(aRngHead.Text)
            Else
                currentHeadingText = "No Heading"
            End If

            ' Insert data into Excel
            xlSheet.Cells(iRow, 1).Value = currentHeadingText
            xlSheet.Cells(iRow, 2).Value = Trim(aRng.Text)

            iRow = iRow + 1 ' Prepare for the next row
            aRng.Collapse Direction:=wdCollapseEnd

            ' Break loop if at the end of the document
            If aRng.Start >= documentEnd Then Exit Do
        Loop
    End With

    ' Optimize Excel appearance
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True ' Make sure Excel is visible after processing
    Exit Sub

ErrorHandler:
    MsgBox "An error has occurred: " & Err.Description
    xlApp.Visible = True ' Ensure Excel is visible if an error occurs
End Sub
