Sub ListRTagsWithSubsequentTagsAndCaptureHeadings()
    ' Define the tags to look for after [R]
    Dim subsequentTags As Variant
    subsequentTags = Array("#SPM", "#SPAM", "#TechLead", "#RSM")
    
    ' Setup Excel
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    ' Setup column titles in Excel
    xlSheet.Cells(1, 1).Value = "Heading"
    xlSheet.Cells(1, 2).Value = "Paragraph Containing [R]"
    Dim i As Integer
    For i = 0 To UBound(subsequentTags)
        xlSheet.Cells(1, i + 3).Value = "Contains " & subsequentTags(i) & "?"
    Next i
    
    ' Initialize Word document variables
    Dim srcDoc As Document
    Set srcDoc = ActiveDocument
    Dim aRng As Range, aRngHead As Range
    Dim lastHeadingText As String: lastHeadingText = ""
    Dim currentHeadingText As String: currentHeadingText = ""
    Dim iRow As Long: iRow = 2 ' Start entering data from the second row
    
    Set aRng = srcDoc.Range
    
    With aRng.Find
        .Text = "[R]"
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            ' Reset range to start of paragraph to capture the heading
            aRng.Start = aRng.Paragraphs(1).Range.Start
            Set aRngHead = aRng.Paragraphs(1).Range.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
            
            If Not aRngHead Is Nothing Then
                currentHeadingText = Trim(aRngHead.Text)
            Else
                currentHeadingText = "No Heading"
            End If
            
            ' Write the current heading and paragraph containing [R] to Excel
            xlSheet.Cells(iRow, 1).Value = currentHeadingText
            xlSheet.Cells(iRow, 2).Value = Trim(aRng.Text)
            
            ' Extract text following [R] and check for subsequent tags
            Dim postRText As String
            postRText = Mid(aRng.Text, InStr(aRng.Text, "[R]") + Len("[R]"))
            
            ' Check for each tag
            For j = 0 To UBound(subsequentTags)
                If InStr(1, postRText, subsequentTags(j), vbTextCompare) > 0 Then
                    xlSheet.Cells(iRow, j + 3).Value = "Yes"
                Else
                    xlSheet.Cells(iRow, j + 3).Value = "No"
                End If
            Next j
            
            ' Move to the next row for the next [R] entry
            iRow = iRow + 1
            aRng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
    
    ' Auto-fit Excel columns for better readability
    xlSheet.Columns.AutoFit
End Sub
