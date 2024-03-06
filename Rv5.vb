Sub ListRTagsWithSubsequentTagsAndCaptureHeadings()
    ' Tags to check for after [R]
    Dim subsequentTags As Variant
    subsequentTags = Array("#SPM", "#SPAM", "#TechLead", "#RSM")
    
    ' Excel setup
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlApp.Visible = True ' Make Excel visible
    
    ' Column titles in Excel
    xlSheet.Cells(1, 1).Value = "Heading"
    xlSheet.Cells(1, 2).Value = "Paragraph Containing [R]"
    For i = 0 To UBound(subsequentTags)
        xlSheet.Cells(1, i + 3).Value = "Contains " & subsequentTags(i) & "?"
    Next i
    
    ' Document setup
    Dim srcDoc As Document
    Set srcDoc = ActiveDocument
    Dim iRow As Long: iRow = 2 ' Start from the second row for data
    
    Dim aRng As Range, aRngHead As Range
    Dim lastHeadingText As String, currentHeadingText As String
    Dim lastHeadingPosition As Long
    
    Set aRng = srcDoc.Range
    lastHeadingPosition = 0 ' Initialize last heading position
    
    With aRng.Find
        .Text = "[R]"
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            aRng.Start = aRng.Paragraphs(1).Range.Start
            
            If aRng.Start >= lastHeadingPosition Then
                ' Retrieve heading only if we've moved past the last heading position
                Set aRngHead = aRng.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
                If Not aRngHead Is Nothing Then
                    aRngHead.End = aRngHead.Paragraphs(1).Range.End - 1
                    currentHeadingText = Trim(aRngHead.ListFormat.ListString & " " & aRngHead.Text)
                    
                    If currentHeadingText <> lastHeadingText Then
                        lastHeadingText = currentHeadingText
                        lastHeadingPosition = aRngHead.Start ' Update last heading position
                    End If
                Else
                    currentHeadingText = "No Heading"
                End If
            End If
            
            ' Output to Excel
            xlSheet.Cells(iRow, 1).Value = currentHeadingText
            xlSheet.Cells(iRow, 2).Value = Trim(aRng.Text) ' Place [R] tag text
            
            ' Check for each specified subsequent tag within the same paragraph
            For i = 0 To UBound(subsequentTags)
                If InStr(aRng.Text, subsequentTags(i)) > 0 Then
                    xlSheet.Cells(iRow, i + 3).Value = "Yes"
                Else
                    xlSheet.Cells(iRow, i + 3).Value = "No"
                End If
            Next i
            
            iRow = iRow + 1 ' Move to the next row for the next entry
            aRng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
    
    ' Auto-fit columns for better visibility
    xlSheet.Columns.AutoFit
End Sub