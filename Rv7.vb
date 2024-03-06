Sub ListRTagsWithSubsequentTagsAndHeadingsEnhanced()
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
    Dim para As Paragraph
    Dim iRow As Long: iRow = 2 ' Start from the second row for data
    Dim lastHeadingText As String: lastHeadingText = ""
    Dim lastHeadingPosition As Long: lastHeadingPosition = 0
    Dim currentHeadingText As String
    
    ' Iterate over each paragraph in the document
    For Each para In srcDoc.Paragraphs
        If InStr(para.Range.Text, "[R]") > 0 Then ' Check if paragraph contains [R]
            ' Only update heading if we've moved past the last heading position
            If para.Range.Start >= lastHeadingPosition Then
                Dim headingRange As Range
                Set headingRange = para.Range.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
                If Not headingRange Is Nothing Then
                    headingRange.End = headingRange.Paragraphs(1).Range.End - 1 ' Adjust range to exclude paragraph mark
                    currentHeadingText = Trim(headingRange.Text)
                    If currentHeadingText <> lastHeadingText Then
                        lastHeadingText = currentHeadingText
                        lastHeadingPosition = headingRange.Start ' Update last heading position
                    End If
                Else
                    currentHeadingText = "No Heading"
                End If
            End If
            
            ' Output to Excel
            xlSheet.Cells(iRow, 1).Value = lastHeadingText
            xlSheet.Cells(iRow, 2).Value = Trim(para.Range.Text) ' Place paragraph text
            
            ' Check for each specified subsequent tag within the same paragraph
            For i = 0 To UBound(subsequentTags)
                If InStr(para.Range.Text, subsequentTags(i)) > 0 Then
                    xlSheet.Cells(iRow, i + 3).Value = "Yes"
                Else
                    xlSheet.Cells(iRow, i + 3).Value = "No"
                End If
            Next i
            
            iRow = iRow + 1 ' Move to the next row for the next entry
        End If
    Next para
    
    ' Auto-fit columns for better visibility
    xlSheet.Columns.AutoFit
End Sub
