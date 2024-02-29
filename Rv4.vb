Sub ListRTagsWithSubsequentTags()
    ' Specify the tags to check for after [R]
    Dim subsequentTags As Variant
    subsequentTags = Array("#SPM", "#SPAM", "#TechLead", "#RSM")
    
    ' Excel setup
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlApp.Visible = True ' Make Excel visible
    
    ' Column titles
    xlSheet.Cells(1, 1).Value = "Paragraph Containing [R]"
    For i = 0 To UBound(subsequentTags)
        xlSheet.Cells(1, i + 2).Value = "Contains " & subsequentTags(i) & "?"
    Next i
    
    ' Document setup
    Dim srcDoc As Document
    Set srcDoc = ActiveDocument
    Dim para As Paragraph
    Dim iRow As Long: iRow = 2 ' Start from the second row for data
    
    ' Iterate over each paragraph in the document
    For Each para In srcDoc.Paragraphs
        If InStr(para.Range.Text, "[R]") > 0 Then ' Check if paragraph contains [R]
            xlSheet.Cells(iRow, 1).Value = para.Range.Text ' Place paragraph text in the first column
            
            ' Check for each specified subsequent tag within the same paragraph
            For i = 0 To UBound(subsequentTags)
                If InStr(para.Range.Text, subsequentTags(i)) > 0 Then
                    xlSheet.Cells(iRow, i + 2).Value = "Yes"
                Else
                    xlSheet.Cells(iRow, i + 2).Value = "No"
                End If
            Next i
            
            iRow = iRow + 1 ' Move to the next row for the next entry
        End If
    Next para
    
    ' Auto-fit columns for better visibility
    xlSheet.Columns.AutoFit
End Sub
