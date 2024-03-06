Sub ListRTagsAndSubsequentTagsWithAccurateHeadingAndSpacing()
    ' Defined constants and subsequent tags array
    Dim subsequentTags As Variant
    subsequentTags = Array("#SPM", "#SPAM", "#TechLead", "#RSM")
    
    ' Excel application setup
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    ' Excel column titles setup
    xlSheet.Cells(1, 1).Value = "Heading"
    xlSheet.Cells(1, 2).Value = "Paragraph Containing [R]"
    Dim i As Long
    For i = 0 To UBound(subsequentTags)
        xlSheet.Cells(1, i + 3).Value = "Contains " & subsequentTags(i) & "?"
    Next i
    
    ' Document and range setup
    Dim srcDoc As Document
    Set srcDoc = ActiveDocument
    Dim aRng As Range, aRngHead As Range
    Dim currentHeadingText As String, lastHeadingText As String
    Dim lastHeadingPosition As Long
    Dim iRow As Long: iRow = 2
    
    Set aRng = srcDoc.Range
    lastHeadingPosition = 0 ' Initialize
    
    With aRng.Find
        .Text = SearchTextR
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            aRng.Start = aRng.Paragraphs(1).Range.Start
            
            ' Heading retrieval as per previous logic
            Set aRngHead = aRng.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
            If Not aRngHead Is Nothing Then
                currentHeadingText = Trim(aRngHead.ListFormat.ListString & " " & aRngHead.Text)
                If currentHeadingText <> lastHeadingText Then
                    lastHeadingText = currentHeadingText
                    lastHeadingPosition = aRngHead.Start
                End If
            Else
                currentHeadingText = "No Heading"
            End If
            
            ' Output the heading
            xlSheet.Cells(iRow, 1).Value = currentHeadingText
            
            ' Extract text following [R] for subsequent tag checks
            Dim textFollowingR As String
            textFollowingR = Mid(aRng.Text, InStr(aRng.Text, "[R]") + Len("[R]"))
            
            ' Output the full paragraph text
            xlSheet.Cells(iRow, 2).Value = Trim(aRng.Text)
            
            ' Check for subsequent tags within the extracted text
            For i = 0 To UBound(subsequentTags)
                xlSheet.Cells(iRow, i + 3).Value = IIf(InStr(textFollowingR, subsequentTags(i)) > 0, "Yes", "No")
            Next i
            
            iRow = iRow + 1 ' Prepare for next entry
            aRng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
    
    ' Auto-fit for better visibility in Excel
    xlSheet.Columns.AutoFit
End Sub
