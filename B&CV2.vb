Sub ExportBookmarksAndCrossReferencesToExcel()
    Dim xlApp As Object, xlBook As Object
    Dim bookmarksSheet As Object, crossRefsSheet As Object
    Dim bookmark As Bookmark, field As Field
    Dim excelRowBm As Long, excelRowCr As Long
    Dim headingText As String, crossRefText As String
    Dim wdRange As Range
    
    ' Initialize Excel Application
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set bookmarksSheet = xlBook.Sheets(1)
    bookmarksSheet.Name = "Bookmarks"
    Set crossRefsSheet = xlBook.Sheets.Add(After:=xlBook.Sheets(xlBook.Sheets.Count))
    crossRefsSheet.Name = "Cross-References"
    
    ' Setup Bookmarks Sheet
    With bookmarksSheet
        .Cells(1, 1).Value = "Name"
        .Cells(1, 2).Value = "Content"
        .Cells(1, 3).Value = "Heading Level Name"
    End With
    excelRowBm = 2
    
    ' Setup Cross-References Sheet
    With crossRefsSheet
        .Cells(1, 1).Value = "Reference Type"
        .Cells(1, 2).Value = "Reference Target"
        .Cells(1, 3).Value = "Surrounding Text"
    End With
    excelRowCr = 2
    
    ' Export Bookmarks
    For Each bookmark In ActiveDocument.Bookmarks
        ' Skip hidden bookmarks
        If Not Left(bookmark.Name, 1) = "_" Then
            Set wdRange = bookmark.Range
            headingText = "" ' Reset heading text for each bookmark
            
            ' Try to find the heading for the current bookmark position
            If wdRange.Paragraphs(1).OutlineLevel <> wdOutlineLevelBodyText Then
                headingText = wdRange.Paragraphs(1).Range.Text
            Else
                Do While wdRange.Start > 0
                    wdRange.MoveStart wdParagraph, -1
                    If wdRange.Paragraphs(1).OutlineLevel <> wdOutlineLevelBodyText Then
                        headingText = wdRange.Paragraphs(1).Range.Text
                        Exit Do
                    End If
                Loop
            End If
            
            ' Write bookmark details to Excel
            With bookmarksSheet
                .Cells(excelRowBm, 1).Value = bookmark.Name
                .Cells(excelRowBm, 2).Value = bookmark.Range.Text
                .Cells(excelRowBm, 3).Value = Trim(headingText)
            End With
            
            excelRowBm = excelRowBm + 1
        End If
    Next bookmark
    
    ' Export Cross-References
    ' Note: You'll need to adjust this section based on how you want to handle cross-references,
    ' as extracting specific details can vary. This part is left as a placeholder to remind you to
    ' implement the logic based on your document's structure and needs.
    
    ' AutoFit Columns for both sheets
    bookmarksSheet.Columns("A:C").AutoFit
    crossRefsSheet.Columns("A:C").AutoFit
    
    ' Notify the user upon completion
    MsgBox "Export Complete!", vbInformation
End Sub

