Sub ExportBookmarksToExcel()
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim wdRange As Range
    Dim bookmark As Bookmark
    Dim excelRow As Long
    Dim headingText As String
    
    ' Create a new Excel instance
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    
    With xlSheet
        .Cells(1, 1).Value = "Name"
        .Cells(1, 2).Value = "Content"
        .Cells(1, 3).Value = "Heading Level Name"
    End With
    
    excelRow = 2
    
    ' Iterate through all bookmarks in the Word document
    For Each bookmark In ActiveDocument.Bookmarks
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
        With xlSheet
            .Cells(excelRow, 1).Value = bookmark.Name
            .Cells(excelRow, 2).Value = bookmark.Range.Text
            .Cells(excelRow, 3).Value = Trim(headingText)
        End With
        
        excelRow = excelRow + 1
    Next bookmark
    
    ' AutoFit Columns in Excel
    xlSheet.Columns("A:C").AutoFit
    
    MsgBox "Export Complete!", vbInformation
End Sub
