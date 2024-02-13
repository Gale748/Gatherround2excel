Sub ExtractListItemsToExcelAfterMarkerOptimized()
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim para As Paragraph
    Dim rowNumber As Long
    Dim currentHeading As String
    Dim capture As Boolean
    Dim aRng As Range, docRange As Range
    Dim lastFound As Long
    
    Set wdApp = Word.Application
    Set wdDoc = wdApp.ActiveDocument
    
    ' Create Excel instance
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' Set to False for performance
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    
    xlSheet.Cells(1, 1).Value = "Heading Name"
    xlSheet.Cells(1, 2).Value = "List Item"
    
    rowNumber = 2
    capture = False
    Set docRange = wdDoc.content
    lastFound = 0
    
    ' Loop through document paragraphs
    For Each para In wdDoc.Paragraphs
        If para.Range.Start >= lastFound Then
            If InStr(para.Range.Text, "[R]") > 0 Then
                capture = True
                ' Find and set the heading using the GoTo method
                Set aRng = para.Range.GoTo(wdGoToHeading, wdGoToPrevious)
                If Not aRng Is Nothing Then
                    currentHeading = aRng.Paragraphs(1).Range.Text
                    ' Adjust for multi-line headings
                    If aRng.Paragraphs.Count > 1 Then
                        Dim i As Long
                        For i = 2 To aRng.Paragraphs.Count
                            currentHeading = currentHeading & " " & aRng.Paragraphs(i).Range.Text
                        Next i
                    End If
                    currentHeading = Trim(currentHeading)
                Else
                    currentHeading = "No Heading"
                End If
                lastFound = para.Range.Start ' Update last found position
            ElseIf capture And para.Range.ListFormat.listType <> WdListType.wdListNoNumbering Then
                ' Write heading and list item to Excel
                xlSheet.Cells(rowNumber, 1).Value = currentHeading
                xlSheet.Cells(rowNumber, 2).Value = Trim(para.Range.Text)
                rowNumber = rowNumber + 1
            End If
        End If
    Next para
    
    ' Final adjustments and cleanup
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True
    MsgBox "Extraction completed successfully!"
    
    ' Cleanup
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
