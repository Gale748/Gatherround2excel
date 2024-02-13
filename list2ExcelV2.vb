Sub ExtractListItemsToExcelAfterMarker()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim para As Paragraph
    Dim rowNumber As Long
    Dim currentHeading As String
    Dim capture As Boolean
    
    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.ActiveDocument
    
    ' Attempt to create a new instance of Excel
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    If xlApp Is Nothing Then
        MsgBox "Excel is not available."
        Exit Sub
    End If
    On Error GoTo 0
    
    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    xlSheet.Cells(1, 1).Value = "Heading Name"
    xlSheet.Cells(1, 2).Value = "List Item"
    
    rowNumber = 2
    currentHeading = ""
    capture = False
    
    For Each para In wdDoc.Paragraphs
        If para.Style Like "Heading*" Then
            ' Capture the heading text including the section number if it's part of the text
            currentHeading = Trim(para.Range.Text)
        End If
        
        If InStr(para.Range.Text, ": [R]") > 0 Then
            capture = True
            ' Reset the heading to ensure it's ready to capture under the new section
            currentHeading = Trim(para.Range.Text)
        ElseIf capture And Not para.Range.ListFormat.List Is Nothing Then
            ' Check for list items specifically; adjust as needed for your document's structure
            xlSheet.Cells(rowNumber, 1).Value = currentHeading
            xlSheet.Cells(rowNumber, 2).Value = Trim(para.Range.Text)
            rowNumber = rowNumber + 1
        End If
    Next para
    
    xlSheet.Columns("A:B").AutoFit
    MsgBox "Extraction completed successfully!"
    
    ' Clean up
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
