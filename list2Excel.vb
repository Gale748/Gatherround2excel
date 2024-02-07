Sub ExtractListItemsToExcelAfterMarker()
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim para As Paragraph
    Dim rowNumber As Long
    Dim currentHeading As String
    Dim capture As Boolean ' Flag to start capturing list items

    Set wdApp = Word.Application
    Set wdDoc = wdApp.ActiveDocument

    ' Create a new instance of Excel
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    If xlApp Is Nothing Then
        MsgBox "Excel is not available."
        Exit Sub
    End If
    On Error GoTo 0

    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)

    ' Headers for Excel columns
    xlSheet.Cells(1, 1).Value = "Heading Name"
    xlSheet.Cells(1, 2).Value = "List Item"

    rowNumber = 2 ' Start from the second row to avoid headers
    currentHeading = "" ' Initialize current heading
    capture = False ' Initialize capture flag

    For Each para In wdDoc.Paragraphs
        If InStr(para.Range.Text, "SAMPLE TEXT") > 0 Then
            capture = True ' Start capturing after finding [R]
        ElseIf capture Then
            ' Check if the paragraph is a heading or subheading
            If para.Style Like "Heading*" Then
                currentHeading = para.Range.Text
                currentHeading = Left(currentHeading, Len(currentHeading) - 1) ' Remove end-of-paragraph mark
            ElseIf para.Range.ListFormat.listType <> WdListType.wdListNoNumbering Then
                ' Check if we have captured a heading yet, if not, try to capture the nearest previous heading
                If currentHeading = "" Then
                    Dim tempPara As Paragraph
                    For Each tempPara In wdDoc.Paragraphs
                        If tempPara.Range.Start < para.Range.Start Then
                            If tempPara.Style Like "Heading*" Then
                                currentHeading = tempPara.Range.Text
                                currentHeading = Left(currentHeading, Len(currentHeading) - 1) ' Remove end-of-paragraph mark
                                Exit For ' Found the nearest previous heading
                            End If
                        Else
                            Exit For ' Reached or passed the current paragraph
                        End If
                    Next tempPara
                End If

                ' Write heading and list item to Excel
                xlSheet.Cells(rowNumber, 1).Value = currentHeading
                xlSheet.Cells(rowNumber, 2).Value = Trim(para.Range.Text)

                rowNumber = rowNumber + 1
            End If
        End If
    Next para

    ' Auto-fit columns in Excel
    xlSheet.Columns("A:B").AutoFit

    MsgBox "Extraction completed successfully!"

    ' Clean up
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub