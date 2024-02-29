Sub FindRWithoutHash()
    Const SearchTextR As String = "[R]"
    Const HeadingColumnTitle As String = "Heading"
    Const TextColumnTitle As String = "Missing Hash"
    
    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim aRng As Range, checkRng As Range
    Dim iRow As Long
    
    Set srcDoc = ActiveDocument
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    xlApp.Visible = True
    
    xlSheet.Cells(1, 1).Value = HeadingColumnTitle
    xlSheet.Cells(1, 2).Value = TextColumnTitle
    iRow = 2
    
    Set aRng = srcDoc.Range
    
    With aRng.Find
        .ClearFormatting
        .Text = SearchTextR
        .Forward = True
        Do While .Execute
            ' Check immediately after [R] for #
            Set checkRng = aRng.Duplicate
            checkRng.MoveStart wdCharacter, Len(SearchTextR)
            checkRng.MoveEnd wdCharacter, 1 ' Adjust this range as needed to check for spacing issues
            
            If Not checkRng.Text Like "#*" Then
                ' If # is not found immediately after [R], output to Excel
                xlSheet.Cells(iRow, 1).Value = "[R] without following #"
                xlSheet.Cells(iRow, 2).Value = aRng.Paragraphs(1).Range.Text
                
                iRow = iRow + 1
            End If
            
            aRng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
    
    xlSheet.Columns("A:B").AutoFit
    xlApp.Visible = True
End Sub
