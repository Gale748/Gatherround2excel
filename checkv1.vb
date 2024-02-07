Sub CheckForValueAndMark()
    Dim ws As Worksheet
    Dim searchRange As Range
    Dim cell As Range
    Dim searchedValue As String
    Dim lastRow As Long
    
    ' Set the worksheet to work on. You can change "ActiveSheet" to something like Worksheets("Sheet1") if you know the name of your sheet.
    Set ws = ActiveSheet
    
    ' Set the searched value
    searchedValue = "R"
    
    ' Find the last row in the second column to avoid checking empty cells
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Set the range to search in, which is the entire second column up to the last row with data
    Set searchRange = ws.Range("B1:B" & lastRow)
    
    ' Add a new column for "Requirements" to the right of the second column
    Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Cells(1, 3).Value = "Requirements"
    
    ' Loop through each cell in the second column
    For Each cell In searchRange
        ' Check if the cell contains the searched value
        If InStr(1, cell.Value, searchedValue, vbTextCompare) > 0 Then
            ' If yes, write "Yes" in the next cell
            cell.Offset(0, 1).Value = "Yes"
        Else
            ' If no, write "No"
            cell.Offset(0, 1).Value = "No"
        End If
    Next cell
End Sub
