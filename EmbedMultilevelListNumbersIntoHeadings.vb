Sub EmbedMultilevelListNumbersIntoHeadings()
    Dim para As Paragraph
    Dim listNum As String
    Dim doc As Document
    Set doc = ActiveDocument
    
    Application.ScreenUpdating = False ' Turn off screen updating to improve performance
    
    For Each para In doc.Paragraphs
        If para.Style Like "Heading*" Then
            If Not para.Range.ListFormat.List Is Nothing Then
                ' Extract the list number (including the dot at the end)
                listNum = para.Range.ListFormat.ListString & " "
                ' Check if the paragraph already starts with the list number to avoid duplication
                If Not para.Range.Text Like listNum & "*" Then
                    ' Insert the list number at the beginning of the paragraph
                    para.Range.InsertBefore listNum
                End If
            End If
        End If
    Next para
    
    Application.ScreenUpdating = True ' Turn on screen updating again
End Sub