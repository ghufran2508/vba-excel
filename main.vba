Sub Button2_Click()
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim WordTable As Object
    Dim ws As Worksheet
    Dim cell As Range
    Dim highlightedCells As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim i As Long
    Dim j As Long
    Dim cellColor As Long
    Dim dateFilter As Date
    Dim dateFilterMonth As Long
    Dim cellDate As Date
    
    Set ws = ThisWorkbook.Sheets("Meeting Monitoring Sheet")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastCol = ws.Cells(9, ws.Columns.Count).End(xlToLeft).Column - 1
    ' Convert the month name in cell E2 to a month number
    dateFilterMonth = Month(DateValue("01 " & ws.Range("E2").Value & " 2000"))
    On Error Resume Next
    
    For Each cell In ws.Cells(10, 2).Resize(lastRow, lastCol)
        If Trim(ws.Cells(cell.Row, 3)) <> "" Then
            ' Check if the cell in column C (2 columns before the current cell) is in the same month as dateFilterMonth
            cellDate = ws.Cells(cell.Row, 3).Value
            If Month(cellDate) = dateFilterMonth Then
                ' Debug.Print "cell " & cell.Value
                ' Check if the cell is highlighted with the specified color
                If highlightedCells Is Nothing Then
                    Set highlightedCells = cell
                Else
                    Set highlightedCells = Union(highlightedCells, cell)
                End If
            End If
        End If
    Next cell
    On Error GoTo 0
    ' Debugging: Check if any cells were highlighted
    If highlightedCells Is Nothing Then
        Debug.Print "No highlighted cells found." & lastRow & " " & lastCol
        Exit Sub
    Else
        Debug.Print "Highlighted cells found: " & highlightedCells.Cells.Count / lastCol
        
          ' Create a new Word Application
            Set WordApp = CreateObject("Word.Application")
            WordApp.Visible = True
            ' Create a new Word Document
            Set WordDoc = WordApp.Documents.Add
            ' Add a table to the Word Document
            rowCount = highlightedCells.Cells.Count / lastCol
            Debug.Print rowCount & " " & lastCol
            ' Assuming you want to display address and value in two columns
            Set WordTable = WordDoc.Tables.Add(WordDoc.Range, rowCount + 1, lastCol)
            
            j = 1
            For Each cell In ws.Cells(9, 2).Resize(1, lastCol)
                WordTable.cell(1, j).Range.Font.Bold = True
                WordTable.cell(1, j).Range.Text = cell.Value
                WordTable.cell(1, j).Range.Shading.BackgroundPatternColor = RGB(18, 80, 27)
                WordTable.cell(1, j).Range.Shading.ForegroundPatternColor = wdColorWhite
                j = j + 1
            Next cell
            
            ' Loop through the highlighted cells and add to the Word table
            i = 2
            j = 1
            For Each cell In highlightedCells
                WordTable.cell(i, j).Range.Text = cell.Value ' Add cell value to the table
                j = j + 1
                
                If j = lastCol + 1 Then
                    j = 1
                    i = i + 1
                End If
            Next cell
            ' Format the Word table (optional)
            ' WordTable.Style = "Table Grid"
            ' Cleanup
            WordTable.Borders.Enable = True
            
            Set WordTable = Nothing
            Set WordDoc = Nothing
            Set WordApp = Nothing
            Set highlightedCells = Nothing
            
            Debug.Print "Doc file created!"
        
    End If
    
End Sub
