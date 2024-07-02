Sub Button2_Click()
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim WordTable As Object
    Dim ws As Worksheet
    Dim cell As Range
    ' Dim highlightedCells As Range
    Dim highlightedCells(1 To 100) As Range ' Adjust the size based on expected number of highlighted cells
    Dim numHighlighted As Integer ' Counter for highlighted cells
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
    Dim str As String
     
    Dim columnIndex(1 To 3) As Integer
    Set dict = CreateObject("Scripting.Dictionary")
    columnIndex(1) = 3
    columnIndex(2) = 4
    columnIndex(3) = 9
    
    Set ws = ThisWorkbook.Sheets("Meeting Monitoring Sheet")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastCol = ws.Cells(9, ws.Columns.Count).End(xlToLeft).Column - 1
    ' Convert the month name in cell E2 to a month number
    dateFilterMonth = Month(DateValue("01 " & ws.Range("E2").Value & " 2000"))
    On Error Resume Next
    
    numHighlighted = 0
    For Each cell In ws.Cells(10, 1).Resize(lastRow, 1)
        If Trim(ws.Cells(cell.Row, 3)) <> "" Then
            ' Check if the cell in column C (2 columns before the current cell) is in the same month as dateFilterMonth
            cellDate = ws.Cells(cell.Row, 3).Value
            If Month(cellDate) = dateFilterMonth Then
                ' Debug.Print "cell " & cell.Value
                ' Check if the cell is highlighted with the specified color
                For c = 1 To 3
                    Debug.Print ws.Cells(cell.Row, columnIndex(c)).Value
                    
                    numHighlighted = numHighlighted + 1
                    Set highlightedCells(numHighlighted) = ws.Cells(cell.Row, columnIndex(c))
                    
                Next c
                
                ' Get the value from column G (assuming column 7)
                Dim valueG As Variant
                valueG = ws.Cells(cell.Row, 7).Value
                
                ' Check if the value already exists in the dictionary
                If dict.Exists(valueG) Then
                    dict(valueG) = dict(valueG) + 1  ' Increment count
                Else
                    dict.Add valueG, 1  ' Add new entry with count 1
                End If
                
            End If
        End If
    Next cell
    On Error GoTo 0
    ' Debugging: Check if any cells were highlighted
    If numHighlighted = 0 Then
        Debug.Print "No highlighted cells found." & lastRow
        Exit Sub
    Else
        Debug.Print "Highlighted cells found: " & numHighlighted / 3
            
        
          ' Create a new Word Application
            Set WordApp = CreateObject("Word.Application")
            WordApp.Visible = True
            ' Create a new Word Document
            Set WordDoc = WordApp.Documents.Add
              
            Set para = WordDoc.Paragraphs.Add
            
            str = ""
            ' Loop through dictionary keys and write to Word doc
            For Each Key In dict.Keys
                str = Key & "  "
                str = str & dict(Key)
                WordDoc.Paragraphs.Add.Range.Text = str
                WordDoc.Paragraphs.Add
            Next Key
            
            WordDoc.Paragraphs.Add
            
            ' Add a table to the Word Document
            rowCount = numHighlighted / 3
            
            ' Assuming you want to display address and value in 3 columns
            Set MyRange = WordDoc.Content
            MyRange.Collapse Direction:=wdCollapseEnd
            Set WordTable = WordDoc.Tables.Add(MyRange, rowCount + 1, 3)
            
            j = 1
            For c = 1 To 3
                WordTable.cell(1, j).Range.Font.Bold = True
                WordTable.cell(1, j).Range.Text = ws.Cells(9, columnIndex(c)).Value
                
                WordTable.cell(1, j).Range.Shading.BackgroundPatternColor = RGB(18, 80, 27)
                WordTable.cell(1, j).Range.Shading.ForegroundPatternColor = wdColorWhite
                j = j + 1
            Next c
            
            ' Loop through the highlighted cells and add to the Word table
            i = 2
            j = 1
            
            For ro = 1 To numHighlighted
                
                WordTable.cell(i, j).Range.Text = highlightedCells(ro).Value ' Add cell value to the table
                j = j + 1
                
                If j > 3 Then
                    j = 1
                    i = i + 1
                End If
            Next ro
            ' Format the Word table (optional)
            ' WordTable.Style = "Table Grid"
            ' Cleanup
            WordTable.Borders.Enable = True
            
            Set WordTable = Nothing
            Set WordDoc = Nothing
            Set WordApp = Nothing
            
            Debug.Print "Doc file created!"
        
    End If
    
End Sub
