Sub ConvertTablesToExcel()
    Dim wdDoc As Document
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    Dim tbl As Table
    Dim row As row
    Dim col As cell
    Dim i As Long
    Dim j As Long
    
    ' Create a new Excel application
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    
    ' Create a new workbook in Excel
    Set xlWB = xlApp.Workbooks.Add
    Set xlSheet = xlWB.Worksheets(1)
    
    ' Set the Word document object
    Set wdDoc = ThisDocument
    
    i = 1 ' Initialize row counter
    
    ' Loop through each table in the Word document
    For Each tbl In wdDoc.Tables
        ' Loop through each row in the table
        For Each row In tbl.Rows
            j = 1 ' Initialize column counter
            
            ' Loop through each cell in the row
            For Each col In row.Cells
                On Error Resume Next ' Skip error and continue to the next cell if it's vertically merged
                
                ' Copy the cell content to the corresponding cell in Excel
                xlSheet.Cells(i, j).Value = col.Range.Text
                
                On Error GoTo 0 ' Reset error handling to default
                
                j = j + 1 ' Increment column counter
            Next col
            
            i = i + 1 ' Increment row counter
        Next row
    Next tbl
    
    ' Clean up Excel objects
    Set xlSheet = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    
    MsgBox "Tables converted to Excel successfully!", vbInformation
End Sub 
