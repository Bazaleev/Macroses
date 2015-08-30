Attribute VB_Name = "Utils"
'Searches and returns table in specified file and sheet
'Searching in current directory
Public Function getTable(sourceFileName As String, sheetName As String) As Range
    Dim tableName As String
    Dim sheet As Worksheet
    tableName = "Таблица1"
    
    Set sheet = getSheet(sourceFileName, sheetName)
    
    'if you get error in line below you specified invalid table name
    Set getTable = sheet.Range(tableName & "[#All]")
End Function

'Open specified file and returns payments sheet
Public Function getSheet(sourceFileName As String, sheetName As String) As Worksheet
    Dim Workbook As Workbook

    'if you get error in line below you specifified invalid file name
    'or file isn't opened now
    Set Workbook = Workbooks(sourceFileName)
    
    'if you get error in line below that means that invalid sheet name was specified
    Set getSheet = Workbook.Worksheets(sheetName)
End Function

'Returns sheet with specified name
Public Function getSheetInCurrentWorkbook(sheetName As String) As Worksheet
    Set getSheetInCurrentWorkbook = ThisWorkbook.Worksheets(sheetName)
End Function

'Search in specified sheet and returns specified column
Public Function getColumn(sheet As Worksheet, tableName As String, columnName As String) As Range
    Set getColumn = sheet.Range(tableName & "[" & columnName & "]")
End Function

'returns an index of speficied column
Public Function getColumnIndex(sheet As Worksheet, tableName As String, columnName As String) As Integer
    getColumnIndex = getColumn(sheet, tableName, columnName).Column
End Function

'returns an first empty cell row indext in a range
Public Function getFirstEmptyCellRowIndex(cellsToSearch As Range) As Integer
    Dim cell As Range
    Dim rowIndex As Integer
    Dim lastIndex As Integer
    
    For Each cell In cellsToSearch.Cells
        lastIndex = cell.Row
        If IsEmpty(cell) Then
            rowIndex = cell.Row
            Exit For
        End If
    Next cell
    
    If rowIndex = 0 Then
        'there are not empty cells write to the end of column
        rowIndex = lastIndex + 1
    End If
    
    getFirstEmptyCellRowIndex = rowIndex
End Function

'return the indexes of specified columns
Public Function getColumnIndexes(sheet As Worksheet, tableName As String, columnNames() As String) As Collection
    Dim columnsCount As Integer
    Dim columnIndex As Integer
    Dim idx As Integer
    Dim columnIndexes As New Collection
    
    columnsCount = UBound(columnNames)
    For idx = 0 To columnsCount - 1
        columnIndex = getColumnIndex(sheet, tableName, columnNames(idx))
        columnIndexes.Add (columnIndex)
    Next idx
    
    Set getColumnIndexes = columnIndexes
End Function

'sort specified table by specified column in ascending order, headers are expected
Public Sub sortTable(sheet As Worksheet, tableName As String, columnName As String)
     sheet.Range(tableName).Sort key1:=sheet.Range(tableName & "[" & columnName & "]"), order1:=xlAscending, Header:=xlYes
End Sub

'returns a row indexes from specified range
Public Function getRowIndexes(source As Range) As Collection
    Dim rowIndexes As New Collection
    Dim currentRow As Range
    
    For Each currentRow In source.Rows
        rowIndexes.Add (currentRow.Row)
    Next currentRow
    
    Set getRowIndexes = rowIndexes
End Function

Public Function AreRowsTheSame(firstSheet As Worksheet, firstRowIndex As Integer, firstColumnIndexes As Collection, secondSheet As Worksheet, secondRowIndex As Integer, secondColumnIndexes As Collection) As Boolean
    Dim result As Boolean
    Dim idx As Integer
    Dim firstValue As Variant
    Dim secondValue As Variant
    
    result = True
    For idx = 1 To firstColumnIndexes.Count
        firstValue = firstSheet.Cells(firstRowIndex, firstColumnIndexes(idx)).Value
        secondValue = secondSheet.Cells(secondRowIndex, secondColumnIndexes(idx)).Value
        If firstValue <> secondValue Then
            result = False
            Exit For
        End If
    Next idx
    
    AreRowsTheSame = result
End Function







