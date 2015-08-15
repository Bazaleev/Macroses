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







