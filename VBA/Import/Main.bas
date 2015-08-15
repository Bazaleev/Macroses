Attribute VB_Name = "Main"
Option Explicit
Sub import()
    Dim sourceSheet As Worksheet
    Dim sourceTableName As String
    Dim targetTableName As String
    Dim sourceColumnNames(6) As String
    Dim targetColumnNames(6) As String
    Dim targetSheetName As String
    Dim accountNo As String
           
    Application.ScreenUpdating = False
    
    'init the list of column names which data should be copied
    sourceColumnNames(0) = "Дата оплаты"
    sourceColumnNames(1) = "CF code"
    sourceColumnNames(2) = "Получатель платежа"
    sourceColumnNames(3) = "Сумма"
    sourceColumnNames(4) = "Комментарии"
    sourceColumnNames(5) = "НДС"
    
    'init the list of column names in this file to which data should be copied
    targetColumnNames(0) = "Date"
    targetColumnNames(1) = "CF code"
    targetColumnNames(2) = "CounterParty"
    targetColumnNames(3) = "Amount acc.cur"
    targetColumnNames(4) = "Comment"
    targetColumnNames(5) = "VAT index"
   
    sourceTableName = "Таблица1"
    Set sourceSheet = Utils.getSheet("payments_list.xlsm", "payments")
    
    'copy to first sheet alpha file
    targetSheetName = "CYB Cash EUR"
    targetTableName = "Таблица82"
    accountNo = "CYB Cash EUR"
    Main.ImportDataToSheet sourceSheet, sourceTableName, targetSheetName, targetTableName, accountNo, sourceColumnNames, targetColumnNames
    
    'copy to second sheet alpha file
    targetTableName = "Таблица823"
    accountNo = "CYB Cash USD"
    targetSheetName = "CYB Cash USD"
    Main.ImportDataToSheet sourceSheet, sourceTableName, targetSheetName, targetTableName, accountNo, sourceColumnNames, targetColumnNames
     
    Application.ScreenUpdating = True
End Sub

'returns a list indexeses of row to copy in file
Private Function getRowsToCopy(sheet As Worksheet, tableName As String, accountNo As String) As Collection
    Dim rowsToCopy As New Collection
    Dim paidPaimentsRows As Collection
    Dim bankColumnIndex As Integer
    Dim rowIdx As Variant
    
    bankColumnIndex = Utils.getColumnIndex(sheet, tableName, "bank")
    Set paidPaimentsRows = getPaidPaimentsRows(sheet, tableName)
    
    For Each rowIdx In paidPaimentsRows
        If sheet.Cells(rowIdx, bankColumnIndex) = accountNo Then
            rowsToCopy.Add (rowIdx)
        End If
    Next rowIdx
    
    Set getRowsToCopy = rowsToCopy
End Function

'returns the row number of paiments that were paid
Private Function getPaidPaimentsRows(sheet As Worksheet, tableName As String) As Collection
    Dim dateColumn As Range
    Dim rowsToCopy As New Collection
    Dim cell As Range
    
    Set dateColumn = Utils.getColumn(sheet, tableName, "Дата оплаты")
    For Each cell In dateColumn.Cells
        If Not IsEmpty(cell.Value) Then
            rowsToCopy.Add (cell.Row)
        End If
    Next cell
    
    Set getPaidPaimentsRows = rowsToCopy
End Function

'copies values with respective indexes from the specified columns to target columns
Private Sub copyTableValues(sourceSheet As Worksheet, sourceTableName As String, sourceColumnNames() As String, targetSheet As Worksheet, targetTableName As String, targetColumnNames() As String, rowsToCopy As Collection)
    Dim sourceColumnIndex As Integer
    Dim targetColumn As Range
    Dim columnsCount As Integer
    Dim sourceColumnName As String
    Dim targetColumnName As String
    Dim columnIndex As Integer
    Dim rowIdx As Integer
    Dim targetRowIndex As Integer
    Dim targetColumnIndex As Integer
    
    columnsCount = UBound(sourceColumnNames)
    If columnsCount = UBound(targetColumnNames) Then
        'the number of names to copy from should be equal number of names to copy to
        
        For columnIndex = 0 To columnsCount - 1
            sourceColumnName = sourceColumnNames(columnIndex)
            sourceColumnIndex = Utils.getColumnIndex(sourceSheet, sourceTableName, sourceColumnName)
            
            targetColumnName = targetColumnNames(columnIndex)
            Set targetColumn = Utils.getColumn(targetSheet, targetTableName, targetColumnName)
            
            targetColumnIndex = targetColumn.Column
            targetRowIndex = Utils.getFirstEmptyCellRowIndex(targetColumn)
            For rowIdx = 1 To rowsToCopy.Count
                targetSheet.Cells(targetRowIndex, targetColumnIndex).Value = sourceSheet.Cells(rowsToCopy(rowIdx), sourceColumnIndex).Value
                targetRowIndex = targetRowIndex + 1
            Next rowIdx
        Next columnIndex
    End If
End Sub

'copyies required data to specified sheet and table
Private Sub ImportDataToSheet(sourceSheet As Worksheet, sourceTableName As String, targetSheetName As String, targetTableName As String, accountNo As String, sourceColumnNames() As String, targetColumnNames() As String)
    Dim targetSheet As Worksheet
    Dim rowsToCopy As Collection
    Dim targetColumnIndexes As Variant
    Dim columnsCount As Integer
    Dim columnIndex As Integer
    Dim idx As Integer
    Dim targetTableRange As Range
    
    columnsCount = UBound(targetColumnNames)
    
    Set rowsToCopy = getRowsToCopy(sourceSheet, sourceTableName, accountNo)
    
    Set targetSheet = Utils.getSheetInCurrentWorkbook(targetSheetName)
    ReDim targetColumnIndexes(0 To columnsCount - 1)
    For idx = 0 To columnsCount - 1
        columnIndex = Utils.getColumnIndex(targetSheet, targetTableName, targetColumnNames(idx))
        targetColumnIndexes(idx) = columnIndex
    Next idx
    
    copyTableValues sourceSheet, sourceTableName, sourceColumnNames, targetSheet, targetTableName, targetColumnNames, rowsToCopy
    
    'remove duplicates rows
    Set targetTableRange = targetSheet.Range(targetTableName)
    targetTableRange.RemoveDuplicates Columns:=(targetColumnIndexes), Header:=xlYes

    'sorting
    targetSheet.Range(targetTableName).Sort key1:=targetSheet.Range(targetTableName & "[" & targetColumnNames(0) & "]"), order1:=xlAscending, Header:=xlYes
End Sub
                            








