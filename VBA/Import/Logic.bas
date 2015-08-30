Attribute VB_Name = "Logic"
Option Explicit

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
    Dim sourceColumnName As String
    Dim targetColumnName As String
    Dim columnIndex As Integer
    Dim rowIdx As Integer
    Dim targetRowIndex As Integer
    Dim targetColumnIndex As Integer
    Dim startRowIndex As Integer
    Dim sourceCell As Range
    Dim targetCell As Range
    Dim columnsCount As Integer
    
    columnsCount = UBound(sourceColumnNames)
    If columnsCount = UBound(targetColumnNames) Then
        'the number of names to copy from should be equal number of names to copy to
        
        For columnIndex = 0 To columnsCount - 1
            sourceColumnName = sourceColumnNames(columnIndex)
            sourceColumnIndex = Utils.getColumnIndex(sourceSheet, sourceTableName, sourceColumnName)
            
            targetColumnName = targetColumnNames(columnIndex)
            Set targetColumn = Utils.getColumn(targetSheet, targetTableName, targetColumnName)
            If startRowIndex = 0 Then
                'all values are copied starting from first empty cell in first column
                startRowIndex = Utils.getFirstEmptyCellRowIndex(targetColumn)
            End If
            
            targetColumnIndex = targetColumn.Column
            targetRowIndex = startRowIndex
            
            For rowIdx = 1 To rowsToCopy.Count
                Set sourceCell = sourceSheet.Cells(rowsToCopy(rowIdx), sourceColumnIndex)
                Set targetCell = targetSheet.Cells(targetRowIndex, targetColumnIndex)
                
                'copy value and format from source cell
                targetCell.Value = sourceCell.Value
                targetCell.NumberFormat = sourceCell.NumberFormat
                
                targetRowIndex = targetRowIndex + 1
            Next rowIdx
        Next columnIndex
    End If
End Sub

'copyies required data to specified sheet and table
Public Sub importDataToSheet(sourceSheet As Worksheet, sourceTableName As String, targetSheetName As String, targetTableName As String, accountNo As String, sourceColumnNames() As String, targetColumnNames() As String)
    Dim targetSheet As Worksheet
    Dim rowsToCopy As Collection
    Dim targetColumnIndexes As Collection
    Dim sourceColumnIndexes As Collection
    Dim idx As Integer
    Dim targetTableRange As Range
    Dim newRows As Collection
    
    Set rowsToCopy = getRowsToCopy(sourceSheet, sourceTableName, accountNo)
    
    Set targetSheet = Utils.getSheetInCurrentWorkbook(targetSheetName)
    Set targetColumnIndexes = Utils.getColumnIndexes(targetSheet, targetTableName, targetColumnNames)
    Set sourceColumnIndexes = Utils.getColumnIndexes(sourceSheet, sourceTableName, sourceColumnNames)
    Set targetTableRange = targetSheet.Range(targetTableName)
    
    'sort table for optimised search
    Utils.sortTable sheet:=targetSheet, tableName:=targetTableName, columnName:=targetColumnNames(0)
    Set newRows = getNewOnlyRows(sourceSheet, sourceColumnIndexes, rowsToCopy, targetSheet, targetTableRange, targetColumnIndexes)
    
    If newRows.Count > 0 Then
        'new values is found copy them
        copyTableValues sourceSheet, sourceTableName, sourceColumnNames, targetSheet, targetTableName, targetColumnNames, rowsToCopy
        
        'now sort again with new data
        Utils.sortTable sheet:=targetSheet, tableName:=targetTableName, columnName:=targetColumnNames(0)
    End If
End Sub

'filters input values and returns new only. It is expected that table range will be sorted by first column
Private Function getNewOnlyRows(sourceSheet As Worksheet, sourceColumnIndexes As Collection, rowsToCheck As Collection, targetSheet As Worksheet, targetTable As Range, targetColumnIndexes As Collection) As Collection
    Dim targetRowIndexes As Collection
    Dim newRows As New Collection
    Dim rowIdx As Integer
    Dim isNew As Boolean
    Dim targetRowIndex As Variant
    Dim firstSourceColumn As Integer
    Dim firstTargetColumn As Integer
    Dim sourceValue As Variant
    Dim targetValue As Variant
    Dim sourceRowIndex As Integer
    
    firstSourceColumn = sourceColumnIndexes(1)
    firstTargetColumn = targetColumnIndexes(1)
    
    Set targetRowIndexes = Utils.getRowIndexes(targetTable)
    For rowIdx = 1 To rowsToCheck.Count
        isNew = True
        
        sourceRowIndex = CInt(rowsToCheck(rowIdx))
        sourceValue = sourceSheet.Cells(sourceRowIndex, firstSourceColumn).Value
        For Each targetRowIndex In targetRowIndexes
            If Utils.AreRowsTheSame(sourceSheet, sourceRowIndex, sourceColumnIndexes, targetSheet, CInt(targetRowIndex), targetColumnIndexes) Then
                isNew = False
                Exit For
            End If
            
            targetValue = targetSheet.Cells(targetRowIndex, firstTargetColumn).Value
            If targetValue > sourceValue Then
                'values further are all different as range already sorted by this column
                Exit For
            End If
        Next targetRowIndex
        
        If isNew Then
            newRows.Add (rowsToCheck(rowIdx))
        End If
    Next rowIdx
    
    Set getNewOnlyRows = newRows
End Function

