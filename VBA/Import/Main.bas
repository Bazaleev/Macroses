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
    sourceColumnNames(0) = "���� ������"
    sourceColumnNames(1) = "CF code"
    sourceColumnNames(2) = "���������� �������"
    sourceColumnNames(3) = "�����"
    sourceColumnNames(4) = "�����������"
    sourceColumnNames(5) = "���"
    
    'init the list of column names in this file to which data should be copied
    targetColumnNames(0) = "Date"
    targetColumnNames(1) = "CF code"
    targetColumnNames(2) = "CounterParty"
    targetColumnNames(3) = "Amount acc.cur"
    targetColumnNames(4) = "Comment"
    targetColumnNames(5) = "VAT index"
   
    sourceTableName = "�������1"
    Set sourceSheet = Utils.getSheet("payments_list.xlsm", "payments")
    
    'copy to first sheet alpha file
    targetSheetName = "CYB Cash EUR"
    targetTableName = "�������82"
    accountNo = "CYB Cash EUR"
    Logic.importDataToSheet sourceSheet, sourceTableName, targetSheetName, targetTableName, accountNo, sourceColumnNames, targetColumnNames
    
    'copy to second sheet alpha file
    targetTableName = "�������823"
    accountNo = "CYB Cash USD"
    targetSheetName = "CYB Cash USD"
    Logic.importDataToSheet sourceSheet, sourceTableName, targetSheetName, targetTableName, accountNo, sourceColumnNames, targetColumnNames
     
    Application.ScreenUpdating = True
    
    MsgBox "Successeful!", vbOKOnly, "Done!"
End Sub


                            








