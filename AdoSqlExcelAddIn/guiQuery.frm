VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} guiQuery 
   Caption         =   "Excel SQL Query"
   ClientHeight    =   4750
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8950
   OleObjectBlob   =   "guiQuery.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "guiQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_qry As clsQuery
Private m_wbkInfo As clsWorkbookInfo
Private m_activeWbk As Workbook
Private m_newWorksheet As Worksheet
Private m_selectedColumnNames As Collection
Private m_rngQueryDataFirstCell As Range

' Adds sheet and sets first cell in new sheet as range for query output
Private Sub AddNewSheet(newSheetName As String)
    Set m_newWorksheet = m_activeWbk.Worksheets.Add
    Set m_rngQueryDataFirstCell = m_newWorksheet.Range("A1")
    m_newWorksheet.Name = newSheetName
End Sub

' Close form
Private Sub btnCloseForm_Click()
    Unload Me
End Sub

' Executes the SQL and adds the query output to the new sheet.
' To Do: Error trapping is crude and it should be re-factored.
Private Sub btnGenerateQueryOutput_Click()
    Dim sql As String: sql = Me.txtGeneratedSQL
    Dim rs As ADODB.Recordset
    Dim colNames As Collection
    Dim newSheetName As String
    Set m_qry = New clsQuery
    On Error GoTo ErrTrap:
    newSheetName = GUI.UserInput("Name for new sheet:", "New Sheet", "")
    AddNewSheet newSheetName
    Set rs = m_qry.GetRecordsetForQuery(sql)
    Set colNames = GUI.GetFieldNamesFromRecordset(rs)
    GUI.CollectionToARow colNames, m_rngQueryDataFirstCell
    m_rngQueryDataFirstCell.Offset(1, 0).CopyFromRecordset rs
    Exit Sub
ErrTrap:
    Application.DisplayAlerts = False
    m_newWorksheet.Delete
    Application.DisplayAlerts = True
    MsgBox Err.Description
End Sub

' Based on the selections made from the sheets and column names listboxes, generate an SQL string
'  and write it to the SQL textbox.
Private Sub btnGenerateSQL_Click()
    Dim targetSheetName As String
    Set m_selectedColumnNames = GUI.GetSelectedLboxItems(Me.lboxColumnNames)
    Dim columnNames As String: columnNames = GUI.JoinCollection(m_selectedColumnNames, "," & vbCrLf, " ")
    targetSheetName = GUI.GetSelectedLboxItem(Me.lboxSheetNames)
    Dim sql As String: sql = "SELECT" & vbCrLf & columnNames & vbCrLf & "FROM" & vbCrLf & "  [" & targetSheetName & "$]"
    Me.txtGeneratedSQL = sql
End Sub

' Open and read an SQL script file and put its contents into the SQL textbox.
Private Sub btnOpenAndReadSqlFile_Click()
    Dim sqlFileName As String
    Dim sqlText As String
    sqlFileName = GUI.GetFileNameToRead
    sqlText = GUI.ReadInSql(sqlFileName)
    Me.txtGeneratedSQL = sqlText
End Sub

' Save the contents of the SQL textbox to an external file for future use.
Private Sub btnSaveSqlToFile_Click()
    Dim sqlScriptFileNameToSave As String
    Dim sqlText As String: sqlText = Me.txtGeneratedSQL
    sqlScriptFileNameToSave = GUI.GetFilNameToSave
    GUI.WriteSqlToFile sqlText, sqlScriptFileNameToSave
End Sub

' Update the contents of the column names listbox whenever a different sheet is selected.
Private Sub lboxSheetNames_Change()
    Dim selectedSheetName As String
    selectedSheetName = GUI.GetSelectedLboxItem(Me.lboxSheetNames)
    GUI.ClearListbox Me.lboxColumnNames
    GUI.AddCollectionToListbox m_wbkInfo.GetColumnNamesForSheet(selectedSheetName), Me.lboxColumnNames
End Sub

' Create required object instances when the form loads.
Private Sub UserForm_Initialize()
    Set m_qry = New clsQuery
    Set m_wbkInfo = New clsWorkbookInfo
    Set m_activeWbk = ActiveWorkbook
    m_wbkInfo.Init m_activeWbk
    GUI.AddCollectionToListbox m_wbkInfo.sheetNames, Me.lboxSheetNames
    GUI.AddCollectionToListbox m_wbkInfo.GetColumnNamesForSheet(m_wbkInfo.sheetNames(1)), Me.lboxColumnNames
End Sub
