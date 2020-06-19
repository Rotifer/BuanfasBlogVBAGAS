Attribute VB_Name = "GUI"
Option Explicit
' Procedures and functions used by the form

' For ribbon to show form
Public Sub ShowAdoExcelSQLQueryForm()
    guiQuery.Show
End Sub

' Add all items in the given collection to the given listbox
Public Sub AddCollectionToListbox(items As Collection, lbox As MSForms.ListBox)
    Dim item As Variant
    For Each item In items
        lbox.AddItem item
    Next item
End Sub

' Remove all items from the given listbox
Public Sub ClearListbox(lbox As MSForms.ListBox)
    lbox.Clear
End Sub

' Return the value of the selected listbox item. Only makes sense for listboxes
'  with multiselect set as False
Public Function GetSelectedLboxItem(lbox As MSForms.ListBox) As String
    Dim i As Long
    For i = 0 To lbox.ListCount - 1
        If lbox.Selected(i) Then
            GetSelectedLboxItem = CStr(lbox.List(i))
            Exit Function
        End If
    Next i
End Function

' Return a collection of all the selected items for the given listbox
Public Function GetSelectedLboxItems(lbox As MSForms.ListBox) As Collection
    Dim item As Variant
    Dim items As Collection: Set items = New Collection
    Dim i As Long
    For i = 0 To lbox.ListCount - 1
        If lbox.Selected(i) Then
            item = lbox.List(i)
            items.Add CStr(item)
        End If
    Next i
    Set GetSelectedLboxItems = items
End Function

' Given a collection, join it with the given delimiter
Public Function JoinCollection(items As Collection, delimiter As String, leftPad As String) As String
    Dim arrTemp() As String
    Dim i As Long
    On Error GoTo ErrTrap
    ReDim arrTemp(items.Count - 1)
    For i = 1 To items.Count
        arrTemp(i - 1) = leftPad & CStr(items(i))
    Next i
    JoinCollection = Join(arrTemp, delimiter)
    Exit Function
ErrTrap:
    If Err.Description = "Subscript out of range" Then
        JoinCollection = "No columns selected"
    Else
        Err.Raise 1000, "GUI.JoinCollection", "Unable to Join given Collection"
    End If
End Function

' Return a collection of column names for the given recordset
Public Function GetFieldNamesFromRecordset(rs As ADODB.Recordset) As Collection
    Dim field As Variant
    Dim fieldNames As Collection: Set fieldNames = New Collection
    For Each field In rs.Fields
        fieldNames.Add field.Name
    Next field
    Set GetFieldNamesFromRecordset = fieldNames
End Function

' Given a collection put each item in a sheet row starting at the given range
Public Sub CollectionToARow(items As Collection, startingCell As Range)
    Dim item As Variant
    Dim idx As Long: idx = 0
    For Each item In items
        startingCell.Offset(0, idx).Value = CStr(item)
        idx = idx + 1
    Next
End Sub

' Display a dialog to get a selected file name and return the file name
Public Function GetFileNameToRead() As String
    Dim fDlg As FileDialog
    Dim result As Integer
    Set fDlg = Application.FileDialog(msoFileDialogFilePicker)
    fDlg.AllowMultiSelect = False
    fDlg.title = "Select an SQL Script"
    fDlg.InitialFileName = "C:\"
    fDlg.Filters.Clear
    fDlg.Filters.Add "SQL Scripts", "*.sql"
    fDlg.Filters.Add "All files", "*.*"
     
    'Show the dialog. -1 means success!
    If fDlg.Show = -1 Then
       GetFileNameToRead = fDlg.SelectedItems(1)
    End If
End Function

' Return the entire text from an SQL script file
Public Function ReadInSql(sqlFullFileName As String) As String
    Dim fso As Scripting.FileSystemObject
    Dim tsSqlScript As Scripting.TextStream
    Dim sqlText As String
    Set fso = New Scripting.FileSystemObject
    Set tsSqlScript = fso.OpenTextFile(sqlFullFileName, ForReading)
    sqlText = tsSqlScript.ReadAll
    tsSqlScript.Close
    Set tsSqlScript = Nothing
    Set fso = Nothing
    ReadInSql = sqlText
End Function

' Display a dialog to return a full file name for an SQL script
' Tried to use msoFileDialogSaveAs but cannot save as .sql
Public Function GetFilNameToSave() As String
    Dim result As Variant
    result = Application.GetSaveAsFilename(, , , "Save SQL to File", "Save As")
    'checks to make sure the user hasn't canceled the dialog
    If result <> False Then
        GetFilNameToSave = result
    End If
End Function

' Writes the given SQL text to the given SQL script file name and returns True
Public Function WriteSqlToFile(sqlText As String, sqlScriptFileNameToSave As String) As Boolean
    Dim fso As Scripting.FileSystemObject
    Dim tsSqlScript As Scripting.TextStream
    Set fso = New Scripting.FileSystemObject
    Set tsSqlScript = fso.OpenTextFile(sqlScriptFileNameToSave, ForWriting, True)
    tsSqlScript.Write sqlText
    tsSqlScript.Close
    Set tsSqlScript = Nothing
    Set fso = Nothing
    WriteSqlToFile = True
End Function

' Handle user input prompts for "Cancel" or empty string
Public Function UserInput(prompt As String, title As String, default As String) As String
    Dim result As String
    result = InputBox(prompt, title, default)
    If StrPtr(result) = 0 Then
        UserInput = ""
    ElseIf result = vbNullString Then
       UserInput = ""
    Else
        UserInput = result
    End If
End Function
