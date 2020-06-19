VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorkbookInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_wbk As Workbook
Private m_sheetNames As Collection

' Initialise the class with a workbook instance
Public Sub Init(wbk As Workbook)
    Set m_wbk = wbk
    Set m_sheetNames = GetSheetNames
End Sub

' Return a collection of sheet names in the workbook instance set in Init()
Private Function GetSheetNames() As Collection
    Dim sheetNames As Collection: Set sheetNames = New Collection
    Dim wSh As Worksheet
    For Each wSh In m_wbk.Worksheets
        sheetNames.Add wSh.Name
    Next wSh
    Set GetSheetNames = sheetNames
End Function

' Return a collection of all values of the first row of the worksheet with the given name
Public Function GetColumnNamesForSheet(sheetName As String) As Collection
    Dim wSheet As Worksheet
    Dim usedSheetRange
    Dim sheetColumnCount As Long
    Dim i As Long
    Dim columnNamesForSheet As Collection: Set columnNamesForSheet = New Collection
    Dim cellValue As String
    Set wSheet = m_wbk.Worksheets(sheetName)
    Set usedSheetRange = wSheet.UsedRange
    sheetColumnCount = usedSheetRange.Columns.Count
    For i = 1 To sheetColumnCount
        cellValue = CStr(usedSheetRange.Cells(1, i).Value)
        columnNamesForSheet.Add cellValue
    Next i
    Set GetColumnNamesForSheet = columnNamesForSheet
End Function

' Return the sheet names collection
Public Property Get sheetNames() As Collection
    Set sheetNames = m_sheetNames
End Property


