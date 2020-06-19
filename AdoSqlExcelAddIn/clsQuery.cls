VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Manages the SQL connection and executes queries
Private m_xlConn As ADODB.Connection
Private m_activeWbk As Workbook
Private m_xlFileName As String

' Create the connection to the active workbook
Private Sub Class_Initialize()
    Dim connStr As String
    Set m_activeWbk = ActiveWorkbook
    m_xlFileName = m_activeWbk.FullName
    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
          m_xlFileName & ";Extended Properties=""Excel 12.0 Macro;HDR=YES"""
    Set m_xlConn = New ADODB.Connection
    m_xlConn.ConnectionString = connStr
    m_xlConn.Open
End Sub

' Return a recordset for the given SQL query
Public Function GetRecordsetForQuery(sql As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open sql, m_xlConn
    Set GetRecordsetForQuery = rs
End Function

' Get the full name with path for the active workbook
Public Property Get ExcelFileName() As String
    ExcelFileName = m_xlFileName
End Property

' Clean-up
Private Sub Class_Terminate()
    m_xlConn.Close
    Set m_xlConn = Nothing
End Sub
