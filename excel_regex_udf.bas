Option Explicit

' All functions in this module assume:
'     "Tools>References>Microsoft VBScript Regular Expressions" is referenced.
' Windows only, will not work on a Mac (Sorry!)


' Returns a Boolean to indicate if a piece of text matches a regular expression.
Public Function REGEXMATCH(text As String, regular_expression As String) As Boolean
    Dim re As New RegExp
    re.Pattern = regular_expression
    re.IgnoreCase = False
    re.MultiLine = False
    REGEXMATCH = re.Test(text)
End Function

' Returns the first match in the text that matches the regular expression.
Public Function REGEXEXTRACT(text As String, regular_expression As String) As String
    Dim re As New RegExp
    Dim matches As MatchCollection
    Dim firstMatch As Match
    If Not REGEXMATCH(text, regular_expression) Then
        REGEXEXTRACT = ""
        Exit Function
    End If
    re.Pattern = regular_expression
    re.Global = False
    re.IgnoreCase = False
    re.MultiLine = False
    Set matches = re.Execute(text)
    Set firstMatch = matches.Item(0) 'This collection is zero-based!!
    REGEXEXTRACT = firstMatch.Value
End Function

' Replaces all matches of the regex in the target text with the replacement and retrns the replaced text.
Public Function REGEXREPLACE(text As String, regular_expression As String, replacement As String) As String
    Dim re As New RegExp
    re.Pattern = regular_expression
    re.Global = True
    re.IgnoreCase = False
    re.MultiLine = False
    REGEXREPLACE = re.Replace(text, replacement)
End Function
