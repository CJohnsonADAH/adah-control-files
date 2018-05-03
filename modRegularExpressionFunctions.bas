Attribute VB_Name = "modRegularExpressionFunctions"
Option Explicit
Option Compare Database

'**
'* RegexMatch: Return True if the given string value matches the given Regex pattern
'*
'* @param Variant value Value to check for a regular expression match
'* @param String pattern Regular expression pattern
'* @param Boolean MatchCase If True, letters in the pattern must match case (/a/ matches "a", not "A")
'*      If False or omitted, letters in the pattern match across upper/lowercase (/a/ matches "a" or "A")
'
'* @return Boolean True if the given string value matches the given Regex pattern, False otherwise
'**
Public Function RegexMatch(Value As Variant, Pattern As String, Optional ByVal MatchCase As Boolean) As Boolean
    If IsNull(Value) Then Exit Function
    ' Using a static, we avoid re-creating the same regex object for every call '
    Static Regex As Object
    ' Initialise the Regex object '
    If Regex Is Nothing Then
        Set Regex = CreateObject("vbscript.regexp")
        With Regex
            .Global = True
            .MultiLine = True
        End With
    End If

    With Regex
            .IgnoreCase = Not MatchCase
    End With

    ' Update the regex pattern if it has changed since last time we were called '
    If Regex.Pattern <> Pattern Then Regex.Pattern = Pattern
    ' Test the value against the pattern '
    RegexMatch = Regex.Test(Value)
End Function

'**
'* RegexComponent: get the contents of a numbered back-reference component
'* provided the given string value matches the given Regex pattern
'*
'* @param Variant value Value to check for a regular expression match
'* @param String pattern Regular expression pattern to match it against
'* @param Integer Part Number of the sub-pattern to return the matching contents for, beginning with 1 for the first ($1)
'* @param Boolean MatchCase If True, letters in the pattern must match case (/a/ matches "a", not "A")
'*      If False or omitted, letters in the pattern match across upper/lowercase (/a/ matches "a" or "A")
'*
'* @return String The contents of the matching back-reference, or an empty string if there is no match.
'**
Public Function RegexComponent(Value As Variant, Pattern As String, Part As Integer, Optional ByVal MatchCase As Boolean) As String
    Dim cMatches As Variant
    Dim iMatch As Variant
    Dim iSubMatch As Variant
    Dim index As Integer
    
    If IsNull(Value) Then Exit Function
    ' Using a static, we avoid re-creating the same regex object for every call '
    Static Regex As Object
    ' Initialise the Regex object '
    If Regex Is Nothing Then
        Set Regex = CreateObject("vbscript.regexp")
        With Regex
            .Global = True
            .MultiLine = True
        End With
    End If
    
    With Regex
        .IgnoreCase = Not MatchCase
    End With
    
    ' Update the regex pattern if it has changed since last time we were called '
    If Regex.Pattern <> Pattern Then Regex.Pattern = Pattern
    ' Test the value against the pattern '
    Set cMatches = Regex.Execute(Value)
    For Each iMatch In cMatches
        If Part = 0 Then
            RegexComponent = iMatch.Value
            Exit For
        Else
            For Each iSubMatch In iMatch.SubMatches
                index = index + 1
                If index = Part Then
                    RegexComponent = iSubMatch
                    Exit For
                End If
            Next iSubMatch
            Exit For
        End If
    Next iMatch
    
End Function

Public Function RegexReplace(Value As Variant, Pattern As String, Replace As String, Optional ByVal MatchCase As Boolean, Optional ByVal OnlyOne As Boolean) As String
    ' Using a static, we avoid re-creating the same regex object for every call '
    Static Regex As RegExp
    
    Dim hasMatch As Boolean
    Dim Result As String
    
    If IsNull(Value) Then Exit Function
    
    ' Initialise the Regex object '
    If Regex Is Nothing Then
        Set Regex = New RegExp
        With Regex
            .MultiLine = True
        End With
    End If
    
    Result = CStr(Value)
    With Regex
        .Pattern = Pattern
        .Global = Not OnlyOne
        .IgnoreCase = Not MatchCase
    End With
            
    RegexReplace = Regex.Replace(Result, Replace)

End Function

Public Function RegexSplit(ByVal Text As String, ByVal Pattern As String, Optional ByVal MatchCase As Boolean, Optional ByVal DelimCapture As Boolean) As String()
    Dim aWords() As String
    Dim sReplace As String
    
    If DelimCapture Then
        Let sReplace = Chr$(0) & "$1" & Chr$(0)
    Else
        Let sReplace = Chr$(0)
    End If
    
    Let Text = RegexReplace(Value:=Text, Pattern:="(" & Pattern & ")", Replace:=sReplace, MatchCase:=MatchCase)
    Let aWords = Split(Expression:=Text, Delimiter:=Chr$(0))
    Let RegexSplit = aWords
End Function
