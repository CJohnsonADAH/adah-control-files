Attribute VB_Name = "ModControlFilesDatabase"
'**
'* @module ModControlFilesDatabase
'*
'* @uses RegExp Tools > References > Microsoft VBScript Regular Expressions 5.5
'*
'* @author Charles Johnson
'* @version 2017.0921
'**
Option Compare Binary
Option Explicit

'**
'* SecureCreatorRecord
'*
'* @uses DAO.Recordset
'*
'* @param String Creator A two- or three-character alphanumeric CurName code
'**
Sub SecureCreatorRecord(Creator As String)
    Dim sSuggestion As String
    Dim sAgencyName As String
    Dim rsAccessions As DAO.Recordset
    
    Set rsAccessions = CurrentDb.OpenRecordset("SELECT * FROM Sheet1 WHERE CurName = '" & Creator & "'")
    If Not rsAccessions.EOF Then
        If Not IsNull(rsAccessions!AgencyName) Then
            sSuggestion = rsAccessions!AgencyName
        Else
            sSuggestion = Creator
        End If
    Else
        sSuggestion = Creator
    End If
    
    Set rsAccessions = CurrentDb.OpenRecordset("SELECT * FROM Creators WHERE CurName = '" & Creator & "'")
    If rsAccessions.EOF Then
        sAgencyName = InputBox(Prompt:=Creator & " = ", Title:="AgencyName for CurName " & Creator, Default:=sSuggestion)
        If Len(sAgencyName) > 0 Then
            Set rsAccessions = CurrentDb.OpenRecordset("Creators")
            rsAccessions.AddNew
            rsAccessions!CurName = Creator
            rsAccessions!AgencyName = sAgencyName
            rsAccessions.Update
        End If
    End If

End Sub

Sub deleteAccnScanAttachments(ACCN As String, ByVal FileName As String, ByVal FilePath As String)
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef
    Dim sFilePathMatch As String
    
    'FIXME: This currently seems to be broken, and returns no records from the WHERE
    sFilePathMatch = "FilePath=[paramFilePath]"
    If Len(FilePath) = 0 Then
        sFilePathMatch = "(" & sFilePathMatch & " OR FilePath IS NULL)"
    End If
    
    'check the database for comments flagged with this CollectionNumber
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qDeleteAccnScans": On Error GoTo 0
    
    Set oQuery = CurrentDb.CreateQueryDef( _
        Name:="qDeleteAccnScans", _
        SQLText:="DELETE FROM AccnScans WHERE " _
        & "ACCN=[paramAccn] AND " _
        & "FileName=[paramFileName] AND " _
        & sFilePathMatch _
    )
    oQuery.Parameters("paramAccn") = ACCN
    oQuery.Parameters("paramFileName") = FileName
    oQuery.Parameters("paramFilePath") = FilePath
    
    oQuery.Execute
    
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qDeleteAccnScans": On Error GoTo 0

    Set oQuery = Nothing

End Sub

Sub putAccnScanAttachments(ACCN As String, ByVal FileName As String, ByVal FilePath As String)
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef
    Dim cScans As Collection
    Dim vScan As Variant
    Dim alreadyGotOne As Boolean
    
    'check the database for comments flagged with this CollectionNumber
    alreadyGotOne = False
    Set cScans = getAccnScanAttachments(ACCN)
    For Each vScan In cScans
        If vScan(2) = FileName Then
            alreadyGotOne = True
            Debug.Print "I DON'T THINK HE'LL BE VERY INTERESTED IN " & FileName & ". HE'S ALREADY GOT ONE, YOU SEE? IT'SA VERY NICE-UH!"
        End If
    Next vScan
    
    If Not alreadyGotOne Then
        On Error Resume Next: CurrentDb.QueryDefs.Delete "qInsertAccnScans": On Error GoTo 0
        
        Set oQuery = CurrentDb.CreateQueryDef( _
            Name:="qInsertAccnScans", _
            SQLText:="INSERT INTO AccnScans (ACCN, NonAccnId, FileName, FilePath) VALUES (" _
            & "[paramAccn], " _
            & "NULL, " _
            & "[paramFileName], " _
            & "[paramFilePath])" _
        )
        oQuery.Parameters("paramAccn") = Trim(UCase(ACCN))
        oQuery.Parameters("paramFileName") = Trim(FileName)
        oQuery.Parameters("paramFilePath") = Trim(FilePath)
        
        oQuery.Execute

        On Error Resume Next: CurrentDb.QueryDefs.Delete "qInsertAccnScans": On Error GoTo 0
        
        Set oQuery = Nothing
    End If
    
End Sub

Function getAccnScanAttachments(ACCN As String) As Collection
    Dim aBits(1 To 2) As String
    Dim cAttachments As New Collection
    
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef
    Dim I As Integer, iCount As Integer

    'check the database for comments flagged with this CollectionNumber
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qScanFilesByAccn": On Error GoTo 0
    Set oQuery = CurrentDb.CreateQueryDef( _
        Name:="qScanFilesByAccn", _
        SQLText:="SELECT * FROM AccnScans " _
        & "WHERE ACCN = [paramAccn]" _
    )
    oQuery.Parameters("paramAccn") = Trim(UCase(ACCN))
        
    Set Rs = oQuery.OpenRecordset
    
    Do Until Rs.EOF
        If Not IsNull(Rs!FilePath) Then
            aBits(1) = Rs!FilePath
        Else
            aBits(1) = ""
        End If
        If Not IsNull(Rs!FileName) Then
            aBits(2) = Rs!FileName
        Else
            aBits(2) = ""
        End If
        cAttachments.Add aBits
        Rs.MoveNext
    Loop

    Set Rs = Nothing
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qScanFilesByAccn": On Error GoTo 0
    
    Set getAccnScanAttachments = cAttachments

End Function

'**
'*
'**
Function getScanFileName(ByVal FileName As String, Optional ByVal FilePath As String) As String
    Dim sFilePath As String
    Dim sRelativePath As String
    Dim f As String
    Dim sDrive As String
    Dim sPattern As String
    Dim aWords() As String
    Dim cSearchDirs As New Collection
    Dim vSearchDir As Variant
    
    sDrive = "S:"
    If Len(FilePath) > 0 Then
        sRelativePath = FilePath
        If Left(sRelativePath, 1) <> "\" Then
            sRelativePath = "\" & sRelativePath
        End If
        
        If Right(sRelativePath, 1) <> "\" Then
            sRelativePath = sRelativePath & "\"
        End If
        
        sFilePath = sRelativePath & FileName
    Else
        Dim aScanDir(1 To 2) As String
        Dim sDirPrefix As String
        Dim I As Integer
        
        aScanDir(1) = "State"
        aScanDir(2) = "Local"
        For I = 1 To 2
            aWords = Split(FileName, "_", 2)
            If LBound(aWords) = 0 And UBound(aWords) > 0 Then
                sPattern = aWords(0) & "_*"
            Else
                sPattern = "*"
            End If
    
            sDirPrefix = "\CollectionsManagement\AgencyFiles\" & aScanDir(I) & "\"
            
            f = Dir(sDrive & sDirPrefix & sPattern, vbDirectory)
            Do While Len(f) > 0
                cSearchDirs.Add f
                f = Dir()
            Loop
        
            For Each vSearchDir In cSearchDirs
                sPattern = sDrive & sDirPrefix & vSearchDir & "\ControlFiles\" & FileName
                If Dir(sPattern, vbNormal) <> "" Then
                    sFilePath = sDirPrefix & vSearchDir & "\ControlFiles\" & FileName
                    Exit For
                End If
            Next vSearchDir
        Next I
    End If
    
    If Len(sFilePath) > 0 Then
        sFilePath = sDrive & sFilePath
    End If
    getScanFileName = sFilePath
End Function

'**
'* getUserName: return a string to use in signature boxes, either the login name
'* of the current user (from $USERNAME) or an alias/initials specified in the
'* user_initials table
'*
'* @return String the signature to use for the current Windows user
'**
Function getUserName() As String
    Dim sUserName As String
    Dim Rs As Object
    
    'Potential signature candidates from $USERNAME
    sUserName = Environ("USERNAME")
        
    'check the database for an alias
    Set Rs = CurrentDb.OpenRecordset("SELECT * FROM user_initials WHERE username = '" & sUserName & "'")
    If Not (Rs.EOF) Then
        sUserName = Rs!initials
    End If

    getUserName = sUserName

End Function

'**
'* GetAccnsFromCurName: given a CurName indicating source, return a collection containing the ACCN numbers
'* of all the Accession sheets generated by that source.
'*
'* @param String CurName The Agency to collect accessions for (e.g. "ATG" for the Attorney General's Office)
'* @return Collection A collection of Strings containing ACCN numbers (e.g. {"2016.0001", "2000.0003", "1989.0123"})
'**
Function GetAccnsFromCurName(CurName As String) As Collection
    Dim cACCNs As New Collection
    
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef

    'check the database for comments flagged with this CollectionNumber
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qAccnsGet": On Error GoTo 0
    Set oQuery = CurrentDb.CreateQueryDef( _
        Name:="qAccnsGet", _
        SQLText:="SELECT * FROM Accessions " _
        & "WHERE Creator = [paramCreator]" _
    )
    oQuery.Parameters("paramCreator") = Trim(UCase(CurName))
        
    Set Rs = oQuery.OpenRecordset
    
    Do Until Rs.EOF
        cACCNs.Add Rs!ACCN.Value
        Rs.MoveNext
    Loop

    Set Rs = Nothing
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qAccnsGet": On Error GoTo 0
    
    Set GetAccnsFromCurName = cACCNs
End Function

Function CountAccnsFromCurName(CurName As String) As Integer
    
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef

    'check the database for comments flagged with this CollectionNumber
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qAccnsGet": On Error GoTo 0
    Set oQuery = CurrentDb.CreateQueryDef( _
        Name:="qAccnsGet", _
        SQLText:="SELECT COUNT(*) AS Count FROM Accessions " _
        & "WHERE Creator = [paramCreator]" _
    )
    oQuery.Parameters("paramCreator") = Trim(UCase(CurName))
        
    Set Rs = oQuery.OpenRecordset
    
    Do Until Rs.EOF
        CountAccnsFromCurName = Rs!Count.Value
        Rs.MoveNext
    Loop

    Set Rs = Nothing
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qAccnsGet": On Error GoTo 0
End Function

Function IsCurNameInCreators(CurName As String) As Boolean
    
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef

    'check the database for comments flagged with this CollectionNumber
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qAccnsGet": On Error GoTo 0
    Set oQuery = CurrentDb.CreateQueryDef( _
        Name:="qAccnsGet", _
        SQLText:="SELECT COUNT(*) AS Count FROM Creators " _
        & "WHERE CurName = [paramCreator]" _
    )
    oQuery.Parameters("paramCreator") = Trim(UCase(CurName))
        
    Set Rs = oQuery.OpenRecordset
    
    Do Until Rs.EOF
        IsCurNameInCreators = (Rs!Count.Value > 0)
        Rs.MoveNext
    Loop

    Set Rs = Nothing
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qAccnsGet": On Error GoTo 0
End Function

Function CurNamesInCreators(Optional ByVal CurName As String) As Dictionary
    Dim dCurNames As New Dictionary
    Dim sPattern As String
    
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef

    If Len(CurName) = 0 Then
        sPattern = "*"
    Else
        sPattern = CurName
    End If
    
    'check the database for comments flagged with this CollectionNumber
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qAccnsGet": On Error GoTo 0
    Set oQuery = CurrentDb.CreateQueryDef( _
        Name:="qAccnsGet", _
        SQLText:="SELECT CurName FROM Creators " _
        & "WHERE CurName LIKE [paramCreator]" _
    )
    oQuery.Parameters("paramCreator") = Trim(UCase(sPattern))
        
    Set Rs = oQuery.OpenRecordset
    
    Do Until Rs.EOF
        dCurNames.Item(Rs!CurName.Value) = True
        Rs.MoveNext
    Loop

    Set Rs = Nothing
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qAccnsGet": On Error GoTo 0

    Set CurNamesInCreators = dCurNames
End Function

'**
'* GetCommentsOnCollection: given a collection number, return an array containing the ID numbers of all
'* the comment/question records that reference that collection number, including all those that directly
'* reference it in the CollectionNumber field, and, optionally, also those that that mention it in comment
'* text.
'*
'* @param String CollectionNumber The CollectionNumber to count comments for (e.g. "GR-FIN-8")
'* @return Collection A collection of Integers containing comment ID numbers (e.g. {1, 17, 23})
'**
Function GetCommentsOnCollection(CollectionNumber As String) As Collection
    Dim aComments() As Integer, iComment As Integer
    Dim cComments As New Collection
    
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef
    Dim I As Integer, iCount As Integer

    'check the database for comments flagged with this CollectionNumber
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qCommentsCount": On Error GoTo 0
    Set oQuery = CurrentDb.CreateQueryDef( _
        Name:="qCommentsCount", _
        SQLText:="SELECT * FROM series_commentsquestions " _
        & "WHERE CollectionNumber = [paramCollectionNumber]" _
        & "OR CommentText LIKE ('*' & [paramCollectionNumber] & '*')" _
    )
    oQuery.Parameters("paramCollectionNumber") = Trim(UCase(CollectionNumber))
        
    Set Rs = oQuery.OpenRecordset
    
    Do Until Rs.EOF
        cComments.Add Rs!ID.Value
        Rs.MoveNext
    Loop

    Set Rs = Nothing
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qCommentsCount": On Error GoTo 0
    
    Set GetCommentsOnCollection = cComments
End Function

'**
'* NumberOfComments: given a collection number, return the number of comments/questions that have been
'* posted referencing that collection number, either directly in the CollectionNumber field, or by
'* mentioning it in comment text
'*
'* @uses GetCommentsOnCollection
'*
'* @param String CollectionNumber The CollectionNumber to count comments for (e.g. "GR-SOS-33")
'* @return Integer The number of comments referencing the given CollectionNumber
'**
Function NumberOfComments(CollectionNumber As String)
    Dim cComments As Collection
    
    Set cComments = GetCommentsOnCollection(CollectionNumber)
    NumberOfComments = cComments.Count
End Function

'**
'* DebugDump: Utility Function mainly for use in the Immediate pane to more easily display a
'* bunch of different kinds of objects and collections of objects in VBA
'*
'* @param Variant v The object to print out a representation of in the Immediate pane
'**
Sub DebugDump(v As Variant)
    Dim vScalar As Variant
    If IsArray(v) Or TypeName(v) = "Collection" Then
        If IsArray(v) Then
            Debug.Print TypeName(v), LBound(v), UBound(v)
        Else
            Debug.Print TypeName(v), v.Count
        End If
        
        For Each vScalar In v
            DebugDump (vScalar)
        Next vScalar
    ElseIf TypeName(v) = "Dictionary" Then
        Debug.Print TypeName(v)
        For Each vScalar In v.Keys
            Debug.Print vScalar & ":"
            DebugDump v.Item(vScalar)
        Next vScalar
    Else
        Debug.Print TypeName(v), v
    End If
End Sub

Function JoinCollection(delim As String, c As Collection)
    Dim first As Boolean
    Dim vItem As Variant
    Dim sConjunction As String
    
    first = True
    For Each vItem In c
        If Not first Then
            sConjunction = sConjunction & delim
        End If
        
        sConjunction = sConjunction & vItem
        
        first = False
    Next vItem
    
    JoinCollection = sConjunction
End Function

Function camelCaseSplitString(ByVal s As String) As Collection
    Dim isUpper As New RegExp
    Dim isLower As New RegExp
    Dim isWhiteSpace As New RegExp
    
    With isUpper
        .IgnoreCase = False
        .Pattern = "^([A-Z])$"
    End With
    
    With isLower
        .IgnoreCase = False
        .Pattern = "^([a-z])$"
    End With
    
    With isWhiteSpace
        .IgnoreCase = False
        .Pattern = "^(\s|[_])$"
    End With

    
    Dim cWords As New Collection
    Dim c As String
    Dim I As Integer
    Dim Anchor As Integer
    Dim State As Integer
    
    Anchor = 0
    I = 1
    GoTo NextWord
    
    'Finite State Machine
NextWord:
    If I > Len(s) Then
        GoTo ExitMachine
    End If
    
    c = Mid(s, I, 1)
    If isUpper.Test(c) Then
        Anchor = I
        GoTo FromUpperToNextWord
    ElseIf isLower.Test(c) Then
        Anchor = I
        GoTo FromLowerToNextWord
    ElseIf isWhiteSpace.Test(c) Then
        Anchor = I
        I = I + 1
        GoTo NextWord
    Else
        Anchor = I
        GoTo FromOtherToNextWord
    End If

FromLowerToNextWord:
    If I > Len(s) Then GoTo NextWord
    c = Mid(s, I, 1)
    
    If isLower.Test(c) Then
        I = I + 1
        GoTo FromLowerToNextWord
    ElseIf isUpper.Test(c) Or isWhiteSpace.Test(c) Then
        GoTo ClipWord
    Else
        I = I + 1
        GoTo FromLowerToNextWord
    End If
    
FromUpperToNextWord:
    If I > Len(s) Then GoTo NextWord
    c = Mid(s, I, 1)

    If isUpper.Test(c) Then
        I = I + 1
        GoTo FromUpperToNextWord
    ElseIf isLower.Test(c) Then
        I = I + 1
        GoTo FromLowerToNextWord
    ElseIf isWhiteSpace.Test(c) Then
        GoTo ClipWord
    Else
        I = I + 1
        GoTo FromUpperToNextWord
    End If
    
FromOtherToNextWord:
    If I > Len(s) Then GoTo NextWord
    c = Mid(s, I, 1)
    
    If isUpper.Test(c) Or isLower.Test(c) Or isWhiteSpace.Test(c) Then
        GoTo ClipWord
    Else
        I = I + 1
        GoTo FromOtherToNextWord
    End If

ClipWord:
    If Anchor < I Then
        cWords.Add Mid(s, Anchor, I - Anchor)
    End If
    GoTo NextWord

ExitMachine:
    If Anchor > 0 Then
        cWords.Add Mid(s, Anchor, I - Anchor)
    End If
    
'    Anchor = 1
'
'    For I = 1 To Len(s)
'        c = Mid(s, I, 1)
'        If oIsUpperAlpha.Test(c) Then
'            If Anchor < I Then
'                cWords.Add Mid(s, Anchor, I - Anchor)
'            End If
'            Anchor = I
'        Else
'
'        End If
'    Next I
    
'    cWords.Add Mid(s, Anchor, Len(s) - Anchor + 1)
    
    Set camelCaseSplitString = cWords
End Function

Function getDefaultDrive() As String
    getDefaultDrive = "\\ADAHFS1\GR-Collections"
End Function
