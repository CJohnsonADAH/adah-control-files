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

Public gPipes As cPipeNetwork

'**
'* SecureCreatorRecord
'*
'* @uses DAO.Recordset
'*
'* @param String Creator A two- or three-character alphanumeric CurName code
'**
Sub SecureCreatorRecord(Creator As String)
    Dim sSuggestion As String
    
    Dim sCurName As String
    Dim sAgencyName As String
    
    Dim sCreatorCode As String
    Dim sDivision As String
    Dim sSection As String
    
    Dim rsAccessions As DAO.Recordset
    
    Let sCurName = Creator
    If InStr(1, Creator, "-") Then
        Let sCurName = Left(sCurName, InStr(1, sCurName, "-") - 1)
    End If
    
    Set rsAccessions = CurrentDb.OpenRecordset("SELECT * FROM Sheet1 WHERE CurName = '" & sCurName & "'")
    If Not rsAccessions.EOF Then
        If Not IsNull(rsAccessions!AgencyName) Then
            sSuggestion = rsAccessions!AgencyName
        Else
            sSuggestion = Creator
        End If
    Else
        sSuggestion = Creator
    End If
    rsAccessions.Close
    
    
    Set rsAccessions = CurrentDb.OpenRecordset("SELECT * FROM Agencies WHERE CurName = '" & sCurName & "'")
    If rsAccessions.EOF Then
        sAgencyName = InputBox(Prompt:=Creator & " = ", Title:="AgencyName for CurName " & sCurName, Default:=sSuggestion)
        If Len(sAgencyName) > 0 Then
            rsAccessions.Close
            
            Set rsAccessions = CurrentDb.OpenRecordset("Agencies")
            rsAccessions.AddNew
            rsAccessions!CurName = sCurName
            rsAccessions!AgencyName = sAgencyName
            rsAccessions.Update
        End If
    Else
        Let sAgencyName = Nz(rsAccessions!AgencyName.Value)
    End If
    rsAccessions.Close
    
    Set rsAccessions = CurrentDb.OpenRecordset("SELECT * FROM Creators WHERE (CreatorCode = '" & Creator & "') OR (CreatorCode LIKE '" & Creator & "-*')")
    If rsAccessions.EOF Then
        sCreatorCode = InputBox(Prompt:=Creator & " = ", Title:="CreatorCode for " & Creator, Default:=Creator & "-00")
        sDivision = InputBox(Prompt:=sCreatorCode & ", Division: ", Title:="Division for " & sCreatorCode, Default:="")
        sSection = InputBox(Prompt:=sCreatorCode & ", Division: " & sDivision & " Section: ", Title:="Section for " & sCreatorCode, Default:="")
        If Len(sCreatorCode) > 0 Then
            rsAccessions.Close
            
            Set rsAccessions = CurrentDb.OpenRecordset("Creators")
            rsAccessions.AddNew
            rsAccessions!CreatorCode = sCreatorCode
            rsAccessions!AgencyName = sAgencyName
            If Len(sDivision) > 0 Then
                rsAccessions!Division = sDivision
            End If
            If Len(sSection) > 0 Then
                rsAccessions!Section = sSection
            End If
            rsAccessions.Update
        End If
    End If
    rsAccessions.Close
    
End Sub

Sub deleteAccnScanAttachments(ACCN As String, ByVal FileName As String, ByVal FilePath As String)
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef
    Dim sFilePathMatch As String
    
    'FIXME: This currently seems to be broken, and returns no records from the WHERE
    sFilePathMatch = "FilePath=[paramFilePath]"
    If Len(FilePath) = 0 Then
        sFilePathMatch = "(LEN(FilePath)=0 OR FilePath IS NULL)"
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
    If Len(FilePath) > 0 Then
        oQuery.Parameters("paramFilePath") = FilePath
    End If
    oQuery.Execute
    
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qDeleteAccnScans": On Error GoTo 0

    Set oQuery = Nothing

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
Function getScanFileName(ByVal FileName As String, Optional ByVal FilePath As String, Optional ByVal ExcludeDrive As Boolean, Optional ByVal ExcludeFileName As Boolean) As String
    Dim oAccnScan As New cAccnScan
    Dim sFullPath As String
    
    Let oAccnScan.FileName = FileName
    Let oAccnScan.FilePath = FilePath
    
    If oAccnScan.Exists Then
        If Not ExcludeDrive Then
            Let sFullPath = oAccnScan.Drive
        End If
        
        Let sFullPath = sFullPath & oAccnScan.FilePath
        
        If Not ExcludeFileName Then
            Let sFullPath = sFullPath & "\" & oAccnScan.FileName
        End If
    End If
    
    Let getScanFileName = sFullPath
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

Public Function camelCaseSplitString(ByVal s As String) As Collection
    Dim isAlpha As New RegExp
    Dim isUpper As New RegExp
    Dim isLower As New RegExp
    Dim isUpperLower As New RegExp
    Dim isWhiteSpace As New RegExp
    
    With isUpper
        .IgnoreCase = False
        .Pattern = "^([A-Z])$"
    End With
    
    With isLower
        .IgnoreCase = False
        .Pattern = "^([a-z])$"
    End With
    
    With isAlpha
        .IgnoreCase = False
        .Pattern = "^([A-Za-z]+)$"
    End With

    With isUpperLower
        .IgnoreCase = False
        .Pattern = "^([A-Z][a-z])$"
    End With
    
    With isWhiteSpace
        .IgnoreCase = False
        .Pattern = "^((\s|[_])+)$"
    End With

    
    Dim cWords As New Collection
    Dim c0 As String, c As String, c2 As String
    Dim I0 As Integer, I As Integer
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
    
    c0 = Mid(s, I, 1)
    If isUpper.Test(c0) Then
        Anchor = I
        GoTo WordBeginsOnUpper
    ElseIf isLower.Test(c0) Then
        Anchor = I
        GoTo WordBeginsOnLower
    ElseIf isWhiteSpace.Test(c0) Then
        Anchor = I
        Let I = I + Len(c)
        GoTo NextWord
    Else
        Anchor = I
        GoTo FromOtherToNextWord
    End If

WordBeginsOnLower:
    If I > Len(s) Then GoTo NextWord
    Let c0 = c: Let c = Mid(s, I, 1)
    
    Let I = I + Len(c)
    GoTo ContinueWordToUpperBreak

ContinueWordToUpperBreak:
    If I > Len(s) Then GoTo NextWord
    Let c0 = c: Let c = Mid(s, I, 1)
    
    If isUpper.Test(c) Or isWhiteSpace.Test(c) Then
        GoTo ClipWord
    Else
        I = I + Len(c)
    End If
    GoTo ContinueWordToUpperBreak
    
WordBeginsOnUpper:
    If I > Len(s) Then GoTo NextWord
    Let c0 = c: c = Mid(s, I, 1)

    'Move ahead to the next character
    'UPPERCase: two uppers in a row
    'MixedCase: one upper, one lower
    Let I = I + 1
    Let c0 = c: c = Mid(s, I, 1)
    
    If isLower.Test(c) Then
        Let I = I + Len(c)
        GoTo ContinueWordToUpperBreak
    ElseIf isUpper.Test(c) Then
        GoTo ContinueWordToUpperLowerBreak
    ElseIf Not isWhiteSpace.Test(c) Then
        GoTo ContinueWordToUpperLowerBreak
    Else
        GoTo ClipWord
    End If
    
ContinueWordToUpperLowerBreak:
    If I > Len(s) Then GoTo NextWord
    Let c0 = c: c = Mid(s, I, 1): c2 = Mid(s, I, 2)

    If isUpperLower.Test(c2) Or isWhiteSpace.Test(c) Then
        GoTo ClipWord
    ElseIf isAlpha.Test(c) Then
        I = I + 1
        GoTo ContinueWordToUpperLowerBreak
    Else
        I = I + 1
        GoTo ContinueWordToUpperLowerBreak
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

Public Function getDefaultDrive() As String
    getDefaultDrive = "\\ADAHFS1\GR-Collections"
End Function

Public Function IsSameLNUMBER(L1 As String, L2 As String) As Boolean
    Dim o1 As New cLNumber
    Dim o2 As New cLNumber
    
    Let o1.Number = L1
    Let o2.Number = L2
    
    Let IsSameLNUMBER = o1.SameAs(o2)
End Function

Public Function ConvertAccnScanFileNames()
    Dim Rs As DAO.Recordset
    Dim oAccnScan As cAccnScan
    Dim sNewFileName As String
    
    Set Rs = CurrentDb.OpenRecordset("AccnScans")
    Do Until Rs.EOF
        
        If Len(Nz(Rs!FilePath)) > 0 Then
            Set oAccnScan = New cAccnScan: With oAccnScan
                .FilePath = Rs!FilePath
                .FileName = Rs!FileName
            End With
            
            oAccnScan.ConvertFileName Result:=sNewFileName
            If Len(sNewFileName) > 0 Then
                Debug.Print Rs!FileName.Value, " => ", oAccnScan.FileName
                Rs.Edit
                Let Rs!FileName.Value = oAccnScan.FileName
                Rs.Update
            End If
        End If
        
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
    
End Function

Public Function RenamePDFs(Optional ByVal LogLevel As Integer)
    Dim oAccnScan As cAccnScan
    
    Dim sRelativePath As String
    Dim f As String
    Dim sPattern As String
    Dim aWords() As String
    
    Dim sCreator As String
    Dim sCurName As String
    
    Dim sDrive As String
    
    Dim sFileName As String
    Dim sNewFileName As String
    Dim sFullPath As String
    Dim sNewFullPath As String
    
    Dim cBits As Collection
    
    Dim cSourcePaths As Collection, vSourcePath As Variant
    Dim cSearchDirs As Collection, vSearchDir As Variant
    Dim cScanFiles As Collection, vScanFile As Variant
    
    Dim sDirPrefix As String
    Dim I As Integer
    
    Dim Rs As DAO.Recordset
        
    Set oAccnScan = New cAccnScan
    
    Set cSourcePaths = New Collection: With cSourcePaths
        .Add "\AgencyState"
        .Add "\AgencyLocal"
        .Add "\AgencyCourts"
        .Add "\AgencyUS"
        .Add "\CollectionsManagement\AgencyFiles\State"
        .Add "\CollectionsManagement\AgencyFiles\Local"
    End With
    Let sDrive = oAccnScan.Drive
    
    Set cSearchDirs = New Collection
    
    For Each vSourcePath In cSourcePaths
        sPattern = "*"
    
        sDirPrefix = CStr(vSourcePath) & "\"
        f = Dir(sDrive & sDirPrefix & sPattern, vbDirectory)
        Do While Len(f) > 0
            If (f <> ".") And (f <> "..") Then
                With cSearchDirs
                    .Add (sDirPrefix & f & "\ContolFile")
                    .Add (sDirPrefix & f & "\ContolFiles")
                    .Add (sDirPrefix & f & "\ControlFile")
                    .Add (sDirPrefix & f & "\ControlFiles")
                End With
            End If
            f = Dir()
        Loop
    Next vSourcePath

    For Each vSearchDir In cSearchDirs
        sPattern = sDrive & vSearchDir & "\*.PDF"
        
        Set cScanFiles = New Collection
        Let sFileName = Dir(sPattern, vbNormal)
        Do While Len(sFileName) > 0
            cScanFiles.Add sFileName
            Let sFileName = Dir
        Loop
        
        For Each vScanFile In cScanFiles
            DoEvents
            
            Let sFileName = CStr(vScanFile)
            
            Set oAccnScan = New cAccnScan: With oAccnScan
                .Url = sDrive & vSearchDir & "\" & sFileName
            End With
            
            ' Is this scan file referenced in AccnScans?
            Set Rs = CurrentDb.OpenRecordset("SELECT * FROM AccnScans LEFT JOIN Accessions ON (AccnScans.ACCN=Accessions.ACCN) WHERE FileName='" & Replace(sFileName, "'", "''") & "'")
                
            ' Yes: Get the meta-data from the database, then use that to enforce the naming convertion as need be
            If Not Rs.EOF Then
                Let sFullPath = sDrive & vSearchDir & "\" & sFileName
                
                ' Is this scan file flagged as needing a fix for the naming convention?
                If Rs!FileNameToBeFixed.Value Then
                    Let oAccnScan.SheetType = Nz(Rs!SheetType.Value)
                    If Nz(Rs!Timestamp.Value) <> 0 Then
                        Let oAccnScan.Timestamp = Nz(Rs!Timestamp.Value)
                    End If
                    If Len(Nz(Rs.Fields("AccnScans.ACCN").Value)) > 0 Then
                        Let oAccnScan.ACCN = Nz(Rs.Fields("AccnScans.ACCN").Value)
                    End If
                        
                    Let sCreator = Nz(Rs!Creator.Value)
                    If Len(sCreator) = 0 Then
                        Let sCreator = oAccnScan.Creator
                    End If
                    Let sCurName = GetCurNameFromCreatorCode(sCreator)
                    
                    oAccnScan.ConvertFileName Result:=sNewFileName
                    If Len(sNewFileName) > 0 Then
                        If LogLevel = 0 Or LogLevel > 0 Then
                            Debug.Print "COPIER SCAN [DB]: ", sFullPath, "=>", sNewFileName
                        End If
                    End If
                Else
                    If LogLevel = 0 Or LogLevel > 1 Then
                        Debug.Print oAccnScan.SheetType, "UNPROCESSED [NO FIX CHECK]: ", sFileName
                    End If
                End If
                
            'No, but we can tell it's an ACCN sheet from the naming convention
            ElseIf oAccnScan.IsAccnSheet Then
                'NOOP
                If LogLevel = 0 Or LogLevel > 1 Then
                    Debug.Print "ACCN SHEET UNPROCESSED [NO DB]: ", sFileName
                End If
                    
            'No, but we can tell it's a documentation sheet from the V1 naming convention
            ElseIf oAccnScan.IsDocumentationSheet(Version:=1) Then
                Set cBits = New Collection
                cBits.Add RegexComponent(sFileName, oAccnScan.MATCH_DOCUMENTATION_V1, 1)
                cBits.Add oAccnScan.SheetTypeSlug
                cBits.Add RegexComponent(sFileName, oAccnScan.MATCH_DOCUMENTATION_V1, 3)
                cBits.Add RegexComponent(sFileName, oAccnScan.MATCH_DOCUMENTATION_V1, 4)
                
                ' Is it referenced in AccnScans?
                Set Rs = CurrentDb.OpenRecordset("SELECT * FROM AccnScans WHERE FileName='" & Replace(sFileName, "'", "''") & "'")
                If Rs.EOF Then
                    Debug.Print RegexComponent(sFileName, oAccnScan.MATCH_DOCUMENTATION_V1, 2) & " [NO DB]: ", sFileName, JoinCollection("", cBits)
                    Name (sDrive & vSearchDir & "\" & sFileName) As (sDrive & vSearchDir & "\" & JoinCollection("", cBits))
                Else
                    Debug.Print RegexComponent(sFileName, oAccnScan.MATCH_DOCUMENTATION_V1, 2) & " [DB]: ", sFileName, JoinCollection("", cBits)
                    Name (sDrive & vSearchDir & "\" & sFileName) As (sDrive & vSearchDir & "\" & JoinCollection("", cBits))
                    Rs.Edit
                    Let Rs!FileName = JoinCollection("", cBits)
                    Rs.Update
                End If
                Rs.Close
                
                Set cBits = Nothing
            
            'No, but we can tell it's a documentation sheet from the V2 naming convention
            ElseIf oAccnScan.IsDocumentationSheet(Version:=2) Then
                'NOOP
            
            'No, and we can't really tell what it is from the naming conventions
            Else
                If LogLevel = 0 Or LogLevel > 2 Then
                    Debug.Print "MISC UNPROCESSED: ", sFileName
                End If
            End If

        Next vScanFile
    Next vSearchDir
    
    If LogLevel = 0 Or LogLevel > 0 Then
        Debug.Print "... Done."
    End If
    Exit Function
    
CatchNameAsFailure:
    Dim Fixed As Boolean
    Dim Ignore As Boolean
    
    Dim Disambiguator_Pattern As String
    Dim Disambiguator As String
    
    If Err.Number = EX_FILEALREADYEXISTS Then
        Let Disambiguator_Pattern = "(-([0-9]+))?([.][^.]+)$"
        Let Disambiguator = RegexComponent(sNewFileName, Disambiguator_Pattern, 2)
        If Len(Disambiguator) > 0 Then
            Let Disambiguator = Format(Val(Disambiguator) + 1, "0")
        Else
            Let Disambiguator = "2"
        End If
        
        Let sNewFileName = RegexReplace(sNewFileName, Disambiguator_Pattern, "-" & Disambiguator & "$3")
        Let sNewFullPath = sDrive & vSearchDir & "\" & sNewFileName
        Let Fixed = True
    ElseIf Err.Number = EX_FILEPERMISSIONDENIED Then
        Debug.Print "!!!", "Could not rename [" & sFullPath & "] to [" & sNewFullPath & "]: permission denied"
    End If
    
    If Fixed Then
        Resume
    ElseIf Ignore Then
        Resume Next
    Else
        Err.Raise EX_RENAMEFAILED, "ModControlFilesDatabase::ConvertPDFs", "Failed to rename [" & sFileName & "] to [" & sNewFileName & "]"
    End If
    
End Function

Public Function GetCurNameFromCreatorCode(Code As String) As String
    Let GetCurNameFromCreatorCode = Left(Code, InStr(1, Code & "-", "-") - 1)
End Function

Public Function GetDateSlug(Timestamp As Variant, Optional ByVal OmitTime As Boolean, Optional ByVal Default As Variant) As String
    Dim sDefault As String
    Dim sDate As String
    Dim sTime As String
    
    If IsMissing(Default) Then
        Let sDefault = "ND"
    Else
        Let sDefault = CStr(Default)
    End If
    
    If IsNull(Timestamp) Or Len(Nz(Timestamp)) = 0 Or (IsDate(Timestamp) And Timestamp = 0) Then
        Let GetDateSlug = sDefault
    Else
        Let sDate = Format(Timestamp, "YYYYmmdd")
        If OmitTime Or Format(Timestamp, "HMS") <> "000" Then
            Let sDate = sDate & "_" & Format(Timestamp, "HHMM")
        End If
        Let GetDateSlug = sDate
    End If
End Function

Public Sub GetCabinetFoldersFromAccessions()
    Dim Rs As DAO.Recordset
    Dim CabinetFolder As Variant
    Dim nCabinetFolder As Long
    
    Set Rs = CurrentDb.OpenRecordset("AccnScans")
    Do Until Rs.EOF
        If IsNull(Rs!CabinetFolder.Value) And Len(Nz(Rs!ACCN.Value)) > 0 Then
            CabinetFolder = DLookup("CabinetFolder", "Accessions", "ACCN='" & Nz(Rs!ACCN.Value) & "'")
            If Len(Nz(CabinetFolder)) > 0 Then
                nCabinetFolder = Nz(DLookup("ID", "CabinetFolders", "Label='" & Replace(CabinetFolder, "'", "''") & "'"))
                Debug.Print Rs!ID.Value, " in folder ", CabinetFolder, nCabinetFolder
                If IsNull(Rs!CabinetFolder.Value) And nCabinetFolder > 0 Then
                    Rs.Edit
                    Rs!CabinetFolder.Value = nCabinetFolder
                    Rs.Update
                End If
            End If
        End If
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
End Sub

Public Sub InitializePipesAndFilters()
    If gPipes Is Nothing Then
        Set gPipes = New cPipeNetwork
    End If
End Sub

Public Sub ReInitializePipesAndFilters()
    If Not gPipes Is Nothing Then
        gPipes.Reset
    End If
    Set gPipes = Nothing
    
    Set gPipes = New cPipeNetwork
End Sub

Public Sub FileRename(ByVal Source As String, ByVal Destination As String)
    Dim FromTo As New Dictionary: With FromTo
        .Add Key:="Source", Item:=Source
        .Add Key:="Destination", Item:=Destination
    End With
    
    DoAction Outlet:="FileToBeRenamed", Parameters:=FromTo
    Name Source As Destination
    DoAction Outlet:="FileHasBeenRenamed", Parameters:=FromTo
End Sub

Public Sub AddAction(ByVal Outlet As String, Plug As IReceiver, Optional ByVal Priority As Integer)
    InitializePipesAndFilters
    gPipes.AddAction Outlet:=Outlet, Plug:=Plug, Priority:=Priority
End Sub

Public Sub DoAction(ByVal Outlet As String, Optional Parameters As Variant)
    InitializePipesAndFilters
    gPipes.DoAction Outlet:=Outlet, Parameters:=Parameters
End Sub

Public Function ApplyFilters(ByVal Outlet As String, InputElement As Variant, Optional Parameters As Variant) As Variant
    InitializePipesAndFilters
    Let ApplyFilters = gPipes.ApplyFilters(Outlet:=Outlet, InputElement:=InputElement, Parameters:=Parameters)
End Function
