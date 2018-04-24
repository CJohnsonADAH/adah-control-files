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
        Let sAgencyName = Nz(rsAccessions!AgencyName.value)
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
    Dim i As Integer, iCount As Integer

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
        cACCNs.Add Rs!ACCN.value
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
        CountAccnsFromCurName = Rs!Count.value
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
        IsCurNameInCreators = (Rs!Count.value > 0)
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
        dCurNames.Item(Rs!CurName.value) = True
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
    Dim i As Integer, iCount As Integer

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
        cComments.Add Rs!ID.value
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
    If IsArray(v) Or TypeName(v) = "Collection" Or TypeName(v) = "ISubMatches" Then
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
    Dim isAlpha As New RegExp
    Dim isUpper As New RegExp
    Dim isLower As New RegExp
    Dim isUpperLower As New RegExp
    Dim isWhiteSpace As New RegExp
    
    With isUpper
        .IgnoreCase = False
        .pattern = "^([A-Z])$"
    End With
    
    With isLower
        .IgnoreCase = False
        .pattern = "^([a-z])$"
    End With
    
    With isAlpha
        .IgnoreCase = False
        .pattern = "^([A-Za-z]+)$"
    End With

    With isUpperLower
        .IgnoreCase = False
        .pattern = "^([A-Z][a-z])$"
    End With
    
    With isWhiteSpace
        .IgnoreCase = False
        .pattern = "^((\s|[_])+)$"
    End With

    
    Dim cWords As New Collection
    Dim c0 As String, c As String, c2 As String
    Dim I0 As Integer, i As Integer
    Dim Anchor As Integer
    Dim State As Integer
    
    Anchor = 0
    i = 1
    GoTo NextWord
    
    'Finite State Machine
NextWord:
    If i > Len(s) Then
        GoTo ExitMachine
    End If
    
    c0 = Mid(s, i, 1)
    If isUpper.Test(c0) Then
        Anchor = i
        GoTo WordBeginsOnUpper
    ElseIf isLower.Test(c0) Then
        Anchor = i
        GoTo WordBeginsOnLower
    ElseIf isWhiteSpace.Test(c0) Then
        Anchor = i
        Let i = i + Len(c)
        GoTo NextWord
    Else
        Anchor = i
        GoTo FromOtherToNextWord
    End If

WordBeginsOnLower:
    If i > Len(s) Then GoTo NextWord
    Let c0 = c: Let c = Mid(s, i, 1)
    
    Let i = i + Len(c)
    GoTo ContinueWordToUpperBreak

ContinueWordToUpperBreak:
    If i > Len(s) Then GoTo NextWord
    Let c0 = c: Let c = Mid(s, i, 1)
    
    If isUpper.Test(c) Or isWhiteSpace.Test(c) Then
        GoTo ClipWord
    Else
        i = i + Len(c)
    End If
    GoTo ContinueWordToUpperBreak
    
WordBeginsOnUpper:
    If i > Len(s) Then GoTo NextWord
    Let c0 = c: c = Mid(s, i, 1)

    'Move ahead to the next character
    'UPPERCase: two uppers in a row
    'MixedCase: one upper, one lower
    Let i = i + 1
    Let c0 = c: c = Mid(s, i, 1)
    
    If isLower.Test(c) Then
        Let i = i + Len(c)
        GoTo ContinueWordToUpperBreak
    ElseIf isUpper.Test(c) Then
        GoTo ContinueWordToUpperLowerBreak
    ElseIf Not isWhiteSpace.Test(c) Then
        GoTo ContinueWordToUpperLowerBreak
    Else
        GoTo ClipWord
    End If
    
ContinueWordToUpperLowerBreak:
    If i > Len(s) Then GoTo NextWord
    Let c0 = c: c = Mid(s, i, 1): c2 = Mid(s, i, 2)

    If isUpperLower.Test(c2) Or isWhiteSpace.Test(c) Then
        GoTo ClipWord
    ElseIf isAlpha.Test(c) Then
        i = i + 1
        GoTo ContinueWordToUpperLowerBreak
    Else
        i = i + 1
        GoTo ContinueWordToUpperLowerBreak
    End If
    
FromOtherToNextWord:
    If i > Len(s) Then GoTo NextWord
    c = Mid(s, i, 1)
    
    If isUpper.Test(c) Or isLower.Test(c) Or isWhiteSpace.Test(c) Then
        GoTo ClipWord
    Else
        i = i + 1
        GoTo FromOtherToNextWord
    End If

ClipWord:
    If Anchor < i Then
        cWords.Add Mid(s, Anchor, i - Anchor)
    End If
    GoTo NextWord

ExitMachine:
    If Anchor > 0 Then
        cWords.Add Mid(s, Anchor, i - Anchor)
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
                Debug.Print Rs!FileName.value, " => ", oAccnScan.FileName
                Rs.Edit
                Let Rs!FileName.value = oAccnScan.FileName
                Rs.Update
            End If
        End If
        
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
    
End Function

Public Function ConvertPDFs()
    Dim oAccnScan As New cAccnScan
    
    Dim sRelativePath As String
    Dim f As String
    Dim sPattern As String
    Dim aWords() As String
    Dim cSearchDirs As New Collection
    Dim vSearchDir As Variant
    
    Dim sFileName As String
    Dim sNewFileName As String
    Dim sFullPath As String
    Dim sNewFullPath As String
    
    Dim cBits As Collection
    
    Dim vScanDir As Variant
    Dim cScanDirs As New Collection
    Dim sDirPrefix As String
    Dim i As Integer
    
    Dim Rs As DAO.Recordset
    
    cScanDirs.Add "\AgencyState"
    cScanDirs.Add "\AgencyLocal"
    cScanDirs.Add "\AgencyCourts"
    cScanDirs.Add "\AgencyUS"
    cScanDirs.Add "\CollectionsManagement\AgencyFiles\State"
    cScanDirs.Add "\CollectionsManagement\AgencyFiles\Local"
    
    For Each vScanDir In cScanDirs
        sPattern = "*"
    
        sDirPrefix = vScanDir & "\"
        f = Dir(oAccnScan.Drive & sDirPrefix & sPattern, vbDirectory)
        Do While Len(f) > 0
            If (f <> ".") And (f <> "..") Then
                cSearchDirs.Add (sDirPrefix & f & "\ContolFile")
                cSearchDirs.Add (sDirPrefix & f & "\ContolFiles")
                cSearchDirs.Add (sDirPrefix & f & "\ControlFile")
                cSearchDirs.Add (sDirPrefix & f & "\ControlFiles")
            End If
            f = Dir()
        Loop
    Next vScanDir

    Dim sMatchDocumentation As String
    Dim sMatchDocumentationV2 As String
    Dim sMatchCopierScan As String
    
    Let sMatchDocumentation = "^([A-Z0-9]{2,3})_(Correspondence|Documentation|Administrative|Administration|Clipping|Microfilm)_(.*)([.]PDF)$"
    Let sMatchDocumentationV2 = "^([A-Z0-9]{2,3})(Corr|Doc|Admin|Clip|Microfilm)(.*)([.]PDF)$"
    
    Let sMatchCopierScan = oAccnScan.MATCH_COPIER_SCAN
    
    For Each vSearchDir In cSearchDirs
        sPattern = oAccnScan.Drive & vSearchDir & "\*.PDF"
        Let sFileName = Dir(sPattern, vbNormal)
        Do While Len(sFileName) > 0
            If RegexMatch(sFileName, oAccnScan.MATCH_FILENAME_V1) Then
                ' NOOP
            ElseIf RegexMatch(sFileName, oAccnScan.MATCH_FILENAME_V2) Then
                ' NOOP
            ElseIf RegexMatch(sFileName, sMatchDocumentationV2) Then
                ' NOOP
            ElseIf RegexMatch(sFileName, sMatchDocumentation) Then
                Set cBits = New Collection
                cBits.Add RegexComponent(sFileName, sMatchDocumentation, 1)
                cBits.Add Abbreviate(RegexComponent(sFileName, sMatchDocumentation, 2))
                cBits.Add RegexComponent(sFileName, sMatchDocumentation, 3)
                cBits.Add RegexComponent(sFileName, sMatchDocumentation, 4)
                
                ' Is it referenced in AccnScans?
                Set Rs = CurrentDb.OpenRecordset("SELECT * FROM AccnScans WHERE FileName='" & Replace(sFileName, "'", "''") & "'")
                If Rs.EOF Then
                    Debug.Print RegexComponent(sFileName, sMatchDocumentation, 2) & " [NO DB]: ", sFileName, JoinCollection("", cBits)
                    Name (oAccnScan.Drive & vSearchDir & "\" & sFileName) As (oAccnScan.Drive & vSearchDir & "\" & JoinCollection("", cBits))
                Else
                    Debug.Print RegexComponent(sFileName, sMatchDocumentation, 2) & " [DB]: ", sFileName, JoinCollection("", cBits)
                    Name (oAccnScan.Drive & vSearchDir & "\" & sFileName) As (oAccnScan.Drive & vSearchDir & "\" & JoinCollection("", cBits))
                    Rs.Edit
                    Let Rs!FileName = JoinCollection("", cBits)
                    Rs.Update
                End If
                Rs.Close
                
                Set cBits = Nothing
            ElseIf RegexMatch(sFileName, sMatchCopierScan) Then
                
                ' Is it referenced in AccnScans?
                Set Rs = CurrentDb.OpenRecordset("SELECT * FROM AccnScans LEFT JOIN Accessions ON (AccnScans.ACCN=Accessions.ACCN) WHERE FileName='" & Replace(sFileName, "'", "''") & "'")
                If Rs.EOF Then
                    Debug.Print "COPIER SCAN UNPROCESSED [NO DB]: ", sFileName
                Else
                    Let sFullPath = oAccnScan.Drive & vSearchDir & "\" & sFileName
                    If Rs!FileNameToBeFixed.value Then
                        If Len(Nz(Rs.Fields("AccnScans.ACCN").value)) > 0 Then
                            
                            Let sNewFileName = GetCurNameFromCreatorCode(Nz(Rs!Creator)) & Replace(Nz(Rs.Fields("AccnScans.ACCN").value), ".", "") & ".PDF"
                        Else
                            Let sNewFileName = sFileName
                        End If
                        
                        Let sNewFullPath = oAccnScan.Drive & vSearchDir & "\" & sNewFileName
                    
                        Debug.Print "COPIER SCAN [DB]: ", sFullPath, "=>", sNewFileName
                        Name sFullPath As sNewFullPath
                        Rs.Edit
                        Let Rs!FileName = sNewFileName
                        Let Rs!FileNameToBeFixed = False
                        Rs.Update
                    Else
                        Debug.Print "COPIER SCAN UNPROCESSED [DB]: ", sFileName
                    End If
                End If
                Rs.Close
                
            Else
                Debug.Print "MISC UNPROCESSED: ", sFileName
            End If
            Let sFileName = Dir
        Loop
    Next vSearchDir
    Debug.Print "... Done."
    
End Function

Public Function Abbreviate(Text As String)
    Let Abbreviate = Text
    
    Select Case Text
    Case "Correspondence":
        Let Abbreviate = "Corr"
    Case "Documentation":
        Let Abbreviate = "Doc"
    Case "Administrative":
        Let Abbreviate = "Admin"
    Case "Administration":
        Let Abbreviate = "Admin"
    Case "Clipping":
        Let Abbreviate = "Clip"
    End Select
End Function

Public Function GetCurNameFromCreatorCode(Code As String) As String
    Let GetCurNameFromCreatorCode = Left(Code, InStr(1, Code & "-", "-") - 1)
End Function
