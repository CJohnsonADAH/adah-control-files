VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAccnScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private sPathStart As String
Private sDrive As String
Private sFilePath As String
Private sFileName As String
Private sOldPath As String
Private sFileNameCorrected As String
Private sSheetType As String
Private bSheetTypeSet As Boolean
Private iCabinetFolder As Long
Private oFolder As cCabinetFolder
Private bCabinetFolderSet As Boolean
Private sACCN As String

Private dSheetTypeSlugs As Dictionary

'**
'* Get cAccnScan::FilePath: get the filesystem path of an AccnScan file
'*
'* @return String A file system path to this Accession sheet scan, either from the database record or
'*      from a quick search of the file system. Does not include drive letter or network authority.
'*      For example, \CollectionsManagement\AgencyFiles\Local\63_TuscaloosaCounty\ControlFiles
'**
Public Property Get FilePath() As String
    If Len(sFilePath) = 0 Then
        FilePath = retrieveScanFilePath
    Else
        FilePath = sFilePath
    End If
End Property

'**
'* Let cAccnScan::FilePath: set a value for the filesystem path of an AccnScan file
'*
'* @param String A file system path to this Accession sheet scan. Does not include drive letter or
'*      network authority.
'*      For example, \CollectionsManagement\AgencyFiles\Local\63_TuscaloosaCounty\ControlFiles
'**
Public Property Let FilePath(sFP As String)
    sFilePath = sFP
End Property

Public Property Get FileName() As String
    FileName = sFileName
End Property

Public Property Let FileName(sFN As String)
    sACCN = ""
    sFileName = sFN
End Property

Public Property Let PathStart(sPS As String)
    sPathStart = sPS
End Property

Public Property Get PathStart() As String
    If Len(sPathStart) = 0 Then
        sPathStart = "CollectionsManagement"
    End If
    
    PathStart = sPathStart
End Property

Public Property Get OldPath() As String
    Let OldPath = sOldPath
End Property

Public Property Let OldPath(s As String)
    Let sOldPath = s
End Property

Public Property Get CabinetFolder() As Long
    If bCabinetFolderSet Then
        Let CabinetFolder = iCabinetFolder
    ElseIf Len(FilePath) > 0 Then
        Let CabinetFolder = getCabinetFolderID
    End If
End Property

Public Property Let CabinetFolder(l As Long)
    Let iCabinetFolder = l
    Let bCabinetFolderSet = True
End Property

Private Function Folder() As cCabinetFolder
    If oFolder Is Nothing Then
        ResetFolderWithSlug
    End If
    Set Folder = oFolder
End Function

Private Sub ResetFolderWithSlug()
    Set oFolder = New cCabinetFolder: With oFolder
        .FilePath = FileFolderSlug
    End With
End Sub

Private Function getCabinetFolderID() As Integer
    ResetFolderWithSlug
    Let getCabinetFolderID = oFolder.ID
End Function

Public Function FileFolder() As String
    Dim ID As Long
    
    Let ID = CabinetFolder
    If ID = 0 Then Folder.Insert
    
    If ID <> 0 Then Let FileFolder = Folder.Label
End Function

Public Function FileFolderSlug(Optional ByVal Old As Boolean) As String
    Dim myPath As String
    Dim i As Integer
    Dim sSlug As String
    Dim asFolders() As String
    
    Let myPath = FilePath
    Let asFolders = Split(myPath, "\")
    
    If UBound(asFolders) >= LBound(asFolders) Then
        Let I = UBound(asFolders)
        Do Until (Len(sSlug) > 0) Or (I < LBound(asFolders))
            If Not RegexMatch(asFolders(I), "Cont[r]?ol\s*File[s]?") Then
                Let sSlug = asFolders(I)
            End If
            Let I = I - 1
        Loop
    End If
    
    Let FileFolderSlug = sSlug
End Function

Public Property Let Url(sUrl As String)
    Dim aPath() As String
    Dim sPath As String
    Dim iStartPath As Integer
    Dim I As Integer
    
    aPath = Split(sUrl, "\")
    FileName = aPath(UBound(aPath))
    
    sDrive = ""
    sPath = ""
    
    Let iStartPath = LBound(aPath)
    
    If RegexMatch(Left(sUrl, 2), "[A-Za-z][:]") Then
        Let sDrive = Left(sUrl, 2)
        Let iStartPath = iStartPath + 1
    ElseIf Left(sUrl, 2) = "\\" Or Left(sUrl, 2) = "//" Then
        Let sDrive = "\\" & aPath(iStartPath + 2) & "\" & aPath(iStartPath + 3)
        Let iStartPath = iStartPath + 4
    End If
    
    For I = iStartPath To UBound(aPath) - 1
        If Len(aPath(I)) = 0 Then
            'NOOP
        ElseIf (Len(sDrive) = 0) And (aPath(I) = PathStart) Then
            sDrive = "\" & sPath
            sPath = "\" & aPath(I)
        Else
            sPath = sPath & "\" & aPath(I)
        End If
    Next I
    FilePath = sPath
End Property

Public Property Get Url() As String
    Dim Normalize As Boolean
    
    Dim sUrl As String
    sUrl = Drive & FilePath & "\" & FileName
    If (Normalize) Then
        sUrl = Replace(sUrl, "\", "/")
        sUrl = "file:///" & sUrl
    End If
    
    Url = sUrl
End Property

Public Property Get Drive() As String
    If Len(sDrive) = 0 Then
        sDrive = getDefaultDrive
    End If
    Drive = sDrive
End Property

Public Property Get SheetType() As String
    Let SheetType = sSheetType
    
    If Not bSheetTypeSet Then
        If IsAccnSheet Then
            Let SheetType = "ACCN"
        ElseIf IsDocumentationSheet(Version:=2) Then
            Let SheetType = GetSheetTypeFromSlug(RegexComponent(FileName, MATCH_DOCUMENTATION_V1, 2))
        ElseIf IsDocumentationSheet(Version:=1) Then
            Let SheetType = RegexComponent(FileName, MATCH_DOCUMENTATION_V2, 2)
        End If
    End If
End Property

Public Property Let SheetType(s As String)
    Let sSheetType = s
    Let bSheetTypeSet = True
End Property

Public Function SheetTypeSlug()
    Let SheetTypeSlug = GetSlugFromSheetType(SheetType)
End Function

Public Function Exists() As Boolean
    Let Exists = Len(Dir(Drive & FilePath & "\" & FileName)) > 0
End Function

Public Function RecordedFilePath() As String
    RecordedFilePath = sFilePath
End Function

Public Function IsAccnSheet(Optional ByVal Version As Integer) As Boolean
    Dim Pattern As String
    
    Select Case Version
    Case 1:
        Let Pattern = MATCH_FILENAME_V1
    Case 2:
        Let Pattern = MATCH_FILENAME_V2
    Case Else:
        Let Pattern = "(" & MATCH_FILENAME_V1 & "|" & MATCH_FILENAME_V2 & ")"
    End Select
    
    Let IsAccnSheet = RegexMatch(FileName, Pattern)

End Function

Public Function IsDocumentationSheet(Optional ByVal Version As Integer) As Boolean
    Dim Pattern As String
    
    Select Case Version
    Case 1:
        Let Pattern = MATCH_DOCUMENTATION_V1
    Case 2:
        Let Pattern = MATCH_DOCUMENTATION_V2
    Case Else:
        Let Pattern = "(" & MATCH_DOCUMENTATION_V1 & "|" & MATCH_DOCUMENTATION_V2 & ")"
    End Select

    Let IsDocumentationSheet = RegexMatch(FileName, Pattern)
End Function

Public Function MATCH_FILENAME_V1() As String
    Let MATCH_FILENAME_V1 = "^([A-Z0-9]+)_([0-9]+)_([0-9]+)_([0-9]{4})?([0-9]{2})?([0-9]{2})?(.*)[.]PDF(|.*)$"
End Function

Public Function MATCH_FILENAME_V2() As String
    Let MATCH_FILENAME_V2 = "^([A-Z0-9]{2,3})([0-9]{4})([0-9]{4})()()()(.*)[.]PDF(|.*)$"
End Function

Public Function MATCH_DOCUMENTATION_V1() As String
    Let MATCH_DOCUMENTATION_V1 = "^([A-Z0-9]{2,3})_(Correspondence|Documentation|Administrative|Administration|Clipping|Microfilm)_(.*)([.]PDF)$"
End Function

Public Function MATCH_DOCUMENTATION_V2() As String
    Let MATCH_DOCUMENTATION_V2 = "^([A-Z0-9]{2,3})(Corr|Doc|Admin|Clip|Microfilm)(.*)([.]PDF)$"
End Function

Public Function MATCH_COPIER_SCAN() As String
    Let MATCH_COPIER_SCAN = "^copier[.]archives@archives[.]alabama[.]gov_([0-9]+)_([0-9]+)([.]PDF)$"
End Function

Public Function MetadataFromFileName() As Variant
    Dim bFound As Boolean
    
    Dim oRef As Variant
    Dim cRefs As Variant
    
    Dim vPattern As Variant
    Dim cPatterns As New Collection: With cPatterns
        .Add MATCH_FILENAME_V2
        .Add MATCH_FILENAME_V1
    End With
    
    Dim oIsAccnFile As New RegExp: With oIsAccnFile
        .IgnoreCase = True
    End With
    
    For Each vPattern In cPatterns
        oIsAccnFile.Pattern = vPattern
        Set cRefs = oIsAccnFile.Execute(FileName)
        If cRefs.Count > 0 Then
            For Each oRef In cRefs
                Set MetadataFromFileName = oRef.SubMatches
                Let bFound = True
                Exit For
            Next oRef
        Else
            Set MetadataFromFileName = Nothing
        End If
        
        If bFound Then
            Exit For
        End If
    Next vPattern
    
End Function

Public Property Get ACCN() As String
    Dim oRefs As Variant
    
    If Len(sACCN) = 0 Then
        Set oRefs = MetadataFromFileName
    
        'Get the ACCN # from the FileName
        Let sACCN = oRefs(1) & "." & oRefs(2)
    End If
    Let ACCN = sACCN
End Property

Public Property Let ACCN(s As String)
    Let sACCN = s
End Property

Public Function Creator() As String
    Dim oRefs As Variant
    Dim cWords As Collection
    Dim sSlug As String
    Dim aPath() As String
    Dim I As Integer
    
    Set oRefs = MetadataFromFileName
    
    'Get the ACCN # from the FileName
    If Not oRefs Is Nothing Then
        Creator = oRefs(0)
    Else
        Let aPath = Split(FilePath, "\")
        Let I = UBound(aPath)
        
        If I >= LBound(aPath) Then
            If RegexMatch(aPath(I), "^Contr?olFiles?$") Then
                Let I = I - 1
            End If
        End If
        
        If I >= LBound(aPath) Then
            Let Creator = aPath(I)
            
            Let sSlug = aPath(I)
            Set cWords = camelCaseSplitString(sSlug)
            
            If cWords.Count > 0 Then
                Let Creator = cWords.Item(1)
            End If
        End If
    End If

End Function

Public Function AccnDate() As Date
    Dim oRefs As Variant
    
    Set oRefs = MetadataFromFileName
    If Not oRefs(3) = Empty Then
        AccnDate = DateSerial(Val(oRefs(3)), Val(oRefs(4)), Val(oRefs(5)))
    End If
End Function

Public Sub InsertIntoAccessions(Optional ByVal OnCurrent As Boolean, Optional ByVal NewRecord As Boolean, Optional ByRef ListView As Variant)
    Dim sSuggestion As String
    Dim rsAccnScan As DAO.Recordset
    Dim rsAccessions As DAO.Recordset
    Dim origPath As String
    Dim origFileName As String
    Dim bFileNameToBeFixed As Boolean
    
    Let bFileNameToBeFixed = False
    If Not OnCurrent Then
        SecureCreatorRecord Creator
        
        'Save old values before updating
        origPath = Drive & FilePath
        origFileName = FileName
        ForceCorrectFileName
            
        If origFileName <> sFileNameCorrected Then
            Name origPath & "\" & origFileName As origPath & "\" & sFileNameCorrected
        End If
    
        Set rsAccessions = CurrentDb.OpenRecordset("SELECT * FROM Accessions WHERE ACCN = '" & ACCN & "'")

        'Use the metadata for this Accession into the Accessions table
        If rsAccessions.EOF Then
            Set rsAccessions = CurrentDb.OpenRecordset("Accessions")
            rsAccessions.AddNew
            rsAccessions!ACCN = ACCN
            rsAccessions!AccnDate = AccnDate
            rsAccessions!Creator = Creator
            rsAccessions!CabinetFolder = FileFolder
            rsAccessions!AddedToDatabaseBy = "cwj"
            rsAccessions!AddedToDatabaseOn = Now
            rsAccessions.Update
        End If
    Else
        If RegexMatch(FileName, MATCH_COPIER_SCAN) And NewRecord Then
            Let bFileNameToBeFixed = True
        End If
    End If
    
    'Use the file name / path for this Accession scan into the AccnScans table
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef
    Dim oAccession As cAccession
    Dim cScans As Collection
    Dim vScan As Variant
    
    'Check to see whether or not this is a duplicate of an existing AccnScan record
    Set oAccession = New cAccession: Let oAccession.ACCN = ACCN
    Set cScans = oAccession.Attachments
    For Each vScan In cScans
        If (vScan.FileName = FileName) And (vScan.RecordedFilePath = FilePath) Then
            Err.Raise vbObjectError + 255, "InsertIntoAccessions: I DON'T THINK HE'LL BE VERY INTERESTED IN " & FileName & ". HE'S ALREADY GOT ONE, YOU SEE? IT'SA VERY NICE-UH!"
        End If
    Next vScan
    Set oAccession = Nothing
    
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qInsertAccnScans": On Error GoTo 0
        
    Set oQuery = CurrentDb.CreateQueryDef( _
        Name:="qInsertAccnScans", _
        SQLText:="INSERT INTO AccnScans (ACCN, VoyagerID, LNUMBER, NonAccnId, CabinetFolder, FileName, FilePath, OldPath, FileNameToBeFixed) VALUES (" _
        & "[paramAccn], " _
        & "[paramVoyagerID], " _
        & "[paramLNUMBER], " _
        & "[paramNonAccnId], " _
        & "[paramCabinetFolder]," _
        & "[paramFileName], " _
        & "[paramFilePath], " _
        & "[paramOldPath], " _
        & "[paramFileNameToBeFixed])" _
    )
    
    With oQuery
        .Parameters("paramAccn") = Trim(UCase(ACCN))
        .Parameters("paramVoyagerID") = Null
        .Parameters("paramLNUMBER") = Null
        .Parameters("paramNonAccnId") = Null
        .Parameters("paramCabinetFolder") = CabinetFolder
        .Parameters("paramFileName") = Trim(FileName)
        .Parameters("paramFilePath") = Trim(FilePath)
        .Parameters("paramOldPath") = IIf(Len(OldPath) > 0, OldPath, Null)
        .Parameters("paramFileNameToBeFixed") = bFileNameToBeFixed
        .Execute
    End With
    
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qInsertAccnScans": On Error GoTo 0
    
    Dim sOldFileName As String
    Dim sNewFileName As String
    If bFileNameToBeFixed Then
        Let sOldFileName = FileName
        
        If Not IsMissing(ListView) Then
            If IsObject(ListView) And (Not ListView Is Nothing) And OnCurrent Then
                ListView.Blank
            End If
        End If
        
        ConvertFileName Result:=sNewFileName
        
        If Len(sNewFileName) > 0 And sNewFileName <> sOldFileName Then
                
            On Error Resume Next: CurrentDb.QueryDefs.Delete "qUpdateAccnScans": On Error GoTo 0
            
            Set oQuery = CurrentDb.CreateQueryDef( _
                Name:="qUpdateAccnScans", _
                SQLText:="SELECT * FROM AccnScans " _
                & "WHERE FileName=[paramFileName] " _
                & "AND FilePath=[paramFilePath]" _
            )
            oQuery.Parameters("paramFileName") = Trim(sOldFileName)
            oQuery.Parameters("paramFilePath") = Trim(FilePath)

            Set rsAccnScan = oQuery.OpenRecordset
            Do Until rsAccnScan.EOF
                rsAccnScan.Edit
                rsAccnScan!FileName = sNewFileName
                rsAccnScan!FileNameToBeFixed = False
                rsAccnScan.Update
                
                rsAccnScan.MoveNext
            Loop
            rsAccnScan.Close
            Set rsAccnScan = Nothing
            
            On Error Resume Next: CurrentDb.QueryDefs.Delete "qUpdateAccnScans": On Error GoTo 0

        End If
    End If
    
    Set oQuery = Nothing

End Sub

Public Sub ForceCorrectFileName()
    Dim sAccnYear As String
    Dim sAccnPoint As String
    Dim sAccnDate As String
    Dim oRefs As Variant
    
    Set oRefs = MetadataFromFileName
    
    sAccnDate = oRefs(3) & oRefs(4) & oRefs(5)
    
    If Val(oRefs(1)) < 1900 Then
        sAccnYear = Format(1900 + Val(oRefs(1)), "0000")
    Else
        sAccnYear = Format(Val(oRefs(1)), "0000")
    End If
    
    sAccnPoint = Format(Val(oRefs(2)), "0000")
    
    Let sFileNameCorrected = oRefs(0) & "_" & sAccnYear & "_" & sAccnPoint & "_" & sAccnDate & oRefs(6) & ".PDF"
End Sub

Public Sub ConvertFileName(ByRef Result As String)
    Dim oIsOldAccnFile As New RegExp
    Dim reIsCopierScan As New RegExp
    
    Dim MetaD As Variant
    Dim sOldNew As String
    Dim sNewFileName As String
    Dim oRef As Variant
    Dim cRefs As Variant
    Dim I As Integer
    
    Dim sOldFull As String
    Dim sNewFull As String
    
    With oIsOldAccnFile
        .IgnoreCase = True
        .Pattern = MATCH_FILENAME_V1
    End With
    
    With reIsCopierScan
        .IgnoreCase = True
        .Pattern = MATCH_COPIER_SCAN
    End With
    
    Set cRefs = oIsOldAccnFile.Execute(FileName)
    If cRefs.Count > 0 Then
        For Each oRef In cRefs
            Set MetaD = oRef.SubMatches
            Exit For
        Next oRef
        
        Let sNewFileName = MetaD.Item(0) & MetaD.Item(1) & MetaD.Item(2) & MetaD.Item(6) & ".PDF" & MetaD.Item(7)
    ElseIf reIsCopierScan.Test(FileName) Then
        Let sNewFileName = Creator & Replace(ACCN, ".", "") & ".PDF"
    End If
    
    If Len(sNewFileName) > 0 Then

TryRename:
        On Error GoTo CatchRename
        
        Let Result = ""
        
        Let sOldFull = Drive & FilePath & "\" & FileName
        Let sNewFull = Drive & FilePath & "\" & sNewFileName
        If Me.Exists Then
            If (sOldFull <> sNewFull) Then
                Name sOldFull As sNewFull
                Let Result = sNewFileName
                Let FileName = sNewFileName
            End If
        Else
            Debug.Print "DID/COULD NOT RENAME: ", Url, "DOES NOT EXIST..."
        End If

FinallyRename:
        On Error GoTo 0
    
    Else
        Set MetaD = Nothing
    End If
    Exit Sub
    
CatchRename:
    Debug.Print "DID/COULD NOT RENAME: ", Url, "TRY AGAIN..."
    Let sOldNew = sNewFileName
    Let I = I + 1
    Let sNewFileName = MetaD.Item(0) & MetaD.Item(1) & MetaD.Item(2) & MetaD.Item(6) & "-" & Format(I, "00") & ".PDF" & MetaD.Item(7)
    
    If sOldNew <> sNewFileName Then Resume TryRename
    Resume FinallyRename

End Sub

Public Function isOnDefaultDrive() As Boolean
    Dim FSO As New FileSystemObject
    Dim FO(1 To 2) As Folder
    Dim driveNames(1 To 2) As String
    
    Set FO(1) = FSO.GetFolder(Drive & FilePath)
    Set FO(2) = FSO.GetFolder(getDefaultDrive & "\")
    
    'We need to account for the possibility of mapped network drives
    Dim I As Integer
    For I = 1 To 2
        Let driveNames(I) = IIf(Len(FO(I).Drive.ShareName) > 0, FO(I).Drive.ShareName, FO(I).Drive.Path)
    Next I
    
    Let isOnDefaultDrive = (driveNames(1) = driveNames(2))
End Function

Public Function FolderHomePath()
    Dim oFolder As cCabinetFolder
    
    Set oFolder = New cCabinetFolder: With oFolder
    .FileSourcePath = FileFolderSlug
    End With
    
    If oFolder.ID <> 0 Then
        Let FolderHomePath = getDefaultDrive & oFolder.FolderPath
    End If
 
End Function

Public Sub doCopyTo(NewPath As String)
    Dim sOldFullPath As String
    Dim sNewFullPath As String
    
    Let sOldFullPath = Drive & FilePath & "\" & FileName
    Let sNewFullPath = NewPath & "\" & FileName
    
    Name sOldFullPath As sNewFullPath
    Let Url = sNewFullPath
End Sub

'**
'*
'**
Private Function retrieveScanFilePath()
    Dim sRelativePath As String
    Dim f As String
    Dim sPattern As String
    Dim aWords() As String
    Dim cSearchDirs As New Collection
    Dim vSearchDir As Variant
    
    Dim vScanDir As Variant
    Dim cScanDirs As New Collection
    Dim sDirPrefix As String
    Dim I As Integer
    
    cScanDirs.Add "\AgencyState"
    cScanDirs.Add "\AgencyLocal"
    cScanDirs.Add "\AgencyCourts"
    cScanDirs.Add "\AgencyUS"
    cScanDirs.Add "\CollectionsManagement\AgencyFiles\State"
    cScanDirs.Add "\CollectionsManagement\AgencyFiles\Local"
    
    For Each vScanDir In cScanDirs
        aWords = Split(FileName, "_", 2)
        If LBound(aWords) = 0 And UBound(aWords) > 0 Then
            sPattern = aWords(0) & "*"
        Else
            sPattern = "*"
        End If
    
        sDirPrefix = vScanDir & "\"
        f = Dir(Drive & sDirPrefix & sPattern, vbDirectory)
        Do While Len(f) > 0
            cSearchDirs.Add (sDirPrefix & f & "\ControlFile")
            cSearchDirs.Add (sDirPrefix & f & "\ControlFiles")
            f = Dir()
        Loop
    Next vScanDir

    For Each vSearchDir In cSearchDirs
        sPattern = Drive & vSearchDir & "\" & FileName
        If Dir(sPattern, vbNormal) <> "" Then
            retrieveScanFilePath = vSearchDir
            Exit For
        End If
    Next vSearchDir
    
End Function

Public Function GetSheetTypeFromSlug(Slug As String) As String
    Dim vType As Variant
    
    Let GetSheetTypeFromSlug = Slug
    
    For Each vType In dSheetTypeSlugs.Keys()
        If UCase(Slug) = UCase(dSheetTypeSlugs.Item(vType)) Then
            Let GetSheetTypeFromSlug = CStr(vType)
            Exit For
        End If
    Next vType
    
End Function

Public Function GetSlugFromSheetType(SheetType As String) As String
    Dim vType As Variant
    
    Let GetSlugFromSheetType = SheetType
    
    For Each vType In dSheetTypeSlugs.Keys()
        If UCase(SheetType) = UCase(CStr(vType)) Then
            Let GetSlugFromSheetType = CStr(dSheetTypeSlugs.Item(vType))
            Exit For
        End If
    Next vType
End Function

Private Sub Class_Initialize()
    Set dSheetTypeSlugs = New Dictionary: With dSheetTypeSlugs
        .Add Key:="Correspondence", Item:="Corr"
        .Add Key:="Documentation", Item:="Doc"
        .Add Key:="Administrative", Item:="Admin"
        .Add Key:="Administration", Item:="Admin"
        .Add Key:="Clipping", Item:="Clip"
    End With
End Sub

Private Sub Class_Terminate()
    dSheetTypeSlugs.RemoveAll
    Set dSheetTypeSlugs = Nothing
End Sub
