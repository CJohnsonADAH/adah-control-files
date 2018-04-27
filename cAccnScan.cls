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

Private bNewRecord As Boolean
Private sPathStart As String
Private sDrive As String
Private sFilePath As String
Private sFileName As String
Private sOldPath As String
Private sFileNameCorrected As String
Private sSheetType As String, bSheetTypeSet As Boolean
Private tTimestamp As Date, bTimestampSet As Boolean
Private iCabinetFolder As Long, oFolder As cCabinetFolder, bCabinetFolderSet As Boolean
Private sACCN As String
Private bFileNameToBeFixed As Boolean

Private dSheetTypeSlugs As Dictionary
Private dFileSystemMetadata As Dictionary
Private dDatabaseMetadata As Dictionary


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
        Let CabinetFolder = retrieveCabinetFolderID
    End If
End Property

Public Property Let CabinetFolder(l As Long)
    Let iCabinetFolder = l
    Let bCabinetFolderSet = True
End Property

Public Property Get NewRecord() As Boolean
    Let NewRecord = bNewRecord
End Property

Public Property Let NewRecord(b As Boolean)
    Let bNewRecord = b
End Property

Private Function Folder() As cCabinetFolder
    If oFolder Is Nothing Then
        ResetFolderWithSlugs
    End If
    Set Folder = oFolder
End Function

Private Sub ResetFolderWithSlugs()
    Set oFolder = New cCabinetFolder: With oFolder
        .FileSourcePath = FileFolderSlug(Old:=True)
    End With
    
    If oFolder.ID = 0 Then
        Set oFolder = New cCabinetFolder: With oFolder
            .FilePath = FileFolderSlug(Old:=False)
        End With
    End If
End Sub

Private Function retrieveCabinetFolderID() As Integer
    ResetFolderWithSlugs
    If oFolder.ID <> 0 Then
        Let retrieveCabinetFolderID = oFolder.ID
    Else
        Let retrieveCabinetFolderID = 0
    End If

End Function

Public Function FileFolder() As String
    Dim ID As Long
    
    Let ID = CabinetFolder
    If ID = 0 Then Folder.Insert
    
    If ID <> 0 Then Let FileFolder = Folder.Label
End Function

Public Function FileFolderSlug(Optional ByVal Old As Variant) As String
    Dim myPath(1 To 2) As String
    
    Dim I As Integer
    Dim J As Integer
    
    Dim sSlug As String
    Dim asFolders() As String
    
    Dim UsePath(1 To 2) As Boolean
    
    If IsMissing(Old) Then
        Let UsePath(1) = True
        Let UsePath(2) = True
    Else
        Let UsePath(1) = CBool(Old)
        Let UsePath(2) = Not CBool(Old)
    End If
    
    Let myPath(1) = OldPath
    Let myPath(2) = FilePath
    
    Let I = LBound(myPath)
    Do Until Len(sSlug) > 0 Or I > UBound(myPath)
        If UsePath(I) Then
            Let asFolders = Split(myPath(I), "\")
            
            Let J = UBound(asFolders)
            Do Until (Len(sSlug) > 0) Or (J < LBound(asFolders))
                If Not RegexMatch(asFolders(J), "(^Cont[r]?ol\s*File[s]?$|[.][A-Za-z0-9]+$)") Then
                    Let sSlug = asFolders(J)
                End If
                Let J = J - 1
            Loop
        End If
        Let I = I + 1
    Loop
    
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

Public Property Get Timestamp() As Date
    Let Timestamp = tTimestamp
    
    If bTimestampSet Then
        Let Timestamp = tTimestamp
    Else
        Let Timestamp = AccnDate
    End If
End Property

Public Property Let Timestamp(t As Date)
    Let tTimestamp = t
    Let bTimestampSet = True
End Property

Public Function SheetTypeSlug()
    Let SheetTypeSlug = GetSlugFromSheetType(SheetType)
End Function

Public Property Get FileNameToBeFixed() As Boolean
    Let FileNameToBeFixed = bFileNameToBeFixed
End Property

Public Property Let FileNameToBeFixed(b As Boolean)
    Let bFileNameToBeFixed = b
End Property

Public Function Exists() As Boolean
    Let Exists = Len(Dir(Url)) > 0
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

Public Function IsCopierScan() As Boolean
    Let IsCopierScan = RegexMatch(FileName, MATCH_COPIER_SCAN)
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
        
        If Not oRefs Is Nothing Then
            'Get the ACCN # from the FileName
            Let sACCN = oRefs(1) & "." & oRefs(2)
        Else
            Let sACCN = ""
        End If
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
    If Not oRefs Is Nothing Then
        If Not oRefs(3) = Empty Then
            AccnDate = DateSerial(Val(oRefs(3)), Val(oRefs(4)), Val(oRefs(5)))
        End If
    End If
End Function

'**
'* cAccnScan::InsertIntoAccnScans: use a parameterized SQL query to INSERT this object into a fresh record in AccnScans
'*
'* @uses cAccnScan::ACCN
'* @uses cAccnScan::CabinetFolder
'* @uses cAccnScan::FileName
'* @uses cAccnScan::FilePath
'* @uses cAccnScan::OldPath
'* @uses cAccnScan::FileNameToBeFixed
'**
Public Sub InsertIntoAccnScans(Optional ByRef Controller As Form)
    Dim oQuery As DAO.QueryDef
    
    If Controller Is Nothing Then
        On Error Resume Next: CurrentDb.QueryDefs.Delete "qInsertAccnScans": On Error GoTo 0
        
        Set oQuery = CurrentDb.CreateQueryDef( _
            Name:="qInsertAccnScans", _
            SQLText:="INSERT INTO AccnScans (ACCN, VoyagerID, LNUMBER, NonAccnId, [Timestamp], SheetType, CabinetFolder, FileName, FilePath, OldPath, FileNameToBeFixed) VALUES (" _
            & "[paramAccn], " _
            & "[paramVoyagerID], " _
            & "[paramLNUMBER], " _
            & "[paramNonAccnId], " _
            & "[paramTimestamp], " _
            & "[paramSheetType], " _
            & "[paramCabinetFolder], " _
            & "[paramFileName], " _
            & "[paramFilePath], " _
            & "[paramOldPath], " _
            & "[paramFileNameToBeFixed])" _
        )
        
        With oQuery
            .Parameters("paramAccn") = IIf(Len(ACCN) > 0, Trim(UCase(ACCN)), Null)
            .Parameters("paramVoyagerID") = Null
            .Parameters("paramLNUMBER") = Null
            .Parameters("paramNonAccnId") = Null
            .Parameters("paramSheetType") = SheetType
            .Parameters("paramTimestamp") = IIf(Timestamp <> 0, Timestamp, Null)
            .Parameters("paramCabinetFolder") = CabinetFolder
            .Parameters("paramFileName") = Trim(FileName)
            .Parameters("paramFilePath") = Trim(FilePath)
            .Parameters("paramOldPath") = IIf(Len(OldPath) > 0, OldPath, Null)
            .Parameters("paramFileNameToBeFixed") = FileNameToBeFixed
            .Execute
        End With
        
        On Error Resume Next: CurrentDb.QueryDefs.Delete "qInsertAccnScans": On Error GoTo 0
    Else
    
        With Controller
            .Controls("ACCN").value = IIf(Len(ACCN) > 0, Trim(UCase(ACCN)), Null)
            .Controls("FileName").value = FileName
            .Controls("FilePath").value = FilePath
            .Controls("SheetType").value = SheetType
            .Controls("Timestamp").value = Timestamp
            .Controls("CabinetFolder").value = CabinetFolder
            .Controls("OldPath").value = OldPath
            .Controls("FileNameToBeFixed").value = FileNameToBeFixed Or (NewRecord And IsCopierScan)
        End With
        
    End If
End Sub

Public Sub InsertIntoAccessions(Optional ByVal OnCurrent As Boolean, Optional ByVal NewAccessionRecord As Boolean, Optional ByRef ListView As Variant)
    Dim sSuggestion As String
    Dim rsAccnScan As DAO.Recordset
    
    'FIXME: Do we ever need to be careful to respect a pre-existing setting?
    Let FileNameToBeFixed = False
    
    If Not OnCurrent Then
        SecureAccessionsRecord
    Else
        Let FileNameToBeFixed = FileNameToBeFixed Or (IsCopierScan And NewAccessionRecord)
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
            Err.Raise EX_ACCNSCAN_DUPLICATE, "InsertIntoAccessions: I DON'T THINK HE'LL BE VERY INTERESTED IN " & FileName & ". HE'S ALREADY GOT ONE, YOU SEE? IT'SA VERY NICE-UH!"
        End If
    Next vScan
    Set oAccession = Nothing
    
    InsertIntoAccnScans
    
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
        
        Let sOldFull = Url
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
    
    Let sOldFullPath = Url
    Let sNewFullPath = NewPath & "\" & FileName
    
    Name sOldFullPath As sNewFullPath
    
    Let Url = sNewFullPath
    Let OldPath = sOldFullPath
End Sub

Private Sub SecureAccessionsRecord()
    Dim oAccession As cAccession
    Dim origPath As String
    Dim origFileName As String
    
    SecureCreatorRecord Creator
        
    'Save old values before updating
    origPath = Drive & FilePath
    origFileName = FileName
    ForceCorrectFileName

    If origFileName <> sFileNameCorrected Then
        doCopyTo NewPath:=(origPath & "\" & sFileNameCorrected)
    End If
    
    Set oAccession = New cAccession: With oAccession
        .ACCN = ACCN
        .AccnDate = AccnDate
        .Creator = Creator
        .CabinetFolder = FileFolder
        .Sign
    End With
    
    oAccession.Secure

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
    
    Set dFileSystemMetadata = New Dictionary
    Set dDatabaseMetadata = New Dictionary
End Sub

Private Sub Class_Terminate()
    dSheetTypeSlugs.RemoveAll
    Set dSheetTypeSlugs = Nothing
    
    Set dFileSystemMetadata = New Dictionary
    Set dDatabaseMetadata = New Dictionary
End Sub
