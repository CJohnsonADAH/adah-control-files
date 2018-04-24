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

Private sPathStart As String
Private sDrive As String
Private sFilePath As String
Private sFileName As String

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
    sFileName = sFN
End Property

Public Property Let PathStart(sPS As String)
    sPathStart = sPS
End Property

Public Property Get FileFolder()
    Dim myPath As String
    Dim aFolders() As String
    Dim prevFolder As String
    Dim curFolder As Variant
    Dim myFolder As String
    Dim oFolder As cCabinetFolder
    
    myPath = FilePath
    aFolders = Split(myPath, "\")
    
    If UBound(aFolders) > 0 Then
        If UBound(aFolders) - LBound(aFolders) > 0 Then
            myFolder = aFolders(UBound(aFolders) - 1)
        Else
            myFolder = aFolders(UBound(aFolders))
        End If
        
        Set oFolder = New cCabinetFolder
        Let oFolder.FilePath = myFolder
        If oFolder.ID = 0 Then
            oFolder.Insert
        End If
        
        If oFolder.ID <> 0 Then
            myFolder = oFolder.Label
        End If
        
    Else
        myFolder = ""
    End If
    
    '
    'prevFolder = ""
    'For Each curFolder In aFolders
    '    myFolder = prevFolder & " " & JoinCollection(" ", camelCaseSplitString(curFolder & ""))
    '
    '    'when we move to the next curFolder, preserve this one as previous
    '    prevFolder = JoinCollection(" ", camelCaseSplitString(curFolder & ""))
    'Next curFolder
    
    FileFolder = myFolder
    
End Property

Public Property Get PathStart() As String
    If Len(sPathStart) = 0 Then
        sPathStart = "CollectionsManagement"
    End If
    
    PathStart = sPathStart
End Property

Public Property Let Url(sUrl As String)
    Dim aPath() As String
    Dim sPath As String
    Dim I As Integer
    
    aPath = Split(sUrl, "\")
    FileName = aPath(UBound(aPath))
    
    sDrive = ""
    sPath = ""
    For I = LBound(aPath) To UBound(aPath) - 1
        If Right(aPath(I), 1) = ":" Then
            sDrive = aPath(I)
        ElseIf Len(aPath(I)) = 0 Then
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

Private Function MetadataFromFileName() As Variant
    Dim oIsAccnFile As New RegExp
    With oIsAccnFile
        .IgnoreCase = True
        .Pattern = "^([A-Z0-9]+)_([0-9]+)_([0-9]+)_([0-9]{4})?([0-9]{2})?([0-9]{2})?(.*)[.]PDF(|.*)$"
    End With
    
    Set cRefs = oIsAccnFile.Execute(FileName)
    If cRefs.Count > 0 Then
        For Each oRef In cRefs
            Set MetadataFromFileName = oRef.SubMatches
            Exit For
        Next oRef
    Else
        Set MetadataFromFileName = Nothing
    End If
    
End Function

Public Function ACCN() As String
    Dim oRefs As Variant
    
    Set oRefs = MetadataFromFileName
    
    'Get the ACCN # from the FileName
    ACCN = oRefs(1) & "." & oRefs(2)

End Function

Public Function Creator() As String
    Dim oRefs As Variant
    
    Set oRefs = MetadataFromFileName
    
    'Get the ACCN # from the FileName
    If Not oRefs Is Nothing Then
        Creator = oRefs(0)
    End If

End Function

Public Function AccnDate() As Date
    Dim oRefs As Variant
    
    Set oRefs = MetadataFromFileName
    If Not oRefs(3) = Empty Then
        AccnDate = DateSerial(Val(oRefs(3)), Val(oRefs(4)), Val(oRefs(5)))
    End If
End Function

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
    
    Let FileName = oRefs(0) & "_" & sAccnYear & "_" & sAccnPoint & "_" & sAccnDate & oRefs(6) & ".PDF"
End Sub

Public Sub InsertIntoAccessions()
    Dim sSuggestion As String
    Dim rsAccessions As DAO.Recordset
    
    SecureCreatorRecord Creator
    
    'Save old values before updating
    origPath = Drive & FilePath
    origFileName = FileName
    ForceCorrectFileName
        
    If origFileName <> FileName Then
        Name origPath & "\" & origFileName As origPath & "\" & FileName
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
    
    'Use the file name / path for this Accession scan into the AccnScans table
    putAccnScanAttachments ACCN:=ACCN, FileName:=FileName, FilePath:=FilePath
    
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
    
    cScanDirs.Add "State"
    cScanDirs.Add "Local"
    For Each vScanDir In cScanDirs
        aWords = Split(FileName, "_", 2)
        If LBound(aWords) = 0 And UBound(aWords) > 0 Then
            sPattern = aWords(0) & "_*"
        Else
            sPattern = "*"
        End If
    
        sDirPrefix = "\CollectionsManagement\AgencyFiles\" & vScanDir & "\"
            
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

