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
Private iID As Long
Private oFolder As cCabinetFolder
Private sACCN As String
Private bFileNameToBeFixed As Boolean

Private dSheetTypeSlugs As Dictionary
Private dFileSystemMetadata As Dictionary
Private dDatabaseMetadata As Dictionary

Private oNamingConvention As CScanFileNamingConvention

'**
'* Get cAccnScan::ID:
'**
Public Property Get ID() As Long
    Let ID = iID
End Property

'**
'* Let cAccnScan::ID:
'**
Public Property Let ID(L As Long)
    Let iID = L
    RetrieveFromDatabase Field:="ID", Value:=L
End Property

'**
'* Get cAccnScan::FilePath: get the filesystem path of an AccnScan file
'*
'* @return String A file system path to this Accession sheet scan, either from the database record or
'*      from a quick search of the file system. Does not include drive letter or network authority.
'*      For example, \CollectionsManagement\AgencyFiles\Local\63_TuscaloosaCounty\ControlFiles
'**
Public Property Get FilePath() As String
    Let FilePath = NamingConvention.RelativePath
End Property

'**
'* Let cAccnScan::FilePath: set a value for the filesystem path of an AccnScan file
'*
'* @param String A file system path to this Accession sheet scan. Does not include drive letter or
'*      network authority.
'*      For example, \CollectionsManagement\AgencyFiles\Local\63_TuscaloosaCounty\ControlFiles
'**
Public Property Let FilePath(sFP As String)
    Let NamingConvention.RelativePath = sFP
End Property

Public Property Get FileName() As String
    Let FileName = NamingConvention.BaseName
End Property

Private Sub SetFileSystemMetaData(Key As String, Item As Variant)
    If Not dFileSystemMetadata.Exists(Key) Then
        dFileSystemMetadata.Add Key:=Key, Item:=Null
    End If
    
    Let dFileSystemMetadata.Item(Key) = Item
End Sub

Private Sub SetDatabaseMetaData(Key As String, Item As Variant)
    If Not dDatabaseMetadata.Exists(Key) Then
        dDatabaseMetadata.Add Key:=Key, Item:=Null
    End If
    
    Let dDatabaseMetadata.Item(Key) = Item
End Sub

Public Property Let FileName(sFN As String)
    Dim cMeta As Variant
    Dim oAx As cAccession
    
    sACCN = ""
    Let NamingConvention.BaseName = sFN
    
    Set cMeta = NamingConvention.Metadata
    If Not cMeta Is Nothing Then
        dFileSystemMetadata.RemoveAll
        
        'Set the Creator
        If Not cMeta(0) = Empty Then
            SetFileSystemMetaData "Creator", cMeta(0)
        End If
        
        'If it's available, set the ACCN number
        If Not cMeta(1) = Empty And Not cMeta(2) = Empty Then
            Set oAx = New cAccession: With oAx
                .Major = cMeta(1)
                .Minor = cMeta(2)
            End With
            SetFileSystemMetaData "ACCN", oAx.ACCN
            Set oAx = Nothing
            
        End If
        
        'If it can be mined from the file naming convention, set the timestamp
        If Not cMeta(3) = Empty And Not cMeta(3) = "ND" Then
            SetFileSystemMetaData "Timestamp", DateSerial(Val(cMeta(3)), Val(cMeta(4)), Val(cMeta(5)))
        End If
        
        'If it can be mined from the file naming convention, set the SheetType
        If NamingConvention.IsAccnSheet Then
            SetFileSystemMetaData "SheetType", "ACCN"
        ElseIf NamingConvention.IsDocumentationSheet(Version:=2) Then
            SetFileSystemMetaData "SheetType", GetSheetTypeFromSlug(RegexComponent(NamingConvention.BaseName, NamingConvention.MATCH_DOCUMENTATION_V1, 2))
        ElseIf NamingConvention.IsDocumentationSheet(Version:=1) Then
            SetFileSystemMetaData "SheetType", RegexComponent(NamingConvention.BaseName, NamingConvention.MATCH_DOCUMENTATION_V2, 2)
        End If
    
    End If

End Property

Public Property Get OldPath() As String
    If dDatabaseMetadata.Exists("OldPath") Then
        Let OldPath = dDatabaseMetadata.Item("OldPath")
    End If
End Property
Public Property Let OldPath(s As String)
    SetDatabaseMetaData Key:="OldPath", Item:=s
    Let NamingConvention.OldPath = s
End Property

Public Property Get CabinetFolder() As Long
    If dDatabaseMetadata.Exists("CabinetFolder") Then
        Let CabinetFolder = dDatabaseMetadata.Exists("CabinetFolder")
    Else
        Let CabinetFolder = NamingConvention.CabinetFolder
    End If
End Property

Public Property Let CabinetFolder(L As Long)
    SetDatabaseMetaData "CabinetFolder", L
    ResetFolderWithID
End Property

Public Property Get NewRecord() As Boolean
    Let NewRecord = bNewRecord
End Property

Public Property Let NewRecord(b As Boolean)
    Let bNewRecord = b
End Property

Private Function Folder() As cCabinetFolder
    If oFolder Is Nothing Then
        resetFolderWithSlugs
    End If
    Set Folder = oFolder
End Function

Private Sub ResetFolderWithID()
    Set oFolder = New cCabinetFolder: With oFolder
        .ID = CabinetFolder
    End With
End Sub

Private Sub resetFolderWithSlugs()
    Set oFolder = New cCabinetFolder: With oFolder
        .FileSourcePath = NamingConvention.FileFolderSlug(Old:=True)
    End With
    
    If oFolder.ID = 0 Then
        Set oFolder = New cCabinetFolder: With oFolder
            .FilePath = NamingConvention.FileFolderSlug(Old:=False)
        End With
    End If
End Sub

Public Function FileFolder() As String
    Dim ID As Long
    
    Let ID = CabinetFolder
    If ID = 0 Then Folder.Insert
    
    If ID <> 0 Then Let FileFolder = Folder.Label
End Function

Public Property Let Url(s As String): Let NamingConvention.Url = s: End Property
Public Property Get Url() As String: Let Url = NamingConvention.Url: End Property

Public Property Get Drive() As String: Let Drive = NamingConvention.Drive: End Property

Public Property Get SheetType() As String
    If dDatabaseMetadata.Exists("SheetType") Then
        Let SheetType = dDatabaseMetadata.Item("SheetType")
    ElseIf dFileSystemMetadata.Exists("SheetType") Then
        Let SheetType = dFileSystemMetadata.Item("SheetType")
    Else
        Let SheetType = ""
    End If
End Property

Public Property Let SheetType(s As String)
    If Not dDatabaseMetadata.Exists("SheetType") Then
        dDatabaseMetadata.Add Key:="SheetType", Item:=s
    End If
    
    Let dDatabaseMetadata.Item("SheetType") = s
End Property

Public Property Get Timestamp() As Date
    If dDatabaseMetadata.Exists("Timestamp") Then
        Let Timestamp = dDatabaseMetadata.Item("Timestamp")
    ElseIf dFileSystemMetadata.Exists("Timestamp") Then
        Let Timestamp = dFileSystemMetadata.Item("Timestamp")
    Else
        Let Timestamp = 0
    End If
End Property

Public Property Let Timestamp(t As Date)
    If Not dDatabaseMetadata.Exists("Timestamp") Then
        dDatabaseMetadata.Add Key:="Timestamp", Item:=t
    End If
    
    Let dDatabaseMetadata.Item("Timestamp") = t
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

Public Property Get NamingConvention() As CScanFileNamingConvention
    Set NamingConvention = oNamingConvention
End Property

Public Property Set NamingConvention(o As CScanFileNamingConvention)
    Set oNamingConvention = o
End Property

Public Function Exists() As Boolean
    Let Exists = Len(Dir(Url)) > 0
End Function

Public Property Get ACCN() As String
    Dim oRefs As Variant
    
    If Len(sACCN) = 0 Then
        If dDatabaseMetadata.Exists("ACCN") Then
            Let sACCN = dDatabaseMetadata.Item("ACCN")
        ElseIf dFileSystemMetadata.Exists("ACCN") Then
            Let sACCN = dFileSystemMetadata.Item("ACCN")
        End If
    End If
    
    Let ACCN = sACCN
End Property

Public Property Let ACCN(s As String)
    Let sACCN = s
End Property

Public Function Creator() As String
    Let Creator = NamingConvention.Creator
End Function

Public Sub CheckFileNameToBeFixed()
    Let FileNameToBeFixed = FileNameToBeFixed Or NamingConvention.IsCopierScan
End Sub

Public Sub InsertIntoAccessions(Optional ByRef Controller As Form)
    Dim oAccession As cAccession
    
    'FIXME: Do we ever need to be careful to respect a pre-existing setting?
    Let FileNameToBeFixed = False
    
    If Controller Is Nothing Then
        SecureCreatorRecord Creator

        Set oAccession = New cAccession: With oAccession
        .ACCN = ACCN
        .AccnDate = Timestamp
        .Creator = Creator
        .CabinetFolder = Folder.Label
        .Sign
        End With
    
        oAccession.Secure
        
    Else
    
        Controller.Dirty = False
        
    End If
End Sub

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
    'Use the file name / path for this Accession scan into the AccnScans table
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef
    Dim oAccession As cAccession
    Dim cScans As Collection
    Dim vScan As Variant
    Dim Bundle As Object
    
    'Check to see whether or not this is a duplicate of an existing AccnScan record
    Set oAccession = New cAccession: Let oAccession.ACCN = ACCN
    Set cScans = oAccession.Attachments
    For Each vScan In cScans
        If (vScan.NamingConvention.Conflicts(Me.NamingConvention)) Then
            Err.Raise Number:=EX_ACCNSCAN_DUPLICATE, Source:="InsertIntoAccessions", Description:="I DON'T THINK HE'LL BE VERY INTERESTED IN " & FileName & ". HE'S ALREADY GOT ONE, YOU SEE? IT'SA VERY NICE-UH!"
        End If
    Next vScan
    Set oAccession = Nothing
    
    If Controller Is Nothing Then
        Set Rs = CurrentDb.OpenRecordset("AccnScans")
        
        With Rs
            .AddNew
            Set Bundle = .Fields
        End With
    Else
        Set Bundle = Controller.Controls
    End If
            
    With Bundle
        .Item("ACCN") = IIf(Len(ACCN) > 0, Trim(UCase(ACCN)), Null)
        .Item("VoyagerID") = Null
        .Item("LNUMBER") = Null
        .Item("NonAccnId") = Null
        .Item("SheetType") = SheetType
        .Item("Timestamp") = IIf(Timestamp <> 0, Timestamp, Null)
        .Item("CabinetFolder") = CabinetFolder
        .Item("FileName") = Trim(FileName)
        .Item("FilePath") = Trim(FilePath)
        .Item("OldPath") = IIf(Len(OldPath) > 0, OldPath, Null)
        .Item("FileNameToBeFixed") = FileNameToBeFixed Or (NewRecord And NamingConvention.IsCopierScan)
    End With

    If Controller Is Nothing Then
        With Rs
            .Update
            .Close
        End With
        
        Set Rs = Nothing
    End If
End Sub

Public Sub ConvertFileName(Optional ByRef Result As String)
    Dim oIsOldAccnFile As New RegExp
    Dim reIsBadDateScan As New RegExp
    
    Dim MetaD As Variant
    
    Dim sNewFileName As String
    Dim sNewFileBaseName As String
    Dim sNewFileExtension As String
    
    Dim sOldFileName As String, sOldFilePath As String
    
    Dim oRef As Variant
    Dim cRefs As Variant
    
    Dim sCurName As String
    
    With oIsOldAccnFile
        .IgnoreCase = True
        .Pattern = NamingConvention.MATCH_FILENAME_V1
    End With
    
    With reIsBadDateScan
        .IgnoreCase = True
        .Pattern = NamingConvention.MATCH_DOCUMENTATION_V2
    End With
    
    'Phase 1. Try to use database meta-data
    If SheetType = "ACCN" Then
        
        sCurName = GetCurNameFromCreatorCode(Creator)
        
        Let sNewFileExtension = ".PDF"
        If Len(ACCN) > 0 Then
            Let sNewFileBaseName = sCurName & Replace(ACCN, ".", "")
        Else
            Let sNewFileBaseName = sCurName & "_NA_" & Format(Timestamp, "YYYYmmdd")
        End If
        Let sNewFileName = sNewFileBaseName & sNewFileExtension

    ElseIf HasNamingConvention(SheetType) Then
    
        Let sCurName = GetCurNameFromCreatorCode(Creator)
        Let sNewFileExtension = ".PDF"
        
        Select Case SheetType
        Case "", "ACCN":
            Let sNewFileBaseName = sCurName & Replace(ACCN, ".", "")
        Case Else:
            Let sNewFileBaseName = sCurName & SheetTypeSlug & GetDateSlug(Timestamp)
        End Select
        Let sNewFileName = sNewFileBaseName & sNewFileExtension
        
    'Phase 2. Try to use file system naming convention meta-data
    ElseIf NamingConvention.IsAccnSheet(Version:=1) Then
        Set cRefs = oIsOldAccnFile.Execute(FileName)
        
        For Each oRef In cRefs
            Set MetaD = oRef.SubMatches
            Exit For
        Next oRef
        
        Let sNewFileBaseName = MetaD.Item(0) & MetaD.Item(1) & MetaD.Item(2) & MetaD.Item(6)
        Let sNewFileExtension = ".PDF" & MetaD.Item(7)
        Let sNewFileName = sNewFileBaseName & sNewFileExtension
    
    ElseIf NamingConvention.IsCopierScan Then

        Let sCurName = GetCurNameFromCreatorCode(Creator)
        Let sNewFileExtension = ".PDF"
        
        Select Case SheetType
        Case "", "ACCN":
            Let sNewFileBaseName = sCurName & Replace(ACCN, ".", "")
        Case Else:
            Let sNewFileBaseName = sCurName & SheetTypeSlug & GetDateSlug(Timestamp)
        End Select
        Let sNewFileName = sNewFileBaseName & sNewFileExtension
        
    ElseIf reIsBadDateScan.Test(FileName) Then
    
        Set cRefs = reIsBadDateScan.Execute(FileName)
        If cRefs.Count > 0 Then
            For Each oRef In cRefs
                Set MetaD = oRef.SubMatches
            Next oRef
            
            If Not MetaD Is Nothing Then
                If MetaD(2) = "18991230" Then
                    Let sNewFileBaseName = MetaD(0) & MetaD(1) & GetDateSlug(Timestamp)
                    Let sNewFileExtension = MetaD(3)
                    Let sNewFileName = sNewFileBaseName & sNewFileExtension
                End If
            End If
        End If
        
    End If
    
    Let sOldFileName = FileName
    Let sOldFilePath = FilePath
    
    FileSystemConvertFileName NewFileName:=sNewFileName, BaseName:=sNewFileBaseName, Extension:=sNewFileExtension, Result:=Result
    If Len(Result) > 0 Then
        DatabaseConvertFileNames OldFileName:=sOldFileName, OldFilePath:=sOldFilePath
    End If
End Sub

Public Sub FileSystemConvertFileName(ByVal NewFileName As String, ByVal BaseName As String, ByVal Extension As String, ByRef Result As String)
    Dim sOldFull As String, sNewFull As String
    Dim I As Integer, sOldNew As String

    Let Result = ""

TryRename:
    On Error GoTo CatchRename
        
    If Len(NewFileName) > 0 Then
        
        Let sOldFull = Url
        Let sNewFull = Drive & FilePath & "\" & NewFileName
        
        If Me.Exists Then
            If (sOldFull <> sNewFull) Then
                FileRename Source:=sOldFull, Destination:=sNewFull
                
                Let Result = NewFileName
                Let FileName = NewFileName
            End If
        Else
            Debug.Print "DID/COULD NOT RENAME: ", Url, "DOES NOT EXIST..."
        End If
    End If

FinallyRename:
    On Error GoTo 0
    Exit Sub
    
CatchRename:
    Let sOldNew = NewFileName
    If Err.Number = EX_FILEALREADYEXISTS Then
        Debug.Print "DID/COULD NOT RENAME: ", Url, "TRY AGAIN..."
        Let I = I + 1
        Let NewFileName = BaseName & "-" & Format(I, "00") & Extension
    ElseIf Err.Number = EX_FILEPERMISSIONDENIED Then
        'NOOP
        MsgBox "Denied"
    ElseIf Err.Number < 0 Then
        MsgBox Err.Source & ":" & Err.Number & ": " & Err.Description
    End If
    
    If sOldNew <> NewFileName Then Resume TryRename
    Resume FinallyRename


End Sub

Public Sub DatabaseConvertFileNames(OldFileName As String, OldFilePath As String)
    Dim Rs As DAO.Recordset, FileNamePath As Dictionary
    Dim FromTo As Dictionary
    
    Set FileNamePath = New Dictionary: With FileNamePath
        .Add Key:="FileName", Item:=Trim(OldFileName)
        .Add Key:="FilePath", Item:=Trim(OldFilePath)
    End With
    
    Set Rs = RsSelect(Parameters:=FileNamePath)
    With Rs
        Do Until .EOF
            .Edit
            !FileName.Value = FileName
            !FilePath.Value = FilePath
            !FileNameToBeFixed.Value = False
            .Update
        
            Set FromTo = New Dictionary: With FromTo
                .Add Key:="Source", Item:=OldFileName
                .Add Key:="Destination", Item:=FileName
            End With
            DoAction Outlet:="FileHasBeenRenamedInDatabase", Parameters:=FromTo
        
            .MoveNext
        Loop
        .Close
    End With
    Set Rs = Nothing

End Sub

Public Sub doCopyTo(NewPath As String)
    Dim sOldFullPath As String
    Dim sNewFullPath As String
    
    Let sOldFullPath = Url
    Let sNewFullPath = NewPath & "\" & FileName
    
    FileRename Source:=sOldFullPath, Destination:=sNewFullPath
    
    Let Url = sNewFullPath
    Let OldPath = sOldFullPath
End Sub

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

Private Sub RetrieveFromDatabase(ByVal Field As String, ByVal Value As Variant)
    Dim Rs As DAO.Recordset
    Dim v As Variant
    
    Set Rs = RsSelect(Field:=Field, Value:=Value)
    If Not Rs.EOF Then
        Let FilePath = Nz(Rs!FilePath.Value)
        Let FileName = Nz(Rs!FileName.Value)
        Let OldPath = Nz(Rs!OldPath.Value)
        Let ACCN = Nz(Rs!ACCN.Value)
        Let SheetType = Nz(Rs!SheetType.Value)
        Let Timestamp = Nz(Rs!Timestamp.Value)
        Let CabinetFolder = Nz(Rs!CabinetFolder.Value)
        Let FileNameToBeFixed = Nz(Rs!FileNameToBeFixed.Value)
    End If
    Rs.Close
    
    Set Rs = Nothing
    
End Sub

Private Function RsSelect(Optional ByVal Field As String, Optional ByVal Value As Variant, Optional Parameters As Dictionary, Optional ByVal Operator As String) As DAO.Recordset
    Dim SQL As String
    Dim WhereCondition As String
    
    Dim vKey As Variant
    Dim sKey As String
    Dim vValue As Variant
    
    Dim Qy As DAO.QueryDef
    
    If Len(Operator) = 0 Then
        Let Operator = "AND"
    End If
    
    If Parameters Is Nothing Then
        Set Parameters = New Dictionary
    End If
    
    If Len(Field) > 0 Then
        If Not Parameters.Exists(Field) Then
            Parameters.Add Key:=Field, Item:=Value
        End If
    End If
    
    Let SQL = "SELECT * FROM [AccnScans]"
    For Each vKey In Parameters.Keys
        
        Let vValue = Parameters.Item(vKey)
        
        Let sKey = Nz(vKey)
        If Len(WhereCondition) = 0 Then
            Let WhereCondition = " WHERE "
        Else
            Let WhereCondition = WhereCondition & " " & Operator & " "
        End If
        
        Let WhereCondition = WhereCondition & "[" & sKey & "] = [param" & sKey & "]"
    Next vKey
    Let SQL = SQL & WhereCondition
    Set Qy = CurrentDb.CreateQueryDef(Name:="", SQLText:=SQL)
    
    For Each vKey In Parameters.Keys
        Let sKey = "param" & Nz(vKey)
        Let vValue = Parameters.Item(vKey)
        
        Let Qy.Parameters(sKey) = vValue
    Next vKey

    Set RsSelect = Qy.OpenRecordset

End Function

Public Function HasNamingConvention(Optional ByVal SheetType As String) As Boolean
    If Len(SheetType) = 0 Then
        Let SheetType = Me.SheetType
    End If
     
    Let HasNamingConvention = dSheetTypeSlugs.Exists(TitleCase(SheetType))
End Function

Private Sub Class_Initialize()
    Set dSheetTypeSlugs = New Dictionary: With dSheetTypeSlugs
        .Add Key:="Accn", Item:="Accn"
        .Add Key:="Correspondence", Item:="Corr"
        .Add Key:="Destruction", Item:="Destruction"
        .Add Key:="Documentation", Item:="Doc"
        .Add Key:="Administrative", Item:="Admin"
        .Add Key:="Administration", Item:="Admin"
        .Add Key:="Clipping", Item:="Clip"
        .Add Key:="Separation", Item:="Separation"
    End With
    
    Set dFileSystemMetadata = New Dictionary
    Set dDatabaseMetadata = New Dictionary
    
    Set oNamingConvention = New CScanFileNamingConvention
End Sub

Private Sub Class_Terminate()
    'Empty and unlink the dictionary objects
    dSheetTypeSlugs.RemoveAll: Set dSheetTypeSlugs = Nothing
    dFileSystemMetadata.RemoveAll: Set dFileSystemMetadata = Nothing
    dDatabaseMetadata.RemoveAll: Set dDatabaseMetadata = Nothing
    
    'Unlike the NamingConvention object
    Set oNamingConvention = Nothing
End Sub
