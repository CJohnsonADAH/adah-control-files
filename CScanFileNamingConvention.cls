VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CScanFileNamingConvention"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private sBaseName As String
Private sRelativePath As String
Private sDrive As String

Public Property Get BaseName() As String: Let BaseName = sBaseName: End Property
Public Property Let BaseName(s As String): Let sBaseName = s: End Property

Public Property Get RelativePath() As String
    If Len(sRelativePath) = 0 Then
        Let RelativePath = retrieveScanFilePath
    Else
        Let RelativePath = sRelativePath
    End If
End Property
Public Property Let RelativePath(s As String): Let sRelativePath = s: End Property

Public Property Get Drive() As String
    If Len(sDrive) = 0 Then
        sDrive = getDefaultDrive
    End If
    Drive = sDrive
End Property
Public Property Let Drive(s As String): Let sDrive = s: End Property

Public Property Let Url(sUrl As String)
    Dim aPath() As String
    Dim sPath As String
    Dim iStartPath As Integer
    Dim I As Integer
    
    aPath = Split(sUrl, "\")
    
    '1. BASENAME: Take from the terminal element
    Let BaseName = aPath(UBound(aPath))
    
    '2. DRIVE AND RELATIVEPATH: Take from prior elements
    Let Drive = ""
    Let RelativePath = ""
    
    Let iStartPath = LBound(aPath)
    
    If RegexMatch(Left(sUrl, 2), "[A-Za-z][:]") Then
        Let Drive = Left(sUrl, 2)
        Let iStartPath = iStartPath + 1
    ElseIf Left(sUrl, 2) = "\\" Or Left(sUrl, 2) = "//" Then
        Let Drive = "\\" & aPath(iStartPath + 2) & "\" & aPath(iStartPath + 3)
        Let iStartPath = iStartPath + 4
    End If
    
    For I = iStartPath To UBound(aPath) - 1
        If Len(aPath(I)) = 0 Then
            'NOOP
        Else
            sPath = sPath & "\" & aPath(I)
        End If
    Next I
    Let RelativePath = sPath
End Property

Public Property Get Url() As String
    'FIXME: Normalize should be a parameter
    Dim Normalize As Boolean
    
    Dim sUrl As String
    
    Let sUrl = AbsolutePath & "\" & BaseName
    If (Normalize) Then
        sUrl = Replace(sUrl, "\", "/")
        sUrl = "file:///" & sUrl
    End If
    
    Url = sUrl
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function AbsolutePath() As String
    Let AbsolutePath = Drive & RelativePath
End Function

Public Function RecordedFilePath() As String
    Let RecordedFilePath = sRelativePath
End Function

Public Function Conflicts(v As Variant) As Boolean
    Dim o As CScanFileNamingConvention
    If IsObject(v) Then
        If TypeOf v Is CScanFileNamingConvention Then
            Set o = v
        Else
            Set o = v.NamingConvention
        End If
        
        Let Conflicts = (Me.BaseName = o.BaseName) And (Me.RecordedFilePath = Me.RelativePath)
    End If
End Function

Public Function Creator() As String
    Dim oRefs As Variant
    Dim cWords As Collection
    Dim sSlug As String
    Dim aPath() As String
    Dim Last As Integer

    Set oRefs = Metadata
    
    'Get the Creator from BaseName, if available (e.g., "SOS" from "SOS20120199.PDF")
    If Not oRefs Is Nothing Then
        Let Creator = oRefs(0)
        
    'It's not in the BaseName, alack, so get the Creator from the RelativePath (e.g., "SOS" from "SOSSecretaryOfState")
    Else
        Let aPath = Split(RelativePath, "\"): Let Last = UBound(aPath)
        If Last >= LBound(aPath) Then
            If RegexMatch(aPath(Last), "^Contr?olFiles?$") Then
                Let Last = Last - 1
            End If
        End If
        
        If Last >= LBound(aPath) Then
            Let Creator = aPath(Last)
            
            Let sSlug = aPath(Last)
            Set cWords = camelCaseSplitString(sSlug)
            
            If cWords.Count > 0 Then
                Let Creator = cWords.Item(1)
            End If
        End If
    End If

End Function

Public Function Metadata() As Variant
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
        Set cRefs = oIsAccnFile.Execute(BaseName)
        If cRefs.Count > 0 Then
            For Each oRef In cRefs
                Set Metadata = oRef.SubMatches
                Let bFound = True
                Exit For
            Next oRef
        Else
            Set Metadata = Nothing
        End If
        
        If bFound Then
            Exit For
        End If
    Next vPattern
    
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
    
    Let IsAccnSheet = RegexMatch(BaseName, Pattern)

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

    Let IsDocumentationSheet = RegexMatch(BaseName, Pattern)
End Function

Public Function IsCopierScan() As Boolean
    Let IsCopierScan = RegexMatch(BaseName, MATCH_COPIER_SCAN)
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
        aWords = Split(BaseName, "_", 2)
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
        sPattern = Drive & vSearchDir & "\" & BaseName
        If Dir(sPattern, vbNormal) <> "" Then
            retrieveScanFilePath = vSearchDir
            Exit For
        End If
    Next vSearchDir
    
End Function

