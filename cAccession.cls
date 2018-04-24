VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAccession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**
'* @class cAccession Class representing Accession records in the Accessions tables as linked to scan images in the AccnScans table
'*
'* @version 2018.0424
'**

Option Explicit
Option Compare Database

Private sMajor As String
Private sMinor As String
Private cAttachments As Collection
Private bAttachmentsSettled As Boolean

Private MajorFormat As String
Private MinorFormat As String

'**
'* Get cAccession::ACCN: get the Accession number of the record to which this object refers
'*
'* @uses cAccession::Major
'* @uses cAccession::Minor
'*
'* @return String A dotted accession number (ACCN), for example "1992.0003"
'**
Public Property Get ACCN() As String
    Let ACCN = Major & "." & Minor
End Property

'**
'* Let cAccession::ACCN: reset the object to point to a given Accession number
'*
'* @uses cAccession::Major
'* @uses cAccession::Minor
'* @uses cAccession::MATCH_WELL_FORMED
'* @uses RegexComponent
'*
'* @param String AccnNo A dotted accession number (ACCN), for example "1992.0003"
'**
Public Property Let ACCN(AccnNo As String)
    'Extract the parts of a dotted accession number and assign
    Let Major = RegexComponent(AccnNo, MATCH_WELL_FORMED, 1)
    Let Minor = RegexComponent(AccnNo, MATCH_WELL_FORMED, 2)
End Property

'**
'* Get cAccession::Major: get the major component of the Accession number
'* (e.g., 2000 from 2000.0123). Per ISO, this is typically the year when
'* the records were accessioned.
'*
'* @return String A formatted, zero-padded number, for example, "1992"
'**
Public Property Get Major() As String
    Dim iMajor As Long
    
    If Len(sMajor) > 0 Then
        Let iMajor = CLng(sMajor)
    Else
        Let iMajor = 0
    End If
    
    Let Major = Format(iMajor, MajorFormat)
End Property

'**
'* Let cAccession::Major: (re)set the major component of the Accession number
'*
'* @uses cAccession::canUseForMajor
'* @uses ResetAttachments
'*
'* @param String AccnNo A well-formed major component of an accession number
'*      for example "1992", from the ACCN number "1992.0003"
'**
Public Property Let Major(s As String)
    Dim iMajor As Long
    
    If canUseForMajor(s) Then
        Let iMajor = CLng(s)
        Let sMajor = Format(iMajor, MajorFormat)
    Else
        Let sMajor = ""
    End If
    
    'Reset the Attachments list
    ResetAttachments
End Property

'**
'* Get cAccession::Minor: get the minor component of the Accession number
'* (e.g., 0123 from 2000.0123). Per ISO, this is typically the order in
'* which the records were accessioned during the year.
'*
'* @return String A formatted, zero-padded number, for example, "0003"
'**
Public Property Get Minor() As String
    Dim iMinor As Long
    
    Let iMinor = CLng(sMinor)
    Let Minor = Format(iMinor, "0000")
End Property

'**
'* Let cAccession::Major: (re)set the minor component of the Accession number
'*
'* @uses cAccession::canUseForMinor
'* @uses ResetAttachments
'*
'* @param String AccnNo A well-formed minor component of an accession number
'*      zero-padding optional; e.g. "3" or "0003", from ACCN "1992.0003"
'**
Public Property Let Minor(s As String)
    Dim iMinor As Long

    If canUseForMinor(s) Then
        Let iMinor = CLng(s)
        Let sMinor = Format(iMinor, MinorFormat)
    Else
        Let sMinor = ""
    End If
    
    'Reset the Attachments list
    ResetAttachments
End Property

'**
'* Get cAccession::Attachments: get a Collection of cAccnScan objects representing all the
'* scanned sheets or related documentation attached to this Accession (by ACCN number)
'*
'* @uses cAccession::SettleAttachments
'*
'* @return Collection of cAccnScan objects, one for each attached scanned document
'**
Public Property Get Attachments() As Collection
    SettleAttachments
    Set Attachments = cAttachments
End Property

Public Function canUseForMajor(s As String) As Boolean
    Let canUseForMajor = RegexMatch(s, MATCH_WELL_FORMED_MAJOR)
End Function

Public Function canUseForMinor(s As String) As Boolean
    Let canUseForMinor = RegexMatch(s, MATCH_WELL_FORMED_MINOR)
End Function

Public Function MATCH_WELL_FORMED() As String
    Let MATCH_WELL_FORMED = "^" & MATCH_WELL_FORMED_MAJOR("(", ")") & "[.]" & MATCH_WELL_FORMED_MINOR("(", ")") & "$"
End Function

Public Function MATCH_WELL_FORMED_MAJOR(Optional ByVal Prefix As String, Optional ByVal Suffix As String) As String
    If Len(Prefix) = 0 Then Let Prefix = "^"
    If Len(Suffix) = 0 Then Let Suffix = "$"
    
    Let MATCH_WELL_FORMED_MAJOR = Prefix & "[0-9]+" & Suffix
End Function

Public Function MATCH_WELL_FORMED_MINOR(Optional ByVal Prefix As String, Optional ByVal Suffix As String) As String
    If Len(Prefix) = 0 Then Let Prefix = "^"
    If Len(Suffix) = 0 Then Let Suffix = "$"
    
    Let MATCH_WELL_FORMED_MINOR = Prefix & "[0-9]+" & Suffix
End Function

'**
'* SettleAttachments: check whether or not the list of attachments to this Accession have been
'* retrieved from the database, and if not, retrieve it from the database and make it available
'* for the cAccession::Attachments() property, etc.
'**
Private Sub SettleAttachments()
    If cAttachments Is Nothing Then
        Set cAttachments = New Collection
    End If
    
    If Not bAttachmentsSettled Then
        RetrieveAttachmentsFromDB
        bAttachmentsSettled = True
    End If
    
End Sub

Private Sub ResetAttachments()
    Set cAttachments = New Collection
    Let bAttachmentsSettled = False
End Sub

Private Sub RetrieveAttachmentsFromDB()
    Dim oAccnScan As cAccnScan
    
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef
    Dim i As Integer, iCount As Integer

    'check the database for comments flagged with this CollectionNumber
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qScanFilesByAccn": On Error GoTo 0
    Set oQuery = CurrentDb.CreateQueryDef( _
        Name:="qScanFilesByAccn", _
        SQLText:="SELECT * FROM AccnScans " _
        & "WHERE TRIM(ACCN) = TRIM([paramAccn])" _
    )
    oQuery.Parameters("paramAccn") = UCase(ACCN)
        
    Set Rs = oQuery.OpenRecordset
    
    Do Until Rs.EOF
        Set oAccnScan = New cAccnScan
        If Not IsNull(Rs!FilePath) Then
            oAccnScan.FilePath = Rs!FilePath
        Else
            oAccnScan.FilePath = ""
        End If
        
        If Not IsNull(Rs!FileName) Then
            oAccnScan.FileName = Rs!FileName
        Else
            oAccnScan.FileName = ""
        End If
        cAttachments.Add oAccnScan
        Rs.MoveNext
    Loop

    Set Rs = Nothing
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qScanFilesByAccn": On Error GoTo 0
    
End Sub

Private Sub Class_Initialize()
    Let MajorFormat = "0000"
    Let MinorFormat = "0000"
End Sub
