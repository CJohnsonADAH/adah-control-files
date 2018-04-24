VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAgency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**
'* @class cAgency Model Class to represent entries in the Creators table. Each Accession in Government Records has a Creator,
'* typically identified by a two- or three-character alphanumeric CurName code. (State agencies have a three letter code,
'* counties have a two-digit numeric code.)
'**
Option Compare Database
Option Explicit

Private sCreatorCode As String
Private sAgencyName As String
Private sShortName As String
Private sDivision As String
Private sSection As String
Private bRetrieved As Boolean

Public Property Get CreatorCode() As String
    Let CreatorCode = sCreatorCode
End Property

Public Property Let CreatorCode(ByVal s As String)
    Let sCreatorCode = s

    If bRetrieved Then
        bRetrieved = False
        sAgencyName = ""
        sShortName = ""
    End If
End Property

Public Property Get AgencyName() As String
    AgencyName = sAgencyName
End Property

Public Property Let AgencyName(ByVal s As String)
    Let sAgencyName = s
End Property

Public Property Get ShortName() As String
    ShortName = sShortName
End Property

Public Property Let ShortName(ByVal s As String)
    Let sShortName = s
End Property

Public Property Let Division(ByVal s As String)
    Let sDivision = s
End Property

Public Property Get Division() As String
    Let Division = sDivision
End Property

Public Property Let Section(ByVal s As String)
    Let sSection = s
End Property

Public Property Get Section() As String
    Let Section = sSection
End Property

Public Function Name(Optional ByVal Short As Boolean)
    Dim sName As String
    
    SecureRecord
    
    If Short And Len(ShortName) > 0 Then
        sName = ShortName
    Else
        sName = AgencyName
    End If
    Name = sName
End Function

Private Sub SecureRecord()
    Dim Rs As DAO.Recordset
    
    If Not bRetrieved Then
        'check the database for an alias
        'FIXME: Properly parameterize this SQL
        Set Rs = CurrentDb.OpenRecordset("SELECT * FROM Creators WHERE CreatorCode = '" & Replace(CreatorCode, "'", "''") & "'")
        
        If Not (Rs.EOF) Then
            Let AgencyName = Nz(Rs!AgencyName.value)
            Let ShortName = Nz(Rs!Agency.value)
            Let Division = Nz(Rs!Division.value)
            Let Section = Nz(Rs!Section.value)
            
        Else
            'FIXME: Pitch a fit somehow
            Let AgencyName = ""
            Let ShortName = ""
            Let Division = ""
            Let Section = ""
            
        End If
        
        Rs.Close
        Set Rs = Nothing
        
        bRetrieved = True
    End If
End Sub
