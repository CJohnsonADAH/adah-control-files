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

Private sCurName As String
Private sAgencyName As String
Private sShortName As String
Private bRetrieved As Boolean

Public Property Get CurName() As String
    CurName = sCurName
End Property

Public Property Let CurName(ByVal s As String)
    Let sCurName = s

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
        Set Rs = CurrentDb.OpenRecordset("SELECT * FROM Creators WHERE CurName = '" & Replace(CurName, "'", "''") & "'")
        
        If Not (Rs.EOF) Then
            If Not IsNull(Rs!AgencyName) Then Let AgencyName = Rs!AgencyName Else Let AgencyName = ""
            If Not IsNull(Rs!ShortName) Then Let ShortName = Rs!ShortName Else Let ShortName = ""
        Else
            'FIXME: Pitch a fit somehow
            Let AgencyName = ""
            Let ShortName = ""
        End If
        
        Set Rs = Nothing
        
        bRetrieved = True
    End If
End Sub
