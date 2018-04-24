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
'**

Option Explicit
Option Compare Database

Private sACCN As String
Private cAttachments As Collection
Private bAttachmentsSettled As Boolean

'**
'* Get cAccession::ACCN: get the Accession number of the record to which this object refers
'*
'* @return String A dotted accession number (ACCN), for example "1992.0003"
'**
Public Property Get ACCN() As String
    ACCN = sACCN
End Property

'**
'* Let cAccession::ACCN: reset the object to point to a given Accession number
'*
'* @param String AccnNo A dotted accession number (ACCN), for example "1992.0003"
'**
Public Property Let ACCN(AccnNo As String)
    sACCN = AccnNo
    
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
    bAttachmentsSettled = False
End Sub

Private Sub RetrieveAttachmentsFromDB()
    Dim oAccnScan As cAccnScan
    
    Dim Rs As DAO.Recordset
    Dim oQuery As DAO.QueryDef
    Dim I As Integer, iCount As Integer

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
        oAccnScan.FilePath = Rs!FilePath
        oAccnScan.FileName = Rs!FileName
        cAttachments.Add oAccnScan
        Rs.MoveNext
    Loop

    Set Rs = Nothing
    On Error Resume Next: CurrentDb.QueryDefs.Delete "qScanFilesByAccn": On Error GoTo 0
    
End Sub

