VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAccnScanListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private oBrowserControl
Private oListControl
Private sRecordTable
Private sACCN

Public Property Get ACCN() As String
    ACCN = sACCN
End Property

Public Property Let ACCN(sA As String)
    Let sACCN = sA
End Property

Public Property Set BrowserControl(oBC As Object)
    Set oBrowserControl = oBC
End Property

Public Property Get BrowserControl() As Object
    Set BrowserControl = oBrowserControl
End Property

Public Property Set ListControl(oLC As Object)
    Set oListControl = oLC
End Property

Public Property Get ListControl() As Object
    Set ListControl = oListControl
End Property

Public Property Get HasScans() As Boolean
    Let HasScans = (ListControl.ListCount > 0)
End Property

Public Property Get Scans() As Collection
    Dim cScans As New Collection
    
    For I = 0 To ListControl.ListCount - 1
        cScans.Add ListControl.ItemData(I)
    Next I
    
    Set Scans = cScans
End Property

Public Property Get RecordTable() As String
    If Len(sRecordTable) = 0 Then
        'DEFAULT VALUE: `Accessions` table
        sRecordTable = "Accessions"
    End If
    RecordTable = sRecordTable
End Property

'**
'* cAccnScanListView::Blank: set the PDF viewer to a blank page and block until it completes the switch
'*
'**
Public Sub Blank()
    BrowserControl.Navigate2 "about:blank"
    Do
        DoEvents
    Loop Until BrowserControl.ReadyState >= 3
End Sub

'**
'* cAccnScanListView::View: set the PDF viewer to a file at a given path (or URL)
'*
'* @param String FileName
'**
Public Sub View(FileName As String)
    BrowserControl.Navigate2 FileName
End Sub

Public Sub Fill(vAttachments As Variant)
    Dim vAccnScan As Variant
    
    ListControl.Requery
End Sub

Public Sub ClearList()
    ListControl.Requery
End Sub

Public Sub DetachSelectedFromRecord(Optional ByVal ID As Integer)
    Dim vIndex As Variant
    Dim iIndex As Integer
    Dim sFileName As String
    Dim sFilePath As String
    Dim aFileParts() As String
    Dim SuccessfullyRemoved As Boolean
    
    SuccessfullyRemoved = False 'So far
    For Each vIndex In ListControl.ItemsSelected
        iIndex = Val(vIndex)
        sFileName = ListControl.Column(0, iIndex)
        sFilePath = ListControl.Column(1, iIndex)
        
        If Len(ACCN) = 0 Then
            'Delete using numeric ID...
            Err.Raise vbObjectError + 255, "DetachSelectedFromRecord Not Implemented Without ACCN"
        Else
            deleteAccnScanAttachments ACCN:=ACCN, FileName:=sFileName, FilePath:=sFilePath
            SuccessfullyRemoved = True
        End If
        
        If SuccessfullyRemoved Then
            'Removed from the database, so let's remove it from the interface control
            ListControl.Requery
        End If
    Next vIndex

End Sub

Public Sub AttachToRecord(ByVal Location As String, Optional ByVal ID As Integer)
    Dim oAccnScan As New cAccnScan

    'Location should contain a full path, e.g. "S:\CollectionsManagement\AgencyFiles\Local\03_MontgomeryCounty\ControlFiles\03_1994_0451_19940912.PDF"
    'Could be sourced to a DOS drive letter ("S:") or to a a network authority ("\\adahfs1\GR-Collections")
    'In any case, cAccnScan was written to help parse all this, so DRY.
    Let oAccnScan.Url = Location

    'INSERT INTO AccnScan VALUES (Me!ACCN.Value, sFilePath, sFileName)
    On Error GoTo HandleInsertError
    If Len(ACCN) > 0 Then
        putAccnScanAttachments ACCN, oAccnScan.FileName, oAccnScan.FilePath
    Else
        'Insert using numeric ID for NonAccessioned
        Err.Raise vbObjectError + 255, "AttachToRecord not implemented without ACCN"
    End If
    On Error GoTo 0
    
    ListControl.Requery

    Exit Sub
    
HandleInsertError:
    MsgBox "FIXME: INSERT ERROR..."
    
    ListControl.Requery
    
End Sub

Public Sub ReloadPDF(Optional ByVal selectedFileName As String, Optional ByVal selectedFilePath As String)
    
    Dim hasPDF As Boolean
    
    Dim cAttachmentPaths As Collection
    Dim vFilePath As Variant
    
    Dim oSelection As cAccnScan
    Dim oComparison As cAccnScan
    Dim oScanFileToView As cAccnScan
    
    Dim sFilePathToView As String
    Dim sComparisonFilePath As String
    Dim sSelectedFilePath As String
    
    Dim selectionProvided As Boolean
    
    Blank
    hasPDF = False
        
    '--| (1) First: Either get an existing file system path, or secure one for an attached blob |---
    'If we have a way to do so then get a Collection of attached files
    If Len(ACCN) > 0 Then
        Set cAttachmentPaths = getAccnScanAttachments(ACCN)
    Else
        Set cAttachmentPaths = New Collection
    End If
    
    selectionProvided = (Len(selectedFileName) + Len(selectedFilePath) > 0)
    
    Set oSelection = New cAccnScan
    If selectionProvided Then
        With oSelection
            .FileName = selectedFileName
            .FilePath = selectedFilePath
        End With
    End If
    
    'Use AccnScans table records to get FilePath and FileName
    If cAttachmentPaths.Count > 0 Then
        'Step through the Collection of attached files, checking each against selection
        For Each vFilePath In cAttachmentPaths
            Set oComparison = New cAccnScan: With oComparison
                .FileName = vFilePath(2)
                .FilePath = vFilePath(1)
            End With

            If (Not selectionProvided) Or (oComparison.Url = oSelection.Url) Then
                sFilePathToView = oComparison.Url
                hasPDF = True
                Exit For
            End If
        Next vFilePath
        
    'There doesn't seem to be anything anywhere, so all we've got is a blank
    Else
        hasPDF = False
        
    End If
        
    '--| (2) Now: Set the PDF viewer control to view the PDF at the secured location |------------
    If hasPDF Then
        View sFilePathToView
    Else
        Blank
    End If

End Sub

Public Sub ReloadPDFFromList()
    Dim aFileParts() As String
       
    ReloadPDF selectedFileName:=ListControl.Column(0), selectedFilePath:=ListControl.Column(1)

End Sub
