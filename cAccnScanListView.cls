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

Private oBrowserControl As Object
Private oListControl As Object
Private oForm As Object
Private sRecordTable As String
Private sACCN As String

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

Public Property Get Form() As Object
    Set Form = oForm
End Property

Public Property Set Form(o As Object)
    Set oForm = o
End Property

Public Property Get HasScans() As Boolean
    If Not ListControl Is Nothing Then
        Let HasScans = (ListControl.ListCount > 0)
    ElseIf Not Form Is Nothing Then
        Let HasScans = (Len(Nz(Form!FileName)) > 0)
    Else
        Let HasScans = False
    End If
End Property

Public Property Get Scans() As Collection
    Dim I As Integer
    Dim cScans As New Collection
    
    If Not ListControl Is Nothing Then
        For I = 0 To ListControl.ListCount - 1
            cScans.Add ListControl.ItemData(I)
        Next I
    ElseIf Not Form Is Nothing Then
        cScans.Add Form!FileName
    End If
    
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
Public Sub View(ByVal FileName As String)
    BrowserControl.Navigate2 FileName
End Sub

Public Sub Fill(vAttachments As Variant)
    Dim vAccnScan As Variant
    
    If Not ListControl Is Nothing Then
        ListControl.Requery
    End If
End Sub

Public Sub ClearList()
    If Not ListControl Is Nothing Then
        ListControl.Requery
    End If
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
            Err.Raise EX_NOHOOK, "DetachSelectedFromRecord Not Implemented Without ACCN"
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

Public Sub AttachToRecord(Location As Variant, Optional NewRecord As Boolean, Optional ByVal ID As Integer)
    Dim I As Integer
    Dim f As Object
    Dim sACCN As String
    Dim sDate As String
    Dim aParts() As String
    Dim sCreator As String
    Dim oAccnScan As New cAccnScan
    
    'Location should contain a full path, e.g. "S:\CollectionsManagement\AgencyFiles\Local\03_MontgomeryCounty\ControlFiles\03_1994_0451_19940912.PDF"
    'Could be sourced to a DOS drive letter ("S:") or to a a network authority ("\\adahfs1\GR-Collections")
    'In any case, cAccnScan was written to help parse all this, so DRY.
    If TypeName(Location) = "String" Then
        Let oAccnScan.Url = Location
    ElseIf TypeName(Location) = "cAccnScan" Then
        Set oAccnScan = Location
    End If
    
    'INSERT INTO AccnScan VALUES (Me!ACCN.Value, sFilePath, sFileName)
    Dim oAx As cAccession
    If Len(ACCN) > 0 Then
        Let oAccnScan.ACCN = ACCN
    End If
    
    If Form Is Nothing Then
        If Len(ACCN) > 0 Then
    
            On Error GoTo InsertIntoAccessionsCatch
            oAccnScan.InsertIntoAccessions OnCurrent:=True, NewRecord:=NewRecord, ListView:=Me
            On Error GoTo 0
    
        Else
            
            'Insert using numeric ID for NonAccessioned
            Err.Raise EX_NOHOOK, "cAccnScanListView.AttachToRecord", "Not implemented for empty ACCN"
        
        End If
        
    End If
    
    If Not ListControl Is Nothing Then
        ListControl.Requery
    End If
    
    If Not Form Is Nothing Then
        Let Form!FileName.value = oAccnScan.FileName
        Let Form!FilePath.value = oAccnScan.FilePath
        Let Form!CabinetFolder.value = oAccnScan.CabinetFolder
    End If
    
    Exit Sub
    
InsertIntoAccessionsCatch:
    MsgBox "INSERT ERROR: " & Err.Source & " - " & Err.Description

    Resume Next
    
End Sub

Public Sub ReloadPDF(Optional ByVal selectedFileName As String, Optional ByVal selectedFilePath As String)
    
    Dim hasPDF As Boolean
    
    Dim cAttachmentPaths As Collection
    Dim vScan As Variant
    
    Dim oAccession As cAccession
    
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
    selectionProvided = (Len(selectedFileName) + Len(selectedFilePath) > 0)
    
    Set oSelection = New cAccnScan
    If selectionProvided Then
        With oSelection
            .FileName = selectedFileName
            .FilePath = selectedFilePath
        End With
    End If
    
    'If we have a way to do so then get a Collection of attached files
    If Len(ACCN) > 0 Then
        Set oAccession = New cAccession: oAccession.ACCN = ACCN
        Set cAttachmentPaths = oAccession.Attachments
    Else
        Set cAttachmentPaths = New Collection
        If Not Form Is Nothing Then
            If Len(Nz(Form!FileName.value)) > 0 Then
                cAttachmentPaths.Add oSelection
            End If
        End If
    End If
    
    'Use AccnScans table records to get FilePath and FileName
    If cAttachmentPaths.Count > 0 Then
        'Step through the Collection of attached files, checking each against selection
        For Each vScan In cAttachmentPaths
            Set oComparison = vScan

            If (Not selectionProvided) Or (oComparison.Url = oSelection.Url) Then
                sFilePathToView = oComparison.Url
                hasPDF = True
                Exit For
            End If
        Next vScan
    
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

Public Sub ReloadPDFFromForm()
    If Not Form Is Nothing Then
        ReloadPDF selectedFileName:=Nz(Form!FileName.value), selectedFilePath:=Nz(Form!FilePath.value)
    End If
End Sub
