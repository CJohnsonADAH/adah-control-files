VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Accessions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'**
'* Property oAccnScanListView: Used to control an interface widget combining a List control of attachment filenames
'* and a Browser control for displaying the contents of those files, typically PDF blobs
'**
Private oAccnScanListView As cAccnScanListView

Private sLastCabinetFolderLabel As String
Private sLastCabinetFolderPath As String

Private Type AccessionDetails
    ACCN As String
    Creator As String
    Date As Date
End Type

Public Sub RefreshAttachmentView()
    Me!lstAccnScans.Requery
    AccnScanListView.ReloadPDFFromList
End Sub

Private Property Get AccnScanListView() As cAccnScanListView

    If oAccnScanListView Is Nothing Then
        Set oAccnScanListView = New cAccnScanListView
        Set oAccnScanListView.Form = Me
        Set oAccnScanListView.BrowserControl = Me!PreviewBrowser
        Set oAccnScanListView.ListControl = Me!lstAccnScans
        
        If Not IsNull(Me!ACCN.value) Then
            Let oAccnScanListView.ACCN = Me!ACCN.value
        Else
            Let oAccnScanListView.ACCN = ""
        End If
    End If
    
    If Not IsNull(Me!ACCN.value) Then
        Let oAccnScanListView.ACCN = Me!ACCN.value
    Else
        Let oAccnScanListView.ACCN = ""
    End If
    
    Set AccnScanListView = oAccnScanListView

End Property

Private Sub Update_txtCreatorName_from(Text As String, Optional NewRecord As Boolean)
    Dim oAgency As New cAgency
    Let oAgency.CreatorCode = Text

    Me!txtCreatorName.value = oAgency.Name(Short:=False)
    
    If Len(Me!txtCreatorName.value) > 0 Or NewRecord Then
        Me!txtCreatorName.BackColor = COLOR_DISABLED
    Else
        Me!txtCreatorName.BackColor = COLOR_ALERT
    End If
    
End Sub

Private Sub cmdAddScan_Click()
    Dim dlgSourceFile As FileDialog
    Dim vItem As Variant
    Dim bAttachmentError As Boolean
    Dim sDestinationFolder As String
    
    Me!ACCN.BackColor = COLOR_UNMARKED
    
    Set dlgSourceFile = Application.FileDialog(msoFileDialogFilePicker): With dlgSourceFile
        .AllowMultiSelect = True
        .Filters.Add Description:="PDF", Extensions:="*.PDF", Position:=1
        .Filters.Add Description:="All", Extensions:="*.*", Position:=2
    End With
    
    If dlgSourceFile.Show Then
        For Each vItem In dlgSourceFile.SelectedItems
            DoAttachScanFile File:=Nz(vItem), Halt:=bAttachmentError, DestinationFolder:=sDestinationFolder
            
            If bAttachmentError Then
                Exit For
            End If
        Next vItem
                
        AccnScanListView.ReloadPDF
    End If
    
End Sub

Private Sub DoAttachScanFile(ByVal File As String, ByRef Halt As Boolean, ByRef DestinationFolder As String)
    Dim dlgDestinationFolder As FileDialog
    Dim asFileFolderSlugs(1 To 2) As String
    Dim bWasNewRecord As Boolean
    
    Dim oAccnScan As New cAccnScan: With oAccnScan
        .Url = File
    End With

    Let asFileFolderSlugs(1) = oAccnScan.FileFolderSlug

    If Not oAccnScan.isOnDefaultDrive Then
        If Len(DestinationFolder) = 0 Then
            Let DestinationFolder = oAccnScan.FolderHomePath
        End If
                
        If Len(DestinationFolder) = 0 Then
            Set dlgDestinationFolder = Application.FileDialog(msoFileDialogFolderPicker): With dlgDestinationFolder
                .Title = "Destination Folder"
            End With
                
            If dlgDestinationFolder.Show Then
                Let DestinationFolder = dlgDestinationFolder.SelectedItems(1)
            End If
        End If
                                
        If Len(DestinationFolder) > 0 Then
            oAccnScan.doCopyTo NewPath:=DestinationFolder
            Let File = oAccnScan.Url
            Let asFileFolderSlugs(2) = oAccnScan.FileFolderSlug
        End If
    End If
    
    'If we have ACCN on the form, use it
    If Not IsNull(Me!ACCN.value) Then
        Let AccnScanListView.ACCN = Me!ACCN.value
    
    'If we don't have ACCN on the form, try to get it from the file meta-data
    Else
        Let AccnScanListView.ACCN = oAccnScan.ACCN

    End If
            
    Let bWasNewRecord = Me.NewRecord 'Status may be changed by error handler
    
AttachToRecordTry:
    On Error GoTo AttachToRecordCatch
    If oAccnScan.Timestamp = 0 And Not IsNull(Me!AccnDate.value) Then
        Let oAccnScan.Timestamp = Me!AccnDate.value
    End If
    AccnScanListView.AttachToRecord Location:=oAccnScan, NewRecord:=bWasNewRecord
    On Error GoTo 0
    
    Let Halt = False
    If bWasNewRecord Then
        If Not Halt And Not IsNull(Me!AccnDate.value) Then
            DoAdvanceToNextRecord
        Else
            Me!AccnDate.SetFocus
            Let Halt = True
        End If
    End If
    GoTo AttachToRecordFinally
    
AttachToRecordCatch:
    Dim Fixed As Boolean
    Dim ErrCode As Long
    
    Let Fixed = False
    If Err.Number = EX_NOHOOK Then
        
        If Me.NewRecord Then
            AccnScanListView.View FileName:=File
            Let oAccnScan.SheetType = "ACCN"
            DoFillOutAndSave Scan:=oAccnScan, OK:=Fixed, ErrCode:=ErrCode
            
            If Not Fixed Then
                If ErrCode = EX_DUPLICATE_KEY_VALUE Then
                    Me.Undo
                    DoAdvanceToNextRecord Record:=oAccnScan.ACCN
                    Let Fixed = True
                End If
            End If
        End If
    End If
    
    If Fixed Then
        Resume AttachToRecordTry
    Else
        Me!ACCN.BackColor = COLOR_MARKEDERROR
        Let Halt = True
        Resume AttachToRecordFinally
    End If

AttachToRecordFinally:

    Exit Sub

End Sub

Private Sub DoAdvanceToNextRecord(Optional ByVal Record As String)
    Me.FilterOn = False
    If Len(Record) > 0 Then
        DoCmd.SearchForRecord Record:=acFirst, WhereCondition:="ACCN='" & Record & "'"
    Else
        DoCmd.GoToRecord Record:=acNewRec
    End If
End Sub

Private Sub DoFillOutAndSave(ByRef Scan As cAccnScan, ByRef OK As Boolean, ByRef ErrCode As Long)
    Dim I As Integer
        
    Dim sDetails As String
    Dim rDetails As AccessionDetails
    
    FillOutFormFromAttachment Scan.Url
    If Len(Nz(Me!ACCN.value)) = 0 Then
        Let sDetails = InputBox("Accession Number")
        Let rDetails = GetDetails(File:=Scan.Url, s:=sDetails)
    Else
        Let rDetails.ACCN = Nz(Me!ACCN.value)
        Let rDetails.Creator = Nz(Me!Creator.value)
        Let rDetails.Date = Me!AccnDate.value
    End If

    'Innocent until proven guilty
    Let OK = True
    Let ErrCode = 0
    
    If Len(rDetails.ACCN) > 0 Then
        If Len(rDetails.Creator) = 0 Then
            Let rDetails.Creator = InputBox("Creator")
        End If
        
        Let Me!ACCN.value = rDetails.ACCN
        Let Me!Creator.value = rDetails.Creator
                        
        If rDetails.Date > 0 Then
            Let Me!AccnDate.value = rDetails.Date
        End If

        Let Me!AddedToDatabaseBy.value = getUserName
        Let Me!AddedToDatabaseOn.value = Now
        
        Let Me!CabinetFolder.value = Scan.FileFolder
        If Len(Nz(Me!CabinetFolder.value)) = 0 Then
            Let Me!CabinetFolder.value = sLastCabinetFolderLabel
        End If
        
        On Error GoTo cmdSaveRecordCatch
        DoCmd.RunCommand acCmdSaveRecord
        On Error GoTo 0
        
        Let Scan.ACCN = rDetails.ACCN
        'Let Scan.Creator = rDetails.Creator
        'Let Scan.Timestamp = rDetails.Date
    Else
        'Proven guilty: no ACCN
        Let OK = False
    End If
    Exit Sub
    
cmdSaveRecordCatch:
    'Proven guilty: error during record save
    Let ErrCode = Err.Number
    Let OK = False
    Resume Next

End Sub

Private Function GetDetails(ByVal File As String, Optional ByVal s As String) As AccessionDetails
    Dim I As Integer
    Dim aParts() As String
    
    Dim oAx As New cAccession
    Dim oAccnScan As cAccnScan
    
    Dim sACCN As String
    Dim sCreator As String
    Dim sDate As String
    
    Dim rDetails As AccessionDetails
    
    Let aParts = Split(s, "/")
    
    If UBound(aParts) > LBound(aParts) Then
        Let I = LBound(aParts)
        Let oAx.ACCN = aParts(I)
                
        Let I = LBound(aParts) + 1
        If CLng(oAx.Minor) = 0 Then
            Let oAx.ACCN = aParts(I) & "." & oAx.Major
        End If
        Let rDetails.ACCN = oAx.ACCN

        'Get the YYYY from the major part of the Accession number
        Let sDate = Format(oAx.Major, "0000")
                
        'Get remaining date elements from the string
        If aParts(I) <> sDate Then
            If RegexMatch(aParts(I), "^[0-9]{4}$") Then
                If CInt(aParts(I)) <= 1231 Then
                    Let aParts(I) = Left(aParts(I), 2) & "/" & Right(aParts(I), 2)
                End If
            End If
            Let sDate = sDate & "/" & aParts(I)
        End If
                
        For I = LBound(aParts) + 2 To UBound(aParts)
            Let sDate = sDate & "/" & aParts(I)
        Next I
        
        Let rDetails.Date = CDate(sDate)
    ElseIf Len(s) > 0 Then
        Let oAx.ACCN = s
        Let rDetails.ACCN = oAx.ACCN
    Else
        'NOOP
    End If
    
    If Len(rDetails.ACCN) > 0 Then
        'Try to extract Creator from the file system
        'information, e.g. from the path
        Set oAccnScan = New cAccnScan: With oAccnScan
            .ACCN = rDetails.ACCN
            .Url = File
        End With
        rDetails.Creator = oAccnScan.Creator
    End If
    
    Let GetDetails = rDetails
End Function

Private Sub cmdBulkAdd_Click()
    Dim iResult As Integer
    iResult = BulkAddDialog
End Sub

Private Sub cmbCreatorFilter_Change()
    'NOOP
End Sub

Private Sub cmdCreatorFilter_Click()
    If Not IsNull(cmbCreatorFilter.value) Then
        If cmbCreatorFilter.value <> "-" Then
            Me.Filter = "Creator = '" & cmbCreatorFilter.value & "'"
            Me.FilterOn = True
        Else
            Me.FilterOn = False
        End If
    Else
        Me.FilterOn = False
    End If
End Sub

Private Sub cmdToDo_Click()
    Me.Filter = "(SeriesTitle IS NULL) OR (LEN(SeriesTitle) = 0)"
    Me.FilterOn = True
    Me!cmbCreatorFilter.value = "-"
End Sub

Private Sub cmdUnfilledCreators_Click()
    ClearCurNames
    FillCurNames
End Sub

Private Sub cmdDeleteScan_Click()
    If Not IsNull(Me!ACCN.value) Then
        Let AccnScanListView.ACCN = Me!ACCN.value
    Else
        Let AccnScanListView.ACCN = ""
    End If
    
    AccnScanListView.DetachSelectedFromRecord
End Sub

Private Sub ACCN_Change()
    Me!ACCN.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Creator_Change()
    Update_txtCreatorName_from Me!Creator.Text
End Sub

Private Sub Form_AfterUpdate()
    If Len(Nz(CabinetFolder.value)) > 0 Then
        Let sLastCabinetFolderLabel = Nz(CabinetFolder.value)
        Let sLastCabinetFolderPath = Nz(CabinetFolder.Column(3))
    End If
End Sub

Private Sub Form_Current()
    
    Me!ACCN.BackColor = RGB(255, 255, 255)

    Dim sCreator As String
    
    If IsNull(Me!Creator.value) Then
        sCreator = ""
    Else
        sCreator = Me!Creator.value
    End If
    Update_txtCreatorName_from sCreator, Me.NewRecord
    
    AccnScanListView.ReloadPDF
    
    Do
        DoEvents
    Loop Until Me!PreviewBrowser.ReadyState >= 1
    
    Me!SeriesTitle.SetFocus
    
    Dim vAccnScan As Variant
    
    AccnScanListView.ClearList
    If Not Me.NewRecord Then
        AccnScanListView.Fill getAccnScanAttachments(Me!ACCN)
    End If
    
End Sub

Private Sub cmdFillOut_Click()
    FillOutFormFromAttachmentControls
End Sub

Private Sub Form_Open(Cancel As Integer)
    'Call sEnableDrop(Me)
    'Call sHook(Me.Hwnd, "sDragDrop")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call sUnhook(Me.Hwnd)
End Sub

Private Sub lstAccnScans_AfterUpdate()
    AccnScanListView.ReloadPDFFromList
End Sub

Private Sub lstCurNames_DblClick(Cancel As Integer)
    Dim sItem As String
    Dim aWords() As String
    Dim iResult As Integer
    
    sItem = lstCurNames.ItemData(lstCurNames.ListIndex)
    
    aWords = Split(sItem, " => ", 2)
    
    iResult = BulkAddDialog(aWords(1) & "\ControlFiles\")
    If iResult = 0 Then
        SecureCreatorRecord Creator:=aWords(0)
    End If

    ClearCurNames
    FillCurNames
End Sub

Private Sub Scan_AfterUpdate()
    Me!cmdFillOut.Enabled = True
End Sub

'**
'* BulkAddDialog
'*
'* @param String Path The initial file path, if any, to start the dialog in (the user may jump out to a different directory if they see fit)
'* @return Integer a code representing the button the user clicked to close the Dialog.
'**
Private Function BulkAddDialog(Optional ByVal Path As String) As Integer
    Dim f    As Object
    Dim iResult As Integer
    Dim sItem As Variant
    Dim asPath() As String
    Dim oAccnScan As cAccnScan
    Dim bHalt As Boolean
    Dim sDestFolder As String
    
    Me!ACCN.BackColor = RGB(255, 255, 255)
    
    Set f = Application.FileDialog(3)
    With f
        .AllowMultiSelect = True
        .ButtonName = "Bulk Add"
        If Len(Path) > 0 Then
            .InitialFileName = Path
        End If
        iResult = .Show
    End With
    BulkAddDialog = iResult
    
    If iResult = 0 Then
        Exit Function
    End If
    
    boxProgress.Width = 10
    For Each sItem In f.SelectedItems
        boxProgress.Visible = True
        
        DoCmd.GoToRecord Record:=acNewRec
        DoAttachScanFile File:=Nz(sItem), Halt:=bHalt, DestinationFolder:=sDestFolder
        
        If bHalt Then
            Exit For
        End If
        
        boxProgress.Width = boxProgress.Width + 10
        
    Next sItem
    boxProgress.Visible = False
    Me.Requery
End Function

'**
'* FillOutFormAttachment
'*
'* @param String AttachPath The name of the attachment file to extract Accession data from
'* @return Boolean True if the filename supplied useful data for filling out the form, False otherwise
'**
Private Function FillOutFormFromAttachment(ByVal AttachPath As String) As Boolean
    Dim cRefs As Variant
    Dim oAccnScan As New cAccnScan
    Dim hasData As Boolean
    Dim oRef As Variant
    
    hasData = False 'So far
    Let oAccnScan.Url = AttachPath
    
    'Get the ACCN #
    Set cRefs = regexpIsAccnScanFile.Execute(oAccnScan.FileName)
    For Each oRef In cRefs
        hasData = True
        Me!ACCN.value = oRef.SubMatches(1) & "." & oRef.SubMatches(2)
        Me!AccnDate.value = DateSerial(Val(oRef.SubMatches(3)), Val(oRef.SubMatches(4)), Val(oRef.SubMatches(5)))
        Me!Creator.value = oRef.SubMatches(0)
    Next oRef
    
    If hasData Then
        Me!AddedToDatabaseBy.value = "cwj"
        Me!AddedToDatabaseOn.value = Now
    End If

    Let FillOutFormFromAttachment = hasData
End Function

'**
'* regexpIsAccnScanFile: get a RegExp object for testing whether a given string (filename) looks like the stereotypical pattern for an Accession sheet scan file
'* (for instance SPF_1998_0189_19990311.PDF for Accession 1998.0189, dated 1999-03-11, Creator CurName SPF)
'*
'* @return Object A newly instantiated RegExp object pre-programmed with the correct pattern for testing and extracting data from filenames
'**
Private Function regexpIsAccnScanFile() As Object
    Dim oIsAccnFile As New RegExp
    
    With oIsAccnFile
        .IgnoreCase = True
        .Pattern = "^([A-Z0-9]+)_([0-9]+)_([0-9]+)_([0-9]{4})([0-9]{2})?([0-9]{2})?.*[.]PDF(|.*)$"
    End With

    Set regexpIsAccnScanFile = oIsAccnFile
End Function

Private Sub FillOutFormFromAttachmentControls()
    Dim cAttachPaths As New Collection
    Dim hasData As Boolean
    Dim cRefs As Variant
    Dim vAttachPath As Variant
    
   'SOURCE CONTROL 1: Me!Scan, an Attachment field control
   If Me!Scan.AttachmentCount > 0 Then
        Me!Scan.CurrentAttachment = 0
        For I = 0 To (Me!Scan.AttachmentCount - 1)
            If regexpIsAccnScanFile.Test(Me!ScanFileName) Then
                cAttachPaths.Add Me!ScanFileName
            End If
            Me!Scan.Forward
        Next I
    End If
    
    'SOURCE CONTROL 2: Me!lstAccnScans, via AccnScanListView object
    If AccnScanListView.HasScans Then
        For Each vAttachPath In AccnScanListView.Scans
            If regexpIsAccnScanFile.Test(Me!ScanFileName) Then
                cAttachPaths.Add vAttachPath
            End If
        Next vAttachPath
    End If
    
    For Each vAttachPath In cAttachPaths
        hasData = FillOutFormFromAttachment(vAttachPath)
        If hasData Then
            Exit For
        End If
    Next vAttachPath
    
    If hasData Then
        AccnScanListView.ReloadPDF
        Me!cmdFillOut.Enabled = False
    End If
    
End Sub

'**
'* ClearCurNames: clear all items form the utility list of not-yet-processed CurNames
'* The list needs to get cleared out and refreshed sometimes (for example after a Bulk Add)
'* but VBA doesn't seem to give us an easy way to do this, so instead we'll loop through and
'* delete each item in the list until we have an empty list.
'**
Private Sub ClearCurNames()
    'Clear the CurNames list
    Do While Me!lstCurNames.ListCount > 0
        Me!lstCurNames.RemoveItem 0
    Loop
End Sub

'**
'* FillCurNames: Fill out a set of not-yet-processed CurNames in the utility list box
'* Scan through our storage location for State and for Local agencies. For each directory
'* we find, check whether or not that agency's CurName appears in Creators database table
'* If so, presumably we've filled in at least one directory's worth of that agency's scan
'* files. If not, probably not. This is highly imperfect as a metric, but should give a
'* rough and ready picture of how much processing or backfilling there is left to do.
'**
Private Sub FillCurNames()
    Dim sDrive As String
    Dim sPattern As String
    Dim aScanDir(1 To 2) As String
    Dim aWords() As String
    
    sDrive = getDefaultDrive
    Dim dCreators As New Dictionary
    Dim vCreator As Variant
    
    aScanDir(1) = "State"
    aScanDir(2) = "Local"
    For I = 1 To 1
        sPattern = "*_*"
        sDirPrefix = "\CollectionsManagement\AgencyFiles\" & aScanDir(I) & "\"
            
        f = Dir(sDrive & sDirPrefix & sPattern, vbDirectory)
        Do While Len(f) > 0
            aWords = Split(f, "_", 2)
            If UBound(aWords) > 0 Then
                vCreator = aWords(0)
                If Not dCreators.Exists(vCreator) Then
                    dCreators.Item(vCreator) = sDrive & sDirPrefix & f
                End If
            End If
            f = Dir()
        Loop
    Next I

    Dim dCurNamesNoted As Dictionary
    Set dCurNamesNoted = CurNamesInCreators
    
    For Each vCreator In dCreators.Keys()
        If Not dCurNamesNoted.Exists(vCreator) Then
            lstCurNames.AddItem vCreator & " => " & dCreators.Item(vCreator)
        End If
    Next vCreator
End Sub

'Code clipped from ReloadPDF:
'---
'
'        If Dir(FilePath) <> "" Then
'            Kill FilePath
'        End If
'
'        hasPDF = True
'        If Me.NewRecord Then
'            iFileNum = FreeFile
'            Open FilePath For Binary As #iFileNum
'            Put #iFileNum, 1, AttachmentControl.FileData
'            Close #iFileNum
'        Else
'            Set RecAtt = FormRS.Fields("Scan").Value
'            RecAtt.OpenRecordset
'            RecAtt.Fields("FileData").SaveToFile FilePath
'        End If

