VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_NonAccessioned"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Blank_Browser()
'    Me!PreviewBrowser.Navigate2 "about:blank"
'    Do
'        DoEvents
'    Loop Until Me!PreviewBrowser.ReadyState >= 3
End Sub

Private Sub Update_txtCreatorName_from(Text As String, Optional NewRecord As Boolean)
    Dim oAgency As New cAgency
    Let oAgency.CurName = Text

    Me!txtCreatorName.Value = oAgency.Name(Short:=False)

    If Len(Me!txtCreatorName.Value) > 0 Or NewRecord Then
        Me!txtCreatorName.BackColor = RGB(192, 192, 192) 'Light grey
    Else
        Me!txtCreatorName.BackColor = RGB(255, 255, 0) 'Yellow!
    End If
End Sub

Private Sub cmdBulkAdd_Click()
'    Dim iResult As Integer
'    iResult = BulkAddDialog
End Sub

Private Function BulkAddDialog(Optional ByVal Path As String) As Integer
'    Dim f    As Object
'    Dim iResult As Integer
'    Dim sItem As Variant
'    Dim asPath() As String
'    Dim sLocator As Variant
'    Dim sPrevious As String
'    Dim bDrive As Boolean
'    Dim sFilePath As String
'    Dim sFileName As String
'    Dim oAccnScan As cAccnScan
'
'    Set f = Application.FileDialog(3)
'    With f
'        .AllowMultiSelect = True
'        .ButtonName = "Bulk Add"
'        If Len(Path) > 0 Then
'            .InitialFileName = Path
'        End If
'        iResult = .Show
'    End With
'    BulkAddDialog = iResult
'
'    If iResult = 0 Then
'        Exit Function
'    End If
'
'    boxProgress.Width = 10
'    For Each sItem In f.SelectedItems
'        boxProgress.Visible = True
'        'Should contain a full path, e.g. "S:\CollectionsManagement\AgencyFiles\Local\03_MontgomeryCounty\ControlFiles\03_1994_0451_19940912.PDF"
'        Set oAccnScan = New cAccnScan
'        oAccnScan.Url = sItem
'
'        'On Error GoTo HandleInsertError
'        oAccnScan.InsertIntoAccessions
'        boxProgress.Width = boxProgress.Width + 10
'        On Error GoTo 0
'
'        GoTo ContinueLoop
'
HandleInsertError:
'        'NOOP
'        On Error GoTo 0

ContinueLoop:
'
'    Next sItem
'    boxProgress.Visible = False
'    Me.Requery
'
'    DoCmd.GoToControl "ACCN"
'    DoCmd.FindRecord FindWhat:=oAccnScan.ACCN, Match:=acEntire, MatchCase:=False, OnlyCurrentField:=acCurrent
End Function

Private Sub cmdUnfilledCreators_Click()
'    ClearCurNames
'    FillCurNames
End Sub

Private Sub Command110_Click()
'    Dim vItem As Variant
'    Dim sData As String
'    Dim aFileParts() As String
'
'    For Each vItem In Me!lstAccnScans.ItemsSelected
'        sData = Me!lstAccnScans.ItemData(vItem)
'        Me!lstAccnScans.RemoveItem vItem
'
'        aFileParts = Split(sData, "|", 2)
'        deleteAccnScanAttachments Me!ACCN.Value, aFileParts(0), aFileParts(1)
'    Next vItem
End Sub

Private Sub Command98_Click()
'    Dim f    As Object
'    Dim sItem As Variant
'    Dim asPath() As String
'    Dim sLocator As Variant
'    Dim sPrevious As String
'    Dim bDrive As Boolean
'    Dim sFilePath As String
'    Dim sFileName As String
'
'    Me!ACCN.BackColor = RGB(255, 255, 255)
'
'    Set f = Application.FileDialog(3)
'    f.AllowMultiSelect = True
'    f.Show
'
'    For Each sItem In f.SelectedItems
'        'Should contain a full path, e.g. "S:\CollectionsManagement\AgencyFiles\Local\03_MontgomeryCounty\ControlFiles\03_1994_0451_19940912.PDF"
'
'        'Split path into array locators, e.g. {"S:", "CollectionsManagement", "AgencyFiles", "Local", "03_MontgomeryCounty", "ControlFiles", "03_1994_0451_19940912.PDF"}
'        asPath = Split(sItem, "\")
'
'        bDrive = True: sPrevious = ""
'        sFileName = "": sFilePath = ""
'        For Each sLocator In asPath
'            If bDrive Then
'                bDrive = False
'                sPrevious = ""
'            Else
'                sFilePath = sFilePath & sPrevious
'                sFileName = sLocator
'                sPrevious = "\" & sFileName
'            End If
'        Next sLocator
'
'        'INSERT INTO AccnScan VALUES (Me!ACCN.Value, sFilePath, sFileName)
'        Me!lstAccnScans.AddItem sFileName & "|" & sFilePath
'        If IsNull(Me!ACCN.Value) Or Len(Me!ACCN.Value) = 0 Then
'            cmdFillOut_Click
'            On Error GoTo HandleInsertError
'            DoCmd.RunCommand acCmdSaveRecord
'            On Error GoTo 0
'        End If
'
'        If Len(Me!ACCN.Value) > 0 Then
'            putAccnScanAttachments Me!ACCN.Value, sFileName, sFilePath
'        End If
'
'        GoTo ContinueLoop
'
HandleInsertError:
'        On Error GoTo 0
'        Me!ACCN.BackColor = RGB(255, 192, 192)
'        Me!lstAccnScans.RemoveItem Me!lstAccnScans.ListCount - 1
'        'NOOP

ContinueLoop:

'    Next sItem
    
'    ReloadPDF
End Sub

Private Sub Creator_Change()
'    Update_txtCreatorName_from Me!Creator.Text
End Sub

Private Sub Form_Current()
'
'    Dim sCreator As String
'
'    If IsNull(Me!Creator.Value) Then
'        sCreator = ""
'    Else
'        sCreator = Me!Creator.Value
'    End If
'    Update_txtCreatorName_from sCreator, Me.NewRecord
'
'    ReloadPDF
'
'    Do
'        DoEvents
'    Loop Until Me!PreviewBrowser.ReadyState >= 1
'
'    Me!SeriesTitle.SetFocus
'
'    Dim vAccnScan As Variant
'
'    'Clear the AccnScans list.
'    Do While Me!lstAccnScans.ListCount > 0
'        DoEvents
'        Me!lstAccnScans.RemoveItem 0
'    Loop
'
'    If Not Me.NewRecord Then
'        'Rebuild it: better, stronger, faster
'        For Each vAccnScan In getAccnScanAttachments(Me!ACCN)
'            Me!lstAccnScans.AddItem vAccnScan(2) & "|" & vAccnScan(1)
'        Next vAccnScan
'    End If
    
End Sub

Private Sub ClearCurNames()
'    'Clear the CurNames list
'    Do While Me!lstCurNames.ListCount > 0
'        Me!lstCurNames.RemoveItem 0
'    Loop
End Sub

Private Sub FillCurNames()
'    Dim sDrive As String
'    Dim sPattern As String
'    Dim aScanDir(1 To 2) As String
'    Dim aWords() As String
'    sDrive = "\\ADAHFS1\GR-Collections"
'    Dim dCreators As New Dictionary
'    Dim vCreator As Variant
'
'    aScanDir(1) = "State"
'    aScanDir(2) = "Local"
'    For i = 1 To 1
'        sPattern = "*_*"
'        sDirPrefix = "\CollectionsManagement\AgencyFiles\" & aScanDir(i) & "\"
'
'        f = Dir(sDrive & sDirPrefix & sPattern, vbDirectory)
'        Do While Len(f) > 0
'            aWords = Split(f, "_", 2)
'            If UBound(aWords) > 0 Then
'                vCreator = aWords(0)
'                If Not dCreators.Exists(vCreator) Then
'                    dCreators.Item(vCreator) = sDrive & sDirPrefix & f
'                End If
'            End If
'            f = Dir()
'        Loop
'    Next i
'
'    Dim dCurNamesNoted As Dictionary
'    Set dCurNamesNoted = CurNamesInCreators
'
'    For Each vCreator In dCreators.Keys()
'        If Not dCurNamesNoted.Exists(vCreator) Then
'            lstCurNames.AddItem vCreator & " => " & dCreators.Item(vCreator)
'        End If
'    Next vCreator
End Sub

Private Sub ReloadPDF(Optional ByVal selectedFileName As String, Optional ByVal selectedFilePath As String)
'    On Error GoTo ExitSub
'    Dim FormRS As DAO.Recordset
'    Set FormRS = Me.Recordset
'
'    Dim hasPDF As Boolean
'
'    Dim RecAtt As DAO.Recordset
'
'    Dim cAttachmentPaths As Collection
'
'    Dim iFileNum As Integer
'    Dim bFileData() As Byte
'    Dim fs As Object
'
'    Blank_Browser
'    hasPDF = False
'
'    Dim FilePath As String
'    Dim vFilePath As Variant
'    Dim sFilePath As String
'    Dim sSelectedFilePath As String
'
'    Set cAttachmentPaths = getAccnScanAttachments(Me!ACCN)
'
'    'Use AccnScans table records to get FilePath and FileName
'    If cAttachmentPaths.Count > 0 Then
'        For Each vFilePath In cAttachmentPaths
'            sFilePath = getScanFileName(vFilePath(2), vFilePath(1))
'            sSelectedFilePath = getScanFileName(selectedFileName, selectedFilePath)
'
'            If (Len(selectedFileName) + Len(selectedFilePath) = 0) Or (sFilePath = sSelectedFilePath) Then
'                altFilePath = getScanFileName(vFilePath(2), vFilePath(1))
'                FilePath = Environ("APPDATA") & "\" & vFilePath(2)
'            End If
'        Next vFilePath
'
'    'Or check out the Me!Scan attachment field in Accessions
'    ElseIf (Me!Scan.AttachmentCount > 0) Then
'        Me!Scan.CurrentAttachment = 0
'
'        FilePath = Environ("APPDATA") & "\" & Me!Scan.FileName
'        altFilePath = getScanFileName(Me!Scan.FileName)
'    Else
'        FilePath = "about:blank"
'    End If
'
'    If Len(altFilePath) > 0 And Dir(altFilePath) <> "" Then
'        FilePath = altFilePath
'        hasPDF = True
'    Else
'        If Dir(FilePath) <> "" Then
'            Kill FilePath
'        End If
'
'        hasPDF = True
'        If Me.NewRecord Then
'            iFileNum = FreeFile
'            Open FilePath For Binary As #iFileNum
'            Put #iFileNum, 1, Me!Scan.FileData
'            Close #iFileNum
'        Else
'            Set RecAtt = FormRS.Fields("Scan").Value
'            RecAtt.OpenRecordset
'            RecAtt.Fields("FileData").SaveToFile FilePath
'        End If
'    End If
'
'    Me!PreviewBrowser.Navigate2 FilePath

ExitSub:

'    If Not RecAtt Is Nothing Then
'        RecAtt.Close
'        Set RecAtt = Nothing
'    End If
    
'    If Not hasPDF Then
'        Blank_Browser
'    End If

End Sub

Private Sub cmdFillOut_Click()
'    Dim cAttachPaths As New Collection
'    Dim oIsAccnFile As New RegExp
'    Dim cRefs As Variant
'    Dim vAttachPath As Variant
'    Dim hasData As Boolean
'
'    With oIsAccnFile
'        .IgnoreCase = True
'        .Pattern = "^([A-Z0-9]+)_([0-9]+)_([0-9]+)_([0-9]{4})([0-9]{2})?([0-9]{2})?.*[.]PDF(|.*)$"
'    End With
'
'    If Me!Scan.AttachmentCount > 0 Then
'        Me!Scan.CurrentAttachment = 0
'        For i = 0 To (Me!Scan.AttachmentCount - 1)
'            cAttachPaths.Add Me!ScanFileName
'            Me!Scan.Forward
'        Next i
'    End If
'
'    If Me!lstAccnScans.ListCount > 0 Then
'        For i = 0 To (Me!lstAccnScans.ListCount - 1)
'            cAttachPaths.Add Me!lstAccnScans.ItemData(i)
'        Next i
'    End If
'
'    For Each vAttachPath In cAttachPaths
'        'Get the ACCN #
'        Set cRefs = oIsAccnFile.Execute(vAttachPath)
'        For Each oRef In cRefs
'            hasData = True
'            Me!ACCN.Value = oRef.SubMatches(1) & "." & oRef.SubMatches(2)
'            Me!AccnDate.Value = DateSerial(Val(oRef.SubMatches(3)), Val(oRef.SubMatches(4)), Val(oRef.SubMatches(5)))
'            Me!Creator.Value = oRef.SubMatches(0)
'        Next oRef
'    Next vAttachPath
'
'    If hasData Then
'        Me!AddedToDatabaseBy.Value = "cwj"
'        Me!AddedToDatabaseOn.Value = Now
'        ReloadPDF
'    End If
'
'    Me!cmdFillOut.Enabled = False
End Sub

Private Sub lstAccnScans_AfterUpdate()
'    Dim aFileParts() As String
'
'    aFileParts = Split(Me!lstAccnScans.Value, "|", 2)
'    ReloadPDF aFileParts(0), aFileParts(1)
End Sub

Private Sub lstCurNames_DblClick(Cancel As Integer)
'    Dim sItem As String
'    Dim aWords() As String
'    Dim iResult As Integer
'
'    sItem = lstCurNames.ItemData(lstCurNames.ListIndex)
'
'    aWords = Split(sItem, " => ", 2)
'
'    iResult = BulkAddDialog(aWords(1) & "\ControlFiles\")
'    If iResult = 0 Then
'        SecureCreatorRecord Creator:=aWords(0)
'    End If
'
'    ClearCurNames
'    FillCurNames
End Sub

Private Sub Scan_AfterUpdate()
'    Me!cmdFillOut.Enabled = True
End Sub

