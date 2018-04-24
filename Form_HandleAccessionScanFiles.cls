VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_HandleAccessionScanFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private oAccnScanListView As cAccnScanListView

Private Property Get AccnScanListView() As cAccnScanListView

    If oAccnScanListView Is Nothing Then
        Set oAccnScanListView = New cAccnScanListView
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
    Let oAgency.CurName = Text

    Me!txtCreatorName.value = oAgency.Name(Short:=False)
    
    If Len(Me!txtCreatorName.value) > 0 Or NewRecord Then
        Me!txtCreatorName.BackColor = RGB(192, 192, 192) 'Light grey
    Else
        Me!txtCreatorName.BackColor = RGB(255, 255, 0) 'Yellow!
    End If
    
End Sub


Private Sub FilterNamesDialog(Optional ByVal Path As String)
    Dim f    As Object
    Dim iResult As Integer
    Dim sItem As Variant
    Dim asPath() As String
    Dim sLocator As Variant
    Dim sPrevious As String
    Dim bDrive As Boolean
    Dim sFilePath As String
    Dim sFileName As String
    Dim oAccnScan As cAccnScan
    
    Set f = Application.FileDialog(3)
    With f
        .AllowMultiSelect = True
        .ButtonName = "Bulk Add"
        If Len(Path) > 0 Then
            .InitialFileName = Path
        End If
        iResult = .Show
    End With
    
    If iResult = 0 Then
        Exit Sub
    End If
    
    Dim origPath As String
    Dim origFileName As String
    
    boxProgress.Width = 10
    For Each sItem In f.SelectedItems
        boxProgress.Visible = True
        'Should contain a full path, e.g. "S:\CollectionsManagement\AgencyFiles\Local\03_MontgomeryCounty\ControlFiles\03_1994_0451_19940912.PDF"
        Set oAccnScan = New cAccnScan
        oAccnScan.Url = sItem
        
        'Save old values before updating
        origPath = oAccnScan.Drive & oAccnScan.FilePath
        origFileName = oAccnScan.FileName
        
        oAccnScan.ForceCorrectFileName
        
        Debug.Print "OLD: ", origPath, origFileName
        Debug.Print "NEW: ", origPath, oAccnScan.FileName
        
        Name origPath & "\" & origFileName As origPath & "\" & oAccnScan.FileName
        
        'On Error GoTo HandleInsertError
        'oAccnScan.InsertIntoAccessions
        boxProgress.Width = boxProgress.Width + 10
        On Error GoTo 0
        
        GoTo ContinueLoop

HandleInsertError:
        'NOOP
        On Error GoTo 0

ContinueLoop:

    Next sItem
    boxProgress.Visible = False
    Me.Requery

End Sub

Private Sub cmdFilterNames_Click()
    FilterNamesDialog
End Sub

Private Sub Command110_Click()
    Dim vItem As Variant
    Dim sData As String
    Dim aFileParts() As String
    
    For Each vItem In Me!lstAccnScans.ItemsSelected
        sData = Me!lstAccnScans.ItemData(vItem)
        Me!lstAccnScans.RemoveItem vItem
    
        aFileParts = Split(sData, "|", 2)
        deleteAccnScanAttachments Me!ACCN.value, aFileParts(0), aFileParts(1)
    Next vItem
End Sub

Private Sub Command98_Click()
    Dim f    As Object
    Dim sItem As Variant
    Dim sLocator As Variant
    Dim sPrevious As String
    Dim sFilePath As String
    Dim sFileName As String
    
    Me!ACCN.BackColor = RGB(255, 255, 255)
    
    Set f = Application.FileDialog(3)
    f.AllowMultiSelect = True
    f.Show
        
    For Each sItem In f.SelectedItems
        If IsNull(Me!ACCN.value) Then
            FillOutFormFromAttachment sItem
            On Error GoTo HandleInsertError
            DoCmd.RunCommand acCmdSaveRecord
            On Error GoTo 0
        End If
        
        If Not IsNull(Me!ACCN.value) Then
            Let AccnScanListView.ACCN = Me!ACCN.value
        Else
            Let AccnScanListView.ACCN = ""
        End If
        AccnScanListView.AttachToRecord Location:=sItem
        
        GoTo ContinueLoop

HandleInsertError:
        On Error GoTo 0
        Me!ACCN.BackColor = RGB(255, 192, 192)
        Resume ContinueLoop
        
ContinueLoop:

    Next sItem
            
    AccnScanListView.ReloadPDF
End Sub

Private Sub ACCN_Change()
    Me!ACCN.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Creator_Change()
    Update_txtCreatorName_from Me!Creator.Text
End Sub

Private Sub Form_Current()
    'Clear the CurNames list
    Do While Me!lstCurNames.ListCount > 0
        Me!lstCurNames.RemoveItem 0
    Loop
    
    FillCurNames
End Sub

Private Sub FillCurNames()
    Dim sDrive As String
    Dim sPattern As String
    Dim aScanDir(1 To 2) As String
    Dim aWords() As String
    sDrive = "S:"
    Dim dCreators As New Dictionary
    Dim vCreator As Variant
    
    aScanDir(1) = "State"
    'aScanDir(2) = "Local"
    For i = 1 To 1
        sPattern = "*_*"
        sDirPrefix = "\CollectionsManagement\AgencyFiles\" & aScanDir(i) & "\"
            
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
    Next i

    Dim dCurNamesNoted As Dictionary
    Set dCurNamesNoted = CurNamesInCreators
    
    For Each vCreator In dCreators.Keys()
        If Not dCurNamesNoted.Exists(vCreator) Then
            lstCurNames.AddItem vCreator & " => " & dCreators.Item(vCreator)
        End If
    Next vCreator
End Sub

Private Sub cmdFillOut_Click()
    FillOutFormFromAttachmentControls
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

Private Sub Scan_AttachmentCurrent()

End Sub

