VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_AccnScans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**
'* @class Form_AccnScans: form for directly handling individual scanned sheets
'*
'* @version 2018.0427
'**
Option Compare Database
Option Explicit

Private oView As Object
Private oScan As cAccnScan
Private bInsertedAccession As Boolean

Private tDelayedSearch As cStopWatch

Private sInstanceId As String

Private Property Get View() As Object
    If oView Is Nothing Then
        Set oView = New cAccnScanListView: With oView
            Set .BrowserControl = Me!webPDF
            Set .Form = Me
        End With
    End If

    Set View = oView
End Property

Public Property Get Scan() As Object
    If oScan Is Nothing Then
        Set oScan = ScanFactory
    End If
    
    Set Scan = oScan
End Property

Private Function ScanFactory() As Object
    Dim o As New cAccnScan: With o
        If Nz(Me!ID.Value) Then
            Let .ID = Nz(Me!ID.Value)
        Else
            Let .FileName = Nz(Me!FileName.Value)
            Let .FilePath = Nz(Me!FilePath.Value)
        End If
    End With
    
    Set ScanFactory = o
End Function

Public Property Get DelayedSearch() As cStopWatch
    If tDelayedSearch Is Nothing Then
        Set tDelayedSearch = New cStopWatch
    End If
    
    Set DelayedSearch = tDelayedSearch
End Property

Private Sub cmdConvertFileName_Click()
    'Before we do anything, we should lock in pending database changes
    Let Me.Dirty = False
    
    If Len(Scan.FileName) > 0 Then
        Scan.ConvertFileName
        View.ReloadPDFFromForm
    End If
End Sub

Private Sub cmdInsert_Click()
    Dim dlgSourceFile As FileDialog
    
    Dim vItem As Variant
    Dim sItem As String
    
    Dim oAccnScan As cAccnScan
    Dim sDestinationFolder As String
    Dim asFileFolderSlugs(1 To 2) As String

    Set dlgSourceFile = Application.FileDialog(msoFileDialogFilePicker): With dlgSourceFile
        .AllowMultiSelect = False
        .Filters.Add Description:="PDF", Extensions:="*.PDF", Position:=1
        .Filters.Add Description:="All", Extensions:="*.*", Position:=2
    End With

    If dlgSourceFile.Show Then
        For Each vItem In dlgSourceFile.SelectedItems
            Let sItem = Nz(vItem)
            Set oAccnScan = New cAccnScan: With oAccnScan
                .Url = sItem
            End With
    
            If Not oAccnScan.NamingConvention.isOnDefaultDrive Then
                If Len(sDestinationFolder) = 0 Then
                    Let sDestinationFolder = oAccnScan.NamingConvention.FileFolderHomePath
                End If
                    
                If Len(sDestinationFolder) = 0 Then
                    Let sDestinationFolder = DestinationFolderDialog
                End If
                                
                If Len(sDestinationFolder) > 0 Then
                    oAccnScan.doCopyTo NewPath:=sDestinationFolder
                    Let sItem = oAccnScan.Url
                End If
            End If
            
            Let asFileFolderSlugs(1) = oAccnScan.NamingConvention.FileFolderSlug(Old:=True)
            Let asFileFolderSlugs(2) = oAccnScan.NamingConvention.FileFolderSlug(Old:=False)
            
            Let View.ACCN = Nz(Me!ACCN.Value)
            
            Let oAccnScan.NewRecord = Me.NewRecord
            View.AttachToRecord Location:=oAccnScan
        Next vItem
        
        View.ReloadPDFFromForm
        
        Me!Timestamp.SetFocus
    End If
    Exit Sub
    
End Sub

Private Sub Form_AfterUpdate()
    RefreshLabelLinks
    
    Set oScan = ScanFactory
    
    If bInsertedAccession Then
        DoEvents
        DoCmd.OpenForm FormName:="Accessions"
        Forms!Accessions.FilterOn = False
        DoCmd.SearchForRecord ObjectType:=acDataForm, ObjectName:="Accessions", Record:=acFirst, WhereCondition:="ACCN='" & Me!ACCN.Value & "'"
        Forms!Accessions.RefreshAttachmentView
    End If
    Let bInsertedAccession = False
End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    Dim sFolderID As String
    Dim iFolderID As Integer
    Dim sFolderLabel As String
    Dim sFilterCaption As String
    
    Select Case ApplyType
    Case acApplyFilter:
        Let sFilterCaption = Me.Filter
        
        Let sFolderID = RegexComponent(Value:=Me.Filter, Pattern:="^CabinetFolder=([0-9]+)$", Part:=1)
        If Len(sFolderID) > 0 Then
            Let iFolderID = Val(sFolderID)
            Let sFolderLabel = DLookup(Expr:="Label", Domain:="CabinetFolders", Criteria:="ID=" & Format(iFolderID, "0"))
            If Len(sFolderLabel) > 0 Then
                Let sFilterCaption = sFolderLabel
            End If
        End If
        
        lblFolderLabel.Caption = sFilterCaption
    
    Case acShowAllRecords:
        lblFolderLabel.Caption = "*"
        
        Let lstFolders.Value = ""
    End Select
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim N As String
    Dim Accession As cAccession
    Dim myScan As cAccnScan
    
    If Len(Nz(Me!ACCN.Value)) > 0 Then
        Set Accession = New cAccession: With Accession
            Let .ACCN = Me!ACCN.Value
        End With
        
        If Accession.ACCN <> Me!ACCN.Value Then
            Let Me!ACCN.Value = Accession.ACCN
        End If
        
        'Check to see if we have anything in the Accessions table already
        Let N = DCount("ACCN", "Accessions", "ACCN='" & Nz(Me!ACCN.Value) & "'")
        If (N = 0) And hasAccnSheet Then
            Dim YesNo As Long
            Let YesNo = MsgBox("Create an Accession record for " & Me!ACCN.Value & "?", Buttons:=vbYesNo)
            If YesNo = 6 Then '6 = Yes button, 7 = No button
            
                Set myScan = New cAccnScan: With myScan
                    .FileName = Nz(Me!FileName.Value)
                    .FilePath = Nz(Me!FilePath.Value)
                    .ACCN = Nz(Me!ACCN.Value)
                    .Timestamp = Nz(Me!Timestamp.Value)
                    .SheetType = "ACCN"
                    If Not IsNull(Me!CabinetFolder.Value) Then
                        .CabinetFolder = Me!CabinetFolder.Value
                    End If
                End With
                
                myScan.InsertIntoAccessions
                Let Me!ACCN.BackColor = COLOR_UNMARKED
                Let bInsertedAccession = True
                
            Else
            
                Me!ACCN.BackColor = COLOR_MARKEDERROR
                Me!ACCN.ControlTipText = ""
                Let Cancel = True
                
            End If
        End If
    End If
End Sub

Private Function hasAccnSheet()
    Let hasAccnSheet = (Len(Nz(Me!SheetType)) = 0 Or Nz(Me!SheetType.Value) = "ACCN")
End Function

Private Sub Form_Close()
    If Not oView Is Nothing Then
        oView.Unload
        Set oView = Nothing
    End If
End Sub

Private Sub Form_Current()
    Let View.ACCN = Nz(Me!ACCN.Value)
    View.ReloadPDFFromForm
    
    Set oScan = ScanFactory
    
    Me!ACCN.BackColor = RGB(255, 255, 255)
    
    RefreshLabelLinks
End Sub

Private Sub Form_Load()
    Me.TimerInterval = 100
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim iCancel As Integer
    
    Form_ApplyFilter Cancel:=iCancel, ApplyType:=acShowAllRecords
End Sub

Private Sub Form_Timer()
    Dim WhereCondition As String
    
    If DelayedSearch.Running Then
        DelayedSearch.StopTimer
        If DelayedSearch.ms >= 300 Then
            Let WhereCondition = ""
            If Len(txtFolders.Text) > 0 Then
                Let WhereCondition = "WHERE Creator LIKE '*" & txtFolders.Text & "*' OR Label LIKE '*" & txtFolders.Text & "*'"
            End If
            
            lstFolders.RowSource = "SELECT ID, Creator, Label FROM CabinetFolders " & WhereCondition & " ORDER BY Creator, Label, ID;"
            lstFolders.Requery
            DelayedSearch.ResetTimer
        End If
    End If
End Sub

Private Sub lblACCN_Click()
    If Len(Nz(ACCN.Value)) > 0 Then
        DoCmd.OpenForm FormName:="Accessions", WindowMode:=acWindowNormal, WhereCondition:="ACCN='" & ACCN.Value & "'"
    End If
End Sub

Private Function HasACCN() As Boolean
    Let HasACCN = (Len(Nz(ACCN.Value)) > 0)
End Function

Private Sub RefreshLabelLinks()
    If HasACCN Then
        Let lblACCN.FontUnderline = True
        Let lblACCN.FontBold = True
        Let lblACCN.ForeColor = RGB(0, 0, 255)
    Else
        Let lblACCN.FontUnderline = False
        Let lblACCN.FontBold = False
        Let lblACCN.ForeColor = RGB(75, 75, 75)
    End If
End Sub

Private Sub lblACCN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If HasACCN Then
        Let lblACCN.ForeColor = RGB(255, 0, 0)
    End If
End Sub

Private Sub lblACCN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If HasACCN Then
        Let lblACCN.ForeColor = RGB(0, 0, 255)
    End If
End Sub

Private Sub lstFolders_Click()
    Dim iCancel As Integer
    
    'First, turn off the filter
    Me.FilterOn = False
    
    'Next, set the filter
    Me.Filter = "CabinetFolder=" & Format(lstFolders.Column(0), "0")
    
    'Finally, turn the filter (back) on
    Me.FilterOn = True
    
    Form_ApplyFilter Cancel:=iCancel, ApplyType:=acApplyFilter
End Sub

Private Sub txtFolders_KeyUp(KeyCode As Integer, Shift As Integer)
    DelayedSearch.StartTimer
End Sub
