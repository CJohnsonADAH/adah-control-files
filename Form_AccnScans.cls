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
Private bInsertedAccession As Boolean

Private Property Get View() As Object
    If oView Is Nothing Then
        Set oView = New cAccnScanListView: With oView
            Set .BrowserControl = Me!webPDF
            Set .Form = Me
        End With
    End If

    Set View = oView
End Property

Private Sub cmdInsert_Click()
    Dim dlgSourceFile As FileDialog
    Dim dlgDestinationFolder As FileDialog
    
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
    
            If Not oAccnScan.isOnDefaultDrive Then
                If Len(sDestinationFolder) = 0 Then
                    Let sDestinationFolder = oAccnScan.FolderHomePath
                End If
                    
                If Len(sDestinationFolder) = 0 Then
                    Set dlgDestinationFolder = Application.FileDialog(msoFileDialogFolderPicker): With dlgDestinationFolder
                        .Title = "Destination Folder"
                    End With
                
                    If dlgDestinationFolder.Show Then
                        Let sDestinationFolder = dlgDestinationFolder.SelectedItems(1)
                    End If
                End If
                                
                If Len(sDestinationFolder) > 0 Then
                    oAccnScan.doCopyTo NewPath:=sDestinationFolder
                    Let sItem = oAccnScan.Url
                End If
            End If
            
            Let asFileFolderSlugs(1) = oAccnScan.FileFolderSlug(Old:=True)
            Let asFileFolderSlugs(2) = oAccnScan.FileFolderSlug(Old:=False)
            
            Let View.ACCN = Nz(Me!ACCN.value)
            
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
    
    If bInsertedAccession Then
        DoEvents
        DoCmd.OpenForm FormName:="Accessions"
        Forms!Accessions.FilterOn = False
        DoCmd.SearchForRecord ObjectType:=acDataForm, ObjectName:="Accessions", Record:=acFirst, WhereCondition:="ACCN='" & Me!ACCN.value & "'"
        Forms!Accessions.RefreshAttachmentView
    End If
    Let bInsertedAccession = False
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim N As String
    Dim Scan As cAccnScan
    
    If Me.NewRecord And Len(Nz(Me!ACCN.value)) > 0 Then
        'Check to see if we have anything in the Accessions table already
        Let N = DCount("ACCN", "Accessions", "ACCN='" & Nz(Me!ACCN.value) & "'")
        If (N = 0) And hasAccnSheet Then
            Dim YesNo As Long
            Let YesNo = MsgBox("Create an Accession record for " & Me!ACCN.value & "?", Buttons:=vbYesNo)
            If YesNo = 6 Then '6 = Yes button, 7 = No button
            
                Set Scan = New cAccnScan: With Scan
                    .FileName = Nz(Me!FileName.value)
                    .FilePath = Nz(Me!FilePath.value)
                    .ACCN = Nz(Me!ACCN.value)
                    .Timestamp = Nz(Me!Timestamp.value)
                    .SheetType = "ACCN"
                    If Not IsNull(Me!CabinetFolder.value) Then
                        .CabinetFolder = Me!CabinetFolder.value
                    End If
                End With
                
                Scan.InsertIntoAccessions
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
    Let hasAccnSheet = (Len(Nz(Me!SheetType)) = 0 Or Nz(Me!SheetType.value) = "ACCN")
End Function

Private Sub Form_Close()
    If Not oView Is Nothing Then
        Set oView.BrowserControl = Nothing
        Set oView.Form = Nothing
        Set oView = Nothing
    End If
End Sub

Private Sub Form_Current()
    Let View.ACCN = Nz(Me!ACCN.value)
    View.ReloadPDFFromForm
    
    Me!ACCN.BackColor = RGB(255, 255, 255)
    
    RefreshLabelLinks
End Sub

Private Sub Form_Open(Cancel As Integer)
    'NOOP
End Sub

Private Sub lblACCN_Click()
    If Len(Nz(ACCN.value)) > 0 Then
        DoCmd.OpenForm FormName:="Accessions", WindowMode:=acWindowNormal, WhereCondition:="ACCN='" & ACCN.value & "'"
    End If
End Sub

Private Function HasACCN() As Boolean
    Let HasACCN = (Len(Nz(ACCN.value)) > 0)
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
