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
    
            Let asFileFolderSlugs(1) = oAccnScan.FileFolderSlug
    
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
                    Let asFileFolderSlugs(2) = oAccnScan.FileFolderSlug
                End If
            End If
            
            Let oAccnScan.OldPath = asFileFolderSlugs(1)
            
            Let View.ACCN = Nz(Me!ACCN.value)
            
            View.AttachToRecord Location:=oAccnScan
        Next vItem
        
        View.ReloadPDFFromForm
    End If
    Exit Sub
    
End Sub

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
End Sub

Private Sub Form_Open(Cancel As Integer)
    'NOOP
End Sub
