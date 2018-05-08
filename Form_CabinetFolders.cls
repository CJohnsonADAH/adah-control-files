VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CabinetFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdFilePathPicker_Click()
    Dim sDestinationFolder As String
    
    'Get a path
    Let sDestinationFolder = DestinationFolderDialog
    
    'Take the slug from the end
    Let FilePath.Value = FileFolderSlug(sDestinationFolder)
End Sub

Private Sub cmdFileSourcePathPicker_Click()
    Dim sDestinationFolder As String
    
    'Get a path
    Let sDestinationFolder = DestinationFolderDialog
    
    'Take the slug from the end
    Let FileSourcePath.Value = FileFolderSlug(sDestinationFolder)
    
End Sub

Private Sub cmdFolderPathPicker_Click()
    Dim sDestinationFolder As String
    Dim NamingConvention As New CScanFileNamingConvention
    
    'Get a path
    Let sDestinationFolder = DestinationFolderDialog
    
    'Feed a hypothetical scan file name into NamingConvention
    If Len(Dir(sDestinationFolder & "\ControlFile\.")) > 0 Then
        Let sDestinationFolder = sDestinationFolder & "\ControlFile"
    ElseIf Len(Dir(sDestinationFolder & "\ControlFiles\.")) > 0 Then
        Let sDestinationFolder = sDestinationFolder & "\ControlFiles"
    End If
    
    Let NamingConvention.Url = sDestinationFolder & "\copier.pdf"
    
    Let FolderPath.Value = NamingConvention.RelativePath
End Sub

Private Sub cmdSign_Click()
    If Len(Nz(AddedToDatabaseBy.Value)) = 0 Then
        Let AddedToDatabaseBy.Value = getUserName
    End If
    
    Let AddedToDatabaseOn.Value = Now
End Sub
