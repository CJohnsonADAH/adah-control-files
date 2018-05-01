Attribute VB_Name = "modMyVBA"
Option Compare Database
Option Explicit

' No VT_GUID available so must declare type GUID
Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As LongPtr
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr

Public Const EX_NOHOOK = (vbObjectError + 255)
Public Const EX_RENAMEFAILED = (vbObjectError + 254)
Public Const EX_ACCNSCAN_DUPLICATE = (vbObjectError + 253)

Public Const EX_FILEALREADYEXISTS = 58
Public Const EX_FILEPERMISSIONDENIED = 75
Public Const EX_DUPLICATE_KEY_VALUE = 3022

Public Const COLOR_ALERT As Long = &HFFFF           'Bright yellow
Public Const COLOR_DISABLED As Long = &HC0C0C0      'Light grey
Public Const COLOR_UNMARKED As Long = &HFFFFFF      'White
Public Const COLOR_MARKEDERROR As Long = &HC0C0FF   'Light red

Public Sub BubbleSortList(ByRef List As Variant)
    Dim Swapped As Boolean
    Dim vSwap As Variant
    Dim I As Integer, J As Integer
    
    If IsArray(List) Or TypeName(List) = "Collection" Then
        Do
            Let Swapped = False
            For I = LBound(List) To UBound(List) - 1
                If List(I + 1) < List(I) Then
                    Let vSwap = List(I)
                    Let List(I) = List(I + 1)
                    Let List(I + 1) = vSwap
                    
                    Let Swapped = True
                End If
            Next I
        Loop Until Not Swapped
    End If
End Sub

'**
'* DebugDump: Utility Function mainly for use in the Immediate pane to more easily display a
'* bunch of different kinds of objects and collections of objects in VBA
'*
'* @param Variant v The object to print out a representation of in the Immediate pane
'**
Sub DebugDump(v As Variant)
    Dim vScalar As Variant
    If IsArray(v) Or TypeName(v) = "Collection" Or TypeName(v) = "ISubMatches" Then
        If IsArray(v) Then
            Debug.Print TypeName(v), LBound(v), UBound(v)
        Else
            Debug.Print TypeName(v), v.Count
        End If
        
        For Each vScalar In v
            DebugDump (vScalar)
        Next vScalar
    ElseIf TypeName(v) = "Dictionary" Then
        Debug.Print TypeName(v)
        For Each vScalar In v.Keys
            Debug.Print vScalar & ":"
            DebugDump v.Item(vScalar)
        Next vScalar
    ElseIf TypeName(v) = "Nothing" Then
        Debug.Print TypeName(v)
    Else
        Debug.Print TypeName(v), v
    End If
End Sub

Public Function JoinCollection(delim As String, c As Collection)
    Dim first As Boolean
    Dim vItem As Variant
    Dim sConjunction As String
    
    first = True
    For Each vItem In c
        If Not first Then
            sConjunction = sConjunction & delim
        End If
        
        sConjunction = sConjunction & vItem
        
        first = False
    Next vItem
    
    JoinCollection = sConjunction
End Function

Public Function CreateGuidString()
    Dim guid As GUID_TYPE
    Dim strGuid As String
    Dim retValue As LongPtr
    Const guidLength As Long = 39 'registry GUID format with null terminator {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}

    retValue = CoCreateGuid(guid)
    If retValue = 0 Then
        strGuid = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(guid, StrPtr(strGuid), guidLength)
        If retValue = guidLength Then
            ' valid GUID as a string
            CreateGuidString = strGuid
        End If
    End If
End Function

'Derived from code posted at https://stackoverflow.com/questions/16948215/exporting-ms-access-forms-and-class-modules-recursively-to-text-files
'Modified to allow the user to set a desired path
Public Sub ExportAllCode(Optional ByVal Path As String)

    Dim c As VBComponent
    Dim Sfx As String

    Dim sDestinationFolder As String
    Dim dlgDestinationFolder As FileDialog
                    
    Set dlgDestinationFolder = Application.FileDialog(msoFileDialogFolderPicker): With dlgDestinationFolder
        .Title = "Export Destination Folder"
        .InitialFileName = IIf(Len(Path) > 0, Path, CurrentProject.Path)
    End With
                
    If dlgDestinationFolder.Show Then
        Let sDestinationFolder = dlgDestinationFolder.SelectedItems(1)

        For Each c In Application.VBE.VBProjects(1).VBComponents
            Select Case c.Type
                Case vbext_ct_ClassModule, vbext_ct_Document
                    Sfx = ".cls"
                Case vbext_ct_MSForm
                    Sfx = ".frm"
                Case vbext_ct_StdModule
                    Sfx = ".bas"
                Case Else
                    Sfx = ""
            End Select
    
            If Sfx <> "" Then
                c.Export _
                    FileName:=sDestinationFolder & "\" & _
                    c.Name & Sfx
            End If
        Next c
    End If
    
End Sub
