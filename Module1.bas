Attribute VB_Name = "Module1"
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
