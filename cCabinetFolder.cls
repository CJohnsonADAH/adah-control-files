VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCabinetFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private iID As Integer
Private sCreator As String
Private sLabel As String
Private sFilePath As String
Private sAddedToDatabaseBy As String
Private sAddedToDatabaseOn As Date

Public Property Let ID(iI As Integer)
    
    'Numeric ID provides a unique record index
    FillOutFromRecordSet ("ID=" & iI)

    iID = iI
    
End Property

Public Property Get ID() As Integer
    ID = iID
End Property

Public Property Get Creator() As String
    Creator = sCreator
End Property

Public Property Get Label() As String
    Label = sLabel
End Property

Public Property Get FilePath() As String
    FilePath = sFilePath
End Property

Public Property Let FilePath(sF As String)
    FillOutFromRecordSet ("FilePath='" & Replace(Replace(sF, "'", "''"), ";", "") & "'")
    
    sFilePath = sF
End Property

Public Property Get AddedToDatabaseBy() As String
    AddedToDatabaseBy = sAddedToDatabaseBy
End Property

Public Property Get AddedToDatabaseOn() As Date
    AddedToDatabaseOn = sAddedToDatabaseOn
End Property

Public Sub Insert()
    Dim rsCabinetFolders As DAO.Recordset
    Dim sDefault As String
    Dim aWords() As String
    
    If (ID = 0) Then
        Set rsCabinetFolders = CurrentDb.OpenRecordset("CabinetFolders")
        rsCabinetFolders.AddNew
        
        If Len(Creator) = 0 Then
            If Len(FilePath) > 0 Then
                aWords = Split(FilePath, "_", 2)
                sDefault = aWords(0)
            Else
                sDefault = ""
            End If
            sCreator = InputBox("CurName:", "Folder Creator", sDefault)
        End If
        rsCabinetFolders!Creator = Creator
        
        If Len(Label) = 0 Then
            If Len(FilePath) > 0 Then
                sDefault = JoinCollection(" ", FilterFilePath(camelCaseSplitString(FilePath))) & " Control Files"
            Else
                sDefault = ""
            End If
            sLabel = InputBox("Label:", "Folder Label", sDefault)
        End If
        rsCabinetFolders!Label = Label
        
        If Len(FilePath) = 0 Then
            sFilePath = InputBox("FilePath?", "File Path")
        End If
        rsCabinetFolders!FilePath = FilePath
        
        'These are automatically generated
        'rsCabinetFolders!ID = ID
        rsCabinetFolders!AddedToDatabaseBy = "cwj"
        rsCabinetFolders!AddedToDatabaseOn = Now
        
        rsCabinetFolders.Update
        
        Set rsCabinetFolders = Nothing
        
        Let FilePath = sFilePath
    End If
End Sub

Private Sub FillOutFromRecordSet(sqlWhere As String)
    Dim oRS As DAO.Recordset
    
    Set oRS = CurrentDb.OpenRecordset("SELECT * FROM CabinetFolders WHERE " & sqlWhere)
    
    If Not oRS.EOF Then
        Let iID = oRS!ID
        Let sCreator = oRS!Creator
        Let sLabel = oRS!Label
        Let sFilePath = oRS!FilePath
        Let sAddedToDatabaseBy = oRS!AddedToDatabaseBy
        Let sAddedToDatabaseOn = oRS!AddedToDatabaseOn
    Else
        Let iID = 0
        Let sCreator = ""
        Let sLabel = ""
        Let sFilePath = ""
        Let sAddedToDatabaseBy = ""
        Let sAddedToDatabaseOn = Now
    End If
    
    Set oRS = Nothing
End Sub

Private Function FilterFilePath(c As Variant) As Collection
    Dim vItem As Variant
    Dim sItem As String
    Dim ret As New Collection
    Dim i As Integer
    Dim re As New RegExp
    Dim rParen As Boolean
    Dim oAgency As cAgency
    
    rParen = False
    
    i = 0
    For Each vItem In c
        i = i + 1: Let sItem = vItem
        
        'Place 1: Probably a CurName
        If i = 1 Then
            With re
                .pattern = "^[A-Za-z]{3}$"
                .IgnoreCase = True
            End With
            If re.Test(Trim(sItem)) Then
                Set oAgency = New cAgency: Let oAgency.CreatorCode = sItem
                Let sItem = oAgency.Name(Short:=True)
                If Len(sItem) = 0 Then
                    Let sItem = vItem
                End If
            End If
        End If
        
        'Place 2: Possibly a date or date range
        If i = 2 Then
            With re
                .pattern = "^[0-9]{4,8}$"
            End With
            If re.Test(Trim(sItem)) Then
                If Len(sItem) > 4 Then
                    sItem = Left(sItem, 4) & "-" & Right(sItem, Len(sItem) - 4)
                End If
                sItem = "(" & sItem & ":"
                rParen = True
            End If
        End If
        
        'Place 3: Usually the Governor's last name, just before "Control Files" or similar
        If i >= 3 Then
            If rParen Then
                rParen = False
                sItem = sItem & ")"
            End If
        End If
        
        'Jam it in to the result
        ret.Add sItem
    Next vItem
    
    Set FilterFilePath = ret
    
End Function
