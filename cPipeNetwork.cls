VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPipeNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private dPipelines As Dictionary
Private dFilters As Dictionary

Private Function EventOf(ByVal Outlet As String, ByVal EventType As String)
    Select Case EventType
    Case "Action":
        Let EventOf = "Do-" & Outlet
    Case "Filter":
        Let EventOf = "Apply-" & Outlet
    Case Else:
        Let EventOf = EventType & "-" & Outlet
    End Select
End Function

Public Sub AddAction(ByVal Outlet As String, Plug As IReceiver, Optional ByVal Priority As Integer)
    AddEvent Outlet:=EventOf(Outlet, "Action"), Plug:=Plug, Priority:=Priority
End Sub

Public Sub RemoveAction(ByVal Outlet As String, Instance As Variant)
    RemoveEvent Outlet:=EventOf(Outlet, "Action"), Instance:=Instance
End Sub

Public Sub DoAction(ByVal Outlet As String, Optional Parameters As Variant)
    Dim dPipe As Dictionary
    Dim dTier As Dictionary
    Dim Tiers As Variant
    Dim Action As Variant
    Dim PriorityLevel As Variant
    Dim Plug As Variant
    Dim rcvPlug As IReceiver
    Dim sEvent As String
    
    Let sEvent = EventOf(Outlet, "Action")
    
    If dPipelines.Exists(sEvent) Then
        Set dPipe = dPipelines.Item(sEvent)
        
        Let Tiers = dPipe.Keys
        BubbleSortList Tiers
        
        For Each PriorityLevel In Tiers
            Set dTier = dPipe.Item(PriorityLevel)
            For Each Action In dTier.Keys
                Set Plug = dTier.Item(Action)
                If TypeOf Plug Is IReceiver Then
                    Set rcvPlug = Plug
                    Call rcvPlug.HandleAction(Parameters:=Parameters, Outlet:=Outlet)
                End If
            Next Action
        Next PriorityLevel
    End If

End Sub

Public Sub AddFilter(ByVal Outlet As String, Plug As IReceiver, Optional ByVal Priority As Integer)
    AddEvent Outlet:=EventOf(Outlet, "Filter"), Plug:=Plug, Priority:=Priority
End Sub

Public Sub RemoveFilter(ByVal Outlet As String, Instance As Variant)
    RemoveEvent Outlet:=EventOf(Outlet, "Filter"), Instance:=Instance
End Sub

Public Function ApplyFilters(ByVal Outlet As String, InputElement As Variant, Optional Parameters As Variant)
    Dim OutputElement As Variant
    Dim dPipe As Dictionary
    Dim dTier As Dictionary
    Dim Tiers As Variant
    Dim Filter As Variant
    Dim PriorityLevel As Variant
    Dim Plug As Variant
    Dim sEvent As String
    
    Let sEvent = EventOf(Outlet, "Filter")
    
    Let OutputElement = InputElement
    If dPipelines.Exists(sEvent) Then
        Set dPipe = dPipelines.Item(sEvent)
        
        Let Tiers = dPipe.Keys
        BubbleSortList Tiers
        
        For Each PriorityLevel In Tiers
            Set dTier = dPipe.Item(PriorityLevel)
            For Each Filter In dTier.Keys
                Set Plug = dTier.Item(Filter)
                Let OutputElement = Plug.Filter(Element:=OutputElement, Parameters:=Parameters, Outlet:=Outlet)
            Next Filter
        Next PriorityLevel
    End If
    Let ApplyFilters = OutputElement
End Function

Public Sub AddEvent(ByVal Outlet As String, Plug As IReceiver, Optional ByVal Priority As Integer)
    Dim dPipe As Dictionary
    Dim dTier As Dictionary
    Dim dLoc As Dictionary
    
    If Not dPipelines.Exists(Outlet) Then
        Set dPipe = New Dictionary
        dPipelines.Add Outlet, dPipe
    End If
    Set dPipe = dPipelines.Item(Outlet)
    
    If Not dPipe.Exists(Priority) Then
        Set dTier = New Dictionary
        dPipe.Add Priority, dTier
    End If
    Set dTier = dPipe.Item(Priority)
    
    If Not dTier.Exists(Plug.InstanceId) Then
        dTier.Add Key:=Plug.InstanceId, Item:=Plug
        
        Set dLoc = New Dictionary
        dLoc.Add Key:="Outlet", Item:=Outlet
        dLoc.Add Key:="Priority", Item:=Priority
        
        dFilters.Add Key:=Outlet & "," & Plug.InstanceId, Item:=dLoc
    End If
End Sub

Public Sub RemoveEvent(ByVal Outlet As String, Instance As Variant)
    Dim dLoc As Dictionary
    Dim dPipe As Dictionary
    Dim dTier As Dictionary
    Dim InstanceId As String
    
    Dim Key As String
    
    If IsObject(Instance) Then
        Let InstanceId = Instance.InstanceId
    Else
        Let InstanceId = Instance
    End If
    
    Let Key = Outlet & "," & InstanceId
    If dFilters.Exists(Key) Then
        Set dLoc = dFilters.Item(Key)
        
        Set dPipe = dPipelines.Item(dLoc.Item("Outlet"))
        Set dTier = dPipe.Item(dLoc.Item("Priority"))
        
        dTier.Remove Key:=InstanceId
        dFilters.Remove Key:=Key
    End If
End Sub

Private Sub Class_Initialize()
    Set dPipelines = New Dictionary
    Set dFilters = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set dPipelines = Nothing
    Set dFilters = New Dictionary
End Sub
