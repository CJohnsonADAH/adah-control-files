VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Property Get ClassId() As String

End Property

Public Property Get InstanceId() As String

End Property

Public Sub HandleAction(Parameters As Variant, Outlet As String)

End Sub

Public Function Filter(Element As Variant, Parameters As Variant, Outlet As String) As Variant

End Function
