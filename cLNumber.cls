VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private sLNumber As String
Private iOffset As Long
Private sCollection As String

Public Property Get Number() As String
    Let Number = sLNumber
End Property

Public Property Let Number(s As String)
    Dim sOffset As String
    Dim sOffsetHead As String
    Dim sOffsetTail As String
    
    If Len(s) > 0 Then
        Let sCollection = RegexComponent(s, "^([A-Z]+)\s*([^A-Z].*)$", 1)
        Let sOffset = RegexComponent(s, "^([A-Z]+)\s*([^A-Z].*)$", 2)
    
        Let sOffsetHead = RegexComponent(sOffset, "^([0-9]+)([^0-9].*)?", 1)
        Let sOffsetTail = RegexComponent(sOffset, "^([0-9]+)([^0-9].*)?", 2)
        
        If Len(sOffsetHead) > 0 Then
            Let iOffset = CLng(sOffsetHead)
            Let sOffset = Format(iOffset, "000000") & sOffsetTail
        Else
            Let sOffset = sOffsetTail
        End If

    Else
        Let sCollection = ""
        Let sOffset = ""
        Let iOffset = 0
        Let sOffsetTail = ""
    End If
    Let sLNumber = sCollection & sOffset
End Property

Public Property Get Collection() As String
    Let Collection = sCollection
End Property

Public Property Get Offset() As Integer
    Let Offset = iOffset
End Property

Public Function SameAs(v As Variant) As Boolean
    Dim o As cLNumber
    
    If TypeName(v) = "String" Then
        Set o = New cLNumber
        o.Number = v
    ElseIf IsObject(v) Then
        Set o = v
    Else
        '-- Raise a custom error message
    End If
    
    Let SameAs = (Me.Number = o.Number)
End Function
