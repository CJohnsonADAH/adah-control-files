VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAccessionsToAccnScans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**
'* @class cAccessionsToAccnScans
'**

Option Explicit
Option Compare Database

Public Sub InsertAccessionScans(AccessionScans As Variant)
    Dim vAccessionScan As Variant
    Dim oAccnScan As cAccnScan
    
    For Each vAccessionScan In AccessionScans
        Set oAccnScan = New cAccnScan
        oAccnScan.Url = vAccessionScan
        oAccnScan.InsertIntoAccessions
    Next vAccessionScan
End Sub
