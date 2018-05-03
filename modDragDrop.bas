Attribute VB_Name = "modDragDrop"
Option Compare Database
Option Explicit

Private fForm As Form

'************* Code Start *************
' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Dev Ashish
'
Private Declare Function apiCallWindowProc Lib "user32" _
    Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, _
    ByVal Hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) _
    As Long
   
Private Declare Function apiSetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
    (ByVal Hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal wNewWord As Long) _
    As Long

Private Declare Function apiGetWindowLong Lib "user32" _
    Alias "GetWindowLongA" _
   (ByVal Hwnd As Long, _
   ByVal nIndex As Long) _
   As Long

Private Declare Sub sapiDragAcceptFiles Lib "shell32.dll" _
    Alias "DragAcceptFiles" _
    (ByVal Hwnd As Long, _
    ByVal fAccept As Long)
    
Private Declare Sub sapiDragFinish Lib "shell32.dll" _
    Alias "DragFinish" _
    (ByVal hDrop As Long)

Private Declare Function apiDragQueryFile Lib "shell32.dll" _
    Alias "DragQueryFileA" _
    (ByVal hDrop As Long, _
    ByVal iFile As Long, _
    ByVal lpszFile As String, _
    ByVal cch As Long) _
    As Long

Private lpPrevWndProc  As Long
Private Const GWL_WNDPROC  As Long = (-4)
Private Const GWL_EXSTYLE = (-20)
Private Const WM_DROPFILES = &H233
Private Const WS_EX_ACCEPTFILES = &H10&
Private hWnd_Frm As Long

Sub sDragDrop(ByVal Hwnd As Long, _
                            ByVal Msg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long)

Dim lngRet As Long, strTmp As String, intLen As Integer
Dim lngCount As Long, I As Long, strOut As String
Const cMAX_SIZE = 50
    On Error Resume Next
    If Msg = WM_DROPFILES Then
        strTmp = String$(255, 0)
        lngCount = apiDragQueryFile(wParam, &HFFFFFFFF, strTmp, Len(strTmp))
        For I = 0 To lngCount - 1
            strTmp = String$(cMAX_SIZE, 0)
            intLen = apiDragQueryFile(wParam, I, strTmp, cMAX_SIZE)
            strOut = strOut & Left$(strTmp, intLen) & ";"
        Next I
        strOut = Left$(strOut, Len(strOut) - 1)
        Call sapiDragFinish(wParam)
        With fForm
            MsgBox "DROP: " & strOut
            '.RowSourceType = "Value List"
            '.RowSource = strOut
            'Forms!frmDragDrop.Caption = "DragDrop: " & _
            '                                       .ListCount & _
            '                                        " files dropped."
        End With
        
    Else
        lngRet = apiCallWindowProc( _
                            ByVal lpPrevWndProc, _
                            ByVal Hwnd, _
                            ByVal Msg, _
                            ByVal wParam, _
                            ByVal lParam)
    End If
End Sub

Sub sEnableDrop(frm As Form)
Dim lngStyle As Long, lngRet As Long
    lngStyle = apiGetWindowLong(frm.Hwnd, GWL_EXSTYLE)
    lngStyle = lngStyle Or WS_EX_ACCEPTFILES
    lngRet = apiSetWindowLong(frm.Hwnd, GWL_EXSTYLE, lngStyle)
    Call sapiDragAcceptFiles(frm.Hwnd, True)
    hWnd_Frm = frm.Hwnd
    
    Set fForm = frm
End Sub


Sub sHook(Hwnd As Long, strFunction As String)
    lpPrevWndProc = apiSetWindowLong(Hwnd, GWL_WNDPROC, AddressOf sDragDrop)
End Sub

Sub sUnhook(Hwnd As Long)
Dim lngTmp As Long
    lngTmp = apiSetWindowLong(Hwnd, _
                    GWL_WNDPROC, _
                    lpPrevWndProc)
    lpPrevWndProc = 0
End Sub

