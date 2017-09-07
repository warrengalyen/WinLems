Attribute VB_Name = "mWheel"
'================================================
' Module:        mWheel.bas
' Author:        Warren Galyen
' Dependencies:
' Last revision: 11.13.2006
'================================================

Option Explicit

'-- API:

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC   As Long = -4
Private Const WM_MOUSEWHEEL As Long = &H20A

'//

Private m_OldWindowProc As Long

Public Sub InitializeWheel()
    '-- New window proc.
    m_OldWindowProc = SetWindowLong(fMain.hWnd, GWL_WNDPROC, AddressOf wgWindowProc)
End Sub

Private Function wgWindowProc( _
                 ByVal hWnd As Long, _
                 ByVal uMsg As Long, _
                 ByVal wParam As Long, _
                 ByVal lParam As Long _
                 ) As Long
    
    If (uMsg = WM_MOUSEWHEEL) Then
        If (wParam > 0) Then
            Call VBA.SendKeys("Z")
          Else
            Call VBA.SendKeys("A")
        End If
    End If
    wgWindowProc = CallWindowProc(m_OldWindowProc, hWnd, uMsg, wParam, lParam)
End Function


