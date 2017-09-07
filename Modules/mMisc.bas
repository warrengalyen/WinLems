Attribute VB_Name = "mMisc"
'================================================
' Module:        mMisc.bas
' Author:        Warren Galyen
' Dependencies:
' Last revision: 01.23.2006
'================================================

Option Explicit

'-- API:

Private Const BITSPIXEL As Long = 12

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Const GW_CHILD      As Long = 5
Private Const GWL_STYLE     As Long = -16
Private Const TB_STYLE_FLAT As Long = &H848
Private Const BS_FLAT       As Long = &H8000& 'Needs Style = [Standard]
Private Const BS_OWNERDRAW  As Long = &HB     'Needs Style = [Graphical]

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const SW_SHOW  As Long = 5

Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



'========================================================================================
' Methods
'========================================================================================

Public Function InIDE( _
                Optional c As Boolean = False _
                ) As Boolean
  
  Static b As Boolean
  
    b = c
    If (b = False) Then
        Debug.Assert InIDE(True)
    End If
    InIDE = b
    
End Function

Public Function FileExists( _
                ByVal Filename As String _
                ) As Boolean
    
    If (Len(Filename)) Then
        FileExists = (Dir$(Filename) <> vbNullString)
    End If
End Function

Public Function ScreenColourDepth( _
                ) As Long
 
 Dim hTmpDC As Long
   
    hTmpDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    ScreenColourDepth = GetDeviceCaps(hTmpDC, BITSPIXEL)
    Call DeleteDC(hTmpDC)
End Function

Public Function AppPath( _
                ) As String

    If (Right$(App.Path, 1) <> "\") Then
        AppPath = App.Path & "\"
      Else
        AppPath = App.Path
    End If
End Function

Public Sub Navigate( _
           ByVal hOwnerWnd As Long, _
           ByVal sURL As String _
           )
    
    Call ShellExecute(hOwnerWnd, "open", sURL, vbNullString, vbNullString, SW_SHOW)
End Sub

Public Sub FlattenToolbar( _
           oToolbar As Toolbar _
           )
    
  Dim hBar As Long
  Dim lRet As Long
    
    hBar = GetWindow(oToolbar.hWnd, GW_CHILD)
    lRet = GetWindowLong(hBar, GWL_STYLE)
    Call SetWindowLong(hBar, GWL_STYLE, lRet Or TB_STYLE_FLAT)
End Sub

Public Sub FlattenButton( _
           oButton As CommandButton _
           )
  
  Dim lRet As Long
  
    lRet = GetWindowLong(oButton.hWnd, GWL_STYLE)
    Call SetWindowLong(oButton.hWnd, GWL_STYLE, lRet Or BS_FLAT)
End Sub

Public Sub RemoveButtonBorderEnhance( _
           oButton As CommandButton _
           )
  
  Dim lRet As Long
  
    lRet = GetWindowLong(oButton.hWnd, GWL_STYLE)
    Call SetWindowLong(oButton.hWnd, GWL_STYLE, lRet And Not BS_OWNERDRAW)
End Sub


