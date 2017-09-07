Attribute VB_Name = "mSettings"
'================================================
' Module:        mSettings.bas
' Author:        Warren Galyen
' Dependencies:  -
' Last revision: 08.02.2006
'================================================

Option Explicit
Option Compare Text

'========================================================================================
' Methods
'========================================================================================

Public Sub SetLevelDone( _
           ByVal ID As Integer, _
           ByVal Done As Boolean _
           )
    
    Call PutINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Levels", _
         "Level_" & Format$(ID, "0000"), _
         IIf(Done, "1", "0") _
         )
End Sub

Public Sub SetLastLevel( _
           ByVal ID As Integer _
           )
    
    Call PutINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Settings", _
         "Last", _
         Format$(ID, "0000") _
         )
End Sub

Public Function GetLastLevel( _
                ) As Integer
    
    GetLastLevel = Val(GetINI( _
                       AppPath & "CONFIG\GAME.ini", _
                       "Settings", _
                       "Last") _
                       )
End Function

Public Function IsLevelDone( _
                ByVal ID As Integer _
                ) As Boolean
    
    IsLevelDone = Val(GetINI( _
                      AppPath & "CONFIG\GAME.ini", _
                      "Levels", _
                      "Level_" & Format$(ID, "0000")) _
                      )
End Function


