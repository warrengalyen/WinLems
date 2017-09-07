Attribute VB_Name = "mSound"
'================================================
' Module:        mSound.bas
' Author:        Warren Galyen
' Dependencies:
' Last revision: 11.13.2006
'================================================

Option Explicit

Private Declare Function PlaySound Lib "winmm" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const SND_ASYNC     As Long = &H1
Private Const SND_NODEFAULT As Long = &H2
Private Const SND_MEMORY    As Long = &H4
Private Const SND_NOWAIT    As Long = &H2000

Private Type uSFXData
    aData() As Byte
End Type

Private m_uSFX(23)      As uSFXData
Private m_bSoundEffects As Boolean

Public Enum eSoundFX
    [sfxBang] = 0
    [sfxChain]
    [sfxChangeOp]
    [sfxChink]
    [sfxDie]
    [sfxDoor]
    [sfxElectric]
    [sfxExplode]
    [sfxFire]
    [sfxGlug]
    [sfxLetsGo]
    [sfxManTrap]
    [sfxMousePre]
    [sfxOhNo]
    [sfxOing]
    [sfxScrape]
    [sfxSlicer]
    [sfxSplash]
    [sfxSplat]
    [sfxTenton]
    [sfxThud]
    [sfxThunk]
    [sfxTing]
    [sfxYipee]
End Enum



'========================================================================================
' Methods
'========================================================================================

Public Sub InitializeSound()
  
  Dim hFile As Long
  Dim sPath As String
  Dim i     As Long
  
    With fMain
        
        '-- Sounds path
        .flFilter.Path = AppPath & "SOUND\"
        .flFilter.Pattern = "*.wav"

        '-- Load available sounds
        For i = 0 To .flFilter.ListCount - 1
            
            '-- Get full path and a free file handle
            sPath = .flFilter.Path & "\" & .flFilter.List(i)
            hFile = VBA.FreeFile()
            
            '-- Open file
            Open sPath For Binary Access Read As #hFile
                
                '-- Resize array and get sound data
                With m_uSFX(i)
                    ReDim .aData(FileLen(sPath) - 1)
                    Get #hFile, , .aData()
                End With
            Close #hFile
        Next i
    End With
    
    '-- Default sound effects
    m_bSoundEffects = True
End Sub

Public Sub PlaySoundFX( _
           ByVal SoundFX As eSoundFX _
           )
    
    '-- Play sound FX
    If (m_bSoundEffects) Then
        Call PlaySound( _
             m_uSFX(SoundFX).aData(0), _
             0, _
             SND_ASYNC Or SND_MEMORY Or SND_NOWAIT _
             )
    End If
End Sub

Public Sub SetSoundEffectsState( _
           ByVal Enable As Boolean _
           )
    
    '-- Enable/disable sound effects
    m_bSoundEffects = Enable
End Sub



