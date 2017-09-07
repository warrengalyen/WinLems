Attribute VB_Name = "mLevel"
'================================================
' Module:        mLevel.bas
' Author:        Warren Galyen
' Dependencies:
' Last revision: 01.23.2006
'================================================

Option Explicit

Private Type RECT2
    x1                      As Integer
    y1                      As Integer
    x2                      As Integer
    y2                      As Integer
End Type

Private Type uObjectGFX
    DIB                     As New cDIB08
    Width                   As Byte
    Height                  As Byte
    StartAnimationFrame     As Byte
    EndAnimationFrame       As Byte
    lpTriggerRect           As RECT2
    TriggerEffect           As Byte
    SoundEffect             As Byte
End Type

Private Type uTerrainGFX
    DIB                     As New cDIB08
    Width                   As Integer
    Height                  As Byte
End Type

Private Type uObject
    ID                      As Byte
    lpRect                  As RECT2
    NotOverlap              As Boolean
    OnTerrain               As Boolean
    UpsideDown              As Boolean
    wgFrameIdxCur           As Byte
    wgFrameIdxMax           As Byte
    wgLoop                  As Boolean
End Type

Private Type uTerrainPiece
    ID                      As Byte
    lpRect                  As RECT2
    NotOverlap              As Boolean
    Black                   As Boolean
    UpsideDown              As Boolean
End Type

Private Type uSteelArea
    lpRect                  As RECT2
End Type

Private Type uLevelData
    Title                   As String * 32
    LemsToLetOut            As Byte
    LemsToBeSaved           As Byte
    ReleaseRate             As Byte
    PlayingTime             As Byte
    MaxClimbers             As Byte
    MaxFloaters             As Byte
    MaxBombers              As Byte
    MaxBlockers             As Byte
    MaxBuilders             As Byte
    MaxBashers              As Byte
    MaxMiners               As Byte
    MaxDiggers              As Byte
    ScreenStart             As Integer
    GraphicSet              As Byte
    GraphicSetEx            As Byte
    Objects                 As Integer
    Object()                As uObject
    TerrainPieces           As Integer
    TerrainPiece()          As uTerrainPiece
    SteelAreas              As Integer
    SteelArea()             As uSteelArea
End Type

Public g_eGamePack          As eGamePack        ' 0, 1,..., 9 [custom]
Public g_nLevelID           As Integer          ' #### [pack#|rating#|level##]
Public g_uLevel             As uLevelData       ' level data
Public g_uObjGFX()          As uObjectGFX       ' object item
Public g_uTerGFX()          As uTerrainGFX      ' terrain item
Public g_oDIBBack           As New cDIB08       ' extended level background image



'========================================================================================
' Methods
'========================================================================================

Public Function LoadLevelTitle( _
                ByVal ID As Integer _
                ) As Boolean

  Dim sPath As String
  Dim hFile As Long
  
    sPath = AppPath & "LEVELS\" & Format$(ID, "0000") & ".dat"
    If (FileExists(sPath)) Then
        With g_uLevel
            hFile = VBA.FreeFile()
            Open sPath For Binary Access Read As #hFile
              Get #hFile, , .Title
            Close #hFile
        End With
        LoadLevelTitle = True
    End If
End Function

Public Function LoadLevel( _
                ByVal ID As Integer, _
                Optional ByVal LoadGraphics As Boolean = True _
                ) As Boolean

  Dim sPath As String
  Dim hFile As Long
  Dim i     As Long
  
    sPath = AppPath & "LEVELS\" & Format$(ID, "0000") & ".dat"
    
    If (FileExists(sPath)) Then
    
        With g_uLevel
            
            hFile = VBA.FreeFile()
            Open sPath For Binary Access Read As #hFile
            
                Get #hFile, , .Title
                
                Get #hFile, , .LemsToLetOut
                Get #hFile, , .LemsToBeSaved
                Get #hFile, , .ReleaseRate
                Get #hFile, , .PlayingTime
                
                Get #hFile, , .MaxClimbers
                Get #hFile, , .MaxFloaters
                Get #hFile, , .MaxBombers
                Get #hFile, , .MaxBlockers
                Get #hFile, , .MaxBuilders
                Get #hFile, , .MaxBashers
                Get #hFile, , .MaxMiners
                Get #hFile, , .MaxDiggers
                
                If (LoadGraphics) Then
                
                    Get #hFile, , .ScreenStart
                    Get #hFile, , .GraphicSet
                    Get #hFile, , .GraphicSetEx
                    
                    Call wgLoadGraphicSet(.GraphicSet, .GraphicSetEx)
                
                    Get #hFile, , .Objects
                    ReDim .Object(0 To .Objects)
                    
                    For i = 1 To .Objects
                        With .Object(i)
                            Get #hFile, , .ID
                            Get #hFile, , .lpRect.x1: .lpRect.x2 = .lpRect.x1 + g_uObjGFX(.ID).Width
                            Get #hFile, , .lpRect.y1: .lpRect.y2 = .lpRect.y1 + g_uObjGFX(.ID).Height
                            Get #hFile, , .NotOverlap
                            Get #hFile, , .OnTerrain
                            Get #hFile, , .UpsideDown
                            
                            .wgFrameIdxCur = g_uObjGFX(.ID).StartAnimationFrame
                            .wgFrameIdxMax = g_uObjGFX(.ID).EndAnimationFrame
                        End With
                    Next i
            
                    Get #hFile, , .TerrainPieces
                    ReDim .TerrainPiece(0 To .TerrainPieces)
        
                    For i = 1 To .TerrainPieces
                        With .TerrainPiece(i)
                            Get #hFile, , .ID
                            Get #hFile, , .lpRect.x1: .lpRect.x2 = .lpRect.x1 + g_uTerGFX(.ID).Width
                            Get #hFile, , .lpRect.y1: .lpRect.y2 = .lpRect.y1 + g_uTerGFX(.ID).Height
                            Get #hFile, , .NotOverlap
                            Get #hFile, , .Black
                            Get #hFile, , .UpsideDown
                        End With
                    Next i
        
                    Get #hFile, , .SteelAreas
                    ReDim .SteelArea(0 To .SteelAreas)
        
                    For i = 1 To .SteelAreas
                        With .SteelArea(i)
                            Get #hFile, , .lpRect.x1
                            Get #hFile, , .lpRect.y1
                            Get #hFile, , .lpRect.x2
                            Get #hFile, , .lpRect.y2
                        End With
                    Next i
                End If
            Close #hFile
        End With
        LoadLevel = True
    End If
End Function

Public Function GetNextLevel( _
                ) As Integer
    
  Dim sTmp As String
    
    If (g_eGamePack = [gpCustom]) Then
        
        '-- Return same level:
        '   Unexpected sequence
        GetNextLevel = g_nLevelID
    
      Else
      
        '-- Get next level (same rating) file path
        sTmp = AppPath & _
               "LEVELS\" & _
               Format$(g_nLevelID + 1, "0000") & ".dat"
        
        '-- Check file
        If (FileExists(sTmp)) Then
            
            '-- Exists: OK
            GetNextLevel = g_nLevelID + 1
          Else
            
            '-- Get first level (next rating)
            sTmp = AppPath & _
                   "LEVELS\" & _
                   Format$((g_nLevelID \ 100 + 1) * 100, "0000") & ".dat"
            
            '-- Check file
            If (FileExists(sTmp)) Then
                
                '-- Exists: OK
                GetNextLevel = (g_nLevelID \ 100 + 1) * 100
              
              Else
                '-- All done: start again
                GetNextLevel = g_eGamePack * 1000
            End If
        End If
    End If
End Function

Public Function GetLevelRatingString( _
                ByVal ID As Integer _
                ) As String
    
  Dim s As String
    
    '-- Get rating ID
    s = Format$(ID, "0000")
    s = Mid$(s, 2, 1)
    
    '-- Available ratings
    Select Case g_eGamePack
        Case [gpLems]
            GetLevelRatingString = Choose(Val(s) + 1, _
                                   "Fun", _
                                   "Tricky", _
                                   "Taxing", _
                                   "Mayhem" _
                                   )
        Case [gpOhNoMoreLems]
            GetLevelRatingString = Choose(Val(s) + 1, _
                                   "Tame", _
                                   "Crazy", _
                                   "Wild", _
                                   "Wicked", _
                                   "Havoc" _
                                   )
        Case [gpCustom]
            GetLevelRatingString = "N/A"
    End Select
End Function

'========================================================================================
' Private
'========================================================================================

Private Sub wgLoadGraphicSet( _
            ByVal GraphicSet As Byte, _
            ByVal GraphicSetEx As Byte _
            )
  
  Dim i    As Long
  Dim sINI As String
  Dim sKey As String
    
    Screen.MousePointer = vbHourglass
    
    With g_uLevel
        
        '-- INI file
        If (GraphicSetEx > 0) Then
            sINI = AppPath & "CONFIG\GS_" & GraphicSetEx & "EX.ini"
          Else
            sINI = AppPath & "CONFIG\GS_" & GraphicSet & ".ini"
        End If
        
        '-- Resize objects collection
        ReDim g_uObjGFX(0 To Val( _
            GetINI(sINI, "main", "ObjectCount")) - 1)
        
        '-- Load available objects
        For i = 0 To UBound(g_uObjGFX())
            
            '-- Create 8bit image
            Call g_uObjGFX(i).DIB.CreateFromBitmapFile( _
                 AppPath & "GFX\" & _
                 "obj_" & GraphicSet & "_" & Format$(i, "00") & ".bmp" _
                 )
            
            '-- Get animation info
            sKey = "obj_" & Format$(i, "00")
            With g_uObjGFX(i)
                
                '-- Animation frame size
                .Width = _
                    Val(GetINI(sINI, sKey, "Width"))
                .Height = _
                    Val(GetINI(sINI, sKey, "Height"))
                
                '-- Start and ending frames
                .StartAnimationFrame = _
                    Val(GetINI(sINI, sKey, "StartAnimationFrame"))
                .EndAnimationFrame = _
                    Val(GetINI(sINI, sKey, "EndAnimationFrame"))
                
                '-- Trigger area and related effect
                .TriggerEffect = _
                    Val(GetINI(sINI, sKey, "TriggerEffect"))
                With .lpTriggerRect
                    .x1 = _
                        Val(GetINI(sINI, sKey, "TriggerLeft"))
                    .x2 = .x1 + _
                        Val(GetINI(sINI, sKey, "TriggerWidth"))
                    .y1 = _
                        Val(GetINI(sINI, sKey, "TriggerTop"))
                    .y2 = .y1 + _
                        Val(GetINI(sINI, sKey, "TriggerHeight"))
                End With
                
                '-- Trap sound effect
                .SoundEffect = _
                    Val(GetINI(sINI, sKey, "SoundEffect"))
            End With
        Next i
        
        '-- Extended level?
        If (GraphicSetEx > 0) Then
            
            '-- Load background image
             Call g_oDIBBack.CreateFromBitmapFile( _
                  AppPath & "GFX\" & _
                  "back_" & Format$(GraphicSetEx, "0") & "ex.bmp" _
                  )
          Else
            
            '-- Resize terrain pieces collection
            ReDim g_uTerGFX(0 To Val( _
                GetINI(sINI, "main", "TerrainCount")) - 1)
            
            '-- Load available terrain pieces
            For i = 0 To UBound(g_uTerGFX())
                Call g_uTerGFX(i).DIB.CreateFromBitmapFile( _
                     AppPath & "GFX\" & _
                     "ter_" & Format$(GraphicSet, "0") & "_" & Format$(i, "00") & ".bmp" _
                     )
                     
                '-- Set item info
                With g_uTerGFX(i)
                    .Width = .DIB.Width
                    .Height = .DIB.Height
                End With
            Next i
        End If
    End With

    '-- Finaly, load/merge level palette
    Call MergePaletteEntries( _
        GetINI(sINI, "main", "BrickColor"), 7)
    Call MergePaletteEntries( _
        GetINI(sINI, "main", "Palette"), 8)
    Call fMain.ucScreen.UpdatePalette(GetGlobalPalette())
    
    Screen.MousePointer = vbDefault
End Sub



