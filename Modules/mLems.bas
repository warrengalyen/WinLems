Attribute VB_Name = "mLems"
'================================================
' Module:        mLems.bas
' Author:        Warren Galyen
' Dependencies:
' Last revision: 03.13.2007
'================================================

Option Explicit

'-- A little bit of API

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    wgData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'-- Lems

Private Type RECT2
    x1           As Integer
    y1           As Integer
    x2           As Integer
    y2           As Integer
End Type

Private Type uParticle
    x            As Integer
    y            As Integer
    vx           As Integer
    vy           As Integer
End Type

Private Type uAnimationData
    FrameOffY    As Integer
    FrameIdxMax  As Byte
    FrameHasDir  As Boolean
End Type

Private Type uLem
    Active       As Boolean
    DieNextFrame As Boolean
    Job          As eLemJob
    Ability      As eLemAbility
    ExplodeCount As Integer
    Frame        As eLemFrame
    FrameSrcY    As Integer
    FrameIdx     As Integer
    FrameIdxMax  As Integer
    FrameOffY    As Integer
    FrameHasDir  As Boolean
    x            As Integer
    y            As Integer
    xs           As Integer
    Counter      As Integer
    Particles    As Boolean
    Particle(24) As uParticle
End Type

Public Enum eLemFrame
    [fWalker] = 0
    [fFalling] = 2
    [fSpliting] = 4
    [fDrowning] = 5
    [fBurning] = 6
    [fExploding] = 7
    [fSurviving] = 8
    [fClimber] = 9
    [fClimberEnd] = 11
    [fFloater] = 13
    [fBlocker] = 15
    [fBuilder] = 16
    [fBuilderEnd] = 18
    [fBasher] = 20
    [fMiner] = 22
    [fDigger] = 24
End Enum

Public Enum eLemJob
    [jNone] = 0
    [jBlocker] = 1
    [jBuilder] = 2
    [jBasher] = 3
    [jMiner] = 4
    [jDigger] = 5
End Enum

Public Enum eLemAbility
    [aNone] = 0
    [aClimber] = 1
    [aFloater] = 2
    [aBomber] = 4
End Enum

Public Enum eGameStage
    [gsLetsGo] = 0
    [gsOpeningDoors]
    [gsPlaying]
    [gsEnding]
End Enum

Private Const MIN_FALL_FLOATER As Long = 20
Private Const MAX_FALL         As Long = 60
Private Const MIN_OBSTACLE     As Long = 7
Private Const MAX_BRICKS       As Long = 12
Private Const MAX_YCHECK       As Long = 175
Private Const EXPLODE_TICKS    As Long = 75
Private Const DOOR_ID          As Byte = 1

Private Const ANIMATION_DATA   As String = "00071000000003100000001500015000130001500007000071000000007100000000910000000150001510000000071000000031100000022310000002151"
Private m_uAnimationData(24)   As uAnimationData
Private Const BASHER_DATA      As String = "12222233345432111222223344543211"
Private m_aBasherData(31)      As Byte

Private m_eGameStage           As eGameStage
Private m_lcGameStage          As Long

Private m_lpoScreen            As ucScreen08
Private m_oDIBLems             As New cDIB08
Private m_oDIBMask             As New cDIB08
Private m_oDIBScreenBuffer     As New cDIB08
Private m_oDIBScreenBKMask     As New cDIB08

Private m_lpoPanView           As ucScreen08
Private m_uPanViewRct          As RECT
Private m_uPanViewSelectionRct As RECT
Private m_oDIBLemPoint         As New cDIB08

Private m_uSAScreenBKMask      As SAFEARRAY2D
Private m_aScreenBKMaskBits()  As Byte
Private m_uScreenBKMaskRct     As RECT

Private m_iCursorPointer       As New StdPicture
Private m_iCursorSelect        As New StdPicture

Private m_xScreen              As Long
Private m_xCur                 As Long
Private m_yCur                 As Long

Private m_ePreparedAbility     As eLemAbility
Private m_ePreparedJob         As eLemJob
Private m_lLem                 As Long
    
Private m_uPtOut()             As POINTAPI
Private m_lPtOut               As Long
Private m_lcPtOut              As Long

Private m_uLems()              As uLem
Private m_lLemsOut             As Long
Private m_lLemsSaved           As Long



'========================================================================================
' Initialization / termination
'========================================================================================

Public Sub InitializeLems()

  Dim i As Long
    
    '-- Initialize constant data
    For i = 0 To 24
        With m_uAnimationData(i)
            .FrameOffY = _
                Mid$(ANIMATION_DATA, 5 * i + 1, 2)
            .FrameIdxMax = _
                Mid$(ANIMATION_DATA, 5 * i + 3, 2)
            .FrameHasDir = _
                Mid$(ANIMATION_DATA, 5 * i + 5, 1)
        End With
    Next i
    For i = 0 To 31
        m_aBasherData(i) = Mid$(BASHER_DATA, i + 1, 1)
    Next i
    
    '-- Private reference to main form 'screens'
    Set m_lpoScreen = fMain.ucScreen
    Set m_lpoPanView = fMain.ucPanView
    
    '-- Pre-load selection cursors
    Set m_iCursorPointer = VB.LoadResPicture( _
        "CUR_POINTER", vbResCursor _
        )
    Set m_iCursorSelect = VB.LoadResPicture( _
        "CUR_SELECT", vbResCursor _
        )
    
    '-- Load Lems' frames & masks
    Call m_oDIBLems.CreateFromBitmapFile( _
         AppPath & "GFX\main_1.bmp" _
         )
    Call m_oDIBMask.CreateFromBitmapFile( _
         AppPath & "GFX\main_2.bmp" _
         )
    
    '-- Create main buffer (terrain) and back-mask DIBs
    Call m_oDIBScreenBuffer.Create( _
         1600, 160 _
         ) ' 250 KB
    Call m_oDIBScreenBKMask.Create( _
         1600, 160 _
         ) ' 250 KB
        
    '-- Define back-mask and panoramic view rects.
    Call SetRect( _
         m_uScreenBKMaskRct, _
         0, 0, 1600, 160 _
         )
    Call SetRect(m_uPanViewRct, _
         0, 0, 640, 80 _
         )
    
    '-- Map back-mask DIB bytes
    Call wgMapDIB( _
         m_uSAScreenBKMask, _
         m_aScreenBKMaskBits(), _
         m_oDIBScreenBKMask _
         )
    
    '-- Create Lem 'point' (preview)
    Call m_oDIBLemPoint.Create( _
         2, 2, _
         BkColorIdx:=IDX_GREEN _
         )
End Sub

Public Sub TerminateLems()

    '-- Unmap 'screen'
    Call wgUnmapDIB(m_aScreenBKMaskBits())
End Sub

'========================================================================================
' Terrain/Mask initialization
'========================================================================================

Public Sub InitializeGame()
    
  Dim i As Long
  
    '-- Reset game state
    Call SetGameStage([gsLetsGo])
    
    '-- Initialize 'Release rate' and 'Playing time'
    g_lReleaseRateMin = g_uLevel.ReleaseRate
    g_lReleaseRate = g_uLevel.ReleaseRate
    g_lcPlayingTime = g_uLevel.PlayingTime * 60
    
    '-- Initialize Lems' collection and related variables
    ReDim m_uLems(0)
    m_lLem = 0
    m_lLemsOut = 0
    m_lLemsSaved = 0
    m_ePreparedAbility = [aNone]
    m_ePreparedJob = [jNone]
    
    '-- Initialize Exits' collection
    ReDim m_uPtOut(0)
    m_lcPtOut = 0
    m_lPtOut = 1
    
    '-- Reset buffer and back-mask
    Call m_oDIBScreenBuffer.Reset
    Call m_oDIBScreenBKMask.Reset
    
    '-- Prepare level...
    With g_uLevel
    
        '-- Initialize terrain
        If (.GraphicSetEx > 0) Then
            Call MaskBlt( _
                 m_oDIBScreenBuffer, _
                 304, 0, _
                 960, 160, _
                 g_oDIBBack, _
                 0, 0, _
                 IDX_TRANS _
                 )
          Else
            For i = 1 To .TerrainPieces
                With .TerrainPiece(i)
                    If (.Black) Then
                        If (.NotOverlap) Then
                            Call MaskBltIdxOverlap( _
                                 m_oDIBScreenBuffer, _
                                 .lpRect.x1, .lpRect.y1, _
                                 g_uTerGFX(.ID).Width, g_uTerGFX(.ID).Height, _
                                 IDX_NONE, IDX_NONE, _
                                 g_uTerGFX(.ID).DIB, _
                                 0, 0, _
                                 IDX_TRANS, _
                                 .UpsideDown _
                                 )
                          Else
                            Call MaskBltIdx( _
                                 m_oDIBScreenBuffer, _
                                 .lpRect.x1, .lpRect.y1, _
                                 g_uTerGFX(.ID).Width, g_uTerGFX(.ID).Height, _
                                 IDX_NONE, _
                                 g_uTerGFX(.ID).DIB, _
                                 0, 0, _
                                 IDX_TRANS, _
                                 .UpsideDown _
                                 )
                        End If
                      Else
                        If (.NotOverlap) Then
                            Call MaskBltOverlap( _
                                 m_oDIBScreenBuffer, _
                                 .lpRect.x1, .lpRect.y1, _
                                 g_uTerGFX(.ID).Width, g_uTerGFX(.ID).Height, _
                                 IDX_NONE, _
                                 g_uTerGFX(.ID).DIB, _
                                 0, 0, _
                                 IDX_TRANS, _
                                 .UpsideDown _
                                 )
                          Else
                            Call MaskBlt( _
                                 m_oDIBScreenBuffer, _
                                 .lpRect.x1, .lpRect.y1, _
                                 g_uTerGFX(.ID).Width, g_uTerGFX(.ID).Height, _
                                 g_uTerGFX(.ID).DIB, _
                                 0, 0, _
                                 IDX_TRANS, _
                                 .UpsideDown _
                                 )
                        End If
                    End If
                End With
            Next i
        End If
        
        '-- Initialize back-mask
        
        '  Terrain...
        Call MaskBltIdx( _
             m_oDIBScreenBKMask, _
             0, 0, _
             1600, 160, _
             IDX_TERRAIN, _
             m_oDIBScreenBuffer, _
             0, 0, _
             IDX_NONE, _
             False _
             )
        
        '   Steel...
        For i = 1 To .SteelAreas
            With .SteelArea(i)
                Call MaskRectIdxOverlap( _
                     m_oDIBScreenBKMask, _
                     .lpRect.x1, .lpRect.y1, _
                     .lpRect.x2 - .lpRect.x1, .lpRect.y2 - .lpRect.y1, _
                     IDX_TERRAIN, IDX_STEEL _
                     )
            End With
        Next i

        '   Objects' trigger areas...
        For i = 1 To .Objects
            With .Object(i)
                If (g_uObjGFX(.ID).TriggerEffect > 0) Then
                    If (.OnTerrain) Then
                        Call MaskRectIdxOverlap( _
                             m_oDIBScreenBKMask, _
                             g_uObjGFX(.ID).lpTriggerRect.x1 + .lpRect.x1, _
                             g_uObjGFX(.ID).lpTriggerRect.y1 + .lpRect.y1, _
                             g_uObjGFX(.ID).lpTriggerRect.x2 - g_uObjGFX(.ID).lpTriggerRect.x1, _
                             g_uObjGFX(.ID).lpTriggerRect.y2 - g_uObjGFX(.ID).lpTriggerRect.y1, _
                             IDX_TERRAIN, g_uObjGFX(.ID).TriggerEffect _
                             )
                      Else
                        Call MaskRectIdxBkMask( _
                             m_oDIBScreenBKMask, _
                             g_uObjGFX(.ID).lpTriggerRect.x1 + .lpRect.x1, _
                             g_uObjGFX(.ID).lpTriggerRect.y1 + .lpRect.y1, _
                             g_uObjGFX(.ID).lpTriggerRect.x2 - g_uObjGFX(.ID).lpTriggerRect.x1, _
                             g_uObjGFX(.ID).lpTriggerRect.y2 - g_uObjGFX(.ID).lpTriggerRect.y1, _
                             g_uObjGFX(.ID).TriggerEffect _
                             )
                    End If
                End If
            End With
        Next i
        
        '-- Set loop mode
        For i = 1 To .Objects
            With .Object(i)
                Select Case .ID
                    '-- Door
                    Case Is = DOOR_ID
                        m_lcPtOut = m_lcPtOut + 1
                        ReDim Preserve m_uPtOut(0 To m_lcPtOut)
                        m_uPtOut(m_lcPtOut).x = .lpRect.x1 + 17
                        m_uPtOut(m_lcPtOut).y = .lpRect.y1 - 2
                    '-- Any other
                    Case Else
                        Select Case g_uObjGFX(.ID).TriggerEffect
                            Case Is <> IDX_TRAP
                                .wgLoop = True
                        End Select
                End Select
            End With
        Next i

        '-- Scroll to start position
        Call DoScrollTo(x:=.ScreenStart, ScaleAndCenter:=False)
    End With
End Sub

'========================================================================================
' Scrolling
'========================================================================================

Public Sub DoScroll( _
           ByVal dx As Long _
           )
    
    '-- Add offset
    m_xScreen = m_xScreen + 2.5 * dx
    
    '-- Scroll
    Call DoScrollTo(x:=m_xScreen, ScaleAndCenter:=False)
End Sub

Public Sub DoScrollTo( _
           ByVal x As Long, _
           Optional ByVal ScaleAndCenter As Boolean = True _
           )
    
    '-- Re-scale and center position
    '   (panoramic view scale 1:2.5)
    If (ScaleAndCenter) Then
        m_xScreen = 2.5 * x - 160
      Else
        m_xScreen = x
    End If
    
    '-- Check bounds
    If (m_xScreen < 0) Then
        m_xScreen = 0
      ElseIf (m_xScreen > 1280) Then
        m_xScreen = 1280
    End If
    
    '-- Define new scroll-selection
    Call SetRect(m_uPanViewSelectionRct, _
                 m_xScreen / 2.5, 0, _
                 m_xScreen / 2.5 + 127, 80 _
                 )
    
    '-- Render frame?
    If (IsTimerPaused) Then
        Call DoFrame
    End If
End Sub

'========================================================================================
' Paint a frame!
'========================================================================================

Public Sub DoFrame()

    '-- Render terrain, then objects
    Call wgDrawTerrain
    Call wgDrawObjects
    
    '-- Depending on stage...
    Select Case m_eGameStage
        
        Case [gsLetsGo]
            
            If (IsTimerPaused = False) Then
            
                '-- Play sound
                If (m_lcGameStage = 0) Then
                    Call PlayMidi(RandomTheme:=True)
                    Call PlaySoundFX([sfxLetsGo])
                End If
                
                '-- Stage frame counter
                m_lcGameStage = m_lcGameStage + 1
                
                '-- 10 frames after...
                If (m_lcGameStage = 10) Then
                    Call SetGameStage([gsOpeningDoors])
                End If
            End If
        
        Case [gsOpeningDoors]
            
            If (IsTimerPaused = False) Then
            
                '-- Stage 1st frame: open doors
                If (m_lcGameStage = 0) Then
                    Call wgSetDoorsState(bOpening:=True)
                End If
                
                '-- Stage frame counter
                m_lcGameStage = m_lcGameStage + 1
                
                '-- 10 frames after...
                Select Case m_lcGameStage
                    Case 10
                        Call wgSetDoorsState(bOpening:=False)
                    Case 15
                        Call SetGameStage([gsPlaying])
                End Select
            End If
            
        Case [gsPlaying]
            
            If (IsTimerPaused = False) Then
                '-- Check/move lems and render them
                Call wgCheckLems
                Call wgDrawLems
              Else
                '-- Only render lems
                Call wgDrawLems
            End If
        
        Case [gsEnding]
            
            '-- Stage 1st  frame: stop 'timer' timer and close doors
            If (m_lcGameStage = 0) Then
                Call wgSetDoorsState(bOpening:=False)
            End If
            
            '-- Stage frame counter
            If (IsTimerPaused = False) Then
                m_lcGameStage = m_lcGameStage + 1
            End If
            
            '-- Restore DIB palette
            Call m_lpoScreen.UpdatePalette( _
                 GetGlobalPalette() _
                 )
            
            '-- 10 frames...
            If (m_lcGameStage < 10) Then
                '-- Fade out
                Call m_lpoScreen.UpdatePalette( _
                     GetFadedOutGlobalPalette(Amount:=m_lcGameStage * 25) _
                     )
              Else
                '-- Level finished: stop timer and stop music
                Call StopTimer
                Call CloseMidi
                '-- 'Report'
                Call fMain.LevelDone
                Exit Sub
            End If
    End Select
    
    '-- Update screen
    Call m_lpoScreen.Refresh
    
    '-- Finaly, render preview window
    Call wgDrawPreview
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub AddLem()
    
    '-- Lems out counter
    m_lLemsOut = m_lLemsOut + 1
    
    '-- Resize Lems' array
    ReDim Preserve m_uLems(m_lLemsOut)
    
    '-- Define new Lem position and properties
    With m_uLems(m_lLemsOut)
        .x = m_uPtOut(m_lPtOut).x
        .y = m_uPtOut(m_lPtOut).y
        .xs = 1
        .Active = True
        Call wgSetLemAnimation(m_lLemsOut, [fFalling])
    End With
    
    '-- Next Lem, next door (if any)
    m_lPtOut = m_lPtOut + 1
    If (m_lPtOut > m_lcPtOut) Then
        m_lPtOut = 1
    End If
    
    '-- Show info
    fMain.ucInfo.Panels(3) = "Out: " & m_lLemsOut & "/" & g_uLevel.LemsToLetOut
End Sub

Public Sub SetArmageddonLem( _
           ByRef LemID As Long _
           )
    
    With m_uLems(LemID)
        
        '-- Activate 'Armageddon Lem'
        If (.Active) Then
            
            '-- Only if it's not already activated
            If (.ExplodeCount = 0) Then
                .Ability = .Ability Or [aBomber]
                .ExplodeCount = EXPLODE_TICKS
            End If
          
          Else
            '-- Not active: find next
            If (LemID < m_lLemsOut) Then
                LemID = LemID + 1
                Call SetArmageddonLem(LemID)
            End If
        End If
    End With
End Sub

Public Function HitTest( _
                ) As String
    
  Dim uPt As POINTAPI
    
    '-- Get cursor position (translate it to our game screen coordinates)
    Call GetCursorPos(uPt)
    Call ScreenToClient(fMain.ucScreen.hWnd, uPt)
    
    '-- Screen is 2x zoomed (also apply offset)
    m_xCur = uPt.x \ 2 + m_xScreen
    m_yCur = uPt.y \ 2
    
    '-- Hit-Test now (return Lem's description and number)
    HitTest = wgHitTest()
End Function

Public Sub PrepareAbility( _
           ByVal Ability As eLemAbility _
           )
    
    '-- Prepare ability to be applied (reset job)
    m_ePreparedAbility = Ability
    m_ePreparedJob = [jNone]
End Sub

Public Sub PrepareJob( _
           ByVal Job As eLemJob _
           )
    
    '-- Prepare job to be applied (reset ability)
    m_ePreparedJob = Job
    m_ePreparedAbility = [aNone]
End Sub

Public Sub ApplyPrepared()
     
  Dim nIdx As Integer
    
    If (IsTimerPaused = False) Then
    
        '-- Apply perpared ability/job
        
        If (m_lLem > 0) Then
            
            With m_uLems(m_lLem)
                
                If (.Frame <> [fExploding]) Then
                    
                    '-- Ability prepared
                    If (m_ePreparedAbility) Then
                        
                        '-- Set ability
                        .Ability = .Ability Or m_ePreparedAbility
                        If (m_ePreparedAbility = [aBomber]) Then
                            .ExplodeCount = EXPLODE_TICKS
                        End If
                        
                        '-- Update remaining
                        nIdx = (m_ePreparedAbility \ 2) + 1
                        With fMain
                            .lblButton(nIdx).Caption = Val(.lblButton(nIdx).Caption) - 1
                            If (Val(.lblButton(nIdx).Caption) = 0) Then
                                .ucToolbar.Buttons(nIdx).Value = [tbrUnpressed]
                                .ucToolbar.Buttons(nIdx).Enabled = False
                                m_ePreparedAbility = [aNone]
                            End If
                        End With
                        
                        '-- Play sound
                        Call PlaySoundFX([sfxMousePre])
                        
                    '-- Job prepared
                    ElseIf (m_ePreparedJob) Then
                        
                        '-- Can do job?
                        If (wgCanDoJobNow(m_lLem, m_ePreparedJob)) Then
                        
                            '-- Set job
                            .Job = m_ePreparedJob
                            Select Case .Job
                                Case [jNone]
                                    Call wgSetLemAnimation(m_lLem, [fWalker])
                                Case [jBlocker]
                                    Call wgSetLemAnimation(m_lLem, [fBlocker])
                                Case [jBuilder]
                                    Call wgSetLemAnimation(m_lLem, [fBuilder])
                                Case [jBasher]
                                    Call wgSetLemAnimation(m_lLem, [fBasher])
                                Case [jMiner]
                                    Call wgSetLemAnimation(m_lLem, [fMiner])
                                Case [jDigger]
                                    Call wgSetLemAnimation(m_lLem, [fDigger])
                            End Select
                            
                            '-- Update remaining
                            nIdx = .Job + 3
                            With fMain
                                .lblButton(nIdx).Caption = Val(.lblButton(nIdx).Caption) - 1
                                If (Val(.lblButton(nIdx).Caption) = 0) Then
                                    .ucToolbar.Buttons(nIdx).Value = [tbrUnpressed]
                                    .ucToolbar.Buttons(nIdx).Enabled = False
                                    m_ePreparedJob = [jNone]
                                End If
                            End With
                            
                            '-- Play sound
                            Call PlaySoundFX([sfxMousePre])
                        End If
                    End If
                End If
            End With
        End If
    End If
End Sub

Public Property Get GetLemsOut( _
                    ) As Long

    '-- All lems let out (+died)
    GetLemsOut = m_lLemsOut
End Property

Public Property Get GetLemsSaved( _
                    ) As Long
  
    '-- Get all saved lems
    GetLemsSaved = m_lLemsSaved
End Property

'========================================================================================
' Private
'========================================================================================

'----------------------------------------------------------------------------------------
' Game stage [0: 'Lets go', 1: 'Opening doors', 2: Playing, 3: 'Ending']
'----------------------------------------------------------------------------------------

Public Sub SetGameStage( _
           ByVal New_GameStage As eGameStage _
           )
    
    '-- Set current stage
    If (New_GameStage <> m_eGameStage) Then
        m_eGameStage = New_GameStage
        m_lcGameStage = 0
    End If
End Sub

Public Function GetGameStage( _
                ) As eGameStage
    
    '-- Return current stage
    GetGameStage = m_eGameStage
End Function

'----------------------------------------------------------------------------------------
' Rendering terrain and objects
'----------------------------------------------------------------------------------------

Private Sub wgDrawTerrain()

    '-- 'Render' terrain (load it from buffer)
    Call BltFast( _
         m_lpoScreen.DIB, _
         0, 0, 320, 160, _
         m_oDIBScreenBuffer, _
         m_xScreen, 0 _
         )
End Sub

Private Sub wgDrawObjects()
    
  Dim i As Long
    
    With g_uLevel
        '-- First draw 'on-terrain' objects
        For i = 1 To .Objects
            If (.Object(i).OnTerrain) Then
                Call wgDrawObject(i)
            End If
        Next i
        '-- Then, any other object
        For i = 1 To .Objects
            If (.Object(i).OnTerrain = False) Then
                Call wgDrawObject(i)
            End If
        Next i
    End With
End Sub

Private Sub wgDrawObject( _
            ByVal Idx As Integer _
            )
        
    With g_uLevel.Object(Idx)
        
        '-- Animate object?
        If ((IsTimerPaused = False) And .wgLoop) Then
            .wgFrameIdxCur = .wgFrameIdxCur + 1
            If (.wgFrameIdxCur = .wgFrameIdxMax) Then
                .wgFrameIdxCur = 0
                '-- Is a trap? (single loop)
                If (g_uObjGFX(.ID).TriggerEffect = IDX_TRAP) Then
                    .wgLoop = False
                End If
            End If
        End If
        
        '-- Rendering object
        If (.OnTerrain) Then
            Call MaskBltOverlapNot( _
                 m_lpoScreen.DIB, _
                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                 g_uObjGFX(.ID).Width, g_uObjGFX(.ID).Height, _
                 IDX_NONE, _
                 g_uObjGFX(.ID).DIB, _
                 0, 1& * .wgFrameIdxCur * g_uObjGFX(.ID).Height, _
                 IDX_TRANS _
                 )
          Else
            If (.NotOverlap) Then
                Call MaskBltOverlap( _
                     m_lpoScreen.DIB, _
                     -m_xScreen + .lpRect.x1, .lpRect.y1, _
                     g_uObjGFX(.ID).Width, g_uObjGFX(.ID).Height, _
                     IDX_NONE, _
                     g_uObjGFX(.ID).DIB, _
                     0, 1& * .wgFrameIdxCur * g_uObjGFX(.ID).Height, _
                     IDX_TRANS _
                     )
              Else
                Call MaskBltOverlapNot( _
                     m_lpoScreen.DIB, _
                     -m_xScreen + .lpRect.x1, .lpRect.y1, _
                     g_uObjGFX(.ID).Width, g_uObjGFX(.ID).Height, _
                     IDX_BRICK, _
                     g_uObjGFX(.ID).DIB, _
                     0, 1& * .wgFrameIdxCur * g_uObjGFX(.ID).Height, _
                     IDX_TRANS _
                     )
            End If
        End If
    End With
End Sub

'----------------------------------------------------------------------------------------
' Check all lems (most important routine... and longest)
'----------------------------------------------------------------------------------------

Private Sub wgCheckLems()

  Dim l As Long, bOneActive As Boolean
  Dim i As Long, j As Long, ix As Long, iy As Long, px As Long
  Dim bSkip As Boolean
    
    '-- Check all lems
    For l = 1 To m_lLemsOut
        
        With m_uLems(l)
            
            '-- Is active?
            If (.Active) Then
                
                '-- Reset 'skip code' flag
                bSkip = False
                
                '-- At least, one is active
                bOneActive = True
                
                '-- Next frame now
                .FrameIdx = .FrameIdx + 1
                If (.FrameIdx > .FrameIdxMax) Then
                    .FrameIdx = 0
                End If
                
                '-- Start checking...
                Select Case .Job
                    
                    Case [jNone]
                
                        Select Case .Frame
                        
                            Case [fWalker]
                                
                                '-- One step forward
                                .x = .x + .xs
                                
                                '-- Check feet
                                ix = 8 + (.xs < 0)
                                iy = 16
                                
                                If (wgCheckPixel(.x + ix, .y + iy, l, True)) Then

                                    '-- Can go up?
                                    For i = 0 To MIN_OBSTACLE - 1
                                        Select Case wgCheckPixel(.x + ix, .y + iy - 1, l, True)
                                            Case IDX_NONE
                                                Exit For
                                            Case IDX_NULL
                                                bSkip = True
                                                Exit For
                                            Case IDX_BLOCKER
                                                .xs = -.xs
                                                .x = .x + .xs
                                                .y = .y + i
                                                bSkip = True
                                                Exit For
                                            Case Else
                                                .y = .y - 1
                                        End Select
                                    Next i
                                    If (bSkip) Then GoTo lblSkip
                                        
                                    '-- Obstacle: can climb?
                                    If (i = MIN_OBSTACLE) Then
                                        .y = .y + i
                                        If (.Ability And [aClimber]) Then
                                            Call wgSetLemAnimation(l, [fClimber])
                                          Else
                                            .xs = -.xs
                                            .x = .x + .xs
                                        End If
                                    End If
                                  
                                  Else
                                    
                                    '-- Almost falling?
                                    For i = 0 To MIN_OBSTACLE - 1
                                        Select Case wgCheckPixel(.x + ix, .y + iy, l, True)
                                            Case IDX_NONE, IDX_BLOCKER
                                                .y = .y + 1
                                            Case IDX_NULL
                                                bSkip = True
                                                Exit For
                                            Case Else
                                                Exit For
                                        End Select
                                    Next i
                                    If (bSkip) Then GoTo lblSkip
                                    
                                    '-- Yes, just falling
                                    If (i = MIN_OBSTACLE) Then
                                        Call wgSetLemAnimation(l, [fFalling])
                                        .y = .y - i + 3
                                        .Counter = 3
                                        GoTo lblSkip
                                    End If
                                End If
                                
                                '-- Check ahead (blocker)
                                ix = 11 + 7 * (.xs < 0)
                                iy = 11
                                If (wgGetPixelLo(.x + .xs + ix, .y + iy, l) = IDX_BLOCKER) Then
                                    .xs = -.xs
                                    .x = .x + .xs
                                End If
                                
                            Case [fFalling]
                                
                                '-- 3-pixel loop
                                ix = 8 + (.xs < 0)
                                iy = 16
                                
                                For i = 1 To 3
                                    
                                    '-- One step down
                                    .y = .y + 1
                                    
                                    '-- Animation counter (frame)
                                    .Counter = .Counter + 1
                                    
                                    '-- Floater?
                                    If (.Counter >= MIN_FALL_FLOATER) Then
                                        If (.Ability And [aFloater]) Then
                                            Call wgSetLemAnimation(l, [fFloater])
                                            Exit For
                                        End If
                                    End If
                                    
                                    '-- A soft landing?
                                    Select Case wgCheckPixel(.x + ix, .y + iy, l, True)
                                        Case IDX_NULL
                                            Exit For
                                        Case IDX_NONE, IDX_BLOCKER
                                        Case Else
                                            If (.Counter > MAX_FALL + 1) Then
                                                Call wgSetLemAnimation(l, [fSpliting])
                                                Exit For
                                              Else
                                                Call wgSetLemAnimation(l, [fWalker])
                                                Exit For
                                            End If
                                    End Select
                                Next i
                                
                            Case [fSpliting]
                                                                      
                                Select Case .FrameIdx
                                    Case 1
                                        Call PlaySoundFX([sfxSplat])
                                    Case .FrameIdxMax
                                        .Active = False
                                End Select
                                
                            Case [fDrowning]
                                    
                                '-- Move?
                                ix = 13 + 11 * (.xs < 0)
                                iy = 16
                                If (wgGetPixelHi(.x + ix, .y + iy, l) = IDX_LIQUID) Then
                                    .x = .x + .xs
                                End If
                                 
                                Select Case .FrameIdx
                                    Case 1
                                        Call PlaySoundFX([sfxGlug])
                                    Case .FrameIdxMax
                                        .Active = False
                                End Select
                            
                            Case [fBurning]
                                
                                Select Case .FrameIdx
                                    Case 1
                                        Call PlaySoundFX([sfxFire])
                                    Case .FrameIdxMax
                                        .Active = False
                                End Select
                                
                            Case [fExploding]
                            
                                '-- Play sound
                                If (.FrameIdx = 11 And IsArmageddonActivated = False) Then
                                    Call PlaySoundFX([sfxOhNo])
                                End If
                                
                                '-- Check feet
                                iy = 16
                                For j = 1 To 3
                                    i = 0
                                    For ix = 8 To 9
                                        Select Case wgCheckPixel(.x + ix, .y + iy, l, True)
                                            Case IDX_NULL
                                                bSkip = True
                                                Exit For
                                            Case IDX_NONE, IDX_BLOCKER
                                                i = i + 1
                                        End Select
                                    Next ix
                                    If (bSkip) Then
                                        Exit For
                                      ElseIf (i = 2) Then
                                        .y = .y + 1
                                    End If
                                Next j
                                
                                '-- Almost...
                                If (.FrameIdx = .FrameIdxMax) Then
                                    
                                    '-- Explode
                                    Call PlaySoundFX([sfxExplode])
                                    Call wgDrawExplosion(l)
                                    Call wgDrawMask(l)
                                    
                                    '-- One less
                                    .Active = False
                                End If
                                
                            Case [fSurviving]
                                
                                Select Case .FrameIdx
                                    
                                    Case 1
                                    
                                        '-- Play sound
                                        Call PlaySoundFX([sfxYipee])
                                        
                                        '-- Save it now
                                        m_lLemsSaved = m_lLemsSaved + 1
                                        fMain.ucInfo.Panels(4) = "Saved: " & Format$(m_lLemsSaved / g_uLevel.LemsToLetOut, "0%")
                                 
                                    Case .FrameIdxMax
                                    
                                        '-- One less (but saved)
                                        .Active = False
                                End Select
                                
                            Case [fClimber]
                                
                                '-- Adjust animation (small obstacle)
                                If (.Counter = 0) Then
                                    
                                    '-- 4-pixel loop
                                    ix = 8 + (.xs < 0)
                                    iy = 7
                                    If (wgGetPixelLo(.x + ix, .y + iy, l) = IDX_NONE) Then
                                        For i = 0 To 4
                                            Select Case wgCheckPixel(.x + ix, .y + iy + i, l)
                                                Case IDX_NULL
                                                    bSkip = True
                                                    Exit For
                                                Case Is <> IDX_NONE
                                                    .y = .y - 1
                                            End Select
                                        Next i
                                        If (bSkip) Then
                                            GoTo lblSkip
                                          Else
                                            .y = .y + 3
                                        End If
                                        
                                        '-- Animation end
                                        Call wgSetLemAnimation(l, [fClimberEnd], 1)
                                        GoTo lblSkip
                                    End If
                                End If
                                
                                Select Case .FrameIdx
                                
                                    Case 4 To .FrameIdxMax
                                    
                                        '-- Climbing
                                        ix = 8 + (.xs < 0)
                                        iy = 6
                                        
                                        '-- Insurmountable obstacle?
                                        Select Case wgCheckPixel(.x + ix, .y + iy + i, l)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case IDX_NONE
                                                .y = .y - 1
                                                Call wgSetLemAnimation(l, [fClimberEnd])
                                                GoTo lblSkip
                                        End Select
                                        
                                        '-- Can't climb
                                        Select Case wgCheckPixel(.x + ix - .xs, .y + iy + i, l)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case Is <> IDX_NONE
                                                .x = .x - .xs
                                                .xs = -.xs
                                                Call wgSetLemAnimation(l, [fFalling])
                                                GoTo lblSkip
                                        End Select
                                        
                                        '-- Must fall?
                                        For iy = 9 To 13
                                            Select Case wgCheckPixel(.x + ix, .y + iy + i, l)
                                                Case IDX_NULL
                                                    bSkip = True
                                                    Exit For
                                                Case IDX_NONE
                                                    Call wgSetLemAnimation(l, [fFalling])
                                                    bSkip = True
                                                    Exit For
                                            End Select
                                        Next iy
                                        If (bSkip) Then GoTo lblSkip
                                        
                                        .y = .y - 1
                                End Select
                                
                                If (.FrameIdx = .FrameIdxMax) Then
                                    '-- Animation counter (loop)
                                    .Counter = .Counter + 1
                                End If
                            
                            Case [fClimberEnd]
                            
                                If (.FrameIdx = .FrameIdxMax) Then
                                    '-- Adjust animation
                                    .y = .y - 7
                                    '-- Walk again
                                    Call wgSetLemAnimation(l, [fWalker])
                                End If
                                
                            Case [fFloater]
                                
                                Select Case .FrameIdx
                                
                                    Case 0
                                    
                                        '-- Adjust animation
                                        .FrameIdx = 4
                                  
                                    Case 4 To 8
                                        
                                        '-- Decelerate when open
                                        If (.Counter = 0) Then
                                            .y = .y - 1
                                        End If
                                             
                                    Case .FrameIdxMax
                                    
                                        '-- Animation counter (loop)
                                        .Counter = .Counter + 1
                                End Select
                                
                                '-- A soft landing?
                                ix = 8 + (.xs < 0)
                                iy = 15
                                For i = 1 To 2
                                    .y = .y + 1
                                    Select Case wgCheckPixel(.x + ix, .y + iy, l, True)
                                        Case IDX_NULL
                                            Exit For
                                        Case Is <> IDX_NONE
                                            .y = .y - 1
                                            Call wgSetLemAnimation(l, [fWalker])
                                            Exit For
                                    End Select
                                Next i
                        End Select
                    
                    Case [jBlocker]
                        
                        '-- Draw mask
                        If (.Counter = 0) Then
                            .Counter = 1
                            Call wgDrawMask(l)
                        End If
                        
                        '-- 'Un-blocked'?
                        Select Case .FrameIdx
                        
                            Case 0, 4, 8, 12
                            
                                '-- Check feet line
                                iy = 16
                                i = 0
                                For ix = 7 To 10
                                    Select Case wgCheckPixel(.x + ix, .y + iy, l, True)
                                        Case IDX_NONE, IDX_BLOCKER
                                            i = i + 1
                                    End Select
                                Next ix
                                If (i = 4) Then
                                    .Job = [jNone]
                                    Call wgRemoveBlockerMask(l)
                                    Call wgSetLemAnimation(l, [fFalling])
                                End If
                        End Select

                    Case [jBuilder]
                        
                        Select Case .Frame
                        
                            Case [fBuilder]
                                
                                '-- Last three bricks?
                                If (.Counter >= MAX_BRICKS - 3) Then
                                    '-- Play sound
                                    If (.FrameIdx = 9) Then
                                        Call PlaySoundFX([sfxChink])
                                    End If
                                End If
                                
                                Select Case .FrameIdx
                                  
                                    Case 0
                                        
                                        '-- One brick up
                                        .x = .x + 2 * .xs
                                        .y = .y - 1
                                                                        
                                        '-- Can go up?
                                        If (.Counter = MAX_BRICKS) Then
                                            Call wgSetLemAnimation(l, [fBuilderEnd])
                                            GoTo lblSkip
                                        End If
                                        
                                        '-- Can continue (feet)?
                                        ix = 8 + (.xs < 0)
                                        iy = 15
                                        Select Case wgCheckPixel(.x + ix, .y + iy, l, True)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case IDX_BLOCKER
                                                .xs = -.xs
                                                GoTo lblSkip
                                            Case Is <> IDX_NONE
                                                .xs = -.xs
                                                .Job = [jNone]
                                                Call wgSetLemAnimation(l, [fWalker])
                                                GoTo lblSkip
                                        End Select
                                        
                                        '-- Can continue (ahead-blocker)?
                                        ix = 12 + 9 * (.xs < 0)
                                        iy = 11
                                        Select Case wgCheckPixel(.x + ix, .y + iy, l)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case IDX_BLOCKER
                                                .xs = -.xs
                                                GoTo lblSkip
                                        End Select
                                        
                                        '-- Can continue (ahead-all)?
                                        ix = 9 + 3 * (.xs < 0)
                                        iy = 11
                                        Select Case wgCheckPixel(.x + ix, .y + iy, l)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case IDX_BLOCKER
                                                .xs = -.xs
                                                GoTo lblSkip
                                            Case Is <> IDX_NONE
                                                .xs = -.xs
                                                .Job = [jNone]
                                                Call wgSetLemAnimation(l, [fWalker])
                                                GoTo lblSkip
                                        End Select
                                        
                                        '-- Can continue (head)?
                                        ix = 10 + 5 * (.xs < 0)
                                        iy = 7
                                        Select Case wgCheckPixel(.x + ix, .y + iy, l)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case Is <> IDX_NONE
                                                .xs = -.xs
                                                .Job = [jNone]
                                                Call wgSetLemAnimation(l, [fWalker])
                                        End Select
                                  
                                    Case 8
                                        Call wgDrawMask(l)
                                  
                                    Case .FrameIdxMax
                                        .Counter = .Counter + 1
                                End Select
                            
                            Case [fBuilderEnd]
                            
                                '-- Hey! Nothing to do?
                                If (.FrameIdx = .FrameIdxMax) Then
                                    .Job = [jNone]
                                    Call wgSetLemAnimation(l, [fWalker])
                                End If
                        End Select
                       
                    Case [jBasher]
                        
                        Select Case .FrameIdx
                          
                            Case 1 To 3
                            
                                '-- Draw mask now
                                Call wgDrawMask(l, .FrameIdx)
                          
                            Case 18 To 20
                            
                                '-- Draw mask now
                                Call wgDrawMask(l, .FrameIdx - 17)
                          
                            Case 4, 21
                            
                                '-- Can continue?
                                ix = 16 + 17 * (.xs < 0)
                                For iy = 11 To 15
                                    px = wgCheckPixel(.x + ix, .y + iy, l)
                                    If (px = IDX_NULL) Then
                                        Exit For
                                      Else
                                        If ((px = IDX_TERRAIN) Or _
                                            (px = IDX_BASHRIGHT And .xs > 0) Or _
                                            (px = IDX_BASHLEFT And .xs < 0)) Then
                                          Else
                                            .Job = [jNone]
                                            Call wgSetLemAnimation(l, [fWalker])
                                            Exit For
                                        End If
                                    End If
                                Next iy
                          
                            Case 11 To 15, 27 To 31
                            
                                '-- Adjust animation
                                .x = .x + .xs
                                
                                '-- Check feet (exit, trap...)
                                ix = 8 + (.xs < 0)
                                iy = 15
                                If (wgCheckPixel(.x + ix, .y + iy, l, True) = IDX_NULL) Then
                                    GoTo lblSkip
                                End If
                                
                                '-- Falling?
                                If (.Counter) Then
                                    ix = 8 + (.xs < 0)
                                    iy = 16
                                    Select Case wgCheckPixel(.x + ix, .y + iy, l)
                                        Case IDX_NULL
                                            GoTo lblSkip
                                        Case IDX_NONE, IDX_BLOCKER
                                            .Job = [jNone]
                                            Call wgSetLemAnimation(l, [fFalling])
                                            GoTo lblSkip
                                    End Select
                                End If
                                
                                '-- Animation counter
                                If (.FrameIdx = .FrameIdxMax) Then
                                    .Counter = .Counter + 1
                                End If
                        End Select
                        
                    Case [jMiner]
                        
                        Select Case .FrameIdx
                          
                            Case 0
                            
                                '-- Advance
                                .x = .x + .xs
                                .y = .y + 2
                            
                            Case 1
                                
                                '-- Check ahead (blocker)
                                iy = 8
                                For ix = 12 + 9 * (.xs < 0) To 15 + 15 * (.xs < 0) Step .xs
                                    If (wgGetPixelLo(.x + ix, .y + iy, l) = IDX_BLOCKER) Then
                                        .xs = -.xs
                                        Exit For
                                    End If
                                Next ix
                            
                                '-- Draw mask now
                                Call wgDrawMask(l)
                          
                            Case 3
                            
                                '-- Can continue (feet)?
                                ix = 11 + 7 * (.xs < 0)
                                iy = 16
                                px = wgCheckPixel(.x + ix, .y + iy, l)
                                If (px = IDX_NULL) Then
                                    GoTo lblSkip
                                  Else
                                    If ((px = IDX_TERRAIN) Or _
                                        (px = IDX_BASHRIGHT And .xs > 0) Or _
                                        (px = IDX_BASHLEFT And .xs < 0)) Then
                                      Else
                                        .Job = [jNone]
                                        Call wgSetLemAnimation(l, [fWalker])
                                        GoTo lblSkip
                                    End If
                                End If
                                
                                '-- Can continue (ahead)?
                                ix = 14 + 13 * (.xs < 0)
                                iy = 10
                                px = wgGetPixelLo(.x + ix, .y + iy, l)
                                If ((px = IDX_NONE) Or _
                                    (px = IDX_TERRAIN) Or _
                                    (px = IDX_BASHRIGHT And .xs > 0) Or _
                                    (px = IDX_BASHLEFT And .xs < 0)) Then
                                  Else
                                    .Job = [jNone]
                                    Call wgSetLemAnimation(l, [fWalker])
                                End If
                                
                            Case 6 To 13
                            
                                '-- Can continue (feet)?
                                ix = 9 + 3 * (.xs < 0)
                                iy = 15
                                px = wgCheckPixel(.x + ix, .y + iy, l, True)
                                If (px = IDX_NULL) Then
                                    GoTo lblSkip
                                  Else
                                    If ((px = IDX_TERRAIN) Or _
                                        (px = IDX_BASHRIGHT And .xs > 0) Or _
                                        (px = IDX_BASHLEFT And .xs < 0)) Then
                                      Else
                                        .Job = [jNone]
                                        Call wgSetLemAnimation(l, [fWalker])
                                        GoTo lblSkip
                                    End If
                                End If
                                
                            Case 14 To .FrameIdxMax
                                
                                If (.FrameIdx = 15) Then
                                    '-- Adjust animation
                                    .x = .x + 2 * .xs
                                End If
                                
                                '-- Can continue (feet)?
                                ix = 9 + 3 * (.xs < 0)
                                iy = 16
                                px = wgCheckPixel(.x + ix, .y + iy, l, True)
                                If (px = IDX_NULL) Then
                                    GoTo lblSkip
                                  Else
                                    If ((px = IDX_TERRAIN) Or _
                                        (px = IDX_BASHRIGHT And .xs > 0) Or _
                                        (px = IDX_BASHLEFT And .xs < 0)) Then
                                      Else
                                        .Job = [jNone]
                                        Call wgSetLemAnimation(l, [fWalker])
                                        GoTo lblSkip
                                    End If
                                End If
                        End Select
                    
                    Case [jDigger]
                            
                        Select Case .FrameIdx
                          
                            Case 1, 9
                            
                                '-- Check feet
                                iy = 14
                                For ix = 7 + (.xs < 0) To 9 + (.xs < 0)
                                     px = wgCheckPixel(.x + ix, .y + iy, l, True)
                                     If (px <> IDX_NONE) Then
                                         Exit For
                                     End If
                                Next ix
                                If (px = IDX_NULL) Then
                                    GoTo lblSkip
                                  Else
                                    If ((px = IDX_TERRAIN) Or _
                                        (px = IDX_BASHRIGHT) Or _
                                        (px = IDX_BASHLEFT)) Then
                                        '-- Draw mask now
                                        Call wgDrawMask(l)
                                        .y = .y + 1
                                      ElseIf (px = IDX_STEEL) Then
                                        .Job = [jNone]
                                        Call wgSetLemAnimation(l, [fWalker])
                                      Else
                                        '-- Falling?
                                        If (.Counter) Then
                                            .Job = [jNone]
                                            Call wgSetLemAnimation(l, [fFalling])
                                        End If
                                    End If
                                End If
                          
                            Case .FrameIdxMax
                            
                                '-- Animation counter (loop)
                                .Counter = .Counter + 1
                        End Select
                End Select
                
                '-- Bomber ability?
lblSkip:        If (.Ability And [aBomber]) Then
                    
                    '-- Already counting-down?
                    If (.ExplodeCount > 0) Then
                        
                        '-- Count-down
                        .ExplodeCount = .ExplodeCount - 1
                        Call wgDrawCountDown(l)
                      
                      Else
                        
                        '-- Remove ability and job
                        .Ability = .Ability And Not [aBomber]
                        .Job = [jNone]
                                        
                        '-- Particles rendering activated
                        .Particles = True
                        
                        '-- Initialize particles
                        For i = 0 To 24
                            .Particle(i).x = .x + 8 + (VBA.Rnd * 16 - 8)
                            .Particle(i).y = .y + 8 + (VBA.Rnd * 16 - 8)
                            .Particle(i).vx = VBA.Rnd * 16 - 8
                            .Particle(i).vy = -VBA.Rnd * 8 - 4
                        Next i
                        
                        '-- Avoid animation?
                        If (.Frame = [fFalling] Or _
                            .Frame = [fFloater] Or _
                            .Frame = [fClimber] Or _
                            .Frame = [fClimberEnd]) Then
                            
                            '-- Explode
                            Call wgSetLemAnimation(l, [fExploding])
                            Call PlaySoundFX([sfxExplode])
                            Call wgDrawExplosion(l)
                            Call wgDrawMask(l)
                            
                            '-- Sorry
                            .Active = False
                            
                          Else
                            '-- Pre-explosion animation
                            Call wgSetLemAnimation(l, [fExploding])
                        End If
                    End If
                End If
                
              Else
                    
                '-- Render particles?
                If (.Particles) Then
                    '-- Active in case remaining particles / Draw particles
                    bOneActive = True
                    Call wgDrawParticles(l)
                End If
            End If
        End With
    Next l
    
    '-- None active?
    If (bOneActive = False) Then
        If ((m_lLemsOut = g_uLevel.LemsToLetOut) Or IsArmageddonActivated) Then
            Call SetGameStage([gsEnding])
        End If
    End If
End Sub

'----------------------------------------------------------------------------------------
' Render all lems
'----------------------------------------------------------------------------------------

Private Sub wgDrawLems()
   
  Dim l    As Long
  Dim ySrc As Long
  
    For l = 1 To m_lLemsOut
        
        With m_uLems(l)
            
            If (.Active) Then
                
                '-- Reversed-direction offset
                If (.xs > 0 Or .FrameHasDir = False) Then
                    ySrc = .FrameSrcY
                  Else
                    ySrc = .FrameSrcY + 17
                End If
                
                '-- Draw frame
                Call MaskBlt( _
                     m_lpoScreen.DIB, _
                     -m_xScreen + .x, .y, _
                     16, 16, _
                     m_oDIBLems, _
                     .FrameIdx * 17, ySrc, _
                     IDX_TRANS _
                     )

                '-- Preserving a last frame (traps)
                If (.DieNextFrame) Then
                    .Active = False
                End If
            End If
        End With
    Next l
End Sub

'----------------------------------------------------------------------------------------
' Set current lem animation
'----------------------------------------------------------------------------------------

Private Sub wgSetLemAnimation( _
            ByVal LemID As Long, _
            ByVal Frame As eLemFrame, _
            Optional ByVal FrameIdx As Long = 0 _
            )
                                        
    With m_uLems(LemID)
    
        '-- Special cases...
        
        '   y offset...
        Select Case .Frame
            Case [fBuilder]
                If (.FrameIdx > 13) Then
                    .y = .y - 1
                End If
            Case [fMiner]
                If (Frame <> [fBasher]) Then
                    Select Case .FrameIdx
                        Case 0
                            .y = .y - 2
                        Case 1 To 8
                            .y = .y - 1
                    End Select
                End If
            Case [fDigger]
                If (Frame <> [fFalling]) Then
                    .y = .y - 2
                End If
        End Select
        
        '   x offset...
        Select Case .Frame
            Case [fBuilder]
                If (.FrameIdx > 13) Then
                    .x = .x + .xs
                End If
            Case [fBasher]
                .x = .x + .xs * m_aBasherData(.FrameIdx)
            Case [fMiner]
                If (.FrameIdx > 2) Then
                    .x = .x + .xs
                End If
            Case [fDigger]
                .x = .x + .xs
        End Select
        
        '-- Get frame (animation) data...
        .Frame = Frame
        .FrameIdx = FrameIdx
        .FrameSrcY = .Frame * 17
        .FrameOffY = m_uAnimationData(Frame).FrameOffY
        .FrameIdxMax = m_uAnimationData(Frame).FrameIdxMax
        .FrameHasDir = m_uAnimationData(Frame).FrameHasDir
        .Counter = 0
        
        '-- Offset?
        .y = .y + .FrameOffY
    End With
End Sub

'----------------------------------------------------------------------------------------
' Can lem do that job now?
'----------------------------------------------------------------------------------------

Private Function wgCanDoJobNow( _
                 ByVal LemID As Long, _
                 ByVal Job As eLemJob _
                 ) As Boolean
    
  Dim ix As Long
  Dim iy As Long
  Dim px As Long
    
    With m_uLems(LemID)
        
        If (.Frame <> [fFalling]) Then
        
            Select Case Job
            
                Case [jBuilder]
                
                    '-- Check ahead
                    ix = 8 + (.xs < 0)
                    iy = 11
                    px = wgGetPixelLo(.x + ix, .y + iy, LemID)
                    wgCanDoJobNow = (px = IDX_NONE)
                    
                    '-- Check feet
                    ix = 8 + (.xs < 0)
                    iy = 16
                    px = wgGetPixelLo(.x + ix, .y + iy, LemID)
                    wgCanDoJobNow = wgCanDoJobNow Or _
                                Not ( _
                                    (px = IDX_LIQUID) Or _
                                    (px = IDX_FIRE) Or _
                                    (px = IDX_TRAP) _
                                    )
            
                Case [jBasher]
                    
                    '-- Check ahead
                    ix = -1 - 17 * (.xs > 0)
                    iy = 16 - MIN_OBSTACLE
                    px = wgGetPixelLo(.x + ix, .y + iy, LemID)
                    wgCanDoJobNow = ( _
                                    (px = IDX_NONE) Or _
                                    (px = IDX_TERRAIN) Or _
                                    (px = IDX_BASHRIGHT And .xs > 0) Or _
                                    (px = IDX_BASHLEFT And .xs < 0) _
                                    )
                
                Case [jMiner]
                    
                    '-- Check feet
                    ix = 11 + 7 * (.xs < 0)
                    iy = 16
                    px = wgGetPixelLo(.x + ix, .y + iy, LemID)
                    wgCanDoJobNow = ( _
                                    (px = IDX_STEEL) Or _
                                    (px = IDX_BASHRIGHT And .xs < 0) Or _
                                    (px = IDX_BASHLEFT And .xs > 0) _
                                    )
                                      
                    '-- Check ahead
                    ix = 15 + 15 * (.xs < 0)
                    iy = 11
                    px = wgGetPixelLo(.x + ix, .y + iy, LemID)
                    wgCanDoJobNow = wgCanDoJobNow Or _
                                    ( _
                                    (px = IDX_STEEL) Or _
                                    (px = IDX_BASHRIGHT And .xs < 0) Or _
                                    (px = IDX_BASHLEFT And .xs > 0) _
                                    )
                    wgCanDoJobNow = Not wgCanDoJobNow
                    
                Case [jDigger]
                    
                    '-- Check feet
                    px = IDX_NONE
                    iy = 16
                    For ix = 7 + (.xs < 0) To 9 + (.xs < 0)
                        px = wgGetPixelLo(.x + ix, .y + iy, LemID)
                        If (px <> IDX_NONE) Then
                            Exit For
                        End If
                    Next ix
                    wgCanDoJobNow = ( _
                                    (px = IDX_TERRAIN) Or _
                                    (px = IDX_BASHRIGHT) Or _
                                    (px = IDX_BASHLEFT) _
                                    )
                
                Case Else
                    
                    '-- OK
                    wgCanDoJobNow = True
            End Select
        End If
    End With
End Function

'----------------------------------------------------------------------------------------
' Get pixel idx at (x,y)
'----------------------------------------------------------------------------------------

Private Function wgGetPixelLo( _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal LemID As Long _
                 ) As Byte
    
    With m_uScreenBKMaskRct
    
        '-- Valid coordinate?
        If (x >= .x1 And x < .x2 And _
            y >= .y1 And y < .y2 _
            ) Then
                
            '-- Return pixel lo-idx (terrain-related idx)
            wgGetPixelLo = m_aScreenBKMaskBits(x, y) And &HF
          
          Else
            '-- Where do you go?!
            If (y > MAX_YCHECK) Then
                m_uLems(LemID).Active = False
            End If
        End If
    End With
End Function

Private Function wgGetPixelHi( _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal LemID As Long _
                 ) As Byte
    
    With m_uScreenBKMaskRct
    
        '-- Valid coordinate?
        If (x >= .x1 And x < .x2 And _
            y >= .y1 And y < .y2 _
            ) Then
                
            '-- Return pixel hi-idx (trigger-related idx)
            wgGetPixelHi = m_aScreenBKMaskBits(x, y) And &HF0
          
          Else
            '-- Where do you go?!
            If (y > MAX_YCHECK) Then
                m_uLems(LemID).Active = False
            End If
        End If
    End With
End Function

'----------------------------------------------------------------------------------------
' Get (and check) pixel idx at (x,y)
'----------------------------------------------------------------------------------------

Private Function wgCheckPixel( _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal LemID As Long, _
                 Optional ByVal Feet As Boolean = False _
                 ) As Byte
    
    With m_uScreenBKMaskRct
    
        '-- Valid coordinate?
        If (x >= .x1 And x < .x2 And _
            y >= .y1 And y < .y2 _
            ) Then
        
            '-- Check trigger ID
            Select Case m_aScreenBKMaskBits(x, y) And &HF0
            
                Case IDX_EXIT
                    
                    m_uLems(LemID).Job = [jNone]
                    Call wgSetLemAnimation(LemID, [fSurviving])
                    wgCheckPixel = IDX_NULL
            
                Case IDX_TRAP
                    
                    If (Feet) Then
                        m_uLems(LemID).Job = [jNone]
                        If (m_uLems(LemID).DieNextFrame = False) Then
                            m_uLems(LemID).DieNextFrame = wgFindAndActivateTrap(x, y)
                        End If
                        wgCheckPixel = IDX_NULL
                    End If
            
                Case IDX_LIQUID
                    
                    m_uLems(LemID).Job = [jNone]
                    Call wgSetLemAnimation(LemID, [fDrowning])
                    wgCheckPixel = IDX_NULL
                
                Case IDX_FIRE
                
                    m_uLems(LemID).Job = [jNone]
                    Call wgSetLemAnimation(LemID, [fBurning])
                    wgCheckPixel = IDX_NULL
                    
                Case Else
                
                    '-- Return pixel terrain ID (hi-nibble = 0)
                    wgCheckPixel = m_aScreenBKMaskBits(x, y)
            End Select
          
          Else
            '-- Where do you go?!
            If (y > MAX_YCHECK) Then
                m_uLems(LemID).Active = False
            End If
        End If
    End With
End Function

'----------------------------------------------------------------------------------------
' Render mask (onto main buffer and back-mask)
'----------------------------------------------------------------------------------------

Private Sub wgDrawMask( _
            ByVal LemID As Long, _
            Optional ByVal lStep As Long = 0 _
            )
    
    With m_uLems(LemID)
    
        Select Case .Frame
        
            Case [fExploding]
                
                '-- Remove blocker mask always
                Call wgRemoveBlockerMask(LemID)
                
                '-- Steel under feet?
                If (wgGetPixelLo(.x + 7, .y + 16, LemID) = IDX_STEEL Or _
                    wgGetPixelLo(.x + 8, .y + 16, LemID) = IDX_STEEL _
                    ) Then
                  
                  Else
                  
                    '-- Draw exploding hole mask (on everything)
                    Call MaskBltIdxBkMask( _
                         m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                         .x, .y + 2, 16, 22, _
                         IDX_NULL, IDX_NONE, IDX_NONE, _
                         m_oDIBMask, _
                         0, 0, _
                         IDX_TRANS _
                         )
                End If
            
            Case [fBlocker]
          
                '-- Draw blocker mask
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                     .x, .y + 2, 16, 16, _
                     IDX_NONE, IDX_BLOCKER, IDX_NONE, _
                     m_oDIBMask, _
                     0, 74, _
                     IDX_TRANS _
                     )
          
            Case [fBuilder]

                '-- Draw brick
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                     .x, .y, 16, 16, _
                     IDX_NONE, IDX_TERRAIN, IDX_BRICK, _
                     m_oDIBMask, _
                     -(.xs < 0) * 17, 91, _
                     IDX_TRANS _
                     )
                     
            Case [fBasher]
          
                '-- Draw basher hole mask (sequence)
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                     .x, .y, 16, 7 + 3 * lStep, _
                     IDX_TERRAIN, IDX_NONE, IDX_NONE, _
                     m_oDIBMask, _
                     -(.xs < 0) * 17, 23, _
                     IDX_TRANS _
                     )
                If (.xs > 0) Then
                    Call MaskBltIdxBkMask( _
                         m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                         .x, .y, 16, 7 + 3 * lStep, _
                         IDX_BASHRIGHT, IDX_NONE, IDX_NONE, _
                         m_oDIBMask, _
                         0, 23, _
                         IDX_TRANS _
                         )
                  Else
                    Call MaskBltIdxBkMask( _
                         m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                         .x, .y, 16, 7 + 3 * lStep, _
                         IDX_BASHLEFT, IDX_NONE, IDX_NONE, _
                         m_oDIBMask, _
                         17, 23, _
                         IDX_TRANS _
                         )
                End If
            
            Case [fMiner]
          
                '-- Draw miner hole mask
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                     .x, .y, 16, 16, _
                     IDX_TERRAIN, IDX_NONE, IDX_NONE, _
                     m_oDIBMask, _
                     -(.xs < 0) * 17, 40, _
                     IDX_TRANS _
                     )
                If (.xs > 0) Then
                    Call MaskBltIdxBkMask( _
                         m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                         .x, .y, 16, 16, _
                         IDX_BASHRIGHT, IDX_NONE, IDX_NONE, _
                         m_oDIBMask, _
                         0, 40, _
                         IDX_TRANS _
                         )
                  Else
                    Call MaskBltIdxBkMask( _
                         m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                         .x, .y, 16, 16, _
                         IDX_BASHLEFT, IDX_NONE, IDX_NONE, _
                         m_oDIBMask, _
                         17, 40, _
                         IDX_TRANS _
                         )
                End If
                
            Case [fDigger]
                
                '-- Draw digger hole mask
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                     .x, .y, 16, 16, _
                     IDX_TERRAIN, IDX_NONE, IDX_NONE, _
                     m_oDIBMask, _
                     -(.xs < 0) * 17, 57, _
                     IDX_TRANS _
                     )
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                     .x, .y, 16, 16, _
                     IDX_BASHRIGHT, IDX_NONE, IDX_NONE, _
                     m_oDIBMask, _
                     -(.xs < 0) * 17, 57, _
                     IDX_TRANS _
                     )
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
                     .x, .y, 16, 16, _
                     IDX_BASHLEFT, IDX_NONE, IDX_NONE, _
                     m_oDIBMask, _
                     -(.xs < 0) * 17, 57, _
                     IDX_TRANS _
                     )
        End Select
    End With
End Sub

'----------------------------------------------------------------------------------------
' Render mask (special case: blocker un-blocked)
'----------------------------------------------------------------------------------------

Private Sub wgRemoveBlockerMask( _
            ByVal LemID As Long _
            )

    With m_uLems(LemID)
        '-- Remove blocker pixels from mask buffer
        Call MaskBltIdxBkMask( _
             m_oDIBScreenBKMask, m_oDIBScreenBuffer, _
             .x, .y + 2, 16, 16, _
             IDX_BLOCKER, IDX_NONE, IDX_NONE, _
             m_oDIBMask, _
             0, 74, _
             IDX_TRANS _
             )
    End With
End Sub

'----------------------------------------------------------------------------------------
' Paint count-down (directly onto screen: not persistent)
'----------------------------------------------------------------------------------------

Private Sub wgDrawCountDown( _
            ByVal LemID As Long _
            )
    
    With m_uLems(LemID)
        
        '-- Need to paint count-down?
        If (.Frame = [fBurning] Or _
            .Frame = [fDrowning] Or _
            .Frame = [fSpliting] Or _
            .Frame = [fSurviving]) Then
            
            '-- Not needed...
            .Ability = .Ability And Not aBomber
          
          Else
            
            '-- Draw count-down directly onto screen (not persistent)
            If (.Frame = [fFloater]) Then
                Call MaskBlt( _
                     m_lpoScreen.DIB, _
                     -m_xScreen + .x, .y - 6, 16, 5, _
                     m_oDIBMask, _
                     (.ExplodeCount \ 15) * 17, 108, _
                     IDX_TRANS _
                     )
              Else
                Call MaskBlt( _
                     m_lpoScreen.DIB, _
                     -m_xScreen + .x, .y - 0, 16, 5, _
                     m_oDIBMask, _
                     (.ExplodeCount \ 15) * 17, 108, _
                     IDX_TRANS _
                     )
            End If
        End If
    End With
End Sub

Private Sub wgDrawExplosion( _
            ByVal LemID As Long _
            )

    With m_uLems(LemID)
        Call MaskBlt( _
             m_lpoScreen.DIB, _
             -m_xScreen + .x - 5, .y - 6, 26, 32, _
             m_oDIBMask, _
             58, 0, _
             IDX_TRANS _
             )
    End With
End Sub

Private Sub wgDrawParticles( _
            ByVal LemID As Long _
            )
            
  Dim i As Long
  Dim c As Long
  
    With m_uLems(LemID)
    
        For i = 0 To 24
            
            With .Particle(i)
                
                '-- Render particle
                Call m_lpoScreen.DIB.SetPixelIdx(-m_xScreen + .x, .y, i \ 9 + 1)
                
                '-- Next position...
                .x = .x + .vx
                .y = .y + .vy
                .vy = .vy + 1
                
                '-- Out of screen particles count
                If (.y > MAX_YCHECK) Then
                    c = c + 1
                End If
            End With
        Next i
        
        '-- Still particles
        .Particles = (c < 25)
    End With
End Sub

'----------------------------------------------------------------------------------------
' Set doors state (open/closed: animation enabled/disabled)
'----------------------------------------------------------------------------------------

Private Sub wgSetDoorsState( _
            ByVal bOpening As Boolean _
            )
    
  Dim i As Long
    
    '-- Play sound
    If (bOpening) Then
        Call PlaySoundFX([sfxDoor])
    End If
    
    '-- Open/close all doors (start/stop animation)
    With g_uLevel
        For i = 1 To .Objects
            With .Object(i)
                '-- Door ID?
                If (.ID = DOOR_ID) Then
                    .wgLoop = bOpening
                End If
            End With
        Next i
    End With
End Sub

'----------------------------------------------------------------------------------------
' Find trap lem is just over and activate it
'----------------------------------------------------------------------------------------

Private Function wgFindAndActivateTrap( _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean
  
  Dim i As Long
    
    With g_uLevel
        
        For i = 1 To .Objects
            
            With .Object(i)
                
                '-- Traps' trigger effect?
                If (g_uObjGFX(.ID).TriggerEffect = IDX_TRAP And .wgLoop = False) Then
                    
                    '-- Our coordinate is in trap rectangle?
                    If (x >= .lpRect.x1 And x < .lpRect.x2 And _
                        y >= .lpRect.y1 And y < .lpRect.y2) Then
                        
                        '-- Start animation
                        .wgLoop = True
                        .wgFrameIdxCur = 0
                        
                        '-- Play FX
                        Call PlaySoundFX(g_uObjGFX(.ID).SoundEffect)
                        wgFindAndActivateTrap = True
                    End If
                End If
            End With
        Next i
    End With
End Function

'----------------------------------------------------------------------------------------
' Hit-test: check if we can apply prepared ability or job just now; also get description
'----------------------------------------------------------------------------------------

Private Function wgHitTest( _
                 ) As String
    
  Dim l      As Long
  Dim lLemID As Long
  Dim lMinx  As Long
  Dim lMiny  As Long
  Dim bSet   As Boolean
    
    '-- Don't check if nothing prepared
    If (m_ePreparedAbility Or m_ePreparedJob) Then
    
        '-- Minimum x/y distances to frame center (16x16 -> ~[8,8] rel.)
        lMinx = 16
        lMiny = 16
        
        '-- Check all lems
        For l = 1 To m_lLemsOut
           
            With m_uLems(l)
                
                '-- Only check active
                If (.Active) Then
                    
                    '-- Over a lem?
                    If (m_xCur >= .x And m_xCur < .x + 16 And _
                        m_yCur >= .y And m_yCur < .y + 16) Then
                        
                        '-- Reset
                        bSet = False
                    
                        '-- Is there a prepared ability? Check if we can apply
                        If (m_ePreparedAbility) Then
                            bSet = ((.Ability And m_ePreparedAbility) = [aNone])
                        
                        '-- Or, is there a prepared job? Check if we can apply
                        ElseIf (m_ePreparedJob) Then
                            If ((.Job <> m_ePreparedJob) Or (.Frame = [fBuilderEnd])) Then
                                bSet = (.Frame = [fWalker] Or _
                                        .Frame = [fBuilder] Or _
                                        .Frame = [fBuilderEnd] Or _
                                        .Frame = [fBasher] Or _
                                        .Frame = [fMiner] Or _
                                        .Frame = [fDigger] _
                                        )
                            End If
                        End If
                        
                        If (bSet) Then
                            '-- Nearest to center...
                            If (Abs(m_xCur - .x - 8) < lMinx) Then
                                lMinx = Abs(m_xCur - .x - 8)
                                lLemID = l
                            End If
                            If (Abs(m_yCur - .y - 8) < lMiny) Then
                                lMiny = Abs(m_yCur - .y - 8)
                                lLemID = l
                            End If
                        End If
                    End If
                End If
            End With
        Next l
        
        '-- One found
        If (lLemID > 0) Then
        
            '-- Something has been prepared: update cursor and return lem description
            Set m_lpoScreen.UserIcon = m_iCursorSelect
            m_lLem = lLemID
            wgHitTest = wgLemDescription(lLemID)
            Exit Function
        End If
    End If
                    
    '-- No lem found: update cursor and 'reset' current lem idx.
    Set m_lpoScreen.UserIcon = m_iCursorPointer
    m_lLem = -1
End Function

Private Function wgLemDescription( _
                 ByVal LemID As Long _
                 ) As String

    With m_uLems(LemID)
        
        '-- Return lem description depending on current lem frame
        Select Case .Frame
            Case [fFalling]
                wgLemDescription = "Faller"
            Case [fBlocker]
                wgLemDescription = "Blocker"
            Case [fBuilder]
                wgLemDescription = "Builder"
            Case [fDigger]
                wgLemDescription = "Digger"
            Case [fBasher]
                wgLemDescription = "Basher"
            Case [fMiner]
                wgLemDescription = "Miner"
            Case Else
                Select Case .Ability
                    Case [aClimber] Or [aFloater]
                        wgLemDescription = "Athlete"
                    Case [aClimber]
                        wgLemDescription = "Climber"
                    Case [aFloater]
                        wgLemDescription = "Floater"
                    Case Else
                        wgLemDescription = "Walker"
                End Select
        End Select
        
        '-- Also add lem idx.
        wgLemDescription = wgLemDescription & Space$(1) & m_lLem
    End With
End Function

'----------------------------------------------------------------------------------------
' Render preview (scroller view)
'----------------------------------------------------------------------------------------

Private Sub wgDrawPreview()
  
  Dim l As Long
    
    '-- Stretch our current buffer
    Call FXStretch(m_lpoPanView.DIB, m_oDIBScreenBuffer)
    
    '-- Normalize to half-grey
    Call FXNormalizeColor(m_lpoPanView.DIB, IDX_NONE, IDX_GREY128)
    
    '-- Draw lems
    If (m_eGameStage = [gsPlaying]) Then
        For l = 1 To m_lLemsOut
            With m_uLems(l)
                If (.Active) Then
                    Call BltFast( _
                         m_lpoPanView.DIB, _
                         (.x + 5) / 2.5 + 1, (.y + 15) \ 2 - 1, _
                         2, 2, _
                         m_oDIBLemPoint, _
                         0, 0 _
                         )
                End If
            End With
        Next l
    End If
    
    '-- Now, draw selection (visible area) rectangle
    With m_uPanViewSelectionRct
        Call FXLineV( _
             m_lpoPanView.DIB, _
             .x1, .y1, .y2, _
             IDX_YELLOW _
             )
        Call FXLineV( _
             m_lpoPanView.DIB, _
             .x2, .y1, .y2, _
             IDX_YELLOW _
             )
    End With
    
    '-- Refresh view
    Call m_lpoPanView.Refresh
End Sub

'----------------------------------------------------------------------------------------
' Mapping and unmapping DIBs
'----------------------------------------------------------------------------------------

Private Sub wgMapDIB(uSA As SAFEARRAY2D, aBits() As Byte, oDIB As cDIB08 _
            )

    With uSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB.BytesPerScanline
        .wgData = oDIB.lpBits
    End With
    Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(uSA), 4)
End Sub

Private Sub wgUnmapDIB(aBits() As Byte)

    Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
End Sub
