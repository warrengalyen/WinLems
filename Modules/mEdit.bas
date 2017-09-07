Attribute VB_Name = "mEdit"
'================================================
' Module:        mEdit.bas
' Author:        Warren Galyen
' Dependencies:  None
' Last revision: 06.30.2006
'================================================
Option Explicit

'-- A little bit of API

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const PS_SOLID As Long = 0

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'-- LemsEdit

Public Enum eSelection
    [eObject] = 0
    [eTerrain] = 1
    [eSteel] = 2
End Enum

Private Const MAX_SCROLL       As Long = 1280
Private Const MAX_POS          As Long = 1600
Private Const MAX_OFFSET       As Long = 1
Private Const VIEW_WIDTH       As Long = 320
Private Const VIEW_HEIGHT      As Long = 160

Private m_lpoScreen            As ucScreen32
Private m_lphDC                As Long
Private m_x                    As Integer
Private m_y                    As Integer
Private m_xScreen              As Integer

Private m_lScreenBackcolor     As Long
Private m_lScreenBackcolorRev  As Long
Private m_bShowObjects         As Boolean
Private m_bShowTerrain         As Boolean
Private m_bShowSteel           As Boolean
Private m_bShowTriggerAreas    As Boolean
Private m_bEnhanceSelected     As Boolean
Private m_bRedBlackPieces      As Boolean
Private m_bShowSelectionBox    As Boolean

Private m_eSelectionPreference As eSelection
Private m_eSelectionType       As eSelection
Private m_nSelectionIdx        As Integer
Private m_uSelectionRct        As RECT2



'========================================================================================
' Main initialization
'========================================================================================

Public Sub Initialize()
    
    '-- Short references...
    Set m_lpoScreen = fEdit.ucScreen
    m_lphDC = fEdit.ucScreen.DIB.hDC
    
    '-- Set default values
    m_lScreenBackcolor = &H10000
    m_lScreenBackcolorRev = &H1
    m_bShowSelectionBox = True
    m_bEnhanceSelected = True
    m_bShowObjects = True
    m_bShowTerrain = True
    m_bShowSteel = True
    m_bShowTriggerAreas = False
    m_bRedBlackPieces = False
    m_eSelectionPreference = [eObject]
    
    '-- Reset
    uLEVEL.GraphicSet = &HFF
End Sub

'========================================================================================
' Info initialization
'========================================================================================

Public Sub InitializeInfo()
    
    '-- Initialize and show info
    Call wgUpdateLevelInfo
    Call wgUpdateStatistics
    Call wgResetSelection
    Call wgUpdateSelectionInfo
End Sub

'========================================================================================
' Render screen
'========================================================================================

Public Sub DoFrame()
    
    '-- Reset (clear) buffer
    Call m_lpoScreen.DIB.Cls(m_lScreenBackcolor)
    
    '-- Render terrain pieces / objects / steel areas
    If (m_bShowTerrain) Then
        Call wgDrawTerrain
    End If
    If (m_bShowObjects) Then
        Call wgDrawObjects
    End If
    If (m_bShowSteel) Then
        Call wgDrawSteel
    End If
    If (m_bShowTriggerAreas) Then
        Call wgDrawTriggerAreas
    End If
    
    '-- Draw selection box
    If (m_bShowSelectionBox) Then
        Call wgDrawSelectionBox
    End If
    
    '-- Refresh from buffer
    Call m_lpoScreen.Refresh
    
    '-- Update info
    Call wgUpdateSelectionInfo
End Sub
 
'========================================================================================
' Scroll screen
'========================================================================================

Public Sub DoScrollTo( _
           ByVal x As Long _
           )
    
  Dim lxScreenPrev As Long
    
    lxScreenPrev = m_xScreen
    m_xScreen = x
    
    If (m_xScreen < 0) Then
        m_xScreen = 0
    ElseIf (m_xScreen > MAX_SCROLL) Then
        m_xScreen = MAX_SCROLL
    End If
    
    If (GetAsyncKeyState(vbKeySpace) < 0) Then
        Call SelectionMove(m_xScreen - lxScreenPrev, 0)
      Else
        Call DoFrame
    End If
End Sub

'========================================================================================
' Adding objects / terrain pieces / steel areas
'========================================================================================

Public Function AddObject( _
       ByVal ID As Integer _
       ) As Boolean
    
    If (uLEVEL.Objects < MAX_OBJECTS) Then
    
        With uLEVEL
        
            .Objects = .Objects + 1
            ReDim Preserve .Object(0 To .Objects)
            
            With .Object(.Objects)
                .ID = ID
                With .lpRect
                    .x1 = m_xScreen + (VIEW_WIDTH - uOBJGFX(ID).Width) \ 2
                    .y1 = (VIEW_HEIGHT - uOBJGFX(ID).Height) \ 2
                    .x2 = .x1 + uOBJGFX(ID).Width
                    .y2 = .y1 + uOBJGFX(ID).Height
                End With
                Call wgCopyRect(m_uSelectionRct, .lpRect)
                m_eSelectionType = [eObject]
            End With
            m_nSelectionIdx = .Objects
        End With
        AddObject = True
        
        Call wgUpdateStatistics
        Call wgUpdateSelectionFlags
        Call DoFrame
    End If
End Function

Public Function AddTerrainPiece( _
       ByVal ID As Integer _
       ) As Boolean
    
    If (uLEVEL.TerrainPieces < MAX_TERRAINPIECES) Then
    
        With uLEVEL
        
            .TerrainPieces = .TerrainPieces + 1
            ReDim Preserve .TerrainPiece(0 To .TerrainPieces)
            
            With .TerrainPiece(.TerrainPieces)
                .ID = ID
                With .lpRect
                    .x1 = m_xScreen + (VIEW_WIDTH - uTERGFX(ID).Width) \ 2
                    .y1 = (VIEW_HEIGHT - uTERGFX(ID).Height) \ 2
                    .x2 = .x1 + uTERGFX(ID).Width
                    .y2 = .y1 + uTERGFX(ID).Height
                End With
                Call wgCopyRect(m_uSelectionRct, .lpRect)
                m_eSelectionType = [eTerrain]
            End With
            m_nSelectionIdx = .TerrainPieces
        End With
        AddTerrainPiece = True
        
        Call wgUpdateStatistics
        Call wgUpdateSelectionFlags
        Call DoFrame
    End If
End Function

Public Function AddSteelArea( _
       ByVal Width As Integer, _
       ByVal Height As Integer _
       ) As Boolean

    If (uLEVEL.SteelAreas < MAX_STEELAREAS) Then
    
        If (Width >= MIN_STEELAREASIZE And Height >= MIN_STEELAREASIZE) Then
            
            With uLEVEL
                
                .SteelAreas = .SteelAreas + 1
                ReDim Preserve .SteelArea(0 To .SteelAreas)
            
                With .SteelArea(.SteelAreas)
                    With .lpRect
                        .x1 = m_xScreen + (VIEW_WIDTH - Width) \ 2
                        .y1 = (VIEW_HEIGHT - Height) \ 2
                        .x2 = .x1 + Width
                        .y2 = .y1 + Height
                    End With
                    Call wgCopyRect(m_uSelectionRct, .lpRect)
                    m_eSelectionType = [eSteel]
                End With
                m_nSelectionIdx = .SteelAreas
            End With
            AddSteelArea = True
        
            Call wgUpdateStatistics
            Call wgUpdateSelectionFlags
            Call DoFrame
        End If
    End If
End Function

Public Sub UpdateStatistics()
    
    Call wgUpdateStatistics
End Sub

'========================================================================================
' Steel area size
'========================================================================================

Public Sub SteelAreaSetWidth( _
           ByVal Width As Integer _
           )

    If (m_nSelectionIdx > 0 And Width >= MIN_STEELAREASIZE) Then
        With uLEVEL
            With .SteelArea(m_nSelectionIdx)
                With .lpRect
                    .x2 = .x1 + Width
                End With
                Call wgCopyRect(m_uSelectionRct, .lpRect)
            End With
        End With
    
        Call DoFrame
    End If
End Sub

Public Sub SteelAreaSetHeight( _
           ByVal Height As Integer _
           )

    If (m_nSelectionIdx > 0 And Height >= MIN_STEELAREASIZE) Then
        With uLEVEL
            With .SteelArea(m_nSelectionIdx)
                With .lpRect
                    .y2 = .y1 + Height
                End With
                Call wgCopyRect(m_uSelectionRct, .lpRect)
            End With
        End With
    
        Call DoFrame
    End If
End Sub

'========================================================================================
' Dragging selection
'========================================================================================

Public Sub MouseDown( _
           Button As Integer, _
           Shift As Integer, _
           x As Long, _
           y As Long _
           )
    
    '-- Store current position
    m_x = x
    m_y = y
    
    '-- Hit-test...
    If (wgHitTest(Button, x, y)) Then
        
        '-- Show context menu?
        If (Button = vbRightButton) Then
            Call fEdit.PopupMenu(fEdit.mnuContextSelectionTop)
        End If
    End If
End Sub

Public Sub MouseMove( _
           Button As Integer, _
           Shift As Integer, _
           x As Long, _
           y As Long _
           )
    
    If (Button = vbLeftButton) Then
        
        '-- Clip position
        If (x < 0) Then x = 0 Else If (x >= VIEW_WIDTH) Then x = VIEW_WIDTH
        If (y < 0) Then y = 0 Else If (y >= VIEW_HEIGHT) Then y = VIEW_HEIGHT
        
        '-- Scroll screen / move selection
        If (GetAsyncKeyState(vbKeyShift) < 0) Then
            fEdit.ucScroll.Value = fEdit.ucScroll.Value - (x - m_x)
          Else
            Call SelectionMove(x - m_x, y - m_y)
        End If
        
        '-- Store current position
        m_x = x
        m_y = y
        
      Else
        If (GetAsyncKeyState(vbKeyControl) < 0) Then
            Call wgHitTest(0, x, y)
        End If
    End If
End Sub

'========================================================================================
' Manipulating selection
'========================================================================================

Public Sub ResetSelection()
    
    Call wgResetSelection
    Call wgUpdateSelectionInfo
End Sub

Public Function SelectionExists( _
                ) As Boolean

    SelectionExists = (m_nSelectionIdx > 0)
End Function

Public Sub SelectionMove( _
           ByVal dx As Long, _
           ByVal dy As Long _
           )
    
    If (m_nSelectionIdx > 0) Then
        
        Select Case m_eSelectionType
            
            Case [eObject]
                With uLEVEL.Object(m_nSelectionIdx)
                    Call wgMoveRect(.lpRect, dx, dy)
                    Call wgCopyRect(m_uSelectionRct, .lpRect)
                End With
            
            Case [eTerrain]
                With uLEVEL.TerrainPiece(m_nSelectionIdx)
                    Call wgMoveRect(.lpRect, dx, dy)
                    Call wgCopyRect(m_uSelectionRct, .lpRect)
                End With
            
            Case [eSteel]
                With uLEVEL.SteelArea(m_nSelectionIdx)
                    Call wgMoveRect(.lpRect, dx, dy)
                    Call wgCopyRect(m_uSelectionRct, .lpRect)
                End With
        End Select
        
        Call DoFrame
    End If
End Sub

Public Function SelectionDuplicate( _
                ) As Boolean
    
    If (m_nSelectionIdx > 0) Then
    
        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    If (.Objects < MAX_OBJECTS) Then
                        
                        .Objects = .Objects + 1
                        ReDim Preserve .Object(0 To .Objects)
                        
                        .Object(.Objects) = .Object(m_nSelectionIdx)
                        m_nSelectionIdx = .Objects
                        
                        With .Object(.Objects)
                            Call wgMoveRect(.lpRect, 2, 2)
                            Call wgCopyRect(m_uSelectionRct, .lpRect)
                        End With
                        
                        SelectionDuplicate = True
                    End If
                        
                Case [eTerrain]
                
                    If (.TerrainPieces < MAX_TERRAINPIECES) Then
                        
                        .TerrainPieces = .TerrainPieces + 1
                        ReDim Preserve .TerrainPiece(0 To .TerrainPieces)
                        
                        .TerrainPiece(.TerrainPieces) = .TerrainPiece(m_nSelectionIdx)
                        m_nSelectionIdx = .TerrainPieces
                        
                        With .TerrainPiece(.TerrainPieces)
                            Call wgMoveRect(.lpRect, 2, 2)
                            Call wgCopyRect(m_uSelectionRct, .lpRect)
                        End With
                        
                        SelectionDuplicate = True
                    End If
                    
                Case [eSteel]
            
                    If (.SteelAreas < MAX_STEELAREAS) Then
                    
                        .SteelAreas = .SteelAreas + 1
                        ReDim Preserve .SteelArea(0 To .SteelAreas)
                        
                        .SteelArea(.SteelAreas) = .SteelArea(m_nSelectionIdx)
                        m_nSelectionIdx = .SteelAreas
                        
                        With .SteelArea(.SteelAreas)
                            Call wgMoveRect(.lpRect, 2, 2)
                            Call wgCopyRect(m_uSelectionRct, .lpRect)
                        End With
                        
                        SelectionDuplicate = True
                    End If
            End Select
        End With
        
        If (SelectionDuplicate) Then
            Call wgUpdateSelectionInfo
            Call wgUpdateStatistics
            Call DoFrame
        End If
        
      Else
        SelectionDuplicate = True
    End If
End Function

Public Sub SelectionRemove()
    
  Dim i As Long
  
    If (m_nSelectionIdx > 0) Then
    
        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    For i = m_nSelectionIdx To .Objects - 1
                        .Object(i) = .Object(i + 1)
                    Next i
                    .Objects = .Objects - 1
                    ReDim Preserve .Object(0 To .Objects)
                        
                Case [eTerrain]
                
                    For i = m_nSelectionIdx To .TerrainPieces - 1
                        .TerrainPiece(i) = .TerrainPiece(i + 1)
                    Next i
                    .TerrainPieces = .TerrainPieces - 1
                    ReDim Preserve .TerrainPiece(0 To .TerrainPieces)
                
                Case [eSteel]
            
                    For i = m_nSelectionIdx To .SteelAreas - 1
                        .SteelArea(i) = .SteelArea(i + 1)
                    Next i
                    .SteelAreas = .SteelAreas - 1
                    ReDim Preserve .SteelArea(0 To .SteelAreas)
            End Select
        End With
        
        Call wgResetSelection
        Call wgUpdateStatistics
        Call DoFrame
    End If
End Sub

Public Sub SelectionBringToTop()
    
  Dim uTmpObject       As uObject
  Dim uTmpTerrainPiece As uTerrainPiece
  Dim uTmpSteelArea    As uSteelArea
  Dim i                As Long
  
    If (m_nSelectionIdx > 0) Then
    
        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    uTmpObject = .Object(m_nSelectionIdx)
                    For i = m_nSelectionIdx To .Objects - 1
                        .Object(i) = .Object(i + 1)
                    Next i
                    .Object(.Objects) = uTmpObject
                    m_nSelectionIdx = .Objects
    
                Case [eTerrain]
                
                    uTmpTerrainPiece = .TerrainPiece(m_nSelectionIdx)
                    For i = m_nSelectionIdx To .TerrainPieces - 1
                        .TerrainPiece(i) = .TerrainPiece(i + 1)
                    Next i
                    .TerrainPiece(.TerrainPieces) = uTmpTerrainPiece
                    m_nSelectionIdx = .TerrainPieces
                
                Case [eSteel]
            
                    uTmpSteelArea = .SteelArea(m_nSelectionIdx)
                    For i = m_nSelectionIdx To .SteelAreas - 1
                        .SteelArea(i) = .SteelArea(i + 1)
                    Next i
                    .SteelArea(.SteelAreas) = uTmpSteelArea
                    m_nSelectionIdx = .SteelAreas
            End Select
        End With
        
        Call wgUpdateSelectionInfo
        Call DoFrame
    End If
End Sub

Public Sub SelectionBringToBottom()
    
  Dim uTmpObject       As uObject
  Dim uTmpTerrainPiece As uTerrainPiece
  Dim uTmpSteelArea    As uSteelArea
  Dim i                As Long
  
    If (m_nSelectionIdx > 0) Then
    
        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    uTmpObject = .Object(m_nSelectionIdx)
                    For i = m_nSelectionIdx To 2 Step -1
                        .Object(i) = .Object(i - 1)
                    Next i
                    .Object(1) = uTmpObject
                    m_nSelectionIdx = 1
                    
                Case [eTerrain]
                
                    uTmpTerrainPiece = .TerrainPiece(m_nSelectionIdx)
                    For i = m_nSelectionIdx To 2 Step -1
                        .TerrainPiece(i) = .TerrainPiece(i - 1)
                    Next i
                    .TerrainPiece(1) = uTmpTerrainPiece
                    m_nSelectionIdx = 1
                
                Case [eSteel]
            
                    uTmpSteelArea = .SteelArea(m_nSelectionIdx)
                    For i = m_nSelectionIdx To 2 Step -1
                        .SteelArea(i) = .SteelArea(i - 1)
                    Next i
                    .SteelArea(1) = uTmpSteelArea
                    m_nSelectionIdx = 1
            End Select
        End With
        
        Call wgUpdateSelectionInfo
        Call DoFrame
    End If
End Sub

Public Sub SelectionZOrderUp()
    
  Dim uTmpObject       As uObject
  Dim uTmpTerrainPiece As uTerrainPiece
  Dim uTmpSteelArea    As uSteelArea
  
    If (m_nSelectionIdx > 0) Then
        
        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    If (m_nSelectionIdx < .Objects) Then
                        uTmpObject = .Object(m_nSelectionIdx)
                        .Object(m_nSelectionIdx) = .Object(m_nSelectionIdx + 1)
                        .Object(m_nSelectionIdx + 1) = uTmpObject
                        m_nSelectionIdx = m_nSelectionIdx + 1
                    End If
                    
                Case [eTerrain]
    
                    If (m_nSelectionIdx < .TerrainPieces) Then
                        uTmpTerrainPiece = .TerrainPiece(m_nSelectionIdx)
                        .TerrainPiece(m_nSelectionIdx) = .TerrainPiece(m_nSelectionIdx + 1)
                        .TerrainPiece(m_nSelectionIdx + 1) = uTmpTerrainPiece
                        m_nSelectionIdx = m_nSelectionIdx + 1
                    End If
                
                Case [eSteel]
            
                    If (m_nSelectionIdx < .SteelAreas) Then
                        uTmpSteelArea = .SteelArea(m_nSelectionIdx)
                        .SteelArea(m_nSelectionIdx) = .SteelArea(m_nSelectionIdx + 1)
                        .SteelArea(m_nSelectionIdx + 1) = uTmpSteelArea
                        m_nSelectionIdx = m_nSelectionIdx + 1
                    End If
            End Select
        End With
        
        Call wgUpdateSelectionInfo
        Call DoFrame
    End If
End Sub

Public Sub SelectionZOrderDown()
    
  Dim uTmpObject       As uObject
  Dim uTmpTerrainPiece As uTerrainPiece
  Dim uTmpSteelArea    As uSteelArea
   
    If (m_nSelectionIdx > 0) Then

        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    If (m_nSelectionIdx > 1) Then
                        uTmpObject = .Object(m_nSelectionIdx)
                        .Object(m_nSelectionIdx) = .Object(m_nSelectionIdx - 1)
                        .Object(m_nSelectionIdx - 1) = uTmpObject
                        m_nSelectionIdx = m_nSelectionIdx - 1
                    End If
                    
                Case [eTerrain]
                
                    If (m_nSelectionIdx > 1) Then
                        uTmpTerrainPiece = .TerrainPiece(m_nSelectionIdx)
                        .TerrainPiece(m_nSelectionIdx) = .TerrainPiece(m_nSelectionIdx - 1)
                        .TerrainPiece(m_nSelectionIdx - 1) = uTmpTerrainPiece
                        m_nSelectionIdx = m_nSelectionIdx - 1
                    End If
                    
                Case [eSteel]
            
                    If (m_nSelectionIdx > 1) Then
                        uTmpSteelArea = .SteelArea(m_nSelectionIdx)
                        .SteelArea(m_nSelectionIdx) = .SteelArea(m_nSelectionIdx - 1)
                        .SteelArea(m_nSelectionIdx - 1) = uTmpSteelArea
                        m_nSelectionIdx = m_nSelectionIdx - 1
                    End If
            End Select
        End With
        
        Call wgUpdateSelectionInfo
        Call DoFrame
    End If
End Sub

Public Sub SelectionFindNextOver()
  
  Dim uPt As POINTAPI
  Dim i   As Long
  
    Call GetCursorPos(uPt)
    Call ScreenToClient(fEdit.ucScreen.hWnd, uPt)
    
    uPt.x = uPt.x \ 2 + m_xScreen
    uPt.y = uPt.y \ 2
    
    If (m_nSelectionIdx > 0) Then

        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    For i = m_nSelectionIdx + 1 To .Objects
                        With .Object(i)
                            If (wgPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call wgCopyRect(m_uSelectionRct, .lpRect)
                                m_nSelectionIdx = i
                                fEdit.cbObject.ListIndex = .ID
                                Exit For
                            End If
                        End With
                    Next i
                    
                Case [eTerrain]
                
                    For i = m_nSelectionIdx + 1 To .TerrainPieces
                        With .TerrainPiece(i)
                            If (wgPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call wgCopyRect(m_uSelectionRct, .lpRect)
                                m_nSelectionIdx = i
                                fEdit.cbTerrainPiece.ListIndex = .ID
                                Exit For
                            End If
                        End With
                    Next i
                    
                Case [eSteel]
            
                    For i = m_nSelectionIdx + 1 To .SteelAreas
                        With .SteelArea(i)
                            If (wgPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call wgCopyRect(m_uSelectionRct, .lpRect)
                                m_nSelectionIdx = i
                                Exit For
                            End If
                        End With
                    Next i
            End Select
        End With
        
        Call wgUpdateSelectionInfo
        Call wgUpdateSelectionFlags
    Call DoFrame
    End If
End Sub

Public Sub SelectionFindNextUnder()

  Dim uPt As POINTAPI
  Dim i   As Long
  
    Call GetCursorPos(uPt)
    Call ScreenToClient(fEdit.ucScreen.hWnd, uPt)
    
    uPt.x = uPt.x \ 2 + m_xScreen
    uPt.y = uPt.y \ 2
    
    If (m_nSelectionIdx > 0) Then

        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    For i = m_nSelectionIdx - 1 To 1 Step -1
                        With .Object(i)
                            If (wgPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call wgCopyRect(m_uSelectionRct, .lpRect)
                                fEdit.cbObject.ListIndex = .ID
                                m_nSelectionIdx = i
                                Exit For
                            End If
                        End With
                    Next i
                    
                Case [eTerrain]
                
                    For i = m_nSelectionIdx - 1 To 1 Step -1
                        With .TerrainPiece(i)
                            If (wgPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call wgCopyRect(m_uSelectionRct, .lpRect)
                                m_nSelectionIdx = i
                                fEdit.cbTerrainPiece.ListIndex = .ID
                                Exit For
                            End If
                        End With
                    Next i
                    
                Case [eSteel]
            
                    For i = m_nSelectionIdx - 1 To 1 Step -1
                        With .SteelArea(i)
                            If (wgPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call wgCopyRect(m_uSelectionRct, .lpRect)
                                m_nSelectionIdx = i
                                Exit For
                            End If
                        End With
                    Next i
            End Select
        End With
        
        Call wgUpdateSelectionInfo
        Call wgUpdateSelectionFlags
        Call DoFrame
    End If
End Sub

'========================================================================================
' Setting property values (preferences)
'========================================================================================

Public Property Let ShowSelectionBox(ByVal New_ShowSelectionBox As Boolean)
    
    m_bShowSelectionBox = New_ShowSelectionBox
    Call DoFrame
End Property

Public Property Let EnhanceSelected(ByVal New_EnhanceSelected As Boolean)
    
    m_bEnhanceSelected = New_EnhanceSelected
    Call DoFrame
End Property

Public Property Let ShowObjects(ByVal New_ShowObjects As Boolean)
    
    m_bShowObjects = New_ShowObjects
    Call wgResetSelection
    Call DoFrame
End Property

Public Property Let ShowTerrain(ByVal New_ShowTerrain As Boolean)
    
    m_bShowTerrain = New_ShowTerrain
    Call wgResetSelection
    Call DoFrame
End Property

Public Property Let ShowSteel(ByVal New_ShowSteel As Boolean)
    
    m_bShowSteel = New_ShowSteel
    Call DoFrame
End Property

Public Property Let ShowTriggerAreas(ByVal New_ShowTriggerAreas As Boolean)
    
    m_bShowTriggerAreas = New_ShowTriggerAreas
    Call DoFrame
End Property

Public Property Let RedBlackPieces(ByVal New_RedBlackPieces As Boolean)
    
    m_bRedBlackPieces = New_RedBlackPieces
    Call DoFrame
End Property

Public Property Let SelectionPreference(ByVal New_SelectionPreference As eSelection)
    
    m_eSelectionPreference = New_SelectionPreference
    Call wgResetSelection
    Call DoFrame
End Property

'========================================================================================
' Special flags
'========================================================================================

Public Sub SetNotOverlap( _
           ByVal New_NotOverlap As Boolean _
           )
    
    If (m_nSelectionIdx > 0) Then
        Select Case m_eSelectionType
            Case [eObject]
                uLEVEL.Object(m_nSelectionIdx).NotOverlap = New_NotOverlap
                Call wgUpdateSelectionFlags
                Call DoFrame
            Case [eTerrain]
                uLEVEL.TerrainPiece(m_nSelectionIdx).NotOverlap = New_NotOverlap
                Call wgUpdateSelectionFlags
                Call DoFrame
        End Select
    End If
End Sub

Public Sub ObjectSetOnTerrain( _
           ByVal New_OnTerrain As Boolean _
           )
    
    If (m_nSelectionIdx > 0) Then
        uLEVEL.Object(m_nSelectionIdx).OnTerrain = New_OnTerrain
        Call wgUpdateSelectionFlags
        Call DoFrame
    End If
End Sub

Public Sub TerrainPieceSetBlack( _
           ByVal New_Black As Boolean _
           )
    
    If (m_nSelectionIdx > 0) Then
        uLEVEL.TerrainPiece(m_nSelectionIdx).Black = New_Black
        Call wgUpdateSelectionFlags
        Call DoFrame
    End If
End Sub

Public Sub TerrainPieceSetUpsideDown( _
           ByVal New_UpsideDown As Boolean _
           )
    
    If (m_nSelectionIdx > 0) Then
        uLEVEL.TerrainPiece(m_nSelectionIdx).UpsideDown = New_UpsideDown
        Call wgUpdateSelectionFlags
        Call DoFrame
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Function wgHitTest( _
                 ByVal Button As Integer, _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean
    
    x = x + m_xScreen
    
    If (Button = vbRightButton And wgPtInRect(m_uSelectionRct, x, y)) Then
    
        wgHitTest = True
        
      Else
        
        m_nSelectionIdx = 0
        Call wgSetRectEmpty(m_uSelectionRct)
        
        Select Case m_eSelectionPreference
            
            Case [eObject]
                
                wgHitTest = wgHitTestObject(x, y)
                If (wgHitTest = False) Then
                    wgHitTest = wgHitTestTerrainPiece(x, y)
                    If (wgHitTest = False) Then
                        wgHitTest = wgHitTestSteelArea(x, y)
                    End If
                End If
                
            Case [eTerrain]
                
                wgHitTest = wgHitTestTerrainPiece(x, y)
                If (wgHitTest = False) Then
                    wgHitTest = wgHitTestObject(x, y)
                    If (wgHitTest = False) Then
                        wgHitTest = wgHitTestSteelArea(x, y)
                    End If
                End If
            
            Case [eSteel]
            
                wgHitTest = wgHitTestSteelArea(x, y)
                If (wgHitTest = False) Then
                    wgHitTest = wgHitTestObject(x, y)
                    If (wgHitTest = False) Then
                        wgHitTest = wgHitTestTerrainPiece(x, y)
                    End If
                End If
        End Select
        
        Call wgUpdateSelectionFlags
        Call DoFrame
    End If
End Function

Private Function wgHitTestObject( _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean
  
  Dim i As Long
                
    If (m_bShowObjects) Then
        For i = uLEVEL.Objects To 1 Step -1
            With uLEVEL.Object(i)
                If (wgPtInRect(.lpRect, x, y)) Then
'                    If (GetPixel(uOBJGFX(.ID).DIB.hDC, x - .lpRect.x1, y - .lpRect.y1) <> CLR_TRANS) Then
                        Call wgCopyRect(m_uSelectionRct, .lpRect)
                        m_nSelectionIdx = i
                        m_eSelectionType = [eObject]
                        fEdit.cbObject.ListIndex = .ID
                        wgHitTestObject = True
                        Exit For
'                    End If
                End If
            End With
        Next i
    End If
End Function

Private Function wgHitTestTerrainPiece( _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean
  
  Dim i  As Long
  Dim px As Long

    If (m_bShowTerrain) Then
        For i = uLEVEL.TerrainPieces To 1 Step -1
            With uLEVEL.TerrainPiece(i)
                If (wgPtInRect(.lpRect, x, y)) Then
                    If (.UpsideDown) Then
                        px = GetPixel(uTERGFX(.ID).DIB.hDC, x - .lpRect.x1, uTERGFX(.ID).Height - (y - .lpRect.y1) - 1)
                      Else
                        px = GetPixel(uTERGFX(.ID).DIB.hDC, x - .lpRect.x1, y - .lpRect.y1)
                    End If
                    If (px <> CLR_TRANS) Then
                        Call wgCopyRect(m_uSelectionRct, .lpRect)
                        m_nSelectionIdx = i
                        m_eSelectionType = [eTerrain]
                        fEdit.cbTerrainPiece.ListIndex = .ID
                        wgHitTestTerrainPiece = True
                        Exit For
                    End If
                End If
            End With
        Next i
    End If
End Function

Private Function wgHitTestSteelArea( _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean
  
  Dim i As Long

    If (m_bShowSteel) Then
        For i = uLEVEL.SteelAreas To 1 Step -1
            With uLEVEL.SteelArea(i)
                If (wgPtInRect(.lpRect, x, y)) Then
                    Call wgCopyRect(m_uSelectionRct, .lpRect)
                    m_nSelectionIdx = i
                    m_eSelectionType = [eSteel]
                    wgHitTestSteelArea = True
                    Exit For
                End If
            End With
        Next i
    End If
End Function

Private Function wgCanMoveRect( _
                 lpRect As RECT2, _
                 ByVal dx As Long, _
                 ByVal dy As Long _
                 ) As Boolean

    With lpRect
        If ((.x1 > MAX_POS - MAX_OFFSET) Or _
            (.x2 < MAX_OFFSET) Or _
            (.y1 > VIEW_HEIGHT - MAX_OFFSET) Or _
            (.y2 < MAX_OFFSET)) Then
            wgCanMoveRect = False
          Else
            wgCanMoveRect = True
        End If
    End With
End Function

Private Sub wgMoveRect( _
            lpRect As RECT2, _
            ByVal dx As Long, _
            ByVal dy As Long _
            )
    
    Call wgOffsetRect(lpRect, dx, dy)
    
    With lpRect
        If (.x1 > MAX_POS - MAX_OFFSET) Then
            Call wgOffsetRect(lpRect, (MAX_POS - MAX_OFFSET) - .x1, 0)
        End If
        If (.x2 < MAX_OFFSET) Then
            Call wgOffsetRect(lpRect, MAX_OFFSET - .x2, 0)
        End If
        If (.y1 > VIEW_HEIGHT - MAX_OFFSET) Then
            Call wgOffsetRect(lpRect, 0, (VIEW_HEIGHT - MAX_OFFSET) - .y1)
        End If
        If (.y2 < MAX_OFFSET) Then
            Call wgOffsetRect(lpRect, 0, MAX_OFFSET - .y2)
        End If
    End With
End Sub

Private Sub wgCopyRect( _
            lpDestRect As RECT2, _
            lpSourceRect As RECT2 _
            )

    With lpSourceRect
        lpDestRect.x1 = .x1
        lpDestRect.y1 = .y1
        lpDestRect.x2 = .x2
        lpDestRect.y2 = .y2
    End With
End Sub

Private Sub wgSetRectEmpty( _
            lpRect As RECT2 _
            )
    
    With lpRect
        .x2 = .x1
        .y2 = .y1
    End With
End Sub

Private Sub wgOffsetRect( _
            lpRect As RECT2, _
            ByVal x As Long, _
            ByVal y As Long _
            )

    With lpRect
        .x1 = .x1 + x
        .y1 = .y1 + y
        .x2 = .x2 + x
        .y2 = .y2 + y
    End With
End Sub

Private Function wgPtInRect( _
                 lpRect As RECT2, _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean

    With lpRect
        wgPtInRect = (x >= .x1 And x < .x2) And (y >= .y1 And y < .y2)
    End With
End Function

Private Function wgIsRectEmpty( _
                 lpRect As RECT2 _
                 ) As Boolean

    With lpRect
        wgIsRectEmpty = (.x1 = .x2) Or (.y1 = .y2)
    End With
End Function

Private Sub wgDrawTerrain()
 
  Dim i As Long
    
    For i = 1 To uLEVEL.TerrainPieces
        With uLEVEL.TerrainPiece(i)
            If (.lpRect.x1 >= m_xScreen And .lpRect.x1 < m_xScreen + VIEW_WIDTH) Or _
               (.lpRect.x2 >= m_xScreen And .lpRect.x2 < m_xScreen + VIEW_WIDTH) Then
                If (m_bEnhanceSelected And _
                   (m_eSelectionType = [eTerrain]) And _
                   (m_nSelectionIdx = i)) Then
                    If (.Black) Then
                        If (.NotOverlap) Then
                            Call mEditRenderer.MaskBltColorOverlap( _
                                 m_lpoScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 m_lScreenBackcolorRev, _
                                 CLR_LIGHTEN, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                          Else
                            Call mEditRenderer.MaskBltColor( _
                                 m_lpoScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 CLR_LIGHTEN, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                        End If
                      Else
                        If (.NotOverlap) Then
                            Call mEditRenderer.MaskBltLightenOverlap( _
                                 m_lpoScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 m_lScreenBackcolorRev, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                          Else
                            Call mEditRenderer.MaskBltLighten( _
                                 m_lpoScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                        End If
                    End If
                  Else
                    If (.Black) Then
                        If (m_bRedBlackPieces) Then
                            If (.NotOverlap) Then
                                Call mEditRenderer.MaskBltColorOverlap( _
                                     m_lpoScreen.DIB, _
                                     -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                     uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                     m_lScreenBackcolorRev, _
                                     CLR_RED, _
                                     uTERGFX(.ID).DIB, _
                                     0, 0, _
                                     CLR_TRANS, _
                                     .UpsideDown _
                                     )
                              Else
                                Call mEditRenderer.MaskBltColor( _
                                     m_lpoScreen.DIB, _
                                     -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                     uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                     CLR_RED, _
                                     uTERGFX(.ID).DIB, _
                                     0, 0, _
                                     CLR_TRANS, _
                                     .UpsideDown _
                                     )
                            End If
                          Else
                            If (.NotOverlap) Then
                                Call mEditRenderer.MaskBltColorOverlap( _
                                     m_lpoScreen.DIB, _
                                     -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                     uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                     m_lScreenBackcolorRev, m_lScreenBackcolorRev, _
                                     uTERGFX(.ID).DIB, _
                                     0, 0, _
                                     CLR_TRANS, _
                                     .UpsideDown _
                                     )
                              Else
                                Call mEditRenderer.MaskBltColor( _
                                     m_lpoScreen.DIB, _
                                     -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                     uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                     m_lScreenBackcolorRev, _
                                     uTERGFX(.ID).DIB, _
                                     0, 0, _
                                     CLR_TRANS, _
                                     .UpsideDown _
                                     )
                            End If
                        End If
                      Else
                        If (.NotOverlap) Then
                            Call mEditRenderer.MaskBltOverlap( _
                                 m_lpoScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 m_lScreenBackcolorRev, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                          Else
                            Call mEditRenderer.MaskBlt( _
                                 m_lpoScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                        End If
                    End If
                End If
            End If
        End With
    Next i
End Sub

Private Sub wgDrawObjects()

  Dim i As Long
    
    For i = 1 To uLEVEL.Objects
        If (uLEVEL.Object(i).OnTerrain) Then
            Call wgDrawObject(i)
        End If
    Next i
    For i = 1 To uLEVEL.Objects
        If (uLEVEL.Object(i).OnTerrain = False) Then
            Call wgDrawObject(i)
        End If
    Next i
End Sub

Private Sub wgDrawObject( _
            ByVal Idx As Integer _
            )
  
    With uLEVEL.Object(Idx)
        If (.lpRect.x1 >= m_xScreen And .lpRect.x1 < m_xScreen + VIEW_WIDTH) Or _
           (.lpRect.x2 >= m_xScreen And .lpRect.x2 < m_xScreen + VIEW_WIDTH) Then
            If (m_bEnhanceSelected And _
               (m_eSelectionType = [eObject]) And _
               (m_nSelectionIdx = Idx)) Then
                If (.OnTerrain) Then
                    Call mEditRenderer.MaskBltLightenOverlapNot( _
                         m_lpoScreen.DIB, _
                         -m_xScreen + .lpRect.x1, .lpRect.y1, _
                         uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                         m_lScreenBackcolorRev, _
                         uOBJGFX(.ID).DIB, _
                         0, 0, _
                         CLR_TRANS _
                         )
                  Else
                    If (.NotOverlap) Then
                        Call mEditRenderer.MaskBltLightenOverlap( _
                             m_lpoScreen.DIB, _
                             -m_xScreen + .lpRect.x1, .lpRect.y1, _
                             uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                             m_lScreenBackcolorRev, _
                             uOBJGFX(.ID).DIB, _
                             0, 0, _
                             CLR_TRANS _
                             )
                      Else
                        Call mEditRenderer.MaskBltLighten( _
                             m_lpoScreen.DIB, _
                             -m_xScreen + .lpRect.x1, .lpRect.y1, _
                             uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                             uOBJGFX(.ID).DIB, _
                             0, 0, _
                             CLR_TRANS _
                             )
                    End If
                End If
              Else
                If (.OnTerrain) Then
                    Call mEditRenderer.MaskBltOverlapNot( _
                         m_lpoScreen.DIB, _
                         -m_xScreen + .lpRect.x1, .lpRect.y1, _
                         uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                         m_lScreenBackcolorRev, _
                         uOBJGFX(.ID).DIB, _
                         0, 0, _
                         CLR_TRANS _
                         )
                  Else
                    If (.NotOverlap) Then
                        Call mEditRenderer.MaskBltOverlap( _
                             m_lpoScreen.DIB, _
                             -m_xScreen + .lpRect.x1, .lpRect.y1, _
                             uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                             m_lScreenBackcolorRev, _
                             uOBJGFX(.ID).DIB, _
                             0, 0, _
                             CLR_TRANS _
                             )
                      Else
                        Call mEditRenderer.MaskBlt( _
                             m_lpoScreen.DIB, _
                             -m_xScreen + .lpRect.x1, .lpRect.y1, _
                             uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                             uOBJGFX(.ID).DIB, _
                             0, 0, _
                             CLR_TRANS _
                             )
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub wgDrawSteel()

  Dim i    As Long
  Dim uRct As RECT2
  
    For i = 1 To uLEVEL.SteelAreas
        With uLEVEL.SteelArea(i)
            If (.lpRect.x1 >= m_xScreen And .lpRect.x1 < m_xScreen + VIEW_WIDTH) Or _
               (.lpRect.x2 >= m_xScreen And .lpRect.x2 < m_xScreen + VIEW_WIDTH) Then
                Call wgCopyRect(uRct, .lpRect)
                Call wgOffsetRect(uRct, -m_xScreen, 0)
                With uRct
                    Call mEditRenderer.MaskRectOr( _
                         m_lpoScreen.DIB, _
                         .x1, .y1, _
                         .x2 - .x1, .y2 - .y1, _
                         CLR_BLUE _
                         )
                End With
            End If
        End With
    Next i
End Sub
        
Private Sub wgDrawTriggerAreas()
  
  Dim i    As Long
  Dim uRct As RECT2
    
    For i = 1 To uLEVEL.Objects
        With uLEVEL.Object(i)
            If (uOBJGFX(.ID).TriggerEffect > 0) Then
                Call wgCopyRect(uRct, uOBJGFX(.ID).lpTriggerRect)
                Call wgOffsetRect(uRct, -m_xScreen + .lpRect.x1, .lpRect.y1)
                With uRct
                    Call mEditRenderer.MaskRectOr( _
                         m_lpoScreen.DIB, _
                         .x1, .y1, _
                         .x2 - .x1, .y2 - .y1, _
                         CLR_GREEN _
                         )
                End With
            End If
        End With
    Next i
End Sub
        
Private Sub wgResetSelection()
    
    m_nSelectionIdx = 0
    Call wgSetRectEmpty(m_uSelectionRct)
    Call wgUpdateSelectionFlags
End Sub

Private Sub wgDrawSelectionBox()

  Dim Clr     As Long
  Dim hPen    As Long
  Dim hOldPen As Long
  Dim uPt     As POINTAPI

    If (wgIsRectEmpty(m_uSelectionRct) = False) Then
        
        '-- Set color
        Select Case m_eSelectionType
            Case [eObject]
                Clr = vbGreen
            Case [eTerrain]
                Clr = vbYellow
            Case [eSteel]
                Clr = vbCyan
        End Select
            
        '-- Create selection box pen
        hPen = CreatePen(PS_SOLID, 1, Clr)
        hOldPen = SelectObject(m_lphDC, hPen)
        
        '-- Draw box
        With m_uSelectionRct
            Call MoveToEx(m_lphDC, -m_xScreen + .x1, .y1, uPt)
            Call LineTo(m_lphDC, -m_xScreen + .x2 - 1, .y1)
            Call LineTo(m_lphDC, -m_xScreen + .x2 - 1, .y2 - 1)
            Call LineTo(m_lphDC, -m_xScreen + .x1, .y2 - 1)
            Call LineTo(m_lphDC, -m_xScreen + .x1, .y1)
        End With
        
        '-- Unselect and destroy pen
        Call SelectObject(m_lphDC, hOldPen)
        Call DeleteObject(hPen)
    End If
End Sub

Private Sub wgUpdateLevelInfo()
    
    With uLEVEL
        fEdit.txtTitle = RTrim$(.Title)
        fEdit.txtLemsToLetOut = .LemsToLetOut
        fEdit.txtLemsToBeSaved = .LemsToBeSaved
        fEdit.txtReleaseRate = .ReleaseRate
        fEdit.txtPlayingTime = .PlayingTime
        fEdit.txtScreenStart = .ScreenStart
        fEdit.txtSkill(0) = .MaxClimbers
        fEdit.txtSkill(1) = .MaxFloaters
        fEdit.txtSkill(2) = .MaxBombers
        fEdit.txtSkill(3) = .MaxBlockers
        fEdit.txtSkill(4) = .MaxBuilders
        fEdit.txtSkill(5) = .MaxBashers
        fEdit.txtSkill(6) = .MaxMiners
        fEdit.txtSkill(7) = .MaxDiggers
    End With
End Sub

Private Sub wgUpdateSelectionInfo()
    
    If (m_nSelectionIdx > 0) Then
        With m_uSelectionRct
            fEdit.lblSelectionPositionVal.Caption = .x1 & "," & .y1
            fEdit.lblSelectionSizeVal.Caption = .x2 - .x1 & "x" & .y2 - .y1
            fEdit.lblzOrderVal.Caption = Format$(m_nSelectionIdx, "000")
            If (m_eSelectionType = [eSteel]) Then
                fEdit.txtSteelAreaWidth = .x2 - .x1
                fEdit.txtSteelAreaHeight = .y2 - .y1
            End If
        End With
      Else
        fEdit.lblSelectionPositionVal.Caption = vbNullString
        fEdit.lblSelectionSizeVal.Caption = vbNullString
        fEdit.lblzOrderVal.Caption = vbNullString
        fEdit.txtSteelAreaWidth = MIN_STEELAREASIZE
        fEdit.txtSteelAreaHeight = MIN_STEELAREASIZE
    End If
End Sub

Private Sub wgUpdateSelectionFlags()
    
    fEdit.cmdUp.Enabled = (m_nSelectionIdx > 0)
    fEdit.cmdDown.Enabled = (m_nSelectionIdx > 0)
    fEdit.cmdLeft.Enabled = (m_nSelectionIdx > 0)
    fEdit.cmdRight.Enabled = (m_nSelectionIdx > 0)
    
    fEdit.cmdZOrderUp.Enabled = (m_nSelectionIdx > 0)
    fEdit.cmdZOrderDown.Enabled = (m_nSelectionIdx > 0)
    
    fEdit.chkNotOverlap.Enabled = (m_nSelectionIdx > 0 And (m_eSelectionType = [eObject] Or m_eSelectionType = [eTerrain]))
    fEdit.chkOnTerrain.Enabled = (m_nSelectionIdx > 0 And m_eSelectionType = [eObject])
    fEdit.chkUpsideDown.Enabled = (m_nSelectionIdx > 0 And m_eSelectionType = [eTerrain])
    fEdit.chkBlack.Enabled = (m_nSelectionIdx > 0 And m_eSelectionType = [eTerrain])
    
    fEdit.mnuContextSelection(3).Enabled = (m_eSelectionType <> [eSteel])
    fEdit.mnuContextSelection(4).Enabled = (m_eSelectionType <> [eSteel])
    fEdit.mnuContextSelection(5).Visible = (m_eSelectionType <> [eSteel])
    fEdit.mnuContextSelection(6).Visible = (m_eSelectionType = [eObject])
    fEdit.mnuContextSelection(7).Visible = (m_eSelectionType = [eObject])
    fEdit.mnuContextSelection(8).Visible = False
    fEdit.mnuContextSelection(9).Visible = (m_eSelectionType = [eTerrain])
    fEdit.mnuContextSelection(10).Visible = (m_eSelectionType = [eTerrain])
    fEdit.mnuContextSelection(11).Visible = (m_eSelectionType = [eTerrain])
    
    With uLEVEL
        If (m_nSelectionIdx > 0) Then
            Select Case m_eSelectionType
                Case [eObject]
                    fEdit.chkNotOverlap = -.Object(m_nSelectionIdx).NotOverlap
                    fEdit.mnuContextSelection(6).Checked = .Object(m_nSelectionIdx).NotOverlap
                    fEdit.chkOnTerrain = -.Object(m_nSelectionIdx).OnTerrain
                    fEdit.mnuContextSelection(7).Checked = .Object(m_nSelectionIdx).OnTerrain
                Case [eTerrain]
                    fEdit.chkNotOverlap = -.TerrainPiece(m_nSelectionIdx).NotOverlap
                    fEdit.mnuContextSelection(9).Checked = .TerrainPiece(m_nSelectionIdx).NotOverlap
                    fEdit.chkBlack = -.TerrainPiece(m_nSelectionIdx).Black
                    fEdit.mnuContextSelection(10).Checked = .TerrainPiece(m_nSelectionIdx).Black
                    fEdit.chkUpsideDown = -.TerrainPiece(m_nSelectionIdx).UpsideDown
                    fEdit.mnuContextSelection(11).Checked = .TerrainPiece(m_nSelectionIdx).UpsideDown
            End Select
        End If
    End With
End Sub

Private Sub wgUpdateStatistics()
    
    With uLEVEL
        fEdit.ucPrgObjects.Value = .Objects
        fEdit.ucPrgObjects.Caption = "Object: " & .Objects & "/" & MAX_OBJECTS
        fEdit.ucPrgTerrainPieces.Value = .TerrainPieces
        fEdit.ucPrgTerrainPieces.Caption = "Terrain: " & .TerrainPieces & "/" & MAX_TERRAINPIECES
        fEdit.ucPrgSteelAreas.Value = .SteelAreas
        fEdit.ucPrgSteelAreas.Caption = "Steel: " & .SteelAreas & "/" & MAX_STEELAREAS
    End With
End Sub
