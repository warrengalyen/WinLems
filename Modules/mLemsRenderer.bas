Attribute VB_Name = "mLemsRenderer"
'================================================
' Module:        mLemsRenderer.bas
' Author:        Warren Galyen
' Dependencies:
' Last revision: 11.18.2006
'================================================

Option Explicit

'-- API:

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    wgData     As Long
    Bounds     As SAFEARRAYBOUND
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

'-- Masking related idxs.:

Public Const IDX_NONE      As Byte = 0
Public Const IDX_TRANS     As Byte = 254
Public Const IDX_NULL      As Byte = 255

'-- Base-palette related idxs.:

Public Const IDX_BLACK     As Byte = 0
Public Const IDX_BLUE      As Byte = 1
Public Const IDX_GREEN     As Byte = 2
Public Const IDX_LEM       As Byte = 3
Public Const IDX_YELLOW    As Byte = 4
Public Const IDX_RED       As Byte = 5
Public Const IDX_GREY128   As Byte = 6
Public Const IDX_BRICK     As Byte = 7

'-- Back-mask related idxs. (terrain and trigger areas layers):

Public Const IDX_TERRAIN   As Byte = 1  ' 0000 0001
Public Const IDX_BASHLEFT  As Byte = 2  ' 0000 0010 = trigger ID
Public Const IDX_BASHRIGHT As Byte = 3  ' 0000 0011 = trigger ID
Public Const IDX_STEEL     As Byte = 4  ' 0000 0100
Public Const IDX_BLOCKER   As Byte = 5  ' 0000 0101

Public Const IDX_EXIT      As Byte = 16 ' 0001 0000 = trigger ID
Public Const IDX_TRAP      As Byte = 32 ' 0010 0000 = trigger ID
Public Const IDX_LIQUID    As Byte = 48 ' 0011 0000 = trigger ID
Public Const IDX_FIRE      As Byte = 64 ' 0100 0000 = trigger ID

'-- Base and menu palettes:

Private Const PALETTE_BASE As String = "000000101038002C003C34343C3C003C0808202020"
Private Const PALETTE_MENU As String = "081020040C1C040000100800180800"

'-- Font related constants:

Private Const CHAR_SEP     As Long = 6
Private Const CHAR_WIDTH   As Long = 8
Private Const CHAR_HEIGHT  As Long = 8

'-- Private objects:

Private m_aPal(1023)       As Byte       ' 8-bit palette
Private m_oDIBFont         As New cDIB08 ' Font DIB
Private m_oDIBLem          As New cDIB08 ' Tiled bitmap: Lem
Private m_oDIBGround       As New cDIB08 ' Tiled bitmap: ground

Public Enum eTileBitmap                  ' Tiled bitmap enum.
    [tbLem] = 0
    [tbGround]
End Enum



'========================================================================================
' Methods
'========================================================================================

Public Sub InitializeLemsRenderer()

    '-- Build base palette
    Call MergePaletteEntries(PALETTE_BASE, 0)
    Call MergePaletteEntries(PALETTE_MENU, 240)
    
    '-- Load tile bitmaps
    Call m_oDIBLem.CreateFromBitmapFile( _
         AppPath & "RES\Lem.bmp" _
         )
    Call m_oDIBGround.CreateFromBitmapFile( _
         AppPath & "RES\Ground.bmp" _
         )
    
    '-- Load font bitmap
    Call m_oDIBFont.CreateFromBitmapFile( _
         AppPath & "RES\Font.bmp" _
         )
End Sub

Public Sub BltFast( _
           oDstDIB As cDIB08, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           oSrcDIB As cDIB08, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           Optional ByVal UpsideDown As Boolean = False _
           )
 
  Dim aDstBits() As Byte
  Dim aSrcBits() As Byte
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
  
    If (wgCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call wgMapDIB(uDstSA, aDstBits(), oDstDIB)
        Call wgMapDIB(uSrcSA, aSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                aDstBits(i + k2) = aSrcBits(i)
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call wgUnmapDIB(aDstBits())
        Call wgUnmapDIB(aSrcBits())
    End If
End Sub

Public Sub MaskBlt( _
           oDstDIB As cDIB08, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           oSrcDIB As cDIB08, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskIdx As Byte, _
           Optional ByVal UpsideDown As Boolean = False _
           )

  Dim aDstBits() As Byte
  Dim aSrcBits() As Byte
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
  
    If (wgCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call wgMapDIB(uDstSA, aDstBits(), oDstDIB)
        Call wgMapDIB(uSrcSA, aSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If (aSrcBits(i) <> SrcMaskIdx) Then
                    aDstBits(i + k2) = aSrcBits(i)
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call wgUnmapDIB(aDstBits())
        Call wgUnmapDIB(aSrcBits())
    End If
End Sub

Public Sub MaskBltOverlap( _
           oDstDIB As cDIB08, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstMaskIdx As Byte, _
           oSrcDIB As cDIB08, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskIdx As Byte, _
           Optional ByVal UpsideDown As Boolean = False _
           )

  Dim aDstBits() As Byte
  Dim aSrcBits() As Byte
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
  
    If (wgCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call wgMapDIB(uDstSA, aDstBits(), oDstDIB)
        Call wgMapDIB(uSrcSA, aSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If (aSrcBits(i) <> SrcMaskIdx) Then
                    If (aDstBits(i + k2) = DstMaskIdx) Then
                        aDstBits(i + k2) = aSrcBits(i)
                    End If
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call wgUnmapDIB(aDstBits())
        Call wgUnmapDIB(aSrcBits())
    End If
End Sub

Public Sub MaskBltOverlapNot( _
           oDstDIB As cDIB08, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstMaskIdx As Byte, _
           oSrcDIB As cDIB08, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskIdx As Byte, _
           Optional ByVal UpsideDown As Boolean = False _
           )

  Dim aDstBits() As Byte
  Dim aSrcBits() As Byte
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
  
    If (wgCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call wgMapDIB(uDstSA, aDstBits(), oDstDIB)
        Call wgMapDIB(uSrcSA, aSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If (aSrcBits(i) <> SrcMaskIdx) Then
                    If (aDstBits(i + k2) <> DstMaskIdx) Then
                        aDstBits(i + k2) = aSrcBits(i)
                    End If
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call wgUnmapDIB(aDstBits())
        Call wgUnmapDIB(aSrcBits())
    End If
End Sub

Public Sub MaskBltIdx( _
           oDstDIB As cDIB08, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstIdx As Byte, _
           oSrcDIB As cDIB08, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskIdx As Byte, _
           Optional ByVal UpsideDown As Boolean = False _
           )

  Dim aDstBits() As Byte
  Dim aSrcBits() As Byte
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
  
    If (wgCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call wgMapDIB(uDstSA, aDstBits(), oDstDIB)
        Call wgMapDIB(uSrcSA, aSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If (aSrcBits(i) <> SrcMaskIdx) Then
                    aDstBits(i + k2) = DstIdx
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call wgUnmapDIB(aDstBits())
        Call wgUnmapDIB(aSrcBits())
    End If
End Sub

Public Sub MaskBltIdxOverlap( _
           oDstDIB As cDIB08, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstMaskIdx As Byte, _
           ByVal DstIdx As Byte, _
           oSrcDIB As cDIB08, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskIdx As Byte, _
           Optional ByVal UpsideDown As Boolean = False _
           )

  Dim aDstBits() As Byte
  Dim aSrcBits() As Byte
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
  
    If (wgCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call wgMapDIB(uDstSA, aDstBits(), oDstDIB)
        Call wgMapDIB(uSrcSA, aSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If (aSrcBits(i) <> SrcMaskIdx) Then
                    If (aDstBits(i + k2) = DstMaskIdx) Then
                        aDstBits(i + k2) = DstIdx
                    End If
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call wgUnmapDIB(aDstBits())
        Call wgUnmapDIB(aSrcBits())
    End If
End Sub

Public Sub MaskRectIdx( _
           oDstDIB As cDIB08, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstIdx As Byte _
           )

  Dim aDstBits() As Byte
  Dim uDstSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, r As Long
  Dim j As Long, i2 As Long, j2 As Long
    
    If (wgCheckDIBRectRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        i1, i2, j1, j2, r _
        )) Then
        
        Call wgMapDIB(uDstSA, aDstBits(), oDstDIB)
        
        For j = j1 To j2
            For i = i1 To i1 + i2
                aDstBits(i) = DstIdx
            Next i
            i1 = i1 + r
        Next j
        
        Call wgUnmapDIB(aDstBits())
    End If
End Sub

Public Sub MaskRectIdxOverlap( _
           oDstDIB As cDIB08, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstOverlapIdx As Byte, _
           ByVal DstIdx As Byte _
           )

  Dim aDstBits() As Byte
  Dim uDstSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, r As Long
  Dim j As Long, i2 As Long, j2 As Long
    
    If (wgCheckDIBRectRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        i1, i2, j1, j2, r _
        )) Then
        
        Call wgMapDIB(uDstSA, aDstBits(), oDstDIB)
        
        For j = j1 To j2
            For i = i1 To i1 + i2
                If (aDstBits(i) = DstOverlapIdx) Then
                    aDstBits(i) = DstIdx
                End If
            Next i
            i1 = i1 + r
        Next j
        
        Call wgUnmapDIB(aDstBits())
    End If
End Sub

Public Sub MaskRectIdxBkMask( _
           oDstDIB As cDIB08, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstIdxAdd As Byte _
           )

  Dim aDstBits() As Byte
  Dim uDstSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, r As Long
  Dim j As Long, i2 As Long, j2 As Long
    
    If (wgCheckDIBRectRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        i1, i2, j1, j2, r _
        )) Then
        
        Call wgMapDIB(uDstSA, aDstBits(), oDstDIB)
        
        For j = j1 To j2
            For i = i1 To i1 + i2
                aDstBits(i) = aDstBits(i) Or DstIdxAdd
            Next i
            i1 = i1 + r
        Next j
        
        Call wgUnmapDIB(aDstBits())
    End If
End Sub

Public Sub MaskBltIdxBkMask( _
           oDstDIBBkMask As cDIB08, _
           oDstDIBBuffer As cDIB08, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstMaskIdxBkMask As Byte, _
           ByVal DstIdxBkMask As Byte, _
           ByVal DstIdxBuffer As Byte, _
           oSrcDIB As cDIB08, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskIdx As Byte, _
           Optional ByVal UpsideDown As Boolean = False _
           )

  Dim aDstBitsBuffer() As Byte
  Dim aDstBitsBkMask() As Byte
  Dim aSrcBits()       As Byte
  Dim uDstBufferSA     As SAFEARRAY1D
  Dim uDstBkMaskSA     As SAFEARRAY1D
  Dim uSrcSA           As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
      
    If (wgCheckDIBDIBRects( _
        oDstDIBBuffer, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call wgMapDIB(uDstBkMaskSA, aDstBitsBkMask(), oDstDIBBkMask)
        Call wgMapDIB(uDstBufferSA, aDstBitsBuffer(), oDstDIBBuffer)
        Call wgMapDIB(uSrcSA, aSrcBits(), oSrcDIB)
        
        If (DstMaskIdxBkMask = IDX_NULL) Then
            Select Case DstIdxBkMask
                Case IDX_NONE
                    For j = j1 To j2
                        k2 = k1 - i1
                        For i = i1 To i1 + i2
                            If (aSrcBits(i) <> SrcMaskIdx) Then
                                aDstBitsBkMask(i + k2) = aDstBitsBkMask(i + k2) And &HF0
                                aDstBitsBuffer(i + k2) = DstIdxBuffer
                            End If
                        Next i
                        i1 = i1 + r2
                        k1 = k1 + r1
                    Next j
                Case Else
                    For j = j1 To j2
                        k2 = k1 - i1
                        For i = i1 To i1 + i2
                            If (aSrcBits(i) <> SrcMaskIdx) Then
                                aDstBitsBkMask(i + k2) = aDstBitsBkMask(i + k2) And &HF0 Or DstIdxBkMask
                                aDstBitsBuffer(i + k2) = DstIdxBuffer
                            End If
                        Next i
                        i1 = i1 + r2
                        k1 = k1 + r1
                    Next j
            End Select
          Else
            Select Case DstIdxBkMask
                Case IDX_NONE
                    For j = j1 To j2
                        k2 = k1 - i1
                        For i = i1 To i1 + i2
                            If (aSrcBits(i) <> SrcMaskIdx) Then
                                If ((aDstBitsBkMask(i + k2) And &HF) = DstMaskIdxBkMask) Then
                                     aDstBitsBkMask(i + k2) = aDstBitsBkMask(i + k2) And &HF0
                                     aDstBitsBuffer(i + k2) = DstIdxBuffer
                                End If
                            End If
                        Next i
                        i1 = i1 + r2
                        k1 = k1 + r1
                    Next j
                Case Else
                    For j = j1 To j2
                        k2 = k1 - i1
                        For i = i1 To i1 + i2
                            If (aSrcBits(i) <> SrcMaskIdx) Then
                                If ((aDstBitsBkMask(i + k2) And &HF) = DstMaskIdxBkMask) Then
                                     aDstBitsBkMask(i + k2) = aDstBitsBkMask(i + k2) And &HF0 Or DstIdxBkMask
                                     aDstBitsBuffer(i + k2) = DstIdxBuffer
                                End If
                            End If
                        Next i
                        i1 = i1 + r2
                        k1 = k1 + r1
                    Next j
            End Select
        End If
        
        Call wgUnmapDIB(aDstBitsBkMask())
        Call wgUnmapDIB(aDstBitsBuffer())
        Call wgUnmapDIB(aSrcBits())
    End If
End Sub

Public Sub FXTile( _
           oDstDIB As cDIB08, _
           ByVal x As Long, ByVal y As Long, _
           ByVal Width As Long, ByVal Height As Long, _
           ByVal TileBitmap As eTileBitmap _
           )

  Dim oSrcDIB As cDIB08
  Dim W As Long
  Dim H As Long
  Dim i As Long, i1 As Long, i2 As Long
  Dim j As Long, j1 As Long, j2 As Long
 
    Select Case TileBitmap
        Case [tbLem]
            Set oSrcDIB = m_oDIBLem
        Case [tbGround]
            Set oSrcDIB = m_oDIBGround
    End Select
  
    W = oSrcDIB.Width
    H = oSrcDIB.Height
    
    i1 = x
    j1 = y
    i2 = x + Width - 1
    j2 = y + Height - 1
        
    For j = j1 To j2 Step H
        For i = i1 To i2 Step W
            Call BltFast( _
                 oDstDIB, i, j, W, H, _
                 oSrcDIB, 0, 0 _
                 )
        Next i
    Next j
End Sub

Public Sub FXStretch( _
           oDstDIB As cDIB08, _
           oSrcDIB As cDIB08 _
           )
           
  Dim aDstBits() As Byte
  Dim aSrcBits() As Byte
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D

  Dim cX    As Single
  Dim cY    As Single
  Dim xLU() As Long
  Dim yLU() As Long
  
  Dim i  As Long, W  As Long
  Dim j  As Long, H  As Long
  Dim r1 As Long, r2 As Long
  Dim po As Long, pn As Long, qn As Long
  
    cX = oSrcDIB.Width / oDstDIB.Width
    cY = oSrcDIB.Height / oDstDIB.Height
    
    r1 = oDstDIB.BytesPerScanline
    r2 = oSrcDIB.BytesPerScanline
    
    W = oDstDIB.Width - 1
    H = oDstDIB.Height - 1

    ReDim xLU(W)
    For i = 0 To W
        xLU(i) = Int(i * cX)
    Next i
    ReDim yLU(H)
    For i = 0 To H
        yLU(i) = Int(i * cY) * r2
    Next i
    
    Call wgMapDIB(uDstSA, aDstBits(), oDstDIB)
    Call wgMapDIB(uSrcSA, aSrcBits(), oSrcDIB)
    
    For j = 0 To H
        po = yLU(j)
        qn = pn
        For i = 0 To W
            aDstBits(qn) = aSrcBits(po + xLU(i))
            qn = qn + 1
        Next i
        pn = pn + r1
    Next j
    
    Call wgUnmapDIB(aDstBits())
    Call wgUnmapDIB(aSrcBits())
End Sub

Public Sub FXNormalizeColor( _
           oDIB As cDIB08, _
           ByVal MaskIdx As Byte, _
           ByVal DstIdx As Byte _
           )

  Dim aBits() As Byte
  Dim uSA     As SAFEARRAY1D
  Dim i       As Long
  Dim j       As Long
  Dim W       As Long
  Dim H       As Long
  Dim r       As Long
  Dim p       As Long
  
    r = oDIB.BytesPerScanline
    W = oDIB.Width - 1
    H = oDIB.Height - 1
    
    Call wgMapDIB(uSA, aBits(), oDIB)
    
    For j = 0 To H
        For i = p To p + W
            If (aBits(i) <> MaskIdx) Then
                aBits(i) = DstIdx
            End If
        Next i
        p = p + r
    Next j
        
    Call wgUnmapDIB(aBits())
End Sub

Public Sub FXText( _
           oDIB As cDIB08, _
           ByVal x As Long, _
           ByVal y As Long, _
           ByVal Text As String, _
           ByVal DstIdx As Byte _
           )

  Dim c As Long
  
    For c = 1 To Len(Text)
    
        Call MaskBltIdx( _
             oDIB, _
             x + CHAR_SEP * (c - 1), y, _
             CHAR_WIDTH, CHAR_HEIGHT, _
             DstIdx, _
             m_oDIBFont, _
             CHAR_WIDTH * (Asc(Mid$(Text, c, 1)) - 32), 0, _
             IDX_TRANS _
             )
    Next c
End Sub

Public Sub FXLineV( _
           oDIB As cDIB08, _
           ByVal x1 As Long, _
           ByVal y1 As Long, _
           ByVal y2 As Long, _
           ByVal DstIdx As Byte _
           )
           
  Dim aBits() As Byte
  Dim uSA     As SAFEARRAY1D
  Dim p       As Long
  Dim r       As Long
  Dim W       As Long
  Dim H       As Long
  
    r = oDIB.BytesPerScanline
    W = oDIB.Width - 1
    H = oDIB.Height - 1
    
    If (x1 < 0) Then x1 = 0 Else If (x1 > W) Then x1 = W
    If (y1 < 0) Then y1 = 0 Else If (y1 > H) Then y1 = H
    If (y2 < 0) Then y2 = 0 Else If (y2 > H) Then y2 = H
    
    Call wgMapDIB(uSA, aBits(), oDIB)
    
    For p = y1 * r + x1 To _
            y2 * r + x1 Step _
            Sgn(y2 - y1) * r
            
        aBits(p) = DstIdx
    Next p
    
    Call wgUnmapDIB(aBits())
End Sub

Public Function GetGlobalPalette( _
                ) As Byte()

    GetGlobalPalette = m_aPal()
End Function

Public Function GetFadedOutGlobalPalette( _
                Optional ByVal Amount As Byte = 1 _
                ) As Byte()

  Dim aPal(1023) As Byte
  Dim i          As Long
    '-- 16 colors only
    Call CopyMemory(aPal(0), m_aPal(0), 64)
    For i = 0 To 63
        If (aPal(i) > Amount) Then
            aPal(i) = aPal(i) - Amount
          Else
            aPal(i) = 0
        End If
    Next i
    
    GetFadedOutGlobalPalette = aPal()
End Function

Public Sub MergePaletteEntries( _
           ByVal HEXStream As String, _
           ByVal EntryStart As Byte _
           )
    
  Dim e  As Long
  Dim p  As Long
  Dim e1 As Long
  Dim e2 As Long
    
    e1 = 4& * EntryStart
    e2 = 4& * EntryStart + 4 * Len(HEXStream) \ 6 - 4
    p = 0
    For e = e1 To e2 Step 4
        m_aPal(e + 2) = 4 * CByte("&H" & Mid$(HEXStream, p * 6 + 1, 2))
        m_aPal(e + 1) = 4 * CByte("&H" & Mid$(HEXStream, p * 6 + 3, 2))
        m_aPal(e + 0) = 4 * CByte("&H" & Mid$(HEXStream, p * 6 + 5, 2))
        p = p + 1
    Next e
End Sub

'========================================================================================
' Private
'========================================================================================

Private Function wgCheckDIBDIBRects( _
                 oDstDIB As cDIB08, _
                 ByVal xDst As Long, ByVal yDst As Long, _
                 ByVal wDst As Long, ByVal hDst As Long, _
                 oSrcDIB As cDIB08, _
                 ByVal xSrc As Long, ByVal ySrc As Long, _
                 ByVal UpsideDown As Boolean, _
                 i1 As Long, i2 As Long, _
                 j1 As Long, j2 As Long, _
                 k1 As Long, k2 As Long, _
                 r1 As Long, r2 As Long _
                 ) As Boolean

  Dim uDstRct As RECT
  Dim uSrcRct As RECT

    Call SetRect(uDstRct, 0, 0, oDstDIB.Width, oDstDIB.Height)
    Call SetRect(uSrcRct, xDst, yDst, xDst + wDst, yDst + hDst)
    Call IntersectRect(uDstRct, uDstRct, uSrcRct)
    
    If (IsRectEmpty(uDstRct) = 0) Then
        
        Call OffsetRect(uDstRct, -xDst, -yDst)
        
        With uDstRct
            i1 = .x1
            i2 = .x2 - .x1 - 1
            If (UpsideDown) Then
                j1 = oSrcDIB.Height - .y2
                j2 = oSrcDIB.Height - .y1 - 1
              Else
                j1 = .y1
                j2 = .y2 - 1
            End If
        End With
        
        If (UpsideDown) Then
            r1 = -oDstDIB.BytesPerScanline
            k1 = (i1 + xDst) - (uDstRct.y2 - 1 + yDst) * r1
          Else
            r1 = oDstDIB.BytesPerScanline
            k1 = (i1 + xDst) + (uDstRct.y1 - 0 + yDst) * r1
        End If
        r2 = oSrcDIB.BytesPerScanline
        i1 = (i1 + xSrc) + (j1 + ySrc) * r2
        
        wgCheckDIBDIBRects = True
    End If
End Function

Private Function wgCheckDIBRectRects( _
                 oDstDIB As cDIB08, _
                 ByVal xDst As Long, ByVal yDst As Long, _
                 ByVal wDst As Long, ByVal hDst As Long, _
                 i1 As Long, i2 As Long, _
                 j1 As Long, j2 As Long, _
                 r As Long _
                 ) As Boolean

  Dim uDstRct As RECT
  Dim uSrcRct As RECT

    Call SetRect(uDstRct, 0, 0, oDstDIB.Width, oDstDIB.Height)
    Call SetRect(uSrcRct, xDst, yDst, xDst + wDst, yDst + hDst)
    Call IntersectRect(uDstRct, uDstRct, uSrcRct)
    
    If (IsRectEmpty(uDstRct) = 0) Then
        
        r = oDstDIB.BytesPerScanline
        
        With uDstRct
            i1 = .y1 * r + .x1
            i2 = .x2 - .x1 - 1
            j1 = .y1
            j2 = .y2 - 1
        End With
        
        wgCheckDIBRectRects = True
    End If
End Function

Private Sub wgMapDIB(uSA As SAFEARRAY1D, aBits() As Byte, oDIB As cDIB08)

    With uSA
        .cbElements = 1
        .cDims = 1
        .Bounds.lLbound = 0
        .Bounds.cElements = oDIB.Size
        .wgData = oDIB.lpBits
    End With
    Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(uSA), 4)
End Sub

Private Sub wgUnmapDIB(aBits() As Byte)
    
    Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
End Sub


