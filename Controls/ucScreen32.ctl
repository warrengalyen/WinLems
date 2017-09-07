VERSION 5.00
Begin VB.UserControl ucScreen32 
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   73
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
End
Attribute VB_Name = "ucScreen32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucScreen32.ctl
' Author:        Carles P.V.
' Dependencies:  cDIB32.cls
' Last revision: 2004.09.15
'================================================

Option Explicit

'-- API:

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const RGN_DIFF As Long = 4

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TranslateColor Lib "olepro32" Alias "OleTranslateColor" (ByVal Clr As OLE_COLOR, ByVal Palette As Long, Col As Long) As Long

'//

'-- Public Enums.:

Public Enum BorderStyleCts
    [None] = 0
    [Fixed Single]
End Enum

Public Enum eWorkModeCts
    [cnvScrollMode] = 0
    [cnvUserMode]
End Enum

'-- Property Variables:

Private m_BackColor       As OLE_COLOR
Private m_EraseBackground As Boolean      'run-time only
Private m_FitMode         As Boolean      'run-time only
Private m_WorkMode        As eWorkModeCts 'run-time only
Private m_Zoom            As Long         'run-time only

'-- Private Variables:

Private m_Width           As Long
Private m_Height          As Long
Private m_Left            As Long
Private m_Top             As Long
Private m_hPos            As Long
Private m_hMax            As Long
Private m_vPos            As Long
Private m_vMax            As Long
Private m_lsthPos         As Single
Private m_lstvPos         As Single
Private m_lsthMax         As Single
Private m_lstvMax         As Single
Private m_Down            As Boolean
Private m_Pt              As POINTAPI

'-- Event Declarations:

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event Scroll()
Public Event Resize()

'-- Public objects:

Public DIB As cDIB32 ' 32-bit DIB section
Attribute DIB.VB_VarMemberFlags = "400"



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()

    '-- Initialize DIB
    Set Me.DIB = New cDIB32
    
    '-- Initial values
    m_EraseBackground = True
    m_WorkMode = [cnvScrollMode]
    m_Zoom = 1
End Sub

Private Sub UserControl_Terminate()

    '-- Destroy DIB
    Set Me.DIB = Nothing
End Sub

Private Sub UserControl_Resize()

    '-- Resize and refresh
    Call wgResizeCanvas
    Call wgRefreshCanvas
    
    RaiseEvent Resize
End Sub

Private Sub UserControl_Paint()

    '-- Refresh Canvas
    Call wgRefreshCanvas
End Sub

'========================================================================================
' Events + Scrolling
'========================================================================================

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '-- Mouse down flag / Store values
    m_Down = (Button = vbLeftButton)
    m_Pt.x = x
    m_Pt.y = y
    
    RaiseEvent MouseDown(Button, Shift, wgDIBx(x), wgDIBy(y))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    If (m_Down And m_WorkMode = [cnvScrollMode]) Then
    
        '-- Apply offsets
        m_hPos = m_hPos + (m_Pt.x - x)
        m_vPos = m_vPos + (m_Pt.y - y)
        
        '-- Check margins
        If (m_hPos < 0) Then m_hPos = 0 Else If (m_hPos > m_hMax) Then m_hPos = m_hMax
        If (m_vPos < 0) Then m_vPos = 0 Else If (m_vPos > m_vMax) Then m_vPos = m_vMax
        
        '-- Save current position
        m_Pt.x = x
        m_Pt.y = y
        
        '-- Refresh and raise event
        If (m_lsthPos <> m_hPos Or m_lstvPos <> m_vPos) Then
            Call wgRefreshCanvas
            RaiseEvent Scroll
        End If
        
        '-- Store
        m_lsthPos = m_hPos
        m_lstvPos = m_vPos
    End If
    
    RaiseEvent MouseMove(Button, Shift, wgDIBx(x), wgDIBy(y))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    '-- Mouse down flag
    m_Down = False
    
    RaiseEvent MouseUp(Button, Shift, wgDIBx(x), wgDIBy(y))
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub Refresh()
    Call wgRefreshCanvas
End Sub

Public Sub Resize()
    Call wgResizeCanvas
End Sub

Public Function Scroll( _
                ByVal x As Long, _
                ByVal y As Long _
                ) As Boolean

    '-- Apply offsets
    m_hPos = m_hPos - x
    m_vPos = m_vPos - y
    
    '-- Check margins
    If (m_hPos < 0) Then m_hPos = 0 Else If (m_hPos > m_hMax) Then m_hPos = m_hMax
    If (m_vPos < 0) Then m_vPos = 0 Else If (m_vPos > m_vMax) Then m_vPos = m_vMax
    
    '-- Need to refresh?
    If (m_lsthPos <> m_hPos Or m_lstvPos <> m_vPos) Then
        Call wgRefreshCanvas: Scroll = True
        RaiseEvent Scroll
    End If
    
    '-- Store
    m_lsthPos = m_hPos
    m_lstvPos = m_vPos
End Function

'========================================================================================
' Properties
'========================================================================================

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    Call Me.Refresh
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Get BorderStyle() As BorderStyleCts
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleCts)
    UserControl.BorderStyle() = New_BorderStyle
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let EraseBackground(ByVal New_EraseBackground As Boolean)
    m_EraseBackground = New_EraseBackground
End Property
Public Property Get EraseBackground() As Boolean
Attribute EraseBackground.VB_MemberFlags = "400"
    EraseBackground = m_EraseBackground
End Property

Public Property Let FitMode(ByVal New_FitMode As Boolean)
    m_FitMode = New_FitMode
End Property
Public Property Get FitMode() As Boolean
Attribute FitMode.VB_MemberFlags = "400"
    FitMode = m_FitMode
End Property

Public Property Get UserIcon() As StdPicture
Attribute UserIcon.VB_MemberFlags = "400"
    Set UserIcon = UserControl.MouseIcon
End Property
Public Property Set UserIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    Call wgUpdatePointer
End Property

Public Property Let WorkMode(ByVal New_WorkMode As eWorkModeCts)
    m_WorkMode = New_WorkMode
    Call wgUpdatePointer
End Property
Public Property Get WorkMode() As eWorkModeCts
Attribute WorkMode.VB_MemberFlags = "400"
    WorkMode = m_WorkMode
End Property

Public Property Let Zoom(ByVal New_Zoom As Long)
    m_Zoom = IIf(New_Zoom < 1, 1, New_Zoom)
End Property
Public Property Get Zoom() As Long
Attribute Zoom.VB_MemberFlags = "400"
    Zoom = m_Zoom
End Property

'//

Public Property Get ScaleWidth() As Long
Attribute ScaleWidth.VB_MemberFlags = "400"
    ScaleWidth = UserControl.ScaleWidth
End Property
Public Property Get ScaleHeight() As Long
Attribute ScaleHeight.VB_MemberFlags = "400"
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Get ScrollHMax() As Long
Attribute ScrollHMax.VB_MemberFlags = "400"
    ScrollHMax = m_hMax
End Property
Public Property Get ScrollVMax() As Long
Attribute ScrollVMax.VB_MemberFlags = "400"
    ScrollVMax = m_vMax
End Property
Public Property Get ScrollHPos() As Long
Attribute ScrollHPos.VB_MemberFlags = "400"
    ScrollHPos = m_hPos
End Property
Public Property Get ScrollVPos() As Long
Attribute ScrollVPos.VB_MemberFlags = "400"
    ScrollVPos = m_vPos
End Property

'//

Public Property Get hWnd() As Long
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = UserControl.hWnd
End Property

'//

Private Sub UserControl_InitProperties()
    UserControl.BorderStyle = [None]
    m_BackColor = vbApplicationWorkspace
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", [None])
    m_BackColor = PropBag.ReadProperty("BackColor", vbApplicationWorkspace)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, [None])
    Call PropBag.WriteProperty("BackColor", m_BackColor, vbApplicationWorkspace)
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub wgEraseBackground()

  Dim hRgn_1 As Long
  Dim hRgn_2 As Long
  Dim lColor As Long
  Dim hBrush As Long
    
    '-- Create brush (background)
    Call TranslateColor(m_BackColor, 0, lColor)
    hBrush = CreateSolidBrush(lColor)

    '-- Create Cls region (Control Rect. - Canvas Rect.)
    hRgn_1 = CreateRectRgn(0, 0, ScaleWidth, ScaleHeight)
    hRgn_2 = CreateRectRgn(m_Left, m_Top, m_Left + m_Width, m_Top + m_Height)
    Call CombineRgn(hRgn_1, hRgn_1, hRgn_2, RGN_DIFF)
    
    '-- Fill it
    Call FillRgn(hDC, hRgn_1, hBrush)
    
    '-- Clear
    Call DeleteObject(hBrush)
    Call DeleteObject(hRgn_1)
    Call DeleteObject(hRgn_2)
End Sub

Private Sub wgRefreshCanvas()
  
  Dim xOff As Long, yOff As Long
  Dim wDst As Long, hDst As Long
  Dim xSrc As Long, ySrc As Long
  Dim wSrc As Long, hSrc As Long
    
    If (Me.DIB.hDIB <> 0) Then
        
        '-- Get Left and Width of source image rectangle:
        If (m_hMax And Not m_FitMode) Then
            xOff = -m_hPos Mod m_Zoom
            wDst = (m_Width \ m_Zoom) * m_Zoom + 2 * m_Zoom
            xSrc = m_hPos \ m_Zoom
            wSrc = m_Width \ m_Zoom + 2
          Else
            xOff = m_Left
            wDst = m_Width
            xSrc = 0
            wSrc = Me.DIB.Width
        End If
        
        '-- Get Top and Height of source image rectangle:
        If (m_vMax And Not m_FitMode) Then
            yOff = -m_vPos Mod m_Zoom
            hDst = (m_Height \ m_Zoom) * m_Zoom + 2 * m_Zoom
            ySrc = m_vPos \ m_Zoom
            hSrc = m_Height \ m_Zoom + 2
          Else
            yOff = m_Top
            hDst = m_Height
            ySrc = 0
            hSrc = Me.DIB.Height
        End If
        
        '-- Erase background
        If (m_EraseBackground) Then
            Call wgEraseBackground
        End If
        '-- Paint visible source rectangle:
        Call Me.DIB.Stretch(hDC, xOff, yOff, wDst, hDst, xSrc, ySrc, wSrc, hSrc)
        
      Else
        '-- Erase background
        Call wgEraseBackground
    End If
End Sub

Private Sub wgResizeCanvas()
    
    With Me.DIB
        
        If (.hDIB <> 0) Then
        
            If (m_FitMode = False) Then
            
                '-- Get new Width
                If (.Width * m_Zoom > ScaleWidth) Then
                    m_hMax = .Width * m_Zoom - ScaleWidth
                    m_Width = ScaleWidth
                  Else
                    m_hMax = 0
                    m_Width = .Width * m_Zoom
                End If
                
                '-- Get new Height
                If (.Height * m_Zoom > ScaleHeight) Then
                    m_vMax = .Height * m_Zoom - ScaleHeight
                    m_Height = ScaleHeight
                  Else
                    m_vMax = 0
                    m_Height = .Height * m_Zoom
                End If
                
                '-- Offsets
                m_Left = (ScaleWidth - m_Width) \ 2
                m_Top = (ScaleHeight - m_Height) \ 2
              
              Else
                '-- Get best-fit info
                Call wgGetBestFitInfo(.Width, .Height, ScaleWidth, ScaleHeight, m_Left, m_Top, m_Width, m_Height)
            End If
                                
            '-- Memorize position
            If (m_lsthMax) Then
                m_hPos = (m_lsthPos * m_hMax) \ m_lsthMax
              Else
                m_hPos = m_hMax \ 2
            End If
            If (m_lstvMax) Then
                m_vPos = (m_lstvPos * m_vMax) \ m_lstvMax
              Else
                m_vPos = m_vMax \ 2
            End If
            m_lsthPos = m_hPos: m_lstvPos = m_vPos
            m_lsthMax = m_hMax: m_lstvMax = m_vMax
          
          Else
            '-- 'Hide' canvas
            m_Width = 0: m_Height = 0
        End If
    End With
    
    '-- Update mouse pointer
    Call wgUpdatePointer
End Sub

Private Sub wgUpdatePointer()

    If (m_WorkMode = [cnvScrollMode]) Then
        If ((m_hMax Or m_vMax) And Not m_FitMode) Then
            UserControl.MousePointer = vbSizeAll
          Else
            UserControl.MousePointer = vbDefault
        End If
      Else
        If (Not UserControl.MouseIcon Is Nothing) Then
            UserControl.MousePointer = vbCustom
        End If
    End If
End Sub

Private Function wgDIBx(ByVal x As Long) As Long

    If (Me.DIB.hDIB <> 0) Then
        If (m_FitMode) Then
            wgDIBx = Int((x - m_Left) / (m_Width / Me.DIB.Width))
          Else
            wgDIBx = Int((m_hPos + x - m_Left) / m_Zoom)
        End If
    End If
End Function

Private Function wgDIBy(ByVal y As Long) As Long

    If (Me.DIB.hDIB <> 0) Then
        If (m_FitMode) Then
            wgDIBy = Int((y - m_Top) / (m_Height / Me.DIB.Height))
          Else
            wgDIBy = Int((m_vPos + y - m_Top) / m_Zoom)
        End If
    End If
End Function

Private Sub wgGetBestFitInfo( _
            ByVal SrcWidth As Long, ByVal SrcHeight As Long, _
            ByVal DstWidth As Long, ByVal DstHeight As Long, _
            xBF As Long, yBF As Long, _
            BFWidth As Long, BFHeight As Long, _
            Optional ByVal StretchFit As Boolean = False _
                        )
                          
  Dim cW As Single
  Dim cH As Single
    
    If ((SrcWidth > DstWidth Or SrcHeight > DstHeight) Or StretchFit) Then
        cW = DstWidth / SrcWidth
        cH = DstHeight / SrcHeight
        If (cW < cH) Then
            BFWidth = DstWidth
            BFHeight = SrcHeight * cW
          Else
            BFHeight = DstHeight
            BFWidth = SrcWidth * cH
        End If
      Else
        BFWidth = SrcWidth
        BFHeight = SrcHeight
    End If
    If (BFWidth < 1) Then BFWidth = 1
    If (BFHeight < 1) Then BFHeight = 1
    
    xBF = (DstWidth - BFWidth) \ 2
    yBF = (DstHeight - BFHeight) \ 2
End Sub


