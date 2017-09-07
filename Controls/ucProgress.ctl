VERSION 5.00
Begin VB.UserControl ucProgress 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
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
   ScaleHeight     =   39
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   219
End
Attribute VB_Name = "ucProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucProgress.ctl
' Author:        Warren Galyen
' Dependencies:  None
' Last revision: 05.25.2003
'================================================

Option Explicit

'-- API:

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare Function TranslateColor Lib "olepro32" Alias "OleTranslateColor" (ByVal Clr As OLE_COLOR, ByVal Palette As Long, Col As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOREDRAW      As Long = &H8
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_SHOWWINDOW    As Long = &H40

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE         As Long = (-16)
Private Const WS_THICKFRAME     As Long = &H40000
Private Const WS_BORDER         As Long = &H800000
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const WS_EX_WINDOWEDGE  As Long = &H100&
Private Const WS_EX_CLIENTEDGE  As Long = &H200&
Private Const WS_EX_STATICEDGE  As Long = &H20000

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Type LOGFONT
   lfHeight            As Long
   lfWidth             As Long
   lfEscapement        As Long
   lfOrientation       As Long
   lfWeight            As Long
   lfItalic            As Byte
   lfUnderline         As Byte
   lfStrikeOut         As Byte
   lfCharSet           As Byte
   lfOutPrecision      As Byte
   lfClipPrecision     As Byte
   lfQuality           As Byte
   lfPitchAndFamily    As Byte
   lfFaceName(1 To 32) As Byte
End Type

Private Const LOGPIXELSY             As Long = 90
Private Const FW_NORMAL              As Long = 400
Private Const FW_BOLD                As Long = 700
Private Const FF_DONTCARE            As Long = 0
Private Const DEFAULT_QUALITY        As Long = 0
Private Const DEFAULT_PITCH          As Long = 0
Private Const DEFAULT_CHARSET        As Long = 1
Private Const ANTIALIASED_QUALITY    As Long = 2
Private Const NONANTIALIASED_QUALITY As Long = 3

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CENTER     As Long = &H1
Private Const DT_NOCLIP     As Long = &H100
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER    As Long = &H4

Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Const TRANSPARENT As Long = 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'//

'-- Public Enums.:

Public Enum pbBorderStyleConstants
    [pbNone] = 0
    [pbThin]
    [pbThick]
End Enum

'-- Default Property Values:
Private Const m_def_BorderStyle = [pbThick]
Private Const m_def_BackColor = vbButtonFace
Private Const m_def_ForeColor = vbHighlight
Private Const m_def_Max = 100

'-- Property Variables:
Private m_BorderStyle As pbBorderStyleConstants
Private m_BackColor   As OLE_COLOR
Private m_ForeColor   As OLE_COLOR
Private m_Max         As Long
Private m_Caption     As String

'-- Private Variables:
Private m_Value       As Long
Private m_ControlRect As RECT
Private m_PrgForeRect As RECT
Private m_PrgBackRect As RECT
Private m_PrgPos      As Long
Private m_LastPrgPos  As Long
Private m_hForeBrush  As Long
Private m_hBackBrush  As Long

'-- Event Declarations:
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()
    If (m_Max = 0) Then m_Max = 1
End Sub

Private Sub UserControl_Terminate()
    If (m_hForeBrush <> 0) Then Call DeleteObject(m_hForeBrush)
    If (m_hBackBrush <> 0) Then Call DeleteObject(m_hBackBrush)
End Sub

'//

Private Sub UserControl_Resize()
    Call wgGetProgress
    Call wgCalcRects
    Call UserControl_Paint
End Sub

Private Sub UserControl_Paint()

  Dim hTmpDC     As Long
  Dim hTmpBmp    As Long
  Dim hOldTmpBmp As Long
  Dim hFont      As Long
  Dim hOldFont   As Long
  Dim hRgn       As Long
  Dim lClr       As Long

    hTmpDC = CreateCompatibleDC(hDC)
    hTmpBmp = CreateCompatibleBitmap(hDC, ScaleWidth, ScaleHeight)
    hOldTmpBmp = SelectObject(hTmpDC, hTmpBmp)

    hFont = CreateFontIndirect(wgOLEFontToLogFont(Font, hTmpDC))
    hOldFont = SelectObject(hTmpDC, hFont)
    Call SetBkMode(hTmpDC, TRANSPARENT)

    Call FillRect(hTmpDC, m_PrgForeRect, m_hForeBrush)
    Call FillRect(hTmpDC, m_PrgBackRect, m_hBackBrush)

    Call TranslateColor(m_BackColor, 0, lClr)
    Call SetTextColor(hTmpDC, lClr)
    hRgn = CreateRectRgn(m_PrgForeRect.x1, m_PrgForeRect.y1, m_PrgForeRect.x2, m_PrgForeRect.y2)
    Call SelectClipRgn(hTmpDC, hRgn)
    Call DrawText(hTmpDC, m_Caption, -1, m_ControlRect, DT_SINGLELINE Or DT_CENTER Or DT_NOCLIP)
    Call DeleteObject(hRgn)
    
    Call TranslateColor(m_ForeColor, 0, lClr)
    Call SetTextColor(hTmpDC, lClr)
    hRgn = CreateRectRgn(m_PrgBackRect.x1, m_PrgBackRect.y1, m_PrgBackRect.x2, m_PrgBackRect.y2)
    Call SelectClipRgn(hTmpDC, hRgn)
    Call DrawText(hTmpDC, m_Caption, -1, m_ControlRect, DT_SINGLELINE Or DT_CENTER Or DT_NOCLIP)
    Call DeleteObject(hRgn)
    
    Call BitBlt(hDC, 0, 0, ScaleWidth, ScaleHeight, hTmpDC, 0, 0, vbSrcCopy)

    Call SelectObject(hTmpDC, hOldFont)
    Call DeleteObject(hFont)

    Call SelectObject(hTmpDC, hOldTmpBmp)
    Call DeleteObject(hTmpBmp)
    Call DeleteDC(hTmpDC)
End Sub

'========================================================================================
' Events
'========================================================================================

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get BorderStyle() As pbBorderStyleConstants
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As pbBorderStyleConstants)
    m_BorderStyle = New_BorderStyle
    Call wgSetBorder
    Call wgGetProgress
    Call wgCalcRects
    Call UserControl_Paint
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    Call wgCreateBackBrush
    Call UserControl_Paint
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    Call wgCreateForeBrush
    Call UserControl_Paint
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    If (New_Max < 1) Then New_Max = 1
    m_Max = New_Max
    Call UserControl_Paint
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    Call UserControl_Paint
End Property

Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)

    m_Value = New_Value
    
    Call wgGetProgress
    If (m_PrgPos <> m_LastPrgPos) Then
        m_LastPrgPos = m_PrgPos
        Call wgCalcRects
        Call UserControl_Paint
    End If
End Property

'//

Private Sub UserControl_InitProperties()

    m_BorderStyle = m_def_BorderStyle
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Max = m_def_Max
    
    Call wgSetBorder
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        m_BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
        m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
        m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
        m_Max = .ReadProperty("Max", m_def_Max)
        UserControl.Enabled = .ReadProperty("Enabled", True)
    End With

    Call wgSetBorder
    Call wgCalcRects
    Call wgCreateForeBrush
    Call wgCreateBackBrush
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
        Call .WriteProperty("BackColor", m_BackColor, m_def_BackColor)
        Call .WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
        Call .WriteProperty("Max", m_Max, m_def_Max)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
    End With
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub wgCreateForeBrush()
    
  Dim lClr As Long
    
    If (m_hForeBrush <> 0) Then
        Call DeleteObject(m_hForeBrush)
        m_hForeBrush = 0
    End If
    Call TranslateColor(ForeColor, 0, lClr)
    m_hForeBrush = CreateSolidBrush(lClr)
End Sub

Private Sub wgCreateBackBrush()

  Dim lClr As Long
  
    If (m_hBackBrush <> 0) Then
        Call DeleteObject(m_hBackBrush)
        m_hBackBrush = 0
    End If
    Call TranslateColor(BackColor, 0, lClr)
    m_hBackBrush = CreateSolidBrush(lClr)
End Sub

Private Sub wgGetProgress()
    
    m_PrgPos = (m_Value * ScaleWidth) \ m_Max
End Sub

Private Sub wgCalcRects()
    
    Call SetRect(m_ControlRect, 0, 0, ScaleWidth, ScaleHeight)
    Call SetRect(m_PrgForeRect, 0, 0, m_PrgPos, ScaleHeight)
    Call SetRect(m_PrgBackRect, m_PrgPos, 0, ScaleWidth, ScaleHeight)
End Sub

Private Sub wgSetBorder()

    Select Case m_BorderStyle
        Case [pbNone]
            Call wgSetWinStyle(GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call wgSetWinStyle(GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE)
        Case [pbThin]
            Call wgSetWinStyle(GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call wgSetWinStyle(GWL_EXSTYLE, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE)
        Case [pbThick]
            Call wgSetWinStyle(GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call wgSetWinStyle(GWL_EXSTYLE, WS_EX_CLIENTEDGE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE)
    End Select
End Sub

Private Sub wgSetWinStyle(ByVal lType As Long, ByVal lStyle As Long, ByVal lStyleNot As Long)

  Dim lS As Long
    
    lS = GetWindowLong(hWnd, lType)
    lS = (lS And Not lStyleNot) Or lStyle
    Call SetWindowLong(hWnd, lType, lS)
    Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub

Private Function wgOLEFontToLogFont(oFont As StdFont, ByVal lhDC As Long) As LOGFONT

    With wgOLEFontToLogFont
        
        Call CopyMemory(.lfFaceName(1), ByVal oFont.Name, Len(oFont.Name) + 1)
        .lfCharSet = oFont.Charset
        .lfItalic = -oFont.Italic
        .lfUnderline = -oFont.Underline
        .lfStrikeOut = -oFont.Strikethrough
        .lfWeight = oFont.Weight
        .lfHeight = -(oFont.Size * GetDeviceCaps(lhDC, LOGPIXELSY) / 72)
        .lfQuality = ANTIALIASED_QUALITY
    End With
End Function

