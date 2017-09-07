VERSION 5.00
Begin VB.UserControl ucScreen08 
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
Attribute VB_Name = "ucScreen08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucScreen08.ctl
' Author:        Warren Galyen
' Dependencies:  cDIB08.cls, mLemsRenderer.bas
' Last revision: 11.17.2006
'================================================

Option Explicit

'-- Public Enums.:

Public Enum BorderStyleCts
    [None] = 0
    [Fixed Single]
End Enum

'-- Private Variables:

Private m_DIBActual  As cDIB08 ' DIB section actual size
Private m_DIBScaled  As cDIB08 ' DIB section scaled size
Private m_xOffset    As Long    'run-time only
Private m_yOffset    As Long    'run-time only
Private m_Zoom       As Long    'run-time only

'-- Event Declarations:

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()

    '-- Initialize DIB
    Set m_DIBActual = New cDIB08
    Set m_DIBScaled = New cDIB08
End Sub

Private Sub UserControl_Terminate()

    '-- Destroy DIB
    Set m_DIBActual = Nothing
    Set m_DIBScaled = Nothing
End Sub

Private Sub UserControl_Paint()

    '-- Refresh Canvas
    Call wgRefresh
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
    RaiseEvent MouseDown(Button, Shift, x \ m_Zoom, y \ m_Zoom)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x \ m_Zoom, y \ m_Zoom)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x \ m_Zoom, y \ m_Zoom)
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub Initialize( _
           ByVal Width As Long, _
           ByVal Height As Long, _
           Optional ByVal Zoom As Long = 1 _
           )
    
    '-- DIB actual size
    Call m_DIBActual.Create(Width, Height)
    
    '-- Create scaled DIB if necessary
    m_Zoom = Zoom
    If (m_Zoom > 1) Then
        Call m_DIBScaled.Create(m_Zoom * Width, m_Zoom * Height)
    End If
End Sub

Public Sub UpdatePalette( _
           ByRef Palette() As Byte _
           )
    
    Call m_DIBActual.SetPalette(Palette())
    Call m_DIBScaled.SetPalette(Palette())
End Sub

Public Sub Refresh()
    Call wgRefresh
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get DIB() As cDIB08
    Set DIB = m_DIBActual
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    Call Me.Refresh
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
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

Public Property Get UserIcon() As StdPicture
Attribute UserIcon.VB_MemberFlags = "400"
    Set UserIcon = UserControl.MouseIcon
End Property
Public Property Set UserIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    Call wgUpdatePointer
End Property

Public Property Let xOffset(ByVal New_xOffset As Long)
    m_xOffset = New_xOffset
End Property
Public Property Get xOffset() As Long
    xOffset = m_xOffset
End Property

Public Property Let yOffset(ByVal New_yOffset As Long)
    m_yOffset = New_yOffset
End Property
Public Property Get yOffset() As Long
    yOffset = m_yOffset
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

'//

Public Property Get hWnd() As Long
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = UserControl.hWnd
End Property

'//

Private Sub UserControl_InitProperties()
    UserControl.BorderStyle = [None]
    UserControl.BackColor = vbApplicationWorkspace
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", [None])
    UserControl.BackColor = PropBag.ReadProperty("BackColor", vbApplicationWorkspace)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, [None])
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbApplicationWorkspace)
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub wgRefresh()
  
    If (m_DIBActual.HasDIB) Then
        '-- Paint
        If (m_DIBScaled.HasDIB) Then
            Call FXStretch(m_DIBScaled, m_DIBActual)
            Call m_DIBScaled.Paint(UserControl.hDC, m_xOffset, m_yOffset)
          Else
            Call m_DIBActual.Paint(UserControl.hDC, m_xOffset, m_yOffset)
        End If
      Else
        '-- Erase background
        Call UserControl.Cls
    End If
End Sub

Private Sub wgUpdatePointer()

    If (Not UserControl.MouseIcon Is Nothing) Then
         UserControl.MousePointer = vbCustom
      Else
         UserControl.MousePointer = vbDefault
    End If
End Sub


