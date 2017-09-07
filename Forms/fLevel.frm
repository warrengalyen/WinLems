VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form fLevel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose level"
   ClientHeight    =   5655
   ClientLeft      =   11325
   ClientTop       =   6525
   ClientWidth     =   5535
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.TreeView tvLevels 
      Height          =   4830
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   8520
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   503
      LabelEdit       =   1
      Style           =   3
      ImageList       =   "ilLevels"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4350
      TabIndex        =   2
      Top             =   5100
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3195
      TabIndex        =   1
      Top             =   5100
      Width           =   1050
   End
   Begin ComctlLib.ImageList ilLevels 
      Left            =   120
      Top             =   5025
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      UseMaskColor    =   0   'False
      _Version        =   327682
   End
End
Attribute VB_Name = "fLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TV_FIRST         As Long = &H1100
Private Const TVM_SETBKCOLOR   As Long = TV_FIRST + 29
Private Const TVM_SETTEXTCOLOR As Long = TV_FIRST + 30

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



'========================================================================================
' Main
'========================================================================================

Private Sub Form_Load()
    
    '-- No icon
    Set Me.Icon = Nothing
    
    '-- Form cursor
    Set Me.MouseIcon = VB.LoadResPicture("CUR_HAND", vbResCursor)
    
    '-- Image-list images
    Call Me.ilLevels.ListImages.Add(, , VB.LoadPicture(AppPath & "RES\Folder_blue.gif"))
    Call Me.ilLevels.ListImages.Add(, , VB.LoadPicture(AppPath & "RES\Folder_blue.gif"))
    Call Me.ilLevels.ListImages.Add(, , VB.LoadPicture(AppPath & "RES\Checkbox_unchecked.gif"))
    Call Me.ilLevels.ListImages.Add(, , VB.LoadPicture(AppPath & "RES\Checkbox_checked.gif"))
    
    '-- Change treeview colors (masking problems with imagelist)
    Call SendMessage(Me.tvLevels.hWnd, TVM_SETBKCOLOR, 0, vbWhite)
    Call SendMessage(Me.tvLevels.hWnd, TVM_SETTEXTCOLOR, 0, vbBlack)
    
    '-- Change buttons style
    Call FlattenButton(Me.cmdOK)
    Call FlattenButton(Me.cmdCancel)
    
    '-- Fill treeview with all levels
    Screen.MousePointer = vbHourglass
    Call wgShowAllLevels
    Screen.MousePointer = vbDefault
    
    '-- Select current level ID
    Call wgSelectCurrent
End Sub

Private Sub cmdOK_Click()

  Dim nRatings As Integer
    
    '-- Available ratings
    Select Case g_eGamePack
        Case [gpLems]
            nRatings = 4
        Case [gpOhNoMoreLems]
            nRatings = 5
        Case [gpCustom]
            nRatings = 1
    End Select
    
    '-- Is a valid node?
    If (tvLevels.SelectedItem.Index > nRatings) Then
        '-- Yes. Store it as current default level ID
        g_nLevelID = Val(Mid$(tvLevels.SelectedItem.Key, 2))
        Call VB.Unload(Me)
      Else
        '-- No
        Call VBA.MsgBox( _
             "No level has been selected." & vbCrLf & vbCrLf & "Please, select a valid level.", _
             vbExclamation _
             )
    End If
End Sub

Private Sub cmdCancel_Click()
    
    '-- Just exit
    Call VB.Unload(Me)
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub wgShowAllLevels()
    
  Dim nRatings As Integer
  Dim r        As Integer
  Dim nLev     As Integer
  Dim sLev     As String
  Dim sKey     As String
  Dim sTxt     As String
    
    '-- Available ratings
    Select Case g_eGamePack
        Case [gpLems]
            Call tvLevels.Nodes.Add(, , "Fun", "Fun", 1)
            Call tvLevels.Nodes.Add(, , "Tricky", "Tricky", 1)
            Call tvLevels.Nodes.Add(, , "Taxing", "Taxing", 1)
            Call tvLevels.Nodes.Add(, , "Mayhem", "Mayhem", 1)
            nRatings = 4
        Case [gpOhNoMoreLems]
            Call tvLevels.Nodes.Add(, , "Tame", "Tame", 1)
            Call tvLevels.Nodes.Add(, , "Crazy", "Crazy", 1)
            Call tvLevels.Nodes.Add(, , "Wild", "Wild", 1)
            Call tvLevels.Nodes.Add(, , "Wicked", "Wicked", 1)
            Call tvLevels.Nodes.Add(, , "Havoc", "Havoc", 1)
            nRatings = 5
        Case [gpCustom]
            Call tvLevels.Nodes.Add(, , "Custom", "Custom", 2)
            nRatings = 1
    End Select
    
    '-- Ratings
    For r = 0 To nRatings - 1
        
        '-- Starting level
        nLev = g_eGamePack * 1000 + r * 100
        
        '-- Get all levels
        Do While FileExists(AppPath & "LEVELS\" & Format$(nLev, "0000") & ".dat")
            
            '-- Extract level ID  ('####')
            sLev = Format$(nLev, "0000")
            
            '-- Generate node key ('k####')
            sKey = "k" & sLev
            
            '-- Load level info
            Call LoadLevelTitle(Val(sLev))
            
            '-- Add node...
            sTxt = Str(Val(Mid$(sLev, 3, 2)) + 1) & ". " & RTrim$(g_uLevel.Title)
            If (IsLevelDone(Val(sLev))) Then
                Call tvLevels.Nodes.Add(r + 1, tvwChild, sKey, sTxt, 4)
              Else
                Call tvLevels.Nodes.Add(r + 1, tvwChild, sKey, sTxt, 3)
                Exit Do
            End If
            
            '-- Next level?
            nLev = nLev + 1
        Loop
    Next r
End Sub

Private Sub wgSelectCurrent()
    
  Dim sKey As String
    
    On Error GoTo errH
    sKey = "k" & Format$(g_nLevelID, "0000")
    With tvLevels.Nodes(sKey)
        .Selected = True
        Call .EnsureVisible
    End With

errH:
    On Error GoTo 0
End Sub


