VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinLems"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9600
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
   ForeColor       =   &H00000000&
   HasDC           =   0   'False
   Icon            =   "fMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin ComctlLib.Toolbar ucToolbar 
      Height          =   630
      Left            =   1170
      TabIndex        =   1
      Top             =   4935
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ilToolbar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   14
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Climber"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Floater"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Bomber"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Blocker"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Builder"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Basher"
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Miner"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Digger"
            Object.Tag             =   ""
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Plus"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Minus"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Pause"
            Object.Tag             =   ""
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Armageddon"
            Object.Tag             =   "0"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin WinLems.ucScreen08 ucPanView 
      Height          =   1200
      Left            =   0
      TabIndex        =   14
      Top             =   5955
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   2117
      BackColor       =   0
   End
   Begin VB.Timer tmrPlusMinus 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8550
      Top             =   5370
   End
   Begin VB.FileListBox flFilter 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   9720
      Pattern         =   "*.bmp"
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin ComctlLib.StatusBar ucInfo 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   7155
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7779
            MinWidth        =   2646
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2249
            MinWidth        =   2249
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2249
            MinWidth        =   2249
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2249
            MinWidth        =   2249
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2249
            MinWidth        =   2249
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin WinLems.ucScreen08 ucScreen 
      Height          =   4800
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   8467
      BackColor       =   0
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   6660
      TabIndex        =   11
      Top             =   5565
      Width           =   375
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   6075
      TabIndex        =   10
      Top             =   5565
      Width           =   375
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   5370
      TabIndex        =   9
      Top             =   5565
      Width           =   375
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   4785
      TabIndex        =   8
      Top             =   5565
      Width           =   375
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   7
      Top             =   5565
      Width           =   375
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3615
      TabIndex        =   6
      Top             =   5565
      Width           =   375
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3030
      TabIndex        =   5
      Top             =   5565
      Width           =   375
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2445
      TabIndex        =   4
      Top             =   5565
      Width           =   375
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1860
      TabIndex        =   3
      Top             =   5565
      Width           =   375
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1275
      TabIndex        =   2
      Top             =   5565
      Width           =   375
   End
   Begin ComctlLib.ImageList ilToolbar 
      Left            =   9015
      Top             =   5220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":13C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":16DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":19F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":1D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":202C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":2346
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":2660
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuGameTop 
      Caption         =   "&Game"
      Begin VB.Menu mnuGame 
         Caption         =   "Pack"
         Index           =   0
         Begin VB.Menu mnuPack 
            Caption         =   "Lems"
            Index           =   0
         End
         Begin VB.Menu mnuPack 
            Caption         =   "Oh No! More Lems"
            Index           =   1
         End
         Begin VB.Menu mnuPack 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuPack 
            Caption         =   "Custom"
            Index           =   9
         End
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuGame 
         Caption         =   "&Play!"
         Index           =   2
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "&Choose level..."
         Index           =   3
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuGame 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
   Begin VB.Menu mnuOptionsTop 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Sound effects"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Music"
         Index           =   1
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Project:       WinLems
' Author:        Warren Galyen
' Dependencies:  -
' First release: 2005.07.08
' Last revision: 06.30.2007
'================================================
'
' History:
'
' 1.0.00: - First release.
'
' 1.0.01: - Improved-fixed adjustment between animations
'         - Added explosion frame.
'
' 1.0.02: - Fixed 'fissure' detection. Now lems fall at
'           'every' pixel.
'         - Added all levels.
'
' 1.0.03: - Fixed rendering order: Terrain-Steel-Objects.
'         - Fixed masking: Rungs were not 'bashed' by
'           bashers, miners and diggers.
'         - Fixed some levels.
'         - Fixed GUI.
'
' 1.0.04: - Fixed builder feet checking (one pixel before).
'         - Fixed 'falling' end. Fallers could land over
'           blockers.
'
' 1.0.05: - Fixed layering (objects vs rungs).
'
' 1.0.06: - Fixed timing! I've just realized that one
'           second is not a second. Don't know how exact
'           is new delay, but it seems to me that rounds
'           1,6 seconds. Sorry for that! I've seen this
'           trying to pass 'Just a minute...' level
'           (Mayhem rating).
'         - Also fixed basher animation. It moved for-
'           ward only 3 pixels every loop. They should
'           be 4.
'
' 1.0.07: - Fixed terrain painting. Now, there is no need to
'           render terrain transparently. All objects
'           are rendered (back* and fore*) onto terrain.
'         - Minor fix on builder animation. Checks if
'           lem can move up to next rung (to avoid lem
'           get 'hooked' 'inside' terrain).
'
' 1.0.08: - Added 'Custom' rating folder. Now it's not
'           necessary to overwrite original levels.
'
' 1.0.09: - Improved x/y offsets between animations.
'
' 1.0.10: - Added music (original Windows game themes).
'
' 1.0.11: - Fixed pause (frame looping).
'         - Frame delay changed to 80 (still slow?).
'
' 1.0.12: - Fixed pause (on starting and ending stages).
'         - Fixed explosion hole (2 pixels down).
'
' 1.0.13: - Fixed 'Take a running jump... (Taxing)' level.
'
' 1.0.14: - Fixed pixel checking in some animations.
'
' 1.0.15: - Fixed some levels (minor offsets).
'
' 1.0.16: - Fixed exploding animation (fall checking).
'         - Fixed some graphics.
'         - Added 'hyperlink' to Lems' PSC page.
'
' 1.0.17: - Fixed masking: liquid and fire (no overlapping).
'
' 1.0.18: - Improved 'Miner' animation (pixel detection).
'
' 1.0.19: - Removed unnecessary screen reset.
'         - Preview screen now 1x (double resolution).
'         - Pre-exploding animation (1 pixel fall more).
'
' 1.0.20: - Fixed traps #11/12 (Graphic Set #3).
'
' 1.0.21: - Fixed wgCheckPixel() sub (.Job set to 'None').
'         - Fixed basher feet checking (now detects exit,
'           traps...). Also fixed 'ahead' checking (steel).
'         - Fixed 'Lost something? (Tricky)' level.
'
' 1.0.22: - Fixed miner's y-offset (before next animation;
'           problems with diggers and builders).
'
' 1.0.23: - Fixed miner vs blocker (reverse direction).
'         - Fixed 'Armageddon' re-activation.
'
' 1.0.24: - Fixed some levels (terrain).
'         - Fixed basher ending job (5 pixels instead of 7).
'
' 1.0.25: - All graphics have been 'cleaned'. All colors have
'           been normalized. Now, color could be reduced to a
'           6-bit palette. If someone wants to convert all code
'           regarding 32-bit processing to 8-bit...
'           A considerable amount of memory could be saved,
'           as well as processing time could be improved.
'           Note: An intermediate step could be avoid mask
'           buffer and make use of 4th byte of main buffer
'           ('alpha' channel).
'
' 1.0.26: - Fixed trap checking. Added 'feet' flag (enable
'           trap only on 'feet checking').
'
' 1.0.27: - Removed cTile class and mTextExt module: Now
'           via bits directly.
'         - Preview window strecth also via same method.
'         - Interframe delay changed from 80 to 75 ms.
'
' 1.0.28: - Fixed rare case*: DieNextFrame flag is checked
'           before is set. Sometimes it could be reseted
'           again on next frame (so Lem does not die).
'           *'Lend a helping hand' level (Taxing).
'
' 1.0.29: - Basher's falling check: done one loop after.
'
' 1.0.30: - Interframe delay changed from 75 to 80 ms.
'           It seems to be the 'official' delay
'
' 1.0.31: - Builder's head check (two pixels after).
'         - Miner's feet checks (one pixel after).
'         - Once again: interframe delay changed from 80
'           to 75 ms. (closer to DOS speed).
'         - 8 frames fade-out (10 before).
'
' 1.0.32: Hope last update!
'         - Improved Lem's description.
'         - Fixed FXTile() routine.
'
' 1.0.33: Not last...
'         Fixed 'Just a Minute (Part Two) (Mayhem)' level.
'
' 1.0.34: Yes, not last...
'         - Timer delay from 1.6 s to 1.35 s. (adjusted to
'           new interframe delay). Not sure how much exact
'           this new delay is...
'
' 1.1.00: - Big update:
'
'           * Graphics and image processing have been trans-
'             lated to 8-bit indexed bitmaps (in fact, 6-bit).
'           * Now, cDIB08 class does not use hDIB->hDC. All
'             is done via bitmap bits (1-D byte array).
'           * 8-bit screen color-depth is not supported.
'
'             This way we get to avoid making use of extra
'           GDI objects, as well as memory usage is conside-
'           rably reduced and processing time considerably
'           improved.
'
'           Note: Level Editor still uses 32-bit processing
'
' 1.1.01: - Minor improvement related to sound synchronization:
'           Trap sound is now played on correct frame index.
'
' 1.1.02: - Screen greyscaled on pause.
'
' 1.1.03: - Greyscaling on pause is now a 'darkening'.
'
' 1.1.04: - All graphics have been verticaly flipped.
'           cDIB08 reads bitmap data directly: faster loading.
'
' 1.1.05: - Removed second timer (time): all done via *frame*
'           timer: synchronized and not depending on system.
'
' 1.1.06: - Added 'Screenshot' feature (F12).
'
' 1.1.07: - Screen size normalized to 1600x160 (from 1680x160)
'           Differences between DOS/Windows version?
'         - All levels fixed: added animations not present in
'           Windows version (?). A 'x' offset (-16) has been
'           applied to all levels, too.
'         ********************************************************
'         - Improved 'hot-lem' description
'         - Fixed ApplyPrepared() routine. Ability/Job could be
'           applied during exploding animation.
'
' 1.1.08: - Minor addition: 'Fast-Forward' (Key [F] during game).
'
' 1.2.00: - New timing method
'
'           Before: API timer (WM_TIMER)
'           Now:    The classic game loop
'
'           Important note: Add always a *Sleep(1)* call to minimize
'           loop's CPU consuming time! (Do-Loop *loop* itself).
'
'           First method (API timer) it's quite neat, but not reliable
'           when dealing with little dts (less than 100ms) and not
'           *high-performance CPUs*.
'
'           Now, it works fine (at desired *exact* fps) on my old
'           200MHz machine.
'
' 1.2.01: - ucScreen zooming now *manualy processed* instead of
'           using API StrechDIBits: faster!!!
'
' 1.2.02: - Added explosion particles.
'
' 1.2.03: - Fixed 'all particles out of screen' check.
'
' 1.2.04: - Graphics storing format has changed.
'           8-bit bitmaps now stored (encoded) as follows:
'           - A 8-bit palette stored separately (\RES\Palette.dat)
'           - Bitmaps' data as RLE compressed bit-maps (\GFX\*.bm):
'             00-01:  W
'             02-03:  H
'             04-EOF: RLE 8-bit-map
'
'           Results:
'           ~390KB total vs 1.67MB (LZW encoding reached ~140KB)


Option Explicit

'========================================================================================
' GUI initialization
'========================================================================================

Public Enum eGamePack
    [gpLems] = 0
    [gpOhNoMoreLems] = 1
    [gpCustom] = 9
End Enum

Public Enum eMode
    [mMenuScreen] = 1
    [mLevelScreen]
    [mPlaying]
    [mGameScreen]
End Enum

Private Sub Form_Load()

    '-- Initialize modules
    Call InitializeLemsRenderer
    Call InitializeLems
    Call InitializeSound
    Call InitializeMusic
    Call InitializeWheel
   
    '-- Initialize main view
    With ucScreen
        Call .Initialize(320, 160, 2)
        Call .UpdatePalette(GetGlobalPalette())
    End With
    
    '-- Initialize panoramic view
    With ucPanView
        Call .Initialize(640, 80, 1)
        Call .UpdatePalette(GetGlobalPalette())
    End With
    
    '-- Start...
    With Me
        
        '-- App. cursor
        Set .MouseIcon = VB.LoadResPicture("CUR_HAND", vbResCursor)
        
        '-- Flatten and disable toolbar
        Call FlattenToolbar(ucToolbar)
        Call wgSetToolbarState(bEnable:=False)
        
        '-- Load settings
        g_nLevelID = GetLastLevel()
        g_eGamePack = g_nLevelID \ 1000
        .mnuPack(g_eGamePack).Checked = True
        
        '-- Update menus' accelerators
        .mnuOptions(0).Caption = mnuOptions(0).Caption & vbTab & "S"
        .mnuOptions(1).Caption = mnuOptions(1).Caption & vbTab & "M"
        
        '-- Show menu screen
        Call wgShowMenuScreen
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '-- Stop timer
    Call StopTimer
    
    '-- Terminate main module
    Call TerminateLems
    
    '-- Stop midi
    Call TerminateMusic
    Call CloseMidi
    
    '-- Save settings
    Call SetLastLevel(g_nLevelID)
End Sub

'========================================================================================
' Menu
'========================================================================================

Private Sub mnuPack_Click(Index As Integer)
      
    '-- Uncheck and check
    mnuPack(g_eGamePack).Checked = False
    mnuPack(Index).Checked = True
    
    '-- Set level pack ID and reset to first level
    g_eGamePack = Index
    g_nLevelID = Index * 1000
End Sub

Private Sub mnuGame_Click(Index As Integer)
    
    Call VBA.DoEvents
    
    Select Case Index
    
        Case 2 '-- Play!
            
            '-- Load current level data and show 'level screen'
            If (LoadLevel(g_nLevelID)) Then
                Call wgShowLevelScreen
              Else
                Call VBA.MsgBox( _
                     "Unable to load level (ID=" & g_nLevelID & ") ", _
                     vbExclamation _
                     )
           End If
            
        Case 3 '-- Choose level...
        
            Call fLevel.Show(vbModal, Me)
        
        Case 5 '-- Exit
        
            Call VB.Unload(Me)
    End Select
End Sub

Private Sub mnuOptions_Click(Index As Integer)
    
    Select Case Index
        
        Case 0 '-- Sound effects on/off
        
            mnuOptions(0).Checked = Not mnuOptions(0).Checked
            Call SetSoundEffectsState(CBool(mnuOptions(0).Checked))
            If (Me.Tag = [mMenuScreen]) Then
                Call wgShowMenuScreen
            End If
            
        Case 1 '-- Music on/off
        
            mnuOptions(1).Checked = Not mnuOptions(1).Checked
            Call SetMusicState(CBool(mnuOptions(1).Checked))
            If (Me.Tag = [mMenuScreen]) Then
                Call wgShowMenuScreen
            End If
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    
    '-- About box
    Call VBA.MsgBox(vbCrLf & _
                "WinLems" & vbCrLf & vbCrLf & _
                "Current version: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
                "Date: 12.05.2011" & vbCrLf & vbCrLf & _
                "Based on the classic DOS game" & Space$(5) & vbCrLf & vbCrLf & _
                "Created by Warren Galyen 2005-2011" & Space$(5) & vbCrLf & _
                 vbCrLf & vbCrLf, vbInformation, "About")
End Sub

'========================================================================================
' Accelerators
'========================================================================================

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    
        Case vbKeyEscape
        
            Select Case Me.Tag

                Case [mPlaying]

                    '-- Playing: quit game
                    If (IsTimerPaused = False) Then
                        Call SetGameStage([gsEnding])
                    End If

                Case Else

                    '-- Show menu screen
                    Call wgShowMenuScreen
            End Select
            
        Case vbKeyA
            
            '-- Scroll view to left
            If (Me.Tag = [mPlaying]) Then
                Call DoScroll(dx:=-10)
            End If
            
        Case vbKeyZ
            
            '-- Scroll view to right
            If (Me.Tag = [mPlaying]) Then
                Call DoScroll(dx:=10)
            End If
        
        Case vbKeyControl
            
            '-- Center view
            If (Me.Tag = [mPlaying]) Then
                Call DoScrollTo(x:=680, ScaleAndCenter:=False)
            End If
            
        Case vbKey1 To vbKey8
            
            '-- Job/Ability selection
            If (Me.Tag = [mPlaying]) Then
                With ucToolbar
                    If (.Buttons(KeyCode - 48).Enabled) Then
                        If (.Buttons(KeyCode - 48).Value = [tbrUnpressed]) Then
                            .Buttons(KeyCode - 48).Value = [tbrPressed]
                            Call ucToolbar_ButtonClick(.Buttons(KeyCode - 48))
                        End If
                    End If
                End With
            End If
        
        Case vbKeySpace
            
            '-- Pause on/off
            If (Me.Tag = [mPlaying]) Then
                With ucToolbar
                    .Buttons("Pause").Value = 1 - .Buttons("Pause").Value
                    Call ucToolbar_ButtonClick(.Buttons("Pause"))
                End With
            End If
            
        Case vbKeyS
            
            '-- Sound effects on/off
            Call mnuOptions_Click(0)
            
        Case vbKeyM
        
            '-- Music on/off
            Call mnuOptions_Click(1)
            
        Case vbKeyF12
            
            '-- Screenshot
            Call ucScreen.DIB.CopyToClipboard
    End Select
End Sub

'========================================================================================
' Main screen
'========================================================================================

Private Sub ucScreen_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    
    Select Case Me.Tag
    
        Case [mMenuScreen]
        
             Select Case Button
            
                Case vbLeftButton
                    
                    '-- Force Play
                    Call mnuGame_Click(2)
           
                Case vbRightButton
        
                    '-- Choose level...
                    Call mnuGame_Click(3)
            End Select
        
        Case [mLevelScreen]
            
            Select Case Button
            
                Case vbLeftButton
            
                    '-- State flag
                    Me.Tag = [mPlaying]
                    
                    '-- Enable toolbar and show features
                    Call wgSetToolbarState(bEnable:=True)
                    Call wgShowLevelFeatures
                        
                    '-- Initialize game
                    Call InitializeGame
                    
                    '-- Start 'framing'
                    Call StartTimer
                
                Case vbRightButton
                
                    '-- Show menu screen
                    Call wgShowMenuScreen
            End Select
        
        Case [mPlaying]
        
            Select Case Button
            
                Case vbLeftButton
                    
                    '-- Apply prepared job or ability, if any
                    Call ApplyPrepared
                
                Case vbRightButton
                
                    '-- Quit game
                    If (IsTimerPaused = False) Then
                        Call SetGameStage([gsEnding])
                    End If
            End Select
            
        Case [mGameScreen]
            
            '-- Success?
            If (GetLemsSaved() >= g_uLevel.LemsToBeSaved) Then
                g_nLevelID = GetNextLevel()
            End If

            Select Case Button
            
                Case vbLeftButton
                    
                    '-- Play level again
                    Call mnuGame_Click(2)
           
                Case vbRightButton
        
                    '-- Show menu screen
                    Call wgShowMenuScreen
            End Select
    End Select
End Sub

'========================================================================================
' Toolbar
'========================================================================================

Private Sub ucToolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    
    If (Me.Tag = [mPlaying]) Then
        
        With Button
        
            Select Case .Index
            
                Case Is = 1, 2, 3
                    
                    '-- Prepare ability
                    Call PrepareAbility(2 ^ (.Index - 1))
                    Call PlaySoundFX([sfxChangeOp])
                
                Case Is = 4, 5, 6, 7, 8
                
                    '-- Prepare job
                    Call PrepareJob(.Index - 3)
                    Call PlaySoundFX([sfxChangeOp])
                                    
                Case Is = 13
                    
                    '-- Pause on/off
                    Call PauseTimer(.Value = [tbrPressed])
                    If (IsTimerPaused) Then
                        Call ucScreen.UpdatePalette( _
                             GetFadedOutGlobalPalette(Amount:=25) _
                             )
                        Call DoFrame
                      Else
                        Call ucScreen.UpdatePalette( _
                             GetGlobalPalette() _
                             )
                    End If
            
                Case Is = 14
                    
                    '-- Activate Armageddon? (Needs double-click)
                    If (IsTimerPaused = False) Then
                        If (VBA.Timer() - .Tag < 0.5) Then
                            If (IsArmageddonActivated = False) Then
                                Call StartArmageddon
                            End If
                        End If
                        .Tag = VBA.Timer()
                    End If
            End Select
        End With
    End If
End Sub

'== Plus/Minus buttons ==================================================================

Private Sub ucToolbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '-- If playing, enable plus-minus timer
    If (Me.Tag = [mPlaying]) Then
        tmrPlusMinus.Enabled = True
    End If
End Sub

Private Sub ucToolbar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '-- Disable plus-minus timer
    tmrPlusMinus.Enabled = False
End Sub

Private Sub tmrPlusMinus_Timer()
    
    Select Case True
        
        Case ucToolbar.Buttons("Plus").Value = [tbrPressed]
            
            '-- Increase release rate by one
            If (g_lReleaseRate < RELEASE_RATE_MAX) Then
                g_lReleaseRate = g_lReleaseRate + 1
                lblButton(9).Caption = g_lReleaseRate
            End If
        
        Case ucToolbar.Buttons("Minus").Value = [tbrPressed]
            
            '-- Decrease release rate by one
            If (g_lReleaseRate > g_lReleaseRateMin) Then
                g_lReleaseRate = g_lReleaseRate - 1
                lblButton(9).Caption = g_lReleaseRate
            End If
    End Select
End Sub

'== Panoramic view scrolling  ===========================================================

Private Sub ucPanView_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Call ucPanView_MouseMove(Button, Shift, x, y)
End Sub

Private Sub ucPanView_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    
    If (Me.Tag = [mPlaying]) Then
        If (Button = vbLeftButton) Then
            '-- Scroll view to x
            Call DoScrollTo(x:=x, ScaleAndCenter:=True)
        End If
    End If
End Sub

'========================================================================================
' Methods
'========================================================================================

Friend Sub LevelDone()
    
    '-- Reset pointer
    Set ucScreen.UserIcon = Nothing
    
    '-- Reset panoramic screen
    With ucPanView
        Call .DIB.Reset
        Call .Refresh
    End With
    
    '-- Disable toolbar
    Call wgSetToolbarState(bEnable:=False)
    
    '-- Show game results
    Call wgShowGameScreen
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub wgSetMenusState(ByVal bEnable As Boolean)

    '-- Enable/disable menu items
    mnuGame(0).Enabled = bEnable
    mnuGame(2).Enabled = bEnable
    mnuGame(3).Enabled = bEnable
End Sub

Private Sub wgSetToolbarState(ByVal bEnable As Boolean)
    
  Dim i As Long
    
    '-- Update toolbar buttons
    For i = 1 To ucToolbar.Buttons.Count
        ucToolbar.Buttons(i).Value = [tbrUnpressed]
        ucToolbar.Buttons(i).Enabled = bEnable
    Next i
    
    '-- Update skills labels
    For i = 1 To lblButton.Count
        lblButton(i).Caption = vbNullString
    Next i
    
    '-- Update panels text
    For i = 1 To ucInfo.Panels.Count
        ucInfo.Panels(i).Text = vbNullString
    Next i
End Sub

Private Sub wgShowLevelFeatures()
    
    With g_uLevel
    
        '-- Disable not available
        ucToolbar.Buttons(1).Enabled = .MaxClimbers
        ucToolbar.Buttons(2).Enabled = .MaxFloaters
        ucToolbar.Buttons(3).Enabled = .MaxBombers
        ucToolbar.Buttons(4).Enabled = .MaxBlockers
        ucToolbar.Buttons(5).Enabled = .MaxBuilders
        ucToolbar.Buttons(6).Enabled = .MaxBashers
        ucToolbar.Buttons(7).Enabled = .MaxMiners
        ucToolbar.Buttons(8).Enabled = .MaxDiggers
    
        '-- Show skills
        lblButton(1).Caption = .MaxClimbers
        lblButton(2).Caption = .MaxFloaters
        lblButton(3).Caption = .MaxBombers
        lblButton(4).Caption = .MaxBlockers
        lblButton(5).Caption = .MaxBuilders
        lblButton(6).Caption = .MaxBashers
        lblButton(7).Caption = .MaxMiners
        lblButton(8).Caption = .MaxDiggers
        lblButton(9).Caption = .ReleaseRate
        lblButton(10).Caption = .ReleaseRate
        
        '-- Show level title
        ucInfo.Panels(1).Text = RTrim$(.Title)
      
        '-- Show out and saved
        ucInfo.Panels(3).Text = "Out: 0/" & .LemsToLetOut
        ucInfo.Panels(4).Text = "Saved: 0%"
        
        '-- Show playing time
        ucInfo.Panels(5).Text = GetMinSecString(.PlayingTime * 60)
    End With
End Sub

Private Sub wgShowMenuScreen()

    '-- State flag/menus
    Me.Tag = [mMenuScreen]
    Call wgSetMenusState(bEnable:=True)
    
    With ucScreen
        
        '-- Tile background
        Call FXTile( _
             .DIB, 0, 0, .DIB.Width, .DIB.Height, _
             [tbLem] _
             )
        
        '-- Paint menu info
        Call FXText( _
             .DIB, 43, 35, _
             "Press [F1] or left mouse button to play", _
             IDX_YELLOW _
             )
             
        Call FXText( _
             .DIB, 16, 50, _
             "Press [F2] or right mouse button to choose level", _
             IDX_YELLOW _
             )
        
        '-- Paint options info
        If (mnuOptions(0).Checked) Then
            Call FXText( _
                 .DIB, 104, 100, _
                 "[S]ound effects  on", _
                 IDX_GREEN _
                 )
          Else
            Call FXText( _
                 .DIB, 104, 100, _
                 "[S]ound effects  off", _
                 IDX_RED _
                 )
        End If
        If (mnuOptions(1).Checked) Then
            Call FXText( _
                 .DIB, 104, 115, _
                 "[M]usic          on", _
                 IDX_GREEN _
                 )
          Else
            Call FXText( _
                 .DIB, 104, 115, _
                 "[M]usic          off", _
                 IDX_RED _
                 )
        End If
    End With
    
    '-- Refresh
    Call ucScreen.Refresh
End Sub

Private Sub wgShowLevelScreen()

    '-- State flag/menus
    Me.Tag = [mLevelScreen]
    Call wgSetMenusState(bEnable:=False)
    
    With ucScreen
        
        '-- Tile background
        Call FXTile( _
             .DIB, 0, 0, .DIB.Width, .DIB.Height, _
             [tbGround] _
             )
        
        '-- Paint title
        Call FXText( _
             .DIB, 10, 10, _
             "Level " & g_nLevelID Mod 100 + 1, _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 160 - 3 * Len(RTrim$(g_uLevel.Title)), 25, _
             g_uLevel.Title, _
             IDX_YELLOW _
             )
        
        '-- Paint level info
        Call FXText( _
             .DIB, 110, 50, _
             "Number of lems " & g_uLevel.LemsToLetOut, _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 110, 65, _
             Format$(g_uLevel.LemsToBeSaved / g_uLevel.LemsToLetOut, "0%") & " to be saved", _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 110, 80, _
             "Release rate " & g_uLevel.ReleaseRate, _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 110, 95, _
             GetMinSecString(g_uLevel.PlayingTime * 60), _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 110, 110, _
             "Rating " & GetLevelRatingString(g_nLevelID), _
             IDX_YELLOW _
             )
        
        '-- Paint 'continue' message
        Call FXText( _
             .DIB, 70, 140, _
             "Press mouse button to continue", _
             IDX_YELLOW _
             )
    End With
    
    '-- Refresh
    Call ucScreen.Refresh
End Sub

Private Sub wgShowGameScreen()

    '-- State flag
    Me.Tag = [mGameScreen]
    
    With ucScreen
        
        '-- Tile background
        Call FXTile( _
             .DIB, 0, 0, .DIB.Width, .DIB.Height, _
             [tbGround] _
             )
        
        '-- Paint game info
        Call FXText( _
             .DIB, 94, 15, _
             "All lems accounted for", _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 115, 50, _
             "You rescued " & Format$(GetLemsSaved() / g_uLevel.LemsToLetOut, "0%"), _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 115, 65, _
             "You needed  " & Format$(g_uLevel.LemsToBeSaved / g_uLevel.LemsToLetOut, "0%"), _
             IDX_YELLOW _
             )
        
        '-- Paint 'retry level' or 'continue' message
        If (GetLemsSaved() < g_uLevel.LemsToBeSaved) Then
            Call FXText( _
                 .DIB, 46, 125, _
                 "Press left mouse button to retry level", _
                 IDX_YELLOW _
                 )
            Call FXText( _
                 .DIB, 61, 140, _
                 "Press right mouse button for menu", _
                 IDX_YELLOW _
                 )
          Else
            Call FXText( _
                 .DIB, 70, 140, _
                 "Press mouse button to continue", _
                 IDX_YELLOW _
                 )
            Call SetLevelDone( _
                 g_nLevelID, Done:=True _
                 )
        End If
    End With
    
    '-- Refresh
    Call ucScreen.Refresh
End Sub
