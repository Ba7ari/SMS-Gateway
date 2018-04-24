VERSION 5.00
Object = "{7FCAEF84-D390-11D0-8849-006097BFD99B}#2.0#0"; "FORMX.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "OT/X DockingForms"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "frmMain"
   StartUpPosition =   3  'Windows Default
   Begin FormXLib.OtxToolbar OtxToolbar1 
      Left            =   360
      Top             =   120
      _Version        =   131072
      _ExtentX        =   4683
      _ExtentY        =   873
      _StockProps     =   0
      Caption         =   "OT/X ToolBar 1"
      CommandList     =   ""
   End
   Begin FormXLib.OtxCommandManager OtxCommandManager1 
      Left            =   1200
      Top             =   2280
      _cx             =   953
      _cy             =   953
      DefaultTextAlignment=   0
      DisplayMode     =   3
      LargeIcons      =   0   'False
      CoolLook        =   -1  'True
      ToolTips        =   -1  'True
      HighWaterMark   =   58368
      SmallImageWidth =   16
      SmallImageHeight=   16
      LargeImageWidth =   32
      LargeImageHeight=   32
      Count           =   1
      TBData1         =   "Caption=OT/X ToolBar 1,Visible=1"
      CommandGroupCount=   0
   End
   Begin FormXLib.MDIFormX frmxMain 
      Left            =   3000
      Top             =   960
      _Version        =   131072
      _ExtentX        =   953
      _ExtentY        =   953
      _StockProps     =   0
      DefCoolLookBorder=   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    ' N.B.  SystemsParametersInfo() is used for screen dimensions instead
    '       instead of Screen.Width and Screen.Height to account for systray.

Private Sub MDIForm_Load()

    ' make main form large:  entire workspace less 10% margins on each side
    Dim workspace(1 To 4) As Long ' (1)=left (2)=top (3)=right (4)=bottom
    SystemParametersInfo &H30, 0, workspace(1), 0 ' &h30 = get workspace size
    Dim x As Long, y As Long, w As Long, h As Long
    x = workspace(1) * Screen.TwipsPerPixelX
    y = workspace(2) * Screen.TwipsPerPixelY
    w = workspace(3) * Screen.TwipsPerPixelX - x
    h = workspace(4) * Screen.TwipsPerPixelY - y
    Move x + 0.1 * w, y + 0.1 * h, 0.8 * w, 0.8 * h ' 10% margins
    
    ' show the child forms
    '      The names of these forms are hard-wired based on the project template "OTX Docking Forms.vbp".
    '      Simply delete the lines that do not apply if you diverge from that template.
    frmDocked.Show
    frmDockedOnly.Show
    frmMDIChild.Show
    frmStdMDIChild.Show
    frmFloating.Show
    frmFloatingOnly.Show
End Sub

Public Property Get MeX() As MDIFormX
    ' Property MeX is the MDIFormX equivalent of the keyword Me
    Set MeX = frmxMain
    ' N.B. A more general definition for MeX that does not hardwire the name of the embedded MDIFormX control
    '      is as follows.  The more general definition is less efficient because it calls a helper function.
    '   Public Property Get MeX() as MDIFormX
    '       Set MeX = otxMDIFormX()
    '   End Property
    '
    '   The general definition is convenient if you paste the definition into MDIForm modules
    '   that use different names for the embedded MDIFormX control.
End Property

Public Property Get DefAutoFrame() As CollectionX
    Set DefAutoFrame = MeX.DefAutoFrame
End Property

Public Property Set DefAutoFrame(ByVal NewValue As CollectionX)
    Set MeX.DefAutoFrame = NewValue
End Property

Public Property Get DefCanDragDocked() As Boolean
    DefCanDragDocked = MeX.DefCanDragDocked
End Property

Public Property Let DefCanDragDocked(ByVal NewValue As Boolean)
    MeX.DefCanDragDocked = NewValue
End Property

Public Property Get DefClientEdge() As Boolean
    DefClientEdge = MeX.DefClientEdge
End Property

Public Property Let DefClientEdge(ByVal NewValue As Boolean)
    MeX.DefClientEdge = NewValue
End Property

Public Property Get DefUseCaptionButton() As Boolean
    DefUseCaptionButton = MeX.DefUseCaptionButton
End Property

Public Property Let DefUseCaptionButton(ByVal NewValue As Boolean)
    MeX.DefUseCaptionButton = NewValue
End Property

Public Property Get DefUseCaptionDblClick() As Boolean
    DefUseCaptionDblClick = MeX.DefUseCaptionDblClick
End Property

Public Property Let DefUseCaptionDblClick(ByVal NewValue As Boolean)
    MeX.DefUseCaptionDblClick = NewValue
End Property

Public Property Get DefUseCaptionRtDblClick() As Boolean
    DefUseCaptionRtDblClick = MeX.DefUseCaptionRtDblClick
End Property

Public Property Let DefUseCaptionRtDblClick(ByVal NewValue As Boolean)
    MeX.DefUseCaptionRtDblClick = NewValue
End Property

Public Property Get FrameType() As otxMDIFormFrameType
    FrameType = MeX.FrameType
End Property

Public Property Let FrameType(ByVal NewValue As otxMDIFormFrameType)
    MeX.FrameType = NewValue
End Property

Public Function Forms(Optional ByVal Types, Optional ByVal Options, Optional ByVal Attribs) As CollectionX
    Set Forms = MeX.Forms(Types, Options, Attribs)
End Function
