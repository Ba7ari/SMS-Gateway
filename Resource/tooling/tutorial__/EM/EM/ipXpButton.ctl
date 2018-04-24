VERSION 5.00
Begin VB.UserControl XpButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   ScaleHeight     =   3705
   ScaleWidth      =   5835
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      Height          =   195
      Left            =   2340
      TabIndex        =   0
      Top             =   1620
      Width           =   285
   End
End
Attribute VB_Name = "XpButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const GRADIENT_FILL_RECT_H  As Long = &H0
Private Const GRADIENT_FILL_RECT_V  As Long = &H1
Private Const GRADIENT_FILL_TRIANGLE As Long = &H2
Private GRADIENT_COLOR1 As Long
Private GRADIENT_COLOR2 As Long
Private GRADIENT_FILL_RECT_DIRECTION As Long
Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Declare Function GradientFill Lib "msimg32" _
    (ByVal hdc As Long, _
     pVertex As Any, _
     ByVal dwNumVertex As Long, _
     pMesh As Any, _
     ByVal dwNumMesh As Long, _
     ByVal dwMode As Long) As Long
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click

     
'Event Click()

Private Sub Label1_Click()
    Call UserControl_Click
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub
'
'Private Sub UserControl_Click()
'    RaiseEvent Click
'End Sub

Private Sub UserControl_EnterFocus()
'    If GRADIENT_COLOR2 = &HFFFFFF Then Exit Sub
'    GRADIENT_COLOR1 = &HFFFFFF
'    GRADIENT_COLOR2 = &H3399CC
'    Call DrawGradientFill(GRADIENT_COLOR1, GRADIENT_COLOR2)
'    UserControl.Line (45, 45)-(ScaleWidth - 60, ScaleHeight - 60), RGB(255, 204, 0), B
'    UserControl.Line (60, 60)-(ScaleWidth - 75, ScaleHeight - 75), RGB(204, 153, 51), B
    Dim Aaa
    Aaa = 15
    Dim Warna As Long
    Warna = RGB(255, 128, 128)
    UserControl.Line ((Aaa + 15), Aaa)-(ScaleWidth - (Aaa + 30), Aaa), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 30), (Aaa + 15)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 15), (Aaa + 15)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 15), ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 30), ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 30), ScaleHeight - (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), ScaleHeight - (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(Aaa, ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(Aaa, (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), Aaa), Warna
    'Debug.Print "Enter"
End Sub

Private Sub UserControl_ExitFocus()
'    GRADIENT_COLOR1 = &HFFFFFF
'    GRADIENT_COLOR2 = &HCC9966
    Call DrawGradientFill(GRADIENT_COLOR1, GRADIENT_COLOR2)
    'Debug.Print "Exit"
End Sub

Private Sub UserControl_Initialize()
    GRADIENT_COLOR1 = &HFFFFFF
    GRADIENT_COLOR2 = &HCC9966
    Call DrawGradientFill(GRADIENT_COLOR1, GRADIENT_COLOR2)
    Label1.Left = 0
    Label1.Top = 50
End Sub

Private Sub UserControl_InitProperties()
    UserControl.Width = 900
    UserControl.Height = 315
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Exit Sub
    
    Dim i As Long
'    If GRADIENT_COLOR1 <> &HFFFFFF Then GRADIENT_COLOR1 = &HFFFFFF
'    If GRADIENT_COLOR2 <> &H3399CC Then GRADIENT_COLOR2 = &H3399CC
    i = GRADIENT_COLOR1
    GRADIENT_COLOR1 = GRADIENT_COLOR2
    GRADIENT_COLOR2 = i
    Call DrawGradientFill(GRADIENT_COLOR1, GRADIENT_COLOR2)
    Dim Aaa
    Aaa = 15
    Dim Warna As Long
    Warna = RGB(255, 255, 255)
    UserControl.Line ((Aaa + 15), Aaa)-(ScaleWidth - (Aaa + 30), Aaa), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 30), (Aaa + 15)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 15), (Aaa + 15)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 15), ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 30), ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 30), ScaleHeight - (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), ScaleHeight - (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(Aaa, ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(Aaa, (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), Aaa), Warna
    'Debug.Print "Down"
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    i = GRADIENT_COLOR1
    GRADIENT_COLOR1 = GRADIENT_COLOR2
    GRADIENT_COLOR2 = i
    Call DrawGradientFill(GRADIENT_COLOR1, GRADIENT_COLOR2)
'    UserControl.Line (45, 45)-(ScaleWidth - 60, ScaleHeight - 60), RGB(255, 204, 0), B
'    UserControl.Line (60, 60)-(ScaleWidth - 75, ScaleHeight - 75), RGB(204, 153, 51), B
    Dim Aaa
    Aaa = 15
    Dim Warna As Long
    Warna = RGB(255, 128, 128)
    UserControl.Line ((Aaa + 15), Aaa)-(ScaleWidth - (Aaa + 30), Aaa), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 30), (Aaa + 15)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 15), (Aaa + 15)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 15), ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 30), ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(ScaleWidth - (Aaa + 30), ScaleHeight - (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), ScaleHeight - (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(Aaa, ScaleHeight - (Aaa + 30)), Warna
    UserControl.Line -(Aaa, (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), (Aaa + 15)), Warna
    UserControl.Line -((Aaa + 15), Aaa), Warna
    'Debug.Print "Up"
End Sub

Private Sub UserControl_Resize()
    Call DrawGradientFill(GRADIENT_COLOR1, GRADIENT_COLOR2)
    Label1.Width = ScaleWidth
    'Label1.Left = (ScaleWidth - Label1.Width) \ 2
    Label1.Top = (ScaleHeight - Label1.Height) \ 2
End Sub

Private Sub DrawGradientFill(ByVal dwColour1 As Long, ByVal dwColour2 As Long)
    Dim vert(0 To 1) As TRIVERTEX
    Dim grRc As GRADIENT_RECT

    'Colour at upper-left corner
    With vert(0)
        .X = 0
        .Y = 0
        .Red = LongToSignedShort((dwColour1 And &HFF&) * 256)
        .Green = LongToSignedShort(((dwColour1 And &HFF00&) \ &H100&) * 256)
        .Blue = LongToSignedShort(((dwColour1 And &HFF0000) \ &H10000) * 256)
        .Alpha = 0
    End With

    'Colour at bottom-right corner
    With vert(1)
        .X = ScaleWidth \ Screen.TwipsPerPixelX
        .Y = ScaleHeight \ Screen.TwipsPerPixelY
        .Red = LongToSignedShort((dwColour2 And &HFF&) * 256)
        .Green = LongToSignedShort(((dwColour2 And &HFF00&) \ &H100&) * 256)
        .Blue = LongToSignedShort(((dwColour2 And &HFF0000) \ &H10000) * 256)
        .Alpha = 0
    End With
    With grRc
        .LowerRight = 0
        .UpperLeft = 1
    End With
    Cls
    Call GradientFill(hdc, vert(0), 2, grRc, 1, Abs(Not GRADIENT_FILL_RECT_DIRECTION))
    UserControl.Line (15, 0)-(ScaleWidth - 30, 0), vb3DDKShadow
    UserControl.Line -(ScaleWidth - 30, 15), vb3DDKShadow
    UserControl.Line -(ScaleWidth - 15, 15), vb3DDKShadow
    UserControl.Line -(ScaleWidth - 15, ScaleHeight - 30), vb3DDKShadow
    UserControl.Line -(ScaleWidth - 30, ScaleHeight - 30), vb3DDKShadow
    UserControl.Line -(ScaleWidth - 30, ScaleHeight - 15), vb3DDKShadow
    UserControl.Line -(15, ScaleHeight - 15), vb3DDKShadow
    UserControl.Line -(15, ScaleHeight - 30), vb3DDKShadow
    UserControl.Line -(0, ScaleHeight - 30), vb3DDKShadow
    UserControl.Line -(0, 15), vb3DDKShadow
    UserControl.Line -(15, 15), vb3DDKShadow
    UserControl.Line -(15, 0), vb3DDKShadow
    On Error Resume Next
    UserControl.Line (0, 0)-(15, 0), UserControl.BackColor
    UserControl.Line (ScaleWidth - 15, 0)-(ScaleWidth, 0), UserControl.BackColor
    UserControl.Line (ScaleWidth - 15, ScaleHeight - 15)-(ScaleWidth - 15, ScaleHeight), UserControl.BackColor
    UserControl.Line (0, ScaleHeight - 15)-(0, ScaleHeight), UserControl.BackColor
    'UserControl.Line (30, 30)-(ScaleWidth - 45, ScaleHeight - 45), UserControl.BackColor, B
End Sub

Private Function LongToSignedShort(dwUnsigned As Long) As Integer
    'convert from long to signed short
    If dwUnsigned < 32768 Then
        LongToSignedShort = CInt(dwUnsigned)
    Else
        LongToSignedShort = CInt(dwUnsigned - &H10000)
    End If
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label1.Caption = PropBag.ReadProperty("Caption", "Add")
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", Label1.Caption, "Add")
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Label1.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

