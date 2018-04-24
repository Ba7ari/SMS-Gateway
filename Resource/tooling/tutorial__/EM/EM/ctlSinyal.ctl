VERSION 5.00
Begin VB.UserControl ctlSinyal 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   930
   ScaleHeight     =   1020
   ScaleWidth      =   930
   Begin VB.Shape shpSinyal 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   870
      Index           =   3
      Left            =   675
      Top             =   60
      Width           =   165
   End
   Begin VB.Shape shpSinyal 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   630
      Index           =   2
      Left            =   480
      Top             =   300
      Width           =   165
   End
   Begin VB.Shape shpSinyal 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   1
      Left            =   270
      Top             =   525
      Width           =   165
   End
   Begin VB.Shape shpSinyal 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   0
      Left            =   75
      Top             =   705
      Width           =   165
   End
End
Attribute VB_Name = "ctlSinyal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_Sinyal = 0
'Property Variables:
Dim m_Sinyal As Integer
'Event Declarations:
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Sinyal() As Integer
Attribute Sinyal.VB_Description = "Jumlah sinyal diterima pada HP"
    Sinyal = m_Sinyal
End Property

Public Property Let Sinyal(ByVal New_Sinyal As Integer)
    m_Sinyal = New_Sinyal
    PropertyChanged "Sinyal"
    
    Dim i As Integer
    For i = 0 To 3
        If m_Sinyal > i Then
            shpSinyal(i).FillStyle = 0
        Else
            shpSinyal(i).FillStyle = 1
        End If
        shpSinyal(i).Refresh
    Next
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Sinyal = m_def_Sinyal
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Sinyal = PropBag.ReadProperty("Sinyal", m_def_Sinyal)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Sinyal", m_Sinyal, m_def_Sinyal)
End Sub

