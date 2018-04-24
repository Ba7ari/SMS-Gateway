VERSION 5.00
Begin VB.UserControl ctlBattery 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   ScaleHeight     =   1020
   ScaleWidth      =   720
   Begin VB.Shape shpBattery 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   60
      Top             =   195
      Width           =   540
   End
   Begin VB.Shape shpBattery 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   60
      Top             =   315
      Width           =   540
   End
   Begin VB.Shape shpBattery 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   60
      Top             =   435
      Width           =   540
   End
   Begin VB.Shape shpBattery 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   60
      Top             =   555
      Width           =   540
   End
   Begin VB.Shape shpBattery 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   60
      Top             =   675
      Width           =   540
   End
   Begin VB.Shape shpBattery 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   60
      Top             =   795
      Width           =   540
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   180
      Top             =   90
      Width           =   315
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   60
      Top             =   195
      Width           =   540
   End
End
Attribute VB_Name = "ctlBattery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_Power = 0
'Property Variables:
Dim m_Power As Integer
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
Public Property Get Power() As Integer
Attribute Power.VB_Description = "Power battery"
    Power = m_Power
End Property

Public Property Let Power(ByVal New_Power As Integer)
    m_Power = New_Power
    PropertyChanged "Power"
    
    Dim i As Integer
    For i = 0 To 5
        If m_Power > i Then
            shpBattery(i).FillStyle = 0
        Else
            shpBattery(i).FillStyle = 1
        End If
        shpBattery(i).Refresh
    Next
    
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Power = m_def_Power
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Power = PropBag.ReadProperty("Power", m_def_Power)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Power", m_Power, m_def_Power)
End Sub

