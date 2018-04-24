VERSION 5.00
Object = "{8E37F849-94CB-11D5-B563-000021FCAE1A}#1.0#0"; "mdpBar.ocx"
Begin VB.Form frmMulai 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerSP 
      Left            =   480
      Top             =   480
   End
   Begin mdBar.mdpBar proses 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   100
      Value           =   50
   End
   Begin VB.Label lblpsnloading 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   4440
      Left            =   0
      Picture         =   "frmMulai.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6660
   End
End
Attribute VB_Name = "frmMulai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Property Get Showsplash() As Boolean
ShowAsSplash = m_mode
End Property

Public Property Let Showsplash(ByVal vNewValue As Boolean)
m_mode = vNewValue
End Property
Private Sub Form_Load()
Call splashCenter(Me)
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
End Sub

