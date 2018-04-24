VERSION 5.00
Object = "{8E37F849-94CB-11D5-B563-000021FCAE1A}#1.0#0"; "mdpBar.ocx"
Begin VB.Form frmkirim 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tunggu pesan sedang dikirim"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5280
      Top             =   960
   End
   Begin mdBar.mdpBar pbar1 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
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
      Max             =   10
      Value           =   1
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   240
      ScaleHeight     =   975
      ScaleWidth      =   5415
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmkirim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim spath As String
Dim anime As New playAVI
Dim lets As Boolean
Private Sub Form_Load()
tengah
spath = App.Path & "\c_m.avi"
lets = anime.PlayAviCtrl(Picture1.hWnd, spath, 0, 1, 1, 1)
End Sub
Function tengah()
Me.Left = (utama.Height - Me.Height) / 2
Me.Top = (utama.Width - Me.Width) / 2
End Function

Private Sub Timer1_Timer()
If pbar1.Value = pbar1.Max Then
        Unload Me
End If
End Sub
