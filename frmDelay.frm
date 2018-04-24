VERSION 5.00
Object = "{8E37F849-94CB-11D5-B563-000021FCAE1A}#1.0#0"; "mdpBar.ocx"
Begin VB.Form frmDelay 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3600
      Top             =   120
   End
   Begin mdBar.mdpBar pb 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
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
      Max             =   30
      Value           =   1
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Tunggu...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmDelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Form_Load()
Call splashCenter(Me)
End Sub

Private Sub Timer1_Timer()
pb.Value = (pb.Value + 1)
If pb.Value = 30 Then Unload Me

End Sub

