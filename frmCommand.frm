VERSION 5.00
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmCommand 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Command Report"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   Icon            =   "frmCommand.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin ucXPButton.XPButton cmdtutup 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Tutup"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCommand.frx":0442
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox ListTrans 
      Height          =   2790
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   5055
   End
   Begin VB.ListBox Listcommand 
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdtutup_Click()
Me.Hide
'utama.AutoShowChildren = False
'Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
'utama.AutoShowChildren = False
'Me.Hide
End Sub
