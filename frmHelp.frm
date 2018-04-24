VERSION 5.00
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help ?"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "1. Reg<spasi>[nama]<spasi>[alamat]"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "2. Info<spasi>rute"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "3. Tarif<spasi>[No.flt]"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "4. Info<spasi>jadwal<spasi>[Nm.tujuan]"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "5. Info<spasi>Booking"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "6. Info<spasi>flt<spasi>[Nm.tujuan]"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         Caption         =   "7. Booking<spasi>[No.flt]<spasi>[Pen.Dws]<spasi>[Pen.Byi]<spasi>[tgl.brkt]"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   5415
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000009&
         Caption         =   "9. Confirm<spasi>[No.Booking]<spasi>[No.Rek]<spasi>[Nm.Bank]"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   5415
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000009&
         Caption         =   "8. Batal<spasi>[No.Booking]"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   5415
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000009&
         Caption         =   "10. Status<spasi>[No.Booking]"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   2295
      End
   End
   Begin ucXPButton.XPButton cmdOK 
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&OK"
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
      MICON           =   "frmHelp.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call centerscreen(Me)
End Sub

