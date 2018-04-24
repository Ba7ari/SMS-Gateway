VERSION 5.00
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmLogout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logout to windows"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2865
   Icon            =   "frmLogout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   Begin ucXPButton.XPButton cmdOK 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmLogout.frx":29C12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton optLog 
         Caption         =   "Keluar aplikasi"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optLog 
         Caption         =   "Shutdown windows"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optLog 
         Caption         =   "Logout"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optLog 
         Caption         =   "Restart"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin ucXPButton.XPButton cmdBatal 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&Batal"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmLogout.frx":29C2E
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
Attribute VB_Name = "frmLogout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBatal_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If optLog(0).Value = True Then
    
    Shell "Shutdown -r -t 00": End
ElseIf optLog(1).Value = True Then
    Unload utama
    
    Load frmLOGIN
    frmLOGIN.Show 1
    Unload Me
ElseIf optLog(2).Value Then
    
    Shell "Shutdown -s -t 00": End
Else
    End
End If
End Sub
Private Sub Form_Load()
Call centerscreen(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
utama.Timer1 = False: utama.Timer2 = False
End Sub
