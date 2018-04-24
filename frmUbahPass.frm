VERSION 5.00
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmUbahPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ubah Password"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   Icon            =   "frmUbahPass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin ucXPButton.XPButton cmdUbah 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&Ubah"
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
      MICON           =   "frmUbahPass.frx":23D2
      PICN            =   "frmUbahPass.frx":23EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtPB 
      Height          =   318
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtPL 
      Height          =   318
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtusrnm 
      Height          =   318
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin ucXPButton.XPButton cmdbatal 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmUbahPass.frx":3440
      PICN            =   "frmUbahPass.frx":345C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdkeluar 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&Keluar"
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
      MICON           =   "frmUbahPass.frx":E896
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "Password baru :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Password lama :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Username :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmUbahPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clspass As New ClassPass

Private Sub cmdBatal_Click()
Call clearscreen(Me)
txtusrnm.SetFocus
End Sub

Private Sub cmdkeluar_Click()
Unload Me
End Sub

Private Sub cmdUbah_Click()
If tbox_kosong(Me) = True Then
        MsgBox "Username dan password harap diisi", vbInformation + vbSystemModal, "Informasi"
Else
    If MsgBox("Username dan password akan diubah,anda yakin?", vbOKCancel + vbSystemModal + vbQuestion, "Konfirmasi") = vbOK Then
        clspass.update_pass txtusrnm.Text, txtPL.Text, txtPB.Text, 1
        Call clearscreen(Me)
        txtusrnm.SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
Call centerscreen(Me)
End Sub

Private Sub txtPL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        clspass.cek_password txtusrnm.Text, txtPL.Text
       
End If
End Sub

Private Sub txtusrnm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPL.SetFocus
End Sub
