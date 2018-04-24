VERSION 5.00
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Password"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   Icon            =   "frmPass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin ucXPButton.XPButton cmdpass 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&Simpan"
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
      MICON           =   "frmPass.frx":23D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chk 
      Caption         =   "Pakai password ?"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtusrnm 
      Height          =   318
      Left            =   1080
      MaxLength       =   13
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtPass 
      Height          =   318
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   13
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin ucXPButton.XPButton cmdkeluar 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
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
      MICON           =   "frmPass.frx":23EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Password :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cek_setting As String

Private Sub chk_Click()
If chk.Value = 1 Then
        SaveSetting "MandalaGateway", "Pass", "Settings", "1"
Else
        SaveSetting "MandalaGateway", "Pass", "Settings", "0"
End If

End Sub

Private Sub cmdkeluar_Click()
Unload Me
End Sub

Private Sub cmdpass_Click()
Dim clsPas As New ClassPass
If tbox_kosong(Me) Then
    MsgBox "Isi username dan password anda"
Else
    clsPas.simpan_pass txtusrnm.Text, txtpass.Text
    txtusrnm.Text = "": txtpass.Text = ""
    txtusrnm.SetFocus
End If
End Sub

Private Sub Form_Load()
Call centerscreen(Me)
AccessRegistrySetting 'cek setting password
End Sub
Function AccessRegistrySetting(Optional ByVal nilai As Boolean)
cek_setting = GetSetting("MandalaGateway", "Pass", "Settings")
If Len(cek_setting) <> 0 Then
        If cek_setting = "1" Then
            chk.Value = 1
        Else
            chk.Value = 0
        End If
Else
       SaveSetting "MandalaGateway", "Pass", "Settings", "0"
End If
End Function

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdpass_Click
End Sub

Private Sub txtusrnm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtpass.SetFocus
End Sub
