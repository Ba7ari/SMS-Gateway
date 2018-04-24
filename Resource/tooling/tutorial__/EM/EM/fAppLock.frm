VERSION 5.00
Begin VB.Form fAppLock 
   BorderStyle     =   0  'None
   Caption         =   "AppLock"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   Icon            =   "fAppLock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin EnterpriseManager.XpButton cmdOk 
      Height          =   315
      Left            =   1935
      TabIndex        =   6
      Top             =   1935
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      Caption         =   "Ok"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtNama 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1905
      TabIndex        =   0
      Top             =   1035
      Width           =   1890
   End
   Begin VB.TextBox txtPwd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1905
      PasswordChar    =   "O"
      TabIndex        =   1
      Top             =   1455
      Width           =   1890
   End
   Begin EnterpriseManager.XpButton cmdBatal 
      Height          =   315
      Left            =   2925
      TabIndex        =   7
      Top             =   1935
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      Caption         =   "Batal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "fAppLock.frx":000C
      Top             =   1065
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Username dan Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   5
      Top             =   165
      Width           =   3870
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   600
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      Height          =   195
      Left            =   975
      TabIndex        =   3
      Top             =   1095
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   1005
      TabIndex        =   2
      Top             =   1515
      Width           =   690
   End
End
Attribute VB_Name = "fAppLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBatal_Click()
LoggedIn = False
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim RS As ADODB.Recordset

On Error GoTo KirimErr

Set RS = New ADODB.Recordset
RS.Open "select * from pengguna where aktif=1 and " & _
        "username=" & VldTeks(txtNama) & " and password=password(" & VldTeks(txtPWD) & ")", OC
If Not RS.EOF Then
    If Not LoggedIn Then
        WriteLog "Pengguna masuk: " & RS("nama")
        TampilkanPesan "Selamat datang " & VldTeks(RS("nama"))
    End If
    LoggedIn = True
    fUtama.lbOperatoir = "User: " & RS("Nama")
    RS.Close
    Set RS = Nothing
    SaveSetting App.Title, "Setting", "LastUser", txtNama
    Unload Me
    Exit Sub
Else
    TampilkanPesan "Username atau Password tidak benar"
    txtPWD = ""
End If
RS.Close
Set RS = Nothing

Exit Sub

KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOk_Click
End Sub

Private Sub Form_Load()
cmdBatal.Visible = Not LoggedIn
txtNama = GetSetting(App.Title, "Setting", "LastUser", "")
If txtNama <> "" Then
    SendKeys vbTab
Else
    txtNama_Change
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub txtNama_Change()
cmdOk.Enabled = (Len(txtNama) > 0) And (Len(txtPWD) > 0)
End Sub

Private Sub txtPwd_Change()
txtNama_Change
End Sub
