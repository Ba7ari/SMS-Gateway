VERSION 5.00
Begin VB.Form fGantiPwd 
   BorderStyle     =   0  'None
   Caption         =   "Ganti Password"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   Icon            =   "fGantiPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPWD 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2130
      PasswordChar    =   "O"
      TabIndex        =   0
      Top             =   885
      Width           =   2220
   End
   Begin VB.TextBox txtPWD 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2130
      PasswordChar    =   "O"
      TabIndex        =   1
      Top             =   1425
      Width           =   2220
   End
   Begin EnterpriseManager.XpButton cmdOk 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   2010
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
   Begin EnterpriseManager.XpButton cmdBatal 
      Height          =   315
      Left            =   3150
      TabIndex        =   7
      Top             =   2010
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Password Baru"
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
      Left            =   405
      TabIndex        =   4
      Top             =   180
      Width           =   3870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password Baru"
      Height          =   195
      Index           =   1
      Left            =   855
      TabIndex        =   3
      Top             =   930
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ulangi Password Baru:"
      Height          =   195
      Index           =   2
      Left            =   315
      TabIndex        =   2
      Top             =   1425
      Width           =   1605
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   600
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4710
   End
End
Attribute VB_Name = "fGantiPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public namauser As String

Private Sub cmdBatal_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo KirimErr

If txtPWD(1) <> "" Then
    If txtPWD(1) <> txtPWD(2) Then
        TampilkanPesan "Kedua Password baru harus sama"
    Else
        OC.Execute "update pengguna set password=password(" & VldTeks(txtPWD(1)) & ") where username=" & VldTeks(namauser)
        TampilkanPesan "Password sukses diganti"
        Unload Me
    End If
Else
    TampilkanPesan "Password harus diisikan"
End If

Exit Sub

KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number

End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub
