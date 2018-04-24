VERSION 5.00
Begin VB.Form fNewCat 
   BorderStyle     =   0  'None
   Caption         =   "New Kategori"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fNewCat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPesan 
      Height          =   960
      Left            =   255
      MaxLength       =   160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1755
      Width           =   3915
   End
   Begin VB.TextBox txtKatKode 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   945
      Width           =   2280
   End
   Begin EnterpriseManager.XpButton cmdOk 
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   2835
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
      Left            =   3270
      TabIndex        =   7
      Top             =   2835
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pesan Balasan:"
      Height          =   240
      Index           =   2
      Left            =   225
      TabIndex        =   5
      Top             =   1455
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Kode Kategori"
      Height          =   240
      Index           =   1
      Left            =   510
      TabIndex        =   4
      Top             =   990
      Width           =   1230
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Kategori Baru"
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
      Left            =   300
      TabIndex        =   2
      Top             =   195
      Width           =   3870
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   600
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4470
   End
End
Attribute VB_Name = "fNewCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBatal_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim RS As ADODB.Recordset
On Error GoTo KirimErr

    If txtKatKode = "" Then
        TampilkanPesan "Kode kategori harus diisikan"
        Exit Sub
    End If
    If txtPesan = "" Then
        TampilkanPesan "Pesan respon harus diisikan"
        Exit Sub
    End If
    Set RS = New ADODB.Recordset
    RS.Open "select teks from kategori where teks=" & VldTeks(txtKatKode), OC
    If Not RS.EOF Then
        TampilkanPesan "Kode " & AddQuote(txtKatKode) & " sudah ada"
        RS.Close
        Exit Sub
    End If
    RS.Close
    OC.Execute "insert into kategori(teks,respond,aktif) values(" & _
                VldTeks(txtKatKode) & "," & VldTeks(txtPesan) & ",1)"
    TampilkanPesan "Kategori baru telah tersimpan"
    Set RS = Nothing
    Unload Me

Exit Sub

KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub
