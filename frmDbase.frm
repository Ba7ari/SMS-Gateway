VERSION 5.00
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmDbase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Setup"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frmDbase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   3975
      Begin VB.Label Label1 
         Caption         =   "Utility ini digunakan untuk menghapus seluruh Data."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Note:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
   End
   Begin ucXPButton.XPButton cmdHapus 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "Pelanggan"
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
      MICON           =   "frmDbase.frx":000C
      PICN            =   "frmDbase.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdHapus 
      Height          =   495
      Index           =   1
      Left            =   2320
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "Reservasi"
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
      MICON           =   "frmDbase.frx":973A
      PICN            =   "frmDbase.frx":9756
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdHapus 
      Height          =   495
      Index           =   2
      Left            =   2320
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "Maskapai"
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
      MICON           =   "frmDbase.frx":26CF0
      PICN            =   "frmDbase.frx":26D0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdHapus 
      Height          =   495
      Index           =   3
      Left            =   2320
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "Route"
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
      MICON           =   "frmDbase.frx":311E6
      PICN            =   "frmDbase.frx":31202
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdHapus 
      Height          =   495
      Index           =   4
      Left            =   4560
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "Class"
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
      MICON           =   "frmDbase.frx":3159C
      PICN            =   "frmDbase.frx":315B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdHapus 
      Height          =   375
      Index           =   5
      Left            =   5040
      TabIndex        =   8
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "OK"
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
      MICON           =   "frmDbase.frx":31B52
      PICN            =   "frmDbase.frx":31B6E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line4 
      X1              =   3600
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line3 
      X1              =   3000
      X2              =   3000
      Y1              =   1560
      Y2              =   2280
   End
   Begin VB.Line Line2 
      X1              =   3000
      X2              =   3000
      Y1              =   600
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   2520
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmDbase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsPlg As New ClassPelanggan
Dim clsReserv As New ClassAmbil
Dim clsmaskapai As New ClassMaskapai
Dim clsrute As New ClassRoute
Dim clsClass As New ClassKelas
Enum cmdHapus
    hpspelanggan = 0
    hpsreservasi = 1
    hpsmaskapai = 2
    hpsroute = 3
    hpsclass = 4
    tmbKeluar = 5
End Enum
Private Sub cmdHapus_Click(Index As Integer)
If Index = tmbKeluar Then
    Unload Me

Else
    If MsgBox("Seluruh data akan dihapus,anda yakin?", vbQuestion + vbYesNo + vbSystemModal, "Konfirmasi") = vbYes Then
        Me.MousePointer = vbHourglass
        Select Case Index
                Case hpspelanggan
                        clsPlg.hapus_pelanggan 'hapus semua data pelanggan
                Case hpsreservasi
                        clsReserv.hapus_reservasi 'hapus semua data reservasi
                Case hpsroute
                        clsrute.hapus_semuaRute 'hapus semua data rute
                Case hpsmaskapai
                        clsmaskapai.hapus_maskapai 'hapus semua data maskapai
                Case hpsclass
                        clsClass.hapus_kelas 'hapus semua data kelas
                
        End Select
        Me.MousePointer = vbDefault
       End If
End If

End Sub

Private Sub Form_Load()
Call centerscreen(Me)
End Sub
