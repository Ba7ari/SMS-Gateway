VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmambil 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Pemesanan"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmPesan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   10695
   End
   Begin VB.ListBox lstNama 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   840
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   10695
      Begin VB.PictureBox cmdkeluar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   2
         ToolTipText     =   "Apakah anda ingin keluar dari tampilan ini?"
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox cmdSemua 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         ScaleHeight     =   315
         ScaleWidth      =   2355
         TabIndex        =   3
         ToolTipText     =   "Tampilkan semua pemesanan"
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSComctlLib.ListView LVTiket 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tgl.Pesan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hp"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Blm.Bayar"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Pass"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total Bayar"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "No.Rekening"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bank"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   5535
      Begin VB.TextBox txthp 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtNama 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cmbkdpsn 
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox cmdcari 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "No.Hp"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Nama"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "No.Booking"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Confirm :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Hold :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Pelanggan sudah melakukan pengambilan tiket"
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Pelanggan belum melakukan pengambilan pada tiket yang dipesan."
      Height          =   495
      Left            =   6480
      TabIndex        =   13
      Top             =   960
      Width           =   3855
   End
End
Attribute VB_Name = "frmambil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim clsPlg As New ClassPelanggan
Dim clsAmbil As New ClassAmbil
Dim xData() As String, bagixData() As String

Private Sub cmbkdpsn_Click()
Dim clsAmbil As New ClassAmbil
Dim bagiData() As String
bagiData = Split(clsAmbil.cari__pemesan(cmbkdpsn.Text))
txthp.Text = bagiData(0)
txtNama.Text = bagiData(1)
End Sub

Private Sub cmdcari_Click()
semua_pemesanan cmbkdpsn.Text

End Sub

Private Sub cmdkeluar_Click()
Unload Me
End Sub

Private Sub cmdSemua_Click()
semua_pemesanan ""
End Sub

Private Sub Form_Click()
lstNama.Visible = False
End Sub

Private Sub Form_Load()
Call centerscreen(Me)
lstNama.Top = 1680
lstNama.Left = 840
ambil_Kode_reservasi
semua_pemesanan ""
End Sub

Private Sub lstNama_DblClick()
txtNama.Text = lstNama.Text
lstNama.Visible = False
txtNama.SetFocus
End Sub

Private Sub txtNama_Change()
'On Error Resume Next
clsPlg.status_kosong = False
    xData = clsPlg.cari_status_pelanggan(txtNama.Text)
    If clsPlg.status_kosong Then
        lstNama.Visible = False
    Else
        If Len(txtNama.Text) = 0 Then
                lstNama.Visible = False
        Else
            lstNama.Visible = True
            lstNama.Clear
                For i = 0 To UBound(xData)
                    lstNama.AddItem (xData(i))
                    lstNama.Selected(i) = True
                    
                Next
        End If
    End If
        
End Sub
Function ambil_Kode_reservasi()
Dim j As Integer
Dim clsreservasi As New ClassAmbil
On Error Resume Next
cmbkdpsn.Clear
xData = clsreservasi.Add_reservasi
For j = 0 To UBound(xData)
        cmbkdpsn.AddItem (xData(j))
Next
End Function
Function semua_pemesanan(ByVal kdPsn As String)

Dim i As Integer
Dim xData() As String, bagixData() As String
On Error Resume Next
LVTiket.ListItems.Clear
'LVStatus.ListItems.Clear
xData = clsAmbil.ambil_tiket(kdPsn)
For i = 0 To UBound(xData)
    bagixData = Split(xData(i), "#")
    With LVTiket.ListItems.Add(, , bagixData(1))
            .ListSubItems.Add , , bagixData(0)
            .ListSubItems.Add , , bagixData(2)
            .ListSubItems.Add , , bagixData(3)
            .ListSubItems.Add , , bagixData(4)
            .ListSubItems.Add , , Format(bagixData(5), "Rp ###,###,###")
            .ListSubItems.Add , , bagixData(6)
            .ListSubItems.Add , , bagixData(7)
    End With
Next














'Dim j As Integer
'xData = clsAmbil.ambil_tiket
'LVTiket.ListItems.Clear
'For j = 1 To UBound(xData)
'    bagixData = Split(xData(j), "#")
'    With LVTiket.ListItems.Add(, , bagixData(1))
'            .ListSubItems.Add , , bagixData(0)
'            .ListSubItems.Add , , bagixData(2)
'            .ListSubItems.Add , , bagixData(3)
'            .ListSubItems.Add , , bagixData(4)
'            .ListSubItems.Add , , bagixData(5)
'            .ListSubItems.Add , , bagixData(6)
'            .ListSubItems.Add , , bagixData(7)
'
'    End With
'Next
            
End Function
