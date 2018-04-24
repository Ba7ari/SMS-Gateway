VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fUtama 
   BorderStyle     =   0  'None
   Caption         =   "SMS Autorespond"
   ClientHeight    =   9000
   ClientLeft      =   1455
   ClientTop       =   1590
   ClientWidth     =   12000
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Square721 BT"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fUtama.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "fUtama.frx":030A
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3060
      Top             =   7665
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Pilih File Untuk Export"
      Filter          =   "Semua "
   End
   Begin EnterpriseManager.XpButton xpCmd 
      Height          =   420
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   7650
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      Caption         =   "Kirim SMS... (F5)"
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
   Begin VB.ComboBox cbStatFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10395
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   6765
      Width           =   1035
   End
   Begin VB.OptionButton optFilter 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Status Tertentu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   8430
      TabIndex        =   15
      Top             =   6735
      Width           =   1950
   End
   Begin VB.OptionButton optFilter 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tanggal Tertentu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8445
      TabIndex        =   14
      Top             =   6375
      Width           =   1980
   End
   Begin VB.OptionButton optFilter 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SMS Minggu Ini"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   8445
      TabIndex        =   13
      Top             =   6060
      Width           =   2970
   End
   Begin VB.OptionButton optFilter 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SMS Hari Ini"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   8445
      TabIndex        =   12
      Top             =   5790
      Width           =   2970
   End
   Begin VB.TextBox txtPesan 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   3615
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   5490
      Width           =   4590
   End
   Begin VB.Timer WaktuBaca 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4815
      Top             =   600
   End
   Begin VB.ComboBox cbStat 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7005
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   6690
      Width           =   1215
   End
   Begin VB.CheckBox chkAktif 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   465
      TabIndex        =   5
      Top             =   5100
      Value           =   1  'Checked
      Width           =   240
   End
   Begin VB.TextBox txtJawaban 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   450
      MaxLength       =   160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   4110
      Top             =   660
   End
   Begin MSComctlLib.ListView LVCat 
      Height          =   3180
      Left            =   450
      TabIndex        =   3
      Top             =   1725
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5609
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LVPesan 
      Height          =   3885
      Left            =   3675
      TabIndex        =   7
      Top             =   1365
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6853
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin EnterpriseManager.XpButton xpCmd 
      Height          =   315
      Index           =   1
      Left            =   10455
      TabIndex        =   18
      Top             =   6420
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "Pilih..."
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
   Begin EnterpriseManager.XpButton xpCmd 
      Height          =   420
      Index           =   7
      Left            =   1980
      TabIndex        =   19
      Top             =   6615
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      Caption         =   "Hapus"
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
   Begin EnterpriseManager.XpButton xpCmd 
      Height          =   420
      Index           =   3
      Left            =   480
      TabIndex        =   20
      Top             =   8175
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   741
      Caption         =   "Export Teks... (F6)"
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
   Begin EnterpriseManager.XpButton xpCmd 
      Height          =   420
      Index           =   4
      Left            =   3700
      TabIndex        =   21
      Top             =   8175
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      Caption         =   "Cetak... (F7)"
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
   Begin EnterpriseManager.XpButton xpCmd 
      Height          =   420
      Index           =   5
      Left            =   5250
      TabIndex        =   22
      Top             =   8175
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      Caption         =   "Statistik... (F8)"
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
   Begin EnterpriseManager.XpButton xpCmd 
      Height          =   420
      Index           =   6
      Left            =   465
      TabIndex        =   23
      Top             =   6615
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      Caption         =   "Tambah... (F2)"
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
   Begin EnterpriseManager.XpButton xpCmd 
      Height          =   420
      Index           =   2
      Left            =   2150
      TabIndex        =   26
      Top             =   8175
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      Caption         =   "Export XML..."
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
   Begin VB.Label lbCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0 SMS"
      Height          =   240
      Left            =   10035
      TabIndex        =   25
      Top             =   5460
      Width           =   1320
   End
   Begin VB.Label lbOperatoir 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator"
      BeginProperty Font 
         Name            =   "Square721 BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   7755
      TabIndex        =   24
      Top             =   150
      Width           =   3990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Batasi Tampilan:"
      Height          =   240
      Left            =   8415
      TabIndex        =   11
      Top             =   5475
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rubah Status"
      Height          =   240
      Left            =   5625
      TabIndex        =   8
      Top             =   6765
      Width           =   1260
   End
   Begin VB.Label lbAktif 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teks Jawaban Kategori"
      Height          =   240
      Left            =   780
      TabIndex        =   6
      Top             =   5100
      Width           =   2130
   End
   Begin VB.Label cmdTombol 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   2
      Left            =   10830
      TabIndex        =   2
      ToolTipText     =   "Keluar dan menutup aplikasi"
      Top             =   7530
      Width           =   990
   End
   Begin VB.Label cmdTombol 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   1
      Left            =   9915
      TabIndex        =   1
      ToolTipText     =   "Mengunci program"
      Top             =   7530
      Width           =   975
   End
   Begin VB.Label cmdTombol 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   0
      Left            =   8715
      TabIndex        =   0
      ToolTipText     =   "Melakukan setting program"
      Top             =   7575
      Width           =   1200
   End
End
Attribute VB_Name = "fUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ConnectOff = 0
Const ConnectWait = 1
Const ConnectOn = 2

Dim cntImgConnect As Integer
Dim cntPrgSinyal As Integer
Dim cntPrgBattery As Integer

Dim PesanSource As String, sFilter As String
Dim GlobalPesanSource1 As String, GlobalPesanSource2 As String
Public TglFilter1 As Date, TglFilter2 As Date
Public DateChanged As Boolean

Dim ItemBaca As MSComctlLib.ListItem

Private Sub ProsesExport(ByVal ItemName As String)
Select Case ItemName
Case "mExportText"
    CD1.Filter = "File Teks|*.txt|Semua File|*.*"
Case "mExportXML"
    CD1.Filter = "File XML|*.xml|Semua File|*.*"
End Select

On Error GoTo Batal
CD1.ShowSave

OC.Execute "delete from temppesan"
OC.Execute "insert into temppesan(pengirim,teks,waktu) " & _
        GlobalPesanSource1
Screen.MousePointer = vbHourglass
Select Case ItemName
Case "mExportText"
    SimpanFileText CD1.FileName
Case "mExportXML"
    SimpanFileXML CD1.FileName
End Select
Screen.MousePointer = vbDefault
TampilkanPesan "Export selesai"
Exit Sub

Batal:
End Sub

Private Sub cbStat_Click()
If LVPesan.SelectedItem.SubItems(3) = cbStat Then Exit Sub
On Error GoTo KirimErr

OC.Execute "update pesan set flag=" & cbStat.ListIndex & " where id=" & Mid(LVPesan.SelectedItem.Key, 2)
LVPesan.SelectedItem.SubItems(3) = cbStat

Exit Sub
KirimErr:
    TampilkanPesan "Kesalahan dalam menulis ke database" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub

Private Sub cbStatFilter_Click()
sFilter = "flag=" & cbStatFilter.ListIndex
TampilkanSMS
End Sub

Private Sub chkAktif_Click()
txtJawaban.Enabled = (chkAktif.Value = 1)
txtJawaban.BackColor = IIf(txtJawaban.Enabled, &HFFFFC0, &HC0C0C0)
KategoriList(KategoriIdx).Aktif = (chkAktif.Value = 1)
End Sub

Private Sub cmdTombol_Click(Index As Integer)
On Error GoTo KirimErr

Select Case Index
Case 0
    fConfig.Show vbModal, Me
Case 1
    fAppLock.Show vbModal, Me
Case 2
    FlushData
    Unload Me
End Select

Exit Sub
KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number

End Sub

Private Sub cmdTombol_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim e As Label
For Each e In cmdTombol
    If e.Index <> Index Then e.BorderStyle = 0
Next
cmdTombol(Index).BorderStyle = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 113: xpCmd_Click 6
Case 116: xpCmd_Click 0
Case 117: xpCmd_Click 3
Case 118: xpCmd_Click 4
Case 119: xpCmd_Click 5
End Select
End Sub

Private Sub optFilter_Click(Index As Integer)
xpcmd(1).Enabled = (Index = 2)
cbStatFilter.Enabled = (Index = 3)

Select Case Index
Case 0
    sFilter = "date_format(waktu,'%Y/%m/%d')='" & Format(Now, "yyyy/MM/dd") & "'"
Case 1
    sFilter = "date_format(waktu,'%Y/%m/%d') between '" & Format(DateAdd("d", -7, Now), "yyyy/MM/dd") & "' and '" & Format(Now, "yyyy/MM/dd") & "'"
Case 2
    sFilter = "waktu between '" & Format(TglFilter1, "yyyy/MM/dd") & "' and '" & Format(TglFilter2, "yyyy/MM/dd") & "'"
Case 3
    sFilter = "flag=" & cbStatFilter.ListIndex
End Select

TampilkanSMS
End Sub

Private Sub xpCmd_Click(Index As Integer)
Dim i As Integer

On Error GoTo KirimErr

Select Case Index
Case 0
    fSendSMS.Show vbModal, Me
    
Case 1
    fPilihTanggal.Show vbModal, Me
    If DateChanged Then
        sFilter = "waktu between '" & Format(TglFilter1, "yyyy/MM/dd") & "' and '" & Format(TglFilter2, "yyyy/MM/dd") & "'"
        TampilkanSMS
    End If
    
Case 2
    ProsesExport "mExportXML"

Case 3
    ProsesExport "mExportText"

Case 4
    OC.Execute "delete from temppesan"
    OC.Execute "insert into temppesan(pengirim,teks,waktu) " & _
               GlobalPesanSource1
               
    DBConnection.rssqlPesan.Open
    drPesan.Show vbModal, Me
    DBConnection.rssqlPesan.Close
    
Case 5
    fStat.Show vbModal, Me
    
Case 6
    fNewCat.Show vbModal, Me
    ShowKategori
    
Case 7
    If TampilkanTanya("Yakin kategori " & AddQuote(KategoriList(KategoriIdx).Teks) & " akan dihapus?") = BS_Button1 Then
        OC.Execute "delete from kategori where id=" & KategoriList(KategoriIdx).Kode
        OC.Execute "update pesan set kategori=0 where kategori=" & KategoriList(KategoriIdx).Kode
        TampilkanPesan "Kategori sudah dihapus"
        ShowKategori
    Else
        TampilkanPesan "Kategori batal dihapus"
    End If
End Select

Exit Sub
KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub

Private Sub Form_Load()
InitApp

SetupListView
Timer1.Interval = 10
Timer1.Enabled = True

TglFilter1 = Now
TglFilter2 = Now

optFilter(0).Value = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim e As Label
For Each e In cmdTombol
    If e.BorderStyle = 1 Then e.BorderStyle = 0
Next
End Sub

Sub SetupListView()
Dim i As Integer

With cbStat
    .Clear
    .AddItem strBaru
    .AddItem strTerbaca
    .AddItem strDiproses
    .AddItem strDitandai
End With

For i = 0 To cbStat.ListCount - 1
    cbStatFilter.AddItem cbStat.List(i)
Next
If cbStatFilter.ListCount > 0 Then cbStatFilter.ListIndex = 0

LVPesan.ColumnHeaders.Clear
LVPesan.ColumnHeaders.Add , , "Pengirim", LVPesan.Width * 0.2
LVPesan.ColumnHeaders.Add , , "Pesan", LVPesan.Width * 0.4
LVPesan.ColumnHeaders.Add , , "Waktu", LVPesan.Width * 0.2
LVPesan.ColumnHeaders.Add , , "Status", LVPesan.Width * 0.15

LVCat.ColumnHeaders.Clear
LVCat.ColumnHeaders.Add , , "Kategori", LVCat.Width - 75
jKategori = 0
ShowKategori
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
DeInitApp
End Sub

Private Sub lbPesan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub LVCat_ItemClick(ByVal Item As MSComctlLib.ListItem)
KategoriIdx = Item.Index - 1
txtJawaban = Trim(KategoriList(KategoriIdx).Pesan)
txtJawaban_Change

PesanSource = "select id,pengirim,teks,waktu,flag as stts from pesan"
GlobalPesanSource2 = "select pengirim,pesan.teks,waktu from pesan"
Select Case Item.Index

Case 1
    txtJawaban.Visible = False
    chkAktif.Visible = False
    lbAktif.Visible = False
    chkAktif.Value = 1
    xpcmd(7).Visible = False
    
Case 2
    txtJawaban.Visible = True
    chkAktif.Visible = True
    chkAktif.Enabled = False
    chkAktif.Value = 1
    lbAktif.Visible = True
    xpcmd(7).Visible = False
    PesanSource = PesanSource & " where kategori=0"
    GlobalPesanSource2 = GlobalPesanSource2 & " where kategori=0"
    
Case Else
    txtJawaban.Visible = True
    chkAktif.Enabled = True
    chkAktif.Visible = True
    chkAktif.Value = IIf(KategoriList(KategoriIdx).Aktif, 1, 0)
    lbAktif.Visible = True
    xpcmd(7).Visible = True
    PesanSource = PesanSource & " where kategori=" & KategoriList(KategoriIdx).Kode
    GlobalPesanSource2 = GlobalPesanSource2 & " where kategori=" & KategoriList(KategoriIdx).Kode
    
End Select

TampilkanSMS
End Sub

Sub TampilkanSMS()
Dim l As MSComctlLib.ListItem
Dim RS As ADODB.Recordset

Dim strSQL As String

If PesanSource = "" Then Exit Sub

Set RS = New ADODB.Recordset
LVPesan.ListItems.Clear
txtPesan = ""

Screen.MousePointer = vbHourglass

On Error GoTo SelectErr

strSQL = PesanSource
If sFilter <> "" Then
    If KategoriIdx = 0 Then
        strSQL = strSQL & " where " & sFilter
        GlobalPesanSource1 = GlobalPesanSource2 & " where " & sFilter
    Else
        strSQL = strSQL & " and " & sFilter
        GlobalPesanSource1 = GlobalPesanSource2 & " and " & sFilter
    End If
End If

RS.Open strSQL, OC
While Not RS.EOF
    Set l = LVPesan.ListItems.Add(, "s" & RS("id"), RS("pengirim"))
    l.SubItems(1) = RS("teks")
    l.SubItems(2) = Format(RS("waktu"), "ddd, dd MMM yy hh:mm")
    l.SubItems(3) = IIf(RS("stts") = 0 Or IsNull(RS("stts")), strBaru, cbStat.List(RS("stts")))
    RS.MoveNext
Wend
lbCount = LVPesan.ListItems.Count & " SMS"
RS.Close
Set RS = Nothing
If LVPesan.ListItems.Count > 0 Then
    LVPesan_ItemClick LVPesan.ListItems(1)
    cbStat.Enabled = True
Else
    cbStat.Enabled = False
End If

Screen.MousePointer = vbDefault
Exit Sub
SelectErr:
    TampilkanPesan "Terjadi kesalahan dalam pembacaan data" & vbCrLf & _
                   "Pesan: " & Err.Description & vbCrLf & _
                   "Kode: " & Err.Number
    Screen.MousePointer = vbDefault
End Sub

Private Sub LVPesan_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtPesan = "Pengirim: " & Item & vbCrLf & "Pesan:" & vbCrLf & Item.SubItems(1)
cbStat = Item.SubItems(3)
WaktuBaca.Enabled = True
End Sub

Private Sub Timer1_Timer()

If Not LoggedIn Then
    Timer1.Enabled = False
    
    fAppLock.Show vbModal, Me
    
    If Not LoggedIn Then
        Unload Me
        Exit Sub
    End If
End If

BacaSMS
End Sub

Sub ShowKategori()
Dim RS As ADODB.Recordset
Dim i As MSComctlLib.ListItem

If jKategori > 2 Then FlushData

On Error GoTo KirimErr

jKategori = 2
LVCat.ListItems.Clear
Set RS = New ADODB.Recordset
ReDim Preserve KategoriList(0 To jKategori)
With KategoriList(0)
    .Kode = 0
    .Pesan = ""
    Set i = LVCat.ListItems.Add(, "semua", "(semua)")
End With

RS.Open "select * from kategori order by id asc", OC
If Not RS.EOF Then
    With KategoriList(1)
        .Kode = 0
        .Pesan = RS("Respond")
        .Aktif = RS("Aktif")
        Set i = LVCat.ListItems.Add(, "default", "Default")
    End With
End If
RS.MoveNext
While Not RS.EOF
    ReDim Preserve KategoriList(0 To jKategori)
    With KategoriList(jKategori)
        .Kode = RS("ID")
        .Teks = RS("Teks")
        .Pesan = RS("Respond")
        .Aktif = RS("Aktif")
        Set i = LVCat.ListItems.Add(, "str" & .Kode, .Teks)
    End With
    RS.MoveNext
    jKategori = jKategori + 1
Wend
RS.Close
Set RS = Nothing

LVCat_ItemClick LVCat.ListItems(1)

Exit Sub
KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub

Private Sub txtJawaban_Change()
KategoriList(KategoriIdx).Pesan = txtJawaban
txtJawaban.BackColor = IIf(Len(txtJawaban) = 0, &H8080FF, &HFFFFC0)
End Sub

Sub FlushData()
Dim j As Integer

On Error GoTo KirimErr

For j = 1 To jKategori - 1
    OC.Execute "update kategori set respond=" & VldTeks(KategoriList(j).Pesan) & _
                ",aktif=" & IIf(KategoriList(j).Aktif, 1, 0) & _
                " where id=" & KategoriList(j).Kode
Next

Exit Sub
KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub

Sub BacaSMS()
Dim i As Integer
Dim Tgl As Date, Pengirim As String, Tujuan As String, Pesan As String

Screen.MousePointer = vbHourglass
'lbStatus = "Membaca SMS..."
Refresh
'lbStatus = ""
Screen.MousePointer = vbDefault
End Sub

Private Sub WaktuBaca_Timer()

On Error GoTo KirimErr
If LVPesan.ListItems.Count > 0 Then
    If cbStat = strBaru Then
        OC.Execute "update pesan set flag=" & smsTerbaca & " where id=" & Mid(LVPesan.SelectedItem.Key, 2)
        LVPesan.SelectedItem.SubItems(3) = strTerbaca
        cbStat = strTerbaca
    End If
End If
WaktuBaca.Enabled = False
Exit Sub

KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub
