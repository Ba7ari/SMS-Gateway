VERSION 5.00
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetailMaskapai 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detail Maskapai"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmDetailMaskapai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbTujuan 
      Height          =   315
      Left            =   9240
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin ucXPButton.XPButton cmbSemua 
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Semua No.Flight"
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
      MICON           =   "frmDetailMaskapai.frx":A4CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbnoflt 
      Height          =   315
      Left            =   9240
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin MSComctlLib.ListView LVDetail 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6165
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No.flt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tujuan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Class"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Jam Berangkat"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Harga Tiket"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Stock"
         Object.Width           =   2540
      EndProperty
   End
   Begin ucXPButton.XPButton cmdclose 
      Height          =   375
      Left            =   9120
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Close"
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
      MICON           =   "frmDetailMaskapai.frx":A4E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tujuan :"
      Height          =   255
      Left            =   9240
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   9120
      Picture         =   "frmDetailMaskapai.frx":A502
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmDetailMaskapai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbnoflt_Click()
Dim clsmaskapai As New ClassMaskapai
Dim i As Integer
Dim xDetail() As String, bagiDetail() As String
On Error Resume Next
xDetail = clsmaskapai.cari_Semua_maskapai(cmbnoflt.Text)
LVDetail.ListItems.Clear
For i = 0 To UBound(xDetail)
        bagiDetail = Split(xDetail(i), "#")
        With LVDetail.ListItems.Add(, , bagiDetail(0))
                .ListSubItems.Add , , bagiDetail(1)
                .ListSubItems.Add , , bagiDetail(2)
                .ListSubItems.Add , , bagiDetail(3)
                .ListSubItems.Add , , Format(bagiDetail(4), "Rp ###,###,###")
                .ListSubItems.Add , , bagiDetail(5)
        End With
Next
                    
End Sub

Private Sub cmbSemua_Click()
ClassDetail
End Sub

Private Sub cmbTujuan_Click()
Call ambil_jadwal_Tujuan
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call centerscreen(Me)
ClassDetail
Call ambilNoflight
Call ambilRute
End Sub
Function ClassDetail()
Dim conVtoDate As Date
Dim clsClass As New ClassKelas
Dim j As Integer
Dim xDetail() As String, bagiDetail() As String
On Error Resume Next
xDetail = clsClass.ClassDetail
LVDetail.ListItems.Clear 'jika data sudah ada ditabel maka refresh
For j = 0 To UBound(xDetail)
        bagiDetail = Split(xDetail(j), "#")
        With LVDetail.ListItems.Add(, , bagiDetail(0))
                .ListSubItems.Add , , bagiDetail(1)
                .ListSubItems.Add , , bagiDetail(2)
                .ListSubItems.Add , , bagiDetail(3)
                .ListSubItems.Add , , Format(bagiDetail(4), "Rp ###,###,###")
                .ListSubItems.Add , , bagiDetail(5)
        End With
Next
End Function
Function ambilNoflight()
Dim i As Integer
Dim xDetail() As String, bagiDetail() As String
Dim clsmaskapai As New ClassMaskapai
On Error Resume Next 'jika database kosong maka lanjut
xDetail = clsmaskapai.panggilMaskapai
cmbnoflt.Clear
For i = 0 To UBound(xDetail)
        bagiDetail = Split(xDetail(i), "#")
        cmbnoflt.AddItem (bagiDetail(0))
Next
End Function
Function ambilRute()
Dim i As Integer
Dim xroute() As String, bagiRoute() As String
Dim clsRoute As New ClassRoute
On Error Resume Next
xroute = clsRoute.ambilRoute
cmbTujuan.Clear
For i = 0 To UBound(xroute)
    bagiRoute = Split(xroute(i), "#")
    cmbTujuan.AddItem (bagiRoute(0))
Next
End Function
Function ambil_jadwal_Tujuan()
Dim i As Integer
Dim xroute() As String, bagiRoute() As String
Dim clsrute As New ClassRoute
On Error Resume Next
xroute = clsrute.cari_jadwal(cmbTujuan.Text)
LVDetail.ListItems.Clear
For i = 0 To UBound(xroute)
    bagiRoute = Split(xroute(i), "#")
        With LVDetail.ListItems.Add(, , bagiRoute(0))
                .ListSubItems.Add , , bagiRoute(1)
                .ListSubItems.Add , , bagiRoute(2)
                .ListSubItems.Add , , bagiRoute(3)
                .ListSubItems.Add , , Format(bagiRoute(4), "Rp ###,###,###")
                .ListSubItems.Add , , bagiRoute(5)
        End With
Next
End Function

