VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form fStat 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   1050
   ClientTop       =   2520
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbBulan 
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5280
      Width           =   2700
   End
   Begin EnterpriseManager.XpButton cmdTutup 
      Height          =   360
      Left            =   4920
      TabIndex        =   1
      Top             =   5280
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   635
      Caption         =   "Tutup"
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
   Begin MSChart20Lib.MSChart MC1 
      Height          =   4425
      Left            =   90
      OleObjectBlob   =   "fStat.frx":0000
      TabIndex        =   0
      Top             =   675
      Width           =   6510
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah SMS Pehari"
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
      Left            =   1290
      TabIndex        =   4
      Top             =   150
      Width           =   3870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pilih Bulan:"
      Height          =   195
      Left            =   330
      TabIndex        =   2
      Top             =   5310
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   600
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6660
   End
End
Attribute VB_Name = "fStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As ADODB.Recordset

Private Sub cbBulan_Click()
Dim i As Integer
RS.Open "select count(*) as jumlah,dayofmonth(waktu) as tgl from pesan where date_format(waktu,'%M %Y')=" & VldTeks(cbBulan) & " group by tgl order by tgl asc", OC
With MC1
    .ColumnCount = 1
    .RowCount = 1
    .Column = 1
    i = 1
    While i < 31
        If Not RS.EOF Then
            While i <> RS("tgl")
                If i > .RowCount Then .RowCount = i
                .Row = i
                .RowLabel = i
                .Data = 0
                i = i + 1
            Wend
            If i = RS("tgl") Then
                If i > .RowCount Then .RowCount = i
                .Row = i
                .RowLabel = i
                .Data = RS("jumlah")
                RS.MoveNext
            End If
        Else
            If i > .RowCount Then .RowCount = i
            .Row = i
            .RowLabel = i
            .Data = 0
        End If
        i = i + 1
    Wend
End With
RS.Close
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim s As String

MC1.RowCount = 0
Set RS = New ADODB.Recordset
On Error GoTo KirimErr

cbBulan.Clear
RS.Open "select date_format(waktu,'%M %Y') as tgl from pesan order by waktu asc", OC
s = ""
While Not RS.EOF
    If s <> RS("tgl") Then cbBulan.AddItem RS("tgl")
    s = RS("tgl")
    RS.MoveNext
Wend
RS.Close

If cbBulan.ListCount > 0 Then cbBulan.ListIndex = cbBulan.ListCount - 1

Exit Sub
KirimErr:
    TampilkanPesan "Gagal membaca database" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RS = Nothing
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub
